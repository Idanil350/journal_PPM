"""Tableau de bord Streamlit : upload du PPM (n'importe quelle date) ->
extraction des marchés IT/Numérique/COLEPS -> Business Intelligence pour
DAREDAB (répertoire des autorités, budget, sources de financement)."""

import os
import tempfile
from datetime import date, datetime, timedelta
from pathlib import Path

import altair as alt
import pandas as pd
import pdfplumber
import streamlit as st
from dotenv import load_dotenv

import db
from ppm_extraction import (
    ATTRIBUTION_COLUMN,
    PPMAnalytics,
    extract_pdf,
    filter_ongoing,
    results_to_dataframe,
    to_excel_bytes,
)

load_dotenv()


def fmt_fcfa(amount) -> str:
    """25017848460 -> "25,02 Md FCFA" ; plus lisible qu'un nombre brut sur une carte KPI."""
    if amount is None:
        return "—"
    amount = int(amount)
    if amount >= 1_000_000_000:
        return f"{amount / 1_000_000_000:.2f} Md FCFA".replace(".", ",")
    if amount >= 1_000_000:
        return f"{amount / 1_000_000:.1f} M FCFA".replace(".", ",")
    return f"{amount:,}".replace(",", " ") + " FCFA"


def priority_bar_chart(df: pd.DataFrame, category_col: str, count_col: str, budget_col: str):
    """Graphique Altair : barres triées par volume, couleur selon le budget
    total (deuxième dimension visible d'un coup d'œil), infobulle complète
    au survol -- remplace les st.bar_chart plats qui ne montraient que le
    nombre de lignes, sans le poids financier de chaque catégorie."""
    chart = (
        alt.Chart(df)
        .mark_bar()
        .encode(
            x=alt.X(f"{count_col}:Q", title="Nombre de marchés"),
            y=alt.Y(f"{category_col}:N", sort="-x", title=None),
            color=alt.Color(
                f"{budget_col}:Q", title="Budget total (FCFA)",
                scale=alt.Scale(scheme="blues"),
                legend=alt.Legend(format=","),
            ),
            tooltip=[
                alt.Tooltip(f"{category_col}:N", title="Catégorie"),
                alt.Tooltip(f"{count_col}:Q", title="Marchés"),
                alt.Tooltip(f"{budget_col}:Q", title="Budget total (FCFA)", format=","),
            ],
        )
        .properties(height=alt.Step(28))
    )
    st.altair_chart(chart)


st.set_page_config(
    page_title="DAREDAB PPM Intelligence",
    page_icon=":material/query_stats:",
    layout="wide",
)


def _get_expected_password():
    """Cherche APP_PASSWORD dans st.secrets (Streamlit Community Cloud) puis
    dans les variables d'environnement (.env en local, secret d'une autre
    plateforme d'hébergement) -- jamais codé en dur dans le fichier. Utilise
    .get() plutôt que l'opérateur "in" -- plus sûr face à st.secrets quand
    aucun secrets.toml n'existe (comportement qui varie selon le contexte)."""
    try:
        value = st.secrets.get("APP_PASSWORD")
        if value:
            return value
    except Exception:
        pass  # Pas de secrets.toml configuré -- normal en local/hors Streamlit Cloud.
    return os.environ.get("APP_PASSWORD")


def check_password() -> bool:
    """Porte d'accès par mot de passe partagé -- suffisant pour un outil interne
    à faible trafic, pas un vrai système multi-compte."""
    if st.session_state.get("authenticated"):
        return True

    expected = _get_expected_password()
    if not expected:
        # Pas de mot de passe configuré (ex: dev local sans .env) -- ne pas
        # bloquer silencieusement l'accès à cause d'une variable manquante,
        # mais prévenir clairement que l'app tourne sans protection.
        st.warning(
            "APP_PASSWORD n'est pas configuré -- l'application tourne sans mot de passe.",
            icon=":material/warning:",
        )
        st.session_state["authenticated"] = True
        return True

    st.title("PPM Intelligence")
    st.caption("Accès réservé à l'équipe DAREDAB.")
    with st.form("login", border=True):
        password = st.text_input("Mot de passe", type="password")
        submitted = st.form_submit_button("Se connecter", type="primary", icon=":material/login:")
    if submitted:
        if password == expected:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Mot de passe incorrect.", icon=":material/error:")
    return False


if not check_password():
    st.stop()

st.title("PPM Intelligence")
st.caption(
    "Marchés publics IT, Numérique et COLEPS -- dépose le Journal de programmation des "
    "marchés (PDF, n'importe quelle date) pour extraire automatiquement les lignes "
    "pertinentes et lire la Business Intelligence associée."
)

collection = db.get_collection()

with st.sidebar:
    st.caption("Filtres")
    hide_past = st.toggle(
        "Marchés pas encore attribués uniquement",
        value=True,
        help=f'Basé sur la colonne "{ATTRIBUTION_COLUMN}" : un marché dont la date '
             "d'attribution prévue est déjà passée est masqué -- une fois attribué, "
             "l'opportunité de soumissionner est perdue. Les dates vides ou illisibles "
             "sont conservées par prudence.",
    )
    show_new_only = st.toggle(
        "Nouveaux et modifiés uniquement",
        value=True,
        disabled=collection is None,
        help="Compare avec la dernière fois que quelqu'un a consulté l'app : masque les "
             "marchés déjà vus et inchangés. Fonctionne dès l'ouverture, sans upload.",
    )
    if collection is None:
        st.caption(
            ":material/info: Mémoire entre éditions désactivée -- configure MONGO_URI "
            "pour comparer avec les journaux déjà chargés.",
        )
    st.caption("DAREDAB · outil d'extraction PPM")

# tab_history est déclaré ici (avant le upload) et rendu tout de suite --
# il répond à "quels sont les nouveaux marchés depuis telle date ?" en lisant
# directement la mémoire MongoDB, sans dépendre d'une extraction lancée dans
# la session en cours. tab_data/tab_bi, eux, ont besoin des résultats d'une
# extraction et sont remplis plus bas dans le script.
#
# on_change="rerun" active .open : sans ça, TOUS les onglets s'exécutent à
# CHAQUE interaction n'importe où dans l'app (upload, filtre...), donc
# mark_consultation() avancerait le repère "dernière consultation" à chaque
# clic ailleurs -- pas seulement quand quelqu'un regarde vraiment cet onglet.
tab_data, tab_bi, tab_history = st.tabs(
    [
        ":material/table_chart: Lignes extraites",
        ":material/insights: Business intelligence",
        ":material/history: Historique",
    ],
    on_change="rerun",
    key="ppm_active_tab",
)

def render_priority_panel():
    """Panneau de décision principal de l'onglet Business Intelligence :
    marchés encore ouverts classés par potentiel commercial, avec une
    action NOMMÉE (pas un score seul à interpréter) et un bouton pour
    enregistrer la décision de DAREDAB. DAREDAB est encore jeune sur ce
    secteur (~5 ans) -- ce panneau sert à observer où se positionner, pas à
    prédire un résultat garanti. Appelé seulement quand l'onglet BI est
    réellement affiché (voir garde `if tab_bi.open` plus bas)."""
    st.subheader(":material/lightbulb: Priorités commerciales")
    st.caption(
        "Marchés encore ouverts, classés par potentiel pour DAREDAB (financement, "
        "catégorie IT, urgence du lancement, acheteur récurrent). Choisis une décision "
        "pour garder une trace -- ça sert aussi à construire l'historique de résultats."
    )
    if collection is None:
        st.info(
            "Mémoire désactivée -- configure MONGO_URI pour activer ce panneau.",
            icon=":material/info:",
        )
        return

    opportunities = db.get_priority_opportunities(collection, limit=10)
    if not opportunities:
        st.info("Aucune opportunité à prioriser pour l'instant.", icon=":material/task_alt:")
        return

    decision_labels = {
        "go": "✅ Décidé : on y va",
        "no-go": "❌ Décidé : on laisse tomber",
    }
    rows = [{
        "Marché": opp.get("Désignation et localisation du projet", ""),
        "Autorité": opp.get("Autorité (en-tête de page)", ""),
        "Action recommandée": opp["_action"],
        "Urgent": "Oui" if opp["_urgent"] else "",
        "Acheteur récurrent": "Oui" if opp["_recurrent"] else "",
        "Décision": decision_labels.get(opp.get("_decision"), "-- à décider --"),
        "Choisir": [":material/thumb_up: On y va", ":material/thumb_down: On laisse tomber"],
        "_market_key": opp.get("_market_key"),
    } for opp in opportunities]
    priority_df = pd.DataFrame(rows)

    def handle_priority_decision():
        click = st.session_state.get("priority_decision")
        if not click:
            return
        row_key = priority_df.iloc[click["row"]]["_market_key"]
        decision = "go" if "on y va" in click["label"].lower() else "no-go"
        db.set_decision(collection, row_key, decision)
        st.toast(f"Décision enregistrée : {click['label']}", icon=":material/check:")

    st.dataframe(
        priority_df.drop(columns=["_market_key"]),
        hide_index=True,
        column_config={
            "Marché": st.column_config.TextColumn(pinned=True),
            "Choisir": st.column_config.ButtonColumn(
                "Décision", on_click=handle_priority_decision, key="priority_decision",
            ),
        },
    )


def render_history_tab():
    """Corps de l'onglet Historique, appelé seulement quand il est réellement
    affiché (voir garde `if tab_history.open` plus bas) -- sans ça,
    mark_consultation() avancerait le repère "dernière consultation" à
    chaque interaction ailleurs dans l'app, pas seulement quand quelqu'un
    regarde vraiment cet onglet."""
    if collection is None:
        st.info(
            "Mémoire entre éditions désactivée -- configure MONGO_URI pour activer cet historique.",
            icon=":material/info:",
        )
    else:
        with st.container(border=True):
            st.subheader("Nouveautés depuis la dernière consultation")
            st.caption(
                "Ce que le DG a demandé explicitement : les projets ajoutés ou modifiés "
                "depuis la dernière fois que quelqu'un chez DAREDAB a consulté cette page -- "
                "automatique, rien à taper. Un repère partagé pour toute l'équipe."
            )
            meta_collection = db.get_meta_collection()
            last_consultation = db.get_last_consultation(meta_collection)
            changes = db.get_changes_since(collection, last_consultation or "0000-00-00")

            if last_consultation:
                last_dt = datetime.fromisoformat(last_consultation)
                st.caption(f"Dernière consultation : {last_dt.strftime('%d/%m/%Y à %H:%M')} (heure UTC).")
            else:
                st.caption("Première consultation -- tout ce qui est en mémoire est affiché.")

            st.metric("Nouveautés depuis la dernière consultation", len(changes), border=True)
            if changes:
                changes_df = results_to_dataframe(changes)
                changes_df.insert(0, "Statut", [c.get("Statut") for c in changes])
                st.dataframe(
                    changes_df,
                    hide_index=True,
                    column_config={
                        "Statut": st.column_config.TextColumn(pinned=True),
                        "Désignation et localisation du projet": st.column_config.TextColumn(pinned=True),
                        "Montant prévisionnel (FCFA)": st.column_config.NumberColumn(format="%d"),
                    },
                )
                st.download_button(
                    "Télécharger en Excel",
                    data=to_excel_bytes(changes_df),
                    file_name="nouveautes_depuis_derniere_consultation.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    icon=":material/download:",
                    key="dl_changes",
                )

                modified = [c for c in changes if c.get("Statut") == db.STATUS_UPDATED and c.get("_diff")]
                if modified:
                    st.markdown("**Détail des modifications** -- quel champ a changé, et vers quoi :")
                    for m in modified:
                        title = m.get("Désignation et localisation du projet", "")[:80]
                        with st.expander(f":material/edit: {title}"):
                            for d in m["_diff"]:
                                st.markdown(f"- **{d['champ']}** : {d['avant']} → {d['après']}")
            else:
                st.info("Rien de nouveau depuis la dernière consultation.", icon=":material/task_alt:")

            # Marquer la consultation APRÈS avoir lu/affiché -- jamais avant,
            # sinon on écrase le repère avant de s'en être servi.
            db.mark_consultation(meta_collection)

        with st.container(border=True):
            st.subheader("Projets dont le lancement de consultation démarre depuis une date")
            st.caption(
                "Répond directement à « quels sont les nouveaux projets depuis le [date] ? » au "
                "sens du DG : filtre sur la colonne « Lancement de consultation / Invitation à "
                "soumissionner » -- pas besoin d'avoir chargé un journal dans cette session."
            )
            since_launch = st.date_input(
                "Lancement de consultation depuis quelle date", key="since_launch",
                value=date.today() - timedelta(days=14), format="DD/MM/YYYY",
            )
            launch_results = db.get_markets_by_launch_date(collection, since_launch)
            st.metric(
                f"Projets dont le lancement démarre depuis le {since_launch.strftime('%d/%m/%Y')}",
                len(launch_results), border=True,
            )
            if launch_results:
                launch_df = results_to_dataframe(launch_results)
                st.dataframe(
                    launch_df,
                    hide_index=True,
                    column_config={
                        "Désignation et localisation du projet": st.column_config.TextColumn(pinned=True),
                        "Montant prévisionnel (FCFA)": st.column_config.NumberColumn(format="%d"),
                    },
                )
                st.download_button(
                    "Télécharger en Excel",
                    data=to_excel_bytes(launch_df),
                    file_name=f"lancement_consultation_depuis_{since_launch.isoformat()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    icon=":material/download:",
                    key="dl_launch",
                )
            else:
                st.info("Aucun projet avec un lancement de consultation prévu depuis cette date.", icon=":material/task_alt:")

        with st.container(border=True):
            st.subheader("Tous les projets IT connus")
            st.caption(
                "L'ensemble des marchés IT actuellement en mémoire, toutes éditions du "
                "journal confondues -- pas besoin d'avoir chargé un journal dans cette session."
            )
            all_markets = db.get_all_markets(collection, last_consultation)
            st.metric("Projets IT en mémoire", len(all_markets), border=True)
            if all_markets:
                all_df = results_to_dataframe(all_markets)
                all_df.insert(0, "Statut", [m.get("Statut") for m in all_markets])
                st.dataframe(
                    all_df,
                    hide_index=True,
                    column_config={
                        "Statut": st.column_config.TextColumn(pinned=True),
                        "Désignation et localisation du projet": st.column_config.TextColumn(pinned=True),
                        "Montant prévisionnel (FCFA)": st.column_config.NumberColumn(format="%d"),
                    },
                )
                st.download_button(
                    "Télécharger en Excel",
                    data=to_excel_bytes(all_df),
                    file_name="tous_les_projets_it.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    icon=":material/download:",
                    key="dl_all",
                )
            else:
                st.info("Aucun projet IT en mémoire pour l'instant.", icon=":material/task_alt:")


with tab_bi:
    if tab_bi.open:
        render_priority_panel()

with tab_history:
    if tab_history.open:
        render_history_tab()

uploaded_file = st.file_uploader("Journal de programmation des marchés (PDF)", type=["pdf"])

if uploaded_file is not None:
    run_clicked = st.button("Extraire et analyser", type="primary", icon=":material/search:")

    if run_clicked:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp.write(uploaded_file.getbuffer())
            tmp_path = Path(tmp.name)

        progress = st.progress(0.0, text="Ouverture du PDF...")
        status = st.empty()

        try:
            with pdfplumber.open(tmp_path) as pdf:
                total_pages = len(pdf.pages)

            def on_progress(page_number, total, found_so_far):
                progress.progress(page_number / total, text=f"Page {page_number}/{total} traitée...")
                status.caption(f"{found_so_far} ligne(s) pertinente(s) trouvée(s) jusque-là.")

            # Sauvegarde en mémoire lot par lot, PENDANT l'extraction -- pas
            # seulement à la toute fin. Un fichier de 1770 pages peut prendre
            # plusieurs minutes ; sur mobile, si la connexion tombe en cours
            # de route (écran verrouillé, changement d'appli), tout ce qui a
            # déjà été traité reste en sécurité au lieu d'être perdu.
            tagged_accumulator = []

            def on_batch(batch_records, batch_edition_date):
                if collection is not None:
                    tagged_accumulator.extend(
                        db.classify_and_store(collection, batch_records, edition_date=batch_edition_date)
                    )
                else:
                    tagged_accumulator.extend(batch_records)

            _raw_results, edition_date = extract_pdf(
                tmp_path, progress_every=25, on_progress=on_progress, on_batch=on_batch,
            )

            progress.empty()
            status.empty()
            st.session_state["ppm_results"] = tagged_accumulator
            st.session_state["ppm_total_pages"] = total_pages
            st.session_state["ppm_edition_date"] = edition_date
            if not tagged_accumulator:
                st.warning(f"{total_pages} pages analysées -- aucune ligne ne correspond aux mots-clés IT/COLEPS.")
        finally:
            tmp_path.unlink(missing_ok=True)

# Affichage principal : toujours depuis la mémoire persistante (pas depuis
# la session en cours) -- pour que les filtres marchent dès l'ouverture de
# l'app, sans upload, et restent cohérents avec l'onglet Historique (même
# repère "dernière consultation" partagé). Repli sur les résultats de cette
# session seulement si MONGO_URI n'est pas configuré du tout.
meta_collection = db.get_meta_collection()
last_consultation = db.get_last_consultation(meta_collection)
total_pages = st.session_state.get("ppm_total_pages")
edition_date = st.session_state.get("ppm_edition_date")

if collection is not None:
    results = db.get_all_markets(collection, last_consultation)
else:
    results = st.session_state.get("ppm_results")

if not results:
    st.info(
        "Aucun marché IT en mémoire pour l'instant -- dépose un PDF ci-dessus pour lancer "
        "une première extraction.",
        icon=":material/upload_file:",
    )
    st.stop()

hidden_count = 0
if hide_past:
    ongoing_results = filter_ongoing(results)
    hidden_count = len(results) - len(ongoing_results)
    results = ongoing_results
    if not results:
        st.warning(
            "Tous les marchés IT retenus sont déjà attribués. Désactive le filtre "
            "dans la barre latérale pour les revoir.",
            icon=":material/warning:",
        )
        st.stop()

new_count = updated_count = None
if collection is not None:
    new_count = sum(1 for r in results if r.get("Statut") == db.STATUS_NEW)
    updated_count = sum(1 for r in results if r.get("Statut") == db.STATUS_UPDATED)
    if show_new_only:
        before_new_filter = len(results)
        results = [r for r in results if r.get("Statut") in (db.STATUS_NEW, db.STATUS_UPDATED)]
        if not results:
            st.info(
                f"Les {before_new_filter} marché(s) IT en mémoire sont déjà connus et inchangés "
                "depuis la dernière consultation. Désactive le filtre dans la barre latérale "
                "pour les revoir.",
                icon=":material/task_alt:",
            )
            st.stop()

analytics = PPMAnalytics(results)
kpis = analytics.kpis()

with st.container(horizontal=True):
    if total_pages is not None:
        st.metric("Pages analysées (dernier upload)", f"{total_pages:,}".replace(",", " "), border=True)
    st.metric("Marchés IT retenus", kpis["lignes"], border=True)
    st.metric("Autorités distinctes", kpis["autorites"], border=True)
    st.metric("Budget total identifié", fmt_fcfa(kpis["budget_total"]), border=True)
    if new_count is not None:
        st.metric("Nouveaux/modifiés depuis la dernière consultation", new_count + updated_count, border=True)

if edition_date:
    st.caption(
        f":material/event: Édition du PPM détectée : {edition_date.strftime('%d/%m/%Y')} -- "
        "c'est cette date qui sert de référence pour le suivi des nouveautés, pas la date d'upload."
    )
elif collection is not None and total_pages is not None:
    st.caption(
        ":material/warning: Date de l'édition introuvable dans ce PDF -- la date d'upload est "
        "utilisée en repli pour le suivi des nouveautés (moins fiable si l'upload est tardif)."
    )

if hide_past and hidden_count:
    st.caption(f":material/schedule: {hidden_count} marché(s) déjà attribué(s) masqué(s).")

with tab_data:
    with st.container(border=True):
        st.subheader("Résultats")
        df = results_to_dataframe(results)
        if new_count is not None:
            df.insert(0, "Statut", [r.get("Statut") for r in results])
        st.dataframe(
            df,
            hide_index=True,
            column_config={
                "Statut": st.column_config.TextColumn(pinned=True),
                "Désignation et localisation du projet": st.column_config.TextColumn(pinned=True),
                "Montant prévisionnel (FCFA)": st.column_config.NumberColumn(format="%d"),
            },
        )

    with st.container(horizontal=True):
        st.download_button(
            "Télécharger en Excel",
            data=to_excel_bytes(df),
            file_name="marches_it_coleps_filtres.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
        )
        st.download_button(
            "Télécharger en CSV",
            data=df.to_csv(index=False).encode("utf-8-sig"),
            file_name="marches_it_coleps_filtres.csv",
            mime="text/csv",
            icon=":material/download:",
        )

with tab_bi:
    with st.container(border=True):
        st.subheader("Répertoire des autorités contractantes")
        st.caption(
            "Chaque ministère/autorité rencontré dans les lignes IT retenues, avec son "
            "volume et son budget prévisionnel cumulé."
        )
        authority_df = analytics.authority_directory()
        st.dataframe(
            authority_df,
            hide_index=True,
            column_config={
                "Autorité": st.column_config.TextColumn(pinned=True),
                "Budget total (FCFA)": st.column_config.NumberColumn(format="%d"),
            },
        )
        if not authority_df.empty:
            top_authorities = authority_df.nlargest(10, "Budget total (FCFA)")
            priority_bar_chart(top_authorities, "Autorité", "Lignes", "Budget total (FCFA)")

    col_a, col_b = st.columns(2)
    with col_a.container(border=True, height="stretch"):
        st.subheader("Type d'acheteur")
        st.caption("Administration centrale vs. communes vs. établissements publics/entreprises.")
        buyer_type_df = analytics.buyer_type_breakdown()
        st.dataframe(buyer_type_df, hide_index=True)
        if not buyer_type_df.empty:
            priority_bar_chart(buyer_type_df, "Type d'acheteur", "Lignes", "Budget total (FCFA)")
    with col_b.container(border=True, height="stretch"):
        st.subheader("Source de financement")
        funding_df = analytics.funding_breakdown()
        st.dataframe(funding_df, hide_index=True)
        if not funding_df.empty:
            priority_bar_chart(funding_df, "Source de financement", "Lignes", "Budget total (FCFA)")

    with st.container(border=True):
        st.subheader("Segmentation par catégorie IT")
        st.caption(
            "Une ligne peut relever de plusieurs catégories (ex : développement logiciel + "
            "infrastructure) -- son budget est alors compté dans chacune, donc la somme des "
            "budgets par catégorie peut dépasser le budget total global."
        )
        category_df = analytics.category_breakdown()
        st.dataframe(
            category_df,
            hide_index=True,
            column_config={
                "Catégorie": st.column_config.TextColumn(pinned=True),
                "Budget total (FCFA)": st.column_config.NumberColumn(format="%d"),
            },
        )
        if not category_df.empty:
            priority_bar_chart(category_df, "Catégorie", "Lignes", "Budget total (FCFA)")
