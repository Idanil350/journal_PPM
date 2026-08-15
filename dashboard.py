"""Tableau de bord Streamlit : upload du PPM (n'importe quelle date) ->
extraction des marchés IT/Numérique/COLEPS -> Business Intelligence pour
DAREDAB (répertoire des autorités, budget, sources de financement)."""

import os
import tempfile
from datetime import date, datetime, timedelta
from pathlib import Path

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
        disabled="ppm_results" not in st.session_state,
        help=f'Basé sur la colonne "{ATTRIBUTION_COLUMN}" : un marché dont la date '
             "d'attribution prévue est déjà passée est masqué -- une fois attribué, "
             "l'opportunité de soumissionner est perdue. Les dates vides ou illisibles "
             "sont conservées par prudence.",
    )
    show_new_only = st.toggle(
        "Nouveaux et modifiés uniquement",
        value=True,
        disabled=collection is None or "ppm_results" not in st.session_state,
        help="Compare avec les éditions précédentes du PPM déjà chargées : masque les "
             "marchés déjà vus et inchangés depuis la dernière fois.",
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
                st.dataframe(
                    changes_df,
                    hide_index=True,
                    column_config={
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
            st.subheader("Marchés jamais vus par DAREDAB depuis une date")
            st.caption(
                "Autre question, différente de celle ci-dessus : quels marchés IT DAREDAB n'avait "
                "encore jamais détectés dans aucune édition précédente du PPM."
            )
            since_seen = st.date_input(
                "Jamais vu depuis quelle date", key="since_seen",
                value=date.today() - timedelta(days=14), format="DD/MM/YYYY",
            )
            seen_results = db.get_markets_since(collection, since_seen)
            st.metric(
                f"Marchés jamais vus, détectés depuis le {since_seen.strftime('%d/%m/%Y')}",
                len(seen_results), border=True,
            )
            if seen_results:
                seen_df = results_to_dataframe(seen_results)
                st.dataframe(
                    seen_df,
                    hide_index=True,
                    column_config={
                        "Désignation et localisation du projet": st.column_config.TextColumn(pinned=True),
                        "Montant prévisionnel (FCFA)": st.column_config.NumberColumn(format="%d"),
                    },
                )
                st.download_button(
                    "Télécharger en Excel",
                    data=to_excel_bytes(seen_df),
                    file_name=f"nouveaux_marches_depuis_{since_seen.isoformat()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    icon=":material/download:",
                    key="dl_seen",
                )
            else:
                st.info("Aucun nouveau marché détecté depuis cette date.", icon=":material/task_alt:")


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
        finally:
            tmp_path.unlink(missing_ok=True)

# Les résultats vivent dans session_state, indépendamment du widget d'upload
# -- rechoisir/effacer le fichier ne fait pas disparaître la dernière analyse.
if "ppm_results" not in st.session_state:
    st.info("Dépose un fichier PDF ci-dessus pour lancer l'extraction.", icon=":material/upload_file:")
    st.stop()

results = st.session_state["ppm_results"]
total_pages = st.session_state["ppm_total_pages"]
edition_date = st.session_state.get("ppm_edition_date")

if not results:
    st.warning(f"{total_pages} pages analysées -- aucune ligne ne correspond aux mots-clés IT/COLEPS.")
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

# Le statut Nouveau/Modifié/Déjà vu a déjà été calculé et enregistré lot par
# lot pendant l'extraction (voir on_batch ci-dessus) -- pas besoin de
# reclassifier ici, ça compterait tout comme "déjà vu" puisque ça vient
# d'être stocké. On se contente de compter/filtrer sur le Statut déjà présent.
new_count = updated_count = None
if collection is not None:
    new_count = sum(1 for r in results if r.get("Statut") == db.STATUS_NEW)
    updated_count = sum(1 for r in results if r.get("Statut") == db.STATUS_UPDATED)
    if show_new_only:
        before_new_filter = len(results)
        results = [r for r in results if r.get("Statut") in (db.STATUS_NEW, db.STATUS_UPDATED)]
        if not results:
            st.info(
                f"Les {before_new_filter} marché(s) IT retenus sont déjà connus et inchangés "
                "depuis la dernière édition chargée. Désactive le filtre dans la barre latérale "
                "pour les revoir.",
                icon=":material/task_alt:",
            )
            st.stop()

analytics = PPMAnalytics(results)
kpis = analytics.kpis()

with st.container(horizontal=True):
    st.metric("Pages analysées", f"{total_pages:,}".replace(",", " "), border=True)
    st.metric("Marchés IT retenus", kpis["lignes"], border=True)
    st.metric("Autorités distinctes", kpis["autorites"], border=True)
    st.metric("Budget total identifié", fmt_fcfa(kpis["budget_total"]), border=True)
    if new_count is not None:
        st.metric("Nouveaux depuis la dernière édition", new_count, border=True)

if edition_date:
    st.caption(
        f":material/event: Édition du PPM détectée : {edition_date.strftime('%d/%m/%Y')} -- "
        "c'est cette date qui sert de référence pour le suivi des nouveautés, pas la date d'upload."
    )
elif collection is not None:
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

    col_a, col_b = st.columns(2)
    with col_a.container(border=True, height="stretch"):
        st.subheader("Type d'acheteur")
        st.caption("Administration centrale vs. communes vs. établissements publics/entreprises.")
        buyer_type_df = analytics.buyer_type_breakdown()
        st.dataframe(buyer_type_df, hide_index=True)
        if not buyer_type_df.empty:
            st.bar_chart(
                buyer_type_df, x="Type d'acheteur", y="Lignes", horizontal=True, x_label="Lignes",
            )
    with col_b.container(border=True, height="stretch"):
        st.subheader("Source de financement")
        funding_df = analytics.funding_breakdown()
        st.dataframe(funding_df, hide_index=True)
        if not funding_df.empty:
            st.bar_chart(
                funding_df, x="Source de financement", y="Lignes", horizontal=True, x_label="Lignes",
            )

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
            st.bar_chart(category_df, x="Catégorie", y="Lignes", horizontal=True, x_label="Lignes")
