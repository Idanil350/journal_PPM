"""Tableau de bord Streamlit : upload du PPM (n'importe quelle date) ->
extraction des marchés IT/Numérique/COLEPS -> Business Intelligence pour
DAREDAB (répertoire des autorités, budget, sources de financement)."""

import os
import tempfile
from datetime import date, timedelta
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
tab_data, tab_bi, tab_history = st.tabs([
    ":material/table_chart: Lignes extraites",
    ":material/insights: Business intelligence",
    ":material/history: Historique",
])

with tab_history:
    st.subheader("Nouveaux marchés depuis une date")
    st.caption(
        "Répond directement à « quels sont les nouveaux projets depuis le [date] ? » -- "
        "lit la mémoire déjà enregistrée, pas besoin d'avoir chargé un journal dans cette session."
    )
    if collection is None:
        st.info(
            "Mémoire entre éditions désactivée -- configure MONGO_URI pour activer cet historique.",
            icon=":material/info:",
        )
    else:
        since = st.date_input(
            "Depuis quelle date", value=date.today() - timedelta(days=14), format="DD/MM/YYYY",
        )
        history_results = db.get_markets_since(collection, since)
        st.metric(
            f"Nouveaux marchés depuis le {since.strftime('%d/%m/%Y')}",
            len(history_results), border=True,
        )
        if history_results:
            history_df = results_to_dataframe(history_results)
            st.dataframe(
                history_df,
                hide_index=True,
                column_config={
                    "Désignation et localisation du projet": st.column_config.TextColumn(pinned=True),
                    "Montant prévisionnel (FCFA)": st.column_config.NumberColumn(format="%d"),
                },
            )
            st.download_button(
                "Télécharger en Excel",
                data=to_excel_bytes(history_df),
                file_name=f"nouveaux_marches_depuis_{since.isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
            )
        else:
            st.info("Aucun nouveau marché détecté depuis cette date.", icon=":material/task_alt:")

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

            results = extract_pdf(tmp_path, progress_every=25, on_progress=on_progress)

            progress.empty()
            status.empty()
            st.session_state["ppm_results"] = results
            st.session_state["ppm_total_pages"] = total_pages
        finally:
            tmp_path.unlink(missing_ok=True)

# Les résultats vivent dans session_state, indépendamment du widget d'upload
# -- rechoisir/effacer le fichier ne fait pas disparaître la dernière analyse.
if "ppm_results" not in st.session_state:
    st.info("Dépose un fichier PDF ci-dessus pour lancer l'extraction.", icon=":material/upload_file:")
    st.stop()

results = st.session_state["ppm_results"]
total_pages = st.session_state["ppm_total_pages"]

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

# Comparaison avec les éditions précédentes du PPM déjà chargées -- appliquée
# APRÈS le filtrage IT/en-cours, jamais sur le contenu lui-même. Sans base
# configurée, tout est marqué "Nouveau" par défaut (comportement identique à
# avant l'ajout de cette mémoire).
new_count = updated_count = None
if collection is not None:
    results = db.classify_and_store(collection, results)
    new_count = sum(1 for r in results if r["Statut"] == db.STATUS_NEW)
    updated_count = sum(1 for r in results if r["Statut"] == db.STATUS_UPDATED)
    if show_new_only:
        before_new_filter = len(results)
        results = [r for r in results if r["Statut"] in (db.STATUS_NEW, db.STATUS_UPDATED)]
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

if hide_past and hidden_count:
    st.caption(f":material/schedule: {hidden_count} marché(s) déjà attribué(s) masqué(s).")

with tab_data:
    with st.container(border=True):
        st.subheader("Résultats")
        df = results_to_dataframe(results)
        if new_count is not None:
            df.insert(0, "Statut", [r["Statut"] for r in results])
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
