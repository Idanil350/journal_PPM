"""
Extraction automatique des marches publics IT / Numerique / Formation
digitale (dont COLEPS) depuis le Journal de Programmation des Marches (PPM)
de l'ARMP, + agregation Business Intelligence (budget, autorites, sources
de financement).

Le PPM est un PDF genere par JasperReports : texte natif (pas un scan),
structure en tableau de 12 colonnes repete sur chaque page, precede d'un
en-tete indiquant l'Autorite Contractante concernee.

Verifie en direct sur "JPM au 10 Aout 2026.pdf" (1770 pages) avant d'etre
finalise. Fonctionne avec n'importe quelle date de journal tant que la
structure du tableau reste la meme (verifiable au nombre de lignes lues).

Point d'encodage important (verifie sur le fichier reel, pas suppose) :
ce PDF a ete genere avec un moteur ancien (iText 2.1.7) dont l'extraction
de texte remplace occasionnellement un caractere accentue ou une apostrophe
par le caractere de remplacement Unicode U+FFFD. Les regex ci-dessous
tolerent explicitement ce caractere a la place de chaque lettre accentuee.

Usage en CLI :
    python ppm_extraction.py "JPM au 10 Aout 2026.pdf"
    python ppm_extraction.py "JPM au 10 Aout 2026.pdf" -o resultat.xlsx
"""

import argparse
import gc
import logging
import re
import sys
import unicodedata
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import pandas as pd
import pdfplumber

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("ppm_extraction")

REPLACEMENT_CHAR = "�"

TABLE_COLUMNS = [
    "N°",
    "Désignation et localisation du projet",
    "Nature des prestations",
    "Montant prévisionnel (FCFA)",
    "Source de financement",
    "Autorité Contractante Compétente (rôle)",
    "Mode de consultation",
    "Lancement de consultation / Invitation à soumissionner",
    "Attribution du marché",
    "Signature du marché",
    "Démarrage des prestations",
    "Réception définitive des prestations",
]

EXTRA_COLUMNS = [
    "Page PDF", "Autorité (en-tête de page)", "Type d'acheteur",
    "Catégories détectées", "À relire (terme générique)",
]
ALL_COLUMNS = EXTRA_COLUMNS + TABLE_COLUMNS

# ---------------------------------------------------------------------------
# Regex tolerantes aux accents / apostrophes mal encodes.
# ---------------------------------------------------------------------------
_ACCENT_FOLD = {
    "é": f"[ée{REPLACEMENT_CHAR}]", "è": f"[èe{REPLACEMENT_CHAR}]",
    "ê": f"[êe{REPLACEMENT_CHAR}]", "ë": f"[ëe{REPLACEMENT_CHAR}]",
    "à": f"[àa{REPLACEMENT_CHAR}]", "â": f"[âa{REPLACEMENT_CHAR}]",
    "î": f"[îi{REPLACEMENT_CHAR}]", "ï": f"[ïi{REPLACEMENT_CHAR}]",
    "ô": f"[ôo{REPLACEMENT_CHAR}]", "ö": f"[öo{REPLACEMENT_CHAR}]",
    "û": f"[ûu{REPLACEMENT_CHAR}]", "ù": f"[ùu{REPLACEMENT_CHAR}]",
    "ü": f"[üu{REPLACEMENT_CHAR}]", "ç": f"[çc{REPLACEMENT_CHAR}]",
    "œ": f"(?:œ|oe|{REPLACEMENT_CHAR})",
    "'": f"['’{REPLACEMENT_CHAR}]?", "’": f"['’{REPLACEMENT_CHAR}]?",
    " ": r"\s+",
}


def fuzzy_pattern(phrase: str) -> str:
    return "".join(_ACCENT_FOLD.get(ch, re.escape(ch)) for ch in phrase)


class Keyword:
    """Motif de recherche avec sensibilite a la casse explicite -- plus
    fiable que de la deviner a partir de la forme du pattern regex."""

    __slots__ = ("pattern", "case_sensitive", "label", "context_required")

    def __init__(self, pattern: str, case_sensitive: bool, label: str, context_required: bool = False):
        self.pattern = re.compile(pattern)
        self.case_sensitive = case_sensitive
        self.label = label
        # True pour un mot ambigu en français (sens IT ET sens courant non-IT)
        # -- un match seul ne suffit pas, il doit co-occurer avec un marqueur
        # IT non ambigu dans la meme ligne. Voir _IT_CONTEXT_MARKERS.
        self.context_required = context_required

    def matches(self, original_text: str, lowered_text: str) -> bool:
        haystack = original_text if self.case_sensitive else lowered_text
        return self.pattern.search(haystack) is not None


def word(phrase: str, context_required: bool = False) -> Keyword:
    pattern = r"(?<!\w)" + fuzzy_pattern(phrase.lower()) + r"s?(?!\w)"
    return Keyword(pattern, case_sensitive=False, label=phrase, context_required=context_required)


def phrase(text: str, context_required: bool = False) -> Keyword:
    pattern = r"(?<!\w)" + fuzzy_pattern(text.lower()) + r"(?!\w)"
    return Keyword(pattern, case_sensitive=False, label=text, context_required=context_required)


def acronym(acro: str, context_required: bool = False) -> Keyword:
    """Sigle court (IA, AMO, MOE, LMS, FOAD...) : recherche exacte en
    majuscules, sensible a la casse -- en minuscule ces sigles matcheraient
    beaucoup trop de mots courants pour etre fiables."""
    pattern = r"(?<!\w)" + re.escape(acro) + r"(?!\w)"
    return Keyword(pattern, case_sensitive=True, label=acro, context_required=context_required)


KEYWORD_CATEGORIES: dict[str, list[Keyword]] = {
    "Développement logiciel": [
        phrase("développement d'application"), word("logiciel"), word("progiciel"),
        phrase("plateforme numérique"), phrase("site web"),
        phrase("système d'information"), phrase("base de données"),
        # "application" seule et "plateforme" seule sont ambigues en francais
        # administratif : "Lycee d'Application" (etablissement scolaire),
        # "plateforme routiere"/"plateforme petroliere" (BTP/industrie) --
        # verifie en direct sur le PPM reel. Retenues seulement avec un
        # marqueur IT non ambigu dans la meme ligne.
        word("application", context_required=True),
        word("plateforme", context_required=True),
    ],
    "Innovation / IA / dématérialisation": [
        word("interopérabilité"), word("dématérialisation"),
        phrase("intelligence artificielle"), acronym("IA"),
    ],
    "Infrastructure IT": [
        phrase("matériel informatique"), word("serveur"), phrase("réseau informatique"),
        word("cybersécurité"), word("cloud"),
        phrase("maintenance informatique"), phrase("maintenance applicative"),
        # "hébergement" seul est ambigu : "bâtiment d'hébergement" d'un
        # centre d'accueil est un dortoir, pas de l'hébergement web/serveur.
        # Verifie en direct sur le PPM reel.
        word("hébergement", context_required=True),
    ],
    # Termes generiques : fort risque de faux positifs (utilises dans TOUS
    # les secteurs -- BTP, sante...). Conserves car demandes explicitement,
    # mais soumis au garde-fou context_required.
    "Conseil / Études IT (termes génériques -- à relire)": [
        phrase("audit informatique"), word("étude", context_required=True),
        phrase("schéma directeur", context_required=True),
        phrase("maîtrise d'ouvrage", context_required=True), acronym("AMO", context_required=True),
        phrase("maîtrise d'œuvre", context_required=True), acronym("MOE", context_required=True),
        phrase("cahier des charges", context_required=True),
    ],
    "Formation digitale": [
        phrase("formation informatique"), word("e-learning"), word("elearning"),
        word("moodle"), acronym("LMS"), acronym("FOAD"),
        phrase("contenu pédagogique numérique"),
    ],
    "COLEPS": [
        word("coleps"),
    ],
}

# Marqueurs IT non ambigus utilisés pour valider les mots-clés
# `context_required=True` ci-dessus. IMPORTANT : "application"/"plateforme"/
# "hébergement" n'y figurent PAS -- ils sont eux-mêmes ambigus, les inclure
# créerait une validation circulaire ("Lycée d'Application" validerait son
# propre match via lui-même).
_IT_CONTEXT_MARKERS = [
    word("informatique"), word("numérique"), word("digital"), word("digitale"),
    phrase("système d'information"), word("logiciel"),
    word("cybersécurité"), acronym("SI"), acronym("TIC"), word("coleps"),
]


def _has_it_context(text: str, lowered: str) -> bool:
    return any(kw.matches(text, lowered) for kw in _IT_CONTEXT_MARKERS)


def detect_categories(text: str):
    """Retourne (categories, uses_generic_term). `uses_generic_term` est
    True si au moins une categorie n'a ete confirmee qu'a travers un mot-cle
    ambigu (context_required) -- utile pour signaler les lignes a relire
    manuellement meme quand elles passent le garde-fou de contexte."""
    lowered = text.lower()
    it_context = None
    found = []
    uses_generic_term = False
    for category, keywords in KEYWORD_CATEGORIES.items():
        matched = False
        for kw in keywords:
            if not kw.matches(text, lowered):
                continue
            if kw.context_required:
                if it_context is None:
                    it_context = _has_it_context(text, lowered)
                if not it_context:
                    continue
                uses_generic_term = True
            matched = True
        if matched:
            found.append(category)
    return found, uses_generic_term


# ---------------------------------------------------------------------------
# Normalisation des champs (pour le tableau ET pour la BI)
# ---------------------------------------------------------------------------

_BUYER_TYPE_RULES = [
    ("Commune / Collectivité locale", [word("commune"), word("mairie"), phrase("conseil régional")]),
    (
        "Administration centrale",
        [word("ministère"), word("ministere"), phrase("présidence"), phrase("premier ministre"),
         phrase("services du premier ministre"), phrase("contrôle supérieure de l'état")],
    ),
]


def classify_buyer_type(authority_text: str) -> str:
    """Classe une autorité contractante en 3 grandes familles, pour la BI --
    utile pour distinguer les projets IT des ministères de ceux des
    communes, hors du cœur de cible DAREDAB (cf. app.rules.daredab_targets
    du projet de veille ARMP)."""
    if not authority_text:
        return "Non identifié"
    lowered = authority_text.lower()
    for label, keywords in _BUYER_TYPE_RULES:
        if any(kw.matches(authority_text, lowered) for kw in keywords):
            return label
    return "Établissement public / Entreprise"


def clean_authority_name(header: str) -> str:
    """Retire le code numerique final ("... - 3700000000") de l'en-tete de
    page pour un nom d'autorite lisible dans les tableaux de bord."""
    if not header:
        return "(Autorité non identifiée)"
    return re.sub(r"\s*-\s*\d{6,}\s*$", "", header).strip() or header


_MONTHS_FR = {
    "janvier": 1, "fevrier": 2, "mars": 3, "avril": 4, "mai": 5, "juin": 6,
    "juillet": 7, "aout": 8, "septembre": 9, "octobre": 10, "novembre": 11, "decembre": 12,
}

_EDITION_DATE_RE = re.compile(r"mise\s*a\s*jour\s*du\s+(\d{1,2})\s+([^\d,]+?)\s+(\d{4})")


def _fold_ascii(text: str) -> str:
    return unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii").lower()


def extract_edition_date(first_page_text: str):
    """Date de "MISE A JOUR DU [date]" imprimée en tête de la 1ère page du
    PPM -- la vraie date de publication de cette édition par l'ARMP,
    indépendante du jour où quelqu'un a pensé à uploader le fichier dans
    l'app. Utilisée comme référence pour "nouveau depuis telle date" au lieu
    de l'heure d'upload -- sinon un retard d'upload ferait passer un projet
    publié le 10 pour "nouveau depuis le 20". None si introuvable (texte
    fourni n'est pas celui de la 1ère page, ou format inattendu)."""
    if not first_page_text:
        return None
    match = _EDITION_DATE_RE.search(_fold_ascii(first_page_text))
    if not match:
        return None
    day, month_word, year = match.groups()
    month = _MONTHS_FR.get(month_word.strip())
    if month is None:
        return None
    try:
        return date(int(year), month, int(day))
    except ValueError:
        return None


def parse_amount(value: str):
    """"40 000 000" -> 40000000. None si aucun chiffre exploitable."""
    if not value:
        return None
    digits = re.sub(r"[^\d]", "", str(value))
    return int(digits) if digits else None


def parse_ppm_date(value):
    """"08/09/2026" -> date(2026, 9, 8). None si vide/illisible -- une date
    non-analysable n'est jamais traitee comme "passee" (voir filter_ongoing),
    pour ne jamais faire disparaitre une ligne a tort faute de savoir la lire."""
    if not value:
        return None
    try:
        return datetime.strptime(str(value).strip(), "%d/%m/%Y").date()
    except ValueError:
        return None


ATTRIBUTION_COLUMN = "Attribution du marché"


def filter_ongoing(results: list[dict], reference_date=None, date_column: str = ATTRIBUTION_COLUMN) -> list[dict]:
    """Ecarte les marchés déjà attribués (demande du DG : ne garder que ce
    qui est encore en cours). S'applique EN AVAL de detect_categories --
    c'est un filtre sur le calendrier des lignes déjà identifiées comme IT,
    jamais sur leur contenu. Une date vide ou illisible est conservée par
    prudence (considérée "encore en cours" plutôt que supprimée à tort)."""
    ref = reference_date or date.today()
    kept = []
    for record in results:
        attribution = parse_ppm_date(record.get(date_column))
        if attribution is not None and attribution < ref:
            continue
        kept.append(record)
    return kept


def extract_buyer_header(page_text: str) -> str:
    """La 1ere ligne de chaque page identifie l'Autorite Contractante --
    verifie en direct : repetee sur CHAQUE page, exploitable page par page."""
    return page_text.split("\n", 1)[0].strip() if page_text else ""


def clean_cell(value) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


# ---------------------------------------------------------------------------
# Extraction du PDF
# ---------------------------------------------------------------------------

def extract_pdf(pdf_path, progress_every: int = 100, on_progress=None, batch_size: int = 150):
    """Parcourt le PDF page par page et retourne la liste des lignes de
    marché dont le texte contient au moins un mot-clé cible.

    Traite le document par lots de `batch_size` pages, en refermant et
    rouvrant le PDF entre chaque lot -- `page.flush_cache()` seul ne suffit
    pas sur un très gros document : le cache de polices/objets embarqués de
    pdfminer est au niveau du DOCUMENT entier, pas de la page, et grossit
    sur toute la durée du parcours. Vérifié en direct : sans ce découpage,
    le processus montait à ~1,7 Go de RAM sur les 1770 pages du PPM réel,
    au-delà de la limite gratuite (~1 Go) de Streamlit Community Cloud --
    l'app plantait (tuée par manque de mémoire) sans erreur Python visible.

    Retourne `(results, edition_date)` -- `edition_date` est la date "MISE A
    JOUR DU ..." imprimée en tête de la 1ère page (None si introuvable),
    utilisée en aval comme référence pour "nouveau depuis telle date" au
    lieu de l'heure d'upload.

    `on_progress(page_number, total_pages, rows_found_so_far)`, si fourni,
    est appelé tous les `progress_every` pages (et sur la dernière) -- permet
    à une interface (CLI ou Streamlit) de piloter sa propre barre de
    progression sans dupliquer la boucle d'extraction."""
    results = []
    total_rows_seen = 0
    edition_date = None

    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
    logger.info(f"Ouverture du PDF : {total_pages} pages a analyser.")

    batch_start = 0
    while batch_start < total_pages:
        batch_end = min(batch_start + batch_size, total_pages)

        with pdfplumber.open(pdf_path) as pdf:
            for page_index in range(batch_start, batch_end):
                page = pdf.pages[page_index]
                page_number = page_index + 1
                page_text = page.extract_text() or ""
                buyer_header = extract_buyer_header(page_text)
                if page_number == 1:
                    edition_date = extract_edition_date(page_text)

                for table in page.extract_tables():
                    if not table or len(table) < 2:
                        continue
                    for row in table[1:]:
                        if not row or all(c is None or str(c).strip() == "" for c in row):
                            continue
                        total_rows_seen += 1
                        cleaned_row = [clean_cell(c) for c in row]
                        if len(cleaned_row) < len(TABLE_COLUMNS):
                            cleaned_row += [""] * (len(TABLE_COLUMNS) - len(cleaned_row))
                        elif len(cleaned_row) > len(TABLE_COLUMNS):
                            cleaned_row = cleaned_row[: len(TABLE_COLUMNS)]

                        row_text = " ".join(cleaned_row)
                        categories, uses_generic_term = detect_categories(row_text)
                        if not categories:
                            continue

                        record = dict(zip(TABLE_COLUMNS, cleaned_row))
                        record["Page PDF"] = page_number
                        record["Autorité (en-tête de page)"] = buyer_header
                        record["Type d'acheteur"] = classify_buyer_type(buyer_header)
                        record["Catégories détectées"] = ", ".join(categories)
                        record["À relire (terme générique)"] = "Oui" if uses_generic_term else ""
                        results.append(record)

                page.flush_cache()

                if on_progress and (page_number % progress_every == 0 or page_number == total_pages):
                    on_progress(page_number, total_pages, len(results))
                if page_number % progress_every == 0 or page_number == total_pages:
                    logger.info(
                        f"Page {page_number}/{total_pages} traitée "
                        f"({len(results)} lignes pertinentes trouvées jusque-là)."
                    )

        # Le "with" ci-dessus referme le PDF et libère le cache du lot --
        # gc.collect() force Python à récupérer cette mémoire immédiatement
        # plutôt que d'attendre le prochain passage du ramasse-miettes.
        gc.collect()
        batch_start = batch_end

    logger.info(
        f"Analyse terminée : {total_rows_seen} lignes de marché lues au total, "
        f"{len(results)} retenues (mention d'un mot-clé IT/Numérique/COLEPS)."
    )
    if edition_date:
        logger.info(f"Date de l'édition du PPM détectée : {edition_date.strftime('%d/%m/%Y')}.")
    else:
        logger.warning(
            "Date de l'édition du PPM introuvable en tête de la 1ère page -- "
            "la date d'upload sera utilisée comme repli pour le suivi des nouveautés."
        )
    return results, edition_date


# ---------------------------------------------------------------------------
# Business Intelligence -- agregations sur les resultats d'extraction
# ---------------------------------------------------------------------------

class PPMAnalytics:
    """Agrege les lignes retenues en indicateurs de pilotage : budget total
    identifie, repartition par autorite/type d'acheteur/categorie IT/source
    de financement. Ne filtre rien -- l'extraction reste exhaustive, la BI
    ne fait que segmenter pour la lecture du DG."""

    def __init__(self, results: list[dict]):
        self.df = pd.DataFrame(results)
        if not self.df.empty:
            self.df["_amount"] = self.df["Montant prévisionnel (FCFA)"].apply(parse_amount)
            self.df["_authority"] = self.df["Autorité (en-tête de page)"].apply(clean_authority_name)

    @property
    def is_empty(self):
        return self.df.empty

    def kpis(self):
        if self.df.empty:
            return {"lignes": 0, "autorites": 0, "budget_total": 0, "categories": 0}
        return {
            "lignes": len(self.df),
            "autorites": self.df["_authority"].nunique(),
            "budget_total": int(self.df["_amount"].dropna().sum()),
            "categories": self.df["Catégories détectées"].str.split(", ").explode().nunique(),
        }

    def authority_directory(self):
        cols = ["Autorité", "Type", "Lignes", "Budget total (FCFA)"]
        if self.df.empty:
            return pd.DataFrame(columns=cols)
        grouped = self.df.groupby(["_authority", "Type d'acheteur"]).agg(
            lignes=("_authority", "count"), budget=("_amount", "sum"),
        ).reset_index()
        grouped.columns = ["Autorité", "Type", "Lignes", "Budget total (FCFA)"]
        return grouped.sort_values("Budget total (FCFA)", ascending=False).reset_index(drop=True)

    def buyer_type_breakdown(self):
        cols = ["Type d'acheteur", "Lignes", "Budget total (FCFA)"]
        if self.df.empty:
            return pd.DataFrame(columns=cols)
        grouped = self.df.groupby("Type d'acheteur").agg(
            lignes=("_authority", "count"), budget=("_amount", "sum"),
        ).reset_index()
        grouped.columns = cols
        return grouped.sort_values("Budget total (FCFA)", ascending=False).reset_index(drop=True)

    def category_breakdown(self):
        """Une ligne peut appartenir a plusieurs categories (ex: dev logiciel
        + infra) -- chaque categorie concernee recoit le budget de la ligne,
        donc la somme des budgets par categorie peut depasser le budget
        total global. C'est le comportement attendu pour une segmentation
        multi-etiquettes, pas une erreur de calcul."""
        cols = ["Catégorie", "Lignes", "Budget total (FCFA)"]
        if self.df.empty:
            return pd.DataFrame(columns=cols)
        exploded = self.df.assign(
            _category=self.df["Catégories détectées"].str.split(", ")
        ).explode("_category")
        grouped = exploded.groupby("_category").agg(
            lignes=("_category", "count"), budget=("_amount", "sum"),
        ).reset_index()
        grouped.columns = cols
        return grouped.sort_values("Lignes", ascending=False).reset_index(drop=True)

    def funding_breakdown(self):
        cols = ["Source de financement", "Lignes", "Budget total (FCFA)"]
        if self.df.empty:
            return pd.DataFrame(columns=cols)
        source = self.df["Source de financement"].replace("", "Non précisée")
        grouped = self.df.assign(_source=source).groupby("_source").agg(
            lignes=("_source", "count"), budget=("_amount", "sum"),
        ).reset_index()
        grouped.columns = cols
        return grouped.sort_values("Budget total (FCFA)", ascending=False).reset_index(drop=True)


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

def results_to_dataframe(results: list[dict]) -> pd.DataFrame:
    if not results:
        return pd.DataFrame(columns=ALL_COLUMNS)
    return pd.DataFrame(results)[ALL_COLUMNS]


def export_results(results: list[dict], output_path: Path):
    if not results:
        logger.warning("Aucune ligne retenue -- aucun fichier ne sera généré.")
        return None
    df = results_to_dataframe(results)
    if output_path.suffix.lower() == ".csv":
        df.to_csv(output_path, index=False, encoding="utf-8-sig")
    else:
        df.to_excel(output_path, index=False, engine="openpyxl")
    logger.info(f"Export terminé : {len(df)} lignes écrites dans {output_path}")
    return df


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl")
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Option OCR (repli pour un futur PPM scanne -- non necessaire ici, verifie :
# c'est du texte natif genere par JasperReports).
# ---------------------------------------------------------------------------

def extract_pdf_with_ocr(pdf_path, dpi: int = 200):
    """Necessite en plus : `pip install pytesseract pdf2image` + Tesseract
    OCR installe sur le systeme + Poppler (pour pdf2image)."""
    import pytesseract
    from pdf2image import convert_from_path

    logger.warning("Mode OCR activé -- nettement plus lent qu'une extraction en texte natif.")
    results = []
    images = convert_from_path(str(pdf_path), dpi=dpi)
    for page_number, image in enumerate(images, start=1):
        text = pytesseract.image_to_string(image, lang="fra")
        categories, uses_generic_term = detect_categories(text)
        if categories:
            results.append({
                "Page PDF": page_number,
                "Catégories détectées": ", ".join(categories),
                "À relire (terme générique)": "Oui" if uses_generic_term else "",
                "Texte OCR (extrait)": text[:1000],
            })
        if page_number % 50 == 0:
            logger.info(f"[OCR] Page {page_number}/{len(images)} traitée.")
    return results


# ---------------------------------------------------------------------------
# Point d'entree CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Extrait du PPM les marchés publics liés au Numérique/IT/COLEPS."
    )
    parser.add_argument("pdf_path", type=Path, help="Chemin du PDF du Journal de Programmation des Marchés.")
    parser.add_argument(
        "-o", "--output", type=Path, default=Path("marches_it_coleps_filtres.xlsx"),
        help="Fichier de sortie (.xlsx ou .csv). Par défaut : marches_it_coleps_filtres.xlsx",
    )
    parser.add_argument("--progress-every", type=int, default=100)
    parser.add_argument("--ocr", action="store_true", help="Force le mode OCR (PDF scanné).")
    parser.add_argument(
        "--ongoing-only", action="store_true",
        help=f'Ecarte les marchés déjà attribués (colonne "{ATTRIBUTION_COLUMN}" déjà passée). '
             "Appliqué après la détection des catégories IT, jamais sur le contenu.",
    )
    args = parser.parse_args()

    if not args.pdf_path.exists():
        logger.error(f"Fichier introuvable : {args.pdf_path}")
        sys.exit(1)

    if args.ocr:
        results = extract_pdf_with_ocr(args.pdf_path)
    else:
        results, _edition_date = extract_pdf(args.pdf_path, progress_every=args.progress_every)
    if args.ongoing_only:
        before = len(results)
        results = filter_ongoing(results)
        logger.info(f"--ongoing-only : {before - len(results)} marché(s) déjà attribué(s) écarté(s) ({len(results)} restants).")
    export_results(results, args.output)

    broad_hits = sum(1 for r in results if r.get("À relire (terme générique)") == "Oui")
    if broad_hits:
        logger.warning(
            f"{broad_hits} ligne(s) retenue(s) au moins en partie via un terme ambigu "
            f"('étude', 'audit', 'AMO', 'MOE', 'cahier des charges', 'application', "
            f"'plateforme', 'hébergement') -- marquées 'Oui' dans la colonne "
            f"« À relire (terme générique) », à vérifier manuellement."
        )


if __name__ == "__main__":
    main()
