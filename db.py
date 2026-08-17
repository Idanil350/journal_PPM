"""Mémoire persistante des marchés IT déjà vus, à travers plusieurs éditions
du PPM (août, septembre, ...) -- pour qu'un nouvel upload ne fasse ressortir
que ce qui est vraiment nouveau ou modifié, pas relire 200 lignes déjà vues.

Le PPM n'a pas d'identifiant stable entre éditions (la colonne "N°" n'est que
la position dans le tableau de CETTE édition -- elle bouge d'un mois à
l'autre si des lignes sont ajoutées/retirées avant). L'identité d'un marché
est donc reconstruite à partir de son autorité + désignation normalisées.
"""

import hashlib
import os
import re
import unicodedata
from collections import Counter
from datetime import date, datetime, timezone

import streamlit as st
from pymongo import MongoClient
from pymongo.errors import PyMongoError

from ppm_extraction import ATTRIBUTION_COLUMN, clean_authority_name, filter_ongoing, parse_ppm_date

LAUNCH_COLUMN = "Lancement de consultation / Invitation à soumissionner"

STATUS_NEW = "Nouveau"
STATUS_UPDATED = "Modifié"
STATUS_SEEN = "Déjà vu"

_SNAPSHOT_FIELDS = [
    "Montant prévisionnel (FCFA)",
    ATTRIBUTION_COLUMN,
    "Lancement de consultation / Invitation à soumissionner",
]


def _normalize(text: str) -> str:
    """Insensible aux accents/casse/espaces multiples, pour que deux éditions
    du PPM avec une reformulation mineure du même projet ne soient pas
    traitées comme deux marchés différents."""
    if not text:
        return ""
    folded = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", folded).strip().lower()


def market_key(record: dict) -> str:
    """Identité stable d'un marché à travers les éditions : hash de
    l'autorité + la désignation du projet, normalisées."""
    authority = _normalize(clean_authority_name(record.get("Autorité (en-tête de page)", "")))
    designation = _normalize(record.get("Désignation et localisation du projet", ""))
    raw = f"{authority}|{designation}"
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()


def _snapshot(record: dict) -> dict:
    """Champs dont le changement doit déclencher le statut "Modifié"."""
    return {field: record.get(field) for field in _SNAPSHOT_FIELDS}


def classify_and_store(collection, results: list[dict], edition_date=None) -> list[dict]:
    """Pour chaque ligne extraite, détermine si le marché est Nouveau,
    Modifié (budget/dates changés depuis la dernière édition vue) ou Déjà vu,
    puis met à jour la base en conséquence. `collection` a l'interface
    pymongo standard (find_one/update_one) -- injectable pour les tests.

    `edition_date` est la date "MISE A JOUR DU ..." imprimée sur le PPM
    (extraite par ppm_extraction.extract_edition_date) -- utilisée comme
    référence pour first_seen/last_seen à la place de l'heure d'upload, pour
    qu'un retard à uploader ne fasse pas passer un projet publié le 10 pour
    "nouveau depuis le 20". Repli sur l'heure actuelle si introuvable.

    `last_changed` n'est mis à jour QUE quand le marché est Nouveau ou
    Modifié -- pas quand il est simplement retrouvé inchangé -- pour que
    get_changes_since() puisse répondre à "qu'est-ce qui a vraiment changé",
    sans les marchés "déjà vus" qui n'apportent aucune information neuve.
    Il utilise l'heure réelle (pas `edition_date`) : c'est ce qui doit rester
    sur la même horloge que mark_consultation()/get_last_consultation() --
    sinon un PPM daté dans le passé (uploadé en retard) ne dépasserait
    jamais un repère de consultation basé sur l'heure actuelle, et les
    changements resteraient invisibles pour toujours."""
    reference = edition_date.isoformat() if edition_date else datetime.now(timezone.utc).isoformat()
    processed_at = datetime.now(timezone.utc).isoformat()
    tagged = []
    for record in results:
        key = market_key(record)
        snapshot = _snapshot(record)
        existing = collection.find_one({"_id": key})

        if existing is None:
            status = STATUS_NEW
        elif existing.get("snapshot") != snapshot:
            status = STATUS_UPDATED
        else:
            status = STATUS_SEEN

        set_fields = {"snapshot": snapshot, "record": record, "last_seen": reference}
        if status in (STATUS_NEW, STATUS_UPDATED):
            set_fields["last_changed"] = processed_at

        collection.update_one(
            {"_id": key},
            {"$set": set_fields, "$setOnInsert": {"first_seen": reference}},
            upsert=True,
        )
        tagged.append({**record, "Statut": status})
    return tagged


def get_markets_since(collection, since_date) -> list[dict]:
    """Tous les marchés détectés comme nouveaux depuis `since_date` (basé sur
    first_seen) -- lit directement la mémoire persistante, sans dépendre
    d'une extraction lancée dans la session en cours. Répond directement à
    "quels sont les nouveaux marchés depuis le [date] ?" sans avoir à
    ré-uploader/ré-analyser un journal."""
    if collection is None:
        return []
    since_key = since_date.isoformat()
    docs = collection.find({"first_seen": {"$gte": since_key}}).sort("first_seen", -1)
    results = []
    for doc in docs:
        record = dict(doc.get("record") or {})
        record["_first_seen"] = doc.get("first_seen")
        results.append(record)
    return results


def get_changes_since(collection, since_key) -> list[dict]:
    """Marchés ajoutés OU modifiés strictement après `since_key` (chaîne ISO
    comparable à `last_changed`) -- exclut les marchés simplement "déjà vus"
    inchangés. C'est la définition exacte donnée par le DG : "les nouvelles
    informations, c'est ce qui a changé depuis la dernière consultation".

    Borne strictement supérieure (pas >=) : `since_key` vient typiquement de
    mark_consultation(), pris juste après avoir lu ce résultat -- une
    comparaison inclusive re-signalerait comme "changé" un marché dont
    l'horodatage tombe exactement sur ce repère."""
    if collection is None:
        return []
    docs = collection.find({"last_changed": {"$gt": since_key}}).sort("last_changed", -1)
    results = []
    for doc in docs:
        record = dict(doc.get("record") or {})
        record["_last_changed"] = doc.get("last_changed")
        results.append(record)
    return results


_CONSULTATION_ID = "last_consultation"


def get_last_consultation(meta_collection):
    """Horodatage (chaîne ISO) de la dernière fois que quelqu'un a consulté
    les nouveautés -- un seul repère partagé pour toute l'équipe DAREDAB
    (le DG a précisé que ce n'est pas personnalisé par employé). None si
    jamais consulté."""
    if meta_collection is None:
        return None
    doc = meta_collection.find_one({"_id": _CONSULTATION_ID})
    return doc.get("timestamp") if doc else None


def mark_consultation(meta_collection) -> None:
    """Enregistre "maintenant" comme nouvelle dernière consultation. À
    appeler APRÈS avoir lu/affiché les changements depuis l'ancien repère --
    jamais avant, sinon on écrase la référence avant de s'en être servi."""
    if meta_collection is None:
        return
    meta_collection.update_one(
        {"_id": _CONSULTATION_ID},
        {"$set": {"timestamp": datetime.now(timezone.utc).isoformat()}},
        upsert=True,
    )


def get_markets_by_launch_date(collection, since_date, date_column: str = LAUNCH_COLUMN) -> list[dict]:
    """Tous les marchés (dans toute la mémoire enregistrée) dont la colonne
    "Lancement de consultation / Invitation à soumissionner" est prévue à
    partir de `since_date` -- critère demandé explicitement par le DG :
    "nouveau depuis une date" au sens où le lancement de la consultation
    a commencé à partir de cette date, pas au sens de notre propre suivi
    Nouveau/Modifié/Déjà vu (qui reste disponible séparément, sans lien
    avec cette fonction). Ne dépend pas d'une extraction en session --
    relit tout ce qui est déjà en mémoire."""
    if collection is None:
        return []
    results = []
    for doc in collection.find({}):
        record = dict(doc.get("record") or {})
        launch = parse_ppm_date(record.get(date_column))
        if launch is not None and launch >= since_date:
            results.append(record)
    results.sort(key=lambda r: parse_ppm_date(r.get(date_column)) or since_date)
    return results


# ---------------------------------------------------------------------------
# Score d'attractivité commerciale -- pas du machine learning : une
# heuristique explicable, fondée sur des écarts mesurés en direct sur les
# vraies données DAREDAB (ticket moyen bailleur ~330M FCFA vs BIP ~57M ;
# développement logiciel ~203M/dossier vs infrastructure ~78M). DAREDAB est
# encore jeune sur ce secteur (~5 ans) -- ce score sert à observer où se
# positionner, pas à prétendre prédire un résultat.
# ---------------------------------------------------------------------------

FUNDING_TIER = {
    "BMI-IDA": 3, "BAD": 3, "FM": 3, "BID": 3, "AFD": 3,
    "Fonds Propres des CTD": 2, "Autres": 2, "etc": 2,
    "BIP": 1, "FEICOM": 1,
}
CATEGORY_TIER = {
    "Développement logiciel": 3,
    "COLEPS": 2,
    "Innovation / IA / dématérialisation": 2,
    "Infrastructure IT": 1,
    "Conseil / Études IT (termes génériques -- à relire)": 1,
    "Formation digitale": 1,
}
RECURRENT_THRESHOLD = 3  # nombre de marchés déjà vus pour qu'un acheteur compte comme récurrent

ACTION_CONTACT = "Contacter maintenant"
ACTION_PREPARE = "Préparer le dossier"
ACTION_WATCH = "Surveiller seulement"


def _funding_tier(source) -> int:
    """Accepte n'importe quel type en entrée (pas seulement str) -- une
    valeur non textuelle dans un vieux document Mongo ne doit pas planter
    tout le panneau de priorités, juste retomber sur le palier par défaut."""
    return FUNDING_TIER.get(str(source or "").strip(), 1)


def _category_tier(categories_str) -> int:
    cats = [c.strip() for c in str(categories_str or "").split(",") if c.strip()]
    return max((CATEGORY_TIER.get(c, 1) for c in cats), default=1)


def _urgency_tier(launch_date_value, today=None) -> int:
    """Plus le lancement de consultation est proche, plus c'est urgent --
    un score élevé sur un marché qui ne se lance que dans 6 mois n'appelle
    pas la même réaction qu'un score élevé sur un marché à 5 jours."""
    launch = parse_ppm_date(launch_date_value)
    if launch is None:
        return 1
    days = (launch - (today or date.today())).days
    if days <= 7:
        return 3
    if days <= 30:
        return 2
    return 1


def buyer_market_counts(all_records: list[dict]) -> Counter:
    """Combien de marchés déjà vus par acheteur (nom normalisé) -- sert à
    repérer les acheteurs récurrents, cibles de relation long terme."""
    counts = Counter()
    for record in all_records:
        authority = clean_authority_name(record.get("Autorité (en-tête de page)", ""))
        counts[authority] += 1
    return counts


def compute_attractiveness(record: dict, buyer_counts: Counter) -> dict:
    """Score = financement × catégorie × urgence, avec un bonus si
    l'acheteur est récurrent. Renvoie aussi une action NOMMÉE, pas un chiffre
    à interpréter -- un score seul laisse toujours la place à "je déciderai
    plus tard"."""
    funding = _funding_tier(record.get("Source de financement"))
    category = _category_tier(record.get("Catégories détectées"))
    urgency = _urgency_tier(record.get(LAUNCH_COLUMN))
    authority = clean_authority_name(record.get("Autorité (en-tête de page)", ""))
    recurrent = buyer_counts.get(authority, 0) >= RECURRENT_THRESHOLD

    score = funding * category * urgency + (2 if recurrent else 0)

    if score >= 15:
        action = ACTION_CONTACT
    elif score >= 8:
        action = ACTION_PREPARE
    else:
        action = ACTION_WATCH

    return {
        "_score": score,
        "_action": action,
        "_urgent": urgency == 3,
        "_recurrent": recurrent,
    }


def get_priority_opportunities(collection, limit: int = 10) -> list[dict]:
    """Marchés encore ouverts (pas déjà attribués), classés par score
    d'attractivité décroissant -- le panneau de décision principal de
    l'onglet Business Intelligence. Fonctionne sur toute la mémoire, pas
    seulement la dernière extraction en session."""
    if collection is None:
        return []
    all_docs = list(collection.find({}))
    all_records = [dict(d.get("record") or {}) for d in all_docs if d.get("record")]
    counts = buyer_market_counts(all_records)

    scored = []
    for doc in all_docs:
        record = dict(doc.get("record") or {})
        if not record:
            continue
        try:
            attractiveness = compute_attractiveness(record, counts)
        except Exception:
            # Un document ancien/inattendu ne doit pas faire disparaître le
            # panneau entier -- on l'ignore, silencieusement plutôt qu'un
            # écran d'erreur pour tout le monde.
            continue
        enriched = {
            **record,
            **attractiveness,
            "_decision": doc.get("decision"),
            "_market_key": doc.get("_id"),
        }
        scored.append(enriched)

    scored = filter_ongoing(scored)
    scored.sort(key=lambda r: r["_score"], reverse=True)
    return scored[:limit]


def set_decision(collection, market_key_value: str, decision: str) -> None:
    """Enregistre la décision de DAREDAB (go / no-go) sur un marché --
    trace explicite de l'action prise, pas seulement du score affiché. Sert
    aussi de graine de donnée de résultat pour un futur entraînement une
    fois qu'on saura aussi si le marché a été gagné ou perdu."""
    if collection is None or not market_key_value:
        return
    collection.update_one(
        {"_id": market_key_value},
        {"$set": {"decision": decision, "decision_at": datetime.now(timezone.utc).isoformat()}},
    )


@st.cache_resource(show_spinner=False)
def _get_client():
    """Connexion Mongo paresseuse et mise en cache pour la durée de vie du
    serveur (évite de reconnecter/ping à chaque rerun Streamlit) -- None si
    MONGO_URI n'est pas configuré (l'app doit continuer à fonctionner sans
    mémoire plutôt que planter)."""
    try:
        uri = st.secrets.get("MONGO_URI")
    except Exception:
        uri = None
    uri = uri or os.environ.get("MONGO_URI")
    if not uri:
        return None

    try:
        client = MongoClient(uri, serverSelectionTimeoutMS=5000)
        client.admin.command("ping")
    except PyMongoError:
        return None
    return client


def get_collection():
    """Collection des marchés suivis."""
    client = _get_client()
    return client["daredab_ppm"]["markets"] if client is not None else None


def get_meta_collection():
    """Collection séparée pour des repères globaux (ex: dernière
    consultation) -- distincte de `markets` pour ne pas polluer les requêtes
    qui parcourent tous les marchés (get_markets_by_launch_date etc.)."""
    client = _get_client()
    return client["daredab_ppm"]["meta"] if client is not None else None
