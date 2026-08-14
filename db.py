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
from datetime import datetime, timezone

import streamlit as st
from pymongo import MongoClient
from pymongo.errors import PyMongoError

from ppm_extraction import ATTRIBUTION_COLUMN, clean_authority_name

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


def classify_and_store(collection, results: list[dict]) -> list[dict]:
    """Pour chaque ligne extraite, détermine si le marché est Nouveau,
    Modifié (budget/dates changés depuis la dernière édition vue) ou Déjà vu,
    puis met à jour la base en conséquence. `collection` a l'interface
    pymongo standard (find_one/update_one) -- injectable pour les tests."""
    now = datetime.now(timezone.utc).isoformat()
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

        collection.update_one(
            {"_id": key},
            {
                "$set": {"snapshot": snapshot, "record": record, "last_seen": now},
                "$setOnInsert": {"first_seen": now},
            },
            upsert=True,
        )
        tagged.append({**record, "Statut": status})
    return tagged


@st.cache_resource(show_spinner=False)
def get_collection():
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
    return client["daredab_ppm"]["markets"]
