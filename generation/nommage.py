"""Règles de nommage.

Chaque champ est nettoyé AVANT d'être assemblé : le « _ » est réservé au séparateur,
et à l'intérieur d'un champ tout caractère qui n'est ni lettre, ni chiffre, ni tiret
devient un tiret. Le nettoyage des UTM de l'export CM ne modifie alors plus rien.
"""
from __future__ import annotations

import re
import unicodedata


def slug(valeur) -> str:
    v = unicodedata.normalize("NFKD", str(valeur or ""))
    v = "".join(c for c in v if not unicodedata.combining(c))
    v = re.sub(r"[^A-Za-z0-9-]+", "-", v)
    return re.sub(r"-+", "-", v).strip("-")


def nom_campagne_libre(nom_pdr: str) -> str:
    """MDC_Sdg_V2 -> MDC-SDG-V2"""
    return slug(str(nom_pdr).replace("_", "-")).upper()
