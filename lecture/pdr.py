"""Lecture du PDR (une ligne par créa).

Deux présentations sont gérées :
- le PDR global (onglet DISPLAY_A79, ligne d'en-tête) : colonnes lues par leur nom ;
- un extrait sans en-tête : colonnes lues par position (ancienne présentation).
"""
from __future__ import annotations

import openpyxl

from tables import normaliser

ONGLET = "DISPLAY_A79"
PAR_NOM = {"pilier": ("pilier",), "objectif": ("objectif",), "campagne": ("campagne",),
           "message": ("produit/info", "produit"), "regie": ("plateforme",), "format": ("format",),
           "origine": ("organique",), "landing": ("redirection",), "lien_crea": ("lien vers la crea",)}
PAR_POSITION = {"pilier": 4, "objectif": 5, "campagne": 6, "message": 10, "regie": 11,
                "format": 12, "origine": 15, "landing": 16}


def lire_pdr(fichier) -> list[dict]:
    wb = openpyxl.load_workbook(fichier, read_only=True, data_only=True)
    ws = wb[ONGLET] if ONGLET in wb.sheetnames else wb.worksheets[0]
    lignes = list(ws.iter_rows(values_only=True))
    if not lignes:
        return []

    entete = [normaliser(v) for v in lignes[0]]
    if "campagne" in entete:
        cols = {}
        for cle, alias in PAR_NOM.items():
            for j, nom in enumerate(entete):
                if nom and any(nom == a or nom.startswith(a) for a in alias):
                    cols[cle] = j
                    break
        donnees = lignes[1:]
    else:
        cols = {k: c - 1 for k, c in PAR_POSITION.items()}
        donnees = lignes

    res = []
    for row in donnees:
        ligne = {k: (row[j] if j < len(row) else None) for k, j in cols.items()}
        if ligne.get("campagne"):
            res.append({k: (str(v).strip() if v is not None else "") for k, v in ligne.items()})
    return res


def campagnes(lignes: list[dict]) -> list[str]:
    vues = []
    for l in lignes:
        if l["campagne"] not in vues:
            vues.append(l["campagne"])
    return vues


def filtrer(lignes: list[dict], campagne: str) -> list[dict]:
    return [l for l in lignes if l["campagne"] == campagne]
