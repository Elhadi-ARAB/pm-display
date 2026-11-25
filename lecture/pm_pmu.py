"""Lecture d'un PM PMU au format actuel (onglet « DIGITAL », présenté comme un devis).

Le PM a un ou plusieurs tableaux (Branding, Considération…), chacun avec une ligne d'en-tête
contenant « Régies ». La section est lue dans la 2e colonne de cette ligne d'en-tête.
"""
from __future__ import annotations

import re
from datetime import datetime

import openpyxl

from tables import normaliser

ENTETES = {"levier": ("levier",), "regie": ("regies", "regie"), "objectif": ("objectif",),
           "emplacement": ("emplacement",), "ciblage": ("ciblage",), "device": ("device",),
           "achat": ("mode d'achat", "mode d achat"), "formats": ("formats", "format"), "creative": ("creative",),
           "debut": ("date de debut",), "fin": ("date de fin",), "volume": ("volume",), "cpm": ("cpm",),
           "budget": ("net media",), "commentaire": ("commentaire",)}


def _colonnes(row) -> dict:
    cols = {}
    for c in row:
        n = normaliser(c.value)
        for k, alias in ENTETES.items():
            if k not in cols and n and any(n == a or n.startswith(a) for a in alias):
                cols[k] = c.column
    return cols


def _fusions(ws) -> dict:
    res = {}
    for plage in ws.merged_cells.ranges:
        v = ws.cell(row=plage.min_row, column=plage.min_col).value
        for r in range(plage.min_row, plage.max_row + 1):
            for c in range(plage.min_col, plage.max_col + 1):
                res[(r, c)] = v
    return res


def lire_pm_pmu(fichier) -> dict:
    wb = openpyxl.load_workbook(fichier, data_only=True)
    ws = wb["DIGITAL"] if "DIGITAL" in wb.sheetnames else wb.worksheets[0]
    fusions = _fusions(ws)
    val = lambda r, c: fusions.get((r, c), ws.cell(row=r, column=c).value)
    lignes, cols, section, levier = [], None, None, None
    for r in range(1, ws.max_row + 1):
        row = [ws.cell(row=r, column=c) for c in range(1, ws.max_column + 1)]
        if any(normaliser(c.value) in ("regies", "regie") for c in row):
            cols = _colonnes(row)
            section = next((str(c.value).strip().capitalize() for c in row[:3] if c.value and normaliser(c.value) not in ("levier",)
                            and normaliser(c.value) not in ("hippique",)), None)
            levier = None
            continue
        if not cols:
            continue
        regie = val(r, cols["regie"])
        if not regie or not str(regie).strip():
            continue                                  # sous-total, ligne vide
        if cols.get("levier") and val(r, cols["levier"]):
            levier = str(val(r, cols["levier"])).strip()
        get = lambda k: val(r, cols[k]) if k in cols else None
        lignes.append({
            "section": section, "levier": levier, "regie": str(regie).strip(), "objectif": get("objectif"),
            "ciblage": re.sub(r"\s+", " ", str(get("ciblage") or "")).strip(), "device": get("device"),
            "emplacement": get("emplacement"), "commentaire": get("commentaire"),
            "achat": get("achat"), "formats": re.sub(r"\s+", " ", str(get("formats") or "")).strip(),
            "debut": get("debut") if isinstance(get("debut"), datetime) else None,
            "fin": get("fin") if isinstance(get("fin"), datetime) else None,
            "budget": get("budget") if isinstance(get("budget"), (int, float)) else None,
            "cpm": get("cpm") if isinstance(get("cpm"), (int, float)) else None,
        })
    return {"lignes": lignes, "sections": [s for s in dict.fromkeys(l["section"] for l in lignes) if s]}


def lire_brief_pmu(fichier) -> list[dict]:
    """Brief créa : une ligne par régie × format × taille."""
    wb = openpyxl.load_workbook(fichier, data_only=True)
    ws = wb.worksheets[0]
    res, cols, regie = [], None, None
    for r in range(1, ws.max_row + 1):
        row = [ws.cell(row=r, column=c) for c in range(1, ws.max_column + 1)]
        noms = [normaliser(c.value) for c in row]
        if "regies" in noms and any(n.startswith("nom du format") for n in noms):
            cols = {n: c.column for n, c in zip(noms, row) if n}
            continue
        if not cols:
            continue
        g = lambda prefixe: next((ws.cell(row=r, column=c).value for n, c in cols.items() if n.startswith(prefixe)), None)
        if g("regies"):
            regie = str(g("regies")).strip()
        taille = g("taille du format") or g("dimensions")
        if regie and taille and g("nom du format"):
            res.append({"regie": regie, "format": str(g("nom du format")).strip(),
                        "taille": re.sub(r"\s*px\s*$", "", str(taille).strip(), flags=re.I),
                        "tracking": g("hebergement")})
    return res
