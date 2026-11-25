"""PM client Carrefour : même mise en forme que le PM Carrefour actuel (onglets GLOBAL et PM DISPLAY).

On part du PM Carrefour d'origine (pm_builder/modeles/PM_CRF_modele.xlsx) :
- PM DISPLAY : l'en-tête est gardé et mis à jour, le tableau est réécrit à partir du PM commun, avec les mêmes
  colonnes, styles et formules (fees DV360 + taxe Google, impressions, vues 100 %, CPVV, reach) et un total juste ;
- GLOBAL : infos de campagne, budget display relié au total, tableau et camembert par plateforme, cases à cocher
  display (audiences, formats, KPI). La partie Social et l'argumentaire sont laissés à compléter.
- L'onglet PM SOCIAL (propre à la campagne d'origine) est retiré : l'app ne gère que le display.
"""
from __future__ import annotations

import io
from copy import copy
from datetime import date, datetime
from pathlib import Path

import openpyxl
import pandas as pd

import tables as T
from pm_builder import frais

MODELE = Path(__file__).parent / "modeles" / "PM_CRF_modele.xlsx"
R_ENTETE, R_LIGNE, R_TOTAL = 12, 15, 21      # lignes de référence de style dans le modèle
NB_COL = 22                                    # A -> V
VIDEO = {"Non-skippable", "In-Stream-bumper", "Shorts", "In-Stream-Classic", "In-Stream-Trueview", "In-Stream-catch-up",
         "In-Stream-enrichi", "Midroll", "Monopreroll", "Pre-Roll", "Hack6", "Connected-TV", "QR-codes-experience",
         "VRC", "VRC-Non-Skip", "Out-Stream-inread", "Custom3D-CTV", "Video-Dynamic-Creative"}


def _v(x):
    return None if x is None or (isinstance(x, float) and pd.isna(x)) or str(x).strip() == "" else x


def _date(x):
    x = _v(x)
    if x is None:
        return None
    if isinstance(x, pd.Timestamp):
        return x.to_pydatetime()
    if isinstance(x, date) and not isinstance(x, datetime):
        return datetime(x.year, x.month, x.day)
    return x


def _style(ws, r):
    return {c: copy(ws.cell(row=r, column=c)._style) for c in range(1, NB_COL + 1)}


def texte_audiences(ligne, audiences: pd.DataFrame) -> str | None:
    """Ciblage écrit pour le client s'il est rempli, sinon les audiences du groupe : « Familles (30%) »."""
    c = _v(ligne.get("Ciblage"))
    if c:
        return c
    g = str(_v(ligne.get("Groupe d'audiences")) or "")
    grp = audiences[audiences["Groupe d'audiences"].astype(str) == g] if g else audiences.iloc[0:0]
    lignes = []
    for _, a in grp.iterrows():
        p = pd.to_numeric(a.get("% du groupe"), errors="coerce")
        lignes.append(f"{a['Audience']} ({p:g}%)" if pd.notna(p) else str(a["Audience"]))
    return "\n".join(lignes) or None


def exporter(camp: dict, lignes: pd.DataFrame, audiences: pd.DataFrame) -> bytes:
    wb = openpyxl.load_workbook(MODELE)
    if "PM SOCIAL" in wb.sheetnames:
        del wb["PM SOCIAL"]
    ws = wb["PM DISPLAY"]
    st_ligne, st_total = _style(ws, R_LIGNE), _style(ws, R_TOTAL)
    h_ligne = ws.row_dimensions[R_LIGNE].height

    # ---------------------------------------------------------------- tableau : on repart de zéro sous l'en-tête
    for m in [m for m in ws.merged_cells.ranges if m.min_row > R_ENTETE]:
        ws.unmerge_cells(str(m))
    ws._images = [im for im in ws._images if im.anchor._from.row + 1 <= R_ENTETE]   # visuels d'exemple de l'original
    ws.delete_rows(R_ENTETE + 1, ws.max_row - R_ENTETE)
    for r in [r for r in list(ws.row_dimensions.keys()) if r > R_ENTETE]:
        del ws.row_dimensions[r]
    ws.data_validations.dataValidation = []

    df = lignes.reset_index(drop=True)
    r0 = R_ENTETE + 1
    for i, l in df.iterrows():
        r = r0 + i
        for c, st_ in st_ligne.items():
            ws.cell(row=r, column=c)._style = copy(st_)
        ws.row_dimensions[r].height = h_ligne
        cle = T.trouver_regie(l.get("Régie"))
        plateforme = T.plateforme(cle, l.get("Type d'achat"))
        taux = frais.taux_fees_crf(l.get("Type d'achat"), l.get("Régie"))
        vals = {2: _date(l.get("Début")), 3: _date(l.get("Fin")), 4: "DV360" if plateforme == "DV360" else "Régie",
                5: str(_v(l.get("Régie")) or "").replace("-", " ").upper() or None, 6: _v(l.get("Type d'achat")),
                7: _v(l.get("Format")), 8: _v(l.get("Specs")), 9: _v(l.get("Device")), 10: texte_audiences(l, audiences),
                11: _v(l.get("KPI")), 14: _v(l.get("Budget net (€)")), 16: _v(l.get("CPM (€)")), 18: _v(l.get("VTR / LTR")),
                20: _v(l.get("Fréquence")), 22: _v(l.get("Études / BLS"))}
        for c, v in vals.items():
            ws.cell(row=r, column=c, value=v)
        for c in (2, 3):
            ws.cell(row=r, column=c).number_format = "DD/MM/YYYY"
        ws.cell(row=r, column=12, value=f"=M{r}+N{r}")
        ws.cell(row=r, column=13, value=f"=({taux}*N{r})/100")
        ws.cell(row=r, column=15, value=f"=N{r}/P{r}*1000")
        ws.cell(row=r, column=17, value=f'=IF(R{r}="","",O{r}*R{r})')
        ws.cell(row=r, column=19, value=f'=IF(OR(Q{r}="",Q{r}=0),"",N{r}/Q{r})')
        ws.cell(row=r, column=21, value=f'=IF(T{r}="","",O{r}/T{r})')
    rl = r0 + len(df) - 1
    rt = rl + 1
    for c, st_ in st_total.items():
        ws.cell(row=rt, column=c)._style = copy(st_)
    for col in "LMNOQ":
        ws[f"{col}{rt}"] = f"=SUM({col}{r0}:{col}{rl})"
    ws[f"P{rt}"] = f"=(N{rt}/O{rt})*1000"
    # régies et audiences identiques sur des lignes qui se suivent : cellules fusionnées, comme dans l'original
    df["_dsp"] = ["DV360" if T.plateforme(T.trouver_regie(r), a) == "DV360" else "Régie" for r, a in zip(df["Régie"], df["Type d'achat"])]
    for col, cle in ((2, "Début"), (3, "Fin"), (4, "_dsp"), (5, "Régie"), (10, "Groupe d'audiences")):
        debut = 0
        for i in range(1, len(df) + 1):
            fin_bloc = i == len(df) or str(df.at[i, cle]) != str(df.at[debut, cle]) or not _v(df.at[debut, cle])
            if fin_bloc:
                if i - 1 > debut:
                    ws.merge_cells(start_row=r0 + debut, start_column=col, end_row=r0 + i - 1, end_column=col)
                debut = i

    # ---------------------------------------------------------------- en-tête de PM DISPLAY
    d0, d1 = _date(camp.get("debut")), _date(camp.get("fin"))
    dates = f"{d0:%d/%m/%Y} au {d1:%d/%m/%Y}" if d0 and d1 else None
    duree = f"{(d1 - d0).days + 1} jours" if d0 and d1 else None
    libelle = _v(camp.get("libelle")) or str(camp.get("nom") or "").replace("_", " ").upper()
    ws["C3"], ws["C5"], ws["C6"] = libelle, dates, duree
    ws["F3"] = str(camp.get("objectif") or "").upper() or None
    ws["G3"], ws["H3"] = f"=L{rt}", "=G3/$G$5"

    # ---------------------------------------------------------------- GLOBAL
    g = wb["GLOBAL"]
    g["C4"], g["C6"], g["C8"] = libelle, dates, _v(camp.get("owner_crf"))
    if d0 and d1:                                  # nombre affiché « 14 jours » : la formule budget / jour fonctionne
        g["C7"] = (d1 - d0).days + 1
        g["C7"].number_format = '0 "jours"'
    else:
        g["C7"] = None
    g["C14"] = f"='PM DISPLAY'!L{rt}"
    for r in (106, 107, 108):                      # budgets social : à compléter par l'équipe Social
        g[f"C{r}"] = None
    g["C15"] = "=SUM(C106:C108)"
    regies = list(dict.fromkeys(str(x) for x in df["Régie"] if _v(x)))
    for r in range(109, 109 + max(len(regies), 5) + 5):
        g[f"B{r}"], g[f"C{r}"] = None, None
    for k, reg in enumerate(regies):
        cellules = "+".join(f"'PM DISPLAY'!L{r0 + i}" for i in df.index if str(df.at[i, "Régie"]) == reg)
        g[f"B{109 + k}"], g[f"C{109 + k}"] = reg.replace("-", " "), f"={cellules}"
    fin = 108 + len(regies)
    for ch in g._charts:
        for s in ch.series:
            s.val.numRef.f = f"GLOBAL!$C$106:$C${fin}"
            if s.cat is not None and s.cat.strRef is not None:
                s.cat.strRef.f = f"GLOBAL!$B$106:$B${fin}"
    g["B26"] = "Argumentaire à compléter."
    g["B28"] = "Display : à compléter."
    g["B38"] = "Social : à compléter."
    td = set(audiences["Targeting details"].dropna().astype(str))
    ft = set(df["Type de format"].dropna().astype(str))
    for r, cles in ((57, {"Broad", "Socio-demo"}), (58, {"Familles", "Affinity-Familles", "1stFamilles", "Familles/Seniors/Mots-clés"}),
                    (59, {"Promovores"}), (60, {"Seniors", "Seniors & CSP+"}), (61, {"Aisés / CSP+", "Gourmets", "Contextual"}),
                    (62, {"Clients Magasins"}), (63, {"RTG"}), (64, {"LAL"})):
        g[f"C{r}"] = bool(td & cles)
    for r in range(58, 64):                         # formats social : à cocher par l'équipe Social
        g[f"F{r}"] = False
    g["H58"] = bool(ft & VIDEO)
    g["H59"] = "In-Stream-bumper" in ft
    g["H60"] = any(f.startswith("Display") for f in ft) or bool(ft & {"Incontent-banner", "Display", "Engage", "Cover", "HPTO", "In-feed"})
    g["H61"] = bool(ft & {"Stream-Audio", "Podcast-Audio"})
    g["H62"] = "Exclu" in td
    g["H63"] = "In-Stream-enrichi" in ft
    g["C66"] = True
    g["C67"] = bool(ft & VIDEO)
    for r in range(68, 73):
        g[f"C{r}"] = False
    for r in range(59, 64):                         # objectifs vs réalisé : propres à la campagne d'origine
        for col in "NOP":
            g[f"{col}{r}"] = None
    g["C75"], g["C76"] = "=IF(C16=0,0,C15/C16)", "=IF(C16=0,0,C14/C16)"
    g["G75"], g["H75"] = str(camp.get("objectif") or "Branding").capitalize(), 1.0

    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()
