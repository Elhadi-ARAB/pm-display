"""PM client PMU : même mise en forme que le PM PMU actuel (devis, logos, blocs par section, frais).

On part du PM PMU d'origine (pm_builder/modeles/PM_PMU_modele.xlsx) : l'en-tête (logos, adresse, devis) est gardé,
puis les sections sont réécrites à partir du PM saisi dans l'app, avec les mêmes colonnes, styles et formules :
Volume = Net media / CPM × 1000, Fees Meta 3 %, sous-totaux par section, DAR par section, puis le récapitulatif
Net media / Total frais (Meta, DV, XPLN, DAR) / Total budget.
"""
from __future__ import annotations

import io
from copy import copy
from datetime import date, datetime
from pathlib import Path

import openpyxl
import pandas as pd

from pm_builder import frais

MODELE = Path(__file__).parent / "modeles" / "PM_PMU_modele.xlsx"
ONGLET = "DIGITAL"
# lignes du modèle d'origine qui servent de référence de style
R_TITRE, R_FEES, R_ENTETE, R_LIGNE, R_TOTAL, R_DAR, R_RECAP, R_RECAP_TOTAL = 16, 17, 18, 19, 22, 24, 41, 45
NB_COL = 22  # A -> V
TITRE_SECTION = {"Branding": "BRANDING", "Considération": "Consideration"}


def _style(ws, r):
    return {c: (copy(ws.cell(row=r, column=c)._style)) for c in range(1, NB_COL + 1)}, ws.row_dimensions[r].height


def _appliquer(ws, r, style):
    styles, h = style
    for c, st_ in styles.items():
        ws.cell(row=r, column=c)._style = copy(st_)
    if h:
        ws.row_dimensions[r].height = h


def _v(x):
    return None if x is None or (isinstance(x, float) and pd.isna(x)) or str(x).strip() == "" else x


def _date(x):
    if _v(x) is None:
        return None
    if isinstance(x, pd.Timestamp):
        return x.to_pydatetime()
    if isinstance(x, date) and not isinstance(x, datetime):
        return datetime(x.year, x.month, x.day)
    return x


def exporter(camp: dict, lignes: pd.DataFrame) -> bytes:
    wb = openpyxl.load_workbook(MODELE)
    ws = wb[ONGLET]
    S = {k: _style(ws, r) for k, r in (("titre", R_TITRE), ("fees", R_FEES), ("entete", R_ENTETE), ("ligne", R_LIGNE),
                                       ("total", R_TOTAL), ("dar", R_DAR), ("recap", R_RECAP), ("recap_total", R_RECAP_TOTAL))}
    entete_valeurs = {c: ws.cell(row=R_ENTETE, column=c).value for c in range(1, NB_COL + 1)}

    # nettoyage de tout ce qui est sous l'en-tête du devis
    for m in [m for m in ws.merged_cells.ranges if m.min_row >= R_TITRE]:
        ws.unmerge_cells(str(m))
    ws.delete_rows(R_TITRE, ws.max_row - R_TITRE + 1)
    for r in [r for r in list(ws.row_dimensions.keys()) if r >= R_TITRE]:   # formats de ligne entière de l'original
        del ws.row_dimensions[r]
    ws.data_validations.dataValidation = []                                  # listes de l'original, sur d'anciennes lignes

    # en-tête du devis
    if _v(camp.get("numero_projet")):
        ws["D6"] = f"DEVIS (N°Projet) : {camp['numero_projet']}"
    ws["F9"] = camp.get("nom") or None
    ws["F11"] = camp.get("consultant") or None

    df = lignes.copy()
    df["Section"] = df["Section"].fillna("Branding")
    sections = list(dict.fromkeys(df["Section"]))
    r = R_TITRE
    cellules = {"total_q": [], "total_o": [], "dar": [], "social_q": [], "dv_o": [], "xpln_o": []}
    dar_o = {}
    for sec in sections:
        grp = df[df["Section"] == sec]
        titre = TITRE_SECTION.get(sec, str(sec))
        _appliquer(ws, r, S["titre"]); _appliquer(ws, r + 1, S["fees"])
        ws.cell(row=r, column=3, value=titre)
        ws.merge_cells(start_row=r, start_column=3, end_row=r + 1, end_column=4)
        ws.merge_cells(start_row=r, start_column=5, end_row=r + 1, end_column=8)
        ws.cell(row=r + 1, column=19, value="Montant Fees Meta 3%")
        re_ = r + 2
        _appliquer(ws, re_, S["entete"])
        for c, v in entete_valeurs.items():
            ws.cell(row=re_, column=c, value=v)
        ws.cell(row=re_, column=2, value=str(titre).upper())
        ws.cell(row=re_, column=3, value="HIPPIQUE" if str(camp.get("hs", "H")).upper() == "H" else "SPORT")
        r1 = re_ + 1
        for k, (_, l) in enumerate(grp.iterrows()):
            rr = r1 + k
            _appliquer(ws, rr, S["ligne"])
            device = str(_v(l.get("Device")) or "")
            vals = {4: _v(l.get("Levier")), 5: _v(l.get("Régie")), 6: _v(l.get("Objectif")), 7: _v(l.get("Emplacement")),
                    8: _v(l.get("Ciblage PM client")) or _v(l.get("Ciblage")),
                    9: device.capitalize() if device.isupper() else device or None,
                    10: _v(l.get("Type d'achat")), 11: _v(l.get("Format")), 12: _v(l.get("Créative")) or "TBD",
                    13: _date(l.get("Début")), 14: _date(l.get("Fin")), 16: _v(l.get("CPM (€)")),
                    17: _v(l.get("Budget net (€)")), 18: _v(l.get("Commentaire"))}
            for c, v in vals.items():
                ws.cell(row=rr, column=c, value=v)
            ws.cell(row=rr, column=15, value=f"=(Q{rr}/P{rr})*1000")
            if str(l.get("Levier") or "").upper() == "SOCIAL":
                ws.cell(row=rr, column=19, value=f"=Q{rr}*0.03")
                cellules["social_q"].append(f"Q{rr}")
            if frais.actif(l.get("Double Verify")):
                cellules["dv_o"].append(f"O{rr}")
            if frais.actif(l.get("XPLN")):
                cellules["xpln_o"].append(f"O{rr}")
            if frais.actif(l.get("DAR")):
                dar_o.setdefault(sec, []).append(f"O{rr}")
        rl = r1 + len(grp) - 1
        rt = rl + 1
        _appliquer(ws, rt, S["total"])
        ws.cell(row=rt, column=15, value=f"=SUM(O{r1}:O{rl})")
        ws.cell(row=rt, column=16, value=f"=Q{rt}/O{rt}*1000")
        ws.cell(row=rt, column=17, value=f"=SUM(Q{r1}:Q{rl})")
        ws.merge_cells(start_row=re_, start_column=2, end_row=rt, end_column=2)
        ws.merge_cells(start_row=re_, start_column=3, end_row=rt, end_column=3)
        cellules["total_q"].append(f"Q{rt}"); cellules["total_o"].append(f"O{rt}")
        rd = rt + 2
        if dar_o.get(sec):
            vol = "(" + "+".join(dar_o[sec]) + ")"
            _appliquer(ws, rd, S["dar"])
            ws.cell(row=rd, column=14, value=f"DAR {titre}")
            ws.cell(row=rd, column=15, value=f"=IF({vol}<={frais.DAR_SEUIL},{frais.DAR_FORFAIT},"
                                             f"{frais.DAR_FORFAIT}+(({vol}-{frais.DAR_SEUIL})*{frais.DAR_TAUX})/1000)")
            cellules["dar"].append(f"O{rd}")
        r = rd + 2

    # récapitulatif
    r += 1
    somme = lambda cs, f: f"=({'+'.join(cs)})" if cs else "=0"
    recap = [("Net media", somme(cellules["total_q"], None)), ("Total Frais", None),
             ("META", f"=({'+'.join(cellules['social_q'])})*{frais.META_TAUX}" if cellules["social_q"] else "=0"),
             ("DV", f"=((({'+'.join(cellules['dv_o'])})*{frais.DV_TAUX})/1000)*{frais.MAJORATION}" if cellules["dv_o"] else "=0"),
             ("XPLN", f"=((({'+'.join(cellules['xpln_o'])})*{frais.XPLN_TAUX})/1000)*{frais.MAJORATION}" if cellules["xpln_o"] else "=0"),
             ("DAR", somme(cellules["dar"], None)), ("Total Budget Campagne", None)]
    r_net, r_frais = r, r + 1
    for i, (lab, formule) in enumerate(recap):
        rr = r + i
        _appliquer(ws, rr, S["recap_total"] if lab.startswith("Total Budget") else S["recap"])
        ws.cell(row=rr, column=15, value=lab)
        if lab == "Total Frais":
            formule = f"=SUM(R{r_frais + 1}:R{r_frais + 4})"
        if lab.startswith("Total Budget"):
            formule = f"=R{r_frais}+R{r_net}"
        ws.cell(row=rr, column=18, value=formule)
        ws.cell(row=rr, column=18).number_format = '#,##0.00 "€"'
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()
