"""Export du PM commun en Excel, prêt à partager avec le client (une mise en forme par client)."""
from __future__ import annotations

import io

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL

from pm_builder.modele import COLS_CLIENT, COLS_COMMUNES

F = "Arial"
HF = PatternFill("solid", start_color="1F3864"); TOT = PatternFill("solid", start_color="D9E1F2")
thin = Side(style="thin", color="BFBFBF"); BD = Border(left=thin, right=thin, top=thin, bottom=thin)


def _feuille(ws, df: pd.DataFrame, r0: int, largeur=16):
    for j, h in enumerate(df.columns, 1):
        c = ws.cell(row=r0, column=j, value=h); c.font = Font(name=F, bold=True, color="FFFFFF"); c.fill = HF; c.border = BD
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.column_dimensions[CL(j)].width = max(largeur, min(40, len(str(h)) + 4))
    for i, (_, r) in enumerate(df.iterrows(), r0 + 1):
        for j, h in enumerate(df.columns, 1):
            v = r[h]
            v = None if (v is None or (isinstance(v, float) and pd.isna(v))) else v
            c = ws.cell(row=i, column=j, value=v); c.border = BD; c.font = Font(name=F, size=10)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if h in ("Début", "Fin"):
                c.number_format = "DD/MM/YYYY"
    return r0 + len(df)


def exporter(camp: dict, lignes: pd.DataFrame, creas: pd.DataFrame, audiences: pd.DataFrame) -> bytes:
    if camp["client"] == "Carrefour":              # PM client Carrefour : mise en forme du PM Carrefour d'origine
        from pm_builder.export_pm_crf import exporter as exporter_crf
        return exporter_crf(camp, lignes, audiences)
    if camp["client"] == "PMU":                    # PM client PMU : mise en forme du PM PMU d'origine
        from pm_builder.export_pm_pmu import exporter as exporter_pmu
        return exporter_pmu(camp, lignes)
    client = camp["client"]
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "PM DISPLAY" if client == "Carrefour" else "DIGITAL"
    ws["A1"] = f"Plan média {client} : {camp['nom']}"; ws["A1"].font = Font(name=F, bold=True, size=14)
    infos = [("Campagne", camp["nom"]), ("Début", camp["debut"]), ("Fin", camp["fin"]), ("Landing page", camp["landing"])]
    infos += ([("Pilier", camp["pilier"]), ("Objectif", camp["objectif"])] if client == "Carrefour" else
              [("Levier", camp["levier"]), ("H/S", camp["hs"]), ("85/2010", camp["univers"]), ("Medium", camp["medium"]), ("Promo/Offre", camp["promo"])])
    for i, (k, v) in enumerate(infos, 3):
        ws.cell(row=i, column=1, value=k).font = Font(name=F, bold=True)
        c = ws.cell(row=i, column=2, value=v)
        if k in ("Début", "Fin"):
            c.number_format = "DD/MM/YYYY"
    df = lignes[["N° ligne"] + COLS_COMMUNES + COLS_CLIENT[client]].copy()
    b = pd.to_numeric(df["Budget net (€)"], errors="coerce"); p = pd.to_numeric(df["CPM (€)"], errors="coerce")
    df.insert(df.columns.get_loc("CPM (€)") + 1, "Volume estimé", (b / p * 1000).round(0))
    r0 = len(infos) + 5
    rl = _feuille(ws, df, r0)
    rt = rl + 1
    ws.cell(row=rt, column=1, value="TOTAL")
    for j, h in enumerate(df.columns, 1):
        c = ws.cell(row=rt, column=j); c.fill = TOT; c.font = Font(name=F, bold=True); c.border = BD
        if h in ("Budget net (€)", "Volume estimé"):
            c.value = f"=SUM({CL(j)}{r0 + 1}:{CL(j)}{rl})"
        if h in ("Budget net (€)", "CPM (€)"):
            for r in range(r0 + 1, rt + 1):
                ws.cell(row=r, column=j).number_format = "#,##0.00 €"
        if h == "Volume estimé":
            for r in range(r0 + 1, rt + 1):
                ws.cell(row=r, column=j).number_format = "#,##0"
    ws.freeze_panes = ws.cell(row=r0 + 1, column=2)
    _feuille(wb.create_sheet("CREAS"), creas, 1)
    if client == "Carrefour":
        _feuille(wb.create_sheet("AUDIENCES"), audiences, 1)
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()
