"""PM commun (PMU) -> fichier média PMU.

Reproduit exactement les formules du fichier média actuel (85_DIB_SEPTEMBRE_PMU_PDV_PROMO_BRAND),
pour que l'app Streamlit PMU qui le consomme reçoive le même fichier qu'aujourd'hui.
Une ligne de fichier média = une ligne de l'onglet Créas (une taille de format).
"""
from __future__ import annotations

import io

import openpyxl
import pandas as pd
from openpyxl.styles import Font, PatternFill

from pm_builder.frais import actif
from pm_builder.pmu_validation import clean_url

COLONNES = ["Régie", "Programmation", "Support", "Emplacement / Ciblage", "Device", "Mode d'Achat", "Type de Tracking",
            "Type de Format", "Format", "Piste Créative", "PLACEMENT NAME DCM", "Créative DCM", "Landing Page",
            "Date de début", "Date de fin", "Type d'achat", "Emplacement DSP", "Créa DSP", "Double Verify", "XPLN",
            "DAR", "Adjust", "URL avec les utms"]


def _s(v) -> str:
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v)


def nom_campagne_cm(camp: dict) -> str:
    """=UPPER(85/2010 & "_" & DIB & "_" & nom & "_" & Promo/Offre & "_" & Levier)"""
    return "_".join(_s(camp[k]) for k in ("univers", "medium", "nom", "promo", "levier")).upper()


def _emplacement_dsp(placement: str, type_achat: str) -> str:
    """Formule du fichier : on garde les 6 premiers champs, on retire le 7e (taille), puis « _ » + type d'achat."""
    morceaux = placement.split("_")
    return "_".join(morceaux[:6]) + "_" + "_".join(morceaux[7:]) + "_" + type_achat


def generer(camp: dict, lignes: pd.DataFrame, creas: pd.DataFrame) -> pd.DataFrame:
    nom_cm = nom_campagne_cm(camp)
    par_id = {str(r["N° ligne"]): r for _, r in lignes.iterrows()}
    res = []
    for _, c in creas.iterrows():
        l = par_id.get(str(c["N° ligne"]))
        if l is None:
            continue
        regie, taille, piste, support = _s(l["Régie"]), _s(c["Taille"]), _s(c["Piste créative / message"]), _s(l["Support"])
        placement = "_".join([_s(camp["levier"]), _s(camp["hs"]), regie, _s(l["Device"]), _s(l["Ciblage"]),
                              _s(l["Type de format"]), taille, _s(l["Type de tracking"])]).upper()
        creative = "_".join([_s(camp["levier"]), _s(camp["hs"]), _s(camp["nom"]), _s(camp["promo"]), taille, piste, support]).lower().replace(" ", "")
        landing = _s(c["Landing (si différente)"]) or _s(camp["landing"])
        url = (f"{landing}?utm_medium={_s(camp['medium'])}&utm_source={regie}&utm_content={placement}"
               f"&utm_campaign={nom_cm}&utm_term={creative}").replace("+", "").replace(" ", "")
        url = clean_url(url)   # mêmes règles que l'app automation PMU : le fichier sort déjà propre
        achat = _s(l["Type d'achat"])
        res.append({
            "Régie": regie, "Programmation": _s(l["Programmation"]), "Support": support,
            "Emplacement / Ciblage": _s(l["Ciblage"]), "Device": _s(l["Device"]), "Mode d'Achat": achat,
            "Type de Tracking": _s(l["Type de tracking"]), "Type de Format": _s(l["Type de format"]), "Format": taille,
            "Piste Créative": piste, "PLACEMENT NAME DCM": placement, "Créative DCM": creative, "Landing Page": landing,
            "Date de début": l["Début"], "Date de fin": l["Fin"], "Type d'achat": achat,
            "Emplacement DSP": _emplacement_dsp(placement, achat),
            "Créa DSP": f"{creative}_{piste}_{nom_cm}".upper(),
            "Double Verify": "OUI" if actif(l.get("Double Verify")) else "",
            "XPLN": "OUI" if actif(l.get("XPLN")) else "",
            "DAR": "Oui" if actif(l.get("DAR")) else "", "Adjust": "", "URL avec les utms": url,
        })
    return pd.DataFrame(res, columns=COLONNES)


def generer_par_levier(camp: dict, lignes: pd.DataFrame, creas: pd.DataFrame) -> dict:
    """Un fichier média par levier de campagne : la section de chaque ligne donne le levier
    (Branding -> BRAND, Considération -> CONSID), sinon le levier de l'en-tête. Renvoie {levier: (campagne, tableau)}."""
    from pm_builder.modele import LEVIER_PAR_SECTION
    leviers = lignes["Section"].map(lambda x: LEVIER_PAR_SECTION.get(_s(x))).fillna(camp["levier"])
    res = {}
    for lev in dict.fromkeys(leviers):
        c = dict(camp, levier=lev, type_campagne=lev)
        res[lev] = (c, generer(c, lignes[leviers == lev], creas))
    return res


def en_excel(camp: dict, df: pd.DataFrame) -> bytes:
    """Même disposition que le fichier média actuel : en-tête lignes 1 à 12, tableau à partir de la ligne 13."""
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "fichier_media"
    nom_cm = nom_campagne_cm(camp)
    entete = [("Consultant email", camp.get("consultant")), ("Levier", camp["levier"]), ("H/S ", camp["hs"]),
              ("Nom de campagne", camp["nom"]), ("85/2010", camp["univers"]), ("DIB", camp["medium"]),
              ("Promo/Offre", camp["promo"]), ("Nom de campagne CM", nom_cm), ("Numéro de projet", camp.get("numero_projet")),
              ("Type de campagne", camp.get("type_campagne") or camp["levier"]),
              ("Campagne DSP", f"{nom_cm}_{_s(camp.get('numero_projet'))}"), ("Advertiser CM", "PMU_MASTER_TAG")]
    for i, (k, v) in enumerate(entete, 1):
        ws.cell(row=i, column=1, value=k).font = Font(bold=True); ws.cell(row=i, column=2, value=v)
    for j, h in enumerate(COLONNES, 1):
        c = ws.cell(row=13, column=j, value=h); c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color="1F3864")
    for i, (_, r) in enumerate(df.iterrows(), 14):
        for j, h in enumerate(COLONNES, 1):
            v = r[h]
            ws.cell(row=i, column=j, value=None if v == "" else v)
            if h.startswith("Date"):
                ws.cell(row=i, column=j).number_format = "DD/MM/YYYY"
    from pm_builder.modele import LP_PMU
    lp = wb.create_sheet("LPS")
    for j, h in enumerate(("Landing Page ID", "Landing Page Name", "URL"), 1):
        lp.cell(row=1, column=j, value=h).font = Font(bold=True)
    for i, ligne in enumerate(LP_PMU, 2):
        for j, v in enumerate(ligne, 1):
            lp.cell(row=i, column=j, value=int(v) if j == 1 else v)
    from pm_builder.modele import ADVERTISERS_PMU, SITES_PMU
    for titre, entetes, lignes_ref in (("Sites", ("Site ID", "Site Name"), SITES_PMU),
                                       ("Advertiser", ("Advertiser ID", "Advertiser Name"), ADVERTISERS_PMU)):
        f = wb.create_sheet(titre)
        for j, h in enumerate(entetes, 1):
            f.cell(row=1, column=j, value=h).font = Font(bold=True)
        for i, (ident, nom) in enumerate(lignes_ref, 2):
            f.cell(row=i, column=1, value=int(ident)); f.cell(row=i, column=2, value=nom)
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()
