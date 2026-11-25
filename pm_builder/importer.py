"""Pré-remplir le constructeur de PM à partir d'un fichier :
- le modèle Excel commun (saisie rapide dans Excel) ;
- un PM au format actuel (Carrefour + PDR, ou PMU + brief créa).
Chaque fonction renvoie (campagne, lignes, créas, audiences, remarques) ; les remarques signalent ce qui est à vérifier.
"""
from __future__ import annotations

import re
from datetime import date, datetime

import openpyxl
import pandas as pd

import tables as T
from generation import setup
from lecture.pdr import filtrer, lire_pdr
from lecture.pm import lire_pm
from lecture.pm_pmu import lire_brief_pmu, lire_pm_pmu
from pm_builder import modele


def _d(v):
    return v.date() if isinstance(v, datetime) else (v if isinstance(v, date) else None)


def _s(v) -> str:
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()


def _cle(t) -> str:
    return re.sub(r"[^a-z0-9]", "", T.normaliser(t))


def _dans_liste(v, liste):
    """Valeur de la liste qui correspond (sans tenir compte des majuscules, espaces, accents), sinon la valeur brute."""
    return next((x for x in liste if _cle(x) == _cle(v)), _s(v) or None)


# ================================================================== modèle Excel commun
def _lp(v):
    """Landing PMU : accepte le nom de la LP (liste CM) ou directement son URL."""
    v = _s(v)
    if not v or v in modele.LP_URLS:
        return v
    for _, nom, url in modele.LP_PMU:
        if T.normaliser(nom) == T.normaliser(v):
            return url
    return v


def est_modele(fichier) -> bool:
    wb = openpyxl.load_workbook(fichier, read_only=True)
    noms = set(wb.sheetnames)
    return {"PM CARREFOUR", "PM PMU"} <= noms or {"PM", "CREAS"} <= noms


def depuis_modele(fichier, client=None):
    """Modèle v2 (un onglet par client) ; l'ancien modèle v1 (onglet PM unique) reste accepté."""
    wb = openpyxl.load_workbook(fichier, data_only=True)
    if "PM CARREFOUR" not in wb.sheetnames:
        return _depuis_modele_v1(fichier)
    if client not in modele.CLIENTS:
        raise ValueError("Choisissez le client (Carrefour ou PMU) avant d'importer le modèle.")
    rem = []
    if client == "Carrefour":
        pm = wb["PM CARREFOUR"]
        camp = modele.campagne_vide("Carrefour")
        camp.update({"nom": _s(pm["C3"].value), "debut": _d(pm["C4"].value), "fin": _d(pm["C5"].value),
                     "landing": _s(pm["C6"].value), "pilier": _s(pm["C7"].value), "objectif": _s(pm["C8"].value)})
        cols = {"A": "N° ligne", "B": "Section", "C": "Levier", "D": "Régie", "E": "Type d'achat", "F": "Format",
                "G": "Type de format", "H": "Ciblage", "I": "Début", "J": "Fin", "K": "Budget net (€)", "L": "CPM (€)",
                "N": "Groupe d'audiences", "O": "Nb messages", "P": "Durée (s)", "Q": "Tracking Type",
                "U": "Device", "V": "Specs", "W": "KPI", "X": "VTR / LTR", "Y": "Fréquence", "Z": "Études / BLS"}
        camp.update({"libelle": _s(pm["C9"].value), "owner_crf": _s(pm["C10"].value)})
    else:
        pm = wb["PM PMU"]
        camp = modele.campagne_vide("PMU")
        camp.update({"nom": _s(pm["C3"].value), "debut": _d(pm["C4"].value), "fin": _d(pm["C5"].value),
                     "landing": _lp(pm["C6"].value), "levier": _s(pm["C7"].value) or "BRAND", "numero_projet": _s(pm["C8"].value),
                     "hs": _s(pm["F3"].value), "univers": _s(pm["F4"].value), "medium": _s(pm["F5"].value),
                     "promo": _s(pm["F6"].value), "consultant": _s(pm["F7"].value), "advertiser": modele.ADVERTISER_PMU})
        cols = {"A": "N° ligne", "B": "Section", "C": "Levier", "D": "Régie", "E": "Type d'achat", "F": "Format",
                "G": "Type de format", "H": "Ciblage", "I": "Ciblage PM client", "J": "Objectif", "K": "Emplacement",
                "L": "Device", "M": "Programmation", "N": "Support", "O": "Type de tracking", "P": "Double Verify",
                "Q": "XPLN", "R": "DAR", "S": "Créative", "T": "Début", "U": "Fin", "V": "Budget net (€)", "W": "CPM (€)",
                "AD": "Commentaire"}
    rows = []
    for r in range(12, pm.max_row + 1):
        if _s(pm[f"A{r}"].value).upper() == "TOTAL":
            break
        if not _s(pm[f"D{r}"].value):
            continue
        d = {nom: pm[f"{col}{r}"].value for col, nom in cols.items()}
        d["N° ligne"] = _s(d["N° ligne"]) or f"L{r - 11}"
        for k in ("Début", "Fin"):
            d[k] = _d(d[k])
        rows.append(d)
    lignes = pd.DataFrame(rows, columns=["N° ligne"] + modele.COLS_COMMUNES + modele.COLS_CLIENT[client])
    if lignes.empty:
        raise ValueError(f"Aucune ligne remplie dans l'onglet PM {client.upper()} du modèle.")

    if client == "Carrefour":
        lignes["Tracking Type"] = lignes["Tracking Type"].fillna("Pixel+Click-Command")
        lignes["Nb messages"] = pd.to_numeric(lignes["Nb messages"], errors="coerce").fillna(1)
        au, me = wb["AUDIENCES CARREFOUR"], wb["MESSAGES CARREFOUR"]
        aud = []
        for r in range(3, au.max_row + 1):
            if not _s(au[f"A{r}"].value):
                continue
            p = pd.to_numeric(au[f"G{r}"].value, errors="coerce")
            aud.append({"Groupe d'audiences": _s(au[f"A{r}"].value), "Audience": au[f"C{r}"].value,
                        "Type de ciblage": au[f"D{r}"].value, "Targeting details": au[f"F{r}"].value,
                        "% du groupe": (round(p * 100, 4) if pd.notna(p) and p <= 1 else p)})
        format_de = dict(zip(lignes["N° ligne"], lignes["Format"]))
        creas = [{"N° ligne": _s(me[f"A{r}"].value), "Nom du format": format_de.get(_s(me[f"A{r}"].value)), "Taille": "1x1",
                  "Piste créative / message": me[f"D{r}"].value, "Code message": me[f"E{r}"].value, "Landing (si différente)": None}
                 for r in range(3, me.max_row + 1) if _s(me[f"A{r}"].value)]
        return camp, lignes, pd.DataFrame(creas, columns=modele.COLS_CREAS), pd.DataFrame(aud, columns=modele.COLS_AUDIENCES), rem

    cp = wb["CRÉAS PMU"]
    creas = []
    for r in range(3, cp.max_row + 1):
        if not _s(cp[f"A{r}"].value):
            continue
        nom, piste = cp[f"D{r}"].value, cp[f"F{r}"].value
        creas.append({"N° ligne": _s(cp[f"A{r}"].value), "Nom du format": nom, "Taille": cp[f"E{r}"].value,
                      "Piste créative / message": piste if _s(piste) else nom, "Code message": None,
                      "Landing (si différente)": _lp(cp[f"G{r}"].value) or None})
    from pm_builder.frais import actif
    for col in ("Double Verify", "XPLN", "DAR"):
        lignes[col] = lignes[col].map(actif)
    lignes, notes = modele.completer_pmu(lignes)
    return camp, lignes, pd.DataFrame(creas, columns=modele.COLS_CREAS), modele.audiences_vides(), rem + notes


def _depuis_modele_v1(fichier):
    wb = openpyxl.load_workbook(fichier, data_only=True)
    pm, cr, au = wb["PM"], wb["CREAS"], wb["AUDIENCES"]
    client = _s(pm["C3"].value)
    if client not in modele.CLIENTS:
        raise ValueError("Client manquant en C3 de l'onglet PM (Carrefour ou PMU).")
    camp = modele.campagne_vide(client)
    camp.update({"nom": _s(pm["C4"].value), "debut": _d(pm["C5"].value), "fin": _d(pm["C6"].value), "landing": _s(pm["C7"].value)})
    if client == "Carrefour":
        camp.update({"pilier": _s(pm["F3"].value), "objectif": _s(pm["F4"].value)})
    else:
        for k, cell in (("levier", "I3"), ("hs", "I4"), ("univers", "I5"), ("medium", "I6"), ("promo", "I7"), ("advertiser", "I8")):
            if _s(pm[cell].value):
                camp[k] = _s(pm[cell].value)
    noms = ["N° ligne"] + modele.COLS_COMMUNES[:10] + ["CPM (€)", "_vol", "Groupe d'audiences", "Nb messages", "Durée (s)",
                                                       "Device", "Programmation", "Support", "Type de tracking", "Double Verify", "DAR"]
    rows = []
    for r in range(12, pm.max_row + 1):
        vals = [pm.cell(row=r, column=c).value for c in range(1, len(noms) + 1)]
        if _s(vals[0]).upper() == "TOTAL":
            break
        if not _s(vals[3]):
            continue
        d = dict(zip(noms, vals))
        d["N° ligne"] = _s(d["N° ligne"]) or f"L{r - 11}"
        d.pop("_vol")
        rows.append(d)
    lignes = pd.DataFrame(rows, columns=["N° ligne"] + modele.COLS_COMMUNES + modele.COLS_CLIENT[client] + ["_x"]).drop(columns="_x")
    if client == "Carrefour":
        lignes["Tracking Type"] = lignes["Tracking Type"].fillna("Pixel+Click-Command")
    creas = pd.DataFrame([{"N° ligne": _s(cr.cell(row=r, column=1).value), "Nom du format": cr.cell(row=r, column=3).value,
                           "Taille": cr.cell(row=r, column=4).value, "Piste créative / message": cr.cell(row=r, column=6).value,
                           "Code message": cr.cell(row=r, column=7).value, "Landing (si différente)": cr.cell(row=r, column=8).value}
                          for r in range(2, cr.max_row + 1) if _s(cr.cell(row=r, column=1).value)], columns=modele.COLS_CREAS)
    vide = creas["Piste créative / message"].isna() | (creas["Piste créative / message"].astype(str).str.strip() == "")
    creas.loc[vide, "Piste créative / message"] = creas.loc[vide, "Nom du format"]   # par défaut : le nom du format
    aud = []
    for r in range(2, au.max_row + 1):
        if not _s(au.cell(row=r, column=1).value):
            continue
        p = pd.to_numeric(au.cell(row=r, column=7).value, errors="coerce")
        aud.append({"Groupe d'audiences": _s(au.cell(row=r, column=1).value), "Audience": au.cell(row=r, column=3).value,
                    "Type de ciblage": au.cell(row=r, column=4).value, "Targeting details": au.cell(row=r, column=6).value,
                    "% du groupe": (p * 100 if pd.notna(p) and p <= 1 else p)})
    return camp, lignes, creas, pd.DataFrame(aud, columns=modele.COLS_AUDIENCES), []


# ================================================================== PM Carrefour actuel (+ PDR)
INVERSE_TC = {v: k for k, v in modele.TYPES_CIBLAGE.items()}


def _achat_crf(t) -> str | None:
    if T.contient(t, "gag"):
        return "GAG"
    if T.contient(t, "pmp") or T.contient(t, "private"):
        return "PMP"
    if T.contient(t, "reservation"):
        return "Réservation"
    if T.contient(t, "pg"):
        return "PG"
    if T.contient(t, "cpm"):
        return "CPM Floor"
    return None


def depuis_pm_crf(fichier_pm, onglet=None, fichier_pdr=None, campagne_pdr=None):
    pm = lire_pm(fichier_pm, onglet)
    pdr = filtrer(lire_pdr(fichier_pdr), campagne_pdr) if (fichier_pdr and campagne_pdr) else []
    remarques = []
    if not pdr:
        remarques.append("Sans PDR : pilier, objectif, landing et messages sont à compléter.")
    camp = modele.campagne_vide("Carrefour")
    c = setup.campagne(pm, pdr, campagne_pdr or pm.get("campagne") or "")
    camp.update({"nom": campagne_pdr or _s(pm.get("campagne")).replace(" ", "_"), "debut": _d(c["debut"]), "fin": _d(c["fin"]),
                 "landing": c["landing"], "pilier": c["pilier"], "objectif": c["objectif"]})
    df_f, df_a = setup.formats(pm), setup.audiences(pm)
    df_m = setup.messages(df_f, pdr)
    L = modele.listes()["Carrefour"]
    # groupe d'audiences : un code court par bloc du PM
    groupes, pris = {}, set()
    for bloc, regie in df_f[["Bloc", "Régie"]].drop_duplicates().itertuples(index=False):
        base = _cle(T.info_regie(T.trouver_regie(regie)).get("code") or regie).upper()[:6] or "G"
        code, k = base, 2
        while code in pris:
            code, k = f"{base}{k}", k + 1
        groupes[bloc] = code; pris.add(code)
    budgets = {(f["bloc"], f["format_pm"]): (f.get("budget"), f.get("cpm")) for f in pm["formats"]}
    lignes, creas = [], []
    for i, (_, f) in enumerate(df_f.iterrows(), 1):
        nid = f"L{i}"
        b, p = budgets.get((f["Bloc"], f["Format PM"]), (None, None))
        msgs = df_m[(df_m["Bloc"] == f["Bloc"]) & (df_m["Code format"] == f["Code format"])] if f["Une ligne par message"] else df_m.iloc[0:0]
        lignes.append({"N° ligne": nid, "Section": "Branding", "Levier": None, "Régie": f["Régie"],
                       "Type d'achat": _achat_crf(f["Type d'achat"]), "Format": f["Format PM"], "Type de format": f["Format Type"],
                       "Ciblage": None, "Début": _d(c["debut"]), "Fin": _d(c["fin"]), "Budget net (€)": b, "CPM (€)": p,
                       "Groupe d'audiences": groupes[f["Bloc"]], "Nb messages": max(len(msgs), 1),
                       "Durée (s)": pd.to_numeric(T.duree(f["Format PM"]), errors="coerce"), "Tracking Type": f["Tracking Type"]})
        for _, m in msgs.iterrows():
            creas.append({"N° ligne": nid, "Nom du format": f["Format PM"], "Taille": "1x1",
                          "Piste créative / message": m["Message PDR"], "Code message": m["Code message"], "Landing (si différente)": None})
        if f["Statut"] != "validé":
            remarques.append(f"{nid} ({f['Régie PM']} · {f['Format PM'][:40]}) : valeurs {f['Statut']}es, à vérifier.")
    aud = []
    for _, a in df_a.iterrows():
        aud.append({"Groupe d'audiences": groupes.get(a["Bloc"]), "Audience": a["Audience PM"],
                    "Type de ciblage": INVERSE_TC.get(a["Targeting Type"]), "Targeting details": a["Targeting details"],
                    "% du groupe": a["% budget"]})
        if a["Statut"] != "validé":
            remarques.append(f"Audience « {a['Audience PM'][:50]} » : valeurs {a['Statut']}es, à vérifier.")
    for g in {x["Groupe d'audiences"] for x in aud}:
        pcts = [x["% du groupe"] for x in aud if x["Groupe d'audiences"] == g]
        if all(p is None or pd.isna(p) for p in pcts):     # pas de % dans le PM : répartition égale, à ajuster
            for x in aud:
                if x["Groupe d'audiences"] == g:
                    x["% du groupe"] = round(100 / len(pcts), 2)
            remarques.append(f"Groupe {g} : pas de % dans le PM, répartition égale proposée.")
    return camp, pd.DataFrame(lignes), pd.DataFrame(creas, columns=modele.COLS_CREAS), pd.DataFrame(aud, columns=modele.COLS_AUDIENCES), remarques


# ================================================================== PM PMU actuel (+ brief créa)
FAMILLES = {"smart cover": ["smart cover", "smartcover"], "habillage": ["habillage"],
            "interstitiel": ["interstitiel", "intersitiel", "interstiel"],
            "iab": ["iab", "pave", "grand angle", "banniere", "skyscraper"], "story": ["story", "stories", "reel"],
            "feed": ["fil d'actualite", "feed"], "bandeau": ["bandeau"]}


def _familles(texte):
    return [f for f, mots in FAMILLES.items() if any(T.contient(texte, m) for m in mots)]


def sections_pmu(fichier_pm):
    return lire_pm_pmu(fichier_pm)["sections"]


def depuis_pm_pmu(fichier_pm, sections=None, fichier_brief=None):
    d = lire_pm_pmu(fichier_pm)
    brief = lire_brief_pmu(fichier_brief) if fichier_brief else []
    L = modele.listes()["PMU"]
    remarques = ["Ciblage : vérifiez qu'il s'agit du libellé court voulu dans les noms CM (ex. Turfistes reguliers)."]
    if not brief:
        remarques.append("Sans brief créa : les tailles de l'onglet Créas sont à remplir.")
    camp = modele.campagne_vide("PMU")
    lignes_pm = [l for l in d["lignes"] if not sections or l["section"] in sections]
    debuts = [l["debut"] for l in lignes_pm if l["debut"]]; fins = [l["fin"] for l in lignes_pm if l["fin"]]
    camp.update({"debut": _d(min(debuts)) if debuts else None, "fin": _d(max(fins)) if fins else None})
    if fichier_brief:
        ws = openpyxl.load_workbook(fichier_brief, data_only=True).worksheets[0]
        infos = {T.normaliser(ws.cell(row=r, column=2).value): _s(ws.cell(row=r, column=3).value) for r in range(1, 10)}
        if infos.get("campagne") and infos.get("plan"):          # ex. « SEPTEMBRE » + « PMU PDV » -> SEPTEMBRE_PMU_PDV
            camp["nom"] = re.sub(r"\s+", "_", f"{infos['campagne']} {infos['plan']}".strip()).upper()
            remarques.append(f"Nom de campagne « {camp['nom']} » repris du brief créa (Campagne + Plan) : à vérifier.")
    if not camp.get("nom"):
        remarques.append("Nom de campagne absent du PM : à remplir dans l'onglet Campagne.")
    remarques.append("Landing page absente du PM : à remplir dans l'onglet Campagne.")
    lignes, creas = [], []
    for i, l in enumerate(lignes_pm, 1):
        nid, lev = f"L{i}", _s(l["levier"]).upper()
        lev = modele.LEVIER_ALIAS_PMU.get(lev, lev)
        regie = _dans_liste(l["regie"], L["Régie"])
        social = lev.startswith("SOCIAL") or T.contient(l["regie"], "meta")
        direct = lev.startswith("DISPLAY")
        lignes.append({"N° ligne": nid, "Section": "Considération" if T.contient(l["section"], "consideration") else _s(l["section"]) or "Branding",
                       "Levier": _dans_liste(lev, L["Levier"]), "Régie": regie, "Type d'achat": _dans_liste(l["achat"], L["Type d'achat"]),
                       "Format": l["formats"], "Type de format": "STATIC", "Ciblage": l["ciblage"],
                       "Début": _d(l["debut"]), "Fin": _d(l["fin"]), "Budget net (€)": l["budget"], "CPM (€)": l["cpm"],
                       "Device": _dans_liste(l["device"], L["Device"]), "Programmation": "Régie" if direct else "79",
                       "Support": "FACEBOOK" if social else ("UN1T" if not direct else None),
                       "Type de tracking": "PCC", "Double Verify": not social, "XPLN": False, "DAR": True,
                       "Objectif": _s(l.get("objectif")), "Emplacement": _s(l.get("emplacement")) or "RG sur site",
                       "Ciblage PM client": l["ciblage"], "Créative": "TBD", "Commentaire": _s(l.get("commentaire")) or None})
        if direct:
            remarques.append(f"{nid} ({regie}) : achat en direct, programmation « Régie » proposée et support à remplir.")
        if brief:
            dispo = [b for b in brief if _cle(b["regie"]) and (_cle(b["regie"]) in _cle(l["regie"]) or _cle(l["regie"]) in _cle(b["regie"]))]
            fam = _familles(l["formats"])
            choisies = [b for b in dispo if not fam or set(_familles(b["format"])) & set(fam)] or dispo
            vues = set()
            for b in choisies:
                if (b["taille"], b["format"]) in vues:
                    continue
                vues.add((b["taille"], b["format"]))
                creas.append({"N° ligne": nid, "Nom du format": b["format"], "Taille": b["taille"],
                              "Piste créative / message": b["format"], "Code message": None, "Landing (si différente)": None})
            if not dispo:
                remarques.append(f"{nid} ({regie}) : aucune taille trouvée dans le brief, à ajouter dans l'onglet Créas.")
    df_l = pd.DataFrame(lignes)
    if not df_l.empty:
        df_l, notes = modele.completer_pmu(df_l)
        remarques = [r for r in remarques if "support à remplir" not in r] + notes
    return camp, df_l, pd.DataFrame(creas, columns=modele.COLS_CREAS), modele.audiences_vides(), remarques
