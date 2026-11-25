"""Construction de la fiche setup pré-remplie (ce que l'expert vérifie dans l'app).

Chaque ligne a un « Statut » :
- validé : trouvé dans les tables, convention confirmée ;
- proposé : trouvé dans les tables, mais proposition encore à confirmer par un expert ;
- inconnu : rien trouvé, cases vides à compléter.
"""
from __future__ import annotations

from collections import Counter

import pandas as pd

import tables as T
from generation.nommage import nom_campagne_libre

COLS_FORMATS = ["Bloc", "Régie PM", "Format PM", "Type d'achat", "DSP", "Platform", "Régie", "Format Type",
                "Format size", "Tracking Type", "Code régie", "Code format", "Une ligne par message", "Statut"]
COLS_AUDIENCES = ["Bloc", "Régie PM", "Audience PM", "% budget", "Targeting Type", "Origine Targeting Type",
                  "Targeting details", "Code créa", "Statut"]
COLS_MESSAGES = ["Bloc", "Régie PM", "Code format", "Message PDR", "Code message", "Statut"]


def _premier(valeurs):
    for v in valeurs:
        if v:
            return v
    return ""


def _statut(*parties) -> str:
    if any(p is None for p in parties):
        return T.INCONNU
    return T.PROPOSE if any(p.get("a_valider") for p in parties) else T.VALIDE


def campagne(pm: dict, pdr_campagne: list[dict], nom_pdr: str, listes: dict | None = None) -> dict:
    L = listes or T.listes()
    objectifs = [o for o in (T.objectif(l["objectif"]) for l in pdr_campagne) if o]
    pilier_pdr = _premier(l["pilier"] for l in pdr_campagne)
    return {
        "nom_pdr": nom_pdr,
        "nom": nom_campagne_libre(nom_pdr),
        "pilier": T.valeur_de_liste(pilier_pdr, L["Pilier"]) or pilier_pdr,
        "objectif": Counter(objectifs).most_common(1)[0][0] if objectifs else "",
        "objectifs_pdr": sorted({l["objectif"] for l in pdr_campagne if l["objectif"]}),
        "debut": pm["debut"],
        "fin": pm["fin"],
        "landing": _premier(l["landing"] for l in pdr_campagne),
    }


def formats(pm: dict, listes: dict | None = None) -> pd.DataFrame:
    L = listes or T.listes()
    c = T.constantes()
    lignes = []
    for f in pm["formats"]:
        cle = T.trouver_regie(f["regie_pm"])
        info = T.info_regie(cle) or None
        trouves = T.trouver_formats(cle, f["format_pm"]) or [None]
        for e in trouves:
            code = (e or {}).get("code")
            if code and "{duree}" in code:
                d = T.duree(f["format_pm"])
                code = code.replace("{duree}", d) if d else code
            tailles = T.tailles(f["format_pm"]) if e and e.get("tailles") else []
            for taille in tailles or [(e or {}).get("format_size") or c["format_size"]]:
                lignes.append({
                    "Bloc": f["bloc"], "Régie PM": f["regie_pm"], "Format PM": f["format_pm"],
                    "Type d'achat": f["type_achat"], "DSP": f["dsp"],
                    "Platform": T.plateforme(cle, f["type_achat"], f["dsp"], L["Plateforme"]),
                    "Régie": (info or {}).get("regie"),
                    "Format Type": (e or {}).get("format_type"),
                    "Format size": taille,
                    "Tracking Type": (e or {}).get("tracking_type") or c["tracking_type"],
                    "Code régie": (info or {}).get("code"),
                    "Code format": code,
                    "Une ligne par message": bool(e and e.get("ligne_par_message")),
                    "Statut": _statut(info, e),
                })
    return pd.DataFrame(lignes, columns=COLS_FORMATS)


def audiences(pm: dict) -> pd.DataFrame:
    lignes = []
    for a in pm["audiences"]:
        cle = T.trouver_regie(a["regie_pm"])
        res = T.trouver_audience(cle, a["audience_pm"])
        lu = T.type_ciblage(a["audience_pm"])          # le PM fait foi s'il nomme le type de ciblage
        if lu:
            tt, origine = lu[0], f"PM (« {lu[1]} »)"
        elif res:
            tt, origine = res["targeting_type"], "table"
        else:
            tt, origine = None, ""
        lignes.append({
            "Bloc": a["bloc"], "Régie PM": a["regie_pm"], "Audience PM": a["audience_pm"], "% budget": a["pct"],
            "Targeting Type": tt, "Origine Targeting Type": origine,
            "Targeting details": res["targeting_details"] if res else None,
            "Code créa": res["code"] if res else None,
            "Statut": res["statut"] if res else T.INCONNU,
        })
    return pd.DataFrame(lignes, columns=COLS_AUDIENCES)


def messages(df_formats: pd.DataFrame, pdr_campagne: list[dict]) -> pd.DataFrame:
    """Messages du PDR pour les formats « une ligne par message »."""
    lignes = []
    for _, f in df_formats[df_formats["Une ligne par message"].astype(bool)].iterrows():
        cle = T.trouver_regie(f["Régie PM"])
        entrees = [e for e in T.trouver_formats(cle, f["Format PM"]) if e.get("code") == f["Code format"]] \
            or T.trouver_formats(cle, f["Format PM"]) or [{"pdr": [str(f["Format PM"])]}]
        vus = []
        for l in pdr_campagne:
            if T.trouver_regie(l["regie"]) == cle and T.format_correspond_pdr(entrees[0], l["format"]) \
                    and l["message"] and l["message"] not in vus:
                vus.append(l["message"])
        for m in vus:
            code = T.code_message(m)
            lignes.append({"Bloc": f["Bloc"], "Régie PM": f["Régie PM"], "Code format": f["Code format"],
                           "Message PDR": m, "Code message": code,
                           "Statut": T.VALIDE if code else T.INCONNU})
    return pd.DataFrame(lignes, columns=COLS_MESSAGES).drop_duplicates()
