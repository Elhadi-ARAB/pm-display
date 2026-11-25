"""PM commun (Carrefour) -> URL Builder, via le générateur existant (generation/url_builder.py)."""
from __future__ import annotations

import pandas as pd

import tables as T
from generation import url_builder
from generation.nommage import nom_campagne_libre, slug
from pm_builder.modele import TYPES_CIBLAGE


def _s(v) -> str:
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()


def code_format(cle_regie, format_libelle, format_type, duree) -> tuple[str, bool]:
    """Code format du nom de créa : table des formats si possible, sinon dérivé du Format Type (signalé)."""
    for e in T.trouver_formats(cle_regie, f"{format_libelle} {format_type}"):
        if e.get("format_type") == format_type or not format_type:
            code = e["code"]
            if "{duree}" in code:
                code = code.replace("{duree}", _s(duree).split(".")[0]) if _s(duree) else code
            return code, True
    return slug(format_type) or slug(format_libelle), False


def code_audience(cle_regie, libelle, targeting_details) -> tuple[str, bool]:
    res = T.trouver_audience(cle_regie, libelle)
    if res and res["targeting_details"] == targeting_details:
        return res["code"], True
    return slug(targeting_details), False


def preparer(camp: dict, lignes: pd.DataFrame, creas: pd.DataFrame, audiences: pd.DataFrame):
    """Construit les tableaux attendus par generation/url_builder.generer. Renvoie aussi les codes dérivés à vérifier."""
    a_verifier = []
    f_rows, a_rows, m_rows = [], [], []
    for _, l in lignes.iterrows():
        cle = T.trouver_regie(l["Régie"])
        info = T.info_regie(cle)
        code_r = info.get("code") or slug(l["Régie"])
        code_f, connu = code_format(cle, _s(l["Format"]), _s(l["Type de format"]), l.get("Durée (s)"))
        if not connu:
            a_verifier.append(f"Code format « {code_f} » déduit pour {l['N° ligne']} ({l['Régie']} · {l['Format']}).")
        msgs = creas[(creas["N° ligne"] == l["N° ligne"]) & creas["Code message"].notna() & (creas["Code message"].astype(str).str.strip() != "")]
        nb = int(pd.to_numeric(l.get("Nb messages"), errors="coerce") or 1)
        f_rows.append({"Bloc": _s(l["Groupe d'audiences"]), "Régie PM": l["Régie"], "Format PM": l["Format"],
                       "Type d'achat": l["Type d'achat"], "DSP": None,
                       "Platform": T.plateforme(cle, l["Type d'achat"]), "Régie": l["Régie"],
                       "Format Type": l["Type de format"], "Format size": T.constantes()["format_size"],
                       "Tracking Type": _s(l.get("Tracking Type")) or T.constantes()["tracking_type"],
                       "Code régie": code_r, "Code format": code_f, "Une ligne par message": nb > 1, "Statut": "validé"})
        if nb > 1:
            for _, m in msgs.iterrows():
                m_rows.append({"Bloc": _s(l["Groupe d'audiences"]), "Régie PM": l["Régie"], "Code format": code_f,
                               "Message PDR": m["Piste créative / message"], "Code message": m["Code message"], "Statut": "validé"})
    regie_du_groupe = {_s(l["Groupe d'audiences"]): l["Régie"] for _, l in lignes.iterrows()}
    for _, a in audiences.iterrows():
        g = _s(a["Groupe d'audiences"])
        cle = T.trouver_regie(regie_du_groupe.get(g))
        code_a, connu = code_audience(cle, _s(a["Audience"]), _s(a["Targeting details"]))
        if not connu:
            a_verifier.append(f"Code créa « {code_a} » déduit pour l'audience « {a['Audience']} » (groupe {g}).")
        a_rows.append({"Bloc": g, "Régie PM": regie_du_groupe.get(g), "Audience PM": a["Audience"],
                       "% budget": a["% du groupe"], "Targeting Type": TYPES_CIBLAGE.get(_s(a["Type de ciblage"])),
                       "Origine Targeting Type": "PM", "Targeting details": a["Targeting details"],
                       "Code créa": code_a, "Statut": "validé"})
    df_f = pd.DataFrame(f_rows); df_a = pd.DataFrame(a_rows)
    df_m = pd.DataFrame(m_rows, columns=["Bloc", "Régie PM", "Code format", "Message PDR", "Code message", "Statut"])
    camp_ub = {"nom": nom_campagne_libre(camp["nom"]), "pilier": camp["pilier"], "objectif": camp["objectif"],
               "debut": camp["debut"], "fin": camp["fin"], "landing": camp["landing"]}
    return camp_ub, df_f, df_a, df_m, a_verifier


def generer(camp, lignes, creas, audiences):
    camp_ub, df_f, df_a, df_m, a_verifier = preparer(camp, lignes, creas, audiences)
    return url_builder.generer(camp_ub, df_f, df_a, df_m), a_verifier
