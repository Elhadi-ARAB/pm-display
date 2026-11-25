"""Fiche setup validée -> lignes de l'URL Builder.

Le DataFrame produit a exactement les colonnes de l'onglet « URL BUILDER »,
ce qui permet de le passer tel quel à l'export CM existant.
"""
from __future__ import annotations

import pandas as pd

import tables as T
from generation.nommage import slug

COLONNES = [
    "Plateforme*", "Pilier", "Objectif", "Nom de la campagne (free form)",
    "Début  JJ/MM/AAAA", "Fin  JJ/MM/AAAA", "Nom de la campagne*", "Message", "code message",
    "Lien", "Lien drive vers la crea", "Platform", "Audience", "Geolocation",
    "Targeting Type", "Targeting details", "Format Type", "Format size (CM only)",
    "Tracking Type", "Régie", "Stratégie", "Complément (Free form)", "XPLN script CM",
    "Placement CM = creative DV360", "Creative Name", "utm_source", "utm_campaign",
    "utm_content", "utm_term", "URL", "UTM", "URL de redirection trackée",
]


def _messages_du_format(df_messages, f):
    return df_messages[(df_messages["Bloc"] == f["Bloc"]) & (df_messages["Code format"] == f["Code format"])]


def plan_de_lignes(df_formats: pd.DataFrame, df_audiences: pd.DataFrame, df_messages: pd.DataFrame):
    """Pour chaque format : × les audiences de son bloc (× les messages si « une ligne par message »)."""
    plan = []
    for _, f in df_formats.iterrows():
        auds = df_audiences[df_audiences["Bloc"] == f["Bloc"]]
        msgs = [None]
        if f["Une ligne par message"]:
            msgs = [r for _, r in _messages_du_format(df_messages, f).iterrows()] or [None]
        for _, a in auds.iterrows():
            for msg in msgs:
                plan.append((f, a, msg))
    return plan


def nombre_attendu(df_formats, df_audiences, df_messages) -> int:
    """Calcul indépendant du nombre de lignes, pour le contrôle."""
    total = 0
    for _, f in df_formats.iterrows():
        n_aud = int((df_audiences["Bloc"] == f["Bloc"]).sum())
        n_msg = max(1, len(_messages_du_format(df_messages, f))) if f["Une ligne par message"] else 1
        total += n_aud * n_msg
    return total


def generer(camp: dict, df_formats, df_audiences, df_messages, constantes: dict | None = None) -> pd.DataFrame:
    c = {**T.constantes(), **(constantes or {})}
    A, B, C, D = c["plateforme_campagne"], camp["pilier"], camp["objectif"], camp["nom"]
    debut, fin = pd.Timestamp(camp["debut"]), pd.Timestamp(camp["fin"])
    nom_campagne = "_".join(["Display", "A79", slug(A), slug(B), slug(C), slug(D), debut.strftime("%d%m%Y")])

    lignes = []
    for f, a, msg in plan_de_lignes(df_formats, df_audiences, df_messages):
        Q, R, S = f["Platform"], c["audience"], c["geolocation"]
        Tt, U, V = a["Targeting Type"], a["Targeting details"], f["Format Type"]
        W, X, Y, Z, AA = f["Format size"], f["Tracking Type"], f["Régie"], C, c["complement"]

        placement = "_".join(slug(v) for v in (Q, R, S, Tt, U, V, W, X)) + f"_{slug(Y)}-{slug(Z)}_{slug(AA)}"
        morceaux = [D, str(f["Code régie"]), str(f["Code format"])]
        if str(W) != str(c["format_size"]):
            morceaux.append(str(W))          # plusieurs tailles d'un même format : la taille distingue la créa
        morceaux.append(str(a["Code créa"]))
        creative = "_".join(morceaux)
        if msg is not None:
            creative += "_" + str(msg["Code message"])

        landing = camp["landing"]
        utm = (f"utm_source={c['utm_source']}&utm_campaign={nom_campagne}"
               f"&utm_content={placement}&utm_term={creative}")
        sep = "&" if "?" in landing else "?"

        lignes.append({
            "Plateforme*": A, "Pilier": B, "Objectif": C, "Nom de la campagne (free form)": D,
            "Début  JJ/MM/AAAA": debut, "Fin  JJ/MM/AAAA": fin, "Nom de la campagne*": nom_campagne,
            "Message": msg["Message PDR"] if msg is not None else None,
            "code message": msg["Code message"] if msg is not None else None,
            "Lien": None, "Lien drive vers la crea": None,
            "Platform": Q, "Audience": R, "Geolocation": S, "Targeting Type": Tt,
            "Targeting details": U, "Format Type": V, "Format size (CM only)": W,
            "Tracking Type": X, "Régie": Y, "Stratégie": Z, "Complément (Free form)": AA,
            "XPLN script CM": None, "Placement CM = creative DV360": placement,
            "Creative Name": creative, "utm_source": c["utm_source"], "utm_campaign": nom_campagne,
            "utm_content": placement, "utm_term": creative, "URL": landing,
            "UTM": sep + utm, "URL de redirection trackée": landing + sep + utm,
        })
    return pd.DataFrame(lignes, columns=COLONNES)
