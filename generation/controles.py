"""Contrôles avant génération. Une erreur bloque, un avertissement informe."""
from __future__ import annotations

import pandas as pd

import tables as T
from export.cm import display_keywords, video_keywords
from generation.url_builder import nombre_attendu, plan_de_lignes

TRACKING_ACCEPTES_CM = set(video_keywords) | set(display_keywords) | {"Stream-Audio"}


def _vide(v) -> bool:
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == ""


def _code_invalide(v) -> str | None:
    if _vide(v):
        return "à compléter"
    s = str(v)
    if "{" in s:
        return "contient « {duree} » : durée introuvable dans le PM, complétez-la"
    if "_" in s or " " in s:
        return "ne doit contenir ni « _ » ni espace"
    return None


def verifier(camp: dict, df_formats, df_audiences, df_messages, constantes: dict | None = None,
             listes: dict | None = None):
    L = listes or T.listes()
    erreurs, avertissements = [], []

    for champ, nom in (("nom", "Nom de campagne"), ("pilier", "Pilier"), ("objectif", "Objectif"),
                       ("debut", "Date de début"), ("fin", "Date de fin"), ("landing", "Landing page")):
        if _vide(camp.get(champ)):
            erreurs.append(f"Campagne : {nom} manquant.")
    if "_" in str(camp.get("nom", "")):
        erreurs.append("Campagne : le nom ne doit pas contenir « _ » (c'est le séparateur).")
    for champ, liste in (("pilier", "Pilier"), ("objectif", "Objectif")):
        if camp.get(champ) and camp[champ] not in L[liste]:
            erreurs.append(f"Campagne : {champ} « {camp[champ]} » absent de la liste.")

    if df_formats.empty:
        erreurs.append("Aucun format : aucune ligne ne sera créée.")

    for _, f in df_formats.iterrows():
        q = f"Format « {f['Régie PM']} · {str(f['Format PM'])[:60]} »"
        for col, liste in (("Platform", "Plateforme"), ("Régie", "Régie"), ("Format Type", "Format Type"),
                           ("Format size", "Format Size"), ("Tracking Type", "Tracking Type")):
            if _vide(f[col]):
                aide = " Si la valeur n'existe pas, ajoutez-la dans l'onglet Inputs A79." if col == "Régie" else ""
                erreurs.append(f"{q} : {col} à compléter.{aide}")
            elif str(f[col]) not in L[liste]:
                erreurs.append(f"{q} : {col} « {f[col]} » absent de la liste.")
        if not _vide(f["Tracking Type"]) and str(f["Tracking Type"]) not in TRACKING_ACCEPTES_CM:
            erreurs.append(f"{q} : Tracking Type « {f['Tracking Type']} » refusé par l'export CM "
                           "(il n'est ni dans la liste vidéo ni dans la liste display du code CM).")
        for col in ("Code régie", "Code format"):
            pb = _code_invalide(f[col])
            if pb:
                erreurs.append(f"{q} : {col} {pb}.")
        if T.contient(f["Type d'achat"], "gag") and f["Platform"] == "DV360":
            avertissements.append(f"{q} : achat en GAG mais Platform = DV360. À vérifier.")
        if not (df_audiences["Bloc"] == f["Bloc"]).any():
            erreurs.append(f"{q} : aucune audience dans son bloc « {f['Bloc']} », aucune ligne ne sera créée.")

    for _, a in df_audiences.iterrows():
        q = f"Audience « {a['Régie PM']} · {str(a['Audience PM'])[:60]} »"
        for col in ("Targeting Type", "Targeting details"):
            if _vide(a[col]):
                erreurs.append(f"{q} : {col} à compléter.")
            elif str(a[col]) not in L[col]:
                erreurs.append(f"{q} : {col} « {a[col]} » absent de la liste.")
        pb = _code_invalide(a["Code créa"])
        if pb:
            erreurs.append(f"{q} : code créa {pb}.")
        if not (df_formats["Bloc"] == a["Bloc"]).any():
            avertissements.append(f"{q} : aucun format dans son bloc « {a['Bloc']} », elle ne créera aucune ligne.")

    for bloc, groupe in df_audiences.groupby("Bloc"):
        pcts = pd.to_numeric(groupe["% budget"], errors="coerce").dropna()
        if len(pcts) == len(groupe) and len(pcts) and abs(pcts.sum() - 100) > 0.5:
            avertissements.append(f"Audiences du bloc « {bloc} » : les % font {pcts.sum():g} % au lieu de 100 %. "
                                  "Une audience a peut-être été mal lue.")

    for _, m in df_messages.iterrows():
        if _vide(m["Code message"]):
            erreurs.append(f"Message « {m['Message PDR']} » ({m['Régie PM']}) : code message à compléter.")

    for _, f in df_formats[df_formats["Une ligne par message"].astype(bool)].iterrows():
        msgs = df_messages[(df_messages["Bloc"] == f["Bloc"]) & (df_messages["Code format"] == f["Code format"])]
        if msgs.empty:
            erreurs.append(
                f"Format « {f['Régie PM']} · {str(f['Format PM'])[:60]} » : une ligne par message est demandée, "
                "mais aucun message n'a été trouvé dans le PDR pour ce format. Vérifiez la campagne choisie et le PDR, "
                "ajoutez les messages dans le tableau Messages, ou décochez « Une ligne par message ».")

    proposes = int((df_formats["Statut"] == T.PROPOSE).sum() + (df_audiences["Statut"] == T.PROPOSE).sum())
    if proposes:
        avertissements.append(f"{proposes} ligne(s) au statut « proposé » : ces valeurs viennent de propositions "
                              "pas encore confirmées. Vérifiez-les avant de générer.")

    attendu = nombre_attendu(df_formats, df_audiences, df_messages)
    obtenu = len(plan_de_lignes(df_formats, df_audiences, df_messages))
    if attendu != obtenu:
        erreurs.append(f"Nombre de lignes incohérent : {obtenu} générées pour {attendu} attendues.")

    return erreurs, avertissements, attendu
