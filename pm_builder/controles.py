"""Contrôles du PM commun. Une erreur bloque la génération, un avertissement informe."""
from __future__ import annotations

import re
import unicodedata

import pandas as pd

from pm_builder.modele import listes


def _v(x) -> str:
    return "" if x is None or (isinstance(x, float) and pd.isna(x)) else str(x).strip()


def _taille_ok(t: str) -> bool:
    return re.fullmatch(r"\d{1,4}\s*[xX×]\s*\d{1,4}", t.strip()) is not None


def _accents(t: str) -> bool:
    return any(unicodedata.combining(c) for c in unicodedata.normalize("NFKD", t))


def verifier(camp: dict, lignes: pd.DataFrame, creas: pd.DataFrame, audiences: pd.DataFrame):
    client = camp["client"]
    L = listes()[client]
    err, avert = [], []

    # ---------------- en-tête
    oblig = {"nom": "Nom de campagne", "debut": "Début", "fin": "Fin", "landing": "Landing page"}
    oblig.update({"pilier": "Pilier", "objectif": "Objectif"} if client == "Carrefour" else
                 {"levier": "Levier campagne", "hs": "H/S", "univers": "85/2010", "medium": "Medium", "promo": "Promo/Offre"})
    for k, nom in oblig.items():
        if not _v(camp.get(k)):
            err.append(f"Campagne : {nom} manquant.")
    if camp.get("debut") and camp.get("fin") and camp["fin"] < camp["debut"]:
        err.append("Campagne : la date de fin est avant la date de début.")
    if _v(camp.get("landing")) and not _v(camp["landing"]).startswith("http"):
        err.append("Campagne : la landing page doit commencer par http.")
    if client == "PMU" and _v(camp.get("landing")):
        from pm_builder.modele import LP_URLS
        if _v(camp["landing"]) not in LP_URLS:
            err.append("Campagne : la landing page doit être choisie dans la liste des LP PMU.")
    if client == "PMU":
        from pm_builder.pmu_validation import EMAIL_RE
        if not _v(camp.get("consultant")):
            err.append("Campagne : Consultant email manquant (obligatoire pour l'app automation PMU).")
        elif not EMAIL_RE.match(_v(camp["consultant"])):
            err.append(f"Campagne : Consultant email invalide (« {camp['consultant']} »).")
    if client == "Carrefour" and " " in _v(camp.get("nom")):
        avert.append("Campagne : le nom contient des espaces. Utilisez exactement le nom du PDR (ex. MDC_Sdg_V2).")

    # ---------------- lignes
    if lignes.empty:
        err.append("Aucune ligne dans le PM.")
    cols_liste = {"Levier": "Levier", "Régie": "Régie", "Type d'achat": "Type d'achat", "Type de format": "Type de format"}
    if client == "Carrefour":
        cols_liste["Tracking Type"] = "Tracking Type"
        oblig_l = ["Régie", "Type d'achat", "Format", "Type de format", "Groupe d'audiences"]
    else:
        cols_liste.update({"Device": "Device", "Programmation": "Programmation", "Type de tracking": "Type de tracking"})
        oblig_l = ["Levier", "Régie", "Type d'achat", "Type de format", "Ciblage", "Device", "Programmation", "Support", "Type de tracking"]
    for _, l in lignes.iterrows():
        q = f"Ligne {l['N° ligne']} ({_v(l['Régie']) or 'régie ?'} · {_v(l['Format'])[:40]})"
        for col in oblig_l:
            if not _v(l.get(col)):
                err.append(f"{q} : {col} à remplir.")
        for col, nom in cols_liste.items():
            if _v(l.get(col)) and _v(l[col]) not in L[nom]:
                err.append(f"{q} : {col} « {l[col]} » absent de la liste.")
        b, p = pd.to_numeric(l.get("Budget net (€)"), errors="coerce"), pd.to_numeric(l.get("CPM (€)"), errors="coerce")
        if pd.isna(b) or pd.isna(p) or b <= 0 or p <= 0:
            avert.append(f"{q} : budget ou CPM manquant, le volume ne peut pas être calculé.")
        if l.get("Début") is not None and l.get("Fin") is not None and not pd.isna(l["Début"]) and not pd.isna(l["Fin"]) and l["Fin"] < l["Début"]:
            err.append(f"{q} : la date de fin est avant la date de début.")
        n_creas = int((creas["N° ligne"] == l["N° ligne"]).sum())
        if client == "PMU":
            if n_creas == 0:
                err.append(f"{q} : aucune taille dans l'onglet Créas, aucune ligne ne sera générée.")
            from pm_builder.modele import SITE_NOMS
            if _v(l.get("Support")) and _v(l["Support"]) not in SITE_NOMS:
                err.append(f"{q} : Support « {l['Support']} » absent de la liste des sites CM PMU.")
            for col in ("Régie", "Device", "Ciblage", "Type de format", "Type de tracking"):
                if "_" in _v(l.get(col)):
                    err.append(f"{q} : « _ » interdit dans {col} (c'est le séparateur des noms).")
            if _accents(_v(l.get("Ciblage"))):
                avert.append(f"{q} : le ciblage contient des accents, ils apparaîtront tels quels dans les noms CM.")
        else:
            g = _v(l.get("Groupe d'audiences"))
            if g and not (audiences["Groupe d'audiences"].astype(str) == g).any():
                err.append(f"{q} : le groupe d'audiences « {g} » n'a aucune audience dans l'onglet Audiences.")
            nb = pd.to_numeric(l.get("Nb messages"), errors="coerce")
            nb = 1 if pd.isna(nb) else int(nb)
            n_msg = int(((creas["N° ligne"] == l["N° ligne"]) & creas["Code message"].astype(str).str.strip().ne("") & creas["Code message"].notna()).sum())
            if nb > 1 and n_msg != nb:
                err.append(f"{q} : {nb} messages annoncés mais {n_msg} code(s) message dans l'onglet Créas.")

    # ---------------- créas
    ids = set(lignes["N° ligne"].astype(str))
    for i, c in creas.iterrows():
        q = f"Créa n°{i + 1} ({_v(c['N° ligne'])} · {_v(c['Nom du format'])})"
        if _v(c["N° ligne"]) not in ids:
            err.append(f"{q} : N° de ligne inconnu dans le PM.")
        if client == "PMU":
            if not _taille_ok(_v(c["Taille"])):
                err.append(f"{q} : taille « {_v(c['Taille'])} » à écrire comme 300x250.")
            from pm_builder.modele import LP_URLS
            if _v(c.get("Landing (si différente)")) and _v(c["Landing (si différente)"]) not in LP_URLS:
                err.append(f"{q} : landing hors de la liste des LP PMU.")
            if not _v(c["Piste créative / message"]):
                err.append(f"{q} : piste créative à remplir (elle entre dans le nom de créa).")
        if _v(c["Code message"]) and ("_" in _v(c["Code message"]) or " " in _v(c["Code message"])):
            err.append(f"{q} : le code message ne doit contenir ni « _ » ni espace.")

    # ---------------- audiences (Carrefour)
    if client == "Carrefour":
        for i, a in audiences.iterrows():
            q = f"Audience « {_v(a['Audience'])} » (groupe {_v(a['Groupe d\'audiences'])})"
            for col, nom in (("Type de ciblage", "Type de ciblage"), ("Targeting details", "Targeting details")):
                if not _v(a.get(col)):
                    err.append(f"{q} : {col} à remplir.")
                elif _v(a[col]) not in L[nom]:
                    err.append(f"{q} : {col} « {a[col]} » absent de la liste.")
        for g, grp in audiences.groupby("Groupe d'audiences"):
            s = pd.to_numeric(grp["% du groupe"], errors="coerce").sum()
            if abs(s - 100) > 0.5:
                err.append(f"Groupe d'audiences « {g} » : les % font {s:g} % au lieu de 100 %.")
            if not (lignes["Groupe d'audiences"].astype(str) == str(g)).any():
                avert.append(f"Groupe d'audiences « {g} » : rattaché à aucune ligne du PM.")
    return err, avert


def nb_lignes_sortie(camp, lignes, creas, audiences) -> int:
    if camp["client"] == "PMU":
        return int(creas["N° ligne"].isin(lignes["N° ligne"]).sum())
    total = 0
    for _, l in lignes.iterrows():
        n_aud = int((audiences["Groupe d'audiences"].astype(str) == str(l["Groupe d'audiences"])).sum())
        nb = pd.to_numeric(l.get("Nb messages"), errors="coerce")
        total += n_aud * (1 if pd.isna(nb) else max(int(nb), 1))
    return total
