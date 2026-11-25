"""Contrôle du fichier généré (URL Builder Carrefour ou fichier média PMU), règle par règle.

Chaque règle renvoie (titre, ok, détails). Sert à afficher « voici ce qui a été généré » avant le téléchargement.
"""
from __future__ import annotations

import pandas as pd

import tables as T
from export.cm import ALLOWED_PLATFORMS, map_tracking_type
from pm_builder import controles
from pm_builder.modele import LP_ID, LP_URLS


def _vides(df, cols):
    return [c for c in cols if df[c].isna().any() or (df[c].astype(str).str.strip() == "").any()]


def _regle(titre, pb):
    return (titre, not pb, pb)


def url_builder(df: pd.DataFrame, camp, lignes, creas, audiences) -> list:
    L = T.listes()
    attendu = controles.nb_lignes_sortie(camp, lignes, creas, audiences)
    res = [_regle("Nombre de lignes conforme au PM", [] if len(df) == attendu else [f"{len(df)} lignes générées pour {attendu} attendues."])]
    vides = _vides(df, ["Platform", "Régie", "Format Type", "Targeting Type", "Targeting details", "Tracking Type",
                        "Creative Name", "Placement CM = creative DV360", "URL de redirection trackée"])
    res.append(_regle("Aucune case obligatoire vide", [f"Colonne « {c} » incomplète." for c in vides]))
    hors = []
    for col, liste in (("Platform", "Plateforme"), ("Régie", "Régie"), ("Format Type", "Format Type"),
                       ("Targeting Type", "Targeting Type"), ("Targeting details", "Targeting details"), ("Tracking Type", "Tracking Type")):
        hors += [f"{col} « {v} » absent de la liste." for v in sorted(set(df[col].dropna().astype(str)) - set(L[liste]))]
    res.append(_regle("Toutes les valeurs sont dans les listes déroulantes", hors))
    doublons = df["Creative Name"][df["Creative Name"].duplicated()].unique().tolist()
    res.append(_regle("Noms de créa uniques", [f"« {d} » apparaît plusieurs fois." for d in doublons[:5]]))
    morceaux = df["utm_content"].astype(str).str.split("_").str.len()
    res.append(_regle("Nom de placement en 10 morceaux (exploitable dans BigQuery)",
                      [f"{n} ligne(s) n'ont pas 10 morceaux." for n in [int((morceaux != 10).sum())] if n]))
    sans_utm = df[~df["URL de redirection trackée"].astype(str).str.contains("utm_source=.*utm_campaign=.*utm_content=.*utm_term=", regex=True)]
    mauvaise = df[~df["URL de redirection trackée"].astype(str).str.startswith("http")]
    res.append(_regle("URL trackées complètes (http + 4 UTM)", ([f"{len(sans_utm)} URL sans les 4 UTM."] if len(sans_utm) else []) +
                      ([f"{len(mauvaise)} URL ne commencent pas par http."] if len(mauvaise) else [])))
    refus = sorted({p for p in df["Platform"].astype(str) if p not in ALLOWED_PLATFORMS}) + \
        sorted({t for t in df["Tracking Type"].astype(str) if map_tracking_type(t) == "OTHER"})
    res.append(_regle("Accepté par l'export CM", [f"« {x} » refusé par l'export CM." for x in refus]))
    return res


def fichier_media(df: pd.DataFrame, lev: str, camp, lignes, creas) -> list:
    from pm_builder.modele import LEVIER_PAR_SECTION
    lev_l = lignes["Section"].map(lambda x: LEVIER_PAR_SECTION.get(str(x))).fillna(camp.get("levier"))
    ids = set(lignes.loc[lev_l == lev, "N° ligne"].astype(str))
    attendu = int(creas["N° ligne"].astype(str).isin(ids).sum())
    res = [_regle("Nombre de lignes conforme aux créas", [] if len(df) == attendu else [f"{len(df)} lignes générées pour {attendu} tailles attendues."])]
    vides = _vides(df, ["Régie", "Programmation", "Support", "Emplacement / Ciblage", "Device", "Type de Tracking",
                        "Type de Format", "Format", "Piste Créative", "Landing Page", "Date de début", "Date de fin"])
    res.append(_regle("Aucune case obligatoire vide", [f"Colonne « {c} » incomplète." for c in vides]))
    n8 = df["PLACEMENT NAME DCM"].astype(str).str.split("_").str.len()
    res.append(_regle("Nom de placement en 8 morceaux", [f"{n} ligne(s) mal découpées (un « _ » dans un champ ?)." for n in [int((n8 != 8).sum())] if n]))
    doublons = df[df.duplicated(["PLACEMENT NAME DCM", "Créative DCM"])]["PLACEMENT NAME DCM"].unique().tolist()
    res.append(_regle("Couples placement / créa uniques", [f"« {d} » en double." for d in doublons[:5]]))
    from pm_builder.modele import SITE_NOMS
    hors_sites = sorted(set(df["Support"].astype(str)) - set(SITE_NOMS))
    res.append(_regle("Supports dans la liste des sites CM PMU", [f"« {x} » n'est pas un site CM." for x in hors_sites]))
    hors_lp = sorted(set(df["Landing Page"].astype(str)) - set(LP_URLS))
    res.append(_regle("Landing pages de la liste CM PMU", [f"« {u} » n'est pas dans la liste." for u in hors_lp]))
    url = df["URL avec les utms"].astype(str)
    pb = [f"{int(n)} URL incomplète(s)." for n in [(~url.str.contains("utm_medium=.*utm_source=.*utm_content=.*utm_campaign=.*utm_term=", regex=True)).sum()] if n]
    pb += [f"{int(n)} URL avec des espaces." for n in [url.str.contains(" ").sum()] if n]
    res.append(_regle("URL avec les 5 UTM, sans espace", pb))
    return res


def recap_regies(df: pd.DataFrame, col_regie="Régie") -> pd.DataFrame:
    return df.groupby(col_regie).size().rename("Lignes générées").reset_index()


def landing_ids(df: pd.DataFrame) -> pd.DataFrame:
    lp = df["Landing Page"].astype(str).value_counts().rename("Lignes").reset_index()
    lp.columns = ["Landing page", "Lignes"]
    lp.insert(1, "Landing Page ID", lp["Landing page"].map(LP_ID))
    return lp
