"""Test de référence : le générateur doit reproduire l'URL Builder rempli à la main pour MDC SDG V2.

Usage : python -m tests.test_mdc <PM.xlsx> <PDR.xlsx> <URL_BUILDER_reference.xlsx>
Compare les colonnes de saisie (Q à AA, Creative Name, URL) ligne par ligne, sans tenir compte de l'ordre.
Les colonnes assemblées (Placement, utm_content) diffèrent volontairement : nouvelle règle de nommage.
"""
import sys
from collections import Counter

import pandas as pd

from generation import setup, url_builder, controles
from lecture.pdr import filtrer, lire_pdr
from lecture.pm import lire_pm

COMPAREES = ["Platform", "Audience", "Geolocation", "Targeting Type", "Targeting details", "Format Type",
             "Format size (CM only)", "Tracking Type", "Régie", "Stratégie", "Complément (Free form)",
             "Creative Name", "URL", "Nom de la campagne*"]


def main(pm_path, pdr_path, ref_path, campagne="MDC_Sdg_V2"):
    pm = lire_pm(pm_path)
    pdr = filtrer(lire_pdr(pdr_path), campagne)
    camp = setup.campagne(pm, pdr, campagne)
    f, a = setup.formats(pm), setup.audiences(pm)
    m = setup.messages(f, pdr)
    erreurs, avert, attendu = controles.verifier(camp, f, a, m)
    print(f"Formats lus : {len(f)} {f['Statut'].value_counts().to_dict()}")
    print(f"Audiences lues : {len(a)} {a['Statut'].value_counts().to_dict()}")
    print(f"Messages : {len(m)} | Lignes attendues : {attendu}")
    print("Erreurs :", erreurs or "aucune")
    print("Avertissements :", avert or "aucun")

    gen = url_builder.generer(camp, f, a, m)
    ref = pd.read_excel(ref_path, sheet_name="URL BUILDER", skiprows=1).dropna(subset=["Creative Name"])
    ref = ref[ref["Creative Name"].astype(str).str.startswith(camp["nom"])]
    print(f"\nLignes générées : {len(gen)} | lignes de référence : {len(ref)}")

    cle = lambda df: Counter(tuple(str(r[c]).strip() for c in COMPAREES) for _, r in df.iterrows())
    cg, cr = cle(gen), cle(ref)
    manquantes, en_trop = cr - cg, cg - cr
    if not manquantes and not en_trop:
        print("✅ Identique à la référence sur les colonnes de saisie.")
    else:
        print(f"❌ {sum(manquantes.values())} ligne(s) attendue(s) absente(s), {sum(en_trop.values())} en trop.")
        for t in list(manquantes.elements())[:3]:
            print("   attendue :", dict(zip(COMPAREES, t)))
        for t in list(en_trop.elements())[:3]:
            print("   générée  :", dict(zip(COMPAREES, t)))
    return gen


if __name__ == "__main__":
    main(*sys.argv[1:4])
