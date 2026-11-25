"""Test de référence du constructeur de PM : les exemples doivent redonner les fichiers réels.

Usage : python -m tests.test_pm_builder <URL_BUILDER_MDC_SDG_V2.xlsx> <Fichier_media_85_DIB_SEPTEMBRE.xlsx>
"""
import sys
from collections import Counter

import openpyxl
import pandas as pd

from pm_builder import controles, modele, vers_crf, vers_pmu


def carrefour(ub_ref):
    camp, l, c, a = modele.exemple_carrefour()
    assert not controles.verifier(camp, l, c, a)[0]
    gen, _ = vers_crf.generer(camp, l, c, a)
    ref = pd.read_excel(ub_ref, sheet_name="URL BUILDER", skiprows=1).dropna(subset=["Creative Name"])
    ref = ref[ref["Creative Name"].astype(str).str.startswith("MDC-SDG-V2")]
    cols = ["Platform", "Targeting Type", "Targeting details", "Format Type", "Tracking Type", "Régie", "Creative Name", "URL", "Nom de la campagne*"]
    k = lambda df: Counter(tuple(str(x[cc]).strip() for cc in cols) for _, x in df.iterrows())
    print(f"Carrefour : {len(gen)} lignes, {sum((k(ref) & k(gen)).values())}/{len(ref)} identiques à l'URL Builder")


def pmu(fm_ref):
    camp, l, c, a = modele.exemple_pmu()
    assert not controles.verifier(camp, l, c, a)[0]
    gen = vers_pmu.generer(camp, l, c)
    ws = openpyxl.load_workbook(fm_ref, data_only=True)["fichier_media"]
    cols = {"PLACEMENT NAME DCM": "K", "Créative DCM": "L", "Emplacement DSP": "Q", "Créa DSP": "R", "URL avec les utms": "W"}
    from pm_builder.pmu_validation import clean_url
    ref = lambda col, v: clean_url(str(v)) if col == "URL avec les utms" else str(v)   # URL : version nettoyée par l'app PMU
    ok = sum(str(r[col]).lower() == ref(col, ws[f"{let}{i}"].value).lower()
             for i, (_, r) in enumerate(gen.iterrows(), 14) for col, let in cols.items())
    print(f"PMU : {len(gen)} lignes, {ok}/{len(gen) * len(cols)} noms identiques au fichier média")


if __name__ == "__main__":
    carrefour(sys.argv[1]); pmu(sys.argv[2])
