"""Compare l'app à des URL Builders remplis à la main, sur plusieurs campagnes.

Usage : python -m tests.test_references <PDR global.xlsx> <PM.xlsx>::<onglet>::<URL_BUILDER.xlsx> ...
Mesure, par campagne :
- le nombre de lignes (app / référence) ;
- l'accord sur les lignes d'achat : (Platform, Régie, Format Type) ;
- l'accord sur les audiences : (Régie, Targeting Type, Targeting details) ;
- l'accord complet sur (Platform, Régie, Format Type, Targeting Type, Targeting details).
Les noms de créa ne sont pas comparés : leur écriture varie d'une campagne à l'autre.
"""
import difflib
import io
import re
import sys
import contextlib
from collections import Counter

import openpyxl

import tables as T
from generation import setup, url_builder
from lecture.pdr import campagnes, filtrer, lire_pdr
from lecture.pm import lire_pm


def lire_url_builder(chemin, max_lignes=1500):
    wb = openpyxl.load_workbook(chemin, read_only=True, data_only=True)
    lignes, entete = [], None
    for i, r in enumerate(wb["URL BUILDER"].iter_rows(min_row=2, max_row=max_lignes, values_only=True)):
        if i == 0:
            entete = [str(h).strip() if h else f"col{j}" for j, h in enumerate(r)]
            continue
        d = dict(zip(entete, r))
        if d.get("Creative Name") and d.get("Platform") and not str(d["Creative Name"]).startswith("DV4"):
            lignes.append({k: ("" if v is None else str(v).strip()) for k, v in d.items()})
    return lignes


def _accord(gen, ref, cols):
    cg = Counter(tuple(str(r[c]) for c in cols) for r in gen)
    cr = Counter(tuple(str(r[c]) for c in cols) for r in ref)
    commun = sum((cg & cr).values())
    return commun / max(sum(cg.values()), sum(cr.values()), 1), cr - cg, cg - cr


def comparer(pm_path, onglet, ub_path, pdr_tout):
    with contextlib.redirect_stderr(io.StringIO()):
        pm = lire_pm(pm_path, onglet or None)
    cle = lambda n: re.sub(r"[^a-z0-9]", "", T.normaliser(n))
    proches = difflib.get_close_matches(cle(pm.get("campagne") or ""), [cle(c) for c in campagnes(pdr_tout)], n=1, cutoff=0.6)
    nom_pdr = next((c for c in campagnes(pdr_tout) if proches and cle(c) == proches[0]), None)
    pdr = filtrer(pdr_tout, nom_pdr) if nom_pdr else []
    camp = setup.campagne(pm, pdr, nom_pdr or "X")
    f, a = setup.formats(pm), setup.audiences(pm)
    m = setup.messages(f, pdr)
    camp.update({"objectif": camp["objectif"] or "Branding", "pilier": camp["pilier"] or "Marque",
                 "landing": camp["landing"] or "https://www.carrefour.fr/", "debut": camp["debut"] or "2026-01-01"})
    gen = url_builder.generer(camp, f, a, m).to_dict("records")
    ref = lire_url_builder(ub_path)
    res = {"campagne": pm.get("campagne"), "pdr": nom_pdr, "n_app": len(gen), "n_ref": len(ref)}
    for nom, cols in (("achat", ["Platform", "Régie", "Format Type"]),
                      ("audience", ["Régie", "Targeting Type", "Targeting details"]),
                      ("complet", ["Platform", "Régie", "Format Type", "Targeting Type", "Targeting details"])):
        res[nom], res[nom + "_manque"], res[nom + "_trop"] = _accord(gen, ref, cols)
    return res


if __name__ == "__main__":
    pdr_tout = lire_pdr(sys.argv[1])
    for arg in sys.argv[2:]:
        pm_path, onglet, ub_path = arg.split("::")
        r = comparer(pm_path, onglet, ub_path, pdr_tout)
        print(f"{r['campagne']!s:22} PDR={r['pdr']!s:14} lignes app/réf = {r['n_app']:3}/{r['n_ref']:3}  "
              f"achat {r['achat']:.0%}  audiences {r['audience']:.0%}  complet {r['complet']:.0%}")
