"""Chargement des tables de correspondance et fonctions de recherche.

Les comparaisons se font sur un texte normalisé (sans accents, en minuscules, espaces regroupés).
Un mot-clé correspond s'il apparaît en début de mot : « senior » trouve « Seniors »,
mais « aise » ne trouve pas « française ». Une sous-liste de mots-clés exige tous les mots.

Les entrées marquées « a_valider: true » sont des propositions : l'app les pré-remplit
avec le statut « proposé » pour que l'expert les vérifie.
"""
from __future__ import annotations

import re
import unicodedata
from functools import lru_cache
from pathlib import Path

import yaml

DOSSIER = Path(__file__).parent
VALIDE, PROPOSE, INCONNU = "validé", "proposé", "inconnu"


def normaliser(texte) -> str:
    if texte is None:
        return ""
    texte = unicodedata.normalize("NFKD", str(texte))
    texte = "".join(c for c in texte if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", texte).strip().lower()


def _motif(mot_cle) -> str:
    return r"(?<![a-z0-9])" + re.escape(normaliser(mot_cle))


def contient(texte, mot_cle) -> bool:
    """Vrai si le mot-clé (ou tous les mots d'une sous-liste) apparaît en début de mot."""
    t = normaliser(texte)
    mots = mot_cle if isinstance(mot_cle, list) else [mot_cle]
    return all(re.search(_motif(m), t) for m in mots)


def _retirer(texte: str, mots) -> str:
    t = normaliser(texte)
    for m in mots:
        for mm in (m if isinstance(m, list) else [m]):
            t = re.sub(_motif(mm), " ", t)
    return t


@lru_cache
def _charger(nom: str):
    with open(DOSSIER / f"{nom}.yaml", encoding="utf-8") as f:
        return yaml.safe_load(f)


# ---------------------------------------------------------------- listes
def listes(surcharge: dict[str, list[str]] | None = None) -> dict[str, list[str]]:
    """Listes déroulantes : celles du YAML, remplacées par celles d'un URL Builder si fournies."""
    res = {k: [str(v) for v in vals] for k, vals in _charger("listes").items()}
    for k, vals in (surcharge or {}).items():
        if k in res and vals:
            res[k] = vals
    return res


def listes_depuis_url_builder(fichier) -> dict[str, list[str]]:
    """Lit l'onglet « Inputs A79 » d'un URL Builder : {en-tête de colonne: [valeurs]}."""
    import openpyxl
    wb = openpyxl.load_workbook(fichier, read_only=True, data_only=True)
    if "Inputs A79" not in wb.sheetnames:
        raise ValueError("Onglet « Inputs A79 » introuvable dans ce fichier.")
    lignes = list(wb["Inputs A79"].iter_rows(values_only=True))
    res = {}
    for j, entete in enumerate(lignes[0] if lignes else []):
        if not entete:
            continue
        vals = []
        for r in lignes[1:]:
            v = r[j] if j < len(r) else None
            if v not in (None, "") and str(v).strip() not in vals:
                vals.append(str(v).strip())
        res[str(entete).strip()] = vals
    return res


def constantes() -> dict:
    return _charger("divers")["constantes"]


def valeur_de_liste(texte, liste: list[str]) -> str | None:
    """Retrouve une valeur de liste malgré accents, « _ » et majuscules (« Attractivité_commerciale »)."""
    cle = lambda v: re.sub(r"[^a-z0-9]", "", normaliser(v))
    for v in liste:
        if cle(v) == cle(texte):
            return v
    return None


# ---------------------------------------------------------------- régies
def trouver_regie(texte) -> str | None:
    """Renvoie la clé de régie (ex. 'france-tv') ou None."""
    for cle, info in _charger("regies").items():
        if any(contient(texte, a) for a in info["alias"]):
            return str(cle)
    return None


def info_regie(cle: str | None) -> dict:
    return _charger("regies").get(cle, {}) if cle else {}


# ---------------------------------------------------------------- formats
def trouver_formats(cle_regie: str | None, texte) -> list[dict]:
    """Tous les formats de la régie présents dans le texte (« Non skippable + Bumper » -> 2)."""
    entrees = _charger("formats").get(cle_regie, []) if cle_regie else []
    reste, trouves = normaliser(texte), []
    for e in [e for e in entrees if not e.get("defaut")]:
        if any(contient(reste, m) for m in e["mots_cles"]):
            trouves.append(e)
            reste = _retirer(reste, e["mots_cles"] + e.get("absorbe", []))
    if not trouves:
        for e in [e for e in entrees if e.get("defaut")]:
            if any(contient(reste, m) for m in e["mots_cles"]):
                trouves.append(e)
                break
    return trouves


def format_correspond_pdr(entree_format: dict, texte_pdr) -> bool:
    return any(contient(texte_pdr, m) for m in entree_format.get("pdr", []))


def duree(texte) -> str | None:
    """Première durée lue dans le texte : « 15 secondes », « 20 s », « 20'' » -> '15'."""
    m = re.search(r"(?<![a-z0-9])(\d{1,3})\s*(?:secondes?|sec\b|s\b|''|’’|\")", normaliser(texte))
    return m.group(1) if m else None


def tailles(texte) -> list[str]:
    """Tailles display lues dans le texte : ['300x250', '728x90']."""
    vues = []
    for t in re.findall(r"(?<!\d)(\d{2,4})\s*[x×]\s*(\d{2,4})(?!\d)", str(texte or "")):
        v = f"{t[0]}x{t[1]}"
        if v not in vues:
            vues.append(v)
    return vues


# ---------------------------------------------------------------- audiences
def trouver_audience(cle_regie: str | None, texte) -> dict | None:
    """Renvoie {targeting_type, targeting_details, code, statut} ou None si inconnue."""
    entrees = _charger("audiences")
    regie_de = lambda e: e.get("regie") if isinstance(e.get("regie"), list) else ([e["regie"]] if e.get("regie") else [])
    specifiques = [e for e in entrees if cle_regie in regie_de(e)]
    generales = [e for e in entrees if not regie_de(e)]
    t = normaliser(texte)
    for e in specifiques + generales:
        ok = any(contient(texte, m) for m in e.get("mots_cles", [])) or \
            any(t.startswith(normaliser(d)) for d in e.get("debut_par", []))
        if ok and any(contient(texte, m) for m in e.get("sauf", [])):
            ok = False
        if ok:
            res = {k: e[k] for k in ("targeting_type", "targeting_details", "code")}
            res.update(e.get("par_regie", {}).get(cle_regie, {}))
            res["statut"] = PROPOSE if e.get("a_valider") else VALIDE
            return res
    return None


def type_ciblage(texte) -> tuple[str, str] | None:
    """Type de ciblage nommé dans le libellé du PM : (Targeting Type, mot trouvé), ou None.
    Si plusieurs mots sont présents, celui qui apparaît en premier dans le libellé gagne."""
    t = normaliser(texte)
    meilleur = None
    for e in _charger("types_ciblage"):
        for m in e["mots_cles"]:
            r = re.search(_motif(m) + r"(?![a-z0-9])", t)
            if r and (meilleur is None or r.start() < meilleur[0]):
                meilleur = (r.start(), e["targeting_type"], m)
    return (meilleur[1], meilleur[2]) if meilleur else None


# ---------------------------------------------------------------- plateforme, objectif, messages
def plateforme(cle_regie: str | None, type_achat, dsp=None, plateformes: list[str] | None = None) -> str | None:
    """1. GAG -> plateforme de la régie (la régie diffuse elle-même) ;
    2. sinon, colonne DSP du PM si elle donne une plateforme connue (DV360, MPP…) ;
    3. sinon PG, CPM, PMP, private auction, curated, réservation -> DV360.
    La colonne DSP est souvent fusionnée sur tout le tableau : elle ne passe jamais avant le GAG."""
    connue = None
    if dsp:
        if contient(dsp, "dv360") or contient(dsp, "dv 360"):
            connue = "DV360"
        else:
            connue = valeur_de_liste(dsp, plateformes or listes()["Plateforme"])
    if contient(type_achat, "gag"):
        gag = info_regie(cle_regie).get("plateforme_gag")
        return gag or (connue if connue and connue != "DV360" else None)
    if connue:
        return connue
    if any(contient(type_achat, m) for m in ("pg", "cpm", "pmp", "private auction", "curated", "reservation")):
        return "DV360"
    return None


def objectif(texte_pdr) -> str | None:
    cle = normaliser(texte_pdr).replace("_", " ")
    return _charger("divers")["objectifs"].get(re.sub(r"\s+", " ", cle).strip())


def code_message(texte_pdr) -> str | None:
    return _charger("divers")["messages"].get(normaliser(texte_pdr))
