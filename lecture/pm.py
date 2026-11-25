"""Lecture de l'onglet display du plan média.

Produit :
- formats : une entrée par morceau de cellule format (le découpage fin « Non skippable + Bumper »
  est fait ensuite avec les tables, dans generation/setup.py) ;
- audiences : une entrée par audience, avec le % de budget s'il est indiqué.

Chaque ligne du PM qui a des formats forme un « bloc » : ses formats et ses audiences vont ensemble.
Une ligne sans audiences reprend celles du bloc précédent de la même régie (ex. ligne Shorts),
et une ligne avec des audiences mais sans format complète le bloc précédent.
"""
from __future__ import annotations

import re
from datetime import datetime

import openpyxl

from tables import normaliser

ENTETES = {
    "debut": ("debut", "date debut"),
    "fin": ("fin", "date fin"),
    "dsp": ("dsp/regie", "dsp"),
    "regie": ("regie/partenaire",),
    "achat": ("type achat",),
    "formats": ("formats", "format"),
    "aud": ("audiences", "ciblages", "audience", "ciblage"),
    "budget": ("budget net media ht", "budget net media", "net media"),
    "cpm": ("cpm regie", "cpm"),
}
EXACT_D_ABORD = ("budget",)   # « budget net media ht » existe aussi en version « (+ frais) » : on préfère l'en-tête exact
LIGNES_IGNOREES = ("exemple", "proposition", "total", "specs", "spotify - spec")


def _correspond(nom: str, alias: tuple) -> bool:
    return any(nom == a or nom.startswith(a + " ") for a in alias)


def _colonnes(cellules) -> dict:
    cols = {}
    for cle in EXACT_D_ABORD:
        for c in cellules:
            if normaliser(c.value) in ENTETES[cle]:
                cols[cle] = c.column
                break
    for c in cellules:
        n = normaliser(c.value)
        for cle, alias in ENTETES.items():
            if cle not in cols and n and _correspond(n, alias):
                cols[cle] = c.column
    return cols


def _est_entete(ligne_cellules) -> bool:
    return any(normaliser(c.value) == "regie/partenaire" for c in ligne_cellules)


def onglets_display(fichier) -> list[str]:
    """Onglets « display » qui contiennent un tableau avec un en-tête REGIE/PARTENAIRE."""
    wb = openpyxl.load_workbook(fichier, read_only=True, data_only=True)
    res = []
    for ws in wb:
        if "disp" not in normaliser(ws.title):
            continue
        for row in ws.iter_rows(min_row=1, max_row=60, values_only=True):
            if any(normaliser(v) == "regie/partenaire" for v in row):
                res.append(ws.title)
                break
    res.sort(key=lambda t: 0 if normaliser(t) == "pm display" else 1)
    return res


def _nettoyer(texte) -> str:
    texte = re.sub(r"\bnouveau\b", "", str(texte or ""), flags=re.I)
    return re.sub(r"\s+", " ", texte).strip(" +&->")


def _premiere_ligne(texte) -> str:
    for l in str(texte or "").split("\n"):
        if l.strip():
            return l.strip()
    return ""


def decouper_formats(cellule) -> list[str]:
    """Coupe sur un « + » en début de ligne. Les autres cas (« A + B » sur une ligne, « A & B »)
    sont découpés ensuite grâce aux tables (plusieurs formats reconnus dans le même morceau)."""
    res = []
    for m in re.split(r"\n\s*\+", str(cellule or "")):
        m = _nettoyer(m)
        if not m:
            continue
        if res and normaliser(m).startswith("companion"):
            res[-1] += " + " + m     # le companion accompagne le format précédent (pas de ligne à part)
        else:
            res.append(m)
    return res


_PCT = re.compile(r"(\d+(?:[.,]\d+)?)\s*%")
_PCT_TEXTE = re.compile(r"[\(:\-–]?\s*\d+(?:[.,]\d+)?\s*%+\s*(?:du budget)?\s*[\)°]?", re.I)


def _nettoyer_audience(texte: str) -> tuple[str, float | None]:
    texte = re.split(r">\s*(?:deal id|creative id)", texte, flags=re.I)[0]
    m = _PCT.search(texte)
    pct = float(m.group(1).replace(",", ".")) if m else None
    libelle = _PCT_TEXTE.sub(" ", texte)
    libelle = re.sub(r"\s+", " ", libelle).strip(" :-–>•")
    return libelle, pct


def decouper_audiences(cellule) -> list[tuple[str, float | None]]:
    """Renvoie [(libellé, % de budget ou None)].

    - Si des lignes commencent par « > », chaque puce est une audience.
    - Sinon, si la cellule a des blocs séparés par une ligne vide, de 1 ou 2 lignes chacun
      (titre + description), chaque bloc donne une audience (sa 1re ligne).
    - Sinon, une audience par ligne : les lignes avec un % si certaines en ont, et la 1re ligne
      est retirée si c'est un titre sans chiffre (ex. « Socio demo » au-dessus des âges Spotify).
    """
    lignes = [l.strip() for l in str(cellule or "").split("\n")]
    non_vides = [l for l in lignes if l]
    if not non_vides:
        return []
    puces = [l for l in non_vides if l.startswith(">")]
    if puces:
        # une puce + ses lignes de détail sans puce ; si la puce finit par « : », le détail complète le libellé
        items = []
        for l in non_vides:
            if l.startswith(">"):
                items.append(l.lstrip(">").strip())
            elif items and items[-1].rstrip().endswith(":"):
                items[-1] = items[-1].rstrip() + " " + l
            elif items and "%" in l and "%" not in items[-1]:
                items[-1] = items[-1] + " " + l[l.rfind(" ", 0, l.find("%")) + 1:]
        # une puce sans % suivie d'une puce avec % partage son budget : une seule ligne d'achat
        if any("%" in i for i in items):
            fusion = []
            for i in items:
                if fusion and "%" not in fusion[-1] and not fusion[-1].endswith(" /"):
                    fusion[-1] = fusion[-1] + " / " + i
                else:
                    fusion.append(i)
            items = fusion
    else:
        blocs, courant = [], []
        for l in lignes:
            if l:
                courant.append(l)
            elif courant:
                blocs.append(courant)
                courant = []
        if courant:
            blocs.append(courant)
        if len(blocs) > 1 and any(len(b) > 1 for b in blocs) and all(len(b) <= 2 for b in blocs):
            items = [b[0] for b in blocs]
        else:
            items = non_vides
            avec_pct = [l for l in items if "%" in l]
            if avec_pct:
                items = avec_pct
            elif len(items) > 1 and not re.search(r"\d", items[0]) and \
                    (items[0].rstrip().endswith(":") or len(items[0].split()) <= 3):
                items = items[1:]   # titre court au-dessus des audiences (ex. « Socio demo » chez Spotify)
    items = [i for i in items if not i.startswith("(")]
    items = [i for i in items if not normaliser(i).startswith("exclusion")]   # une exclusion n'est pas une audience
    if len(items) > 1:
        items = [i for i in items if not i.rstrip().endswith(":")]
    res = []
    for i in items:
        libelle, pct = _nettoyer_audience(i)
        if libelle:
            res.append((libelle, pct))
    return res


def _fusions(ws) -> dict:
    """{(ligne, colonne): (valeur de la cellule en haut à gauche, est_cette_cellule_en_haut_a_gauche)}"""
    res = {}
    for plage in ws.merged_cells.ranges:
        v = ws.cell(row=plage.min_row, column=plage.min_col).value
        for r in range(plage.min_row, plage.max_row + 1):
            for c in range(plage.min_col, plage.max_col + 1):
                res[(r, c)] = (v, r == plage.min_row and c == plage.min_col)
    return res


def lire_pm(fichier, onglet: str | None = None) -> dict:
    wb = openpyxl.load_workbook(fichier, data_only=True)
    if onglet is None:
        candidats = onglets_display(fichier)
        if not candidats:
            raise ValueError("Aucun onglet display avec un en-tête « REGIE/PARTENAIRE » dans ce PM.")
        onglet = candidats[0]
    ws = wb[onglet]

    ligne_entete, cols = None, {}
    for row in ws.iter_rows(min_row=1, max_row=60):
        if _est_entete(row):
            ligne_entete, cols = row[0].row, _colonnes(row)
            break
    if not ligne_entete:
        raise ValueError(f"En-tête « REGIE/PARTENAIRE » introuvable dans l'onglet « {onglet} ».")
    for cle in ("regie", "formats"):
        if cle not in cols:
            raise ValueError(f"Colonne « {ENTETES[cle][0]} » introuvable dans l'onglet « {onglet} ».")

    nom_campagne = None
    for row in ws.iter_rows(min_row=1, max_row=ligne_entete):
        for cell in row:
            if normaliser(cell.value) == "campagne":
                voisin = ws.cell(row=cell.row, column=cell.column + 1).value
                nom_campagne = str(voisin).strip() if voisin else None
                break
        if nom_campagne:
            break

    fusions = _fusions(ws)
    dans_tableau = True
    formats, audiences = [], []
    debut = fin = None
    regie = achat = dsp = None
    bloc = None                     # bloc courant
    dernier_bloc_regie: dict[str, str] = {}

    for r in range(ligne_entete + 1, ws.max_row + 1):
        cellules = [ws.cell(row=r, column=c) for c in range(1, ws.max_column + 1)]
        if _est_entete(cellules):   # 2e tableau dans le même onglet (ex. bloc trafic)
            cols = _colonnes(cellules)
            regie = achat = dsp = bloc = None
            dans_tableau = True
            continue
        def val(cle, resoudre=True):
            if cle not in cols:
                return None
            if (r, cols[cle]) in fusions:
                v, premiere = fusions[(r, cols[cle])]
                return v if (resoudre or premiere) else None
            return ws.cell(row=r, column=cols[cle]).value
        v_regie, v_achat, v_dsp = val("regie"), val("achat"), val("dsp")
        v_fmt, v_aud = val("formats", resoudre=False), val("aud", resoudre=False)

        textes = [normaliser(_premiere_ligne(v)) for v in (v_regie, v_fmt, v_dsp)]
        if any(t.startswith(LIGNES_IGNOREES) for t in textes if t):
            continue
        if isinstance(v_achat, (int, float)) or normaliser(v_achat).startswith("repartition"):
            continue                # tableau de répartition budgétaire
        if not v_regie and not v_achat:
            if dans_tableau:
                dans_tableau = False   # ligne de total ou texte sous le tableau : fin du tableau
            continue
        if not dans_tableau and not v_regie:
            continue                   # pas de nouveau tableau sans en-tête ni régie
        if not v_fmt and not v_aud:
            continue
        dans_tableau = True

        if v_regie:
            nouvelle = _premiere_ligne(v_regie)
            if nouvelle != regie:
                regie, dsp = nouvelle, None
        if not regie:
            continue
        if v_achat:
            achat = re.sub(r"\s+", " ", str(v_achat)).strip()
        if v_dsp:
            dsp = re.sub(r"\s+", " ", str(v_dsp)).strip()
        for cle in ("debut", "fin"):
            v = val(cle)
            if isinstance(v, datetime):
                if cle == "debut" and debut is None:
                    debut = v
                if cle == "fin" and fin is None:
                    fin = v

        auds = decouper_audiences(v_aud)
        if v_fmt:
            if not auds and regie in dernier_bloc_regie:
                bloc = dernier_bloc_regie[regie]          # reprend les audiences du bloc précédent
            else:
                bloc = f"{regie} · L{r}"
                for libelle, pct in auds:
                    audiences.append({"bloc": bloc, "regie_pm": regie, "audience_pm": libelle, "pct": pct})
            dernier_bloc_regie[regie] = bloc
            budget, cpm = val("budget", resoudre=False), val("cpm", resoudre=False)
            for k, f in enumerate(decouper_formats(v_fmt)):
                formats.append({"bloc": bloc, "regie_pm": regie, "format_pm": f,
                                "type_achat": achat, "dsp": dsp, "ligne": r,
                                "budget": budget if k == 0 and isinstance(budget, (int, float)) else None,
                                "cpm": cpm if isinstance(cpm, (int, float)) else None})
        elif auds and bloc:                               # audiences seules : complètent le bloc
            for libelle, pct in auds:
                audiences.append({"bloc": bloc, "regie_pm": regie, "audience_pm": libelle, "pct": pct})

    return {"formats": formats, "audiences": audiences, "debut": debut, "fin": fin,
            "campagne": nom_campagne, "onglet": onglet}
