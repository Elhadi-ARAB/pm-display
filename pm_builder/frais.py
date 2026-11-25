"""Frais PMU calculés à partir des impressions (mêmes règles que le PM PMU).

- Volume (impressions) = budget net / CPM × 1000
- Double Verify = impressions × 0,09 / 1000 × 1,25   (lignes dont la case Double Verify est remplie)
- XPLN          = impressions × 0,2  / 1000 × 1,25   (lignes dont la case XPLN est remplie)
- DAR, par section, sur les impressions des lignes dont la case DAR est remplie :
  800 € jusqu'à 15 M d'impressions, puis 800 + (impressions − 15 M) × 0,052 / 1000
Une case remplie (coche, « x », n'importe quelle valeur) = la ligne a ce frais ; vide = non.
- Fees Meta = 3 % du net média des lignes SOCIAL
"""
from __future__ import annotations

import pandas as pd

DV_TAUX, XPLN_TAUX, MAJORATION = 0.09, 0.2, 1.25
DAR_FORFAIT, DAR_SEUIL, DAR_TAUX = 800, 15_000_000, 0.052
META_TAUX = 0.03


def actif(v) -> bool:
    """Case remplie = oui (coche, « x », n'importe quel texte) ; vide = non. « Non », « False », « 0 » restent non."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return False
    if isinstance(v, bool):
        return v
    return str(v).strip().lower() not in ("", "non", "false", "0", "nan", "none")


_oui = actif  # ancien nom


def volume(budget, cpm) -> float:
    b, p = pd.to_numeric(budget, errors="coerce"), pd.to_numeric(cpm, errors="coerce")
    return float(b / p * 1000) if pd.notna(b) and pd.notna(p) and p > 0 else 0.0


def dar_section(impressions: float) -> float:
    return DAR_FORFAIT if impressions <= DAR_SEUIL else DAR_FORFAIT + (impressions - DAR_SEUIL) * DAR_TAUX / 1000


def par_ligne(lignes: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, l in lignes.iterrows():
        vol = volume(l.get("Budget net (€)"), l.get("CPM (€)"))
        budget = float(pd.to_numeric(l.get("Budget net (€)"), errors="coerce") or 0)
        rows.append({"N° ligne": l["N° ligne"], "Section": l.get("Section"), "Régie": l.get("Régie"), "Volume": vol,
                     "DAR": actif(l.get("DAR")),
                     "Double Verify (€)": vol * DV_TAUX / 1000 * MAJORATION if _oui(l.get("Double Verify")) else 0.0,
                     "XPLN (€)": vol * XPLN_TAUX / 1000 * MAJORATION if _oui(l.get("XPLN")) else 0.0,
                     "Fees Meta (€)": budget * META_TAUX if str(l.get("Levier") or "").upper() == "SOCIAL" else 0.0})
    return pd.DataFrame(rows, columns=["N° ligne", "Section", "Régie", "Volume", "DAR", "Double Verify (€)", "XPLN (€)", "Fees Meta (€)"])


def synthese(camp: dict, lignes: pd.DataFrame) -> dict:
    pl = par_ligne(lignes)
    net = float(pd.to_numeric(lignes["Budget net (€)"], errors="coerce").fillna(0).sum())
    dar = {}
    avec_dar = pl[pl["DAR"].astype(bool)] if len(pl) else pl
    for sec, grp in avec_dar.groupby(avec_dar["Section"].fillna("Branding"), sort=False):
        dar[sec] = dar_section(grp["Volume"].sum())
    meta, dv, xpln = pl["Fees Meta (€)"].sum(), pl["Double Verify (€)"].sum(), pl["XPLN (€)"].sum()
    frais = meta + dv + xpln + sum(dar.values())
    return {"net": net, "meta": meta, "dv": dv, "xpln": xpln, "dar": dar, "frais": frais, "total": net + frais,
            "volume": pl["Volume"].sum(), "par_ligne": pl}


# ======================================================================= Carrefour
# Fees DV360 + taxe Google (barème du PM Carrefour) : YouTube et PG -> 3,6 % + 2 % ; reste du DV360 -> 9,5 % + 2 % ; GAG -> 0.
FEES_YOUTUBE_PG, FEES_DV360 = 5.6, 11.5


def taux_fees_crf(type_achat, regie) -> float:
    """Taux (en %) des fees DV360 + taxe Google d'une ligne Carrefour."""
    achat, reg = str(type_achat or "").strip().lower(), str(regie or "").strip().lower()
    if achat == "gag":
        return 0.0
    if reg.startswith("youtube") or achat in ("pg", "cpm pg", "réservation", "reservation"):
        return FEES_YOUTUBE_PG
    return FEES_DV360


def synthese_crf(lignes: pd.DataFrame) -> dict:
    net = pd.to_numeric(lignes["Budget net (€)"], errors="coerce").fillna(0)
    taux = [taux_fees_crf(a, r) for a, r in zip(lignes["Type d'achat"], lignes["Régie"])]
    fees = sum(n * t / 100 for n, t in zip(net, taux))
    vol = sum(volume(b, c) for b, c in zip(lignes["Budget net (€)"], lignes["CPM (€)"]))
    return {"net": float(net.sum()), "fees": float(fees), "total": float(net.sum() + fees), "volume": vol}
