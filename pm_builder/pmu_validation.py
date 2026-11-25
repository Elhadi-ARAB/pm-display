"""Règles de l'app « automation PMU » (validation du fichier média et renommage des créas), reprises telles quelles.

- Validation : en-têtes exacts, consultant email, nom de campagne CM, advertiser, support dans Sites,
  landing dans LPS, tracking PCC/RD, et dimension obligatoire pour les lignes RD.
- Nettoyage des URL : seules les valeurs utm_* sont nettoyées (accents retirés, espaces et caractères
  spéciaux -> _), avant et après le « # ». Le reste de l'URL est conservé à l'identique.
- Renommage des créas : lignes RD uniquement, appariement par dimension, une copie par Créative DCM.
"""
from __future__ import annotations

import os
import re
import struct
import unicodedata
import zipfile
from io import BytesIO

import pandas as pd
from openpyxl import load_workbook

SHEET_MAIN = "fichier_media"
HEADER_ROW = 12
META_COL_VALUE = 1
URL_COLUMN = "URL avec les utms"
DEFAULT_TRACKING = ["PCC", "RD"]
REQUIRED_COLUMNS = ["Support", "Type de Tracking", "Format", "Landing Page", "URL avec les utms"]
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
DIM_RE = re.compile(r"^\d{1,4}\s*[xX]\s*\d{1,4}$")

# ------------------------------------------------------------------ nettoyage des URL
_UTM_KEYS = {"utm_source", "utm_medium", "utm_campaign", "utm_term", "utm_content"}
_SPECIALS = set(list('!"#$%&\'()*+,/:;<=?>@[\\]^`{|}~'))


def _strip_accents(s: str) -> str:
    return "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))


def sanitize_utm_value(v) -> str:
    if v is None:
        return v
    s = _strip_accents(str(v))
    s = "".join("_" if (ch in _SPECIALS or ch.isspace()) else ch for ch in s)
    return re.sub(r"_+", "_", s).strip("_")


def _clean_qs(qs: str) -> str:
    if not qs:
        return qs
    out = []
    for p in qs.split("&"):
        if "=" in p:
            k, v = p.split("=", 1)
            out.append(f"{k}={sanitize_utm_value(v)}" if k.lower() in _UTM_KEYS else p)
        else:
            out.append(p)
    return "&".join(out)


def clean_url(url: str) -> str:
    if not url or not isinstance(url, str):
        return url
    s = url.strip()
    before, after = (s.split("#", 1) + [None])[:2] if "#" in s else (s, None)
    if "?" in before:
        base, q = before.split("?", 1)
        before = base + "?" + _clean_qs(q)
    if after is not None:
        if "?" in after:
            fp, fq = after.split("?", 1)
            after = fp + "?" + _clean_qs(fq)
        return before + "#" + after
    return before


# ------------------------------------------------------------------ validation d'un fichier média
def _ref(xls) -> dict:
    ref = {"tracking": DEFAULT_TRACKING, "sites": set(), "advertisers": set(), "lps_urls": set()}

    def noms(sheet, motcle):
        try:
            d = pd.read_excel(xls, sheet_name=sheet, header=0)
            col = [c for c in d.columns if motcle in str(c).lower()]
            return set(str(x).strip() for x in d[col[0]].dropna()) if col else set()
        except Exception:
            return set()

    if "Type de Tracking" in xls.sheet_names:
        v = [str(x).strip() for x in pd.read_excel(xls, sheet_name="Type de Tracking", header=None).iloc[:, 0].dropna()]
        ref["tracking"] = v or DEFAULT_TRACKING
    if "Sites" in xls.sheet_names:
        ref["sites"] = noms("Sites", "name")
    if "Advertiser" in xls.sheet_names:
        ref["advertisers"] = noms("Advertiser", "name")
    if "LPS" in xls.sheet_names:
        ref["lps_urls"] = noms("LPS", "url")
    return ref


def valider_fichier(contenu: bytes) -> dict:
    """Valide un fichier média PMU. Renvoie {erreurs, corrections_url, df_corrige, nom_campagne, nb_lignes}."""
    xls = pd.ExcelFile(BytesIO(contenu))
    if SHEET_MAIN not in xls.sheet_names:
        return {"erreurs": [f"L'onglet « {SHEET_MAIN} » est introuvable."], "corrections_url": [], "df_corrige": None,
                "nom_campagne": None, "nb_lignes": 0}
    raw = pd.read_excel(xls, sheet_name=SHEET_MAIN, header=None)
    df = pd.read_excel(xls, sheet_name=SHEET_MAIN, skiprows=HEADER_ROW).dropna(how="all").reset_index(drop=True)
    ref = _ref(xls)
    err = []
    cols = {str(c).strip(): str(c) for c in df.columns}
    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            err.append(f"En-tête « {cols[col]} » avec des espaces parasites (attendu « {col} »)." if col in cols
                       else f"Colonne requise manquante : « {col} ».")
    meta = lambda i: raw.iloc[i, META_COL_VALUE] if raw.shape[0] > i else None
    vide = lambda x: x is None or (isinstance(x, float) and pd.isna(x)) or str(x).strip() == ""
    email, camp, adv = meta(0), meta(7), meta(11)
    if vide(email):
        err.append("Consultant email manquant (cellule B1).")
    elif not EMAIL_RE.match(str(email).strip()):
        err.append(f"Consultant email invalide : « {email} » (cellule B1).")
    if vide(camp):
        err.append("Nom de campagne CM manquant (cellule B8).")
    if vide(adv):
        err.append("Advertiser CM manquant (cellule B12).")
    elif ref["advertisers"] and str(adv).strip() not in ref["advertisers"]:
        err.append(f"Advertiser CM « {adv} » introuvable dans l'onglet Advertiser (cellule B12).")
    for idx, row in df.iterrows():
        q = f"Ligne {HEADER_ROW + 2 + idx}"
        sup, lp, trk, fmt = row.get("Support"), row.get("Landing Page"), row.get("Type de Tracking"), row.get("Format")
        if "Support" in df.columns:
            if vide(sup):
                err.append(f"{q} : Support vide.")
            elif ref["sites"] and str(sup).strip() not in ref["sites"]:
                err.append(f"{q} : Support « {sup} » introuvable dans l'onglet Sites.")
        if "Landing Page" in df.columns:
            if vide(lp):
                err.append(f"{q} : Landing Page vide.")
            elif ref["lps_urls"] and str(lp).strip() not in ref["lps_urls"]:
                err.append(f"{q} : Landing Page « {lp} » introuvable dans l'onglet LPS.")
        if "Type de Tracking" in df.columns:
            t = "" if vide(trk) else str(trk).strip().upper()
            if not t:
                err.append(f"{q} : Type de Tracking vide (PCC ou RD attendu).")
            elif t not in [x.upper() for x in ref["tracking"]]:
                err.append(f"{q} : Type de Tracking « {trk} » non autorisé.")
            elif t == "RD" and "Format" in df.columns and (vide(fmt) or not DIM_RE.match(str(fmt).strip())):
                err.append(f"{q} : Format « {'' if vide(fmt) else fmt} » invalide pour une ligne RD (ex. 300x250).")
    corrections, df2 = [], df.copy()
    if URL_COLUMN in df.columns:
        for idx, u in df[URL_COLUMN].items():
            if vide(u):
                continue
            c = clean_url(str(u))
            if c != str(u):
                corrections.append({"ligne": HEADER_ROW + 2 + idx, "avant": str(u), "apres": c})
                df2.at[idx, URL_COLUMN] = c
    return {"erreurs": err, "corrections_url": corrections, "df_corrige": df2,
            "nom_campagne": None if vide(camp) else str(camp).strip(), "nb_lignes": len(df)}


def fichier_propre(contenu: bytes, df_corrige: pd.DataFrame) -> bytes:
    """Réécrit le fichier : formules figées en valeurs, seule la colonne URL est remplacée par la version corrigée."""
    wb = load_workbook(BytesIO(contenu), data_only=True)
    ws = wb[SHEET_MAIN]
    col = next((c.column for c in ws[HEADER_ROW + 1] if c.value is not None and str(c.value).strip() == URL_COLUMN), None)
    if col:
        for i, (_, r) in enumerate(df_corrige.iterrows()):
            if not pd.isna(r.get(URL_COLUMN)):
                ws.cell(row=HEADER_ROW + 2 + i, column=col, value=str(r[URL_COLUMN]))
    out = BytesIO(); wb.save(out); return out.getvalue()


# ------------------------------------------------------------------ renommage des créas (lignes RD)
IMG_EXTS = {".jpg", ".jpeg", ".png", ".gif"}
VIDEO_EXTS = {".mp4", ".mov", ".webm", ".avi"}


def _dim_image(b):
    try:
        from PIL import Image
        return Image.open(BytesIO(b)).size
    except Exception:
        return None


def _dim_mp4(b):
    try:
        i = b.find(b"tkhd")
        if i == -1:
            return None
        taille = struct.unpack(">I", b[i - 4:i])[0]
        fin = i - 4 + taille
        w, h = struct.unpack(">I", b[fin - 8:fin - 4])[0] >> 16, struct.unpack(">I", b[fin - 4:fin])[0] >> 16
        return (w, h) if w > 0 and h > 0 else None
    except Exception:
        return None


def index_formats(df: pd.DataFrame) -> dict:
    """{(largeur, hauteur): [Créative DCM, …]} pour les lignes RD (une créa par support)."""
    idx = {}
    for _, r in df.iterrows():
        if str(r.get("Type de Tracking", "")).strip().upper() != "RD":
            continue
        m = re.match(r"^\s*(\d{1,4})\s*[xX]\s*(\d{1,4})\s*$", str(r.get("Format", "")))
        crea = r.get("Créative DCM")
        if m and crea and not pd.isna(crea):
            l = idx.setdefault((int(m.group(1)), int(m.group(2))), [])
            if str(crea).strip() not in l:
                l.append(str(crea).strip())
    return idx


def renommer_creas(fichiers: list, df: pd.DataFrame):
    """fichiers : [(nom, octets)]. Renvoie (zip, renommés, dimensions absentes, à la main)."""
    idx = index_formats(df)
    ok, absentes, manuel = [], [], []
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for nom, b in fichiers:
            ext = os.path.splitext(nom)[1].lower()
            dims = _dim_image(b) if ext in IMG_EXTS else (_dim_mp4(b) if ext in VIDEO_EXTS else None)
            if ext not in IMG_EXTS | VIDEO_EXTS:
                manuel.append((nom, f"extension {ext} non gérée")); continue
            if dims is None:
                manuel.append((nom, "dimension non détectée")); continue
            if dims in idx:
                noms = [f"{c}{ext}" for c in idx[dims]]
                for n in noms:
                    zf.writestr(n, b)
                ok.append((nom, noms, f"{dims[0]}x{dims[1]}"))
            else:
                absentes.append((nom, f"{dims[0]}x{dims[1]}"))
    return buf.getvalue(), ok, absentes, manuel
