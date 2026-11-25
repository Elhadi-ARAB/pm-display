"""Export CM : code existant de l'app « Générateur de fichiers Media Carrefour », repris tel quel.

Seule l'interface Streamlit a été retirée (elle est désormais dans app.py).
Ne pas modifier la logique ici sans valider avec l'équipe : ce fichier est la référence de l'export CM.
"""
import os
import re
import urllib.parse
import unicodedata

import pandas as pd

# === Listes de mots-clés pour le mapping ===
video_keywords = [
    "In-Stream-Classic", "In-Stream-Trueview", "Video-Dynamic-Creative", "Connected-TV",
    "Synchro-TV", "In-Stream-trueview-for-reach", "In-Stream-trueview-for-action",
    "In-Stream-trueview-for-lead", "In-Stream-non-skippable", "In-Stream-bumper",
    "In-Stream-catch-up", "Out-Stream-inread", "Out-Stream-pave-video", "Video-Ad-Serving-Template"
]

display_keywords = [
    "Display-IAB", "Display-Native", "Redirect", "Display-Social", "Display-Dynamic-Creative",
    "Digital-Out-of-Home", "Emailing", "Display-panorama", "Display-scratch", "Display-slider",
    "Display-inread", "Display-swipe-to-site", "Display-habillage", "Display-habillage-video",
    "Display-habillage-sliding", "Display-habillage-swapping", "Pixel+Click-Command"
]

ALLOWED_PLATFORMS = ["DV360", "Ogury", "Seedtag_FR", "Spotify_FRA", "M6+", "TF1", "SNCF", "SportEasy",
                     "Teads_FRA", "Dailymotion_FRA", "MPP", "Rakuten_FR", "K-words", "weborama"]


def map_tracking_type(value):
    val = str(value).strip()
    if val in video_keywords:
        return "VIDEO"
    elif val == "Stream-Audio":
        return "AUDIO"
    elif val in display_keywords:
        return "DISPLAY"
    else:
        print(f"⚠️ Format Type non reconnu: {val}. Contactez l'équipe Data.")
        return "OTHER"


_SPECIALS = set(list('!"#$%&\'()*+,/:;<=?>@[\\]^{|}~%'))  # caractères à remplacer par "_"
_UTM_KEYS = {"utm_source", "utm_medium", "utm_campaign", "utm_term", "utm_content"}


def _strip_accents(s: str) -> str:
    """Supprime les accents de manière générique (é→e, ñ→n, etc.)."""
    nfkd = unicodedata.normalize("NFKD", s)
    return "".join(ch for ch in nfkd if not unicodedata.combining(ch))


def _sanitize_text_for_utm(s: str) -> str:
    if s is None:
        return s
    s = str(s)
    s = _strip_accents(s)
    s = "".join("_" if (ch in _SPECIALS or ch.isspace()) else ch for ch in s)
    s = re.sub(r"_+", "_", s).strip("_")
    return s


def _idna_hostname(host: str) -> str:
    if not host:
        return host
    host = host.strip().lower()
    try:
        return host.encode("idna").decode("ascii")
    except Exception:
        return host


def _normalize_path(path: str) -> str:
    if path is None or path == "":
        return "/"
    unquoted = urllib.parse.unquote(path)
    return urllib.parse.quote(unquoted, safe="/%:@-._~!$&'()*+,;=")


def _encode_q_value(v: str) -> str:
    if v is None:
        return ""
    return urllib.parse.quote_plus(v, safe="%:@-._~!$'()*,")


def normalize_url_preserving_utm(url: str) -> str:
    if not url or not isinstance(url, str):
        return url
    u = url.strip()
    if u.startswith("www."):
        u = "https://" + u
    parts = urllib.parse.urlsplit(u)
    scheme = (parts.scheme or "https").lower()
    netloc = _idna_hostname(parts.netloc)
    path = _normalize_path(parts.path)
    pairs = urllib.parse.parse_qsl(parts.query, keep_blank_values=True)
    new_pairs = []
    for k, v in pairs:
        k_enc = urllib.parse.quote(k, safe="%:@-._~")
        if k.lower() in _UTM_KEYS:
            v_safe = (v or "")
            v_safe = v_safe.replace("&", "%26").replace("=", "%3D")
            new_pairs.append(f"{k_enc}" if v_safe == "" else f"{k_enc}={v_safe}")
        else:
            v_enc = _encode_q_value(v or "")
            new_pairs.append(f"{k_enc}" if v_enc == "" else f"{k_enc}={v_enc}")
    new_query = "&".join(new_pairs)
    fragment = urllib.parse.quote(urllib.parse.unquote(parts.fragment or ""), safe="%:@-._~!$&'()*+,;=/")
    return urllib.parse.urlunsplit((scheme, netloc, path, new_query, fragment))


def sanitize_url_utm_values(url: str) -> str:
    if not url or not isinstance(url, str):
        return url
    try:
        parts = urllib.parse.urlsplit(url)
        pairs = urllib.parse.parse_qsl(parts.query, keep_blank_values=True)
        new_pairs = []
        for k, v in pairs:
            k_enc = urllib.parse.quote(k, safe="%:@-._~")
            if k.lower() in _UTM_KEYS:
                raw = _sanitize_text_for_utm(v)
                raw = (raw or "").replace("&", "%26").replace("=", "%3D")
                new_pairs.append(f"{k_enc}" if raw == "" else f"{k_enc}={raw}")
            else:
                new_pairs.append(f"{k_enc}" if v == "" else f"{k_enc}={v}")
        new_query = "&".join(new_pairs)
        return urllib.parse.urlunsplit((parts.scheme, parts.netloc, parts.path, new_query, parts.fragment))
    except Exception:
        return url


def generate_files(df, output_folder="exports_cm"):
    os.makedirs(output_folder, exist_ok=True)
    grouped = df.groupby(["Nom de la campagne*", "Début  JJ/MM/AAAA", "Fin  JJ/MM/AAAA"])
    filepaths = []
    ignored_platforms = []
    unknown_tracking = []

    for (campagne, start, end), group in grouped:
        if pd.isna(campagne) or pd.isna(start) or pd.isna(end):
            continue

        header_data = {
            "Consultant_Email": "",
            "Campaign_Name": campagne,
            "Advertiser Name": "NEW Carrefour One - Havas",
            "Landing page": "https://www.carrefour.fr/",
            "Start_Date": start,
            "End_Date": end
        }

        rows = []
        for _, row in group.iterrows():
            if pd.isna(row["Format size (CM only)"]):
                continue

            platform = str(row["Platform"]).strip()
            if platform not in ALLOWED_PLATFORMS:
                ignored_platforms.append((campagne, platform))
                continue

            if platform == "DV360":
                site_name = "DV360 - CRF Marketing - Agence 79"
            elif platform == "TF1":
                site_name = "TF1_FRA"
            elif platform == "SNCF":
                site_name = "Groupe Sncf"
            else:
                site_name = platform

            raw_tracking = row.get("Tracking Type", "")
            tracking_type = map_tracking_type(raw_tracking)
            if tracking_type == "OTHER":
                unknown_tracking.append((campagne, str(raw_tracking).strip()))
                continue

            original_url = row.get("URL de redirection trackée")
            formatted_url = normalize_url_preserving_utm(original_url)
            safe_url = sanitize_url_utm_values(formatted_url)

            rows.append({
                "Site Name": site_name,
                "Type de Tracking": tracking_type,
                "Format": row["Format size (CM only)"],
                "Placement_Name": row["Placement CM = creative DV360"],
                "Creative_Name": row["Creative Name"],
                "Add_ClickTroughUrl": safe_url,
                "IMPRESSION_JAVASCRIPT_EVENT_TAG": row["XPLN script CM"]
            })

        if not rows:
            continue

        df_table = pd.DataFrame(rows)
        safe_name = str(campagne).replace(" ", "_").replace("/", "_")
        filename = f"CM_{safe_name}_carrefour.xlsx"
        filepath = os.path.join(output_folder, filename)

        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            for i, (key, val) in enumerate(header_data.items()):
                pd.DataFrame([[key, val]]).to_excel(writer, index=False, header=False, startrow=i, startcol=0)
            df_table.to_excel(writer, index=False, startrow=9)

        filepaths.append(filepath)

    return filepaths, ignored_platforms, unknown_tracking
