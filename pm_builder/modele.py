"""Définition des champs du PM commun, par client, et exemples de démonstration."""
from __future__ import annotations

from datetime import date

import pandas as pd

import tables as T

CLIENTS = ["Carrefour", "PMU"]
SECTIONS = ["Branding", "Considération", "Trafic", "Performance"]
TYPES_CIBLAGE = {  # vocabulaire du PM -> Targeting Type de l'URL Builder Carrefour
    "Audience prio": "3rd - Audience prio", "Intérêts": "3rd - Interets", "Contextuel": "3rd - Contextuel",
    "Broad": "3rd - Broad", "Socio démo": "3rd - Socio demo", "Intention": "3rd - Intentioniste",
    "3rd party": "3rd-Party-Targeting", "1st party": "1st-Party-Targeting", "Géoloc": "3rd - Geoloc"}


def listes() -> dict:
    L = T.listes()
    srt = lambda v: sorted({str(x) for x in v if str(x) not in ("x", "")}, key=str.lower)
    return {
        "Carrefour": {
            "Levier": ["Vidéo", "Audio", "Display", "CTV", "Social"],
            "Régie": srt(L["Régie"]),
            "Type d'achat": ["PG", "CPM Floor", "PMP", "GAG", "Réservation"],
            "Type de format": srt(L["Format Type"]),
            "Tracking Type": [v for v in L["Tracking Type"] if v != "x"],
            "Pilier": L["Pilier"], "Objectif": L["Objectif"],
            "Type de ciblage": list(TYPES_CIBLAGE), "Targeting details": srt(L["Targeting details"]),
        },
        "PMU": {
            "Levier": ["VOL", "DISPLAY", "SOCIAL"],
            "Régie": ["366", "Le Parisien", "Meta", "Zone Turf", "Paris Turf", "Geny", "Paris Turf - Geny",
                      "Canal Turf", "Zinfo", "Le Veinard", "TTD & Adsquare", "Equidia", "Google"],
            "Type d'achat": ["CPM", "CPM Dyn", "FORFAIT"],
            "Type de format": ["STATIC", "IAB", "VIDEO"],
            "Device": ["MULTIDEVICE", "MOBILE/DESKTOP", "WEB"], "Programmation": ["79", "Régie"],
            "Type de tracking": ["PCC", "RD"], "Oui/Non": ["Oui", "Non"],
            "Levier campagne": ["BRAND", "CONSID"], "H/S": ["H", "S"], "85/2010": ["85", "2010"],
            "Medium": ["DIB"], "Promo/Offre": ["Promo", "Offre", "Filrouge"],
        },
    }


# PMU : landing pages autorisées (liste CM : ID, nom, URL). Le fichier média reprend l'URL, l'ID fait la correspondance côté CM.
LP_PMU = [
    ("7668217", "pmu", "https://info.pmu.fr/"),
    ("40544563", "LP", "https://www.pmu.fr/"),
    ("9971400", "LP EPIQE SSL", "https://epiqe.fr/"),
    ("30978874", "HP", "https://www.equidia.fr/"),
    ("33547046", "LP Inscription", "https://inscription.pmu.fr/"),
    ("35279969", "lp offline", "https://www.pmu.fr/point-de-vente/"),
    ("39349736", "LP DEFAULT", "http://google.com"),
    ("43301433", "lp_inscription_redirection_civilite", "https://inscription.pmu.fr/?typeCompte=2010&clientApi=78&codePromo=SPORT&univers=SPORT&redirectionUrl=https%3A%2F%2Fpmu.fr%2Fsport#/civilite"),
    ("43425225", "Inscription_hippique", "https://inscription.pmu.fr/?typeCompte=2010&codePromo=HIPPIQUE"),
    ("44233906", "lp_inscription_redirection_api", "https://inscription.pmu.fr/?typeCompte=2010&clientApi=61&redirectionUrl=https%3A%2F%2Fpmu.fr%2Fturf#/civilite"),
    ("44785841", "LP Superpick", "https://www.pmu.fr/turf/cote-courses/lancement-e-superpick5-promo/"),
]
LP_URLS = [u for _, _, u in LP_PMU]
LP_LIBELLE = {u: f"{n} · {u}" for _, n, u in LP_PMU}
LP_ID = {u: i for i, _, u in LP_PMU}
ADVERTISER_PMU = "PMU_MASTER_TAG"
# PMU : sites CM (onglet Sites du fichier média) et advertisers (onglet Advertiser)
SITES_PMU = [["2569439", "ZONETURRF"], ["2569627", "BILTOO"], ["2569885", "GENYCOURSE"], ["2570035", "TEST"], ["2570233", "UN1TRTBCONQUETE"], ["2571237", "PARISTURF"], ["2571467", "PARISCOURSES"], ["2571648", "TIERCEMAGAZINE"], ["2571889", "FOOTMERCATOADVERSPORT"], ["2571892", "LEQUIIPE"], ["2572081", "WEEKENDTURF"], ["2572083", "TF1"], ["2572233", "RMC"], ["2572868", "EQUIDIA"], ["2572869", "LEBONCOIN"], ["2572872", "PMUSITESPMU"], ["2573023", "FACEBOOK"], ["3633464", "ORANGE"], ["4335399", "UN1T"], ["5036909", "GENYYSPORT"], ["6276469", "GravityFR"], ["6343144", "Dailymotion"], ["6906915", "TABMO"], ["6930635", "Veepee"], ["6998361", "SeedTag"], ["6998451", "Youtube"], ["7280038", "CANALTURF"], ["7332558", "DARTSearchGoogle"], ["7351422", "DARTSearchMSN"], ["7845890", "Turf"], ["7845902", "Turfomania"], ["7999181", "RTLFrance"], ["8023342", "funradio"], ["8229832", "AdrenaleadFR"], ["8408322", "AlticeMedia"], ["8575905", "sirdata"], ["8576763", "InviteMedia"], ["8579396", "AdsMovil"], ["8801054", "havasmedia"], ["9010015", "ShowroomPrive"], ["9039881", "PMPrecisionFRXandr"], ["9054808", "Prisma-DBM"], ["9062600", "PresseQuotidienneRégionale"], ["9073384", "RoktFR"], ["9147280", "Snapchat"], ["9161817", "Dataventure"], ["9306656", "SportLocalMediaFRA"], ["9405432", "Manageo"], ["9409272", "Infopro"], ["9409278", "PLS"], ["9445172", "maxifoot"], ["9524748", "UberEats"], ["9589358", "ButFootball"], ["9895265", "gravity"], ["9895645", "Quarry"], ["9899293", "M6"], ["9899302", "Rakuten"], ["9909459", "Ogury"], ["4455901", "TURFOO_TURFOO"], ["10197876", "Amazon"], ["10414627", "Displayce"], ["2571891", "BEINSPORTS_BEINSPORTS"], ["5420832", "Teads"], ["8042744", "CanalPlus"]]
SITE_NOMS = [n for _, n in SITES_PMU]
ADVERTISERS_PMU = [["5453254", "PMU_MASTER_TAG"], ["5456749", "2010_WEB"], ["5459600", "2010_APP"], ["5459800", "2010_MOB"], ["5460200", "85_MOB"], ["5460400", "85_WEB"], ["5460800", "85_APP"], ["5626045", "EPIQE"], ["8276712", "BidManager_SeparateSpotlightAdvertiser_DO_NOT_EDIT"], ["8284559", "DO NOT USE - EQUIDIA"], ["8284562", "BidManager_Advertiser_DO_NOT_EDIT_2187640"]]


# PMU : la section du PM donne le levier de campagne (un fichier média par levier)
LEVIER_PAR_SECTION = {"Branding": "BRAND", "Considération": "CONSID"}
# PMU : site CM des régies achetées en direct (programmation « Régie »)
SITES_DIRECT = {"zone turf": "ZONETURRF", "paris turf - geny": "PARISTURF", "paris turf": "PARISTURF", "geny": "GENYCOURSE",
                "canal turf": "CANALTURF", "equidia": "EQUIDIA", "le parisien": "UN1T"}


# anciens libellés de levier -> nouveaux
LEVIER_ALIAS_PMU = {"TTD": "VOL", "DV360": "VOL", "DISPLAY TURF": "DISPLAY"}


def completer_pmu(lignes: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """Complète les cases vides des lignes PMU :
    - Programmation / Support selon le levier : VOL -> 79 + UN1T ; SOCIAL -> 79 + FACEBOOK ; DISPLAY -> Régie + site CM ;
    (Double Verify, XPLN et DAR ne sont jamais remplis automatiquement : case remplie = oui, vide = non.)
    - Emplacement : « RG sur site » ; Créative : « TBD » ; Ciblage PM client : le ciblage court."""
    df, notes = lignes.copy(), []
    vide = lambda v: v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == ""
    for col in ("Objectif", "Emplacement", "Ciblage PM client", "Double Verify", "XPLN", "DAR", "Créative", "Commentaire"):
        if col not in df.columns:
            df[col] = None
    for i, l in df.iterrows():
        lev = LEVIER_ALIAS_PMU.get(str(l.get("Levier") or "").upper(), str(l.get("Levier") or "").upper())
        if lev and lev != l.get("Levier"):
            df.at[i, "Levier"] = lev
        if lev == "VOL":
            prog, sup = "79", "UN1T"
        elif lev == "SOCIAL":
            prog, sup = "79", "FACEBOOK"
        elif lev == "DISPLAY":
            prog, sup = "Régie", SITES_DIRECT.get(T.normaliser(l.get("Régie")))
        else:
            prog, sup = None, None
        if prog and vide(l.get("Programmation")):
            df.at[i, "Programmation"] = prog
        if prog and vide(l.get("Support")):
            if sup:
                df.at[i, "Support"] = sup
            else:
                notes.append(f"Ligne {l['N° ligne']} ({l.get('Régie')}) : site CM inconnu pour cette régie, Support à remplir.")
        if vide(l.get("Emplacement")):
            df.at[i, "Emplacement"] = "RG sur site"
        if vide(l.get("Créative")):
            df.at[i, "Créative"] = "TBD"
        if vide(l.get("Ciblage PM client")) and not vide(l.get("Ciblage")):
            df.at[i, "Ciblage PM client"] = l["Ciblage"]
    return df, notes


COLS_COMMUNES = ["Section", "Levier", "Régie", "Type d'achat", "Format", "Type de format", "Ciblage",
                 "Début", "Fin", "Budget net (€)", "CPM (€)"]
COLS_CLIENT = {"Carrefour": ["Groupe d'audiences", "Nb messages", "Durée (s)", "Tracking Type",
                             "Device", "Specs", "KPI", "VTR / LTR", "Fréquence", "Études / BLS"],
               "PMU": ["Objectif", "Emplacement", "Ciblage PM client", "Device", "Programmation", "Support", "Type de tracking",
                       "Double Verify", "XPLN", "DAR", "Créative", "Commentaire"]}
COLS_CREAS = ["N° ligne", "Nom du format", "Taille", "Piste créative / message", "Code message", "Landing (si différente)"]
COLS_AUDIENCES = ["Groupe d'audiences", "Audience", "Type de ciblage", "Targeting details", "% du groupe"]


def campagne_vide(client: str) -> dict:
    c = {"client": client, "nom": "", "debut": None, "fin": None, "landing": ""}
    if client == "Carrefour":
        c.update({"pilier": "", "objectif": "", "libelle": "", "owner_crf": ""})
    else:
        c.update({"levier": "BRAND", "hs": "H", "univers": "85", "medium": "DIB", "promo": "Promo",
                  "numero_projet": "", "advertiser": ADVERTISER_PMU, "consultant": "", "type_campagne": "BRAND"})
    return c


def lignes_vides(client: str) -> pd.DataFrame:
    return pd.DataFrame(columns=["N° ligne"] + COLS_COMMUNES + COLS_CLIENT[client])


def creas_vides() -> pd.DataFrame:
    return pd.DataFrame(columns=COLS_CREAS)


def audiences_vides() -> pd.DataFrame:
    return pd.DataFrame(columns=COLS_AUDIENCES)


def numeroter(df: pd.DataFrame) -> pd.DataFrame:
    """Donne un N° de ligne stable (L1, L2…) aux lignes qui n'en ont pas, sans renuméroter les autres."""
    df = df.copy().reset_index(drop=True)
    pris = {str(v) for v in df["N° ligne"] if isinstance(v, str) and v.startswith("L")}
    n = 1
    for i in df.index:
        v = df.at[i, "N° ligne"]
        if not (isinstance(v, str) and v.strip()):
            while f"L{n}" in pris:
                n += 1
            df.at[i, "N° ligne"] = f"L{n}"
            pris.add(f"L{n}")
    return df


# ------------------------------------------------------------------ exemples
def exemple_carrefour():
    camp = campagne_vide("Carrefour")
    camp.update({"nom": "MDC_Sdg_V2", "debut": date(2026, 10, 13), "fin": date(2026, 10, 26),
                 "landing": "https://www.carrefour.fr/evenements/semaine-du-gout", "pilier": "Marque", "objectif": "Branding",
                 "libelle": "MDC SDG V2"})
    d0, d1 = date(2026, 10, 13), date(2026, 10, 26)
    YT = None   # vide : le PM client affiche les audiences du groupe (onglet Audiences)
    rows = [
        ("Branding", "Vidéo", "Youtube", "CPM Floor", "In stream non skippable", "Non-skippable", YT, 15000, 6, "YT", 1, 20),
        ("Branding", "Vidéo", "Youtube", "CPM Floor", "Bumper", "In-Stream-bumper", None, 5000, 6, "YT", 1, 6),
        ("Branding", "Vidéo", "Youtube", "CPM Floor", "Shorts", "Shorts", None, 12015, 3, "YT", 1, 10),
        ("Branding", "Audio", "Spotify", "GAG", "Audio + companion banner", "Stream-Audio", None, 30000, 6.01, "SP", 3, 30),
        ("Branding", "Vidéo", "M6", "GAG", "HACK6", "Hack6", None, 30000, 16.5, "M6", 1, 20),
        ("Branding", "Vidéo", "TF1", "PG", "Midroll", "Midroll", None, 15000, 17.17, "TF1", 1, 20),
        ("Branding", "Vidéo", "France-TV", "PG", "Vidéo enrichie skin préroll QR code", "In-Stream-enrichi", None, 42000, 16.2, "FTV", 1, 20),
        ("Branding", "Vidéo", "Weborama", "GAG", "Vidéo enrichie", "In-Stream-enrichi", None, 30000, 10, "WEB", 1, 20),
    ]
    lignes = pd.DataFrame([{"N° ligne": f"L{i+1}", "Section": s, "Levier": l, "Régie": r, "Type d'achat": a, "Format": f,
                            "Type de format": ft, "Ciblage": c, "Début": d0, "Fin": d1, "Budget net (€)": b, "CPM (€)": p,
                            "Groupe d'audiences": g, "Nb messages": m, "Durée (s)": du, "Tracking Type": "Pixel+Click-Command",
                            "Device": "100% CTV" if r == "France-TV" else "Multidevice", "Specs": None,
                            "KPI": "Reach/Completion" if r in ("Youtube", "TF1", "France-TV") else "Reach",
                            "VTR / LTR": {"Youtube": 0.83, "Spotify": 0.87, "M6": 0.94, "TF1": 0.93, "France-TV": 0.93, "Weborama": 0.78}[r]
                            if f != "Shorts" else 0.019, "Fréquence": None, "Études / BLS": None}
                           for i, (s, l, r, a, f, ft, c, b, p, g, m, du) in enumerate(rows)])
    aud = [("YT", "Parents 25-65 ans", "Audience prio", "Familles", 35), ("YT", "Value shoppers", "Contextuel", "Promovores", 15),
           ("YT", "Seniors +55 ans", "Audience prio", "Seniors", 20), ("YT", "Food & Dining", "Contextuel", "Contextual", 10),
           ("YT", "Light TV viewers", "Intérêts", "Petit-conso-TV", 10), ("YT", "Aisés, CSP+", "Audience prio", "Aisés / CSP+", 10),
           ("SP", "25-54 ans", "Broad", "Socio-demo", 70), ("SP", "55+ ans", "Broad", "Socio-demo", 30)]
    base = [("Audience prio", "Familles", 30), ("3rd party", "Promovores", 15), ("Audience prio", "Seniors", 20),
            ("Audience prio", "Aisés / CSP+", 20), ("Contextuel", "Contextual", 15)]
    libs = {"M6": ["Foyers avec enfants", "Promophiles", "Data 60+", "Data CSP", "Gourmets"],
            "TF1": ["Familles avec enfants", "Promophiles", "Individus 60+", "Hauts revenus", "Food lovers"],
            "FTV": ["Familles avec enfants", "Chasseurs de promo", "Seniors", "CSP+", "Data gourmets"],
            "WEB": ["Familles organisées", "Chasseurs de bonnes affaires", "Seniors gourmands", "Hauts revenus", "Les épicuriens"]}
    for g, ls in libs.items():
        aud += [(g, lib, tc, td, p) for lib, (tc, td, p) in zip(ls, base)]
    audiences = pd.DataFrame(aud, columns=COLS_AUDIENCES)
    creas = pd.DataFrame([("L4", "Audio 30 s", "1x1", m, code, None) for m, code in
                          (("pomme_de_terre", "PDT"), ("huile", "Huile-olive"), ("raclette", "raclette"))], columns=COLS_CREAS)
    return camp, lignes, creas, audiences


def exemple_pmu():
    camp = campagne_vide("PMU")
    camp.update({"nom": "SEPTEMBRE_PMU_PDV", "debut": date(2026, 9, 1), "fin": date(2026, 9, 30),
                 "landing": "https://www.pmu.fr/point-de-vente/", "consultant": "lea.le-provost@agence79.io"})
    d0, d1 = date(2026, 9, 1), date(2026, 9, 30)
    rows = [("VOL", "366", " Visibilité", "Static Smart cover", "Turfistes reguliers", "Turfistes reguliers", 5163.64, 10.33, "UN1T", "Oui"),
            ("VOL", "Le Parisien", "Reach & Visibilité", "Habillage & IAB", "Sous-Rubrique Hippisme", "Sous-Rubrique Hippisme 100% SOV", 5171.64, 11.38, "UN1T", "Oui"),
            ("SOCIAL", "Meta", "Reach & Visibilité", "Instream, Reels, Stories", "Turfistes reguliers", "ALL35-55 ans - Turfistes reguliers", 10000, 3, "FACEBOOK", "Non")]
    lignes = pd.DataFrame([{"N° ligne": f"L{i+1}", "Section": "Branding", "Levier": lev, "Régie": reg, "Type d'achat": "CPM",
                            "Format": fmt, "Type de format": "STATIC", "Ciblage": cib, "Début": d0, "Fin": d1,
                            "Budget net (€)": bud, "CPM (€)": cpm, "Objectif": obj.strip(), "Emplacement": "RG sur site",
                            "Ciblage PM client": cib_pm, "Device": "MULTIDEVICE", "Programmation": "79", "Support": sup,
                            "Type de tracking": "PCC", "Double Verify": dv == "Oui", "XPLN": False, "DAR": True,
                            "Créative": "TBD", "Commentaire": None}
                           for i, (lev, reg, obj, fmt, cib, cib_pm, bud, cpm, sup, dv) in enumerate(rows)])
    creas = pd.DataFrame([("L1", "Smartcover", "2560x430", "Smartcover", None, None), ("L1", "Smartcover", "1440x575", "Smartcover", None, None),
                          ("L2", "Habillage", "1800 x 1000", "Habillage", None, None), ("L2", "Pavé", "300x250", "Pae", None, None),
                          ("L2", "Grand Angle", "300x600", "Grand Angle", None, None),
                          ("L3", "Fil d'actualité", "1440x1440", "Fil d'actualite 1440x1440", None, None),
                          ("L3", "Story-Reels", "1440x2560", "Story-Reels 1440x2560", None, None)], columns=COLS_CREAS)
    return camp, lignes, creas, audiences_vides()
