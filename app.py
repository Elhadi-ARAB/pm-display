"""Générateur URL Builder + CM Carrefour.

Deux modes :
- Créer depuis le PM et le PDR : fiche setup -> contrôles -> URL Builder -> fichier CM.
- Depuis un URL Builder existant : mode historique, génère le fichier CM.
"""
import difflib
import io
import os
import re
import time
from datetime import date

import pandas as pd
import streamlit as st

import tables as T
from export.cm import generate_files, map_tracking_type
from generation import controles, setup, url_builder
from lecture.pdr import campagnes, filtrer, lire_pdr
from lecture.pm import lire_pm, onglets_display

st.set_page_config(page_title="URL Builder & CM Carrefour", page_icon="📊", layout="wide")

# Listes déroulantes : par défaut celles de tables/listes.yaml, ou celles d'un URL Builder à jour.
with st.sidebar.expander("Listes déroulantes"):
    st.caption("Pour utiliser une valeur ajoutée récemment dans l'onglet Inputs A79 (nouvelle régie, "
               "nouveau format, nouveau tracking…), déposez ici un URL Builder à jour.")
    f_listes = st.file_uploader("URL Builder de référence (optionnel)", type=["xlsx"], key="listes")
    surcharge = {}
    if f_listes:
        try:
            surcharge = T.listes_depuis_url_builder(f_listes)
            st.success("Listes chargées depuis l'onglet Inputs A79.")
        except Exception as e:
            st.error(f"Lecture impossible : {e}")
L = T.listes(surcharge)


def telecharger_cm(df):
    paths, ignored_platforms, unknown_tracking = generate_files(df)
    if ignored_platforms:
        st.error("Des lignes ont été ignorées car la plateforme n'est pas autorisée :")
        for campagne, plat in ignored_platforms:
            st.write(f"- Campagne **{campagne}** → plateforme non reconnue : **{plat}**")
    if unknown_tracking:
        st.error("Des lignes ont été ignorées car le Tracking Type n'est pas reconnu :")
        for campagne, trk in unknown_tracking:
            st.write(f"- Campagne **{campagne}** → Tracking Type non reconnu : **{trk}**")
    if ignored_platforms or unknown_tracking:
        st.warning("Aucun fichier CM généré. Corrigez la source puis réessayez.")
        return
    st.info("Pensez à renseigner votre adresse mail en colonne B (ligne Consultant_Email).")
    for i, p in enumerate(paths):
        with open(p, "rb") as f:
            st.download_button(f"Télécharger {os.path.basename(p)}", f.read(),
                               file_name=os.path.basename(p), key=f"cm_{i}")


def en_excel(df) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="URL BUILDER", index=False, startrow=1)
    return buf.getvalue()


def selecteur(options, label):
    return st.column_config.SelectboxColumn(label, options=options, required=True)


# ---------------------------------------------------------------------------
st.title("URL Builder & CM Carrefour")
mode = st.sidebar.radio("Mode", ["Créer depuis le PM et le PDR", "Depuis un URL Builder existant"])

# ======================== Mode historique ========================
if mode == "Depuis un URL Builder existant":
    st.caption("Déposez un URL Builder déjà rempli pour générer le fichier CM.")
    fichier = st.file_uploader("URL Builder (Excel)", type=["xlsx"])
    if fichier:
        try:
            df = pd.read_excel(fichier, sheet_name="URL BUILDER", skiprows=1).dropna(how="all")
            df["Début  JJ/MM/AAAA"] = pd.to_datetime(df["Début  JJ/MM/AAAA"], errors="coerce")
            df["Fin  JJ/MM/AAAA"] = pd.to_datetime(df["Fin  JJ/MM/AAAA"], errors="coerce")
            df = df.dropna(subset=["Nom de la campagne*", "Début  JJ/MM/AAAA", "Fin  JJ/MM/AAAA", "Format size (CM only)"])
            st.success(f"{len(df)} lignes chargées.")
            if st.button("Générer le fichier CM", type="primary"):
                telecharger_cm(df)
        except Exception as e:
            st.error(f"Lecture impossible : {e}")
    st.stop()

# ======================== Étape 1 : fichiers ========================
st.subheader("1. Déposer le PM et le PDR")
c1, c2 = st.columns(2)
f_pm = c1.file_uploader("Plan média", type=["xlsx"])
f_pdr = c2.file_uploader("PDR (plan de roulement global ou extrait)", type=["xlsx"])
if not (f_pm and f_pdr):
    st.info("Déposez les deux fichiers pour préparer la fiche setup.")
    st.stop()

try:
    onglets = onglets_display(f_pm)
    if not onglets:
        st.error("Aucun onglet display avec un en-tête « REGIE/PARTENAIRE » dans ce PM.")
        st.stop()
    onglet = c1.selectbox("Onglet du PM", onglets) if len(onglets) > 1 else onglets[0]
    pm = lire_pm(f_pm, onglet)
    pdr_tout = lire_pdr(f_pdr)
except Exception as e:
    st.error(f"Lecture impossible : {e}")
    st.stop()

liste_camp = campagnes(pdr_tout)
if not liste_camp:
    st.error("Aucune campagne trouvée dans le PDR.")
    st.stop()


def _cle_nom(nom) -> str:
    """« MDC SDG V2 » (PM) et « MDC_Sdg_V2 » (PDR) donnent la même clé."""
    return re.sub(r"[^a-z0-9]", "", T.normaliser(nom))


nom_pm = pm.get("campagne") or ""
defaut = next((i for i, c in enumerate(liste_camp) if nom_pm and _cle_nom(c) == _cle_nom(nom_pm)), None)
if defaut is None and nom_pm:
    proches = difflib.get_close_matches(_cle_nom(nom_pm), [_cle_nom(c) for c in liste_camp], n=1, cutoff=0.6)
    if proches:
        defaut = [_cle_nom(c) for c in liste_camp].index(proches[0])
        c2.warning(f"Campagne du PM « {nom_pm} » : la plus proche dans le PDR est présélectionnée. Vérifiez-la.")
if defaut is None:
    c2.warning(f"La campagne du PM (« {nom_pm or 'non trouvée'} ») n'a pas été retrouvée dans le PDR. "
               "Choisissez-la dans la liste.")
nom_pdr = c2.selectbox("Campagne du PDR", liste_camp, index=defaut if defaut is not None else 0,
                       help=f"Campagne indiquée dans le PM : {nom_pm or 'non trouvée'}")
pdr = filtrer(pdr_tout, nom_pdr)

cle_session = (f_pm.name, f_pm.size, onglet, f_pdr.name, f_pdr.size, nom_pdr)
if st.session_state.get("cle") != cle_session:
    st.session_state.cle = cle_session
    st.session_state.camp = setup.campagne(pm, pdr, nom_pdr, L)
    st.session_state.formats = setup.formats(pm, L)
    st.session_state.audiences = setup.audiences(pm)
    st.session_state.messages = setup.messages(st.session_state.formats, pdr)
    st.session_state.genere = None

# ======================== Étape 2 : fiche setup ========================
st.subheader("2. Vérifier la fiche setup")
camp = st.session_state.camp
statuts = pd.concat([st.session_state.formats["Statut"], st.session_state.audiences["Statut"],
                     st.session_state.messages["Statut"]])
n_inconnus, n_proposes = int((statuts == T.INCONNU).sum()), int((statuts == T.PROPOSE).sum())
if n_inconnus:
    st.warning(f"{n_inconnus} élément(s) inconnu(s) : complétez les cases vides (statut « inconnu »).")
if n_proposes:
    st.info(f"{n_proposes} élément(s) au statut « proposé » : valeurs pré-remplies à partir de propositions "
            "pas encore confirmées. Vérifiez-les.")
if not n_inconnus and not n_proposes:
    st.success("Tous les libellés sont validés : vérifiez simplement les valeurs proposées.")
st.caption("Statut : validé = convention confirmée · proposé = à vérifier · inconnu = à compléter. "
           "Un « bloc » regroupe les formats et les audiences d'une même ligne du PM.")

with st.expander("Campagne", expanded=True):
    a, b, c = st.columns(3)
    camp["nom"] = a.text_input("Nom de campagne (sans « _ »)", camp["nom"])
    camp["pilier"] = b.selectbox("Pilier", L["Pilier"], index=L["Pilier"].index(camp["pilier"]) if camp["pilier"] in L["Pilier"] else 0)
    camp["objectif"] = c.selectbox("Objectif", L["Objectif"], index=L["Objectif"].index(camp["objectif"]) if camp["objectif"] in L["Objectif"] else 0,
                                   help=f"Objectifs dans le PDR : {', '.join(camp['objectifs_pdr']) or 'aucun'}")
    d, e, f = st.columns(3)
    camp["debut"] = d.date_input("Début", camp["debut"] or date.today(), format="DD/MM/YYYY")
    camp["fin"] = e.date_input("Fin", camp["fin"] or date.today(), format="DD/MM/YYYY")
    camp["landing"] = f.text_input("Landing page", camp["landing"])

statut_col = st.column_config.TextColumn("Statut", disabled=True)

st.markdown("**Formats** · une ligne par format (ajoutez ou supprimez des lignes si le PM a été mal découpé)")
st.session_state.formats = st.data_editor(
    st.session_state.formats, key="ed_formats", hide_index=True, width="stretch", num_rows="dynamic",
    disabled=["Type d'achat", "DSP", "Statut"],
    column_config={
        "Bloc": st.column_config.TextColumn("Bloc", help="Les formats et audiences d'un même bloc sont croisés."),
        "Platform": selecteur(L["Plateforme"], "Platform"),
        "Régie": selecteur(L["Régie"], "Régie"),
        "Format Type": selecteur(L["Format Type"], "Format Type"),
        "Format size": selecteur(L["Format Size"], "Format size"),
        "Tracking Type": st.column_config.SelectboxColumn(
            "Tracking Type", options=L["Tracking Type"], required=True,
            help="Pixel+Click-Command (PCC) → DISPLAY dans le fichier CM ; "
                 "Video-Ad-Serving-Template (VAST) → VIDEO."),
        "Une ligne par message": st.column_config.CheckboxColumn("Une ligne par message"),
        "Statut": statut_col,
    })

st.markdown("**Audiences** · une ligne par audience, rattachée à un bloc")
st.session_state.audiences = st.data_editor(
    st.session_state.audiences, key="ed_audiences", hide_index=True, width="stretch",
    num_rows="dynamic", disabled=["Statut", "Origine Targeting Type"],
    column_config={
        "Origine Targeting Type": st.column_config.TextColumn(
            "Origine Targeting Type", help="PM = type de ciblage écrit dans le libellé du PM (il fait foi) ; "
                                           "table = valeur par défaut de l'audience."),
        "Bloc": st.column_config.SelectboxColumn("Bloc", options=sorted(st.session_state.formats["Bloc"].dropna().unique().tolist())),
        "Targeting Type": selecteur(L["Targeting Type"], "Targeting Type"),
        "Targeting details": selecteur(L["Targeting details"], "Targeting details"),
        "% budget": st.column_config.NumberColumn("% budget", format="%g %%"),
        "Statut": statut_col,
    })

if len(st.session_state.messages) or st.session_state.formats["Une ligne par message"].astype(bool).any():
    st.markdown("**Messages** · pour les formats « une ligne par message »")
    st.session_state.messages = st.data_editor(
        st.session_state.messages, key="ed_messages", hide_index=True, width="stretch",
        num_rows="dynamic", disabled=["Statut"])

# ======================== Étape 3 : contrôles ========================
st.subheader("3. Contrôler")
erreurs, avertissements, attendu = controles.verifier(
    camp, st.session_state.formats, st.session_state.audiences, st.session_state.messages, listes=L)
for msg in erreurs:
    st.error(msg)
for msg in avertissements:
    st.warning(msg)
if not erreurs:
    st.success(f"Contrôles OK : {attendu} lignes vont être créées.")
    recap = st.session_state.formats.drop_duplicates(["Régie PM", "Tracking Type"])
    st.caption("Tracking dans le fichier CM : " + " · ".join(
        f"{r['Régie PM']} → {r['Tracking Type']} ({map_tracking_type(r['Tracking Type'])})" for _, r in recap.iterrows()))

# ======================== Étape 4 : génération ========================
st.subheader("4. Générer")
if st.button("Générer l'URL Builder", type="primary", disabled=bool(erreurs)):
    t0 = time.time()
    st.session_state.genere = url_builder.generer(
        camp, st.session_state.formats, st.session_state.audiences, st.session_state.messages)
    st.success(f"{len(st.session_state.genere)} lignes générées en {time.time() - t0:.2f} s.")

df = st.session_state.get("genere")
if df is not None and len(df):
    st.dataframe(df[["Platform", "Régie", "Format Type", "Tracking Type", "Targeting details", "Creative Name",
                     "Placement CM = creative DV360"]], width="stretch", hide_index=True)
    c1, c2 = st.columns(2)
    c1.download_button("Télécharger l'URL Builder", en_excel(df),
                       file_name=f"URL_BUILDER_{camp['nom']}.xlsx", width="stretch")
    if c2.button("Générer le fichier CM", width="stretch"):
        telecharger_cm(df)
