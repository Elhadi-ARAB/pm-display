"""Construire un PM display (Carrefour ou PMU), puis générer l'URL Builder ou le fichier média.

Lancer : streamlit run pm_builder_app.py
"""
import io
import re

import pandas as pd
import streamlit as st

from export.cm import generate_files
from pm_builder import controles, export_pm, frais, importer, modele, pmu_validation, vers_crf, vers_pmu, verif_sortie

st.set_page_config(page_title="PM Display · Carrefour / PMU", page_icon="📋", layout="wide",
                   initial_sidebar_state="expanded")
COULEUR = {"Carrefour": "#2E75B6", "PMU": "#C55A11"}
SORTIE = {"Carrefour": "URL Builder", "PMU": "fichier média"}

st.markdown("""
<style>
.block-container {padding-top: 1.6rem; max-width: 1500px;}
.etape {border-left: 4px solid #D0D7E2; padding: .35rem .7rem; border-radius: 4px; background: #F7F9FC; font-size: .9rem;}
.etape.ok {border-left-color: #2E8B57;}
.etape.ko {border-left-color: #D9822B;}
.etape b {display: block; font-size: .95rem;}
.client {display: inline-block; padding: .15rem .6rem; border-radius: 4px; color: white; font-weight: 600; font-size: .85rem;}
</style>""", unsafe_allow_html=True)


# ------------------------------------------------------------------ état
def charger(camp, lignes, creas, audiences, remarques=None):
    for col in ("Début", "Fin"):
        lignes[col] = pd.to_datetime(lignes[col], errors="coerce")
    if camp.get("client") == "PMU":                       # DV, XPLN, DAR : cases à cocher (remplie = oui)
        from pm_builder.frais import actif
        for col in ("Double Verify", "XPLN", "DAR"):
            if col in lignes.columns:
                lignes[col] = lignes[col].map(actif).astype(bool)
    st.session_state.update(camp=camp, lignes_base=lignes.reset_index(drop=True), creas_base=creas.reset_index(drop=True),
                            aud_base=audiences.reset_index(drop=True), ver=st.session_state.get("ver", 0) + 1,
                            sortie=None, remarques=remarques or [])


def nouveau(client):
    charger(modele.campagne_vide(client), modele.lignes_vides(client), modele.creas_vides(), modele.audiences_vides())


if "camp" not in st.session_state:
    nouveau("Carrefour")


def vide(x) -> bool:
    return x is None or (isinstance(x, float) and pd.isna(x)) or str(x).strip() == ""


# ------------------------------------------------------------------ barre latérale
with st.sidebar:
    st.subheader("PM Display")
    choix = st.segmented_control("Client", modele.CLIENTS, default=st.session_state.camp["client"], key="seg_client")
    if choix and choix != st.session_state.camp["client"]:
        nouveau(choix); st.rerun()
    client_sb = st.session_state.camp["client"]

    if st.button("Nouveau PM vide", width="stretch", type="primary"):
        nouveau(client_sb); st.rerun()

    st.markdown("**Importer le modèle Excel**")
    st.caption(f"Le modèle PM commun rempli dans Excel. Seuls les onglets {client_sb} sont lus.")
    f_imp = st.file_uploader("Modèle PM commun (.xlsx)", type=["xlsx"], key=f"imp_{st.session_state.ver}",
                             label_visibility="collapsed")
    if f_imp is not None and st.button("Importer", width="stretch"):
        try:
            charger(*importer.depuis_modele(f_imp, client_sb)); st.rerun()
        except Exception as e:
            st.error(f"Import impossible : {e}")

    st.markdown("**Voir un exemple**")
    if st.button(f"Exemple {client_sb}", width="stretch"):
        charger(*(modele.exemple_carrefour() if client_sb == "Carrefour" else modele.exemple_pmu())); st.rerun()

camp = st.session_state.camp
client = camp["client"]
L = modele.listes()[client]
v = st.session_state.ver

# ------------------------------------------------------------------ en-tête
st.markdown(f"<span class='client' style='background:{COULEUR[client]}'>{client}</span>", unsafe_allow_html=True)
st.title(camp["nom"] or "Nouveau plan média")
st.caption(f"Remplissez les onglets dans l'ordre. Le dernier onglet vérifie tout et génère le PM et le {SORTIE[client]}.")

if st.session_state.get("remarques"):
    with st.expander(f"À vérifier après l'import ({len(st.session_state.remarques)})", expanded=True):
        for r in st.session_state.remarques:
            st.markdown(f"- {r}")
        if st.button("J'ai vérifié", key="ok_rem"):
            st.session_state.remarques = []; st.rerun()

etapes = ["Campagne", "Lignes", "Créas"] + (["Audiences"] if client == "Carrefour" else []) + ["Vérifier et générer"]
progression = st.container()
_noms_onglets = [f"{i}. {e}" for i, e in enumerate(etapes, 1)] + (["🔎 Valider un fichier média existant"] if client == "PMU" else [])
_onglets = st.tabs(_noms_onglets)
tabs = dict(zip(etapes, _onglets[:len(etapes)]))
onglet_validation = _onglets[-1] if client == "PMU" else None


def afficher_controle(regles):
    """Affiche les règles du fichier généré : ✅ conforme / ❌ à regarder."""
    ok = sum(1 for _, r, _ in regles if r)
    if ok == len(regles):
        st.success(f"Contrôle du fichier généré : les {len(regles)} règles sont respectées.")
    else:
        st.error(f"Contrôle du fichier généré : {len(regles) - ok} règle(s) sur {len(regles)} à regarder avant de l'utiliser.")
    for titre, r, details in regles:
        if r:
            st.markdown(f"✅ {titre}")
        else:
            with st.expander(f"❌ {titre}", expanded=True):
                for d in details:
                    st.markdown(f"- {d}")


def sel(options, label, default=None, aide=None):
    return st.column_config.SelectboxColumn(label, options=options, default=default, help=aide)


idx = lambda opts, val: opts.index(val) if val in opts else None

# ------------------------------------------------------------------ 1. Campagne
with tabs["Campagne"]:
    a, b, c = st.columns([2, 1, 1])
    camp["nom"] = a.text_input("Nom de campagne", camp["nom"], key=f"nom_{v}",
                               help="Carrefour : exactement le nom du PDR (ex. MDC_Sdg_V2). PMU : ex. SEPTEMBRE_PMU_PDV.")
    camp["debut"] = b.date_input("Début", camp["debut"], format="DD/MM/YYYY", key=f"deb_{v}")
    camp["fin"] = c.date_input("Fin", camp["fin"], format="DD/MM/YYYY", key=f"fin_{v}")
    if client == "Carrefour":
        camp["landing"] = st.text_input("Landing page", camp["landing"], key=f"lp_{v}")
        a, b = st.columns(2)
        camp["libelle"] = a.text_input("Nom affiché dans le PM client", camp.get("libelle", ""), key=f"lib_{v}",
                                       help="Ex. MDC PÂQUES. Vide : le nom de campagne, en majuscules.")
        camp["owner_crf"] = b.text_input("Owner Carrefour", camp.get("owner_crf", ""), key=f"own_{v}")
        a, b = st.columns(2)
        camp["pilier"] = a.selectbox("Pilier", L["Pilier"], index=idx(L["Pilier"], camp["pilier"]), key=f"pil_{v}", placeholder="Choisir…")
        camp["objectif"] = b.selectbox("Objectif", L["Objectif"], index=idx(L["Objectif"], camp["objectif"]), key=f"obj_{v}", placeholder="Choisir…")
    else:
        a, b = st.columns([4, 1])
        camp["landing"] = a.selectbox("Landing page (URL)", modele.LP_URLS, index=idx(modele.LP_URLS, camp["landing"]),
                                      key=f"lp_{v}", placeholder="Choisir l'URL dans la liste CM…",
                                      help="Seules les URL de la liste CM (onglet LPS) sont acceptées par l'app automation PMU.")
        b.text_input("Landing Page ID", modele.LP_ID.get(camp["landing"], ""), disabled=True, key=f"lpid_{v}_{camp['landing']}")

        cols = st.columns(5)
        for col, (k, lab, opts) in zip(cols, (("levier", "Levier par défaut", L["Levier campagne"]), ("hs", "H/S", L["H/S"]),
                                              ("univers", "85/2010", L["85/2010"]), ("medium", "Medium", L["Medium"]),
                                              ("promo", "Promo/Offre", L["Promo/Offre"]))):
            camp[k] = col.selectbox(lab, opts, index=idx(opts, camp.get(k)), key=f"{k}_{v}", placeholder="Choisir…")
        a, b, c = st.columns(3)
        camp["consultant"] = a.text_input("Consultant email *", camp.get("consultant", ""), key=f"cons_{v}",
                                          help="Obligatoire : l'app automation PMU le vérifie (cellule B1 du fichier média).")
        camp["numero_projet"] = b.text_input("Numéro de projet", camp.get("numero_projet", ""), key=f"proj_{v}")
        c.text_input("Advertiser CM", modele.ADVERTISER_PMU, disabled=True, key=f"adv_{v}")
        camp["advertiser"] = modele.ADVERTISER_PMU
        st.caption("Chaque ligne prend le levier BRAND ou CONSID selon sa section (Branding / Considération) : "
                   "un fichier média est généré par levier. Le levier par défaut sert aux autres sections.")

# ------------------------------------------------------------------ 2. Lignes
with tabs["Lignes"]:
    st.caption("Une ligne par régie et par format. Ajoutez une ligne avec le bouton + sous le tableau ; "
               "vous pouvez aussi coller plusieurs cellules copiées depuis Excel.")
    cfg = {
        "N° ligne": st.column_config.TextColumn("N°", disabled=True, width="small"),
        "Section": sel(modele.SECTIONS, "Section", "Branding"),
        "Levier": sel(L["Levier"], "Levier"), "Régie": sel(L["Régie"], "Régie"),
        "Type d'achat": sel(L["Type d'achat"], "Achat"),
        "Format": st.column_config.TextColumn("Format", help="Libellé du format pour le client."),
        "Type de format": sel(L["Type de format"], "Type de format"),
        "Ciblage": st.column_config.TextColumn("Ciblage", help="PMU : libellé court, il entre dans les noms CM."),
        "Début": st.column_config.DateColumn("Début", format="DD/MM/YYYY", default=camp["debut"]),
        "Fin": st.column_config.DateColumn("Fin", format="DD/MM/YYYY", default=camp["fin"]),
        "Budget net (€)": st.column_config.NumberColumn("Budget net (€)", min_value=0, format="%.2f"),
        "CPM (€)": st.column_config.NumberColumn("CPM (€)", min_value=0, format="%.2f"),
    }
    if client == "Carrefour":
        cfg.update({"Groupe d'audiences": st.column_config.TextColumn("Groupe d'audiences", help="Les formats d'un même groupe partagent les audiences."),
                    "Nb messages": st.column_config.NumberColumn("Nb messages", min_value=1, step=1, default=1),
                    "Durée (s)": st.column_config.NumberColumn("Durée (s)", min_value=0, step=1),
                    "Tracking Type": sel(L["Tracking Type"], "Tracking", "Pixel+Click-Command",
                                         "PCC → DISPLAY dans le fichier CM ; VAST (Video-Ad-Serving-Template) → VIDEO."),
                    "Ciblage": st.column_config.TextColumn("Ciblage (PM client)", help="Texte des audiences pour le client. "
                                                           "Vide : les audiences du groupe avec leur %."),
                    "Device": st.column_config.TextColumn("Device", default="Multidevice"),
                    "Specs": st.column_config.TextColumn("Specs"),
                    "KPI": st.column_config.TextColumn("KPI", help="Ex. Reach/Completion."),
                    "VTR / LTR": st.column_config.NumberColumn("VTR / LTR", min_value=0, max_value=1, format="%.2f",
                                                               help="Taux de complétion (0,79 = 79 %). Sert au calcul des vues et du CPVV."),
                    "Fréquence": st.column_config.NumberColumn("Fréquence", min_value=0, help="Sert au calcul du reach."),
                    "Études / BLS": st.column_config.TextColumn("Études / BLS")})
    else:
        cfg.update({"Ciblage": st.column_config.TextColumn("Ciblage (noms CM)", help="Libellé court, il entre dans les noms CM (ex. Turfistes reguliers)."),
                    "Ciblage PM client": st.column_config.TextColumn("Ciblage (PM client)", help="Libellé complet affiché dans le PM client. Vide : le ciblage court."),
                    "Objectif": st.column_config.TextColumn("Objectif", help="Ex. Visibilité, Reach & Visibilité, CTR."),
                    "Emplacement": st.column_config.TextColumn("Emplacement", help="Cadre de diffusion. Vide : « RG sur site »."),
                    "Device": sel(L["Device"], "Device", "MULTIDEVICE"), "Programmation": sel(L["Programmation"], "Programmation"),
                    "Support": st.column_config.SelectboxColumn("Support", options=modele.SITE_NOMS,
                                                                help="Site CM (onglet Sites). Vide : rempli automatiquement selon le levier."),
                    "Type de tracking": sel(L["Type de tracking"], "Tracking", "PCC"),
                    "Double Verify": st.column_config.CheckboxColumn("Double Verify", default=False,
                                                                     help="Coché = la ligne a du DV. Coût = impressions × 0,09 / 1000 × 1,25."),
                    "XPLN": st.column_config.CheckboxColumn("XPLN", default=False,
                                                            help="Coché = la ligne a du XPLN. Coût = impressions × 0,2 / 1000 × 1,25."),
                    "DAR": st.column_config.CheckboxColumn("DAR", default=False,
                                                           help="Coché = la ligne entre dans le DAR de sa section : 800 € jusqu'à 15 M "
                                                                "d'impressions, puis + (impressions − 15 M) × 0,052 / 1000."),
                    "Créative": st.column_config.TextColumn("Créative", help="Vide : TBD."),
                    "Commentaire": st.column_config.TextColumn("Commentaire")})
    lignes = st.data_editor(st.session_state.lignes_base, key=f"ed_l_{v}", num_rows="dynamic", hide_index=True,
                            width="stretch", column_config=cfg)
    lignes = modele.numeroter(lignes)
    if client == "PMU":
        lignes, notes_pmu = modele.completer_pmu(lignes)
        st.caption("Double Verify, XPLN, DAR : case cochée = la ligne a ce frais. Remplis automatiquement s'ils sont vides : "
                   "Programmation et Support (VOL → 79 + UN1T, SOCIAL → 79 + FACEBOOK, DISPLAY → Régie + site CM), "
                   "Emplacement (RG sur site), Créative (TBD).")
        for n_ in notes_pmu:
            st.warning(n_)
    bud = pd.to_numeric(lignes["Budget net (€)"], errors="coerce"); cpm = pd.to_numeric(lignes["CPM (€)"], errors="coerce")
    m1, m2, m3 = st.columns(3)
    m1.metric("Lignes du PM", len(lignes))
    m2.metric("Budget net", f"{bud.sum():,.2f} €".replace(",", " "))
    m3.metric("Volume estimé", f"{(bud / cpm * 1000).sum():,.0f}".replace(",", " "))
    if client == "Carrefour" and len(lignes):
        fc = frais.synthese_crf(lignes)
        eurc = lambda x: f"{x:,.2f} €".replace(",", " ")
        st.markdown("**Fees DV360 + taxe Google** (YouTube et PG : 5,6 % · reste du DV360 : 11,5 % · gré à gré : 0)")
        c1, c2 = st.columns(2)
        c1.metric("Fees", eurc(fc["fees"])); c2.metric("Budget net média HT (+ frais tech)", eurc(fc["total"]))
    if client == "PMU" and len(lignes):
        fr = frais.synthese(camp, lignes)
        eur = lambda x: f"{x:,.2f} €".replace(",", " ")
        st.markdown("**Frais calculés à partir des impressions**")
        f1, f2, f3, f4, f5 = st.columns(5)
        f1.metric("Fees Meta (3 %)", eur(fr["meta"])); f2.metric("Double Verify", eur(fr["dv"]))
        f3.metric("XPLN", eur(fr["xpln"])); f4.metric("DAR", eur(sum(fr["dar"].values())),
                                                    help=" · ".join(f"{k} : {eur(x)}" for k, x in fr["dar"].items()))
        f5.metric("Total budget campagne", eur(fr["total"]), help=f"Net média {eur(fr['net'])} + frais {eur(fr['frais'])}")

ids = [i for i in lignes["N° ligne"].astype(str).tolist() if i]
apercu_lignes = {str(r["N° ligne"]): f"{r['N° ligne']} · {r['Régie'] or '?'} · {r['Format'] or ''}" for _, r in lignes.iterrows()}

# ------------------------------------------------------------------ 3. Créas
with tabs["Créas"]:
    if client == "PMU":
        st.caption("Une ligne par taille de format : chacune donne une ligne du fichier média.")
    else:
        st.caption("Seulement pour une ligne du PM avec plusieurs messages créatifs (ex. 3 audios Spotify) : une ligne par message.")
    with st.expander("N° des lignes du PM", expanded=False):
        st.markdown("\n".join(f"- {t}" for t in apercu_lignes.values()) or "Aucune ligne pour l'instant.")
    cfg_c = {"N° ligne": sel(ids, "N° ligne"),
             "Taille": st.column_config.TextColumn("Taille", help="Écrire 300x250."),
             "Code message": st.column_config.TextColumn("Code message", help="Carrefour : code court (ex. PDT), sans « _ » ni espace.")}
    if client == "PMU":
        cfg_c["Landing (si différente)"] = st.column_config.SelectboxColumn("Landing (si différente)", options=modele.LP_URLS,
                                                                            help="Une LP de la liste CM, si différente de celle de la campagne.")
        cfg_c["Code message"] = None
    else:
        cfg_c["Landing (si différente)"] = st.column_config.LinkColumn("Landing (si différente)")
    creas = st.data_editor(st.session_state.creas_base, key=f"ed_c_{v}", num_rows="dynamic", hide_index=True, width="stretch",
                           column_config=cfg_c)
    if client == "PMU" and len(creas):
        apercu = vers_pmu.generer_par_levier(camp, lignes, creas)
        if apercu:
            st.markdown("**Aperçu des lignes du fichier média**")
            st.dataframe(pd.concat([d for _, d in apercu.values()], ignore_index=True)[
                ["Régie", "Format", "Piste Créative", "PLACEMENT NAME DCM", "Créative DCM"]], hide_index=True, width="stretch")

# ------------------------------------------------------------------ 4. Audiences (Carrefour)
audiences = st.session_state.aud_base
if client == "Carrefour":
    with tabs["Audiences"]:
        st.caption("Une ligne par audience, rattachée au groupe d'audiences des lignes. Les % d'un groupe font 100.")
        groupes = sorted({str(g) for g in lignes["Groupe d'audiences"].dropna() if str(g).strip()})
        audiences = st.data_editor(st.session_state.aud_base, key=f"ed_a_{v}", num_rows="dynamic", hide_index=True, width="stretch",
                                   column_config={"Groupe d'audiences": sel(groupes, "Groupe d'audiences"),
                                                  "Type de ciblage": sel(L["Type de ciblage"], "Type de ciblage"),
                                                  "Targeting details": sel(L["Targeting details"], "Targeting details"),
                                                  "% du groupe": st.column_config.NumberColumn("% du groupe", min_value=0, max_value=100)})
        if len(audiences):
            tot = audiences.groupby("Groupe d'audiences")["% du groupe"].apply(lambda s: pd.to_numeric(s, errors="coerce").sum())
            st.caption("Total par groupe : " + " · ".join(f"{g} {t:g} %" + ("" if abs(t - 100) < .5 else " ⚠️") for g, t in tot.items()))

# ------------------------------------------------------------------ contrôles (utilisés par la progression et le dernier onglet)
err, avert = controles.verifier(camp, lignes, creas, audiences)
n_sortie = controles.nb_lignes_sortie(camp, lignes, creas, audiences)
famille = lambda e: ("Campagne" if e.startswith("Campagne") else "Lignes" if e.startswith("Ligne") else
                     "Créas" if e.startswith("Créa") else "Audiences")
par_etape = {e: [x for x in err if famille(x) == e] for e in ("Campagne", "Lignes", "Créas", "Audiences")}
rempli = {"Campagne": not vide(camp["nom"]), "Lignes": len(lignes) > 0,
          "Créas": len(creas) > 0 or client == "Carrefour", "Audiences": len(audiences) > 0}
with progression:
    cols = st.columns(len(etapes))
    for col, e in zip(cols, etapes):
        if e == "Vérifier et générer":
            ok, txt = not err, ("Prêt à générer" if not err else f"{len(err)} point(s) à corriger")
        else:
            ok = rempli.get(e, True) and not par_etape.get(e)
            txt = ("Complet" if ok else ("À remplir" if not rempli.get(e, True) else f"{len(par_etape[e])} point(s)"))
        col.markdown(f"<div class='etape {'ok' if ok else 'ko'}'><b>{e}</b>{txt}</div>", unsafe_allow_html=True)

# ------------------------------------------------------------------ 5. Vérifier et générer
with tabs["Vérifier et générer"]:
    regles = [("En-tête de campagne complet", par_etape["Campagne"]),
              ("Lignes complètes et valeurs des listes", par_etape["Lignes"]),
              ("Créas : tailles et rattachements" if client == "PMU" else "Messages créatifs", par_etape["Créas"])]
    if client == "Carrefour":
        regles.append(("Audiences : listes et % à 100 par groupe", par_etape["Audiences"]))

    st.subheader("Récapitulatif")
    a, b, c, d = st.columns(4)
    a.metric("Lignes du PM", len(lignes))
    b.metric("Budget net", f"{bud.sum():,.0f} €".replace(",", " "))
    c.metric(f"Lignes du {SORTIE[client]}", n_sortie)
    if client == "PMU":
        b.metric("Total budget (avec frais)", f"{frais.synthese(camp, lignes)['total']:,.0f} €".replace(",", " "),
                 help=f"Dont net média {bud.sum():,.0f} €".replace(",", " "))
        leviers = lignes["Section"].map(lambda x: modele.LEVIER_PAR_SECTION.get(str(x))).fillna(camp.get("levier") or "?")
        d.metric("Fichiers média", len(set(leviers)), help=", ".join(sorted(set(leviers))))
    else:
        d.metric("Audiences", len(audiences))

    st.markdown("**Règles vérifiées**")
    for titre, pb in regles:
        if pb:
            with st.expander(f"❌ {titre} ({len(pb)})", expanded=True):
                for x in pb:
                    st.markdown(f"- {x}")
        else:
            st.markdown(f"✅ {titre}")
    for w in avert:
        st.warning(w)

    st.divider()
    nom_f = (camp["nom"] or "campagne").replace(" ", "_")
    g1, g2 = st.columns(2)
    with g1:
        st.markdown("**Le PM**")
        st.caption("Le PM à envoyer au client : mise en forme du devis PMU avec les frais (Meta, DV, XPLN, DAR) et le total budget."
                   if client == "PMU" else "Le PM à envoyer au client : mise en forme du PM Carrefour (GLOBAL + PM DISPLAY), "
                   "avec les fees DV360 + taxe Google.")
        st.download_button("Télécharger le PM (Excel)", export_pm.exporter(camp, lignes, creas, audiences),
                           file_name=f"PM_{client}_{nom_f}.xlsx", width="stretch", disabled=bool(err),
                           help="Disponible quand toutes les règles sont respectées." if err else None)
    with g2:
        st.markdown(f"**Le {SORTIE[client]}**")
        st.caption("Généré à partir du PM, prêt pour " + ("l'app CM et la création DV360." if client == "Carrefour" else "l'app Streamlit PMU."))
        if st.button(f"Générer le {SORTIE[client]}", type="primary", width="stretch", disabled=bool(err)):
            try:
                if client == "Carrefour":
                    df, a_verif = vers_crf.generer(camp, lignes, creas, audiences)
                    st.session_state.sortie = {"client": client, "df": df, "a_verif": a_verif}
                else:
                    par_lev = vers_pmu.generer_par_levier(camp, lignes, creas)
                    st.session_state.sortie = {"client": client, "par_levier": par_lev, "a_verif": []}
            except Exception as e:
                st.error(f"Génération impossible : {e}")
    if err:
        st.info("Corrigez les points en rouge ci-dessus : les téléchargements s'activent quand tout est conforme.")

    s = st.session_state.get("sortie")
    if s and s["client"] == client:
        st.divider()
        st.subheader(f"{SORTIE[client].capitalize()} généré")
        for x in s["a_verif"]:
            st.warning(x)
        if client == "Carrefour":
            df = s["df"]
            m1, m2, m3 = st.columns(3)
            m1.metric("Lignes générées", len(df)); m2.metric("Régies", df["Régie"].nunique())
            m3.metric("Noms de créa", df["Creative Name"].nunique())
            c_reg, c_ctl = st.columns([1, 2])
            with c_reg:
                st.markdown("**Lignes par régie**")
                st.dataframe(verif_sortie.recap_regies(df), hide_index=True, width="stretch")
            with c_ctl:
                afficher_controle(verif_sortie.url_builder(df, camp, lignes, creas, audiences))
            st.markdown("**Aperçu**")
            st.dataframe(df[["Platform", "Régie", "Format Type", "Targeting details", "Creative Name", "Placement CM = creative DV360"]],
                         hide_index=True, width="stretch")
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                df.to_excel(w, sheet_name="URL BUILDER", index=False, startrow=1)
            a, b = st.columns(2)
            a.download_button("Télécharger l'URL Builder", buf.getvalue(), file_name=f"URL_BUILDER_{nom_f}.xlsx", width="stretch")
            if b.button("Générer le fichier CM", width="stretch"):
                paths, ign, unk = generate_files(df)
                if ign or unk:
                    st.error(f"Fichier CM non généré : plateformes refusées {ign}, tracking refusés {unk}.")
                for p in paths:
                    b.download_button(f"Télécharger {p.split('/')[-1]}", open(p, "rb").read(), file_name=p.split("/")[-1], width="stretch")
        else:
            for lev, (c_lev, d_lev) in s["par_levier"].items():
                st.markdown(f"#### Fichier média {lev}")
                st.caption(f"Campagne CM : {vers_pmu.nom_campagne_cm(c_lev)} · Advertiser : {modele.ADVERTISER_PMU}")
                m1, m2, m3 = st.columns(3)
                m1.metric("Lignes générées", len(d_lev)); m2.metric("Régies", d_lev["Régie"].nunique())
                m3.metric("Landing pages", d_lev["Landing Page"].nunique())
                c_reg, c_ctl = st.columns([1, 2])
                with c_reg:
                    st.markdown("**Lignes par régie**")
                    st.dataframe(verif_sortie.recap_regies(d_lev), hide_index=True, width="stretch")
                    st.markdown("**Landing pages utilisées**")
                    st.dataframe(verif_sortie.landing_ids(d_lev), hide_index=True, width="stretch")
                with c_ctl:
                    afficher_controle(verif_sortie.fichier_media(d_lev, lev, camp, lignes, creas))
                st.markdown("**Aperçu**")
                st.dataframe(d_lev[["Régie", "Format", "Piste Créative", "PLACEMENT NAME DCM", "Créative DCM", "Landing Page"]],
                             hide_index=True, width="stretch")
                st.download_button(f"Télécharger le fichier média {lev}", vers_pmu.en_excel(c_lev, d_lev),
                                   file_name=f"Fichier_media_{vers_pmu.nom_campagne_cm(c_lev)}.xlsx", width="stretch", key=f"dl_{lev}")
                n_rd = int((d_lev["Type de Tracking"].astype(str).str.upper() == "RD").sum())
                with st.expander(f"Renommer les créas reçues ({n_rd} ligne(s) RD dans ce fichier)", expanded=False):
                    st.caption("Comme dans l'app automation PMU : seules les lignes RD ont une créa. Chaque image ou vidéo est "
                               "reconnue par sa dimension et copiée une fois par Créative DCM de cette dimension (une par support).")
                    if n_rd == 0:
                        st.info("Aucune ligne RD : rien à renommer (les lignes PCC n'ont pas de créa).")
                    else:
                        fichiers = st.file_uploader("Créas reçues (jpg, png, gif, mp4…)", accept_multiple_files=True,
                                                    type=["jpg", "jpeg", "png", "gif", "mp4", "mov", "webm", "avi"], key=f"cr_{lev}")
                        if fichiers and st.button("Renommer les créas", key=f"ren_{lev}"):
                            zipb, ok, absentes, manuel = pmu_validation.renommer_creas([(f.name, f.getvalue()) for f in fichiers], d_lev)
                            if ok:
                                st.success(f"{len(ok)} créa(s) reçue(s) → {sum(len(n) for _, n, _ in ok)} fichier(s) renommé(s).")
                                for ancien, noms, dim in ok:
                                    st.markdown(f"- `{ancien}` ({dim}) → " + ", ".join(f"`{n}`" for n in noms))
                                st.download_button("Télécharger les créas renommées (ZIP)", zipb, file_name=f"creas_{lev}.zip", key=f"zip_{lev}")
                            for nom_f, dim in absentes:
                                st.error(f"`{nom_f}` : dimension {dim} absente des lignes RD de ce fichier.")
                            for nom_f, raison in manuel:
                                st.warning(f"`{nom_f}` : {raison}, à renommer à la main.")
            st.caption("Même disposition que le fichier média actuel (avec l'onglet LPS) : il se dépose tel quel dans l'app Streamlit PMU.")


# ------------------------------------------------------------------ PMU : valider un fichier média existant
if onglet_validation is not None:
    with onglet_validation:
        st.caption("Pour un fichier média rempli à la main : mêmes règles que l'app automation PMU (en-têtes, consultant email, "
                   "nom de campagne CM, advertiser, supports, landing pages, tracking PCC/RD, dimension des lignes RD). "
                   "Les URL sont nettoyées automatiquement ; le reste se corrige dans le fichier source.")
        f_val = st.file_uploader("Fichier média PMU (.xlsx)", type=["xlsx"], key="val_pmu")
        if f_val is not None and st.button("Valider le fichier", type="primary"):
            contenu = f_val.getvalue()
            try:
                r = pmu_validation.valider_fichier(contenu)
            except Exception as e:
                st.error(f"Lecture impossible : {e}")
                r = None
            if r:
                a, b = st.columns(2)
                a.metric("Lignes lues", r["nb_lignes"]); b.metric("Erreurs bloquantes", len(r["erreurs"]))
                if r["corrections_url"]:
                    with st.expander(f"Corrections d'URL appliquées ({len(r['corrections_url'])})"):
                        for c in r["corrections_url"]:
                            st.markdown(f"**Ligne {c['ligne']}**"); st.code(c["avant"], language="text"); st.code(c["apres"], language="text")
                else:
                    st.info("Aucune correction d'URL nécessaire.")
                if r["erreurs"]:
                    st.error(f"{len(r['erreurs'])} erreur(s) à corriger dans le fichier source :")
                    for e in r["erreurs"]:
                        st.markdown(f"- {e}")
                else:
                    st.success("Fichier valide : prêt pour l'upload GCP.")
                    nom = re.sub(r"\s+", "_", re.sub(r'[\\/:*?"<>|]+', "_", r["nom_campagne"] or "fichier_media"))
                    st.download_button("Télécharger le fichier propre", pmu_validation.fichier_propre(contenu, r["df_corrige"]),
                                       file_name=f"{nom}.xlsx")
