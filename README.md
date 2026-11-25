# PM Display & URL Builder (Carrefour, PMU)

## Nouveau : constructeur de PM (`pm_builder_app.py`)

```bash
streamlit run pm_builder_app.py
```

L'expert média construit son PM directement dans l'app, pour Carrefour ou PMU :
1. **Campagne** : en-tête commun + champs du client (pilier/objectif pour Carrefour ; levier, H/S, 85/2010, medium, promo/offre pour PMU).
2. **Lignes** : une ligne par régie et par format, listes déroulantes propres au client, valeurs par défaut.
3. **Créas** : une ligne par taille (PMU) ou par message (Carrefour).
4. **Audiences** : Carrefour uniquement, une ligne par audience rattachée à un groupe.
5. **Contrôles et exports** : erreurs bloquantes en direct, PM Excel pour le client, brouillon à reprendre,
   puis URL Builder + fichier CM (Carrefour) ou fichier média (PMU, même disposition que l'actuel).

Les noms (placements, créas, URL) sont calculés par le code : plus de formules à étendre.
Saisie rapide : remplir le modèle Excel v2 (un onglet par client), puis « Importer un fichier pour pré-remplir ».
Test : `python -m tests.test_pm_builder URL_BUILDER_MDC.xlsx Fichier_media_PMU.xlsx`
(résultat actuel : Carrefour 44/44, PMU 35/35 noms identiques aux fichiers réels).

| Dossier | Rôle |
|---|---|
| `pm_builder/modele.py` | Champs par client, listes, exemples |
| `pm_builder/controles.py` | Contrôles de saisie |
| `pm_builder/vers_crf.py` | PM Carrefour -> URL Builder (réutilise `generation/`) |
| `pm_builder/vers_pmu.py` | PM PMU -> fichier média (reproduit les formules actuelles) |
| `pm_builder/export_pm.py` | Export du PM Excel pour le client |

---

# URL Builder & CM Carrefour (lecture d'un PM existant : `app.py`)

Application Streamlit qui génère l'URL Builder d'une campagne Carrefour à partir du plan média (PM) et du PDR, puis le fichier d'import CM.

## Lancer

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Parcours

1. **Déposer** le PM (onglet PM DISPLAY) et le PDR, choisir la campagne.
2. **Vérifier la fiche setup** : formats, audiences et messages sont pré-remplis grâce aux tables de correspondance. Les libellés jamais vus ont la case « Connu » décochée et des cases vides à compléter dans les listes déroulantes.
3. **Contrôler** : valeurs dans les listes, pas de « _ » dans les codes, nombre de lignes cohérent, Tracking Type accepté par l'export CM, % de budget à 100 %.
4. **Générer** l'URL Builder (Excel) puis le fichier CM.

Le mode « Depuis un URL Builder existant » reprend l'ancienne app : il génère le fichier CM à partir d'un URL Builder rempli à la main.

## Organisation

| Dossier | Rôle |
|---|---|
| `lecture/` | Lecture du PM (découpage formats / audiences) et du PDR |
| `tables/` | Tables de correspondance (YAML) et listes déroulantes d'Inputs A79 |
| `generation/` | Fiche setup, génération de l'URL Builder, contrôles, règles de nommage |
| `export/cm.py` | Export CM existant, repris tel quel |
| `tests/test_mdc.py` | Test de référence sur MDC SDG V2 |

## Statut des valeurs

Chaque ligne de la fiche setup a un statut :
- **validé** : trouvé dans les tables, convention confirmée par un expert ;
- **proposé** : trouvé dans les tables, mais la correspondance est une proposition (`a_valider: true` dans le YAML) ;
- **inconnu** : rien trouvé, cases vides à compléter.

Pour passer une proposition en validé : retirer `a_valider: true` de l'entrée dans le YAML.

## Type de ciblage lu dans le PM

Le PM fait foi. Si le libellé d'une audience contient un mot de `tables/types_ciblage.yaml` (1st party, broad, intérêts, contextuel), ce mot décide du Targeting Type ; sinon, c'est la valeur de la table des audiences. La colonne « Origine Targeting Type » de l'app indique laquelle des deux a été utilisée.

## Blocs

Chaque ligne du PM qui contient des formats forme un bloc (ex. `TF1 · L17`). Les formats d'un bloc sont croisés avec les audiences du même bloc. Une ligne sans audiences reprend celles du bloc précédent de la même régie (ex. ligne Shorts) ; une ligne avec des audiences mais sans format complète le bloc précédent.

## Fichiers acceptés

- **PM** : tout onglet dont le nom contient « display » et qui a un en-tête « REGIE/PARTENAIRE ». Les cellules fusionnées sont gérées, ainsi que les en-têtes « AUDIENCES » ou « CIBLAGES », « DEBUT » ou « DATE DEBUT ». Les lignes de total, de section, d'exemple et les tableaux de répartition budgétaire sont ignorés.
- **PDR** : le plan de roulement global (onglet DISPLAY_A79, colonnes lues par leur nom) ou un extrait sans en-tête (colonnes lues par position).

## Ajouter une règle

Quand un expert complète un libellé inconnu, ajouter le mot-clé dans le YAML concerné (`audiences.yaml`, `formats.yaml`, `regies.yaml`, `divers.yaml`). Les mots-clés sont comparés sans accents ni majuscules, en début de mot. L'ordre des entrées compte : la première qui correspond gagne.

## Règle de nommage

Chaque champ est nettoyé avant assemblage : accents retirés, tout caractère hors lettres, chiffres et tiret remplacé par un tiret. Le « _ » est réservé au séparateur. Le placement CM (= utm_content) a donc toujours 10 morceaux, et le nettoyage UTM de l'export CM ne le modifie plus. Le placement utilise la plateforme de la ligne (col. Q) et non la colonne A.

## Test de référence

```bash
python -m tests.test_mdc PM.xlsx PDR.xlsx URL_BUILDER_reference.xlsx
```

Compare les colonnes de saisie (Q à AA, Creative Name, URL, nom de campagne) avec l'URL Builder rempli à la main. Résultat actuel sur MDC SDG V2 : 44 lignes sur 44 identiques, avec l'extrait de PDR comme avec le plan de roulement global.

## Test sur plusieurs campagnes

```bash
python -m tests.test_references PDR_global.xlsx "PM.xlsx::ONGLET::URL_BUILDER.xlsx" ...
```

Compare l'app à des URL Builders remplis à la main : nombre de lignes, accord sur les lignes d'achat (Platform, Régie, Format Type) et sur les audiences (Régie, Targeting Type, Targeting details). Résultats actuels : MDC SDG V2 et FAV Automne 100 %, RDC V2 94 %, BDP V5 100 % / 68 %, MDC Pâques 74 % / 67 %, TTF1 48 % / 76 %. Les écarts restants viennent surtout de conventions qui varient d'une campagne à l'autre (voir le fichier Conventions_URL_Builder_a_trancher.xlsx).

## Limites connues (POC)

- L'URL Builder est exporté en valeurs, sans les formules ni les listes déroulantes du modèle original.
- Les tables sont pré-remplies à partir de MDC uniquement, en attente de validation par les experts.
- Les nouvelles règles validées dans l'app ne sont pas encore enregistrées automatiquement dans les tables.
- La colonne XPLN script CM et le lien créa (col. K) restent vides.
