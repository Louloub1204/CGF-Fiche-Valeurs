# Fiche valeur BRVM

Application d'aide à la décision pour les actions cotées à la BRVM. À partir du classeur des
états financiers des sociétés (`Etats_Financiers_2021_2025.xlsx`), elle produit pour chaque valeur :

- une **fiche signalétique PDF** imprimable (5 pages, ou 1 page en version synthétique) ;
- un **rapport PowerPoint** de 12 diapositives (graphiques natifs, modifiables) ;
- un **classeur Excel** reprenant tous les calculs ;
- un **signal argumenté** : `ACHAT`, `CONSERVER`, `ATTENDRE`, `ALLÉGER` ou `VENDRE`, avec niveau de
  conviction, zone d'accumulation et niveau d'invalidation.

L'analyse croise les états financiers (SIG, ratios, solidité), la valorisation (multiples, Gordon,
actualisation des dividendes), l'analyse technique du titre, l'environnement macroéconomique
(UEMOA, BCEAO, FMI) et les actualités de l'émetteur. Les données de marché sont récupérées en ligne
quand c'est possible, sinon lues dans une base de marché locale ou saisies à la main.

**Tous les titres de la BRVM sont analysables** (48 valeurs dans le référentiel), y compris ceux qui
ne figurent pas dans votre classeur : l'analyse s'adapte aux données disponibles (voir « Niveaux
d'analyse »). Un tableau de bord multi-valeurs traite toute la cote en une fois et exporte les signaux
et les fiches dans une archive.

---

## Démarrage rapide

```bash
git clone https://github.com/<votre-compte>/fiche-valeur-brvm.git
cd fiche-valeur-brvm
python -m venv .venv && source .venv/bin/activate      # Windows : .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

L'application s'ouvre sur http://localhost:8501. Chargez votre classeur des états financiers dans
la barre latérale, choisissez une valeur parmi les 48 titres de la cote, puis cliquez sur **Lancer
l'analyse**. Sans classeur, les titres sont analysés sur chiffres clés ou données de marché lues en
ligne. L'option **Compléter avec les données en ligne** active les sources listées plus bas, et
**Utiliser les classeurs de démonstration** ouvre six sociétés fictives pour découvrir l'outil.
Chaque champ de la barre latérale est expliqué (icône d'aide et guide « Quelles données fournir ? »).

Python 3.10 ou plus récent est requis.

### En ligne de commande

```bash
# Une valeur
python cli.py --ef Etats_Financiers_2021_2025.xlsx --ticker SNTS

# Un titre hors classeur, sans fichier (chiffres clés et cours lus en ligne)
python cli.py --ticker BOAB

# Toute la cote (titres du classeur et titres hors classeur), avec la base de marché, le PowerPoint et le détail Excel
python cli.py --ef Etats_Financiers_2021_2025.xlsx --smf "Base de données - SMF.xlsx" --all --pptx --xlsx

# Démonstration hors ligne
python cli.py --ef sample_data/demo_etats_financiers.xlsx --smf sample_data/demo_base_marche.xlsx --all --hors-ligne
```

Options : `--compact` (fiche d'une page), `--hors-ligne`, `--rf 6.5` (taux sans risque, en pourcentage),
`--prime 5.5` (prime de risque actions), `--g 3` (croissance perpétuelle),
`--maison "CGF Gestion"`, `--out sorties`. Les fichiers sont écrits dans `sorties/` avec un
`Tableau_des_signaux.xlsx` récapitulatif.

---

## Déploiement sur GitHub et Streamlit Community Cloud

1. Créez un dépôt GitHub (privé conseillé) et poussez le projet :
   ```bash
   git init && git add . && git commit -m "Fiche valeur BRVM"
   git branch -M main
   git remote add origin https://github.com/<votre-compte>/fiche-valeur-brvm.git
   git push -u origin main
   ```
   Le `.gitignore` exclut vos classeurs `Etats_Financiers*.xlsx` et `Base*de*donn*.xlsx` : vos
   données restent sur votre poste et sont chargées dans l'application au moment de l'analyse.
2. Sur https://share.streamlit.io, choisissez **Create app**, sélectionnez le dépôt, la branche
   `main` et le fichier `app.py`, puis déployez.
3. Facultatif : pour les commentaires rédigés par Claude, ajoutez dans **Settings > Secrets** :
   ```toml
   ANTHROPIC_API_KEY = "sk-ant-..."
   ```
   En local, placez la même ligne dans `.streamlit/secrets.toml` (fichier ignoré par git).

Le workflow `.github/workflows/tests.yml` exécute les tests à chaque push.

---

## Niveaux d'analyse selon les données disponibles

Les titres n'ont pas tous la même information. Plutôt que d'écarter ceux qui sont incomplets,
l'application mesure ce qui est disponible, fait tourner les moteurs sur ce qui existe, neutralise
le reste et plafonne la conviction du signal.

| Niveau | Données | Ce qui est évalué | Garde-fous sur le signal |
|---|---|---|---|
| **1 - Complet** | États financiers du classeur (3 exercices ou plus, bilan inclus) | SIG, ratios, solidité, croissance, pairs, valorisation (PER, P/B, Gordon, dividendes, VE/EBE), technique, macro, actualités | Aucun |
| **1b - Complet avec lacunes** | Classeur incomplet (exercices ou postes manquants) | Idem, sur les données présentes | Conviction plafonnée à Moyenne si elle était Forte |
| **2 - Chiffres clés** | Chiffre d'affaires, résultat net, BNPA, dividendes (lus sur la fiche société sikafinance ou saisis dans l'application), capitaux propres et total bilan facultatifs | Croissance, marge nette, PER, rendement, valorisation (PER, dividendes, rendement de référence), ROE et P/B si fonds propres saisis, technique, macro, actualités | Conviction plafonnée à Moyenne ; solidité financière non évaluée |
| **3 - Marché seul** | Cours, volumes, dividendes | Technique, valorisation par les dividendes (actualisation et rendement de référence), macro, actualités | Vue fondamentale neutre par construction ; conviction Faible ; VENDRE ramené à ALLÉGER ; ACHAT seulement si potentiel de 25% ou plus et tendance haussière |

Chaque fiche affiche son niveau, un **taux de couverture des données** (de 0 à 100%) et le tableau
« Données disponibles et compléments à apporter » : pour chaque donnée absente, l'effet sur
l'analyse et le moyen de la compléter. Sans comparables suffisants dans le classeur, les multiples
sont remplacés par des **multiples de référence par secteur** (`SECTOR_MULTIPLES` dans `config.py`,
modifiables et signalés comme tels dans la fiche).

### Ajouter un titre

Le référentiel (`fiche_valeur/config.py`) compte 48 valeurs. Pour une introduction récente, utilisez
« Ajouter ou mettre à jour un titre » dans la barre latérale, ou « Rechercher les nouveaux titres »
qui compare la cote publiée sur brvm.org au référentiel.

---

## Fichiers d'entrée

### 1. Classeur des états financiers (obligatoire)

Une feuille par société, nommée par son ticker BRVM (`SNTS`, `SGBC`, `ORAC`...). Dans chaque
feuille :

- une ligne d'en-tête `Poste | 2021 | 2022 | 2023 | 2024 | 2025` (le libellé « Poste » est facultatif,
  deux années au moins suffisent) ;
- l'unité dans le titre de la feuille ou au-dessus de l'en-tête (« en millions FCFA », « en milliers
  de FCFA », « en FCFA », « USD »...) ;
- des intertitres sans valeur qui délimitent les sections : compte de résultat, bilan, hors bilan,
  ratios, notes ;
- une ligne par poste, avec ou sans code SYSCOHADA (`XB - Chiffre d'affaires`, `TA Ventes de
  marchandises`...).

Les libellés sont reconnus par règles normalisées (accents, ponctuation, codes et parenthèses
ignorés), ce qui couvre les présentations bancaires (PNB, résultat brut d'exploitation, coût du
risque), SYSCOHADA (chiffre d'affaires, valeur ajoutée, EBE) et IFRS (résultat part du groupe).
Les formules non recalculées sont évaluées (sommes, références de cellules).

Contrôles automatiques, signalés dans la fiche : exercice saisi dans une autre unité que les
autres (harmonisation par ordre de grandeur), unité de feuille incohérente, bilan déséquilibré,
résultat net supérieur au chiffre d'affaires, EBE reconstitué (résultat d'exploitation + dotations),
capitaux propres reconstitués.

Le référentiel (`fiche_valeur/config.py`, dictionnaire `TICKERS`) fixe pour les 48 valeurs de la cote
le nom, le pays, le secteur selon la classification retenue par CGF Gestion et les autres noms usuels
(« BANK OF AFRICA NG », « CROWN SIEM CI »...) qui servent à reconnaître la société dans les
publications, les dividendes et la presse. Une feuille `SYNTHÈSE` du classeur ne classe que les titres
absents de ce référentiel.

### 2. Base de données SMF (facultative, à charger après le classeur)

Classeur « Base de données - SMF » ; les feuilles sont reconnues par leur nom :

| Feuille | Contenu | Usage |
|---|---|---|
| `Cours` | `Date` puis une colonne par ticker ; `.BRVMCI` (ou `BRVMC`) pour le BRVM Composite | historique, technique, bêta |
| `Volume moyen` ou `Volumes` | même structure ; montants en FCFA détectés et convertis en nombre de titres | liquidité |
| `Tableau de bord` | blocs par agrégat (capitaux propres, résultat net, chiffre d'affaires, résultat d'exploitation, flux de trésorerie), années en lignes, montants en FCFA ; 0 = non renseigné | chiffres clés (niveau 2), comparables, exercices manquants du classeur |
| `Historique_dividende` ou `Dividendes` | dividende net par action, **année de paiement** en lignes (exercice N-1) | rendement, actualisation des dividendes |
| `Nb_titres` | lignes « Titres flottants » et « Nombre de titres » | capitalisation, flottant |
| `Moyenne_cours` | cours moyen par année | historique |

Apports : les titres hors classeur passent au niveau 2 avec capitaux propres (ROE, P/B, P/B justifié),
les exercices absents du classeur sont complétés sans jamais écraser une valeur existante (signalé dans
la fiche), et les comparables sectoriels couvrent toute la cote. Un dividende plus de 2,5 fois supérieur
aux versements précédents est traité comme exceptionnel : le dividende ordinaire médian sert aux calculs.

### 3. Historique d'un titre (facultatif)

CSV ou Excel avec les colonnes `Date`, `Ouverture`, `Plus haut`, `Plus bas`, `Clôture`, `Volume`
(libellés français ou anglais acceptés), par exemple un export sikafinance.

### Paramètres automatiques selon le titre

Avec l'option « Paramètres automatiques selon le titre » (activée par défaut), chaque champ de la
barre latérale est prérempli pour le titre choisi, avec sa source affichée sous le champ :

| Champ | Valeur proposée |
|---|---|
| Cours, nombre de titres, flottant, dernier dividende | ceux que l'analyse utilisera (fichiers d'abord, web ensuite) |
| Événement majeur | coché si une annonce des six derniers mois porte sur une offre publique, une augmentation de capital, une suspension, une radiation, une fusion ou une cession de bloc |
| Informations communiquées | intitulés des annonces officielles récentes (dividendes, AG, rapports, notation...) |
| Taux sans risque | taux souverain de référence du pays de l'émetteur (`TAUX_SOUVERAINS`, indicatif) |
| Bêta par défaut | bêta médian des titres du même secteur dans la base de marché |
| Seuils ACHAT / VENDRE | moitié de la volatilité annualisée du titre, entre 12% et 30% |
| Pondération des piliers | technique ramenée à 10% pour un titre peu liquide, neutralisée sans historique |

Seules les valeurs que vous modifiez sont traitées comme des saisies manuelles. Les suggestions sont
mises en cache une heure.

### Ordre de priorité des données

Les fichiers que vous chargez sont toujours la source première ; le web ne fait que compléter ce qui
manque ou ce qui est plus récent.

| Donnée | Ordre de priorité |
|---|---|
| États financiers | classeur > base SMF (Tableau de bord, exercices manquants seulement) > fiche société sikafinance > états financiers publiés (richbourse) |
| Historique de cours | fichier d'historique > base SMF > sikafinance pour les seules séances postérieures au fichier |
| Cours retenu | saisie manuelle > dernier cours de l'historique > cours du jour brvm.org s'il est plus récent |
| Nombre de titres, flottant | saisie manuelle > base SMF > brvm.org |
| Dividendes | saisie manuelle > base SMF > brvm.org et sikafinance > classeur |
| Soldes de l'exercice en cours | PDF ou adresse fournis > recherche automatique sur richbourse |
| Secteur | référentiel de l'application (classification CGF Gestion) > feuille SYNTHÈSE pour un titre hors référentiel |

Un fractionnement d'action postérieur à la base (écart d'un facteur entier entre le dernier cours du
fichier et le cours en ligne) est détecté : cours, volumes, dividendes et nombre de titres de la base
sont alors ajustés, et la fiche le signale. La provenance de chaque donnée figure dans la section
« Sources » de la fiche.

---

## Soldes de l'exercice en cours : publications officielles

Une fois un titre choisi (données en ligne activées), l'application parcourt la rubrique « États
financiers » de richbourse.com, qui relaie les avis publiés à la BRVM, repère les publications de la
société par leur intitulé (« ORANGE CI : Rapport d'activités - 1er semestre 2026 »), télécharge le PDF
le plus récent et lit son tableau d'indicateurs : chiffre d'affaires ou PNB, EBE, résultat
d'exploitation, résultat financier, résultat des activités ordinaires, coût du risque, résultat avant
impôt, résultat net, période N face à N-1.

Ces soldes figurent dans la fiche PDF, le rapport PowerPoint (diapositive « Exercice en cours »),
l'export Excel et l'application. Leur évolution sur un an entre pour 30% dans le pilier croissance. Pour
un titre sans états financiers détaillés, des états financiers annuels publiés complètent l'exercice
manquant.

Si le site ne répond pas ou si le document est scanné, chargez le PDF ou collez son adresse dans
l'onglet d'analyse (« Dernière publication de la société ») : le même analyseur est utilisé.

## Sources en ligne

| Donnée | Source |
|---|---|
| Historique des cours (OHLCV) et indice BRVM Composite | sikafinance.com |
| Nombre de titres, cours du jour, capitalisation flottante | brvm.org (capitalisations) |
| Dividendes publiés | brvm.org (paiement de dividendes), puis sikafinance.com |
| Chiffres clés (CA, résultat net, BNPA, PER, dividende, actionnariat) des titres hors classeur | sikafinance.com (fiche société) |
| Rapports d'activités et états financiers publiés (soldes de l'exercice en cours) | richbourse.com, rubrique États financiers (PDF) |
| Croissance du PIB et inflation par pays (2024-2027) | FMI, API DataMapper (valeurs de repli intégrées) |
| Parité EUR/USD (sociétés en USD) | frankfurter.app (parité fixe EUR/XOF 655,957) |
| Actualités de l'émetteur | Google Actualités (flux RSS) |
| Politique monétaire | BCEAO, paramètres modifiables dans l'application |

Chaque source est interrogée avec un délai court et mise en cache une heure. En cas
d'indisponibilité, l'analyse se poursuit avec les autres données et l'incident est consigné dans la
fiche. Le mode hors ligne n'interroge aucune source.

---

## Méthode d'analyse

### Analyse fondamentale

- **Soldes intermédiaires de gestion** adaptés au type de société : banque (PNB, frais généraux,
  RBE, coût du risque, résultat net), SYSCOHADA (CA, marge commerciale, valeur ajoutée, EBE,
  résultat d'exploitation, résultat financier, HAO) et IFRS.
- **Ratios** : ROE, ROA, marges, coefficient d'exploitation, coût du risque, crédits/dépôts,
  levier, dette nette/EBE, couverture des frais financiers, BFR, taux de distribution, CAF.
- **Solidité financière** : neuf tests de type Piotroski (adaptés aux banques) et score d'Altman
  Z'' marchés émergents pour les sociétés non financières.
- **Croissance** : taux de croissance annuel moyen du revenu, du résultat et des fonds propres.
- **Comparaison aux pairs** du classeur (même secteur, sinon même famille) : ROE, marge, PER, P/B,
  rendement, médianes calculées sur deux sociétés au moins.

### Valorisation

- Coût des fonds propres : k = taux sans risque + bêta × prime de risque (bêta ajusté de Blume,
  calculé contre l'indice BRVM Composite, borné entre 0,6 et 1,4).
- Méthodes et pondérations : PER et P/B sectoriels, P/B justifié (ROE normalisé - g) / (k - g),
  modèle d'actualisation des dividendes à deux étapes, VE/EBE pour les sociétés non financières.
- Sorties : juste valeur, fourchette, potentiel, rentabilité totale attendue (potentiel +
  rendement), ROE implicite dans le cours.

### Analyse technique

Moyennes mobiles 20/50/100/200 séances, RSI 14 (Wilder), MACD, bandes de Bollinger, ATR,
stochastique, volatilité annualisée, drawdown, force relative face à l'indice, supports et
résistances par pivots, liquidité (taux de séances actives, rotation du flottant).

### Environnement macroéconomique et actualités

Croissance et inflation du pays d'origine (FMI), taux directeurs de la BCEAO et sensibilité du
secteur au cycle. Tonalité des articles récents par lexique financier français, pondérée par
l'ancienneté, et repérage des événements sur titre (dividende, AG, opération sur capital, OPA,
gouvernance, contentieux).

### Construction du signal

Six piliers notés sur 100, pondérations modifiables dans l'application :

| Pilier | Poids par défaut |
|---|---|
| Valorisation | 30% |
| Qualité et solidité | 20% |
| Croissance | 15% |
| Technique | 20% |
| Macroéconomie | 10% |
| Actualités | 5% |

1. La **vue fondamentale** (tous les piliers sauf la technique) est classée favorable (≥ 62),
   neutre (≥ 42) ou défavorable.
2. Le **potentiel** classe le titre en décote (≥ +15%), proche de sa juste valeur, ou en surcote
   (≤ -15%).
3. La matrice fondamentaux × valorisation donne le signal de base :

   | | Décote | Juste valeur | Surcote |
   |---|---|---|---|
   | **Favorable** | ACHAT | CONSERVER | ALLÉGER |
   | **Neutre** | ACHAT si tendance haussière, sinon ATTENDRE | CONSERVER | ALLÉGER, VENDRE si tendance baissière |
   | **Défavorable** | ATTENDRE (piège de valeur) | ALLÉGER, VENDRE si tendance baissière | VENDRE |

4. La **technique sert de déclencheur** : un achat en tendance baissière (score < 35) devient
   ATTENDRE ; une conservation sans potentiel en tendance baissière devient ALLÉGER.
5. **Garde-fous** : un événement majeur signalé par l'analyste (OPA, augmentation de capital,
   suspension) impose ATTENDRE, sauf signal VENDRE ; la conviction est abaissée en cas d'actualité
   majeure récente, de moins de quatre piliers renseignés, de forte dispersion des méthodes de
   valorisation ou de liquidité faible.
6. La **zone d'accumulation** (entre le premier support et la moyenne mobile 50 jours) et le
   **niveau d'invalidation** (sous le deuxième support) complètent le signal.

Tous les seuils sont réglables dans `fiche_valeur/config.py` (classe `Settings`) ou dans la barre
latérale de l'application.

### Commentaires rédigés par Claude (facultatif)

Avec une clé `ANTHROPIC_API_KEY`, l'option « Reformuler la thèse avec Claude » reformule la
thèse d'investissement et liste catalyseurs et risques à partir des chiffres calculés. Elle ne
modifie jamais le signal, qui reste issu de la méthode ci-dessus.

---

## Structure du dépôt

```
fiche-valeur-brvm/
├── app.py                     interface Streamlit
├── cli.py                     génération en ligne de commande
├── requirements.txt
├── fiche_valeur/
│   ├── config.py              référentiel des valeurs, paramètres, BCEAO, pondérations
│   ├── parser.py              lecture et contrôle du classeur des états financiers
│   ├── market.py              cours, titres, dividendes (fichiers et sources en ligne)
│   ├── partial.py             titres à données partielles : chiffres clés, saisie, base SMF, couverture
│   ├── publications.py        publications officielles : liste richbourse, lecture des PDF
│   ├── autoparams.py          paramètres proposés automatiquement selon le titre
│   ├── fundamentals.py        SIG, ratios, solidité, pairs
│   ├── valuation.py           juste valeur multi-méthodes
│   ├── technical.py           indicateurs techniques et niveaux
│   ├── macro.py               environnement macroéconomique
│   ├── news.py                actualités, tonalité, événements
│   ├── scoring.py             piliers et signal
│   ├── commentary.py          textes d'analyse (et enrichissement Claude)
│   ├── pipeline.py            orchestration
│   ├── charts.py              graphiques
│   ├── report_pdf.py          fiche PDF
│   ├── report_pptx.py         rapport PowerPoint
│   └── export.py              exports Excel
├── assets/fonts/              polices Carlito et Caladea (licence SIL OFL)
├── sample_data/               données de démonstration fictives et leur générateur
├── tests/                     tests hors ligne (pytest)
└── .streamlit/config.toml     thème de l'application
```

## Tests

```bash
pip install pytest
python -m pytest -q
```

Les tests fonctionnent hors ligne : lecture du classeur de démonstration et d'un classeur à
formules, harmonisation des unités, signaux, rapports PDF, PowerPoint et Excel, et analyse
d'échantillons des réponses de brvm.org, sikafinance et Google Actualités.

Pour régénérer les données de démonstration : `python sample_data/make_demo_workbook.py`.

## Limites

- Les sites de données peuvent changer leur présentation ; les analyseurs sont isolés
  (`parse_capitalisations`, `parse_sika_payload`, `parse_sika_societe`, `parse_listing`, `parse_report_text`, `parse_rss`) pour être ajustés facilement.
  Une base de marché locale garantit des résultats reproductibles.
- Les paramètres BCEAO et les valeurs macroéconomiques de repli datent de la dernière mise à jour
  du code : vérifiez-les dans la barre latérale avant diffusion.
- La tonalité des actualités repose sur les titres d'articles ; les informations communiquées par
  la société (communiqués, AG, rapports) peuvent être saisies dans l'application pour être prises en
  compte.
- - Les chiffres clés lus en ligne (niveau 2) n'incluent pas le bilan : sans capitaux propres saisis, ni
  ROE ni P/B. Les multiples de référence par secteur sont indicatifs.
- Les rapports d'activités n'ont pas tous la même mise en page : les lignes non reconnues sont
  ignorées, et un PDF scanné (image) n'est pas lu. Les périodes trimestrielles peuvent être cumulées
  depuis le début de l'exercice : se référer à l'intitulé du document.
- Les sociétés de démonstration sont fictives.

## Avertissement

Ce document est produit à titre d'information et ne constitue ni une offre ni une sollicitation
d'achat ou de vente de valeurs mobilières. Les performances passées ne préjugent pas des
performances futures. Les signaux sont générés selon une méthode quantitative et doivent être
validés par l'analyste avant toute décision d'investissement, dans le respect de la réglementation
de l'AMF-UMOA.

## Licences

Polices Carlito et Caladea : SIL Open Font License 1.1 (`assets/fonts/OFL-LICENSES.txt`).
