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

Un tableau de bord multi-valeurs analyse tout le classeur en une fois et exporte les signaux et les
fiches dans une archive.

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
la barre latérale, choisissez une valeur, puis cliquez sur **Lancer l'analyse**. Sans fichier,
l'application s'ouvre sur les classeurs de démonstration (sociétés fictives) pour découvrir l'outil.
L'option **Compléter avec les données en ligne** active les sources listées plus bas.

Python 3.10 ou plus récent est requis.

### En ligne de commande

```bash
# Une valeur
python cli.py --ef Etats_Financiers_2021_2025.xlsx --ticker SNTS

# Tout le classeur, avec la base de marché, le PowerPoint et le détail Excel
python cli.py --ef Etats_Financiers_2021_2025.xlsx --smf "Base de données - SMF.xlsx" --all --pptx --xlsx

# Démonstration hors ligne
python cli.py --ef sample_data/demo_etats_financiers.xlsx --smf sample_data/demo_base_marche.xlsx --all --hors-ligne
```

Options : `--compact` (fiche d'une page), `--hors-ligne`, `--rf 6.5` (taux sans risque en %),
`--prime 5.5` (prime de risque actions en %), `--g 3` (croissance perpétuelle en %),
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

Une feuille `SYNTHÈSE` facultative (colonnes ticker et secteur) permet d'imposer le secteur de
chaque société. Le référentiel interne (`fiche_valeur/config.py`) couvre les 48 valeurs de la cote
(nom, pays, secteur).

### 2. Base de marché (facultative)

Classeur type « Base de données - SMF » ; les feuilles sont reconnues par leur nom :

| Feuille | Contenu |
|---|---|
| `Cours` | colonne `Date` puis une colonne par ticker (cours de clôture) ; une colonne `BRVMC` pour l'indice composite |
| `Volumes` | même structure, volumes échangés |
| `Nb_titres` | ticker, nombre total de titres, flottant |
| `Moyenne_cours` | ticker, cours moyen par année |
| `Dividendes` | ticker, exercice, dividende net par action, date de paiement |

### 3. Historique d'un titre (facultatif)

CSV ou Excel avec les colonnes `Date`, `Ouverture`, `Plus haut`, `Plus bas`, `Clôture`, `Volume`
(libellés français ou anglais acceptés), par exemple un export sikafinance.

### Ordre de priorité des données de marché

Saisie manuelle dans l'application > base de marché ou fichier d'historique > sources en ligne >
états financiers (nombre d'actions, dividende). La provenance de chaque donnée figure dans la
section « Sources » de la fiche.

---

## Sources en ligne

| Donnée | Source |
|---|---|
| Historique des cours (OHLCV) et indice BRVM Composite | sikafinance.com |
| Nombre de titres, cours du jour, capitalisation flottante | brvm.org (capitalisations) |
| Dividendes publiés | brvm.org (paiement de dividendes), puis sikafinance.com |
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
  (`parse_capitalisations`, `parse_sika_payload`, `parse_rss`) pour être ajustés facilement.
  Une base de marché locale garantit des résultats reproductibles.
- Les paramètres BCEAO et les valeurs macroéconomiques de repli datent de la dernière mise à jour
  du code : vérifiez-les dans la barre latérale avant diffusion.
- La tonalité des actualités repose sur les titres d'articles ; les informations communiquées par
  la société (communiqués, AG, rapports) peuvent être saisies dans l'application pour être prises en
  compte.
- Les sociétés de démonstration sont fictives.

## Avertissement

Ce document est produit à titre d'information et ne constitue ni une offre ni une sollicitation
d'achat ou de vente de valeurs mobilières. Les performances passées ne préjugent pas des
performances futures. Les signaux sont générés selon une méthode quantitative et doivent être
validés par l'analyste avant toute décision d'investissement, dans le respect de la réglementation
de l'AMF-UMOA.

## Licences

Polices Carlito et Caladea : SIL Open Font License 1.1 (`assets/fonts/OFL-LICENSES.txt`).
