"""Paramètres, référentiels et valeurs par défaut de l'application.

Tout ce qui est susceptible d'être ajusté par le gérant (taux, pondérations,
seuils de signal) est centralisé ici et peut être surchargé depuis
l'interface ou la ligne de commande.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FONTS_DIR = ROOT / "assets" / "fonts"

APP_NAME = "Fiche de valeur BRVM"
APP_VERSION = "1.0.0"
DEFAULT_HOUSE = "CGF Gestion"

# ---------------------------------------------------------------------------
# Charte graphique CGF Gestion (modèles « Modèle de sortie CGF GESTION » et présentations maison) :
# bleu institutionnel, cyan d'accent, cartes gris-bleu clair ; police Inter (Aptos dans PowerPoint)
# ---------------------------------------------------------------------------
NAVY = "#004B77"          # titres, en-têtes de tableaux
NAVY_DARK = "#0B2239"     # fond sombre (dégradé des couvertures)
GOLD = "#00B0F0"          # accent (filets, puces) - nom historique conservé
GOLD_LIGHT = "#C9E2F0"
SKY = "#59BCE9"           # sous-titres sur fond sombre
INK = "#1A1A1A"
MUTED = "#595959"
GRID = "#DCE4EA"
PAPER = "#F3F6F8"         # fond des cartes
INFO = "#EAF4FA"          # encadré d'information
WHITE = "#FFFFFF"
POS = "#1B7F5A"      # vert hausse
NEG = "#B4372F"      # rouge baisse
NEUTRAL = "#8A8F9C"
SERIES = [NAVY, GOLD, "#4F7CAC", "#8C6D3F", "#7A8FB8", "#B9A06A"]
BRAND_DIR = ROOT / "assets" / "brand"
TEMPLATES_DIR = ROOT / "assets" / "templates"
PPTX_TEMPLATE = TEMPLATES_DIR / "modele_cgf_gestion.pptx"
LETTERHEAD = BRAND_DIR / "entete_cgf_gestion.jpg"
LEGAL_MENTION = ("Société de gestion collective au capital de 500 000 000 FCFA - N° agrément CREPMF SG-003/2001")
CONTACTS = [
    ("web", "www.cgfgestion.com", "Site de la société de gestion"),
    ("courriel", "cgfgestion@cgfgestion.com", "Société de gestion - CGF Gestion"),
    ("telephone", "+221 33 864 97 97", "Km 6 avenue Cheikh Anta Diop, Dakar"),
]

SIGNAL_COLORS = {
    "ACHAT": "#1B7F5A",
    "CONSERVER": "#2F5D9E",
    "ATTENDRE": "#C08A1E",
    "ALLÉGER": "#C8622B",
    "VENDRE": "#B4372F",
}

SIGNAL_DEFINITIONS = {
    "ACHAT": "Potentiel de revalorisation significatif, fondamentaux solides et "
             "configuration technique favorable : constituer ou renforcer la position.",
    "CONSERVER": "Valorisation proche de la juste valeur ou potentiel modéré : "
                 "maintenir la position, encaisser le rendement.",
    "ATTENDRE": "Thèse favorable mais point d'entrée ou information à confirmer "
                "(tendance technique, événement en cours, fondamentaux à valider) : "
                "placer le titre sous surveillance.",
    "ALLÉGER": "Potentiel épuisé ou signaux de dégradation : réduire progressivement "
               "l'exposition, prendre des bénéfices.",
    "VENDRE": "Surévaluation marquée et fondamentaux ou tendance défavorables : "
              "solder la position.",
}

# ---------------------------------------------------------------------------
# Pays de l'UEMOA
# ---------------------------------------------------------------------------
COUNTRIES = {
    "CI": {"nom": "CÔTE D'IVOIRE", "iso3": "CIV", "suffixe": ".ci", "gl": "CI"},
    "SN": {"nom": "SÉNÉGAL", "iso3": "SEN", "suffixe": ".sn", "gl": "SN"},
    "BF": {"nom": "BURKINA FASO", "iso3": "BFA", "suffixe": ".bf", "gl": "BF"},
    "ML": {"nom": "MALI", "iso3": "MLI", "suffixe": ".ml", "gl": "ML"},
    "BJ": {"nom": "BENIN", "iso3": "BEN", "suffixe": ".bj", "gl": "BJ"},
    "TG": {"nom": "TOGO", "iso3": "TGO", "suffixe": ".tg", "gl": "TG"},
    "NE": {"nom": "NIGER", "iso3": "NER", "suffixe": ".ne", "gl": "NE"},
    "GW": {"nom": "GUINÉE-BISSAU", "iso3": "GNB", "suffixe": ".gw", "gl": "GW"},
}

# ---------------------------------------------------------------------------
# Référentiel des valeurs cotées (nomenclature sectorielle BRVM à 7 secteurs)
# Le secteur lu dans la feuille SYNTHÈSE du classeur prime s'il est présent.
# ---------------------------------------------------------------------------
TEL = "Télécommunications"
FIN = "Services financiers"
CDB = "Consommation de base"
CDI = "Consommation discrétionnaire"
ENE = "Énergie"
IND = "Industriels"
SPU = "Services publics"
SECTORS = [FIN, TEL, CDB, CDI, ENE, SPU, IND]

# Référentiel des titres : secteurs et dénominations retenus par CGF Gestion (classification de
# référence de l'application, prioritaire sur la feuille SYNTHÈSE du classeur). « autres_noms » sert à
# reconnaître la société dans les publications, les dividendes et la presse.
TICKERS: dict[str, dict] = {
    "ABJC": {"nom": "SERVAIR ABIDJAN COTE D'IVOIRE", "pays": "CI", "secteur": CDI, "autres_noms": ['SERVAIR ABIDJAN']},
    "BBGC": {"nom": "BRIDGE BANK GROUP COTE D'IVOIRE", "pays": "CI", "secteur": FIN, "autres_noms": []},
    "BICB": {"nom": "BIIC BENIN", "pays": "BJ", "secteur": FIN, "autres_noms": []},
    "BICC": {"nom": "BICI COTE D'IVOIRE", "pays": "CI", "secteur": FIN, "autres_noms": ['BICI CI']},
    "BNBC": {"nom": "BERNABE COTE D'IVOIRE", "pays": "CI", "secteur": CDI, "autres_noms": ['BERNABE CI']},
    "BOAB": {"nom": "BANK OF AFRICA BENIN", "pays": "BJ", "secteur": FIN, "autres_noms": ['BANK OF AFRICA BN']},
    "BOABF": {"nom": "BANK OF AFRICA BURKINA FASO", "pays": "BF", "secteur": FIN, "autres_noms": ['BANK OF AFRICA BF']},
    "BOAC": {"nom": "BANK OF AFRICA COTE D'IVOIRE", "pays": "CI", "secteur": FIN, "autres_noms": ['BANK OF AFRICA CI']},
    "BOAM": {"nom": "BANK OF AFRICA MALI", "pays": "ML", "secteur": FIN, "autres_noms": ['BANK OF AFRICA ML']},
    "BOAN": {"nom": "BANK OF AFRICA NIGER", "pays": "NE", "secteur": FIN, "autres_noms": ['BANK OF AFRICA NG']},
    "BOAS": {"nom": "BANK OF AFRICA SENEGAL", "pays": "SN", "secteur": FIN, "autres_noms": ['BANK OF AFRICA SN']},
    "CABC": {"nom": "SICABLE COTE D'IVOIRE", "pays": "CI", "secteur": IND, "autres_noms": ['SICABLE CI']},
    "CBIBF": {"nom": "CORIS BANK INTERNATIONAL BURKINA FASO", "pays": "BF", "secteur": FIN, "autres_noms": ['CORIS BANK INTERNATIONAL BF']},
    "CFAC": {"nom": "CFAO MOTORS COTE D'IVOIRE", "pays": "CI", "secteur": CDI, "autres_noms": ['CFAO CI']},
    "CIEC": {"nom": "CIE COTE D'IVOIRE", "pays": "CI", "secteur": SPU, "autres_noms": ['CIE CI']},
    "ECOC": {"nom": "ECOBANK COTE D'IVOIRE", "pays": "CI", "secteur": FIN, "autres_noms": ['ECOBANK COTE DIVOIRE']},
    "ETIT": {"nom": "ECOBANK TRANSNATIONAL INCORPORATED TOGO", "pays": "TG", "secteur": FIN, "autres_noms": ['ECOBANK ETI']},
    "FTSC": {"nom": "FILTISAC COTE D'IVOIRE", "pays": "CI", "secteur": IND, "autres_noms": ['FILTISAC']},
    "LNBB": {"nom": "LOTERIE NATIONALE DU BENIN", "pays": "BJ", "secteur": CDI, "autres_noms": []},
    "NEIC": {"nom": "NEI-CEDA COTE D'IVOIRE", "pays": "CI", "secteur": CDI, "autres_noms": []},
    "NSBC": {"nom": "NSIA BANQUE COTE D'IVOIRE", "pays": "CI", "secteur": FIN, "autres_noms": ['NSIA BANQUE CI']},
    "NTLC": {"nom": "NESTLE COTE D'IVOIRE", "pays": "CI", "secteur": CDB, "autres_noms": ['NESTLE CI']},
    "ONTBF": {"nom": "ONATEL BURKINA FASO", "pays": "BF", "secteur": TEL, "autres_noms": ['ONATEL BF']},
    "ORAC": {"nom": "ORANGE COTE D'IVOIRE", "pays": "CI", "secteur": TEL, "autres_noms": ['Orange CI']},
    "ORGT": {"nom": "ORAGROUP TOGO", "pays": "TG", "secteur": FIN, "autres_noms": []},
    "PALC": {"nom": "PALM COTE D'IVOIRE", "pays": "CI", "secteur": CDB, "autres_noms": ['PALM CI']},
    "PRSC": {"nom": "TRACTAFRIC MOTORS COTE D'IVOIRE", "pays": "CI", "secteur": CDI, "autres_noms": ['TRACTAFRIC MOTORS CI']},
    "SAFC": {"nom": "SAFCA COTE D'IVOIRE", "pays": "CI", "secteur": FIN, "autres_noms": ['SAFCA']},
    "SCRC": {"nom": "SUCRIVOIRE COTE D'IVOIRE", "pays": "CI", "secteur": CDB, "autres_noms": ['SUCRIVOIRE']},
    "SDCC": {"nom": "SODE COTE D'IVOIRE", "pays": "CI", "secteur": SPU, "autres_noms": ['SODECI']},
    "SDSC": {"nom": "BOLLORE TRANSPORT & LOGISTICS COTE D'IVOIRE", "pays": "CI", "secteur": IND, "autres_noms": ['BOLLORE AFRICA LOGISTIC']},
    "SEMC": {"nom": "EVIOSYS PACKAGING SIEM COTE D'IVOIRE", "pays": "CI", "secteur": IND, "autres_noms": ['CROWN SIEM CI']},
    "SGBC": {"nom": "SOCIETE GENERALE COTE D'IVOIRE", "pays": "CI", "secteur": FIN, "autres_noms": []},
    "SHEC": {"nom": "VIVO ENERGY COTE D'IVOIRE", "pays": "CI", "secteur": ENE, "autres_noms": ['VIVO ENERGY CI']},
    "SIBC": {"nom": "SOCIETE IVOIRIENNE DE BANQUE COTE D'IVOIRE", "pays": "CI", "secteur": FIN, "autres_noms": ['SOCIETE IVOIRIENNE DE BANQUE']},
    "SICC": {"nom": "SICOR COTE D'IVOIRE", "pays": "CI", "secteur": CDB, "autres_noms": ['SICOR CI']},
    "SIVC": {"nom": "AIR LIQUIDE COTE D'IVOIRE", "pays": "CI", "secteur": IND, "autres_noms": ['AIR LIQUIDE CI']},
    "SLBC": {"nom": "SOLIBRA COTE D'IVOIRE", "pays": "CI", "secteur": CDB, "autres_noms": []},
    "SMBC": {"nom": "SMB COTE D'IVOIRE", "pays": "CI", "secteur": ENE, "autres_noms": ['SMB CI']},
    "SNTS": {"nom": "SONATEL SENEGAL", "pays": "SN", "secteur": TEL, "autres_noms": []},
    "SOGC": {"nom": "SOGB COTE D'IVOIRE", "pays": "CI", "secteur": CDB, "autres_noms": ['SOGB CI']},
    "SPHC": {"nom": "SAPH COTE D'IVOIRE", "pays": "CI", "secteur": CDB, "autres_noms": ['SAPH CI']},
    "STAC": {"nom": "SETAO COTE D'IVOIRE", "pays": "CI", "secteur": IND, "autres_noms": ['SETAO CI']},
    "STBC": {"nom": "SITAB COTE D'IVOIRE", "pays": "CI", "secteur": CDB, "autres_noms": ['SITAB CI']},
    "TTLC": {"nom": "TOTAL COTE D'IVOIRE", "pays": "CI", "secteur": ENE, "autres_noms": ['TOTAL CI']},
    "TTLS": {"nom": "TOTAL SENEGAL", "pays": "SN", "secteur": ENE, "autres_noms": ['TOTAL SN']},
    "UNLC": {"nom": "UNILEVER COTE D'IVOIRE", "pays": "CI", "secteur": CDB, "autres_noms": []},
    "UNXC": {"nom": "UNIWAX COTE D'IVOIRE", "pays": "CI", "secteur": CDI, "autres_noms": ['UNIWAX CI']},
}

# Comptes consolidés IFRS (résultat et capitaux propres part du groupe)
IFRS_TICKERS = {"SNTS", "ORAC", "ORGT", "ETIT"}

# Libellés de secteur non normalisés rencontrés dans les classeurs, ramenés
# à la nomenclature BRVM à 7 secteurs.
SECTOR_ALIASES = {
    "banque": FIN, "banques": FIN, "finance": FIN, "services financiers": FIN,
    "groupe bancaire": FIN, "groupe bancaire ifrs": FIN, "assurance": FIN,
    "telecom": TEL, "telecoms": TEL, "telecommunications": TEL,
    "agro industrie": CDB, "agro-industrie": CDB, "agroalimentaire": CDB,
    "agriculture": CDB, "consommation de base": CDB, "brasserie": CDB,
    "commerce": CDI, "distribution": CDI, "medias": CDI, "consommation discretionnaire": CDI,
    "transport": IND, "industrie": IND, "industriels": IND, "construction": IND,
    "concession": SPU, "services publics": SPU, "energie": ENE, "petrole": ENE,
}

# ---------------------------------------------------------------------------
# Environnement monétaire et macroéconomique (valeurs de repli)
# Utilisées lorsque les sources en ligne sont indisponibles. Modifiables
# dans l'interface. Sources : communiqués du Comité de Politique Monétaire
# de la BCEAO (10 juin et 9 septembre 2026), FMI World Economic Outlook
# via l'API DataMapper (consultée le 18/10/2026).
# ---------------------------------------------------------------------------
BCEAO = {
    "date_reference": "9 septembre 2026",
    "taux_directeur": 3.00,
    "guichet_pret_marginal": 5.00,
    "reserves_obligatoires": 3.00,
    "inflation_uemoa_prevue": 1.0,
    "croissance_uemoa_prevue": 6.1,
    "commentaire": (
        "Taux directeur maintenu à 3,00% (en vigueur depuis le 16 mars 2026, après une "
        "baisse de 25 pb) ; guichet de prêt marginal à 5,00%. Inflation attendue à 1,0% "
        "en 2026 après 0,0% en 2025 ; croissance de l'Union attendue à 6,1%."
    ),
}

# Croissance du PIB réel (%) et inflation moyenne (%) - FMI, WEO
MACRO_FALLBACK = {
    "CIV": {"croissance": {2024: 6.0, 2025: 6.5, 2026: 6.0, 2027: 6.1},
            "inflation": {2024: 3.4, 2025: 0.1, 2026: 2.2, 2027: 2.5}},
    "SEN": {"croissance": {2024: 6.5, 2025: 6.7},
            "inflation": {2024: 0.8, 2025: 1.4}},
    "BFA": {"croissance": {2024: 4.8, 2025: 5.3, 2026: 4.4, 2027: 4.8},
            "inflation": {2024: 4.2, 2025: -0.5, 2026: 2.1, 2027: 2.2}},
    "MLI": {"croissance": {2024: 5.0, 2025: 4.9, 2026: 5.5, 2027: 5.7},
            "inflation": {2024: 3.2, 2025: 2.3, 2026: 3.0, 2027: 2.5}},
    "BEN": {"croissance": {2024: 7.5, 2025: 8.1, 2026: 7.0, 2027: 6.8},
            "inflation": {2024: 1.2, 2025: 1.1, 2026: 0.8, 2027: 2.0}},
    "TGO": {"croissance": {2024: 6.3, 2025: 6.3, 2026: 5.0, 2027: 5.3},
            "inflation": {2024: 2.9, 2025: 0.4, 2026: 1.1, 2027: 2.3}},
    "NER": {"croissance": {2024: 8.3, 2025: 6.9, 2026: 7.0, 2027: 6.5},
            "inflation": {2024: 9.1, 2025: -4.7, 2026: -1.9, 2027: 2.0}},
    "GNB": {"croissance": {2024: 4.8, 2025: 5.8, 2026: 4.7, 2027: 4.9},
            "inflation": {2024: 3.7, 2025: 0.9, 2026: 2.5, 2027: 2.1}},
}

IMF_INDICATORS = {
    "NGDP_RPCH": "croissance",
    "PCPIPCH": "inflation",
    "GGXWDG_NGDP": "dette_publique",
    "GGXCNL_NGDP": "solde_budgetaire",
    "BCA_NGDPD": "solde_courant",
}

# Taux de change de repli (1 EUR = 655,957 FCFA, parité fixe)
EUR_XOF = 655.957
EURUSD_FALLBACK = 1.15

def guess_country(ticker: str) -> str:
    """Pays probable d'après la convention de code BRVM (dernière lettre : C Côte d'Ivoire, S Sénégal...)."""
    t = ticker.upper()
    if t.endswith("BF"):
        return "BF"
    return {"C": "CI", "S": "SN", "B": "BJ", "T": "TG", "N": "NE", "M": "ML", "G": "GW"}.get(t[-1:], "CI")


def register_ticker(ticker: str, nom: str, pays: str = "CI", secteur: str = "Non classé") -> None:
    """Ajoute (ou met à jour) un titre absent du référentiel : introduction récente, saisie libre."""
    TICKERS[ticker.upper()] = {"nom": nom or ticker.upper(), "pays": pays, "secteur": secteur}


# Taux sans risque de référence par pays (%, rendement souverain long UMOA-Titres). Valeurs indicatives
# servant de valeur par défaut au coût des fonds propres : à mettre à jour d'après les dernières adjudications.
TAUX_SOUVERAINS = {"CI": 6.25, "SN": 6.50, "BJ": 6.25, "TG": 6.50, "BF": 7.25, "ML": 7.50, "NE": 7.75, "GW": 7.75}

# Multiples de référence par secteur, utilisés quand le classeur ne contient pas assez de
# sociétés comparables (moins de deux). Valeurs indicatives de la cote régionale, à ajuster.
SECTOR_MULTIPLES = {
    "Services financiers": {"pe": 7.5, "pb": 1.0, "ev_ebitda": None, "div_yield": 0.070},
    "Télécommunications": {"pe": 10.5, "pb": 3.0, "ev_ebitda": 5.0, "div_yield": 0.070},
    "Consommation de base": {"pe": 10.0, "pb": 1.5, "ev_ebitda": 6.0, "div_yield": 0.060},
    "Consommation discrétionnaire": {"pe": 9.0, "pb": 1.3, "ev_ebitda": 5.5, "div_yield": 0.060},
    "Industriels": {"pe": 9.0, "pb": 1.2, "ev_ebitda": 5.0, "div_yield": 0.060},
    "Énergie": {"pe": 9.0, "pb": 1.8, "ev_ebitda": 5.0, "div_yield": 0.065},
    "Services publics": {"pe": 10.0, "pb": 2.5, "ev_ebitda": 5.5, "div_yield": 0.060},
    "Non classé": {"pe": 9.0, "pb": 1.3, "ev_ebitda": 5.5, "div_yield": 0.060},
}

# Niveaux de complétude des données d'un titre
LEVELS = {
    "complet": ("Niveau 1 - Analyse complète",
                "États financiers détaillés sur au moins trois exercices (compte de résultat et bilan)."),
    "partiel": ("Niveau 1b - Analyse complète avec lacunes",
                "États financiers présents mais incomplets (exercices ou postes manquants)."),
    "chiffres_cles": ("Niveau 2 - Chiffres clés",
                      "Chiffre d'affaires, résultat net, fonds propres ou BNPA et dividendes, sans états financiers détaillés."),
    "marche": ("Niveau 3 - Données de marché",
               "Cours, volumes et dividendes uniquement : analyse technique et valorisation par les "
               "dividendes, signal indicatif prudent."),
}

# ---------------------------------------------------------------------------
# Paramètres d'analyse par défaut
# ---------------------------------------------------------------------------


@dataclass
class Settings:
    # Coût des fonds propres
    taux_sans_risque: float = 6.50      # %, rendement souverain UMOA-Titres de référence
    prime_risque_marche: float = 5.50   # %, prime de risque actions BRVM
    croissance_long_terme: float = 3.00  # %, croissance perpétuelle g
    beta_defaut: float = 1.00
    horizon_mois: int = 12
    # Pondérations des piliers (somme libre, renormalisée)
    poids: dict = field(default_factory=lambda: {
        "valorisation": 30.0,
        "qualite": 20.0,
        "croissance": 15.0,
        "technique": 20.0,
        "macro": 10.0,
        "sentiment": 5.0,
    })
    # Seuils de décision
    seuil_potentiel_achat: float = 15.0     # % de potentiel pour un signal d'achat
    seuil_potentiel_vente: float = -15.0    # % de potentiel pour un signal de vente
    seuil_fondamental_fort: float = 62.0
    seuil_fondamental_faible: float = 42.0
    seuil_technique_favorable: float = 45.0
    seuil_technique_defavorable: float = 35.0
    # Données en ligne
    en_ligne: bool = True
    multiples_defaut: bool = True       # multiples sectoriels de référence si pairs insuffisants
    annees_historique: int = 3
    fx_eurusd: float = EURUSD_FALLBACK
    # Commentaire assisté par Claude (facultatif)
    claude_api_key: str = ""
    claude_model: str = "claude-sonnet-5-5"
    maison: str = DEFAULT_HOUSE
    analyste: str = ""

    @property
    def cout_fonds_propres_base(self) -> float:
        return self.taux_sans_risque + self.beta_defaut * self.prime_risque_marche


PILLAR_LABELS = {
    "valorisation": "Valorisation",
    "qualite": "Qualité & rentabilité",
    "croissance": "Croissance",
    "technique": "Analyse technique",
    "macro": "Environnement macro",
    "sentiment": "Actualités & sentiment",
}

DISCLAIMER = (
    "Document d'aide à la décision à usage interne, établi à partir des états financiers "
    "publiés, des données de marché et d'informations publiques. Les signaux résultent d'un "
    "modèle multicritère paramétrable et doivent être validés par le gérant et le comité "
    "d'investissement. Ils ne constituent ni une offre ni un conseil en investissement au sens "
    "de la réglementation AMF-UMOA. Les performances passées ne préjugent pas des performances "
    "futures."
)
