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
# Charte graphique (navy / or, Cambria-Calibri ; Caladea-Carlito à l'écran)
# ---------------------------------------------------------------------------
NAVY = "#1E2761"
NAVY_DARK = "#141A45"
GOLD = "#C9A34E"
GOLD_LIGHT = "#E8D6A8"
INK = "#1F2430"
MUTED = "#6B7280"
GRID = "#E3E6EE"
PAPER = "#F6F7FB"
WHITE = "#FFFFFF"
POS = "#1B7F5A"      # vert hausse
NEG = "#B4372F"      # rouge baisse
NEUTRAL = "#8A8F9C"
SERIES = [NAVY, GOLD, "#4F7CAC", "#8C6D3F", "#7A8FB8", "#B9A06A"]

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
    "CI": {"nom": "Côte d'Ivoire", "iso3": "CIV", "suffixe": ".ci", "gl": "CI"},
    "SN": {"nom": "Sénégal", "iso3": "SEN", "suffixe": ".sn", "gl": "SN"},
    "BF": {"nom": "Burkina Faso", "iso3": "BFA", "suffixe": ".bf", "gl": "BF"},
    "ML": {"nom": "Mali", "iso3": "MLI", "suffixe": ".ml", "gl": "ML"},
    "BJ": {"nom": "Bénin", "iso3": "BEN", "suffixe": ".bj", "gl": "BJ"},
    "TG": {"nom": "Togo", "iso3": "TGO", "suffixe": ".tg", "gl": "TG"},
    "NE": {"nom": "Niger", "iso3": "NER", "suffixe": ".ne", "gl": "NE"},
    "GW": {"nom": "Guinée-Bissau", "iso3": "GNB", "suffixe": ".gw", "gl": "GW"},
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

TICKERS: dict[str, dict] = {
    "ABJC": {"nom": "Servair Abidjan Côte d'Ivoire", "pays": "CI", "secteur": CDI},
    "BBGC": {"nom": "Bridge Bank Group Côte d'Ivoire", "pays": "CI", "secteur": FIN},
    "BICB": {"nom": "BIIC Bénin", "pays": "BJ", "secteur": FIN},
    "BICC": {"nom": "BICI Côte d'Ivoire", "pays": "CI", "secteur": FIN},
    "BNBC": {"nom": "Bernabé Côte d'Ivoire", "pays": "CI", "secteur": CDI},
    "BOAB": {"nom": "Bank of Africa Bénin", "pays": "BJ", "secteur": FIN},
    "BOABF": {"nom": "Bank of Africa Burkina Faso", "pays": "BF", "secteur": FIN},
    "BOAC": {"nom": "Bank of Africa Côte d'Ivoire", "pays": "CI", "secteur": FIN},
    "BOAM": {"nom": "Bank of Africa Mali", "pays": "ML", "secteur": FIN},
    "BOAN": {"nom": "Bank of Africa Niger", "pays": "NE", "secteur": FIN},
    "BOAS": {"nom": "Bank of Africa Sénégal", "pays": "SN", "secteur": FIN},
    "CABC": {"nom": "Sicable Côte d'Ivoire", "pays": "CI", "secteur": IND},
    "CBIBF": {"nom": "Coris Bank International Burkina Faso", "pays": "BF", "secteur": FIN},
    "CFAC": {"nom": "CFAO Motors Côte d'Ivoire", "pays": "CI", "secteur": CDI},
    "CIEC": {"nom": "CIE Côte d'Ivoire", "pays": "CI", "secteur": SPU},
    "ECOC": {"nom": "Ecobank Côte d'Ivoire", "pays": "CI", "secteur": FIN},
    "ETIT": {"nom": "Ecobank Transnational Incorporated", "pays": "TG", "secteur": FIN},
    "FTSC": {"nom": "Filtisac Côte d'Ivoire", "pays": "CI", "secteur": IND},
    "LNBB": {"nom": "Loterie Nationale du Bénin", "pays": "BJ", "secteur": CDI},
    "NEIC": {"nom": "NEI-CEDA Côte d'Ivoire", "pays": "CI", "secteur": CDI},
    "NSBC": {"nom": "NSIA Banque Côte d'Ivoire", "pays": "CI", "secteur": FIN},
    "NTLC": {"nom": "Nestlé Côte d'Ivoire", "pays": "CI", "secteur": CDB},
    "ONTBF": {"nom": "Onatel Burkina Faso", "pays": "BF", "secteur": TEL},
    "ORAC": {"nom": "Orange Côte d'Ivoire", "pays": "CI", "secteur": TEL},
    "ORGT": {"nom": "Oragroup Togo", "pays": "TG", "secteur": FIN},
    "PALC": {"nom": "Palm Côte d'Ivoire", "pays": "CI", "secteur": CDB},
    "PRSC": {"nom": "Tractafric Motors Côte d'Ivoire", "pays": "CI", "secteur": CDI},
    "SAFC": {"nom": "SAFCA Côte d'Ivoire", "pays": "CI", "secteur": FIN},
    "SCRC": {"nom": "Sucrivoire Côte d'Ivoire", "pays": "CI", "secteur": CDB},
    "SDCC": {"nom": "SODECI Côte d'Ivoire", "pays": "CI", "secteur": SPU},
    "SDSC": {"nom": "Africa Global Logistics Côte d'Ivoire", "pays": "CI", "secteur": IND},
    "SEMC": {"nom": "Eviosys Packaging SIEM Côte d'Ivoire", "pays": "CI", "secteur": IND},
    "SGBC": {"nom": "Société Générale Côte d'Ivoire", "pays": "CI", "secteur": FIN},
    "SHEC": {"nom": "Vivo Energy Côte d'Ivoire", "pays": "CI", "secteur": ENE},
    "SIBC": {"nom": "Société Ivoirienne de Banque", "pays": "CI", "secteur": FIN},
    "SICC": {"nom": "SICOR Côte d'Ivoire", "pays": "CI", "secteur": CDB},
    "SIVC": {"nom": "Erium Côte d'Ivoire", "pays": "CI", "secteur": IND},
    "SLBC": {"nom": "Solibra Côte d'Ivoire", "pays": "CI", "secteur": CDB},
    "SMBC": {"nom": "SMB Côte d'Ivoire", "pays": "CI", "secteur": ENE},
    "SNTS": {"nom": "Sonatel", "pays": "SN", "secteur": TEL},
    "SOGC": {"nom": "SOGB Côte d'Ivoire", "pays": "CI", "secteur": CDB},
    "SPHC": {"nom": "SAPH Côte d'Ivoire", "pays": "CI", "secteur": CDB},
    "STAC": {"nom": "SETAO Côte d'Ivoire", "pays": "CI", "secteur": IND},
    "STBC": {"nom": "SITAB Côte d'Ivoire", "pays": "CI", "secteur": CDB},
    "TTLC": {"nom": "TotalEnergies Marketing Côte d'Ivoire", "pays": "CI", "secteur": ENE},
    "TTLS": {"nom": "TotalEnergies Marketing Sénégal", "pays": "SN", "secteur": ENE},
    "UNLC": {"nom": "Unilever Côte d'Ivoire", "pays": "CI", "secteur": CDB},
    "UNXC": {"nom": "Uniwax Côte d'Ivoire", "pays": "CI", "secteur": CDI},
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
