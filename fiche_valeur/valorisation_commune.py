"""Moteur commun de valeur fondamentale CGF Gestion.

Ce fichier est IDENTIQUE dans les deux applications (application SMF « MF_TEST » et application des fiches
de valeur « CGF-Fiche-Valeurs »). Il lit les deux fichiers de référence :

- le classeur des états financiers (une feuille par titre, années en colonnes, unité dans le titre de la
  feuille : FCFA, milliers, millions ou milliards) ;
- la base de données SMF (feuilles Cours, Nb_titres, Historique_dividende, Moyenne_cours, Tableau de bord) ;

et calcule la valeur fondamentale d'un titre à une date d'arrêté. Les deux applications appellent les mêmes
fonctions avec les mêmes paramètres par défaut : elles affichent donc la même valeur. L'empreinte
``empreinte_moteur()`` permet de vérifier que les deux copies du fichier sont identiques.

Méthode (version 1) :
- PER : BPA du dernier exercice clos x PER médian historique du titre (cours moyen annuel / BPA de la
  même année, 5 derniers exercices, borné 4x-30x ; à défaut, PER de référence du secteur). Le cours moyen
  est calculé sur la feuille Cours, ajustée des fractionnements (détectés automatiquement) ;
- P/B : fonds propres par action x P/B médian historique (borné 0,3x-8x ; à défaut, référence du secteur) ;
- actualisation des dividendes à deux étapes : 5 ans à la croissance historique du dividende (0%-10%), puis
  croissance perpétuelle ; un exercice sans dividende rend la méthode non applicable ; un dividende
  supérieur à 2,5 fois la médiane des précédents ET au bénéfice par action est traité comme exceptionnel
  (dividende ordinaire médian retenu) ;
- entreprises : DCF simplifié (flux disponible = résultat d'exploitation après impôt + 20% des dotations,
  croissance historique du résultat d'exploitation bornée 0%-8% ramenée linéairement à la croissance
  perpétuelle sur 5 ans, CMPC borné 7%-20%) ;
- banques : P/B justifié (ROE normatif des 3 derniers exercices - g) / (k - g) à la place du DCF ;
- coût des fonds propres : taux souverain du pays + bêta x prime de risque ; bêta hebdomadaire sur 2 ans
  contre le BRVM Composite, ajusté (Blume), borné 0,6-1,4 ; bêta médian du secteur si la régression n'est
  pas significative ;
- écart minimal de 3 points entre le taux d'actualisation et la croissance perpétuelle ;
- pondérations : entreprises PER 30, P/B 20, dividendes 20, DCF 30 ; banques PER 30, P/B 20,
  dividendes 25, P/B justifié 25 (renormalisées sur les méthodes applicables) ;
- une méthode donnant moins de 0,2 fois ou plus de 5 fois le cours est écartée (donnée source à vérifier) ;
- résultats antérieurs de plus de 3 ans à la date d'arrêté : seule l'actualisation des dividendes est retenue ;
- date d'arrêté : exercices clos à cette date, dividendes payés à cette date (exercice N réputé payé le 1er
  juillet N+1), cours jusqu'à cette date ; elle est plafonnée à la dernière séance de la base SMF.
"""
from __future__ import annotations

import hashlib
import io
import math
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

VERSION = "1.0"

# ---------------------------------------------------------------------------
# Paramètres communs (modifiables dans les deux applications, alignées tant qu'ils restent identiques)
# ---------------------------------------------------------------------------
PARAMS = {
    "prime_risque": 0.055,            # prime de risque actions BRVM
    "croissance_perpetuelle": 0.03,   # croissance nominale de long terme
    "ecart_min_k_g": 0.03,            # écart minimal entre taux d'actualisation et croissance perpétuelle
    "beta_min": 0.6, "beta_max": 1.4,
    "beta_semaines": 104,
    "pe_min": 4.0, "pe_max": 30.0, "pb_min": 0.3, "pb_max": 8.0,
    "g_dividende_max": 0.10, "g_exploitation_max": 0.08,
    "capex_sur_dotations": 0.80,
    "cmpc_min": 0.07, "cmpc_max": 0.20,
    "kd_defaut": 0.07, "impot_defaut": 0.25,
    "garde_fou_bas": 0.2, "garde_fou_haut": 5.0,
    "poids_entreprise": {"PER": 30, "P/B": 20, "Dividendes": 20, "DCF": 30},
    "poids_banque": {"PER": 30, "P/B": 20, "Dividendes": 25, "P/B justifié": 25},
}

# Taux souverains de référence (%) par pays de cotation
TAUX_SOUVERAINS = {"CI": 6.25, "SN": 6.50, "BJ": 6.25, "TG": 6.50, "BF": 7.25, "ML": 7.50, "NE": 7.75, "GW": 7.75}

# Référentiel minimal : pays et secteur (classification BRVM) de chaque titre
_FIN, _TEL, _CDB, _CDI, _IND, _ENE, _SPU = ("Services financiers", "Télécommunications", "Consommation de base",
                                             "Consommation discrétionnaire", "Industriels", "Énergie",
                                             "Services publics")
REFERENTIEL = {
    "ABJC": ("CI", _CDI), "BBGC": ("CI", _FIN), "BICB": ("BJ", _FIN), "BICC": ("CI", _FIN), "BNBC": ("CI", _CDI),
    "BOAB": ("BJ", _FIN), "BOABF": ("BF", _FIN), "BOAC": ("CI", _FIN), "BOAM": ("ML", _FIN), "BOAN": ("NE", _FIN),
    "BOAS": ("SN", _FIN), "CABC": ("CI", _CDI), "CBIBF": ("BF", _FIN), "CFAC": ("CI", _CDI), "CIEC": ("CI", _SPU),
    "ECOC": ("CI", _FIN), "ETIT": ("TG", _FIN), "FTSC": ("CI", _IND), "LNBB": ("BJ", _CDI), "NEIC": ("CI", _CDI),
    "NSBC": ("CI", _FIN), "NTLC": ("CI", _CDB), "ONTBF": ("BF", _TEL), "ORAC": ("CI", _TEL), "ORGT": ("TG", _FIN),
    "PALC": ("CI", _CDB), "PRSC": ("CI", _CDI), "SAFC": ("CI", _FIN), "SCRC": ("CI", _CDB), "SDCC": ("CI", _SPU),
    "SDSC": ("CI", _IND), "SEMC": ("CI", _CDB), "SGBC": ("CI", _FIN), "SHEC": ("CI", _ENE), "SIBC": ("CI", _FIN),
    "SICC": ("CI", _CDB), "SIVC": ("CI", _IND), "SLBC": ("CI", _CDB), "SMBC": ("CI", _ENE), "SNTS": ("SN", _TEL),
    "SOGC": ("CI", _CDB), "SPHC": ("CI", _CDB), "STAC": ("CI", _IND), "STBC": ("CI", _CDB), "SVOC": ("CI", _CDI),
    "TTLC": ("CI", _ENE), "TTLS": ("SN", _ENE), "UNLC": ("CI", _CDB), "UNXC": ("CI", _CDI),
}
# Multiples de référence du secteur, utilisés faute d'historique propre suffisant
MULTIPLES_SECTEUR = {
    _FIN: (7.5, 1.0), _TEL: (10.5, 3.0), _CDB: (10.0, 1.5), _CDI: (9.0, 1.3), _IND: (9.0, 1.2),
    _ENE: (9.0, 1.8), _SPU: (10.0, 2.5),
}
ALIAS_FEUILLES = {"BIIC": "BICB", "ETI": "ETIT", "SOGB": "SOGC"}
INDICE = ".BRVMCI"
NON_TITRES = {"DATE", "ANNEE", "JOUR", "TOTAL", "MOYENNE", "BRVM30", "BRVM PREST", "BRVM-PRINC", "BRVM-C TR",
              "BRVM-CB", "BRVM-CD", "BRVM-ENER", "BRVM-SFIN", "BRVM-SPUB"}


def empreinte_moteur() -> str:
    """Empreinte du présent fichier (8 caractères) : identique dans les deux applications si elles sont alignées."""
    return hashlib.sha256(Path(__file__).read_bytes().replace(b"\r\n", b"\n")).hexdigest()[:8]


def pays_secteur(ticker: str) -> tuple[str, str]:
    return REFERENTIEL.get(ticker, ("CI", "Non classé"))


def dernier_exercice_clos(d: date) -> int:
    return d.year if (d.month, d.day) == (12, 31) else d.year - 1


def _num(v) -> bool:
    return isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool) and math.isfinite(float(v))


def _norm(s) -> str:
    s = unicodedata.normalize("NFD", str(s)).encode("ascii", "ignore").decode().lower()
    s = re.sub(r"[^a-z0-9/%]+", " ", s)
    return " ".join(s.split())


# ---------------------------------------------------------------------------
# Lecture du classeur des états financiers
# ---------------------------------------------------------------------------
_LIB_BANQUE = {
    "pnb": ["produit net bancaire"],
    "charges_generales": ["charges generales d exploitation", "frais generaux"],
    "rbe": ["resultat brut d exploitation"],
    "cout_risque": ["cout net du risque", "cout du risque"],
    "rex": ["resultat d exploitation"],
    "rai": ["resultat avant impot", "resultat avant impots"],
    "impots": ["impots sur les benefices", "impot sur les societes", "impots sur le resultat"],
    "rn": ["resultat net", "resultat net consolide"],
    "rn_mere": ["attributable proprietaires mere", "rn attribuable proprietaires mere"],
    "total_actif": ["total actif"],
    "cp": ["capitaux propres et ressources assimilees", "total capitaux propres", "total des capitaux propres",
           "capitaux propres"],
    "cp_mere": ["capitaux propres part du groupe", "capitaux propres mere"],
    "creances_clientele": ["creances sur la clientele", "prets et creances sur la clientele",
                           "dont prets et creances sur la clientele"],
    "depots_clientele": ["dettes a l egard de la clientele", "depots de la clientele",
                         "dont dettes envers la clientele"],
    "interets_charges": ["interets et charges assimilees", "charges d interets"],
    "dettes_interbancaires": ["dettes interbancaires et assimilees"],
    "tresorerie_actif": ["caisse banque centrale ccp", "caisse et banques centrales",
                         "tresorerie et equivalents de tresorerie"],
}
_LIB_ENTREPRISE = {
    "ca": ["chiffre d affaires"],
    "va": ["valeur ajoutee"],
    "ebe": ["excedent brut d exploitation", "ebe", "ebitdaal"],
    "dotations": ["dotations amort /prov", "dotations aux amortissements",
                  "dotations amortissements et provisions nettes"],
    "rex": ["resultat d exploitation"],
    "frais_financiers": ["frais financiers"],
    "revenus_financiers": ["revenus financiers"],
    "resultat_financier": ["resultat financier"],
    "rao": ["resultat activites ordinaires", "resultat des activites ordinaires"],
    "hao": ["resultat hao"],
    "impots": ["impots sur le resultat", "impot sur les societes", "impots sur les benefices"],
    "rn": ["resultat net", "resultat net consolide"],
    "rn_mere": ["attributable proprietaires mere", "rn attribuable proprietaires mere"],
    "total_actif": ["total actif"],
    "cp": ["capitaux propres", "total capitaux propres", "total des capitaux propres"],
    "cp_mere": ["capitaux propres mere", "capitaux propres part du groupe"],
    "dettes_financieres": ["dettes financieres", "dettes financieres et ressources assimilees"],
    "passifs_fin_nc": ["passifs financiers non courants"],
    "passifs_fin_c": ["passifs financiers courants"],
    "tresorerie_actif": ["tresorerie actif", "disponibilites", "disponibilites et quasi disponibilites",
                         "tresorerie et equivalents de tresorerie"],
    "tresorerie_passif": ["tresorerie passif"],
}
_PAR_ACTION = {"dps": ["dividende/action fcfa"],
               "bpa": ["rn/action de base fcfa", "rn par action fcfa", "resultat net par action fcfa"]}


@dataclass
class Societe:
    ticker: str
    feuille: str
    nom: str
    banque: bool
    unite: str
    facteur: float
    annees: dict = field(default_factory=dict)      # année -> {clé: valeur en FCFA (par action : en FCFA)}

    def valeur(self, cle: str, annee: int):
        return self.annees.get(annee, {}).get(cle)

    def rn(self, annee: int):
        v = self.valeur("rn_mere", annee)
        return v if _num(v) else self.valeur("rn", annee)

    def cp(self, annee: int):
        v = self.valeur("cp_mere", annee)
        return v if _num(v) else self.valeur("cp", annee)


def _facteur(titre: str) -> tuple[str, float]:
    t = _norm(titre)
    if "milliard" in t:
        return "milliards FCFA", 1e9
    if "million" in t:
        return "millions FCFA", 1e6
    if "millier" in t:
        return "milliers FCFA", 1e3
    return "FCFA", 1.0


class _Evaluateur:
    """Évalue les formules arithmétiques simples d'une feuille (+ - * /, SUM, références de cellules)
    lorsque le classeur n'a pas de valeurs calculées enregistrées."""

    def __init__(self, ws_f, ws_v):
        self.f, self.v, self.cache = ws_f, ws_v, {}

    def get(self, coord: str, depth: int = 0):
        if coord in self.cache:
            return self.cache[coord]
        val = self.v[coord].value
        if _num(val):
            self.cache[coord] = float(val)
            return float(val)
        form = self.f[coord].value
        res = None
        if isinstance(form, str) and form.startswith("=") and depth < 40:
            res = self._eval(form[1:], depth)
        elif _num(form):
            res = float(form)
        self.cache[coord] = res
        return res

    def _eval(self, expr: str, depth: int):
        from openpyxl.utils import range_boundaries, get_column_letter
        e = expr.replace("$", "").lstrip("+")
        if re.search(r"[A-Z]{2,}\(", e.upper().replace("SUM(", "")):
            return None                                  # fonction non gérée

        def rng(m):
            c1, r1, c2, r2 = range_boundaries(m.group(1))
            tot = 0.0
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    x = self.get(f"{get_column_letter(c)}{r}", depth + 1)
                    tot += x if _num(x) else 0.0
            return repr(tot)
        e = re.sub(r"(?i)SUM\(([A-Z]+\d+:[A-Z]+\d+)\)", rng, e)

        def ref(m):
            x = self.get(m.group(0).upper(), depth + 1)
            return repr(x) if _num(x) else "0.0"
        e = re.sub(r"(?<![A-Za-z])[A-Za-z]{1,3}\d+", ref, e)
        if not re.fullmatch(r"[0-9eE.+\-*/() ]+", e):
            return None
        try:
            return float(eval(e, {"__builtins__": {}}, {}))
        except Exception:
            return None


def lire_etats_financiers(source) -> dict[str, Societe]:
    """Classeur des états financiers -> {ticker: Societe}. ``source`` : chemin ou octets."""
    import openpyxl
    data = source if isinstance(source, (bytes, bytearray)) else Path(source).read_bytes()
    wf = openpyxl.load_workbook(io.BytesIO(data), data_only=False)
    wv = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    out: dict[str, Societe] = {}
    for ws in wf.worksheets:
        name = ws.title.strip()
        if _norm(name).startswith("synthese") or len(name) > 6:
            continue
        ticker = ALIAS_FEUILLES.get(name.upper(), name.upper())
        hdr_row, years = None, {}
        for r in range(1, min(ws.max_row, 8) + 1):
            ys = {}
            for c in range(2, ws.max_column + 1):
                v = ws.cell(r, c).value
                s = str(v).strip() if v is not None else ""
                if re.fullmatch(r"(19|20)\d{2}(\.0)?", s):
                    ys[c] = int(float(s))
            if len(ys) >= 1:
                hdr_row, years = r, ys
                break
        if hdr_row is None:
            continue
        unite, fac = _facteur(ws.cell(1, 1).value or "")
        ev = _Evaluateur(ws, wv[ws.title])
        labels = {r: _norm(ws.cell(r, 1).value) for r in range(hdr_row + 1, ws.max_row + 1) if ws.cell(r, 1).value}
        banque = any(lab == "produit net bancaire" for lab in labels.values())
        table = _LIB_BANQUE if banque else _LIB_ENTREPRISE
        rows: dict[str, int] = {}
        for key, pats in list(table.items()) + list(_PAR_ACTION.items()):
            for p in pats:                               # priorité à l'ordre des libellés, puis à l'ordre des lignes
                hit = next((r for r, lab in labels.items() if lab == p), None)
                if hit is not None:
                    rows[key] = hit
                    break
        soc = Societe(ticker=ticker, feuille=ws.title, nom=str(ws.cell(1, 1).value or "").split("—")[0].split(" - ")[0].strip(),
                      banque=banque, unite=unite, facteur=fac)
        from openpyxl.utils import get_column_letter
        for c, y in years.items():
            vals = {}
            for key, r in rows.items():
                x = ev.get(f"{get_column_letter(c)}{r}")
                if _num(x):
                    vals[key] = float(x) * (1.0 if key in _PAR_ACTION else fac)
            if not banque and "dettes_financieres" not in vals and ("passifs_fin_nc" in vals or "passifs_fin_c" in vals):
                vals["dettes_financieres"] = vals.get("passifs_fin_nc", 0.0) + vals.get("passifs_fin_c", 0.0)
            if vals:
                soc.annees[y] = vals
        if soc.annees:
            out[ticker] = soc
    return out


# ---------------------------------------------------------------------------
# Lecture de la base de données SMF
# ---------------------------------------------------------------------------
@dataclass
class BaseMarche:
    cours: pd.DataFrame                                   # index Date, colonnes tickers (dont .BRVMCI)
    nb_titres: dict = field(default_factory=dict)
    flottant: dict = field(default_factory=dict)
    dividendes: dict = field(default_factory=dict)        # ticker -> {exercice: dividende net (0 = aucun)}
    moyenne_cours: dict = field(default_factory=dict)     # ticker -> {année: cours moyen}
    tableau_de_bord: dict = field(default_factory=dict)   # ticker -> {indicateur: {année: FCFA}}
    cours_ajustes: pd.DataFrame | None = None             # cours ajustés des fractionnements d'actions
    fractionnements: list = field(default_factory=list)   # (ticker, date, facteur)

    @property
    def derniere_date(self) -> date | None:
        return None if self.cours is None or self.cours.empty else pd.Timestamp(self.cours.index.max()).date()


FACTEURS_FRACTIONNEMENT = (2, 3, 4, 5, 10, 20, 25, 50, 100)


def ajuster_fractionnements(cours: pd.DataFrame) -> tuple[pd.DataFrame, list]:
    """Détecte les fractionnements d'actions (baisse d'un facteur usuel en une séance, sans retour ensuite ; la
    variation quotidienne de la BRVM étant plafonnée, un tel saut ne peut venir du marché) et ajuste les cours
    antérieurs. Les dividendes de la base sont supposés déjà ajustés."""
    adj = cours.copy()
    found = []
    for t in cours.columns:
        if t == INDICE:
            continue
        s = cours[t].dropna()
        if len(s) < 20:
            continue
        ratio = s.shift(1) / s
        for d, r in ratio[ratio >= 1.8].items():
            n = min(FACTEURS_FRACTIONNEMENT, key=lambda f: abs(math.log(r / f)))
            if abs(math.log(r / n)) > math.log(1.15):
                continue
            prev = s[s.index < d].iloc[-1]
            after = s[s.index >= d].iloc[:10]
            if after.max() > prev / 1.5:
                continue                                   # retour rapide : erreur de saisie, pas un fractionnement
            adj.loc[adj.index < d, t] = adj.loc[adj.index < d, t] / n
            found.append((t, pd.Timestamp(d).date(), n))
    return adj, found


def _est_titre(c) -> bool:
    s = str(c).strip()
    return s == INDICE or (bool(re.fullmatch(r"[A-Z]{3,6}", s)) and s not in NON_TITRES)


def lire_base_smf(source) -> BaseMarche:
    data = source if isinstance(source, (bytes, bytearray)) else Path(source).read_bytes()
    xl = pd.ExcelFile(io.BytesIO(data))
    raw = pd.read_excel(xl, sheet_name="Cours")
    raw["Date"] = pd.to_datetime(raw["Date"], errors="coerce")
    raw = raw.dropna(subset=["Date"]).set_index("Date").sort_index()
    cols = [c for c in raw.columns if _est_titre(c)]
    cours = raw[cols].apply(pd.to_numeric, errors="coerce")
    cours = cours.where(cours > 0)
    cours = cours[~cours.index.duplicated(keep="last")]
    base = BaseMarche(cours=cours)
    base.cours_ajustes, base.fractionnements = ajuster_fractionnements(cours)
    # Nombre de titres
    nb = pd.read_excel(xl, sheet_name="Nb_titres", header=None)
    col_t = {str(nb.iloc[0, c]).strip(): c for c in range(2, nb.shape[1]) if _est_titre(nb.iloc[0, c])}
    for r in range(1, nb.shape[0]):
        lab = _norm(nb.iloc[r, 1])
        for t, c in col_t.items():
            v = pd.to_numeric(nb.iloc[r, c], errors="coerce")
            if _num(v) and v > 0:
                (base.flottant if "flottant" in lab else base.nb_titres if "nombre" in lab else {})[t] = float(v)
    # Dividendes (une ligne par exercice ; cellule vide d'un exercice renseigné pour au moins 3 titres = aucun)
    dv = pd.read_excel(xl, sheet_name="Historique_dividende")
    dv = dv.dropna(subset=["Date"])
    tick = [c for c in dv.columns if _est_titre(c) and c != INDICE]
    for _, row in dv.iterrows():
        try:
            y = int(float(row["Date"]))
        except (TypeError, ValueError):
            continue
        vals = {t: pd.to_numeric(row.get(t), errors="coerce") for t in tick}
        filled = sum(1 for v in vals.values() if _num(v))
        for t, v in vals.items():
            if _num(v):
                base.dividendes.setdefault(t, {})[y] = max(0.0, float(v))
            elif filled >= 3:
                base.dividendes.setdefault(t, {})[y] = 0.0
    # Cours moyens annuels
    try:
        mc = pd.read_excel(xl, sheet_name="Moyenne_cours").dropna(subset=["Date"])
        for _, row in mc.iterrows():
            try:
                y = int(float(row["Date"]))
            except (TypeError, ValueError):
                continue
            for t in mc.columns:
                v = pd.to_numeric(row.get(t), errors="coerce")
                if _est_titre(t) and _num(v) and v > 0:
                    base.moyenne_cours.setdefault(t, {})[y] = float(v)
    except Exception:
        pass
    # Tableau de bord (indicateurs annuels en FCFA)
    try:
        tb = pd.read_excel(xl, sheet_name="Tableau de bord", header=None)
        col_t = {str(tb.iloc[0, c]).strip(): c for c in range(2, tb.shape[1]) if _est_titre(tb.iloc[0, c])}
        bloc = None
        for r in range(1, tb.shape[0]):
            y = pd.to_numeric(tb.iloc[r, 1], errors="coerce")
            lab = tb.iloc[r, 2]
            if not _num(y):
                if isinstance(lab, str) and not _est_titre(lab):
                    bloc = _norm(lab)
                continue
            if bloc is None:
                continue
            for t, c in col_t.items():
                v = pd.to_numeric(tb.iloc[r, c], errors="coerce")
                if _num(v):
                    base.tableau_de_bord.setdefault(t, {}).setdefault(bloc, {})[int(y)] = float(v)
    except Exception:
        pass
    return base


# ---------------------------------------------------------------------------
# Valorisation
# ---------------------------------------------------------------------------
@dataclass
class Methode:
    nom: str
    valeur: float
    poids: float
    detail: str


@dataclass
class Resultat:
    ticker: str
    date_arrete: date
    cours: float | None = None
    valeur: float | None = None
    potentiel: float | None = None
    methodes: list = field(default_factory=list)          # méthodes retenues
    ecartees: list = field(default_factory=list)          # (nom, raison)
    annee_ref: int | None = None
    bpa: float | None = None
    fpa: float | None = None
    dpa: float | None = None
    dpa_exercice: int | None = None
    beta: float | None = None
    beta_source: str = ""
    k: float | None = None
    g: float | None = None
    cmpc: float | None = None
    pe_cible: float | None = None
    pe_source: str = ""
    pb_cible: float | None = None
    pb_source: str = ""
    notes: list = field(default_factory=list)
    version: str = VERSION

    @property
    def fourchette(self):
        vals = [m.valeur for m in self.methodes]
        return (min(vals), max(vals)) if vals else (None, None)


def _cagr(serie: dict, n_max: int = 5) -> float | None:
    ys = sorted(y for y, v in serie.items() if _num(v) and v > 0)[-(n_max + 1):]
    if len(ys) < 3 or ys[-1] == ys[0]:
        return None
    return (serie[ys[-1]] / serie[ys[0]]) ** (1 / (ys[-1] - ys[0])) - 1


def cours_a_date(base: BaseMarche, ticker: str, d: date) -> float | None:
    if base is None or ticker not in base.cours:
        return None
    s = base.cours[ticker][base.cours.index <= pd.Timestamp(d)].dropna()
    return float(s.iloc[-1]) if not s.empty else None


def _beta_regression(base: BaseMarche, ticker: str, d: date, p: dict) -> tuple[float | None, float | None]:
    if base is None or ticker not in base.cours or INDICE not in base.cours:
        return None, None
    src = base.cours_ajustes if base.cours_ajustes is not None else base.cours
    df = src[[ticker, INDICE]]
    df = df[(df.index <= pd.Timestamp(d)) & (df.index > pd.Timestamp(d) - pd.Timedelta(weeks=p["beta_semaines"]))]
    w = df.ffill().resample("W-FRI").last().pct_change().dropna()
    w = w[(w[ticker] != 0)]
    if len(w) < 40 or w[INDICE].var() == 0:
        return None, None
    corr = float(w[ticker].corr(w[INDICE]))
    b = float(w[ticker].cov(w[INDICE]) / w[INDICE].var())
    return b, corr


def beta(base: BaseMarche, ticker: str, d: date, p: dict | None = None) -> tuple[float, str]:
    p = p or PARAMS
    b, corr = _beta_regression(base, ticker, d, p)
    if b is not None and corr is not None and corr >= 0.15:
        return min(p["beta_max"], max(p["beta_min"], 0.67 * b + 0.33)), "régression hebdomadaire sur 2 ans (Blume)"
    sect = pays_secteur(ticker)[1]
    bs = []
    for t, (_, s) in REFERENTIEL.items():
        if s == sect and t != ticker:
            bt, ct = _beta_regression(base, t, d, p)
            if bt is not None and ct is not None and ct >= 0.15:
                bs.append(0.67 * bt + 0.33)
    if len(bs) >= 2:
        return (min(p["beta_max"], max(p["beta_min"], float(np.median(bs)))),
                f"bêta médian du secteur ({len(bs)} titres), régression du titre non significative")
    return 1.0, "valeur neutre (régression non significative, secteur insuffisant)"


def _cours_moyen(base: BaseMarche, ticker: str, y: int, d: date) -> float | None:
    # moyenne des séances de la feuille Cours (ajustée des fractionnements), plus sûre que la feuille Moyenne_cours
    src = base.cours_ajustes if base.cours_ajustes is not None else base.cours
    if ticker in src:
        s = src[ticker]
        s = s[(s.index.year == y) & (s.index <= pd.Timestamp(d))].dropna()
        if not s.empty:
            return float(s.mean())
    return None


def _multiple_historique(soc: Societe, base, ticker, shares, y_ref, d, kind, p):
    obs = []
    for y in sorted(soc.annees):
        if y > y_ref or y < y_ref - 4:
            continue
        num = soc.rn(y) if kind == "pe" else soc.cp(y)
        if not _num(num) or num <= 0:
            continue
        px = _cours_moyen(base, ticker, y, d)
        if not _num(px):
            continue
        m = px / (num / shares)
        if (1 < m < 100) if kind == "pe" else (0.1 < m < 20):
            obs.append(m)
    lo, hi = (p["pe_min"], p["pe_max"]) if kind == "pe" else (p["pb_min"], p["pb_max"])
    if len(obs) >= 2:
        return min(hi, max(lo, float(np.median(obs)))), f"médiane historique du titre ({len(obs)} exercices)"
    ref = MULTIPLES_SECTEUR.get(pays_secteur(ticker)[1], (9.0, 1.3))
    return (ref[0] if kind == "pe" else ref[1]), f"référence du secteur {pays_secteur(ticker)[1]} (historique insuffisant)"


def dividende_retenu(base: BaseMarche, ticker: str, d: date, bpa_par_exercice: dict | None = None):
    """(dividende retenu, exercice, note) : dernier exercice dont le dividende est connu à la date d'arrêté
    (réputé payé le 1er juillet de l'année suivante)."""
    divs = {y: v for y, v in (base.dividendes.get(ticker, {}) if base else {}).items() if date(y + 1, 7, 1) <= d}
    if not divs:
        return None, None, ""
    y = max(divs)
    d0 = divs[y]
    if d0 <= 0:
        return 0.0, y, f"aucun dividende au titre de l'exercice {y}"
    prev = [divs[k] for k in sorted(divs) if k < y and divs[k] > 0][-4:]
    eps = (bpa_par_exercice or {}).get(y)
    if len(prev) >= 2 and d0 > 2.5 * float(np.median(prev)) and (not _num(eps) or d0 > eps):
        med = float(np.median(prev))
        return med, y, (f"dividende {y} de {d0:,.0f} FCFA jugé exceptionnel (plus de 2,5 fois les précédents et "
                        f"supérieur au bénéfice par action) : dividende ordinaire de {med:,.0f} FCFA retenu").replace(",", " ")
    return d0, y, ""


# ---------------------------------------------------------------------------
# Sérialisation (persistance des données entre deux sessions, utilisée par l'application SMF)
# ---------------------------------------------------------------------------
_ALIAS_ANCIENS = {"cp": "capitaux_propres", "cp_mere": "capitaux_propres_mere", "dettes_financieres": "dette_financiere",
                  "tresorerie_actif": "tresorerie", "dotations": "amortissements", "bpa": "eps", "rex": "rex",
                  "creances_clientele": "creances_clientele", "interets_charges": "interets_charges",
                  "revenus_financiers": "revenus_financiers", "frais_financiers": "interets_charges"}


def ef_vers_json(ef: dict[str, Societe]) -> str:
    import json
    return json.dumps({"version": VERSION, "societes": {
        t: {"feuille": s.feuille, "nom": s.nom, "banque": s.banque, "unite": s.unite, "facteur": s.facteur,
            "annees": {str(y): v for y, v in s.annees.items()}} for t, s in ef.items()}}, ensure_ascii=False)


def ef_depuis_json(txt: str) -> dict[str, Societe]:
    import json
    d = json.loads(txt)
    return {t: Societe(ticker=t, feuille=v["feuille"], nom=v["nom"], banque=v["banque"], unite=v["unite"],
                       facteur=v["facteur"], annees={int(y): x for y, x in v["annees"].items()})
            for t, v in d.get("societes", {}).items()}


def ef_vers_fin_data(ef: dict[str, Societe]) -> dict:
    """Format {ticker: {année: postes}} des anciens modules de l'application SMF (screening Charia), alimenté par
    le lecteur commun."""
    out = {}
    for t, s in ef.items():
        for y, v in s.annees.items():
            p = dict(v)
            for k, alias in _ALIAS_ANCIENS.items():
                if k in v and alias not in p:
                    p[alias] = v[k]
            if s.banque and "pnb" in v:
                p["ca"] = v["pnb"]
            p.update({"type": "banque" if s.banque else "societe", "unite": s.unite,
                      "secteur": pays_secteur(t)[1], "source": f"moteur commun v{VERSION}"})
            out.setdefault(t, {})[y] = p
    return out


def base_vers_json(base: BaseMarche) -> str:
    import json
    def frame(df):
        if df is None:
            return None
        x = df.copy()
        x.index = [pd.Timestamp(i).strftime("%Y-%m-%d") for i in x.index]
        return x.to_json(orient="split")
    return json.dumps({"version": VERSION, "cours": frame(base.cours), "nb_titres": base.nb_titres,
                       "flottant": base.flottant,
                       "dividendes": {t: {str(y): v for y, v in d.items()} for t, d in base.dividendes.items()},
                       "moyenne_cours": {t: {str(y): v for y, v in d.items()} for t, d in base.moyenne_cours.items()},
                       "tableau_de_bord": {t: {k: {str(y): v for y, v in s.items()} for k, s in d.items()}
                                           for t, d in base.tableau_de_bord.items()}})


def base_depuis_json(txt: str) -> BaseMarche:
    import json
    d = json.loads(txt)
    cours = pd.read_json(io.StringIO(d["cours"]), orient="split")
    cours.index = pd.to_datetime(cours.index)
    cours = cours.sort_index().astype(float)
    base = BaseMarche(cours=cours, nb_titres=d["nb_titres"], flottant=d.get("flottant", {}),
                      dividendes={t: {int(y): v for y, v in x.items()} for t, x in d["dividendes"].items()},
                      moyenne_cours={t: {int(y): v for y, v in x.items()} for t, x in d["moyenne_cours"].items()},
                      tableau_de_bord={t: {k: {int(y): v for y, v in s.items()} for k, s in x.items()}
                                       for t, x in d.get("tableau_de_bord", {}).items()})
    base.cours_ajustes, base.fractionnements = ajuster_fractionnements(base.cours)
    return base


def valeur_fondamentale(ticker: str, ef: dict[str, Societe] | None, base: BaseMarche | None,
                        date_arrete: date | None = None, params: dict | None = None,
                        taux_sans_risque: float | None = None) -> Resultat:
    """Valeur fondamentale d'un titre. ``taux_sans_risque`` (décimal) remplace le taux souverain du pays."""
    p = dict(PARAMS)
    p.update(params or {})
    d = date_arrete or (base.derniere_date if base is not None else None) or date.today()
    if base is not None and base.derniere_date and d > base.derniere_date:
        d = base.derniere_date                            # date d'arrêté plafonnée à la dernière séance de la base
    res = Resultat(ticker=ticker, date_arrete=d)
    soc = (ef or {}).get(ticker)
    ly = dernier_exercice_clos(d)
    res.cours = cours_a_date(base, ticker, d)
    shares = base.nb_titres.get(ticker) if base is not None else None
    pays, secteur = pays_secteur(ticker)
    rf = taux_sans_risque if taux_sans_risque is not None else TAUX_SOUVERAINS.get(pays, 6.5) / 100
    res.beta, res.beta_source = beta(base, ticker, d, p) if base is not None else (1.0, "valeur neutre")
    k = rf + res.beta * p["prime_risque"]
    g = min(p["croissance_perpetuelle"], k - p["ecart_min_k_g"])
    res.k, res.g = k, g
    if not _num(shares) or shares <= 0:
        res.notes.append("Nombre de titres absent de la base SMF : valorisation impossible.")
        return res
    years = sorted(y for y in (soc.annees if soc else {}) if y <= ly and _num(soc.rn(y)))
    banque = bool(soc and soc.banque) or secteur == _FIN
    poids = p["poids_banque"] if banque else p["poids_entreprise"]
    cands: list[Methode] = []
    bpa_ex = {y: soc.rn(y) / shares for y in years} if soc else {}
    if years:
        y = years[-1]
        res.annee_ref = y
        rn, cp = soc.rn(y), soc.cp(y)
        res.bpa = rn / shares
        res.fpa = cp / shares if _num(cp) else None
        stale = y < d.year - 3
        if stale:
            res.notes.append(f"Derniers résultats connus datant de {y} : seules les méthodes fondées sur le "
                             f"dividende sont retenues.")
        if not stale and res.bpa > 0:
            res.pe_cible, res.pe_source = _multiple_historique(soc, base, ticker, shares, y, d, "pe", p)
            cands.append(Methode("PER", res.bpa * res.pe_cible, poids["PER"],
                                 f"BPA {y} {res.bpa:,.0f} FCFA x PER {res.pe_cible:.1f}x ({res.pe_source})"))
        if not stale and _num(res.fpa) and res.fpa > 0:
            res.pb_cible, res.pb_source = _multiple_historique(soc, base, ticker, shares, y, d, "pb", p)
            cands.append(Methode("P/B", res.fpa * res.pb_cible, poids["P/B"],
                                 f"Fonds propres par action {res.fpa:,.0f} FCFA x P/B {res.pb_cible:.2f}x ({res.pb_source})"))
        if not stale and banque and "P/B justifié" in poids:
            roes = []
            for yy in years[-3:]:
                c1 = soc.cp(yy)
                if _num(c1) and c1 > 0:
                    roes.append(soc.rn(yy) / c1)
            if roes and _num(res.fpa) and res.fpa > 0:
                roe = float(np.mean(roes))
                if roe > g:
                    pbj = (roe - g) / (k - g)
                    cands.append(Methode("P/B justifié", res.fpa * pbj, poids["P/B justifié"],
                                         f"(ROE normatif {roe:.1%} - g {g:.1%}) / (k {k:.1%} - g) = {pbj:.2f}x"))
                else:
                    res.ecartees.append(("P/B justifié", f"ROE normatif {roe:.1%} inférieur à la croissance"))
        if not stale and not banque and "DCF" in poids:
            m = _dcf(soc, y, years, k, g, shares, p, res)
            if m is not None:
                cands.append(Methode("DCF", m[0], poids["DCF"], m[1]))
    elif soc is None:
        res.notes.append("Titre absent du classeur des états financiers : seules les méthodes fondées sur le "
                         "dividende sont applicables.")
    else:
        res.notes.append(f"Aucun exercice clos au plus tard en {ly} dans le classeur des états financiers.")
    # Dividendes
    d0, dy, dnote = dividende_retenu(base, ticker, d, bpa_ex)
    res.dpa, res.dpa_exercice = d0, dy
    if dnote:
        res.notes.append(dnote[0].upper() + dnote[1:] + ".")
    if _num(d0) and d0 > 0:
        divs = {yy: v for yy, v in base.dividendes.get(ticker, {}).items() if yy <= dy and v > 0}
        g1 = _cagr(divs)
        g1 = min(p["g_dividende_max"], max(0.0, g1)) if g1 is not None else g
        pv, div = 0.0, d0
        for t in range(1, 6):
            div *= 1 + g1
            pv += div / (1 + k) ** t
        pv += div * (1 + g) / (k - g) / (1 + k) ** 5
        cands.append(Methode("Dividendes", pv, poids["Dividendes"],
                             f"DPA {dy} {d0:,.0f} FCFA, croissance {g1:.1%} sur 5 ans puis {g:.1%}, k {k:.1%}"))
    elif d0 == 0:
        res.ecartees.append(("Dividendes", dnote or "aucun dividende"))
    # Garde-fou et combinaison
    retenues = cands
    if _num(res.cours) and res.cours > 0:
        lo, hi = p["garde_fou_bas"] * res.cours, p["garde_fou_haut"] * res.cours
        retenues = [m for m in cands if lo <= m.valeur <= hi]
        for m in cands:
            if m not in retenues:
                res.ecartees.append((m.nom, f"valeur {m.valeur:,.0f} FCFA hors de [{p['garde_fou_bas']}x ; "
                                            f"{p['garde_fou_haut']}x] le cours : donnée source à vérifier".replace(",", " ")))
    retenues = [m for m in retenues if _num(m.valeur) and m.valeur > 0 and m.poids > 0]
    res.methodes = retenues
    if retenues:
        tw = sum(m.poids for m in retenues)
        res.valeur = sum(m.valeur * m.poids for m in retenues) / tw
        if _num(res.cours) and res.cours > 0:
            res.potentiel = res.valeur / res.cours - 1
    else:
        res.notes.append("Aucune méthode applicable : valeur fondamentale indéterminable.")
    for m in res.methodes:
        m.detail = m.detail.replace(",", " ")
    return res


def _dcf(soc: Societe, y: int, years: list, k: float, g: float, shares: float, p: dict, res: Resultat):
    v = soc.annees[y]
    rex = v.get("rex")
    if not _num(rex) or rex <= 0:
        res.ecartees.append(("DCF", "résultat d'exploitation négatif ou absent"))
        return None
    base_imp = v.get("rao") if _num(v.get("rao")) and v.get("rao") > 0 else v.get("rex")
    imp = abs(v["impots"]) if _num(v.get("impots")) else None
    t = min(0.40, max(0.10, imp / base_imp)) if imp and base_imp else p["impot_defaut"]
    dot = abs(v.get("dotations") or 0.0)
    fcf0 = rex * (1 - t) + dot * (1 - p["capex_sur_dotations"])
    debt = abs(v.get("dettes_financieres") or 0.0) + abs(v.get("tresorerie_passif") or 0.0)
    cash = abs(v.get("tresorerie_actif") or 0.0)
    eq = soc.cp(y)
    ff = abs(v.get("frais_financiers") or 0.0)
    kd = min(0.15, max(0.03, ff / debt)) if debt > 0 and ff > 0 else p["kd_defaut"]
    if _num(eq) and eq > 0:
        wacc = (k * eq + kd * (1 - t) * debt) / (eq + debt)
    else:
        wacc = k
    wacc = min(p["cmpc_max"], max(p["cmpc_min"], wacc))
    gt = min(g, wacc - p["ecart_min_k_g"])
    g1 = _cagr({yy: soc.valeur("rex", yy) for yy in years})
    g1 = min(p["g_exploitation_max"], max(0.0, g1)) if g1 is not None else gt
    pv, f = 0.0, fcf0
    for i in range(1, 6):
        gi = g1 + (gt - g1) * (i - 1) / 4
        f *= 1 + gi
        pv += f / (1 + wacc) ** i
    pv += f * (1 + gt) / (wacc - gt) / (1 + wacc) ** 5
    equity = pv - debt + cash
    res.cmpc = wacc
    if equity <= 0:
        res.ecartees.append(("DCF", "valeur d'entreprise inférieure à la dette nette"))
        return None
    return equity / shares, (f"Flux disponible {fcf0 / 1e9:,.1f} Md FCFA, croissance {g1:.1%} ramenée à {gt:.1%}, "
                             f"CMPC {wacc:.1%}, dette nette {(debt - cash) / 1e9:,.1f} Md FCFA")


def valoriser_tout(ef, base, date_arrete=None, params=None, tickers=None) -> dict[str, Resultat]:
    tickers = tickers or sorted(set(ef or {}) | {t for t in (base.nb_titres if base else {})})
    return {t: valeur_fondamentale(t, ef, base, date_arrete, params) for t in tickers}
