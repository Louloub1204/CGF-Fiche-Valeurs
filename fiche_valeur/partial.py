"""Titres à données incomplètes : chiffres clés, saisie manuelle et mesure de la couverture.

Trois niveaux d'analyse sont gérés :

* niveau 1 : états financiers détaillés du classeur ;
* niveau 2 : chiffres clés (chiffre d'affaires, résultat net, BNPA, dividendes, éventuellement
  capitaux propres et total du bilan), lus sur sikafinance ou saisis dans l'application ;
* niveau 3 : données de marché seules (cours, volumes, dividendes).

Chaque niveau fait tourner les mêmes moteurs d'analyse ; ce qui manque est neutralisé et signalé,
et la conviction du signal est plafonnée en conséquence.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

import pandas as pd

from . import config
from .parser import CompanyData, normalise_sector
from .utils import is_num, norm_label, to_float

SIKA_SOCIETE = "https://www.sikafinance.com/marches/societe/{sym}"


# ---------------------------------------------------------------------------
# Chiffres clés
# ---------------------------------------------------------------------------
@dataclass
class KeyFigures:
    """Chiffres clés annuels, en FCFA (montants) et FCFA par action (BNPA, dividende)."""
    revenue: dict[int, float] = field(default_factory=dict)
    net_income: dict[int, float] = field(default_factory=dict)
    equity: dict[int, float] = field(default_factory=dict)
    total_assets: dict[int, float] = field(default_factory=dict)
    eps: dict[int, float] = field(default_factory=dict)
    dps: dict[int, float] = field(default_factory=dict)
    per: dict[int, float] = field(default_factory=dict)
    shares: float | None = None
    float_pct: float | None = None
    shareholders: list[tuple[str, float]] = field(default_factory=list)
    source: str = ""

    @property
    def years(self) -> list[int]:
        ys = set(self.revenue) | set(self.net_income) | set(self.equity) | set(self.total_assets)
        return sorted(ys)

    def has_data(self) -> bool:
        return bool(self.revenue or self.net_income)


MANUAL_COLUMNS = {
    "ca": "Chiffre d'affaires ou PNB (millions FCFA)",
    "rn": "Résultat net part du groupe (millions FCFA)",
    "cp": "Capitaux propres (millions FCFA)",
    "ta": "Total du bilan (millions FCFA)",
    "bpa": "BNPA (FCFA par action)",
    "dps": "Dividende net par action (FCFA)",
}


def empty_manual_table(last_year: int | None = None, n: int = 5) -> pd.DataFrame:
    last = last_year or (pd.Timestamp.today().year - 1)
    df = pd.DataFrame({"Exercice": list(range(last - n + 1, last + 1))})
    for col in MANUAL_COLUMNS.values():
        df[col] = pd.Series([None] * n, dtype="float64")
    return df


def figures_from_table(df: pd.DataFrame | None) -> KeyFigures | None:
    """Tableau saisi dans l'application (montants en millions de FCFA) -> KeyFigures."""
    if df is None or df.empty:
        return None
    kf = KeyFigures(source="Saisie manuelle")
    cols = {k: v for k, v in MANUAL_COLUMNS.items() if v in df.columns}
    for _, row in df.iterrows():
        y = row.get("Exercice")
        if not is_num(y):
            continue
        y = int(y)
        for key, col in cols.items():
            v = row.get(col)
            if not is_num(v):
                continue
            v = float(v)
            if key in ("ca", "rn", "cp", "ta"):
                v *= 1e6
            target = {"ca": kf.revenue, "rn": kf.net_income, "cp": kf.equity, "ta": kf.total_assets,
                      "bpa": kf.eps, "dps": kf.dps}[key]
            target[y] = v
    return kf if (kf.has_data() or kf.dps) else None


# ---------------------------------------------------------------------------
# Lecture de la fiche société de sikafinance
# ---------------------------------------------------------------------------
_ROWS = {
    "revenue": ("chiffre d affaires", "produit net bancaire", "pnb", "produits d exploitation bancaire"),
    "net_income": ("resultat net",),
    "eps": ("bnpa", "bpa", "benefice net par action"),
    "per": ("per",),
    "dps": ("dividende",),
}


def _row_key(label: str) -> str | None:
    n = norm_label(label)
    if n.startswith("croissance"):
        return None
    for key, starts in _ROWS.items():
        if any(n == s.upper() or n.startswith(s.upper() + " ") for s in starts):
            return key
    return None


def parse_sika_societe(html: str) -> KeyFigures | None:
    """Fiche société sikafinance -> chiffres clés (tableau années en colonnes, nombre de titres, flottant)."""
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text(" ", strip=True)
    n_text = norm_label(text[:4000]) if text else ""
    mult = 1e9 if "MILLIARDS DE FCFA" in n_text else 1e6
    kf = KeyFigures(source="sikafinance.com (fiche société)")
    for tb in soup.find_all("table"):
        rows = [[c.get_text(" ", strip=True) for c in tr.find_all(["td", "th"])] for tr in tb.find_all("tr")]
        rows = [r for r in rows if r]
        head = next((r for r in rows if sum(1 for c in r if re.fullmatch(r"20\d\d", c.strip())) >= 2), None)
        if head is None:
            continue
        idx = {i: int(c) for i, c in enumerate(head) if re.fullmatch(r"20\d\d", c.strip())}
        for r in rows:
            if r is head or not r:
                continue
            key = _row_key(r[0])
            if key is None:
                continue
            for i, y in idx.items():
                if i >= len(r):
                    continue
                v = to_float(r[i])
                if not is_num(v):
                    continue
                if key == "revenue":
                    kf.revenue[y] = v * mult
                elif key == "net_income":
                    kf.net_income[y] = v * mult
                elif key == "eps":
                    kf.eps[y] = v
                elif key == "per":
                    kf.per[y] = v
                elif key == "dps":
                    kf.dps[y] = v
    m = re.search(r"Nombre de titres\s*:?\s*([\d\s  .,]+)", text, re.I)
    if m:
        v = to_float(m.group(1).strip())
        if is_num(v) and v > 1000:
            kf.shares = float(v)
    m = re.search(r"Flottant\s*:?\s*([\d.,]+)\s*%", text, re.I)
    if m:
        v = to_float(m.group(1))
        if is_num(v) and 0 < v <= 100:
            kf.float_pct = v / 100.0
    for tb in soup.find_all("table"):
        tb.decompose()
    free_text = soup.get_text(" ; ", strip=True)
    for name, pct in re.findall(r"([A-ZÉÈÀ][A-ZÉÈÀ0-9 '&.()\-]{2,40}?)\s*:?\s*(\d{1,3}(?:[.,]\d+)?)\s*%", free_text):
        v = to_float(pct)
        if is_num(v) and 0 < v <= 100 and norm_label(name) not in {"FLOTTANT", "CROISSANCE", "PER"}:
            kf.shareholders.append((name.strip(), v / 100.0))
    kf.shareholders = kf.shareholders[:6]
    return kf if (kf.has_data() or kf.dps) else None


def fetch_sika_societe(ticker: str, country: str | None = None, errors: list[str] | None = None) -> KeyFigures | None:
    from .market import TIMEOUT, _cached, _session, sika_symbol

    def run():
        try:
            r = _session().get(SIKA_SOCIETE.format(sym=sika_symbol(ticker, country)), timeout=TIMEOUT)
            r.raise_for_status()
        except Exception as exc:
            if errors is not None:
                errors.append(f"sikafinance fiche société {ticker} : {type(exc).__name__}"[:200])
            return None
        kf = parse_sika_societe(r.text)
        if kf is None and errors is not None:
            errors.append(f"sikafinance fiche société {ticker} : aucun chiffre clé reconnu.")
        return kf
    return _cached(f"sikasoc:{ticker}", run)


# ---------------------------------------------------------------------------
# Construction d'une société à partir de chiffres clés
# ---------------------------------------------------------------------------
def kind_for(ticker: str, sector: str) -> str:
    if sector == config.FIN and ticker not in config.IFRS_TICKERS:
        return "banque"
    return "ifrs" if ticker in config.IFRS_TICKERS else "syscohada"


def build_company(ticker: str, kf: KeyFigures | None, level: str | None = None) -> CompanyData:
    ref = config.TICKERS.get(ticker, {})
    sector = normalise_sector(None, ticker)
    kind = kind_for(ticker, sector)
    fields: dict[str, dict[int, float]] = {}
    notes: list[str] = []
    years: list[int] = []
    if kf is not None:
        rev_key = "pnb" if kind == "banque" else "ca"
        if kf.revenue:
            fields[rev_key] = dict(kf.revenue)
        if kf.net_income:
            fields["rn"] = dict(kf.net_income)
        if kf.equity:
            fields["capitaux_propres"] = dict(kf.equity)
        if kf.total_assets:
            fields["total_actif"] = dict(kf.total_assets)
        if kf.eps:
            fields["bpa_pub"] = dict(kf.eps)
        if kf.dps:
            fields["dps"] = dict(kf.dps)
        if kf.shares and kf.years:
            fields["nb_actions"] = {max(kf.years): kf.shares}
        elif kf.shares:
            fields["nb_actions"] = {pd.Timestamp.today().year - 1: kf.shares}
        years = kf.years
    if level is None:
        level = "chiffres_cles" if years else "marche"
    if level == "chiffres_cles":
        notes.append("Analyse sur chiffres clés : pas d'états financiers détaillés (bilan, soldes intermédiaires "
                     "de gestion). Les ratios de structure et la solidité financière ne sont pas évalués.")
        if kf is not None and not kf.equity:
            notes.append("Capitaux propres non renseignés : ROE et P/B non calculés.")
    else:
        notes.append("Analyse sur données de marché seules : aucun état financier disponible. Signal indicatif, "
                     "conviction plafonnée.")
    return CompanyData(
        ticker=ticker, name=ref.get("nom", ticker), sheet="", kind=kind, consolidated=False, sector=sector,
        country=ref.get("pays", "CI"), currency="XOF", unit_label="millions FCFA", multiplier=1e6,
        years=years, lines=[], fields=fields, field_labels={}, notes=notes, level=level,
        data_source=(kf.source if kf is not None and kf.source else "Données de marché"),
    )


# ---------------------------------------------------------------------------
# Couverture des données
# ---------------------------------------------------------------------------
@dataclass
class CoverageItem:
    name: str
    status: str            # "ok" | "partiel" | "absent"
    detail: str
    impact: str            # effet de l'absence sur l'analyse
    action: str            # comment compléter


@dataclass
class Coverage:
    level: str
    items: list[CoverageItem]
    score: float           # 0-100

    @property
    def label(self) -> str:
        return config.LEVELS[self.level][0]

    @property
    def missing(self) -> list[CoverageItem]:
        return [i for i in self.items if i.status != "ok"]


def _n_years(c: CompanyData, name: str) -> int:
    return sum(1 for v in c.fields.get(name, {}).values() if is_num(v))


def assess(c: CompanyData, md, fd) -> Coverage:
    """Mesure la couverture des données et fixe le niveau d'analyse."""
    rev = c.revenue_field
    n_rev, n_rn = _n_years(c, rev), _n_years(c, c.net_income_field)
    n_cp, n_ta = _n_years(c, c.equity_field), _n_years(c, "total_actif")
    items: list[CoverageItem] = []

    def add(name, status, detail, impact, action):
        items.append(CoverageItem(name, status, detail, impact, action))

    if c.level in ("complet", "partiel"):
        full = min(n_rev, n_rn, n_cp, n_ta) >= 3
        add("États financiers", "ok" if full else "partiel",
            f"{len(c.years)} exercices ({n_rev} de revenus, {n_rn} de résultat, {n_cp} de fonds propres, "
            f"{n_ta} de total bilan)",
            "Ratios, solidité et croissance calculés sur moins d'exercices.",
            "Compléter la feuille du titre dans le classeur (au moins trois exercices, bilan inclus).")
        if not full:
            c.level = "partiel"
        sig_n = sum(1 for k in ("va", "ebe", "re", "rao", "marge_interets", "rbe") if c.fields.get(k))
        add("Soldes intermédiaires de gestion", "ok" if sig_n >= 3 else ("partiel" if sig_n else "absent"),
            f"{sig_n} soldes reconnus", "Tableau des SIG incomplet.",
            "Ajouter les lignes valeur ajoutée, EBE, résultat d'exploitation (ou marge d'intérêts, RBE).")
    elif c.level == "chiffres_cles":
        add("Chiffres clés (CA, résultat net)", "ok" if min(n_rev, n_rn) >= 3 else "partiel",
            f"{n_rev} exercices de revenus, {n_rn} de résultat net ({c.data_source})",
            "Croissance et PER sur une période plus courte.",
            "Charger l'onglet du titre dans le classeur ou compléter le tableau de saisie.")
        add("Bilan (capitaux propres, total actif)", "ok" if min(n_cp, n_ta) >= 2 else "absent",
            f"{n_cp} exercices de fonds propres, {n_ta} de total bilan",
            "ROE, P/B, P/B justifié, autonomie financière et solidité non évalués.",
            "Saisir capitaux propres et total du bilan (états financiers publiés sur brvm.org).")
        add("Soldes intermédiaires de gestion", "absent", "Non disponibles",
            "Pas de tableau des SIG ni de marge d'exploitation.", "Charger les états financiers détaillés.")
    else:
        add("États financiers", "absent", "Aucun chiffre financier disponible",
            "Qualité et croissance non évaluées ; vue fondamentale neutre par construction.",
            "Saisir les chiffres clés dans le panneau latéral ou charger le classeur.")

    # Marché -----------------------------------------------------------------
    h = md.history
    n_h = 0 if h is None else len(h)
    add("Historique de cours", "ok" if n_h >= 250 else ("partiel" if n_h >= 60 else "absent"),
        f"{n_h} séances" + (f", du {h.index.min():%d/%m/%Y} au {h.index.max():%d/%m/%Y}" if n_h else ""),
        "Analyse technique, bêta et liquidité neutralisés ou moins fiables." if n_h < 250 else "",
        "Activer les données en ligne, charger la base de marché ou un fichier d'historique.")
    add("Cours actuel", "ok" if is_num(md.price) else "absent",
        f"{md.sources.get('cours', 'n.d.')}", "Aucune valorisation ni potentiel sans cours.",
        "Saisir le cours dans le panneau latéral.")
    add("Nombre de titres", "ok" if is_num(md.shares) else "absent", md.sources.get("nombre de titres", "n.d."),
        "Capitalisation, BPA, PER et P/B non calculables.", "Saisir le nombre de titres (brvm.org, fiche société).")
    nd = len([d for d in md.dividends if is_num(d.montant_net) and d.montant_net > 0])
    add("Dividendes", "ok" if nd >= 3 else ("partiel" if nd else "absent"), f"{nd} exercice(s)",
        "Rendement et actualisation des dividendes moins fiables." if nd < 3 else "",
        "Charger l'onglet Dividendes de la base de marché ou saisir le dernier dividende net.")
    n_peers = (fd.medians or {}).get("n", 0) if fd is not None else 0
    add("Sociétés comparables", "ok" if n_peers >= 3 else ("partiel" if n_peers >= 2 else "absent"),
        f"{n_peers} comparable(s) dans le classeur",
        "Multiples sectoriels remplacés par des valeurs de référence par défaut." if n_peers < 2 else "",
        "Charger le classeur complet des états financiers pour disposer de pairs.")
    ok_idx = md.index_history is not None and len(md.index_history) >= 100
    add("Indice BRVM Composite", "ok" if ok_idx else "absent",
        md.sources.get("indice", "n.d."), "Bêta par défaut, force relative non calculée." if not ok_idx else "",
        "Activer les données en ligne ou ajouter la colonne BRVMC dans la feuille Cours.")
    weights = {"ok": 1.0, "partiel": 0.5, "absent": 0.0}
    core = {"États financiers": 3, "Chiffres clés (CA, résultat net)": 2, "Bilan (capitaux propres, total actif)": 1,
            "Soldes intermédiaires de gestion": 1, "Historique de cours": 2, "Cours actuel": 2,
            "Nombre de titres": 1.5, "Dividendes": 1, "Sociétés comparables": 1, "Indice BRVM Composite": 0.5}
    num = sum(weights[i.status] * core.get(i.name, 1) for i in items)
    den = sum(core.get(i.name, 1) for i in items)
    return Coverage(level=c.level, items=items, score=100 * num / den if den else 0.0)
