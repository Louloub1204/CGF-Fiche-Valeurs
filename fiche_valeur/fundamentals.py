"""Analyse fondamentale : soldes intermédiaires de gestion, ratios, solidité,
croissance, positionnement sectoriel."""
from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

from . import config
from .parser import CompanyData, Workbook
from .utils import cagr, clamp, growth, is_num, mean_of, median_of, safe_div, scale

# ---------------------------------------------------------------------------
# Soldes intermédiaires de gestion
# ---------------------------------------------------------------------------
SIG_BANK = [
    ("marge_interets", "Marge d'intérêts"),
    ("commissions_nettes", "Commissions nettes"),
    ("pnb", "Produit net bancaire"),
    ("charges_generales", "Charges générales d'exploitation"),
    ("dotations", "Dotations aux amortissements"),
    ("rbe", "Résultat brut d'exploitation"),
    ("cout_risque", "Coût du risque"),
    ("re", "Résultat d'exploitation"),
    ("impots", "Impôts sur les bénéfices"),
    ("rn", "Résultat net"),
    ("rn_pg", "Résultat net part du groupe"),
]
SIG_SYSCOHADA = [
    ("ca", "Chiffre d'affaires"),
    ("marge_commerciale", "Marge commerciale"),
    ("va", "Valeur ajoutée"),
    ("charges_personnel", "Charges de personnel"),
    ("ebe", "Excédent brut d'exploitation"),
    ("dotations", "Dotations aux amortissements et provisions"),
    ("re", "Résultat d'exploitation"),
    ("rf", "Résultat financier"),
    ("rao", "Résultat des activités ordinaires"),
    ("rhao", "Résultat hors activités ordinaires"),
    ("impots", "Impôts sur le résultat"),
    ("rn", "Résultat net"),
]
SIG_IFRS = [
    ("ca", "Chiffre d'affaires"),
    ("ebe", "EBITDA / EBE"),
    ("dotations", "Dotations aux amortissements"),
    ("re", "Résultat opérationnel"),
    ("rf", "Résultat financier"),
    ("frais_financiers", "Frais financiers"),
    ("rai", "Résultat avant impôt"),
    ("impots", "Impôts sur le résultat"),
    ("rn", "Résultat net consolidé"),
    ("rn_pg", "Résultat net part du groupe"),
]


def sig_rows(c: CompanyData) -> list[tuple[str, str, dict[int, float]]]:
    spec = SIG_BANK if c.is_bank else (SIG_IFRS if c.kind == "ifrs" else SIG_SYSCOHADA)
    out = []
    for key, label in spec:
        series = c.fields.get(key)
        if series and any(is_num(v) and v != 0 for v in series.values()):
            lab = label + (" (estimé)" if key in c.estimated and key not in ("marge_interets", "commissions_nettes", "rai") else "")
            out.append((key, lab, series))
    return out


def sig_table(c: CompanyData, years: list[int] | None = None) -> pd.DataFrame:
    years = years or c.years[-5:]
    rows = sig_rows(c)
    data = {label: [s.get(y) for y in years] for _, label, s in rows}
    return pd.DataFrame(data, index=years).T


# ---------------------------------------------------------------------------
# Ratios
# ---------------------------------------------------------------------------
RATIO_META = {
    # clé: (libellé, format, sens)  format: pct | x | jours ; sens: up = plus c'est haut mieux c'est
    "croissance_ca": ("Croissance du chiffre d'affaires", "pct", "up"),
    "croissance_pnb": ("Croissance du PNB", "pct", "up"),
    "croissance_rn": ("Croissance du résultat net", "pct", "up"),
    "marge_nette": ("Marge nette", "pct", "up"),
    "marge_ebe": ("Marge d'EBE", "pct", "up"),
    "marge_op": ("Marge opérationnelle", "pct", "up"),
    "taux_va": ("Taux de valeur ajoutée", "pct", "up"),
    "poids_personnel": ("Charges de personnel / VA", "pct", "down"),
    "roe": ("ROE (rentabilité des fonds propres)", "pct", "up"),
    "roa": ("ROA (rentabilité des actifs)", "pct", "up"),
    "autonomie": ("Fonds propres / total bilan", "pct", "up"),
    "gearing": ("Gearing (dette nette / fonds propres)", "pct", "down"),
    "dette_nette_ebe": ("Dette nette / EBE", "x", "down"),
    "couverture_ff": ("Couverture des frais financiers", "x", "up"),
    "liquidite_generale": ("Liquidité générale", "x", "up"),
    "bfr_jours": ("BFR d'exploitation (jours de CA)", "jours", "down"),
    "rotation_actif": ("Rotation de l'actif (CA / total bilan)", "x", "up"),
    "taux_impot": ("Taux d'imposition apparent", "pct", "down"),
    "coef_exploitation": ("Coefficient d'exploitation", "pct", "down"),
    "cout_risque_pnb": ("Coût du risque / PNB", "pct", "down"),
    "cout_risque_credits": ("Coût du risque / encours de crédits", "pct", "down"),
    "ltd": ("Crédits / dépôts (LTD)", "pct", "neutral"),
    "croissance_credits": ("Croissance des crédits", "pct", "up"),
    "croissance_depots": ("Croissance des dépôts", "pct", "up"),
    "poids_marge_interets": ("Marge d'intérêts / PNB", "pct", "neutral"),
    "rendement_actifs": ("PNB / total bilan", "pct", "up"),
    "npl": ("Créances en souffrance (publié)", "pct", "down"),
    "solvabilite": ("Ratio de solvabilité (publié)", "pct", "up"),
}


def _avg(series: dict[int, float], y: int) -> float | None:
    v = series.get(y)
    p = series.get(y - 1)
    if is_num(v) and is_num(p):
        return (v + p) / 2
    return v if is_num(v) else None


def compute_ratios(c: CompanyData) -> dict[str, dict[int, float]]:
    f = c.fields
    rev = f.get(c.revenue_field, {})
    rn = f.get(c.net_income_field, {})
    rn_tot = f.get("rn", {})
    cp = f.get(c.equity_field, {})
    ta = f.get("total_actif", {})
    out: dict[str, dict[int, float]] = {}

    def put(key, y, val):
        if is_num(val):
            out.setdefault(key, {})[y] = float(val)

    for y in c.years:
        put("croissance_pnb" if c.is_bank else "croissance_ca", y, growth(rev.get(y), rev.get(y - 1)))
        put("croissance_rn", y, growth(rn.get(y), rn.get(y - 1)))
        put("marge_nette", y, safe_div(rn.get(y), rev.get(y)))
        put("roe", y, safe_div(rn.get(y), _avg(cp, y)) if (_avg(cp, y) or 0) > 0 else None)
        put("roa", y, safe_div(rn_tot.get(y, rn.get(y)), _avg(ta, y)))
        put("autonomie", y, safe_div(f.get("capitaux_propres", {}).get(y), ta.get(y)))
        if c.is_bank:
            cge = f.get("charges_generales", {}).get(y)
            da = f.get("dotations", {}).get(y, 0.0)
            if is_num(cge):
                put("coef_exploitation", y, safe_div(-(cge + (da or 0.0)), rev.get(y)))
            cdr = f.get("cout_risque", {}).get(y)
            if is_num(cdr):
                put("cout_risque_pnb", y, safe_div(-cdr, rev.get(y)))
                put("cout_risque_credits", y, safe_div(-cdr, _avg(f.get("credits_clientele", {}), y)))
            put("ltd", y, safe_div(f.get("credits_clientele", {}).get(y), f.get("depots_clientele", {}).get(y)))
            put("croissance_credits", y, growth(f.get("credits_clientele", {}).get(y), f.get("credits_clientele", {}).get(y - 1)))
            put("croissance_depots", y, growth(f.get("depots_clientele", {}).get(y), f.get("depots_clientele", {}).get(y - 1)))
            put("poids_marge_interets", y, safe_div(f.get("marge_interets", {}).get(y), rev.get(y)))
            put("rendement_actifs", y, safe_div(rev.get(y), _avg(ta, y)))
            if "npl_pub" in f:
                put("npl", y, f["npl_pub"].get(y))
            if "solvabilite_pub" in f:
                put("solvabilite", y, f["solvabilite_pub"].get(y))
            if "coef_exploitation" not in out or y not in out.get("coef_exploitation", {}):
                if "coef_expl_pub" in f:
                    put("coef_exploitation", y, f["coef_expl_pub"].get(y))
        else:
            ebe = f.get("ebe", {}).get(y)
            re_ = f.get("re", {}).get(y)
            put("marge_ebe", y, safe_div(ebe, rev.get(y)))
            put("marge_op", y, safe_div(re_, rev.get(y)))
            put("taux_va", y, safe_div(f.get("va", {}).get(y), rev.get(y)))
            pers = f.get("charges_personnel", {}).get(y)
            put("poids_personnel", y, safe_div(abs(pers) if is_num(pers) else None, f.get("va", {}).get(y)))
            nd = net_debt(c, y)
            put("gearing", y, safe_div(nd, f.get("capitaux_propres", {}).get(y)) if (f.get("capitaux_propres", {}).get(y) or 0) > 0 else None)
            if is_num(nd) and is_num(ebe) and ebe > 0:
                put("dette_nette_ebe", y, nd / ebe)
            ff = f.get("frais_financiers", {}).get(y)
            if is_num(ff) and ff != 0 and is_num(re_):
                put("couverture_ff", y, re_ / abs(ff))
            put("liquidite_generale", y, safe_div(f.get("actif_circulant", {}).get(y), f.get("passif_circulant", {}).get(y)))
            st, cl, fo = (f.get(k, {}).get(y) for k in ("stocks", "creances_clients", "fournisseurs"))
            if is_num(st) or is_num(cl):
                bfr = (st or 0) + (cl or 0) - abs(fo or 0)
                put("bfr_jours", y, safe_div(bfr * 360, rev.get(y)))
            put("rotation_actif", y, safe_div(rev.get(y), _avg(ta, y)))
        imp = f.get("impots", {}).get(y)
        rai = f.get("rai", {}).get(y)
        if not is_num(rai) and is_num(rn_tot.get(y)) and is_num(imp):
            rai = rn_tot[y] - imp
        if is_num(imp) and is_num(rai) and rai > 0:
            put("taux_impot", y, abs(imp) / rai)
    if "roe_pub" in f:
        for y, v in f["roe_pub"].items():
            if y not in out.get("roe", {}):
                put("roe", y, v)
    return out


def net_debt(c: CompanyData, y: int) -> float | None:
    f = c.fields
    df = f.get("dettes_financieres", {}).get(y)
    tp = f.get("tresorerie_passif", {}).get(y)
    ta_ = f.get("tresorerie_actif", {}).get(y)
    if not any(is_num(v) for v in (df, tp, ta_)):
        return None
    return abs(df or 0) + abs(tp or 0) - abs(ta_ or 0)


def caf(c: CompanyData, y: int) -> float | None:
    rn = c.fields.get("rn", {}).get(y)
    da = c.fields.get("dotations", {}).get(y)
    if not is_num(rn):
        return None
    return rn + abs(da or 0)


# ---------------------------------------------------------------------------
# Scores de solidité
# ---------------------------------------------------------------------------

@dataclass
class Solidity:
    score: int
    maximum: int
    tests: list[tuple[str, bool | None]]
    altman: float | None = None
    altman_zone: str | None = None


def solidity(c: CompanyData, r: dict[str, dict[int, float]]) -> Solidity:
    """Score de solidité inspiré de Piotroski (9 tests), adapté aux banques."""
    y = c.last_year(c.net_income_field) or c.last_year()
    p = y - 1 if y else None
    f = c.fields
    tests: list[tuple[str, bool | None]] = []

    def cmp(a, b, better="up"):
        if not is_num(a) or not is_num(b):
            return None
        return a > b if better == "up" else a < b

    rn = f.get(c.net_income_field, {})
    tests.append(("Résultat net positif", is_num(rn.get(y)) and rn.get(y) > 0 if is_num(rn.get(y)) else None))
    tests.append(("ROA en amélioration", cmp(r.get("roa", {}).get(y), r.get("roa", {}).get(p))))
    tests.append(("ROE en amélioration", cmp(r.get("roe", {}).get(y), r.get("roe", {}).get(p))))
    if c.is_bank:
        tests.append(("Coefficient d'exploitation en baisse", cmp(r.get("coef_exploitation", {}).get(y), r.get("coef_exploitation", {}).get(p), "down")))
        tests.append(("Coût du risque / PNB en baisse", cmp(r.get("cout_risque_pnb", {}).get(y), r.get("cout_risque_pnb", {}).get(p), "down")))
        tests.append(("PNB en croissance", cmp(f.get("pnb", {}).get(y), f.get("pnb", {}).get(p))))
        ltd = r.get("ltd", {}).get(y)
        tests.append(("Transformation maîtrisée (LTD entre 55% et 95%)", (0.55 <= ltd <= 0.95) if is_num(ltd) else None))
        tests.append(("Fonds propres / bilan en hausse", cmp(r.get("autonomie", {}).get(y), r.get("autonomie", {}).get(p))))
        tests.append(("Marge nette en hausse", cmp(r.get("marge_nette", {}).get(y), r.get("marge_nette", {}).get(p))))
    else:
        cf = caf(c, y)
        tests.append(("Capacité d'autofinancement positive", (cf > 0) if is_num(cf) else None))
        tests.append(("CAF supérieure au résultat net (qualité des résultats)",
                      (cf > rn.get(y)) if is_num(cf) and is_num(rn.get(y)) else None))
        g_now, g_prev = r.get("gearing", {}).get(y), r.get("gearing", {}).get(p)
        tests.append(("Endettement net en baisse", cmp(g_now, g_prev, "down")))
        tests.append(("Liquidité générale en hausse", cmp(r.get("liquidite_generale", {}).get(y), r.get("liquidite_generale", {}).get(p))))
        tests.append(("Marge d'EBE en hausse", cmp(r.get("marge_ebe", {}).get(y), r.get("marge_ebe", {}).get(p))))
        tests.append(("Rotation de l'actif en hausse", cmp(r.get("rotation_actif", {}).get(y), r.get("rotation_actif", {}).get(p))))
    score = sum(1 for _, t in tests if t is True)
    maximum = sum(1 for _, t in tests if t is not None)
    sol = Solidity(score=score, maximum=maximum, tests=tests)
    if not c.is_bank:
        sol.altman, sol.altman_zone = altman_em(c, y)
    return sol


def altman_em(c: CompanyData, y: int) -> tuple[float | None, str | None]:
    """Score Z'' d'Altman pour marchés émergents (sociétés non financières)."""
    f = c.fields
    ta = f.get("total_actif", {}).get(y)
    ac, pc = f.get("actif_circulant", {}).get(y), f.get("passif_circulant", {}).get(y)
    cp = f.get("capitaux_propres", {}).get(y)
    re_ = f.get("re", {}).get(y)
    if not all(is_num(v) for v in (ta, cp, re_)) or not ta:
        return None, None
    reserves = (f.get("reserves", {}).get(y) or 0) + (f.get("report_a_nouveau", {}).get(y) or 0)
    if reserves == 0 and is_num(f.get("capital", {}).get(y)):
        reserves = cp - f["capital"][y]
    dettes = ta - cp
    x1 = ((ac - pc) / ta) if is_num(ac) and is_num(pc) else 0.0
    x2 = reserves / ta
    x3 = re_ / ta
    x4 = cp / dettes if dettes > 0 else 3.0
    z = 3.25 + 6.56 * x1 + 3.26 * x2 + 6.72 * x3 + 1.05 * x4
    zone = "sécurité" if z > 5.85 else ("vigilance" if z >= 4.35 else "détresse")
    return z, zone


# ---------------------------------------------------------------------------
# Croissance
# ---------------------------------------------------------------------------

def growth_stats(c: CompanyData) -> dict[str, float | None]:
    rev = c.fields.get(c.revenue_field, {})
    rn = c.fields.get(c.net_income_field, {})
    out: dict[str, float | None] = {}
    for key, s in (("rev", rev), ("rn", rn)):
        ys = sorted(y for y, v in s.items() if is_num(v))
        if len(ys) >= 2:
            first = ys[0] if len(ys) <= 5 else ys[-5]
            out[f"tcam_{key}"] = cagr(s[ys[-1]], s[first], ys[-1] - first)
            out[f"tcam_{key}_periode"] = f"{first}-{ys[-1]}"
            out[f"last_{key}"] = growth(s[ys[-1]], s[ys[-2]])
        else:
            out[f"tcam_{key}"] = None
            out[f"last_{key}"] = None
    return out


# ---------------------------------------------------------------------------
# Pairs et positionnement sectoriel
# ---------------------------------------------------------------------------

@dataclass
class PeerRow:
    ticker: str
    name: str
    sector: str
    kind: str
    year: int | None
    revenue: float | None
    net_income: float | None
    equity: float | None
    roe: float | None
    net_margin: float | None
    growth_rev: float | None
    ebitda: float | None
    net_debt: float | None
    price: float | None = None
    shares: float | None = None
    dps: float | None = None

    @property
    def market_cap(self):
        return self.price * self.shares if is_num(self.price) and is_num(self.shares) else None

    @property
    def pe(self):
        return safe_div(self.market_cap, self.net_income) if is_num(self.net_income) and self.net_income > 0 else None

    @property
    def pb(self):
        return safe_div(self.market_cap, self.equity) if is_num(self.equity) and self.equity > 0 else None

    @property
    def ev_ebitda(self):
        if not is_num(self.market_cap) or not is_num(self.ebitda) or self.ebitda <= 0:
            return None
        return (self.market_cap + (self.net_debt or 0)) / self.ebitda

    @property
    def div_yield(self):
        return safe_div(self.dps, self.price) if is_num(self.dps) and self.dps > 0 else None


def peer_rows(wb: Workbook, prices: dict[str, float] | None = None,
              shares: dict[str, float] | None = None) -> list[PeerRow]:
    prices = prices or {}
    shares = shares or {}
    rows = []
    for t, c in wb.companies.items():
        y = c.last_year(c.net_income_field)
        r = compute_ratios(c)
        dps_series = c.fields.get("dps", {})
        rows.append(PeerRow(
            ticker=t, name=c.name, sector=c.sector, kind=c.kind, year=y,
            revenue=c.fields.get(c.revenue_field, {}).get(y),
            net_income=c.fields.get(c.net_income_field, {}).get(y),
            equity=c.fields.get(c.equity_field, {}).get(y),
            roe=r.get("roe", {}).get(y), net_margin=r.get("marge_nette", {}).get(y),
            growth_rev=r.get("croissance_pnb" if c.is_bank else "croissance_ca", {}).get(y),
            ebitda=None if c.is_bank else c.fields.get("ebe", {}).get(y),
            net_debt=None if c.is_bank else net_debt(c, y),
            price=prices.get(t), shares=shares.get(t),
            dps=dps_series.get(max(dps_series)) if dps_series else None,
        ))
    return rows


def peer_group(rows: list[PeerRow], target: CompanyData, min_size: int = 3) -> tuple[list[PeerRow], str]:
    same = [r for r in rows if r.sector == target.sector and r.ticker != target.ticker]
    if len(same) >= min_size - 1:
        return same, f"secteur {target.sector}"
    family = "banques" if target.is_bank else "sociétés non financières"
    fam = [r for r in rows if (r.kind == "banque") == target.is_bank and r.ticker != target.ticker]
    return fam, f"{family} du classeur"


def peer_medians(rows: list[PeerRow]) -> dict[str, float | None]:
    def med(attr, lo=None, hi=None):
        vals = [getattr(r, attr) for r in rows]
        vals = [v for v in vals if is_num(v) and (lo is None or v >= lo) and (hi is None or v <= hi)]
        return median_of(vals) if len(vals) >= 2 else None
    return {
        "pe": med("pe", 0.5, 60), "pb": med("pb", 0.05, 15), "ev_ebitda": med("ev_ebitda", 0.5, 40),
        "roe": med("roe", -1, 1), "net_margin": med("net_margin", -1, 1), "growth_rev": med("growth_rev", -0.5, 1),
        "div_yield": med("div_yield", 0, 0.3), "n": len(rows),
    }


def percentile(value, values) -> float | None:
    vals = [v for v in values if is_num(v)]
    if not is_num(value) or len(vals) < 2:
        return None
    below = sum(1 for v in vals if v < value)
    equal = sum(1 for v in vals if v == value)
    return (below + 0.5 * equal) / len(vals)


# ---------------------------------------------------------------------------
# Scores (0-100)
# ---------------------------------------------------------------------------

def _weighted(parts: list[tuple[float | None, float]]) -> float | None:
    num = sum(s * w for s, w in parts if is_num(s))
    den = sum(w for s, w in parts if is_num(s))
    return num / den if den else None


def quality_score(c: CompanyData, r: dict, sol: Solidity, peers: list[PeerRow]) -> tuple[float | None, list]:
    y = c.last_year(c.net_income_field)
    g = lambda k: r.get(k, {}).get(y)  # noqa: E731
    parts: list[tuple[str, float | None, float]] = []
    if c.is_bank:
        ltd = g("ltd")
        ltd_s = None
        if is_num(ltd):
            ltd_s = 100.0 if 0.65 <= ltd <= 0.90 else clamp(100 - abs(ltd - (0.65 if ltd < 0.65 else 0.90)) * 300)
        parts += [
            ("ROE", scale(g("roe"), 0.04, 0.24), 25), ("ROA", scale(g("roa"), 0.004, 0.025), 10),
            ("Coefficient d'exploitation", scale(g("coef_exploitation"), 0.40, 0.75, reverse=True), 20),
            ("Coût du risque / PNB", scale(g("cout_risque_pnb"), 0.05, 0.35, reverse=True), 15),
            ("Fonds propres / bilan", scale(g("autonomie"), 0.05, 0.13), 10),
            ("Transformation (LTD)", ltd_s, 5),
            ("Créances en souffrance", scale(g("npl"), 0.03, 0.15, reverse=True), 5),
        ]
    else:
        gear = g("gearing")
        gear_s = 100.0 if is_num(gear) and gear <= 0 else scale(gear, 0.0, 1.5, reverse=True)
        az = {"sécurité": 100.0, "vigilance": 55.0, "détresse": 15.0}.get(sol.altman_zone) if sol.altman_zone else None
        parts += [
            ("ROE", scale(g("roe"), 0.0, 0.25), 25), ("Marge nette", scale(g("marge_nette"), 0.0, 0.15), 15),
            ("Marge d'EBE", scale(g("marge_ebe"), 0.03, 0.30), 15), ("Endettement", gear_s, 15),
            ("Altman Z''", az, 5),
        ]
    if sol.maximum:
        parts.append(("Score de solidité", sol.score / sol.maximum * 100, 15))
    pct = percentile(g("roe"), [p.roe for p in peers] + [g("roe")])
    parts.append(("ROE relatif aux pairs", pct * 100 if is_num(pct) else None, 10))
    score = _weighted([(s, w) for _, s, w in parts])
    return score, parts


def growth_score(c: CompanyData, r: dict, gs: dict) -> tuple[float | None, list]:
    y = c.last_year(c.net_income_field)
    mn = r.get("marge_nette", {})
    ys = sorted(mn)
    trend = None
    if len(ys) >= 3 and y in mn:
        trend = mn[y] - mean_of([mn[k] for k in ys[-4:-1]])
    parts = [
        ("TCAM des revenus", scale(gs.get("tcam_rev"), -0.05, 0.15), 30),
        ("TCAM du résultat net", scale(gs.get("tcam_rn"), -0.10, 0.20), 25),
        ("Croissance du résultat net (dernier exercice)", scale(gs.get("last_rn"), -0.20, 0.30), 25),
        ("Tendance de la marge nette", scale(trend, -0.03, 0.03), 20),
    ]
    if c.is_bank:
        parts.append(("Croissance des crédits", scale(r.get("croissance_credits", {}).get(y), -0.02, 0.15), 10))
    return _weighted([(s, w) for _, s, w in parts]), parts


@dataclass
class Fundamentals:
    ratios: dict[str, dict[int, float]]
    solidity: Solidity
    growth: dict
    peers: list[PeerRow]
    peer_label: str
    medians: dict
    quality_score: float | None
    quality_parts: list
    growth_score: float | None
    growth_parts: list
    year: int | None
    notes: list[str] = field(default_factory=list)
    default_multiples: list[str] = field(default_factory=list)   # multiples issus des références par défaut


def analyse_fundamentals(c: CompanyData, wb: Workbook | None, prices=None, shares=None,
                         use_defaults: bool = True) -> Fundamentals:
    r = compute_ratios(c)
    sol = solidity(c, r)
    gs = growth_stats(c)
    rows = peer_rows(wb, prices, shares) if wb is not None and wb.companies else []
    peers, label = peer_group(rows, c) if rows else ([], "aucun comparable dans le classeur")
    med = peer_medians(peers)
    defaults: list[str] = []
    if use_defaults:
        ref = config.SECTOR_MULTIPLES.get(c.sector, config.SECTOR_MULTIPLES["Non classé"])
        for key in ("pe", "pb", "ev_ebitda", "div_yield"):
            if not is_num(med.get(key)) and is_num(ref.get(key)):
                med[key] = ref[key]
                defaults.append(key)
        if defaults:
            label = f"{label} ; multiples de référence du secteur {c.sector} pour " + ", ".join(
                {"pe": "le PER", "pb": "le P/B", "ev_ebitda": "le VE/EBE", "div_yield": "le rendement"}[k]
                for k in defaults)
    qs, qparts = quality_score(c, r, sol, peers)
    gsc, gparts = growth_score(c, r, gs)
    return Fundamentals(ratios=r, solidity=sol, growth=gs, peers=peers, peer_label=label, medians=med,
                        quality_score=qs, quality_parts=qparts, growth_score=gsc, growth_parts=gparts,
                        year=c.last_year(c.net_income_field), default_multiples=defaults)
