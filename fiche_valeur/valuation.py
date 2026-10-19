"""Valorisation : multiples, valeur intrinsèque par plusieurs méthodes, potentiel."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date

from .config import Settings
from .fundamentals import Fundamentals, net_debt
from .market import MarketData
from .parser import CompanyData
from .technical import Technical
from .utils import clamp, fmt_num, fmt_pct, is_num, mean_of, safe_div, scale


@dataclass
class Method:
    name: str
    value: float | None
    weight: float
    detail: str


@dataclass
class Valuation:
    price: float | None
    shares: float | None
    market_cap: float | None
    eps: float | None = None
    eps_ltm: float | None = None
    pe_ltm: float | None = None
    ltm_label: str = ""
    bvps: float | None = None
    dps: float | None = None
    dps_exercice: int | None = None
    pe: float | None = None
    pb: float | None = None
    ps: float | None = None
    ev_ebitda: float | None = None
    div_yield: float | None = None
    payout: float | None = None
    cost_of_equity: float | None = None
    beta: float | None = None
    beta_source: str = ""
    g: float | None = None
    roe_norm: float | None = None
    pb_justified: float | None = None
    implied_roe: float | None = None
    implied_div_growth: float | None = None
    sustainable_growth: float | None = None
    methods: list[Method] = field(default_factory=list)
    fair_value: float | None = None
    fv_low: float | None = None
    fv_high: float | None = None
    upside: float | None = None
    total_return: float | None = None
    score: float | None = None
    parts: list = field(default_factory=list)
    notes: list[str] = field(default_factory=list)


def _dividend_growth(md: MarketData) -> float | None:
    """Taux de croissance annuel moyen du dividende net sur l'historique disponible."""
    byy = {}
    for d in md.dividends:
        if is_num(d.montant_net) and d.montant_net > 0 and d.exercice:
            byy[d.exercice] = d.montant_net
    ys = sorted(byy)
    if len(ys) < 3 or ys[-1] == ys[0]:
        return None
    return (byy[ys[-1]] / byy[ys[0]]) ** (1 / (ys[-1] - ys[0])) - 1


def analyse_valuation(c: CompanyData, fd: Fundamentals, md: MarketData, tech: Technical,
                      st: Settings, dyn=None) -> Valuation:
    y = fd.year
    price, shares = md.price, md.shares
    v = Valuation(price=price, shares=shares, market_cap=md.market_cap)
    rn = c.fields.get(c.net_income_field, {}).get(y)
    cp = c.fields.get(c.equity_field, {}).get(y)
    rev = c.fields.get(c.revenue_field, {}).get(y)
    ebitda = None if c.is_bank else c.fields.get("ebe", {}).get(y)
    nd = None if c.is_bank else net_debt(c, y)

    if not is_num(shares) or shares <= 0:
        v.notes.append("Nombre de titres inconnu : les données par action et les multiples ne peuvent être "
                       "calculés (renseigner le nombre de titres dans le panneau latéral).")
    v.eps = safe_div(rn, shares)
    ltm_rn = dyn.ltm.get("rn") if dyn is not None else None
    if is_num(ltm_rn) and is_num(shares) and shares > 0:
        v.eps_ltm = ltm_rn / shares
        v.ltm_label = f"12 mois à {dyn.label}"
    v.bvps = safe_div(cp, shares)
    last_div = md.last_dividend()
    if last_div:
        v.dps, v.dps_exercice = last_div.montant_net, last_div.exercice
        if "exceptionnel" in (last_div.source or ""):
            v.notes.append("Dernier dividende jugé exceptionnel (plus de 2,5 fois les versements précédents) : "
                           "le dividende ordinaire médian est retenu pour le rendement et l'actualisation.")
    elif md.dividends and any(is_num(d_.montant_net) for d_ in md.dividends):
        # dividendes renseignés mais aucun versement au dernier exercice : rendement nul, pas de repli sur un exercice ancien
        latest = max((d_ for d_ in md.dividends if is_num(d_.montant_net)),
                     key=lambda d_: (d_.exercice or 0, d_.date_paiement or date(1900, 1, 1)))
        v.dps, v.dps_exercice = 0.0, latest.exercice
        v.notes.append(f"Aucun dividende recensé au titre de l'exercice {latest.exercice} : rendement nul, "
                       "actualisation des dividendes écartée. Saisir le dividende dans le panneau latéral s'il a été versé.")
    elif c.fields.get("dps"):
        yy = max(c.fields["dps"])
        v.dps, v.dps_exercice = c.fields["dps"][yy], yy

    mcap = v.market_cap
    if is_num(mcap):
        v.pe = safe_div(mcap, rn) if is_num(rn) and rn > 0 else None
        if is_num(ltm_rn) and ltm_rn > 0:
            v.pe_ltm = mcap / ltm_rn
        v.pb = safe_div(mcap, cp) if is_num(cp) and cp > 0 else None
        v.ps = safe_div(mcap, rev)
        if is_num(ebitda) and ebitda > 0:
            v.ev_ebitda = (mcap + (nd or 0)) / ebitda
    if is_num(v.dps) and is_num(price) and price > 0:
        v.div_yield = v.dps / price
    if is_num(v.dps) and is_num(v.eps) and v.eps > 0:
        v.payout = v.dps / v.eps

    # Coût des fonds propres ---------------------------------------------------
    beta = tech.indicators.get("beta") if tech and tech.available else None
    if is_num(beta) and tech.indicators.get("correlation", 0) and tech.indicators.get("correlation", 0) > 0.15:
        v.beta = min(1.4, max(0.6, beta))
        v.beta_source = "régression hebdomadaire sur le BRVM Composite (2 ans, ajustement de Blume)"
    else:
        v.beta = st.beta_defaut
        v.beta_source = "valeur par défaut (corrélation à l'indice trop faible ou indisponible)"
    k = (st.taux_sans_risque + v.beta * st.prime_risque_marche) / 100.0
    g = min(st.croissance_long_terme / 100.0, k - 0.02)
    v.cost_of_equity, v.g = k, g

    roe_series = fd.ratios.get("roe", {})
    roes = [roe_series[yy] for yy in sorted(roe_series)[-3:] if is_num(roe_series[yy])]
    v.roe_norm = mean_of(roes)
    if is_num(v.roe_norm) and is_num(v.payout):
        v.sustainable_growth = v.roe_norm * max(0.0, 1 - min(v.payout, 1.0))
    if is_num(v.roe_norm) and v.roe_norm > g:
        v.pb_justified = (v.roe_norm - g) / (k - g)
    if is_num(v.pb):
        v.implied_roe = g + v.pb * (k - g)
    if is_num(v.div_yield):
        v.implied_div_growth = (k - v.div_yield) / (1 + v.div_yield)

    # Méthodes de valorisation ---------------------------------------------------
    med = fd.medians
    bank = c.is_bank
    w = ({"pe": 25, "pb": 20, "pbj": 30, "ddm": 25, "ev": 0} if bank else
         {"pe": 30, "pb": 10, "pbj": 10, "ddm": 25, "ev": 25})
    dflt = set(fd.default_multiples)
    methods: list[Method] = []
    if is_num(v.eps) and v.eps > 0 and is_num(med.get("pe")):
        use_ltm = is_num(v.eps_ltm) and v.eps_ltm > 0
        # la dynamique de l'exercice en cours (12 mois glissants) prend la moitié du poids du PER ; un seul trimestre en pèse 30%
        share_ltm = (0.3 if dyn is not None and dyn.period == "T1" else 0.5) if use_ltm else 0.0
        methods.append(Method("PER sectoriel", med["pe"] * v.eps, w["pe"] * (1 - share_ltm),
                              f"PER {'médian des pairs' if 'pe' not in dflt else 'de référence'} {med['pe']:.1f}x "
                              f"appliqué au BPA {y}".replace(".", ",")))
        if use_ltm:
            methods.append(Method("PER sur 12 mois glissants", med["pe"] * v.eps_ltm, w["pe"] * share_ltm,
                                  f"PER {'médian des pairs' if 'pe' not in dflt else 'de référence'} {med['pe']:.1f}x "
                                  f"appliqué au BPA des {v.ltm_label} ({fmt_num(v.eps_ltm)} FCFA)".replace(".", ",")))
    if is_num(v.bvps) and v.bvps > 0 and is_num(med.get("pb")):
        methods.append(Method("P/B sectoriel", med["pb"] * v.bvps, w["pb"],
                              f"P/B {'médian des pairs' if 'pb' not in dflt else 'de référence'} {med['pb']:.2f}x "
                              f"appliqué aux fonds propres par action".replace(".", ",")))
    if is_num(v.pb_justified) and is_num(v.bvps) and v.bvps > 0:
        methods.append(Method("P/B justifié (Gordon)", v.pb_justified * v.bvps, w["pbj"],
                              f"(ROE normatif {v.roe_norm * 100:.1f}% - g) / (k - g) = {v.pb_justified:.2f}x"
                              .replace(".", ",")))
    if is_num(v.dps) and v.dps > 0:
        g1 = fd.growth.get("tcam_rn")
        if not is_num(g1):
            g1 = _dividend_growth(md)
        g1 = min(0.12, max(0.0, g1)) if is_num(g1) else g
        pv, d = 0.0, v.dps
        for t in range(1, 6):
            d *= 1 + g1
            pv += d / (1 + k) ** t
        tv = d * (1 + g) / (k - g)
        pv += tv / (1 + k) ** 5
        methods.append(Method("Actualisation des dividendes", pv, w["ddm"],
                              f"DPA {fmt_num(v.dps)} FCFA, croissance {fmt_pct(g1)} sur 5 ans puis "
                              f"{fmt_pct(g)}, k = {fmt_pct(k)}"))
    if not bank and is_num(ebitda) and ebitda > 0 and is_num(med.get("ev_ebitda")) and is_num(shares) and shares > 0:
        eq = med["ev_ebitda"] * ebitda - (nd or 0)
        if eq > 0:
            methods.append(Method("VE/EBE sectoriel", eq / shares, w["ev"],
                                  f"VE/EBE {'médian des pairs' if 'ev_ebitda' not in dflt else 'de référence'} "
                                  f"{med['ev_ebitda']:.1f}x, dette nette déduite".replace(".", ",")))
    # Rendement cible : dividende capitalisé au rendement de référence (utile sans états financiers)
    if (c.level in ("chiffres_cles", "marche") and is_num(v.dps) and v.dps > 0
            and is_num(med.get("div_yield")) and med["div_yield"] > 0):
        methods.append(Method("Rendement de référence", v.dps / med["div_yield"], 25 if not is_num(v.eps) else 15,
                              f"DPA {fmt_num(v.dps)} FCFA capitalisé au rendement {'médian des pairs' if 'div_yield' not in dflt else 'de référence du secteur'} "
                              f"de {fmt_pct(med['div_yield'])}"))
    if dflt:
        v.notes.append("Multiples de référence du secteur utilisés faute de comparables suffisants dans le classeur : "
                       "charger le classeur complet pour des multiples issus des pairs.")
    methods = [m for m in methods if is_num(m.value) and m.value > 0 and m.weight > 0]
    if is_num(y) and y < date.today().year - 3:
        kept = [m for m in methods if m.name in ("Actualisation des dividendes", "Rendement de référence")]
        if len(kept) < len(methods):
            v.notes.append(f"Derniers résultats connus datant de {y} : méthodes fondées sur les résultats et les "
                           f"fonds propres écartées.")
        methods = kept
    if is_num(price) and price > 0:
        sane = [m for m in methods if 0.15 * price <= m.value <= 8 * price]
        dropped = [m.name for m in methods if m not in sane]
        if dropped and sane:
            v.notes.append("Méthode(s) écartée(s) pour valeur aberrante au regard du cours : "
                           + ", ".join(dropped).lower() + " (donnée source à vérifier).")
            methods = sane
        elif dropped:
            v.notes.append("Toutes les méthodes donnent des valeurs aberrantes au regard du cours : données "
                           "financières à vérifier, aucune juste valeur retenue.")
            methods = []
    v.methods = methods
    if methods:
        tw = sum(m.weight for m in methods)
        v.fair_value = sum(m.value * m.weight for m in methods) / tw
        vals = sorted(m.value for m in methods)
        v.fv_low, v.fv_high = vals[0], vals[-1]
        if len(methods) == 1:
            v.notes.append("Une seule méthode de valorisation disponible : juste valeur à interpréter avec prudence.")
    else:
        v.notes.append("Aucune méthode de valorisation applicable (cours, nombre de titres, pairs cotés ou "
                       "dividende manquants).")
    if is_num(v.fair_value) and is_num(price) and price > 0:
        v.upside = v.fair_value / price - 1
        v.total_return = v.upside + (v.div_yield or 0)

    # Score de valorisation ------------------------------------------------------
    pe_rel = safe_div(v.pe, med.get("pe")) if is_num(v.pe) else None
    pb_rel = safe_div(v.pb, v.pb_justified) if is_num(v.pb) and is_num(v.pb_justified) else None
    v.parts = [
        ("Potentiel vs juste valeur", scale(v.upside, -0.30, 0.40), 50),
        ("PER relatif aux pairs", scale(pe_rel, 1.6, 0.6) if is_num(pe_rel) else None, 15),
        ("P/B vs P/B justifié", scale(pb_rel, 1.6, 0.6) if is_num(pb_rel) else None, 15),
        ("Rendement du dividende", scale(v.div_yield, 0.02, 0.10), 20),
    ]
    num = sum(s * wt for _, s, wt in v.parts if is_num(s))
    den = sum(wt for _, s, wt in v.parts if is_num(s))
    v.score = clamp(num / den) if den else None
    return v
