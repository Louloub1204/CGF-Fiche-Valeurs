"""Orchestration : du classeur des états financiers à l'analyse complète."""
from __future__ import annotations

import copy
from dataclasses import dataclass, field
from datetime import date

from . import config
from .commentary import build_commentary, claude_enrich
from .config import Settings
from .fundamentals import Fundamentals, analyse_fundamentals
from .macro import Macro, analyse_macro
from .market import MarketBase, MarketData, build_market_data
from .news import News, add_user_notes, fetch_news
from .parser import CompanyData, Workbook
from .partial import (Coverage, KeyFigures, assess, build_company, fetch_sika_societe, figures_from_market_base,
                      fill_company_gaps, merge_figures, smf_peer_rows)
from .publications import InterimReport, Publication, latest_report, momentum
from .scoring import Decision, clamp_pillars, decide, pillar_dict
from .technical import Technical, analyse_technical
from .utils import is_num, scale
from .valuation import Valuation, analyse_valuation


@dataclass
class Analysis:
    company: CompanyData
    market: MarketData
    fundamentals: Fundamentals
    technical: Technical
    valuation: Valuation
    macro: Macro
    news: News
    decision: Decision
    texts: dict = field(default_factory=dict)
    settings: Settings = field(default_factory=Settings)
    as_of: date = field(default_factory=date.today)
    ai_used: bool = False
    errors: list[str] = field(default_factory=list)
    coverage: Coverage | None = None
    key_figures: KeyFigures | None = None
    interim: InterimReport | None = None
    publications: list[Publication] = field(default_factory=list)


def resolve_company(wb: Workbook | None, ticker: str, st: Settings, figures: KeyFigures | None,
                    errors: list[str], market_base: MarketBase | None = None) -> tuple[CompanyData, KeyFigures | None]:
    """Société du classeur si elle y figure (exercices manquants complétés par la base SMF) ; sinon chiffres
    clés (saisis, puis base SMF, puis lecture en ligne) ; sinon données de marché seules."""
    smf = figures_from_market_base(market_base, ticker)
    if wb is not None and ticker in wb.companies and figures is None:
        c = copy.deepcopy(wb.companies[ticker])
        fill_company_gaps(c, smf)
        return c, None
    kf = merge_figures(copy.deepcopy(figures), smf) if figures is not None else smf
    if st.en_ligne and (kf is None or not kf.eps):
        online = fetch_sika_societe(ticker, config.TICKERS.get(ticker, {}).get("pays"), errors)
        kf = merge_figures(kf, online) if kf is not None else online
    return build_company(ticker, kf), kf


def run_analysis(wb: Workbook | None, ticker: str, settings: Settings | None = None,
                 market_base: MarketBase | None = None, price_file=None, overrides: dict | None = None,
                 user_notes: str = "", major_event: bool = False, bceao: dict | None = None,
                 use_claude: bool = False, progress=None, figures: KeyFigures | None = None,
                 custom: dict | None = None, report: InterimReport | None = None) -> Analysis:
    """Analyse d'un titre. ``wb`` peut être vide ou absent : le titre est alors traité au niveau 2 ou 3.

    ``figures`` : chiffres clés saisis (prioritaires sur ceux du classeur) ; ``custom`` : nom, pays et
    secteur d'un titre absent du référentiel.
    """
    st = settings or Settings()
    if custom:
        config.register_ticker(ticker, custom.get("nom", ticker), custom.get("pays", "CI"),
                               custom.get("secteur", "Non classé"))
    pre_errors: list[str] = []
    c, kf = resolve_company(wb, ticker, st, figures, pre_errors, market_base)

    def step(msg):
        if progress:
            progress(msg)

    step("Données de marché")
    md = build_market_data(ticker, c, market_base, price_file, online=st.en_ligne,
                           years=st.annees_historique, overrides=overrides)
    md.errors[:0] = pre_errors
    if kf is not None and kf.float_pct and md.float_shares is None and is_num(md.shares):
        md.float_shares = kf.float_pct * md.shares
        md.sources["flottant"] = kf.source
    interim, pubs = report, []
    if st.en_ligne:
        step("Publications officielles")
        found, pubs = latest_report(ticker, md.errors)
        if interim is None:
            interim = found
    if interim is not None:
        _merge_annual(c, interim)
    step("Analyse fondamentale")
    extra = smf_peer_rows(market_base, wb, md.peer_prices, md.peer_shares)
    fd = analyse_fundamentals(c, wb, md.peer_prices, md.peer_shares, use_defaults=st.multiples_defaut,
                              extra_rows=extra)
    if interim is not None:
        _apply_momentum(fd, interim)
    step("Analyse technique")
    tech = analyse_technical(md.history, md.index_history, md.float_shares)
    step("Valorisation")
    val = analyse_valuation(c, fd, md, tech, st)
    step("Environnement macroéconomique")
    errors: list[str] = list(md.errors)
    mac = analyse_macro(c.country, c.sector, online=st.en_ligne, bceao=bceao, errors=errors)
    step("Actualités")
    news = fetch_news(ticker, c.name, c.country) if st.en_ligne else News(error="Mode hors ligne : actualités non consultées.")
    if news.error and st.en_ligne:
        errors.append(news.error)
    add_user_notes(news, user_notes)
    step("Synthèse et signal")
    pillars = clamp_pillars(pillar_dict(val, fd, tech, mac, news))
    cov = assess(c, md, fd, interim)
    dec = decide(pillars, st, val, tech, news, tech.liquidity if tech.available else None, major_event,
                 level=c.level)
    a = Analysis(company=c, market=md, fundamentals=fd, technical=tech, valuation=val, macro=mac, news=news,
                 decision=dec, settings=st, errors=errors, coverage=cov, key_figures=kf, interim=interim,
                 publications=pubs)
    a.texts = build_commentary(a)
    if use_claude and st.claude_api_key:
        step("Rédaction assistée")
        enriched = claude_enrich(a, a.texts, st.claude_api_key, st.claude_model)
        if enriched and enriched.get("these"):
            a.texts["these"] = enriched["these"]
            if enriched.get("catalyseurs"):
                a.texts["catalyseurs"] = enriched["catalyseurs"]
            if enriched.get("risques"):
                a.texts["risques_ia"] = enriched["risques"]
            a.ai_used = True
        else:
            a.errors.append("Rédaction assistée indisponible : commentaire automatique conservé.")
    return a


def _apply_momentum(fd: Fundamentals, rep: InterimReport) -> None:
    """Dynamique de l'exercice en cours (dernière publication) intégrée au pilier croissance."""
    mo = momentum(rep)
    parts = [(scale(mo.get("rev"), -0.10, 0.20), 50), (scale(mo.get("res"), -0.20, 0.30), 50)]
    num = sum(s_ * w for s_, w in parts if is_num(s_))
    den = sum(w for s_, w in parts if is_num(s_))
    if den:
        dyn = num / den
        fd.growth_parts = list(fd.growth_parts) + [(f"Dynamique {rep.period or 'publication récente'}", dyn, 30)]
        fd.growth_score = dyn if fd.growth_score is None else (0.7 * fd.growth_score + 0.3 * dyn)


def _merge_annual(c: CompanyData, rep: InterimReport) -> None:
    """Exercice annuel publié absent des données d'un titre sans états financiers détaillés."""
    if rep.kind == "annuel" and rep.year and c.level in ("chiffres_cles", "marche"):
        mapping = {"pnb": "pnb", "ca": "ca", "rn": "rn", "re": "re", "rao": "rao", "ebe": "ebe",
                   "rbe": "rbe", "cout_risque": "cout_risque"}
        added = False
        for ln in rep.lines:
            key = mapping.get(ln.key)
            if key and is_num(ln.current) and not is_num(c.fields.get(key, {}).get(rep.year)):
                c.fields.setdefault(key, {})[rep.year] = ln.current
                added = True
                if is_num(ln.previous) and not is_num(c.fields[key].get(rep.year - 1)):
                    c.fields[key][rep.year - 1] = ln.previous
        if added:
            c.years = sorted(set(c.years) | {rep.year, rep.year - 1})
            if c.level == "marche":
                c.level = "chiffres_cles"
            c.notes.append(f"Exercice {rep.year} complété par la publication « {rep.title[:90]} ».")


def company_label(c: CompanyData) -> str:
    country = config.COUNTRIES.get(c.country, {}).get("nom", c.country)
    return f"{c.name} ({c.ticker}) - {c.sector} - {country}"
