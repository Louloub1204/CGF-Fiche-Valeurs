"""Orchestration : du classeur des états financiers à l'analyse complète."""
from __future__ import annotations

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
from .partial import Coverage, KeyFigures, assess, build_company, fetch_sika_societe
from .scoring import Decision, clamp_pillars, decide, pillar_dict
from .technical import Technical, analyse_technical
from .utils import is_num
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


def resolve_company(wb: Workbook | None, ticker: str, st: Settings, figures: KeyFigures | None,
                    errors: list[str]) -> tuple[CompanyData, KeyFigures | None]:
    """Société du classeur si elle y figure ; sinon chiffres clés (saisis ou en ligne) ; sinon marché seul."""
    if wb is not None and ticker in wb.companies and figures is None:
        return wb.companies[ticker], None
    kf = figures
    if kf is None and st.en_ligne:
        kf = fetch_sika_societe(ticker, config.TICKERS.get(ticker, {}).get("pays"), errors)
    return build_company(ticker, kf), kf


def run_analysis(wb: Workbook | None, ticker: str, settings: Settings | None = None,
                 market_base: MarketBase | None = None, price_file=None, overrides: dict | None = None,
                 user_notes: str = "", major_event: bool = False, bceao: dict | None = None,
                 use_claude: bool = False, progress=None, figures: KeyFigures | None = None,
                 custom: dict | None = None) -> Analysis:
    """Analyse d'un titre. ``wb`` peut être vide ou absent : le titre est alors traité au niveau 2 ou 3.

    ``figures`` : chiffres clés saisis (prioritaires sur ceux du classeur) ; ``custom`` : nom, pays et
    secteur d'un titre absent du référentiel.
    """
    st = settings or Settings()
    if custom:
        config.register_ticker(ticker, custom.get("nom", ticker), custom.get("pays", "CI"),
                               custom.get("secteur", "Non classé"))
    pre_errors: list[str] = []
    c, kf = resolve_company(wb, ticker, st, figures, pre_errors)

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
    step("Analyse fondamentale")
    fd = analyse_fundamentals(c, wb, md.peer_prices, md.peer_shares, use_defaults=st.multiples_defaut)
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
    cov = assess(c, md, fd)
    dec = decide(pillars, st, val, tech, news, tech.liquidity if tech.available else None, major_event,
                 level=c.level)
    a = Analysis(company=c, market=md, fundamentals=fd, technical=tech, valuation=val, macro=mac, news=news,
                 decision=dec, settings=st, errors=errors, coverage=cov, key_figures=kf)
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


def company_label(c: CompanyData) -> str:
    country = config.COUNTRIES.get(c.country, {}).get("nom", c.country)
    return f"{c.name} ({c.ticker}) - {c.sector} - {country}"
