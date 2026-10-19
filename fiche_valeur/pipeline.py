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
from .scoring import Decision, clamp_pillars, decide, pillar_dict
from .technical import Technical, analyse_technical
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


def run_analysis(wb: Workbook, ticker: str, settings: Settings | None = None,
                 market_base: MarketBase | None = None, price_file=None, overrides: dict | None = None,
                 user_notes: str = "", major_event: bool = False, bceao: dict | None = None,
                 use_claude: bool = False, progress=None) -> Analysis:
    st = settings or Settings()
    c = wb.companies[ticker]

    def step(msg):
        if progress:
            progress(msg)

    step("Données de marché")
    md = build_market_data(ticker, c, market_base, price_file, online=st.en_ligne,
                           years=st.annees_historique, overrides=overrides)
    step("Analyse fondamentale")
    fd = analyse_fundamentals(c, wb, md.peer_prices, md.peer_shares)
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
    dec = decide(pillars, st, val, tech, news, tech.liquidity if tech.available else None, major_event)
    a = Analysis(company=c, market=md, fundamentals=fd, technical=tech, valuation=val, macro=mac, news=news,
                 decision=dec, settings=st, errors=errors)
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
