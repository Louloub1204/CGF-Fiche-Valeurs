"""Paramètres proposés automatiquement pour le titre choisi.

Chaque champ de la barre latérale reçoit une valeur par défaut propre au titre, avec sa source :

* cours, nombre de titres, flottant, dernier dividende : ceux que l'analyse utilisera (fichiers d'abord,
  web ensuite) ; l'utilisateur ne saisit plus que ce qu'il veut corriger ;
* taux sans risque : taux souverain de référence du pays de cotation de l'émetteur ;
* bêta par défaut : bêta médian des titres du même secteur (base de marché), utilisé si la régression du
  titre n'est pas significative ;
* seuils ACHAT / VENDRE : proportionnés à la volatilité du titre ;
* pondération des piliers : technique réduite pour un titre peu liquide, qualité et croissance neutralisées
  sans états financiers ;
* événement majeur et informations communiquées : annonces officielles récentes de la société.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, timedelta

import numpy as np
import pandas as pd

from . import config
from .market import INDEX_TICKER, MarketBase, build_market_data
from .technical import analyse_technical
from .utils import clamp, fmt_num, fmt_pct, is_num, norm_label

# Événements qui justifient une position d'attente (le fractionnement, neutre pour la valeur, n'en fait pas partie)
MAJOR_KEYWORDS = {
    "OFFRE PUBLIQUE": "offre publique", "OPA": "offre publique d'achat", "OPR": "offre publique de retrait",
    "OPE": "offre publique d'échange", "AUGMENTATION DE CAPITAL": "augmentation de capital",
    "SUSPENSION": "suspension de cotation", "RADIATION": "radiation", "RETRAIT DE LA COTE": "retrait de la cote",
    "FUSION": "fusion", "SCISSION": "scission", "CESSION DE BLOC": "cession de bloc",
}
INFO_KEYWORDS = ("DIVIDENDE", "ASSEMBLEE", "AGO", "AGE", "FRACTIONNEMENT", "RAPPORT D ACTIVITE", "ETATS FINANCIERS",
                 "NOTATION", "COMMUNIQUE", "NOMINATION", "CONSEIL D ADMINISTRATION")


@dataclass
class AutoParams:
    ticker: str
    price: float | None = None
    price_src: str = ""
    shares: float | None = None
    shares_src: str = ""
    float_pct: float | None = None
    float_src: str = ""
    dps: float | None = None
    dps_exercice: int | None = None
    dps_src: str = ""
    rf: float = 6.5
    rf_src: str = ""
    beta: float = 1.0
    beta_src: str = ""
    seuil_achat: float = 15.0
    seuil_vente: float = -15.0
    seuil_src: str = ""
    poids: dict = field(default_factory=dict)
    poids_src: str = ""
    major_event: bool = False
    major_src: str = ""
    notes: str = ""
    notes_src: str = ""


def sector_beta(mb: MarketBase | None, sector: str, exclude: str = "") -> tuple[float | None, int]:
    """Bêta médian (ajusté de Blume) des titres du secteur, en hebdomadaire sur deux ans, contre le BRVM Composite."""
    if mb is None or mb.closes is None or INDEX_TICKER not in mb.closes:
        return None, 0
    closes = mb.closes
    start = closes.index.max() - pd.Timedelta(days=730)
    wk = closes[closes.index >= start].resample("W-FRI").last()
    rets = wk.pct_change(fill_method=None)
    idx = rets[INDEX_TICKER]
    betas = []
    for t, ref in config.TICKERS.items():
        if t == exclude or ref.get("secteur") != sector or t not in rets:
            continue
        pair = pd.concat([rets[t], idx], axis=1).dropna()
        pair = pair[(pair.iloc[:, 0] != 0)]
        if len(pair) < 40 or pair.iloc[:, 1].var() == 0:
            continue
        corr = pair.iloc[:, 0].corr(pair.iloc[:, 1])
        if not is_num(corr) or corr < 0.15:
            continue
        b = pair.iloc[:, 0].cov(pair.iloc[:, 1]) / pair.iloc[:, 1].var()
        betas.append(0.67 * b + 0.33)
    if len(betas) < 2:
        return None, len(betas)
    return float(clamp(float(np.median(betas)), 0.6, 1.4)), len(betas)


def find_announcements(ticker: str, pages: int = 6, days: int = 180, errors: list[str] | None = None) -> list:
    """Annonces officielles récentes de la société (toutes rubriques), d'après la liste générale richbourse."""
    from .market import TIMEOUT, _cached, _session
    from .publications import BASE, _ticker_of_pub, parse_listing

    def page(n: int):
        def run():
            try:
                r = _session().get(f"{BASE}/common/actualite/index?page={n}", timeout=TIMEOUT)
                r.raise_for_status()
            except Exception as exc:
                if errors is not None:
                    errors.append(f"richbourse.com (annonces, page {n}) : {type(exc).__name__}")
                return None
            return parse_listing(r.text)
        return _cached(f"rb:all:{n}", run)

    limit = date.today() - timedelta(days=days)
    out = []
    for n in range(1, pages + 1):
        pubs = page(n)
        if not pubs:
            break
        out += [p for p in pubs if _ticker_of_pub(p) == ticker and (p.published is None or p.published >= limit)]
        oldest = min((p.published for p in pubs if p.published), default=None)
        if oldest and oldest < limit:
            break
    return out


def classify_announcements(pubs: list) -> tuple[list[str], list[str]]:
    """(événements majeurs, informations utiles) à partir des intitulés."""
    major, info = [], []
    for p in pubs:
        n = " " + norm_label(p.title) + " "
        when = p.published.strftime("%d/%m/%Y") if p.published else ""
        hit = next((lab for k, lab in MAJOR_KEYWORDS.items() if f" {k} " in n), None)
        if hit:
            major.append(f"{when} : {p.title}")
        elif any(f" {k}" in n for k in INFO_KEYWORDS):
            info.append(f"{when} : {p.title}")
    return major, info


def suggest(ticker: str, wb=None, mb: MarketBase | None = None, online: bool = False) -> AutoParams:
    ref = config.TICKERS.get(ticker, {})
    c = wb.companies.get(ticker) if wb is not None else None
    sector = ref.get("secteur") or (c.sector if c else "Non classé")
    country = ref.get("pays") or (c.country if c else "CI")
    ap = AutoParams(ticker=ticker)

    md = build_market_data(ticker, c, mb, None, online=online, years=3)
    ap.price, ap.price_src = md.price, md.sources.get("cours", "")
    if md.price_date:
        ap.price_src += f", au {md.price_date:%d/%m/%Y}"
    ap.shares, ap.shares_src = md.shares, md.sources.get("nombre de titres", "")
    ap.float_pct, ap.float_src = md.float_pct, md.sources.get("flottant", "")
    d = md.last_dividend()
    if d is not None:
        ap.dps, ap.dps_exercice, ap.dps_src = d.montant_net, d.exercice, d.source

    # Taux sans risque du pays de l'émetteur
    ap.rf = config.TAUX_SOUVERAINS.get(country, 6.5)
    ap.rf_src = (f"taux souverain de référence {config.COUNTRIES.get(country, {}).get('nom', country)} "
                 f"(indicatif, config.TAUX_SOUVERAINS)")

    # Bêta par défaut : médiane du secteur
    sb, n = sector_beta(mb, sector, exclude=ticker)
    if sb is not None:
        ap.beta, ap.beta_src = round(sb, 2), f"bêta médian de {n} titres du secteur {sector}"
    else:
        ap.beta_src = "valeur neutre (pas assez de titres du secteur dans la base)"

    # Volatilité et liquidité du titre
    tech = analyse_technical(md.history, md.index_history, md.float_shares)
    vol = tech.indicators.get("volatilite") if tech.available else None
    if is_num(vol):
        thr = float(clamp(round(0.5 * vol * 100), 12, 30))
        ap.seuil_achat, ap.seuil_vente = thr, -thr
        ap.seuil_src = f"la moitié de la volatilité annualisée ({fmt_pct(vol)}), bornée entre 12% et 30%"
    else:
        ap.seuil_src = "valeur standard (volatilité non mesurable)"

    poids = dict(config.Settings().poids)
    reasons = []
    liq = tech.liquidity.get("taux_seances_actives") if tech.available else None
    if is_num(liq) and liq < 0.35:
        poids["technique"], poids["valorisation"] = 10.0, 40.0
        reasons.append(f"titre peu liquide ({fmt_num(liq * 100, 0)}% de séances actives) : technique allégée")
    if c is None and not (mb is not None and ticker in getattr(mb, "fundamentals", {})):
        poids["valorisation"] += poids["qualite"] / 2
        reasons.append("pas d'états financiers chargés : la qualité et la croissance dépendront des données en ligne")
    if not tech.available:
        poids["technique"] = 0.0
        reasons.append("historique de cours insuffisant : pilier technique neutralisé")
    ap.poids = poids
    ap.poids_src = "; ".join(reasons) if reasons else "pondération standard"

    # Annonces officielles
    if online:
        pubs = find_announcements(ticker)
        major, info = classify_announcements(pubs)
        if major:
            ap.major_event = True
            ap.major_src = "annonce récente : " + major[0]
        lines = major + info
        if lines:
            ap.notes = "\n".join(lines[:6])
            ap.notes_src = "annonces officielles des six derniers mois (richbourse.com)"
    return ap
