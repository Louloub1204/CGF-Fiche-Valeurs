"""Actualités de l'émetteur : flux Google Actualités (RSS, sans clé), tonalité
lexicale en français et repérage des événements sur titre (dividende, AG,
augmentation de capital, OPA, cession...)."""
from __future__ import annotations

import re
import urllib.parse
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime

import requests

from . import config
from .utils import clamp, is_num, strip_accents

RSS = "https://news.google.com/rss/search?q={q}&hl=fr&gl={gl}&ceid={gl}:fr"

POSITIVE = [
    "hausse", "progresse", "progression", "croissance", "benefice", "record", "amelior", "rebond",
    "solide", "renforce", "succes", "augmente", "dividende en hausse", "acquisition", "contrat",
    "partenariat", "expansion", "investit", "investissement", "leader", "performant", "excedent",
    "notation relevee", "relevement", "s'envole", "bondit", "surperform", "optimis", "dynamique",
    "resultats en hausse", "chiffre d'affaires en hausse", "agrement", "lancement", "inaugur",
]
NEGATIVE = [
    "baisse", "recul", "perte", "chute", "deficit", "litige", "sanction", "suspension", "suspendu",
    "degradation", "retard", "difficulte", "crise", "fraude", "avertissement", "contentieux",
    "demission", "plonge", "decroche", "effondr", "greve", "amende", "redressement", "liquidation",
    "defaut", "impaye", "provision", "deprecia", "abaisse", "sous-perform", "inquiet", "tension",
    "resultats en baisse", "conflit", "penalite", "radiation", "alerte",
]
EVENTS = {
    "Dividende / détachement": ["dividende", "detachement", "mise en paiement"],
    "Assemblée générale": ["assemblee generale", " ago ", " age ", "convocation"],
    "Résultats publiés": ["resultats", "etats financiers", "chiffre d'affaires", "rapport annuel", "semestriel", "trimestriel"],
    "Opération sur capital": ["augmentation de capital", "fractionnement", "split", "emission d'actions", "attribution gratuite"],
    "Changement de contrôle / OPA": ["opa", "ope ", "offre publique", "cession", "rachat", "prise de controle", "changement de controle", "fusion", "acquisition"],
    "Gouvernance": ["directeur general", "nomination", "conseil d'administration", "demission", "president du conseil"],
    "Réglementaire / judiciaire": ["sanction", "amende", "litige", "contentieux", "tribunal", "suspension", "amf-umoa", "crepmf"],
}
MAJOR_EVENTS = {"Changement de contrôle / OPA", "Opération sur capital", "Réglementaire / judiciaire"}


@dataclass
class NewsItem:
    title: str
    link: str
    published: datetime | None
    source: str
    sentiment: float
    events: list[str] = field(default_factory=list)


@dataclass
class News:
    items: list[NewsItem] = field(default_factory=list)
    query: str = ""
    score: float | None = None
    sentiment: float | None = None
    events: dict[str, list[NewsItem]] = field(default_factory=dict)
    pending_major: list[str] = field(default_factory=list)
    user_notes: str = ""
    error: str = ""


def _norm(t: str) -> str:
    return " " + strip_accents(t).lower().replace("’", "'") + " "


def sentiment_of(text: str) -> float:
    t = _norm(text)
    pos = sum(1 for w in POSITIVE if w in t)
    neg = sum(1 for w in NEGATIVE if w in t)
    if pos == neg == 0:
        return 0.0
    return (pos - neg) / (pos + neg)


def events_of(text: str) -> list[str]:
    t = _norm(text)
    return [k for k, words in EVENTS.items() if any(w in t for w in words)]


def _short_name(name: str) -> str:
    s = re.sub(r"(?i)\s+(c[oô]te d'ivoire|s[eé]n[eé]gal|burkina faso|b[eé]nin|mali|niger|togo|ci)$", "", name).strip()
    return s or name


def fetch_news(ticker: str, name: str, country: str = "CI", days: int = 365, limit: int = 25) -> News:
    gl = config.COUNTRIES.get(country, {}).get("gl", "CI")
    short = _short_name(name)
    q = f'("{name}" OR "{short}" OR {ticker}) when:{days}d'
    try:
        url = RSS.format(q=urllib.parse.quote(q), gl=gl)
        r = requests.get(url, timeout=15, headers={"User-Agent": "Mozilla/5.0 FicheValeurBRVM/1.0"})
        r.raise_for_status()
        content = r.content
    except Exception as exc:
        return News(query=q, error=f"Actualités indisponibles ({type(exc).__name__}).")
    news = parse_rss(content, ticker, name, limit)
    news.query = q
    return news


GENERIC = {"cote", "ivoire", "ivoirienne", "societe", "banque", "bank", "africa", "afrique", "international",
           "internationale", "group", "groupe", "nationale", "national", "compagnie", "marketing", "energie",
           "energy", "motors", "senegal", "burkina", "faso", "benin", "togo", "mali", "niger", "general", "generale"}


def _relevant(t_norm: str, ticker: str, short: str) -> bool:
    if re.search(rf"\b{re.escape(ticker.lower())}[a-z]?\b", t_norm):
        return True
    words = re.findall(r"[a-z0-9]{3,}", _norm(short))
    distinctive = [w for w in words if w not in GENERIC and len(w) >= 4]
    if distinctive:
        return any(w in t_norm for w in distinctive)
    phrase = " ".join(words)
    return bool(phrase) and phrase in " ".join(re.findall(r"[a-z0-9]{3,}", t_norm))


def parse_rss(content: bytes, ticker: str, name: str, limit: int = 25) -> News:
    """Flux RSS Google Actualités -> articles pertinents, tonalité et événements."""
    news = News()
    short = _short_name(name)
    try:
        root = ET.fromstring(content)
    except Exception as exc:
        news.error = f"Flux d'actualités illisible ({type(exc).__name__})."
        return news
    seen = set()
    for it in root.iter("item"):
        title = (it.findtext("title") or "").strip()
        if not title:
            continue
        src_el = it.find("source")
        source = (src_el.text or "").strip() if src_el is not None else ""
        if source and title.endswith(" - " + source):
            title = title[: -len(source) - 3].strip()
        key = strip_accents(title).lower()[:80]
        if key in seen:
            continue
        seen.add(key)
        pub = None
        try:
            pub = parsedate_to_datetime(it.findtext("pubDate") or "")
        except Exception:
            pass
        # pertinence : le titre doit citer la société (mot distinctif) ou le ticker
        t_norm = _norm(title)
        if not _relevant(t_norm, ticker, short):
            continue
        news.items.append(NewsItem(title=title, link=it.findtext("link") or "", published=pub, source=source,
                                   sentiment=sentiment_of(title), events=events_of(title)))
    news.items.sort(key=lambda n: n.published or datetime(1970, 1, 1, tzinfo=timezone.utc), reverse=True)
    news.items = news.items[:limit]
    _aggregate(news)
    return news


def _aggregate(news: News) -> None:
    now = datetime.now(timezone.utc)
    num, den = 0.0, 0.0
    for n in news.items:
        age = (now - n.published).days if n.published else 180
        w = 1.0 if age <= 30 else (0.6 if age <= 120 else 0.3)
        num += n.sentiment * w
        den += w
        for e in n.events:
            news.events.setdefault(e, []).append(n)
            if e in MAJOR_EVENTS and age <= 60 and e not in news.pending_major:
                news.pending_major.append(e)
    if den:
        news.sentiment = num / den
        # Peu d'articles : le score est rapproché de la neutralité
        conf = min(1.0, len(news.items) / 8)
        news.score = clamp(50 + 50 * news.sentiment * conf)


def add_user_notes(news: News, text: str) -> None:
    """Informations communiquées par la société, saisies par l'analyste."""
    if not text or not text.strip():
        return
    news.user_notes = text.strip()
    s = sentiment_of(text)
    if news.score is None:
        news.score = clamp(50 + 35 * s)
    else:
        news.score = clamp(0.6 * news.score + 0.4 * (50 + 50 * s))
