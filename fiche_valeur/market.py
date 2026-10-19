"""Données de marché : cours, volumes, nombre de titres, dividendes, indice.

Sources, par ordre de priorité (la première disponible l'emporte) :

1. saisie manuelle dans l'interface (cours, nombre de titres, dividende) ;
2. base de marché chargée par l'utilisateur (classeur « Base de données -
   SMF » : feuilles Cours, Nb_titres, Moyenne_cours, et le cas échéant
   Volumes et Dividendes), ou fichier d'historique d'un titre ;
3. sources publiques en ligne :
   * historique OHLCV séance par séance : API de sikafinance.com
     (POST /api/general/GetHistos, tranches de 89 jours, Referer requis) ;
   * nombre de titres, cours du jour, capitalisation flottante :
     brvm.org/fr/capitalisations/0 ;
   * calendrier des dividendes : brvm.org/fr/esv/paiement-de-dividendes et
     sikafinance.com/marches/dividendes ;
4. données déjà présentes dans le classeur des états financiers (dividende
   par action, nombre d'actions).

Toutes les fonctions réseau échouent proprement : elles renvoient ``None``
ou un tableau vide et consignent la raison dans ``errors``.
"""
from __future__ import annotations

import io
import re
import time
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta

import numpy as np
import pandas as pd
import requests

from . import config
from .utils import is_num, norm_header, norm_label, strip_accents, to_float

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0 Safari/537.36 FicheValeurBRVM/1.0")
TIMEOUT = 20

SIKA_API = "https://www.sikafinance.com/api/general/GetHistos"
SIKA_PAGE = "https://www.sikafinance.com/marches/historiques/{sym}"
SIKA_DIV = "https://www.sikafinance.com/marches/dividendes"
BRVM_CAPI = "https://www.brvm.org/fr/capitalisations/0"
BRVM_DIV = "https://www.brvm.org/fr/esv/paiement-de-dividendes"
FX_API = "https://api.frankfurter.app/latest?from=EUR&to=USD"
INDEX_TICKER = "BRVMC"


# ---------------------------------------------------------------------------
# Modèles
# ---------------------------------------------------------------------------


@dataclass
class Dividend:
    exercice: int | None
    montant_net: float
    date_paiement: date | None = None
    source: str = ""


@dataclass
class MarketData:
    ticker: str
    history: pd.DataFrame | None = None        # index Date ; open, high, low, close, volume
    index_history: pd.DataFrame | None = None  # BRVM Composite
    price: float | None = None
    price_date: date | None = None
    shares: float | None = None
    float_shares: float | None = None
    dividends: list[Dividend] = field(default_factory=list)
    avg_prices: dict[int, float] = field(default_factory=dict)
    peer_prices: dict[str, float] = field(default_factory=dict)   # cours de tous les titres
    peer_shares: dict[str, float] = field(default_factory=dict)   # nombre de titres de tous les titres
    sources: dict[str, str] = field(default_factory=dict)
    errors: list[str] = field(default_factory=list)

    @property
    def market_cap(self) -> float | None:
        if is_num(self.price) and is_num(self.shares):
            return float(self.price) * float(self.shares)
        return None

    @property
    def float_pct(self) -> float | None:
        if is_num(self.float_shares) and is_num(self.shares) and self.shares:
            return float(self.float_shares) / float(self.shares)
        return None

    def last_dividend(self) -> Dividend | None:
        """Dernier dividende connu ; None si le dernier exercice renseigné n'a pas donné lieu à distribution."""
        divs = [d for d in self.dividends if is_num(d.montant_net) and d.montant_net >= 0]
        if not divs:
            return None
        divs = sorted(divs, key=lambda d: (d.exercice or 0, d.date_paiement or date(1900, 1, 1)))
        last = divs[-1]
        if last.montant_net <= 0:
            return None
        prev = [d.montant_net for d in divs[-5:-1] if d.montant_net > 0]
        if len(prev) >= 2:
            med = float(np.median(prev))
            if last.montant_net > 2.5 * med:
                # versement exceptionnel : le dividende ordinaire sert aux calculs de rendement et d'actualisation
                return Dividend(last.exercice, med, last.date_paiement,
                                f"{last.source} ; dernier versement de {last.montant_net:,.0f} FCFA jugé "
                                f"exceptionnel, dividende ordinaire retenu".replace(",", " "))
        return last


@dataclass
class MarketBase:
    """Contenu d'un classeur de marché type « Base de données - SMF »."""
    closes: pd.DataFrame | None = None        # index Date, colonnes tickers
    volumes: pd.DataFrame | None = None
    shares_total: dict[str, float] = field(default_factory=dict)
    shares_float: dict[str, float] = field(default_factory=dict)
    avg_prices: dict[str, dict[int, float]] = field(default_factory=dict)
    dividends: dict[str, list[Dividend]] = field(default_factory=dict)
    fundamentals: dict[str, dict[str, dict[int, float]]] = field(default_factory=dict)   # Tableau de bord
    notes: list[str] = field(default_factory=list)
    volumes_in_value: bool = False

    @property
    def tickers(self) -> list[str]:
        return list(self.closes.columns) if self.closes is not None else []


# ---------------------------------------------------------------------------
# Lecture de la base de marché (classeur)
# ---------------------------------------------------------------------------
TICKER_RE = re.compile(r"^[A-Z]{3,6}$")
NOT_TICKERS = {"DATE", "DATES", "ANNEE", "JOUR", "MOIS", "NONE", "NAN", "TOTAL", "MOYENNE", "SEANCE"}


def _ticker_of(h) -> str | None:
    """En-tête de colonne -> ticker (ou BRVMC pour l'indice composite), sinon None."""
    if h is None or (isinstance(h, float) and not is_num(h)):
        return None
    t = re.sub(r"[^A-Z0-9]", "", strip_accents(str(h)).upper())
    if t in {"BRVMCOMPOSITE", "BRVMC", "BRVMCOMP", "BRVMCI"}:
        return INDEX_TICKER
    if t in NOT_TICKERS:
        return None
    return t if TICKER_RE.match(t) or t.startswith("BRVM") else None


def _to_date(v) -> date | None:
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)) and is_num(v) and 20000 < float(v) < 80000:  # numéro de série Excel
        return (datetime(1899, 12, 30) + timedelta(days=float(v))).date()
    s = str(v).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s[:19], fmt).date()
        except ValueError:
            continue
    return None


def _wide_series(raw: pd.DataFrame) -> pd.DataFrame | None:
    """Feuille large (Date + une colonne par ticker) -> DataFrame indexé par date."""
    header_row, date_col = None, None
    for i in range(min(10, len(raw))):
        row = [norm_header(x) for x in raw.iloc[i].tolist()]
        if "date" in row or "dates" in row or "seance" in row:
            header_row = i
            date_col = row.index("date") if "date" in row else (row.index("dates") if "dates" in row else row.index("seance"))
            break
    if header_row is None:
        return None
    headers = raw.iloc[header_row].tolist()
    body = raw.iloc[header_row + 1:]
    dates = body.iloc[:, date_col].map(_to_date)
    out = {}
    for j, h in enumerate(headers):
        if j == date_col:
            continue
        t = _ticker_of(h)
        if t is None or t in out:       # premier bloc seulement si la feuille répète les tickers
            continue
        out[t] = pd.to_numeric(body.iloc[:, j].map(to_float), errors="coerce").values
    if not out:
        return None
    df = pd.DataFrame(out, index=pd.to_datetime(dates.values, errors="coerce"))
    df = df[~df.index.isna()].sort_index()
    df = df[~df.index.duplicated(keep="last")]
    return df


def load_market_base(source) -> MarketBase:
    data = source.read() if hasattr(source, "read") else open(source, "rb").read()
    xls = pd.ExcelFile(io.BytesIO(data))
    mb = MarketBase()
    for sheet in xls.sheet_names:
        key = norm_header(sheet)
        raw = pd.read_excel(xls, sheet_name=sheet, header=None)
        if raw.empty:
            continue
        if key in {"cours", "cours de cloture", "cours cloture", "historique", "prix"}:
            mb.closes = _wide_series(raw)
        elif key.startswith("volume"):
            mb.volumes = _wide_series(raw)
        elif key.startswith("nb titres") or key.startswith("nombre de titres") or key == "titres":
            _read_shares(raw, mb)
        elif key.startswith("moyenne"):
            _read_avg(raw, mb)
        elif "divid" in key:
            _read_dividends_sheet(raw, mb)
        elif key.startswith("tableau de bord") or key.startswith("fondamentaux") or key.startswith("donnees financieres"):
            _read_dashboard(raw, mb)
    if mb.closes is None:
        mb.notes.append("Feuille « Cours » introuvable dans la base de marché.")
    if mb.closes is not None:
        mb.closes = mb.closes.dropna(how="all")
    if mb.volumes is not None and mb.closes is not None:
        mb.volumes_in_value = _volumes_are_values(mb.volumes, mb.closes)
        if mb.volumes_in_value:
            vol = mb.volumes.reindex(mb.closes.index)
            common = [c for c in vol.columns if c in mb.closes.columns]
            mb.volumes = (vol[common] / mb.closes[common].replace(0, np.nan)).round()
            mb.notes.append("Volumes de la base exprimés en FCFA : convertis en nombre de titres (montant / cours).")
    return mb


def _volumes_are_values(vol: pd.DataFrame, closes: pd.DataFrame) -> bool:
    """Vrai si les volumes sont des montants échangés : montant / cours tombe sur un nombre entier de titres."""
    common = [c for c in vol.columns if c in closes.columns and c != INDEX_TICKER]
    if not common:
        return False
    v = vol[common].reindex(closes.index)
    ratio = (v / closes[common]).stack()
    ratio = ratio[(ratio > 0.5) & np.isfinite(ratio)]
    if len(ratio) < 30:
        return False
    ratio = ratio.sample(min(len(ratio), 3000), random_state=0)
    near_int = (np.abs(ratio - ratio.round()) < 0.02).mean()
    return bool(near_int > 0.6)


DASHBOARD_FIELDS = {
    "CAPITAUX PROPRES": "capitaux_propres", "RESULTAT NET": "rn", "CHIFFRE D AFFAIRES": "ca",
    "PRODUIT NET BANCAIRE": "pnb", "RESULTAT D EXPLOITATION": "re", "FLUX DE TRESORERIE": "flux_tresorerie",
    "TOTAL ACTIF": "total_actif", "TOTAL BILAN": "total_actif", "EBE": "ebe", "EXCEDENT BRUT D EXPLOITATION": "ebe",
    "VALEUR AJOUTEE": "va", "RESULTAT DES ACTIVITES ORDINAIRES": "rao", "PNB": "pnb",
}


def _read_dashboard(raw: pd.DataFrame, mb: MarketBase) -> None:
    """Feuille « Tableau de bord » : blocs (titre de l'agrégat, ligne des tickers, années en lignes), en FCFA."""
    header: list | None = None
    current: str | None = None
    for i in range(len(raw)):
        row = raw.iloc[i].tolist()
        texts = [norm_label(x) for x in row if isinstance(x, str)]
        tick_cells = sum(1 for x in row if isinstance(x, str) and _ticker_of(x) not in (None, INDEX_TICKER))
        if tick_cells >= 2:
            header = row
            continue
        label = next((DASHBOARD_FIELDS[t] for t in texts if t in DASHBOARD_FIELDS), None)
        if label is not None:
            current = label
            continue
        y = next((int(v) for v in row[:3] if is_num(v) and 1990 < float(v) < 2100), None)
        if y is None or header is None or current is None:
            continue
        for j, h in enumerate(header):
            t = _ticker_of(h) if isinstance(h, str) else None
            if t is None or t == INDEX_TICKER or j >= len(row):
                continue
            v = to_float(row[j])
            if is_num(v) and v != 0:          # 0 = donnée non renseignée dans la base
                mb.fundamentals.setdefault(t, {}).setdefault(current, {})[y] = float(v)


def _read_shares(raw: pd.DataFrame, mb: MarketBase) -> None:
    header = raw.iloc[0].tolist()
    rows = {norm_header(raw.iloc[i, 0]): i for i in range(1, len(raw)) if raw.iloc[i, 0] is not None}
    i_float = next((i for k, i in rows.items() if "flott" in k), 1 if len(raw) > 1 else None)
    i_total = next((i for k, i in rows.items() if "total" in k or "nombre" in k), 2 if len(raw) > 2 else None)
    for j, h in enumerate(header):
        t = _ticker_of(h)
        if t is None or t == INDEX_TICKER or not TICKER_RE.match(t):
            continue
        if i_total is not None:
            v = to_float(raw.iloc[i_total, j])
            if is_num(v):
                mb.shares_total[t] = v
        if i_float is not None:
            v = to_float(raw.iloc[i_float, j])
            if is_num(v):
                mb.shares_float[t] = v


def _read_avg(raw: pd.DataFrame, mb: MarketBase) -> None:
    header = raw.iloc[0].tolist()
    for i in range(1, len(raw)):
        y = None
        for v in raw.iloc[i, :3].tolist():
            if is_num(v) and 1990 < float(v) < 2100:
                y = int(v)
        if y is None:
            continue
        for j, h in enumerate(header):
            t = _ticker_of(h)
            if t and t != INDEX_TICKER and TICKER_RE.match(t):
                v = to_float(raw.iloc[i, j])
                if is_num(v):
                    mb.avg_prices.setdefault(t, {})[y] = v


def _read_dividends_sheet(raw: pd.DataFrame, mb: MarketBase) -> None:
    hdr_i = None
    for i in range(min(8, len(raw))):
        row = [norm_header(x) for x in raw.iloc[i].tolist()]
        if any(("ticker" in c or "symbole" in c or "code" in c) for c in row) and any("divid" in c or "montant" in c for c in row):
            hdr_i = i
            break
    if hdr_i is not None:  # format long
        cols = [norm_header(x) for x in raw.iloc[hdr_i].tolist()]

        def col(*keys):
            for k in keys:
                for j, c in enumerate(cols):
                    if k in c:
                        return j
            return None
        jt, jy, jm, jd = col("ticker", "symbole", "code"), col("exercice", "annee"), col("net", "montant", "divid"), col("paiement", "date")
        for i in range(hdr_i + 1, len(raw)):
            t = strip_accents(str(raw.iloc[i, jt] or "")).upper().strip() if jt is not None else ""
            m = to_float(raw.iloc[i, jm]) if jm is not None else None
            if not TICKER_RE.match(t) or not is_num(m):
                continue
            y = to_float(raw.iloc[i, jy]) if jy is not None else None
            d = _to_date(raw.iloc[i, jd]) if jd is not None else None
            mb.dividends.setdefault(t, []).append(
                Dividend(int(y) if is_num(y) else (d.year - 1 if d else None), m, d, "Base de marché"))
        return
    # format large : années de paiement en lignes, tickers en colonnes (dividende de l'exercice N-1 payé en N)
    header = raw.iloc[0].tolist()
    for i in range(1, len(raw)):
        y = next((int(v) for v in raw.iloc[i, :3].tolist() if is_num(v) and 1990 < float(v) < 2100), None)
        if y is None:
            continue
        cells = []
        for j, h in enumerate(header):
            t = _ticker_of(h)
            if t and t != INDEX_TICKER and TICKER_RE.match(t):
                cells.append((t, to_float(raw.iloc[i, j])))
        filled = sum(1 for _, v in cells if is_num(v))
        for t, v in cells:
            if is_num(v) and v >= 0:      # 0 = pas de dividende cette année-là
                mb.dividends.setdefault(t, []).append(
                    Dividend(y - 1, v, date(y, 7, 1), "Base de marché (Historique_dividende)"))
            elif not is_num(v) and filled >= 3 and t in {c_[0] for c_ in cells if is_num(c_[1])} | set(mb.dividends):
                # saison renseignée pour d'autres titres : une cellule vide vaut absence de versement
                mb.dividends.setdefault(t, []).append(
                    Dividend(y - 1, 0.0, date(y, 7, 1), "Base de marché (Historique_dividende, cellule vide = aucun versement)"))


def load_price_file(source, filename: str = "") -> pd.DataFrame | None:
    """Historique d'un titre (CSV ou Excel) : Date, Ouverture, Haut, Bas, Clôture, Volume."""
    data = source.read() if hasattr(source, "read") else open(source, "rb").read()
    name = (filename or getattr(source, "name", "")).lower()
    try:
        if name.endswith(".csv") or name.endswith(".txt"):
            text = data.decode("utf-8", errors="ignore")
            sep = ";" if text.count(";") > text.count(",") else ","
            df = pd.read_csv(io.StringIO(text), sep=sep)
        else:
            df = pd.read_excel(io.BytesIO(data))
    except Exception:
        return None
    return normalise_ohlcv(df)


ALIASES_OHLCV = {
    "date": ("date", "seance", "jour"),
    "open": ("ouverture", "open", "premier"),
    "high": ("plus haut", "high", "haut"),
    "low": ("plus bas", "low", "bas"),
    "close": ("cloture", "close", "dernier", "cours"),
    "volume": ("volume titres", "volume", "quantite", "titres"),
}


def normalise_ohlcv(df: pd.DataFrame) -> pd.DataFrame | None:
    cols = {c: norm_header(c) for c in df.columns}
    mapping, used = {}, set()
    for target, keys in ALIASES_OHLCV.items():
        for k in keys:
            hit = next((c for c, n in cols.items() if c not in used and (n == k or n.startswith(k))), None)
            if hit is not None:
                mapping[target] = hit
                used.add(hit)
                break
    if "date" not in mapping or "close" not in mapping:
        return None
    out = pd.DataFrame({k: df[v] for k, v in mapping.items()})
    out["date"] = pd.to_datetime(out["date"].map(_to_date), errors="coerce")
    for c in out.columns:
        if c != "date":
            out[c] = pd.to_numeric(out[c].map(to_float), errors="coerce")
    out = out.dropna(subset=["date", "close"]).sort_values("date").drop_duplicates("date", keep="last")
    out = out.set_index("date")
    for c in ("open", "high", "low"):
        if c not in out:
            out[c] = np.nan
    if "volume" not in out:
        out["volume"] = np.nan
    return out[["open", "high", "low", "close", "volume"]]


# ---------------------------------------------------------------------------
# Accès réseau
# ---------------------------------------------------------------------------
_CACHE: dict[str, tuple[float, object]] = {}
CACHE_TTL = 3600


def _cached(key: str, fn, ttl: int = CACHE_TTL):
    now = time.time()
    hit = _CACHE.get(key)
    if hit and now - hit[0] < ttl:
        return hit[1]
    val = fn()
    _CACHE[key] = (now, val)
    return val


def _session() -> requests.Session:
    s = requests.Session()
    s.headers.update({"User-Agent": UA, "Accept-Language": "fr-FR,fr;q=0.9"})
    return s


def sika_symbol(ticker: str, country: str | None = None) -> str:
    if ticker.startswith("BRVM"):
        return ticker
    suffix = config.COUNTRIES.get(country or config.TICKERS.get(ticker, {}).get("pays", "CI"), {}).get("suffixe", ".ci")
    return f"{ticker}{suffix}"


def _sika_chunk(s: requests.Session, sym: str, d0: date, d1: date) -> list[dict]:
    r = s.post(SIKA_API, json={"ticker": sym, "datedeb": d0.isoformat(), "datefin": d1.isoformat(),
                               "xperiod": "0"},
               headers={"Referer": SIKA_PAGE.format(sym=sym), "Origin": "https://www.sikafinance.com",
                        "Accept": "application/json, text/plain, */*"},
               timeout=TIMEOUT)
    r.raise_for_status()
    payload = r.json()
    rows = payload.get("lst") if isinstance(payload, dict) else payload
    return rows or []


def fetch_sikafinance_history(ticker: str, country: str | None = None, years: int = 3,
                              errors: list[str] | None = None) -> pd.DataFrame | None:
    """Historique quotidien OHLCV depuis sikafinance (volume en nombre de titres)."""
    def run():
        s = _session()
        end = date.today()
        start = end - timedelta(days=int(365.25 * years) + 5)
        sym = sika_symbol(ticker, country)
        rows: list[dict] = []
        cur = start
        chunks = []
        while cur <= end:
            nxt = min(cur + timedelta(days=88), end)
            chunks.append((cur, nxt))
            cur = nxt + timedelta(days=1)
        # Commence par la tranche la plus récente pour détecter vite un mauvais suffixe
        chunks.reverse()
        for k, (a, b) in enumerate(chunks):
            try:
                got = _sika_chunk(s, sym, a, b)
            except Exception as exc:  # réseau, 403, JSON invalide
                if errors is not None:
                    errors.append(f"sikafinance ({sym}) : {type(exc).__name__} {exc}"[:200])
                if k == 0:
                    return None
                break
            if k == 0 and not got and not ticker.startswith("BRVM"):
                # Suffixe pays alternatif
                for suf in (".ci", ".sn", ".bf", ".bj", ".tg", ".ml", ".ne"):
                    alt = f"{ticker}{suf}"
                    if alt == sym:
                        continue
                    try:
                        got = _sika_chunk(s, alt, a, b)
                    except Exception:
                        got = []
                    if got:
                        sym = alt
                        break
            rows.extend(got)
            time.sleep(0.25)
        return parse_sika_payload({"lst": rows})
    return _cached(f"sika:{ticker}:{years}", run)


def _html_tables(html: str) -> list[pd.DataFrame]:
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html, "html.parser")
    out = []
    for tb in soup.find_all("table"):
        heads = [th.get_text(" ", strip=True) for th in tb.select("thead th")]
        body = tb.select("tbody tr") or tb.find_all("tr")
        rows = []
        for tr in body:
            cells = [td.get_text(" ", strip=True) for td in tr.find_all(["td", "th"])]
            if cells:
                rows.append(cells)
        if not heads and rows:
            heads, rows = rows[0], rows[1:]
        if not rows or not heads:
            continue
        n = len(heads)
        rows = [r[:n] + [""] * (n - len(r)) for r in rows if len(r) >= max(2, n - 1)]
        if rows:
            out.append(pd.DataFrame(rows, columns=heads))
    return out


def parse_capitalisations(html: str) -> pd.DataFrame | None:
    """Tableau des capitalisations de brvm.org -> DataFrame indexé par ticker."""
    for tb in _html_tables(html):
        cols = {c: norm_header(c) for c in tb.columns}
        c_tick = next((c for c, n in cols.items() if n.startswith("code") or "symbole" in n), None)
        c_nb = next((c for c, n in cols.items() if "nombre de titres" in n), None)
        c_px = next((c for c, n in cols.items() if "cours" in n), None)
        c_flot = next((c for c, n in cols.items() if "flottante" in n), None)
        c_glob = next((c for c, n in cols.items() if "globale" in n and "%" not in c), None)
        c_nom = next((c for c, n in cols.items() if n == "nom" or "societe" in n), None)
        if not (c_tick and c_nb and c_px):
            continue
        df = pd.DataFrame({
            "ticker": tb[c_tick].str.strip().str.upper(),
            "nom": tb[c_nom] if c_nom else "",
            "nb_titres": tb[c_nb].map(to_float),
            "cours": tb[c_px].map(to_float),
            "cap_flottante": tb[c_flot].map(to_float) if c_flot else np.nan,
            "cap_globale": tb[c_glob].map(to_float) if c_glob else np.nan,
        })
        df = df[df["ticker"].str.match(r"^[A-Z]{3,6}$")]
        if not df.empty:
            return df.set_index("ticker")
    return None


def fetch_brvm_capitalisations(errors: list[str] | None = None) -> pd.DataFrame | None:
    """Nombre de titres, cours du jour et capitalisation flottante de toutes les valeurs (brvm.org)."""
    def run():
        try:
            r = _session().get(BRVM_CAPI, timeout=TIMEOUT)
            r.raise_for_status()
        except Exception as exc:
            if errors is not None:
                errors.append(f"brvm.org capitalisations : {type(exc).__name__} {exc}"[:200])
            return None
        df = parse_capitalisations(r.text)
        if df is None and errors is not None:
            errors.append("brvm.org capitalisations : tableau non reconnu.")
        return df
    return _cached("brvm:capi", run)


def parse_sika_payload(payload) -> pd.DataFrame | None:
    """Réponse JSON de l'API d'historique sikafinance -> OHLCV."""
    rows = payload.get("lst") if isinstance(payload, dict) else payload
    if not rows:
        return None
    df = pd.DataFrame(rows)
    df.columns = [str(c) for c in df.columns]
    return normalise_ohlcv(df.rename(columns={"Date": "date", "Open": "open", "High": "high", "Low": "low",
                                              "Close": "close", "Volume": "volume"}))


def _name_tokens(name: str) -> set[str]:
    stop = {"COTE", "D", "IVOIRE", "CI", "SA", "SENEGAL", "BURKINA", "FASO", "BENIN", "MALI", "NIGER", "TOGO",
            "DE", "DU", "LA", "LE", "LES", "ET", "BANQUE", "BANK", "OF", "AFRICA", "SOCIETE", "GROUPE",
            "INTERNATIONAL", "INTERNATIONALE", "THE", "INCORPORATED"}
    return {t for t in norm_label(name).split() if t not in stop and len(t) > 1}


def _name_match(target: str, candidate: str, ticker: str = "") -> bool:
    """Rapproche une raison sociale publiée du titre, par le nom puis les autres noms du référentiel."""
    others = config.TICKERS.get(ticker, {}).get("autres_noms", []) if ticker else []
    return any(_name_match_one(n, candidate, ticker) for n in [target] + list(others))


def _name_match_one(target: str, candidate: str, ticker: str = "") -> bool:
    c = norm_label(candidate)
    if ticker and re.search(rf"\b{re.escape(ticker)}\b", c):
        return True
    a, b = _name_tokens(target), _name_tokens(candidate)
    if not a or not b:
        return False
    # Les banques BOA se distinguent par leur pays : on l'exige
    country = config.TICKERS.get(ticker, {}).get("pays")
    if ticker.startswith("BOA") and country:
        pays = norm_label(config.COUNTRIES[country]["nom"]).split()[0]
        if pays not in c:
            return False
    return len(a & b) / len(a) >= 0.6


def fetch_dividends_online(ticker: str, name: str, errors: list[str] | None = None) -> list[Dividend]:
    """Dividendes publiés (brvm.org puis sikafinance), rapprochés par la raison sociale."""
    def run():
        found: list[Dividend] = []
        s = _session()
        for page in range(0, 3):
            url = BRVM_DIV + (f"?page={page}" if page else "")
            try:
                r = s.get(url, timeout=TIMEOUT)
                r.raise_for_status()
            except Exception as exc:
                if errors is not None and page == 0:
                    errors.append(f"brvm.org dividendes : {type(exc).__name__}"[:200])
                break
            tables = _html_tables(r.text)
            hit_any = False
            for tb in tables:
                cols = {c: norm_header(c) for c in tb.columns}
                c_em = next((c for c, n in cols.items() if "emetteur" in n or n in ("societe", "nom")), None)
                c_mt = next((c for c, n in cols.items() if "montant" in n or "dividende net" in n), None)
                c_ex = next((c for c, n in cols.items() if "exercice" in n), None)
                c_dp = next((c for c, n in cols.items() if "paiement" in n), None)
                if not (c_em and c_mt):
                    continue
                hit_any = True
                for _, row in tb.iterrows():
                    if not _name_match(name, str(row[c_em]), ticker):
                        continue
                    m = to_float(row[c_mt])
                    if not is_num(m) or m <= 0:
                        continue
                    ex = to_float(re.sub(r"\D", "", str(row[c_ex]))[:4]) if c_ex else None
                    d = _to_date(row[c_dp]) if c_dp else None
                    found.append(Dividend(int(ex) if is_num(ex) else (d.year - 1 if d else None), m, d, "brvm.org"))
            if not hit_any:
                break
        if found:
            return found
        try:
            r = s.get(SIKA_DIV, timeout=TIMEOUT)
            r.raise_for_status()
            for tb in _html_tables(r.text):
                cols = {c: norm_header(c) for c in tb.columns}
                c_nom = next((c for c, n in cols.items() if n in ("nom", "societe", "valeur")), None)
                c_mt = next((c for c, n in cols.items() if "montant" in n), None)
                c_dt = next((c for c, n in cols.items() if "detachement" in n or "date" in n), None)
                if not (c_nom and c_mt):
                    continue
                for _, row in tb.iterrows():
                    if not _name_match(name, str(row[c_nom]), ticker):
                        continue
                    m = to_float(row[c_mt])
                    d = _to_date(row[c_dt]) if c_dt else None
                    if is_num(m) and m > 0:
                        found.append(Dividend(d.year - 1 if d else None, m, d, "sikafinance"))
        except Exception as exc:
            if errors is not None:
                errors.append(f"sikafinance dividendes : {type(exc).__name__}"[:200])
        return found
    return _cached(f"div:{ticker}", run)


def fetch_usd_xof(errors: list[str] | None = None) -> float | None:
    def run():
        try:
            r = _session().get(FX_API, timeout=10)
            r.raise_for_status()
            eurusd = float(r.json()["rates"]["USD"])
            return config.EUR_XOF / eurusd
        except Exception as exc:
            if errors is not None:
                errors.append(f"taux EUR/USD : {type(exc).__name__}")
            return None
    return _cached("fx:usd", run)


# ---------------------------------------------------------------------------
# Assemblage
# ---------------------------------------------------------------------------

def _extend_history(file_h: pd.DataFrame, web_h: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Ajoute à l'historique du fichier les séances postérieures du web. Un écart d'un facteur entier
    (au moins 2) entre les deux sources trahit un fractionnement : l'historique du fichier est ajusté
    (la variation quotidienne de la BRVM est plafonnée, un tel saut ne peut venir du marché)."""
    last = file_h.index.max()
    common = file_h.index.intersection(web_h.index)
    ratio = None
    if len(common) >= 3:
        ratio = float((file_h.loc[common, "close"] / web_h.loc[common, "close"]).median())
        if abs(ratio - 1) < 0.08:
            ratio = None
    if ratio is None:
        after = web_h[web_h.index > last]
        if not after.empty and file_h["close"].iloc[-1] > 0:
            ratio = float(file_h["close"].iloc[-1] / after["close"].iloc[0])
    split = 1
    if ratio and ratio >= 1.8:
        n = round(ratio)
        if abs(ratio - n) < 0.12 * n:
            split = n
    base = file_h.copy()
    if split > 1:
        for col in ("open", "high", "low", "close"):
            if col in base:
                base[col] = base[col] / split
        if "volume" in base:
            base["volume"] = base["volume"] * split
    out = pd.concat([base, web_h[web_h.index > last]])
    return out[~out.index.duplicated(keep="first")].sort_index(), split


def build_market_data(ticker: str, company=None, base: MarketBase | None = None,
                      price_file: pd.DataFrame | None = None, online: bool = True,
                      years: int = 3, overrides: dict | None = None) -> MarketData:
    md = MarketData(ticker=ticker)
    overrides = overrides or {}
    country = getattr(company, "country", None) or config.TICKERS.get(ticker, {}).get("pays")
    name = getattr(company, "name", None) or config.TICKERS.get(ticker, {}).get("nom", ticker)

    # Historique de cours : fichiers d'abord, le web ne complète que les séances postérieures ----------
    split = 1
    if price_file is not None and not price_file.empty:
        md.history = price_file
        md.sources["historique"] = "Fichier d'historique chargé"
    if md.history is None and base is not None and base.closes is not None and ticker in base.closes:
        h = pd.DataFrame({"close": base.closes[ticker]})
        h["open"] = np.nan
        h["high"] = np.nan
        h["low"] = np.nan
        h["volume"] = base.volumes[ticker] if base.volumes is not None and ticker in base.volumes else np.nan
        h = h.dropna(subset=["close"])
        h = h[h["close"] > 0]
        if not h.empty:
            md.history = h[["open", "high", "low", "close", "volume"]]
            md.sources["historique"] = "Base de marché (feuille Cours)"
    if online:
        stale = md.history is None or md.history.index.max() < pd.Timestamp(date.today() - timedelta(days=4))
        if stale:
            h = fetch_sikafinance_history(ticker, country, years, md.errors)
            if h is not None and not h.empty:
                if md.history is None:
                    md.history = h
                    md.sources["historique"] = "sikafinance.com (API d'historique)"
                else:
                    md.history, split = _extend_history(md.history, h)
                    if len(md.history) and md.history.index.max() > h.index.min():
                        md.sources["historique"] = (md.sources["historique"] + ", complété par sikafinance.com "
                                                    "pour les séances postérieures")
                    if split > 1:
                        md.errors.append(f"Fractionnement de l'action détecté (1 ancienne pour {split} nouvelles) : "
                                         f"cours, volumes, dividendes et nombre de titres de la base ajustés.")
    # Indice BRVM Composite
    if base is not None and base.closes is not None and INDEX_TICKER in base.closes:
        ih = base.closes[[INDEX_TICKER]].rename(columns={INDEX_TICKER: "close"}).dropna()
        if not ih.empty:
            md.index_history = ih
            md.sources["indice"] = "Base de marché"
    if online and (md.index_history is None or md.index_history.index.max() < pd.Timestamp(date.today() - timedelta(days=4))):
        ih = fetch_sikafinance_history(INDEX_TICKER, None, years, md.errors)
        if ih is not None and not ih.empty:
            if md.index_history is None:
                md.index_history = ih[["close"]]
                md.sources["indice"] = "sikafinance.com (BRVM Composite)"
            else:
                last = md.index_history.index.max()
                md.index_history = pd.concat([md.index_history, ih.loc[ih.index > last, ["close"]]])
                md.sources["indice"] += ", complété par sikafinance.com"

    if md.history is not None and not md.history.empty:
        md.price = float(md.history["close"].iloc[-1])
        md.price_date = md.history.index[-1].date()
        md.sources["cours"] = md.sources.get("historique", "")

    # Nombre de titres : base de marché d'abord, brvm.org ensuite ---------------------------------
    if base is not None:
        if ticker in base.shares_total:
            md.shares = base.shares_total[ticker] * split
            md.sources["nombre de titres"] = "Base de marché (Nb_titres)" + (" ajusté du fractionnement" if split > 1 else "")
        if ticker in base.shares_float:
            md.float_shares = base.shares_float[ticker] * split
            md.sources["flottant"] = "Base de marché (Nb_titres)"
        if ticker in base.avg_prices:
            md.avg_prices = dict(base.avg_prices[ticker])
        if base.closes is not None:
            last = base.closes.ffill().iloc[-1]
            md.peer_prices = {t: float(v) for t, v in last.items() if is_num(v)}
        md.peer_shares = dict(base.shares_total)
    capi = fetch_brvm_capitalisations(md.errors) if online else None
    if capi is not None:
        # pairs : couple cours / titres de la base s'il est complet, sinon celui de brvm.org (même source)
        for t, row_ in capi.iterrows():
            if not (is_num(md.peer_prices.get(t)) and is_num(md.peer_shares.get(t))):
                if is_num(row_["cours"]) and is_num(row_["nb_titres"]):
                    md.peer_prices[t], md.peer_shares[t] = float(row_["cours"]), float(row_["nb_titres"])
        if ticker in capi.index:
            row = capi.loc[ticker]
            web_n = float(row["nb_titres"]) if is_num(row["nb_titres"]) else None
            if web_n and md.shares is None:
                md.shares = web_n
                md.sources["nombre de titres"] = "brvm.org (capitalisations)"
            elif web_n and md.shares:
                ratio = web_n / md.shares
                n = round(ratio)
                if n >= 2 and abs(ratio - n) < 0.05 * n:
                    md.shares, md.float_shares = web_n, (md.float_shares * n if md.float_shares else None)
                    md.sources["nombre de titres"] = "brvm.org (capitalisations), fractionnement postérieur à la base"
            if md.float_shares is None and is_num(row["cap_flottante"]) and is_num(row["cours"]) and row["cours"]:
                md.float_shares = float(row["cap_flottante"]) / float(row["cours"])
                md.sources["flottant"] = "brvm.org (capitalisation flottante / cours)"
            if is_num(row["cours"]) and (md.price is None or (md.price_date and md.price_date < date.today() - timedelta(days=4))):
                md.price = float(row["cours"])
                md.price_date = date.today()
                md.sources["cours"] = "brvm.org (cours du jour, plus récent que l'historique)"
    if company is not None and md.shares is None:
        nb = company.fields.get("nb_actions", {})
        if nb:
            md.shares = nb[max(nb)]
            md.sources["nombre de titres"] = getattr(company, "data_source", "Classeur des états financiers")

    # Dividendes ------------------------------------------------------------
    if base is not None and base.dividends.get(ticker):
        md.dividends = [Dividend(d.exercice, d.montant_net / split, d.date_paiement, d.source)
                        for d in base.dividends[ticker]]
        md.sources["dividendes"] = "Base de marché (Dividendes)" + (" ajustés du fractionnement" if split > 1 else "")
    if not md.dividends and online:
        divs = fetch_dividends_online(ticker, name, md.errors)
        if divs:
            md.dividends = divs
            md.sources["dividendes"] = divs[0].source
    if not md.dividends and company is not None and company.fields.get("dps"):
        dps = company.fields["dps"]
        origin = getattr(company, "data_source", "Classeur des états financiers")
        md.dividends = [Dividend(y, v, None, origin) for y, v in sorted(dps.items())]
        md.sources["dividendes"] = origin

    # Dernier exercice clos sans versement recensé : les sources (sikafinance, brvm.org) ne listent que les versements
    # effectifs ; passé la saison des assemblées (30 septembre N+1), l'absence de ligne vaut absence de dividende.
    ly = company.last_year() if company is not None and getattr(company, "years", None) else None
    if md.dividends and ly and not overrides.get("dps"):
        known = [d.exercice for d in md.dividends if d.exercice]
        from_base = md.sources.get("dividendes", "").startswith("Base de marché")
        # base de marché : les zéros y sont explicites ; l'absence du dernier exercice ne vaut absence de dividende que si
        # la saison de ce dernier exercice y est déjà renseignée pour d'autres titres
        season_in_base = (not from_base) or (base is not None and any(
            d_.exercice and d_.exercice >= ly for ds in base.dividends.values() for d_ in ds))
        if known and max(known) < ly and season_in_base and date.today() >= date(ly + 1, 9, 30):
            md.dividends.append(Dividend(ly, 0.0, None, f"Aucun versement recensé au titre de {ly}"))
            md.sources["dividendes"] = (md.sources.get("dividendes", "") + f" ; aucun versement recensé au titre de {ly}").lstrip(" ;")

    # Saisies manuelles (prioritaires) -------------------------------------------
    if is_num(overrides.get("price")):
        md.price = float(overrides["price"])
        md.price_date = overrides.get("price_date") or date.today()
        md.sources["cours"] = "Saisie manuelle"
    if is_num(overrides.get("shares")):
        md.shares = float(overrides["shares"])
        md.sources["nombre de titres"] = "Saisie manuelle"
    if is_num(overrides.get("float_pct")) and is_num(md.shares):
        md.float_shares = float(overrides["float_pct"]) * md.shares
        md.sources["flottant"] = "Saisie manuelle"
    if is_num(overrides.get("dps")):
        md.dividends.append(Dividend(int(overrides.get("dps_exercice") or date.today().year - 1),
                                     float(overrides["dps"]), None, "Saisie manuelle"))
        md.sources["dividendes"] = "Saisie manuelle"
    if md.history is None:
        md.errors.append("Aucun historique de cours disponible : l'analyse technique sera neutralisée.")
    return md
