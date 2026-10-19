"""Analyse technique du titre : tendance, momentum, oscillateurs, volatilité,
liquidité, supports et résistances.

Particularités BRVM prises en compte : cotation en continu mais liquidité
souvent faible (séances sans échange), variation quotidienne plafonnée à
+/-7,5%, plus haut / plus bas parfois absents (on retombe alors sur les
clôtures).
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date

import numpy as np
import pandas as pd

from .utils import clamp, is_num, scale


@dataclass
class Technical:
    available: bool
    last: float | None = None
    last_date: date | None = None
    indicators: dict = field(default_factory=dict)
    performance: dict = field(default_factory=dict)
    levels: dict = field(default_factory=dict)
    liquidity: dict = field(default_factory=dict)
    signals: list[tuple[str, str]] = field(default_factory=list)   # (libellé, "pos"|"neg"|"neu")
    trend: str = "indéterminée"
    score: float | None = None
    parts: list = field(default_factory=list)
    frame: pd.DataFrame | None = None


def _rsi(close: pd.Series, n: int = 14) -> pd.Series:
    d = close.diff()
    up = d.clip(lower=0).ewm(alpha=1 / n, adjust=False).mean()
    dn = (-d.clip(upper=0)).ewm(alpha=1 / n, adjust=False).mean()
    rs = up / dn.replace(0, np.nan)
    rsi = 100 - 100 / (1 + rs)
    rsi = rsi.where(dn > 0, 100.0)
    return rsi.where((up > 0) | (dn > 0), 50.0)


def _atr(df: pd.DataFrame, n: int = 14) -> pd.Series:
    h = df["high"].fillna(df["close"])
    lo = df["low"].fillna(df["close"])
    pc = df["close"].shift(1)
    tr = pd.concat([h - lo, (h - pc).abs(), (lo - pc).abs()], axis=1).max(axis=1)
    return tr.ewm(alpha=1 / n, adjust=False).mean()


def _perf(close: pd.Series, days: int) -> float | None:
    if len(close) < 2:
        return None
    end = close.index[-1]
    start = end - pd.Timedelta(days=days)
    past = close[close.index <= start]
    if past.empty:
        return None
    return float(close.iloc[-1] / past.iloc[-1] - 1)


def _ytd(close: pd.Series) -> float | None:
    end = close.index[-1]
    prev = close[close.index < pd.Timestamp(end.year, 1, 1)]
    if prev.empty:
        return None
    return float(close.iloc[-1] / prev.iloc[-1] - 1)


def _levels(close: pd.Series, price: float) -> dict:
    """Supports / résistances : extrêmes locaux regroupés (fenêtre 1 an)."""
    s = close[close.index >= close.index[-1] - pd.Timedelta(days=365)]
    if len(s) < 30:
        return {}
    w = 10
    piv_hi, piv_lo = [], []
    vals = s.values
    for i in range(w, len(vals) - w):
        seg = vals[i - w:i + w + 1]
        if vals[i] == seg.max():
            piv_hi.append(vals[i])
        if vals[i] == seg.min():
            piv_lo.append(vals[i])

    def cluster(points):
        pts = sorted(points)
        groups = []
        for p in pts:
            if groups and abs(p - np.mean(groups[-1])) / np.mean(groups[-1]) < 0.03:
                groups[-1].append(p)
            else:
                groups.append([p])
        return [(float(np.mean(g)), len(g)) for g in groups]

    supports = [lv for lv, n in cluster(piv_lo + piv_hi) if lv < price * 0.995]
    resist = [lv for lv, n in cluster(piv_hi + piv_lo) if lv > price * 1.005]
    out = {
        "plus_haut_52s": float(s.max()), "plus_bas_52s": float(s.min()),
        "support_1": max(supports) if supports else float(s.min()),
        "resistance_1": min(resist) if resist else float(s.max()),
    }
    lower = sorted([x for x in supports if x < out["support_1"] * 0.985], reverse=True)
    higher = sorted([x for x in resist if x > out["resistance_1"] * 1.015])
    out["support_2"] = lower[0] if lower else out["plus_bas_52s"]
    out["resistance_2"] = higher[0] if higher else out["plus_haut_52s"]
    return out


def analyse_technical(history: pd.DataFrame | None, index_history: pd.DataFrame | None = None,
                      float_shares: float | None = None) -> Technical:
    if history is None or history.empty or history["close"].dropna().shape[0] < 30:
        return Technical(available=False)
    df = history.copy().sort_index()
    df = df[~df.index.duplicated(keep="last")]
    df["close"] = df["close"].astype(float)
    close = df["close"].dropna()
    price = float(close.iloc[-1])
    t = Technical(available=True, last=price, last_date=close.index[-1].date())

    # Moyennes mobiles et indicateurs ------------------------------------
    for n in (20, 50, 100, 200):
        df[f"mm{n}"] = close.rolling(n, min_periods=int(n * 0.8)).mean()
    df["rsi"] = _rsi(close)
    ema12 = close.ewm(span=12, adjust=False).mean()
    ema26 = close.ewm(span=26, adjust=False).mean()
    df["macd"] = ema12 - ema26
    df["macd_signal"] = df["macd"].ewm(span=9, adjust=False).mean()
    df["macd_hist"] = df["macd"] - df["macd_signal"]
    mid = close.rolling(20).mean()
    sd = close.rolling(20).std()
    df["bb_mid"], df["bb_up"], df["bb_lo"] = mid, mid + 2 * sd, mid - 2 * sd
    df["atr"] = _atr(df)
    hh = df["high"].fillna(df["close"]).rolling(14).max()
    ll = df["low"].fillna(df["close"]).rolling(14).min()
    df["stoch_k"] = 100 * (close - ll) / (hh - ll).replace(0, np.nan)
    df["stoch_d"] = df["stoch_k"].rolling(3).mean()
    t.frame = df

    last = df.iloc[-1]
    ind = {k: (float(last[k]) if is_num(last.get(k)) else None)
           for k in ("mm20", "mm50", "mm100", "mm200", "rsi", "macd", "macd_signal", "macd_hist",
                     "bb_up", "bb_lo", "bb_mid", "atr", "stoch_k", "stoch_d")}
    rets = close.pct_change().dropna()
    r1y = rets[rets.index >= rets.index[-1] - pd.Timedelta(days=365)]
    ind["volatilite"] = float(r1y.std() * np.sqrt(252)) if len(r1y) > 20 else None
    cum = close[close.index >= close.index[-1] - pd.Timedelta(days=365)]
    ind["drawdown_max"] = float((cum / cum.cummax() - 1).min()) if len(cum) > 5 else None
    ind["atr_pct"] = ind["atr"] / price if is_num(ind["atr"]) else None
    ind["bb_position"] = ((price - ind["bb_lo"]) / (ind["bb_up"] - ind["bb_lo"])
                          if is_num(ind["bb_up"]) and is_num(ind["bb_lo"]) and ind["bb_up"] != ind["bb_lo"] else None)
    # Pente de la MM50 sur 20 séances
    mm50 = df["mm50"].dropna()
    ind["pente_mm50"] = float(mm50.iloc[-1] / mm50.iloc[-21] - 1) if len(mm50) > 21 else None
    # Croisements MM50 / MM200 récents
    cross = None
    if df["mm200"].notna().sum() > 30:
        diff = (df["mm50"] - df["mm200"]).dropna()
        sign = np.sign(diff)
        changes = sign.ne(sign.shift()).iloc[1:]
        if changes.any():
            last_cross = changes[changes].index[-1]
            cross = ("croisement haussier" if diff.loc[last_cross] > 0 else "croisement baissier",
                     last_cross.date())
    ind["dernier_croisement"] = cross

    # Bêta et force relative vs BRVM Composite ----------------------------
    if index_history is not None and not index_history.empty:
        idx = index_history["close"].astype(float).sort_index()
        both = pd.concat([close.rename("t"), idx.rename("i")], axis=1).dropna()
        both = both[both.index >= both.index[-1] - pd.Timedelta(days=730)]
        if len(both) > 60:
            wk = both.resample("W-FRI").last().pct_change().dropna()
            if len(wk) > 20 and wk["i"].var() > 0:
                beta = float(wk["t"].cov(wk["i"]) / wk["i"].var())
                ind["beta_brut"] = beta
                ind["beta"] = 0.67 * beta + 0.33  # ajustement de Blume
                ind["correlation"] = float(wk["t"].corr(wk["i"]))
        for lab, d in (("3m", 91), ("6m", 182), ("12m", 365)):
            pt, pi = _perf(close, d), _perf(idx, d)
            if is_num(pt) and is_num(pi):
                t.performance[f"relative_{lab}"] = (1 + pt) / (1 + pi) - 1
        t.performance["indice_ytd"] = _ytd(idx)
    t.indicators = ind

    # Performances --------------------------------------------------------
    for lab, d in (("1s", 7), ("1m", 30), ("3m", 91), ("6m", 182), ("12m", 365), ("3a", 1095)):
        t.performance[lab] = _perf(close, d)
    t.performance["ytd"] = _ytd(close)

    # Niveaux -------------------------------------------------------------
    t.levels = _levels(close, price)

    # Liquidité -----------------------------------------------------------
    vol = df["volume"] if "volume" in df else pd.Series(dtype=float)
    v3 = vol[vol.index >= vol.index[-1] - pd.Timedelta(days=91)] if len(vol) else vol
    if len(v3) and v3.notna().sum() > 5:
        traded = (v3.fillna(0) > 0).mean()
        avg_vol = float(v3.fillna(0).mean())
        t.liquidity = {
            "volume_moyen_3m": avg_vol,
            "valeur_moyenne_3m": avg_vol * price,
            "taux_seances_actives": float(traded),
        }
        if is_num(float_shares) and float_shares > 0:
            t.liquidity["rotation_flottant_annuelle"] = avg_vol * 252 / float_shares
    else:
        # Approximation : séances avec variation de cours
        chg = close[close.index >= close.index[-1] - pd.Timedelta(days=91)].diff().fillna(0)
        if len(chg) > 5:
            t.liquidity = {"taux_seances_actives": float((chg != 0).mean()), "approximation": True}

    # Tendance et score --------------------------------------------------------
    mm50v, mm200v, mm20v = ind["mm50"], ind["mm200"], ind["mm20"]
    trend_pts = []
    if is_num(mm200v):
        trend_pts.append(1 if price > mm200v else -1)
    if is_num(mm50v):
        trend_pts.append(1 if price > mm50v else -1)
    if is_num(mm50v) and is_num(mm200v):
        trend_pts.append(1 if mm50v > mm200v else -1)
    if is_num(ind["pente_mm50"]):
        trend_pts.append(1 if ind["pente_mm50"] > 0.005 else (-1 if ind["pente_mm50"] < -0.005 else 0))
    ts = np.mean(trend_pts) if trend_pts else 0
    t.trend = "haussière" if ts >= 0.5 else ("baissière" if ts <= -0.5 else "neutre")
    trend_score = (ts + 1) * 50 if trend_pts else None

    mom_parts = [scale(t.performance.get("3m"), -0.15, 0.20), scale(t.performance.get("6m"), -0.20, 0.30),
                 scale(t.performance.get("relative_6m"), -0.15, 0.15)]
    mom_parts = [m for m in mom_parts if is_num(m)]
    mom_score = float(np.mean(mom_parts)) if mom_parts else None

    rsi = ind["rsi"]
    osc = []
    if is_num(rsi):
        # zone idéale 45-65 ; survente = opportunité modérée ; surachat = risque de repli
        if rsi < 30:
            osc.append(60)
        elif rsi < 45:
            osc.append(45)
        elif rsi <= 65:
            osc.append(75)
        elif rsi <= 75:
            osc.append(55)
        else:
            osc.append(30)
    if is_num(ind["macd_hist"]) and is_num(ind["macd"]):
        osc.append(75 if ind["macd_hist"] > 0 and ind["macd"] > 0 else (60 if ind["macd_hist"] > 0 else
                                                                          (40 if ind["macd"] > 0 else 25)))
    osc_score = float(np.mean(osc)) if osc else None
    risk_score = None
    if is_num(ind["volatilite"]) or is_num(ind["drawdown_max"]):
        risk_score = np.nanmean([scale(ind["volatilite"], 0.10, 0.45, reverse=True) if is_num(ind["volatilite"]) else np.nan,
                                 scale(ind["drawdown_max"], -0.40, -0.05) if is_num(ind["drawdown_max"]) else np.nan])
    t.parts = [("Tendance (moyennes mobiles)", trend_score, 40), ("Momentum et force relative", mom_score, 30),
               ("Oscillateurs (RSI, MACD)", osc_score, 20), ("Risque (volatilité, repli max.)", risk_score, 10)]
    num = sum(s * w for _, s, w in t.parts if is_num(s))
    den = sum(w for _, s, w in t.parts if is_num(s))
    t.score = clamp(num / den) if den else None

    # Signaux lisibles ------------------------------------------------------
    sig = t.signals
    if is_num(mm200v):
        sig.append((f"Cours {'au-dessus' if price > mm200v else 'en dessous'} de la moyenne mobile 200 jours",
                    "pos" if price > mm200v else "neg"))
    if is_num(mm50v):
        sig.append((f"Cours {'au-dessus' if price > mm50v else 'en dessous'} de la moyenne mobile 50 jours",
                    "pos" if price > mm50v else "neg"))
    if cross:
        sig.append((f"Dernier {cross[0]} MM50/MM200 le {cross[1].strftime('%d/%m/%Y')}",
                    "pos" if "haussier" in cross[0] else "neg"))
    if is_num(rsi):
        if rsi >= 70:
            sig.append((f"RSI à {rsi:.0f} : zone de surachat, risque de consolidation", "neg"))
        elif rsi <= 30:
            sig.append((f"RSI à {rsi:.0f} : zone de survente, rebond technique possible", "pos"))
        else:
            sig.append((f"RSI à {rsi:.0f} : zone neutre", "neu"))
    if is_num(ind["macd_hist"]):
        sig.append(("MACD au-dessus de sa ligne de signal" if ind["macd_hist"] > 0 else
                    "MACD sous sa ligne de signal", "pos" if ind["macd_hist"] > 0 else "neg"))
    bp = ind["bb_position"]
    if is_num(bp):
        if bp > 1:
            sig.append(("Cours au-dessus de la bande de Bollinger supérieure (excès haussier)", "neg"))
        elif bp < 0:
            sig.append(("Cours sous la bande de Bollinger inférieure (excès baissier)", "pos"))
    return t
