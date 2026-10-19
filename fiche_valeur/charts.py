"""Graphiques (matplotlib, rendu PNG) partagés par la fiche PDF, le rapport
PowerPoint et l'interface.

Règles appliquées : une seule échelle par graphique (petits multiples plutôt
que double axe), traits fins, quadrillage discret, légende dès deux séries,
étiquettes sélectives, palette de séries validée pour les daltonismes.
"""
from __future__ import annotations

import io
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib import font_manager  # noqa: E402
from matplotlib.ticker import FuncFormatter, MaxNLocator  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402

from . import config  # noqa: E402
from .utils import fmt_num, is_num  # noqa: E402

# Palette des séries (vérifiée : bande de luminosité, chroma, séparation daltonisme)
S1, S2, S3, S4 = "#3A4AA3", "#A97A1F", "#1F8BBE", "#C2556A"
GRAY = "#B8BDC9"
INK = "#1F2430"
MUTED = "#6B7280"
GRID = "#E6E8EF"
SURFACE = "#FFFFFF"
WASH = "#3A4AA3"

_FONT_READY = False


def _setup_fonts() -> None:
    global _FONT_READY
    if _FONT_READY:
        return
    family = "DejaVu Sans"
    for f in Path(config.FONTS_DIR).glob("Carlito-*.ttf"):
        try:
            font_manager.fontManager.addfont(str(f))
            family = "Carlito"
        except Exception:
            pass
    plt.rcParams.update({
        "font.family": family, "font.size": 9, "axes.edgecolor": GRID, "axes.linewidth": 0.8,
        "axes.labelcolor": MUTED, "xtick.color": MUTED, "ytick.color": MUTED, "axes.titlesize": 10,
        "axes.titleweight": "bold", "axes.titlecolor": INK, "axes.titlelocation": "left",
        "axes.spines.top": False, "axes.spines.right": False, "legend.frameon": False,
        "legend.fontsize": 8, "figure.facecolor": SURFACE, "axes.facecolor": SURFACE,
        "axes.grid": True, "grid.color": GRID, "grid.linewidth": 0.6, "grid.linestyle": "-",
        "xtick.major.size": 0, "ytick.major.size": 0, "axes.axisbelow": True,
    })
    _FONT_READY = True


def _png(fig, dpi: int = 200) -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight", facecolor=SURFACE)
    plt.close(fig)
    return buf.getvalue()


def _fr(dec=0):
    return FuncFormatter(lambda x, _: fmt_num(x, dec))


def _pct_fmt(dec=0):
    return FuncFormatter(lambda x, _: fmt_num(x * 100, dec) + "%")


MOIS = ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc."]


def _date_axis(ax, idx):
    span_days = (idx[-1] - idx[0]).days if len(idx) > 1 else 0
    step = 3 if span_days > 400 else (2 if span_days > 200 else 1)
    ticks = pd.date_range(idx[0].normalize(), idx[-1], freq=f"{step}MS")
    ax.set_xticks(ticks)
    ax.set_xticklabels([f"{MOIS[t.month - 1]} {str(t.year)[2:]}" for t in ticks])


# ---------------------------------------------------------------------------
# Cours et indicateurs techniques
# ---------------------------------------------------------------------------

def price_chart(tech, months: int = 18, size=(7.2, 3.6), with_rsi: bool = False, levels: bool = True) -> bytes | None:
    _setup_fonts()
    if tech is None or not tech.available or tech.frame is None:
        return None
    df = tech.frame
    df = df[df.index >= df.index[-1] - pd.DateOffset(months=months)]
    has_vol = df["volume"].notna().sum() > 10 and df["volume"].fillna(0).sum() > 0
    rows = [3] + ([1] if has_vol else []) + ([1.1] if with_rsi else [])
    fig, axes = plt.subplots(len(rows), 1, figsize=size, sharex=True,
                             gridspec_kw={"height_ratios": rows, "hspace": 0.12})
    axes = np.atleast_1d(axes)
    ax = axes[0]
    if df["bb_up"].notna().any():
        ax.fill_between(df.index, df["bb_lo"], df["bb_up"], color=WASH, alpha=0.08, linewidth=0,
                        label="Bandes de Bollinger (20 j, 2 écarts-types)")
    ax.plot(df.index, df["close"], color=S1, lw=1.6, label="Cours de clôture", solid_capstyle="round")
    if df["mm50"].notna().any():
        ax.plot(df.index, df["mm50"], color=S2, lw=1.2, label="Moyenne mobile 50 j")
    if df["mm200"].notna().any():
        ax.plot(df.index, df["mm200"], color=S3, lw=1.2, label="Moyenne mobile 200 j")
    last = df["close"].iloc[-1]
    ax.scatter([df.index[-1]], [last], s=28, color=S1, zorder=5, edgecolor=SURFACE, linewidth=1.5)
    ax.annotate(fmt_num(last), (df.index[-1], last), xytext=(6, 0), textcoords="offset points",
                va="center", fontsize=8.5, color=INK, fontweight="bold")
    if levels and tech.levels:
        for key, lab, va, dy in (("support_1", "Support", "top", -3), ("resistance_1", "Résistance", "bottom", 3)):
            lv = tech.levels.get(key)
            if is_num(lv) and df["close"].min() * 0.9 < lv < df["close"].max() * 1.1:
                ax.axhline(lv, color=MUTED, lw=0.7, alpha=0.7)
                ax.annotate(f"{lab} {fmt_num(lv)}", (df.index[len(df) // 2], lv), xytext=(0, dy),
                            textcoords="offset points", fontsize=7, color=MUTED, va=va, ha="center",
                            bbox=dict(boxstyle="square,pad=0.1", fc=SURFACE, ec="none", alpha=0.85))
    ax.yaxis.set_major_formatter(_fr())
    ax.yaxis.set_major_locator(MaxNLocator(5))
    ax.legend(loc="lower left", bbox_to_anchor=(0, 1.0), ncol=4, fontsize=7.5, handlelength=1.6,
              columnspacing=1.2, borderaxespad=0.2)
    ax.set_ylabel("FCFA")
    i = 1
    if has_vol:
        axv = axes[i]
        axv.bar(df.index, df["volume"].fillna(0), width=1.0, color=GRAY, linewidth=0)
        axv.set_ylabel("Titres")
        axv.yaxis.set_major_formatter(FuncFormatter(lambda x, _: fmt_num(x / 1000, 0) + "k" if x >= 1000 else fmt_num(x)))
        axv.yaxis.set_major_locator(MaxNLocator(3))
        i += 1
    if with_rsi:
        axr = axes[i]
        axr.plot(df.index, df["rsi"], color=S1, lw=1.1)
        axr.axhspan(70, 100, color=S4, alpha=0.07, lw=0)
        axr.axhspan(0, 30, color=S3, alpha=0.07, lw=0)
        axr.set_ylim(0, 100)
        axr.set_yticks([30, 50, 70])
        axr.set_ylabel("RSI 14")
    _date_axis(axes[-1], df.index)
    return _png(fig)


def macd_chart(tech, months: int = 12, size=(7.2, 1.8)) -> bytes | None:
    _setup_fonts()
    if tech is None or not tech.available:
        return None
    df = tech.frame
    df = df[df.index >= df.index[-1] - pd.DateOffset(months=months)]
    fig, ax = plt.subplots(figsize=size)
    hist = df["macd_hist"]
    ax.bar(df.index, hist, width=1.0, color=np.where(hist >= 0, S3, S4), linewidth=0, alpha=0.6,
           label="Histogramme")
    ax.plot(df.index, df["macd"], color=S1, lw=1.2, label="MACD (12, 26)")
    ax.plot(df.index, df["macd_signal"], color=S2, lw=1.2, label="Signal (9)")
    ax.axhline(0, color=MUTED, lw=0.6)
    ax.legend(loc="upper left", ncol=3, fontsize=7.5)
    ax.yaxis.set_major_formatter(_fr())
    ax.yaxis.set_major_locator(MaxNLocator(4))
    _date_axis(ax, df.index)
    return _png(fig)


# ---------------------------------------------------------------------------
# Fondamentaux
# ---------------------------------------------------------------------------

def financials_chart(c, ratios, size=(7.2, 2.6)) -> bytes | None:
    """Petits multiples : revenus et résultat net (Mds FCFA) | marges et ROE (%)."""
    _setup_fonts()
    rev = c.fields.get(c.revenue_field, {})
    rn = c.fields.get(c.net_income_field, {})
    years = [y for y in c.years if y in rev or y in rn][-5:]
    if not years:
        return None
    fig, (a1, a2) = plt.subplots(1, 2, figsize=size, gridspec_kw={"width_ratios": [1.25, 1], "wspace": 0.28})
    x = np.arange(len(years))
    w = 0.36
    rv = [rev.get(y, np.nan) / 1e9 for y in years]
    nv = [rn.get(y, np.nan) / 1e9 for y in years]
    a1.bar(x - w / 2 - 0.01, rv, w, color=S1, label="PNB" if c.is_bank else "Chiffre d'affaires", linewidth=0)
    a1.bar(x + w / 2 + 0.01, nv, w, color=S2, label="Résultat net" + (" part du groupe" if c.net_income_field == "rn_pg" else ""), linewidth=0)
    a1.set_xticks(x)
    a1.set_xticklabels([str(y) for y in years])
    a1.yaxis.set_major_formatter(_fr())
    a1.yaxis.set_major_locator(MaxNLocator(4))
    a1.set_title("Activité et résultats (Mds FCFA)", pad=16)
    a1.legend(loc="lower left", bbox_to_anchor=(0, 1.0), ncol=2, fontsize=7.5, borderaxespad=0.1)
    a1.axhline(0, color=MUTED, lw=0.6)
    if is_num(rv[-1]):
        a1.annotate(fmt_num(rv[-1], 1), (x[-1] - w / 2, rv[-1]), xytext=(0, 3), textcoords="offset points",
                    ha="center", fontsize=7.5, color=INK)
    series = [("roe", "ROE", S1), ("marge_nette", "Marge nette", S2)]
    series.append(("coef_exploitation", "Coefficient d'exploitation", S3) if c.is_bank else ("marge_ebe", "Marge d'EBE", S3))
    plotted = False
    for key, lab, col in series:
        s = ratios.get(key, {})
        ys = [y for y in years if is_num(s.get(y))]
        if len(ys) >= 1:
            a2.plot([years.index(y) for y in ys], [s[y] for y in ys], color=col, lw=1.6, marker="o", ms=4,
                    markeredgecolor=SURFACE, markeredgewidth=1.2, label=lab)
            plotted = True
    a2.set_xticks(x)
    a2.set_xticklabels([str(y) for y in years])
    a2.yaxis.set_major_formatter(_pct_fmt())
    a2.yaxis.set_major_locator(MaxNLocator(4))
    a2.set_title("Rentabilité (%)", pad=16)
    if plotted:
        a2.legend(loc="lower left", bbox_to_anchor=(0, 1.0), ncol=3, fontsize=7.5, borderaxespad=0.1,
                  handlelength=1.4, columnspacing=0.8)
    return _png(fig)


def peers_chart(fd, c, val, size=(7.2, 2.7)) -> bytes | None:
    """ROE et PER : la société mise en évidence parmi ses pairs."""
    _setup_fonts()
    y = fd.year
    rows = [(p.ticker, p.roe, p.pe) for p in fd.peers]
    rows.append((c.ticker, fd.ratios.get("roe", {}).get(y), val.pe))
    roe_rows = sorted([(t, r) for t, r, _ in rows if is_num(r) and -0.5 < r < 1.0], key=lambda z: z[1])
    pe_rows = sorted([(t, p) for t, _, p in rows if is_num(p) and 0 < p < 60], key=lambda z: z[1])
    if len(roe_rows) < 2:
        return None
    panels = [("ROE du dernier exercice", roe_rows, _pct_fmt(), lambda v: fmt_num(v * 100, 1) + "%")]
    if len(pe_rows) >= 2:
        panels.append(("PER (cours actuel)", pe_rows, _fr(), lambda v: fmt_num(v, 1) + "x"))
    n = max(len(roe_rows), len(pe_rows))
    h = min(max(1.6, 0.24 * n + 0.9), 4.2)
    fig, axes = plt.subplots(1, len(panels), figsize=(size[0], h), gridspec_kw={"wspace": 0.45})
    axes = np.atleast_1d(axes)
    for ax, (title, data, fmt, lab) in zip(axes, panels):
        names = [t for t, _ in data]
        vals = [v for _, v in data]
        cols = [S1 if t == c.ticker else GRAY for t in names]
        ax.barh(range(len(vals)), vals, color=cols, height=0.55 if len(vals) > 3 else 0.35, linewidth=0)
        ax.set_yticks(range(len(vals)))
        ax.set_yticklabels(names, fontsize=7.5)
        for tl in ax.get_yticklabels():
            if tl.get_text() == c.ticker:
                tl.set_fontweight("bold")
                tl.set_color(INK)
        ax.xaxis.set_major_formatter(fmt)
        ax.xaxis.set_major_locator(MaxNLocator(4))
        ax.grid(axis="y", visible=False)
        ax.set_title(title, pad=14)
        med = np.median(vals)
        ax.axvline(med, color=MUTED, lw=0.8)
        ax.text(med, 1.01, f"médiane {lab(med)}", transform=ax.get_xaxis_transform(), fontsize=7,
                color=MUTED, ha="center", va="bottom")
        i = names.index(c.ticker) if c.ticker in names else None
        if i is not None:
            ax.annotate(lab(vals[i]), (vals[i], i), xytext=(4, 0), textcoords="offset points", va="center",
                        fontsize=7.5, color=INK, fontweight="bold")
    return _png(fig)


# ---------------------------------------------------------------------------
# Valorisation et synthèse
# ---------------------------------------------------------------------------

def valuation_chart(val, tech=None, size=(7.2, 2.6)) -> bytes | None:
    """Méthodes de valorisation (points) face au cours et à la juste valeur."""
    _setup_fonts()
    if not val.methods or not is_num(val.price):
        return None
    items = [(m.name, m.value) for m in val.methods]
    fig, ax = plt.subplots(figsize=(size[0], max(size[1], 0.42 * (len(items) + 1) + 0.6)))
    ys = np.arange(len(items) + (1 if tech is not None and tech.available and tech.levels else 0))
    labels = []
    k = 0
    if tech is not None and tech.available and tech.levels.get("plus_bas_52s"):
        lo, hi = tech.levels["plus_bas_52s"], tech.levels["plus_haut_52s"]
        ax.plot([lo, hi], [k, k], color=GRAY, lw=6, solid_capstyle="round")
        labels.append("Fourchette 52 semaines")
        k += 1
    for name, v in items:
        ax.scatter([v], [k], s=46, color=S1, zorder=4, edgecolor=SURFACE, linewidth=1.5)
        ax.annotate(fmt_num(v), (v, k), xytext=(7, 0), textcoords="offset points", va="center", fontsize=7.5,
                    color=INK)
        labels.append(name)
        k += 1
    ax.axvline(val.price, color=MUTED, lw=1.0)
    ax.annotate(f"Cours {fmt_num(val.price)}", (val.price, k - 0.45), xytext=(3, 0), textcoords="offset points",
                fontsize=7.5, color=MUTED)
    if is_num(val.fair_value):
        ax.axvline(val.fair_value, color=S2, lw=1.4)
        ax.annotate(f"Juste valeur {fmt_num(val.fair_value)}", (val.fair_value, -0.6), xytext=(3, 0),
                    textcoords="offset points", fontsize=7.5, color=INK, fontweight="bold")
    ax.set_yticks(range(len(labels)))
    ax.set_yticklabels(labels)
    ax.set_ylim(-0.9, len(labels) - 0.2)
    ax.invert_yaxis()
    ax.xaxis.set_major_formatter(_fr())
    ax.grid(axis="y", visible=False)
    ax.set_title("Valorisation par méthode (FCFA par action)")
    return _png(fig)


def pillars_chart(pillars: dict, size=(3.6, 2.3)) -> bytes:
    _setup_fonts()
    keys = list(config.PILLAR_LABELS)
    labels = [config.PILLAR_LABELS[k] for k in keys]
    vals = [pillars.get(k) for k in keys]
    fig, ax = plt.subplots(figsize=size)
    yy = np.arange(len(keys))
    ax.barh(yy, [100] * len(keys), color="#EEF0F5", height=0.55, linewidth=0)
    ax.barh(yy, [v if is_num(v) else 0 for v in vals], color=S1, height=0.55, linewidth=0)
    for i, v in enumerate(vals):
        ax.annotate("n.d." if not is_num(v) else fmt_num(v), (min(v or 0, 100) if is_num(v) else 0, i),
                    xytext=(4, 0), textcoords="offset points", va="center", fontsize=7.5,
                    color=INK if is_num(v) else MUTED)
    ax.axvline(50, color=MUTED, lw=0.7)
    ax.set_yticks(yy)
    ax.set_yticklabels(labels, fontsize=7.5)
    ax.invert_yaxis()
    ax.set_xlim(0, 110)
    ax.set_xticks([0, 50, 100])
    ax.grid(False)
    ax.spines["left"].set_visible(False)
    ax.spines["bottom"].set_visible(False)
    return _png(fig)


def decision_matrix_chart(decision, st, size=(4.6, 3.2)) -> bytes | None:
    """Vue fondamentale (axe vertical) croisée avec le potentiel (axe horizontal)."""
    _setup_fonts()
    F, U = decision.fundamental, decision.upside
    if not is_num(F):
        return None
    xs = st.seuil_potentiel_vente / 100, st.seuil_potentiel_achat / 100
    ys = st.seuil_fondamental_faible, st.seuil_fondamental_fort
    xmin, xmax = min(-0.6, (U or 0) - 0.1), max(0.6, (U or 0) + 0.1)
    fig, ax = plt.subplots(figsize=size)
    zones = {
        ("défavorable", "surcote"): "VENDRE", ("défavorable", "proche"): "ALLÉGER", ("défavorable", "décote"): "ATTENDRE",
        ("neutre", "surcote"): "ALLÉGER", ("neutre", "proche"): "CONSERVER", ("neutre", "décote"): "ACHAT / ATTENDRE",
        ("favorable", "surcote"): "ALLÉGER", ("favorable", "proche"): "CONSERVER", ("favorable", "décote"): "ACHAT",
    }
    xb = {"surcote": (xmin, xs[0]), "proche": (xs[0], xs[1]), "décote": (xs[1], xmax)}
    yb = {"défavorable": (0, ys[0]), "neutre": (ys[0], ys[1]), "favorable": (ys[1], 100)}
    for (fz, vz), lab in zones.items():
        x0, x1 = xb[vz]
        y0, y1 = yb[fz]
        base = lab.split(" ")[0]
        col = config.SIGNAL_COLORS.get(base, GRAY)
        ax.add_patch(plt.Rectangle((x0, y0), x1 - x0, y1 - y0, color=col, alpha=0.10, lw=0))
        ax.text(x0 + (x1 - x0) * 0.04, y1 - 2.5, lab, ha="left", va="top", fontsize=6.5, color=MUTED)
    for x in xs:
        ax.axvline(x, color=SURFACE, lw=2)
    for y in ys:
        ax.axhline(y, color=SURFACE, lw=2)
    if is_num(U):
        ax.scatter([U], [F], s=80, color=S1, zorder=5, edgecolor=SURFACE, linewidth=2)
        ax.annotate(decision.signal, (U, F), xytext=(0, -14), textcoords="offset points", fontsize=8.5,
                    fontweight="bold", color=INK, ha="center", va="top")
    ax.set_xlim(xmin, xmax)
    ax.set_ylim(0, 100)
    ax.xaxis.set_major_formatter(_pct_fmt())
    ax.set_xlabel("Potentiel vers la juste valeur")
    ax.set_ylabel("Vue fondamentale (/100)")
    ax.grid(False)
    return _png(fig)
