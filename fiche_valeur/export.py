"""Exports Excel : détail d'une analyse et tableau de bord multi-valeurs."""
from __future__ import annotations

import io

import pandas as pd

from . import config
from .fundamentals import RATIO_META, sig_rows
from .utils import clean_text, is_num


def analysis_frames(a) -> dict[str, pd.DataFrame]:
    c, v, fd, t, d = a.company, a.valuation, a.fundamentals, a.technical, a.decision
    years = c.years[-5:]
    synth = pd.DataFrame([
        ("Société", c.name), ("Ticker", c.ticker), ("Secteur", c.sector),
        ("Pays", config.COUNTRIES.get(c.country, {}).get("nom", c.country)), ("Date d'analyse", a.as_of),
        ("Signal", d.signal), ("Conviction", d.conviction), ("Note globale /100", d.composite),
        ("Vue fondamentale /100", d.fundamental), ("Score technique /100", d.technical),
        ("Cours (FCFA)", v.price), ("Juste valeur (FCFA)", v.fair_value), ("Potentiel", v.upside),
        ("Rendement net", v.div_yield), ("Rentabilité totale attendue", v.total_return),
        ("Capitalisation (FCFA)", a.market.market_cap), ("Nombre de titres", a.market.shares),
        ("PER", v.pe), ("P/B", v.pb), ("VE/EBE", v.ev_ebitda), ("BPA (FCFA)", v.eps), ("DPA net (FCFA)", v.dps),
        ("Coût des fonds propres", v.cost_of_equity), ("Bêta", v.beta), ("Croissance perpétuelle", v.g),
        ("Zone d'accumulation", f"{d.entry_zone[0]:.0f} - {d.entry_zone[1]:.0f}" if d.entry_zone else None),
        ("Niveau d'invalidation (FCFA)", d.stop),
    ] + [(f"Pilier : {config.PILLAR_LABELS[k]}", d.pillars.get(k)) for k in config.PILLAR_LABELS],
        columns=["Indicateur", "Valeur"])
    sig = pd.DataFrame({lab: [s.get(y) for y in years] for _, lab, s in sig_rows(c)}, index=years).T
    sig.index.name = "Solde (FCFA)"
    ratios = pd.DataFrame({RATIO_META[k][0]: [s.get(y) for y in years] for k, s in fd.ratios.items()
                           if k in RATIO_META}, index=years).T
    ratios.index.name = "Ratio"
    val = pd.DataFrame([(m.name, m.value, m.weight, clean_text(m.detail)) for m in v.methods],
                       columns=["Méthode", "Valeur (FCFA)", "Poids", "Hypothèse"])
    tech = pd.DataFrame(
        [(k, vv) for k, vv in t.indicators.items() if is_num(vv)] +
        [(f"performance {k}", vv) for k, vv in t.performance.items() if is_num(vv)] +
        [(k, vv) for k, vv in t.levels.items() if is_num(vv)], columns=["Indicateur", "Valeur"]) if t.available else pd.DataFrame()
    peers = pd.DataFrame([(p.ticker, p.name, p.sector, p.roe, p.net_margin, p.pe, p.pb, p.div_yield)
                          for p in fd.peers], columns=["Ticker", "Société", "Secteur", "ROE", "Marge nette", "PER",
                                                       "P/B", "Rendement"])
    hist = t.frame[["close", "volume", "mm50", "mm200", "rsi"]].copy() if t.available else pd.DataFrame()
    if not hist.empty:
        hist.index.name = "Date"
    notes = pd.DataFrame({"Contrôles et accès aux données": c.notes + a.errors + d.warnings})
    return {"Synthèse": synth, "SIG": sig, "Ratios": ratios, "Valorisation": val, "Technique": tech,
            "Pairs": peers, "Historique": hist, "Contrôles": notes}


def build_excel(a) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for name, df in analysis_frames(a).items():
            if df is None or df.empty:
                continue
            df.to_excel(xw, sheet_name=name[:31], index=name in ("SIG", "Ratios", "Historique"))
            ws = xw.sheets[name[:31]]
            for col in ws.columns:
                width = max(len(str(cell.value or "")) for cell in col[:60])
                ws.column_dimensions[col[0].column_letter].width = min(60, max(10, width + 2))
    return buf.getvalue()


def screening_frame(analyses: list) -> pd.DataFrame:
    rows = []
    for a in analyses:
        d, v = a.decision, a.valuation
        rows.append({
            "Ticker": a.company.ticker, "Société": a.company.name, "Secteur": a.company.sector,
            "Signal": d.signal, "Conviction": d.conviction, "Note globale": d.composite,
            "Vue fondamentale": d.fundamental, "Technique": d.technical, "Cours": v.price,
            "Juste valeur": v.fair_value, "Potentiel": v.upside, "Rendement": v.div_yield, "PER": v.pe,
            "P/B": v.pb, "ROE": a.fundamentals.ratios.get("roe", {}).get(a.fundamentals.year),
            "Tendance": a.technical.trend if a.technical.available else "n.d.",
        })
    df = pd.DataFrame(rows)
    order = {"ACHAT": 0, "ATTENDRE": 1, "CONSERVER": 2, "ALLÉGER": 3, "VENDRE": 4}
    if not df.empty:
        df["_o"] = df["Signal"].map(order)
        df = df.sort_values(["_o", "Note globale"], ascending=[True, False]).drop(columns="_o")
    return df


def build_screening_excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="Signaux", index=False)
        ws = xw.sheets["Signaux"]
        for col in ws.columns:
            width = max(len(str(cell.value or "")) for cell in col[:80])
            ws.column_dimensions[col[0].column_letter].width = min(45, max(9, width + 2))
        pct_cols = {"Potentiel", "Rendement", "ROE"}
        for j, name in enumerate(df.columns, start=1):
            for i in range(2, len(df) + 2):
                cell = ws.cell(row=i, column=j)
                if name in pct_cols:
                    cell.number_format = "0.0%"
                elif name in ("PER", "P/B"):
                    cell.number_format = "0.0"
                elif name in ("Cours", "Juste valeur", "Note globale", "Vue fondamentale", "Technique"):
                    cell.number_format = "#,##0"
    return buf.getvalue()
