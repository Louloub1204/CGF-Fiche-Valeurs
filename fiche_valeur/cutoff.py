"""Date d'arrêté : restreint toutes les données chargées à ce qui était connu à une date donnée.

Règles :
- états financiers annuels : exercices clos au plus tard à la date d'arrêté (au 31/10/2026, exercices
  jusqu'au 31/12/2025 ; au 31/12/2026, l'exercice 2026 est inclus) ;
- publications intermédiaires : celles publiées au plus tard à la date d'arrêté (date de publication si
  elle est connue, sinon date de fin de période : 31/03, 30/06, 30/09) ;
- cours, volumes et indice : séances jusqu'à la date d'arrêté ;
- dividendes : versements payés au plus tard à la date d'arrêté (sans date de paiement, un dividende
  de l'exercice N est réputé payé le 1er juillet N+1).
"""
from __future__ import annotations

import copy
from datetime import date, timedelta

import pandas as pd

PERIOD_END = {"T1": (3, 31), "T2": (6, 30), "S1": (6, 30), "T3": (9, 30), "T4": (12, 31), "S2": (12, 31)}


def last_closed_exercise(as_of: date) -> int:
    """Dernier exercice clos (exercice calendaire) à la date d'arrêté."""
    return as_of.year if (as_of.month, as_of.day) == (12, 31) else as_of.year - 1


def period_end(year: int, period: str) -> date:
    m, d = PERIOD_END.get(period, (12, 31))
    return date(year, m, d)


def is_historical(as_of: date | None) -> bool:
    """Date d'arrêté antérieure aux données du jour (les sources en ligne « du jour » sont alors écartées)."""
    return as_of is not None and as_of < date.today() - timedelta(days=4)


def dividend_known(d, as_of: date) -> bool:
    if d.date_paiement is not None:
        return d.date_paiement <= as_of
    if d.exercice:
        return date(int(d.exercice) + 1, 7, 1) <= as_of
    return True


def _cut_frame(df: pd.DataFrame | None, as_of: date) -> pd.DataFrame | None:
    if df is None or df.empty:
        return df
    out = df[df.index <= pd.Timestamp(as_of)]
    return out if not out.empty else None


def cut_history(df: pd.DataFrame | None, as_of: date) -> pd.DataFrame | None:
    return _cut_frame(df, as_of)


def cut_company(c, as_of: date):
    """Société restreinte aux exercices clos à la date d'arrêté (modifiée sur place et renvoyée)."""
    ly = last_closed_exercise(as_of)
    if not c.years or max(c.years) <= ly:
        return c
    c.years = [y for y in c.years if y <= ly]
    c.fields = {k: {y: v for y, v in s.items() if y <= ly} for k, s in c.fields.items()}
    for ln in c.lines:
        ln.values = {y: v for y, v in ln.values.items() if y <= ly}
    c.estimated = {e for e in c.estimated if not (isinstance(e, tuple) and len(e) > 1 and isinstance(e[1], int)
                                                   and e[1] > ly)}
    c.notes.append(f"Date d'arrêté : exercices postérieurs à {ly} écartés.")
    return c


def cut_workbook(wb, as_of: date):
    if wb is None:
        return None
    out = copy.deepcopy(wb)
    ly = last_closed_exercise(as_of)
    for t in list(out.companies):
        c = cut_company(out.companies[t], as_of)
        if not c.years:
            del out.companies[t]          # aucun exercice clos à cette date : traité hors classeur
    out.notes = list(out.notes) + [f"Date d'arrêté du {as_of:%d/%m/%Y} : exercices jusqu'à {ly}."]
    return out


def cut_market_base(mb, as_of: date):
    if mb is None:
        return None
    out = copy.copy(mb)
    ly = last_closed_exercise(as_of)
    out.closes = _cut_frame(mb.closes, as_of)
    out.volumes = _cut_frame(mb.volumes, as_of)
    out.avg_prices = {t: {y: v for y, v in s.items() if y <= as_of.year} for t, s in mb.avg_prices.items()}
    out.dividends = {t: [d for d in ds if dividend_known(d, as_of)] for t, ds in mb.dividends.items()}
    out.fundamentals = {t: {k: {y: v for y, v in s.items() if y <= ly} for k, s in f.items()}
                        for t, f in mb.fundamentals.items()}
    out.notes = list(mb.notes)
    return out


def interim_known(year: int, period: str, row, as_of: date) -> bool:
    end = period_end(year, period)
    if end > as_of:
        return False
    pub = getattr(row, "published", None)
    return pub is None or pub <= as_of


def cut_interim(data, as_of: date):
    if data is None:
        return None
    out = copy.copy(data)
    out.rows = {t: {k: r for k, r in rows.items() if interim_known(k[0], k[1], r, as_of)}
                for t, rows in data.rows.items()}
    out.n_rows = sum(len(r) for r in out.rows.values())
    return out


def report_known(rep, as_of: date) -> bool:
    """Publication en ligne (InterimReport) connue à la date d'arrêté."""
    if rep is None:
        return False
    if rep.published is not None and rep.published > as_of:
        return False
    import re
    m = re.search(r"([TS]\d|Exercice)\s*(\d{4})", rep.period or "")
    if m:
        y = int(m.group(2))
        end = date(y, 12, 31) if m.group(1) == "Exercice" else period_end(y, m.group(1))
        if end > as_of or (m.group(1) == "Exercice" and y > last_closed_exercise(as_of)):
            return False
    return True


def default_as_of(market_base=None, interim=None, price_file=None) -> date:
    """Dernière date à jour dans les fichiers chargés (base de marché, historique, publications) ;
    à défaut, la date du jour."""
    cands: list[date] = []
    for df in (getattr(market_base, "closes", None), price_file):
        if df is not None and not df.empty:
            cands.append(pd.Timestamp(df.index.max()).date())
    if interim is not None:
        for t, rows in interim.rows.items():
            for (y, p), r in rows.items():
                cands.append(r.published or period_end(y, p))
    today = date.today()
    cands = [d for d in cands if d <= today]
    return max(cands) if cands else today
