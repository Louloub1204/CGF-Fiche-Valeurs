"""Dynamique de l'exercice en cours : publications intermédiaires (T1, S1, T3).

Les sociétés cotées publient, en cours d'exercice, des indicateurs trimestriels (T1, T3) et des états
semestriels (S1). Chaque publication donne la période de l'exercice N et la même période de N-1. À partir de
ces chiffres, ce module construit :

* le tableau « période N / période N-1 » avec variations, pour chaque période publiée de l'exercice en cours ;
* les **12 mois glissants** : exercice N-1 + cumul N - cumul N-1 (neutralise la saisonnalité) ;
* une estimation de l'exercice en cours (exercice N-1 prolongé de la croissance cumulée) ;
* la tendance d'une publication à l'autre (accélération ou ralentissement).

Sources, par priorité : fichier des publications intermédiaires tenu par CGF Gestion > document chargé ou
lu en ligne (richbourse.com). Montants en FCFA, cumulés depuis le 1er janvier (T1 = 3 mois, S1 = 6 mois,
T3 = 9 mois).
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date

from .utils import is_num

PERIOD_ORDER = {"T1": 1, "S1": 2, "T3": 3}
PERIOD_LABEL = {"T1": "1er trimestre", "S1": "1er semestre", "T3": "3e trimestre (9 mois)"}
PERIOD_MONTHS = {"T1": 3, "S1": 6, "T3": 9}

# (clé, libellé, sens : +1 = hausse favorable, -1 = hausse défavorable)
ENTREPRISE_KEYS = [("ca", "Chiffre d'affaires", 1), ("va", "Valeur ajoutée", 1), ("ebe", "EBE", 1),
                   ("re", "Résultat d'exploitation", 1), ("rao", "Résultat des activités ordinaires", 1),
                   ("rn", "Résultat net", 1)]
BANQUE_KEYS = [("pnb", "Produit net bancaire", 1), ("charges_generales", "Frais de gestion", -1),
               ("rbe", "Résultat brut d'exploitation", 1), ("cout_risque", "Coût du risque", -1),
               ("rai", "Résultat avant impôt", 1), ("rn", "Résultat net", 1),
               ("credits_clientele", "Crédits à la clientèle", 1), ("depots_clientele", "Dépôts de la clientèle", 1)]
STOCK_KEYS = {"credits_clientele", "depots_clientele"}          # encours : pas de glissement sur 12 mois
LTM_KEYS = ("ca", "pnb", "ebe", "re", "rbe", "rn")


def keys_for(is_bank: bool):
    return BANQUE_KEYS if is_bank else ENTREPRISE_KEYS


def normalise_period(raw) -> str | None:
    """« T1 », « S1 », « T3 », « 1er trimestre », « 1er semestre », « 3e trimestre », « 9 mois » -> code."""
    if raw is None:
        return None
    t = re.sub(r"[^A-Z0-9]", "", str(raw).upper().replace("É", "E"))
    if t in PERIOD_ORDER:
        return t
    if t in ("TRIM1", "1ERTRIMESTRE", "1TRIMESTRE", "Q1", "3MOIS"):
        return "T1"
    if t in ("SEM1", "1ERSEMESTRE", "1SEMESTRE", "H1", "6MOIS", "SEMESTRE1"):
        return "S1"
    if t in ("TRIM3", "3ETRIMESTRE", "3EMETRIMESTRE", "3TRIMESTRE", "Q3", "9MOIS"):
        return "T3"
    return None


@dataclass
class PeriodRow:
    values: dict[str, float] = field(default_factory=dict)         # clé -> FCFA, cumul depuis le 1er janvier
    statut: str = ""
    published: date | None = None
    source: str = ""
    origin: str = ""                                               # "fichier" | "en ligne" | "document"


@dataclass
class InterimData:
    """Contenu du fichier des publications intermédiaires : ticker -> (exercice, période) -> ligne."""
    rows: dict[str, dict[tuple[int, str], PeriodRow]] = field(default_factory=dict)
    notes: list[str] = field(default_factory=list)
    unknown: list[str] = field(default_factory=list)
    n_rows: int = 0

    @property
    def tickers(self) -> list[str]:
        return sorted(t for t, r in self.rows.items() if r)

    def periods(self, ticker: str) -> list[tuple[int, str]]:
        return sorted(self.rows.get(ticker, {}), key=lambda k: (k[0], PERIOD_ORDER[k[1]]))

    def latest(self, ticker: str) -> tuple[int, str] | None:
        p = self.periods(ticker)
        return p[-1] if p else None


@dataclass
class Compare:
    year: int
    period: str
    metrics: dict[str, tuple[float | None, float | None, float | None]] = field(default_factory=dict)  # cur, prev, var
    published: date | None = None
    source: str = ""
    statut: str = ""
    origin: str = ""

    @property
    def label(self) -> str:
        return f"{self.period} {self.year}"


@dataclass
class Dynamics:
    ticker: str
    is_bank: bool
    year: int
    period: str
    compares: list[Compare] = field(default_factory=list)
    ltm: dict[str, float] = field(default_factory=dict)             # 12 mois glissants à la dernière période
    ltm_base: dict[str, float] = field(default_factory=dict)        # exercice N-1 utilisé
    fy_est: dict[str, float] = field(default_factory=dict)          # exercice N estimé
    trend: str = ""
    trend_delta: float | None = None
    notes: list[str] = field(default_factory=list)

    @property
    def latest(self) -> Compare:
        return self.compares[-1]

    @property
    def label(self) -> str:
        return f"{self.period} {self.year}"

    @property
    def source(self) -> str:
        c = self.latest
        parts = {"fichier": "fichier des publications intermédiaires", "en ligne": "richbourse.com (publication officielle)",
                 "document": "document chargé"}
        return parts.get(c.origin, c.source or "publication intermédiaire")

    def var(self, key: str) -> float | None:
        m = self.latest.metrics.get(key)
        return m[2] if m else None


def variation(cur, prev) -> float | None:
    """Variation relative ; ``None`` si le signe change ou si la base est nulle."""
    if not is_num(cur) or not is_num(prev) or prev == 0:
        return None
    if (cur > 0) != (prev > 0):
        return None
    return abs(cur) / abs(prev) - 1


# ---------------------------------------------------------------------------
# Fusion des sources
# ---------------------------------------------------------------------------
def rows_from_report(rep) -> dict[tuple[int, str], PeriodRow]:
    """Une publication (InterimReport) -> lignes (exercice N, période) et (exercice N-1, période)."""
    out: dict[tuple[int, str], PeriodRow] = {}
    if rep is None or getattr(rep, "kind", "") != "interim" or not rep.year:
        return out
    m = re.match(r"\s*(T1|S1|T3)\b", rep.period or "")
    per = m.group(1) if m else None
    if per is None:
        return out
    origin = "document" if "chargé" in (rep.source or "").lower() else "en ligne"
    for yr, attr in ((rep.year, "current"), (rep.year - 1, "previous")):
        vals = {ln.key: getattr(ln, attr) for ln in rep.lines if is_num(getattr(ln, attr))}
        if vals:
            out[(yr, per)] = PeriodRow(values=vals, statut="Publié", published=rep.published if yr == rep.year else None,
                                       source=rep.url or rep.title, origin=origin)
    return out


def merged_rows(ticker: str, data: InterimData | None, report=None) -> dict[tuple[int, str], PeriodRow]:
    """Lignes de la société : fichier d'abord, puis publication lue en ligne ou chargée pour les postes manquants."""
    store: dict[tuple[int, str], PeriodRow] = {}
    if data is not None:
        for k, r in data.rows.get(ticker, {}).items():
            store[k] = PeriodRow(dict(r.values), r.statut, r.published, r.source, "fichier")
    for k, r in rows_from_report(report).items():
        cur = store.setdefault(k, PeriodRow(origin=r.origin, published=r.published, source=r.source, statut=r.statut))
        for key, v in r.values.items():
            cur.values.setdefault(key, v)
        if not cur.source:
            cur.source = r.source
        if cur.published is None:
            cur.published = r.published
    return store


# ---------------------------------------------------------------------------
# Construction
# ---------------------------------------------------------------------------
def build_dynamics(ticker: str, c, data: InterimData | None = None, report=None) -> Dynamics | None:
    """Dynamique de l'exercice en cours d'une société (``c`` : CompanyData des comptes annuels)."""
    store = merged_rows(ticker, data, report)
    if not store:
        return None
    is_bank = c.is_bank
    keys = keys_for(is_bank)
    rev_key = "pnb" if is_bank else "ca"
    annual_last = max((y for y, v in c.fields.get(c.revenue_field, {}).items() if is_num(v)), default=None)
    if annual_last is None and c.years:
        annual_last = max(c.years)
    # exercice courant : le plus récent de la base ayant au moins un indicateur, et non encore clos dans les comptes annuels
    usable = [k for k, r in store.items() if any(is_num(r.values.get(kk)) for kk in (rev_key, "rn"))]
    if not usable:
        return None
    year = max(k[0] for k in usable)
    if annual_last is not None and year <= annual_last:
        return None
    periods = sorted({p for (y, p) in usable if y == year}, key=lambda p: PERIOD_ORDER[p])
    dyn = Dynamics(ticker=ticker, is_bank=is_bank, year=year, period=periods[-1])
    for p in periods:
        cur, prv = store[(year, p)], store.get((year - 1, p))
        cmp_ = Compare(year=year, period=p, published=cur.published, source=cur.source, statut=cur.statut,
                       origin=cur.origin)
        for key, _lab, _s in keys:
            a = cur.values.get(key)
            b = prv.values.get(key) if prv is not None else None
            if is_num(a) or is_num(b):
                cmp_.metrics[key] = (a if is_num(a) else None, b if is_num(b) else None, variation(a, b))
        if cmp_.metrics:
            dyn.compares.append(cmp_)
    if not dyn.compares:
        return None
    dyn.period = dyn.compares[-1].period
    last = dyn.latest
    # 12 mois glissants et estimation de l'exercice
    for key in LTM_KEYS:
        if key not in last.metrics or key in STOCK_KEYS:
            continue
        cur_v, prev_v, g = last.metrics[key]
        fy = c.fields.get(key, {}).get(year - 1)
        if is_num(cur_v) and is_num(prev_v) and is_num(fy):
            dyn.ltm[key] = fy + cur_v - prev_v
            dyn.ltm_base[key] = fy
            if is_num(g) and fy > 0 and cur_v > 0:
                dyn.fy_est[key] = fy * (1 + g)
    if not dyn.ltm and annual_last is not None and annual_last != year - 1:
        dyn.notes.append(f"Exercice {year - 1} absent des comptes annuels : 12 mois glissants non calculés.")
    # tendance sur le résultat net, à défaut sur le revenu
    for key in ("rn", rev_key):
        gs = [(cp.period, cp.metrics[key][2]) for cp in dyn.compares if key in cp.metrics and is_num(cp.metrics[key][2])]
        if len(gs) >= 2:
            delta = gs[-1][1] - gs[0][1]
            dyn.trend_delta = delta
            lab = "résultat net" if key == "rn" else ("PNB" if is_bank else "chiffre d'affaires")
            if delta >= 0.05:
                dyn.trend = f"Croissance du {lab} en accélération de {gs[0][0]} à {gs[-1][0]}."
            elif delta <= -0.05:
                dyn.trend = f"Croissance du {lab} en ralentissement de {gs[0][0]} à {gs[-1][0]}."
            else:
                dyn.trend = f"Croissance du {lab} stable de {gs[0][0]} à {gs[-1][0]}."
            break
    if any(cp.statut and cp.statut.lower().startswith("à vérifier") for cp in dyn.compares):
        dyn.notes.append("Certaines lignes du fichier sont marquées « À vérifier ».")
    return dyn


def momentum_from(dyn: Dynamics | None) -> dict:
    """Croissance cumulée de la dernière période, pour le pilier croissance."""
    if dyn is None:
        return {}
    rev_key = "pnb" if dyn.is_bank else "ca"
    res_key = "rn" if "rn" in dyn.latest.metrics and is_num(dyn.var("rn")) else ("rao" if is_num(dyn.var("rao")) else "re")
    lab = {k: l for k, l, _ in keys_for(dyn.is_bank)}
    return {"rev": dyn.var(rev_key), "res": dyn.var(res_key), "rev_label": lab.get(rev_key, ""),
            "res_label": lab.get(res_key, ""), "trend": dyn.trend, "trend_delta": dyn.trend_delta}
