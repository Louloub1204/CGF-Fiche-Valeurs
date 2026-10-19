"""Outils transverses : normalisation de libellés, formats français, calculs sûrs.

Conventions de l'application
----------------------------
* Les montants sont stockés en FCFA (unités), quelle que soit l'unité du
  classeur source.
* Les ratios et taux sont stockés en fraction (0,125 = 12,5%).
* Les pourcentages s'écrivent collés au nombre (« 12,5% ») et le tiret long
  est banni des textes produits.
"""
from __future__ import annotations

import math
import re
import unicodedata
from typing import Iterable

NBSP = " "


# ---------------------------------------------------------------------------
# Normalisation de texte
# ---------------------------------------------------------------------------

def strip_accents(text: str) -> str:
    text = unicodedata.normalize("NFKD", str(text))
    return "".join(c for c in text if not unicodedata.combining(c))


def norm_label(text: object) -> str:
    """Libellé comptable normalisé : majuscules, sans accents ni ponctuation."""
    if text is None:
        return ""
    s = strip_accents(str(text)).upper()
    s = s.replace("’", "'").replace("`", "'")
    s = re.sub(r"[^A-Z0-9 ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


_CODE_SEP = re.compile(r"^\s*[A-Z]{1,2}\d{0,2}\s*[-:.)]\s+")
_CODE_SYSCOHADA = re.compile(r"^(?:X[A-I]|T[A-N]|R[A-S]|H[A-L])\s+(?=[A-Z]{3,})")


def norm_label_core(text: object) -> str:
    """Forme « nue » du libellé : sans code SYSCOHADA en tête ni parenthèses.

    « XB - Chiffre d'affaires (A+B+C) » -> « CHIFFRE D AFFAIRES »
    """
    if text is None:
        return ""
    s = strip_accents(str(text)).upper().replace("’", "'")
    s = _CODE_SEP.sub("", s)
    s = re.sub(r"\(.*?\)|\[.*?\]", " ", s)
    s = re.sub(r"[^A-Z0-9 ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    s = _CODE_SYSCOHADA.sub("", s)
    return s


def norm_header(text: object) -> str:
    s = strip_accents(str(text or "")).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def clean_text(text: str) -> str:
    """Texte final : pas de tiret long, pas d'espace avant le signe %."""
    if text is None:
        return ""
    t = str(text).replace("—", "-").replace("–", "-")
    t = re.sub(r"(\d)[\s  ]+%", r"\1%", t)
    return t


# ---------------------------------------------------------------------------
# Nombres
# ---------------------------------------------------------------------------

def is_num(x: object) -> bool:
    if x is None or isinstance(x, bool):
        return False
    try:
        return not math.isnan(float(x)) and not math.isinf(float(x))
    except (TypeError, ValueError):
        return False


def to_float(value: object) -> float | None:
    """Convertit un nombre écrit à la française (« 12 345,50 », « (1 200) »)."""
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value) if is_num(value) else None
    s = str(value).strip()
    if not s or s.lower() in {"-", "n.d.", "nd", "n/a", "na", "ns", "n.s.", "none", "nan"}:
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    s = s.replace(NBSP, "").replace(" ", "").replace(" ", "")
    s = s.replace("%", "").replace("FCFA", "").replace("F", "") if re.search(r"\d", s) else s
    if s.count(",") == 1 and s.count(".") == 0:
        s = s.replace(",", ".")
    elif s.count(",") >= 1 and s.count(".") == 1:
        s = s.replace(",", "")
    elif s.count(".") > 1 and s.count(",") == 0:
        s = s.replace(".", "")
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if neg else v


def safe_div(a: object, b: object) -> float | None:
    if not is_num(a) or not is_num(b) or float(b) == 0:
        return None
    return float(a) / float(b)


def growth(new: object, old: object) -> float | None:
    """Variation relative, robuste aux bases négatives."""
    if not is_num(new) or not is_num(old) or float(old) == 0:
        return None
    return (float(new) - float(old)) / abs(float(old))


def cagr(last: object, first: object, years: float) -> float | None:
    if not is_num(last) or not is_num(first) or years <= 0:
        return None
    if float(first) <= 0 or float(last) <= 0:
        return None
    return (float(last) / float(first)) ** (1.0 / years) - 1.0


def mean_of(values: Iterable) -> float | None:
    vals = [float(v) for v in values if is_num(v)]
    return sum(vals) / len(vals) if vals else None


def median_of(values: Iterable) -> float | None:
    vals = sorted(float(v) for v in values if is_num(v))
    if not vals:
        return None
    n = len(vals)
    mid = n // 2
    return vals[mid] if n % 2 else (vals[mid - 1] + vals[mid]) / 2


def clamp(x: float, lo: float = 0.0, hi: float = 100.0) -> float:
    return max(lo, min(hi, x))


def scale(x: object, lo: float, hi: float, reverse: bool = False) -> float | None:
    """Projette x linéairement de [lo, hi] vers [0, 100], borné."""
    if not is_num(x):
        return None
    x = float(x)
    if hi == lo:
        return 50.0
    s = (x - lo) / (hi - lo) * 100.0
    s = clamp(s)
    return 100.0 - s if reverse else s


# ---------------------------------------------------------------------------
# Formats français
# ---------------------------------------------------------------------------

def _group(int_str: str) -> str:
    out = []
    while len(int_str) > 3:
        out.insert(0, int_str[-3:])
        int_str = int_str[:-3]
    out.insert(0, int_str)
    return NBSP.join(out)


def fmt_num(x: object, dec: int = 0, sign: bool = False, na: str = "n.d.") -> str:
    if not is_num(x):
        return na
    v = float(x)
    neg = v < 0
    v = abs(v)
    s = f"{v:.{dec}f}"
    if "." in s:
        ip, dp = s.split(".")
        s = _group(ip) + "," + dp
    else:
        s = _group(s)
    if neg and float(s.replace(NBSP, "").replace(",", ".")) != 0:
        return "-" + s
    return ("+" + s) if sign and float(x) > 0 else s


def fmt_pct(x: object, dec: int = 1, sign: bool = False, na: str = "n.d.") -> str:
    """Fraction -> pourcentage collé : 0.125 -> « 12,5% »."""
    if not is_num(x):
        return na
    return fmt_num(float(x) * 100.0, dec, sign=sign) + "%"


def fmt_pts(x: object, dec: int = 1, sign: bool = True, na: str = "n.d.") -> str:
    """Valeur déjà en pourcentage : 6.1 -> « 6,1% »."""
    if not is_num(x):
        return na
    return fmt_num(float(x), dec, sign=sign) + "%"


def fmt_bp(x: object, na: str = "n.d.") -> str:
    """Écart en fraction -> points de base : 0.012 -> « +120 pb »."""
    if not is_num(x):
        return na
    return fmt_num(float(x) * 10000.0, 0, sign=True) + f"{NBSP}pb"


def fmt_x(x: object, dec: int = 1, na: str = "n.d.") -> str:
    if not is_num(x):
        return na
    if float(x) <= 0:
        return "n.s."
    return fmt_num(x, dec) + "x"


def fmt_fcfa(x: object, na: str = "n.d.", short: bool = True) -> str:
    """Montant en FCFA, abrégé (Mds / M FCFA)."""
    if not is_num(x):
        return na
    v = float(x)
    a = abs(v)
    if not short:
        return fmt_num(v, 0) + f"{NBSP}FCFA"
    if a >= 1e9:
        return fmt_num(v / 1e9, 1 if a < 1e12 else 0) + f"{NBSP}Mds FCFA"
    if a >= 1e6:
        return fmt_num(v / 1e6, 1) + f"{NBSP}M FCFA"
    return fmt_num(v, 0) + f"{NBSP}FCFA"


def fmt_mds(x: object, dec: int = 1, na: str = "n.d.") -> str:
    """Montant en milliards de FCFA, sans unité (pour les tableaux)."""
    if not is_num(x):
        return na
    return fmt_num(float(x) / 1e9, dec)


def fmt_price(x: object, na: str = "n.d.") -> str:
    if not is_num(x):
        return na
    dec = 0 if abs(float(x)) >= 100 else 2
    return fmt_num(x, dec) + f"{NBSP}FCFA"


def fmt_date_fr(d) -> str:
    mois = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août",
            "septembre", "octobre", "novembre", "décembre"]
    try:
        return f"{d.day} {mois[d.month - 1]} {d.year}"
    except Exception:
        return str(d)
