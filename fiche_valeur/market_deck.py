"""Présentation PowerPoint « Contexte de marché actions » produite depuis le tableau de bord multi-valeurs.

Diapositive 1 : modèle CGF Gestion « Le marché des actions de la BRVM » (fichier modèle chargé par l'analyste),
mis à jour à la date d'arrêté :
- données quantitatives : Bulletin Officiel de la Cote (BOC) de la BRVM du jour d'arrêté (repli sur la séance
  précédente), complétées par la base SMF (performance par titre, capitalisation de fin d'année, liquidité) ;
- données qualitatives : actualités du marché (Google Actualités, fenêtre se terminant à la date d'arrêté) et
  fiches de valeur (fondamentaux des titres en forte hausse), rédigées par Claude si une clé API est fournie,
  sinon par des règles fixes.
Diapositives 2 à 4 : synthèse des fiches de valeur (signaux, valorisation par secteur, convictions et vigilance).
"""
from __future__ import annotations

import copy
import io
import json
import re
import statistics
import urllib.parse
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path

import pandas as pd
import requests

from . import config
from .utils import clean_text, is_num

TEMPLATE = Path(__file__).with_name("templates") / "modele_contexte_marche.pptx"
BOC_URL = "https://bfin.brvm.org/boc/BOC_JOUR/BOC_{d:%Y%m%d}.pdf"
UA = {"User-Agent": "Mozilla/5.0 FicheValeurBRVM/1.0"}
MOIS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre",
        "novembre", "décembre"]
SECTEURS_BOC = {"CB": "Consommation base", "CD": "Consommation discr.", "FIN": "Finances", "IND": "Industrie",
                "ENE": "Énergie", "SPU": "Services pub.", "TEL": "Télécoms"}
SECTEURS_INDICES = {"TELECOMMUNICATIONS": "Télécoms", "CONSOMMATION DISCRETIONNAIRE": "Consommation discr.",
                    "SERVICES FINANCIERS": "Finances", "CONSOMMATION DE BASE": "Consommation base",
                    "INDUSTRIELS": "Industrie", "ENERGIE": "Énergie", "SERVICES PUBLICS": "Services pub."}
INDICES_BASE = {"BRVMTEL": "Télécoms", "BRVMCD": "Consommation discr.", "BRVMSFIN": "Finances",
                "BRVMCB": "Consommation base", "BRVMIND": "Industrie", "BRVMENER": "Énergie", "BRVMSPUB": "Services pub."}
SECTEURS_REF = {"Services financiers": "Finances", "Télécommunications": "Télécoms",
                "Consommation de base": "Consommation base", "Consommation discrétionnaire": "Consommation discr.",
                "Industriels": "Industrie", "Énergie": "Énergie", "Services publics": "Services pub."}


# ---------------------------------------------------------------------------
# Formats
# ---------------------------------------------------------------------------
def pct(x, d=2, sign=True) -> str:
    if not is_num(x):
        return "n.d."
    s = f"{x * 100:+.{d}f}%" if sign else f"{x * 100:.{d}f}%"
    return s.replace(".", ",")


def num(x, d=0) -> str:
    if not is_num(x):
        return "n.d."
    return f"{x:,.{d}f}".replace(",", " ").replace(".", ",")


def date_fr(d: date, mois_seul=False) -> str:
    return f"{MOIS[d.month - 1]} {d.year}" if mois_seul else f"{d.day} {MOIS[d.month - 1]} {d.year}"


# ---------------------------------------------------------------------------
# Bulletin Officiel de la Cote
# ---------------------------------------------------------------------------
@dataclass
class Cotation:
    symbole: str
    nom: str
    secteur: str
    cours: float | None = None
    var_jour: float | None = None
    volume: float | None = None
    valeur: float | None = None
    ytd: float | None = None
    dividende: float | None = None
    rendement: float | None = None
    per: float | None = None


@dataclass
class Boc:
    date: date | None = None
    url: str = ""
    composite: float | None = None
    composite_jour: float | None = None
    composite_ytd: float | None = None
    prestige: float | None = None
    prestige_ytd: float | None = None
    capitalisation: float | None = None
    volume: float | None = None
    valeur: float | None = None
    n_transiges: int | None = None
    n_hausse: int | None = None
    n_baisse: int | None = None
    n_inchanges: int | None = None
    secteurs: list = field(default_factory=list)      # (nom, n sociétés, valeur, évol. jour, évol. annuelle, PER)
    cotations: dict = field(default_factory=dict)     # symbole -> Cotation


_NUM = r"-?\d{1,3}(?: \d{3})*(?:,\d+)?"


def _f(s: str | None) -> float | None:
    if s is None:
        return None
    s = s.replace(" ", " ").replace("\xa0", " ").strip()
    try:
        return float(s.replace(" ", "").replace(",", "."))
    except ValueError:
        return None


def _p(s: str | None) -> float | None:
    v = _f(s)
    return None if v is None else v / 100


def _groups(tokens: list[str], k: int, check) -> list[float] | None:
    """Répartit des jetons numériques (séparateur de milliers = espace) en k nombres, en retenant le découpage
    le plus plausible (``check`` note un découpage, None = impossible)."""
    best, best_score = None, None

    def rec(i, acc):
        nonlocal best, best_score
        if len(acc) == k:
            if i == len(tokens):
                sc = check(acc)
                if sc is not None and (best_score is None or sc < best_score):
                    best, best_score = list(acc), sc
            return
        if i >= len(tokens) or not re.fullmatch(r"-?\d{1,3}(,\d+)?", tokens[i]):
            return
        j = i + 1
        while True:
            val = _f(" ".join(tokens[i:j]))
            rec(j, acc + [val])
            if j >= len(tokens) or not re.fullmatch(r"\d{3}(,\d+)?", tokens[j]) or "," in tokens[j - 1]:
                break
            j += 1
    rec(0, [])
    return best


def _prix_ok(vals):
    if any(v is None or v <= 0 for v in vals):
        return None
    r = max(vals) / min(vals)
    return None if r > 1.3 else r


def _parse_queue(queue: str) -> dict | None:
    """Champs numériques d'une ligne de cotation : cours précédent, ouverture, clôture, variation du jour, volume,
    valeur, cours de référence, variation annuelle, [dividende, date], [rendement], [PER]."""
    q = re.sub(r"\s+", " ", queue).strip()
    parts = re.split(r"(-?\d+,\d+) ?%", q)
    # parts : [avant, pct1, entre, pct2, après, pct3?, fin?]
    if len(parts) < 5:
        return None
    a = parts[0].split()
    prix = _groups(a, 3, _prix_ok)
    if not prix:
        return None
    var_j = _p(parts[1])
    b = parts[2].split()
    close = prix[2]

    def check_b(v):
        vol, val, ref = v
        if ref is None or ref <= 0 or vol is None or val is None:
            return None
        sc = abs(ref / close - 1)
        if vol > 0:
            ratio = val / (vol * close)
            if not 0.6 <= ratio <= 1.6:
                return None
            sc += abs(ratio - 1)
        elif val > 0:
            return None
        return sc
    vvr = _groups(b, 3, check_b) or (lambda r: [0.0, 0.0, r[0]] if r else None)(_groups(b, 1, lambda v: 0))
    if not vvr:
        return None
    out = {"cours": close, "var_jour": var_j, "volume": vvr[0], "valeur": vvr[1], "ytd": _p(parts[3])}
    reste = parts[4].strip()
    m = re.search(r"\d{1,2}-[a-zéûô]+\.?-\d{2}", reste)
    if m:
        out["dividende"] = _f(reste[:m.start()].strip()) if reste[:m.start()].strip() else None
        after = reste[m.end():].strip()
    else:
        after = reste
    if len(parts) >= 7:
        out["rendement"] = _p(parts[5])
        after = parts[6].strip()
    per = re.match(rf"\s*({_NUM})\s*$", after)
    if per:
        out["per"] = _f(per.group(1))
    return out


_LIGNE = re.compile(r"(?<![A-Z])(CB|CD|FIN|IND|ENE|SPU|TEL)\s+([A-Z]{3,6})\s+(.+?)(?=(?<![A-Z])(?:CB|CD|FIN|IND|ENE|SPU|TEL)"
                    r"\s+[A-Z]{3,6}\s|TOTAL|$)", re.S)


def parse_boc_text(text: str, symboles=None) -> Boc:
    t = text.replace(" ", " ").replace("\xa0", " ")
    b = Boc()
    m = re.search(r"(lundi|mardi|mercredi|jeudi|vendredi|samedi|dimanche)\s+(\d{1,2})\s+([a-zéû]+)\s+(\d{4})", t, re.I)
    if m and m.group(3).lower() in MOIS:
        b.date = date(int(m.group(4)), MOIS.index(m.group(3).lower()) + 1, int(m.group(2)))

    def idx(label):
        mm = re.search(rf"{label}\s+({_NUM})\s+Variation Jour\s+({_NUM}) ?%\s+Variation annuelle\s+({_NUM}) ?%", t)
        return (_f(mm.group(1)), _p(mm.group(2)), _p(mm.group(3))) if mm else (None, None, None)
    b.composite, b.composite_jour, b.composite_ytd = idx("BRVM COMPOSITE")
    b.prestige, _, b.prestige_ytd = idx("BRVM PRESTIGE")

    def ligne(label):
        mm = re.search(rf"{label}\s+((?:\d{{1,3}})(?: \d{{3}})*)(?:\s+-?[\d ]*,\d+ ?%)?", t)
        return _f(mm.group(1)) if mm else None
    b.capitalisation = ligne(r"Capitalisation boursière \(FCFA\)\s*\(Actions & Droits\)")
    b.volume = ligne(r"Volume échangé \(Actions & Droits\)")
    b.valeur = ligne(r"Valeur transigée \(FCFA\)\s*\(Actions & Droits\)")
    for attr, lab in (("n_transiges", "transigés"), ("n_hausse", "en hausse"), ("n_baisse", "en baisse"),
                      ("n_inchanges", "inchangés")):
        mm = re.search(rf"Nombre de titres {lab}\s+(\d+)", t)
        setattr(b, attr, int(mm.group(1)) if mm else None)
    for mm in re.finditer(rf"BRVM - ([A-ZÉÈ ]+?)(?: \(\*\*\))?\s+(\d+)\s+({_NUM})\s+({_NUM}) ?%\s+({_NUM}) ?%"
                          rf"(.*)$", t, re.M):
        per = re.search(r"\s(\d{1,3},\d+)\s*$", mm.group(6) or "")
        nom = mm.group(1).strip().upper()
        b.secteurs.append((SECTEURS_INDICES.get(nom, nom.capitalize()), int(mm.group(2)), _f(mm.group(3)),
                           _p(mm.group(4)), _p(mm.group(5)), _f(per.group(1)) if per else None))
    flat = re.sub(r"\s+", " ", t)
    # Lignes de cotation repérées par le symbole (connu du référentiel) : le code secteur n'est pas toujours
    # imprimé sur chaque ligne du BOC (cellule fusionnée par groupe), il n'est donc pas utilisé pour le découpage.
    syms = sorted(set(symboles or config.TICKERS), key=len, reverse=True)
    pat = re.compile(r"(?<![A-Z0-9])(" + "|".join(map(re.escape, syms)) + r")(?![A-Z0-9])")
    occ = [m_ for m_ in pat.finditer(flat)]
    fin_tab = [m_.start() for m_ in re.finditer(r"\bTOTAL\b", flat)]
    for i, m_ in enumerate(occ):
        sym = m_.group(1)
        if sym in b.cotations:
            continue
        end = occ[i + 1].start() if i + 1 < len(occ) else len(flat)
        end = min([end] + [f for f in fin_tab if f > m_.end()])
        seg = flat[m_.end():end]
        seg = re.sub(r"\s(?:CB|CD|FIN|IND|ENE|SPU|TEL)\s*$", " ", seg)        # code secteur de la ligne suivante
        first = re.search(r"(?<![A-Za-z(])\d", seg)
        if not first:
            continue
        nom = re.sub(r"^(?:CB|CD|FIN|IND|ENE|SPU|TEL)\s+", "", seg[:first.start()].strip())
        champs = _parse_queue(seg[first.start():])
        if not champs or not is_num(champs.get("ytd")):
            continue
        ref = config.TICKERS.get(sym, {})
        secteur = SECTEURS_REF.get(ref.get("secteur"), ref.get("secteur", "n.d."))
        b.cotations[sym] = Cotation(sym, clean_text(nom) or ref.get("nom", sym), secteur, **champs)
    return b


def parse_boc(data: bytes) -> Boc:
    import pdfplumber
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        text = "\n".join((p.extract_text() or "") for p in pdf.pages[:8])
    return parse_boc_text(text)


def fetch_boc(d: date, errors: list[str] | None = None, max_back: int = 8) -> Boc | None:
    """BOC du jour d'arrêté ou, à défaut (week-end, jour férié), de la séance précédente la plus proche."""
    motifs = []
    for i in range(max_back):
        dd = d - timedelta(days=i)
        if dd.weekday() >= 5:
            continue
        url = BOC_URL.format(d=dd)
        try:
            r = requests.get(url, timeout=30, headers=UA)
            if r.status_code != 200:
                motifs.append(f"{dd:%d/%m} : réponse HTTP {r.status_code}")
                continue
            if not r.content.startswith(b"%PDF"):
                motifs.append(f"{dd:%d/%m} : le site n'a pas renvoyé de PDF")
                continue
            b = parse_boc(r.content)
            b.url, b.date = url, b.date or dd
            if b.composite is not None or b.cotations:
                return b
            motifs.append(f"{dd:%d/%m} : PDF lu mais sans données reconnues")
        except Exception as exc:
            motifs.append(f"{dd:%d/%m} : {type(exc).__name__} ({str(exc)[:80]})")
    if errors is not None:
        errors.append(f"BOC indisponible entre le {d - timedelta(days=max_back - 1):%d/%m/%Y} et le {d:%d/%m/%Y} "
                      f"({' ; '.join(motifs[:3]) or 'aucune séance'}) : données de la base SMF utilisées.")
    return None


# ---------------------------------------------------------------------------
# Instantané de marché (BOC complété par la base SMF)
# ---------------------------------------------------------------------------
@dataclass
class Instantane:
    date: date
    composite: float | None = None
    composite_ytd: float | None = None
    composite_annee_prec: float | None = None
    capitalisation: float | None = None          # FCFA
    capitalisation_fin_annee: float | None = None
    titres: list = field(default_factory=list)   # Cotation (avec ytd), toute la cote
    secteurs: list = field(default_factory=list)
    volume_moyen: float | None = None
    valeur_moyenne: float | None = None
    seance_volume: float | None = None
    seance_valeur: float | None = None
    sources: list = field(default_factory=list)
    boc: Boc | None = None

    @property
    def n_hausse(self):
        return sum(1 for c in self.titres if is_num(c.ytd) and c.ytd > 0)

    def top(self, n=5):
        return sorted([c for c in self.titres if is_num(c.ytd)], key=lambda c: c.ytd, reverse=True)[:n]

    def flop(self, n=5):
        return sorted([c for c in self.titres if is_num(c.ytd)], key=lambda c: c.ytd)[:n]


ABREV = [(r"\bCOTE D'?IVOIRE\b", "CI"), (r"\bBURKINA FASO\b", "BF"), (r"\bSENEGAL\b", "SN"), (r"\bBENIN\b", "BJ"),
         (r"\bINTERNATIONAL(E)?\b", "INTL"), (r"\bECOBANK TRANSNATIONAL INCORPORATED\b", "ETI"),
         (r"\bSOCIETE GENERALE\b", "SG"), (r"\bBANK OF AFRICA\b", "BOA"), (r"\s*\(EX [^)]*\)", ""),
         (r"\bNATIONALE\b", "NAT."), (r"\bCOMPAGNIE\b", "CIE")]


def _nom(t: str) -> str:
    return config.TICKERS.get(t, {}).get("nom", t)


def instantane(d: date, mb=None, boc: Boc | None = None, boc_fin: Boc | None = None) -> Instantane:
    """Chiffres de marché à la date d'arrêté : BOC en priorité, base SMF (cours, nombre de titres, volumes) sinon."""
    ins = Instantane(date=(boc.date if boc and boc.date else d), boc=boc)
    closes = getattr(mb, "closes", None)
    if closes is not None:
        closes = closes[closes.index <= pd.Timestamp(ins.date)]
    last = closes.ffill().iloc[-1] if closes is not None and not closes.empty else None
    fin = closes[closes.index <= pd.Timestamp(date(ins.date.year - 1, 12, 31))].ffill() if last is not None else None
    fin_ = fin.iloc[-1] if fin is not None and not fin.empty else None
    fin2 = closes[closes.index <= pd.Timestamp(date(ins.date.year - 2, 12, 31))].ffill() if last is not None else None
    fin2_ = fin2.iloc[-1] if fin2 is not None and not fin2.empty else None
    from .market import INDEX_TICKER
    idx = INDEX_TICKER if closes is not None and INDEX_TICKER in closes.columns else ".BRVMCI"
    if boc and boc.composite:
        ins.composite, ins.composite_ytd = boc.composite, boc.composite_ytd
        ins.sources.append(f"BRVM, Bulletin Officiel de la Cote du {ins.date:%d/%m/%Y}")
    elif last is not None and idx in last and is_num(last[idx]):
        ins.composite = float(last[idx])
        if fin_ is not None and is_num(fin_.get(idx)):
            ins.composite_ytd = ins.composite / float(fin_[idx]) - 1
    if fin_ is not None and fin2_ is not None and is_num(fin_.get(idx)) and is_num(fin2_.get(idx)):
        ins.composite_annee_prec = float(fin_[idx]) / float(fin2_[idx]) - 1
    shares = getattr(mb, "shares_total", {}) or {}
    if boc and boc.capitalisation:
        ins.capitalisation = boc.capitalisation
    elif last is not None and shares:
        ins.capitalisation = sum(float(last[t]) * shares[t] for t in shares if t in last and is_num(last[t]))
    if boc_fin and boc_fin.capitalisation:
        ins.capitalisation_fin_annee = boc_fin.capitalisation
    elif fin_ is not None and shares:
        ins.capitalisation_fin_annee = sum(float(fin_[t]) * shares[t] for t in shares if t in fin_ and is_num(fin_[t]))
    # performances par titre
    if boc and boc.cotations:
        for c in boc.cotations.values():
            if c.symbole in config.TICKERS:
                c.nom = _nom(c.symbole)
        ins.titres = list(boc.cotations.values())
    if last is not None and fin_ is not None:
        deja = {c.symbole for c in ins.titres}
        for t in closes.columns:
            if t in deja or t == idx or str(t).upper().startswith("BRVM") or not is_num(last[t]) or not is_num(fin_.get(t)) or not fin_[t]:
                continue
            ref = config.TICKERS.get(t, {})
            ins.titres.append(Cotation(t, _nom(t), SECTEURS_REF.get(ref.get("secteur"), ref.get("secteur", "n.d.")),
                                       cours=float(last[t]), ytd=float(last[t]) / float(fin_[t]) - 1))
        if not ins.sources:
            ins.sources.append("Base de données SMF (cours de clôture)")
    if boc and boc.secteurs:
        ins.secteurs = boc.secteurs
        ins.seance_volume, ins.seance_valeur = boc.volume, boc.valeur
    elif last is not None and fin_ is not None:
        # indices sectoriels de la base SMF
        for col, lab in INDICES_BASE.items():
            if col in last and is_num(last[col]) and is_num(fin_.get(col)) and fin_[col]:
                ins.secteurs.append((lab, None, float(last[col]), None, float(last[col]) / float(fin_[col]) - 1, None))
    vols = getattr(mb, "volumes", None)
    if vols is not None and closes is not None:
        cols = [t for t in vols.columns if t in closes.columns and t != idx]
        v = vols[cols][(vols.index <= pd.Timestamp(ins.date)) & (vols.index.year == ins.date.year)]
        if not v.empty:
            ins.volume_moyen = float(v.sum(axis=1).mean())
            ins.valeur_moyenne = float((v * closes[cols].reindex(v.index)).sum(axis=1).mean())
    return ins


# ---------------------------------------------------------------------------
# Actualités du marché
# ---------------------------------------------------------------------------
@dataclass
class Actu:
    titre: str
    source: str
    date: datetime | None
    lien: str = ""


def actualites_marche(d: date, jours: int = 45, limite: int = 8, errors: list[str] | None = None) -> list[Actu]:
    q = (f'(BRVM OR "Bourse régionale des valeurs mobilières") after:{d - timedelta(days=jours):%Y-%m-%d} '
         f'before:{d + timedelta(days=1):%Y-%m-%d}')
    url = f"https://news.google.com/rss/search?q={urllib.parse.quote(q)}&hl=fr&gl=CI&ceid=CI:fr"
    try:
        r = requests.get(url, timeout=20, headers=UA)
        r.raise_for_status()
        root = ET.fromstring(r.content)
    except Exception as exc:
        if errors is not None:
            errors.append(f"Actualités du marché indisponibles ({type(exc).__name__}).")
        return []
    fin = datetime(d.year, d.month, d.day, 23, 59, tzinfo=timezone.utc)
    out, vus = [], set()
    for it in root.iter("item"):
        titre = (it.findtext("title") or "").strip()
        src_el = it.find("source")
        src = (src_el.text or "").strip() if src_el is not None else ""
        if src and titre.endswith(" - " + src):
            titre = titre[: -len(src) - 3].strip()
        cle = titre.lower()[:70]
        if not titre or cle in vus:
            continue
        try:
            pub = parsedate_to_datetime(it.findtext("pubDate") or "")
        except Exception:
            pub = None
        if pub is not None and pub > fin:
            continue
        vus.add(cle)
        out.append(Actu(clean_text(titre), src, pub, it.findtext("link") or ""))
    out.sort(key=lambda a: a.date or datetime(1970, 1, 1, tzinfo=timezone.utc), reverse=True)
    return out[:limite]


# ---------------------------------------------------------------------------
# Lecture des fiches de valeur
# ---------------------------------------------------------------------------
def _rn(a):
    c = a.company
    f = c.fields.get(c.net_income_field, {})
    ys = [y for y, v in f.items() if is_num(v)]
    return (max(ys), f[max(ys)]) if ys else (None, None)


def speculatifs(ins: Instantane, analyses: list, n: int = 5) -> list[tuple]:
    """Titres en forte hausse sans ancrage fondamental : perte au dernier exercice, aucun dividende, ou fiche en
    forte surcote (potentiel inférieur à -30%)."""
    par = {a.company.ticker: a for a in analyses}
    out = []
    for c in ins.top(15):
        a = par.get(c.symbole)
        if a is None or not is_num(c.ytd) or c.ytd < 0.25:
            continue
        y, rn = _rn(a)
        motifs = []
        if is_num(rn) and rn < 0:
            motifs.append(f"perte nette de {num(abs(rn) / 1e9, 2)} Mds FCFA en {y}")
        if a.valuation.dps == 0:
            motifs.append(f"aucun dividende au titre de {a.valuation.dps_exercice}")
        up = a.valuation.upside
        if is_num(up) and up < -0.30:
            motifs.append(f"cours {pct(-up, 0, sign=False)} au-dessus de la juste valeur")
        if motifs:
            out.append((c, a, motifs))
    return out[:n]


# ---------------------------------------------------------------------------
# Rédaction
# ---------------------------------------------------------------------------
def _perspectives_calendrier(d: date) -> tuple[str, str]:
    y, m = d.year, d.month
    if m <= 3:
        return (f"Perspectives {y}", f"La publication des résultats annuels {y - 1} et les propositions de dividendes, "
                                     f"attendues d'ici fin avril, constitueront le principal catalyseur.")
    if m <= 6:
        return ("Perspectives 2ème semestre " + str(y), "La publication des résultats semestriels constituera le principal "
                                                         "test de durabilité de la tendance.")
    if m <= 9:
        return (f"Perspectives fin {y}", f"Les résultats semestriels et les publications du troisième trimestre "
                                          f"permettront de vérifier si la progression des cours reflète une amélioration "
                                          f"durable des bénéfices.")
    return (f"Perspectives {y + 1}", f"Les résultats annuels {y}, attendus au premier trimestre {y + 1}, et les annonces "
                                      f"de dividendes diront si la progression des cours repose sur une amélioration "
                                      f"durable des bénéfices ou sur un simple effet de revalorisation.")


def textes_regles(ins: Instantane, analyses: list, actus: list[Actu]) -> dict:
    d = ins.date
    n = len([c for c in ins.titres if is_num(c.ytd)])
    t = {}
    seg = [("normal", "Le BRVM Composite s'établit à "), ("bold", f"{num(ins.composite, 2)} points"),
           ("normal", f" au {d:%d/%m/%Y}, soit "), ("bold", f"{pct(ins.composite_ytd)} depuis le début de l'année")]
    if is_num(ins.composite_annee_prec):
        seg.append(("normal", f", après {pct(ins.composite_annee_prec)} en {d.year - 1}"))
    seg.append(("normal", ". La capitalisation du marché des actions atteint "))
    seg.append(("bold", f"{num(ins.capitalisation / 1e9)} milliards FCFA"))
    if is_num(ins.capitalisation_fin_annee) and ins.capitalisation_fin_annee:
        seg.append(("normal", f" (contre {num(ins.capitalisation_fin_annee / 1e9)} Mds au 31/12/{d.year - 1}, "
                              f"{pct(ins.capitalisation / ins.capitalisation_fin_annee - 1, 0)})."))
    else:
        seg.append(("normal", "."))
    t["evolution"] = seg
    # fondamentaux
    par = {a.company.ticker: a for a in analyses}
    seg = [("normal", f"{ins.n_hausse} des {n} valeurs affichent une hausse depuis le 1er janvier. ")]
    sect = sorted([s for s in ins.secteurs if is_num(s[4])], key=lambda s: s[4], reverse=True)
    if sect:
        seg += [("normal", "Secteurs moteurs : "),
                ("bold", ", ".join(f"{s[0].lower()} ({pct(s[4])})" for s in sect[:2])),
                ("normal", f" ; en retrait : {sect[-1][0].lower()} ({pct(sect[-1][4])}). ")]
    banques = [a for a in analyses if a.company.is_bank]
    if banques:
        tot = sum(_rn(a)[1] for a in banques if is_num(_rn(a)[1]))
        ys = sorted({_rn(a)[0] for a in banques if _rn(a)[0]})
        hausse = [c for c in ins.top(48) if c.symbole in par and par[c.symbole].company.is_bank and is_num(c.ytd)][:3]
        seg.append(("normal", f"Le secteur bancaire ancre sa performance dans des résultats solides "
                              f"({num(tot / 1e9)} Mds FCFA de résultat net {ys[-1] if ys else ''} pour les "
                              f"{len(banques)} banques analysées)"))
        if hausse:
            seg += [("normal", " : "), ("bold", ", ".join(f"{_court(c.nom)} ({pct(c.ytd)})" for c in hausse))]
        seg.append(("normal", ". "))
    spec = speculatifs(ins, analyses)
    if spec:
        txt = "À surveiller : plusieurs valeurs progressent sans ancrage fondamental identifiable. "
        txt += " ; ".join(f"{_court(c.nom)} ({pct(c.ytd)} depuis janvier : {', '.join(m)})" for c, _, m in spec[:3])
        txt += ". Ces hausses, déconnectées des résultats, revêtent un caractère spéculatif."
        seg.append(("vigilance", txt))
    t["fondamentaux"] = seg
    lab, cal = _perspectives_calendrier(d)
    liq = ""
    if is_num(ins.volume_moyen) and is_num(ins.valeur_moyenne):
        liq = (f" La liquidité reste l'enjeu structurel du marché, avec en moyenne {num(ins.volume_moyen)} titres et "
               f"{num(ins.valeur_moyenne / 1e9, 2)} milliards FCFA échangés par séance depuis janvier.")
    sur = [a for a in analyses if is_num(a.valuation.upside)]
    val = ""
    if sur:
        n_s = sum(1 for a in sur if a.valuation.upside < -0.15)
        val = (f" Sur {len(sur)} titres valorisés, {n_s} se traitent au-dessus de leur juste valeur de plus de 15% : "
               f"la sélectivité devient déterminante.")
    t["perspectives_titre"] = lab
    t["perspectives"] = cal + liq + val
    t["actualites"] = [a.titre for a in actus[:4]]
    return t


def _court(nom: str, n: int = 24) -> str:
    """Nom court en capitales, comme dans le modèle (« CORIS BANK INTL BF »)."""
    x = re.sub(r"\s+", " ", nom).strip().upper()
    for a, b in ABREV[:4]:
        x = re.sub(a, b, x)
    for a, b in ABREV[4:]:
        if len(x) <= n:
            break
        x = re.sub(a, b, x)
    return x if len(x) <= n else x[:n - 1].rstrip() + "."


def textes_claude(ins: Instantane, analyses: list, actus: list[Actu], api_key: str, model: str) -> dict | None:
    """Rédaction de la diapositive 1 par Claude, à partir des seuls chiffres calculés et des titres d'actualité."""
    if not api_key:
        return None
    base = textes_regles(ins, analyses, actus)
    donnees = {
        "date_arrete": ins.date.isoformat(), "brvm_composite": ins.composite, "composite_ytd": ins.composite_ytd,
        "composite_annee_precedente": ins.composite_annee_prec,
        "capitalisation_mds": ins.capitalisation / 1e9 if ins.capitalisation else None,
        "capitalisation_fin_annee_prec_mds": ins.capitalisation_fin_annee / 1e9 if ins.capitalisation_fin_annee else None,
        "valeurs_en_hausse": ins.n_hausse, "nombre_valeurs": len(ins.titres),
        "top5": [(c.nom, round(c.ytd * 100, 2)) for c in ins.top()], "flop5": [(c.nom, round(c.ytd * 100, 2)) for c in ins.flop()],
        "secteurs": [(s[0], round(s[4] * 100, 2) if is_num(s[4]) else None) for s in ins.secteurs],
        "volume_moyen_seance": ins.volume_moyen, "valeur_moyenne_seance_mds": (ins.valeur_moyenne or 0) / 1e9,
        "titres_speculatifs": [(c.nom, round(c.ytd * 100, 1), m) for c, _, m in speculatifs(ins, analyses)],
        "signaux_fiches": {s: sum(1 for a in analyses if a.decision.signal == s) for s in
                           ("ACHAT", "CONSERVER", "ATTENDRE", "ALLÉGER", "VENDRE")},
        "actualites_marche": [f"{a.titre} ({a.source}, {a.date:%d/%m/%Y})" if a.date else a.titre for a in actus],
        "calendrier": base["perspectives"],
    }
    prompt = (
        "Tu es gérant actions chez CGF Gestion (BRVM). À partir UNIQUEMENT des données JSON ci-dessous (n'invente "
        "aucun chiffre ni fait ; les actualités ne servent qu'au contexte qualitatif), rédige en français "
        "institutionnel, phrases complètes :\n"
        "1) evolution : 45 à 60 mots sur l'évolution générale du marché depuis janvier ;\n"
        "2) fondamentaux : 50 à 70 mots sur la diffusion de la hausse, les secteurs et les fondamentaux ;\n"
        "3) vigilance : 40 à 70 mots commençant par « À surveiller : » sur les titres spéculatifs listés (vide s'il n'y "
        "en a pas) ;\n4) perspectives : 60 à 85 mots sur les perspectives, la liquidité et les actualités pertinentes.\n"
        "Pourcentages collés au nombre avec virgule décimale (+54,65%), pas de tiret long. Réponds en JSON strict : "
        '{"evolution": "...", "fondamentaux": "...", "vigilance": "...", "perspectives": "..."}\n\n'
        + json.dumps(donnees, ensure_ascii=False, default=str))
    try:
        from .commentary import ANTHROPIC_URL
        r = requests.post(ANTHROPIC_URL, timeout=90,
                          headers={"x-api-key": api_key, "anthropic-version": "2023-06-01",
                                   "content-type": "application/json"},
                          json={"model": model, "max_tokens": 1500, "messages": [{"role": "user", "content": prompt}]})
        r.raise_for_status()
        txt = "".join(b.get("text", "") for b in r.json().get("content", []) if b.get("type") == "text")
        data = json.loads(txt[txt.find("{"): txt.rfind("}") + 1])
        out = dict(base)
        out["evolution"] = [("normal", clean_text(data["evolution"]))]
        out["fondamentaux"] = [("normal", clean_text(data["fondamentaux"]) + " ")]
        if data.get("vigilance"):
            out["fondamentaux"].append(("vigilance", clean_text(data["vigilance"])))
        out["perspectives"] = clean_text(data["perspectives"])
        out["redaction"] = "Claude"
        return out
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Construction du PowerPoint
# ---------------------------------------------------------------------------
NAVY, NAVY2, INK, MUTED = "0E3A5C", "08263F", "2B3A45", "5B6B76"
GREEN, GOLD, RED, BLUE, LIGHT = "00AF52", "C4972A", "E8928A", "2E89A8", "EBF0F4"
SIGNAL_COL = {"ACHAT": "00AF52", "CONSERVER": "2E6FA3", "ATTENDRE": "C4972A", "ALLÉGER": "E67E22", "VENDRE": "C0392B"}


def _shape(slide, name):
    return next((s for s in slide.shapes if s.name == name), None)


def _set_text(shape, text):
    """Remplace le texte d'une forme en conservant le format de son premier segment."""
    if shape is None:
        return
    p = shape.text_frame.paragraphs[0]
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        p.add_run().text = text


def _set_runs(shape, segments, protos):
    """Réécrit le premier paragraphe d'une forme : segments (style, texte), styles copiés des segments du modèle."""
    p = shape.text_frame.paragraphs[0]
    proto = {k: copy.deepcopy(p.runs[i]._r) for k, i in protos.items() if i < len(p.runs)}
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    end = p._p.find("{http://schemas.openxmlformats.org/drawingml/2006/main}endParaRPr")
    for style, txt in segments:
        r = copy.deepcopy(proto.get(style, proto.get("normal")))
        r.find("{http://schemas.openxmlformats.org/drawingml/2006/main}t").text = txt
        if end is not None:
            end.addprevious(r)
        else:
            p._p.append(r)


def _len(segments) -> int:
    return sum(len(t) for _, t in segments)


def _fit(shape, n_chars, w_in, base_pt, min_pt=9.0, h_in=None):
    """Taille de police tenant dans la forme (estimation : largeur moyenne 0,5 em, interligne 1,2)."""
    from pptx.util import Pt
    h = h_in or shape.height / 914400
    pt = base_pt
    while pt > min_pt:
        cpl = w_in * 72 / (pt * 0.42)
        lines = n_chars / cpl + 0.7
        if lines * pt * 1.32 / 72 <= h:
            break
        pt -= 0.25
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            r.font.size = Pt(pt)
    return pt


def _hauteur(n_chars, w_in, pt):
    cpl = w_in * 72 / (pt * 0.42)
    return (n_chars / cpl + 0.7) * pt * 1.32 / 72


def _table_row(tbl_shape, values, sizes=(11, 10.5, 9, 11)):
    from pptx.util import Pt
    for j, v in enumerate(values):
        cell = tbl_shape.table.cell(0, j)
        cell.text_frame.word_wrap = False
        p = cell.text_frame.paragraphs[0]
        if p.runs:
            p.runs[0].text = v
            for r in p.runs[1:]:
                r._r.getparent().remove(r._r)
        else:
            p.add_run().text = v
        p.runs[0].font.size = Pt(sizes[j])


def slide_marche(prs, ins: Instantane, textes: dict, sources: list[str]):
    from pptx.util import Emu, Inches
    s = prs.slides[0]
    d = ins.date
    lab, _ = _perspectives_calendrier(d)
    suite = {"Perspectives 2ème": "pour le second semestre", "Perspectives fin": "pour la fin d'année"}
    _set_text(_shape(s, "Text 2"), f"Bilan janvier - {date_fr(d, True)} et perspectives "
              + next((v for k, v in suite.items() if lab.startswith(k)), f"pour {lab.split()[-1]}"))
    _set_text(_shape(s, "Text 5"), pct(ins.composite_ytd))
    _set_text(_shape(s, "Text 8"), num(ins.capitalisation / 1e9) if ins.capitalisation else "n.d.")
    n = len([c for c in ins.titres if is_num(c.ytd)])
    _set_text(_shape(s, "Text 11"), f"{ins.n_hausse} / {n}")
    cap_var = (ins.capitalisation / ins.capitalisation_fin_annee - 1) if ins.capitalisation and ins.capitalisation_fin_annee else None
    _set_text(_shape(s, "Text 14"), pct(cap_var, 0))
    _set_text(_shape(s, "Text 15"), f"Capitalisation vs déc. {d.year - 1}")
    if is_num(ins.composite_ytd) and ins.composite_ytd < 0:
        _shape(s, "Text 5").text_frame.paragraphs[0].runs[0].font.color.rgb = __import__("pptx").dml.color.RGBColor.from_string("C0392B")
    _set_text(_shape(s, "Text 17"), f"BILAN JANVIER-{MOIS[d.month - 1].upper()} {d.year}")
    # Textes de gauche : hauteur estimée pour empiler les deux paragraphes dans le cadre
    t18, t19 = _shape(s, "Text 18"), _shape(s, "Text 19")
    seg18 = [("label", "Évolution générale "), ("sep", "· ")] + textes["evolution"]
    seg19 = [("label", "Fondamentaux & signaux de vigilance "), ("sep", "· ")] + textes["fondamentaux"]
    _set_runs(t18, seg18, {"label": 0, "sep": 1, "normal": 2, "bold": 3})
    _set_runs(t19, seg19, {"label": 0, "sep": 1, "normal": 2, "bold": 3, "vigilance": 6})
    w = t18.width / 914400
    top, bottom = t18.top / 914400, (_shape(s, "Shape 16").top + _shape(s, "Shape 16").height) / 914400 - 0.08
    total = _len(seg18) + _len(seg19)
    pt = 11.5
    while pt > 8.5 and _hauteur(_len(seg18), w, pt) + _hauteur(_len(seg19), w, pt) + 0.05 > bottom - top:
        pt -= 0.25
    h18 = _hauteur(_len(seg18), w, pt)
    t18.height = Emu(int(h18 * 914400))
    t19.top = Emu(int((top + h18 + 0.04) * 914400))
    t19.height = Emu(int((bottom - top - h18 - 0.04) * 914400))
    _fit(t18, _len(seg18), w, pt, h_in=h18 + 0.5)
    _fit(t19, _len(seg19), w, pt, h_in=bottom - top - h18)
    # Top 5 / Flop 5
    _set_text(_shape(s, "Text 21"), f"TOP 5 & FLOP 5 - {date_fr(d)}")
    tops = ins.top(5)
    for i in range(5):
        c = tops[i] if i < len(tops) else None
        _table_row(_shape(s, f"Table {i + 1}"), [str(i + 1), _court(c.nom), c.secteur, pct(c.ytd)] if c else ["", "", "", ""])
    flops = ins.flop(5)
    t9 = _shape(s, "Table 9")
    if len(flops) >= 5 and t9 is not None and _shape(s, "Table 10") is None:
        new = copy.deepcopy(t9._element)
        t9._element.addnext(new)
        tbl10 = s.shapes[-1] if False else None
        from pptx.shapes.graphfrm import GraphicFrame
        tbl10 = GraphicFrame(new, t9._parent)
        tbl10.name = "Table 10"
        tbl10.top = Emu(t9.top + int(0.27 * 914400))
        fond = _shape(s, "Shape 20")
        bas = tbl10.top + tbl10.height + int(0.06 * 914400)
        if fond is not None and fond.top + fond.height < bas:
            fond.height = Emu(bas - fond.top)
        ids = [sh.shape_id for sh in s.shapes]
        new.nvGraphicFramePr.cNvPr.id = max(ids) + 1
    for i in range(5):
        sh = _shape(s, f"Table {i + 6}")
        c = flops[i] if i < len(flops) else None
        if sh is not None:
            _table_row(sh, [str(i + 1), _court(c.nom), c.secteur, pct(c.ytd)] if c else ["", "", "", ""])
    # Perspectives
    t25 = _shape(s, "Text 25")
    seg25 = [("label", f"{textes['perspectives_titre']} · "), ("normal", textes["perspectives"])]
    _set_runs(t25, seg25, {"label": 0, "normal": 1})
    _fit(t25, _len(seg25), t25.width / 914400, 11.5, min_pt=9.0, h_in=0.86)
    _set_text(_shape(s, "Text 27"), "Sources : " + ", ".join(dict.fromkeys(sources)))
    # Notes de l'orateur
    if s.has_notes_slide:
        notes = [f"Données arrêtées au {d:%d/%m/%Y}."]
        notes += ["".join(t for _, t in textes["evolution"]), "".join(t for _, t in textes["fondamentaux"]),
                  textes["perspectives"]]
        notes += [f"Actualité : {a}" for a in textes.get("actualites", [])]
        s.notes_slide.notes_text_frame.text = "\n".join(notes)


# ----- diapositives 2 à 4 ---------------------------------------------------------------------------------
def _rgb(h):
    from pptx.dml.color import RGBColor
    return RGBColor.from_string(h)


def _tb(slide, x, y, w, h, text="", size=12, bold=False, color=INK, font="Calibri", italic=False, align=None,
        anchor=None):
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Inches, Pt
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if anchor == "middle":
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    lines = text if isinstance(text, list) else [text]
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if align == "center":
            p.alignment = PP_ALIGN.CENTER
        elif align == "right":
            p.alignment = PP_ALIGN.RIGHT
        runs = ln if isinstance(ln, list) else [(ln, {})]
        for txt, opt in runs:
            r = p.add_run()
            r.text = txt
            f = r.font
            f.size = Pt(opt.get("size", size))
            f.bold = opt.get("bold", bold)
            f.italic = opt.get("italic", italic)
            f.name = opt.get("font", font)
            f.color.rgb = _rgb(opt.get("color", color))
        p.space_after = Pt(opt.get("after", 4) if runs else 4)
    return box


def _rect(slide, x, y, w, h, fill, line=None, rounded=True):
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches, Pt
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
                                 Inches(x), Inches(y), Inches(w), Inches(h))
    if rounded:
        shp.adjustments[0] = 0.04
    shp.fill.solid()
    shp.fill.fore_color.rgb = _rgb(fill)
    if line:
        shp.line.color.rgb = _rgb(line)
        shp.line.width = Pt(0.75)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def _entete(slide, titre, sous_titre):
    _tb(slide, 1.16, 0.18, 11.65, 0.62, titre, size=30, bold=True, color=NAVY, font="Cambria")
    _tb(slide, 1.2, 0.74, 11.63, 0.34, sous_titre, size=14, italic=True, color=BLUE)


def _pied(slide, texte, page):
    _tb(slide, 0.52, 7.12, 10.8, 0.25, texte, size=8.5, italic=True, color=MUTED)
    _tb(slide, 12.2, 7.12, 0.65, 0.25, str(page), size=9, color=MUTED, align="right")


def _blank(prs):
    lay = next((l for l in prs.slide_layouts if l.name.lower() in ("vide", "blank")), prs.slide_layouts[-1])
    return prs.slides.add_slide(lay)


def _table(slide, x, y, w, rows, widths, head, row_h=0.3, size=11, dark=True, first_bold=False, colors=None):
    from pptx.util import Inches, Pt
    shp = slide.shapes.add_table(len(rows) + 1, len(head), Inches(x), Inches(y), Inches(w),
                                 Inches(row_h * (len(rows) + 1)))
    tbl = shp.table
    tblPr = tbl._tbl.tblPr
    tblPr.set("firstRow", "0")
    tblPr.set("bandRow", "0")
    style = tblPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}tableStyleId")
    if style is not None:
        style.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"   # sans style
    for j, wd in enumerate(widths):
        tbl.columns[j].width = Inches(wd)
    for i in range(len(rows) + 1):
        tbl.rows[i].height = Inches(row_h)
        for j in range(len(head)):
            c = tbl.cell(i, j)
            c.margin_left = c.margin_right = Inches(0.06)
            c.margin_top = c.margin_bottom = Inches(0.02)
            c.fill.solid()
            if i == 0:
                c.fill.fore_color.rgb = _rgb(NAVY2 if dark else NAVY)
            else:
                c.fill.fore_color.rgb = _rgb(("0E3A5C" if i % 2 else "0B324F") if dark else ("FFFFFF" if i % 2 else LIGHT))
            txt = head[j] if i == 0 else rows[i - 1][j]
            p = c.text_frame.paragraphs[0]
            from pptx.enum.text import PP_ALIGN
            p.alignment = PP_ALIGN.LEFT if j <= 1 else PP_ALIGN.RIGHT
            r = p.add_run()
            r.text = str(txt)
            r.font.size = Pt(size - (1 if i == 0 else 0))
            r.font.name = "Calibri"
            r.font.bold = i == 0 or (first_bold and j == 1)
            col = ("BFD3E2" if dark else "FFFFFF") if i == 0 else ("FFFFFF" if dark else INK)
            if i > 0 and colors and colors.get((i - 1, j)):
                col = colors[(i - 1, j)]
            r.font.color.rgb = _rgb(col)
            _bordures(c, ("0B324F" if dark else "D5DEE7"))
    return shp


def _bordures(cell, couleur):
    from pptx.oxml.ns import qn
    from lxml import etree
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for old in tcPr.findall(qn(tag)):
            tcPr.remove(old)
        ln = etree.SubElement(tcPr, qn(tag), w="6350")
        fill = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(fill, qn("a:srgbClr"), val=couleur)
        tcPr.insert(0, ln)


def _chart_bar(slide, x, y, w, h, cats, series, colors, title="", pct_axis=False, horizontal=False, vary=None):
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
    from pptx.util import Inches, Pt
    cd = CategoryChartData()
    cd.categories = cats
    for name, vals in series:
        cd.add_series(name, vals)
    kind = XL_CHART_TYPE.BAR_CLUSTERED if horizontal else XL_CHART_TYPE.COLUMN_CLUSTERED
    gf = slide.shapes.add_chart(kind, Inches(x), Inches(y), Inches(w), Inches(h), cd)
    ch = gf.chart
    ch.has_title = bool(title)
    if title:
        ch.chart_title.text_frame.text = title
        tr = ch.chart_title.text_frame.paragraphs[0].runs[0].font
        tr.size, tr.bold, tr.color.rgb = Pt(12), True, _rgb(NAVY)
    ch.has_legend = len(series) > 1
    if ch.has_legend:
        ch.legend.position = XL_LEGEND_POSITION.TOP
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(10)
        ch.legend.font.color.rgb = _rgb(INK)
    for k, s in enumerate(ch.series):
        s.format.fill.solid()
        s.format.fill.fore_color.rgb = _rgb(colors[k % len(colors)])
        s.invert_if_negative = False
        if vary and k == 0:
            for i, col in enumerate(vary):
                pt = s.points[i]
                pt.format.fill.solid()
                pt.format.fill.fore_color.rgb = _rgb(col)
        s.data_labels.show_value = True
        s.data_labels.font.size = Pt(9)
        s.data_labels.font.color.rgb = _rgb(INK)
        s.data_labels.number_format = '0%' if pct_axis else '0'
        s.data_labels.number_format_is_linked = False
        s.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    plot = ch.plots[0]
    plot.gap_width = 60
    va, ca = ch.value_axis, ch.category_axis
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = _rgb("D5DEE7")
    va.format.line.fill.background()
    va.tick_labels.font.size = Pt(9)
    va.tick_labels.font.color.rgb = _rgb(MUTED)
    if pct_axis:
        va.tick_labels.number_format = '0%'
        va.tick_labels.number_format_is_linked = False
    from pptx.enum.chart import XL_TICK_LABEL_POSITION
    ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    ca.tick_labels.font.size = Pt(9.5)
    ca.tick_labels.font.color.rgb = _rgb(INK)
    ca.format.line.color.rgb = _rgb("BFD3E2")
    return gf


def slide_signaux(prs, analyses, ins, page):
    s = _blank(prs)
    d = ins.date
    _entete(s, "Les signaux de la cote", f"{len(analyses)} valeurs analysées par les fiches de valeur au {d:%d/%m/%Y}")
    sig = ["ACHAT", "CONSERVER", "ATTENDRE", "ALLÉGER", "VENDRE"]
    cnt = [sum(1 for a in analyses if a.decision.signal == x) for x in sig]
    _rect(s, 0.5, 1.25, 5.0, 4.65, "FFFFFF", line="D5DEE7")
    _chart_bar(s, 0.65, 1.35, 4.7, 4.45, [x.capitalize() for x in sig], [("Titres", cnt)], [NAVY],
               title="Répartition des signaux", vary=[SIGNAL_COL[x] for x in sig])
    # meilleures opportunités
    cand = [a for a in analyses if is_num(a.valuation.upside) and a.decision.signal in ("ACHAT", "CONSERVER", "ATTENDRE")]
    cand.sort(key=lambda a: a.valuation.upside, reverse=True)
    cand = [a for a in cand if a.valuation.upside > 0] or cand[:3]
    rows, colors = [], {}
    for i, a in enumerate(cand[:9]):
        v = a.valuation
        rows.append([a.company.ticker, _court(a.company.name)[:22], num(v.price), num(v.fair_value), pct(v.upside, 1),
                     pct(v.div_yield, 1, sign=False) if is_num(v.div_yield) else "n.d.", a.decision.signal.capitalize()])
        colors[(i, 4)] = "6EE7A8" if v.upside > 0 else "F4A6A0"
        colors[(i, 6)] = {"ACHAT": "6EE7A8", "CONSERVER": "9CC9F0", "ATTENDRE": "F2C66D"}.get(a.decision.signal, "FFFFFF")
    _rect(s, 5.75, 1.25, 7.08, 4.65, NAVY2)
    _tb(s, 5.95, 1.38, 6.7, 0.35, "MEILLEURES OPPORTUNITÉS", size=14.5, bold=True, color="FFFFFF", font="Cambria")
    if rows:
        _table(s, 5.95, 1.85, 6.68, rows, [0.7, 1.8, 0.85, 1.0, 0.85, 0.75, 0.73],
               ["Ticker", "Valeur", "Cours", "Juste val.", "Potentiel", "Rdt", "Signal"], row_h=0.38, size=10.5,
               colors=colors)
    else:
        _tb(s, 5.95, 2.0, 6.6, 0.5, "Aucune valeur en décote parmi les signaux d'achat ou de conservation.",
            size=12, color="FFFFFF")
    # message clé
    up = [a.valuation.upside for a in analyses if is_num(a.valuation.upside)]
    med = statistics.median(up) if up else None
    n_dec = sum(1 for x in up if x > 0.15)
    n_sur = sum(1 for x in up if x < -0.15)
    _rect(s, 0.5, 6.1, 12.33, 0.85, NAVY2)
    _tb(s, 0.78, 6.2, 11.8, 0.68, [[("Lecture · ", {"bold": True, "color": GREEN}),
                                    (f"potentiel médian de la cote {pct(med, 1)} ; {n_dec} valeurs en décote de plus de 15% "
                                     f"contre {n_sur} en surcote de plus de 15%. Les signaux combinent valorisation, qualité, "
                                     f"croissance, tendance, environnement et actualités ; ils sont plafonnés lorsque les "
                                     f"données sont incomplètes.", {"color": "E0EAF2"})]], size=11.5, anchor="middle")
    _pied(s, "Sources : fiches de valeur CGF Gestion (moteur commun de valorisation), Base de données SMF", page)


def slide_secteurs(prs, analyses, ins, page):
    s = _blank(prs)
    _entete(s, "Performance et valorisation par secteur",
            "Hausse depuis janvier et potentiel médian vers la juste valeur")
    par_sect = {}
    for a in analyses:
        lab = SECTEURS_REF.get(a.company.sector, a.company.sector)
        par_sect.setdefault(lab, []).append(a)
    perf = {}
    for c in ins.titres:
        if is_num(c.ytd):
            perf.setdefault(c.secteur, []).append(c.ytd)
    rows, cats, s_ytd, s_up = [], [], [], []
    for lab, lst in sorted(par_sect.items(), key=lambda kv: -len(kv[1])):
        ups = [a.valuation.upside for a in lst if is_num(a.valuation.upside)]
        pes = [a.valuation.pe for a in lst if is_num(a.valuation.pe) and 0 < a.valuation.pe < 80]
        dys = [a.valuation.div_yield for a in lst if is_num(a.valuation.div_yield)]
        ytd = statistics.median(perf[lab]) if perf.get(lab) else None
        sig = max(("ACHAT", "CONSERVER", "ATTENDRE", "ALLÉGER", "VENDRE"),
                  key=lambda x: sum(1 for a in lst if a.decision.signal == x))
        cats.append(lab)
        s_ytd.append(ytd or 0)
        s_up.append(statistics.median(ups) if ups else 0)
        rows.append([lab, str(len(lst)), pct(ytd, 0), pct(statistics.median(ups), 0) if ups else "n.d.",
                     f"{statistics.median(pes):.1f}x".replace(".", ",") if pes else "n.d.",
                     pct(statistics.median(dys), 1, sign=False) if dys else "n.d.", sig.capitalize()])
    _rect(s, 0.5, 1.25, 7.1, 4.65, "FFFFFF", line="D5DEE7")
    if cats:
        _chart_bar(s, 0.62, 1.35, 6.86, 4.45, cats, [("Hausse depuis janvier (médiane)", s_ytd),
                                                     ("Potentiel vers la juste valeur (médiane)", s_up)],
                   [NAVY, GOLD], pct_axis=True)
    _rect(s, 7.83, 1.25, 5.0, 4.65, LIGHT, line="D5DEE7")
    _tb(s, 8.03, 1.38, 4.6, 0.35, "SYNTHÈSE SECTORIELLE", size=14.5, bold=True, color=NAVY, font="Cambria")
    if rows:
        _table(s, 8.03, 1.85, 4.6, [[r[0], r[1], r[4], r[5], r[6]] for r in rows], [1.55, 0.4, 0.75, 0.75, 1.15],
               ["Secteur", "Nb", "PER", "Rdt", "Signal dominant"], row_h=0.36, size=10, dark=False)
    best = max(zip(cats, s_ytd, s_up), key=lambda x: x[1]) if cats else None
    cheap = max(zip(cats, s_ytd, s_up), key=lambda x: x[2]) if cats else None
    msg = ""
    if best and cheap:
        msg = f"{best[0]} mène la hausse ({pct(best[1], 0)} en médiane) avec un potentiel médian de {pct(best[2], 0)} ; "
        msg += (f"{cheap[0]} offre le potentiel médian le plus élevé ({pct(cheap[2], 0)})." if cheap[2] > 0 else
                f"aucun secteur n'offre de potentiel médian positif, le moins surévalué étant {cheap[0].lower()} "
                f"({pct(cheap[2], 0)}).")
        msg += " Un écart marqué entre hausse et potentiel signale une revalorisation déjà avancée."
    _rect(s, 0.5, 6.1, 12.33, 0.85, NAVY2)
    _tb(s, 0.78, 6.2, 11.8, 0.68, [[("Lecture · ", {"bold": True, "color": GREEN}), (msg, {"color": "E0EAF2"})]],
        size=11.5, anchor="middle")
    _pied(s, "Sources : fiches de valeur CGF Gestion, BRVM (BOC), Base de données SMF. PER et rendement : médianes "
             "des titres analysés.", page)


def slide_convictions(prs, analyses, ins, actus, page):
    s = _blank(prs)
    _entete(s, "Convictions et points de vigilance", "Sélection issue des fiches de valeur et actualité du marché")
    ordre = {"Forte": 0, "Moyenne": 1, "Faible": 2}
    conv = [a for a in analyses if a.decision.signal == "ACHAT" and is_num(a.valuation.upside)]
    if len(conv) < 3:
        conv += [a for a in analyses if a.decision.signal == "CONSERVER" and is_num(a.valuation.upside)
                 and a.valuation.upside > 0 and a not in conv]
    conv.sort(key=lambda a: (a.decision.signal != "ACHAT", ordre.get(a.decision.conviction, 3),
                             -(a.valuation.upside or 0)))
    _tb(s, 0.5, 1.22, 6.2, 0.35, "CONVICTIONS", size=14.5, bold=True, color=NAVY, font="Cambria")
    y = 1.65
    for a in conv[:3]:
        v = a.valuation
        forces = [clean_text(x).rstrip(".") for x in (a.texts.get("forces") or [])][:2]
        vig = [clean_text(x).rstrip(".") for x in (a.texts.get("vigilances") or [])
               if not x.startswith(("Qualité des données", "Hors périmètre Charia"))][:1]
        these = (("Atouts : " + " ; ".join(forces) + ".") if forces else "") + \
                ((" Vigilance : " + vig[0] + ".") if vig else "")
        these = these.strip()
        if len(these) > 175:
            these = these[:172].rsplit(" ", 1)[0] + "..."
        _rect(s, 0.5, y, 6.2, 1.28, "FFFFFF", line="D5DEE7")
        _tb(s, 0.7, y + 0.1, 4.2, 0.3, f"{_court(a.company.name)} ({a.company.ticker})", size=13, bold=True, color=NAVY)
        _tb(s, 4.9, y + 0.1, 1.6, 0.3, a.decision.signal.capitalize() + " · " + a.decision.conviction.lower(),
            size=11, bold=True, color=SIGNAL_COL.get(a.decision.signal, NAVY), align="right")
        _tb(s, 0.7, y + 0.43, 5.8, 0.25, f"Cours {num(v.price)} · juste valeur {num(v.fair_value)} · potentiel "
                                         f"{pct(v.upside, 1)} · rendement {pct(v.div_yield, 1, sign=False)}",
            size=10.5, color=MUTED)
        _tb(s, 0.7, y + 0.72, 5.8, 0.5, these, size=10.5, color=INK)
        y += 1.42
    if not conv:
        _tb(s, 0.7, 1.8, 5.8, 0.5, "Aucun signal d'achat à la date d'arrêté.", size=12, color=INK)
    # vigilance
    _rect(s, 7.0, 1.22, 5.83, 2.75, NAVY2)
    _tb(s, 7.2, 1.32, 5.4, 0.35, "VIGILANCE", size=14.5, bold=True, color="FFFFFF", font="Cambria")
    spec = speculatifs(ins, analyses, 4)
    lignes = [[(f"{_court(c.nom)} {pct(c.ytd, 0)} · ", {"bold": True, "color": RED}),
               (", ".join(m) + ".", {"color": "E0EAF2"})] for c, _, m in spec]
    if not lignes:
        surc = sorted([a for a in analyses if is_num(a.valuation.upside)], key=lambda a: a.valuation.upside)[:4]
        lignes = [[(f"{_court(a.company.name)} · ", {"bold": True, "color": RED}),
                   (f"potentiel {pct(a.valuation.upside, 0)}, signal {a.decision.signal.lower()}.", {"color": "E0EAF2"})]
                  for a in surc]
    _tb(s, 7.2, 1.75, 5.45, 2.1, lignes, size=10.5)
    # actualités
    _rect(s, 7.0, 4.12, 5.83, 2.83, LIGHT, line="D5DEE7")
    _tb(s, 7.2, 4.22, 5.4, 0.35, "ACTUALITÉ DU MARCHÉ", size=14.5, bold=True, color=NAVY, font="Cambria")
    if actus:
        lignes = [[(f"{a.date:%d/%m} · " if a.date else "", {"bold": True, "color": BLUE}),
                   (a.titre[:120], {"color": INK}), (f" ({a.source})" if a.source else "", {"color": MUTED, "italic": True})]
                  for a in actus[:5]]
    else:
        lignes = [[("Actualités en ligne non consultées (mode hors ligne ou source indisponible).", {"color": MUTED})]]
    _tb(s, 7.2, 4.65, 5.45, 2.2, lignes, size=10)
    _pied(s, "Sources : fiches de valeur CGF Gestion, Google Actualités" + (", " + ", ".join(
        dict.fromkeys(a.source for a in actus[:5] if a.source)) if actus else ""), page)


def build_market_deck(analyses: list, as_of: date, mb=None, online: bool = True, template: bytes | None = None,
                      api_key: str = "", model: str = "claude-sonnet-5-5", errors: list[str] | None = None,
                      boc: Boc | None = None, boc_fin: Boc | None = None, actus: list[Actu] | None = None,
                      progress=None) -> tuple[bytes, dict]:
    """Présentation « Contexte de marché actions » (4 diapositives). Renvoie le fichier et un résumé des sources."""
    from pptx import Presentation
    errors = errors if errors is not None else []

    def step(m):
        if progress:
            progress(m)
    if online and boc is None:
        step("Bulletin Officiel de la Cote")
        boc = fetch_boc(as_of, errors)
        step("BOC de fin d'année précédente")
        boc_fin = fetch_boc(date(as_of.year - 1, 12, 31), errors) if boc is not None else None
    ins = instantane(as_of, mb, boc, boc_fin)
    if actus is None:
        step("Actualités du marché")
        actus = actualites_marche(ins.date, errors=errors) if online else []
    step("Rédaction")
    textes = (textes_claude(ins, analyses, actus, api_key, model) if api_key else None) or textes_regles(ins, analyses, actus)
    sources = list(ins.sources) + ([] if any(x.startswith("Base de données SMF") for x in ins.sources)
                                   else ["Base de données SMF"]) + ["fiches de valeur CGF Gestion"]
    sources += [a.source for a in actus[:4] if a.source]
    prs = Presentation(io.BytesIO(template) if template else str(TEMPLATE))
    step("Diapositive de marché")
    slide_marche(prs, ins, textes, sources)
    step("Diapositives des fiches de valeur")
    if analyses:
        slide_signaux(prs, analyses, ins, 2)
        slide_secteurs(prs, analyses, ins, 3)
        slide_convictions(prs, analyses, ins, actus, 4)
    prs.core_properties.title = f"Contexte de marché actions BRVM - {ins.date:%d/%m/%Y}"
    buf = io.BytesIO()
    prs.save(buf)
    info = {"date": ins.date, "boc": bool(boc), "boc_url": boc.url if boc else "", "actualites": len(actus),
            "redaction": textes.get("redaction", "règles"), "sources": sources}
    return buf.getvalue(), info
