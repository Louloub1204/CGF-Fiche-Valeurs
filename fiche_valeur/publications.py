"""Publications officielles des sociétés cotées (richbourse.com, relais des avis BRVM).

Une fois un titre choisi :

1. les pages de la rubrique « États financiers » sont parcourues (rapports d'activités trimestriels et
   semestriels, états financiers annuels certifiés) et les publications de la société sont repérées par
   leur intitulé (« ORANGE CI : Rapport d'activités - 1er semestre 2026 ») ;
2. le PDF de la publication la plus récente est téléchargé et son texte extrait ;
3. les lignes du tableau des indicateurs (chiffre d'affaires ou PNB, EBE, résultat d'exploitation,
   résultat des activités ordinaires, coût du risque, résultat net...) sont lues avec la période N et
   la période N-1, ce qui donne les soldes intermédiaires de l'exercice en cours.

Les mêmes analyseurs servent pour un PDF chargé à la main ou une adresse de document collée dans
l'application, ce qui permet de contourner une indisponibilité du site.
"""
from __future__ import annotations

import io
import re
import time
from dataclasses import dataclass, field
from datetime import date

from . import config
from .utils import is_num, norm_label, strip_accents, to_float

BASE = "https://www.richbourse.com"
LIST_URL = BASE + "/common/actualite-categorie/index/etats-financiers?page={page}"
MAX_PDF_BYTES = 20 * 1024 * 1024


# ---------------------------------------------------------------------------
# Modèles
# ---------------------------------------------------------------------------
@dataclass
class Publication:
    title: str
    published: date | None
    pdf_url: str
    detail_url: str = ""
    company: str = ""          # préfixe de l'intitulé (« ORANGE CI »)
    kind: str = ""             # "interim" | "annuel" | "autre"
    period: str = ""           # "S1 2026", "T3 2026", "Exercice 2025"


@dataclass
class ReportLine:
    key: str
    label: str
    current: float | None      # FCFA
    previous: float | None     # FCFA

    @property
    def variation(self) -> float | None:
        if not is_num(self.current) or not is_num(self.previous) or self.previous == 0:
            return None
        if (self.current > 0) != (self.previous > 0):
            return None
        return abs(self.current) / abs(self.previous) - 1


@dataclass
class InterimReport:
    title: str
    period: str
    period_prev: str
    kind: str
    unit_label: str
    lines: list[ReportLine] = field(default_factory=list)
    source: str = ""
    url: str = ""
    published: date | None = None
    notes: list[str] = field(default_factory=list)

    def get(self, key: str) -> ReportLine | None:
        return next((ln for ln in self.lines if ln.key == key), None)

    @property
    def year(self) -> int | None:
        m = re.search(r"(20\d\d)", self.period)
        return int(m.group(1)) if m else None


# ---------------------------------------------------------------------------
# Intitulés -> société et période
# ---------------------------------------------------------------------------
ALIASES = {
    "ABJC": ["SERVAIR ABIDJAN CI", "SERVAIR ABIDJAN"],
    "BBGC": ["BRIDGE BANK GROUP CI", "BRIDGE BANK CI"],
    "BICB": ["BIIC BN", "BIIC BENIN", "BIIC"],
    "BICC": ["BICI CI", "BICICI"],
    "BNBC": ["BERNABE CI"],
    "BOAB": ["BOA BENIN", "BOA BN", "BANK OF AFRICA BENIN"],
    "BOABF": ["BOA BURKINA FASO", "BOA BF", "BANK OF AFRICA BURKINA FASO"],
    "BOAC": ["BOA CI", "BOA COTE D IVOIRE", "BANK OF AFRICA CI"],
    "BOAM": ["BOA MALI", "BOA ML", "BANK OF AFRICA MALI"],
    "BOAN": ["BOA NIGER", "BOA NG", "BOA NE", "BANK OF AFRICA NIGER"],
    "BOAS": ["BOA SENEGAL", "BOA SN", "BANK OF AFRICA SENEGAL"],
    "CABC": ["SICABLE CI"],
    "CBIBF": ["CORIS BANK INTERNATIONAL BF", "CORIS BANK INTERNATIONAL", "CBI BF"],
    "CFAC": ["CFAO MOTORS CI", "CFAO CI"],
    "CIEC": ["CIE CI"],
    "ECOC": ["ECOBANK CI", "ECOBANK COTE D IVOIRE"],
    "ETIT": ["ETI TG", "ECOBANK TRANSNATIONAL INCORPORATED TG", "ECOBANK TRANSNATIONAL INCORPORATED", "ETI"],
    "FTSC": ["FILTISAC CI"],
    "LNBB": ["LNB BN", "LOTERIE NATIONALE DU BENIN"],
    "NEIC": ["NEI CEDA CI", "NEI-CEDA CI"],
    "NSBC": ["NSIA BANQUE CI", "NSIA BANQUE COTE D IVOIRE"],
    "NTLC": ["NESTLE CI", "NESTLE COTE D IVOIRE"],
    "ONTBF": ["ONATEL BF", "ONATEL"],
    "ORAC": ["ORANGE CI", "ORANGE COTE D IVOIRE"],
    "ORGT": ["ORAGROUP TG", "ORAGROUP"],
    "PALC": ["PALM CI"],
    "PRSC": ["TRACTAFRIC MOTORS CI", "TRACTAFRIC CI"],
    "SAFC": ["SAFCA CI"],
    "SCRC": ["SUCRIVOIRE CI", "SUCRIVOIRE"],
    "SDCC": ["SODE CI", "SODECI CI", "SODECI"],
    "SDSC": ["AFRICA GLOBAL LOGISTICS CI", "AGL CI", "BOLLORE TRANSPORT LOGISTICS CI"],
    "SEMC": ["EVIOSYS PACKAGING SIEM CI", "SIEM CI", "CROWN SIEM CI"],
    "SGBC": ["SGB CI", "SOCIETE GENERALE CI", "SOCIETE GENERALE COTE D IVOIRE", "SGBCI"],
    "SHEC": ["VIVO ENERGY CI", "VIVO ENERGY COTE D IVOIRE"],
    "SIBC": ["SIB CI", "SOCIETE IVOIRIENNE DE BANQUE"],
    "SICC": ["SICOR CI"],
    "SIVC": ["ERIUM CI", "AIR LIQUIDE CI"],
    "SLBC": ["SOLIBRA CI"],
    "SMBC": ["SMB CI"],
    "SNTS": ["SONATEL SN", "SONATEL"],
    "SOGC": ["SOGB CI"],
    "SPHC": ["SAPH CI"],
    "STAC": ["SETAO CI"],
    "STBC": ["SITAB CI"],
    "TTLC": ["TOTALENERGIES MARKETING CI", "TOTAL CI", "TOTALENERGIES CI"],
    "TTLS": ["TOTALENERGIES MARKETING SN", "TOTAL SN", "TOTALENERGIES SN", "TOTALENERGIES MARKETING SENEGAL"],
    "UNLC": ["UNILEVER CI"],
    "UNXC": ["UNIWAX CI"],
}
def _alias_index() -> dict[str, str]:
    """Dénominations connues -> ticker : table ci-dessus, noms et autres noms du référentiel."""
    idx = {norm_label(a): t for t, al in ALIASES.items() for a in al}
    for t, ref in config.TICKERS.items():
        for a in [ref.get("nom", "")] + list(ref.get("autres_noms", [])):
            if a:
                idx.setdefault(norm_label(a), t)
    return idx


def ticker_of_company(prefix: str) -> str | None:
    """Préfixe d'intitulé (« ORANGE CI ») -> ticker."""
    return _alias_index().get(norm_label(prefix))


_ORD = {"1ER": 1, "PREMIER": 1, "1": 1, "2E": 2, "2EME": 2, "DEUXIEME": 2, "2": 2, "3E": 3, "3EME": 3,
        "TROISIEME": 3, "3": 3, "4E": 4, "4EME": 4, "QUATRIEME": 4, "4": 4}


def period_of_title(title: str) -> tuple[str, str]:
    """Intitulé -> (type, période) : ("interim", "S1 2026"), ("annuel", "Exercice 2025")..."""
    n = norm_label(title)
    m = re.search(r"\b(1ER|PREMIER|2E|2EME|DEUXIEME|1|2)\s+SEMESTRE\s+(20\d\d)", n)
    if m:
        return "interim", f"S{_ORD.get(m.group(1), 1)} {m.group(2)}"
    m = re.search(r"\b(1ER|PREMIER|2E|2EME|DEUXIEME|3E|3EME|TROISIEME|4E|4EME|QUATRIEME|1|2|3|4)\s+TRIMESTRE\s+(20\d\d)", n)
    if m:
        return "interim", f"T{_ORD.get(m.group(1), 1)} {m.group(2)}"
    m = re.search(r"EXERCICE\s+(20\d\d)", n)
    if m and ("ETATS FINANCIERS" in n or "RAPPORT ANNUEL" in n or "COMPTES" in n):
        return "annuel", f"Exercice {m.group(1)}"
    if "RAPPORT D ACTIVITE" in n:
        m = re.search(r"(20\d\d)", n)
        return "interim", m.group(1) if m else ""
    return "autre", ""


# ---------------------------------------------------------------------------
# Liste des publications (pages HTML)
# ---------------------------------------------------------------------------
def parse_listing(html: str) -> list[Publication]:
    """Page de liste richbourse -> publications (intitulé, date, lien du PDF)."""
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html, "html.parser")
    pubs: dict[str, Publication] = {}
    for a in soup.find_all("a", href=True):
        href = a["href"]
        m = re.search(r"/common/actualite/(details|afficher-fichier)/((\d{2})-(\d{2})-(\d{4})-[\w\-]+)", href)
        if not m:
            continue
        slug = m.group(2)
        pub = pubs.get(slug)
        if pub is None:
            try:
                d = date(int(m.group(5)), int(m.group(4)), int(m.group(3)))
            except ValueError:
                d = None
            pub = Publication(title="", published=d, pdf_url=f"{BASE}/common/actualite/afficher-fichier/{slug}")
            pubs[slug] = pub
        if m.group(1) == "details":
            pub.detail_url = f"{BASE}/common/actualite/details/{slug}"
        text = a.get_text(" ", strip=True)
        if text and " : " in text and len(text) > len(pub.title):
            pub.title = text
    out = []
    for slug, pub in pubs.items():
        if not pub.title:      # intitulé reconstruit à partir de l'adresse
            pub.title = slug[11:].replace("-", " ").upper()
        pub.company = pub.title.split(" : ")[0].strip() if " : " in pub.title else ""
        pub.kind, pub.period = period_of_title(pub.title)
        out.append(pub)
    out.sort(key=lambda p: p.published or date(1900, 1, 1), reverse=True)
    return out


def _ticker_of_pub(pub: Publication) -> str | None:
    if pub.company:
        t = ticker_of_company(pub.company)
        if t:
            return t
    n = norm_label(pub.title)
    for alias, t in sorted(_alias_index().items(), key=lambda kv: -len(kv[0])):
        if n.startswith(alias + " "):
            return t
    return None


def find_publications(ticker: str, max_pages: int = 15, months: int = 18,
                      errors: list[str] | None = None) -> list[Publication]:
    """Publications de la société (les plus récentes d'abord), sur les pages de liste récentes."""
    from .market import TIMEOUT, _cached, _session

    def page(n: int) -> list[Publication] | None:
        def run():
            try:
                r = _session().get(LIST_URL.format(page=n), timeout=TIMEOUT,
                                   headers={"Referer": BASE + "/common/actualite/index"})
                r.raise_for_status()
            except Exception as exc:
                if errors is not None:
                    errors.append(f"richbourse.com (liste des publications, page {n}) : {type(exc).__name__}"[:200])
                return None
            pubs = parse_listing(r.text)
            if not pubs and errors is not None and n == 1:
                errors.append("richbourse.com : aucune publication reconnue (page protégée ou structure modifiée).")
            return pubs
        return _cached(f"rb:list:{n}", run)

    limit = date.today().toordinal() - months * 31
    found: list[Publication] = []
    for n in range(1, max_pages + 1):
        pubs = page(n)
        if not pubs:
            break
        found += [p for p in pubs if _ticker_of_pub(p) == ticker]
        oldest = min((p.published for p in pubs if p.published), default=None)
        if oldest and oldest.toordinal() < limit:
            break
        kinds = {p.kind for p in found}
        if "annuel" in kinds and "interim" in kinds and len(found) >= 3:
            break
        time.sleep(0.2)
    return found


# ---------------------------------------------------------------------------
# Lecture des PDF
# ---------------------------------------------------------------------------
def pdf_text(data: bytes) -> str:
    """Texte d'un PDF (pdfplumber, sinon pypdf). Chaîne vide si le PDF est une image scannée."""
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            return "\n".join((pg.extract_text() or "") for pg in pdf.pages[:25])
    except Exception:
        pass
    try:
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(data))
        return "\n".join((pg.extract_text() or "") for pg in reader.pages[:25])
    except Exception:
        return ""


LINE_RULES: list[tuple[str, str, str]] = [
    # clé, motif sur le libellé normalisé (début de ligne), libellé affiché
    ("pnb", r"^(PRODUIT NET BANCAIRE|PNB)\b", "Produit net bancaire"),
    ("ca", r"^(CHIFFRE D ?AFFAIRES|REVENUS?( CONSOLIDES?)?$|PRODUITS DES ACTIVITES ORDINAIRES)", "Chiffre d'affaires"),
    ("va", r"^VALEUR AJOUTEE", "Valeur ajoutée"),
    ("ebe", r"^(EBITDA\w*|EBE|EXCEDENT BRUT D ?EXPLOITATION)\b", "EBE / EBITDA"),
    ("charges_generales", r"^(CHARGES GENERALES D ?EXPLOITATION|FRAIS GENERAUX|CHARGES D ?EXPLOITATION)", "Charges d'exploitation"),
    ("rbe", r"^(RESULTAT BRUT D ?EXPLOITATION|RBE)\b", "Résultat brut d'exploitation"),
    ("cout_risque", r"^COUT (NET )?DU RISQUE", "Coût du risque"),
    ("re", r"^RESULTAT (D ?EXPLOITATION|OPERATIONNEL)", "Résultat d'exploitation"),
    ("rf", r"^RESULTAT FINANCIER", "Résultat financier"),
    ("rao", r"^(RESULTAT DES ACTIVITES ORDINAIRES|RAO)\b", "Résultat des activités ordinaires"),
    ("rai", r"^RESULTAT AVANT IMPOTS?", "Résultat avant impôt"),
    ("impots", r"^(IMPOTS? SUR (LE|LES) (RESULTATS?|BENEFICES?)|IMPOTS? SUR RESULTAT)", "Impôt sur le résultat"),
    ("rn", r"^RESULTAT NET\b", "Résultat net"),
]
_NUM = re.compile(r"\(?-?\d{1,3}(?:[   .]\d{3})*(?:,\d+)?\)?%?|\(?-?\d+(?:,\d+)?\)?%?")
_CODE = re.compile(r"^(?:[A-Z]{1,2}\d?\s*[-:.)]\s+|X[A-I]\s+|T[A-N]\s+|R[A-S]\s+)")


def _unit(text: str) -> tuple[float, str]:
    n = strip_accents(text[:6000]).lower()
    if re.search(r"milliards? (de )?(f ?cfa|fcfa|xof|francs)|en milliards", n):
        return 1e9, "milliards FCFA"
    if re.search(r"millions? (de )?(f ?cfa|fcfa|xof|francs)|en millions|en m fcfa|mfcfa|m xof", n):
        return 1e6, "millions FCFA"
    if re.search(r"milliers? (de )?(f ?cfa|fcfa|xof|francs)|en milliers|k ?fcfa", n):
        return 1e3, "milliers FCFA"
    if re.search(r"en (f ?cfa|fcfa|xof)\b", n):
        return 1.0, "FCFA"
    return 0.0, ""


def _numbers(rest: str) -> list[float]:
    vals = []
    for tok in _NUM.findall(rest):
        if tok.endswith("%"):
            continue
        neg = tok.startswith("(") and tok.endswith(")")
        v = to_float(tok.strip("()"))
        if is_num(v):
            vals.append(-abs(v) if neg else v)
    return vals


def parse_report_text(text: str, title: str = "") -> InterimReport | None:
    """Texte d'un rapport d'activités ou d'états financiers -> indicateurs N et N-1."""
    if not text or len(text.strip()) < 50:
        return None
    mult, unit_label = _unit(text)
    kind, period = period_of_title(title) if title else ("autre", "")
    if not period:
        kind, period = period_of_title(text[:1500])
    lines: dict[str, ReportLine] = {}
    raw: list[tuple[str, str, float, float]] = []
    for line in text.splitlines():
        line = line.strip()
        if len(line) < 6:
            continue
        m = re.search(r"\(?-?\d", line)
        if not m:
            continue
        label_raw = line[:m.start()].strip(" .:-|")
        label = _CODE.sub("", norm_label(label_raw))
        if not label or len(label) < 3:
            continue
        nums = _numbers(line[m.start():])
        # années éventuellement collées au libellé (« Résultat net 2026 2025 ... ») : on les ignore
        nums = [v for v in nums if not (1990 <= v <= 2100 and float(v).is_integer() and len(nums) > 2)]
        if len(nums) < 2:
            continue
        for key, pat, lab in LINE_RULES:
            if key in lines:
                continue
            if re.search(pat, label):
                if key == "rn" and re.search(r"PAR ACTION|/|MARGE", label):
                    continue
                raw.append((key, lab, nums[0], nums[1]))
                break
    if not raw:
        return None
    if mult == 0.0:
        big = max(abs(v) for _, _, a, b in raw for v in (a, b))
        mult, unit_label = (1.0, "FCFA (déduit)") if big > 1e8 else (1e6, "millions FCFA (présumés)")
    for key, lab, cur, prev in raw:
        if key not in lines:
            lines[key] = ReportLine(key, lab, cur * mult, prev * mult)
    order = [k for k, _, _ in LINE_RULES]
    rep = InterimReport(title=title, period=period, period_prev=_prev_period(period), kind=kind,
                        unit_label=unit_label, lines=sorted(lines.values(), key=lambda ln: order.index(ln.key)))
    if unit_label.endswith("(présumés)"):
        rep.notes.append("Unité non indiquée dans le document : millions de FCFA présumés.")
    return rep


def _prev_period(period: str) -> str:
    m = re.search(r"(20\d\d)", period)
    return period.replace(m.group(1), str(int(m.group(1)) - 1)) if m else "N-1"


def parse_report_pdf(data: bytes, title: str = "", source: str = "", url: str = "",
                     published: date | None = None) -> InterimReport | None:
    text = pdf_text(data)
    rep = parse_report_text(text, title)
    if rep is None:
        return None
    rep.source, rep.url, rep.published = source, url, published
    return rep


def fetch_report(pub: Publication, errors: list[str] | None = None) -> InterimReport | None:
    from .market import TIMEOUT, _cached, _session

    def run():
        try:
            r = _session().get(pub.pdf_url, timeout=TIMEOUT + 10, headers={"Referer": pub.detail_url or BASE})
            r.raise_for_status()
            data = r.content[:MAX_PDF_BYTES]
        except Exception as exc:
            if errors is not None:
                errors.append(f"richbourse.com (document {pub.period or pub.title[:40]}) : {type(exc).__name__}"[:200])
            return None
        if not data.startswith(b"%PDF"):
            if errors is not None:
                errors.append("richbourse.com : le document n'est pas un PDF (accès protégé ?).")
            return None
        rep = parse_report_pdf(data, pub.title, "richbourse.com (publication officielle)", pub.pdf_url, pub.published)
        if rep is None and errors is not None:
            errors.append(f"Publication « {pub.title[:80]} » : tableau d'indicateurs non reconnu "
                          f"(document scanné ou mise en page atypique).")
        return rep
    return _cached(f"rb:pdf:{pub.pdf_url}", run, ttl=24 * 3600)


def fetch_url_report(url: str, errors: list[str] | None = None) -> InterimReport | None:
    """Document dont l'adresse est collée dans l'application (PDF richbourse, brvm.org, site de la société)."""
    from .market import TIMEOUT, _session
    try:
        r = _session().get(url, timeout=TIMEOUT + 10)
        r.raise_for_status()
    except Exception as exc:
        if errors is not None:
            errors.append(f"Document {url[:60]} : {type(exc).__name__}")
        return None
    if not r.content.startswith(b"%PDF"):
        if errors is not None:
            errors.append("L'adresse fournie ne renvoie pas un PDF.")
        return None
    slug = url.rstrip("/").split("/")[-1]
    title = slug[11:].replace("-", " ") if re.match(r"\d{2}-\d{2}-\d{4}-", slug) else ""
    return parse_report_pdf(r.content, title, "Document en ligne", url)


def latest_report(ticker: str, errors: list[str] | None = None) -> tuple[InterimReport | None, list[Publication]]:
    """Publications de la société et indicateurs de la plus récente exploitable (intérimaire ou annuelle)."""
    pubs = find_publications(ticker, errors=errors)
    usable = [p for p in pubs if p.kind in ("interim", "annuel")]
    # rapport d'activités (texte) avant attestation des commissaires aux comptes
    usable.sort(key=lambda p: (p.published or date(1900, 1, 1), "RAPPORT" in norm_label(p.title)), reverse=True)
    for pub in usable[:3]:
        rep = fetch_report(pub, errors)
        if rep is not None and rep.lines:
            return rep, pubs
    return None, pubs


def momentum(rep: InterimReport | None) -> dict:
    """Croissance de la période en cours (revenus et résultat) tirée d'une publication."""
    if rep is None:
        return {}
    rev = rep.get("pnb") or rep.get("ca")
    res = rep.get("rn") or rep.get("rao") or rep.get("re")
    return {"rev": rev.variation if rev else None, "res": res.variation if res else None,
            "rev_label": rev.label if rev else "", "res_label": res.label if res else ""}
