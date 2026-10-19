"""Fiche de valeur imprimable (PDF A4) : version synthétique (1 page) ou
complète (4 pages)."""
from __future__ import annotations

import io
from datetime import date
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (BaseDocTemplate, CondPageBreak, Frame, Image, KeepTogether, NextPageTemplate,
                                PageBreak, PageTemplate, Paragraph, Spacer, Table, TableStyle)

from . import charts, config
from .fundamentals import RATIO_META, sig_rows
from .utils import (clean_text, fmt_date_fr, fmt_fcfa, fmt_mds, fmt_num, fmt_pct, fmt_price, fmt_pts, fmt_x,
                    growth, is_num)

PAGE_W, PAGE_H = A4
MARGIN = 14 * mm
CONTENT_W = PAGE_W - 2 * MARGIN
NAVY = colors.HexColor(config.NAVY)
GOLD = colors.HexColor(config.GOLD)
INK = colors.HexColor(config.INK)
MUTED = colors.HexColor(config.MUTED)
PAPER = colors.HexColor(config.PAPER)
GRID = colors.HexColor("#DDE1EA")

_FONTS = {"body": "Helvetica", "bold": "Helvetica-Bold", "italic": "Helvetica-Oblique", "title": "Times-Bold"}


def _register_fonts() -> None:
    d = config.FONTS_DIR
    try:
        pdfmetrics.registerFont(TTFont("Carlito", str(d / "Carlito-Regular.ttf")))
        pdfmetrics.registerFont(TTFont("Carlito-Bold", str(d / "Carlito-Bold.ttf")))
        pdfmetrics.registerFont(TTFont("Carlito-Italic", str(d / "Carlito-Italic.ttf")))
        pdfmetrics.registerFont(TTFont("Carlito-BoldItalic", str(d / "Carlito-BoldItalic.ttf")))
        pdfmetrics.registerFontFamily("Carlito", normal="Carlito", bold="Carlito-Bold",
                                      italic="Carlito-Italic", boldItalic="Carlito-BoldItalic")
        _FONTS.update(body="Carlito", bold="Carlito-Bold", italic="Carlito-Italic")
    except Exception:
        pass
    try:
        pdfmetrics.registerFont(TTFont("Caladea-Bold", str(d / "Caladea-Bold.ttf")))
        _FONTS["title"] = "Caladea-Bold"
    except Exception:
        pass


def _styles() -> dict:
    b, bo, it, ti = _FONTS["body"], _FONTS["bold"], _FONTS["italic"], _FONTS["title"]
    return {
        "body": ParagraphStyle("body", fontName=b, fontSize=8.6, leading=11.4, textColor=INK, alignment=TA_LEFT),
        "small": ParagraphStyle("small", fontName=b, fontSize=7.3, leading=9.2, textColor=INK),
        "muted": ParagraphStyle("muted", fontName=it, fontSize=7, leading=8.8, textColor=MUTED),
        "h1": ParagraphStyle("h1", fontName=ti, fontSize=12.5, leading=15, textColor=NAVY, spaceBefore=3, spaceAfter=4),
        "h2": ParagraphStyle("h2", fontName=bo, fontSize=9.4, leading=12, textColor=NAVY, spaceBefore=4, spaceAfter=3),
        "cell": ParagraphStyle("cell", fontName=b, fontSize=7.4, leading=9, textColor=INK),
        "cellb": ParagraphStyle("cellb", fontName=bo, fontSize=7.4, leading=9, textColor=INK),
        "cellr": ParagraphStyle("cellr", fontName=b, fontSize=7.4, leading=9, textColor=INK, alignment=TA_RIGHT),
        "head": ParagraphStyle("head", fontName=bo, fontSize=7.4, leading=9, textColor=colors.white),
        "headr": ParagraphStyle("headr", fontName=bo, fontSize=7.4, leading=9, textColor=colors.white, alignment=TA_RIGHT),
        "bullet": ParagraphStyle("bullet", fontName=b, fontSize=8.2, leading=10.6, textColor=INK, leftIndent=9,
                                 bulletIndent=0, spaceAfter=1.5, bulletFontName=b, bulletFontSize=8.2,
                                 bulletColor=GOLD),
        "cellbr": ParagraphStyle("cellbr", fontName=bo, fontSize=7.4, leading=9, textColor=INK, alignment=TA_RIGHT),
        "sig_label": ParagraphStyle("sig_label", fontName=bo, fontSize=7.5, leading=9, textColor=colors.white),
        "sig": ParagraphStyle("sig", fontName=ti, fontSize=22, leading=25, textColor=colors.white),
        "sig_conv": ParagraphStyle("sig_conv", fontName=b, fontSize=8, leading=10, textColor=colors.white),
        "kpi_v": ParagraphStyle("kpi_v", fontName=bo, fontSize=11, leading=13, textColor=INK, alignment=TA_CENTER),
        "kpi_l": ParagraphStyle("kpi_l", fontName=b, fontSize=6.8, leading=8.4, textColor=MUTED, alignment=TA_CENTER),
    }


def P(text, style) -> Paragraph:
    return Paragraph(escape(clean_text(str(text))), style)


def _img(png: bytes | None, width: float, max_h: float | None = None):
    if not png:
        return None
    from reportlab.lib.utils import ImageReader
    ir = ImageReader(io.BytesIO(png))
    iw, ih = ir.getSize()
    h = width * ih / iw
    if max_h and h > max_h:
        width, h = max_h * iw / ih, max_h
    return Image(io.BytesIO(png), width=width, height=h)


def _table(rows, col_widths, st, header=True, num_cols_from=1, zebra=True, left_cols=(), bold_cols=()):
    data = []
    for i, row in enumerate(rows):
        out = []
        for j, cell in enumerate(row):
            if isinstance(cell, (Paragraph, Image, Table)):
                out.append(cell)
                continue
            right = j >= num_cols_from and j not in left_cols
            if header and i == 0:
                out.append(P(cell, st["headr"] if right else st["head"]))
            elif j in bold_cols:
                out.append(P(cell, st["cellb"]))
            else:
                out.append(P(cell, st["cellr"] if right else st["cell"]))
        data.append(out)
    t = Table(data, colWidths=col_widths, repeatRows=1 if header else 0)
    cmds = [("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 2.2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2.2), ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4), ("LINEBELOW", (0, 0), (-1, -1), 0.4, GRID)]
    if header:
        cmds += [("BACKGROUND", (0, 0), (-1, 0), NAVY)]
    if zebra:
        for i in range(1 if header else 0, len(data)):
            if (i % 2 == 0) == header:
                cmds.append(("BACKGROUND", (0, i), (-1, i), PAPER))
    t.setStyle(TableStyle(cmds))
    return t


# ---------------------------------------------------------------------------
# En-tête et pied de page
# ---------------------------------------------------------------------------

def _decorations(a, total_pages_holder):
    c = a.company
    country = config.COUNTRIES.get(c.country, {}).get("nom", c.country)

    def first(canvas, doc):
        canvas.saveState()
        band = 27 * mm
        canvas.setFillColor(NAVY)
        canvas.rect(0, PAGE_H - band, PAGE_W, band, stroke=0, fill=1)
        canvas.setFillColor(GOLD)
        canvas.rect(0, PAGE_H - band - 1.2 * mm, PAGE_W, 1.2 * mm, stroke=0, fill=1)
        canvas.setFillColor(colors.white)
        canvas.setFont(_FONTS["title"], 19)
        name = clean_text(c.name)
        while canvas.stringWidth(name, _FONTS["title"], 19) > CONTENT_W - 58 * mm and len(name) > 10:
            name = name[:-2]
        canvas.drawString(MARGIN, PAGE_H - 12.5 * mm, name)
        canvas.setFont(_FONTS["body"], 9)
        canvas.drawString(MARGIN, PAGE_H - 19 * mm,
                          clean_text(f"{c.ticker}  |  {c.sector}  |  {country}  |  BRVM"))
        canvas.setFillColor(GOLD)
        canvas.setFont(_FONTS["bold"], 8.5)
        canvas.drawRightString(PAGE_W - MARGIN, PAGE_H - 11 * mm, "FICHE DE VALEUR")
        canvas.setFillColor(colors.white)
        canvas.setFont(_FONTS["body"], 8.5)
        canvas.drawRightString(PAGE_W - MARGIN, PAGE_H - 16.5 * mm, fmt_date_fr(a.as_of))
        canvas.drawRightString(PAGE_W - MARGIN, PAGE_H - 21.5 * mm, clean_text(a.settings.maison))
        _footer(canvas, doc)
        canvas.restoreState()

    def later(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(NAVY)
        canvas.rect(0, PAGE_H - 11 * mm, PAGE_W, 11 * mm, stroke=0, fill=1)
        canvas.setFillColor(GOLD)
        canvas.rect(0, PAGE_H - 11.6 * mm, PAGE_W, 0.6 * mm, stroke=0, fill=1)
        canvas.setFillColor(colors.white)
        canvas.setFont(_FONTS["bold"], 9.5)
        canvas.drawString(MARGIN, PAGE_H - 7.2 * mm, clean_text(f"{c.name} ({c.ticker})"))
        canvas.setFont(_FONTS["body"], 8.5)
        canvas.drawRightString(PAGE_W - MARGIN, PAGE_H - 7.2 * mm,
                               clean_text(f"Signal : {a.decision.signal}  |  {fmt_date_fr(a.as_of)}"))
        _footer(canvas, doc)
        canvas.restoreState()

    def _footer(canvas, doc):
        canvas.setStrokeColor(GRID)
        canvas.setLineWidth(0.5)
        canvas.line(MARGIN, 11 * mm, PAGE_W - MARGIN, 11 * mm)
        canvas.setFont(_FONTS["body"], 6.8)
        canvas.setFillColor(MUTED)
        canvas.drawString(MARGIN, 7 * mm, clean_text(
            f"{a.settings.maison}  |  Document interne d'aide à la décision, ne constitue pas un conseil en investissement"))
        canvas.drawRightString(PAGE_W - MARGIN, 7 * mm, f"Page {doc.page}")

    return first, later


# ---------------------------------------------------------------------------
# Blocs
# ---------------------------------------------------------------------------

def _signal_box(a, st) -> Table:
    d = a.decision
    col = colors.HexColor(config.SIGNAL_COLORS.get(d.signal, config.NAVY))
    inner = [[P("SIGNAL", st["sig_label"])], [P(d.signal, st["sig"])],
             [P(f"Conviction {d.conviction.lower()}  |  horizon {a.settings.horizon_mois} mois", st["sig_conv"])],
             [P(f"Note globale {fmt_num(d.composite)}/100" if is_num(d.composite) else "Note globale n.d.",
                st["sig_conv"])]]
    t = Table(inner, colWidths=[56 * mm])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), col), ("LEFTPADDING", (0, 0), (-1, -1), 7),
                           ("TOPPADDING", (0, 0), (-1, -1), 1.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
                           ("TOPPADDING", (0, 0), (0, 0), 6), ("BOTTOMPADDING", (0, -1), (0, -1), 6)]))
    return t


def _kpis(a, st) -> Table:
    v, d = a.valuation, a.decision
    items = [
        (fmt_price(v.price), f"Cours au {a.market.price_date.strftime('%d/%m/%Y') if a.market.price_date else 'n.d.'}"),
        (fmt_price(d.target), "Juste valeur / cible 12 mois"),
        (fmt_pct(d.upside, sign=True), "Potentiel"),
        (fmt_pct(v.div_yield), "Rendement net"),
    ]
    row_v = [P(x, st["kpi_v"]) for x, _ in items]
    row_l = [P(y, st["kpi_l"]) for _, y in items]
    t = Table([row_v, row_l], colWidths=[(CONTENT_W - 60 * mm) / 4] * 4)
    t.setStyle(TableStyle([("LINEAFTER", (0, 0), (-2, -1), 0.5, GRID), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                           ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]))
    return t


def _keydata(a, st) -> Table:
    v, m, t, fd = a.valuation, a.market, a.technical, a.fundamentals
    y = fd.year
    liq = t.liquidity if t.available else {}
    pairs = [
        ("Capitalisation", fmt_fcfa(m.market_cap)), ("PER " + str(y or ""), fmt_x(v.pe)),
        ("Nombre de titres", fmt_num(m.shares)), ("P/B", fmt_x(v.pb, 2)),
        ("Flottant", fmt_pct(m.float_pct, 0)), ("ROE " + str(y or ""), fmt_pct(fd.ratios.get("roe", {}).get(y))),
        ("BPA " + str(y or ""), fmt_price(v.eps)), ("DPA net", fmt_price(v.dps) + (f" ({v.dps_exercice})" if v.dps_exercice else "")),
        ("Perf. depuis le 1er janv.", fmt_pct(t.performance.get("ytd"), sign=True) if t.available else "n.d."),
        ("Perf. 12 mois", fmt_pct(t.performance.get("12m"), sign=True) if t.available else "n.d."),
        ("Plus haut / bas 52 sem.", (fmt_num(t.levels.get("plus_haut_52s")) + " / " + fmt_num(t.levels.get("plus_bas_52s")))
         if t.available and t.levels else "n.d."),
        ("Volume moyen 3 mois", (fmt_num(liq.get("volume_moyen_3m")) + " titres") if is_num(liq.get("volume_moyen_3m")) else "n.d."),
        ("Bêta (BRVM-C)", fmt_num(v.beta, 2)), ("Séances actives 3 mois", fmt_pct(liq.get("taux_seances_actives"), 0)),
    ]
    rows = []
    for i in range(0, len(pairs), 2):
        (l1, v1), (l2, v2) = pairs[i], pairs[i + 1]
        rows.append([l1, v1, l2, v2])
    w = CONTENT_W - 60 * mm
    return _table(rows, [w * 0.27, w * 0.23, w * 0.27, w * 0.23], st, header=False, left_cols=(2,))


def _bullets(items, st, color=None) -> list:
    out = []
    for it in items:
        out.append(Paragraph(escape(clean_text(it)), st["bullet"], bulletText="▪"))
    return out


def _two_cols(left, right, gap=6 * mm, ratio=0.5):
    w1 = (CONTENT_W - gap) * ratio
    w2 = CONTENT_W - gap - w1
    t = Table([[left, "", right]], colWidths=[w1, gap, w2])
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                           ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0)]))
    return t


def _sig_table(a, st):
    c = a.company
    years = c.years[-5:]
    rows = [["En Mds FCFA"] + [str(y) for y in years] + [f"Var. {str(years[-1])[2:]}/{str(years[-2])[2:]}" if len(years) > 1 else "Var."]]
    for key, label, s in sig_rows(c):
        rows.append([label] + [fmt_mds(s.get(y)) for y in years] + [_var_txt(s, years)])
    rev = c.fields.get(c.revenue_field, {})
    for key, lab in (("ebe", "Marge d'EBE"), ("re", "Marge opérationnelle"), (c.net_income_field, "Marge nette")):
        s = c.fields.get(key, {})
        if not s or (c.is_bank and key == "ebe"):
            continue
        vals = [fmt_pct(s.get(y) / rev.get(y)) if is_num(s.get(y)) and is_num(rev.get(y)) and rev.get(y) else "n.d."
                for y in years]
        rows.append([lab + (" (% du PNB)" if c.is_bank else " (% du CA)")] + vals + [""])
    n = len(years)
    first = 58 * mm
    other = (CONTENT_W - first) / (n + 1)
    return _table(rows, [first] + [other] * (n + 1), st)


def _var_txt(s, years) -> str:
    """Variation du dernier exercice ; charges comparées en valeur absolue, n.s. si changement de signe."""
    if len(years) < 2:
        return ""
    new, old = s.get(years[-1]), s.get(years[-2])
    if not is_num(new) or not is_num(old) or old == 0:
        return "n.d."
    if (new > 0) != (old > 0):
        return "n.s."
    return fmt_pct(abs(new) / abs(old) - 1, sign=True)


RATIO_ORDER_BANK = ["croissance_pnb", "coef_exploitation", "cout_risque_pnb", "cout_risque_credits", "marge_nette",
                    "roe", "roa", "autonomie", "ltd", "croissance_credits", "croissance_depots", "poids_marge_interets",
                    "npl", "taux_impot"]
RATIO_ORDER_CORP = ["croissance_ca", "taux_va", "marge_ebe", "marge_op", "marge_nette", "roe", "roa", "poids_personnel",
                    "autonomie", "gearing", "dette_nette_ebe", "couverture_ff", "liquidite_generale", "bfr_jours",
                    "rotation_actif", "taux_impot"]


def _fmt_ratio(key, v):
    kind = RATIO_META.get(key, ("", "pct", ""))[1]
    if not is_num(v):
        return "n.d."
    if kind == "pct":
        return fmt_pct(v)
    if kind == "x":
        return fmt_num(v, 1) + "x"
    if kind == "jours":
        return fmt_num(v, 0) + " j"
    return fmt_num(v, 2)


def _ratio_table(a, st):
    c, r = a.company, a.fundamentals.ratios
    years = c.years[-5:]
    order = RATIO_ORDER_BANK if c.is_bank else RATIO_ORDER_CORP
    med = a.fundamentals.medians
    rows = [["Ratio"] + [str(y) for y in years]]
    for k in order:
        s = r.get(k, {})
        if not any(is_num(s.get(y)) for y in years):
            continue
        rows.append([RATIO_META[k][0]] + [_fmt_ratio(k, s.get(y)) for y in years])
    first = 66 * mm
    other = (CONTENT_W - first) / len(years)
    return _table(rows, [first] + [other] * len(years), st)


def _valuation_table(a, st):
    v = a.valuation
    rows = [["Méthode", "Valeur (FCFA)", "Poids", "Hypothèse"]]
    tw = sum(m.weight for m in v.methods) or 1
    for m in v.methods:
        rows.append([m.name, fmt_num(m.value), fmt_pct(m.weight / tw, 0), P(m.detail, st["cell"])])
    rows.append([P("Juste valeur pondérée", st["cellb"]), P(fmt_num(v.fair_value), st["cellbr"]), "100%",
                 P(f"Potentiel {fmt_pct(v.upside, sign=True)} ; rentabilité totale attendue "
                   f"{fmt_pct(v.total_return, sign=True)} (dividende inclus)", st["cell"])])
    return _table(rows, [40 * mm, 24 * mm, 14 * mm, CONTENT_W - 78 * mm], st, num_cols_from=1, left_cols=(3,))


def _hypotheses(a, st):
    v, s = a.valuation, a.settings
    rows = [
        ["Taux sans risque (UMOA-Titres)", fmt_pts(s.taux_sans_risque, 2, sign=False),
         "Prime de risque actions", fmt_pts(s.prime_risque_marche, 2, sign=False)],
        ["Bêta retenu", fmt_num(v.beta, 2), "Coût des fonds propres", fmt_pct(v.cost_of_equity)],
        ["Croissance perpétuelle", fmt_pct(v.g), "ROE normatif (3 ans)", fmt_pct(v.roe_norm)],
        ["P/B justifié", fmt_x(v.pb_justified, 2), "Taux de distribution", fmt_pct(v.payout, 0)],
        ["ROE implicite dans le cours", fmt_pct(v.implied_roe), "Croissance implicite du dividende", fmt_pct(v.implied_div_growth)],
    ]
    return _table(rows, [CONTENT_W * 0.30, CONTENT_W * 0.18, CONTENT_W * 0.32, CONTENT_W * 0.20], st,
                  header=False, left_cols=(2,))


def _tech_table(a, st):
    t = a.technical
    if not t.available:
        return P("Historique de cours insuffisant pour l'analyse technique.", st["body"])
    i, p, lv = t.indicators, t.performance, t.levels
    rows = [
        ["Tendance", t.trend.capitalize(), "Score technique", fmt_num(t.score) + "/100" if is_num(t.score) else "n.d."],
        ["Moyenne mobile 50 j", fmt_num(i.get("mm50")), "Moyenne mobile 200 j", fmt_num(i.get("mm200"))],
        ["RSI 14", fmt_num(i.get("rsi")), "MACD - signal", fmt_num(i.get("macd_hist"), 1)],
        ["Volatilité annualisée", fmt_pct(i.get("volatilite")), "Repli maximal 12 mois", fmt_pct(i.get("drawdown_max"))],
        ["Support / résistance", f"{fmt_num(lv.get('support_1'))} / {fmt_num(lv.get('resistance_1'))}",
         "Supports et résistances de 2nd rang", f"{fmt_num(lv.get('support_2'))} / {fmt_num(lv.get('resistance_2'))}"],
        ["Perf. 1 mois / 3 mois", f"{fmt_pct(p.get('1m'), sign=True)} / {fmt_pct(p.get('3m'), sign=True)}",
         "Perf. 6 mois / 12 mois", f"{fmt_pct(p.get('6m'), sign=True)} / {fmt_pct(p.get('12m'), sign=True)}"],
        ["Perf. relative 6 mois (BRVM-C)", fmt_pct(p.get("relative_6m"), sign=True),
         "BRVM Composite depuis le 1er janv.", fmt_pct(p.get("indice_ytd"), sign=True)],
    ]
    return _table(rows, [CONTENT_W * 0.27, CONTENT_W * 0.20, CONTENT_W * 0.31, CONTENT_W * 0.22], st,
                  header=False, left_cols=(2,))


def _macro_table(a, st):
    m = a.macro
    y0 = m.year or date.today().year
    years = [y0 - 1, y0, y0 + 1]
    labels = {"croissance": "Croissance du PIB réel", "inflation": "Inflation moyenne",
              "dette_publique": "Dette publique / PIB", "solde_budgetaire": "Solde budgétaire / PIB",
              "solde_courant": "Solde courant / PIB"}
    rows = [[f"{m.country_name} (%)"] + [str(y) + ("p" if y >= y0 else "") for y in years]]
    for k, lab in labels.items():
        s = m.data.get(k, {})
        if s:
            rows.append([lab] + [fmt_pts(s.get(y), 1, sign=False) if is_num(s.get(y)) else "n.d." for y in years])
    return _table(rows, [CONTENT_W * 0.4] + [CONTENT_W * 0.2] * 3, st)


def _news_table(a, st):
    items = a.news.items[:8]
    if not items:
        return P(a.news.error or "Aucune actualité recensée sur la période.", st["body"])
    rows = [["Date", "Source", "Titre", "Ton"]]
    for n in items:
        tone = "+" if n.sentiment > 0.15 else ("-" if n.sentiment < -0.15 else "=")
        rows.append([n.published.strftime("%d/%m/%Y") if n.published else "", P(n.source, st["cell"]),
                     P(n.title, st["cell"]), tone])
    t = _table(rows, [17 * mm, 30 * mm, CONTENT_W - 57 * mm, 10 * mm], st, num_cols_from=3)
    return t


def _sources(a, st):
    s = dict(a.market.sources)
    s["états financiers"] = f"Classeur {a.company.sheet} ({a.company.unit_label})"
    s["macroéconomie"] = a.macro.source
    s["politique monétaire"] = f"BCEAO, Comité de Politique Monétaire du {a.macro.bceao.get('date_reference')}"
    if a.news.items:
        s["actualités"] = "Google Actualités (flux RSS), tonalité lexicale"
    rows = [[k.capitalize(), P(v, st["cell"])] for k, v in s.items()]
    return _table(rows, [40 * mm, CONTENT_W - 40 * mm], st, header=False, bold_cols=(0,))


# ---------------------------------------------------------------------------
# Construction
# ---------------------------------------------------------------------------

def build_pdf(a, compact: bool = False) -> bytes:
    _register_fonts()
    st = _styles()
    buf = io.BytesIO()
    first, later = _decorations(a, None)
    doc = BaseDocTemplate(buf, pagesize=A4, leftMargin=MARGIN, rightMargin=MARGIN, topMargin=MARGIN,
                          bottomMargin=15 * mm, title=f"Fiche de valeur {a.company.ticker}",
                          author=a.settings.maison, subject="Fiche de valeur BRVM")
    f1 = Frame(MARGIN, 15 * mm, CONTENT_W, PAGE_H - 15 * mm - 31 * mm, id="f1", leftPadding=0, rightPadding=0,
               topPadding=0, bottomPadding=0)
    f2 = Frame(MARGIN, 15 * mm, CONTENT_W, PAGE_H - 15 * mm - 16 * mm, id="f2", leftPadding=0, rightPadding=0,
               topPadding=0, bottomPadding=0)
    doc.addPageTemplates([PageTemplate("first", [f1], onPage=first), PageTemplate("later", [f2], onPage=later)])
    story = [NextPageTemplate("later")]
    d, v, c = a.decision, a.valuation, a.company
    tx = a.texts

    # ---------------- Page 1 : synthèse ----------------
    top = Table([[_signal_box(a, st), "", [_kpis(a, st), Spacer(1, 3), _keydata(a, st)]]],
                colWidths=[56 * mm, 4 * mm, CONTENT_W - 60 * mm])
    top.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                             ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0),
                             ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
    story += [top, Spacer(1, 5)]
    story.append(P("Thèse d'investissement" + (" (rédaction assistée)" if a.ai_used else ""), st["h2"]))
    for para in tx.get("these", []):
        story.append(P(para, st["body"]))
        story.append(Spacer(1, 2.5))
    img = _img(charts.price_chart(a.technical, months=18, size=(7.4, 2.9)), CONTENT_W, 64 * mm)
    if img:
        story += [P("Évolution du cours (18 mois)", st["h2"]), img]
    left = [P("Points forts", st["h2"])] + _bullets(tx.get("forces") or ["Aucun point saillant détecté."], st)
    right = [P("Points de vigilance", st["h2"])] + _bullets(tx.get("vigilances") or ["Aucun point saillant détecté."], st)
    story.append(_two_cols(left, right))
    pil = _img(charts.pillars_chart(d.pillars, size=(3.7, 2.0)), CONTENT_W * 0.48, 44 * mm)
    just = [P("Fondement du signal", st["h2"])] + _bullets(tx.get("justification", [])[:4], st)
    story.append(KeepTogether(_two_cols([P("Notation par pilier (/100)", st["h2"]), pil] if pil else [], just)))

    if not compact:
        # ---------------- Page 2 : analyse financière ----------------
        story += [PageBreak(), P("Soldes intermédiaires de gestion", st["h1"]), _sig_table(a, st), Spacer(1, 3)]
        fin = _img(charts.financials_chart(c, a.fundamentals.ratios, size=(7.4, 2.35)), CONTENT_W, 56 * mm)
        if fin:
            story.append(fin)
        story += [P("Ratios financiers", st["h1"]), _ratio_table(a, st)]

        # ---------------- Page 3 : solidité, pairs, valorisation ----------------
        sol = a.fundamentals.solidity
        sol_rows = [["Test de solidité", "Résultat"]] + [
            [lab, "Validé" if ok else ("Non validé" if ok is False else "n.d.")] for lab, ok in sol.tests]
        sol_tb = _table(sol_rows, [CONTENT_W * 0.37, CONTENT_W * 0.13], st)
        sol_side = [P(f"Score de solidité : {sol.score}/{sol.maximum}", st["h2"]), sol_tb]
        if sol.altman is not None:
            sol_side.append(P(f"Score Z'' d'Altman (marchés émergents) : {fmt_num(sol.altman, 2)}, zone de "
                              f"{sol.altman_zone}.", st["small"]))
        med = a.fundamentals.medians
        med_rows = [["Médianes des pairs", ""], ["Groupe", P(a.fundamentals.peer_label, st["cellr"])],
                    ["Nombre de pairs", fmt_num(med.get("n"))], ["PER", fmt_x(med.get("pe"))],
                    ["P/B", fmt_x(med.get("pb"), 2)], ["ROE", fmt_pct(med.get("roe"))],
                    ["Marge nette", fmt_pct(med.get("net_margin"))], ["Rendement du dividende", fmt_pct(med.get("div_yield"))]]
        if not c.is_bank:
            med_rows.insert(5, ["VE/EBE", fmt_x(med.get("ev_ebitda"))])
        med_tb = _table(med_rows, [CONTENT_W * 0.25, CONTENT_W * 0.20], st)
        story += [PageBreak(), P("Solidité financière et positionnement sectoriel", st["h1"]),
                  _two_cols(sol_side, [P("Positionnement sectoriel", st["h2"]), med_tb], ratio=0.52)]
        peers = _img(charts.peers_chart(a.fundamentals, c, v, size=(7.4, 2.1)), CONTENT_W, 52 * mm)
        if peers:
            story += [Spacer(1, 2), peers]
        story += [P("Valorisation", st["h1"]), P(tx.get("valorisation", ""), st["body"]), Spacer(1, 3)]
        if v.methods:
            story += [_valuation_table(a, st), Spacer(1, 2)]
        vimg = _img(charts.valuation_chart(v, a.technical, size=(7.4, 2.2)), CONTENT_W, 52 * mm)
        if vimg:
            story.append(vimg)
        story += [KeepTogether([P("Hypothèses", st["h2"]), _hypotheses(a, st)])]
        for n in v.notes:
            story.append(P(n, st["muted"]))

        # ---------------- Page 4 : technique et environnement ----------------
        story += [PageBreak(), P("Analyse technique", st["h1"]), P(tx.get("technique", ""), st["body"]), Spacer(1, 2)]
        timg = _img(charts.price_chart(a.technical, months=12, size=(7.4, 3.1), with_rsi=True), CONTENT_W, 70 * mm)
        if timg:
            story.append(timg)
        story += [Spacer(1, 2), _tech_table(a, st)]
        if a.technical.available and a.technical.signals:
            sigs = [s_ for s_, _ in a.technical.signals]
            half = (len(sigs) + 1) // 2
            story += [P("Signaux techniques", st["h2"]),
                      _two_cols(_bullets(sigs[:half], st), _bullets(sigs[half:], st))]
        story += [P("Environnement macroéconomique et monétaire", st["h1"]), _macro_table(a, st),
                  Spacer(1, 3), P(tx.get("macro", ""), st["body"])]
        for n in a.macro.notes:
            story.append(P(n, st["muted"]))

        # ---------------- Page 5 : actualités, décision, sources ----------------
        story += [PageBreak(), P("Actualités et informations communiquées", st["h1"])]
        if a.news.items:
            story += [_news_table(a, st), Spacer(1, 3)]
        story.append(P(tx.get("actualites", ""), st["body"]))
        if tx.get("catalyseurs"):
            story += [P("Catalyseurs (rédaction assistée)", st["h2"])] + _bullets(tx["catalyseurs"], st)
        if tx.get("risques_ia"):
            story += [P("Risques (rédaction assistée)", st["h2"])] + _bullets(tx["risques_ia"], st)
        mat = _img(charts.decision_matrix_chart(d, a.settings, size=(4.2, 3.0)), CONTENT_W * 0.46, 66 * mm)
        just_all = [P("Décision", st["h2"])] + _bullets(tx.get("justification", []), st)
        if d.warnings:
            just_all += [P("Réserves", st["h2"])] + _bullets(d.warnings, st)
        story += [P("Matrice de décision", st["h1"]), _two_cols([mat] if mat else [], just_all, ratio=0.48)]
        story += [CondPageBreak(40 * mm), P("Sources et qualité des données", st["h1"]), _sources(a, st)]
        if c.notes:
            story += [Spacer(1, 2)] + [P("Contrôle des états financiers : " + n, st["muted"]) for n in c.notes]
        if a.errors:
            story += [P("Accès aux données : " + "; ".join(a.errors[:5]), st["muted"])]
        story += [P("Méthodologie", st["h2"]), P(METHODO, st["small"]), Spacer(1, 3), P(config.DISCLAIMER, st["muted"])]
    else:
        story += [Spacer(1, 3), P(config.DISCLAIMER, st["muted"])]
    doc.build(story)
    return buf.getvalue()


METHODO = (
    "Six piliers sont notés de 0 à 100 : valorisation (potentiel vers une juste valeur multicritère, PER relatif, "
    "P/B rapporté au P/B justifié, rendement), qualité et rentabilité (ROE, marges, structure financière, score de "
    "solidité à 9 tests inspiré de Piotroski, Z'' d'Altman, positionnement vis-à-vis des pairs), croissance (TCAM des "
    "revenus et du résultat, tendance des marges), analyse technique (moyennes mobiles 50 et 200 jours, momentum et "
    "force relative au BRVM Composite, RSI, MACD, volatilité), environnement macroéconomique (FMI, BCEAO) et "
    "actualités (tonalité des titres de presse). La juste valeur pondère jusqu'à cinq méthodes : PER et P/B "
    "médians des pairs, P/B justifié de Gordon, actualisation des dividendes à deux étapes et VE/EBE sectoriel. Le "
    "signal croise la vue fondamentale avec le potentiel ; la tendance technique sert de déclencheur."
)
