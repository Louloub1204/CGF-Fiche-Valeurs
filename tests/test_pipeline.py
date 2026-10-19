"""Tests hors ligne : lecture du classeur, analyses, rapports et analyseurs de flux.

Lancement : python -m pytest -q
"""
from __future__ import annotations

import io
import json
import re
from datetime import date
from pathlib import Path

import openpyxl
import pandas as pd
import pytest

from fiche_valeur import config
from fiche_valeur.config import Settings
from fiche_valeur.market import load_market_base, parse_capitalisations, parse_sika_payload
from fiche_valeur.news import parse_rss
from fiche_valeur.parser import parse_workbook
from fiche_valeur.pipeline import run_analysis
from fiche_valeur.utils import clean_text, fmt_num, fmt_pct, norm_label_core

ROOT = Path(__file__).resolve().parents[1]
EF = ROOT / "sample_data" / "demo_etats_financiers.xlsx"
SMF = ROOT / "sample_data" / "demo_base_marche.xlsx"
SIGNALS = set(config.SIGNAL_COLORS)


@pytest.fixture(scope="module")
def wb():
    return parse_workbook(EF)


@pytest.fixture(scope="module")
def mb():
    return load_market_base(SMF)


@pytest.fixture(scope="module")
def offline():
    return Settings(en_ligne=False)


@pytest.fixture(scope="module")
def analyses(wb, mb, offline):
    return {t: run_analysis(wb, t, offline, market_base=mb) for t in wb.tickers}


# --------------------------------------------------------------------------- formats

def test_percent_attached_and_no_long_dash():
    assert fmt_pct(0.125) == "12,5%"
    assert " %" not in fmt_pct(-0.034)
    assert clean_text("Marge — en hausse – 10 %") == "Marge - en hausse - 10%"
    assert fmt_num(1234567.8, 0).replace(" ", " ").replace("\xa0", " ") == "1 234 568"


def test_label_normalisation():
    assert norm_label_core("XB - Chiffre d'affaires (A+B)") == "CHIFFRE D AFFAIRES"
    assert norm_label_core("Résultat net") == "RESULTAT NET"


# --------------------------------------------------------------------------- lecture

def test_demo_workbook(wb):
    assert set(wb.tickers) == {"DBQC", "DBQS", "DAGC", "DDIC", "DTLS", "DENC"}
    kinds = {t: c.kind for t, c in wb.companies.items()}
    assert kinds["DBQC"] == "banque" and kinds["DAGC"] == "syscohada" and kinds["DTLS"] == "ifrs"
    for c in wb.companies.values():
        assert c.years[-1] == 2025
        assert c.get(c.net_income_field, 2025) is not None
        assert c.get(c.equity_field, 2025) is not None


def test_unit_harmonisation(wb):
    c = wb.companies["DDIC"]
    assert any("2021" in n and "unité" in n for n in c.notes)
    ca = [c.get("ca", y) for y in c.years]
    # Après harmonisation, aucun exercice ne s'écarte d'un facteur 10 de la médiane
    med = sorted(ca)[len(ca) // 2]
    assert all(0.1 < v / med < 10 for v in ca)


def test_formula_evaluation(tmp_path):
    """Classeur saisi avec formules non recalculées (pas de valeurs en cache)."""
    book = openpyxl.Workbook()
    ws = book.active
    ws.title = "TEST"
    ws.append(["Test Industries SA - états financiers (en millions FCFA)"])
    ws.append(["Poste", 2023, 2024, 2025])
    ws.append(["COMPTE DE RÉSULTAT"])
    ws.append(["Chiffre d'affaires", 50000, 55000, 60500])
    ws.append(["Achats consommés", -30000, -32000, -35000])
    ws.append(["Valeur ajoutée", "=B4+B5", "=C4+C5", "=D4+D5"])
    ws.append(["Charges de personnel", -6000, -6500, -7000])
    ws.append(["Excédent brut d'exploitation", "=B6+B7", "=C6+C7", "=D6+D7"])
    ws.append(["Dotations aux amortissements", -3000, -3200, -3400])
    ws.append(["Résultat d'exploitation", "=SUM(B8:B9)", "=SUM(C8:C9)", "=SUM(D8:D9)"])
    ws.append(["Résultat net", 7000, 8100, 9400])
    ws.append(["BILAN"])
    ws.append(["Total actif", 90000, 95000, 101000])
    ws.append(["Capitaux propres", 40000, 44000, 48000])
    ws.append(["Dettes financières", 15000, 14000, 13000])
    path = tmp_path / "formules.xlsx"
    book.save(path)
    c = parse_workbook(path).companies["TEST"]
    assert c.kind == "syscohada"
    assert c.get("va", 2025) == pytest.approx(25500e6)
    assert c.get("ebe", 2025) == pytest.approx(18500e6)
    assert c.get("re", 2024) == pytest.approx(13300e6)


def test_market_base(mb):
    assert "DBQC" in mb.tickers
    assert mb.closes is not None and len(mb.closes) > 200
    assert mb.shares_total.get("DBQC")


# --------------------------------------------------------------------------- analyses

def test_signals_offline(analyses):
    for t, a in analyses.items():
        d = a.decision
        assert d.signal in SIGNALS, t
        assert 0 <= d.composite <= 100
        assert d.trace and "signal de base" in d.trace[0]
        assert a.valuation.fair_value and a.valuation.fair_value > 0
        assert a.technical.available
        for txt in a.texts.values():
            flat = " ".join(txt) if isinstance(txt, list) else str(txt)
            assert "—" not in flat, t
            assert not re.search(r"\d[ \xa0\u202f]%", flat), (t, flat)


def test_signal_diversity(analyses):
    signals = {a.decision.signal for a in analyses.values()}
    assert "ACHAT" in signals and "VENDRE" in signals


def test_major_event_forces_wait(wb, mb, offline):
    a = run_analysis(wb, "DAGC", offline, market_base=mb, major_event=True)
    assert a.decision.signal in {"ATTENDRE", "VENDRE"}


def test_manual_price_override(wb, mb, offline):
    base = run_analysis(wb, "DAGC", offline, market_base=mb)
    cheap = run_analysis(wb, "DAGC", offline, market_base=mb, overrides={"price": base.valuation.price * 0.6})
    assert cheap.valuation.upside > base.valuation.upside


# --------------------------------------------------------------------------- rapports

def test_reports(analyses):
    from pptx import Presentation

    from fiche_valeur.export import build_excel, build_screening_excel, screening_frame
    from fiche_valeur.report_pdf import build_pdf
    from fiche_valeur.report_pptx import build_pptx

    a = analyses["DBQC"]
    pdf = build_pdf(a)
    assert pdf[:4] == b"%PDF" and len(pdf) > 50_000
    assert build_pdf(a, compact=True)[:4] == b"%PDF"
    prs = Presentation(io.BytesIO(build_pptx(analyses["DAGC"])))
    assert len(prs.slides) >= 10
    xl = pd.read_excel(io.BytesIO(build_excel(a)), sheet_name=None)
    assert {"Synthèse", "SIG", "Ratios", "Valorisation"} <= set(xl)
    scr = screening_frame(list(analyses.values()))
    assert len(scr) == len(analyses)
    assert build_screening_excel(scr)[:2] == b"PK"


# --------------------------------------------------------------------------- flux en ligne (hors ligne)

CAPI_HTML = """
<table><thead><tr><th>Symbole</th><th>Nom</th><th>Nombre de titres</th><th>Cours du jour</th>
<th>Capitalisation flottante</th><th>Capitalisation globale</th><th>Capitalisation globale (%)</th></tr></thead>
<tbody>
<tr><td>SNTS</td><td>SONATEL SN</td><td>100 000 000</td><td>25 500</td><td>675 750 000 000</td>
<td>2 550 000 000 000</td><td>20,1</td></tr>
<tr><td>SGBC</td><td>SOCIETE GENERALE COTE D'IVOIRE</td><td>31 111 110</td><td>28 000</td>
<td>145 000 000 000</td><td>871 111 080 000</td><td>6,9</td></tr>
</tbody></table>"""


def test_parse_capitalisations():
    df = parse_capitalisations(CAPI_HTML)
    assert list(df.index) == ["SNTS", "SGBC"]
    assert df.loc["SNTS", "nb_titres"] == 100_000_000
    assert df.loc["SGBC", "cours"] == 28_000
    assert df.loc["SNTS", "cap_globale"] == 2.55e12


def test_parse_sika_payload():
    payload = json.loads("""{"lst": [
        {"Date": "02/01/2025", "Open": 100, "High": 104, "Low": 99, "Close": 103, "Volume": 1200},
        {"Date": "03/01/2025", "Open": 103, "High": 106, "Low": 101, "Close": 105, "Volume": 800},
        {"Date": "06/01/2025", "Open": 105, "High": 105, "Low": 98, "Close": 99, "Volume": 1500}]}""")
    df = parse_sika_payload(payload)
    assert len(df) == 3 and df["close"].iloc[-1] == 99
    assert df.index.is_monotonic_increasing
    assert parse_sika_payload({"lst": []}) is None


RSS = """<?xml version="1.0" encoding="UTF-8"?><rss version="2.0"><channel>
<item><title>Sonatel : le chiffre d'affaires progresse de 8% - Sika Finance</title><link>https://ex.org/1</link>
<pubDate>Mon, 05 Oct 2026 08:00:00 GMT</pubDate><source url="https://ex.org">Sika Finance</source></item>
<item><title>Sonatel propose un dividende record aux actionnaires - Agence Ecofin</title><link>https://ex.org/2</link>
<pubDate>Tue, 15 Sep 2026 08:00:00 GMT</pubDate><source url="https://ex.org">Agence Ecofin</source></item>
<item><title>Le marché régional des télécoms en Afrique de l'Ouest - Jeune Afrique</title><link>https://ex.org/3</link>
<pubDate>Tue, 15 Sep 2026 08:00:00 GMT</pubDate><source url="https://ex.org">Jeune Afrique</source></item>
</channel></rss>"""


def test_parse_rss():
    news = parse_rss(RSS.encode(), "SNTS", "Sonatel")
    titles = [n.title for n in news.items]
    assert len(titles) == 2 and all("Sonatel" in t for t in titles)
    assert not titles[0].endswith("Sika Finance")
    assert "Dividende / détachement" in news.events
    assert news.sentiment > 0 and news.score > 50
    assert parse_rss(b"pas du xml", "SNTS", "Sonatel").error


# --------------------------------------------------------------------------- titres à données partielles

SIKA_SOCIETE_HTML = """<html><body><div>Les chiffres sont en millions de FCFA</div>
<table><tr><td></td><td>2021</td><td>2022</td><td>2023</td><td>2024</td><td>2025</td></tr>
<tr><td>Chiffre d'affaires</td><td>80 000</td><td>88 000</td><td>95 000</td><td>104 000</td><td>112 000</td></tr>
<tr><td>Croissance CA</td><td></td><td>10,00 %</td><td>7,95 %</td><td>9,47 %</td><td>7,69 %</td></tr>
<tr><td>Résultat net</td><td>18 000</td><td>20 000</td><td>22 000</td><td>24 000</td><td>27 000</td></tr>
<tr><td>BNPA</td><td>720,00</td><td>800,00</td><td>880,00</td><td>960,00</td><td>1 080,00</td></tr>
<tr><td>PER</td><td>10,5</td><td>9,8</td><td>9,1</td><td>8,9</td><td>8,6</td></tr>
<tr><td>Dividende</td><td>400,00</td><td>450,00</td><td>500,00</td><td>550,00</td><td>600,00</td></tr></table>
<p>Nombre de titres : 25 000 000</p><p>Flottant : 22,5 %</p>
<p>Principaux actionnaires : GROUPE X 60 % ; PUBLIC (BRVM) 22,5 %</p></body></html>"""


def test_parse_sika_societe():
    from fiche_valeur.partial import parse_sika_societe
    kf = parse_sika_societe(SIKA_SOCIETE_HTML)
    assert kf.revenue[2025] == 112_000e6 and kf.net_income[2021] == 18_000e6
    assert kf.dps[2023] == 500 and kf.eps[2025] == 1080 and kf.per[2021] == 10.5
    assert kf.shares == 25_000_000 and kf.float_pct == pytest.approx(0.225)
    assert parse_sika_societe("<html><body>rien</body></html>") is None


def test_manual_figures_table():
    from fiche_valeur.partial import empty_manual_table, figures_from_table
    assert figures_from_table(empty_manual_table()) is None
    df = empty_manual_table(2025, 3)
    df.iloc[:, 1] = [100.0, 110.0, 120.0]       # chiffre d'affaires, millions
    df.iloc[:, 2] = [10.0, 12.0, 15.0]          # résultat net
    kf = figures_from_table(df)
    assert kf.revenue[2025] == 120e6 and kf.net_income[2023] == 10e6


@pytest.fixture(scope="module")
def partial_env(mb):
    config.register_ticker("DAGC", "Agro démo", "CI", "Consommation de base")
    config.register_ticker("DBQC", "Banque démo", "CI", "Services financiers")
    return mb


def test_level2_key_figures(partial_env, offline):
    from fiche_valeur.partial import parse_sika_societe
    kf = parse_sika_societe(SIKA_SOCIETE_HTML)
    a = run_analysis(None, "DAGC", offline, market_base=partial_env, figures=kf)
    assert a.company.level == "chiffres_cles"
    assert a.decision.signal in SIGNALS and a.decision.conviction in {"Moyenne", "Faible"}
    assert a.valuation.pe and a.valuation.fair_value
    assert a.coverage.level == "chiffres_cles" and 0 < a.coverage.score < 100
    assert any(i.name.startswith("Bilan") and i.status == "absent" for i in a.coverage.items)


def test_level3_market_only(partial_env, offline):
    a = run_analysis(None, "DBQC", offline, market_base=partial_env)
    assert a.company.level == "marche"
    d = a.decision
    assert d.signal in {"ATTENDRE", "CONSERVER", "ALLÉGER", "ACHAT"} and d.signal != "VENDRE"
    assert d.conviction == "Faible" and d.fundamental is None
    assert a.technical.available and a.valuation.fair_value


def test_level3_without_any_data(offline):
    config.register_ticker("QQXQ", "Titre sans données", "CI", "Non classé")
    a = run_analysis(None, "QQXQ", offline)
    assert a.decision.signal == "ATTENDRE" and a.decision.conviction == "Faible"
    assert a.coverage.score < 40


def test_partial_reports(partial_env, offline):
    from pptx import Presentation

    from fiche_valeur.export import build_excel
    from fiche_valeur.partial import parse_sika_societe
    from fiche_valeur.report_pdf import build_pdf
    from fiche_valeur.report_pptx import build_pptx
    cases = [run_analysis(None, "DAGC", offline, market_base=partial_env, figures=parse_sika_societe(SIKA_SOCIETE_HTML)),
             run_analysis(None, "DBQC", offline, market_base=partial_env),
             run_analysis(None, "QQXQ", offline)]
    for a in cases:
        assert build_pdf(a)[:4] == b"%PDF" and build_pdf(a, compact=True)[:4] == b"%PDF"
        assert len(Presentation(io.BytesIO(build_pptx(a))).slides) >= 9
        assert "Couverture" in pd.read_excel(io.BytesIO(build_excel(a)), sheet_name=None)


def test_guess_country():
    assert config.guess_country("BOABF") == "BF" and config.guess_country("BOAS") == "SN"
    assert config.guess_country("SGBC") == "CI" and config.guess_country("ETIT") == "TG"


def test_full_referential():
    assert len(config.TICKERS) >= 47
    assert all(v["pays"] in config.COUNTRIES and v["secteur"] in config.SECTOR_MULTIPLES
               for v in config.TICKERS.values())


# --------------------------------------------------------------------------- base SMF réelle (format) et publications

def _smf_like(path):
    """Classeur au format de la « Base de données - SMF » (colonne d'index, .BRVMCI, volumes en FCFA, blocs)."""
    import numpy as np
    book = openpyxl.Workbook()
    ws = book.active
    ws.title = "Cours"
    ws.append([None, "Date", "AAAC", "BBBC", ".BRVMCI", "ANNEE"])
    days = pd.bdate_range("2024-01-02", periods=400)
    rng = np.random.default_rng(1)
    pa = 1000 * np.cumprod(1 + rng.normal(0.0005, 0.01, 400))
    pb = 5000 * np.cumprod(1 + rng.normal(0.0003, 0.012, 400))
    pi = 250 * np.cumprod(1 + rng.normal(0.0004, 0.007, 400))
    pa, pb = np.round(pa), np.round(pb)
    for i, d in enumerate(days):
        ws.append([i + 2, d.to_pydatetime(), float(pa[i]), float(pb[i]), float(pi[i]), d.year])
    wv = book.create_sheet("Volume moyen")
    wv.append([None, "Date", "AAAC", "BBBC"])
    wv.append([None, "Moyenne", 1, 2])
    for i, d in enumerate(days):
        wv.append([i, d.to_pydatetime(), float(pa[i] * (100 + i % 7)), float(pb[i] * (20 + i % 5))])
    tb = book.create_sheet("Tableau de bord")
    tb.append([None, 1000000, "AAAC", "BBBC"])
    tb.append([2, None, "CAPITAUX PROPRES"])
    for y, a_, b_ in ((2023, 8e9, 40e9), (2024, 9e9, 44e9), (2025, 10e9, 47e9)):
        tb.append([None, y, a_, b_])
    tb.append([None, None, "RESULTAT NET"])
    tb.append([None, None, "AAAC", "BBBC"])
    for y, a_, b_ in ((2023, 1.5e9, 6e9), (2024, 1.7e9, 0), (2025, 2.0e9, 7e9)):
        tb.append([None, y, a_, b_])
    tb.append([None, None, "CHIFFRE D'AFFAIRES"])
    tb.append([None, None, "AAAC", "BBBC"])
    for y, a_, b_ in ((2023, 20e9, 60e9), (2024, 22e9, 63e9), (2025, 25e9, 66e9)):
        tb.append([None, y, a_, b_])
    hd = book.create_sheet("Historique_dividende")
    hd.append([None, "Date", "AAAC", "BBBC"])
    for y, a_, b_ in ((2023, 40, 300), (2024, 45, 320), (2025, 50, 2000)):
        hd.append([None, y, a_, b_])
    nb = book.create_sheet("Nb_titres")
    nb.append([None, "Date", "AAAC", "BBBC"])
    nb.append([None, "Titres flottants", 4_000_000, 2_000_000])
    nb.append([None, "Nombre de titres", 20_000_000, 10_000_000])
    book.save(path)
    return path


def test_smf_format(tmp_path):
    mb = load_market_base(_smf_like(tmp_path / "smf.xlsx"))
    assert "BRVMC" in mb.closes.columns and "ANNEE" not in mb.closes.columns
    assert mb.volumes_in_value and 90 <= mb.volumes["AAAC"].dropna().iloc[5] <= 110
    assert mb.fundamentals["AAAC"]["capitaux_propres"][2025] == 10e9
    assert 2024 not in mb.fundamentals["BBBC"]["rn"]           # 0 = non renseigné
    assert [d.exercice for d in mb.dividends["AAAC"]] == [2023, 2024, 2025]   # ligne N = dividende de l'exercice N
    assert mb.shares_total["BBBC"] == 10_000_000 and mb.shares_float["AAAC"] == 4_000_000


def test_smf_key_figures_and_exceptional_dividend(tmp_path, offline):
    mb = load_market_base(_smf_like(tmp_path / "smf.xlsx"))
    config.register_ticker("AAAC", "Société A", "CI", "Industriels")
    config.register_ticker("BBBC", "Société B", "CI", "Industriels")
    a = run_analysis(None, "AAAC", offline, market_base=mb)
    assert a.company.level == "chiffres_cles" and "SMF" in a.company.data_source
    assert a.valuation.pb and a.valuation.pe and a.fundamentals.ratios["roe"][2025] > 0
    b = run_analysis(None, "BBBC", offline, market_base=mb)
    assert b.valuation.dps == pytest.approx(310)           # 2 000 jugé exceptionnel : médiane des précédents
    assert any("exceptionnel" in n for n in b.valuation.notes)


def test_smf_fills_workbook_gaps(wb, offline, tmp_path):
    import copy as _copy
    from fiche_valeur.partial import KeyFigures, fill_company_gaps
    c = _copy.deepcopy(wb.companies["DAGC"])
    kf = KeyFigures(source="Base de données SMF", revenue={2026: 1e12}, net_income={2025: 123.0})
    added = fill_company_gaps(c, kf)
    assert added == [2026] and c.fields["ca"][2026] == 1e12
    assert c.fields["rn"][2025] != 123.0                     # jamais d'écrasement


REPORT_TXT = """Montants en milliards de FCFA
Indicateurs S1 2026 S1 2025 Variation
Chiffre d'affaires 672,8 579,3 16,1% 93,4
EBITDAaL 239,3 200,8 19,2% 38,5
Résultat d'exploitation 159,6 131,8 21,1% 27,8
Résultat financier (12,8) (13,6) -5,9% 0,8
Résultat des activités ordinaires 146,8 118,2 24,2% 28,6
Impôts sur le résultat (56,6) (41,9) 35,0% (14,7)
Résultat net 90,3 76,3 18,3% 13,9"""

LISTING_HTML = """<div>
<a href="/common/actualite/details/08-10-2026-orange-ci-rapport-d-activites-1er-semestre-2026">ORANGE CI : Rapport d'activités - 1er semestre 2026</a>
<a href="/common/actualite/afficher-fichier/08-10-2026-orange-ci-rapport-d-activites-1er-semestre-2026">Télécharger</a>
<a href="/common/actualite/afficher-fichier/17-09-2026-africa-global-logistics-ci-etats-financiers-certifies-exercice-2025">AFRICA GLOBAL LOGISTICS CI : Etats Financiers certifiés - Exercice 2025</a>
<a href="/common/actualite/details/25-09-2026-boa-niger-rapport-dactivites-1er-semestre-2026">BOA NIGER : Rapport d'activités - 1er semestre 2026</a>
</div>"""


def test_publications_parsers():
    from fiche_valeur.publications import parse_listing, parse_report_text, ticker_of_company
    pubs = parse_listing(LISTING_HTML)
    assert [(ticker_of_company(p.company), p.kind, p.period) for p in pubs] == [
        ("ORAC", "interim", "S1 2026"), ("BOAN", "interim", "S1 2026"), ("SDSC", "annuel", "Exercice 2025")]
    assert pubs[0].pdf_url.endswith("afficher-fichier/08-10-2026-orange-ci-rapport-d-activites-1er-semestre-2026")
    rep = parse_report_text(REPORT_TXT, "ORANGE CI : Rapport d'activités - 1er semestre 2026")
    assert rep.period == "S1 2026" and rep.period_prev == "S1 2025" and rep.unit_label == "milliards FCFA"
    assert rep.get("ca").current == pytest.approx(672.8e9) and rep.get("rf").current == pytest.approx(-12.8e9)
    assert rep.get("rn").variation == pytest.approx(90.3 / 76.3 - 1)
    assert parse_report_text("rien d'exploitable", "x") is None


def test_report_pdf_roundtrip(tmp_path):
    """Un PDF texte est relu par pdfplumber puis analysé."""
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas

    from fiche_valeur.publications import parse_report_pdf
    path = tmp_path / "rapport.pdf"
    cv = canvas.Canvas(str(path), pagesize=A4)
    y = 800
    for line in REPORT_TXT.splitlines():
        cv.drawString(50, y, line)
        y -= 16
    cv.save()
    rep = parse_report_pdf(path.read_bytes(), "ORANGE CI : Rapport d'activités - 1er semestre 2026")
    assert rep is not None and rep.get("rn").current == pytest.approx(90.3e9)


def test_interim_in_reports(partial_env, offline):
    from fiche_valeur.publications import parse_report_text
    from fiche_valeur.report_pdf import build_pdf
    from fiche_valeur.report_pptx import build_pptx
    rep = parse_report_text(REPORT_TXT, "AGRO DEMO : Rapport d'activités - 1er semestre 2026")
    a = run_analysis(None, "DAGC", offline, market_base=partial_env, report=rep)
    assert a.interim is rep and any(p[0].startswith("Dynamique") for p in a.fundamentals.growth_parts)
    assert a.dynamics is not None and a.dynamics.label == "S1 2026"
    assert "S1" in a.texts["activite"] and "sur un an" in a.texts["activite"]
    assert any(i.name.startswith("Publication") and i.status == "ok" for i in a.coverage.items)
    assert build_pdf(a)[:4] == b"%PDF" and build_pptx(a)[:2] == b"PK"


def test_file_first_and_split_detection():
    from fiche_valeur.market import _extend_history
    idx = pd.bdate_range("2026-06-01", periods=60)
    f = pd.DataFrame({"open": None, "high": None, "low": None, "close": 40000.0, "volume": 100.0}, index=idx[:40])
    w = pd.DataFrame({"open": None, "high": None, "low": None, "close": 4000.0, "volume": 1000.0}, index=idx[35:])
    out, split = _extend_history(f, w)          # web ajusté : 10 pour 1
    assert split == 10 and out["close"].iloc[0] == 4000 and out.index.max() == idx[-1]
    w2 = pd.DataFrame({"open": None, "high": None, "low": None, "close": 41000.0, "volume": 5.0}, index=idx[35:])
    out2, split2 = _extend_history(f, w2)       # pas de fractionnement : le fichier prime sur les dates communes
    assert split2 == 1 and out2.loc[idx[36], "close"] == 40000 and out2.loc[idx[50], "close"] == 41000


def test_referential_sectors():
    assert config.TICKERS["SDSC"]["secteur"] == config.IND and config.TICKERS["LNBB"]["secteur"] == config.CDI
    assert "CROWN SIEM CI" in config.TICKERS["SEMC"]["autres_noms"]
    from fiche_valeur.publications import ticker_of_company
    assert ticker_of_company("BANK OF AFRICA NG") == "BOAN" and ticker_of_company("ECOBANK ETI") == "ETIT"


def test_auto_params(tmp_path):
    from fiche_valeur.autoparams import classify_announcements, suggest
    from fiche_valeur.publications import Publication
    mb = load_market_base(_smf_like(tmp_path / "smf.xlsx"))
    config.register_ticker("AAAC", "Société A", "CI", "Industriels")
    ap = suggest("AAAC", None, mb, online=False)
    assert ap.price == mb.closes["AAAC"].dropna().iloc[-1] and ap.shares == 20_000_000
    assert ap.float_pct == pytest.approx(0.2) and ap.dps == 50 and ap.dps_exercice == 2025
    assert ap.rf == config.TAUX_SOUVERAINS["CI"] and 12 <= ap.seuil_achat <= 30
    major, info = classify_announcements([
        Publication("SOCIETE A CI : Offre publique d'achat", date(2026, 10, 1), "u"),
        Publication("SOCIETE A CI : Paiement de dividendes - Exercice 2025", date(2026, 9, 1), "u"),
        Publication("SOCIETE A CI : Fractionnement de l'action", date(2026, 9, 2), "u")])
    assert len(major) == 1 and "Offre publique" in major[0] and len(info) == 2


def test_sheet_name_resolution(tmp_path):
    from fiche_valeur.parser import resolve_ticker
    assert resolve_ticker("Filtisac") == "FTSC" and resolve_ticker("FTSC CI") == "FTSC"
    assert resolve_ticker("BOA Niger") == "BOAN" and resolve_ticker("Total Sénégal") == "TTLS"
    # libellés en colonne B, feuille nommée par la raison sociale
    book = openpyxl.Workbook()
    ws = book.active
    ws.title = "FILTISAC COTE D'IVOIRE"
    ws.append(["", "Filtisac - états financiers (en millions FCFA)"])
    ws.append(["Code", "Poste", 2023, 2024, 2025])
    ws.append(["XB", "Chiffre d'affaires", 38214, 30694, 32108])
    ws.append(["XI", "Résultat net", 3075, 18595, 465])
    ws.append(["", "Capitaux propres", 26903, 43454, 15703])
    ws.append(["", "Total actif", 60000, 70000, 50000])
    path = tmp_path / "ef.xlsx"
    book.save(path)
    wb_ = parse_workbook(path)
    assert "FTSC" in wb_.companies and wb_.companies["FTSC"].get("ca", 2025) == pytest.approx(32108e6)
    assert wb_.sheet_report[0][1] == "FTSC" and wb_.sheet_report[0][2].startswith("lue")



# ---------------------------------------------------------------------------
# Publications intermédiaires (T1, S1, T3) et dynamique de l'exercice en cours
# ---------------------------------------------------------------------------
def _interim_records():
    ent = lambda y, p, ca, ebe, rn: {"ticker": "FTSC", "year": y, "period": p, "statut": "Publié", "source": "test",
                                      "values": {"ca": ca, "ebe": ebe, "rn": rn}}
    return [ent(2026, "T1", 9000, 700, 300), ent(2025, "T1", 8000, 600, 250),
            ent(2026, "S1", 19000, 1500, 700), ent(2025, "S1", 16000, 1200, 500)]


def test_interim_file_roundtrip():
    from fiche_valeur.interim_file import build_template, parse_interim_file
    data = build_template(_interim_records())
    book = openpyxl.load_workbook(io.BytesIO(data))
    assert book.sheetnames == ["Lisez-moi", "Entreprises", "Banques", "Suivi", "Titres"]
    d = parse_interim_file(data)
    assert d.n_rows == 4 and not d.unknown and not d.notes
    assert d.latest("FTSC") == (2026, "S1")
    assert d.rows["FTSC"][(2026, "S1")].values["ca"] == pytest.approx(19000e6)       # millions -> FCFA
    assert parse_interim_file(build_template(None)).notes                              # modèle vierge signalé


def test_dynamics_ltm_and_trend():
    from fiche_valeur.dynamics import InterimData, PeriodRow, build_dynamics
    from fiche_valeur.interim_file import build_template, parse_interim_file
    from fiche_valeur.partial import KeyFigures, build_company
    kf = KeyFigures(revenue={2024: 30000e6, 2025: 34000e6}, net_income={2024: 1000e6, 2025: 1300e6},
                    equity={2025: 9000e6}, source="test")
    c = build_company("FTSC", kf)
    d = parse_interim_file(build_template(_interim_records()))
    dyn = build_dynamics("FTSC", c, d)
    assert dyn.label == "S1 2026" and [cp.label for cp in dyn.compares] == ["T1 2026", "S1 2026"]
    assert dyn.var("ca") == pytest.approx(19000 / 16000 - 1)
    # 12 mois glissants = exercice N-1 + cumul N - cumul N-1
    assert dyn.ltm["ca"] == pytest.approx((34000 + 19000 - 16000) * 1e6)
    assert dyn.ltm["rn"] == pytest.approx((1300 + 700 - 500) * 1e6)
    assert dyn.fy_est["ca"] == pytest.approx(34000e6 * 19000 / 16000)
    assert "accélération" in dyn.trend                       # croissance du résultat net : +20% (T1) puis +40% (S1)
    # exercice déjà clos dans les comptes annuels : pas de dynamique
    kf2 = KeyFigures(revenue={2025: 34000e6, 2026: 40000e6}, net_income={2025: 1300e6, 2026: 1500e6}, source="test")
    assert build_dynamics("FTSC", build_company("FTSC", kf2), d) is None


def test_interim_file_priority_over_online():
    from fiche_valeur.dynamics import InterimData, PeriodRow, merged_rows
    from fiche_valeur.publications import parse_report_text
    rep = parse_report_text(REPORT_TXT, "AGRO DEMO : Rapport d'activités - 1er semestre 2026")
    d = InterimData(rows={"DAGC": {(2026, "S1"): PeriodRow(values={"ca": 1.0}, statut="Publié")}})
    rows = merged_rows("DAGC", d, rep)
    assert rows[(2026, "S1")].values["ca"] == 1.0                      # le fichier prime
    assert "rn" in rows[(2026, "S1")].values and (2025, "S1") in rows  # la lecture en ligne comble le reste


def test_ltm_per_in_valuation(partial_env, offline):
    from fiche_valeur.dynamics import InterimData, PeriodRow
    cur, prev = {"ca": 700e9, "rn": 100e9}, {"ca": 600e9, "rn": 80e9}
    d = InterimData(rows={"DAGC": {(2026, "S1"): PeriodRow(values=cur), (2025, "S1"): PeriodRow(values=prev)}})
    a = run_analysis(None, "DAGC", offline, market_base=partial_env, interim_data=d)
    assert a.dynamics is not None and a.dynamics.label == "S1 2026"
    names = [m.name for m in a.valuation.methods]
    if "PER sectoriel" in names:                              # BPA glissant disponible : le PER est scindé en deux méthodes
        assert "PER sur 12 mois glissants" in names and a.valuation.pe_ltm
    assert any(p[0].startswith("Dynamique S1 2026") for p in a.fundamentals.growth_parts)
    from fiche_valeur.report_pdf import build_pdf
    from fiche_valeur.report_pptx import build_pptx
    assert build_pdf(a)[:4] == b"%PDF" and build_pptx(a)[:2] == b"PK"


# ---------------------------------------------------------------------------
# Conformité Charia
# ---------------------------------------------------------------------------
def _charia_book(rows):
    book = openpyxl.Workbook()
    ws = book.active
    ws.title = "TEST SCREENING UMOA"
    for r, (sym, j, k, l, y) in enumerate(rows, start=6):
        ws.cell(r, 27, sym)
        for col, v in zip((10, 11, 12, 14, 15, 16, 18, 19, 20, 22, 23), (j, k, l, j, k, l, j, k, l, j, l)):
            ws.cell(r, col, v)
        ws.cell(r, 25, y)
    buf = io.BytesIO()
    book.save(buf)
    return buf.getvalue()


def test_charia_file_rules():
    from fiche_valeur.charia import parse_charia_file
    data = parse_charia_file(_charia_book([
        ("PALC", 0.004, 0.393, 0.048, 0), ("SMBC", 0.40, 0.61, 0.13, 0), ("NTLC", 0.05, 0.2, 0.05, 0.08),
        ("UNXC", 0, 0, 0, 0), ("SOGB", 0.0, 0.089, 0.058, 0), ("BICC", 0.0, 0.1, 0.1, 0), ("LNBB", 0.0, 0.1, 0.1, 0)]))
    r = data.results
    assert r["PALC"].statut == "compatible" and r["PALC"].n_pass == 3          # DJIM : créances > 33%
    assert r["SMBC"].statut == "non_compatible" and r["SMBC"].n_pass == 0
    assert r["NTLC"].statut == "non_compatible" and r["NTLC"].filtre_ok is False   # revenus d'intérêt > 5%
    assert r["UNXC"].statut == "non_evaluable"
    assert r["SOGC"].statut == "compatible"                                    # symbole SOGB du fichier
    assert r["BICC"].statut == "exclu" and r["LNBB"].statut == "exclu"
    bad = parse_charia_file(_charia_book([]))
    assert not bad.results


def test_charia_in_reports(partial_env, offline):
    from fiche_valeur.charia import parse_charia_file
    from fiche_valeur.export import build_excel
    from fiche_valeur.report_pdf import build_pdf
    from fiche_valeur.report_pptx import build_pptx
    data = parse_charia_file(_charia_book([("SNTS", 0.1, 0.1, 0.1, 0)]))
    a = run_analysis(None, "SNTS", offline, charia_data=data)
    assert a.charia.statut == "compatible"
    assert any("Charia" in x for x in a.texts["forces"])
    pdf = build_pdf(a)
    assert pdf[:4] == b"%PDF"
    assert build_pptx(a)[:2] == b"PK"
    xl = openpyxl.load_workbook(io.BytesIO(build_excel(a)))
    assert "Charia" in xl.sheetnames
    a2 = run_analysis(None, "SNTS", offline)
    assert a2.charia.statut == "non_evaluable"


def test_no_dividend_last_year_gives_zero_yield(wb, mb, offline):
    from datetime import date as _d
    from fiche_valeur.market import Dividend
    t = wb.tickers[0]
    a0 = run_analysis(wb, t, offline, market_base=mb)
    mb2 = load_market_base(SMF)
    mb2.dividends[t] = [Dividend(2023, 190.0, _d(2024, 7, 1), "test"), Dividend(2024, 190.0, _d(2025, 7, 1), "test"),
                        Dividend(2025, 0.0, _d(2026, 7, 1), "test")]
    a = run_analysis(wb, t, offline, market_base=mb2)
    assert a.valuation.dps == 0 and a.valuation.dps_exercice == 2025
    assert not a.valuation.div_yield
    assert not any(m.name == "Actualisation des dividendes" for m in a.valuation.methods)
    assert "rendement net" not in a.texts["valorisation"]


def test_stale_dividend_means_no_payment(wb, mb, offline):
    from datetime import date as _d
    from fiche_valeur.market import Dividend
    t = wb.tickers[0]
    ly = wb.companies[t].last_year()
    mb2 = load_market_base(SMF)
    mb2.dividends[t] = [Dividend(ly - 2, 190.0, _d(ly - 1, 7, 1), "test"), Dividend(ly - 1, 190.0, _d(ly, 7, 1), "test")]
    a = run_analysis(wb, t, offline, market_base=mb2)
    if _d.today() >= _d(ly + 1, 9, 30):
        assert a.valuation.dps == 0 and a.valuation.dps_exercice == ly
        assert not a.valuation.div_yield
