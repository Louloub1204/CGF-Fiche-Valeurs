"""Application Streamlit : fiche de valeur BRVM et rapport PowerPoint.

Lancer en local :  streamlit run app.py
"""
from __future__ import annotations

import io
import zipfile
from dataclasses import replace
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

from fiche_valeur import charts, config
from fiche_valeur.config import Settings
from fiche_valeur.export import build_excel, build_screening_excel, screening_frame
from fiche_valeur.fundamentals import RATIO_META, sig_table
from fiche_valeur.market import load_market_base, load_price_file
from fiche_valeur.parser import parse_workbook
from fiche_valeur.pipeline import run_analysis
from fiche_valeur.report_pdf import build_pdf
from fiche_valeur.report_pptx import build_pptx
from fiche_valeur.utils import (clean_text, fmt_fcfa, fmt_num, fmt_pct, fmt_price, fmt_pts, fmt_x, is_num)

ROOT = Path(__file__).resolve().parent
DEMO_EF = ROOT / "sample_data" / "demo_etats_financiers.xlsx"
DEMO_MKT = ROOT / "sample_data" / "demo_base_marche.xlsx"

st.set_page_config(page_title="Fiche de valeur BRVM", page_icon="📈", layout="wide",
                   initial_sidebar_state="expanded")

st.markdown(f"""
<style>
.block-container {{padding-top: 1.6rem; padding-bottom: 2rem;}}
h1, h2, h3 {{color: {config.NAVY}; font-family: Cambria, Caladea, Georgia, serif;}}
.fv-head {{background:{config.NAVY}; color:white; padding:14px 20px; border-radius:6px;
           border-bottom:4px solid {config.GOLD}; margin-bottom:12px;}}
.fv-head .t {{font-family: Cambria, Caladea, Georgia, serif; font-size:1.7rem; font-weight:700;}}
.fv-head .s {{opacity:.85; font-size:.95rem;}}
.fv-badge {{display:inline-block; color:white; padding:10px 18px; border-radius:6px; font-weight:700;
            font-size:1.5rem; letter-spacing:.5px;}}
.fv-small {{color:{config.MUTED}; font-size:.85rem;}}
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Chargements mis en cache
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def _parse(data: bytes, name: str):
    return parse_workbook(data, filename=name)


@st.cache_data(show_spinner=False)
def _market(data: bytes):
    return load_market_base(io.BytesIO(data))


@st.cache_data(show_spinner=False)
def _prices(data: bytes, name: str):
    return load_price_file(io.BytesIO(data), name)


def _secret(key: str) -> str:
    try:
        return st.secrets.get(key, "")
    except Exception:
        return ""


# ---------------------------------------------------------------------------
# Barre latérale
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### 1. Données")
    ef_file = st.file_uploader("États financiers (Etats_Financiers_2021_2025.xlsx)", type=["xlsx", "xlsm"],
                               help="Classeur à une feuille par ticker : ligne d'en-tête « Poste | 2021 | ... | 2025 ».")
    mkt_file = st.file_uploader("Base de marché (facultatif)", type=["xlsx", "xlsm"],
                                help="Classeur type « Base de données - SMF » : feuilles Cours, Nb_titres, "
                                     "Moyenne_cours, et le cas échéant Volumes et Dividendes.")
    px_file = st.file_uploader("Historique d'un titre (facultatif)", type=["csv", "xlsx", "txt"],
                               help="Colonnes Date, Ouverture, Plus haut, Plus bas, Clôture, Volume.")
    demo = False
    if ef_file is None:
        demo = st.toggle("Utiliser les classeurs de démonstration (données fictives)", value=True)

    if ef_file is not None:
        ef_bytes, ef_name = ef_file.getvalue(), ef_file.name
    elif demo and DEMO_EF.exists():
        ef_bytes, ef_name = DEMO_EF.read_bytes(), DEMO_EF.name
    else:
        ef_bytes, ef_name = None, ""
    if mkt_file is not None:
        mkt_bytes = mkt_file.getvalue()
    elif demo and ef_file is None and DEMO_MKT.exists():
        mkt_bytes = DEMO_MKT.read_bytes()
    else:
        mkt_bytes = None

    wb = _parse(ef_bytes, ef_name) if ef_bytes else None
    mb = _market(mkt_bytes) if mkt_bytes else None

    st.markdown("### 2. Valeur analysée")
    ticker = None
    if wb and wb.companies:
        opts = wb.tickers
        ticker = st.selectbox("Société", opts, format_func=lambda t: f"{t} - {wb.companies[t].name}")
    else:
        st.info("Chargez le classeur des états financiers pour commencer.")

    st.markdown("### 3. Paramètres")
    online = st.toggle("Compléter avec les données en ligne", value=not demo,
                       help="Cours et volumes (sikafinance), nombre de titres (brvm.org), dividendes, "
                            "indicateurs FMI, actualités (Google Actualités).")
    with st.expander("Saisies manuelles (prioritaires)"):
        m_price = st.number_input("Cours (FCFA)", min_value=0.0, value=0.0, step=5.0)
        m_shares = st.number_input("Nombre de titres", min_value=0.0, value=0.0, step=1000.0)
        m_float = st.number_input("Flottant (%)", min_value=0.0, max_value=100.0, value=0.0, step=1.0)
        m_dps = st.number_input("Dividende net par action (FCFA)", min_value=0.0, value=0.0, step=5.0)
        m_dps_y = st.number_input("Exercice du dividende", min_value=2015, max_value=2035, value=date.today().year - 1)
        major = st.checkbox("Événement majeur en cours (OPA, augmentation de capital, litige...)",
                            help="Place le titre en position d'attente jusqu'à clarification.")
        notes = st.text_area("Informations communiquées par la société",
                             placeholder="Faits marquants, guidance, communiqué de résultats, projet de dividende...")
    with st.expander("Hypothèses de valorisation"):
        rf = st.number_input("Taux sans risque (%)", value=6.50, step=0.05,
                             help="Rendement souverain UMOA-Titres de maturité longue.")
        erp = st.number_input("Prime de risque actions (%)", value=5.50, step=0.25)
        g = st.number_input("Croissance perpétuelle (%)", value=3.00, step=0.25)
        beta = st.number_input("Bêta par défaut", value=1.00, step=0.05,
                               help="Utilisé si la régression sur le BRVM Composite n'est pas significative.")
        s_buy = st.number_input("Potentiel minimal pour ACHAT (%)", value=15.0, step=1.0)
        s_sell = st.number_input("Potentiel maximal pour VENDRE (%)", value=-15.0, step=1.0)
    with st.expander("Pondération des piliers"):
        w = {}
        defaults = Settings().poids
        for k, lab in config.PILLAR_LABELS.items():
            w[k] = st.slider(lab, 0, 50, int(defaults[k]), step=5)
    with st.expander("Politique monétaire BCEAO"):
        b = dict(config.BCEAO)
        b["taux_directeur"] = st.number_input("Taux directeur (%)", value=float(b["taux_directeur"]), step=0.25)
        b["guichet_pret_marginal"] = st.number_input("Guichet de prêt marginal (%)", value=float(b["guichet_pret_marginal"]), step=0.25)
        b["inflation_uemoa_prevue"] = st.number_input("Inflation UEMOA attendue (%)", value=float(b["inflation_uemoa_prevue"]), step=0.1)
        b["croissance_uemoa_prevue"] = st.number_input("Croissance UEMOA attendue (%)", value=float(b["croissance_uemoa_prevue"]), step=0.1)
        b["date_reference"] = st.text_input("Réunion du CPM de référence", value=b["date_reference"])
    with st.expander("Rédaction assistée par Claude (facultatif)"):
        key_secret = _secret("ANTHROPIC_API_KEY")
        use_ai = st.toggle("Reformuler la thèse avec Claude", value=False,
                           help="Seules les données calculées sont transmises ; le signal n'est jamais modifié.")
        api_key = key_secret or st.text_input("Clé API Anthropic", type="password")
        model = st.text_input("Modèle", value="claude-sonnet-5-5")
    with st.expander("Identité du document"):
        house = st.text_input("Société de gestion", value=config.DEFAULT_HOUSE)
        analyst = st.text_input("Analyste", value="")

settings = Settings(taux_sans_risque=rf, prime_risque_marche=erp, croissance_long_terme=g, beta_defaut=beta,
                    poids=w, seuil_potentiel_achat=s_buy, seuil_potentiel_vente=s_sell, en_ligne=online,
                    claude_api_key=api_key or "", claude_model=model, maison=house or config.DEFAULT_HOUSE,
                    analyste=analyst)
overrides = {}
if m_price > 0:
    overrides["price"] = m_price
if m_shares > 0:
    overrides["shares"] = m_shares
if m_float > 0:
    overrides["float_pct"] = m_float / 100
if m_dps > 0:
    overrides["dps"] = m_dps
    overrides["dps_exercice"] = int(m_dps_y)

# ---------------------------------------------------------------------------
# En-tête
# ---------------------------------------------------------------------------
st.markdown("# Fiche de valeur BRVM")
st.markdown('<span class="fv-small">Analyse fondamentale (SIG, ratios, solidité, pairs), valorisation '
            'multicritère, analyse technique, environnement macroéconomique et actualités, avec un signal '
            'ACHAT / CONSERVER / ATTENDRE / ALLÉGER / VENDRE. Fiche imprimable en PDF et rapport PowerPoint.</span>',
            unsafe_allow_html=True)
if demo and ef_file is None:
    st.warning("Mode démonstration : sociétés et chiffres fictifs. Chargez votre classeur d'états financiers "
               "dans la barre latérale pour analyser les valeurs de la BRVM.")
if wb is not None and wb.notes:
    for n in wb.notes:
        st.error(n)
if not ticker:
    st.stop()

tab_one, tab_batch = st.tabs(["Analyse d'une valeur", "Tableau de bord multi-valeurs"])

# ---------------------------------------------------------------------------
# Analyse d'une valeur
# ---------------------------------------------------------------------------
with tab_one:
    run_key = (ticker, ef_name, len(ef_bytes or b""), online, tuple(sorted(overrides.items())), rf, erp, g, beta,
               tuple(sorted(w.items())), s_buy, s_sell, major, notes, use_ai, tuple(sorted(b.items())), house, analyst,
               px_file.name if px_file else None)
    c1, c2 = st.columns([1, 3])
    go = c1.button("Lancer l'analyse", type="primary", width="stretch")
    if go or (st.session_state.get("run_key") != run_key and st.session_state.get("auto", False)):
        price_df = _prices(px_file.getvalue(), px_file.name) if px_file is not None else None
        with st.status("Analyse en cours...", expanded=False) as status:
            a = run_analysis(wb, ticker, settings, market_base=mb, price_file=price_df, overrides=overrides,
                             user_notes=notes, major_event=major, bceao=b, use_claude=use_ai,
                             progress=lambda msg: status.update(label=msg))
            status.update(label="Analyse terminée", state="complete")
        st.session_state["analysis"] = a
        st.session_state["run_key"] = run_key
        st.session_state["auto"] = True
        st.session_state.pop("pdf", None)
        st.session_state.pop("pptx", None)
    a = st.session_state.get("analysis")
    if a is None or a.company.ticker != ticker:
        c2.info("Cliquez sur « Lancer l'analyse » pour produire la fiche de cette valeur.")
        st.stop()
    if st.session_state.get("run_key") != run_key:
        c2.warning("Paramètres modifiés : relancez l'analyse pour les prendre en compte.")

    c, v, fd, t, m, d, tx = a.company, a.valuation, a.fundamentals, a.technical, a.macro, a.decision, a.texts
    country = config.COUNTRIES.get(c.country, {}).get("nom", c.country)
    color = config.SIGNAL_COLORS.get(d.signal, config.NAVY)
    h1, h2 = st.columns([3, 1])
    h1.markdown(f'<div class="fv-head"><div class="t">{clean_text(c.name)}</div>'
                f'<div class="s">{c.ticker} | {c.sector} | {country} | comptes {c.kind.upper()} '
                f'en {c.unit_label}</div></div>', unsafe_allow_html=True)
    h2.markdown(f'<div style="text-align:center"><span class="fv-badge" style="background:{color}">{d.signal}'
                f'</span><div class="fv-small">Conviction {d.conviction.lower()} | note '
                f'{fmt_num(d.composite)}/100</div></div>', unsafe_allow_html=True)

    k = st.columns(6)
    k[0].metric("Cours (FCFA)", fmt_num(v.price), help=a.market.sources.get("cours"))
    k[1].metric("Juste valeur (FCFA)", fmt_num(d.target))
    k[2].metric("Potentiel", fmt_pct(d.upside, sign=True))
    k[3].metric("Rendement net", fmt_pct(v.div_yield))
    k[4].metric("PER", fmt_x(v.pe))
    k[5].metric("P/B", fmt_x(v.pb, 2))
    st.caption(f"Capitalisation {fmt_fcfa(a.market.market_cap)} | {fmt_num(a.market.shares)} titres | "
               f"flottant {fmt_pct(a.market.float_pct, 0)} | cours au "
               f"{a.market.price_date.strftime('%d/%m/%Y') if a.market.price_date else 'n.d.'} "
               f"({a.market.sources.get('cours', 'source n.d.')})")

    # Téléchargements ----------------------------------------------------------
    dl = st.columns(4)
    if "pdf" not in st.session_state:
        st.session_state["pdf"] = build_pdf(a)
        st.session_state["pdf1"] = build_pdf(a, compact=True)
        st.session_state["pptx"] = build_pptx(a)
        st.session_state["xlsx"] = build_excel(a)
    stamp = date.today().strftime("%Y%m%d")
    dl[0].download_button("Fiche PDF complète", st.session_state["pdf"], f"Fiche_{c.ticker}_{stamp}.pdf",
                          "application/pdf", width="stretch", on_click="ignore", type="primary")
    dl[1].download_button("Fiche PDF 1 page", st.session_state["pdf1"], f"Fiche_{c.ticker}_{stamp}_synthese.pdf",
                          "application/pdf", width="stretch", on_click="ignore")
    dl[2].download_button("Rapport PowerPoint", st.session_state["pptx"], f"Rapport_{c.ticker}_{stamp}.pptx",
                          "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                          width="stretch", on_click="ignore")
    dl[3].download_button("Données (Excel)", st.session_state["xlsx"], f"Analyse_{c.ticker}_{stamp}.xlsx",
                          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                          width="stretch", on_click="ignore")

    t1, t2, t3, t4, t5, t6 = st.tabs(["Synthèse", "Fondamentaux", "Valorisation", "Technique",
                                      "Macro et actualités", "Données et méthode"])
    with t1:
        l, r = st.columns([3, 2])
        with l:
            st.subheader("Thèse d'investissement" + (" (rédaction assistée)" if a.ai_used else ""))
            for p in tx.get("these", []):
                st.write(p)
            img = charts.price_chart(t, months=18, size=(7.4, 3.0))
            if img:
                st.image(img, width="stretch")
        with r:
            st.subheader("Notation par pilier")
            st.image(charts.pillars_chart(d.pillars, size=(4.2, 2.5)), width="stretch")
            st.subheader("Fondement du signal")
            for j in tx.get("justification", []):
                st.markdown(f"- {j}")
            for wn in d.warnings:
                st.markdown(f"- :orange[{wn}]")
        f1, f2 = st.columns(2)
        f1.subheader("Points forts")
        for s_ in tx.get("forces") or ["Aucun point saillant détecté."]:
            f1.markdown(f"- {s_}")
        f2.subheader("Points de vigilance")
        for s_ in tx.get("vigilances") or ["Aucun point saillant détecté."]:
            f2.markdown(f"- {s_}")

    with t2:
        st.subheader("Soldes intermédiaires de gestion (Mds FCFA)")
        sig = sig_table(c) / 1e9
        st.dataframe(sig.style.format(lambda x: fmt_num(x, 1) if is_num(x) else "n.d."), width="stretch")
        st.write(tx.get("activite", ""))
        img = charts.financials_chart(c, fd.ratios, size=(8, 2.6))
        if img:
            st.image(img, width="stretch")
        st.subheader("Ratios")
        years = c.years[-5:]
        rows = {}
        for key_, s_ in fd.ratios.items():
            if key_ in RATIO_META:
                kind = RATIO_META[key_][1]
                rows[RATIO_META[key_][0]] = [
                    (fmt_pct(s_.get(y_)) if kind == "pct" else fmt_num(s_.get(y_), 1)) for y_ in years]
        st.dataframe(pd.DataFrame(rows, index=[str(y_) for y_ in years]).T, width="stretch")
        a1_, a2_ = st.columns(2)
        sol = fd.solidity
        a1_.subheader(f"Score de solidité : {sol.score}/{sol.maximum}")
        a1_.dataframe(pd.DataFrame([(lab, "Validé" if ok else ("Non validé" if ok is False else "n.d."))
                                    for lab, ok in sol.tests], columns=["Test", "Résultat"]),
                      hide_index=True, width="stretch")
        if sol.altman is not None:
            a1_.caption(f"Z'' d'Altman : {fmt_num(sol.altman, 2)} (zone de {sol.altman_zone})")
        a2_.subheader("Pairs : " + fd.peer_label)
        img = charts.peers_chart(fd, c, v, size=(6, 2.6))
        if img:
            a2_.image(img, width="stretch")

    with t3:
        st.write(tx.get("valorisation", ""))
        if v.methods:
            tw = sum(m_.weight for m_ in v.methods)
            st.dataframe(pd.DataFrame([(m_.name, fmt_num(m_.value), fmt_pct(m_.weight / tw, 0), m_.detail)
                                       for m_ in v.methods], columns=["Méthode", "FCFA / action", "Poids", "Hypothèse"]),
                         hide_index=True, width="stretch")
        img = charts.valuation_chart(v, t, size=(8, 2.6))
        if img:
            st.image(img, width="stretch")
        for n_ in v.notes:
            st.warning(n_)
        st.caption(f"Coût des fonds propres {fmt_pct(v.cost_of_equity)} ; bêta {fmt_num(v.beta, 2)} "
                   f"({v.beta_source}) ; croissance perpétuelle {fmt_pct(v.g)}.")

    with t4:
        if t.available:
            st.write(tx.get("technique", ""))
            st.image(charts.price_chart(t, months=12, size=(8, 3.8), with_rsi=True), width="stretch")
            st.image(charts.macd_chart(t, size=(8, 1.8)), width="stretch")
            for s_, kind in t.signals:
                icon = {"pos": ":green[▲]", "neg": ":red[▼]"}.get(kind, ":gray[■]")
                st.markdown(f"{icon} {s_}")
        else:
            st.info("Historique de cours indisponible : chargez une base de marché, un historique, ou activez les "
                    "données en ligne.")

    with t5:
        st.subheader(f"{m.country_name}")
        st.write(tx.get("macro", ""))
        y0 = m.year
        md_rows = {lab: [m.data.get(key_, {}).get(y_) for y_ in (y0 - 1, y0, y0 + 1)]
                   for key_, lab in (("croissance", "Croissance du PIB réel (%)"), ("inflation", "Inflation (%)"),
                                     ("dette_publique", "Dette publique / PIB (%)"),
                                     ("solde_budgetaire", "Solde budgétaire / PIB (%)"),
                                     ("solde_courant", "Solde courant / PIB (%)")) if m.data.get(key_)}
        if md_rows:
            st.dataframe(pd.DataFrame(md_rows, index=[str(y0 - 1), f"{y0}p", f"{y0 + 1}p"]).T, width="stretch")
        st.caption(m.source)
        st.subheader("Actualités")
        if a.news.items:
            st.dataframe(pd.DataFrame([(n_.published.strftime("%d/%m/%Y") if n_.published else "", n_.source,
                                        n_.title, n_.link) for n_ in a.news.items],
                                      columns=["Date", "Source", "Titre", "Lien"]),
                         hide_index=True, width="stretch",
                         column_config={"Lien": st.column_config.LinkColumn("Lien", display_text="Ouvrir")})
        st.write(tx.get("actualites", ""))

    with t6:
        st.subheader("Provenance des données")
        src = dict(a.market.sources)
        src["états financiers"] = f"{ef_name}, feuille {c.sheet} ({c.unit_label})"
        st.table(pd.DataFrame(list(src.items()), columns=["Donnée", "Source"]))
        if c.notes:
            st.subheader("Contrôle des états financiers")
            for n_ in c.notes:
                st.markdown(f"- {n_}")
        if a.errors:
            st.subheader("Accès aux sources")
            for e_ in a.errors:
                st.markdown(f"- {e_}")
        with st.expander("Libellés reconnus dans la feuille"):
            st.table(pd.DataFrame([(k_, v_) for k_, v_ in c.field_labels.items()], columns=["Champ", "Libellé source"]))
        st.caption(config.DISCLAIMER)

# ---------------------------------------------------------------------------
# Tableau de bord multi-valeurs
# ---------------------------------------------------------------------------
with tab_batch:
    st.write("Analyse en série des valeurs du classeur : tableau des signaux, export Excel et archive des fiches PDF.")
    chosen = st.multiselect("Valeurs", wb.tickers, default=wb.tickers,
                            format_func=lambda t_: f"{t_} - {wb.companies[t_].name}")
    with_pdf = st.checkbox("Inclure les fiches PDF dans l'archive", value=True)
    if st.button("Analyser la sélection", type="primary"):
        results, errors = [], []
        bar = st.progress(0.0, text="Démarrage...")
        batch_settings = replace(settings, claude_api_key="")
        for i, t_ in enumerate(chosen):
            bar.progress(i / max(1, len(chosen)), text=f"{t_} ({i + 1}/{len(chosen)})")
            try:
                results.append(run_analysis(wb, t_, batch_settings, market_base=mb, bceao=b))
            except Exception as exc:  # une valeur en erreur n'arrête pas le lot
                errors.append(f"{t_} : {type(exc).__name__} {exc}")
        bar.progress(1.0, text="Terminé")
        st.session_state["batch"] = results
        st.session_state["batch_errors"] = errors
        zbuf = io.BytesIO()
        with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
            df = screening_frame(results)
            z.writestr("Tableau_des_signaux.xlsx", build_screening_excel(df))
            if with_pdf:
                for r_ in results:
                    z.writestr(f"Fiche_{r_.company.ticker}.pdf", build_pdf(r_))
        st.session_state["batch_zip"] = zbuf.getvalue()
    results = st.session_state.get("batch")
    if results:
        df = screening_frame(results)
        show = df.copy()
        for col in ("Potentiel", "Rendement", "ROE"):
            show[col] = show[col].map(lambda x: fmt_pct(x) if is_num(x) else "n.d.")
        for col in ("Note globale", "Vue fondamentale", "Technique", "Cours", "Juste valeur"):
            show[col] = show[col].map(lambda x: fmt_num(x) if is_num(x) else "n.d.")
        for col in ("PER", "P/B"):
            show[col] = show[col].map(lambda x: fmt_x(x, 1) if is_num(x) else "n.d.")
        st.dataframe(show, hide_index=True, width="stretch")
        counts = df["Signal"].value_counts()
        st.caption(" | ".join(f"{k_} : {v_}" for k_, v_ in counts.items()))
        st.download_button("Télécharger l'archive (tableau Excel + fiches)", st.session_state["batch_zip"],
                           f"Fiches_BRVM_{date.today():%Y%m%d}.zip", "application/zip", type="primary",
                           on_click="ignore")
        for e_ in st.session_state.get("batch_errors", []):
            st.error(e_)
