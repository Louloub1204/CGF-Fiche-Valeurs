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
from fiche_valeur.partial import MANUAL_COLUMNS, empty_manual_table, figures_from_market_base, figures_from_table
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


@st.cache_data(show_spinner=False)
def _report_from_pdf(data: bytes, title: str):
    from fiche_valeur.publications import parse_report_pdf
    return parse_report_pdf(data, title, "Document chargé par l'utilisateur")


@st.cache_data(show_spinner=False, ttl=3600)
def _report_from_url(url: str):
    from fiche_valeur.publications import fetch_url_report
    return fetch_url_report(url)


def _secret(key: str) -> str:
    try:
        return st.secrets.get(key, "")
    except Exception:
        return ""


# ---------------------------------------------------------------------------
# Titres ajoutés par l'utilisateur (conservés pendant la session)
# ---------------------------------------------------------------------------
for _t, _v in st.session_state.get("custom_tickers", {}).items():
    config.register_ticker(_t, _v["nom"], _v["pays"], _v["secteur"])

LEVEL_TAG = {"complet": "états financiers", "chiffres_cles": "chiffres clés", "marche": "marché seul"}


def _universe(wb, demo: bool) -> list[str]:
    """Tous les titres de la cote (référentiel) et ceux du classeur ; en démonstration, ceux du classeur."""
    in_wb = set(wb.companies) if wb is not None else set()
    for t_, c_ in (wb.companies if wb is not None else {}).items():
        if t_ not in config.TICKERS and not demo:
            config.register_ticker(t_, c_.name, c_.country, c_.sector)
    if demo:
        return sorted(in_wb)
    return sorted(set(config.TICKERS) | in_wb)


# ---------------------------------------------------------------------------
# Barre latérale
# ---------------------------------------------------------------------------
with st.sidebar:
    with st.expander("Guide : quelles données fournir ?", expanded=False):
        st.markdown(
            "L'application analyse **tous les titres de la BRVM**, même sans état financier détaillé. "
            "Plus vous fournissez de données, plus l'analyse est complète et la conviction élevée.\n\n"
            "**Niveau 1 - Analyse complète.** Classeur des états financiers (une feuille par ticker, au moins "
            "trois exercices). Donne : soldes intermédiaires de gestion, ratios, solidité, croissance, comparaison "
            "aux pairs, valorisation par PER, P/B, Gordon, dividendes et VE/EBE.\n\n"
            "**Niveau 2 - Chiffres clés.** Chiffre d'affaires, résultat net, capitaux propres, résultat "
            "d'exploitation (feuille Tableau de bord de la base SMF), sinon BNPA et dividendes lus sur sikafinance, "
            "ou saisis dans l'onglet d'analyse. "
            "Donne : croissance, PER, rendement, valorisation. Pas de solidité financière ; conviction plafonnée.\n\n"
            "**Niveau 3 - Marché seul.** Cours, volumes, dividendes. Donne : analyse technique, valorisation par "
            "les dividendes, macro et actualités. Signal indicatif prudent, conviction faible, pas de VENDRE "
            "et achat seulement avec un fort potentiel.\n\n"
            "**Exercice en cours.** À tous les niveaux, la dernière publication (rapport d'activités trimestriel ou "
            "semestriel, états financiers annuels) est lue sur richbourse.com ou dans le PDF que vous chargez : ses "
            "soldes N / N-1 alimentent la fiche et le pilier croissance.\n\n"
            "Chaque fiche indique le niveau atteint, le taux de couverture des données et ce qu'il faut "
            "ajouter pour monter d'un niveau.")

    st.markdown("### 1. Données")
    ef_file = st.file_uploader(
        "États financiers (classeur Excel)", type=["xlsx", "xlsm"],
        help="Ex. Etats_Financiers_2021_2025.xlsx. Une feuille par société, nommée par son ticker (SNTS, SGBC...). "
             "Chaque feuille contient une ligne d'en-tête « Poste | 2021 | 2022 | ... », l'unité (millions FCFA...) "
             "dans le titre, puis les lignes du compte de résultat et du bilan. Les titres absents du classeur "
             "restent analysables aux niveaux 2 et 3.")
    st.caption("Facultatif mais recommandé. Alimente le niveau 1 et les comparaisons entre sociétés.")
    mkt_file = st.file_uploader(
        "Base de données SMF (classeur Excel)", type=["xlsx", "xlsm"],
        help="Classeur « Base de données - SMF ». Feuilles reconnues par leur nom : Cours (Date puis une colonne "
             "par ticker, .BRVMCI pour le BRVM Composite), Volume moyen (montants échangés, convertis en titres), "
             "Tableau de bord (capitaux propres, résultat net, chiffre d'affaires, résultat d'exploitation, flux "
             "de trésorerie par année), Historique_dividende (dividende net par action, année de paiement), "
             "Moyenne_cours, Nb_titres (titres flottants, nombre de titres).")
    st.caption("À charger après le classeur des états financiers. Apporte cours, volumes, dividendes, nombre de "
               "titres et chiffres clés de toute la cote : les titres hors classeur passent au niveau 2 avec "
               "capitaux propres (ROE, P/B), et les exercices manquants du classeur sont complétés.")
    px_file = st.file_uploader(
        "Historique du titre analysé (CSV ou Excel)", type=["csv", "xlsx", "txt"],
        help="Colonnes Date, Ouverture, Plus haut, Plus bas, Clôture, Volume (français ou anglais). Remplace "
             "l'historique de la base de marché et des sources en ligne pour le titre sélectionné uniquement.")
    st.caption("Facultatif. Export d'un historique de cours pour un seul titre.")
    demo = False
    if ef_file is None:
        demo = st.toggle("Utiliser les classeurs de démonstration", value=False,
                         help="Six sociétés fictives avec leurs cours : permet de tester l'application sans "
                              "aucune donnée réelle.")

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
    in_wb = set(wb.companies) if wb is not None else set()
    if mb is not None and mb.closes is not None and not mb.closes.empty:
        st.success(f"Base SMF : {len([t_ for t_ in mb.tickers if t_ in config.TICKERS])} titres cotés, cours "
                   f"jusqu'au {mb.closes.index.max():%d/%m/%Y}, chiffres clés pour {len(mb.fundamentals)} titres.")

    st.markdown("### 2. Valeur analysée")
    all_tickers = _universe(wb, demo)

    def _fmt_ticker(t_: str) -> str:
        tag = "états financiers" if t_ in in_wb else "hors classeur"
        return f"{t_} - {config.TICKERS.get(t_, {}).get('nom', t_)} ({tag})"

    ticker = st.selectbox("Société", all_tickers, format_func=_fmt_ticker,
                          index=0 if all_tickers else None,
                          help=f"{len(all_tickers)} titres. « états financiers » : présent dans le classeur "
                               f"(niveau 1). « hors classeur » : analysé sur chiffres clés (niveau 2) ou sur "
                               f"données de marché (niveau 3).")
    if not demo:
        with st.expander("Ajouter ou mettre à jour un titre"):
            st.caption("Pour une introduction récente ou un titre absent de la liste. Le pays sert à "
                       "l'environnement macroéconomique, le secteur aux multiples de référence.")
            n_t = st.text_input("Ticker BRVM", placeholder="Ex. ABCD", max_chars=8).strip().upper()
            n_n = st.text_input("Raison sociale", placeholder="Ex. Société Anonyme Exemple")
            n_c = st.selectbox("Pays", list(config.COUNTRIES), format_func=lambda k_: config.COUNTRIES[k_]["nom"])
            n_s = st.selectbox("Secteur", config.SECTORS + ["Non classé"])
            if st.button("Enregistrer le titre", disabled=not n_t):
                st.session_state.setdefault("custom_tickers", {})[n_t] = {"nom": n_n or n_t, "pays": n_c, "secteur": n_s}
                st.rerun()
            st.markdown("---")
            st.caption("Recherche des titres cotés sur brvm.org qui ne figurent pas dans la liste.")
            if st.button("Rechercher les nouveaux titres"):
                from fiche_valeur.market import fetch_brvm_capitalisations
                errs: list[str] = []
                capi = fetch_brvm_capitalisations(errs)
                if capi is None:
                    st.warning("brvm.org n'a pas répondu : " + (errs[0] if errs else "source indisponible."))
                else:
                    new = [t_ for t_ in capi.index if t_ not in config.TICKERS]
                    if not new:
                        st.success("Aucun titre nouveau : la liste est à jour.")
                    else:
                        for t_ in new:
                            nm = str(capi.loc[t_, "nom"]).strip() or t_
                            st.session_state.setdefault("custom_tickers", {})[t_] = {
                                "nom": nm, "pays": config.guess_country(t_), "secteur": "Non classé"}
                        st.rerun()

    st.markdown("### 3. Paramètres")
    online = st.toggle("Compléter avec les données en ligne", value=not demo,
                       help="Active la consultation de sikafinance (cours, volumes, chiffres clés, dividendes), "
                            "brvm.org (nombre de titres, capitalisation, dividendes), FMI (croissance, inflation) et "
                            "Google Actualités. Désactivez pour travailler uniquement sur vos fichiers.")
    with st.expander("Saisies manuelles (prioritaires)"):
        st.caption("Chaque valeur saisie remplace toutes les sources automatiques. Laissez à 0 pour utiliser les "
                   "données du classeur, de la base de marché ou en ligne.")
        m_price = st.number_input("Cours (FCFA)", min_value=0.0, value=0.0, step=5.0,
                                  help="Dernier cours de clôture du titre. Sert au potentiel, au PER, au rendement.")
        m_shares = st.number_input("Nombre de titres", min_value=0.0, value=0.0, step=1000.0,
                                   help="Nombre total d'actions du capital (ex. 100 000 000). Source : brvm.org, "
                                        "rubrique Capitalisations, ou rapport annuel. Nécessaire au BPA, PER, P/B.")
        m_float = st.number_input("Flottant (%)", min_value=0.0, max_value=100.0, value=0.0, step=1.0,
                                  help="Part du capital librement négociable, en %. Sert à mesurer la liquidité "
                                       "(rotation du flottant).")
        m_dps = st.number_input("Dividende net par action (FCFA)", min_value=0.0, value=0.0, step=5.0,
                                help="Dividende net d'impôt versé par action au titre du dernier exercice. Remplace "
                                     "les dividendes de la base de marché ou en ligne.")
        m_dps_y = st.number_input("Exercice du dividende", min_value=2015, max_value=2035,
                                  value=date.today().year - 1, help="Exercice auquel se rapporte le dividende saisi.")
        major = st.checkbox("Événement majeur en cours (OPA, augmentation de capital, litige...)",
                            help="Place le titre en position d'attente jusqu'à clarification, sauf s'il est déjà à "
                                 "vendre. À cocher quand un événement peut bouleverser la valorisation.")
        notes = st.text_area("Informations communiquées par la société",
                             placeholder="Faits marquants, guidance, communiqué de résultats, projet de dividende...",
                             help="Texte libre (communiqué, résultats, assemblée générale). Il figure dans la fiche et sa "
                                  "tonalité est intégrée à la note actualités.")
    with st.expander("Hypothèses de valorisation"):
        st.caption("Ces paramètres fixent le coût des fonds propres k = taux sans risque + bêta x prime de risque.")
        rf = st.number_input("Taux sans risque (%)", value=6.50, step=0.05,
                             help="Rendement des titres souverains UMOA de maturité longue (5 à 10 ans).")
        erp = st.number_input("Prime de risque actions (%)", value=5.50, step=0.25,
                              help="Rendement supplémentaire exigé pour détenir des actions plutôt que le "
                                   "sans-risque. Usuel : 5 à 7% sur la BRVM.")
        g = st.number_input("Croissance perpétuelle (%)", value=3.00, step=0.25,
                            help="Croissance des dividendes à très long terme. Ne doit pas dépasser la "
                                 "croissance nominale de l'économie (3 à 4%).")
        beta = st.number_input("Bêta par défaut", value=1.00, step=0.05,
                               help="Utilisé si la régression sur le BRVM Composite n'est pas significative.")
        s_buy = st.number_input("Potentiel minimal pour ACHAT (%)", value=15.0, step=1.0,
                                help="Potentiel de hausse vers la juste valeur à partir duquel un titre est "
                                     "considéré en décote.")
        s_sell = st.number_input("Potentiel maximal pour VENDRE (%)", value=-15.0, step=1.0,
                                 help="Potentiel en dessous duquel un titre est considéré en surcote.")
        use_def = st.toggle("Multiples de référence si pairs insuffisants", value=True,
                            help="Quand le classeur compte moins de deux comparables, utilise des multiples "
                                 "sectoriels de référence (config.SECTOR_MULTIPLES), signalés dans la fiche.")
    with st.expander("Pondération des piliers"):
        st.caption("Poids relatifs des six piliers dans la note globale (renormalisés automatiquement). "
                   "Un pilier non évalué faute de données est ignoré.")
        w = {}
        defaults = Settings().poids
        helps = {"valorisation": "Potentiel vers la juste valeur, PER relatif, rendement.",
                 "qualite": "ROE, marges, structure financière, solidité (niveau 1 uniquement).",
                 "croissance": "Croissance du chiffre d'affaires et du résultat.",
                 "technique": "Tendance, momentum, force relative, volatilité.",
                 "macro": "Croissance, inflation, politique monétaire, sensibilité du secteur.",
                 "sentiment": "Tonalité des actualités récentes."}
        for k, lab in config.PILLAR_LABELS.items():
            w[k] = st.slider(lab, 0, 50, int(defaults[k]), step=5, help=helps.get(k))
    with st.expander("Politique monétaire BCEAO"):
        st.caption("Dernière décision du Comité de Politique Monétaire. À vérifier avant diffusion.")
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
        api_key = key_secret or st.text_input("Clé API Anthropic", type="password",
                                              help="Clé personnelle (console.anthropic.com). Non enregistrée. "
                                                   "En déploiement, préférez le secret ANTHROPIC_API_KEY.")
        model = st.text_input("Modèle", value="claude-sonnet-5-5")
    with st.expander("Identité du document"):
        house = st.text_input("Société de gestion", value=config.DEFAULT_HOUSE,
                              help="Affichée en en-tête et pied de page des fiches et rapports.")
        analyst = st.text_input("Analyste", value="", help="Nom affiché sur la couverture du rapport PowerPoint.")

settings = Settings(taux_sans_risque=rf, prime_risque_marche=erp, croissance_long_terme=g, beta_defaut=beta,
                    poids=w, seuil_potentiel_achat=s_buy, seuil_potentiel_vente=s_sell, en_ligne=online,
                    multiples_defaut=use_def, claude_api_key=api_key or "", claude_model=model,
                    maison=house or config.DEFAULT_HOUSE, analyste=analyst)
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
            'ACHAT / CONSERVER / ATTENDRE / ALLÉGER / VENDRE. Fiche imprimable en PDF et rapport PowerPoint. '
            f'{len(all_tickers)} titres analysables.</span>', unsafe_allow_html=True)
if demo and ef_file is None:
    st.warning("Mode démonstration : sociétés et chiffres fictifs. Décochez-le dans la barre latérale pour analyser "
               "les valeurs réelles de la BRVM.")
if wb is None and not demo:
    st.info("Aucun classeur d'états financiers chargé : les titres sont analysés sur chiffres clés (niveau 2 : "
            "base SMF, saisie ou lecture en ligne) ou sur données de marché (niveau 3). Chargez votre classeur "
            "dans la barre latérale pour passer les titres qu'il contient au niveau 1.")
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
    figures = None
    if ticker not in in_wb:
        with st.expander("Chiffres clés du titre (hors classeur) : à saisir pour une analyse plus complète",
                         expanded=False):
            st.markdown(
                "Ce titre n'a pas d'états financiers dans le classeur. Si vous activez les données en ligne, ses "
                "chiffres clés sont lus automatiquement sur sikafinance. Vous pouvez aussi les saisir ou les "
                "corriger ici ; **la saisie prime sur la lecture en ligne** et est prise en compte au clic sur "
                "« Lancer l'analyse ».")
            st.caption("Montants en millions de FCFA ; BNPA et dividende en FCFA par action. Laissez vide ce que "
                       "vous ne connaissez pas. Sources : rapport annuel, états financiers publiés sur brvm.org, "
                       "fiche société sikafinance. Les capitaux propres et le total du bilan ajoutent le ROE, le "
                       "P/B et le P/B justifié.")
            tkey = f"fig_{ticker}"
            if tkey not in st.session_state:
                st.session_state[tkey] = empty_manual_table()
            cfg = {col: st.column_config.NumberColumn(col, min_value=0.0 if "dividende" in col.lower() else None,
                                                      format="%.0f" if "millions" in col else "%.2f")
                   for col in MANUAL_COLUMNS.values()}
            cfg["Exercice"] = st.column_config.NumberColumn("Exercice", format="%d", step=1)
            edited = st.data_editor(st.session_state[tkey], hide_index=True, width="stretch", num_rows="dynamic",
                                    column_config=cfg, key=f"editor_{ticker}")
            st.session_state[tkey] = edited
            figures = figures_from_table(edited)
            st.caption("Aucune saisie : lecture en ligne (si activée), sinon analyse sur données de marché seules."
                       if figures is None else "Saisie détectée : elle sera utilisée à la prochaine analyse.")
    report = None
    with st.expander("Dernière publication de la société (soldes de l'exercice en cours)", expanded=False):
        st.markdown(
            "Les rapports d'activités trimestriels et semestriels et les états financiers annuels donnent les soldes "
            "de l'exercice en cours (chiffre d'affaires ou PNB, EBE, résultat d'exploitation, RAO, coût du risque, "
            "résultat net) face à la même période de l'an passé. **Données en ligne activées** : la publication la "
            "plus récente est recherchée automatiquement sur richbourse.com (relais des avis BRVM). Vous pouvez "
            "aussi charger le PDF ou coller son adresse ; votre document est alors prioritaire.")
        rep_file = st.file_uploader("PDF du rapport d'activités ou des états financiers", type=["pdf"],
                                    key=f"rep_{ticker}",
                                    help="Document publié par la société (brvm.org, richbourse.com, site de la "
                                         "société). Le PDF doit contenir du texte : un document scanné n'est pas lu.")
        rep_url = st.text_input("Ou adresse du document (lien « afficher-fichier » richbourse ou PDF)",
                                key=f"repurl_{ticker}", placeholder="https://www.richbourse.com/common/actualite/afficher-fichier/...")
        rep_title = st.text_input("Période couverte (facultatif)", key=f"reptitle_{ticker}",
                                  placeholder="Ex. Rapport d'activités - 1er semestre 2026",
                                  help="Sert à nommer la période si le document ne l'indique pas clairement.")
        if rep_file is not None:
            report = _report_from_pdf(rep_file.getvalue(), rep_title or rep_file.name.rsplit(".", 1)[0])
            if report is None:
                st.error("Aucun tableau d'indicateurs reconnu dans ce PDF (document scanné ou mise en page atypique).")
        elif rep_url.strip():
            report = _report_from_url(rep_url.strip())
            if report is None:
                st.error("Document inaccessible ou illisible à cette adresse.")
        if report is not None:
            st.caption(f"Document lu : {report.period or 'période non identifiée'}, {len(report.lines)} indicateurs, "
                       f"montants en {report.unit_label}.")
            st.dataframe(pd.DataFrame([(ln.label, fmt_num(ln.current / 1e9, 1), fmt_num(ln.previous / 1e9, 1),
                                        fmt_pct(ln.variation, sign=True) if is_num(ln.variation) else "n.s.")
                                       for ln in report.lines],
                                      columns=["Indicateur (Mds FCFA)", report.period or "N", report.period_prev or "N-1",
                                               "Variation"]), hide_index=True, width="stretch")
    run_key = (ticker, ef_name, len(ef_bytes or b""), online, tuple(sorted(overrides.items())), rf, erp, g, beta,
               tuple(sorted(w.items())), s_buy, s_sell, major, notes, use_ai, tuple(sorted(b.items())), house, analyst,
               px_file.name if px_file else None, use_def, len(mkt_bytes or b""),
               (report.title, report.period, len(report.lines)) if report is not None else None)
    c1, c2 = st.columns([1, 3])
    go = c1.button("Lancer l'analyse", type="primary", width="stretch")
    if go or (st.session_state.get("run_key") != run_key and st.session_state.get("auto", False)):
        if go:
            st.session_state["figs_used"] = (ticker, figures)
        used = st.session_state.get("figs_used")
        figs_used = used[1] if used and used[0] == ticker else None
        price_df = _prices(px_file.getvalue(), px_file.name) if px_file is not None else None
        with st.status("Analyse en cours...", expanded=False) as status:
            a = run_analysis(wb, ticker, settings, market_base=mb, price_file=price_df, overrides=overrides,
                             user_notes=notes, major_event=major, bceao=b, use_claude=use_ai,
                             progress=lambda msg: status.update(label=msg), figures=figs_used, report=report)
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
    basis = (f"comptes {c.kind.upper()} en {c.unit_label}" if c.level in ("complet", "partiel")
             else config.LEVELS[c.level][0])
    h1.markdown(f'<div class="fv-head"><div class="t">{clean_text(c.name)}</div>'
                f'<div class="s">{c.ticker} | {c.sector} | {country} | {basis}</div></div>', unsafe_allow_html=True)
    h2.markdown(f'<div style="text-align:center"><span class="fv-badge" style="background:{color}">{d.signal}'
                f'</span><div class="fv-small">Conviction {d.conviction.lower()} | note '
                f'{(fmt_num(d.composite) + "/100") if is_num(d.composite) else "n.d."}</div></div>', unsafe_allow_html=True)

    if a.coverage is not None and c.level != "complet":
        miss = ", ".join(i.name.lower() for i in a.coverage.missing[:4])
        st.warning(f"**{config.LEVELS[c.level][0]}** | couverture des données {fmt_num(a.coverage.score, 0)}%. "
                   f"{config.LEVELS[c.level][1]} Conviction plafonnée."
                   + (f" À compléter : {miss}. Détail dans l'onglet « Données et méthode »." if miss else ""))
    elif a.coverage is not None:
        st.caption(f"Niveau 1, analyse complète | couverture des données {fmt_num(a.coverage.score, 0)}%")
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
        if c.level in ("chiffres_cles", "marche"):
            if c.level == "chiffres_cles":
                st.subheader("Chiffres clés")
                yrs = c.years[-5:]
                rows = {}
                for lab_, key_, div_ in ((("PNB" if c.is_bank else "Chiffre d'affaires") + " (Mds FCFA)", c.revenue_field, 1e9),
                                         ("Résultat net (Mds FCFA)", c.net_income_field, 1e9),
                                         ("Capitaux propres (Mds FCFA)", c.equity_field, 1e9),
                                         ("Total du bilan (Mds FCFA)", "total_actif", 1e9),
                                         ("BNPA (FCFA)", "bpa_pub", 1.0), ("Dividende net (FCFA)", "dps", 1.0)):
                    ser_ = c.fields.get(key_, {})
                    if any(is_num(ser_.get(y_)) for y_ in yrs):
                        rows[lab_] = [fmt_num(ser_[y_] / div_, 1) if is_num(ser_.get(y_)) else "n.d." for y_ in yrs]
                for key_, s_ in fd.ratios.items():
                    if key_ in ("marge_nette", "croissance_pnb", "croissance_ca", "croissance_rn", "roe") and key_ in RATIO_META:
                        rows[RATIO_META[key_][0]] = [fmt_pct(s_.get(y_)) if is_num(s_.get(y_)) else "n.d." for y_ in yrs]
                st.dataframe(pd.DataFrame(rows, index=[str(y_) for y_ in yrs]).T, width="stretch")
                st.caption(f"Source : {c.data_source}.")
            st.info("Soldes intermédiaires de gestion, ratios de structure et solidité financière non disponibles "
                    "à ce niveau. Chargez le classeur des états financiers pour les obtenir.")
            st.write(tx.get("activite", ""))
        else:
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
        st.subheader("Publications officielles de la société")
        if a.interim is not None and a.interim.lines:
            rp = a.interim
            st.markdown(f"**{rp.title or rp.period}**" + (f" (publiée le {rp.published:%d/%m/%Y})" if rp.published else "")
                        + (f" | [document]({rp.url})" if rp.url else ""))
            st.dataframe(pd.DataFrame([(ln.label, fmt_num(ln.current / 1e9, 1), fmt_num(ln.previous / 1e9, 1),
                                        fmt_pct(ln.variation, sign=True) if is_num(ln.variation) else "n.s.")
                                       for ln in rp.lines],
                                      columns=["Indicateur (Mds FCFA)", rp.period or "N", rp.period_prev or "N-1",
                                               "Variation"]), hide_index=True, width="stretch")
        else:
            st.caption("Aucune publication exploitée : activez les données en ligne ou chargez le PDF du dernier "
                       "rapport d'activités dans l'onglet d'analyse.")
        if a.publications:
            st.dataframe(pd.DataFrame([(p_.published.strftime("%d/%m/%Y") if p_.published else "", p_.title, p_.pdf_url)
                                       for p_ in a.publications[:12]], columns=["Date", "Publication", "Document"]),
                         hide_index=True, width="stretch",
                         column_config={"Document": st.column_config.LinkColumn("Document", display_text="Ouvrir")})
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
        if c.level in ("complet", "partiel"):
            src["états financiers"] = f"{ef_name}, feuille {c.sheet} ({c.unit_label})"
        else:
            src["chiffres clés"] = c.data_source
        st.table(pd.DataFrame(list(src.items()), columns=["Donnée", "Source"]))
        if a.coverage is not None:
            st.subheader(f"Couverture des données : {fmt_num(a.coverage.score, 0)}% ({a.coverage.label})")
            lab_ = {"ok": "Disponible", "partiel": "Partiel", "absent": "Absent"}
            st.dataframe(pd.DataFrame([(i.name, lab_[i.status], i.detail, i.impact, i.action if i.status != "ok" else "")
                                       for i in a.coverage.items],
                                      columns=["Donnée", "Statut", "Détail", "Effet sur l'analyse", "Pour compléter"]),
                         hide_index=True, width="stretch")
        if c.notes:
            st.subheader("Contrôles et remarques sur les données")
            for n_ in c.notes:
                st.markdown(f"- {n_}")
        if a.errors:
            st.subheader("Accès aux sources")
            for e_ in a.errors:
                st.markdown(f"- {e_}")
        if c.field_labels:
            with st.expander("Libellés reconnus dans la feuille"):
                st.table(pd.DataFrame([(k_, v_) for k_, v_ in c.field_labels.items()],
                                      columns=["Champ", "Libellé source"]))
        st.caption(config.DISCLAIMER)

# ---------------------------------------------------------------------------
# Tableau de bord multi-valeurs
# ---------------------------------------------------------------------------
with tab_batch:
    st.write("Analyse en série de tous les titres de la cote : tableau des signaux, export Excel et archive des "
             "fiches PDF. Les titres hors classeur sont analysés aux niveaux 2 ou 3 (chiffres clés en ligne, "
             "sinon données de marché).")
    mkt_tickers = set(mb.tickers) if mb is not None else set()
    cover = pd.DataFrame([{
        "Ticker": t_, "Société": config.TICKERS.get(t_, {}).get("nom", t_),
        "Secteur": config.TICKERS.get(t_, {}).get("secteur", "n.d."),
        "Pays": config.COUNTRIES.get(config.TICKERS.get(t_, {}).get("pays", ""), {}).get("nom", "n.d."),
        "États financiers": "Oui" if t_ in in_wb else "Non",
        "Base de marché": "Oui" if t_ in mkt_tickers else "Non",
        "Chiffres clés SMF": "Oui" if figures_from_market_base(mb, t_) is not None else "Non",
        "Niveau attendu": ("Niveau 1" if t_ in in_wb else
                           ("Niveau 2" if figures_from_market_base(mb, t_) is not None else
                            ("Niveau 2 ou 3" if online else "Niveau 3"))),
    } for t_ in all_tickers])
    n1 = int((cover["États financiers"] == "Oui").sum())
    st.caption(f"{len(cover)} titres : {n1} avec états financiers (niveau 1), {len(cover) - n1} hors classeur.")
    with st.expander("Liste des titres et données disponibles"):
        st.dataframe(cover, hide_index=True, width="stretch")
    chosen = st.multiselect("Valeurs à analyser", all_tickers, default=all_tickers,
                            format_func=lambda t_: f"{t_} - {config.TICKERS.get(t_, {}).get('nom', t_)}")
    with_pdf = st.checkbox("Inclure les fiches PDF dans l'archive", value=True)
    if st.button("Analyser la sélection", type="primary", disabled=not chosen):
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
        for col in ("Potentiel", "Rendement", "ROE", "Couverture"):
            show[col] = show[col].map(lambda x: fmt_pct(x, 0 if col == "Couverture" else 1) if is_num(x) else "n.d.")
        for col in ("Note globale", "Vue fondamentale", "Technique", "Cours", "Juste valeur"):
            show[col] = show[col].map(lambda x: fmt_num(x) if is_num(x) else "n.d.")
        for col in ("PER", "P/B"):
            show[col] = show[col].map(lambda x: fmt_x(x, 1) if is_num(x) else "n.d.")
        st.dataframe(show, hide_index=True, width="stretch")
        counts = df["Signal"].value_counts()
        st.caption(" | ".join(f"{k_} : {v_}" for k_, v_ in counts.items())
                   + " | N1 : états financiers, N2 : chiffres clés, N3 : marché seul (conviction plafonnée).")
        st.download_button("Télécharger l'archive (tableau Excel + fiches)", st.session_state["batch_zip"],
                           f"Fiches_BRVM_{date.today():%Y%m%d}.zip", "application/zip", type="primary",
                           on_click="ignore")
        for e_ in st.session_state.get("batch_errors", []):
            st.error(e_)
