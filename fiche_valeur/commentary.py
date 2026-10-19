"""Rédaction automatique des commentaires (français institutionnel, concis,
écarts en points de base), avec enrichissement facultatif par l'API Claude."""
from __future__ import annotations

import json

import requests

from .config import SIGNAL_DEFINITIONS
from . import config
from .utils import (clean_text, fmt_bp, fmt_fcfa, fmt_num, fmt_pct, fmt_price, fmt_pts, fmt_x,
                    growth, is_num)

ANTHROPIC_URL = "https://api.anthropic.com/v1/messages"


def _rev_word(c) -> str:
    return "produit net bancaire" if c.is_bank else "chiffre d'affaires"


def _delta(series: dict, y: int):
    if y in series and (y - 1) in series and is_num(series[y]) and is_num(series[y - 1]):
        return series[y] - series[y - 1]
    return None


def _bp_txt(d) -> str:
    """Écart en points de base, « stable » sous 5 pb."""
    if not is_num(d):
        return ""
    return "stable" if abs(d) < 0.0005 else fmt_bp(d) + " sur un an"


def build_commentary(a) -> dict:
    """a : objet Analysis (pipeline). Retourne un dictionnaire de textes."""
    c, fd, val, tech, mac, news, dec = a.company, a.fundamentals, a.valuation, a.technical, a.macro, a.news, a.decision
    y = fd.year
    r = fd.ratios
    f = c.fields
    rev = f.get(c.revenue_field, {})
    rn = f.get(c.net_income_field, {})
    out: dict = {}

    # ---- Activité et résultats ------------------------------------------------
    s = []
    if y and is_num(rev.get(y)):
        g_rev = growth(rev.get(y), rev.get(y - 1))
        s.append(f"En {y}, {c.name} a réalisé un {_rev_word(c)} de {fmt_fcfa(rev[y])}"
                 + (f" ({fmt_pct(g_rev, sign=True)} sur un an" if is_num(g_rev) else "")
                 + (f" ; TCAM de {fmt_pct(fd.growth.get('tcam_rev'))} sur {fd.growth.get('tcam_rev_periode')})"
                    if is_num(fd.growth.get("tcam_rev")) and is_num(g_rev) else (")" if is_num(g_rev) else ""))
                 + ".")
    if c.is_bank:
        ce = r.get("coef_exploitation", {})
        cr = r.get("cout_risque_pnb", {})
        if is_num(ce.get(y)):
            d = _delta(ce, y)
            s.append(f"Le coefficient d'exploitation s'établit à {fmt_pct(ce[y])}"
                     + (f" ({_bp_txt(d)})" if is_num(d) else "") + ".")
        if is_num(cr.get(y)):
            s.append(f"Le coût du risque absorbe {fmt_pct(cr[y])} du PNB.")
    else:
        me = r.get("marge_ebe", {})
        if is_num(f.get("ebe", {}).get(y)) and is_num(me.get(y)):
            d = _delta(me, y)
            s.append(f"L'EBE ressort à {fmt_fcfa(f['ebe'][y])}, soit une marge de {fmt_pct(me[y])}"
                     + (f" ({_bp_txt(d)})" if is_num(d) else "") + ".")
    if is_num(rn.get(y)):
        g_rn = growth(rn.get(y), rn.get(y - 1))
        lab = "Le résultat net part du groupe" if c.net_income_field == "rn_pg" else "Le résultat net"
        mn = r.get("marge_nette", {}).get(y)
        roe = r.get("roe", {}).get(y)
        s.append(f"{lab} atteint {fmt_fcfa(rn[y])}" + (f" ({fmt_pct(g_rn, sign=True)})" if is_num(g_rn) else "")
                 + (f", pour une marge nette de {fmt_pct(mn)}" if is_num(mn) else "")
                 + (f" et un ROE de {fmt_pct(roe)}" if is_num(roe) else "") + ".")
    if c.level == "marche":
        s = [f"Aucun état financier n'est disponible pour {c.name} : l'analyse repose sur les données de marché "
             f"(cours, volumes, dividendes), l'environnement macroéconomique et les actualités. La qualité et la "
             f"croissance des résultats ne sont pas évaluées."]
    elif c.level == "chiffres_cles":
        s.append("Analyse sur chiffres clés : le bilan et les soldes intermédiaires de gestion ne sont pas "
                 "disponibles.")
    rep = getattr(a, "interim", None)
    pub_txt = []
    if rep is not None and rep.lines:
        rev = rep.get("pnb") or rep.get("ca")
        res = rep.get("rn") or rep.get("rao") or rep.get("re")
        bits = []
        for ln in (rev, res):
            if ln is not None and is_num(ln.current):
                bits.append(f"{ln.label.lower()} de {fmt_fcfa(ln.current)}"
                            + (f" ({fmt_pct(ln.variation, sign=True)} sur un an)" if is_num(ln.variation) else ""))
        if bits:
            when = f" ({rep.published:%d/%m/%Y})" if rep.published else ""
            pub_txt.append(f"Dernière publication, {rep.period or 'période récente'}{when} : " + " et ".join(bits) + ".")
        s += pub_txt
    out["publication"] = pub_txt
    out["activite"] = " ".join(s)

    # ---- Valorisation -----------------------------------------------------------
    v = []
    if is_num(val.price):
        mult = []
        if is_num(val.pe):
            mult.append(f"un PER de {fmt_x(val.pe)}")
        if is_num(val.pb):
            mult.append(f"un P/B de {fmt_x(val.pb, 2)}")
        if not c.is_bank and is_num(val.ev_ebitda):
            mult.append(f"un VE/EBE de {fmt_x(val.ev_ebitda)}")
        med = fd.medians
        medtxt = []
        if is_num(med.get("pe")):
            medtxt.append(f"PER {fmt_x(med['pe'])}")
        if is_num(med.get("pb")):
            medtxt.append(f"P/B {fmt_x(med['pb'], 2)}")
        v.append(f"Au cours de {fmt_price(val.price)}, "
                 + ("le titre se traite sur " + ", ".join(mult) if mult else
                    "les multiples de résultat ne sont pas calculables faute de résultat net ou de nombre de titres")
                 + (f" (références {fd.peer_label} : {', '.join(medtxt)})" if medtxt else "")
                 + (f" et offre un rendement net de {fmt_pct(val.div_yield)}" if is_num(val.div_yield) else "")
                 + ".")
    if is_num(val.fair_value):
        v.append(f"La juste valeur multicritère ressort à {fmt_price(val.fair_value)} "
                 f"(fourchette {fmt_price(val.fv_low)} - {fmt_price(val.fv_high)}), soit un potentiel de "
                 f"{fmt_pct(val.upside, sign=True)}"
                 + (f" et une rentabilité totale attendue de {fmt_pct(val.total_return, sign=True)} dividende inclus"
                    if is_num(val.total_return) and is_num(val.div_yield) else "") + ".")
    if is_num(val.implied_roe) and is_num(val.roe_norm):
        v.append(f"Ce que le cours suppose : un ROE durable de {fmt_pct(val.implied_roe)} "
                 f"pour un ROE normatif observé de {fmt_pct(val.roe_norm)} (coût des fonds propres "
                 f"{fmt_pct(val.cost_of_equity)}, bêta {fmt_num(val.beta, 2)}).")
    out["valorisation"] = " ".join(v)

    # ---- Technique ----------------------------------------------------------------
    t = []
    if tech.available:
        p = tech.performance
        perf = [f"{lab} {fmt_pct(p.get(k), sign=True)}" for k, lab in (("ytd", "depuis le 1er janvier"),
                                                                       ("12m", "sur un an")) if is_num(p.get(k))]
        t.append(f"La tendance de fond est {tech.trend}" + (f" ({', '.join(perf)})" if perf else "") + ".")
        rel = p.get("relative_6m")
        if is_num(rel):
            t.append(f"Sur six mois, le titre {'surperforme' if rel > 0 else 'sous-performe'} le BRVM Composite "
                     f"de {fmt_pct(abs(rel))}.")
        rsi = tech.indicators.get("rsi")
        if is_num(rsi):
            zone = "surachat" if rsi >= 70 else ("survente" if rsi <= 30 else "neutre")
            t.append(f"Le RSI (14) s'établit à {fmt_num(rsi)} (zone {zone}).")
        lv = tech.levels
        if lv.get("support_1") and lv.get("resistance_1"):
            t.append(f"Support à {fmt_price(lv['support_1'])}, résistance à {fmt_price(lv['resistance_1'])}.")
        vol = tech.indicators.get("volatilite")
        if is_num(vol):
            t.append(f"Volatilité annualisée de {fmt_pct(vol)}"
                     + (f", repli maximal sur un an de {fmt_pct(tech.indicators.get('drawdown_max'))}"
                        if is_num(tech.indicators.get("drawdown_max")) else "") + ".")
    else:
        t.append("Historique de cours insuffisant : l'analyse technique n'a pas pu être conduite.")
    out["technique"] = " ".join(t)

    # ---- Macro ----------------------------------------------------------------------
    m = []
    g, gy = mac.value("croissance", mac.year)
    inf, iy = mac.value("inflation", mac.year)
    debt, dy = mac.value("dette_publique", mac.year)
    if is_num(g):
        m.append(f"{mac.country_name} : croissance du PIB réel attendue à {fmt_pts(g, sign=False)} en {gy}"
                 + (f", inflation à {fmt_pts(inf, sign=False)}" if is_num(inf) else "")
                 + (f", dette publique à {fmt_pts(debt, 0, sign=False)} du PIB" if is_num(debt) else "") + ".")
    b = mac.bceao
    m.append(f"BCEAO : taux directeur à {fmt_pts(b.get('taux_directeur'), 2, sign=False)}, guichet de prêt "
             f"marginal à {fmt_pts(b.get('guichet_pret_marginal'), 2, sign=False)} ; inflation de l'Union "
             f"attendue à {fmt_pts(b.get('inflation_uemoa_prevue'), sign=False)} et croissance à "
             f"{fmt_pts(b.get('croissance_uemoa_prevue'), sign=False)} (CPM du {b.get('date_reference')}).")
    if mac.sector_note:
        m.append(mac.sector_note)
    out["macro"] = " ".join(m)

    # ---- Actualités -----------------------------------------------------------------
    n = []
    if news.items:
        tone = news.sentiment or 0
        n.append(f"{len(news.items)} article(s) recensé(s) sur douze mois ; tonalité "
                 f"{'positive' if tone > 0.15 else 'négative' if tone < -0.15 else 'neutre'}.")
        if news.events:
            n.append("Thèmes relevés : " + ", ".join(f"{k.lower()} ({len(v)})" for k, v in news.events.items()) + ".")
    elif news.error:
        n.append(news.error)
    else:
        n.append("Aucune actualité significative recensée.")
    if news.user_notes:
        n.append("Informations communiquées par la société : " + news.user_notes)
    out["actualites"] = " ".join(n)

    # ---- Forces et vigilances ---------------------------------------------------------
    out["forces"], out["vigilances"] = _strengths_risks(a)

    # ---- Thèse en deux paragraphes ----------------------------------------------------
    p1 = out["activite"]
    sig = dec.signal
    p2 = []
    if out["valorisation"]:
        p2.append(out["valorisation"].split(". Ce que")[0].rstrip(".") + ".")
    if tech.available:
        p2.append(f"Sur le plan technique, la tendance est {tech.trend}.")
    p2.append(f"Nous retenons un signal {sig} (conviction {dec.conviction.lower()})"
              + (f", cours cible à 12 mois de {fmt_price(dec.target)}" if is_num(dec.target) else "") + ".")
    out["these"] = [clean_text(p1), clean_text(" ".join(p2))]

    # ---- Justification du signal --------------------------------------------------------
    j = list(dec.trace)
    if dec.entry_zone:
        j.append(f"Zone d'accumulation : {fmt_price(dec.entry_zone[0])} - {fmt_price(dec.entry_zone[1])}.")
    if is_num(dec.stop):
        j.append(f"Niveau d'invalidation : {fmt_price(dec.stop)} (clôture sous ce seuil).")
    j.append(SIGNAL_DEFINITIONS.get(sig, ""))
    out["justification"] = [clean_text(x) for x in j if x]
    for k in ("activite", "valorisation", "technique", "macro", "actualites"):
        out[k] = clean_text(out[k])
    return out


def _strengths_risks(a) -> tuple[list[str], list[str]]:
    c, fd, val, tech, mac, news, dec = a.company, a.fundamentals, a.valuation, a.technical, a.macro, a.news, a.decision
    y = fd.year
    r = fd.ratios
    g = lambda k: r.get(k, {}).get(y)  # noqa: E731
    S, R = [], []
    roe = g("roe")
    if is_num(roe):
        if roe >= 0.18:
            S.append(f"Rentabilité élevée : ROE de {fmt_pct(roe)}.")
        elif roe < 0.08:
            R.append(f"Rentabilité faible : ROE de {fmt_pct(roe)}, inférieur au coût des fonds propres.")
    tr = fd.growth.get("tcam_rev")
    if is_num(tr):
        if tr >= 0.08:
            S.append(f"Croissance soutenue des revenus (TCAM de {fmt_pct(tr)}).")
        elif tr < 0:
            R.append(f"Revenus en contraction (TCAM de {fmt_pct(tr)}).")
    lr = fd.growth.get("last_rn")
    if is_num(lr) and lr <= -0.10:
        R.append(f"Recul du résultat net de {fmt_pct(abs(lr))} sur le dernier exercice.")
    elif is_num(lr) and lr >= 0.15:
        S.append(f"Résultat net en progression de {fmt_pct(lr)} sur le dernier exercice.")
    if c.is_bank:
        ce, cr, ltd = g("coef_exploitation"), g("cout_risque_pnb"), g("ltd")
        if is_num(ce) and ce <= 0.50:
            S.append(f"Efficacité opérationnelle : coefficient d'exploitation de {fmt_pct(ce)}.")
        elif is_num(ce) and ce >= 0.65:
            R.append(f"Coefficient d'exploitation élevé ({fmt_pct(ce)}).")
        if is_num(cr) and cr >= 0.25:
            R.append(f"Coût du risque élevé ({fmt_pct(cr)} du PNB).")
        if is_num(ltd) and ltd > 0.95:
            R.append(f"Transformation tendue : crédits/dépôts à {fmt_pct(ltd)}.")
        npl = g("npl")
        if is_num(npl) and npl > 0.10:
            R.append(f"Créances en souffrance élevées ({fmt_pct(npl)}).")
    else:
        me = r.get("marge_ebe", {})
        d = _delta(me, y)
        if is_num(d) and d <= -0.01:
            R.append(f"Érosion de la marge d'EBE ({fmt_bp(d)} sur un an).")
        elif is_num(d) and d >= 0.01:
            S.append(f"Amélioration de la marge d'EBE ({fmt_bp(d)} sur un an).")
        gear = g("gearing")
        if is_num(gear) and gear <= 0:
            S.append("Trésorerie nette positive : bilan non endetté.")
        elif is_num(gear) and gear > 1.0:
            R.append(f"Endettement net élevé (gearing de {fmt_pct(gear)}).")
        if fd.solidity.altman_zone == "détresse":
            R.append("Score Z'' d'Altman en zone de détresse.")
        elif fd.solidity.altman_zone == "sécurité":
            S.append("Score Z'' d'Altman en zone de sécurité.")
    if fd.solidity.maximum:
        sc = fd.solidity
        if sc.score / sc.maximum >= 0.7:
            S.append(f"Score de solidité de {sc.score}/{sc.maximum}.")
        elif sc.score / sc.maximum <= 0.35:
            R.append(f"Score de solidité faible ({sc.score}/{sc.maximum}).")
    if is_num(val.div_yield) and val.div_yield >= 0.06:
        S.append(f"Rendement du dividende attractif ({fmt_pct(val.div_yield)} net).")
    if is_num(val.upside):
        if val.upside >= 0.15:
            S.append(f"Décote de valorisation : potentiel de {fmt_pct(val.upside, sign=True)}.")
        elif val.upside <= -0.15:
            R.append(f"Valorisation exigeante : {fmt_pct(abs(val.upside))} au-dessus de la juste valeur.")
    if tech.available:
        if tech.trend == "haussière":
            S.append("Tendance technique haussière.")
        elif tech.trend == "baissière":
            R.append("Tendance technique baissière.")
        rsi = tech.indicators.get("rsi")
        if is_num(rsi) and rsi >= 70:
            R.append(f"Titre en surachat (RSI de {fmt_num(rsi)}).")
    taux = tech.liquidity.get("taux_seances_actives") if tech.available else None
    if is_num(taux) and taux < 0.35:
        R.append("Liquidité réduite sur le marché secondaire.")
    debt, _ = mac.value("dette_publique", mac.year)
    if is_num(debt) and debt > 70:
        R.append(f"Endettement public élevé du pays d'implantation ({fmt_pts(debt, 0, sign=False)} du PIB).")
    for w in dec.warnings:
        if "Actualité" in w:
            R.append(w)
    rep = getattr(a, "interim", None)
    if rep is not None:
        for ln in (rep.get("pnb") or rep.get("ca"), rep.get("rn")):
            if ln is None or not is_num(ln.variation):
                continue
            if ln.variation >= 0.08:
                S.append(f"{ln.label} {rep.period} en hausse de {fmt_pct(ln.variation)} sur un an.")
            elif ln.variation <= -0.08:
                R.append(f"{ln.label} {rep.period} en recul de {fmt_pct(abs(ln.variation))} sur un an.")
    if c.level in ("chiffres_cles", "marche"):
        R.insert(0, f"Données limitées ({config.LEVELS[c.level][0].split(' - ')[1].lower()}) : conviction plafonnée, "
                    f"signal à confirmer avec les états financiers.")
    elif c.notes:
        R.append("Qualité des données : " + str(len(c.notes)) + " point(s) relevé(s) dans les états financiers "
                 "(voir annexe méthodologique).")
    return [clean_text(x) for x in S[:6]], [clean_text(x) for x in R[:6]]


# ---------------------------------------------------------------------------
# Enrichissement facultatif par Claude
# ---------------------------------------------------------------------------

def claude_enrich(a, texts: dict, api_key: str, model: str = "claude-sonnet-5-5", timeout: int = 60) -> dict | None:
    """Reformule la thèse d'investissement à partir des seules données calculées.

    Le modèle ne reçoit que des chiffres déjà calculés par l'application et la
    consigne de ne rien inventer ; il ne modifie jamais le signal.
    """
    if not api_key:
        return None
    payload = {
        "societe": a.company.name, "ticker": a.company.ticker, "secteur": a.company.sector,
        "pays": a.macro.country_name, "signal": a.decision.signal, "conviction": a.decision.conviction,
        "activite": texts.get("activite"), "valorisation": texts.get("valorisation"),
        "technique": texts.get("technique"), "macro": texts.get("macro"), "actualites": texts.get("actualites"),
        "forces": texts.get("forces"), "vigilances": texts.get("vigilances"),
        "titres_actualites": [n.title for n in a.news.items[:10]],
    }
    prompt = (
        "Tu es analyste actions sur la BRVM pour une société de gestion. À partir UNIQUEMENT des éléments JSON "
        "ci-dessous (n'invente aucun chiffre, aucune information), rédige en français institutionnel concis :\n"
        "1) une thèse d'investissement en deux paragraphes (performance puis perspectives), 90 à 140 mots au total ;\n"
        "2) trois catalyseurs et trois risques, une phrase chacun.\n"
        "Le signal est fixé par le modèle quantitatif : ne le modifie pas. Écris les pourcentages collés au nombre "
        "(12,5%), n'utilise pas de tiret long. Réponds en JSON strict : "
        '{"these": ["...", "..."], "catalyseurs": ["..."], "risques": ["..."]}\n\n'
        + json.dumps(payload, ensure_ascii=False)
    )
    try:
        resp = requests.post(
            ANTHROPIC_URL, timeout=timeout,
            headers={"x-api-key": api_key, "anthropic-version": "2023-06-01", "content-type": "application/json"},
            json={"model": model, "max_tokens": 1200, "messages": [{"role": "user", "content": prompt}]},
        )
        resp.raise_for_status()
        text = "".join(b.get("text", "") for b in resp.json().get("content", []) if b.get("type") == "text")
        start, end = text.find("{"), text.rfind("}")
        data = json.loads(text[start:end + 1])
        return {
            "these": [clean_text(x) for x in data.get("these", [])][:2],
            "catalyseurs": [clean_text(x) for x in data.get("catalyseurs", [])][:4],
            "risques": [clean_text(x) for x in data.get("risques", [])][:4],
        }
    except Exception:
        return None
