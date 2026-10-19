"""Synthèse multicritère et signal (ACHAT / CONSERVER / ATTENDRE / ALLÉGER / VENDRE).

Méthode
-------
1. Six piliers notés de 0 à 100 : valorisation, qualité & rentabilité,
   croissance, analyse technique, environnement macro, actualités.
2. Vue fondamentale F = moyenne pondérée des piliers hors technique.
   Note globale = moyenne pondérée des six piliers (pondérations réglables).
3. Matrice de décision : F (favorable / neutre / défavorable) croisée avec le
   potentiel U vis-à-vis de la juste valeur (décote / proche / surcote).
4. Le pilier technique sert de déclencheur (timing) : un achat sur tendance
   baissière devient « ATTENDRE », une position neutre sur tendance baissière
   et sans potentiel devient « ALLÉGER ».
5. Surcharges : événement majeur en cours déclaré par l'analyste (ATTENDRE),
   données insuffisantes (ATTENDRE), liquidité faible (conviction réduite).
"""
from __future__ import annotations

from dataclasses import dataclass, field

from .config import PILLAR_LABELS, Settings
from .utils import clamp, fmt_pct, is_num


@dataclass
class Decision:
    signal: str
    conviction: str
    composite: float | None
    fundamental: float | None
    technical: float | None
    pillars: dict
    weights: dict
    upside: float | None
    target: float | None
    total_return: float | None
    entry_zone: tuple | None = None
    stop: float | None = None
    fund_zone: str = ""
    val_zone: str = ""
    tech_zone: str = ""
    trace: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    level: str = "complet"


def _wavg(pillars: dict, weights: dict, keys) -> float | None:
    num = sum(pillars[k] * weights.get(k, 0) for k in keys if is_num(pillars.get(k)))
    den = sum(weights.get(k, 0) for k in keys if is_num(pillars.get(k)))
    return num / den if den else None


def decide(pillars: dict, st: Settings, valuation, technical, news=None, liquidity: dict | None = None,
           major_event: bool = False, level: str = "complet") -> Decision:
    w = st.poids
    fund_keys = [k for k in PILLAR_LABELS if k != "technique"]
    F = _wavg(pillars, w, fund_keys)
    F_shown = F
    if level == "marche":
        # Sans états financiers, la vue fondamentale est neutre par construction
        F = 50.0
        F_shown = None
    T = pillars.get("technique")
    composite = _wavg(pillars, w, list(PILLAR_LABELS))
    if sum(1 for k in PILLAR_LABELS if is_num(pillars.get(k))) < 3:
        composite = None      # trop peu de piliers renseignés pour une note globale significative
    U = valuation.upside if valuation else None
    d = Decision(signal="ATTENDRE", conviction="Moyenne", composite=composite, fundamental=F_shown, technical=T,
                 pillars=pillars, weights=dict(w), upside=U,
                 target=valuation.fair_value if valuation else None,
                 total_return=valuation.total_return if valuation else None, level=level)

    # Zones --------------------------------------------------------------------
    if is_num(F):
        d.fund_zone = ("favorable" if F >= st.seuil_fondamental_fort else
                       "neutre" if F >= st.seuil_fondamental_faible else "défavorable")
    if is_num(U):
        d.val_zone = ("décote" if U >= st.seuil_potentiel_achat / 100 else
                      "surcote" if U <= st.seuil_potentiel_vente / 100 else "proche de la juste valeur")
    if is_num(T):
        d.tech_zone = ("haussière" if T >= 55 else "baissière" if T < st.seuil_technique_defavorable else "neutre")
    tech_bear = d.tech_zone == "baissière"
    tech_bull = d.tech_zone == "haussière"

    # Données insuffisantes -----------------------------------------------------
    available = sum(1 for k in PILLAR_LABELS if is_num(pillars.get(k)))
    if not is_num(F) or (not is_num(U) and not is_num(T)):
        d.signal, d.conviction = "ATTENDRE", "Faible"
        d.trace.append("Données insuffisantes pour statuer (cours, nombre de titres ou états financiers incomplets).")
        return _finalise(d, valuation, technical, liquidity, available, news, major_event)

    if not is_num(U):
        d.trace.append("Juste valeur indisponible : décision fondée sur la vue fondamentale et la tendance.")
        if d.fund_zone == "favorable":
            d.signal = "CONSERVER" if not tech_bear else "ATTENDRE"
        elif d.fund_zone == "défavorable":
            d.signal = "ALLÉGER" if not tech_bull else "CONSERVER"
        else:
            d.signal = "CONSERVER" if tech_bull else "ATTENDRE"
        d.conviction = "Faible"
        return _finalise(d, valuation, technical, liquidity, available, news, major_event)

    # Matrice fondamentaux x valorisation ------------------------------------------
    matrix = {
        ("favorable", "décote"): "ACHAT",
        ("favorable", "proche de la juste valeur"): "CONSERVER",
        ("favorable", "surcote"): "ALLÉGER",
        ("neutre", "décote"): "ACHAT" if tech_bull else "ATTENDRE",
        ("neutre", "proche de la juste valeur"): "CONSERVER",
        ("neutre", "surcote"): "VENDRE" if tech_bear else "ALLÉGER",
        ("défavorable", "décote"): "ATTENDRE",
        ("défavorable", "proche de la juste valeur"): "VENDRE" if tech_bear else "ALLÉGER",
        ("défavorable", "surcote"): "VENDRE",
    }
    base = matrix[(d.fund_zone, d.val_zone)]
    phrase = {"décote": "en décote", "surcote": "en surcote",
              "proche de la juste valeur": "proche de sa juste valeur"}[d.val_zone]
    if level == "marche":
        d.trace.append(f"Sans états financiers, vue fondamentale neutre ; titre {phrase} "
                       f"(potentiel de {fmt_pct(U, sign=True)}) : signal de base {base}.")
    else:
        d.trace.append(f"Fondamentaux {d.fund_zone}s ({F:.0f}/100), titre {phrase} "
                       f"(potentiel de {fmt_pct(U, sign=True)}) : signal de base {base}.")
    sig = base
    if sig == "ACHAT" and tech_bear:
        sig = "ATTENDRE"
        d.trace.append("Tendance technique baissière : l'achat est différé dans l'attente d'un point d'entrée.")
    elif sig == "CONSERVER" and tech_bear and U < 0:
        sig = "ALLÉGER"
        d.trace.append("Tendance baissière sans potentiel de revalorisation : allègement préconisé.")
    elif sig == "CONSERVER" and tech_bull and U >= 0.08 and d.fund_zone == "favorable":
        d.trace.append("Tendance haussière et potentiel positif : renforcement possible sur repli.")
    elif sig == "ATTENDRE" and d.fund_zone == "défavorable":
        d.trace.append("Décote apparente mais fondamentaux fragiles : risque de « piège de valeur », "
                       "attendre une amélioration des résultats.")
    d.signal = sig

    # Conviction -------------------------------------------------------------------
    strong = (sig == "ACHAT" and U >= 1.8 * st.seuil_potentiel_achat / 100 and is_num(composite) and composite >= 65
              and not tech_bear) or \
             (sig == "VENDRE" and U <= 1.8 * st.seuil_potentiel_vente / 100 and is_num(composite) and composite <= 40)
    d.conviction = "Forte" if strong else "Moyenne"
    return _finalise(d, valuation, technical, liquidity, available, news, major_event)


def _downgrade(level: str) -> str:
    return {"Forte": "Moyenne", "Moyenne": "Faible"}.get(level, "Faible")


def _apply_level(d: Decision) -> None:
    """Prudence proportionnelle à la qualité des données."""
    if d.level == "marche":
        if d.signal == "ACHAT" and not (d.tech_zone == "haussière" and is_num(d.upside) and d.upside >= 0.25):
            d.signal = "ATTENDRE"
            d.trace.append("Données de marché seules : l'achat exige un potentiel d'au moins 25% et une tendance "
                           "haussière ; signal ramené à ATTENDRE.")
        elif d.signal == "VENDRE":
            d.signal = "ALLÉGER"
            d.trace.append("Données de marché seules : la vente intégrale n'est pas préconisée sans états "
                           "financiers ; signal ramené à ALLÉGER.")
        d.conviction = "Faible"
        d.warnings.append("Analyse sur données de marché seules : signal indicatif, à confirmer avec les états "
                          "financiers.")
    elif d.level in ("chiffres_cles", "partiel"):
        if d.conviction == "Forte":
            d.conviction = "Moyenne"
        if d.level == "chiffres_cles":
            d.warnings.append("Analyse sur chiffres clés : bilan et solidité financière non évalués, conviction "
                              "plafonnée.")


def _finalise(d: Decision, valuation, technical, liquidity, available, news, major_event) -> Decision:
    _apply_level(d)
    if major_event and d.signal != "VENDRE":
        d.trace.append("Événement majeur en cours signalé par l'analyste : position d'attente jusqu'à "
                       "clarification.")
        d.signal = "ATTENDRE"
    if news is not None and news.pending_major:
        d.warnings.append("Actualité récente à vérifier : " + ", ".join(news.pending_major).lower() + ".")
        d.conviction = _downgrade(d.conviction)
    if available < 4:
        d.warnings.append(f"Seulement {available} piliers sur 6 renseignés.")
        d.conviction = _downgrade(d.conviction)
    if valuation is not None and is_num(valuation.fv_low) and is_num(valuation.fv_high) and valuation.fv_low > 0:
        if valuation.fv_high / valuation.fv_low > 2.2:
            d.warnings.append("Forte dispersion entre méthodes de valorisation.")
            d.conviction = _downgrade(d.conviction)
    liq = liquidity or {}
    taux = liq.get("taux_seances_actives")
    if is_num(taux) and taux < 0.35:
        d.warnings.append(f"Liquidité faible ({taux * 100:.0f}% de séances actives sur 3 mois) : "
                          "exécution progressive et taille de position limitée.")
        if d.signal in ("ACHAT", "VENDRE"):
            d.conviction = _downgrade(d.conviction)
    # Zone d'entrée et niveau d'invalidation ------------------------------------
    if technical is not None and technical.available:
        lv = technical.levels
        price = technical.last
        mm50 = technical.indicators.get("mm50")
        if d.signal in ("ACHAT", "ATTENDRE") and lv.get("support_1"):
            hi = min(price, mm50) if is_num(mm50) else price
            lo = lv["support_1"]
            if lo < hi:
                d.entry_zone = (lo, hi)
        s2 = lv.get("support_2") or lv.get("support_1")
        if is_num(s2) and s2 < price:
            d.stop = s2 * 0.98
        elif is_num(price):
            d.stop = price * 0.88
    return d


def pillar_dict(valuation, fundamentals, technical, macro, news) -> dict:
    return {
        "valorisation": valuation.score if valuation else None,
        "qualite": fundamentals.quality_score if fundamentals else None,
        "croissance": fundamentals.growth_score if fundamentals else None,
        "technique": technical.score if technical and technical.available else None,
        "macro": macro.score if macro else None,
        "sentiment": news.score if news else None,
    }


def clamp_pillars(p: dict) -> dict:
    return {k: (clamp(v) if is_num(v) else None) for k, v in p.items()}
