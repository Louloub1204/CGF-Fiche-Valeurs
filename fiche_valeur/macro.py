"""Environnement macroéconomique et monétaire du pays de l'émetteur.

* Indicateurs pays : API publique DataMapper du FMI (World Economic Outlook),
  sans clé ; repli sur les valeurs embarquées dans ``config.MACRO_FALLBACK``.
* Politique monétaire : décisions du Comité de Politique Monétaire de la
  BCEAO (paramètres modifiables dans l'interface).
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date

import requests

from . import config
from .utils import clamp, is_num, scale

IMF_API = "https://www.imf.org/external/datamapper/api/v1/{ind}/{iso3}"

SECTOR_SENSITIVITY = {
    config.FIN: "Banques : sensibles à la croissance du crédit, au coût du risque et à la liquidité "
                "régionale ; un taux directeur stable à bas niveau soutient les volumes mais comprime "
                "les marges d'intermédiation. Exposition aux titres souverains UMOA à surveiller.",
    config.TEL: "Télécommunications : activité défensive, portée par la data et le mobile money ; "
                "risques réglementaires et fiscaux (taxes sectorielles, prix plafonds).",
    config.CDB: "Consommation de base et agro-industrie : demande résiliente ; marges exposées aux cours "
                "des matières premières (caoutchouc, huile de palme, sucre, cacao) et aux intrants.",
    config.CDI: "Consommation discrétionnaire : sensible au pouvoir d'achat, au crédit et au climat des "
                "affaires.",
    config.ENE: "Énergie (distribution de produits pétroliers) : marges encadrées par les structures de "
                "prix administrées ; sensibilité aux cours du pétrole et aux arriérés de compensation "
                "publique.",
    config.SPU: "Services publics : revenus régulés et récurrents, profil obligataire ; dépendance aux "
                "décisions tarifaires et aux créances sur l'État.",
    config.IND: "Industriels et transport : activité liée à l'investissement public, au commerce "
                "extérieur et à l'activité portuaire.",
}


@dataclass
class Macro:
    country: str
    country_name: str
    iso3: str
    data: dict[str, dict[int, float]] = field(default_factory=dict)
    source: str = ""
    year: int | None = None
    bceao: dict = field(default_factory=dict)
    score: float | None = None
    parts: list = field(default_factory=list)
    sector_note: str = ""
    notes: list[str] = field(default_factory=list)

    def value(self, key: str, year: int | None = None):
        s = self.data.get(key, {})
        if not s:
            return None, None
        if year is not None and year in s:
            return s[year], year
        ys = [y for y in s if y <= (year or date.today().year)]
        y = max(ys) if ys else max(s)
        return s[y], y


def fetch_imf(iso3: str, errors: list[str] | None = None) -> dict[str, dict[int, float]]:
    out: dict[str, dict[int, float]] = {}
    for ind, key in config.IMF_INDICATORS.items():
        try:
            r = requests.get(IMF_API.format(ind=ind, iso3=iso3), timeout=15,
                             headers={"User-Agent": "FicheValeurBRVM/1.0"})
            r.raise_for_status()
            vals = r.json().get("values", {}).get(ind, {}).get(iso3, {})
            series = {int(y): float(v) for y, v in vals.items() if is_num(v) and 2015 <= int(y) <= 2031}
            if series:
                out[key] = series
        except Exception as exc:
            if errors is not None:
                errors.append(f"FMI {ind} : {type(exc).__name__}")
            break
    return out


def analyse_macro(country: str, sector: str, online: bool = True, bceao: dict | None = None,
                  errors: list[str] | None = None) -> Macro:
    info = config.COUNTRIES.get(country, config.COUNTRIES["CI"])
    m = Macro(country=country, country_name=info["nom"], iso3=info["iso3"],
              bceao=dict(bceao or config.BCEAO), sector_note=SECTOR_SENSITIVITY.get(sector, ""))
    data = fetch_imf(m.iso3, errors) if online else {}
    if data:
        m.data = data
        m.source = "FMI, World Economic Outlook (API DataMapper)"
    else:
        fb = config.MACRO_FALLBACK.get(m.iso3, {})
        m.data = {k: dict(v) for k, v in fb.items()}
        m.source = "FMI, World Economic Outlook (valeurs embarquées, consultées le 18/10/2026)"
    year = date.today().year
    m.year = year
    g, gy = m.value("croissance", year)
    inf, iy = m.value("inflation", year)
    debt, dy = m.value("dette_publique", year)
    fb_, fy = m.value("solde_budgetaire", year)
    ca_, cy = m.value("solde_courant", year)
    if gy is not None and gy < year:
        m.notes.append(f"Dernière donnée de croissance disponible pour {m.country_name} : {gy}.")
    rate = m.bceao.get("taux_directeur")
    inf_s = clamp(100 - abs(inf - 2.0) * 20) if is_num(inf) else None
    m.parts = [
        ("Croissance du PIB réel", scale(g, 2.0, 8.0), 30),
        ("Inflation (proximité de la cible 1-3%)", inf_s, 15),
        ("Dette publique / PIB", scale(debt, 40.0, 90.0, reverse=True), 20),
        ("Solde budgétaire / PIB", scale(fb_, -7.0, 0.0), 10),
        ("Solde courant / PIB", scale(ca_, -10.0, 0.0), 10),
        ("Orientation monétaire (taux directeur)", scale(rate, 5.0, 2.0), 15),
    ]
    num = sum(s * w for _, s, w in m.parts if is_num(s))
    den = sum(w for _, s, w in m.parts if is_num(s))
    m.score = clamp(num / den) if den else None
    return m
