"""Génération en ligne de commande (une valeur ou tout le classeur).

Exemples
--------
    python cli.py --ef Etats_Financiers_2021_2025.xlsx --ticker SNTS
    python cli.py --ef Etats_Financiers_2021_2025.xlsx --smf "Base de données - SMF.xlsx" --all --pptx
    python cli.py --ef sample_data/demo_etats_financiers.xlsx --smf sample_data/demo_base_marche.xlsx --all --hors-ligne
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from fiche_valeur import config
from fiche_valeur.config import Settings
from fiche_valeur.export import build_excel, build_screening_excel, screening_frame
from fiche_valeur.market import load_market_base
from fiche_valeur.parser import parse_workbook
from fiche_valeur.pipeline import run_analysis
from fiche_valeur.report_pdf import build_pdf
from fiche_valeur.report_pptx import build_pptx


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Fiches de valeur BRVM (PDF, PowerPoint, Excel)")
    ap.add_argument("--ef", help="Classeur des états financiers (facultatif : sans lui, niveaux 2 et 3)")
    ap.add_argument("--pub", help="Fichier des publications intermédiaires (T1, S1, T3)")
    ap.add_argument("--smf", help="Base de marché (Cours, Nb_titres, Moyenne_cours, Dividendes)")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--ticker", nargs="+", help="Ticker(s) à analyser")
    g.add_argument("--all", action="store_true",
                   help="Toute la cote : titres du classeur et titres du référentiel hors classeur")
    ap.add_argument("--out", default="sorties", help="Dossier de sortie (défaut : sorties)")
    ap.add_argument("--pptx", action="store_true", help="Produire aussi le rapport PowerPoint")
    ap.add_argument("--xlsx", action="store_true", help="Produire aussi le détail Excel de chaque analyse")
    ap.add_argument("--compact", action="store_true", help="Fiche PDF sur une page")
    ap.add_argument("--hors-ligne", action="store_true", help="Ne consulter aucune source en ligne")
    ap.add_argument("--rf", type=float, default=6.5, help="Taux sans risque, en pourcentage (défaut 6,5)")
    ap.add_argument("--prime", type=float, default=5.5, help="Prime de risque actions (défaut 5,5)")
    ap.add_argument("--g", type=float, default=3.0, help="Croissance perpétuelle (défaut 3)")
    ap.add_argument("--maison", default="CGF Gestion", help="Société de gestion affichée sur les documents")
    args = ap.parse_args(argv)

    wb = parse_workbook(args.ef) if args.ef else None
    if args.ef and not wb.companies:
        print("Aucune société reconnue dans le classeur.", file=sys.stderr)
        return 1
    in_wb = set(wb.companies) if wb is not None else set()
    known = set(config.TICKERS) | in_wb
    mb = load_market_base(args.smf) if args.smf else None
    from fiche_valeur.interim_file import parse_interim_file
    idata = parse_interim_file(args.pub) if args.pub else None
    st = Settings(taux_sans_risque=args.rf, prime_risque_marche=args.prime, croissance_long_terme=args.g,
                  en_ligne=not args.hors_ligne, maison=args.maison)
    if args.all:
        # en démonstration (classeur sans titre du référentiel), seules les sociétés du classeur
        tickers = sorted(known) if set(config.TICKERS) & in_wb or not in_wb else sorted(in_wb)
    else:
        tickers = [t.upper() for t in args.ticker]
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    results = []
    for t in tickers:
        if t not in known:
            print(f"{t} : titre inconnu (absent du référentiel et du classeur)", file=sys.stderr)
            continue
        try:
            a = run_analysis(wb, t, st, market_base=mb, interim_data=idata)
        except Exception as exc:
            print(f"{t} : échec de l'analyse ({type(exc).__name__}: {exc})", file=sys.stderr)
            continue
        results.append(a)
        (out / f"Fiche_{t}.pdf").write_bytes(build_pdf(a, compact=args.compact))
        if args.pptx:
            (out / f"Rapport_{t}.pptx").write_bytes(build_pptx(a))
        if args.xlsx:
            (out / f"Analyse_{t}.xlsx").write_bytes(build_excel(a))
        d = a.decision
        up = f"{d.upside * 100:+.1f}%".replace(".", ",") if d.upside is not None else "n.d."
        print(f"{t:6s} {d.signal:10s} conviction {d.conviction.lower():8s} potentiel {up:9s} "
              f"{config.LEVELS[a.company.level][0].split(' - ')[0]}")
    if results:
        (out / "Tableau_des_signaux.xlsx").write_bytes(build_screening_excel(screening_frame(results)))
        print(f"{len(results)} fiche(s) écrite(s) dans {out.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
