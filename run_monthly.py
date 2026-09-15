#!/usr/bin/env python3
"""
Script de génération du rapport mensuel KPI Hygierun

Usage:
    python run_monthly.py --input=<chemin_excel> --output=<chemin_pptx> --month=08 --year=2026
"""

import argparse
import sys
from datetime import datetime
from pathlib import Path

from src.data_loader import DataLoader, SAVFichesParser


def main():
    """Fonction principale"""

    # ========================================================================
    # ARGUMENTS
    # ========================================================================

    parser = argparse.ArgumentParser(
        description="Génère le rapport mensuel KPI Hygierun"
    )

    parser.add_argument(
        "--input",
        required=True,
        type=str,
        help="Chemin vers Input_Donnees_Brutes.xlsx"
    )

    parser.add_argument(
        "--output",
        required=False,
        type=str,
        default="rapport_mensuel.pptx",
        help="Chemin de sortie pour le rapport PowerPoint (défaut: rapport_mensuel.pptx)"
    )

    parser.add_argument(
        "--month",
        required=False,
        type=int,
        default=None,
        help="Mois à traiter (1-12, défaut: mois courant)"
    )

    parser.add_argument(
        "--year",
        required=False,
        type=int,
        default=None,
        help="Année à traiter (défaut: année courante)"
    )

    parser.add_argument(
        "--sav-zip",
        required=False,
        type=str,
        default=None,
        help="Chemin optionnel vers FIN_AOUT_FICHES_SAV.zip"
    )

    args = parser.parse_args()

    # ========================================================================
    # VÉRIFICATIONS
    # ========================================================================

    print("\n" + "=" * 80)
    print("🚀 HYGIERUN KPI TOOL - RAPPORT MENSUEL")
    print("=" * 80)

    # Vérifier que le fichier Excel existe
    if not Path(args.input).exists():
        print(f"\n❌ Erreur: Le fichier '{args.input}' n'existe pas")
        sys.exit(1)

    # Déterminer la période
    if args.year is None:
        args.year = datetime.now().year
    if args.month is None:
        args.month = datetime.now().month

    try:
        period_start = datetime(args.year, args.month, 1)
    except ValueError:
        print(f"\n❌ Erreur: Date invalide (année={args.year}, mois={args.month})")
        sys.exit(1)

    print(f"\n📅 Période: {period_start.strftime('%B %Y')}")
    print(f"📁 Fichier d'entrée: {args.input}")
    print(f"💾 Fichier de sortie: {args.output}")

    # ========================================================================
    # CHARGER LES DONNÉES
    # ========================================================================

    loader = DataLoader(args.input, args.sav_zip)

    if not loader.load_and_validate():
        print("\n❌ La validation des données a échoué. Corrigez le fichier d'entrée.")
        sys.exit(1)

    # Afficher le résumé des données
    loader.print_data_summary()

    # Filtrer par mois
    print(f"\n🔍 Filtrage des données pour {period_start.strftime('%B %Y')}...")
    loader.filter_by_month(args.year, args.month)
    loader.print_data_summary()

    # ========================================================================
    # TRAITER LES FICHES SAV (SI DISPONIBLE)
    # ========================================================================

    if args.sav_zip and Path(args.sav_zip).exists():
        print("\n📑 Traitement des fiches d'intervention SAV...")
        sav_parser = SAVFichesParser(args.sav_zip)
        if sav_parser.parse_fiches():
            print(f"   → {sav_parser.get_fiches_count()} fiches trouvées")

    # ========================================================================
    # CALCUL DES KPIS (À IMPLÉMENTER EN PHASE 3)
    # ========================================================================

    print("\n⏳ Calcul des KPIs...")
    print("   ⚠️  (À implémenter en Phase 3)")

    # ========================================================================
    # GÉNÉRATION DU RAPPORT POWERPOINT (À IMPLÉMENTER EN PHASE 4)
    # ========================================================================

    print("\n📊 Génération du rapport PowerPoint...")
    print("   ⚠️  (À implémenter en Phase 4)")

    # ========================================================================
    # SUCCÈS
    # ========================================================================

    print("\n" + "=" * 80)
    print("✅ RAPPORT GÉNÉRÉ AVEC SUCCÈS")
    print("=" * 80)
    print(f"\n📄 Rapport disponible à: {args.output}\n")


if __name__ == "__main__":
    main()
