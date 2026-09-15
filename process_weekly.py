#!/usr/bin/env python3
"""
Script d'intégration complet - Calcul des KPIs hebdomadaires et remplissage du snapshot

Usage:
    python process_weekly.py --week=37 --input=Input_Donnees_Brutes_Hebdo.xlsx --snapshot=KPIs_weekly.xlsx

Ce script :
1. Calcule tous les KPIs pour une semaine donnée
2. Crée des feuilles de détail par thématique (Commandes, Factures, Livraisons, Créances)
3. Remplit la ligne correspondante dans le fichier snapshot
4. Sauvegarde le fichier snapshot mis à jour
"""

import argparse
import sys
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from src.weekly_loader import WeeklyDataLoader
from src.kpi_calculator import KPICalculator


def get_week_dates(week_number: int, year: int = 2026) -> tuple:
    """
    Retourne (lundi, dimanche) d'une semaine donnée (ISO 8601)

    Args:
        week_number: Numéro de semaine (1-53)
        year: Année (défaut: 2026)

    Returns:
        (datetime_lundi, datetime_dimanche)
    """
    # Semaine ISO : lundi = jour 1, dimanche = jour 7
    jan_4 = datetime(year, 1, 4)
    week_1_monday = jan_4 - timedelta(days=jan_4.weekday())
    target_monday = week_1_monday + timedelta(weeks=week_number - 1)
    target_sunday = target_monday + timedelta(days=6)
    return target_monday, target_sunday


def get_row_for_week(week_number: int) -> int:
    """
    Calcule le numéro de ligne Excel pour une semaine donnée (s33-s52)

    Args:
        week_number: Numéro de semaine

    Returns:
        Numéro de ligne Excel (3 pour s33, 4 pour s34, ..., 22 pour s52)
    """
    # s33 est à la ligne 3, s34 à la ligne 4, etc.
    # Les lignes 1-2 sont les en-têtes
    return week_number - 33 + 3


def create_detail_sheets(
    file_input: str,
    file_output: str,
    week_number: int,
    calculator: KPICalculator,
    start_date: datetime,
    end_date: datetime
) -> None:
    """
    Crée les feuilles de détail par thématique dans le fichier snapshot

    Args:
        file_input: Chemin du fichier d'entrée
        file_output: Chemin du fichier snapshot à modifier
        week_number: Numéro de semaine
        calculator: Instance du calculateur KPI
        start_date: Date de début de la semaine
        end_date: Date de fin de la semaine
    """

    # Charger le fichier d'entrée
    loader = WeeklyDataLoader(file_input)
    loader.load_and_validate()
    dfs = loader.dfs

    # Charger le fichier snapshot
    wb = load_workbook(file_output)

    # Nom des feuilles de détail
    sheet_names = {
        'Commandes': f'Semaine_{week_number:02d}_Commandes',
        'Factures': f'Semaine_{week_number:02d}_Factures',
        'Livraisons': f'Semaine_{week_number:02d}_Livraisons',
        'Créances': f'Semaine_{week_number:02d}_Créances',
        'Delai': f'Semaine_{week_number:02d}_Delai',
        'EnAttente': f'Semaine_{week_number:02d}_EnAttente',
    }

    # Supprimer les feuilles existantes si elles existent
    for sheet_name in sheet_names.values():
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]

    print(f"\n📊 CRÉATION DES FEUILLES DE DÉTAIL")
    print(f"   Semaine {week_number:02d} ({start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')})")

    # ========================================================================
    # COMMANDES
    # ========================================================================

    print(f"\n   🔹 Feuille Commandes...")
    dfs_cmd = []
    for sheet in ["Commandes_ALivrer", "Commandes_entournées", "Commandes_Arch"]:
        if sheet in dfs:
            dfs_cmd.append(dfs[sheet].copy())

    if dfs_cmd:
        df_cmd = pd.concat(dfs_cmd, ignore_index=True)
        df_cmd['Date'] = pd.to_datetime(df_cmd['Date'], errors='coerce')

        # Filtrer
        df_cmd_filtered = df_cmd[
            (df_cmd['Date'] >= start_date) &
            (df_cmd['Date'] <= end_date) &
            ((df_cmd['Représentant'].isin(calculator.SALES_REPS_6)) | (df_cmd['Représentant'].isna()))
        ]

        # Créer la feuille
        ws = wb.create_sheet(sheet_names['Commandes'])

        # En-têtes
        headers = ['Date', 'Représentant', 'N°', 'Client', 'Total HT', 'Total TTC']
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.value = header
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")

        # Données
        for row_idx, (_, row) in enumerate(df_cmd_filtered.iterrows(), 2):
            ws.cell(row=row_idx, column=1).value = row.get('Date')
            ws.cell(row=row_idx, column=2).value = row.get('Représentant')
            ws.cell(row=row_idx, column=3).value = row.get('N°')
            ws.cell(row=row_idx, column=4).value = row.get('Client')
            ws.cell(row=row_idx, column=5).value = row.get('Total HT')
            ws.cell(row=row_idx, column=6).value = row.get('Total TTC')

        # Auto-largeur
        for col in ['A', 'B', 'C', 'D', 'E', 'F']:
            ws.column_dimensions[col].width = 15

        print(f"      ✅ {len(df_cmd_filtered)} lignes")

    # ========================================================================
    # FACTURES
    # ========================================================================

    print(f"   🔹 Feuille Factures...")
    dfs_fact = []
    for sheet in ["Fact_Integrer", "Fact_Arch"]:
        if sheet in dfs:
            dfs_fact.append(dfs[sheet].copy())

    if dfs_fact:
        df_fact = pd.concat(dfs_fact, ignore_index=True)
        df_fact['Date'] = pd.to_datetime(df_fact['Date'], errors='coerce')

        # Filtrer
        df_fact_filtered = df_fact[
            (df_fact['Date'] >= start_date) &
            (df_fact['Date'] <= end_date) &
            ((df_fact['Représentant'].isin(calculator.SALES_REPS_6)) | (df_fact['Représentant'].isna()))
        ]

        # Créer la feuille
        ws = wb.create_sheet(sheet_names['Factures'])

        # En-têtes
        headers = ['Date', 'Représentant', 'N°', 'Client', 'Total HT', 'Total TTC']
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.value = header
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")

        # Données
        for row_idx, (_, row) in enumerate(df_fact_filtered.iterrows(), 2):
            ws.cell(row=row_idx, column=1).value = row.get('Date')
            ws.cell(row=row_idx, column=2).value = row.get('Représentant')
            ws.cell(row=row_idx, column=3).value = row.get('N°')
            ws.cell(row=row_idx, column=4).value = row.get('Client')
            ws.cell(row=row_idx, column=5).value = row.get('Total HT')
            ws.cell(row=row_idx, column=6).value = row.get('Total TTC')

        # Auto-largeur
        for col in ['A', 'B', 'C', 'D', 'E', 'F']:
            ws.column_dimensions[col].width = 15

        print(f"      ✅ {len(df_fact_filtered)} lignes")

    # ========================================================================
    # LIVRAISONS
    # ========================================================================

    print(f"   🔹 Feuille Livraisons...")
    dfs_livr = []
    for sheet in ["Livr_AFact", "Livr_Arch"]:
        if sheet in dfs:
            dfs_livr.append(dfs[sheet].copy())

    if dfs_livr:
        df_livr = pd.concat(dfs_livr, ignore_index=True)
        df_livr['Date'] = pd.to_datetime(df_livr['Date'], errors='coerce')
        df_livr['Tournée'] = pd.to_numeric(df_livr['Tournée'], errors='coerce')

        # Filtrer
        df_livr_filtered = df_livr[
            (df_livr['Date'] >= start_date) &
            (df_livr['Date'] <= end_date) &
            (df_livr['Tournée'] > 0) &
            ((df_livr['Représentant'].isin(calculator.SALES_REPS_6)) | (df_livr['Représentant'].isna()))
        ]

        # Créer la feuille
        ws = wb.create_sheet(sheet_names['Livraisons'])

        # En-têtes
        headers = ['Date', 'Représentant', 'Tournée', 'Client', 'Total HT', 'Total TTC']
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.value = header
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")

        # Données
        for row_idx, (_, row) in enumerate(df_livr_filtered.iterrows(), 2):
            ws.cell(row=row_idx, column=1).value = row.get('Date')
            ws.cell(row=row_idx, column=2).value = row.get('Représentant')
            ws.cell(row=row_idx, column=3).value = row.get('Tournée')
            ws.cell(row=row_idx, column=4).value = row.get('Client')
            ws.cell(row=row_idx, column=5).value = row.get('Total HT')
            ws.cell(row=row_idx, column=6).value = row.get('Total TTC')

        # Auto-largeur
        for col in ['A', 'B', 'C', 'D', 'E', 'F']:
            ws.column_dimensions[col].width = 15

        print(f"      ✅ {len(df_livr_filtered)} lignes")

    # ========================================================================
    # CRÉANCES
    # ========================================================================

    print(f"   🔹 Feuille Créances...")
    df_creances = dfs.get("Creances", pd.DataFrame()).copy()

    if not df_creances.empty:
        df_creances['Nb JEch'] = pd.to_numeric(df_creances['Nb JEch'], errors='coerce')

        # Filtrer
        df_creances_filtered = df_creances[
            (df_creances['Nb JEch'] > 0) &
            (df_creances['Nb JEch'].notna())
        ]

        # Créer la feuille
        ws = wb.create_sheet(sheet_names['Créances'])

        # En-têtes
        headers = ['Client', 'N° Facture', 'Montant', 'Nb JEch', 'Restant dû']
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.value = header
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")

        # Données
        for row_idx, (_, row) in enumerate(df_creances_filtered.iterrows(), 2):
            ws.cell(row=row_idx, column=1).value = row.get('Client')
            ws.cell(row=row_idx, column=2).value = row.get('N° Facture')
            ws.cell(row=row_idx, column=3).value = row.get('Montant')
            ws.cell(row=row_idx, column=4).value = row.get('Nb JEch')
            ws.cell(row=row_idx, column=5).value = row.get('Restant dû')

        # Auto-largeur
        for col in ['A', 'B', 'C', 'D', 'E']:
            ws.column_dimensions[col].width = 15

        print(f"      ✅ {len(df_creances_filtered)} lignes")

    # ========================================================================
    # DÉLAI (Expl_Loc_Delais avec tous les filtres du KPI)
    # ========================================================================

    print(f"   🔹 Feuille Délai...")
    df_delai = dfs.get("Expl_Loc_Delais", pd.DataFrame()).copy()

    if not df_delai.empty:
        # Convertir dates
        for col in ['Date', 'Date Creation Cde', 'Liv. souhaitée']:
            if col in df_delai.columns:
                df_delai[col] = pd.to_datetime(df_delai[col], errors='coerce')

        # Convertir Tournée
        df_delai['Tournée'] = pd.to_numeric(df_delai['Tournée'], errors='coerce')

        # Calculer le délai livraison souhaité
        df_delai['délai_livraison_souhaité'] = (
            df_delai['Liv. souhaitée'] - df_delai['Date Creation Cde']
        ).dt.days

        # Calculer le délai cde 2 livr
        df_delai['delai_cde_2_livr'] = (
            (df_delai['Date'] - df_delai['Date Creation Cde']).dt.days -
            (df_delai['délai_livraison_souhaité'] - 1)
        )

        # Appliquer les mêmes filtres que le KPI
        df_delai_filtered = df_delai[
            (df_delai['Date'] >= start_date) &
            (df_delai['Date'] <= end_date) &
            (df_delai['délai_livraison_souhaité'] >= 0) &
            (df_delai['delai_cde_2_livr'] >= 0) &
            (df_delai['delai_cde_2_livr'] <= 14) &
            (df_delai['Tournée'] > 3000)
        ]

        # Créer la feuille
        ws = wb.create_sheet(sheet_names['Delai'])

        # En-têtes
        headers = ['Date', 'Cde N°', 'Cde date', 'Liv. souhaitée', 'Date réelle', 'Tournée', 'Délai souhaité (j)', 'Délai réel (j)']
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.value = header
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")

        # Données
        for row_idx, (_, row) in enumerate(df_delai_filtered.iterrows(), 2):
            ws.cell(row=row_idx, column=1).value = row.get('Date')
            ws.cell(row=row_idx, column=2).value = row.get('Cde N°')
            ws.cell(row=row_idx, column=3).value = row.get('Cde date')
            ws.cell(row=row_idx, column=4).value = row.get('Liv. souhaitée')
            ws.cell(row=row_idx, column=5).value = row.get('Date')
            ws.cell(row=row_idx, column=6).value = row.get('Tournée')
            ws.cell(row=row_idx, column=7).value = row.get('délai_livraison_souhaité')
            ws.cell(row=row_idx, column=8).value = row.get('delai_cde_2_livr')

        # Auto-largeur
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
            ws.column_dimensions[col].width = 15

        print(f"      ✅ {len(df_delai_filtered)} lignes")

    # ========================================================================
    # EN ATTENTE (Commandes_ALivrer avec filtres du KPI)
    # ========================================================================

    print(f"   🔹 Feuille En Attente...")
    df_attente = dfs.get("Commandes_ALivrer", pd.DataFrame()).copy()

    if not df_attente.empty:
        df_attente['Livraison'] = pd.to_datetime(df_attente['Livraison'], errors='coerce')

        # Filtres: Livraison [01/01/2026; end_date] + Reps 6
        date_min = datetime(2026, 1, 1)
        df_attente_filtered = df_attente[
            (df_attente['Livraison'] >= date_min) &
            (df_attente['Livraison'] <= end_date) &
            (df_attente['Représentant'].isin(calculator.SALES_REPS_6))
        ]

        # Créer la feuille
        ws = wb.create_sheet(sheet_names['EnAttente'])

        # En-têtes
        headers = ['Date', 'Représentant', 'N°', 'Client', 'Livraison prévue', 'Total HT', 'Total TTC']
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.value = header
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")

        # Données
        for row_idx, (_, row) in enumerate(df_attente_filtered.iterrows(), 2):
            ws.cell(row=row_idx, column=1).value = row.get('Date')
            ws.cell(row=row_idx, column=2).value = row.get('Représentant')
            ws.cell(row=row_idx, column=3).value = row.get('N°')
            ws.cell(row=row_idx, column=4).value = row.get('Client')
            ws.cell(row=row_idx, column=5).value = row.get('Livraison')
            ws.cell(row=row_idx, column=6).value = row.get('Total HT')
            ws.cell(row=row_idx, column=7).value = row.get('Total TTC')

        # Auto-largeur
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
            ws.column_dimensions[col].width = 15

        print(f"      ✅ {len(df_attente_filtered)} lignes")

    # Sauvegarder
    wb.save(file_output)
    print(f"\n   ✅ Feuilles de détail créées")


def main():
    """Fonction principale"""

    parser = argparse.ArgumentParser(
        description="Calcule les KPIs hebdomadaires et remplit le snapshot"
    )

    parser.add_argument(
        "--week",
        required=True,
        type=int,
        help="Numéro de semaine (1-52)"
    )

    parser.add_argument(
        "--input",
        required=True,
        type=str,
        help="Chemin du fichier d'entrée (Input_Donnees_Brutes_Hebdo.xlsx)"
    )

    parser.add_argument(
        "--snapshot",
        required=True,
        type=str,
        help="Chemin du fichier snapshot (KPIs_weekly.xlsx)"
    )

    parser.add_argument(
        "--year",
        required=False,
        type=int,
        default=2026,
        help="Année (défaut: 2026)"
    )

    args = parser.parse_args()

    print("\n" + "=" * 80)
    print("🚀 PROCESS WEEKLY - CALCUL DES KPIs")
    print("=" * 80)

    # Vérifier que les fichiers existent
    if not Path(args.input).exists():
        print(f"\n❌ Fichier d'entrée '{args.input}' non trouvé")
        sys.exit(1)

    if not Path(args.snapshot).exists():
        print(f"\n❌ Fichier snapshot '{args.snapshot}' non trouvé")
        sys.exit(1)

    # Calculer les dates de la semaine
    try:
        start_date, end_date = get_week_dates(args.week, args.year)
    except ValueError as e:
        print(f"\n❌ Erreur: Semaine invalide ({args.week})")
        sys.exit(1)

    print(f"\n📅 Semaine {args.week:02d} ({args.year})")
    print(f"   Période: {start_date.strftime('%A %d/%m/%Y')} - {end_date.strftime('%A %d/%m/%Y')}")
    print(f"\n📁 Fichiers:")
    print(f"   Input:    {args.input}")
    print(f"   Snapshot: {args.snapshot}")

    # ========================================================================
    # CHARGER LES DONNÉES ET CALCULER LES KPIs
    # ========================================================================

    print(f"\n📊 CALCUL DES KPIs")
    print("-" * 80)

    loader = WeeklyDataLoader(args.input)
    if not loader.load_and_validate():
        print("\n❌ Erreur: Impossible de charger les données")
        sys.exit(1)

    calculator = KPICalculator(loader)
    calculator.set_period(start_date, end_date)

    # Calculer tous les KPIs
    kpis = calculator.get_all_kpis()

    print(f"\n✅ KPIs calculés:")
    print(f"   Commandes: {kpis['commandes']['nb_avec']} / {kpis['commandes']['ca_avec']:,.0f}€ / {kpis['commandes']['ca_sans']:,.0f}€")
    print(f"   Factures: {kpis['factures']['nb_avec']} / {kpis['factures']['ca_avec']:,.0f}€ / {kpis['factures']['nb_sans']} / {kpis['factures']['ca_sans']:,.0f}€")
    print(f"   Livraisons: {kpis['livraisons']['nb_avec']} / {kpis['livraisons']['ca_avec']:,.0f}€ / {kpis['livraisons']['nb_sans']} / {kpis['livraisons']['ca_sans']:,.0f}€")
    print(f"   Délai: {kpis['livraisons']['delai']:.2f} jours")
    print(f"   En attente: {kpis['livraisons']['en_attente_nb']} / {kpis['livraisons']['en_attente_ca']:,.0f}€")
    print(f"   Créances: {kpis['creances']['nb']} / {kpis['creances']['ca']:,.0f}€")

    # ========================================================================
    # CRÉER LES FEUILLES DE DÉTAIL
    # ========================================================================

    create_detail_sheets(args.input, args.snapshot, args.week, calculator, start_date, end_date)

    # ========================================================================
    # REMPLIR LE SNAPSHOT
    # ========================================================================

    print(f"\n📝 REMPLISSAGE DU SNAPSHOT")
    print("-" * 80)

    wb = load_workbook(args.snapshot)
    ws = wb.active

    row = get_row_for_week(args.week)
    print(f"   Remplissage ligne {row} (Semaine {args.week:02d})")

    # Remplir les KPIs (en k€)
    # Mapping des colonnes du snapshot KPI:
    # B: Commandes Nb
    # C: Commandes CA (k€)
    # D: Commandes CA hors Franck
    # E: Factures Nb
    # F: Factures CA (k€)
    # G: Factures CA hors Franck
    # H: Livraisons Nb
    # I: Livraisons CA (k€)
    # J: Livraisons CA hors Franck
    # K: Livraisons Délai
    # L: Livraisons En attente
    # M: Créances Nb
    # N: Créances CA (k€)

    col_mapping = {
        'B': kpis['commandes']['nb_avec'],              # Commandes Nb
        'C': kpis['commandes']['ca_avec'] / 1000,      # Commandes CA (k€)
        'D': kpis['commandes']['ca_sans'] / 1000,      # Commandes CA hors Franck
        'E': kpis['factures']['nb_avec'],              # Factures Nb
        'F': kpis['factures']['ca_avec'] / 1000,       # Factures CA (k€)
        'G': kpis['factures']['ca_sans'] / 1000,       # Factures CA hors Franck
        'H': kpis['livraisons']['nb_avec'],            # Livraisons Nb
        'I': kpis['livraisons']['ca_avec'] / 1000,     # Livraisons CA (k€)
        'J': kpis['livraisons']['ca_sans'] / 1000,     # Livraisons CA hors Franck
        'K': kpis['livraisons']['delai'],              # Délai (jours)
        'L': kpis['livraisons']['en_attente_nb'],      # En attente
        'M': kpis['creances']['nb'],                   # Créances Nb
        'N': kpis['creances']['ca'] / 1000,            # Créances CA (k€)
    }

    for col, value in col_mapping.items():
        cell = ws[f'{col}{row}']
        cell.value = value
        if isinstance(value, float):
            # Délai: 2 decimals, autres: 1 decimal
            cell.number_format = '0.00' if col == 'K' else '0.0'

    wb.save(args.snapshot)
    print(f"   ✅ Snapshot rempli et sauvegardé")

    # ========================================================================
    # SUCCÈS
    # ========================================================================

    print("\n" + "=" * 80)
    print("✅ PROCESS COMPLÉTÉ AVEC SUCCÈS")
    print("=" * 80)
    print(f"\n📄 Fichier snapshot: {args.snapshot}")
    print(f"   - Ligne semaine {args.week:02d} remplie")
    print(f"   - Feuilles de détail créées:")
    print(f"     • Semaine_{args.week:02d}_Commandes")
    print(f"     • Semaine_{args.week:02d}_Factures")
    print(f"     • Semaine_{args.week:02d}_Livraisons")
    print(f"     • Semaine_{args.week:02d}_Créances")
    print(f"     • Semaine_{args.week:02d}_Delai (lignes filtrées)")
    print(f"     • Semaine_{args.week:02d}_EnAttente (lignes filtrées)")
    print()


if __name__ == "__main__":
    main()
