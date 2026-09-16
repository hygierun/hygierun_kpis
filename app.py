#!/usr/bin/env python3
"""
Application Streamlit - Générateur de Rapports KPI Hygierun
Interface simple pour générer les rapports KPI hebdomadaires
"""

import streamlit as st
import pandas as pd
import tempfile
import os
from datetime import datetime, timedelta
from pathlib import Path

from src import WeeklyDataLoader, KPICalculator
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

# Configuration Streamlit
st.set_page_config(
    page_title="KPI Hygierun",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Styles
st.markdown("""
    <style>
    .title-main {
        text-align: center;
        color: #366092;
        font-size: 2.5em;
        font-weight: bold;
    }
    .subtitle {
        text-align: center;
        color: #666;
        font-size: 1.1em;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="title-main">📊 Générateur KPI Hygierun</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Automatisez vos rapports hebdomadaires</div>', unsafe_allow_html=True)
st.divider()

# Colonnes pour meilleure présentation
col1, col2 = st.columns([1, 1], gap="medium")

with col1:
    st.subheader("📁 Fichiers")
    uploaded_input = st.file_uploader(
        "Fichier de données brutes",
        type="xlsx",
        help="Input_Donnees_Brutes_Hebdo.xlsx"
    )

with col2:
    st.subheader("📋 Configuration")
    uploaded_snapshot = st.file_uploader(
        "Fichier snapshot",
        type="xlsx",
        help="KPIs_weekly.xlsx"
    )
    week_number = st.number_input(
        "Numéro de semaine",
        min_value=1,
        max_value=52,
        value=37,
        help="Semaine ISO (1-52)"
    )

st.divider()

# Bouton principal
col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
with col_btn2:
    generate_button = st.button(
        "🚀 Générer Rapport",
        use_container_width=True,
        type="primary"
    )

# Processing
if generate_button:
    if not uploaded_input or not uploaded_snapshot:
        st.error("❌ Veuillez charger les 2 fichiers Excel", icon="📋")
    else:
        try:
            with st.spinner("⏳ Traitement en cours..."):
                # Créer des fichiers temporaires
                with tempfile.TemporaryDirectory() as tmpdir:
                    # Sauvegarder les uploads
                    input_path = Path(tmpdir) / "input.xlsx"
                    snapshot_path = Path(tmpdir) / "snapshot.xlsx"
                    output_path = Path(tmpdir) / "output.xlsx"

                    with open(input_path, "wb") as f:
                        f.write(uploaded_input.getbuffer())

                    with open(snapshot_path, "wb") as f:
                        f.write(uploaded_snapshot.getbuffer())

                    # Calculer dates
                    year = 2026
                    jan_4 = datetime(year, 1, 4)
                    week_1_monday = jan_4 - timedelta(days=jan_4.weekday())
                    start_date = week_1_monday + timedelta(weeks=week_number - 1)
                    end_date = start_date + timedelta(days=6)

                    # Charger et valider les données
                    st.info(f"📅 Traitement semaine {week_number:02d} ({start_date.strftime('%d/%m')} - {end_date.strftime('%d/%m')})")

                    loader = WeeklyDataLoader(str(input_path))
                    if not loader.load_and_validate():
                        st.error("❌ Erreur lors du chargement des données")
                        st.stop()

                    # Calculer KPIs
                    calculator = KPICalculator(loader)
                    calculator.set_period(start_date, end_date)
                    kpis = calculator.get_all_kpis()

                    # Diagnostic du délai
                    delai_value = kpis['livraisons']['delai']
                    if delai_value == 0:
                        st.warning(f"⚠️ Délai cde→livraison: Pas de données disponibles pour cette période")
                        st.info(f"💡 Vérifiez que la feuille 'Expl_Loc_Delais' contient des données pour S{week_number:02d}")

                    # Afficher les KPIs calculés
                    with st.expander("📊 KPIs calculés", expanded=True):
                        kpi_cols = st.columns(4)

                        with kpi_cols[0]:
                            st.metric(
                                "Commandes",
                                f"{int(kpis['commandes']['nb_avec'])}",
                                f"{kpis['commandes']['ca_avec']/1000:.1f}k€"
                            )

                        with kpi_cols[1]:
                            st.metric(
                                "Factures",
                                f"{int(kpis['factures']['nb_avec'])}",
                                f"{kpis['factures']['ca_avec']/1000:.1f}k€"
                            )

                        with kpi_cols[2]:
                            st.metric(
                                "Livraisons",
                                f"{int(kpis['livraisons']['nb_avec'])}",
                                f"{kpis['livraisons']['ca_avec']/1000:.1f}k€"
                            )

                        with kpi_cols[3]:
                            st.metric(
                                "Créances",
                                f"{int(kpis['creances']['nb'])}",
                                f"{kpis['creances']['ca']/1000:.1f}k€"
                            )

                    # Créer les feuilles de détail et remplir snapshot
                    # IMPORTANT: Charger avec data_only=False pour garder les formules et les données
                    wb = load_workbook(str(snapshot_path))
                    ws = wb.active

                    # Snapshot backup: Lire les données existantes AVANT modification
                    existing_data = {}
                    for row in ws.iter_rows(min_row=3, max_row=100, min_col=1, max_col=14):
                        row_num = row[0].row
                        if row[0].value and str(row[0].value).startswith('S'):
                            existing_data[row_num] = [cell.value for cell in row]

                    st.info(f"📌 Snapshot input chargé - {len(existing_data)} semaines existantes détectées")

                    # Afficher les semaines détectées dans la sauvegarde
                    if existing_data:
                        detected_weeks = [row[0] for row in existing_data.values() if row[0]]
                        st.info(f"📊 Semaines détectées: {', '.join(str(w) for w in detected_weeks[:10])}")

                    # Fonction pour créer feuilles de détail
                    def create_detail_sheets(wb, loader, calculator, week_num, start, end):
                        dfs = loader.dfs

                        # COMMANDES
                        dfs_cmd = [dfs[sheet].copy() for sheet in ["Commandes_ALivrer", "Commandes_entournées", "Commandes_Arch"] if sheet in dfs]
                        if dfs_cmd:
                            df = pd.concat(dfs_cmd, ignore_index=True)
                            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
                            df_filtered = df[(df['Date'] >= start) & (df['Date'] <= end) & ((df['Représentant'].isin(calculator.SALES_REPS_6)) | (df['Représentant'].isna()))]
                            ws_detail = wb.create_sheet(f'Semaine_{week_num:02d}_Commandes')
                            headers = ['Date', 'Représentant', 'N°', 'Client', 'Total HT', 'Total TTC']
                            for col_idx, header in enumerate(headers, 1):
                                cell = ws_detail.cell(row=1, column=col_idx)
                                cell.value = header
                                cell.font = Font(bold=True, color="FFFFFF")
                                cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                            for row_idx, (_, row) in enumerate(df_filtered.iterrows(), 2):
                                ws_detail.cell(row=row_idx, column=1).value = row.get('Date')
                                ws_detail.cell(row=row_idx, column=2).value = row.get('Représentant')
                                ws_detail.cell(row=row_idx, column=3).value = row.get('N°')
                                ws_detail.cell(row=row_idx, column=4).value = row.get('Client')
                                ws_detail.cell(row=row_idx, column=5).value = row.get('Total HT')
                                ws_detail.cell(row=row_idx, column=6).value = row.get('Total TTC')
                            for col in ['A', 'B', 'C', 'D', 'E', 'F']:
                                ws_detail.column_dimensions[col].width = 15

                        # FACTURES
                        dfs_fact = [dfs[sheet].copy() for sheet in ["Fact_Integrer", "Fact_Arch"] if sheet in dfs]
                        if dfs_fact:
                            df = pd.concat(dfs_fact, ignore_index=True)
                            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
                            df_filtered = df[(df['Date'] >= start) & (df['Date'] <= end) & ((df['Représentant'].isin(calculator.SALES_REPS_6)) | (df['Représentant'].isna()))]
                            ws_detail = wb.create_sheet(f'Semaine_{week_num:02d}_Factures')
                            headers = ['Date', 'Représentant', 'N°', 'Client', 'Total HT', 'Total TTC']
                            for col_idx, header in enumerate(headers, 1):
                                cell = ws_detail.cell(row=1, column=col_idx)
                                cell.value = header
                                cell.font = Font(bold=True, color="FFFFFF")
                                cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                            for row_idx, (_, row) in enumerate(df_filtered.iterrows(), 2):
                                ws_detail.cell(row=row_idx, column=1).value = row.get('Date')
                                ws_detail.cell(row=row_idx, column=2).value = row.get('Représentant')
                                ws_detail.cell(row=row_idx, column=3).value = row.get('N°')
                                ws_detail.cell(row=row_idx, column=4).value = row.get('Client')
                                ws_detail.cell(row=row_idx, column=5).value = row.get('Total HT')
                                ws_detail.cell(row=row_idx, column=6).value = row.get('Total TTC')
                            for col in ['A', 'B', 'C', 'D', 'E', 'F']:
                                ws_detail.column_dimensions[col].width = 15

                        # LIVRAISONS
                        dfs_livr = [dfs[sheet].copy() for sheet in ["Livr_AFact", "Livr_Arch"] if sheet in dfs]
                        if dfs_livr:
                            df = pd.concat(dfs_livr, ignore_index=True)
                            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
                            df['Tournée'] = pd.to_numeric(df['Tournée'], errors='coerce')
                            df_filtered = df[(df['Date'] >= start) & (df['Date'] <= end) & (df['Tournée'] > 0) & ((df['Représentant'].isin(calculator.SALES_REPS_6)) | (df['Représentant'].isna()))]
                            ws_detail = wb.create_sheet(f'Semaine_{week_num:02d}_Livraisons')
                            headers = ['Date', 'Représentant', 'Tournée', 'Client', 'Total HT', 'Total TTC']
                            for col_idx, header in enumerate(headers, 1):
                                cell = ws_detail.cell(row=1, column=col_idx)
                                cell.value = header
                                cell.font = Font(bold=True, color="FFFFFF")
                                cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                            for row_idx, (_, row) in enumerate(df_filtered.iterrows(), 2):
                                ws_detail.cell(row=row_idx, column=1).value = row.get('Date')
                                ws_detail.cell(row=row_idx, column=2).value = row.get('Représentant')
                                ws_detail.cell(row=row_idx, column=3).value = row.get('Tournée')
                                ws_detail.cell(row=row_idx, column=4).value = row.get('Client')
                                ws_detail.cell(row=row_idx, column=5).value = row.get('Total HT')
                                ws_detail.cell(row=row_idx, column=6).value = row.get('Total TTC')
                            for col in ['A', 'B', 'C', 'D', 'E', 'F']:
                                ws_detail.column_dimensions[col].width = 15

                        # CRÉANCES
                        df_creances = dfs.get("Creances", pd.DataFrame()).copy()
                        if not df_creances.empty:
                            df_creances['Nb JEch'] = pd.to_numeric(df_creances['Nb JEch'], errors='coerce')
                            df_filtered = df_creances[(df_creances['Nb JEch'] > 0) & (df_creances['Nb JEch'].notna())]
                            ws_detail = wb.create_sheet(f'Semaine_{week_num:02d}_Créances')
                            headers = ['Client', 'N° Facture', 'Montant', 'Nb JEch', 'Restant dû']
                            for col_idx, header in enumerate(headers, 1):
                                cell = ws_detail.cell(row=1, column=col_idx)
                                cell.value = header
                                cell.font = Font(bold=True, color="FFFFFF")
                                cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                            for row_idx, (_, row) in enumerate(df_filtered.iterrows(), 2):
                                ws_detail.cell(row=row_idx, column=1).value = row.get('Client')
                                ws_detail.cell(row=row_idx, column=2).value = row.get('N° Facture')
                                ws_detail.cell(row=row_idx, column=3).value = row.get('Montant')
                                ws_detail.cell(row=row_idx, column=4).value = row.get('Nb JEch')
                                ws_detail.cell(row=row_idx, column=5).value = row.get('Restant dû')
                            for col in ['A', 'B', 'C', 'D', 'E']:
                                ws_detail.column_dimensions[col].width = 15

                    # Créer feuilles et remplir snapshot
                    create_detail_sheets(wb, loader, calculator, week_number, start_date, end_date)
                    # Créer feuille "délais" - détails des délais calculés
                    df_delais = loader.dfs.get("Expl_Loc_Delais", pd.DataFrame()).copy()
                    if not df_delais.empty:
                        ws_delais = wb.create_sheet("délais")
                        for r_idx, row in enumerate(dataframe_to_rows(df_delais, index=False, header=True), 1):
                            for c_idx, value in enumerate(row, 1):
                                ws_delais.cell(row=r_idx, column=c_idx, value=value)
                        st.info(f"✅ Feuille 'délais': {len(df_delais)} lignes")
                    
                    # Créer feuille "EnAttente" - commandes en attente de livraison
                    df_en_attente = loader.dfs.get("Commandes_ALivrer", pd.DataFrame()).copy()
                    if not df_en_attente.empty:
                        ws_en_attente = wb.create_sheet("EnAttente")
                        for r_idx, row in enumerate(dataframe_to_rows(df_en_attente, index=False, header=True), 1):
                            for c_idx, value in enumerate(row, 1):
                                ws_en_attente.cell(row=r_idx, column=c_idx, value=value)
                        st.info(f"✅ Feuille 'EnAttente': {len(df_en_attente)} lignes")
                    # Calculer ligne pour écrire (colonne A = numéro semaine, colonnes B-N = KPIs)
                    # Trouver la première ligne vide dans le snapshot
                    row_to_fill = None
                    for row in range(2, 201):
                        if ws[f'A{row}'].value is None or str(ws[f'A{row}'].value).strip() == '':
                            row_to_fill = row
                            break
                    
                    if row_to_fill is None:
                        st.error("❌ Snapshot plein (toutes les lignes de 3 à 100 sont remplies)")
                        return

                    st.write(f"📝 Modification ligne {row_to_fill} pour semaine {week_number}")

                    # Diagnostic: Vérifier quelle semaine est actuellement à cette ligne
                    current_week_at_row = ws[f'A{row_to_fill}'].value
                    st.info(f"📋 Ligne {row_to_fill} contient actuellement: {current_week_at_row}")

                    # Remplir colonne A avec le numéro de semaine
                    ws[f'A{row_to_fill}'].value = f'S{week_number:02d}'

                    # Remplir snapshot - SEULEMENT la ligne demandée
                    col_mapping = {
                        'B': kpis['commandes']['nb_avec'],
                        'C': kpis['commandes']['ca_avec'] / 1000,
                        'D': kpis['commandes']['ca_sans'] / 1000,
                        'E': kpis['factures']['nb_avec'],
                        'F': kpis['factures']['ca_avec'] / 1000,
                        'G': kpis['factures']['ca_sans'] / 1000,
                        'H': kpis['livraisons']['nb_avec'],
                        'I': kpis['livraisons']['ca_avec'] / 1000,
                        'J': kpis['livraisons']['ca_sans'] / 1000,
                        'K': kpis['livraisons']['delai'],
                        'L': kpis['livraisons']['en_attente_nb'],
                        'M': kpis['creances']['nb'],
                        'N': kpis['creances']['ca'] / 1000,
                    }

                    # IMPORTANT: Ne modifier QUE la ligne demandée
                    for col, value in col_mapping.items():
                        cell = ws[f'{col}{row_to_fill}']
                        cell.value = value
                        if isinstance(value, float):
                            cell.number_format = '0.00' if col == 'K' else '0.0'

                    st.success(f"✅ Ligne {row_to_fill} remplie - Autres semaines préservées")

                    wb.save(str(output_path))

                    # VERIFICATION: Vérifier que les données ont bien été préservées
                    wb_verify = load_workbook(str(output_path))
                    ws_verify = wb_verify.active

                    preserved_count = 0
                    corrupted_rows = []
                    for row_num, original_data in existing_data.items():
                        if row_num != row_to_fill:  # Ne pas vérifier la ligne qu'on vient de modifier
                            current_data = [ws_verify.cell(row=row_num, column=col).value for col in range(1, 15)]
                            # Vérifier que au moins la colonne A (semaine) est intacte
                            if current_data[0] == original_data[0]:
                                preserved_count += 1
                            else:
                                corrupted_rows.append(row_num)

                    if corrupted_rows:
                        st.warning(f"⚠️ Attention: Lignes potentiellement corrompues détectées: {corrupted_rows}")
                    else:
                        st.success(f"✅ Vérification: {preserved_count} semaines intactes")

                    # Succès
                    st.success(f"✅ Rapport généré avec succès!")

                    # Afficher détails
                    with st.expander("📋 Détails du traitement", expanded=False):
                        st.write(f"""
                        - ✅ Semaine {week_number:02d} remplie (ligne {row_to_fill})
                        - ✅ 4 feuilles de détail créées
                        - ✅ {len(loader.dfs)} feuilles sources chargées
                        """)

                    # Téléchargement
                    with open(str(output_path), "rb") as f:
                        st.download_button(
                            label="📥 Télécharger rapport",
                            data=f.read(),
                            file_name=f"KPIs_semaine_{week_number:02d}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

        except Exception as e:
            st.error(f"❌ Erreur: {str(e)}")
            st.write("Vérifiez que les fichiers sont au bon format")

st.divider()
st.caption("💡 Astuce: Vous pouvez télécharger les fichiers d'exemple pour tester l'application")
