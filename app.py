#!/usr/bin/env python3
"""
Application Streamlit - Générateur de Rapports KPI Hygierun
Interface pour générer les rapports KPI hebdomadaires et le bilan mensuel
"""

import streamlit as st
import pandas as pd
import tempfile
import os
from datetime import datetime, timedelta
from pathlib import Path

from src import WeeklyDataLoader, KPICalculator
from src.excel_helpers import SUMMARY_HEADERS
from src.monthly_loader import MONTHS_FR, bilan_total_required_columns
from src.monthly_report import MONTHLY_SECTIONS
from src.monthly_bilan import build_bilan_report
from src.monthly_bilan_template import generate_bilan_template_excel
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
st.markdown('<div class="subtitle">Automatisez vos rapports KPI</div>', unsafe_allow_html=True)
st.divider()


def run_weekly_report():
    """Onglet Rapport Hebdomadaire (code existant, inchangé fonctionnellement)"""

    # Colonnes pour meilleure présentation
    col1, col2 = st.columns([1, 1], gap="medium")

    with col1:
        st.subheader("📁 Fichiers")
        uploaded_input = st.file_uploader(
            "Fichier de données brutes",
            type="xlsx",
            help="Input_Donnees_Brutes_Hebdo.xlsx",
            key="weekly_input"
        )

    with col2:
        st.subheader("📋 Configuration")
        uploaded_snapshot = st.file_uploader(
            "Fichier snapshot",
            type="xlsx",
            help="KPIs_weekly.xlsx",
            key="weekly_snapshot"
        )
        week_number = st.number_input(
            "Numéro de semaine",
            min_value=1,
            max_value=52,
            value=37,
            help="Semaine ISO (1-52)",
            key="weekly_week_number"
        )

    st.divider()

    # Bouton principal
    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        generate_button = st.button(
            "🚀 Générer Rapport",
            use_container_width=True,
            type="primary",
            key="weekly_generate"
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

                        # Supprimer les feuilles de détail existantes (garder seulement la feuille KPI principale)
                        feuilles_a_garder = [ws.title]  # La feuille active (KPI)
                        for sheet_name in wb.sheetnames:
                            if sheet_name not in feuilles_a_garder:
                                wb.remove(wb[sheet_name])
                                st.info(f"🗑️ Feuille supprimée: {sheet_name}")

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
                                    ws_detail.cell(row=row_idx, column=1).number_format = 'dd/mm/yyyy'  # ← Format date
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
                                    ws_detail.cell(row=row_idx, column=1).number_format = 'dd/mm/yyyy'  # ← Format date
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
                                    ws_detail.cell(row=row_idx, column=1).number_format = 'dd/mm/yyyy'  # ← Format date
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

                            # ===== DELAIS =====
                            df_delais = dfs.get("Expl_Loc_Delais", pd.DataFrame()).copy()
                            if not df_delais.empty:
                                # Conversions
                                for col in ['Date', 'Date Creation Cde', 'Liv. souhaitée']:
                                    if col in df_delais.columns:
                                        df_delais[col] = pd.to_datetime(df_delais[col], errors='coerce')
                                df_delais['Tournée'] = pd.to_numeric(df_delais['Tournée'], errors='coerce')

                                # Calculs
                                df_delais['délai_livraison_souhaité'] = (
                                    df_delais['Liv. souhaitée'] - df_delais['Date Creation Cde']
                                ).dt.days
                                df_delais['delai_cde_2_livr'] = (
                                    (df_delais['Date'] - df_delais['Date Creation Cde']).dt.days -
                                    (df_delais['délai_livraison_souhaité'] - 1)
                                )

                                # Filtres
                                df_filtered = df_delais[
                                    (df_delais['Date'] >= start_date) &
                                    (df_delais['Date'] <= end_date) &
                                    (df_delais['délai_livraison_souhaité'] >= 0) &
                                    (df_delais['delai_cde_2_livr'] >= 0) &
                                    (df_delais['delai_cde_2_livr'] <= 14) &
                                    (df_delais['Tournée'] > 3000)
                                ]

                                ws_detail = wb.create_sheet(f'Semaine_{week_number:02d}_Delais')
                                headers = list(df_filtered.columns)
                                for col_idx, header in enumerate(headers, 1):
                                    cell = ws_detail.cell(row=1, column=col_idx)
                                    cell.value = header
                                    cell.font = Font(bold=True, color="FFFFFF")
                                    cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")

                                for row_idx, (_, row) in enumerate(df_filtered.iterrows(), 2):
                                    for col_idx, (col_name, value) in enumerate(row.items(), 1):
                                        cell = ws_detail.cell(row=row_idx, column=col_idx)
                                        cell.value = value
                                        if isinstance(value, (datetime, pd.Timestamp)):
                                            cell.number_format = 'dd/mm/yyyy'
                                        if col_name == 'Tournée' and isinstance(value, (int, float)):
                                            cell.number_format = '0'

                            # ===== EN ATTENTE =====
                            df_en_attente = dfs.get("Commandes_ALivrer", pd.DataFrame()).copy()
                            if not df_en_attente.empty:
                                df_en_attente['Livraison'] = pd.to_datetime(df_en_attente['Livraison'], errors='coerce')

                                date_min = datetime(2026, 1, 1)
                                df_filtered = df_en_attente[
                                    (df_en_attente['Livraison'] >= date_min) &
                                    (df_en_attente['Livraison'] <= end_date) &
                                    (df_en_attente['Représentant'].isin(calculator.SALES_REPS_6))
                                ]

                                ws_detail = wb.create_sheet(f'Semaine_{week_number:02d}_EnAttente')
                                headers = list(df_filtered.columns)
                                for col_idx, header in enumerate(headers, 1):
                                    cell = ws_detail.cell(row=1, column=col_idx)
                                    cell.value = header
                                    cell.font = Font(bold=True, color="FFFFFF")
                                    cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")

                                for row_idx, (_, row) in enumerate(df_filtered.iterrows(), 2):
                                    for col_idx, (col_name, value) in enumerate(row.items(), 1):
                                        cell = ws_detail.cell(row=row_idx, column=col_idx)
                                        cell.value = value
                                        if isinstance(value, (datetime, pd.Timestamp)):
                                            cell.number_format = 'dd/mm/yyyy'

                        # Créer feuilles et remplir snapshot
                        create_detail_sheets(wb, loader, calculator, week_number, start_date, end_date)

                        # Calculer ligne pour écrire (colonne A = numéro semaine, colonnes B-N = KPIs)
                        # Trouver la première ligne vide dans le snapshot
                        row_to_fill = None
                        for row in range(3, 201):
                            if ws[f'A{row}'].value is None or str(ws[f'A{row}'].value).strip() == '':
                                row_to_fill = row
                                break

                        if row_to_fill is None:
                            st.error("❌ Snapshot plein (toutes les lignes de 3 à 100 sont remplies)")
                            st.stop()

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


def _fmt_number(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, float):
        return f"{value:,.1f}".replace(",", " ").replace(".", ",")
    return f"{value:,}".replace(",", " ")


def _fmt_evolution(value):
    if value is None or pd.isna(value):
        return ""
    return f"{value:+.1f} %".replace(".", ",")


def _summary_dataframe(rows):
    records = []
    for row in rows:
        if row is None:
            continue
        if len(row) == 1:
            records.append({SUMMARY_HEADERS[0]: f"— {row[0]} —", **{h: "" for h in SUMMARY_HEADERS[1:]}})
            continue
        label, value, n1, evo_n1, m1, evo_m1 = row
        records.append({
            SUMMARY_HEADERS[0]: label, SUMMARY_HEADERS[1]: _fmt_number(value), SUMMARY_HEADERS[2]: _fmt_number(n1),
            SUMMARY_HEADERS[3]: _fmt_evolution(evo_n1), SUMMARY_HEADERS[4]: _fmt_number(m1), SUMMARY_HEADERS[5]: _fmt_evolution(evo_m1),
        })
    return pd.DataFrame(records)


def _bilan_total_section(year: int, month: int):
    """Onglet 'Bilan total' : 1 seul fichier Input (19 feuilles fusionnées) -> 1 zip (1 PowerPoint de
    11 diapos + 4 Excel regroupés : Commerce / Préparation+Livraison / Compta+Achats / SAV). Mois-1
    optionnel (deck Bilan Mensuel du mois précédent)."""
    state_key = "bilan_total"
    st.caption("Génère le Bilan Mensuel complet (Commerce, Préparation, Livraison, Compta, SAV, Achat) "
              "à partir d'un seul fichier Input consolidé.")
    st.download_button(
        f"📄 Télécharger le template Input vierge ({MONTHS_FR[month - 1]} {year})",
        data=generate_bilan_template_excel(year, month),
        file_name=f"Input_Mensuel_Template_{MONTHS_FR[month - 1]}{year}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True, key=f"{state_key}_dl_template",
        help="19 feuilles avec les colonnes obligatoires déjà en en-tête (adaptées au mois/année "
             "sélectionnés ci-dessus) — à remplir en collant les exports ERP correspondants.",
    )
    uploaded = st.file_uploader("Input Mensuel (toutes sections)", type="xlsx", key=f"{state_key}_input",
                                help="19 feuilles : les feuilles partagées entre sections (Cdes_Arch, "
                                     "Fact_Arch, Livr_Arch) ne doivent apparaître qu'une seule fois, "
                                     "avec toutes les colonnes nécessaires à chaque section.")
    mois1_pptx = st.file_uploader("PowerPoint Bilan Mensuel du mois précédent (optionnel)", type="pptx",
                                  key=f"{state_key}_mois1_pptx",
                                  help="Pour les évolutions Mois-1 non recalculables depuis l'historique "
                                       "(clients bloqués, factures dues, GPS, nombre d'interventions, "
                                       "productivité, valorisation du stock, articles à épuisement). "
                                       "Optionnel — sans lui, ces évolutions restent 'n/a'.")

    with st.expander("📋 Colonnes obligatoires dans le fichier input", expanded=False):
        st.caption("Ces colonnes doivent exister (avec ces noms exacts) pour que les calculs fonctionnent.")
        for sheet, cols in bilan_total_required_columns(year, month).items():
            st.markdown(f"**{sheet}** : " + ", ".join(f"`{c}`" for c in cols))

    if st.button("🚀 Générer le Bilan Total", use_container_width=True, type="primary", key=f"{state_key}_generate"):
        if not uploaded:
            st.error("❌ Veuillez charger le fichier Input Mensuel consolidé", icon="📋")
        else:
            try:
                with st.spinner("⏳ Traitement en cours..."):
                    with tempfile.TemporaryDirectory() as tmpdir:
                        input_path = Path(tmpdir) / "input_bilan.xlsx"
                        input_path.write_bytes(uploaded.getbuffer())
                        mois1_path = None
                        if mois1_pptx is not None:
                            mois1_path = Path(tmpdir) / "mois1_bilan.pptx"
                            mois1_path.write_bytes(mois1_pptx.getbuffer())
                        st.session_state[state_key] = build_bilan_report(
                            str(input_path), year, month,
                            mois1_pptx_path=str(mois1_path) if mois1_path else None,
                        )
            except Exception as e:
                st.session_state.pop(state_key, None)
                st.error(f"❌ Erreur: {e}")

    report = st.session_state.get(state_key)
    if report:
        suffix = f"{MONTHS_FR[month - 1]}{year}"
        st.success("✅ Bilan Mensuel généré (11 diapos + 4 Excel)")
        st.download_button(
            "📥 Télécharger le zip complet (PowerPoint + 4 Excel)", data=report["zip"],
            file_name=f"Bilan_Mensuel_{suffix}.zip", mime="application/zip",
            use_container_width=True, type="primary",
        )
        with st.expander("Télécharger les fichiers séparément", expanded=False):
            st.download_button(
                "📥 PowerPoint Bilan Mensuel", data=report["pptx"],
                file_name=f"Bilan_Mensuel_{suffix}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True, key=f"{state_key}_dl_pptx",
            )
            excel_labels = {
                "commerce": ("Commerce", f"Commerce_{suffix}.xlsx"),
                "preparation_livraison": ("Préparation + Livraison", f"Preparation_Livraison_{suffix}.xlsx"),
                "compta_achats": ("Compta + Achats", f"Compta_Achats_{suffix}.xlsx"),
                "sav": ("SAV", f"SAV_{suffix}.xlsx"),
            }
            for key, (label, file_name) in excel_labels.items():
                st.download_button(
                    f"📥 Excel {label}", data=report["excels"][key], file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True, key=f"{state_key}_dl_{key}",
                )

        for key, summary in report["summaries"].items():
            with st.expander(f"Résumé — {key.capitalize()}", expanded=False):
                st.dataframe(_summary_dataframe(summary), use_container_width=True, hide_index=True)


def _monthly_section(key: str, section: dict, year: int, month: int):
    """Une section du Bilan Mensuel : upload de son input, génération, aperçu et téléchargements."""
    state_key = f"monthly_{key}"
    uploaded = st.file_uploader(f"Input mensuel {section['name']}", type="xlsx", help=section["input_help"],
                                key=f"{state_key}_input")

    mois1_pptx = None
    if "mois1_pptx_help" in section:
        mois1_pptx = st.file_uploader("PowerPoint Mois-1 (optionnel)", type="pptx", help=section["mois1_pptx_help"],
                                      key=f"{state_key}_mois1_pptx")

    with st.expander("📋 Colonnes obligatoires dans le fichier input", expanded=False):
        st.caption("Ces colonnes doivent exister (avec ces noms exacts) pour que les calculs fonctionnent.")
        for sheet, cols in section["required"](year, month).items():
            st.markdown(f"**{sheet}** : " + ", ".join(f"`{c}`" for c in cols))

    if st.button(f"🚀 Générer {section['name']}", use_container_width=True, type="primary", key=f"{state_key}_generate"):
        if not uploaded:
            st.error(f"❌ Veuillez charger le fichier Input mensuel {section['name']}", icon="📋")
        else:
            try:
                with st.spinner("⏳ Traitement en cours..."):
                    with tempfile.TemporaryDirectory() as tmpdir:
                        input_path = Path(tmpdir) / f"input_{key}.xlsx"
                        input_path.write_bytes(uploaded.getbuffer())
                        build_kwargs = {}
                        if mois1_pptx is not None:
                            mois1_path = Path(tmpdir) / f"mois1_{key}.pptx"
                            mois1_path.write_bytes(mois1_pptx.getbuffer())
                            build_kwargs["mois1_pptx_path"] = str(mois1_path)
                        st.session_state[state_key] = section["build"](str(input_path), year, month, **build_kwargs)
            except Exception as e:
                st.session_state.pop(state_key, None)
                st.error(f"❌ Erreur: {e}")

    report = st.session_state.get(state_key)
    if report:
        result = report["result"]
        period = f"{MONTHS_FR[result['month'] - 1]} {result['year']}"
        st.success(f"✅ Section {section['name']} générée pour {period}")

        st.dataframe(_summary_dataframe(report["summary"]), hide_index=True, use_container_width=True, height=min(600, 40 + 35 * len(report["summary"])))

        col_x, col_p = st.columns(2)
        with col_x:
            st.download_button(
                f"📥 Télécharger l'Excel {section['name']}", data=report["excel"],
                file_name=f"{key.capitalize()}_{result['year']}-{result['month']:02d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, key=f"{state_key}_dl_excel",
            )
        with col_p:
            st.download_button(
                f"📥 Télécharger les diapos {section['name']}", data=report["pptx"],
                file_name=f"KPI_{key.capitalize()}_{MONTHS_FR[result['month'] - 1]}{result['year']}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True, key=f"{state_key}_dl_pptx",
            )


def run_monthly_report():
    """Onglet Bilan Mensuel : une sous-section par domaine (les autres arrivent une par une)"""

    today = datetime.now()
    default_year, default_month = (today.year - 1, 12) if today.month == 1 else (today.year, today.month - 1)

    col_month, col_year = st.columns(2, gap="medium")
    with col_month:
        month = st.selectbox("Mois", list(range(1, 13)), index=default_month - 1,
                             format_func=lambda m: MONTHS_FR[m - 1], key="monthly_month")
    with col_year:
        year = int(st.number_input("Année", min_value=2020, max_value=2100, value=default_year, step=1, key="monthly_year"))

    labels = ["🗂️ Bilan total"] + [section["label"] for section in MONTHLY_SECTIONS.values()]
    all_tabs = st.tabs(labels)
    with all_tabs[0]:
        _bilan_total_section(year, month)
    for tab, (key, section) in zip(all_tabs[1:], MONTHLY_SECTIONS.items()):
        with tab:
            _monthly_section(key, section, year, month)


tab_weekly, tab_monthly = st.tabs(["📊 Rapport Hebdomadaire", "📈 Bilan Mensuel"])

with tab_weekly:
    run_weekly_report()

with tab_monthly:
    run_monthly_report()
