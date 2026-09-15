#!/usr/bin/env python3
"""
Script de test pour valider les calculs KPI

Test contre les valeurs connues de la semaine 37 (7-11 septembre 2026)
"""

import sys
from datetime import datetime
from pathlib import Path

from src.data_loader import DataLoader
from src.kpi_calculator import KPICalculator


def test_week_37():
    """
    Test avec les données réelles de semaine 37
    (Sept 7-11, 2026)
    """

    print("\n" + "=" * 80)
    print("🧪 TEST KPI CALCULATOR - SEMAINE 37 (7-11 Septembre 2026)")
    print("=" * 80)

    # Vérifier que le fichier d'entrée existe
    input_file = "Input_Donnees_Brutes_Hebdo.xlsx"
    if not Path(input_file).exists():
        print(f"\n❌ Fichier '{input_file}' non trouvé")
        print("   Veuillez fournir le fichier Excel de données")
        return False

    print(f"\n📁 Chargement des données: {input_file}")

    # Charger les données
    loader = DataLoader(input_file)
    if not loader.load_and_validate():
        print("\n❌ Erreur lors du chargement des données")
        return False

    print("✅ Données chargées avec succès")

    # Créer le calculateur
    calculator = KPICalculator(loader)

    # Définir la période: semaine 37 (7-11 septembre 2026)
    start_date = datetime(2026, 9, 7)
    end_date = datetime(2026, 9, 11)
    calculator.set_period(start_date, end_date)

    print(f"\n📅 Période: {start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}")

    # ========================================================================
    # RÉSULTATS ATTENDUS (VALIDÉS PAR ANTOINE)
    # ========================================================================

    expected = {
        "nb_commandes": 63,
        "ca_commandes_avec": 21_771.06,
        "ca_commandes_sans": 18_178.13,
        "nb_factures_avec": 248,
        "nb_factures_sans": 212,
        "ca_factures_avec": 101_540.73,
        "ca_factures_sans": 60_325.12,
        "nb_livraisons_avec": 183,
        "nb_livraisons_sans": 171,
        "ca_livraisons_avec": 61_746.79,
        "ca_livraisons_sans": 54_106.71,
        "livraisons_en_attente": 30,
        "delai_livraison": 3.66,
        "creances_count": 417,
        "creances_montant": 238_423.62,
    }

    # ========================================================================
    # CALCULS
    # ========================================================================

    print("\n🔢 Calcul des KPIs...")
    print("-" * 80)

    # Commerce
    print("\n📊 SECTION COMMERCE")
    print("-" * 80)

    nb_cmd = calculator.calc_commandes_clients()
    ca_cmd_avec = calculator.calc_ca_facture_avec_franck()  # Approximation
    ca_cmd_sans = calculator.calc_ca_facture_sans_franck()  # Approximation
    nb_fact_avec = calculator.calc_nb_factures_avec_franck()
    nb_fact_sans = calculator.calc_nb_factures_sans_franck()
    ca_fact_avec = calculator.calc_ca_facture_avec_franck()
    ca_fact_sans = calculator.calc_ca_facture_sans_franck()

    print(f"Commandes clients: {nb_cmd} (attendu: {expected['nb_commandes']})")
    print(f"Factures (avec Franck): {nb_fact_avec} (attendu: {expected['nb_factures_avec']})")
    print(f"Factures (sans Franck): {nb_fact_sans} (attendu: {expected['nb_factures_sans']})")
    print(f"CA Factures (avec Franck): {ca_fact_avec:,.2f}€ (attendu: {expected['ca_factures_avec']:,.2f}€)")
    print(f"CA Factures (sans Franck): {ca_fact_sans:,.2f}€ (attendu: {expected['ca_factures_sans']:,.2f}€)")

    # Livraison
    print("\n📊 SECTION LIVRAISON")
    print("-" * 80)

    ca_liv_avec = calculator.calc_ca_livre_avec_franck()
    ca_liv_sans = calculator.calc_ca_livre_sans_franck()
    nb_liv_avec = calculator.calc_nb_livraisons_avec_franck()
    nb_liv_sans = calculator.calc_nb_livraisons_sans_franck()
    delai = calculator.calc_delai_livraison()
    liv_attente = calculator.calc_livraisons_en_attente()

    print(f"Livraisons (avec Franck): {nb_liv_avec} (attendu: {expected['nb_livraisons_avec']})")
    print(f"Livraisons (sans Franck): {nb_liv_sans} (attendu: {expected['nb_livraisons_sans']})")
    print(f"CA Livraisons (avec Franck): {ca_liv_avec:,.2f}€ (attendu: {expected['ca_livraisons_avec']:,.2f}€)")
    print(f"CA Livraisons (sans Franck): {ca_liv_sans:,.2f}€ (attendu: {expected['ca_livraisons_sans']:,.2f}€)")
    print(f"Délai livraison: {delai:.2f} jours (attendu: {expected['delai_livraison']} jours)")
    print(f"Livraisons en attente: {liv_attente} (attendu: {expected['livraisons_en_attente']})")

    # Créances
    print("\n📊 SECTION CRÉANCES")
    print("-" * 80)

    creances_count = calculator.calc_creances_echues_count()
    creances_montant = calculator.calc_creances_echues_montant()

    print(f"Créances échues (nb): {creances_count} (attendu: {expected['creances_count']})")
    print(f"Créances échues (montant): {creances_montant:,.2f}€ (attendu: {expected['creances_montant']:,.2f}€)")

    # ========================================================================
    # RAPPORT DE TEST
    # ========================================================================

    print("\n" + "=" * 80)
    print("✅ TEST COMPLÉTÉ")
    print("=" * 80)

    # Vérifier les écarts
    errors = []

    if abs(nb_fact_avec - expected["nb_factures_avec"]) > 0:
        errors.append(f"Factures avec Franck: {nb_fact_avec} vs {expected['nb_factures_avec']}")

    if abs(nb_fact_sans - expected["nb_factures_sans"]) > 0:
        errors.append(f"Factures sans Franck: {nb_fact_sans} vs {expected['nb_factures_sans']}")

    if abs(ca_fact_avec - expected["ca_factures_avec"]) > 1:
        errors.append(f"CA Factures avec: {ca_fact_avec:,.2f}€ vs {expected['ca_factures_avec']:,.2f}€")

    if abs(ca_fact_sans - expected["ca_factures_sans"]) > 1:
        errors.append(f"CA Factures sans: {ca_fact_sans:,.2f}€ vs {expected['ca_factures_sans']:,.2f}€")

    if abs(nb_liv_avec - expected["nb_livraisons_avec"]) > 0:
        errors.append(f"Livraisons avec: {nb_liv_avec} vs {expected['nb_livraisons_avec']}")

    if abs(nb_liv_sans - expected["nb_livraisons_sans"]) > 0:
        errors.append(f"Livraisons sans: {nb_liv_sans} vs {expected['nb_livraisons_sans']}")

    if abs(ca_liv_avec - expected["ca_livraisons_avec"]) > 1:
        errors.append(f"CA Livraisons avec: {ca_liv_avec:,.2f}€ vs {expected['ca_livraisons_avec']:,.2f}€")

    if abs(ca_liv_sans - expected["ca_livraisons_sans"]) > 1:
        errors.append(f"CA Livraisons sans: {ca_liv_sans:,.2f}€ vs {expected['ca_livraisons_sans']:,.2f}€")

    if abs(delai - expected["delai_livraison"]) > 0.1:
        errors.append(f"Délai livraison: {delai:.2f} vs {expected['delai_livraison']}")

    if liv_attente != expected["livraisons_en_attente"]:
        errors.append(f"Livraisons en attente: {liv_attente} vs {expected['livraisons_en_attente']}")

    if creances_count != expected["creances_count"]:
        errors.append(f"Créances count: {creances_count} vs {expected['creances_count']}")

    if abs(creances_montant - expected["creances_montant"]) > 1:
        errors.append(f"Créances montant: {creances_montant:,.2f}€ vs {expected['creances_montant']:,.2f}€")

    if errors:
        print("\n⚠️  ÉCARTS DÉTECTÉS:")
        for error in errors:
            print(f"   - {error}")
        return False
    else:
        print("\n✅ Tous les KPIs correspondent aux valeurs attendues!")
        return True


if __name__ == "__main__":
    success = test_week_37()
    sys.exit(0 if success else 1)
