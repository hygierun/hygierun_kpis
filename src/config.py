"""Configuration pour le tool KPI Hygierun"""

# ============================================================================
# CONFIGURATION DES FEUILLES EXCEL
# ============================================================================

EXCEL_SHEETS_CONFIG = {
    "Factures_Echues": {
        "required_columns": ["Date", "Client", "Client (réf.)", "Echéance", "Nb JEch", "Restant dû"],
        "description": "Factures impayées et clients bloqués"
    },
    "Tournees": {
        "required_columns": ["N°", "Jour", "Date", "Désignation", "Chauffeur", "Total HT",
                            "Nb cdes", "Nb bl EP", "Nb bl EC", "Nb bl AF", "Nb bl A", "Nb fact"],
        "description": "Détails des tournées de livraison"
    },
    "Livraison": {
        "required_columns": ["N°", "Date", "Liv. souhaitée", "Client", "Total HT",
                            "Total TTC", "Tournée", "Edité"],
        "description": "Détails de chaque livraison"
    },
    "Factures": {
        "required_columns": ["N°", "Date", "Client", "Total HT", "Total TTC", "Echéance"],
        "description": "Détails des factures émises"
    },
    "Commandes_Archivees": {
        "required_columns": ["N°", "Date", "Client", "Total HT", "Total TTC"],
        "description": "Commandes archivées/livrées"
    },
    "Commandes_ALivrer": {
        "required_columns": ["Bloqué", "N°", "Date", "Client", "Total TTC"],
        "description": "Commandes en attente de livraison"
    },
    "DocsNuls_SAV": {
        "required_columns": ["N°", "Date", "Client", "Total HT", "Total TTC"],
        "description": "Factures nullifiées SAV"
    },
    "Achats_Contenaires": {
        "required_columns": ["N°", "Date", "Fournisseur", "Total HT"],
        "description": "Commandes d'achat de contenaires"
    },
    "Nouveaux_Client": {
        "required_columns": ["Client", "1F année", "08/26 HT"],
        "description": "Nouveaux clients de la période"
    },
    "Expl_locale_Delais": {
        "required_columns": ["N°", "Date", "Cde N°", "Date Creation Cde",
                            "Fact N°", "Délai Cde2Livraison"],
        "description": "Détails des délais commande-livraison-facture"
    }
}

# ============================================================================
# CONFIGURATION DES SECTIONS ET KPIs
# ============================================================================

KPI_SECTIONS = {
    "Commerce": {
        "couleur": "#2C5282",  # Bleu Hygierun
        "kpis": [
            "commandes_clients_passees",
            "ca_ht_facture",
            "panier_moyen",
            "nouveaux_clients",
            "cdes_sous_seuil_100",
            "cdes_sous_seuil_150"
        ]
    },
    "Preparation": {
        "couleur": "#2C5282",
        "kpis": [
            "total_cdes_preparees",
            "ca_preparation",
            "cdes_clients",
            "cdes_reassort",
            "depotage_container"
        ]
    },
    "Livraison": {
        "couleur": "#2C5282",
        "kpis": [
            "ca_livre",
            "nombre_arrets_moyen",
            "distance_moyenne",
            "delai_livraison_total",
            "multiples_livraisons_pct",
            "enleves_clients",
            "chaufieurs_stats"
        ]
    },
    "SAV": {
        "couleur": "#A0522D",  # Marron/rouge
        "kpis": [
            "factures_sav",
            "ca_sav",
            "deplacements_factures",
            "main_oeuvre_heures",
            "nombre_interventions",
            "productivite"
        ]
    },
    "Compta": {
        "couleur": "#5C2E7F",  # Pourpre
        "kpis": [
            "clients_bloques",
            "factures_impayees_total",
            "ca_impayee",
            "factures_impayees_60j",
            "delai_liv_fact"
        ]
    },
    "Achat_Appro": {
        "couleur": "#2F5233",  # Vert
        "kpis": [
            "valorisation_stock",
            "couverture_stock",
            "commandes_fournisseur_recues",
            "commandes_fournisseur_passees"
        ]
    }
}

# ============================================================================
# CONFIGURATION DU STYLE POWERPOINT
# ============================================================================

PPTX_CONFIG = {
    "slide_width": 10,  # pouces
    "slide_height": 7.5,  # pouces
    "fonts": {
        "title": "Segoe UI",
        "body": "Segoe UI",
        "data": "Segoe UI"
    },
    "colors": {
        "background": "#FFFFFF",
        "text_dark": "#1F3A5F",
        "text_light": "#666666",
        "accent": "#2C5282",
        "positive": "#28A745",  # Vert pour +
        "negative": "#DC3545"   # Rouge pour -
    }
}

# ============================================================================
# PERIODES PAR DEFAUT
# ============================================================================

PERIODS = {
    "monthly": {
        "name": "Mensuel",
        "columns": ["Date"]  # Filtrer par mois
    },
    "weekly": {
        "name": "Hebdomadaire",
        "columns": ["Date"]  # Filtrer par semaine
    }
}
