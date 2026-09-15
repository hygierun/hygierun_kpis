# 📋 KPI Recipes - Hygierun KPI Tool

> **Important**: Chaque KPI est calculé selon une recette précise. Ce document documente EXACTEMENT comment obtenir chaque valeur, y compris les filtres exacts et les sources de données.
> 
> Chaque recette doit être validée par Antoine avant implémentation en Phase 3.

---

## Format d'une Recette KPI

```yaml
KPI_ID: "unique_identifier"
Nom: "Nom du KPI"
Section: "Commerce | Livraison | SAV | Compta | Préparation | Achat_Appro"
Source: "Nom de la feuille Excel"
Colonnes: ["Col1", "Col2"]
Filtres: 
  - "Description du filtre (ex: Représentant IN [...])"
Agrégation: "SUM | AVG | COUNT | custom"
Comparaisons:
  N_Moins_1: false  # Même mois année précédente (août 2026 vs août 2025)
  Mois_Moins_1: false  # Mois précédent (août 2026 vs juillet 2026)
Notes: "Details additionnels sur le calcul"
Validation_Antoine: false  # À remplir après discussion
```

**Important**: 
- **N-1 (Year-over-Year)**: Même mois année précédente → vérifier la saisonnalité et la tendance long terme
- **Mois-1 (Month-over-Month)**: Mois précédent → mesurer les variations court terme et les actions récentes

---

## 🏪 SECTION COMMERCE

### KPI-COM-001: Panier Moyen (avec Franck)

```yaml
KPI_ID: "KPI-COM-001"
Nom: "Panier Moyen (avec Franck Leclancher)"
Section: "Commerce"
Source: "Factures"
Colonnes: ["Représentant", "Total HT", "Date"]
Filtres:
  - Représentant IN [
      "STRAZZULLA Francesco",
      "HUGUET Anthony", 
      "QUERE Julien",
      "SUMATRA Alexandre",
      "CAZAL Damien",
      "MARIE-LOUISE Caroline",
      "LECLANCHER Franck"
    ] OU Représentant = VIDE (NULL/NaN)
  - Période: Mois/semaine sélectionné
Agrégation: "AVG(Total HT)"
Formule_Texte: "Factures / filtre Repr = Commerciaux (x7) + Représentant vide"
Comparaisons:
  N_Moins_1: true
  Mois_Moins_1: false
Notes: |
  - Source: Feuille "Factures" (pas Commandes_Archivees)
  - Format des noms: "NOM Prénom" en majuscules (ex: LECLANCHER Franck)
  - IMPORTANT: Inclure aussi les factures avec Représentant = VIDE
  - Franck Leclancher inclus pour la vue globale de l'équipe
  - Août 2026: 364.60€ (966 factures)
Validation_Antoine: true
```

**Recette validée par Antoine** (14/09/2026):
- Source: Feuille **Factures** (pas Commandes_Archivees)
- Filtre: 7 commerciaux + **Représentant vide**
- Calcul: AVERAGE(Total HT) 
- Résultat août 2026: **364.60€** (966 factures, CA 352,200.77€)

---

### KPI-COM-002: Panier Moyen (sans Franck)

```yaml
KPI_ID: "KPI-COM-002"
Nom: "Panier Moyen (sans Franck Leclancher)"
Section: "Commerce"
Source: "Factures"
Colonnes: ["Représentant", "Total HT", "Date"]
Filtres:
  - Représentant IN [
      "STRAZZULLA Francesco",
      "HUGUET Anthony", 
      "QUERE Julien",
      "SUMATRA Alexandre",
      "CAZAL Damien",
      "MARIE-LOUISE Caroline"
    ] OU Représentant = VIDE (NULL/NaN)
  - Période: Mois/semaine sélectionné
Agrégation: "AVG(Total HT)"
Formule_Texte: "Factures / filtre Repr = Commerciaux (x6) + Représentant vide, SANS Franck"
Comparaisons:
  N_Moins_1: true
  Mois_Moins_1: false
Notes: |
  - Source: Feuille "Factures" (pas Commandes_Archivees)
  - Format des noms: "NOM Prénom" en majuscules
  - IMPORTANT: Inclure aussi les factures avec Représentant = VIDE
  - Franck Leclancher EXCLU car son activité machines haute valeur biaise la moyenne
  - Permet de voir la performance des commerciaux "standards"
  - Août 2026: 345.08€ (817 factures)
Validation_Antoine: true
```

**Recette validée par Antoine** (14/09/2026):
- Source: Feuille **Factures**
- Filtre: 6 commerciaux (sans Franck) + **Représentant vide**
- Calcul: AVERAGE(Total HT)
- Résultat août 2026: **345.08€** (817 factures, CA 281,933.69€)

---

### KPI-COM-003: Commandes clients passées

```yaml
KPI_ID: "KPI-COM-003"
Nom: "Commandes clients passées"
Section: "Commerce"
Source: "Commandes_Archivees + Commandes_ALivrer"
Colonnes: ["N° Commande"]
Filtres:
  - Période: Mois/semaine sélectionné
Agrégation: "COUNT(N° Commande distinct)"
Comparaison_N_Moins_1: true
Notes: |
  - Combine les commandes archivées ET celles en attente
  - Compte les commandes uniques
Validation_Antoine: false
```

---

### KPI-COM-004: CA HT facturé

```yaml
KPI_ID: "KPI-COM-004"
Nom: "CA HT facturé"
Section: "Commerce"
Source: "Factures"
Colonnes: ["Total HT", "Date"]
Filtres:
  - Période: Mois/semaine sélectionné
  - Status: Factures valides (non annulées)
Agrégation: "SUM(Total HT)"
Comparaison_N_Moins_1: true
Notes: |
  - Basé sur la feuille Factures (toutes les factures émises)
  - À noter: SAV peut générer des factures supplémentaires
Validation_Antoine: false
```

---

### KPI-COM-005: Nouveaux clients

```yaml
KPI_ID: "KPI-COM-005"
Nom: "Nouveaux clients"
Section: "Commerce"
Source: "Nouveaux_Client"
Colonnes: ["Client", "Total HT", "Date"]
Filtres:
  - Période: Mois/semaine sélectionné
Agrégation: "COUNT(Client distinct) + SUM(Total HT)"
Comparaison_N_Moins_1: true
Notes: |
  - Feuille dédiée aux nouveaux clients
  - Affiche nombre de nouveaux clients ET leur chiffre d'affaires
Validation_Antoine: false
```

---

### KPI-COM-006: Commandes sous seuil livraison

```yaml
KPI_ID: "KPI-COM-006"
Nom: "Commandes sous seuil livraison"
Section: "Commerce"
Source: "Commandes_ALivrer"
Colonnes: ["Total TTC", "Date"]
Filtres:
  - Période: Mois/semaine sélectionné
Agrégation: |
  COUNT(Total TTC < 100€) AS "Sous 100€"
  COUNT(Total TTC < 150€) AS "Sous 150€"
Notes: |
  - Important pour analyser la rentabilité des livraisons
  - Ces petites commandes peuvent être non rentables
Validation_Antoine: false
```

---

## 🚚 SECTION LIVRAISON

### KPI-LIV-001: CA livré par camion

```yaml
KPI_ID: "KPI-LIV-001"
Nom: "CA livré par camion"
Section: "Livraison"
Source: "Tournees"
Colonnes: ["N° Tournée", "Total HT", "Camion", "Date"]
Filtres:
  - Période: Mois/semaine sélectionné
Agrégation: "SUM(Total HT) grouped by Camion"
Notes: |
  - Affiche le CA généré par chaque camion
  - Peut révéler des différences de rendement entre véhicules
Validation_Antoine: false
```

---

### KPI-LIV-002: Nombre d'arrêts moyen

```yaml
KPI_ID: "KPI-LIV-002"
Nom: "Nombre d'arrêts moyen par tournée"
Section: "Livraison"
Source: "Tournees"
Colonnes: ["N° Tournée", "Nb Cdes", "Date"]
Filtres:
  - Période: Mois/semaine sélectionné
Agrégation: "AVG(Nb Cdes)"
Notes: |
  - Indicateur de densité et efficacité des tournées
  - Plus élevé = meilleure optimisation des routes
Validation_Antoine: false
```

---

### KPI-LIV-003: Distance moyenne par tournée

```yaml
KPI_ID: "KPI-LIV-003"
Nom: "Distance moyenne par tournée"
Section: "Livraison"
Source: "Tournees"
Colonnes: ["N° Tournée", "Distance", "Date"]
Filtres:
  - Période: Mois/semaine sélectionné
Agrégation: "AVG(Distance)"
Notes: |
  - ⚠️ À valider: colonne "Distance" existe-t-elle dans Tournees?
  - Sinon, peut être calculée à partir de Livraison + coordonnées clients
Validation_Antoine: false
```

---

### KPI-LIV-004: Délai livraison (Commande → Livraison)

```yaml
KPI_ID: "KPI-LIV-004"
Nom: "Délai moyen Commande → Livraison"
Section: "Livraison"
Source: "Expl_locale_Delais"
Colonnes: ["Délai Cde2Livraison", "Date Livraison"]
Filtres:
  - Période: Mois/semaine sélectionné
Agrégation: "AVG(Délai Cde2Livraison) en jours"
Notes: |
  - Mesure la réactivité opérationnelle
  - Objectif habituel: < X jours
Validation_Antoine: false
```

---

### KPI-LIV-005: Délai Livraison → Facturation

```yaml
KPI_ID: "KPI-LIV-005"
Nom: "Délai moyen Livraison → Facturation"
Section: "Livraison"
Source: "Expl_locale_Delais"
Colonnes: ["Délai Liv2Fact", "Date Facturation"]
Filtres:
  - Période: Mois/semaine sélectionné
Agrégation: "AVG(Délai Liv2Fact) en jours"
Comparaison_N_Moins_1: true
Notes: |
  - Indicateur administratif important
  - Impacte le trésorerie (combien de temps avant encaissement)
Validation_Antoine: false
```

---

### KPI-LIV-006: Pourcentage livraisons multiples

```yaml
KPI_ID: "KPI-LIV-006"
Nom: "Pourcentage de clients avec multiples livraisons"
Section: "Livraison"
Source: "Livraison + Tournees"
Colonnes: ["N° Client", "Date Livraison"]
Filtres:
  - Période: Mois/semaine sélectionné
Agrégation: |
  COUNT(N° Client avec >1 livraison) / COUNT(N° Client distinct) * 100
Notes: |
  - Indique la fidélité et la fréquence de commande
  - Haut = clients réguliers
Validation_Antoine: false
```

---

### KPI-LIV-007: Détails par chauffeur

```yaml
KPI_ID: "KPI-LIV-007"
Nom: "Performance par chauffeur"
Section: "Livraison"
Source: "Tournees"
Colonnes: ["Chauffeur", "N° Tournée", "Total HT", "Nb Cdes", "Distance", "Date"]
Filtres:
  - Période: Mois/semaine sélectionné
Agrégation: |
  Per Chauffeur:
    - COUNT(N° Tournée)
    - SUM(Total HT)
    - AVG(Nb Cdes)
    - AVG(Distance)
Notes: |
  - Tableau de bord détaillé par chauffeur
  - Permet d'identifier les meilleurs performeurs et ceux qui ont besoin de support
Validation_Antoine: false
```

---

## 🔧 SECTION SAV (Service Après-Vente)

### KPI-SAV-001: Factures d'intervention

```yaml
KPI_ID: "KPI-SAV-001"
Nom: "Factures d'intervention SAV"
Section: "SAV"
Source: "DocsNuls_SAV + Fiches PDF (ZIP)"
Colonnes: ["N° Facture", "Total HT", "Date", "Type Intervention"]
Filtres:
  - Période: Mois/semaine sélectionné
  - Status: Factures valides
Agrégation: |
  - COUNT(N° Facture) AS "Nombre"
  - SUM(Total HT) AS "CA SAV"
Comparaison_N_Moins_1: true
Notes: |
  - Combine les factures SAV listées + le contenu des fiches d'intervention
  - Les fiches PDF (73 en août) à parser pour extraire montants/types
  - ⚠️ Phase 3: Parser PDFs avec pdfplumber pour extraire ces infos
Validation_Antoine: false
```

---

### KPI-SAV-002: Déplacements facturés

```yaml
KPI_ID: "KPI-SAV-002"
Nom: "Déplacements facturés"
Section: "SAV"
Source: "Fiches PDF (ZIP)"
Colonnes: ["Type", "Déplacement Y/N", "Date"]
Filtres:
  - Période: Mois/semaine sélectionné
  - Type Intervention CONTAINS "Déplacement"
Agrégation: "COUNT(Déplacement) + SUM(Montant Déplacement)"
Notes: |
  - À extraire des fiches PDF d'intervention
  - Ligne distincte "Déplacement" ou incluse dans le tarif?
  - ⚠️ À clarifier avec Antoine sur la structure des fiches PDF
Validation_Antoine: false
```

---

### KPI-SAV-003: Main d'oeuvre (Heures)

```yaml
KPI_ID: "KPI-SAV-003"
Nom: "Main d'oeuvre (Heures)"
Section: "SAV"
Source: "Fiches PDF (ZIP)"
Colonnes: ["Heures Travail", "Technicien", "Date"]
Filtres:
  - Période: Mois/semaine sélectionné
Agrégation: "SUM(Heures Travail)"
Notes: |
  - À extraire des fiches PDF
  - Chaque fiche liste les heures par technicien
  - ⚠️ Phase 3: Parser PDFs pour extraire "heures" et "technicien"
Validation_Antoine: false
```

---

### KPI-SAV-004: Nombre d'interventions

```yaml
KPI_ID: "KPI-SAV-004"
Nom: "Nombre d'interventions"
Section: "SAV"
Source: "Fiches PDF (ZIP)"
Colonnes: ["N° Intervention", "Date"]
Filtres:
  - Période: Mois/semaine sélectionné
Agrégation: "COUNT(N° Intervention distinct)"
Notes: |
  - À extraire des fiches PDF (73 en août)
  - Une fiche = une intervention
  - ⚠️ À valider: comment sont numérotées les interventions?
Validation_Antoine: false
```

---

### KPI-SAV-005: Productivité (Heures / Interventions)

```yaml
KPI_ID: "KPI-SAV-005"
Nom: "Productivité SAV"
Section: "SAV"
Source: "Fiches PDF (ZIP)"
Columns: ["Heures Travail", "N° Intervention"]
Filtres:
  - Période: Mois/semaine sélectionné
Agrégation: "AVG(Heures Travail / Intervention) in heures/intervention"
Notes: |
  - Indica efficacité du service SAV
  - Ratio heures par intervention
  - ⚠️ À clarifier: interventions simples vs complexes?
Validation_Antoine: false
```

---

## 💰 SECTION COMPTABILITÉ

### KPI-COMPTA-001: Clients bloqués

```yaml
KPI_ID: "KPI-COMPTA-001"
Nom: "Clients bloqués"
Section: "Comptabilité"
Source: "Commandes_ALivrer"
Colonnes: ["N° Client", "Statut Blocage", "Date"]
Filtres:
  - Statut = "Bloqué"
  - Période: Mois/semaine sélectionné (ou snapshot courant)
Agrégation: "COUNT(N° Client distinct)"
Notes: |
  - À confirmer: colonne exacte du statut dans Commandes_ALivrer
  - Snapshot du moment (pas de filtre par période vraiment?)
Validation_Antoine: false
```

---

### KPI-COMPTA-002: Factures impayées

```yaml
KPI_ID: "KPI-COMPTA-002"
Nom: "Factures impayées"
Section: "Comptabilité"
Source: "Factures_Echues"
Colonnes: ["N° Facture", "Montant Restant Dû", "Date Échéance", "Nb Jours Retard"]
Filtres:
  - Statut = "Impayée"
  - Période: Mois/semaine sélectionné
Agrégation: |
  - COUNT(N° Facture) AS "Nombre"
  - SUM(Montant Restant Dû) AS "Total Impayé"
Comparaison_N_Moins_1: true
Notes: |
  - Feuille "Factures_Echues" dédiée aux impayés
  - À ajouter: nombre > 60 jours (KPI-COMPTA-003)
Validation_Antoine: false
```

---

### KPI-COMPTA-003: Factures impayées > 60 jours

```yaml
KPI_ID: "KPI-COMPTA-003"
Nom: "Factures impayées > 60 jours"
Section: "Comptabilité"
Source: "Factures_Echues"
Colonnes: ["N° Facture", "Montant Restant Dû", "Nb Jours Retard"]
Filtres:
  - Statut = "Impayée"
  - Nb Jours Retard > 60
  - Période: Mois/semaine sélectionné
Agrégation: |
  - COUNT(N° Facture) AS "Nombre"
  - SUM(Montant Restant Dû) AS "Total Impayé > 60j"
Notes: |
  - Plus critique: relances urgentes nécessaires
  - Impacte significativement la trésorerie
Validation_Antoine: false
```

---

### KPI-COMPTA-004: CA impayée

```yaml
KPI_ID: "KPI-COMPTA-004"
Nom: "Chiffre d'affaires impayé"
Section: "Comptabilité"
Source: "Factures_Echues"
Colonnes: ["Montant Restant Dû", "Date"]
Filtres:
  - Statut = "Impayée"
  - Période: Mois/semaine sélectionné
Agrégation: "SUM(Montant Restant Dû)"
Comparaison_N_Moins_1: true
Notes: |
  - Agregation du KPI-COMPTA-002
  - Vue du préjudice financier global
Validation_Antoine: false
```

---

## 📦 SECTION PRÉPARATION

### KPI-PREP-001: Total commandes préparées

```yaml
KPI_ID: "KPI-PREP-001"
Nom: "Total commandes préparées"
Section: "Préparation"
Source: "À déterminer avec Antoine"
Colonnes: []
Filtres:
  - Période: Mois/semaine sélectionné
Agrégation: "COUNT"
Notes: |
  - ⚠️ À clarifier avec Antoine
  - Quelle source? Commandes_Archivees filtrées?
  - Ou feuille dédiée "Préparation"?
Validation_Antoine: false
```

---

### KPI-PREP-002: CA préparation

```yaml
KPI_ID: "KPI-PREP-002"
Nom: "Chiffre d'affaires préparation"
Section: "Préparation"
Source: "À déterminer avec Antoine"
Colonnes: []
Filtres:
  - Période: Mois/semaine sélectionné
Agrégation: "SUM"
Notes: |
  - ⚠️ À clarifier avec Antoine
  - Comment distinguer "préparation" d'autres activités?
Validation_Antoine: false
```

---

### KPI-PREP-003: Détails par type de commande

```yaml
KPI_ID: "KPI-PREP-003"
Nom: "Détails préparation (Clients / Réassorts / Dépôtages)"
Section: "Préparation"
Source: "À déterminer avec Antoine"
Colonnes: ["Type Commande"]
Filtres:
  - Type IN ["Client", "Réassort", "Dépôtage"]
  - Période: Mois/semaine sélectionné
Agrégation: "COUNT + SUM by Type"
Notes: |
  - ⚠️ À clarifier:
    - Quelle colonne indique le type?
    - Comment les données sont-elles structurées?
  - Détail important pour suivre les ruptures et l'activité
Validation_Antoine: false
```

---

## 🛒 SECTION ACHAT / APPROVISIONNEMENT

### KPI-ACHAT-001: Valorisation du stock

```yaml
KPI_ID: "KPI-ACHAT-001"
Nom: "Valorisation du stock"
Section: "Achat/Approvisionnement"
Source: "Achats_Contenaires"
Colonnes: ["Ref Produit", "Quantité", "Prix Unitaire", "Date"]
Filtres:
  - Période: Mois/semaine sélectionné (ou snapshot courant)
Agrégation: "SUM(Quantité * Prix Unitaire)"
Notes: |
  - À confirmer: comment est calculé le stock (FIFO/LIFO/Moyenne)?
  - Snapshot ou évolution sur la période?
Validation_Antoine: false
```

---

### KPI-ACHAT-002: Couverture de stock

```yaml
KPI_ID: "KPI-ACHAT-002"
Nom: "Couverture de stock (jours)"
Section: "Achat/Approvisionnement"
Source: "Achats_Contenaires + Livraison (consommation)"
Colonnes: ["Stock", "Consommation Quotidienne"]
Filtres:
  - Période: Mois/semaine sélectionné
Agrégation: "AVG(Stock / Consommation Quotidienne) in jours"
Notes: |
  - ⚠️ À clarifier:
    - Comment mesurer la consommation?
    - Par produit ou global?
  - Important pour le cash flow et la gestion des ruptures
Validation_Antoine: false
```

---

### KPI-ACHAT-003: Commandes fournisseur

```yaml
KPI_ID: "KPI-ACHAT-003"
Nom: "Commandes fournisseur"
Section: "Achat/Approvisionnement"
Source: "Achats_Contenaires"
Colonnes: ["N° Commande Fournisseur", "Montant", "Date"]
Filtres:
  - Période: Mois/semaine sélectionné
Agrégation: |
  - COUNT(N° Commande) AS "Nombre"
  - SUM(Montant) AS "Total Achats"
Comparaison_N_Moins_1: true
Notes: |
  - Vue du volume d'approvisionnement
  - À croiser avec les données de livraison pour vérifier cohérence
Validation_Antoine: false
```

---

## 📝 Statut de Validation

| Section | KPIs | Validés | À Valider |
|---------|------|---------|-----------|
| 🏪 Commerce | 6 | ✅ 2 (COM-001, COM-002) | ❌ 4 |
| 🚚 Livraison | 7 | ❌ 0 | ✅ 7 |
| 🔧 SAV | 5 | ❌ 0 | ✅ 5 |
| 💰 Comptabilité | 4 | ❌ 0 | ✅ 4 |
| 📦 Préparation | 3 | ❌ 0 | ✅ 3 |
| 🛒 Achat | 3 | ❌ 0 | ✅ 3 |
| **TOTAL** | **28** | **2** | **26** |

**Validations complétées** :
- ✅ 14/09/2026 - KPI-COM-001 (Panier Moyen AVEC Franck): 364.60€
- ✅ 14/09/2026 - KPI-COM-002 (Panier Moyen SANS Franck): 345.08€

**Découvertes importantes** :
- Format des représentants: "NOM Prénom" (ex: LECLANCHER Franck)
- Inclure TOUJOURS les factures avec Représentant vide dans les filtres
- Source correcte: Feuille "Factures" (pas Commandes_Archivees)

---

## 🔄 Processus de Validation

Pour chaque KPI à valider:

1. **Antoine valide la recette**: Sources, filtres, agrégations ✅
2. **Claude implémente le calcul** dans `kpi_calculator.py`
3. **Test sur données août 2026**: Résultats vs. dashboard Antoine
4. **Ajustements si nécessaire**
5. **Marquer comme validé**: `Validation_Antoine: true`

Une fois tous les KPIs validés → **Phase 3 complétée** → **Phase 4: PPTX Generation**

---

## 📚 Notes d'Architecture

Cette documentation KPI servira de base pour :

1. **Créer une classe `KPIRecipe`** en Python qui encapsule:
   - Metadata (ID, nom, section)
   - Source (feuille Excel)
   - Filtres (représentants, périodes, etc.)
   - Agrégation (SUM, AVG, COUNT, custom)
   - Comparaisons N-1

2. **Charger les recettes depuis config.py**:
   ```python
   KPI_RECIPES = {
       "KPI-COM-001": {
           "nom": "Panier Moyen (avec Franck)",
           "source": "Commandes_Archivees",
           "filtres": {"representant": [...]},
           "aggregation": "avg",
           "colonne": "Total HT"
       }
   }
   ```

3. **Faire tourner les calculs dynamiquement**:
   ```python
   for kpi_id, recipe in KPI_RECIPES.items():
       df = loader.get_dataframe(recipe["source"])
       value = calculate_kpi(df, recipe)
   ```

---

**Prochaine étape**: Antoine valide les 2 KPIs Commerce (COM-001, COM-002) puis on valide ensemble les autres sections.
