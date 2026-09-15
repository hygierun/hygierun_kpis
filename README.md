# 📊 Générateur de Rapports KPI Hygierun

Application automatisée pour générer les rapports KPI hebdomadaires de Hygierun en quelques clics.

**Avant** : 2-3 heures de copier/coller Excel + calculs  
**Après** : 5 minutes en cliquant sur un bouton ✨

---

## 🚀 Démarrage rapide

### Option 1 : Script de lancement (recommandé)

```bash
./run.sh
```

Le script va :
- ✅ Vérifier que tous les outils sont installés
- ✅ Lancer l'application
- ✅ Afficher le lien d'accès

### Option 2 : Commande directe

```bash
streamlit run app.py
```

---

## 📖 Utilisation

### Étape 1️⃣ : Ouvrir l'application
L'application s'ouvre automatiquement dans votre navigateur à : **http://localhost:8501**

### Étape 2️⃣ : Charger les fichiers

**Fichier de données brutes** 📁
- Fichier Excel exporté depuis l'ERP
- Contient toutes les commandes, factures, livraisons, etc.
- Exemple : `Input_Donnees_Brutes_Hebdo.xlsx`

**Fichier snapshot** 📋
- Modèle Excel de rapport (avec headers pré-formatés)
- Exemple : `KPIs_weekly.xlsx`

### Étape 3️⃣ : Sélectionner la semaine

- Entrez le **numéro de semaine ISO** (1-52)
- Exemple : `37` pour la semaine 37

### Étape 4️⃣ : Générer le rapport

Cliquez sur le bouton **"🚀 Générer Rapport"**

L'application va :
1. ✅ Charger et valider les données
2. ✅ Calculer tous les KPIs
3. ✅ Créer 6 feuilles de détail
4. ✅ Remplir le snapshot avec les valeurs

### Étape 5️⃣ : Télécharger le résultat

Cliquez sur **"📥 Télécharger rapport"** pour récupérer le fichier Excel complète.

---

## 📊 Fichiers générés

Le rapport contient :

| Feuille | Description |
|---------|-------------|
| **Snapshot** (principale) | Vue d'ensemble avec tous les KPIs |
| **Semaine_XX_Commandes** | Détail de toutes les commandes filtrées |
| **Semaine_XX_Factures** | Détail de toutes les factures filtrées |
| **Semaine_XX_Livraisons** | Détail de toutes les livraisons filtrées |
| **Semaine_XX_Délai** | Délai moyen commande→livraison |
| **Semaine_XX_EnAttente** | Commandes en attente de livraison |

---

## 🔍 KPIs calculés

### Commandes
- Nombre de commandes
- CA HT avec représentant
- CA HT sans représentant

### Factures
- Nombre de factures
- CA HT avec représentant
- CA HT sans représentant

### Livraisons
- Nombre de livraisons
- CA HT avec représentant
- CA HT sans représentant
- **Délai moyen** (commande → livraison)
- Nombre de commandes en attente

### Créances
- Nombre de créances échues
- Montant total créances

---

## ⚙️ Configuration système

### Prérequis
- Python 3.8+
- 100 MB d'espace disque

### Installation automatique
Le script `run.sh` installe automatiquement les dépendances si nécessaire :
- `streamlit` - Interface web
- `pandas` - Traitement de données
- `openpyxl` - Manipulation Excel

### Installation manuelle
```bash
pip install streamlit pandas openpyxl
```

---

## ⚠️ Messages d'erreur

### "Veuillez charger les 2 fichiers Excel"
→ Vérifiez que vous avez uploadé **les deux** fichiers

### "Erreur lors du chargement des données"
→ Vérifiez que les fichiers Excel sont au bon format (xlsx)

### Les colonnes manquent de données
→ Vérifiez que le fichier brut contient les feuilles requises

---

## 💡 Conseils d'utilisation

✅ **À faire**
- Vérifier que le fichier brut est récemment extrait
- Utiliser un numéro de semaine existant (33-52)
- Sauvegarder le fichier généré immédiatement

❌ **À éviter**
- Charger un fichier brut corrompu
- Modifier le fichier pendant la génération
- Utiliser un numéro de semaine invalide (0 ou >52)

---

## 🛑 Arrêt de l'application

Appuyez sur **Ctrl+C** dans le terminal pour arrêter l'application.

---

## 📞 Support

Si vous rencontrez un problème :
1. Vérifiez que Python 3.8+ est installé
2. Relancez le script `run.sh`
3. Vérifiez le format de vos fichiers Excel

---

**Version** : 1.0  
**Dernière mise à jour** : Septembre 2026
