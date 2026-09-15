# 📘 Guide de déploiement - Streamlit Cloud

## PHASE 2 : Créer le repo GitHub privé

### Étape 2.1 : Créer un compte GitHub pour Hygierun

1. Aller sur https://github.com/signup
2. **Email** : saisissez un email professionnel Hygierun (ex: `contact@hygierun.re`)
3. **Username** : `hygierun-kpi` (ou ce que vous préférez)
4. **Mot de passe** : créer un mot de passe fort
5. Cliquer **"Create account"**
6. Valider l'email

### Étape 2.2 : Créer un nouveau repo privé

1. Une fois connecté sur GitHub, cliquer **"+"** (en haut à droite)
2. Sélectionner **"New repository"**
3. **Repository name** : `hygierun-kpi-tool`
4. **Description** : `Générateur de rapports KPI Hygierun`
5. **Visibilité** : ⚫ **PRIVATE** (très important!)
6. ☑️ Cocher **"Add a README file"**
7. Cliquer **"Create repository"**

### Étape 2.3 : Uploader le code

Option A (Recommandé - interface GitHub) :

1. Sur la page du repo, cliquer **"Add file"** → **"Upload files"**
2. Glisser-déposer le dossier `hygierun_kpi_tool` entier
3. OU cliquer et sélectionner les fichiers :
   - `app.py`
   - `requirements.txt`
   - `README.md`
   - Dossier `src/` (entier avec tous les fichiers)
   - `.streamlit/config.toml`

4. Ajouter un message : `Initial commit - KPI tool`
5. Cliquer **"Commit changes"**

Option B (Git en ligne de commande - plus technique) :

```bash
cd /home/claude/hygierun_kpi_tool
git init
git add .
git commit -m "Initial commit - KPI tool"
git branch -M main
git remote add origin https://github.com/VOTRE_USERNAME/hygierun-kpi-tool.git
git push -u origin main
```

---

## PHASE 3 : Déployer sur Streamlit Cloud

### Étape 3.1 : Créer un compte Streamlit Cloud

1. Aller sur https://streamlit.io/cloud
2. Cliquer **"Sign in"** (ou **"Get started for free"**)
3. Cliquer **"Sign in with GitHub"**
4. Autoriser Streamlit à accéder à votre GitHub
5. Sélectionner le repository `hygierun-kpi-tool`
6. Sélectionner la branche `main`
7. Spécifier le chemin principal :
   - **Main file path** : `app.py`

8. Cliquer **"Deploy"** 🚀

Streamlit va automatiquement :
- ✅ Lire `requirements.txt`
- ✅ Installer les dépendances
- ✅ Lancer l'application
- ✅ Générer une URL publique

### Étape 3.2 : Récupérer le lien public

Une fois déployée, vous recevrez une URL comme :
```
https://hygierun-kpi.streamlit.app
```

Cette URL est **publique** et accessible à tout le monde ! 🌐

---

## PHASE 4 : Test & partage

### Étape 4.1 : Tester l'application

1. Ouvrir le lien reçu (ex: `https://hygierun-kpi.streamlit.app`)
2. Uploader un fichier Excel de test
3. Uploader un fichier snapshot de test
4. Entrer un numéro de semaine
5. Cliquer **"🚀 Générer Rapport"**
6. Télécharger et vérifier le fichier généré

### Étape 4.2 : Partager le lien

Partager cette URL aux utilisateurs :
```
https://hygierun-kpi.streamlit.app
```

Aucune installation nécessaire pour eux ! ✨

---

## 🔄 Mises à jour futures

Si vous modifiez le code :

1. Modifier les fichiers localement
2. Uploader les changements sur GitHub
3. Streamlit Cloud redéploie automatiquement (quelques secondes)

---

## ⚠️ Dépannage

### L'app affiche une erreur à l'ouverture

→ Vérifier que `requirements.txt` contient :
```
streamlit==1.63.0
pandas>=1.4.0
openpyxl>=3.10.0
```

### Le déploiement échoue

→ Vérifier que :
- ✅ Le repo est **privé** mais accessible par Streamlit
- ✅ Le fichier `app.py` est à la racine
- ✅ Les permissions GitHub sont correctes

### L'utilisateur final n'arrive pas à uploader

→ Vérifier que le navigateur accepte les uploads (Firefox/Chrome/Safari)

---

## 📞 Support Streamlit Cloud

Si problèmes :
- Documentation : https://docs.streamlit.io/streamlit-cloud
- Forum : https://discuss.streamlit.io/

---

**Durée totale** : ~15-20 minutes  
**Coût** : Gratuit (tier Streamlit Cloud Community)  
**Disponibilité** : 24/7 une fois déployé
