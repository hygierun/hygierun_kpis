#!/bin/bash

# Script de lancement de l'application KPI Hygierun
# Usage: ./run.sh

set -e

echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "🚀 Démarrage de l'Application KPI Hygierun"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""

# Vérifier que Streamlit est installé
if ! command -v streamlit &> /dev/null; then
    echo "❌ Streamlit n'est pas installé"
    echo ""
    echo "Installation de Streamlit..."
    pip install --break-system-packages streamlit
    echo ""
fi

# Obtenir le répertoire du script
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

echo "📁 Répertoire: $SCRIPT_DIR"
echo "🔧 Vérification des dépendances..."

# Vérifier les modules Python
python3 -c "
import sys
try:
    import streamlit
    import pandas
    import openpyxl
    print('✅ Toutes les dépendances sont disponibles')
except ImportError as e:
    print(f'❌ Erreur: {e}')
    sys.exit(1)
"

echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "📊 Application lancée!"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""
echo "L'application est disponible à: http://localhost:8501"
echo ""
echo "Pour arrêter l'application, appuyez sur Ctrl+C"
echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""

# Lancer l'application
streamlit run app.py --client.showErrorDetails=true
