@echo off
REM Script de lancement de l'application KPI Hygierun (Windows)
REM Usage: run.bat

echo ================================================================
echo 🚀 Demarrage de l'Application KPI Hygierun
echo ================================================================
echo.

REM Verifier que Python est installe
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python n'est pas installe ou pas dans le PATH
    echo.
    echo Veuillez installer Python 3.8+ depuis https://www.python.org
    pause
    exit /b 1
)

echo ✅ Python detecte
echo.

REM Verifier que Streamlit est installe
python -m pip show streamlit >nul 2>&1
if errorlevel 1 (
    echo 📦 Installation de Streamlit et dependances...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo ❌ Erreur lors de l'installation
        pause
        exit /b 1
    )
    echo ✅ Dependances installes
    echo.
)

echo ================================================================
echo 📊 Application lancee!
echo ================================================================
echo.
echo L'application est disponible a: http://localhost:8501
echo.
echo Pour arreter l'application, appuyez sur Ctrl+C
echo.
echo ================================================================
echo.

REM Lancer l'application
python -m streamlit run app.py --client.showErrorDetails=true

pause
