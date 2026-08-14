@echo off
title DAREDAB PPM Intelligence
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo.
    echo ERREUR : le dossier .venv est introuvable dans ce dossier.
    echo Contacte la personne qui a installe l'application.
    echo.
    pause
    exit /b 1
)

echo Demarrage de PPM Intelligence...
echo Une page va s'ouvrir dans le navigateur dans quelques secondes.
echo Pour arreter l'application, ferme cette fenetre noire.
echo.

".venv\Scripts\python.exe" -m streamlit run dashboard.py

pause
