@echo off
title SINAPI Explorer Launcher
cd /d "%~dp0"
echo Verificando dependencias...
pip install --upgrade streamlit pandas openpyxl text-unidecode --quiet
echo.
echo Iniciando... O navegador abrira em instantes.
streamlit run app.py
pause