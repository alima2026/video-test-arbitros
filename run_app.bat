@echo off
title Video Test Arbitros - Configurable V8
cd /d "%~dp0"
echo Instalando/actualizando dependencias...
python -m pip install -r requirements.txt
echo.
echo Iniciando programa...
python -m streamlit run app.py
pause
