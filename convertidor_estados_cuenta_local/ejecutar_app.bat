@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Convertidor de estados de cuenta
where python >nul 2>nul
if errorlevel 1 (
  echo No se encontro Python. Ejecuta primero instalar_dependencias.bat
  pause
  exit /b 1
)
python -c "import streamlit" >nul 2>nul
if errorlevel 1 (
  echo Faltan dependencias. Ejecuta primero instalar_dependencias.bat
  pause
  exit /b 1
)
python -m streamlit run app.py --server.headless false --browser.gatherUsageStats false
pause
