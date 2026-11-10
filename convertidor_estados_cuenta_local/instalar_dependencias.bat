@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Instalador - Convertidor de estados de cuenta
echo Instalando dependencias del programa local...
where python >nul 2>nul
if errorlevel 1 (
  echo No se encontro Python. Instala Python 3.12 o posterior y vuelve a ejecutar este archivo.
  pause
  exit /b 1
)
python -m pip install --upgrade pip
if errorlevel 1 goto :error
python -m pip install -r requirements.txt
if errorlevel 1 goto :error
echo.
echo Instalacion terminada correctamente. Ya puedes abrir ejecutar_app.bat
pause
exit /b 0

:error
echo.
echo No se pudo completar la instalacion. Revisa el mensaje anterior.
pause
exit /b 1
