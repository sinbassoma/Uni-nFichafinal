@echo off
chcp 65001 >nul
title Actualizar script PDF - Ficha Social UNCP
cd /d "%~dp0"

echo.
echo ============================================================
echo   ACTUALIZAR SCRIPT PARA GUARDAR PDF EN DRIVE
echo ============================================================
echo.
echo  Problema: Google aun usa la version VIEJA del script.
echo  Solo guarda filas en la hoja, NO crea PDF en Drive.
echo.
echo  Voy a:
echo   1) Copiar Codigo.gs al portapapeles
echo   2) Abrir el proyecto de Apps Script
echo   3) Abrir la carpeta de Drive
echo   4) Abrir la URL de prueba
echo.
pause

REM Copiar codigo al portapapeles
powershell -NoProfile -Command "Get-Content -Raw -Encoding UTF8 '%~dp0apps-script\Codigo.gs' | Set-Clipboard"
if errorlevel 1 (
  echo ERROR: no se pudo copiar el codigo.
  pause
  exit /b 1
)

echo.
echo  [OK] Codigo copiado al portapapeles.
echo.
echo  AHORA EN LA PAGINA DE APPS SCRIPT QUE SE ABRE:
echo.
echo   1. Click dentro del editor de codigo
echo   2. Ctrl+A  (seleccionar todo)
echo   3. Ctrl+V  (pegar)
echo   4. Ctrl+S  (guardar)
echo   5. Arriba derecha: Implementar
echo   6. Administrar implementaciones
echo   7. Icono del LAPIZ
echo   8. Version = "Nueva version"
echo   9. Implementar  (y Permitir si pide Drive)
echo.
echo  Luego en la pestana de prueba debe verse:
echo     "version":"2026-03-pdf-v3"
echo     "pdfReady":true
echo.
pause

start "" "https://script.google.com/home/projects/1cgMbCZB7YmHcbMWBZbvINQ2dhO9P_jwPjXPTCu73_YzhtaxrQQcsJGgW/edit"
timeout /t 2 >nul
start "" "https://drive.google.com/drive/folders/1dqy2HuWf6IYVGasuyrtff2id7Jf3aRXv"
timeout /t 1 >nul
start "" "https://script.google.com/macros/s/AKfycbxb0plUajtN66LSpK65O1_k-RNM5oYpv2qLisaar_-ukG_eRBUVtr6IfxmSqqDTx6No/exec"

echo.
echo  Cuando la URL de prueba muestre pdfReady true,
echo  vuelva a la entrevista y guarde una ficha.
echo.
pause
