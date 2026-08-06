@echo off
chcp 65001 >nul
title INSTALAR script v8 - guardar filas en la hoja
cd /d "%~dp0"

echo.
echo ============================================================
echo  POR QUE FALLA LA FILA
echo ============================================================
echo  Su URL /exec sigue con script v4 (viejo).
echo  El PDF SI se guarda. La fila NO, hasta instalar v8.
echo.
echo  Este asistente:
echo   1. Copia el codigo v8 al portapapeles
echo   2. Abre la hoja correcta
echo   3. Abre el Bloc de notas con el codigo
echo.
pause

copy /Y "%~dp0apps-script\Codigo.gs" "%~dp0PEGAR-EN-APPS-SCRIPT-v8.gs" >nul
powershell -NoProfile -Command "Get-Content -Raw -Encoding UTF8 '%~dp0apps-script\Codigo.gs' | Set-Clipboard"

echo.
echo  [OK] Codigo v8 copiado al portapapeles.
echo.
echo ============================================================
echo  HAGA ESTO EN GOOGLE (en orden)
echo ============================================================
echo.
echo  1. Se abrira la HOJA. Espere a que cargue.
echo.
echo  2. Menu: Extensiones  →  Apps Script
echo.
echo  3. En el editor:
echo       Ctrl + A   (seleccionar todo)
echo       Ctrl + V   (pegar v8)
echo       Ctrl + S   (guardar)
echo.
echo  4. Arriba derecha: Implementar
echo       → Nueva implementacion
echo       → Tipo: Aplicacion web
echo       → Ejecutar como: Yo
echo       → Quien tiene acceso: Cualquier persona
echo       → Implementar
echo       → Permitir permisos
echo.
echo  5. COPIE la URL que termina en /exec
echo.
echo  6. Abra esa URL en el navegador. DEBE decir:
echo       "version":"2026-03-pdf-v8"
echo.
echo     Si dice v4, NO se implemento bien. Repita el paso 4.
echo.
echo  7. En la entrevista → Configurar → pegue la URL /exec
echo     → Guardar URL → Diagnosticar PDF
echo.
pause

start "" "https://docs.google.com/spreadsheets/d/1obx8kVIXXxk2P65LVXxHju3I1iUf3SsPgozlxKmZ5ek/edit"
timeout /t 2 >nul
start "" notepad.exe "%~dp0PEGAR-EN-APPS-SCRIPT-v8.gs"

echo.
echo  Cuando tenga la URL /exec nueva, peguela aqui y pulse Enter
echo  (o solo Enter para saltar):
echo.
set /p NEWURL=URL /exec: 

if not "%NEWURL%"=="" (
  echo.
  echo  Abriendo la URL para probar...
  start "" "%NEWURL%"
  echo.
  echo  Si ve version v8, guarde esa URL en la entrevista:
  echo  Configurar → pegar → Guardar URL
  echo.
  echo  URL guardada en portapapeles tambien.
  echo %NEWURL%| clip
)

echo.
echo  Listo. Pruebe guardar una ficha.
echo  Debe crear pestana "Fichas" en la hoja + PDF en Drive.
echo.
pause
