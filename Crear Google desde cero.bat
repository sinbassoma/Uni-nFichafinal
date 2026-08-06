@echo off
chcp 65001 >nul
title Crear Google Sheet + Drive + Script (desde cero)
cd /d "%~dp0"

echo.
echo ============================================================
echo   FICHA SOCIAL UNCP - CREAR GOOGLE DESDE CERO
echo ============================================================
echo.
echo  El PDF no se guardaba porque el script VIEJO en Google
echo  solo escribia la hoja. Vamos a crear todo NUEVO.
echo.
echo  Pasos que hara usted (guiados):
echo   1. Crear hoja nueva
echo   2. Pegar el script (codigo ya en el portapapeles)
echo   3. Implementar como Aplicacion web (NUEVA)
echo   4. Copiar la URL /exec a la entrevista
echo.
pause

powershell -NoProfile -Command "Get-Content -Raw -Encoding UTF8 '%~dp0apps-script\Codigo.gs' | Set-Clipboard"
echo.
echo  [OK] Codigo.gs copiado al portapapeles.
echo.

echo ------------------------------------------------------------
echo  PASO 1 - HOJA NUEVA
echo ------------------------------------------------------------
echo  Se abrira sheets.new
echo  - Nombre: Fichas Social UNCP
echo  - Guarde (Ctrl+S)
echo  - Deje la pestana abierta
echo.
pause
start "" "https://sheets.new"
echo.
echo  Cuando tenga la hoja creada y guardada, pulse una tecla...
pause >nul

echo.
echo ------------------------------------------------------------
echo  PASO 2 - APPS SCRIPT DENTRO DE LA HOJA
echo ------------------------------------------------------------
echo  En la hoja de Google:
echo   Menu: Extensiones  -^>  Apps Script
echo.
echo  En el editor:
echo   1. Ctrl+A  (seleccionar todo el ejemplo)
echo   2. Ctrl+V  (pegar el codigo - ya esta en el portapapeles)
echo   3. Ctrl+S  (guardar)
echo   4. Nombre del proyecto: Ficha Social UNCP
echo.
echo  NOTA: SPREADSHEET_ID y PDF_FOLDER_ID pueden quedar vacios.
echo  El script usara esta hoja y creara la carpeta PDF solo.
echo.
pause

echo.
echo ------------------------------------------------------------
echo  PASO 3 - IMPLEMENTAR (NUEVA aplicacion web)
echo ------------------------------------------------------------
echo  En Apps Script, arriba a la derecha:
echo.
echo   Implementar  -^>  Nueva implementacion
echo.
echo   Tipo: Aplicacion web   (engranaje si no sale)
echo   Descripcion: ficha uncp
echo   Ejecutar como: Yo
echo   Quien tiene acceso: Cualquier persona
echo.
echo   Implementar
echo   - Si sale "Google no ha verificado esta app":
echo     Avanzado  -^>  Ir a Ficha Social UNCP (no seguro)  -^>  Permitir
echo.
echo   COPIE la "URL de la aplicacion web"
echo   Debe terminar en:  /exec
echo.
echo   Ejemplo:
echo   https://script.google.com/macros/s/AKfycb...../exec
echo.
pause

echo.
echo ------------------------------------------------------------
echo  PASO 4 - PROBAR LA URL
echo ------------------------------------------------------------
echo  Pegue la URL /exec en el navegador y Enter.
echo  Debe verse algo como:
echo    "version":"2026-03-pdf-v4"
echo    "pdfReady":true
echo    "folderAccessible":true
echo.
echo  Si folderAccessible es false, ejecute en Apps Script
echo  la funcion probarGuardarPdf y acepte permisos de Drive.
echo.
pause

echo.
echo ------------------------------------------------------------
echo  PASO 5 - PEGAR URL EN LA ENTREVISTA
echo ------------------------------------------------------------
echo  1. Abra con el Bloc de notas:
echo     ficha-entrevista.js
echo.
echo  2. Busque (Ctrl+F):
echo     SHEETS_WEBAPP_URL_DEFAULT
echo.
echo  3. Deje asi (con SU url entre comillas):
echo.
echo     const SHEETS_WEBAPP_URL_DEFAULT =
echo       'https://script.google.com/macros/s/SU_CODIGO/exec';
echo.
echo  4. Guarde el archivo.
echo.
echo  5. En la entrevista, al final:
echo     Configurar -^> pegar la misma URL -^> Guardar URL
echo     (por si el navegador tenia una URL vieja)
echo.
echo  6. Pulse "Diagnosticar PDF" - debe decir Script listo.
echo  7. Complete una ficha de prueba.
echo.
echo  Resultado esperado:
echo   - Nueva fila en la hoja
echo   - PDF en carpeta "Fichas Social UNCP - PDF" en Drive
echo   - Columna "Link PDF" con el enlace
echo.
start "" notepad "%~dp0ficha-entrevista.js"
echo.
echo  Listo. Cuando tenga la URL /exec, peguela en ficha-entrevista.js
echo  y tambien en la pantalla Configurar de la entrevista.
echo.
pause
