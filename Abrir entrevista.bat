@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo  Ficha Social UNCP - servidor local
echo ============================================
echo.
echo El micrófono solo funciona en http://localhost
echo (no abriendo el HTML con doble clic).
echo.
echo Se abrirá el navegador en unos segundos...
echo Para cerrar: pulse Ctrl+C en esta ventana.
echo ============================================
echo.

REM Preferir Python
where python >nul 2>&1
if %errorlevel%==0 (
  start "" "http://127.0.0.1:8765/index.html"
  python -m http.server 8765 --bind 127.0.0.1
  goto :eof
)

where py >nul 2>&1
if %errorlevel%==0 (
  start "" "http://127.0.0.1:8765/index.html"
  py -m http.server 8765 --bind 127.0.0.1
  goto :eof
)

echo ERROR: No se encontró Python.
echo Instale Python desde https://www.python.org/  o abra index.html
echo en Chrome sabiendo que el micrófono puede no funcionar.
echo.
pause
start "" "%~dp0index.html"
