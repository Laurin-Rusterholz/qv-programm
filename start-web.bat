@echo off
REM Startet die Browser-App lokal (Windows). Einfach doppelklicken.
REM Alles bleibt in deinem Browser; nur ein kleiner lokaler Webserver wird gestartet.

cd /d "%~dp0web"
set PORT=8000

set "PY="
where python >nul 2>nul && set "PY=python"
if not defined PY ( where py >nul 2>nul && set "PY=py" )
if not defined PY (
  echo Python 3 wird benoetigt. Bitte von https://www.python.org installieren
  echo und beim Installieren "Add Python to PATH" anhaken.
  pause
  exit /b 1
)

echo ============================================================
echo   QV-Pruefungsassistent (Browser-App)
echo   Browser oeffnet sich. Falls nicht: http://localhost:%PORT%
echo   Zum Beenden dieses Fenster schliessen.
echo ============================================================

start "" "http://localhost:%PORT%"
%PY% -m http.server %PORT%
pause
