@echo off
REM Startskript fuer Windows. Einfach doppelklicken.
REM Beim ersten Mal werden eine virtuelle Umgebung angelegt und alle Pakete
REM installiert. Danach startet die App einfach.

cd /d "%~dp0"

echo ============================================================
echo   QV-Pruefungsassistent wird vorbereitet ...
echo ============================================================

REM 1. Python suchen (zuerst "python", dann der Launcher "py").
set "PY="
where python >nul 2>nul && set "PY=python"
if not defined PY (
  where py >nul 2>nul && set "PY=py"
)
if not defined PY (
  echo.
  echo FEHLER: Python wurde nicht gefunden.
  echo Bitte Python 3.10 oder neuer von https://www.python.org installieren
  echo und beim Installieren "Add Python to PATH" anhaken.
  echo Danach diese Datei erneut doppelklicken.
  echo.
  pause
  exit /b 1
)

REM 2. Virtuelle Umgebung anlegen (nur beim ersten Mal).
if not exist ".venv\" (
  echo Erstelle virtuelle Umgebung (einmalig) ...
  %PY% -m venv .venv
  if errorlevel 1 (
    echo Konnte keine virtuelle Umgebung anlegen.
    pause
    exit /b 1
  )
)

REM 3. Virtuelle Umgebung aktivieren.
call ".venv\Scripts\activate.bat"

REM 4. Pakete installieren (nur beim ersten Mal).
if not exist ".venv\.installiert" (
  echo Installiere Pakete (einmalig, das kann ein paar Minuten dauern) ...
  python -m pip install --upgrade pip
  python -m pip install -r requirements.txt
  if errorlevel 1 (
    echo.
    echo FEHLER: Die Pakete konnten nicht installiert werden (Internet pruefen).
    pause
    exit /b 1
  )
  echo fertig> ".venv\.installiert"
)

echo.
echo ============================================================
echo   Die App startet jetzt im Browser.
echo   Falls sich kein Fenster oeffnet, oeffne von Hand:
echo       http://localhost:8501
echo   Zum Beenden hier Strg+C druecken oder das Fenster schliessen.
echo ============================================================
echo.

REM 5. App starten.
streamlit run app.py

pause
