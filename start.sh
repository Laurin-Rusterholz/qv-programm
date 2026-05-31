#!/usr/bin/env bash
# Startskript fuer macOS und Linux.
# Beim ersten Mal werden eine virtuelle Umgebung angelegt und alle Pakete
# installiert. Danach startet die App einfach.

# In den Ordner dieses Skripts wechseln.
cd "$(dirname "$0")" || exit 1

echo "============================================================"
echo "  QV-Pruefungsassistent wird vorbereitet ..."
echo "============================================================"

# 1. Python suchen.
if command -v python3 >/dev/null 2>&1; then
  PY=python3
elif command -v python >/dev/null 2>&1; then
  PY=python
else
  echo ""
  echo "FEHLER: Python wurde nicht gefunden."
  echo "Bitte Python 3.10 oder neuer von https://www.python.org installieren,"
  echo "danach dieses Startskript erneut ausfuehren."
  echo ""
  read -r -p "Mit Enter schliessen ..." _
  exit 1
fi

# 2. Virtuelle Umgebung anlegen (nur beim ersten Mal).
if [ ! -d ".venv" ]; then
  echo "Erstelle virtuelle Umgebung (einmalig) ..."
  "$PY" -m venv .venv || { echo "Konnte keine virtuelle Umgebung anlegen."; read -r -p "Enter ..." _; exit 1; }
fi

# 3. Virtuelle Umgebung aktivieren.
# shellcheck disable=SC1091
source .venv/bin/activate

# 4. Pakete installieren (nur beim ersten Mal oder wenn requirements.txt neuer ist).
if [ ! -f ".venv/.installiert" ] || [ "requirements.txt" -nt ".venv/.installiert" ]; then
  echo "Installiere Pakete (einmalig, das kann ein paar Minuten dauern) ..."
  python -m pip install --upgrade pip
  if python -m pip install -r requirements.txt; then
    touch ".venv/.installiert"
  else
    echo ""
    echo "FEHLER: Die Pakete konnten nicht installiert werden (Internet pruefen)."
    read -r -p "Mit Enter schliessen ..." _
    exit 1
  fi
fi

echo ""
echo "============================================================"
echo "  Die App startet jetzt im Browser."
echo "  Falls sich kein Fenster oeffnet, oeffne von Hand:"
echo "      http://localhost:8501"
echo "  Zum Beenden hier Strg+C druecken oder das Fenster schliessen."
echo "============================================================"
echo ""

# 5. App starten.
streamlit run app.py
