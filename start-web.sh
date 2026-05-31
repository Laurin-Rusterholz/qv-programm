#!/usr/bin/env bash
# Startet die Browser-App lokal (macOS/Linux). Alles bleibt in deinem Browser.
# Es wird nur ein kleiner lokaler Webserver gestartet (kein Internet ausser fuer die KI).

cd "$(dirname "$0")/web" || exit 1
PORT=8000

if command -v python3 >/dev/null 2>&1; then PY=python3
elif command -v python >/dev/null 2>&1; then PY=python
else
  echo "Python 3 wird benoetigt. Bitte von https://www.python.org installieren."
  read -r -p "Mit Enter schliessen ..." _
  exit 1
fi

echo "============================================================"
echo "  QV-Pruefungsassistent (Browser-App)"
echo "  Oeffne im Browser:  http://localhost:$PORT"
echo "  Zum Beenden hier Strg+C druecken."
echo "============================================================"

# Browser nach kurzer Wartezeit oeffnen (best effort).
( sleep 1
  if command -v xdg-open >/dev/null 2>&1; then xdg-open "http://localhost:$PORT"
  elif command -v open >/dev/null 2>&1; then open "http://localhost:$PORT"
  fi ) >/dev/null 2>&1 &

exec "$PY" -m http.server "$PORT"
