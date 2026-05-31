# Browser-App (Ordner `web/`)

Reine Browser-Variante des QV-Pruefungsassistenten: laeuft ohne Server, speichert
API-Schluessel und Chat nur im Browser (localStorage) und ruft die Anthropic-API direkt
aus dem Browser auf (Header `anthropic-dangerous-direct-browser-access: true`).

## Aufbau
- `index.html`, `styles.css` – Oberflaeche
- `core.js` – reine Logik (System-Prompt, Werkzeug, Datei-Ausfuehrung, Excel-Pruefung). Laeuft im Browser und in Node.
- `app.js` – Browser-Glue (DOM, fetch, localStorage, Datei-Upload, Downloads)
- `vendor/` – lokal eingebundene Bibliotheken: SheetJS (Excel), docx (Word), PptxGenJS (PowerPoint), jsPDF (PDF), mammoth (Word lesen). Kein CDN noetig.

## Datei-Erstellung
Die KI ruft das Werkzeug `javascript_ausfuehren` auf und schreibt JavaScript, das mit den
Bibliotheken eine Datei baut und ueber `speichern(name, daten)` bereitstellt. Erstellte
`.xlsx` werden automatisch auf deutsche Funktionsnamen/Semikolon geprueft (Ursache fuer
`#NAME?`); bei einem Fund korrigiert die KI die Datei selbst.

## Lokal starten
Im Projekt-Hauptordner `start-web.bat` (Windows) bzw. `start-web.sh` (Mac/Linux) ausfuehren,
oder manuell: `cd web && python3 -m http.server 8000`, dann `http://localhost:8000` oeffnen.

## Tests (nur fuer die Entwicklung)
```bash
cd web
npm install
node tests/core.test.mjs        # reine Logik + Datei-Erstellung (Node)
npx playwright install chromium # einmalig, fuer den Browser-Test
node tests/browser.spec.mjs     # echter Browser-Test mit gemockter API
```
Die App selbst braucht **kein** node_modules – die Bibliotheken liegen in `vendor/`.
