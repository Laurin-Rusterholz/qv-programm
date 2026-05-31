# QV-Prüfungsassistent

Eine KI-Chat-App für die schriftliche **Geleitete Fallarbeit** (Kaufleute EFZ, HKB B/C/E).
Sie funktioniert wie dein Claude-Projekt und erstellt zusätzlich **echte Office-Dateien**
(Excel mit Formeln, Word, PowerPoint, PDF). Du kannst sie **lokal** auf deinem Computer
oder **online** als Webseite betreiben.

> **Wichtig zu Netlify:** Netlify zeigt nur reine Webseiten an und kann **kein Python**.
> Diese App ist eine Python-/Streamlit-App, darum lässt sie sich auf Netlify **nicht** öffnen.
> Nimm stattdessen **Streamlit Community Cloud** (Variante B unten) – das ist gratis und
> läuft Python nativ. Lokal funktioniert sie ohnehin (Variante A).

---

## Inhalt
- [Was die App kann](#was-die-app-kann)
- [Variante A: Lokal starten](#variante-a-lokal-starten)
- [Variante B: Online stellen (gratis, empfohlen für „von überall“)](#variante-b-online-stellen)
- [So arbeitest du damit](#so-arbeitest-du-damit)
- [Dateien, Theorie und Kosten](#dateien-theorie-und-kosten)
- [Wenn etwas nicht klappt](#wenn-etwas-nicht-klappt)
- [Datenschutz und Sicherheit](#datenschutz-und-sicherheit)

---

## Was die App kann
- Chat mit vollem Verlauf, genau nach deinem Ablauf (Ausgangslage → «Verstanden.» → Teilaufgaben).
- **Datei-Upload** direkt im Eingabefeld (📎), auch **mehrere auf einmal**: PDF und Bilder gehen direkt an die KI; Word/Excel/TXT werden als Text gelesen.
- **Datei-Erstellung**: Excel mit echten Formeln, Word (breite Tabellen im Querformat),
  PowerPoint (Folienmaster), PDF – mit automatischer Selbstkorrektur und Download-Knopf.
- **Theorie-Ordner**: wird nur einbezogen, wenn deine Nachricht das Wort **«Theorie»** enthält.
- **Modellwahl**: «Schnell (Sonnet 4.6)» oder «Beste Qualität (Opus 4.8)».
- **API-Schlüssel** im Feld eingeben (oder online dauerhaft in den Secrets hinterlegen).
- **Kosten-/Token-Anzeige** pro Antwort und pro Sitzung.
- **Passwortschutz** für den Online-Betrieb.

---

## Variante A: Lokal starten

### 1. Python installieren (nur beim ersten Mal)
- **https://www.python.org/downloads/** → **Python 3.10 oder neuer**.
- **Windows:** beim Installieren **«Add Python to PATH» anhaken**.

### 2. Starten
- **Windows:** Doppelklick auf **`start.bat`**.
- **Mac/Linux:** **`start.sh`** ausführen (im Terminal: `bash start.sh`; evtl. einmal `chmod +x start.sh`).

Beim ersten Mal installiert sich alles selbst (ein paar Minuten, braucht Internet).
Danach öffnet sich der Browser; falls nicht, von Hand **http://localhost:8501** öffnen.

### 3. Schlüssel eintragen
Links in der Seitenleiste den **API-Schlüssel** eingeben und **«Verbindung testen»** drücken.
Mit **«In .env speichern»** musst du ihn beim nächsten Mal nicht erneut eintippen.

---

## Variante B: Online stellen
(Streamlit Community Cloud – gratis, läuft Python, ideal wenn du von jedem Gerät willst.)

### Was du brauchst
- Ein **GitHub-Konto** (gratis). Dein Code liegt bereits unter `Laurin-Rusterholz/qv-programm`.
- Einen **Anthropic API-Schlüssel** mit Guthaben (console.anthropic.com).

### Schritt für Schritt
1. Gehe auf **https://share.streamlit.io** und melde dich mit **GitHub** an.
2. Klicke **«Create app»** → **«Deploy a public app from GitHub»**.
3. Wähle:
   - **Repository:** `Laurin-Rusterholz/qv-programm`
   - **Branch:** `main`
   - **Main file path:** `app.py`
   - (Unter «Advanced settings» kannst du **Python 3.11** wählen.)
4. **Secrets eintragen** (sehr empfohlen!): unter «Advanced settings» → **Secrets** das hier einfügen
   und ausfüllen (Vorlage: `.streamlit/secrets.toml.example`):
   ```toml
   APP_PASSWORT = "ein-eigenes-passwort"
   ANTHROPIC_API_KEY = "sk-ant-dein-schluessel"
   ```
   - `APP_PASSWORT` schützt die Seite, damit **nur du** sie nutzen kannst (sonst ist die URL offen!).
   - `ANTHROPIC_API_KEY` ist optional – wenn gesetzt, musst du ihn in der App nicht mehr eintippen.
5. **«Deploy»** klicken. Nach ein paar Minuten läuft die App unter einer
   `…streamlit.app`-Adresse, die du speichern (z. B. als Lesezeichen) kannst.
6. Seite öffnen → **Passwort** eingeben → falls kein Schlüssel in den Secrets steht, links den
   **API-Schlüssel** eintragen → **«Verbindung testen»** → loslegen.

> **Tipp:** Mach diesen Schritt **schon heute Abend** und drücke einmal «Verbindung testen»,
> damit morgen sicher alles läuft.

> **Hinweis:** Auf Streamlit Cloud sind hochgeladene und erstellte Dateien nur während der
> Sitzung gespeichert. Lade fertige Dateien also gleich herunter. Die Datei-Erstellung
> (Excel, Word usw.) funktioniert online genauso wie lokal.

---

## So arbeitest du damit
1. **Ausgangslage** hochladen (oft ein PDF) oder in den Chat schreiben. Die KI antwortet nur mit **«Verstanden.»**.
2. **Teilaufgabe** starten, z. B. **«T1»**, mit **Fach** (HKB B, C oder E) und der Aufgabenstellung.
3. Die KI schreibt **direkt in den Chat**. Ein Dokument erstellt sie **nur, wenn du es ausdrücklich verlangst**
   (z. B. *«Mach mir dazu ein Excel»*). Nennst du kein Programm, fragt sie kurz nach. **Nie automatisch HTML.**
4. **Rückfragen** stellen oder Anpassungen geben – die KI überarbeitet.
5. Mit **«weiter»** oder einer neuen Teilaufgabe geht es weiter.

Die KI bezieht **immer den ganzen bisherigen Chat** mit ein.

---

## Dateien, Theorie und Kosten

### Dateien anhängen (eine oder mehrere)
Direkt im Chat-Eingabefeld unten auf das **Büroklammer-Symbol 📎** klicken und **eine oder
mehrere Dateien** auswählen. Sie werden zusammen mit deinem Text gesendet (PDF/Bild gehen
direkt an die KI, Word/Excel/TXT werden als Text gelesen).

### Verwalten (Seitenleiste)
- **📚 Theorie-Dateien:** hier deine Theorie hochladen. Sie wird **nur** genutzt, wenn deine
  Nachricht das Wort **«Theorie»** enthält. Dateien lassen sich hier auch löschen.
- **📂 Erstellte Dateien:** alle erzeugten Dokumente herunterladen oder löschen.
- **📁 Hochgeladene Dateien:** Hochgeladenes ansehen oder löschen.

### Kosten
Pro Antwort siehst du eine grobe Schätzung (z. B. *„ca. US$ 0.012 · 3’400 Tokens“*), und links
die Summe der Sitzung. Die API rechnet in US-Dollar ab; **es entstehen echte Kosten je Nutzung**
(meist sehr wenig). Voraussetzung ist **Guthaben** auf console.anthropic.com.

---

## Wenn etwas nicht klappt
- **„Kein API-Schlüssel“:** Schlüssel links eintragen (oder online in den Secrets hinterlegen).
- **„Verbindung … Fehler“:** Internet prüfen. Schlüssel korrekt? **Guthaben** vorhanden?
- **Lokal: Browser zeigt nichts:** **http://localhost:8501** von Hand öffnen.
- **Online: „Falsches Passwort“:** das in den Secrets gesetzte `APP_PASSWORT` verwenden.
- **App neu starten:** lokal Fenster schliessen und Startskript erneut ausführen; online im
  Streamlit-Menü oben rechts **«Reboot app»**.
- **Netlify zeigt nichts an:** richtig so – Netlify kann kein Python. Nutze Variante B.

---

## Datenschutz und Sicherheit
- Keine eigene Datenbank, keine Telemetrie (Streamlit-Statistik ist abgeschaltet).
- Deine Texte und Dateien gehen nur an die **Anthropic-KI**, damit sie antworten kann.
- Der API-Schlüssel steht in deiner lokalen `.env` bzw. (online) in den **Secrets** – beide
  werden **nicht** ins Git aufgenommen.
- **Online unbedingt ein `APP_PASSWORT` setzen**, sonst könnte jede Person mit der URL die App
  (und damit dein API-Guthaben) nutzen.
