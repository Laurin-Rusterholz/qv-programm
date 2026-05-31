# QV-Prüfungsassistent

Eine KI-Chat-App für die schriftliche **Geleitete Fallarbeit** (Kaufleute EFZ, HKB B/C/E).
Sie funktioniert wie dein Claude-Projekt und erstellt zusätzlich **echte Office-Dateien**
(Excel mit Formeln, Word, PowerPoint, PDF).

Es gibt **zwei Ausführungen** – nimm die, die dir passt:

| Ausführung | Wofür | Speicherung |
|---|---|---|
| **🌐 Browser-App** (Ordner `web/`) | Läuft komplett im Browser, ganz ohne Server. Lokal per Doppelklick **oder** auf **Netlify**. | Schlüssel und Chat nur **in deinem Browser** (localStorage). |
| **🐍 Python-App** (Streamlit) | Läuft lokal (`start.bat`/`start.sh`) oder online auf **Streamlit Community Cloud**. | Während der Sitzung auf deinem Rechner bzw. Server. |

> **Browser-App ist die einfachste Variante, wenn du „alles im Browser, nur lokal" willst** –
> springe direkt zu [Variante C](#variante-c-browser-app).

---

## Inhalt
- [Was die App kann](#was-die-app-kann)
- [Variante C: Browser-App (alles im Browser, lokal oder Netlify)](#variante-c-browser-app)
- [Variante A: Python-App lokal starten](#variante-a-lokal-starten)
- [Variante B: Python-App online stellen (Streamlit Cloud)](#variante-b-online-stellen)
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

## Variante C: Browser-App
(Alles im Browser, ohne Server. Schlüssel und Chat bleiben **nur in deinem Browser** gespeichert.)

### Lokal öffnen (Doppelklick)
- **Windows:** Doppelklick auf **`start-web.bat`**.
- **Mac/Linux:** **`start-web.sh`** ausführen (im Terminal: `bash start-web.sh`).

Es öffnet sich der Browser unter **http://localhost:8000**. (Es wird nur ein winziger lokaler
Webserver gestartet, damit der Browser die App laden darf – es geht nichts ins Internet ausser
den Anfragen an die KI.) Dann links den **API-Schlüssel** eintragen, **«Verbindung testen»**,
loslegen. Der Schlüssel bleibt gespeichert; beim nächsten Mal ist er schon da.

### Auf Netlify legen (von überall erreichbar)
1. Auf **https://app.netlify.com** anmelden (gratis).
2. Entweder **das GitHub-Repo verbinden** – Netlify nutzt automatisch die Datei `netlify.toml`
   und veröffentlicht den Ordner `web/` – **oder** den Ordner `web/` einfach per Drag-and-drop
   auf die Netlify-Seite ziehen.
3. Seite öffnen, **API-Schlüssel eintragen** (er bleibt nur in deinem Browser), loslegen.

> **Sicherheit:** Der Schlüssel wird in deinem Browser gespeichert und ist dort technisch
> sichtbar (das ist bei reinen Browser-Apps so und für die private Nutzung mit deinem eigenen
> Schlüssel in Ordnung). Trage den Schlüssel **niemals fest in den Code** ein. Auf einer
> öffentlichen Netlify-Adresse gibt **jede Person ihren eigenen** Schlüssel ein – deiner wird
> dabei nicht weitergegeben.

> **Hinweis:** Dateien anhängen mit dem **📎-Symbol** im Eingabefeld (auch mehrere). Erstellte
> Dateien gleich herunterladen – nach einem Neuladen müssen sie ggf. neu erstellt werden (der
> Chat-Text selbst bleibt erhalten).

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
4. **Secrets eintragen** (empfohlen): unter «Advanced settings» → **Secrets** nur das Passwort
   setzen (Vorlage: `.streamlit/secrets.toml.example`):
   ```toml
   APP_PASSWORT = "ein-eigenes-passwort"
   ```
   - `APP_PASSWORT` schützt die Seite, damit **nur du** sie nutzen kannst (sonst ist die URL offen!).
   - **Den API-Schlüssel NICHT in die Secrets eintragen.** Sonst wäre er auf jedem Gerät
     aktiv/sichtbar. Du gibst ihn stattdessen **in der App pro Gerät einzeln** ein (Feld
     «API-Schlüssel»); er wird nur in dieser Sitzung gehalten und nicht geteilt.
5. **«Deploy»** klicken. Nach ein paar Minuten läuft die App unter einer
   `…streamlit.app`-Adresse, die du speichern (z. B. als Lesezeichen) kannst.
6. Seite öffnen → **Passwort** eingeben → links den **API-Schlüssel** eintragen (auf jedem Gerät
   einmal) → **«Verbindung testen»** → loslegen.

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
- Der API-Schlüssel wird **pro Gerät einzeln** eingegeben und **nicht in den Secrets** abgelegt,
  damit er nicht geteilt oder auf anderen Geräten angezeigt wird. Lokal kann er in deiner eigenen
  `.env` liegen (mit «In .env speichern»); diese wird **nicht** ins Git aufgenommen. In der
  Browser-App liegt er nur im localStorage des jeweiligen Geräts.
- **Online unbedingt ein `APP_PASSWORT` setzen**, sonst könnte jede Person mit der URL die App
  nutzen.
