# QV-Prüfungsassistent (lokale Backup-App)

Eine einfache KI-Chat-App für die schriftliche **Geleitete Fallarbeit** (Kaufleute EFZ, HKB B/C/E).
Sie funktioniert wie dein bisheriges Claude-Projekt und kann zusätzlich **echte Office-Dateien**
(Excel, Word, PowerPoint, PDF) erstellen. Alles läuft **lokal auf deinem Computer**, ohne Datenbank,
ohne Telemetrie.

---

## In aller Kürze (3 Schritte)

1. **Python installieren** (einmalig), falls noch nicht vorhanden.
2. **Startskript ausführen**: `start.bat` (Windows, Doppelklick) oder `start.sh` (Mac/Linux).
3. Im Browser **API-Schlüssel eintragen**, **«Verbindung testen»** drücken, loslegen.

> Für die KI brauchst du einen **API-Schlüssel** von **console.anthropic.com**.
> Dort ist **Guthaben** nötig, und **pro Nutzung entstehen kleine Kosten** (je nach Länge der Antworten meist wenige Rappen bis einige Franken).

---

## Schritt 1: Python installieren (nur beim ersten Mal)

- Gehe auf **https://www.python.org/downloads/** und installiere **Python 3.10 oder neuer**.
- **Windows:** Beim Installieren unten **«Add Python to PATH» anhaken**, dann «Install Now».
- **Mac:** Lade den Installer von python.org und führe ihn aus.

Falls Python schon installiert ist, kannst du diesen Schritt überspringen.

## Schritt 2: Die App starten

- **Windows:** Doppelklick auf **`start.bat`**.
- **Mac/Linux:** Doppelklick auf **`start.sh`**, oder im Terminal im Ordner: `bash start.sh`.
  - Falls Mac das Skript nicht ausführen will, einmal im Terminal: `chmod +x start.sh` eingeben.

Beim **ersten Start** installiert sich die App selbst (das dauert ein paar Minuten und braucht Internet).
Bei jedem weiteren Start geht es **sofort** los.

Danach öffnet sich automatisch dein Browser. Falls nicht, öffne von Hand:
**http://localhost:8501**

## Schritt 3: API-Schlüssel eintragen

Es gibt zwei Wege:

**Weg A (am einfachsten):** In der App links in der Seitenleiste das Feld **«API-Schlüssel»**
ausfüllen. Optional **«In .env speichern»** drücken, dann musst du ihn morgen nicht erneut eintippen.

**Weg B (vorab):** Die Datei `.env.example` zu **`.env`** kopieren und den Schlüssel eintragen:
```
ANTHROPIC_API_KEY=sk-ant-dein-schluessel
```

Zum Schluss in der App **«Verbindung testen»** drücken. Es sollte **«Verbindung erfolgreich»** erscheinen.
**Mach das am besten schon heute Abend**, damit morgen sicher alles läuft.

---

## So arbeitest du damit

1. **Ausgangslage** hochladen (oft ein PDF) oder in den Chat schreiben. Die KI antwortet nur mit **«Verstanden.»**.
2. **Teilaufgabe** starten, z. B. **«T1»** oder **«Teilaufgabe 1»**, mit **Fach** (HKB B, C oder E) und der Aufgabenstellung.
3. Die KI schreibt die Lösung **direkt in den Chat**. Ein Dokument erstellt sie **nur, wenn du es ausdrücklich verlangst**
   (z. B. *«Mach mir dazu ein Excel»*). Nennst du kein Programm, fragt sie kurz nach. Sie erstellt **nie automatisch HTML**.
4. **Rückfragen** stellen oder Anpassungen geben – die KI überarbeitet.
5. Mit **«weiter»** oder einer neuen Teilaufgabe geht es weiter.

Die KI bezieht **immer den ganzen bisherigen Chat** mit ein.

### Dateien erstellen

Verlangst du z. B. ein Excel, baut die KI es mit **echten Formeln** und legt es im Ordner **`outputs/`** ab.
Unter der Antwort erscheint ein **Download-Knopf**. Wenn der Code einmal nicht klappt, korrigiert sich die KI
automatisch (bis zu drei Versuche).

### Die drei Ordner

| Ordner       | Wofür                                                                                  |
|--------------|----------------------------------------------------------------------------------------|
| `uploads/`   | Hierhin landen deine hochgeladenen Dateien. Die KI kann daraus lesen (z. B. Vorlagen). |
| `outputs/`   | Hier speichert die KI die erstellten Dateien (Excel, Word, PowerPoint, PDF).           |
| `theorie/`   | Deine Theorie-PDFs/-Dokumente. Wird **nur** genutzt, wenn deine Nachricht das Wort **«Theorie»** enthält. |

### Modell umschalten

Links in der Seitenleiste:
- **Schnell (Sonnet 4.6)** – Standard, schnell und stark genug für die Prüfung.
- **Beste Qualität (Opus 4.8)** – beste Qualität, etwas langsamer.

---

## Wenn etwas nicht klappt

- **«Kein API-Schlüssel»:** Schlüssel links eintragen (Schritt 3).
- **«Verbindung … Fehler»:** Internet prüfen. Schlüssel korrekt? Ist **Guthaben** auf dem Konto?
- **Browser zeigt nichts:** Von Hand **http://localhost:8501** öffnen.
- **App neu starten:** Fenster schliessen und Startskript erneut ausführen.
- **Pakete-Fehler beim ersten Start:** Internet prüfen und Startskript nochmals ausführen.

---

## Datenschutz

Die App läuft komplett lokal. Deine Texte und Dateien werden nur an die **Anthropic-KI** geschickt
(damit sie antworten kann) – sonst an niemanden. Es gibt keine eigene Datenbank und keine Telemetrie.
Der API-Schlüssel steht nur in deiner lokalen `.env`-Datei und wird nicht weitergegeben.
