"""Automatische Abnahmetests fuer den QV-Pruefungsassistenten.

Diese Tests pruefen alles, was OHNE gueltigen API-Schluessel pruefbar ist -
insbesondere die entscheidende Datei-Erstellung. Dabei wird der KI-Code durch
bekannten, repraesentativen Code ersetzt (genau die Art Code, die die KI ueber
das Werkzeug erzeugt) und das Ergebnis Datei fuer Datei geprueft.

Start:  python test_app.py
"""

import re
import sys
import base64
from pathlib import Path

import app

bestanden = 0
fehlgeschlagen = 0
TESTPRAEFIX = "_test_"


def pruefe(bedingung, beschreibung):
    global bestanden, fehlgeschlagen
    if bedingung:
        bestanden += 1
        print(f"  [OK]   {beschreibung}")
    else:
        fehlgeschlagen += 1
        print(f"  [FEHL] {beschreibung}")


def abschnitt(titel):
    print("\n" + titel)


# --------------------------------------------------------------------------
# Test 1: Konstanten, Modell, Werkzeugname
# --------------------------------------------------------------------------
def test_konstanten():
    abschnitt("Test 1: Konstanten und Konfiguration")
    pruefe(app.MODELL == "claude-sonnet-4-6", "Standardmodell ist claude-sonnet-4-6")
    pruefe(
        app.MODELL_OPTIONEN.get("Schnell (Sonnet 4.6)") == "claude-sonnet-4-6"
        and app.MODELL_OPTIONEN.get("Beste Qualitaet (Opus 4.8)") == "claude-opus-4-8",
        "Beide Modell-Optionen (Sonnet 4.6 / Opus 4.8) vorhanden",
    )
    pruefe(app.MAX_TOKENS >= 8000, f"max_tokens >= 8000 (ist {app.MAX_TOKENS})")
    pruefe(app.ANTHROPIC_VERSION == "2023-06-01", "anthropic-version 2023-06-01")
    # Werkzeugname muss gueltiges API-ASCII sein.
    pruefe(
        re.fullmatch(r"[a-zA-Z0-9_-]{1,64}", app.TOOL_NAME) is not None,
        f"Werkzeugname ist API-gueltig (ASCII): {app.TOOL_NAME}",
    )
    pruefe(app.TOOLS[0]["name"] == app.TOOL_NAME, "Werkzeugdefinition nutzt diesen Namen")
    pruefe(
        "Verstanden" in app.SYSTEM_PROMPT and "python_ausführen" in app.SYSTEM_PROMPT,
        "System-Prompt ist woertlich eingebettet",
    )


# --------------------------------------------------------------------------
# Test 2: Excel mit echten Formeln (Abnahmetest 4)
# --------------------------------------------------------------------------
EXCEL_CODE = r'''
import openpyxl
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Budget"
ws["A1"] = "Budget Sommerfest"
# Eingabewerte (anpassbare Zellen)
ws["E1"] = "MWST-Satz"; ws["F1"] = 0.081
ws["A3"] = "Position"; ws["B3"] = "Menge"; ws["C3"] = "Preis pro Stueck"; ws["D3"] = "Total"
ws["A4"] = "Getraenke"; ws["B4"] = 100; ws["C4"] = 2.5;  ws["D4"] = "=B4*C4"
ws["A5"] = "Snacks";    ws["B5"] = 50;  ws["C5"] = 4.0;  ws["D5"] = "=B5*C5"
# Ergebniszeilen als Formeln
ws["A7"] = "Total exkl. MWST"; ws["D7"] = "=SUM(D4:D5)"
ws["A8"] = "MWST 8.1 %";       ws["D8"] = "=ROUND(D7*F1,2)"
ws["A9"] = "Total inkl. MWST"; ws["D9"] = "=D7+D8"
wb.save("outputs/_test_budget.xlsx")
print("Excel gespeichert")
'''


def test_excel():
    abschnitt("Test 2: Excel-Budget mit echten Formeln (Abnahmetest 4)")
    ergebnis = app.python_ausfuehren(EXCEL_CODE)
    pruefe(ergebnis["fehler"] is None, "Code lief ohne Fehler")
    pruefe("_test_budget.xlsx" in ergebnis["neue_dateien"], "Neue Datei wird erkannt")
    pfad = app.OUTPUT_DIR / "_test_budget.xlsx"
    pruefe(pfad.exists(), "Datei _test_budget.xlsx existiert in outputs/")
    if not pfad.exists():
        return
    # Erneut OEFFNEN und Formeln pruefen (data_only=False -> Formeln statt Werte).
    import openpyxl
    wb = openpyxl.load_workbook(str(pfad), data_only=False)
    ws = wb["Budget"]
    formeln = {
        koord: zelle.value
        for zeile in ws.iter_rows()
        for koord, zelle in ((z.coordinate, z) for z in zeile)
        if isinstance(zelle.value, str) and zelle.value.startswith("=")
    }
    pruefe(len(formeln) >= 4, f"Mehrere echte Formeln vorhanden ({len(formeln)})")
    pruefe(
        ws["D4"].value == "=B4*C4",
        "Zeilen-Total verweist auf Eingabezellen (=B4*C4), kein fester Wert",
    )
    pruefe(
        isinstance(ws["D9"].value, str) and ws["D9"].value.startswith("="),
        "Total inkl. MWST ist eine Formel, kein eingetippter Wert",
    )
    pruefe("F1" in str(ws["D8"].value), "MWST-Formel verweist auf Eingabezelle F1 (MWST-Satz)")
    # Eingabezellen sind echte Zahlen, keine Formeln.
    pruefe(
        ws["B4"].value == 100 and ws["C4"].value == 2.5 and ws["F1"].value == 0.081,
        "Eingabewerte (Menge, Preis, MWST-Satz) liegen als Zahlen in eigenen Zellen",
    )
    wb.close()


# --------------------------------------------------------------------------
# Test 3: Word mit breiter Tabelle im Querformat (Abnahmetest 5)
# --------------------------------------------------------------------------
WORD_CODE = r'''
import docx
from docx.enum.section import WD_ORIENT
doc = docx.Document()
section = doc.sections[0]
# Querformat einstellen (Breite und Hoehe tauschen).
section.orientation = WD_ORIENT.LANDSCAPE
breite, hoehe = section.page_height, section.page_width
section.page_width = breite
section.page_height = hoehe
doc.add_heading("Projektplan", level=1)
spalten = ["Aufgabe", "Start", "Ende", "Verantwortlich", "Status", "Meilenstein"]
tab = doc.add_table(rows=1, cols=len(spalten))
tab.style = "Table Grid"
tab.autofit = False
for i, t in enumerate(spalten):
    tab.rows[0].cells[i].text = t
nutzbar = section.page_width - section.left_margin - section.right_margin
spaltenbreite = int(nutzbar / len(spalten))
for zeile in tab.rows:
    for zelle in zeile.cells:
        zelle.width = spaltenbreite
doc.save("outputs/_test_tabelle.docx")
print("Word gespeichert")
'''


def test_word():
    abschnitt("Test 3: Word mit breiter Tabelle im Querformat (Abnahmetest 5)")
    ergebnis = app.python_ausfuehren(WORD_CODE)
    pruefe(ergebnis["fehler"] is None, "Code lief ohne Fehler")
    pfad = app.OUTPUT_DIR / "_test_tabelle.docx"
    pruefe(pfad.exists(), "Datei _test_tabelle.docx existiert")
    if not pfad.exists():
        return
    import docx
    from docx.enum.section import WD_ORIENT
    doc = docx.Document(str(pfad))
    section = doc.sections[0]
    pruefe(section.orientation == WD_ORIENT.LANDSCAPE, "Abschnitt ist im Querformat")
    pruefe(section.page_width > section.page_height, "Seite ist breiter als hoch (Querformat)")
    tab = doc.tables[0]
    pruefe(len(tab.columns) == 6, "Tabelle hat 6 Spalten")
    nutzbar = section.page_width - section.left_margin - section.right_margin
    breiten = [zelle.width for zelle in tab.rows[0].cells if zelle.width is not None]
    pruefe(len(breiten) == 6, "Alle Spaltenbreiten sind gesetzt")
    pruefe(
        sum(breiten) <= nutzbar + 5000,  # kleine Toleranz (EMU)
        "Tabelle passt in die nutzbare Seitenbreite (keine Spalte abgeschnitten)",
    )


# --------------------------------------------------------------------------
# Test 4: PowerPoint mit Folienmaster/Layouts und Stichpunkten (Abnahmetest 6)
# --------------------------------------------------------------------------
PPTX_CODE = r'''
from pptx import Presentation
prs = Presentation()  # nutzt Vorlage mit Folienmaster und Layouts
titel_layout = prs.slide_layouts[0]
inhalt_layout = prs.slide_layouts[1]
s1 = prs.slides.add_slide(titel_layout)
s1.shapes.title.text = "Sommerfest 2026"
s2 = prs.slides.add_slide(inhalt_layout)
s2.shapes.title.text = "Aufgaben"
rahmen = s2.placeholders[1].text_frame
rahmen.text = "Technik aufbauen"
for stich in ["Catering bestellen", "Einladungen senden", "Abbau organisieren"]:
    p = rahmen.add_paragraph()
    p.text = stich
prs.save("outputs/_test_folien.pptx")
print("PowerPoint gespeichert")
'''


def test_pptx():
    abschnitt("Test 4: PowerPoint mit Folienmaster und Stichpunkten (Abnahmetest 6)")
    ergebnis = app.python_ausfuehren(PPTX_CODE)
    pruefe(ergebnis["fehler"] is None, "Code lief ohne Fehler")
    pfad = app.OUTPUT_DIR / "_test_folien.pptx"
    pruefe(pfad.exists(), "Datei _test_folien.pptx existiert")
    if not pfad.exists():
        return
    from pptx import Presentation
    prs = Presentation(str(pfad))
    pruefe(len(prs.slide_masters) >= 1, "Praesentation hat einen Folienmaster")
    # Vergleich ueber die Teilnamen (SlideLayout-Objekte sind nicht hashbar).
    master_layout_namen = {
        str(layout.part.partname)
        for master in prs.slide_masters
        for layout in master.slide_layouts
    }
    pruefe(
        all(str(folie.slide_layout.part.partname) in master_layout_namen for folie in prs.slides),
        "Alle Folien nutzen Layouts aus dem Folienmaster",
    )
    pruefe(len(prs.slides) == 2, "Zwei Folien vorhanden")
    folie2 = prs.slides[1]
    absaetze = folie2.placeholders[1].text_frame.paragraphs
    pruefe(len(absaetze) >= 4, f"Inhaltsfolie hat mehrere Stichpunkte ({len(absaetze)})")
    pruefe(
        all(len(p.text) < 60 for p in absaetze if p.text),
        "Stichpunkte sind kurz (keine Fliesstexte)",
    )


# --------------------------------------------------------------------------
# Test 5: PDF erstellen (reportlab)
# --------------------------------------------------------------------------
PDF_CODE = r'''
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
c = canvas.Canvas("outputs/_test_dokument.pdf", pagesize=A4)
c.setFont("Helvetica-Bold", 16)
c.drawString(72, 770, "Infoblatt Sommerfest")
c.setFont("Helvetica", 11)
c.drawString(72, 740, "Datum: 1. August 2026")
c.save()
print("PDF gespeichert")
'''


def test_pdf():
    abschnitt("Test 5: PDF erstellen (reportlab)")
    ergebnis = app.python_ausfuehren(PDF_CODE)
    pruefe(ergebnis["fehler"] is None, "Code lief ohne Fehler")
    pfad = app.OUTPUT_DIR / "_test_dokument.pdf"
    pruefe(pfad.exists(), "Datei _test_dokument.pdf existiert")
    if pfad.exists():
        pruefe(pfad.read_bytes()[:5] == b"%PDF-", "Datei ist eine echte PDF (beginnt mit %PDF-)")


# --------------------------------------------------------------------------
# Test 6: Fehlerbehandlung und Selbstkorrektur-Eingang (Abnahmetest robust)
# --------------------------------------------------------------------------
def test_fehler():
    abschnitt("Test 6: Fehlerhafter Code stuerzt die App nicht ab")
    ergebnis = app.python_ausfuehren("x = 1 / 0\n")
    pruefe(ergebnis["fehler"] is not None, "Fehler wird gemeldet (kein Absturz)")
    pruefe(
        "ZeroDivisionError" in (ergebnis["fehler"] or ""),
        "Vollstaendige Fehlermeldung wird zurueckgegeben",
    )
    pruefe(ergebnis["neue_dateien"] == [], "Bei Fehler werden keine Dateien als erstellt gemeldet")

    # Zeitueberschreitung wird sauber abgefangen.
    ergebnis2 = app.python_ausfuehren("import time; time.sleep(5)\n", timeout=1)
    pruefe(
        ergebnis2["fehler"] is not None and "Zeit" in ergebnis2["fehler"],
        "Zeitueberschreitung wird sauber gemeldet",
    )


# --------------------------------------------------------------------------
# Test 7: Theorie wird nur bei Bedarf geladen (Abnahmetest 8)
# --------------------------------------------------------------------------
def test_theorie():
    abschnitt("Test 7: Theorie-Ordner und Ausloesewort (Abnahmetest 8)")
    testdatei = app.THEORIE_DIR / "_test_theorie.txt"
    testdatei.write_text(
        "Gewaltfreie Kommunikation: Wahrnehmung, Gefuehl, Beduerfnis, Bitte.",
        encoding="utf-8",
    )
    try:
        text = app.lade_theorie()
        pruefe("Gewaltfreie Kommunikation" in text, "lade_theorie() liest Dateien aus theorie/")
        # Ausloese-Logik (so wie in behandle_eingabe verwendet).
        pruefe("theorie" in "Loese T2 mit der Theorie".lower(), "Wort «Theorie» wird erkannt")
        pruefe("theorie" not in "Loese bitte T2".lower(), "Ohne «Theorie» kein Ausloeser")
        # baue_user_inhalt haengt Theorie-Block nur an, wenn Text uebergeben wird.
        mit = app.baue_user_inhalt("Frage", [], text)
        ohne = app.baue_user_inhalt("Frage", [], "")
        pruefe(
            any("THEORIE-KONTEXT" in b.get("text", "") for b in mit),
            "Theorie-Kontext wird als Block mitgegeben, wenn vorhanden",
        )
        pruefe(
            not any("THEORIE-KONTEXT" in b.get("text", "") for b in ohne),
            "Ohne Theorie kein Theorie-Block",
        )
    finally:
        testdatei.unlink(missing_ok=True)


# --------------------------------------------------------------------------
# Test 8: Datei-Inhalte extrahieren (Word/Excel/TXT lesen)
# --------------------------------------------------------------------------
def test_extraktion():
    abschnitt("Test 8: Hochgeladene Dateien als Text lesen")
    # Word erzeugen und wieder lesen.
    app.python_ausfuehren(
        'import docx\nd=docx.Document()\nd.add_paragraph("Geheimwort Alpha")\n'
        'd.save("uploads/_test_lesen.docx")\n'
    )
    text_docx = app.extrahiere_datei_text(app.UPLOAD_DIR / "_test_lesen.docx")
    pruefe(text_docx and "Geheimwort Alpha" in text_docx, "Word-Datei wird korrekt gelesen")

    # Excel erzeugen und wieder lesen.
    app.python_ausfuehren(
        'import openpyxl\nwb=openpyxl.Workbook()\nwb.active["A1"]="Geheimwort Beta"\n'
        'wb.save("uploads/_test_lesen.xlsx")\n'
    )
    text_xlsx = app.extrahiere_datei_text(app.UPLOAD_DIR / "_test_lesen.xlsx")
    pruefe(text_xlsx and "Geheimwort Beta" in text_xlsx, "Excel-Datei wird korrekt gelesen")

    # TXT.
    (app.UPLOAD_DIR / "_test_lesen.txt").write_text("Geheimwort Gamma", encoding="utf-8")
    text_txt = app.extrahiere_datei_text(app.UPLOAD_DIR / "_test_lesen.txt")
    pruefe(text_txt and "Geheimwort Gamma" in text_txt, "Textdatei wird korrekt gelesen")


# --------------------------------------------------------------------------
# Test 9: PDF und Bild werden nativ als Inhalt uebergeben (Abnahmetest 7)
# --------------------------------------------------------------------------
# 1x1-PNG (transparent), gueltige PNG-Bytes als Base64.
PNG_1x1 = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9"
    "awAAAABJRU5ErkJggg=="
)


def test_native_eingabe():
    abschnitt("Test 9: PDF und Bild als native KI-Eingabe (Abnahmetest 7)")
    # PDF erzeugen.
    app.python_ausfuehren(
        'from reportlab.pdfgen import canvas\n'
        'c=canvas.Canvas("uploads/_test_eingabe.pdf")\nc.drawString(72,800,"Ausgangslage")\nc.save()\n'
    )
    pdf_pfad = app.UPLOAD_DIR / "_test_eingabe.pdf"
    bloecke = app.datei_zu_content_bloecke(pdf_pfad)
    doc_block = next((b for b in bloecke if b.get("type") == "document"), None)
    pruefe(doc_block is not None, "PDF wird als 'document'-Block uebergeben (nicht als Text)")
    pruefe(
        doc_block and doc_block["source"]["media_type"] == "application/pdf",
        "PDF-Block hat media_type application/pdf",
    )
    pruefe(
        doc_block and isinstance(doc_block["source"]["data"], str) and len(doc_block["source"]["data"]) > 0,
        "PDF-Block enthaelt Base64-Daten",
    )

    # Bild erzeugen.
    bild_pfad = app.UPLOAD_DIR / "_test_bild.png"
    bild_pfad.write_bytes(base64.b64decode(PNG_1x1))
    bild_bloecke = app.datei_zu_content_bloecke(bild_pfad)
    img_block = next((b for b in bild_bloecke if b.get("type") == "image"), None)
    pruefe(img_block is not None, "Bild wird als 'image'-Block uebergeben")
    pruefe(
        img_block and img_block["source"]["media_type"] == "image/png",
        "Bild-Block hat media_type image/png",
    )


# --------------------------------------------------------------------------
# Test 10: kleine Hilfsfunktionen
# --------------------------------------------------------------------------
def test_hilfsfunktionen():
    abschnitt("Test 10: Hilfsfunktionen")
    pruefe(app.sicherer_dateiname("../../etc/passwd") == "passwd", "Pfad-Trickserei wird entfernt")
    pruefe(app.sicherer_dateiname("Budget 2026.xlsx") == "Budget 2026.xlsx", "Normaler Name bleibt")
    pruefe(
        app.mime_fuer("x.xlsx").endswith("spreadsheetml.sheet"),
        "MIME-Typ fuer .xlsx korrekt",
    )
    pruefe(app.mime_fuer("x.pdf") == "application/pdf", "MIME-Typ fuer .pdf korrekt")
    # block_zu_dict mit einem einfachen Objekt.
    pruefe(app.text_aus_bloecken([]) == "", "text_aus_bloecken mit leerer Liste")
    # Kostenschaetzung (Sonnet: 3 USD Eingabe + 15 USD Ausgabe pro 1 Mio).
    pruefe(
        abs(app.geschaetzte_kosten("claude-sonnet-4-6", 1_000_000, 1_000_000) - 18.0) < 1e-6,
        "Kostenschaetzung Sonnet korrekt (3+15 = 18 USD pro 1 Mio)",
    )
    pruefe(
        abs(app.geschaetzte_kosten("claude-opus-4-8", 1_000_000, 0) - 5.0) < 1e-6,
        "Kostenschaetzung Opus korrekt (5 USD Eingabe pro 1 Mio)",
    )


# --------------------------------------------------------------------------
# Test 11 & 12: Werkzeug-Schleife (Tool-Use) mit nachgebautem Client
# --------------------------------------------------------------------------
class FakeText:
    type = "text"

    def __init__(self, text):
        self.text = text

    def model_dump(self, **_):
        return {"type": "text", "text": self.text}


class FakeToolUse:
    type = "tool_use"

    def __init__(self, id, name, eingabe):
        self.id = id
        self.name = name
        self.input = eingabe

    def model_dump(self, **_):
        return {"type": "tool_use", "id": self.id, "name": self.name, "input": self.input}


class FakeUsage:
    def __init__(self, input_tokens=0, output_tokens=0):
        self.input_tokens = input_tokens
        self.output_tokens = output_tokens
        self.cache_read_input_tokens = 0
        self.cache_creation_input_tokens = 0


class FakeResponse:
    def __init__(self, content, stop_reason, usage=None):
        self.content = content
        self.stop_reason = stop_reason
        self.usage = usage


class FakeMessages:
    def __init__(self, antworten):
        self._antworten = list(antworten)
        self.calls = []

    def create(self, **kwargs):
        self.calls.append(kwargs)
        return self._antworten.pop(0)


class FakeClient:
    def __init__(self, antworten):
        self.messages = FakeMessages(antworten)


def test_werkzeugschleife():
    abschnitt("Test 11: Werkzeug-Schleife erstellt Datei und fuehrt Verlauf (mit Mock)")
    code = (
        'import openpyxl\nwb=openpyxl.Workbook()\nwb.active["A1"]="x"\n'
        'wb.save("outputs/_test_mockbudget.xlsx")\nprint("ok")\n'
    )
    antworten = [
        FakeResponse(
            [FakeText("Ich erstelle die Datei."), FakeToolUse("t1", app.TOOL_NAME, {"code": code})],
            "tool_use",
            usage=FakeUsage(1000, 200),
        ),
        FakeResponse([FakeText("Die Datei ist bereit.")], "end_turn", usage=FakeUsage(1200, 50)),
    ]
    client = FakeClient(antworten)
    messages = [{"role": "user", "content": "Mach mir ein Excel"}]
    text, dateien, verbrauch = app.fuehre_konversation(client, "claude-sonnet-4-6", messages)

    pruefe(text == "Die Datei ist bereit.", "Schlusstext der KI wird zurueckgegeben")
    pruefe("_test_mockbudget.xlsx" in dateien, "Erstellte Datei wird gemeldet (Download moeglich)")
    pruefe(
        verbrauch["input"] == 2200 and verbrauch["output"] == 250,
        f"Token-Verbrauch wird ueber alle Anfragen summiert ({verbrauch})",
    )
    pruefe(len(messages) == 4, "Verlauf enthaelt user, assistant, tool_result, assistant")
    pruefe(messages[1]["role"] == "assistant", "Assistenten-Antwort im Verlauf")
    pruefe(
        messages[2]["role"] == "user"
        and messages[2]["content"][0]["type"] == "tool_result"
        and messages[2]["content"][0]["is_error"] is False,
        "Werkzeug-Ergebnis (Erfolg) wird in den Verlauf gelegt",
    )
    pruefe(
        client.messages.calls[0]["tools"] == app.TOOLS
        and isinstance(client.messages.calls[0]["system"], list)
        and client.messages.calls[0]["system"][0].get("cache_control") is not None,
        "System-Prompt (mit Caching) und Werkzeug werden bei jeder Anfrage gesendet",
    )


def test_selbstkorrektur():
    abschnitt("Test 12: Selbstkorrektur nach fehlerhaftem Code (mit Mock)")
    schlechter_code = 'raise ValueError("kaputt")\n'
    guter_code = (
        'import openpyxl\nwb=openpyxl.Workbook()\nwb.active["A1"]="ok"\n'
        'wb.save("outputs/_test_korrigiert.xlsx")\n'
    )
    antworten = [
        FakeResponse([FakeToolUse("t1", app.TOOL_NAME, {"code": schlechter_code})], "tool_use"),
        FakeResponse([FakeToolUse("t2", app.TOOL_NAME, {"code": guter_code})], "tool_use"),
        FakeResponse([FakeText("Jetzt ist die Datei bereit.")], "end_turn"),
    ]
    client = FakeClient(antworten)
    messages = [{"role": "user", "content": "Mach ein Excel"}]
    text, dateien, verbrauch = app.fuehre_konversation(client, "claude-sonnet-4-6", messages)

    pruefe("_test_korrigiert.xlsx" in dateien, "Nach Korrektur wird die Datei erstellt")
    pruefe(text == "Jetzt ist die Datei bereit.", "KI meldet Erfolg nach Korrektur")
    # Im Verlauf muss ein Fehler-Werkzeugergebnis stehen (is_error True).
    fehler_ergebnisse = [
        block
        for nachricht in messages
        if nachricht["role"] == "user" and isinstance(nachricht["content"], list)
        for block in nachricht["content"]
        if isinstance(block, dict) and block.get("type") == "tool_result" and block.get("is_error")
    ]
    pruefe(len(fehler_ergebnisse) == 1, "Fehler wurde der KI zur Korrektur zurueckgemeldet")
    pruefe(
        "FEHLER" in fehler_ergebnisse[0]["content"],
        "Fehlermeldung enthaelt den vollstaendigen Fehlertext",
    )


class FakeUpload:
    """Ahmt eine von Streamlit hochgeladene Datei nach (.name, .getvalue())."""

    def __init__(self, name, daten):
        self.name = name
        self._daten = daten

    def getvalue(self):
        return self._daten


class FakeChatEingabe:
    def __init__(self, text, files):
        self.text = text
        self.files = files


def test_mehrere_dateien():
    abschnitt("Test 13: Mehrere Dateien in EINER Nachricht (Upload im Chat-Feld)")
    # Chat-Eingabe entpacken: reiner Text und Text+Dateien.
    pruefe(app.entpacke_chat_eingabe("Hallo") == ("Hallo", []), "Reiner Text wird erkannt")
    text, dateien = app.entpacke_chat_eingabe(
        FakeChatEingabe("Frage", [FakeUpload("a.txt", b"x"), FakeUpload("b.txt", b"y")])
    )
    pruefe(text == "Frage" and len(dateien) == 2, "Text + mehrere Dateien werden entpackt")

    # Mehrere Uploads werden gespeichert.
    pfade = app.speichere_uploads(
        [FakeUpload("_test_a.txt", b"AAA"), FakeUpload("_test_b.txt", b"BBB")]
    )
    pruefe(len(pfade) == 2 and all(p.exists() for p in pfade), "Mehrere Uploads gespeichert")

    # baue_user_inhalt mit ZWEI Dateien (PDF + Excel) -> beide enthalten.
    app.python_ausfuehren(
        'from reportlab.pdfgen import canvas\nc=canvas.Canvas("uploads/_test_m.pdf")\n'
        'c.drawString(72,800,"PDFINHALT")\nc.save()\n'
    )
    app.python_ausfuehren(
        'import openpyxl\nwb=openpyxl.Workbook()\nwb.active["A1"]="XLSXINHALT"\n'
        'wb.save("uploads/_test_m.xlsx")\n'
    )
    inhalt = app.baue_user_inhalt(
        "Vergleiche beide Dateien",
        [app.UPLOAD_DIR / "_test_m.pdf", app.UPLOAD_DIR / "_test_m.xlsx"],
        "",
    )
    typen = [b.get("type") for b in inhalt]
    pruefe("document" in typen, "PDF ist als nativer document-Block enthalten")
    pruefe(
        any(b.get("type") == "text" and "XLSXINHALT" in b.get("text", "") for b in inhalt),
        "Excel-Inhalt ist als Text enthalten",
    )
    pruefe(
        any(b.get("type") == "text" and "Vergleiche beide" in b.get("text", "") for b in inhalt),
        "Beide Dateien UND der Nutzertext sind in einer Nachricht (mehrere auf einmal)",
    )

    # Nur Dateien, kein Text: Datei-Block vorhanden, kein leerer Textblock.
    nur_datei = app.baue_user_inhalt("", [app.UPLOAD_DIR / "_test_m.pdf"], "")
    pruefe(
        any(b.get("type") == "document" for b in nur_datei),
        "Nachricht ohne Text, nur Datei, enthaelt die Datei",
    )
    pruefe(
        all(b.get("text", "") != "" for b in nur_datei if b.get("type") == "text"),
        "Kein leerer Textblock bei leerem Nachrichtentext",
    )


def aufraeumen():
    """Loescht alle Testdateien (_test_*) aus den Arbeitsordnern."""
    for ordner in (app.OUTPUT_DIR, app.UPLOAD_DIR, app.THEORIE_DIR):
        for pfad in ordner.glob(f"{TESTPRAEFIX}*"):
            try:
                pfad.unlink()
            except OSError:
                pass


def main():
    print("=" * 64)
    print(" Abnahmetests QV-Pruefungsassistent")
    print("=" * 64)
    try:
        test_konstanten()
        test_excel()
        test_word()
        test_pptx()
        test_pdf()
        test_fehler()
        test_theorie()
        test_extraktion()
        test_native_eingabe()
        test_hilfsfunktionen()
        test_werkzeugschleife()
        test_selbstkorrektur()
        test_mehrere_dateien()
    finally:
        aufraeumen()

    print("\n" + "=" * 64)
    print(f" Ergebnis: {bestanden} bestanden, {fehlgeschlagen} fehlgeschlagen")
    print("=" * 64)
    return 0 if fehlgeschlagen == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
