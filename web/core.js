// core.js - Reine Logik des QV-Pruefungsassistenten (Browser-Variante).
// Laeuft im Browser UND in Node (fuer Tests). Bibliotheken (XLSX, docx, ...)
// werden als Parameter uebergeben, damit dieses Modul keine Abhaengigkeiten hat.

export const ANTHROPIC_VERSION = "2023-06-01";
export const MAX_TOKENS = 8000;
export const MAX_SCHLEIFE = 12;
export const MAX_FEHLER_VERSUCHE = 3;

export const MODELL_OPTIONEN = {
  "Schnell (Sonnet 4.6)": "claude-sonnet-4-6",
  "Beste Qualitaet (Opus 4.8)": "claude-opus-4-8",
};
export const STANDARD_MODELL = "claude-sonnet-4-6";

// Preise in US-Dollar pro 1 Mio Tokens (Eingabe, Ausgabe).
export const PREISE_PRO_MIO = {
  "claude-sonnet-4-6": [3.0, 15.0],
  "claude-opus-4-8": [5.0, 25.0],
};

export const TOOL_NAME = "javascript_ausfuehren";

export const TOOLS = [
  {
    name: TOOL_NAME,
    description:
      "Erstellt pruefungstaugliche Dateien, indem du JavaScript schreibst, das im " +
      "Browser ausgefuehrt wird. Als Parameter stehen bereit: XLSX (SheetJS, Excel), " +
      "docx (Word), PptxGenJS (PowerPoint), jspdf (PDF; Klasse jspdf.jsPDF), " +
      "uploads (Objekt Dateiname->Uint8Array der hochgeladenen Dateien) und " +
      "speichern(dateiname, daten) zum Bereitstellen der fertigen Datei (daten als " +
      "ArrayBuffer, Uint8Array oder Blob). Du kannst await verwenden. " +
      "Excel: const wb=XLSX.utils.book_new(); const ws=XLSX.utils.aoa_to_sheet([[\"Menge\",\"Preis\",\"Total\"],[10,2.5,{t:\"n\",f:\"A2*B2\",v:25}]]); XLSX.utils.book_append_sheet(wb,ws,\"Budget\"); speichern(\"Budget.xlsx\", wb); (uebergib das Workbook DIREKT an speichern, nicht XLSX.write) " +
      "Word: const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell}=docx; const doc=new Document({sections:[{children:[new Paragraph(\"Text\")]}]}); speichern(\"Brief.docx\", await Packer.toBlob(doc)); " +
      "PowerPoint: const p=new PptxGenJS(); const s=p.addSlide(); s.addText(\"Titel\",{x:1,y:1}); speichern(\"Folien.pptx\", await p.write({outputType:\"arraybuffer\"})); " +
      "PDF: const {jsPDF}=jspdf; const doc=new jsPDF(); doc.text(\"Hallo\",10,10); speichern(\"Info.pdf\", doc.output(\"arraybuffer\")); " +
      "WICHTIG (Excel/SheetJS): Formeln im Feld f IMMER mit englischen Funktionsnamen " +
      "(SUM, SUMPRODUCT, COUNTIF, COUNTIFS, SUMIF, ROUND, IF, AVERAGE, VLOOKUP) und " +
      "Komma als Trennzeichen schreiben, ohne fuehrendes '='. Niemals deutsche Namen " +
      "(SUMME, SUMMENPRODUKT, ZAEHLENWENN) oder Semikolon, sonst zeigt Excel #NAME?. " +
      "Jede Formelzelle braucht zusaetzlich einen berechneten Wert v, z. B. " +
      "{t:'n',f:'SUM(D2:D3)',v:75}, sonst geht die Formel verloren. Uebergib das " +
      "SheetJS-Workbook direkt an speichern(name, wb). " +
      "Eingabewerte (Preise, Mengen, Saetze) als Zahlen in eigene Zellen, Formeln " +
      "verweisen darauf. Berechne die Resultate zusaetzlich in JavaScript und nenne im " +
      "Begleittext genau diese Zahlen. Diagramme/Tabellen vollstaendig beschriften.",
    input_schema: {
      type: "object",
      properties: {
        code: {
          type: "string",
          description:
            "Der auszufuehrende JavaScript-Code. Stellt die Datei(en) ueber speichern(name, daten) bereit.",
        },
      },
      required: ["code"],
    },
  },
];

// Deutsche Excel-Funktionsnamen, die zu #NAME? fuehren (-> englische Form).
export const DEUTSCHE_EXCEL_FUNKTIONEN = {
  SUMMENPRODUKT: "SUMPRODUCT",
  SUMMEWENNS: "SUMIFS",
  SUMMEWENN: "SUMIF",
  SUMME: "SUM",
  "ZÄHLENWENNS": "COUNTIFS",
  "ZÄHLENWENN": "COUNTIF",
  ZAEHLENWENNS: "COUNTIFS",
  ZAEHLENWENN: "COUNTIF",
  ANZAHL2: "COUNTA",
  ANZAHL: "COUNT",
  MITTELWERT: "AVERAGE",
  RUNDEN: "ROUND",
  WENNFEHLER: "IFERROR",
  WENN: "IF",
  SVERWEIS: "VLOOKUP",
  WVERWEIS: "HLOOKUP",
  HEUTE: "TODAY",
  JETZT: "NOW",
};

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// Prueft eine Liste von Formeln ({ref, f}) auf deutsche Namen/Semikolon.
export function findeFormelProbleme(formeln) {
  const probleme = [];
  for (const { ref, f } of formeln) {
    const text = String(f);
    const gross = text.toUpperCase();
    for (const [deutsch, englisch] of Object.entries(DEUTSCHE_EXCEL_FUNKTIONEN)) {
      const re = new RegExp("(?<![A-ZÄÖÜ0-9_.])" + escapeRegExp(deutsch) + "\\s*\\(");
      if (re.test(gross)) {
        probleme.push(`${ref}: deutsche Funktion '${deutsch}' -> bitte '${englisch}' verwenden`);
        break;
      }
    }
    if (text.includes(";")) {
      probleme.push(`${ref}: Semikolon in Formel -> Komma als Trennzeichen verwenden`);
    }
  }
  return [...new Set(probleme)].slice(0, 20);
}

// Liest die Formeln aus einer .xlsx (Bytes) mit SheetJS.
export function extrahiereXlsxFormeln(XLSX, bytes) {
  const wb = XLSX.read(bytes, { type: "array" });
  const out = [];
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    for (const ref of Object.keys(ws)) {
      if (ref[0] === "!") continue;
      const zelle = ws[ref];
      if (zelle && zelle.f) out.push({ ref: `${name}!${ref}`, f: zelle.f });
    }
  }
  return out;
}

// Komfort: prueft .xlsx-Bytes direkt.
export function pruefeXlsxBytes(XLSX, bytes) {
  try {
    return findeFormelProbleme(extrahiereXlsxFormeln(XLSX, bytes));
  } catch (e) {
    return [];
  }
}

export function geschaetzteKosten(modell, inputTokens, outputTokens) {
  const [ein, aus] = PREISE_PRO_MIO[modell] || [3.0, 15.0];
  return (inputTokens / 1e6) * ein + (outputTokens / 1e6) * aus;
}

// --- Bytes/Base64-Helfer (Browser + Node) ---

export function bytesZuBase64(bytes) {
  let bin = "";
  const u = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const chunk = 0x8000;
  for (let i = 0; i < u.length; i += chunk) {
    bin += String.fromCharCode.apply(null, u.subarray(i, i + chunk));
  }
  if (typeof btoa !== "undefined") return btoa(bin);
  return Buffer.from(u).toString("base64");
}

export function base64ZuBytes(b64) {
  const clean = String(b64).replace(/^data:[^,]+,/, "");
  if (typeof atob !== "undefined") {
    const bin = atob(clean);
    const u = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    return u;
  }
  return new Uint8Array(Buffer.from(clean, "base64"));
}

export async function normalisiereBytes(data) {
  if (data == null) return new Uint8Array();
  if (typeof Blob !== "undefined" && data instanceof Blob) {
    return new Uint8Array(await data.arrayBuffer());
  }
  if (data instanceof Uint8Array) return data;
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  if (ArrayBuffer.isView(data)) return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  if (Array.isArray(data)) return Uint8Array.from(data);
  if (typeof data === "string") {
    try {
      return base64ZuBytes(data);
    } catch (_) {
      return new TextEncoder().encode(data);
    }
  }
  return new Uint8Array();
}

// Fuehrt KI-JavaScript aus. libs = {XLSX, docx, PptxGenJS, jspdf}. uploads = {name: Uint8Array}.
// Rueckgabe: {dateien: {name: Uint8Array}, output: string}.
export async function fuehreJsAus(code, libs, uploads) {
  const roh = {};
  const logs = [];
  function speichern(name, daten) {
    roh[String(name)] = daten;
  }
  function print(...args) {
    logs.push(args.map((x) => (typeof x === "string" ? x : JSON.stringify(x))).join(" "));
  }
  const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
  const fn = new AsyncFunction(
    "XLSX",
    "docx",
    "PptxGenJS",
    "jspdf",
    "uploads",
    "speichern",
    "print",
    code
  );
  const origLog = console.log;
  console.log = (...a) => print(...a);
  try {
    await fn(libs.XLSX, libs.docx, libs.PptxGenJS, libs.jspdf, uploads || {}, speichern, print);
  } finally {
    console.log = origLog;
  }
  const dateien = {};
  for (const [name, daten] of Object.entries(roh)) {
    if (istWorkbook(daten)) {
      // SheetJS-Workbook: fehlende Formelwerte ergaenzen (sonst Formelverlust) und schreiben.
      dateien[name] = workbookZuBytes(libs.XLSX, daten);
    } else {
      dateien[name] = await normalisiereBytes(daten);
    }
  }
  return { dateien, output: logs.join("\n") };
}

export function istWorkbook(daten) {
  return !!(daten && typeof daten === "object" && Array.isArray(daten.SheetNames) && daten.Sheets);
}

// Schreibt ein SheetJS-Workbook in Bytes. Ergaenzt fuer jede Formelzelle ohne Wert
// einen Platzhalterwert (sonst verwirft SheetJS die Formel) und erzwingt Neuberechnung.
export function workbookZuBytes(XLSX, wb) {
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    if (!ws) continue;
    for (const ref of Object.keys(ws)) {
      if (ref[0] === "!") continue;
      const c = ws[ref];
      if (c && c.f != null) {
        if (c.v == null) c.v = 0;
        if (!c.t) c.t = "n";
      }
    }
  }
  wb.Workbook = wb.Workbook || {};
  wb.Workbook.CalcPr = Object.assign({}, wb.Workbook.CalcPr, { fullCalcOnLoad: true });
  const out = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  return out instanceof Uint8Array ? out : new Uint8Array(out);
}

// Sammelt den reinen Text aus einer Antwort-Inhaltsliste (plain JSON).
export function textAusContent(content) {
  if (!Array.isArray(content)) return "";
  return content
    .filter((b) => b && b.type === "text")
    .map((b) => b.text)
    .join("\n")
    .trim();
}

// Baut den Inhalt einer Nutzer-Nachricht. files = Liste von Deskriptoren:
//   {kind:"pdf", name, base64} | {kind:"image", name, mediaType, base64} | {kind:"text", name, text}
export function baueUserContent(text, files, theorieText) {
  const inhalt = [];
  if (theorieText) {
    inhalt.push({
      type: "text",
      text:
        "[THEORIE-KONTEXT - nur mitgegeben, weil die Nachricht das Wort «Theorie» " +
        "enthaelt. Nutze ihn fuer Methode und Inhalt, nie zur Begruendung eines " +
        "Dateiformats.]\n\n" + theorieText,
    });
  }
  for (const f of files || []) {
    if (f.kind === "pdf") {
      inhalt.push({ type: "text", text: `[Hochgeladene Datei: ${f.name}]` });
      inhalt.push({
        type: "document",
        source: { type: "base64", media_type: "application/pdf", data: f.base64 },
      });
    } else if (f.kind === "image") {
      inhalt.push({ type: "text", text: `[Hochgeladenes Bild: ${f.name}]` });
      inhalt.push({
        type: "image",
        source: { type: "base64", media_type: f.mediaType, data: f.base64 },
      });
    } else if (f.kind === "text") {
      inhalt.push({
        type: "text",
        text: `[Inhalt der hochgeladenen Datei ${f.name}]:\n${f.text}`,
      });
    }
  }
  if (text && text.trim()) inhalt.push({ type: "text", text });
  if (inhalt.length === 0) inhalt.push({ type: "text", text: "(leere Nachricht)" });
  return inhalt;
}

// Orchestriert die Werkzeug-Schleife (inkl. Selbstkorrektur und Excel-Pruefung).
// deps = { callClaude(model, messages, system, tools)->antwort,
//          runCode(code)->{error, output, dateien:{name:Uint8Array}},
//          lintXlsx(bytes)->string[] }
// Rueckgabe: { text, dateien:{name:Uint8Array}, verbrauch:{input,output} }.
export async function fuehreKonversation(deps, modell, messages, onStatus) {
  const { callClaude, runCode, lintXlsx } = deps;
  const verbrauch = { input: 0, output: 0 };
  const neueDateien = {};
  let fehlerAnzahl = 0;
  let letzterText = "";
  const system = [
    { type: "text", text: SYSTEM_PROMPT, cache_control: { type: "ephemeral" } },
  ];
  const status = (t) => {
    if (onStatus) try { onStatus(t); } catch (_) {}
  };
  const zaehle = (a) => {
    const u = a && a.usage;
    if (!u) return;
    verbrauch.input += (u.input_tokens || 0) + (u.cache_read_input_tokens || 0) + (u.cache_creation_input_tokens || 0);
    verbrauch.output += u.output_tokens || 0;
  };

  for (let i = 0; i < MAX_SCHLEIFE; i++) {
    status("Die KI denkt nach ...");
    const antwort = await callClaude(modell, messages, system, TOOLS);
    zaehle(antwort);
    messages.push({ role: "assistant", content: antwort.content });
    const t = textAusContent(antwort.content);
    if (t) letzterText = t;

    if (antwort.stop_reason !== "tool_use") {
      return { text: letzterText, dateien: neueDateien, verbrauch };
    }

    const toolResults = [];
    for (const block of antwort.content) {
      if (!block || block.type !== "tool_use") continue;
      if (block.name !== TOOL_NAME) {
        toolResults.push({ type: "tool_result", tool_use_id: block.id, content: "Unbekanntes Werkzeug.", is_error: true });
        continue;
      }
      const code = (block.input && block.input.code) || "";
      status("Die KI erstellt eine Datei ...");
      let res;
      try {
        res = await runCode(code);
      } catch (e) {
        res = { error: String((e && e.stack) || e), output: "", dateien: {} };
      }

      if (res.error) {
        fehlerAnzahl++;
        let inhalt = "FEHLER bei der Ausfuehrung:\n" + res.error;
        inhalt += fehlerAnzahl >= MAX_FEHLER_VERSUCHE
          ? "\n\nDies war der letzte Korrekturversuch. Erklaere mir kurz in einem Satz, woran es liegt."
          : "\n\nBitte korrigiere den Code und rufe das Werkzeug erneut auf.";
        toolResults.push({ type: "tool_result", tool_use_id: block.id, content: inhalt, is_error: true });
        continue;
      }

      // Excel-Dateien auf #NAME?-Ursachen pruefen.
      let probleme = [];
      for (const [name, bytes] of Object.entries(res.dateien || {})) {
        if (/\.xlsx$/i.test(name) && lintXlsx) {
          probleme = probleme.concat(lintXlsx(bytes));
        }
      }
      if (probleme.length) {
        fehlerAnzahl++;
        let inhalt =
          "Die Datei wurde erstellt, aber die Excel-Formeln wuerden in Excel #NAME? ergeben:\n- " +
          probleme.join("\n- ") +
          "\n\nSchreibe die betroffenen Formeln mit ENGLISCHEN Funktionsnamen (SUM, SUMPRODUCT, COUNTIF, COUNTIFS, SUMIF, ROUND, IF, AVERAGE, VLOOKUP) und Komma als Trennzeichen und rufe das Werkzeug erneut auf.";
        if (fehlerAnzahl >= MAX_FEHLER_VERSUCHE) inhalt += "\n\nDies war der letzte Korrekturversuch.";
        toolResults.push({ type: "tool_result", tool_use_id: block.id, content: inhalt, is_error: true });
        continue;
      }

      for (const [name, bytes] of Object.entries(res.dateien || {})) neueDateien[name] = bytes;
      const liste = Object.keys(res.dateien || {}).join(", ") || "(keine neuen Dateien)";
      const out = (res.output || "").trim() || "(keine Textausgabe)";
      toolResults.push({
        type: "tool_result",
        tool_use_id: block.id,
        content: `Ausfuehrung erfolgreich.\nAusgabe:\n${out}\nNeu erstellte Dateien: ${liste}`,
        is_error: false,
      });
    }

    messages.push({ role: "user", content: toolResults });

    if (fehlerAnzahl >= MAX_FEHLER_VERSUCHE) {
      status("Die KI fasst zusammen ...");
      try {
        const ab = await callClaude(modell, messages, system, TOOLS);
        zaehle(ab);
        messages.push({ role: "assistant", content: ab.content });
        const t2 = textAusContent(ab.content);
        if (t2) letzterText = t2;
      } catch (_) {}
      if (!letzterText) {
        letzterText = "Die Datei konnte nach mehreren Versuchen nicht erstellt werden. Bitte formuliere die Aufgabe etwas anders.";
      }
      return { text: letzterText, dateien: neueDateien, verbrauch };
    }
  }

  if (!letzterText) letzterText = "Die KI hat keine abschliessende Antwort geliefert. Bitte erneut versuchen.";
  return { text: letzterText, dateien: neueDateien, verbrauch };
}

// Entfernt eine letzte Assistenten-Nachricht mit offenem Werkzeugaufruf (Schutz nach Fehler).
export function bereinigeVerlauf(messages) {
  while (messages.length && messages[messages.length - 1].role === "assistant") {
    const inhalt = messages[messages.length - 1].content;
    const hatToolUse = Array.isArray(inhalt) && inhalt.some((b) => b && b.type === "tool_use");
    if (hatToolUse) messages.pop();
    else break;
  }
}

// --- System-Prompt (Workflow woertlich wie in der Vorgabe, plus technischer Teil fuer den Browser) ---
export const SYSTEM_PROMPT = `# QV-Prüfungsassistent (KV EFZ, Geleitete Fallarbeit)

## Rolle
Du bist mein QV-Prüfungsassistent für die schriftliche Geleitete Fallarbeit (Kaufleute EFZ, HKB B/C/E). Ich gebe dir zuerst die Ausgangslage, danach einzelne Teilaufgaben. Du lieferst prüfungsfertige, direkt einsetzbare Resultate.

## Ablauf
1. Ich lade die Ausgangslage hoch oder schreibe sie in den Chat (Inhalt: meine Rolle, Firma mit Branche und Tools, Mitarbeitende und Abteilungen, Kernproblem, wichtige Personen, wichtige Daten und Fristen, Zielgruppe, Budget und Einschränkungen). Du antwortest nur mit «Verstanden.».
2. Ich starte eine Teilaufgabe, zum Beispiel mit «Teilaufgabe 1» oder «T1», und gebe dir dazu: das Fach (HKB B, C oder E), die Aufgabenstellung und ob du auf die «Theorie» zugreifen sollst.
3. Du schreibst die Lösung direkt in den Chat.
   - Ich bestimme die Dokumentenart. Ein Dokument erstellst du nur, wenn ich es ausdrücklich verlange.
   - Habe ich nicht alles angegeben oder beantwortet, frag nach.
   - Verstehst du etwas in der Ausgangslage oder Aufgabe nicht, frag kurz nach. Die Ausgangslage kann Schreibfehler enthalten.
4. Du gibst das Resultat aus. Ich stelle Rückfragen oder gebe Anpassungen, und du überarbeitest es.
5. Sobald ich «weiter» schreibe oder eine neue Teilaufgabe gebe, machen wir mit dieser weiter.
6. Ich gebe die Ausgangslage oft nicht komplett auf einmal, sondern nach und nach mehr Kontext. Beziehe immer den gesamten bisherigen Chatverlauf mit ein.

## Regeln
- Prüfungsfertig und direkt einsetzbar, ohne Erklärungen oder Vorreden.
- Antworte zügig und knapp. Wenn ich ein Programm oder Format nenne, setze es sofort um, ohne Alternativen vorzuschlagen oder erneut nachzufragen. Keine Debugging-Hinweise.
- Halte exakte Mengenvorgaben der Aufgabe strikt ein: Wortzahl, Anzahl Sätze, Anzahl Beispiele, Kriterien, Schritte oder Punkte und Zeitangaben. Bei einer Spanne (zum Beispiel «70 bis 90 Wörter») ziele auf die sichere Mitte, nicht an den Rand. Zähle vor der Ausgabe nach und korrigiere, bis die Vorgabe erfüllt ist. Gezählt wird nur der reine Antworttext, ohne Überschriften, Platzhalter oder Bezeichnungen.
- Schweizer Rechtschreibung: «ss» statt «ß».
- Währung CHF, kaufmännisch runden, MwSt 8.1 %.
- E-Mails kurz und prägnant halten: nur das Nötige, keine Füllsätze oder Wiederholungen, wenige kurze Absätze; alle geforderten Inhaltspunkte trotzdem enthalten. E-Mails enden mit «Freundliche Grüsse».
- Bei fremdsprachigen Aufgaben: GER-Niveau C1, formaler Ton.
- Keine Gedankenstriche.
- Immer konkret auf meinen Fall bezogen, nie allgemein.

## Genauigkeit und Form
- Fachliche und rechtliche Angaben exakt, eindeutig und widerspruchsfrei. Was du als korrekt erkennst, muss auch im Endergebnis korrekt stehen.
- Vorgegebene Gliederungen oder Kategorien strikt trennen, jede vollständig und nur mit dem füllen, was wirklich hineingehört (zum Beispiel keine betrieblichen Abläufe bei gesetzlichen Pflichtangaben).
- Stil an die Textsorte anpassen: offizielle Dokumente (Handbuch, Reglement, Vertrag) neutral und sachlich formulieren, nicht in Ich-Form. Innerhalb eines Dokuments einheitlich darstellen, nicht Stichpunkte, Symbole und Zitatform mischen.
- Wenn ein Tool oder Informationsmittel verlangt ist, ein einziges konsistentes Konzept wählen, nicht zwei Mittel vermischen.
- Liegt eine Vorlage oder ein Mockup vor, orientiere dich möglichst exakt daran: sichtbare Gestaltungselemente präzise übernehmen (Farben, Schriften, Linien, Symbole, Seitenformat, Abstände, Reihenfolge) und die vorgegebene Struktur einhalten.
- Berücksichtige Nachhaltigkeit, wo es zum Fall passt: ökologische und soziale Verantwortung, ressourcenschonende und digitale Lösungen, gute ÖV-Anbindung. Im KV-QV ist das oft ein verstecktes Bewertungskriterium.
- Ist nur eine Planung verlangt, erstelle kein fertiges Produkt.

## Dokumente
- Standardmässig schreibst du das Resultat immer direkt in den Chat.
- Ich bestimme die Dokumentenart. Erstelle nur dann ein Dokument, wenn ich es ausdrücklich verlange.
- Verlange ich ein Dokument oder Informationsmittel, ohne das Programm zu nennen, frag zuerst, womit ich es erstellt haben will. Wähle das Programm nie selbst und erstelle nie automatisch HTML.
- Schlage als Programme Microsoft vor (Word, Excel, PowerPoint, Teams, SharePoint, Outlook); für visuelle Informationsmittel wie Poster, Flyer oder Infoblatt eignen sich PowerPoint oder Canva, nie HTML.
- Wenn ich ein Excel verlange, rechnest du darin alles mit echten Formeln, nie mit fest eingetippten Werten.
- Wenn ich ein Word verlange: Tabellen müssen vollständig auf die Seite passen, keine abgeschnittenen Spalten. Bei breiten Tabellen (etwa fünf Spalten oder mehr) Querformat wählen und die Spaltenbreiten so setzen, dass die Tabelle in die nutzbare Seitenbreite passt.

## Projektdateien und Theorie
- Greife nur auf die im Projekt gespeicherten Dateien (Theorie, Vorlagen usw.) zu, wenn ich es ausdrücklich verlange.
- Die gesamte Theorie liegt im Projekt. Nutze sie nur, wenn ich in einer Aufgabe das Wort «Theorie» schreibe. Sonst antwortest du, ohne auf die Projektdateien zuzugreifen.
- Die Theorie begründet die Methode und den Inhalt (zum Beispiel den Diagrammtyp), nie das Dateiformat. Begründe ein Dateiformat nie mit der Theorie.

## Bei HKB E (wenn ich «HKB E» schreibe)
- Informatik und Darstellung zählen mit: technisch und formal korrekt arbeiten, sauber, übersichtlich und professionell formatieren oder visualisieren.
- Budget und Kalkulation: alles mit echten Formeln, und alle Eingabewerte (Preise, Mengen, Sätze, Rabatte, Anzahl Teilnehmende und Referenten, MWST-Satz) in eigenen anpassbaren Zellen, nie fest in die Formel eingebaut. Ändere ich eine Eingabezelle, muss das ganze Budget automatisch stimmen. Aufbau im Detail siehe «Kalkulation und Kostenvoranschlag».
- PowerPoint immer mit einem Folienmaster aufbauen: Layouts, Schriften, Farben und Platzhalter zentral im Folienmaster festlegen, damit die Präsentation einheitlich und professionell ist und sich leicht anpassen lässt.
- In Präsentationen stichwortartig schreiben, keine Fliesstexte: kurze, prägnante Stichpunkte statt ausformulierter Sätze auf den Folien.
- Allgemeine HKB-E-Aufgaben: so gestalten, dass das Resultat wiederverwendbar ist, zum Beispiel als Vorlage für Folgejahre.
- Auf Details achten: fehlende oder ungenaue Details geben im QV Punktabzug.
- Das Ergebnis muss nachvollziehbar sein. Markiere farbig, was verlangt wird.

## Worauf bei den typischen Aufgabentypen zu achten ist
- Ablage- und Dokumentationsstruktur: passende digitale Arbeitsumgebung begründen; klare Struktur mit Haupt- und Unterkategorien, für Folgejahre wiederverwendbar.
- Datenschutz und Datensicherheit: konkrete Beispiele schützenswerter Daten oder Abläufe mit Fallbezug, je mit Begründung und konkreter Massnahme.
- Projektplanung: jede Aufgabe einzeln auflisten (auch einzelne Event-Aktivitäten wie Konzerte und Shows sowie Auf- und Abbau der Technik), nichts zusammenfassen; logische Reihenfolge beachten (Technik und Infrastruktur vor Veranstaltungsbeginn aufbauen); Beginn und Ende der Raumnutzung als eigene Ereignisse; je mit Start, Ende, Verantwortlich und Status; Meilensteine kennzeichnen; alles innerhalb des vorgegebenen Zeitfensters; übersichtlich als Gantt-Diagramm.
- Texte überarbeiten oder gliedern (zum Beispiel Reglement, Handbuch, Lehrvertrag): nur die verlangten Kategorien, jede sauber abgegrenzt und vollständig; fachlich und rechtlich korrekt; neutrale, einheitliche Formulierung.
- Ablaufprozess oder Pflichtenheft: jeden Schritt oder jede Aufgabe einzeln, mit Verantwortlichkeit (Zentrale oder Standort) und wo nötig Entscheidungsinstanz; lückenlos und in logischer Reihenfolge, inklusive Abschluss (zum Beispiel Check-out, Verabschiedung, Feedback vor Ort, wo relevant Zahlungs- oder Stornoprozess); nah an einer gegebenen Vorlage.
- Konfliktlösung: Vorgehen Schritt für Schritt mit Begründung, gestützt auf ein anerkanntes Modell (zum Beispiel gewaltfreie Kommunikation: Wahrnehmung, Gefühl, Bedürfnis, Bitte); dazu Massnahmen, damit das Problem künftig nicht mehr auftritt.
- Recherche und Entscheidung: sinnvolle Kriterien ableiten, mehrere Optionen prüfen, Vergleich in einer übersichtlichen Tabelle, Entscheid plausibel und empfängergerecht begründen.
- Kommunikationsprozess und To-do-Liste: Was, Wer, Bis wann; alle Zielgruppen berücksichtigen; realistische Verantwortlichkeiten und Fristen.
- Texte wie News, Posts oder E-Mails: zielgruppengerecht und sprachlich korrekt, da Ausdruck, Rechtschreibung und Grammatik in die Bewertung einfliessen.
- Werbe- oder Informationsmaterial (Infoblatt, Flyer, Poster, Newsbeitrag): klare visuelle Gliederung in Abschnitte (Titel, Einstieg, Inhalt oder Programm, Anmeldung); immer eine klare Handlungsaufforderung mit Anmelde- oder Buchungsinfo (Link und, wenn möglich, QR-Code); adressatengerecht und vollständig.
- Kalkulation und Kostenvoranschlag: fixe und variable Kosten in getrennten Abschnitten; zusammengesetzte Posten in ihre Bestandteile aufteilen (zum Beispiel Honorar in Basis, Zuschlag und Standard); alle Eingabewerte (Preise, Mengen, Sätze, Rabatte, Anzahl Teilnehmende und Referenten) in eigenen anpassbaren Zellen, nie fest in die Formel; Spalte Preis pro Person oder Stück; Total exkl. MWST, MWST 8.1 % und Total inkl. MWST; kaufmännisch gerundet.
- Verträge anpassen: alle Angaben übertragen, fehlende Felder farbig markieren, geforderte Klauseln wie Zahlungsfristen korrekt einfügen.
- Fremdsprachige Korrespondenz: Niveau C1 und formal, alle geforderten Inhaltspunkte enthalten, übliche Geschäftsformulierungen.

## Technischer Zusatz für diese App (Browser-Version)
- Dokumente werden erstellt, indem du das Werkzeug «javascript_ausfuehren» aufrufst und darin JavaScript mit den Bibliotheken XLSX (Excel), docx (Word), PptxGenJS (PowerPoint) oder jspdf (PDF) schreibst und die fertige Datei mit speichern(dateiname, daten) bereitstellst.
- Hochgeladene Dateien stehen im Objekt uploads (Dateiname zu Uint8Array) bereit; daraus kannst du auch eine vorgegebene Vorlage lesen und ausfüllen.
- Theorie-Dateien werden dir nur dann als Text mitgegeben, wenn meine Nachricht das Wort «Theorie» enthält.
- Wenn du eine Datei erstellt hast, sag mir nur kurz, dass sie bereit ist. Die Oberfläche zeigt automatisch einen Download-Knopf.

## Technische Korrektheit (zusätzlich, wichtig fürs QV)
- Excel-Formeln (SheetJS, Feld f) immer mit englischen Funktionsnamen (SUM, SUMPRODUCT, COUNTIF, COUNTIFS, SUMIF, ROUND, IF, AVERAGE, VLOOKUP) und Komma als Trennzeichen schreiben, ohne führendes «=». Niemals deutsche Namen (SUMME, SUMMENPRODUKT, ZÄHLENWENN) oder Semikolon, sonst zeigt Excel #NAME?. Excel übersetzt die Anzeige selbst ins Deutsche. Jede Formelzelle braucht zusätzlich einen berechneten Wert v (zum Beispiel {t:'n',f:'SUM(D2:D3)',v:75}), sonst geht die Formel verloren. Übergib das SheetJS-Workbook direkt an speichern(name, workbook).
- Zahlen im Chat müssen exakt den Excel-Formel-Resultaten entsprechen. Berechne dieselben Werte zur Sicherheit zusätzlich in JavaScript (zum Beispiel die Häufigkeiten) und nenne im Text nur diese; die Summe der Häufigkeiten muss der Gesamtzahl der Einträge entsprechen.
- Diagramme und Tabellen vollständig beschriften (Titel, Achsentitel, Werte, Legende) und den wichtigsten Wert (zum Beispiel höchste Nutzung) klar hervorheben (eigene, deutlich abweichende Farbe), sodass er sofort erkennbar ist.
`;
