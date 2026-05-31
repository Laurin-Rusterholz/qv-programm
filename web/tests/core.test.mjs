// Node-Tests fuer core.js (reine Logik + Datei-Erstellung mit echten Bibliotheken).
// Start:  npm test   (im Ordner web/)

import * as Core from "../core.js";
import XLSX from "xlsx";
import * as docx from "docx";
import PptxGenJS from "pptxgenjs";
import * as jspdfNS from "jspdf";

let bestanden = 0;
let fehlgeschlagen = 0;

function pruefe(bedingung, beschreibung) {
  if (bedingung) {
    bestanden++;
    console.log(`  [OK]   ${beschreibung}`);
  } else {
    fehlgeschlagen++;
    console.log(`  [FEHL] ${beschreibung}`);
  }
}
function abschnitt(t) {
  console.log("\n" + t);
}

const libs = { XLSX, docx, PptxGenJS, jspdf: jspdfNS };

async function runCode(code) {
  try {
    const r = await Core.fuehreJsAus(code, libs, {});
    return { error: null, output: r.output, dateien: r.dateien };
  } catch (e) {
    return { error: String((e && e.stack) || e), output: "", dateien: {} };
  }
}
const lintXlsx = (bytes) => Core.pruefeXlsxBytes(XLSX, bytes);

class FakeClaude {
  constructor(antworten) {
    this.antworten = antworten;
    this.calls = [];
  }
  async call(model, messages, system, tools) {
    this.calls.push({ model, system, tools });
    return this.antworten.shift();
  }
}

async function main() {
  console.log("================================================================");
  console.log(" Browser-Core-Tests (Node)");
  console.log("================================================================");

  // 1) Konstanten / Kosten
  abschnitt("Test 1: Konstanten und Kosten");
  pruefe(Core.STANDARD_MODELL === "claude-sonnet-4-6", "Standardmodell Sonnet 4.6");
  pruefe(
    Core.MODELL_OPTIONEN["Beste Qualitaet (Opus 4.8)"] === "claude-opus-4-8",
    "Opus 4.8 als Option vorhanden"
  );
  pruefe(Core.MAX_TOKENS >= 8000, "max_tokens >= 8000");
  pruefe(/^[a-zA-Z0-9_-]{1,64}$/.test(Core.TOOL_NAME), `Werkzeugname gueltig: ${Core.TOOL_NAME}`);
  pruefe(
    Math.abs(Core.geschaetzteKosten("claude-sonnet-4-6", 1e6, 1e6) - 18.0) < 1e-6,
    "Kosten Sonnet (3+15=18 USD/1Mio)"
  );
  pruefe(Core.SYSTEM_PROMPT.includes("Verstanden") && Core.SYSTEM_PROMPT.includes("javascript_ausfuehren"),
    "System-Prompt enthaelt Workflow und Werkzeugname");

  // 2) Formel-Pruefung (#NAME?)
  abschnitt("Test 2: Excel-Formeln gegen #NAME? pruefen");
  const probleme = Core.findeFormelProbleme([
    { ref: "A3", f: "SUMME(A1:A2)" },
    { ref: "A4", f: "ZÄHLENWENN(A1:A2,1)" },
    { ref: "A5", f: "SUM(A1;A2)" },
    { ref: "A6", f: "SUM(A1:A2)" },
  ]);
  pruefe(probleme.some((p) => p.includes("SUMME") && p.includes("SUM")), "SUMME erkannt (-> SUM)");
  pruefe(probleme.some((p) => p.includes("COUNTIF")), "ZÄHLENWENN erkannt (-> COUNTIF)");
  pruefe(probleme.some((p) => p.includes("Semikolon")), "Semikolon erkannt");
  pruefe(!probleme.some((p) => p.startsWith("A6")), "Korrekte englische Formel nicht beanstandet");

  // 3) Excel-Datei mit echten Formeln erzeugen (englisch) und pruefen
  abschnitt("Test 3: Excel-Erstellung mit echten Formeln (SheetJS)");
  const excelCode = `
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
      ["Position","Menge","Preis","Total"],
      ["Getraenke",100,2.5,{t:"n",f:"B2*C2",v:250}],
      ["Snacks",50,4,{t:"n",f:"B3*C3",v:200}],
      ["Total exkl.",null,null,{t:"n",f:"SUM(D2:D3)",v:450}],
      ["MWST 8.1%",null,null,{t:"n",f:"ROUND(D4*0.081,2)",v:36.45}],
      ["Total inkl.",null,null,{t:"n",f:"D4+D5",v:486.45}],
    ]);
    XLSX.utils.book_append_sheet(wb, ws, "Budget");
    print("Total:", 100*2.5 + 50*4);
    speichern("Budget.xlsx", wb);
  `;
  const exRes = await runCode(excelCode);
  pruefe(exRes.error === null, "Excel-Code lief ohne Fehler");
  pruefe(!!exRes.dateien["Budget.xlsx"] && exRes.dateien["Budget.xlsx"].length > 0, "Budget.xlsx erzeugt");
  pruefe(exRes.output.includes("450"), "print-Ausgabe (Total 450) erfasst");
  const formeln = Core.extrahiereXlsxFormeln(XLSX, exRes.dateien["Budget.xlsx"]);
  pruefe(formeln.some((f) => f.f === "B2*C2"), "Zeilen-Total verweist auf Eingabezellen (B2*C2)");
  pruefe(formeln.some((f) => /SUM\(/.test(f.f)), "SUM-Formel vorhanden");
  pruefe(Core.pruefeXlsxBytes(XLSX, exRes.dateien["Budget.xlsx"]).length === 0, "Keine #NAME?-Probleme (englisch)");

  // 4) Word, PowerPoint, PDF erzeugen
  abschnitt("Test 4: Word/PowerPoint/PDF-Erstellung");
  const wordRes = await runCode(`
    const {Document,Packer,Paragraph,HeadingLevel}=docx;
    const doc=new Document({sections:[{children:[
      new Paragraph({text:"Projektplan",heading:HeadingLevel.HEADING_1}),
      new Paragraph("Inhalt"),
    ]}]});
    speichern("Plan.docx", await Packer.toBlob(doc));
  `);
  pruefe(wordRes.error === null && wordRes.dateien["Plan.docx"]?.length > 0, "Word-Datei erzeugt");
  // .docx ist ein ZIP -> beginnt mit PK
  pruefe(wordRes.dateien["Plan.docx"][0] === 0x50 && wordRes.dateien["Plan.docx"][1] === 0x4b, "docx ist gueltiges ZIP (PK)");

  const pptRes = await runCode(`
    const p=new PptxGenJS();
    const s1=p.addSlide(); s1.addText("Sommerfest",{x:1,y:1,fontSize:24});
    const s2=p.addSlide(); s2.addText([{text:"Aufgabe 1"},{text:"Aufgabe 2"}].map(t=>({text:t.text,options:{bullet:true}})),{x:1,y:1});
    speichern("Folien.pptx", await p.write({outputType:"arraybuffer"}));
  `);
  pruefe(pptRes.error === null && pptRes.dateien["Folien.pptx"]?.length > 0, "PowerPoint-Datei erzeugt");
  pruefe(pptRes.dateien["Folien.pptx"][0] === 0x50, "pptx ist gueltiges ZIP (PK)");

  const pdfRes = await runCode(`
    const {jsPDF}=jspdf;
    const doc=new jsPDF();
    doc.setFontSize(16); doc.text("Infoblatt Sommerfest",10,20);
    speichern("Info.pdf", doc.output("arraybuffer"));
  `);
  pruefe(pdfRes.error === null && pdfRes.dateien["Info.pdf"]?.length > 0, "PDF-Datei erzeugt");
  const pdfKopf = new TextDecoder().decode(pdfRes.dateien["Info.pdf"].slice(0, 5));
  pruefe(pdfKopf === "%PDF-", "PDF beginnt mit %PDF-");

  // 5) Fehlerbehandlung
  abschnitt("Test 5: Fehlerhafter Code wird sauber gemeldet");
  const fehlerRes = await runCode(`throw new Error("kaputt");`);
  pruefe(fehlerRes.error !== null && fehlerRes.error.includes("kaputt"), "Fehler wird gemeldet (kein Absturz)");
  pruefe(Object.keys(fehlerRes.dateien).length === 0, "Bei Fehler keine Dateien");

  // 6) Nachrichten-Inhalt (mehrere Dateien + Theorie)
  abschnitt("Test 6: Nachrichten-Inhalt (PDF/Bild/Text + Theorie)");
  const content = Core.baueUserContent(
    "Vergleiche die Dateien",
    [
      { kind: "pdf", name: "a.pdf", base64: "AAAA" },
      { kind: "image", name: "b.png", mediaType: "image/png", base64: "BBBB" },
      { kind: "text", name: "c.xlsx", text: "XLSXINHALT" },
    ],
    "THEORIETEXT"
  );
  const typen = content.map((b) => b.type);
  pruefe(typen.filter((t) => t === "document").length === 1, "PDF als document-Block");
  pruefe(typen.filter((t) => t === "image").length === 1, "Bild als image-Block");
  pruefe(content.some((b) => b.type === "text" && b.text.includes("XLSXINHALT")), "Excel-Text enthalten");
  pruefe(content.some((b) => b.type === "text" && b.text.includes("THEORIE")), "Theorie-Kontext enthalten");
  pruefe(content.some((b) => b.type === "text" && b.text.includes("Vergleiche die Dateien")), "Nutzertext enthalten");
  const nurDatei = Core.baueUserContent("", [{ kind: "pdf", name: "a.pdf", base64: "AAAA" }], "");
  pruefe(nurDatei.some((b) => b.type === "document"), "Nur-Datei-Nachricht enthaelt die Datei");

  // 7) Konversations-Schleife (Mock-Claude) erstellt Datei
  abschnitt("Test 7: Werkzeug-Schleife erstellt Datei (Mock-Claude)");
  const claude7 = new FakeClaude([
    {
      stop_reason: "tool_use",
      usage: { input_tokens: 1000, output_tokens: 200 },
      content: [
        { type: "text", text: "Ich erstelle die Datei." },
        { type: "tool_use", id: "t1", name: Core.TOOL_NAME, input: { code: excelCode } },
      ],
    },
    {
      stop_reason: "end_turn",
      usage: { input_tokens: 1200, output_tokens: 50 },
      content: [{ type: "text", text: "Die Datei ist bereit." }],
    },
  ]);
  const messages7 = [{ role: "user", content: "Mach ein Excel" }];
  const erg7 = await Core.fuehreKonversation(
    { callClaude: (m, ms, s, t) => claude7.call(m, ms, s, t), runCode, lintXlsx },
    "claude-sonnet-4-6",
    messages7
  );
  pruefe(erg7.text === "Die Datei ist bereit.", "Schlusstext zurueckgegeben");
  pruefe(!!erg7.dateien["Budget.xlsx"], "Datei aus der Schleife vorhanden");
  pruefe(erg7.verbrauch.input === 2200 && erg7.verbrauch.output === 250, "Token-Verbrauch summiert");

  // 8) Automatische #NAME?-Korrektur in der Schleife
  abschnitt("Test 8: Automatische Korrektur deutscher Excel-Funktion (Mock-Claude)");
  const deutschCode = `
    const wb=XLSX.utils.book_new();
    const ws=XLSX.utils.aoa_to_sheet([[1],[2],[{t:"n",f:"SUMME(A1:A2)",v:3}]]);
    XLSX.utils.book_append_sheet(wb,ws,"Blatt1");
    speichern("Auto.xlsx", wb);
  `;
  const englischCode = `
    const wb=XLSX.utils.book_new();
    const ws=XLSX.utils.aoa_to_sheet([[1],[2],[{t:"n",f:"SUM(A1:A2)",v:3}]]);
    XLSX.utils.book_append_sheet(wb,ws,"Blatt1");
    speichern("Auto.xlsx", wb);
  `;
  const claude8 = new FakeClaude([
    { stop_reason: "tool_use", content: [{ type: "tool_use", id: "t1", name: Core.TOOL_NAME, input: { code: deutschCode } }] },
    { stop_reason: "tool_use", content: [{ type: "tool_use", id: "t2", name: Core.TOOL_NAME, input: { code: englischCode } }] },
    { stop_reason: "end_turn", content: [{ type: "text", text: "Jetzt ist die Datei bereit." }] },
  ]);
  const messages8 = [{ role: "user", content: "Excel mit SUMME" }];
  const erg8 = await Core.fuehreKonversation(
    { callClaude: (m, ms, s, t) => claude8.call(m, ms, s, t), runCode, lintXlsx },
    "claude-sonnet-4-6",
    messages8
  );
  pruefe(!!erg8.dateien["Auto.xlsx"], "Datei nach Korrektur geliefert");
  pruefe(Core.pruefeXlsxBytes(XLSX, erg8.dateien["Auto.xlsx"]).length === 0, "Gelieferte Datei ist sauber (englisch)");
  const hinweise = messages8.filter(
    (m) => m.role === "user" && Array.isArray(m.content)
  ).flatMap((m) => m.content).filter((b) => b && b.type === "tool_result" && b.is_error && String(b.content).includes("NAME?"));
  pruefe(hinweise.length >= 1, "Deutsche Funktion wurde automatisch zur Korrektur zurueckgemeldet");

  console.log("\n================================================================");
  console.log(` Ergebnis: ${bestanden} bestanden, ${fehlgeschlagen} fehlgeschlagen`);
  console.log("================================================================");
  process.exit(fehlgeschlagen === 0 ? 0 : 1);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
