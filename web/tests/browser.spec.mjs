// Echte Browser-Tests (Playwright + Chromium) mit gemockter Anthropic-API.
// Start:  node tests/browser.spec.mjs   (im Ordner web/)

import { chromium } from "@playwright/test";
import http from "node:http";
import fs from "node:fs";
import path from "node:path";
import os from "node:os";
import { fileURLToPath } from "node:url";
import XLSX from "xlsx";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");

let bestanden = 0, fehlgeschlagen = 0;
function pruefe(b, t) { if (b) { bestanden++; console.log(`  [OK]   ${t}`); } else { fehlgeschlagen++; console.log(`  [FEHL] ${t}`); } }

const MIME = { ".html": "text/html", ".js": "text/javascript", ".mjs": "text/javascript", ".css": "text/css", ".json": "application/json" };

function starteServer() {
  const server = http.createServer((req, res) => {
    let p = decodeURIComponent(req.url.split("?")[0]);
    if (p === "/") p = "/index.html";
    const datei = path.join(ROOT, p);
    if (!datei.startsWith(ROOT) || !fs.existsSync(datei) || fs.statSync(datei).isDirectory()) {
      res.writeHead(404); res.end("not found"); return;
    }
    res.writeHead(200, { "content-type": MIME[path.extname(datei)] || "application/octet-stream" });
    fs.createReadStream(datei).pipe(res);
  });
  return new Promise((resolve) => server.listen(0, () => resolve(server)));
}

// Echter Browser-JS-Code, den die gemockte KI "schreibt" (laeuft im Browser mit window.XLSX).
const EXCEL_CODE = `
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ["Position","Menge","Preis","Total"],
    ["Getraenke",100,2.5,{t:"n",f:"B2*C2",v:250}],
    ["Snacks",50,4,{t:"n",f:"B3*C3",v:200}],
    ["Total exkl.",null,null,{t:"n",f:"SUM(D2:D3)",v:450}],
    ["MWST 8.1%",null,null,{t:"n",f:"ROUND(D4*0.081,2)",v:36.45}],
    ["Total inkl.",null,null,{t:"n",f:"D4+D5",v:486.45}]
  ]);
  XLSX.utils.book_append_sheet(wb, ws, "Budget");
  print("Total exkl. MWST:", 450);
  speichern("Budget.xlsx", wb);
`;

const ANTWORTEN = [
  {
    id: "m1", type: "message", role: "assistant", model: "claude-sonnet-4-6",
    stop_reason: "tool_use", usage: { input_tokens: 100, output_tokens: 20 },
    content: [
      { type: "text", text: "Ich erstelle die Excel-Datei." },
      { type: "tool_use", id: "tool_1", name: "javascript_ausfuehren", input: { code: EXCEL_CODE } },
    ],
  },
  {
    id: "m2", type: "message", role: "assistant", model: "claude-sonnet-4-6",
    stop_reason: "end_turn", usage: { input_tokens: 120, output_tokens: 12 },
    content: [{ type: "text", text: "Die Excel-Datei ist bereit." }],
  },
];

async function main() {
  console.log("================================================================");
  console.log(" Browser-Tests (Playwright/Chromium, gemockte API)");
  console.log("================================================================\n");

  const server = await starteServer();
  const port = server.address().port;
  const browser = await chromium.launch();
  const page = await browser.newPage();

  const konsolenfehler = [];
  page.on("pageerror", (e) => konsolenfehler.push(String(e)));

  // Anthropic-API mocken (inkl. CORS-Preflight).
  let aufrufe = 0;
  await page.route("https://api.anthropic.com/v1/messages", async (route) => {
    const req = route.request();
    if (req.method() === "OPTIONS") {
      await route.fulfill({ status: 204, headers: {
        "access-control-allow-origin": "*",
        "access-control-allow-headers": "x-api-key,anthropic-version,content-type,anthropic-dangerous-direct-browser-access",
        "access-control-allow-methods": "POST,OPTIONS",
      }, body: "" });
      return;
    }
    const body = ANTWORTEN[Math.min(aufrufe, ANTWORTEN.length - 1)];
    aufrufe++;
    await route.fulfill({ status: 200, headers: { "access-control-allow-origin": "*", "content-type": "application/json" }, body: JSON.stringify(body) });
  });

  await page.goto(`http://localhost:${port}/index.html`);

  console.log("Test 1: Seite und Bibliotheken laden");
  const globals = await page.evaluate(() => ({
    XLSX: !!window.XLSX, docx: !!window.docx, PptxGenJS: !!window.PptxGenJS,
    jspdf: !!(window.jspdf && window.jspdf.jsPDF), mammoth: !!window.mammoth,
  }));
  pruefe(globals.XLSX, "XLSX (SheetJS) geladen");
  pruefe(globals.docx, "docx geladen");
  pruefe(globals.PptxGenJS, "PptxGenJS geladen");
  pruefe(globals.jspdf, "jspdf.jsPDF geladen");
  pruefe(globals.mammoth, "mammoth geladen");
  pruefe(!!(await page.$("#eingabe")), "Eingabefeld vorhanden");

  console.log("\nTest 2: Schluessel eintragen, Nachricht senden, Datei entsteht");
  await page.fill("#apikey", "sk-ant-testkey");
  pruefe(await page.evaluate(() => localStorage.getItem("qv_apikey")) === "sk-ant-testkey", "Schluessel wird in localStorage gespeichert");

  await page.fill("#eingabe", "Erstelle ein Budget als Excel.");
  await page.click("#senden-btn");

  // Auf die Assistenten-Antwort warten.
  await page.waitForFunction(() => {
    const els = document.querySelectorAll("#verlauf .nachricht.assistant .bubble");
    return [...els].some((e) => e.textContent.includes("Die Excel-Datei ist bereit"));
  }, { timeout: 15000 });
  pruefe(true, "KI-Antwort erscheint im Chat");
  pruefe(aufrufe === 2, `Werkzeug-Schleife hat 2 API-Aufrufe gemacht (war ${aufrufe})`);

  const dlBtn = await page.$("#verlauf .dl-knopf");
  pruefe(!!dlBtn, "Download-Knopf erscheint");
  const btnText = dlBtn ? await dlBtn.textContent() : "";
  pruefe(btnText.includes("Budget.xlsx"), "Download-Knopf nennt Budget.xlsx");

  console.log("\nTest 3: Heruntergeladene Excel-Datei ist gueltig und hat echte Formeln");
  const [download] = await Promise.all([
    page.waitForEvent("download"),
    dlBtn.click(),
  ]);
  const ziel = path.join(os.tmpdir(), "qv_browser_test.xlsx");
  await download.saveAs(ziel);
  const bytes = new Uint8Array(fs.readFileSync(ziel));
  pruefe(bytes[0] === 0x50 && bytes[1] === 0x4b, "Datei ist ein gueltiges xlsx (ZIP/PK)");
  const wb = XLSX.read(bytes, { type: "array" });
  const ws = wb.Sheets["Budget"];
  pruefe(ws && ws["D2"] && ws["D2"].f === "B2*C2", "Zeilen-Total ist Formel B2*C2 (echte Formel)");
  pruefe(ws && ws["D4"] && /SUM\(/.test(ws["D4"].f || ""), "Summen-Formel SUM(...) vorhanden");
  const formeln = [];
  for (const ref of Object.keys(ws)) if (ref[0] !== "!" && ws[ref].f) formeln.push({ ref, f: ws[ref].f });
  pruefe(formeln.length >= 4, `Mehrere Formeln erhalten (${formeln.length})`);
  fs.unlinkSync(ziel);

  console.log("\nTest 4: Verlauf bleibt nach Neuladen gespeichert (localStorage)");
  await page.reload();
  await page.waitForFunction(() => document.querySelectorAll("#verlauf .nachricht").length >= 2, { timeout: 8000 }).catch(() => {});
  const anzahl = await page.evaluate(() => document.querySelectorAll("#verlauf .nachricht").length);
  pruefe(anzahl >= 2, `Nachrichten nach Reload noch da (${anzahl})`);

  pruefe(konsolenfehler.length === 0, "Keine JavaScript-Fehler auf der Seite" + (konsolenfehler.length ? ": " + konsolenfehler[0] : ""));

  await browser.close();
  server.close();

  console.log("\n================================================================");
  console.log(` Ergebnis: ${bestanden} bestanden, ${fehlgeschlagen} fehlgeschlagen`);
  console.log("================================================================");
  process.exit(fehlgeschlagen === 0 ? 0 : 1);
}

main().catch((e) => { console.error(e); process.exit(1); });
