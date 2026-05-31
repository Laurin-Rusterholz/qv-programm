// app.js - Browser-Oberflaeche des QV-Pruefungsassistenten.
// Alles laeuft lokal im Browser; Schluessel und Chat werden in localStorage gespeichert.

import * as Core from "./core.js";

const API_URL = "https://api.anthropic.com/v1/messages";

const BILD_TYPEN = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", webp: "image/webp" };
const MIME = {
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  pdf: "application/pdf", csv: "text/csv", txt: "text/plain",
};

// --- Zustand ---
const state = {
  apiKey: "",
  model: Core.STANDARD_MODELL,
  apiMessages: [],     // voller Verlauf fuer die API
  displayMessages: [], // {role, text, files:[{name,id}], info}
  theorie: [],         // [{name, text}]
  kostenTotal: 0,
  tokensTotal: 0,
};
const uploadsBytes = {};        // name -> Uint8Array (fuer das Werkzeug)
const createdFiles = new Map(); // id -> {name, bytes}
let pendingAttachments = [];     // [{descriptor, name}]
let busy = false;
let fileCounter = 0;

const $ = (id) => document.getElementById(id);

// --- Persistenz (localStorage) ---
function speichern() {
  try { localStorage.setItem("qv_apikey", state.apiKey); } catch (_) {}
  try { localStorage.setItem("qv_model", state.model); } catch (_) {}
  try { localStorage.setItem("qv_theorie", JSON.stringify(state.theorie)); } catch (_) {}
  try {
    localStorage.setItem("qv_conv", JSON.stringify({
      apiMessages: state.apiMessages,
      displayMessages: state.displayMessages,
      kostenTotal: state.kostenTotal,
      tokensTotal: state.tokensTotal,
    }));
    setStatus("");
  } catch (e) {
    // Quota ueberschritten: Schluessel/Theorie bleiben gespeichert, Verlauf nur in dieser Sitzung.
    setStatus("Hinweis: Der Chat ist zu gross zum dauerhaften Speichern (bleibt nur in dieser Sitzung).");
  }
}
function laden() {
  state.apiKey = localStorage.getItem("qv_apikey") || "";
  state.model = localStorage.getItem("qv_model") || Core.STANDARD_MODELL;
  try { state.theorie = JSON.parse(localStorage.getItem("qv_theorie") || "[]"); } catch (_) { state.theorie = []; }
  try {
    const conv = JSON.parse(localStorage.getItem("qv_conv") || "null");
    if (conv) {
      state.apiMessages = conv.apiMessages || [];
      state.displayMessages = conv.displayMessages || [];
      state.kostenTotal = conv.kostenTotal || 0;
      state.tokensTotal = conv.tokensTotal || 0;
    }
  } catch (_) {}
}

// --- Hilfen ---
function endung(name) { return (String(name).split(".").pop() || "").toLowerCase(); }
function mimeFor(name) { return MIME[endung(name)] || "application/octet-stream"; }

function setStatus(text, klasse) {
  const s = $("status-zeile");
  s.textContent = text || "";
  s.className = "hinweis" + (klasse ? " " + klasse : "");
}
function setBusy(b) {
  busy = b;
  $("senden-btn").disabled = b;
  $("eingabe").disabled = b;
  $("anhang-btn").disabled = b;
}

// --- API ---
function headers() {
  return {
    "content-type": "application/json",
    "x-api-key": state.apiKey,
    "anthropic-version": Core.ANTHROPIC_VERSION,
    "anthropic-dangerous-direct-browser-access": "true",
  };
}
async function callClaude(model, messages, system, tools) {
  let resp;
  try {
    resp = await fetch(API_URL, {
      method: "POST",
      headers: headers(),
      body: JSON.stringify({ model, max_tokens: Core.MAX_TOKENS, system, tools, messages }),
    });
  } catch (e) {
    const err = new Error("Keine Internetverbindung zur KI. Bitte Internet pruefen.");
    err.netz = true;
    throw err;
  }
  if (!resp.ok) {
    let msg = "";
    try { msg = (await resp.json())?.error?.message || ""; } catch (_) {}
    const err = new Error(msg || ("HTTP " + resp.status));
    err.status = resp.status;
    throw err;
  }
  return await resp.json();
}
function deutscherFehler(e) {
  if (e && e.netz) return "Keine Internetverbindung zur KI. Bitte Internet pruefen und erneut senden.";
  const s = e && e.status;
  if (s === 401) return "Der API-Schluessel ist ungueltig. Bitte links korrigieren.";
  if (s === 429) return "Zu viele Anfragen (Rate-Limit). Bitte kurz warten und erneut senden.";
  if (s === 400) return "Ungueltige Anfrage: " + (e.message || "");
  if (s && s >= 500) return "Die KI meldet einen Serverfehler. Eventuell fehlt API-Guthaben. Bitte erneut versuchen.";
  return "Unerwarteter Fehler: " + (e && e.message ? e.message : e);
}

async function runCode(code) {
  try {
    const libs = { XLSX: window.XLSX, docx: window.docx, PptxGenJS: window.PptxGenJS, jspdf: window.jspdf };
    const r = await Core.fuehreJsAus(code, libs, uploadsBytes);
    return { error: null, output: r.output, dateien: r.dateien };
  } catch (e) {
    return { error: String((e && e.stack) || e), output: "", dateien: {} };
  }
}
const lintXlsx = (bytes) => Core.pruefeXlsxBytes(window.XLSX, bytes);

// --- Datei-Upload -> Deskriptor + Rohbytes ---
async function dateiZuDeskriptor(file) {
  const name = file.name;
  const ext = endung(name);
  const ab = await file.arrayBuffer();
  const bytes = new Uint8Array(ab);
  uploadsBytes[name] = bytes;
  if (ext === "pdf") return { kind: "pdf", name, base64: Core.bytesZuBase64(bytes) };
  if (BILD_TYPEN[ext]) return { kind: "image", name, mediaType: BILD_TYPEN[ext], base64: Core.bytesZuBase64(bytes) };
  if (ext === "docx") {
    try {
      const r = await window.mammoth.extractRawText({ arrayBuffer: ab });
      return { kind: "text", name, text: r.value || "" };
    } catch (_) { return { kind: "text", name, text: "(Word konnte nicht gelesen werden)" }; }
  }
  if (ext === "xlsx" || ext === "xlsm") {
    try {
      const wb = window.XLSX.read(bytes, { type: "array" });
      let text = "";
      for (const sn of wb.SheetNames) text += `# ${sn}\n` + window.XLSX.utils.sheet_to_csv(wb.Sheets[sn]) + "\n";
      return { kind: "text", name, text };
    } catch (_) { return { kind: "text", name, text: "(Excel konnte nicht gelesen werden)" }; }
  }
  try { return { kind: "text", name, text: new TextDecoder().decode(bytes) }; }
  catch (_) { return { kind: "text", name, text: "(nicht lesbar)" }; }
}

// --- Theorie ---
function theorieText() {
  const teile = state.theorie.filter((t) => t.text && t.text.trim()).map((t) => `===== ${t.name} =====\n${t.text.trim()}`);
  let s = teile.join("\n\n");
  if (s.length > 24000) s = s.slice(0, 24000) + "\n[... Theorie gekuerzt ...]";
  return s;
}

// --- Rendering ---
function registriereDatei(name, bytes) {
  const id = "f" + (++fileCounter);
  createdFiles.set(id, { name, bytes });
  return id;
}
function downloadDatei(id) {
  const eintrag = createdFiles.get(id);
  if (!eintrag) return false;
  const blob = new Blob([eintrag.bytes], { type: mimeFor(eintrag.name) });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = eintrag.name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
  return true;
}

function renderVerlauf() {
  const v = $("verlauf");
  v.innerHTML = "";
  if (!state.apiKey) {
    const w = document.createElement("div");
    w.className = "warnung";
    w.textContent = "Trage links deinen API-Schluessel ein (von console.anthropic.com, Guthaben noetig) und druecke «Verbindung testen». Dann kannst du loslegen.";
    v.appendChild(w);
  }
  if (state.displayMessages.length === 0 && state.apiKey) {
    const w = document.createElement("div");
    w.className = "warnung";
    w.style.background = "#eef6ff"; w.style.borderColor = "#cfe0ff";
    w.textContent = "Bereit. Lade zuerst die Ausgangslage hoch (📎) oder schreibe sie in den Chat. Die KI antwortet dann mit «Verstanden.».";
    v.appendChild(w);
  }
  for (const m of state.displayMessages) {
    const div = document.createElement("div");
    div.className = "nachricht " + m.role;
    const rolle = document.createElement("div");
    rolle.className = "rolle";
    rolle.textContent = m.role === "user" ? "🧑" : "🤖";
    const bubble = document.createElement("div");
    bubble.className = "bubble";
    const txt = document.createElement("div");
    txt.textContent = m.text || "";
    bubble.appendChild(txt);
    for (const f of m.files || []) {
      const b = document.createElement("button");
      const da = createdFiles.has(f.id);
      b.className = "dl-knopf" + (da ? "" : " weg");
      b.textContent = (da ? "⬇️ " : "⚠️ ") + f.name + (da ? "" : " (neu erstellen)");
      if (da) b.onclick = () => downloadDatei(f.id);
      else b.title = "Nach dem Neuladen nicht mehr verfuegbar. Bitte die Datei neu erstellen lassen.";
      bubble.appendChild(b);
    }
    if (m.info) {
      const info = document.createElement("div");
      info.className = "info"; info.textContent = m.info;
      bubble.appendChild(info);
    }
    div.appendChild(rolle); div.appendChild(bubble);
    v.appendChild(div);
  }
  v.scrollTop = v.scrollHeight;
}

function renderAnhaenge() {
  const a = $("anhaenge");
  a.innerHTML = "";
  pendingAttachments.forEach((p, i) => {
    const chip = document.createElement("span");
    chip.className = "chip";
    chip.textContent = "📎 " + p.name + " ";
    const x = document.createElement("button");
    x.textContent = "✕";
    x.onclick = () => { pendingAttachments.splice(i, 1); renderAnhaenge(); };
    chip.appendChild(x);
    a.appendChild(chip);
  });
}

function renderTheorie() {
  const ul = $("theorie-liste");
  ul.innerHTML = "";
  state.theorie.forEach((t, i) => {
    const li = document.createElement("li");
    li.textContent = t.name;
    const x = document.createElement("button");
    x.textContent = "🗑"; x.title = "loeschen";
    x.onclick = () => { state.theorie.splice(i, 1); speichern(); renderTheorie(); };
    li.appendChild(x);
    ul.appendChild(li);
  });
}

function renderKosten() {
  $("kosten").innerHTML = `<b>Verbrauch diese Sitzung</b><br>ca. US$ ${state.kostenTotal.toFixed(3)} · ${state.tokensTotal.toLocaleString("de-CH")} Tokens`;
}

// --- Senden ---
function anzeigeText(text, anhaenge, theorie) {
  let t = text && text.trim() ? text : "(nur Datei(en) gesendet)";
  if (anhaenge.length) t += "\n\n📎 " + anhaenge.map((a) => a.name).join(", ");
  if (theorie) t += "\n\n(Theorie-Kontext mitgegeben)";
  return t;
}

async function senden() {
  if (busy) return;
  const eingabe = $("eingabe");
  const text = eingabe.value.trim();
  const anhaenge = pendingAttachments.slice();
  if (!state.apiKey) { setStatus("Bitte zuerst den API-Schluessel eintragen.", "fehler"); return; }
  if (!text && anhaenge.length === 0) return;

  const theorie = /theorie/i.test(text) ? theorieText() : "";
  const deskriptoren = anhaenge.map((a) => a.descriptor);
  const content = Core.baueUserContent(text, deskriptoren, theorie);
  state.apiMessages.push({ role: "user", content });
  state.displayMessages.push({ role: "user", text: anzeigeText(text, anhaenge, theorie), files: [] });

  eingabe.value = ""; eingabe.style.height = "auto";
  pendingAttachments = []; renderAnhaenge();
  renderVerlauf(); speichern();

  setBusy(true); setStatus("Die KI denkt nach ...");
  let ergebnis;
  try {
    ergebnis = await Core.fuehreKonversation(
      { callClaude, runCode, lintXlsx },
      state.model,
      state.apiMessages,
      (t) => setStatus(t)
    );
  } catch (e) {
    Core.bereinigeVerlauf(state.apiMessages);
    ergebnis = { text: deutscherFehler(e), dateien: {}, verbrauch: { input: 0, output: 0 } };
  }
  setBusy(false); setStatus("");

  const fileRefs = [];
  for (const [name, bytes] of Object.entries(ergebnis.dateien || {})) {
    fileRefs.push({ name, id: registriereDatei(name, bytes) });
  }
  const kosten = Core.geschaetzteKosten(state.model, ergebnis.verbrauch.input, ergebnis.verbrauch.output);
  const tokens = ergebnis.verbrauch.input + ergebnis.verbrauch.output;
  state.kostenTotal += kosten; state.tokensTotal += tokens;
  const info = tokens ? `ca. US$ ${kosten.toFixed(3)} · ${tokens.toLocaleString("de-CH")} Tokens` : "";
  state.displayMessages.push({ role: "assistant", text: ergebnis.text, files: fileRefs, info });

  renderVerlauf(); renderKosten(); speichern();
}

async function verbindungTesten() {
  if (!state.apiKey) { setTestStatus("Bitte zuerst einen Schluessel eingeben.", "fehler"); return; }
  setTestStatus("Teste Verbindung ...", "");
  try {
    const resp = await fetch(API_URL, {
      method: "POST", headers: headers(),
      body: JSON.stringify({ model: state.model, max_tokens: 16, messages: [{ role: "user", content: "Antworte nur mit OK" }] }),
    });
    if (resp.ok) { setTestStatus("Verbindung erfolgreich.", "ok"); return; }
    let msg = ""; try { msg = (await resp.json())?.error?.message || ""; } catch (_) {}
    const e = new Error(msg); e.status = resp.status;
    setTestStatus(deutscherFehler(e), "fehler");
  } catch (_) {
    setTestStatus("Keine Verbindung. Bitte Internet pruefen.", "fehler");
  }
}
function setTestStatus(text, klasse) {
  const s = $("test-status");
  s.textContent = text; s.className = "hinweis" + (klasse ? " " + klasse : "");
}

// --- Initialisierung ---
function fuelleModelle() {
  const sel = $("modell");
  sel.innerHTML = "";
  for (const [label, id] of Object.entries(Core.MODELL_OPTIONEN)) {
    const o = document.createElement("option");
    o.value = id; o.textContent = label;
    if (id === state.model) o.selected = true;
    sel.appendChild(o);
  }
}

function bindeEvents() {
  $("modell").addEventListener("change", (e) => { state.model = e.target.value; speichern(); });
  $("apikey").addEventListener("input", (e) => { state.apiKey = e.target.value.trim(); speichern(); renderVerlauf(); });
  $("test-btn").addEventListener("click", verbindungTesten);
  $("neu-btn").addEventListener("click", () => {
    if (!confirm("Neuen Chat starten? Der bisherige Verlauf wird geleert.")) return;
    state.apiMessages = []; state.displayMessages = []; state.kostenTotal = 0; state.tokensTotal = 0;
    createdFiles.clear();
    speichern(); renderVerlauf(); renderKosten();
  });
  $("senden-btn").addEventListener("click", senden);
  $("anhang-btn").addEventListener("click", () => $("datei-input").click());
  $("datei-input").addEventListener("change", async (e) => {
    for (const file of e.target.files) {
      const descriptor = await dateiZuDeskriptor(file);
      pendingAttachments.push({ descriptor, name: file.name });
    }
    e.target.value = "";
    renderAnhaenge();
  });
  $("theorie-input").addEventListener("change", async (e) => {
    for (const file of e.target.files) {
      const d = await dateiZuDeskriptor(file);
      const text = d.kind === "text" ? d.text : "(Datei: " + file.name + ")";
      state.theorie.push({ name: file.name, text });
    }
    e.target.value = "";
    speichern(); renderTheorie();
  });
  const eingabe = $("eingabe");
  eingabe.addEventListener("input", () => {
    eingabe.style.height = "auto";
    eingabe.style.height = Math.min(eingabe.scrollHeight, 200) + "px";
  });
  eingabe.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); senden(); }
  });
}

function pruefeBibliotheken() {
  const fehlt = [];
  if (!window.XLSX) fehlt.push("XLSX");
  if (!window.docx) fehlt.push("docx");
  if (!window.PptxGenJS) fehlt.push("PptxGenJS");
  if (!window.jspdf) fehlt.push("jspdf");
  if (!window.mammoth) fehlt.push("mammoth");
  if (fehlt.length) setStatus("Hinweis: Bibliotheken nicht geladen (" + fehlt.join(", ") + "). Datei-Erstellung evtl. eingeschraenkt.", "fehler");
}

function init() {
  laden();
  fuelleModelle();
  $("apikey").value = state.apiKey;
  bindeEvents();
  renderVerlauf(); renderAnhaenge(); renderTheorie(); renderKosten();
  pruefeBibliotheken();
}

// Fuer Tests/Automation zugaenglich machen.
window.__qv = { state, senden, Core };

if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
else init();
