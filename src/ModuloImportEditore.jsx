import { useState, useMemo, useCallback, useEffect } from "react";
import { resolveGiri } from "./ModuloImport.jsx";
import { tema, cv } from "./tema.js";

// ═══════════════════════════════════════════════════════════════════════════
// TRASCINAMENTO FILE GLOBALE
// Rende "trascinabili" tutti i caricamenti file di GiroManager senza toccare i
// singoli moduli: il file lasciato cadere viene passato al campo file più vicino
// al punto di rilascio (o all'unico visibile a schermo), come se fosse stato
// scelto con "Sfoglia". I moduli che gestiscono già il drop da soli
// (es. Import da file editore) hanno la precedenza.
// ═══════════════════════════════════════════════════════════════════════════

const CLASSE = "gm-drop-target";

function iniettaStile() {
  if (document.getElementById("gm-drop-style")) return;
  const st = document.createElement("style");
  st.id = "gm-drop-style";
  st.textContent = `.${CLASSE}{outline:2px dashed #c8a96e!important;outline-offset:4px;border-radius:4px;background-color:#c8a96e14!important;transition:all .15s}`;
  document.head.appendChild(st);
}

const haFile = (e) => Array.from(e.dataTransfer?.types || []).includes("Files");
const visibile = (el) => !!el && el.getClientRects().length > 0;
// un input nascosto (display:none) conta come visibile se lo è il suo contenitore / label
const inputUsabile = (inp) => !inp.disabled && (visibile(inp) || visibile(inp.parentElement) || (inp.id && visibile(document.querySelector(`label[for="${CSS.escape(inp.id)}"]`))));

function trovaInput(x, y) {
  // 1) il campo file più vicino risalendo dal punto di rilascio
  let el = document.elementFromPoint(x, y);
  while (el && el !== document.body) {
    const inp = Array.from(el.querySelectorAll?.('input[type="file"]') || []).filter(inputUsabile);
    if (inp.length) return inp[0];
    el = el.parentElement;
  }
  // 2) altrimenti l'unico campo file presente a schermo
  const tutti = Array.from(document.querySelectorAll('input[type="file"]')).filter(inputUsabile);
  return tutti.length === 1 ? tutti[0] : null;
}

function accetta(inp, file) {
  const acc = (inp.getAttribute("accept") || "").split(",").map(s => s.trim().toLowerCase()).filter(Boolean);
  if (!acc.length) return true;
  const nome = file.name.toLowerCase(), tipo = (file.type || "").toLowerCase();
  return acc.some(a => a.startsWith(".") ? nome.endsWith(a) : a.endsWith("/*") ? tipo.startsWith(a.slice(0, -1)) : tipo === a);
}

const contenitore = (inp) => (inp.id && document.querySelector(`label[for="${CSS.escape(inp.id)}"]`)?.parentElement) || inp.parentElement;

function attivaDragDropGlobale() {
  if (window.__gmDragDrop) return;
  window.__gmDragDrop = true;
  iniettaStile();
  let evidenziato = null;
  const evidenzia = (el) => {
    if (evidenziato === el) return;
    evidenziato?.classList.remove(CLASSE);
    evidenziato = el;
    el?.classList.add(CLASSE);
  };

  window.addEventListener("dragover", (e) => {
    if (!haFile(e)) return;
    e.preventDefault(); // evita che il browser apra il file al posto dell'app
    const inp = trovaInput(e.clientX, e.clientY);
    e.dataTransfer.dropEffect = inp ? "copy" : "none";
    evidenzia(inp ? contenitore(inp) : null);
  });
  window.addEventListener("dragleave", (e) => { if (!e.relatedTarget) evidenzia(null); });
  window.addEventListener("drop", (e) => {
    if (!haFile(e)) return;
    const giaGestito = e.defaultPrevented; // un modulo con drop proprio l'ha già preso
    e.preventDefault();
    evidenzia(null);
    if (giaGestito) return;
    const inp = trovaInput(e.clientX, e.clientY);
    if (!inp) return;
    const files = Array.from(e.dataTransfer.files || []).filter(f => accetta(inp, f));
    if (!files.length) { alert(`Formato non supportato. File accettati: ${inp.getAttribute("accept") || "qualsiasi"}`); return; }
    const dt = new DataTransfer();
    (inp.multiple ? files : files.slice(0, 1)).forEach(f => dt.items.add(f));
    inp.files = dt.files;
    inp.dispatchEvent(new Event("change", { bubbles: true }));
  });
}

// attivo una volta sola all'avvio di GiroManager
attivaDragDropGlobale();

// ═══════════════════════════════════════════════════════════════════════════
// IMPORT CEDOLA DA FILE EDITORE
// Legge file "liberi" inviati dagli editori (xlsx/xls/csv, qualsiasi layout):
// - individua da solo foglio, riga intestazioni e colonne (modificabili a mano)
// - abbina i nomi editore all'anagrafica ranking_editori anche se scritti diversi
//   (fuzzy + alias salvati in `alias_editori` + storico EAN via RPC `prefissi_ean_editori`)
// - mantiene l'ordine del file (posizione per giro+editore) e prende il ranking da Supabase
// - al re-import non cancella i campi già presenti che il file non fornisce
// ═══════════════════════════════════════════════════════════════════════════

const SUPABASE_URL = "https://tdflwenlylhctxssatax.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InRkZmx3ZW5seWxoY3R4c3NhdGF4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYzMzgyNzYsImV4cCI6MjA5MTkxNDI3Nn0.l35qEL7LOvyYuI1McQlVqj4vbyTqmlevcmqWbTGYi2Q";

const T = tema({
  bg: "#0f0f0f", surface: "#161616", border: "#252525", borderHi: "#333333",
  text: "#e8e8e8", textMid: "#888888", textDim: "#444444",
  accent: "#c8a96e", green: "#4caf7d", red: "#e05c5c", blue: "#5b8fd4", orange: "#e0a24c",
});
const css = {
  btn: (v = "default") => ({ padding: "6px 14px", border: `1px solid ${v === "accent" ? T.accent : v === "danger" ? T.red : T.border}`, background: v === "accent" ? T.accent : v === "danger" ? T.red + "22" : "transparent", color: v === "accent" ? "#000" : v === "danger" ? T.red : T.text, cursor: "pointer", fontSize: "12px", fontFamily: "inherit", borderRadius: 3, fontWeight: v === "accent" ? "700" : "400" }),
  th: { padding: "8px 10px", textAlign: "left", color: T.textMid, fontWeight: "400", fontSize: "11px", letterSpacing: "0.08em", textTransform: "uppercase", borderBottom: `1px solid ${T.border}`, whiteSpace: "nowrap", background: T.surface, position: "sticky", top: 0 },
  td: { padding: "6px 10px", borderBottom: `1px solid ${T.border}22`, verticalAlign: "middle", fontSize: "12px" },
  input: { background: T.bg, border: `1px solid ${T.borderHi}`, color: T.text, padding: "5px 8px", fontSize: "12px", fontFamily: "inherit", borderRadius: 3 },
  box: (c = T.border) => ({ background: T.surface, border: `1px solid ${c}`, borderRadius: 4, padding: "10px 16px" }),
  panel: { background: T.surface, border: `1px solid ${T.border}`, borderRadius: 4, padding: 16, marginBottom: 16 },
};

const hdr = (token) => ({ apikey: SUPABASE_KEY, Authorization: `Bearer ${token}` });

// ─── Campi riconosciuti ──────────────────────────────────────────────────────
const CAMPI = [
  { key: "ean", label: "EAN", req: true },
  { key: "titolo", label: "Titolo", req: true },
  { key: "autore", label: "Autore" },
  { key: "editore", label: "Editore / marchio" },
  { key: "prezzo", label: "Prezzo" },
  { key: "uscita", label: "Uscita" },
  { key: "top_100", label: "Top 100" },
  { key: "obiettivo", label: "Obiettivo" },
  { key: "tiratura", label: "Tiratura" },
  { key: "note", label: "Note" },
  { key: "gem_ean_1", label: "EAN gemello 1" }, { key: "gem_tit_1", label: "Titolo gemello 1" },
  { key: "gem_ean_2", label: "EAN gemello 2" }, { key: "gem_tit_2", label: "Titolo gemello 2" },
  { key: "gem_ean_3", label: "EAN gemello 3" }, { key: "gem_tit_3", label: "Titolo gemello 3" },
  { key: "gemelli_testo", label: "Gemelli (testo libero)" },
];

// Colonne dell'inserimento manuale: k = campo della riga, map = campo dell'import, w = larghezza
const CAMPI_MANUALI = [
  { k: "ean", label: "EAN", map: "ean", w: 130 },
  { k: "titolo", label: "Titolo", map: "titolo", w: 220 },
  { k: "autore", label: "Autore", map: "autore", w: 150 },
  { k: "editore", label: "Editore", map: "editore", w: 170 },
  { k: "prezzo", label: "Prezzo", map: "prezzo", w: 70 },
  { k: "obiettivo", label: "Obiettivo", map: "obiettivo", w: 80 },
  { k: "note", label: "Note", map: "note", w: 200 },
  { k: "g1e", label: "EAN gemello 1", map: "gem_ean_1", w: 130, gem: 1 },
  { k: "g1t", label: "Titolo gemello 1", map: "gem_tit_1", w: 170, gem: 1 },
  { k: "g2e", label: "EAN gemello 2", map: "gem_ean_2", w: 130, gem: 2 },
  { k: "g2t", label: "Titolo gemello 2", map: "gem_tit_2", w: 170, gem: 2 },
  { k: "g3e", label: "EAN gemello 3", map: "gem_ean_3", w: 130, gem: 3 },
  { k: "g3t", label: "Titolo gemello 3", map: "gem_tit_3", w: 170, gem: 3 },
];
const rigaManualeVuota = () => Object.fromEntries(CAMPI_MANUALI.map(c => [c.k, ""]));
const NOME_MANUALE = "Inserimento manuale";

// Riconoscimento intestazioni (su testo normalizzato minuscolo senza accenti)
const RX = {
  gemello: /(gemell|\bgem\b|twin|abbinat|collegat|correlat|comparab|affin|titol[oi] (simil|di riferimento)|(ean|isbn) (simil|di riferimento))/,
  ean: /\b(ean|isbn|ean ?13|isbn ?13|barcode|codice a barre)\b/,
  titolo: /^(titolo|title|titoli)\b|titolo (opera|libro|volume)/,
  autore: /(autor|author|a cura)/,
  editore: /(editore|marchio|casa editrice|publisher|sigla|imprint|brand|editrice)/,
  prezzo: /(prezzo|pvp|p v p|price|copertina|euro)/,
  uscita: /(uscita|pubblicazione|release|data pub|mese)/,
  top_100: /(top ?100|\btop\b|punta|di punta)/,
  obiettivo: /(obiettiv|\bobj\b|target)/,
  tiratura: /(tiratura|\btir\b|prima tiratura)/,
  note: /(\bnot[ae]\b|comment|osservaz|argoment|info|descrizione|sinossi)/,
};

const MESI = ["GENNAIO", "FEBBRAIO", "MARZO", "APRILE", "MAGGIO", "GIUGNO", "LUGLIO", "AGOSTO", "SETTEMBRE", "OTTOBRE", "NOVEMBRE", "DICEMBRE"];

// ─── Utility testo / numeri ─────────────────────────────────────────────────
const deaccent = (s) => String(s ?? "").normalize("NFD").replace(/[̀-ͯ]/g, "");
const normHeader = (s) => deaccent(s).toLowerCase().replace(/[^a-z0-9€]+/g, " ").trim();
const cleanText = (v) => {
  const s = String(v ?? "").replace(/[ \s]+/g, " ").trim();
  return s === "" ? null : s.toUpperCase();
};
const colLetter = (i) => { let s = ""; i += 1; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };

function isbn10to13(d) {
  const core = "978" + d.slice(0, 9);
  let sum = 0;
  for (let i = 0; i < 12; i++) sum += parseInt(core[i]) * (i % 2 ? 3 : 1);
  return core + ((10 - (sum % 10)) % 10);
}
function parseEan(v) {
  if (v === null || v === undefined || v === "") return null;
  let s = typeof v === "number" ? Math.round(v).toString() : String(v).trim();
  if (/^\d+(\.\d+)?e\+?\d+$/i.test(s)) s = Math.round(Number(s)).toString(); // notazione scientifica da CSV
  const d = s.replace(/[^0-9Xx]/g, "");
  if (/^97[89]\d{10}$/.test(d)) return d;
  if (/^\d{9}[\dXx]$/.test(d)) return isbn10to13(d);
  if (/^\d{13}$/.test(d)) return d;
  return null;
}
const eanChecksumOk = (e) => {
  let sum = 0;
  for (let i = 0; i < 12; i++) sum += parseInt(e[i]) * (i % 2 ? 3 : 1);
  return (10 - (sum % 10)) % 10 === parseInt(e[12]);
};
function parsePrezzo(v) {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return Math.round(v * 100) / 100;
  let s = String(v).replace(/[€\s]|eur(o)?/gi, "");
  if (s.includes(",")) s = s.replace(/\./g, "").replace(",", ".");
  const n = parseFloat(s);
  return isNaN(n) ? null : Math.round(n * 100) / 100;
}
function parseIntero(v) {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return Math.round(v);
  const d = String(v).replace(/[^\d]/g, "");
  return d ? parseInt(d) : null;
}
function parseTop(v) {
  if (v === null || v === undefined || v === "") return false;
  if (typeof v === "number") return v > 0;
  if (typeof v === "boolean") return v;
  const s = deaccent(v).trim().toUpperCase();
  return /^(SI|S|X|1|TRUE|VERO|YES|Y|★|\*+|TOP|TOP ?100|OK|V|✓|✔)$/.test(s) || /^\d+$/.test(s);
}
function parseUscita(v) {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number" && v > 20000 && v < 80000 && window.XLSX?.SSF) {
    const d = window.XLSX.SSF.parse_date_code(v);
    return d ? MESI[d.m - 1] : null;
  }
  if (v instanceof Date) return MESI[v.getMonth()];
  const s = deaccent(v).trim().toUpperCase();
  const mese = MESI.find(m => s.includes(m) || s.includes(m.slice(0, 3) + " ") || s === m.slice(0, 3));
  if (mese) return mese;
  const m = s.match(/^\d{1,2}[/.-](\d{1,2})[/.-]\d{2,4}$/) || s.match(/^(\d{1,2})[/.-]\d{4}$/);
  if (m) { const n = parseInt(m[1]); if (n >= 1 && n <= 12) return MESI[n - 1]; }
  return s || null;
}
// Cella gemelli in qualsiasi forma → [{ean, titolo}]
// "9788845912345" · "9788845912345 - Titolo A" · "Titolo A (978-88-459-1234-5)"
// "978... Titolo A / 978... Titolo B" · con a capo, ; | / , come separatori
// Il testo tra due EAN viene attribuito all'EAN "più vicino": prima quello che segue
// l'EAN (formato EAN + titolo), altrimenti quello che lo precede (formato titolo + EAN).
const RX_EAN_IN_TESTO = /97[89](?:[\s-]?\d){10}/g;
const pulisciTitoloGemello = (s) => cleanText(String(s ?? "").replace(/^[\s\-–—:;,.()/|\[\]]+|[\s\-–—:;,.()/|\[\]]+$/g, "").replace(/\b(ean|isbn)\b\s*:?/gi, " "));
function estraiGemelliCella(v) {
  if (v === null || v === undefined || v === "") return [];
  if (typeof v === "number") { const e = parseEan(v); return e ? [{ ean: e, titolo: null }] : []; }
  const s = String(v);
  const hits = [...s.matchAll(RX_EAN_IN_TESTO)];
  if (!hits.length) return [];
  const out = hits.map(h => ({ ean: parseEan(h[0]), start: h.index, end: h.index + h[0].length, titolo: null })).filter(x => x.ean);
  if (!out.length) return [];
  const segDopo = out.map((x, i) => pulisciTitoloGemello(s.slice(x.end, out[i + 1]?.start ?? s.length)));
  const segPrima = pulisciTitoloGemello(s.slice(0, out[0].start));
  if (!segPrima) out.forEach((x, i) => { x.titolo = segDopo[i]; });          // EAN Titolo, EAN Titolo
  else out.forEach((x, i) => { x.titolo = i === 0 ? segPrima : segDopo[i - 1]; }); // Titolo EAN, Titolo EAN
  return out.map(({ ean, titolo }) => ({ ean, titolo: titolo || null }));
}
// Colonna "titolo gemello" che per errore contiene anche l'EAN: tiene solo il testo
const soloTitolo = (v) => {
  if (v === null || v === undefined || v === "" || typeof v === "number") return null;
  return pulisciTitoloGemello(String(v).replace(RX_EAN_IN_TESTO, " ")) || null;
};

// ─── Matching editori ────────────────────────────────────────────────────────
const GENERICHE = new Set(["EDIZIONI", "EDIZIONE", "EDITORE", "EDITORI", "EDITRICE", "EDITORIALE", "CASA", "SRL", "SRLS", "SPA", "SAS", "SNC", "AD", "ED", "EDIT", "GRUPPO", "LIBRI", "PUBLISHING", "BOOKS", "ITALIA", "ITALY", "ITALIANA", "IL", "LO", "LA", "GLI", "LE", "I", "L", "DI", "DEL", "DELLA", "DEI", "DEGLI", "D", "E", "AND", "THE"]);
// "S.R.L." / "S.P.A." ecc. vanno tolti come sigla intera, non lettera per lettera (ADELPHI_R ≠ ADELPHI)
const normEd = (s) => deaccent(s).toUpperCase().replace(/[^A-Z0-9]+/g, " ").replace(/\bS R L S?\b|\bS P A\b|\bS A S\b|\bS N C\b/g, " ").replace(/\s+/g, " ").trim();
const coreTokens = (s) => normEd(s).split(" ").filter(t => t && !GENERICHE.has(t));

function bigrams(s) { const r = []; for (let i = 0; i < s.length - 1; i++) r.push(s.slice(i, i + 2)); return r; }
function dice(a, b) {
  if (!a || !b) return 0;
  if (a === b) return 1;
  const A = bigrams(a), B = bigrams(b);
  const m = new Map(); A.forEach(x => m.set(x, (m.get(x) || 0) + 1));
  let inter = 0; B.forEach(x => { const c = m.get(x); if (c) { inter++; m.set(x, c - 1); } });
  return (2 * inter) / (A.length + B.length || 1);
}
function levRatio(a, b) {
  const dp = Array.from({ length: a.length + 1 }, (_, i) => [i]);
  for (let j = 1; j <= b.length; j++) dp[0][j] = j;
  for (let i = 1; i <= a.length; i++) for (let j = 1; j <= b.length; j++)
    dp[i][j] = Math.min(dp[i - 1][j] + 1, dp[i][j - 1] + 1, dp[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
  return 1 - dp[a.length][b.length] / Math.max(a.length, b.length, 1);
}
function scoreEditore(raw, anag) {
  const lit = (x) => deaccent(x).toUpperCase().replace(/\s+/g, " ").trim();
  if (lit(raw) && lit(raw) === lit(anag)) return 1.01; // identico lettera per lettera: vince su varianti (+, _R, ...)
  const na = normEd(raw), nb = normEd(anag);
  if (!na || !nb) return 0;
  if (na === nb) return 1;
  const ta = coreTokens(raw), tb = coreTokens(anag);
  const ja = ta.join(" "), jb = tb.join(" ");
  if (ja && ja === jb) return 0.96;
  if (ja && ja.replace(/ /g, "") === jb.replace(/ /g, "")) return 0.95; // HARPERCOLLINS vs HARPER COLLINS
  if (!ta.length || !tb.length) return dice(na, nb) * 0.7;
  // un nome contiene interamente l'altro (es. GALLUCCI ⊂ GALLUCCI BROS): buono ma non conclusivo
  const small = ta.length <= tb.length ? ta : tb, large = ta.length <= tb.length ? tb : ta;
  if (small.every(t => large.includes(t))) return 0.8 + 0.1 * (small.length / large.length);
  const sa = ja.replace(/ /g, ""), sb = jb.replace(/ /g, "");
  if (Math.min(sa.length, sb.length) >= 5 && (sa.includes(sb) || sb.includes(sa))) return 0.82;
  // match token con tolleranza refusi
  let hit = 0;
  ta.forEach(t => { if (tb.some(u => u === t || (t.length > 3 && u.length > 3 && levRatio(t, u) >= 0.8))) hit++; });
  const tokRatio = hit / Math.max(ta.length, tb.length);
  return Math.min(0.88, 0.55 * dice(ja, jb) + 0.45 * tokRatio);
}
function candidati(raw, anagrafica) {
  return anagrafica
    .map(a => ({ nome: a.editore_nome, score: scoreEditore(raw, a.editore_nome) }))
    .filter(c => c.score >= 0.4)
    .sort((x, y) => y.score - x.score)
    .slice(0, 6);
}

// Prefissi EAN → editore unico (≥90% delle occorrenze, almeno 2 titoli)
function buildPrefixIndex(rows) {
  const tot = {}, best = {};
  rows.forEach(({ prefisso, editore_nome, n }) => {
    n = Number(n);
    tot[prefisso] = (tot[prefisso] || 0) + n;
    if (!best[prefisso] || n > best[prefisso].n) best[prefisso] = { editore_nome, n };
  });
  const idx = {};
  Object.keys(best).forEach(p => { if (tot[p] >= 2 && best[p].n / tot[p] >= 0.9) idx[p] = best[p].editore_nome; });
  return idx;
}
function editoreDaEan(ean, storicoEan, prefixIdx) {
  if (!ean) return null;
  if (storicoEan[ean]) return { nome: storicoEan[ean], fonte: "ean" };
  for (let l = 10; l >= 7; l--) { const e = prefixIdx[ean.slice(0, l)]; if (e) return { nome: e, fonte: "prefisso" }; }
  return null;
}

// ─── Analisi foglio ─────────────────────────────────────────────────────────
function rilevaIntestazioni(aoa) {
  let best = { row: -1, score: 0 };
  const lim = Math.min(aoa.length, 40);
  for (let r = 0; r < lim; r++) {
    const hs = (aoa[r] || []).map(normHeader);
    const found = new Set();
    hs.forEach(h => {
      if (!h) return;
      Object.entries(RX).forEach(([k, rx]) => { if (k !== "gemello" && rx.test(h)) found.add(k); });
    });
    // l'intestazione vera ha di solito EAN/ISBN o TITOLO
    const score = found.size + (found.has("ean") ? 2 : 0) + (found.has("titolo") ? 1 : 0);
    if (score > best.score) best = { row: r, score };
  }
  return best.score >= 3 ? best.row : -1;
}

function rilevaMapping(aoa, headerRow) {
  const ncol = Math.max(0, ...aoa.slice(0, 200).map(r => (r || []).length));
  const headers = Array.from({ length: ncol }, (_, i) => headerRow >= 0 ? String(aoa[headerRow]?.[i] ?? "").trim() : "");
  const map = {};
  CAMPI.forEach(c => { map[c.key] = -1; });
  const used = new Set();
  const gemCols = [];
  headers.forEach((h, i) => {
    const n = normHeader(h);
    if (!n) return;
    if (RX.gemello.test(n)) { gemCols.push({ i, isEan: /(ean|isbn|cod)/.test(n) }); used.add(i); }
  });
  // ordine di priorità: ean e titolo prima, "note" per ultimo (più generico)
  ["ean", "titolo", "tiratura", "obiettivo", "top_100", "prezzo", "autore", "editore", "uscita", "note"].forEach(k => {
    const i = headers.findIndex((h, idx) => !used.has(idx) && RX[k].test(normHeader(h)) && !(k === "editore" && /cod/.test(normHeader(h))) && !(k === "titolo" && /sotto/.test(normHeader(h))));
    if (i >= 0) { map[k] = i; used.add(i); }
  });
  // gemelli
  let ne = 0, nt = 0;
  const dati = aoa.slice(headerRow + 1, headerRow + 200);
  const conEan = (i) => dati.filter(r => estraiGemelliCella(r?.[i]).length).length;
  const testuale = (i) => {
    const v = dati.map(r => r?.[i]).filter(x => typeof x === "string" && x.trim());
    return v.length > 0 && v.filter(x => !estraiGemelliCella(x).length && !parseEan(x)).length >= v.length * 0.6;
  };
  // Dopo i campi del titolo (EAN, titolo, autore, editore, prezzo…), ogni altra colonna con EAN
  // seguita da una colonna di titolo è un gemello, anche senza la parola "gemello" nell'intestazione
  // (es. "ISBN13 | Titolo | … | EAN | TITOLO | EAN | TITOLO" o "ean gem | titolo gem")
  const pieni = (i) => dati.map(r => r?.[i]).filter(x => x !== "" && x !== null && x !== undefined).length;
  const extraEan = headers.map((_, i) => i).filter(i => {
    if (used.has(i) || i === map.ean || /ebook|epub|digital/.test(normHeader(headers[i]))) return false;
    const n = conEan(i);
    return n > 0 && (RX.ean.test(normHeader(headers[i])) || n >= pieni(i) * 0.5);
  });
  const extraTit = [];
  extraEan.forEach(i => {
    const j = [i + 1, i - 1].find(k => k >= 0 && k < ncol && !used.has(k) && !extraEan.includes(k) && !extraTit.includes(k)
      && (RX.titolo.test(normHeader(headers[k])) || (k === i + 1 && testuale(k))));
    if (j !== undefined) extraTit.push(j);
  });
  [...extraEan.map(i => ({ i, isEan: true })), ...extraTit.map(i => ({ i, isEan: false }))].forEach(g => { gemCols.push(g); used.add(g.i); });
  gemCols.sort((a, b) => a.i - b.i);
  if (gemCols.length === 1) {
    // colonna unica: EAN, titolo o entrambi nella stessa cella → parser misto
    map.gemelli_testo = gemCols[0].i;
  } else {
    gemCols.forEach(g => {
      // il tipo lo decide il contenuto, non solo l'intestazione ("Gemello 1" può contenere EAN)
      const isEan = conEan(g.i) > 0 && (g.isEan || !testuale(g.i));
      if (isEan && ne < 3) map[`gem_ean_${++ne}`] = g.i;
      else if (!isEan && nt < 3) map[`gem_tit_${++nt}`] = g.i;
    });
  }
  // EAN gemello senza colonna titolo: prova la colonna accanto (destra, poi sinistra),
  // es. "EAN GEMELLO | TITOLO" dove la seconda intestazione non dice "gemello"
  const giaMappate = () => new Set(Object.values(map).filter(i => i >= 0));
  const unica = map.gemelli_testo >= 0 && ne === 0 ? [["gemelli_testo", "gem_tit_1"]] : [];
  [...[1, 2, 3].map(k => [`gem_ean_${k}`, `gem_tit_${k}`]), ...unica].forEach(([ke, kt]) => {
    const ie = map[ke];
    if (!(ie >= 0) || map[kt] >= 0) return;
    // se le celle hanno già EAN + titolo insieme non serve cercare altrove
    const celle = dati.map(r => estraiGemelliCella(r?.[ie])).filter(x => x.length);
    if (!celle.length || celle.filter(x => x.every(g => !g.titolo)).length < celle.length * 0.6) return;
    const occ = giaMappate();
    const adj = [ie + 1, ie - 1].find(i => i >= 0 && i < ncol && !occ.has(i) && !used.has(i) && testuale(i));
    if (adj !== undefined) { map[kt] = adj; used.add(adj); }
  });
  // fallback EAN da contenuto
  if (map.ean < 0) {
    let bestI = -1, bestN = 0;
    for (let i = 0; i < ncol; i++) {
      if (used.has(i)) continue;
      const n = dati.filter(r => parseEan(r?.[i]) && /^97[89]/.test(parseEan(r?.[i]))).length;
      if (n > bestN) { bestN = n; bestI = i; }
    }
    if (bestI >= 0) { map.ean = bestI; used.add(bestI); }
  }
  // fallback titolo: colonna testuale più "lunga" tra le libere
  if (map.titolo < 0) {
    let bestI = -1, bestL = 0;
    for (let i = 0; i < ncol; i++) {
      if (used.has(i)) continue;
      const vals = dati.map(r => r?.[i]).filter(v => typeof v === "string" && v.trim() && !parseEan(v));
      const avg = vals.reduce((s, v) => s + v.length, 0) / (vals.length || 1);
      if (vals.length > dati.length * 0.3 && avg > bestL) { bestL = avg; bestI = i; }
    }
    if (bestI >= 0) map.titolo = bestI;
  }
  return { headers, map };
}

function estraiRighe(aoa, headerRow, map) {
  const get = (r, k) => (map[k] >= 0 ? r[map[k]] : null);
  const righe = [], scartate = [];
  let sezione = null; // riga "titolo di sezione" (es. nome marchio) sopra un gruppo di titoli
  for (let r = headerRow + 1; r < aoa.length; r++) {
    const row = aoa[r] || [];
    const pieni = row.map((v, i) => ({ v, i })).filter(x => x.v !== "" && x.v !== null && x.v !== undefined);
    if (!pieni.length) continue;
    const ean = parseEan(get(row, "ean"));
    if (!ean) {
      if (pieni.length <= 2 && typeof pieni[0].v === "string" && pieni[0].v.trim().length <= 60) sezione = pieni[0].v.trim();
      else if (get(row, "titolo")) scartate.push({ riga: r + 1, motivo: "EAN mancante o non valido", testo: String(get(row, "titolo")).slice(0, 60) });
      continue;
    }
    const gemelli = [];
    for (let g = 1; g <= 3; g++) {
      if (g === 1 && map.gem_ean_1 < 0 && map.gemelli_testo >= 0) continue; // titolo accanto alla colonna unica: gestito sotto
      const vE = get(row, `gem_ean_${g}`), vT = get(row, `gem_tit_${g}`);
      // la cella EAN può contenere anche il titolo (o più gemelli); la colonna titolo può contenere anche l'EAN
      const daE = estraiGemelliCella(vE);
      const daT = daE.length ? [] : estraiGemelliCella(vT);
      const items = daE.length ? daE : daT;
      const tit = soloTitolo(vT);
      if (!items.length) { if (tit) gemelli.push({ ean: null, titolo: tit }); continue; }
      items.forEach((x, j) => gemelli.push({ ean: x.ean, titolo: (j === 0 && tit) || x.titolo || null }));
    }
    const vTxt = get(row, "gemelli_testo");
    const daTxt = estraiGemelliCella(vTxt);
    if (daTxt.length) {
      // titolo da colonna accanto (se rilevata) solo per il primo gemello senza titolo
      const titAdj = map.gem_ean_1 >= 0 ? null : soloTitolo(get(row, "gem_tit_1"));
      daTxt.forEach((x, j) => gemelli.push({ ean: x.ean, titolo: x.titolo || (j === 0 ? titAdj : null) }));
    } else if (map.gem_ean_1 < 0 && map.gem_tit_1 < 0 && cleanText(vTxt)) {
      gemelli.push({ ean: null, titolo: cleanText(vTxt) });
    }
    const obiettivo = parseIntero(get(row, "obiettivo"));
    const tiratura = parseIntero(get(row, "tiratura"));
    righe.push({
      _riga: r + 1,
      ean,
      eanWarn: !eanChecksumOk(ean),
      titolo: cleanText(get(row, "titolo")),
      autore: cleanText(get(row, "autore")),
      rawEditore: String(get(row, "editore") ?? "").replace(/[ \s]+/g, " ").trim() || null,
      sezione,
      prezzo: parsePrezzo(get(row, "prezzo")),
      uscita: parseUscita(get(row, "uscita")),
      // cella vuota = dato assente (al re-import resta il Top 100 già salvato)
      top_100: map.top_100 >= 0 && String(get(row, "top_100") ?? "").trim() !== "" ? parseTop(get(row, "top_100")) : null,
      obiettivo, tiratura,
      obiettivoFonte: obiettivo ? "O" : tiratura ? "T" : null,
      note: cleanText(get(row, "note")),
      gemelli: gemelli.slice(0, 3),
    });
  }
  return { righe, scartate };
}

// Foglio con più EAN validi = foglio dei titoli
function sceltaFoglio(wb) {
  let best = null;
  wb.SheetNames.forEach(n => {
    const aoa = window.XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: "", raw: true });
    const nEan = aoa.reduce((s, r) => s + ((r || []).some(v => { const e = parseEan(v); return e && /^97[89]/.test(e); }) ? 1 : 0), 0);
    if (!best || nEan > best.nEan) best = { nome: n, nEan };
  });
  return best?.nome ?? wb.SheetNames[0];
}

// ─── Fetch Supabase ─────────────────────────────────────────────────────────
async function fetchJson(url, token, opts = {}) {
  const res = await fetch(url, { ...opts, headers: { ...hdr(token), ...(opts.headers || {}) } });
  if (!res.ok) throw new Error(`HTTP ${res.status} — ${(await res.text()).slice(0, 200)}`);
  const t = await res.text();
  return t ? JSON.parse(t) : null;
}
async function fetchStoricoEan(token, eans) {
  const out = {};
  const list = [...new Set(eans)];
  for (let i = 0; i < list.length; i += 150) {
    const chunk = list.slice(i, i + 150);
    const rows = await fetchJson(`${SUPABASE_URL}/rest/v1/titoli?select=ean,editore_nome&ean=in.(${chunk.join(",")})&order=updated_at.desc`, token);
    rows.forEach(r => { if (r.editore_nome && !out[r.ean]) out[r.ean] = r.editore_nome; });
  }
  return out;
}
// Titoli già noti in Supabase per gli EAN gemelli arrivati senza titolo nel file
async function fetchTitoliNoti(token, eans) {
  const out = {};
  const list = [...new Set(eans)].filter(Boolean);
  for (let i = 0; i < list.length; i += 150) {
    const chunk = list.slice(i, i + 150).join(",");
    const q = [
      `titoli?select=ean,titolo&ean=in.(${chunk})&titolo=not.is.null`,
      `titoli_novita?select=ean,titolo&ean=in.(${chunk})&titolo=not.is.null`,
      ...[1, 2, 3].map(k => `titoli?select=ean:ean_gemello_${k},titolo:titolo_gemello_${k}&ean_gemello_${k}=in.(${chunk})&titolo_gemello_${k}=not.is.null`),
    ];
    const res = await Promise.all(q.map(u => fetchJson(`${SUPABASE_URL}/rest/v1/${u}`, token).catch(() => [])));
    res.flat().forEach(r => { if (r?.ean && r.titolo && !out[r.ean]) out[r.ean] = String(r.titolo).replace(/\s+/g, " ").trim().toUpperCase(); });
  }
  return out;
}
async function fetchEsistentiExtra(token, nCedola, eans) {
  const out = {};
  const list = [...new Set(eans)];
  const nome = encodeURIComponent(nCedola);
  for (let i = 0; i < list.length; i += 150) {
    const chunk = list.slice(i, i + 150);
    const rows = await fetchJson(`${SUPABASE_URL}/rest/v1/titoli?select=*&giro_label=eq.EXTRA&n_cedola=eq.${nome}&ean=in.(${chunk.join(",")})&order=id.asc`, token);
    rows.forEach(r => { if (!out[r.ean]) out[r.ean] = r; });
  }
  return out;
}
async function fetchEsistenti(token, giroIds, eans) {
  const out = {};
  const list = [...new Set(eans)];
  if (!giroIds.length) return out;
  for (let i = 0; i < list.length; i += 150) {
    const chunk = list.slice(i, i + 150);
    const rows = await fetchJson(`${SUPABASE_URL}/rest/v1/titoli?select=*&giro_id=in.(${giroIds.join(",")})&ean=in.(${chunk.join(",")})`, token);
    rows.forEach(r => { out[`${r.giro_id}|${r.ean}`] = r; });
  }
  return out;
}

// ─── Componente ─────────────────────────────────────────────────────────────
export default function ModuloImportEditore({ token, onImportDone, manuale = false }) {
  const [step, setStep] = useState("upload"); // upload | preview | result
  const [loading, setLoading] = useState(false);
  const [fileName, setFileName] = useState("");
  const [wb, setWb] = useState(null);
  const [foglio, setFoglio] = useState("");
  const [headerRow, setHeaderRow] = useState(-1);
  const [headers, setHeaders] = useState([]);
  const [map, setMap] = useState({});
  const [aoa, setAoa] = useState([]);
  const [anagrafica, setAnagrafica] = useState([]);
  const [alias, setAlias] = useState({});
  const [prefixIdx, setPrefixIdx] = useState({});
  const [storicoEan, setStoricoEan] = useState({});
  const [scelte, setScelte] = useState({}); // chiave gruppo → editore_nome confermato dall'utente
  const [editoreDefault, setEditoreDefault] = useState("");
  const [giroNum, setGiroNum] = useState("");
  const [giroAnno, setGiroAnno] = useState(String(new Date().getFullYear()));
  // destinazione: giro numerato (GIRO N AAAA A/B/C…) oppure cedola extra (giro_label "EXTRA", nome libero)
  const [modo, setModo] = useState("giro"); // giro | extra
  const [nomeExtra, setNomeExtra] = useState("");
  const [extraEsistenti, setExtraEsistenti] = useState([]);
  useEffect(() => {
    if (modo !== "extra" || extraEsistenti.length) return;
    fetchJson(`${SUPABASE_URL}/rest/v1/titoli?giro_label=eq.EXTRA&select=n_cedola&order=updated_at.desc&limit=20000`, token)
      .then(r => setExtraEsistenti([...new Set((r || []).map(x => x.n_cedola).filter(Boolean))]))
      .catch(() => {});
  }, [modo, token, extraEsistenti.length]);
  // nome definitivo della cedola extra: maiuscolo, anno aggiunto se manca (come nel Carico Semplice)
  const nCedolaExtra = useMemo(() => {
    const n = String(nomeExtra || "").replace(/\s+/g, " ").trim().toUpperCase();
    if (!n) return "";
    return /\b20\d{2}\b/.test(n) || !/^\d{4}$/.test(giroAnno) ? n : `${n} ${giroAnno}`;
  }, [nomeExtra, giroAnno]);
  const [showMapping, setShowMapping] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  // un file lasciato fuori dal riquadro non deve far aprire il file nel browser
  useEffect(() => {
    const blocca = (e) => { if (e.dataTransfer?.types?.includes("Files")) e.preventDefault(); };
    window.addEventListener("dragover", blocca);
    window.addEventListener("drop", blocca);
    return () => { window.removeEventListener("dragover", blocca); window.removeEventListener("drop", blocca); };
  }, []);
  const [importing, setImporting] = useState(false);
  const [done, setDone] = useState(null);

  const anagByNome = useMemo(() => {
    const m = {};
    anagrafica.forEach(a => { const k = a.editore_nome; if (!m[k] || a.ranking < m[k].ranking) m[k] = a; });
    return m;
  }, [anagrafica]);
  const nomiAnagrafica = useMemo(() => Object.values(anagByNome).sort((a, b) => (a.ranking ?? 999) - (b.ranking ?? 999)), [anagByNome]);

  const { righe: righeFile, scartate } = useMemo(() => (aoa.length ? estraiRighe(aoa, headerRow, map) : { righe: [], scartate: [] }), [aoa, headerRow, map]);

  // Gemelli senza titolo nel file → titolo recuperato da Supabase (titoli / titoli_novita / gemelli già salvati)
  const [titoliNoti, setTitoliNoti] = useState({});
  const eanGemelliSenzaTitolo = useMemo(() => [...new Set(righeFile.flatMap(r => r.gemelli.filter(g => g.ean && !g.titolo).map(g => g.ean)))].sort().join(","), [righeFile]);
  useEffect(() => {
    if (!eanGemelliSenzaTitolo) { setTitoliNoti({}); return; }
    let vivo = true;
    fetchTitoliNoti(token, eanGemelliSenzaTitolo.split(",")).then(m => { if (vivo) setTitoliNoti(m); }).catch(() => {});
    return () => { vivo = false; };
  }, [eanGemelliSenzaTitolo, token]);
  const righe = useMemo(() => righeFile.map(r => ({
    ...r,
    gemelli: r.gemelli.map(g => (g.ean && !g.titolo && titoliNoti[g.ean] ? { ...g, titolo: titoliNoti[g.ean], titoloDaDb: true } : g)),
  })), [righeFile, titoliNoti]);
  const gemStat = useMemo(() => {
    const all = righe.flatMap(r => r.gemelli.filter(g => g.ean));
    return { tot: all.length, daDb: all.filter(g => g.titoloDaDb).length, senza: all.filter(g => !g.titolo).length };
  }, [righe]);

  // Gruppi editore: per nome nel file (colonna o riga di sezione), altrimenti "senza nome"
  const gruppi = useMemo(() => {
    const g = {};
    righe.forEach(r => {
      const raw = r.rawEditore || r.sezione || "";
      const key = raw ? normEd(raw) : "__NONAME__";
      if (!g[key]) g[key] = { key, raw, righe: [] };
      g[key].righe.push(r);
    });
    // proposta automatica per ciascun gruppo
    Object.values(g).forEach(gr => {
      const voti = {};
      gr.righe.forEach(r => { const e = editoreDaEan(r.ean, storicoEan, prefixIdx); if (e && anagByNome[e.nome]) voti[e.nome] = (voti[e.nome] || 0) + 1; });
      const [eanTop, eanN] = Object.entries(voti).sort((a, b) => b[1] - a[1])[0] || [null, 0];
      const eanForte = eanTop && eanN / gr.righe.length >= 0.7 ? eanTop : null;
      gr.eanVoto = eanForte;
      if (!gr.raw) { gr.cand = []; gr.auto = eanForte; gr.fonte = eanForte ? "EAN" : null; return; }
      gr.cand = candidati(gr.raw, nomiAnagrafica);
      const al = alias[normEd(gr.raw)];
      if (al && anagByNome[al]) { gr.auto = al; gr.fonte = "alias"; return; }
      const [b1, b2] = gr.cand;
      const nomeSicuro = b1 && b1.score >= 0.9 && (!b2 || b1.score - b2.score >= 0.05);
      if (nomeSicuro && (!eanForte || eanForte === b1.nome)) { gr.auto = b1.nome; gr.fonte = b1.score >= 1 ? "esatto" : "simile"; return; }
      if (eanForte && (!b1 || gr.cand.some(c => c.nome === eanForte && c.score >= 0.5) || !nomeSicuro)) {
        // lo storico EAN scioglie l'ambiguità (es. LA NAVE DI TESEO vs LA NAVE DI TESEO +)
        const inCand = gr.cand.some(c => c.nome === eanForte);
        gr.auto = inCand ? eanForte : null; gr.proposta = inCand || !b1 ? eanForte : b1.nome; gr.fonte = inCand ? "EAN" : null; return;
      }
      gr.auto = null; gr.proposta = b1?.nome ?? null; gr.fonte = null;
    });
    return Object.values(g).sort((a, b) => a.righe[0]._riga - b.righe[0]._riga);
  }, [righe, storicoEan, prefixIdx, alias, anagByNome, nomiAnagrafica]);

  const editoreDiGruppo = (gr) => scelte[gr.key] ?? gr.auto ?? (gr.key === "__NONAME__" ? editoreDefault || null : null);
  const daVerificare = gruppi.filter(gr => !editoreDiGruppo(gr));

  // Righe finali (ordine file) con anagrafica, cedola, posizione
  const finali = useMemo(() => {
    const cont = {};
    const visti = new Set();
    const extra = modo === "extra";
    const giroOk = extra ? !!nCedolaExtra : /^\d+$/.test(giroNum) && /^\d{4}$/.test(giroAnno);
    const out = [];
    gruppi.forEach(gr => gr.righe.forEach(r => {
      let ed = scelte[gr.key] ?? gr.auto ?? null;
      if (!ed && gr.key === "__NONAME__") ed = editoreDaEan(r.ean, storicoEan, prefixIdx)?.nome || editoreDefault || null;
      const a = ed ? anagByNome[ed] : null;
      const errs = [];
      if (!a) errs.push("editore da abbinare");
      else if (!a.cedola && !extra) errs.push("editore senza categoria cedola");
      if (!r.titolo) errs.push("titolo mancante");
      if (!giroOk) errs.push(extra ? "nome cedola extra mancante" : "giro non impostato");
      const dupKey = extra ? r.ean : `${a?.cedola}|${r.ean}`;
      const dup = visti.has(dupKey); visti.add(dupKey);
      const n_cedola = !giroOk ? null : extra ? nCedolaExtra : a?.cedola ? `GIRO ${giroNum} ${giroAnno} ${a.cedola}` : null;
      out.push({ ...r, editore_nome: ed, anag: a, errs, dup, n_cedola });
    }));
    // ordine esatto del file; posizione progressiva per giro+editore
    out.sort((x, y) => x._riga - y._riga);
    out.forEach(r => { if (r.dup) return; const k = `${r.n_cedola}|${r.editore_nome}`; cont[k] = (cont[k] || 0) + 1; r.posizione = cont[k]; });
    return out;
  }, [gruppi, scelte, editoreDefault, giroNum, giroAnno, modo, nCedolaExtra, anagByNome, storicoEan, prefixIdx]);

  const importabili = finali.filter(r => !r.errs.length && !r.dup);
  const conErrori = finali.filter(r => r.errs.length);

  // ─── Caricamento file ─────────────────────────────────────────────────────
  const caricaFoglio = (workbook, nome) => {
    const data = window.XLSX.utils.sheet_to_json(workbook.Sheets[nome], { header: 1, defval: "", raw: true });
    const hr = rilevaIntestazioni(data);
    const { headers: h, map: m } = rilevaMapping(data, hr);
    setFoglio(nome); setAoa(data); setHeaderRow(hr); setHeaders(h); setMap(m);
    return { data, hr, m };
  };

  // Anagrafica, alias e prefissi EAN: servono sia al file sia all'inserimento manuale
  const caricaAnagrafica = async () => {
    const [anag, al, pref] = await Promise.all([
      fetchJson(`${SUPABASE_URL}/rest/v1/ranking_editori?select=editore_nome,codice_editore,ranking,account_editore,promozione,cedola,attivo&attivo=is.true`, token),
      fetchJson(`${SUPABASE_URL}/rest/v1/alias_editori?select=alias,editore_nome`, token).catch(() => []),
      fetchJson(`${SUPABASE_URL}/rest/v1/rpc/prefissi_ean_editori`, token, { method: "POST", headers: { "Content-Type": "application/json" }, body: "{}" }).catch(() => []),
    ]);
    setAnagrafica(anag.map(a => ({ ...a, editore_nome: String(a.editore_nome ?? "").replace(/\s+/g, " ").trim().toUpperCase(), cedola: String(a.cedola ?? "").trim().toUpperCase() })));
    const am = {}; (al || []).forEach(x => { am[x.alias] = x.editore_nome; }); setAlias(am);
    setPrefixIdx(buildPrefixIndex(pref || []));
    return anag;
  };

  // ─── Inserimento manuale: le righe diventano un foglio con colonne fisse e seguono lo stesso percorso del file ──
  const [righeManuali, setRigheManuali] = useState(() => Array.from({ length: 5 }, rigaManualeVuota));
  const [nomiEditori, setNomiEditori] = useState([]);
  useEffect(() => {
    if (!manuale) return;
    fetchJson(`${SUPABASE_URL}/rest/v1/ranking_editori?select=editore_nome&attivo=is.true&order=editore_nome.asc`, token)
      .then(r => setNomiEditori([...new Set((r || []).map(x => String(x.editore_nome ?? "").trim().toUpperCase()).filter(Boolean))]))
      .catch(() => {});
  }, [manuale, token]);
  const confermaManuale = async () => {
    const piene = righeManuali.filter(r => CAMPI_MANUALI.some(c => String(r[c.k] ?? "").trim()));
    if (!piene.length) { alert("Inserisci almeno un titolo."); return; }
    const senzaEan = piene.filter(r => !parseEan(r.ean));
    if (senzaEan.length) { alert(`${senzaEan.length} ${senzaEan.length === 1 ? "riga ha" : "righe hanno"} l'EAN mancante o non valido: correggi prima di confermare.`); return; }
    setLoading(true);
    try {
      await caricaAnagrafica();
      const data = [CAMPI_MANUALI.map(c => c.label), ...piene.map(r => CAMPI_MANUALI.map(c => String(r[c.k] ?? "").trim()))];
      const X = window.XLSX;
      const workbook = X.utils.book_new();
      X.utils.book_append_sheet(workbook, X.utils.aoa_to_sheet(data), "Manuale");
      setWb(workbook);
      setFileName(NOME_MANUALE);
      // colonne note in partenza: niente riconoscimento automatico
      const m = {}; CAMPI.forEach(c => { m[c.key] = -1; }); CAMPI_MANUALI.forEach((c, i) => { m[c.map] = i; });
      setFoglio("Manuale"); setAoa(data); setHeaderRow(0); setHeaders(data[0]); setMap(m);
      setStoricoEan(await fetchStoricoEan(token, piene.map(r => parseEan(r.ean))));
      setEditoreDefault(""); setScelte({});
      setStep("preview");
    } catch (err) {
      alert("Errore: " + err.message);
    }
    setLoading(false);
  };

  // accetta sia l'evento dell'input file sia un File trascinato
  const handleFile = useCallback(async (e) => {
    const f = e instanceof File ? e : e?.target?.files?.[0];
    if (!f) return;
    if (!(e instanceof File)) e.target.value = "";
    if (!/\.(xlsx|xls|xlsm|csv)$/i.test(f.name)) { alert("Formato non supportato: usa .xlsx, .xls, .xlsm o .csv"); return; }
    setLoading(true);
    try {
      const anag = await caricaAnagrafica();

      const buf = await f.arrayBuffer();
      const workbook = window.XLSX.read(buf, { type: "array", cellDates: false });
      setWb(workbook);
      setFileName(f.name);
      const { data, hr, m } = caricaFoglio(workbook, sceltaFoglio(workbook));
      const { righe: rr } = estraiRighe(data, hr, m);
      setStoricoEan(await fetchStoricoEan(token, rr.map(r => r.ean)));

      // giro e editore di default suggeriti dal nome file (es. "Giro 5 2026 - Neri Pozza.xlsx")
      const gm = f.name.match(/giro[\s_-]*(\d{1,2})(?:[\s_-]+(\d{4}))?/i);
      if (gm) { setGiroNum(gm[1]); if (gm[2]) setGiroAnno(gm[2]); }
      // file di campagna/cedola extra (senza "giro N" nel nome): propone la modalità extra col nome file
      else if (/extra|campagna|promo/i.test(f.name)) {
        setModo("extra");
        setNomeExtra(f.name.replace(/\.[^.]+$/, "").replace(/[_]+/g, " ").replace(/\s+/g, " ").trim().toUpperCase());
      }
      const base = f.name.replace(/\.[^.]+$/, "").replace(/giro[\s_-]*\d+([\s_-]+\d{4})?/i, "");
      const guess = anag.map(a => ({ n: String(a.editore_nome).trim().toUpperCase(), s: Math.max(...base.split(/[_\-–.]+/).map(p => scoreEditore(p, a.editore_nome))) })).sort((x, y) => y.s - x.s)[0];
      setEditoreDefault(guess && guess.s >= 0.9 ? guess.n : "");
      setScelte({});
      setStep("preview");
    } catch (err) {
      alert("Errore lettura file: " + err.message);
    }
    setLoading(false);
  }, [token]);

  const cambiaFoglio = async (nome) => {
    const { data, hr, m } = caricaFoglio(wb, nome);
    const { righe: rr } = estraiRighe(data, hr, m);
    setStoricoEan(await fetchStoricoEan(token, rr.map(r => r.ean)));
    setScelte({});
  };
  const cambiaHeaderRow = (n) => {
    const hr = parseInt(n) - 1;
    const { headers: h, map: m } = rilevaMapping(aoa, hr);
    setHeaderRow(hr); setHeaders(h); setMap(m);
  };

  const confermaSuggerimenti = () => {
    const s = { ...scelte };
    gruppi.forEach(gr => { if (!editoreDiGruppo(gr) && (gr.proposta || gr.cand?.[0])) s[gr.key] = gr.proposta || gr.cand[0].nome; });
    setScelte(s);
  };

  // ─── Import ───────────────────────────────────────────────────────────────
  const handleImport = async () => {
    setImporting(true);
    try {
      const extra = modo === "extra";
      let giriMap = {}, esistenti = {};
      if (extra) {
        // cedola extra: giro_id NULL, i titoli già presenti si riconoscono per nome cedola + EAN
        esistenti = await fetchEsistentiExtra(token, nCedolaExtra, importabili.map(r => r.ean));
      } else {
        const combos = new Set(importabili.map(r => `${giroNum}|${giroAnno}|${r.anag.cedola}`));
        giriMap = await resolveGiri(token, combos);
        esistenti = await fetchEsistenti(token, Object.values(giriMap), importabili.map(r => r.ean));
      }
      let aggiornati = 0;
      const payload = importabili.map(r => {
        const giro_id = extra ? null : giriMap[`${giroNum}|${giroAnno}|${r.anag.cedola}`];
        const ex = (extra ? esistenti[r.ean] : esistenti[`${giro_id}|${r.ean}`]) || {};
        if (ex.id) aggiornati++;
        // dato nel file → scrive; dato assente → resta quello già salvato (gemelli slot per slot,
        // titolo del gemello conservato solo se l'EAN non cambia)
        const gem = (k) => {
          const x = r.gemelli[k - 1];
          if (!x) return { ean: ex[`ean_gemello_${k}`] ?? null, titolo: ex[`titolo_gemello_${k}`] ?? null };
          const ean = x.ean ?? ex[`ean_gemello_${k}`] ?? null;
          return { ean, titolo: x.titolo ?? (ean === ex[`ean_gemello_${k}`] ? ex[`titolo_gemello_${k}`] ?? null : null) };
        };
        const [g1, g2, g3] = [gem(1), gem(2), gem(3)];
        const noteTir = r.tiratura && r.obiettivoFonte !== "T" ? `TIRATURA ${r.tiratura}` : null;
        const note = [r.note, noteTir].filter(Boolean).join(" · ") || null;
        return {
          _id: ex.id ?? null,
          giro_id, giro_label: extra ? "EXTRA" : `${giroNum} ${giroAnno}`, n_cedola: r.n_cedola,
          ean: r.ean, titolo: r.titolo ?? ex.titolo ?? null, autore: r.autore ?? ex.autore ?? null,
          editore_nome: r.editore_nome, codice_editore: r.anag.codice_editore ?? ex.codice_editore ?? null,
          ranking_editore: r.anag.ranking ?? ex.ranking_editore ?? null, account_editore: r.anag.account_editore ?? ex.account_editore ?? null,
          promozione: ex.promozione ?? r.anag.promozione ?? null,
          prezzo: r.prezzo ?? ex.prezzo ?? null, uscita: r.uscita ?? ex.uscita ?? null,
          formato: ex.formato ?? "Cover", eta: ex.eta ?? null, il_triangolo: ex.il_triangolo ?? null,
          note_comunicazione: ex.note_comunicazione ?? null,
          posizione: r.posizione, ranking_titolo: ex.ranking_titolo ?? null,
          obiettivo_assegnato: r.obiettivo ?? r.tiratura ?? ex.obiettivo_assegnato ?? 0,
          top_100: r.top_100 ?? ex.top_100 ?? false,
          note: note ?? ex.note ?? null,
          ean_gemello_1: g1.ean, titolo_gemello_1: g1.titolo,
          ean_gemello_2: g2.ean, titolo_gemello_2: g2.titolo,
          ean_gemello_3: g3.ean, titolo_gemello_3: g3.titolo,
        };
      });
      if (!extra) {
        await fetchJson(`${SUPABASE_URL}/rest/v1/rpc/upsert_titoli`, token, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ payload: payload.map(({ _id, ...p }) => p) }) });
      } else {
        // UNIQUE(ean, giro_id) non scatta con giro_id NULL: upsert duplicherebbe le righe a ogni reimport.
        // Quindi: aggiornamento per id dei titoli già nella cedola, inserimento dei nuovi.
        const upd = payload.filter(p => p._id), ins = payload.filter(p => !p._id).map(({ _id, ...p }) => p);
        for (let i = 0; i < upd.length; i += 10) {
          await Promise.all(upd.slice(i, i + 10).map(({ _id, ...p }) => fetchJson(`${SUPABASE_URL}/rest/v1/titoli?id=eq.${_id}`, token, { method: "PATCH", headers: { "Content-Type": "application/json", Prefer: "return=minimal" }, body: JSON.stringify(p) })));
        }
        for (let i = 0; i < ins.length; i += 200) {
          await fetchJson(`${SUPABASE_URL}/rest/v1/titoli`, token, { method: "POST", headers: { "Content-Type": "application/json", Prefer: "return=minimal" }, body: JSON.stringify(ins.slice(i, i + 200)) });
        }
      }

      // memorizza gli abbinamenti nome-file → anagrafica per i prossimi import
      const nuoviAlias = gruppi
        .filter(gr => gr.raw && editoreDiGruppo(gr) && (scelte[gr.key] || gr.fonte === "EAN"))
        .map(gr => ({ alias: normEd(gr.raw), editore_nome: editoreDiGruppo(gr) }));
      if (nuoviAlias.length) {
        await fetchJson(`${SUPABASE_URL}/rest/v1/alias_editori?on_conflict=alias`, token, { method: "POST", headers: { "Content-Type": "application/json", Prefer: "resolution=merge-duplicates,return=minimal" }, body: JSON.stringify(nuoviAlias) }).catch(() => {});
      }
      setDone({ tot: payload.length, nuovi: payload.length - aggiornati, aggiornati, alias: nuoviAlias.length, cedole: [...new Set(payload.map(p => p.n_cedola))] });
      setStep("result");
      onImportDone && onImportDone();
    } catch (err) {
      alert("Errore import: " + err.message);
    }
    setImporting(false);
  };

  const reset = () => {
    // dopo un import riuscito la griglia manuale riparte vuota; tornando indietro dall'anteprima resta com'era
    if (step === "result" && manuale) setRigheManuali(Array.from({ length: 5 }, rigaManualeVuota));
    setStep("upload"); setWb(null); setAoa([]); setScelte({}); setDone(null); setFileName("");
  };

  const setCella = (i, k, v) => setRigheManuali(rr => rr.map((r, j) => (j === i ? { ...r, [k]: v } : r)));
  // incolla da Excel: più celle (tab / a capo) riempiono la griglia a partire dalla cella selezionata
  const incolla = (e, i, col) => {
    const testo = e.clipboardData.getData("text/plain");
    if (!/[\t\n]/.test(testo.replace(/\r?\n$/, ""))) return;
    e.preventDefault();
    const blocco = testo.replace(/\r/g, "").replace(/\n$/, "").split("\n").map(l => l.split("\t"));
    setRigheManuali(rr => {
      const out = rr.map(r => ({ ...r }));
      blocco.forEach((celle, di) => {
        while (out.length <= i + di) out.push(rigaManualeVuota());
        celle.forEach((v, dc) => { const c = CAMPI_MANUALI[col + dc]; if (c) out[i + di][c.k] = v.trim(); });
      });
      return out;
    });
  };

  // ─── UI ───────────────────────────────────────────────────────────────────
  if (step === "upload" && manuale) {
    const piena = (r) => CAMPI_MANUALI.some(c => String(r[c.k] ?? "").trim());
    const statoRiga = (r) => {
      if (!piena(r)) return null;
      if (!parseEan(r.ean)) return { ok: false, msg: String(r.ean).trim() ? "EAN non valido" : "EAN mancante" };
      if (!String(r.titolo).trim()) return { ok: false, msg: "Titolo mancante" };
      return { ok: true, msg: "Pronta" };
    };
    const stati = righeManuali.map(statoRiga);
    const compilate = stati.filter(Boolean).length;
    const pronte = stati.filter(s => s?.ok).length;
    const daSistemare = compilate - pronte;
    const LARG_NUM = 44;
    const bordo = `1px solid ${T.border}`;
    const thBase = { position: "sticky", background: T.surface, color: T.textMid, fontWeight: 500, fontSize: "11px", textAlign: "left", whiteSpace: "nowrap", borderBottom: bordo, borderRight: bordo, padding: "7px 10px", zIndex: 2 };
    const fisso = (left) => ({ position: "sticky", left, zIndex: 3 });
    const cellaInput = { width: "100%", boxSizing: "border-box", background: "transparent", border: "none", outline: "none", color: T.text, font: "inherit", fontSize: "12px", padding: "9px 10px", borderRadius: 0 };
    const pulsante = (extra = {}) => ({ ...css.btn(), padding: "7px 14px", borderRadius: 6, ...extra });
    return (
      <div style={{ background: T.surface, border: bordo, borderRadius: 10, overflow: "hidden" }}>
        <style>{`
          .gm-cella:focus-within { box-shadow: inset 0 0 0 2px ${T.accent}; background: ${T.accent}12 !important; }
          .gm-riga:hover > td { background: ${T.border}55; }
          .gm-cella input::placeholder { color: ${T.textDim}; }
          .gm-elimina { opacity: 0; transition: opacity .12s; }
          .gm-riga:hover .gm-elimina { opacity: 1; }
        `}</style>

        {/* Intestazione */}
        <div style={{ display: "flex", alignItems: "flex-start", gap: 16, padding: "18px 20px 14px", borderBottom: bordo }}>
          <div style={{ fontSize: "22px", lineHeight: 1 }}>✍️</div>
          <div style={{ flex: 1 }}>
            <div style={{ color: T.text, fontSize: "14px", fontWeight: 700, marginBottom: 4 }}>Inserimento manuale</div>
            <div style={{ color: T.textMid, fontSize: "12px", lineHeight: 1.6 }}>
              Scrivi un titolo per riga, oppure copia un blocco di celle da Excel e incollalo nella prima casella. Obbligatori EAN e titolo.
              Con <b style={{ color: T.text }}>Conferma</b> scegli giro o cedola extra e il nome della cedola, come per il file editore.
            </div>
          </div>
        </div>

        {/* Griglia */}
        <datalist id="editori-anagrafica">{nomiEditori.map(n => <option key={n} value={n} />)}</datalist>
        <div style={{ overflow: "auto", maxHeight: "58vh", scrollPaddingLeft: LARG_NUM + 8, scrollPaddingTop: 64 }}>
          <table style={{ borderCollapse: "separate", borderSpacing: 0, minWidth: "100%" }}>
            <thead>
              <tr>
                <th rowSpan={2} style={{ ...thBase, ...fisso(0), top: 0, width: LARG_NUM, minWidth: LARG_NUM, textAlign: "center", zIndex: 4 }}>#</th>
                <th colSpan={7} style={{ ...thBase, top: 0, color: T.text, fontWeight: 700, letterSpacing: "0.04em" }}>Titolo</th>
                {[1, 2, 3].map(n => (
                  <th key={n} colSpan={2} style={{ ...thBase, top: 0, color: T.blue, fontWeight: 700, letterSpacing: "0.04em", background: cv("#14203a", "#eef3ff") }}>Gemello {n}</th>
                ))}
                <th rowSpan={2} style={{ ...thBase, top: 0, width: 40, borderRight: "none" }}></th>
              </tr>
              <tr>
                {CAMPI_MANUALI.map((c, ci) => (
                  <th key={c.k} style={{ ...thBase, top: 31, minWidth: c.w, ...(c.gem ? { background: cv("#14203a", "#eef3ff") } : {}) }}>
                    {c.gem ? (c.k.endsWith("e") ? "EAN" : "Titolo") : c.label}
                    {["ean", "titolo"].includes(c.k) && <span style={{ color: T.accent }}> *</span>}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {righeManuali.map((r, i) => {
                const st = stati[i];
                const sfondoRiga = i % 2 ? cv("#131313", "#fafbfe") : T.bg;
                return (
                  <tr key={i} className="gm-riga">
                    <td style={{ ...fisso(0), background: sfondoRiga, borderBottom: bordo, borderRight: bordo, textAlign: "center", fontSize: "11px", color: T.textDim, padding: "0 6px" }} title={st?.msg || "Riga vuota"}>
                      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 5 }}>
                        <span style={{ width: 7, height: 7, borderRadius: "50%", background: !st ? "transparent" : st.ok ? T.green : T.red, border: st ? "none" : `1px solid ${T.borderHi}` }} />
                        {i + 1}
                      </div>
                    </td>
                    {CAMPI_MANUALI.map((c, ci) => {
                      const errore = c.k === "ean" && st && !st.ok && st.msg.startsWith("EAN") || c.k === "titolo" && st && st.msg === "Titolo mancante";
                      return (
                        <td key={c.k} className="gm-cella" style={{ background: c.gem ? cv("#121a2e", "#f5f8ff") : sfondoRiga, borderBottom: bordo, borderRight: bordo, padding: 0, ...(errore ? { boxShadow: `inset 0 -2px 0 ${T.red}` } : {}) }}>
                          <input value={r[c.k]} onChange={e => setCella(i, c.k, e.target.value)} onPaste={e => incolla(e, i, ci)}
                            list={c.k === "editore" ? "editori-anagrafica" : undefined}
                            inputMode={["prezzo", "obiettivo"].includes(c.k) ? "decimal" : undefined}
                            placeholder={i === 0 ? ({ ean: "978…", prezzo: "0,00", obiettivo: "copie" }[c.k] || "") : ""}
                            style={{ ...cellaInput, textAlign: ["prezzo", "obiettivo"].includes(c.k) ? "right" : "left", fontVariantNumeric: ["ean", "g1e", "g2e", "g3e", "prezzo", "obiettivo"].includes(c.k) ? "tabular-nums" : undefined }}
                            title={errore ? st.msg : ""} />
                        </td>
                      );
                    })}
                    <td style={{ background: sfondoRiga, borderBottom: bordo, textAlign: "center", padding: 0 }}>
                      <button className="gm-elimina" title="Elimina riga"
                        onClick={() => setRigheManuali(rr => (rr.length > 1 ? rr.filter((_, j) => j !== i) : [rigaManualeVuota()]))}
                        style={{ background: "transparent", border: "none", color: T.textMid, cursor: "pointer", fontSize: "14px", padding: "6px 10px" }}>✕</button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <button onClick={() => setRigheManuali(rr => [...rr, rigaManualeVuota()])}
          style={{ display: "block", width: "100%", textAlign: "left", background: "transparent", border: "none", borderTop: bordo, color: T.accent, cursor: "pointer", font: "inherit", fontSize: "12px", padding: "10px 20px" }}>
          + Aggiungi riga
        </button>

        {/* Barra in fondo */}
        <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "14px 20px", borderTop: bordo, background: T.bg, flexWrap: "wrap" }}>
          <button style={pulsante()} onClick={() => setRigheManuali(rr => [...rr, ...Array.from({ length: 10 }, rigaManualeVuota)])}>+ 10 righe</button>
          <button style={pulsante({ color: T.textMid })} onClick={() => { if (confirm("Svuotare tutta la griglia?")) setRigheManuali(Array.from({ length: 5 }, rigaManualeVuota)); }}>Svuota</button>
          <div style={{ display: "flex", gap: 14, marginLeft: 12, fontSize: "12px" }}>
            <span style={{ color: T.textMid }}><b style={{ color: T.text }}>{compilate}</b> {compilate === 1 ? "titolo" : "titoli"}</span>
            {pronte > 0 && <span style={{ color: T.green }}>● {pronte} {pronte === 1 ? "pronto" : "pronti"}</span>}
            {daSistemare > 0 && <span style={{ color: T.red }}>● {daSistemare} da sistemare</span>}
          </div>
          <button onClick={confermaManuale} disabled={loading || !compilate || daSistemare > 0}
            style={{ ...css.btn("accent"), marginLeft: "auto", padding: "9px 22px", borderRadius: 6, fontSize: "13px", opacity: loading || !compilate || daSistemare > 0 ? 0.45 : 1, cursor: loading || !compilate || daSistemare > 0 ? "default" : "pointer" }}
            title={daSistemare > 0 ? "Sistema le righe segnate in rosso" : ""}>
            {loading ? "Analisi in corso…" : `Conferma ${pronte || ""} →`}
          </button>
        </div>
      </div>
    );
  }

  if (step === "upload") return (
    <div style={{ maxWidth: 560 }}>
      <div
        onDragOver={e => { e.preventDefault(); e.dataTransfer.dropEffect = "copy"; if (!dragOver) setDragOver(true); }}
        onDragLeave={e => { if (!e.currentTarget.contains(e.relatedTarget)) setDragOver(false); }}
        onDrop={e => { e.preventDefault(); setDragOver(false); if (!loading && e.dataTransfer.files?.[0]) handleFile(e.dataTransfer.files[0]); }}
        style={{ border: `2px dashed ${dragOver ? T.accent : T.borderHi}`, background: dragOver ? T.accent + "14" : "transparent", borderRadius: 6, padding: 40, textAlign: "center", marginBottom: 16, transition: "all .15s" }}>
        <div style={{ fontSize: "32px", marginBottom: 12 }}>📑</div>
        <div style={{ color: T.text, marginBottom: 8 }}>{loading ? "Lettura e analisi in corso..." : dragOver ? "Rilascia qui il file" : "Trascina qui il file ricevuto dall'editore, oppure"}</div>
        <div style={{ color: T.textMid, fontSize: "11px", marginBottom: 20 }}>.xlsx, .xls o .csv, qualsiasi layout: colonne ed editori vengono riconosciuti in automatico</div>
        <input type="file" accept=".xlsx,.xls,.xlsm,.csv" onChange={handleFile} style={{ display: "none" }} id="file-editore" disabled={loading} />
        <label htmlFor="file-editore" style={{ ...css.btn("accent"), cursor: loading ? "default" : "pointer", padding: "8px 20px", opacity: loading ? 0.6 : 1 }}>Scegli file</label>
      </div>
      <div style={{ color: T.textMid, fontSize: "11px", lineHeight: 1.6 }}>
        Legge: EAN, titolo, autore, editore, prezzo, uscita, top 100, obiettivo/tiratura, note, gemelli.<br />
        Ordine titoli = ordine del file · ranking editore = anagrafica Supabase.
      </div>
    </div>
  );

  if (step === "result" && done) return (
    <div style={{ textAlign: "center", padding: 60 }}>
      <div style={{ fontSize: "48px", marginBottom: 16 }}>✅</div>
      <div style={{ color: T.green, fontSize: "20px", fontWeight: "700", marginBottom: 8 }}>Import completato</div>
      <div style={{ color: T.textMid, marginBottom: 8 }}>{done.tot} titoli · {done.nuovi} nuovi · {done.aggiornati} aggiornati</div>
      <div style={{ color: T.textMid, marginBottom: 8, fontSize: "12px" }}>{done.cedole.join(" · ")}</div>
      {done.alias > 0 && <div style={{ color: T.blue, fontSize: "11px", marginBottom: 24 }}>{done.alias} abbinamenti editore memorizzati per i prossimi import</div>}
      <button style={css.btn("accent")} onClick={reset}>Nuovo import</button>
    </div>
  );

  const fonteBadge = (f) => {
    const cfg = { esatto: [T.green, "✓ esatto"], simile: [T.green, "≈ simile"], alias: [T.blue, "↺ memorizzato"], EAN: [T.blue, "# da EAN"], manuale: [T.accent, "✎ scelto"] }[f];
    return cfg ? <span style={{ color: cfg[0], fontSize: "11px", whiteSpace: "nowrap" }}>{cfg[1]}</span> : <span style={{ color: T.orange, fontSize: "11px", fontWeight: 700 }}>⚠ da verificare</span>;
  };

  return (
    <div>
      {/* Barra superiore */}
      <div style={{ display: "flex", gap: 12, marginBottom: 16, alignItems: "center", flexWrap: "wrap" }}>
        <div style={css.box()}><span style={{ color: T.textMid, fontSize: "11px" }}>File: </span><span style={{ color: T.text, fontWeight: 700 }}>{fileName}</span></div>
        <div style={css.box()}><span style={{ color: T.textMid, fontSize: "11px" }}>Titoli letti: </span><span style={{ fontWeight: 700 }}>{finali.length}</span></div>
        <div style={css.box(daVerificare.length ? T.orange : T.green)}><span style={{ color: T.textMid, fontSize: "11px" }}>Editori da verificare: </span><span style={{ fontWeight: 700, color: daVerificare.length ? T.orange : T.green }}>{daVerificare.length}</span></div>
        {gemStat.tot > 0 && (
          <div style={css.box(gemStat.senza ? T.orange : T.green)} title={gemStat.daDb ? `${gemStat.daDb} titoli gemelli non presenti nel file, recuperati da Supabase` : ""}>
            <span style={{ color: T.textMid, fontSize: "11px" }}>Gemelli: </span><span style={{ fontWeight: 700 }}>{gemStat.tot}</span>
            {gemStat.daDb > 0 && <span style={{ color: T.blue, fontSize: "11px" }}> · {gemStat.daDb} titoli da Supabase</span>}
            {gemStat.senza > 0 && <span style={{ color: T.orange, fontSize: "11px", fontWeight: 700 }}> · {gemStat.senza} senza titolo</span>}
          </div>
        )}
        <div style={css.box(conErrori.length ? T.red : T.green)}><span style={{ color: T.textMid, fontSize: "11px" }}>Righe con errori: </span><span style={{ fontWeight: 700, color: conErrori.length ? T.red : T.green }}>{conErrori.length}</span></div>
        <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
          <button style={css.btn()} onClick={reset}>{fileName === NOME_MANUALE ? "← Torna a modificare" : "← Ricarica"}</button>
          <button style={{ ...css.btn("accent"), opacity: importing || !importabili.length || conErrori.length ? 0.5 : 1 }} onClick={handleImport} disabled={importing || !importabili.length || conErrori.length > 0}>
            {importing ? "Import in corso..." : `Importa ${importabili.length} titoli`}
          </button>
        </div>
      </div>

      {/* Giro + foglio */}
      <div style={{ ...css.panel, display: "flex", gap: 20, alignItems: "center", flexWrap: "wrap" }}>
        <div style={{ display: "flex", border: `1px solid ${T.borderHi}`, borderRadius: 3, overflow: "hidden" }}>
          {[["giro", "Giro numerato"], ["extra", "Cedola extra"]].map(([k, l]) => (
            <button key={k} onClick={() => setModo(k)} style={{ ...css.btn(modo === k ? "accent" : "default"), border: "none", borderRadius: 0 }}>{l}</button>
          ))}
        </div>
        {modo === "giro" ? (
          <label style={{ color: T.textMid, fontSize: "12px" }}>Giro&nbsp;
            <input style={{ ...css.input, width: 50 }} value={giroNum} onChange={e => setGiroNum(e.target.value.replace(/\D/g, ""))} placeholder="N" />
          </label>
        ) : (
          <label style={{ color: T.textMid, fontSize: "12px" }} title="Scrivi un nome nuovo per creare la cedola, oppure scegline una esistente per aggiungere/aggiornare titoli">Nome cedola&nbsp;
            <input style={{ ...css.input, width: 280, borderColor: nCedolaExtra ? T.borderHi : T.orange }} value={nomeExtra} onChange={e => setNomeExtra(e.target.value)} placeholder="es. CAMPAGNA NATALE GALLUCCI" list="extra-esistenti" />
            <datalist id="extra-esistenti">{extraEsistenti.map(n => <option key={n} value={n} />)}</datalist>
          </label>
        )}
        <label style={{ color: T.textMid, fontSize: "12px" }} title={modo === "extra" ? "Aggiunto al nome della cedola se non c'è già un anno" : ""}>Anno&nbsp;
          <input style={{ ...css.input, width: 64 }} value={giroAnno} onChange={e => setGiroAnno(e.target.value.replace(/\D/g, "").slice(0, 4))} />
        </label>
        {wb && wb.SheetNames.length > 1 && (
          <label style={{ color: T.textMid, fontSize: "12px" }}>Foglio&nbsp;
            <select style={css.input} value={foglio} onChange={e => cambiaFoglio(e.target.value)}>{wb.SheetNames.map(n => <option key={n}>{n}</option>)}</select>
          </label>
        )}
        <label style={{ color: T.textMid, fontSize: "12px" }}>Riga intestazioni&nbsp;
          <input style={{ ...css.input, width: 50 }} type="number" min="1" value={headerRow + 1} onChange={e => cambiaHeaderRow(e.target.value)} />
        </label>
        <button style={css.btn()} onClick={() => setShowMapping(s => !s)}>{showMapping ? "▾" : "▸"} Colonne riconosciute</button>
        <span style={{ color: T.textMid, fontSize: "11px" }}>{modo === "extra"
          ? (nCedolaExtra ? <>Tutti i titoli vanno nella cedola extra <b style={{ color: T.accent }}>{nCedolaExtra}</b>{extraEsistenti.includes(nCedolaExtra) ? " (esistente: titoli aggiunti/aggiornati)" : " (nuova)"}</> : "Scrivi il nome della cedola extra")
          : "La cedola (A/B/C/KIDS/SERVICE) è presa dall'anagrafica di ciascun editore"}</span>
      </div>

      {showMapping && (
        <div style={{ ...css.panel, display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(230px, 1fr))", gap: 10 }}>
          {CAMPI.map(c => (
            <label key={c.key} style={{ display: "flex", flexDirection: "column", gap: 4, fontSize: "11px", color: map[c.key] >= 0 ? T.text : c.req ? T.red : T.textMid }}>
              {c.label}{c.req ? " *" : ""}
              <select style={css.input} value={map[c.key] ?? -1} onChange={e => setMap(m => ({ ...m, [c.key]: parseInt(e.target.value) }))}>
                <option value={-1}>— non presente —</option>
                {headers.map((h, i) => <option key={i} value={i}>{colLetter(i)} · {h || "(senza intestazione)"}</option>)}
              </select>
            </label>
          ))}
        </div>
      )}

      {/* Abbinamento editori */}
      <div style={css.panel}>
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 10 }}>
          <div style={{ color: T.accent, fontWeight: 700, fontSize: "12px", textTransform: "uppercase", letterSpacing: "0.08em" }}>Editori nel file → anagrafica</div>
          {daVerificare.some(gr => gr.proposta || gr.cand?.[0]) && <button style={css.btn()} onClick={confermaSuggerimenti}>Accetta i suggerimenti</button>}
        </div>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr>{["Nome nel file", "Titoli", "Editore in anagrafica", "Cedola", "Rk", "Stato"].map(h => <th key={h} style={{ ...css.th, position: "static" }}>{h}</th>)}</tr></thead>
          <tbody>
            {gruppi.map(gr => {
              const scelto = editoreDiGruppo(gr);
              const a = scelto ? anagByNome[scelto] : null;
              const fonte = scelte[gr.key] ? "manuale" : gr.auto ? gr.fonte : null;
              const isNoName = gr.key === "__NONAME__";
              const valore = scelto ?? "";
              return (
                <tr key={gr.key} style={{ background: scelto ? "transparent" : T.orange + "11" }}>
                  <td style={{ ...css.td, fontWeight: 600 }}>{isNoName ? <span style={{ color: T.textMid, fontStyle: "italic" }}>nessun editore indicato{gr.auto ? " (riconosciuto da EAN)" : ""}</span> : gr.raw}</td>
                  <td style={{ ...css.td, color: T.textMid }}>{gr.righe.length}</td>
                  <td style={css.td}>
                    <select style={{ ...css.input, minWidth: 260, borderColor: scelto ? T.borderHi : T.orange }} value={valore}
                      onChange={e => setScelte(s => ({ ...s, [gr.key]: e.target.value || undefined }))}>
                      <option value="">{gr.proposta ? `— scegli (suggerito: ${gr.proposta}) —` : "— scegli editore —"}</option>
                      {gr.cand?.length > 0 && <optgroup label="Più simili">{gr.cand.map(c => <option key={"c" + c.nome} value={c.nome}>{c.nome} · {Math.min(100, Math.round(c.score * 100))}%</option>)}</optgroup>}
                      {gr.eanVoto && !gr.cand?.some(c => c.nome === gr.eanVoto) && <optgroup label="Da storico EAN"><option value={gr.eanVoto}>{gr.eanVoto}</option></optgroup>}
                      <optgroup label="Tutti">{nomiAnagrafica.map(x => <option key={x.editore_nome} value={x.editore_nome}>{x.editore_nome}</option>)}</optgroup>
                    </select>
                  </td>
                  <td style={{ ...css.td, color: T.accent }}>{a?.cedola || "—"}</td>
                  <td style={{ ...css.td, color: T.accent, fontWeight: 700 }}>{a?.ranking ?? "—"}</td>
                  <td style={css.td}>{isNoName && !gr.auto && !scelte[gr.key] ? fonteBadge(valore ? "manuale" : null) : fonteBadge(fonte)}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {scartate.length > 0 && (
        <div style={{ ...css.panel, borderColor: T.orange + "66" }}>
          <div style={{ color: T.orange, fontWeight: 700, fontSize: "12px", marginBottom: 6 }}>ℹ {scartate.length} righe ignorate (senza EAN valido)</div>
          {scartate.slice(0, 15).map((s, i) => <div key={i} style={{ color: T.textMid, fontSize: "11px" }}>Riga {s.riga}: {s.testo}</div>)}
          {scartate.length > 15 && <div style={{ color: T.textDim, fontSize: "11px" }}>…e altre {scartate.length - 15}</div>}
        </div>
      )}

      {/* Anteprima titoli nell'ordine del file */}
      <div style={{ overflow: "auto", maxHeight: "60vh", border: `1px solid ${T.border}`, borderRadius: 4 }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr>{["Riga", "N. Cedola", "Pos.", "Rk Ed.", "EAN", "Titolo", "Autore", "Editore", "Prezzo", "Uscita", "★", "Obj", "Note", "Gemelli", ""].map(h => <th key={h} style={css.th}>{h}</th>)}</tr></thead>
          <tbody>
            {finali.map((r, i) => (
              <tr key={i} style={{ background: r.errs.length ? T.red + "11" : r.dup ? T.textDim + "33" : i % 2 ? T.surface + "66" : "transparent" }}>
                <td style={{ ...css.td, color: T.textDim, fontSize: "11px" }}>{r._riga}</td>
                <td style={{ ...css.td, color: r.n_cedola ? T.accent : T.red, fontWeight: 600, fontSize: "11px", whiteSpace: "nowrap" }}>{r.n_cedola ?? "—"}</td>
                <td style={{ ...css.td, color: T.textMid, textAlign: "center" }}>{r.dup ? "—" : r.posizione}</td>
                <td style={{ ...css.td, color: T.accent, fontWeight: 700, fontSize: "11px" }}>{r.anag?.ranking ?? "—"}</td>
                <td style={{ ...css.td, fontFamily: "monospace", fontSize: "11px", color: r.eanWarn ? T.orange : T.textMid }} title={r.eanWarn ? "Cifra di controllo EAN non valida: verifica" : ""}>{r.ean}{r.eanWarn ? " ⚠" : ""}</td>
                <td style={{ ...css.td, maxWidth: 240, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontWeight: 600 }} title={r.titolo}>{r.titolo ?? <span style={{ color: T.red }}>—</span>}</td>
                <td style={{ ...css.td, color: T.textMid, maxWidth: 140, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.autore}</td>
                <td style={{ ...css.td, fontSize: "11px", whiteSpace: "nowrap" }}>{r.editore_nome ?? <span style={{ color: T.orange }}>?</span>}</td>
                <td style={{ ...css.td, whiteSpace: "nowrap" }}>{r.prezzo != null ? `€ ${r.prezzo.toFixed(2)}` : "—"}</td>
                <td style={{ ...css.td, color: T.textMid, fontSize: "11px" }}>{r.uscita}</td>
                <td style={{ ...css.td, color: T.accent }}>{r.top_100 ? "★" : ""}</td>
                <td style={css.td} title={r.obiettivoFonte === "T" ? "Da tiratura" : ""}>{r.obiettivo ?? r.tiratura ?? ""}{r.obiettivoFonte === "T" && <span style={{ color: T.blue, fontSize: "10px" }}> T</span>}</td>
                <td style={{ ...css.td, color: T.textMid, fontSize: "11px", maxWidth: 180, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={r.note ?? ""}>{r.note}</td>
                <td style={{ ...css.td, fontSize: "11px", color: r.gemelli.some(g => !g.titolo) ? T.orange : T.textMid, whiteSpace: "nowrap" }} title={r.gemelli.map(g => `${g.ean ?? ""} ${g.titolo ?? "(titolo mancante)"}${g.titoloDaDb ? " [da Supabase]" : ""}`).join("\n")}>{r.gemelli.length ? `${r.gemelli.length} ⇄${r.gemelli.some(g => !g.titolo) ? " ⚠" : ""}` : ""}</td>
                <td style={{ ...css.td, color: T.red, fontSize: "11px", whiteSpace: "nowrap" }}>{r.dup ? <span style={{ color: T.textMid }}>EAN duplicato, ignorato</span> : r.errs.join(", ")}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
