import { useState, useEffect, useMemo, useCallback, useRef } from "react";
import { tema, cv } from "./tema.js";

// ─── Pesi Spalmatura ─────────────────────────────────────────────────────────
// Vista e modifica diretta della tabella spalmatura_obiettivo: per ogni editore+formato, la quota
// (in % 0-100) dell'obiettivo assegnata a ciascun canale. Da qui GiroManager calcola gli obiettivi per
// canale/gruppo (dashboard, cedola direzionale/agenti, Fine Giro) e BookUp gli obiettivi agente.
// • Una riga per editore+formato, una colonna per canale. Cella vuota = nessun peso su quel canale.
// • Le modifiche restano in bozza finché non premi Salva (come Ranking editori).
// • "Importa Excel" carica il template SPALMATURA nella griglia come modifiche non salvate;
//   "Esporta Excel" produce lo stesso template, quindi si può scaricare, modificare e ricaricare.
// • "Correggi con resa" carica il file delle rese sulle novità (stesso formato: editore, tipo edizione,
//   una colonna per canale + Totale = resa media editore) e corregge i pesi in griglia: chi rende meno
//   della media dell'editore guadagna peso, chi rende di più ne perde. Due metodi: "Regole a soglie"
//   (resa critica/alta → taglio, resa virtuosa → riceve) o "Formula proporzionale". Resta in bozza fino a Salva.

const SUPABASE_URL = "https://tdflwenlylhctxssatax.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InRkZmx3ZW5seWxoY3R4c3NhdGF4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYzMzgyNzYsImV4cCI6MjA5MTkxNDI3Nn0.l35qEL7LOvyYuI1McQlVqj4vbyTqmlevcmqWbTGYi2Q";

const T = tema({
  bg: "#1a2140", surface: "#212d54", border: "#2e3d6b", borderHi: "#3d4f82",
  text: "#f0f2f8", textMid: "#8b9cc8", textDim: "#4a5a8a",
  accent: "#7b9fe8", green: "#4caf7d", red: "#e05c5c", amber: "#e0a84c",
});

const css = {
  btn: (v = "default", disabled = false) => ({ padding: "6px 12px", border: `1px solid ${v === "accent" ? T.accent : v === "green" ? T.green : v === "danger" ? T.red : T.border}`, background: v === "accent" ? T.accent : v === "green" ? T.green : "transparent", color: v === "accent" || v === "green" ? "#000" : v === "danger" ? T.red : T.text, cursor: disabled ? "default" : "pointer", fontSize: "12px", fontFamily: "inherit", borderRadius: 3, fontWeight: v === "accent" || v === "green" ? 700 : 400, opacity: disabled ? 0.5 : 1, whiteSpace: "nowrap" }),
  mini: { padding: "2px 6px", border: `1px solid ${T.border}`, background: "transparent", color: T.textMid, cursor: "pointer", fontSize: "11px", fontFamily: "inherit", borderRadius: 3 },
  card: { background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, padding: 16, marginBottom: 14 },
  th: { padding: "7px 6px", textAlign: "left", color: T.textMid, fontWeight: 400, fontSize: "10px", letterSpacing: "0.06em", textTransform: "uppercase", borderBottom: `1px solid ${T.border}`, whiteSpace: "nowrap", position: "sticky", top: 0, background: T.surface, zIndex: 1 },
  td: { padding: "4px 6px", borderBottom: `1px solid ${T.border}55`, fontSize: "12px", color: T.text, whiteSpace: "nowrap", verticalAlign: "middle" },
  input: { background: T.bg, border: `1px solid ${T.borderHi}`, color: T.text, padding: "4px 6px", fontSize: "12px", fontFamily: "inherit", borderRadius: 3 },
  label: { display: "flex", flexDirection: "column", gap: 4, color: T.textMid, fontSize: "11px" },
};

// Ordine colonne: per gruppo di canale. I codici presenti nei dati ma non elencati qui vengono aggiunti in coda.
const CANALI_BASE = [
  "FELTRINELLI", "GIUNTI", "MONDADORI", "UBIK", "LIBRACCIO",
  "INDIPENDENTI_ALTRE_CATENE", "LIB_COOP", "LIB_RELIGIOSE",
  "AMAZON", "IBS", "ALTRI_ONLINE",
  "FASTBOOK", "CENTROLIBRI", "GROSSISTI",
];
const FORMATI = ["Cover", "Tascabile"];
// Formati come arrivano nei file obiettivi (es. "Tipo Edizione": Cover / Economici)
const FORMATO_DA_FILE = { COVER: "Cover", TASCABILE: "Tascabile", TASCABILI: "Tascabile", ECONOMICI: "Tascabile", ECONOMICO: "Tascabile" };
// Colonne di totale da ignorare (intestazioni normalizzate con "_")
const COLONNE_TOTALE = new Set(["SOMMA_%", "SOMMA", "TOTALE", "TOTALE_%", "TOTALE_COMPLESSIVO"]);

const norm = (s) => String(s ?? "").replace(/\s+/g, " ").trim().toUpperCase();
const deaccent = (s) => String(s ?? "").normalize("NFD").replace(/[̀-ͯ]/g, "");
// Stessa normalizzazione delle chiavi di alias_editori usata in ModuloImportEditore (normEd)
const normEd = (s) => deaccent(s).toUpperCase().replace(/[^A-Z0-9]+/g, " ").replace(/\bS R L S?\b|\bS P A\b|\bS A S\b|\bS N C\b/g, " ").replace(/\s+/g, " ").trim();
const PAROLE_GENERICHE = new Set(["EDITORE", "EDITORI", "EDIZIONI", "EDITRICE", "ED", "ITALIA", "SRL", "SRLS", "SPA", "SAS", "SNC"]);
// Nome ridotto per l'abbinamento: senza parole generiche né forma societaria (normEd non toglie "S.R.L." in fondo al nome)
const coreEd = (s) => normEd(s).replace(/ (S R L( S)?|S P A|S A S|S N C)$/, "").split(" ").filter(t => t && !PAROLE_GENERICHE.has(t)).join(" ");
const r2 = (n) => Math.round(n * 100) / 100;
const chiave = (editore, formato) => `${editore}|${formato}`;
const somma = (pesi) => r2(Object.values(pesi).reduce((s, v) => s + (Number(v) || 0), 0));
const sommaOk = (s) => s >= 99.5 && s <= 100.5;
const fmt = (v) => (v === "" || v == null ? "" : String(Number(v)));
const numOVuoto = (raw) => {
  if (raw === "" || raw == null) return "";
  const n = parseFloat(String(raw).replace(",", "."));
  return isNaN(n) ? null : r2(n);
};

const headers = (token, extra = {}) => ({ apikey: SUPABASE_KEY, Authorization: `Bearer ${token}`, ...extra });

async function fetchTutto(path, token) {
  const out = [];
  for (let from = 0; ; from += 1000) {
    const r = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, { headers: headers(token, { "Range-Unit": "items", Range: `${from}-${from + 999}` }) });
    if (!r.ok) throw new Error(`Errore caricamento ${path.split("?")[0]} (${r.status})`);
    const page = await r.json();
    out.push(...page);
    if (page.length < 1000) return out;
  }
}

export default function ModuloSpalmatura({ token, onDataChange }) {
  const [orig, setOrig] = useState({});          // chiave → { canale: percentuale } come da DB
  const [righe, setRighe] = useState([]);        // [{ key, editore_nome, formato, pesi }]
  const [canaliInfo, setCanaliInfo] = useState({});
  const [ranking, setRanking] = useState([]);    // anagrafica editori (ordine e categoria)
  const [loading, setLoading] = useState(true);
  const [errore, setErrore] = useState("");
  const [msg, setMsg] = useState("");
  const [saving, setSaving] = useState(false);
  const [filtro, setFiltro] = useState("");
  const [filtroFormato, setFiltroFormato] = useState("");
  const [soloFuori100, setSoloFuori100] = useState(false);
  const [showDiff, setShowDiff] = useState(false);
  const [showNuovo, setShowNuovo] = useState(false);
  const [showCopia, setShowCopia] = useState(null);   // null | { da, a } per il pannello "Copia pesi"
  const [nuoviEntranti, setNuoviEntranti] = useState([]);
  const [showSenzaPesi, setShowSenzaPesi] = useState(false);
  const [aliasMap, setAliasMap] = useState({});  // alias normalizzato → editore del Ranking
  const [linee, setLinee] = useState([]);        // [{ linea, madre }] linee figlie → casa madre
  const [abbina, setAbbina] = useState(null);    // import in attesa: editori del file da abbinare a mano
  const fileRef = useRef(null);
  const resaRef = useRef(null);
  const [resa, setResa] = useState(null);        // file rese caricato: { nomeFile, rese: chiave → { perCanale, totale }, nonAbbinati }
  const [parResa, setParResa] = useState({
    metodo: "regole",                       // "regole" (soglie di resa) | "formula" (proporzionale)
    critica: 25, taglioCritica: 30,         // resa ≥ 25% → il canale perde il 30% del suo peso
    alta: 10, taglioAlta: 10,               // resa ≥ 10% → perde il 10%
    virtuosa: 5, aumentoMax: 40,            // resa ≤ 5% → riceve le quote tolte, fino a +40% del suo peso
    pesoMinimo: 2,                          // canali sotto il 2% di peso: troppo piccoli per giudicare la resa
    // Eccezione Amazon (decisione di Alberto, 01/10/2026): peso fisso 10% per tutti, gli altri canali
    // riproporzionati sul resto; esclusi gli editori sotto (e le loro linee figlie). 0 = eccezione spenta.
    amazonFisso: 10,
    amazonEsclusi: ["ALPHA TEST", "ALPHA TEST TU", "RAFFAELLO CORTINA", "SOLFERINO", "SOLFERINO RAGAZZI", "CAIRO EDITORE"],
    tetto: 30, prudenza: 5, neutri: ["IBS", "ALTRI_ONLINE", "AMAZON"],
  });

  const carica = useCallback(async () => {
    setLoading(true); setErrore("");
    try {
      const [sp, can, rk, ne, al, li] = await Promise.all([
        fetchTutto("spalmatura_obiettivo?select=editore_nome,formato,canale_codice,percentuale&order=id", token),
        fetchTutto("canali?select=codice,nome,gruppo", token),
        fetchTutto("ranking_editori?select=editore_nome,ranking,cedola,attivo", token),
        fetchTutto("editori_new_entry?select=nome_editore&attivo=eq.true", token).catch(() => []),
        fetchTutto("alias_editori?select=alias,editore_nome", token).catch(() => []),
        fetchTutto("spalmatura_linee?select=linea,madre", token).catch(() => []),
      ]);
      setAliasMap(Object.fromEntries(al.map(a => [normEd(a.alias), a.editore_nome])));
      setLinee(li.map(l => ({ linea: norm(l.linea), madre: norm(l.madre) })));
      setNuoviEntranti([...new Set(ne.map(x => norm(x.nome_editore)).filter(Boolean))]);
      const o = {};
      sp.forEach(x => {
        const k = chiave(norm(x.editore_nome), x.formato);
        (o[k] = o[k] || {})[x.canale_codice] = x.percentuale == null ? "" : Number(x.percentuale);
      });
      setOrig(o);
      setRighe(Object.entries(o).map(([k, pesi]) => { const [e, f] = k.split("|"); return { key: k, editore_nome: e, formato: f, pesi: { ...pesi } }; }));
      setCanaliInfo(Object.fromEntries(can.map(c => [c.codice, c])));
      setRanking(rk.map(r => ({ ...r, editore_nome: norm(r.editore_nome) })));
    } catch (e) { setErrore(e.message); }
    setLoading(false);
  }, [token]);
  useEffect(() => { carica(); }, [carica]);

  const rankingPer = useMemo(() => Object.fromEntries(ranking.map(r => [r.editore_nome, r])), [ranking]);

  const canali = useMemo(() => {
    const presenti = new Set();
    Object.values(orig).forEach(p => Object.keys(p).forEach(c => presenti.add(c)));
    righe.forEach(r => Object.keys(r.pesi).forEach(c => presenti.add(c)));
    return [...CANALI_BASE, ...[...presenti].filter(c => !CANALI_BASE.includes(c)).sort()];
  }, [orig, righe]);

  // Ordine righe: come la sequenza di Ranking editori, poi nome, poi formato
  const ordinate = useMemo(() => [...righe].sort((a, b) => {
    const ra = rankingPer[a.editore_nome]?.ranking ?? 1e9, rb = rankingPer[b.editore_nome]?.ranking ?? 1e9;
    return (ra - rb) || a.editore_nome.localeCompare(b.editore_nome) || FORMATI.indexOf(a.formato) - FORMATI.indexOf(b.formato);
  }), [righe, rankingPer]);

  // ─── Modifiche ─────────────────────────────────────────────────────────────
  const setCella = (key, canale, raw) => {
    const v = numOVuoto(raw);
    if (v === null) return;
    setRighe(rs => rs.map(r => {
      if (r.key !== key) return r;
      const pesi = { ...r.pesi };
      if (v === "") delete pesi[canale]; else pesi[canale] = v;
      return { ...r, pesi };
    }));
  };

  // Riporta la riga a somma 100 mantenendo le proporzioni; lo scarto di arrotondamento va sul canale più pesante
  const normalizza = (key) => setRighe(rs => rs.map(r => {
    if (r.key !== key) return r;
    const s = somma(r.pesi);
    if (!s) return r;
    const pesi = Object.fromEntries(Object.entries(r.pesi).map(([c, v]) => [c, r2((Number(v) || 0) * 100 / s)]));
    const max = Object.keys(pesi).reduce((a, c) => (pesi[c] > (pesi[a] ?? -1) ? c : a), null);
    if (max) pesi[max] = r2(pesi[max] + 100 - somma(pesi));
    return { ...r, pesi };
  }));

  const ripristina = (key) => setRighe(rs => rs.map(r => (r.key === key && orig[key] ? { ...r, pesi: { ...orig[key] } } : r)));
  const elimina = (key) => setRighe(rs => rs.filter(r => r.key !== key));

  const aggiungi = ({ editore_nome, formato, copiaDa }) => {
    const k = chiave(editore_nome, formato);
    const base = copiaDa ? righe.find(r => r.key === copiaDa)?.pesi : null;
    setRighe(rs => [...rs, { key: k, editore_nome, formato, pesi: base ? { ...base } : {} }]);
    setShowNuovo(false);
    setFiltro(editore_nome);
  };

  // Copia i pesi di un editore su un altro (per i formati scelti): crea le righe mancanti,
  // sovrascrive quelle esistenti. Resta in bozza fino a Salva.
  const copiaPesi = ({ da, a, formati }) => {
    setRighe(rs => {
      const out = [...rs];
      formati.forEach(f => {
        const src = rs.find(r => r.key === chiave(da, f));
        if (!src) return;
        const k = chiave(a, f);
        const i = out.findIndex(r => r.key === k);
        const riga = { key: k, editore_nome: a, formato: f, pesi: { ...src.pesi } };
        if (i >= 0) out[i] = riga; else out.push(riga);
      });
      return out;
    });
    setShowCopia(null);
    setFiltro(a);
    setMsg(`Pesi di ${da} copiati su ${a} (${formati.join(" e ")}): controlla e premi Salva.`);
  };

  // ─── Differenze da salvare ─────────────────────────────────────────────────
  const diff = useMemo(() => {
    const upsert = [], cancellaCelle = [], nuove = [], modificate = [];
    const presenti = new Set(righe.map(r => r.key));
    righe.forEach(r => {
      const o = orig[r.key];
      if (!o) nuove.push(r);
      let cambiata = false;
      Object.entries(r.pesi).forEach(([c, v]) => {
        if (!o || String(o[c] ?? "") !== String(v)) { upsert.push({ editore_nome: r.editore_nome, formato: r.formato, canale_codice: c, percentuale: v }); cambiata = true; }
      });
      if (o) Object.keys(o).forEach(c => { if (!(c in r.pesi)) { cancellaCelle.push({ editore_nome: r.editore_nome, formato: r.formato, canale_codice: c }); cambiata = true; } });
      if (o && cambiata) modificate.push(r);
    });
    const eliminate = Object.keys(orig).filter(k => !presenti.has(k)).map(k => { const [e, f] = k.split("|"); return { key: k, editore_nome: e, formato: f }; });
    return { upsert, cancellaCelle, nuove, modificate, eliminate };
  }, [righe, orig]);
  const nModifiche = diff.nuove.length + diff.modificate.length + diff.eliminate.length;

  const salva = async () => {
    setSaving(true); setErrore(""); setMsg("");
    try {
      for (let i = 0; i < diff.upsert.length; i += 300) {
        const res = await fetch(`${SUPABASE_URL}/rest/v1/spalmatura_obiettivo?on_conflict=editore_nome,formato,canale_codice`, {
          method: "POST",
          headers: headers(token, { "Content-Type": "application/json", Prefer: "resolution=merge-duplicates,return=minimal" }),
          body: JSON.stringify(diff.upsert.slice(i, i + 300)),
        });
        if (!res.ok) { const j = await res.json().catch(() => ({})); throw new Error(j.message || `Errore salvataggio (${res.status})`); }
      }
      const cancella = async (e, f, c) => {
        const q = `editore_nome=eq.${encodeURIComponent(e)}&formato=eq.${encodeURIComponent(f)}${c ? `&canale_codice=eq.${encodeURIComponent(c)}` : ""}`;
        const res = await fetch(`${SUPABASE_URL}/rest/v1/spalmatura_obiettivo?${q}`, { method: "DELETE", headers: headers(token, { Prefer: "return=minimal" }) });
        if (!res.ok) throw new Error(`Errore eliminazione ${e} ${f} (${res.status})`);
      };
      for (const x of diff.cancellaCelle) await cancella(x.editore_nome, x.formato, x.canale_codice);
      for (const x of diff.eliminate) await cancella(x.editore_nome, x.formato);
      setMsg(`Salvato: ${diff.nuove.length} nuove righe · ${diff.modificate.length} modificate · ${diff.eliminate.length} eliminate (${diff.upsert.length} pesi scritti).`);
      setShowDiff(false);
      await carica();
      onDataChange && onDataChange();
    } catch (e) { setErrore(e.message); }
    setSaving(false);
  };

  // ─── Excel (stesso formato del template SPALMATURA: dati da riga 5) ────────
  const esporta = () => {
    const XLSX = window.XLSX;
    if (!XLSX) return;
    const aoa = [
      ["PESI SPALMATURA OBIETTIVO PER CANALE"],
      ["Una riga per editore+formato. Pesi in % (0-100): la somma di ogni riga deve fare 100. Cella vuota = nessun peso sul canale."],
      [`Esportato il ${new Date().toLocaleDateString("it")}`],
      ["EDITORE", "FORMATO", ...canali, "SOMMA %"],
      ...ordinate.map(r => [r.editore_nome, r.formato, ...canali.map(c => (c in r.pesi ? r.pesi[c] : "")), somma(r.pesi)]),
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 36 }, { wch: 10 }, ...canali.map(c => ({ wch: Math.max(9, c.length + 2) })), { wch: 9 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "SPALMATURA");
    XLSX.writeFile(wb, `pesi_spalmatura_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  // Legge il file e restituisce le righe come array di celle. Supporta xlsx/xls, CSV UTF-8 o
  // Windows-1252 con ; o , e i CSV UTF-16 separati da TAB (export dal gestionale obiettivi).
  const leggiRighe = (buf, nomeFile) => {
    const XLSX = window.XLSX;
    const b = new Uint8Array(buf);
    const utf16 = (b[0] === 0xff && b[1] === 0xfe) ? "utf-16le" : (b[0] === 0xfe && b[1] === 0xff) ? "utf-16be" : null;
    let wb;
    if (utf16 || /\.(csv|txt)$/i.test(nomeFile)) {
      let testo = new TextDecoder(utf16 || "utf-8").decode(b);
      if (!utf16 && testo.includes("�")) testo = new TextDecoder("windows-1252").decode(b);
      testo = testo.replace(/^﻿/, "");
      const inizio = testo.slice(0, 4000);
      const conta = (s) => inizio.split(s).length - 1;
      const sep = ["\t", ";", ","].sort((x, y) => conta(y) - conta(x))[0];
      wb = XLSX.read(testo, { type: "string", FS: sep, raw: true });
    } else {
      wb = XLSX.read(buf, { type: "array" });
    }
    // Foglio "SPALMATURA" (template); se manca ma il file ha un solo foglio (es. CSV), si usa quello.
    const ws = wb.Sheets["SPALMATURA"] || (wb.SheetNames.length === 1 ? wb.Sheets[wb.SheetNames[0]] : null);
    if (!ws) throw new Error(`Foglio 'SPALMATURA' non trovato nel file (fogli presenti: ${wb.SheetNames.join(", ")}).`);
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
  };

  // Dal nome editore del file all'editore del Ranking: nome uguale → alias salvato → nome ridotto
  // (senza EDITORE/EDIZIONI/SRL…) se c'è un solo candidato.
  const risolviEditore = (nomeFile) => {
    const n = normEd(nomeFile);
    // prima il nome identico: normEd rende uguali ad es. "LA NAVE DI TESEO" e "LA NAVE DI TESEO +"
    const perNorm = ranking.find(r => r.editore_nome === norm(nomeFile)) || ranking.find(r => normEd(r.editore_nome) === n);
    if (perNorm) return { rk: perNorm, via: "uguale" };
    const alias = aliasMap[n];
    const perAlias = alias && ranking.find(r => r.editore_nome === norm(alias));
    if (perAlias) return { rk: perAlias, via: "alias" };
    const c = coreEd(nomeFile);
    const cand = c ? ranking.filter(r => coreEd(r.editore_nome) === c) : [];
    if (cand.length === 1) return { rk: cand[0], via: "ridotto" };
    return { rk: null, via: null };
  };

  // Dalle righe del file (template o export obiettivi/rese) alle righe { nomeFile, formato, pesi, totale }.
  // somma=true: più colonne sullo stesso canale si sommano (pesi); false: vale la prima (rese).
  const leggiFileCanali = (data, { somma: sommaCol }) => {
    // Riga intestazioni: EDITORE | FORMATO (template) oppure EDITORE | TIPO EDIZIONE (file obiettivi),
    // cercata nelle prime 10 righe.
    let hIdx = data.slice(0, 10).findIndex(r => norm(r[0]) === "EDITORE" && ["FORMATO", "TIPO EDIZIONE"].includes(norm(r[1])));
    if (hIdx < 0) hIdx = 3;
    // Colonne canale: per codice (template) o per nome del canale (file obiettivi); GDO è dentro Fastbook
    const perNome = Object.fromEntries(Object.values(canaliInfo).map(c => [norm(c.nome), c.codice]));
    const colCanali = [], sconosciute = [];
    let colTotale = -1;
    (data[hIdx] || []).forEach((h, i) => {
      if (i < 2 || !norm(h)) return;
      const hc = norm(h).replace(/\s+/g, "_");
      if (COLONNE_TOTALE.has(hc)) { if (colTotale < 0) colTotale = i; return; }
      const cod = hc === "GDO" ? "FASTBOOK" : canaliInfo[hc] ? hc : perNome[norm(h)];
      if (cod) colCanali.push([cod, i]); else sconosciute.push(String(h).trim());
    });
    if (sconosciute.length) throw new Error(`Colonne non riconosciute come canali: ${sconosciute.join(", ")}. Usa i codici o i nomi dei canali.`);
    if (!colCanali.length) throw new Error("Intestazioni dei canali non trovate (serve una riga EDITORE | FORMATO o TIPO EDIZIONE | canali).");

    const errs = [];
    const lette = [];
    data.slice(hIdx + 1).forEach((r, idx) => {
      if (!r.some(v => String(v).trim() !== "")) return;
      const nRiga = idx + hIdx + 2;
      const nomeFile = String(r[0] ?? "").trim();
      const formato = FORMATO_DA_FILE[norm(r[1])];
      if (!nomeFile || !formato) { errs.push(`riga ${nRiga}: editore o formato non valido ("${r[1]}")`); return; }
      const pesi = {};
      colCanali.forEach(([c, i]) => {
        const v = numOVuoto(r[i]);
        if (v === null) errs.push(`riga ${nRiga}: valore non valido per ${c}`);
        else if (sommaCol) { if (v !== "" && v !== 0) pesi[c] = r2((pesi[c] || 0) + v); }
        else if (v !== "" && !(c in pesi)) pesi[c] = v;
      });
      const totale = colTotale >= 0 ? numOVuoto(r[colTotale]) : "";
      lette.push({ nomeFile, formato, pesi, totale: totale === null ? "" : totale });
    });
    if (errs.length) throw new Error(`File non caricato, correggi: ${errs.slice(0, 5).join("; ")}${errs.length > 5 ? ` e altri ${errs.length - 5}` : ""}`);
    return { lette };
  };

  // ─── Correzione con le rese ────────────────────────────────────────────────
  const caricaRese = (e) => {
    const f = e.target.files[0];
    e.target.value = "";
    if (!f) return;
    setErrore(""); setMsg("");
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const { lette } = leggiFileCanali(leggiRighe(evt.target.result, f.name), { somma: false });
        const rese = {}, nonAbbinati = new Set(), via = {};
        lette.forEach(l => {
          const { rk, via: v } = risolviEditore(l.nomeFile);
          if (!rk) { nonAbbinati.add(l.nomeFile); return; }
          const k = chiave(rk.editore_nome, l.formato);
          // se due righe del file finiscono sullo stesso editore, vince quella col nome identico
          if (rese[k] && v !== "uguale") return;
          rese[k] = { perCanale: l.pesi, totale: l.totale };
          via[k] = v;
        });
        // linee figlie senza riga propria nel file: usano la resa della casa madre
        linee.forEach(({ linea, madre }) => FORMATI.forEach(fm => {
          if (!rese[chiave(linea, fm)] && rese[chiave(madre, fm)]) rese[chiave(linea, fm)] = rese[chiave(madre, fm)];
        }));
        if (!Object.keys(rese).length) throw new Error("Nessun editore del file rese corrisponde a quelli del Ranking.");
        setResa({ nomeFile: f.name, rese, nonAbbinati: [...nonAbbinati] });
      } catch (err) { setErrore(err.message); }
    };
    reader.readAsArrayBuffer(f);
  };

  const applicaResa = () => {
    const perKey = Object.fromEntries(correzione.righe.map(c => [c.key, c.nuovi]));
    setRighe(rs => rs.map(r => (perKey[r.key] ? { ...r, pesi: perKey[r.key] } : r)));
    setMsg(`Pesi corretti con le rese su ${correzione.righe.length} righe: i valori cambiati sono evidenziati (passa sopra una cella per vedere il valore di prima). Controlla e premi Salva.`);
    setResa(null);
  };

  const importa = (e) => {
    const f = e.target.files[0];
    e.target.value = "";
    if (!f) return;
    setErrore(""); setMsg(""); setAbbina(null);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const { lette } = leggiFileCanali(leggiRighe(evt.target.result, f.name), { somma: true });
        const risolte = lette.map(l => ({ ...l, ...risolviEditore(l.nomeFile) }));
        const ignoti = [...new Set(risolte.filter(r => !r.rk).map(r => r.nomeFile))];
        if (ignoti.length) {
          // alcuni nomi vanno abbinati a mano: la griglia si aggiorna dopo la conferma
          setAbbina({ risolte, scelte: Object.fromEntries(ignoti.map(n => [n, ""])) });
        } else applicaImport(risolte, {});
      } catch (err) { setErrore(err.message); }
    };
    reader.readAsArrayBuffer(f);
  };

  // Porta in griglia (come bozza) le righe risolte. `scelte`: nome del file → editore del Ranking ("" = salta)
  const applicaImport = async (risolte, scelte) => {
    setErrore("");
    try {
      const nuoviAlias = Object.entries(scelte).filter(([, ed]) => ed).map(([nome, ed]) => ({ alias: normEd(nome), editore_nome: ed }));
      if (nuoviAlias.length) {
        const res = await fetch(`${SUPABASE_URL}/rest/v1/alias_editori?on_conflict=alias`, {
          method: "POST",
          headers: headers(token, { "Content-Type": "application/json", Prefer: "resolution=merge-duplicates,return=minimal" }),
          body: JSON.stringify(nuoviAlias),
        });
        if (!res.ok) throw new Error(`Errore salvataggio abbinamenti (${res.status})`);
        setAliasMap(m => ({ ...m, ...Object.fromEntries(nuoviAlias.map(a => [a.alias, a.editore_nome])) }));
      }
      const conteggi = { uguale: 0, ridotto: 0, alias: 0, manuale: 0 };
      const inattivi = new Set(), saltati = new Set(), doppioni = new Set();
      const finali = new Map();
      risolte.forEach(r => {
        let rk = r.rk, via = r.via;
        if (!rk && scelte[r.nomeFile]) { rk = ranking.find(x => x.editore_nome === scelte[r.nomeFile]); via = "manuale"; }
        if (!rk) { saltati.add(r.nomeFile); return; }
        if (rk.attivo === false) { inattivi.add(r.nomeFile); return; }
        const k = chiave(rk.editore_nome, r.formato);
        if (finali.has(k)) doppioni.add(`${rk.editore_nome} ${r.formato}`);
        else conteggi[via]++;
        finali.set(k, { key: k, editore_nome: rk.editore_nome, formato: r.formato, pesi: r.pesi });
      });
      // Linee figlie (spalmatura_linee): se nel file c'è la casa madre e non la linea, la linea eredita i pesi
      let ereditate = 0;
      linee.forEach(({ linea, madre }) => {
        const rkLinea = ranking.find(x => x.editore_nome === linea);
        if (rkLinea?.attivo === false) return;
        FORMATI.forEach(fm => {
          const m = finali.get(chiave(madre, fm));
          if (m && !finali.has(chiave(linea, fm))) { finali.set(chiave(linea, fm), { key: chiave(linea, fm), editore_nome: linea, formato: fm, pesi: { ...m.pesi }, _ereditata: true }); ereditate++; }
        });
      });
      const nuove = [...finali.values()].map(({ _ereditata, ...r }) => r);
      setRighe(rs => {
        const m = new Map(rs.map(r => [r.key, r]));
        nuove.forEach(r => m.set(r.key, r));
        return [...m.values()];
      });
      setAbbina(null);
      const parti = [
        `${conteggi.uguale} con nome uguale al Ranking`,
        conteggi.ridotto && `${conteggi.ridotto} abbinate togliendo EDITORE/EDIZIONI/SRL`,
        conteggi.alias && `${conteggi.alias} tramite alias`,
        conteggi.manuale && `${conteggi.manuale} abbinate ora (salvate per le prossime volte)`,
        ereditate && `${ereditate} linee figlie con i pesi della casa madre`,
      ].filter(Boolean);
      const esclusi = [
        inattivi.size && `inattivi nel Ranking: ${[...inattivi].join(", ")}`,
        saltati.size && `non abbinati: ${[...saltati].join(", ")}`,
        doppioni.size && `presenti due volte (tenuta l'ultima riga): ${[...doppioni].join(", ")}`,
      ].filter(Boolean);
      setMsg(`Caricate ${nuove.length} righe dal file (${parti.join(", ")}).${esclusi.length ? ` Saltati — ${esclusi.join("; ")}.` : ""} Controlla le modifiche evidenziate e premi Salva.`);
    } catch (err) { setErrore(err.message); }
  };

  // ─── Vista ─────────────────────────────────────────────────────────────────
  const stat = useMemo(() => {
    const fuori = righe.filter(r => !sommaOk(somma(r.pesi))).length;
    const conPesi = new Set(righe.map(r => r.editore_nome));
    const senzaPesi = ranking.filter(r => r.attivo !== false && !conPesi.has(r.editore_nome)).sort((a, b) => a.ranking - b.ranking);
    return { fuori, editori: conPesi.size, senzaPesi };
  }, [righe, ranking]);

  // Pesi corretti con le rese (calcolo spiegato in calcolaCorrezione, in fondo al file)
  const correzione = useMemo(() => {
    if (!resa) return null;
    const out = [];
    const calcola = parResa.metodo === "regole" ? calcolaRegole : calcolaCorrezione;
    righe.forEach(r => {
      const rs = resa.rese[r.key];
      if (!somma(r.pesi)) return;
      // senza rese nel file la riga riceve solo l'eccezione Amazon (se prevista)
      const { nuovi: calcolati, info } = rs ? calcola(r.pesi, rs, parResa) : { nuovi: { ...r.pesi }, info: {} };
      const nuovi = applicaAmazonFisso(r.editore_nome, calcolati, parResa);
      if (!rs && somma(Object.fromEntries(Object.keys(nuovi).map(c => [c, Math.abs((nuovi[c] || 0) - (Number(r.pesi[c]) || 0))]))) === 0) return;
      out.push({ key: r.key, prima: r.pesi, nuovi, info, rese: rs ? rs.perCanale : {} });
    });
    const senzaPesi = Object.keys(resa.rese).filter(k => !righe.some(r => r.key === k));
    const perCanale = canali.map(c => {
      const presenti = out.filter(x => c in x.prima);
      if (!presenti.length) return null;
      const media = (k) => r2(presenti.reduce((s, x) => s + (Number(x[k][c]) || 0), 0) / presenti.length);
      const delta = presenti.map(x => (Number(x.nuovi[c]) || 0) - (Number(x.prima[c]) || 0));
      const conResa = presenti.filter(x => x.rese[c] !== undefined && x.rese[c] !== "");
      const resaMedia = conResa.length ? r2(conResa.reduce((s, x) => s + Number(x.rese[c]), 0) / conResa.length) : null;
      const fascia = (f) => presenti.filter(x => x.info.fasce?.[c] === f).length;
      return { c, prima: media("prima"), dopo: media("nuovi"), su: delta.filter(d => d > 0.05).length, giu: delta.filter(d => d < -0.05).length,
        resaMedia, critica: fascia("critica"), alta: fascia("alta"), virtuosa: fascia("virtuosa") };
    }).filter(Boolean);
    const senzaBeneficiari = out.filter(x => x.info.senzaBeneficiari).map(x => x.key.replace("|", " · "));
    const conTagli = out.filter(x => x.info.tagli).length;
    const tagliRidotti = out.filter(x => x.info.tagliRidotti).length;
    return { righe: out, senzaPesi, perCanale, senzaBeneficiari, conTagli, tagliRidotti };
  }, [resa, righe, parResa, canali]);

  const filtrate = ordinate.filter(r =>
    (!filtro || r.editore_nome.includes(norm(filtro))) &&
    (!filtroFormato || r.formato === filtroFormato) &&
    (!soloFuori100 || !sommaOk(somma(r.pesi))));

  if (loading) return <div style={{ padding: 24, color: T.textMid }}>Caricamento pesi spalmatura…</div>;

  const Kpi = ({ label, value, color = T.text, onClick, title }) => (
    <div onClick={onClick} title={title} style={{ border: `1px solid ${color === T.text ? T.border : color + "66"}`, borderRadius: 4, padding: "6px 10px", minWidth: 80, cursor: onClick ? "pointer" : "default" }}>
      <div style={{ color: color === T.text ? T.textMid : color, fontSize: "10px", fontWeight: 700 }}>{label}</div>
      <div style={{ color: T.text, fontWeight: 700 }}>{value}</div>
    </div>
  );

  return (
    <div style={{ flex: 1, overflowY: "auto", padding: 24 }}>
      <div style={css.card}>
        <div style={{ display: "flex", alignItems: "flex-start", gap: 16, flexWrap: "wrap" }}>
          <div style={{ flex: "1 1 420px" }}>
            <div style={{ color: T.text, fontWeight: 700, fontSize: "13px", marginBottom: 6 }}>⚖ Pesi spalmatura</div>
            <div style={{ color: T.textMid, fontSize: "11px", lineHeight: 1.6 }}>
              Per ogni editore e formato, quanta parte dell'obiettivo (in %) va a ciascun canale: da qui escono gli obiettivi per canale e per gruppo
              di dashboard, cedole e Fine Giro. La somma di ogni riga deve fare <b style={{ color: T.text }}>100</b>: usa <b style={{ color: T.text }}>→100</b> per riproporzionarla.
              Cella vuota = nessun peso su quel canale. Nulla viene salvato finché non premi <b style={{ color: T.text }}>Salva</b>.
            </div>
          </div>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
            <Kpi label="EDITORI" value={stat.editori} />
            <Kpi label="RIGHE" value={righe.length} />
            <Kpi label="SOMMA ≠ 100" value={stat.fuori} color={stat.fuori ? T.amber : T.green} onClick={() => setSoloFuori100(s => !s)} title="Mostra solo le righe con somma diversa da 100" />
            <Kpi label="SENZA PESI" value={stat.senzaPesi.length} color={stat.senzaPesi.length ? T.red : T.green} onClick={() => setShowSenzaPesi(s => !s)} title="Editori attivi in Ranking editori senza pesi: i loro obiettivi per canale risultano 0" />
          </div>
        </div>
        {showSenzaPesi && stat.senzaPesi.length > 0 && (
          <div style={{ marginTop: 12, fontSize: "11px", color: T.textMid, lineHeight: 1.8 }}>
            <b style={{ color: T.red }}>Editori attivi senza pesi</b> (clicca per aggiungerli):{" "}
            {stat.senzaPesi.map(r => (
              <button key={r.editore_nome} style={{ ...css.mini, margin: "0 4px 4px 0" }} title="Copia su questo editore i pesi di un altro editore" onClick={() => { setShowNuovo(false); setShowCopia({ da: "", a: r.editore_nome }); setShowSenzaPesi(false); }}>{r.editore_nome}</button>
            ))}
          </div>
        )}
      </div>

      {/* Barra strumenti */}
      <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", marginBottom: 12 }}>
        <input style={{ ...css.input, width: 200, padding: "6px 8px" }} placeholder="Cerca editore…" value={filtro} onChange={e => setFiltro(e.target.value)} />
        <select style={{ ...css.input, padding: "6px 8px" }} value={filtroFormato} onChange={e => setFiltroFormato(e.target.value)}>
          <option value="">Tutti i formati</option>
          {FORMATI.map(f => <option key={f}>{f}</option>)}
        </select>
        <label style={{ color: T.textMid, fontSize: "12px", display: "flex", gap: 4, alignItems: "center" }}><input type="checkbox" checked={soloFuori100} onChange={e => setSoloFuori100(e.target.checked)} /> solo somma ≠ 100</label>
        <div style={{ flex: 1 }} />
        <button style={css.btn("green")} onClick={() => { setShowCopia(null); setShowNuovo(s => (s ? false : true)); }}>+ Nuovo editore</button>
        <button style={css.btn()} onClick={() => { setShowNuovo(false); setShowCopia(c => (c ? null : { da: "", a: "" })); }} title="Copia i pesi di un editore su un altro, anche nuovo entrante">⧉ Copia pesi</button>
        <button style={css.btn()} onClick={() => fileRef.current?.click()} title="Carica un file (.xlsx, .xls o .csv) nel formato del template SPALMATURA: le righe del file sostituiscono quelle in griglia (da salvare)">Importa Excel / CSV</button>
        <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv,.txt" style={{ display: "none" }} onChange={importa} />
        <button style={css.btn()} onClick={() => resaRef.current?.click()} title="Carica il file delle rese sulle novità per editore e canale: i pesi in griglia vengono corretti (chi rende meno guadagna peso). Resta in bozza fino a Salva.">↺ Correggi con resa</button>
        <input ref={resaRef} type="file" accept=".xlsx,.xls,.csv,.txt" style={{ display: "none" }} onChange={caricaRese} />
        <button style={css.btn()} onClick={esporta}>Esporta Excel</button>
      </div>

      {resa && correzione && (
        <PannelloResa resa={resa} correzione={correzione} par={parResa} setPar={setParResa} canaliInfo={canaliInfo}
          onApplica={applicaResa} onAnnulla={() => setResa(null)} />
      )}

      {abbina && (
        <div style={{ ...css.card, borderColor: T.amber }}>
          <div style={{ color: T.amber, fontWeight: 700, fontSize: "12px", marginBottom: 6 }}>Editori del file da abbinare ({Object.keys(abbina.scelte).length})</div>
          <div style={{ color: T.textMid, fontSize: "11px", marginBottom: 10 }}>
            Questi nomi non corrispondono a nessun editore del Ranking. Scegli a chi appartengono, oppure lascia "salta".
            L'abbinamento viene ricordato: la prossima volta lo stesso nome si riconosce da solo.
          </div>
          <table style={{ borderCollapse: "collapse", marginBottom: 12 }}>
            <tbody>
              {Object.entries(abbina.scelte).map(([nome, scelta]) => (
                <tr key={nome}>
                  <td style={{ ...css.td, fontWeight: 600 }}>{nome}</td>
                  <td style={css.td}>→</td>
                  <td style={css.td}>
                    <select style={{ ...css.input, minWidth: 260, color: scelta ? T.text : T.textDim }} value={scelta}
                      onChange={e => setAbbina(a => ({ ...a, scelte: { ...a.scelte, [nome]: e.target.value } }))}>
                      <option value="">— salta —</option>
                      {ranking.filter(r => r.attivo !== false).map(r => r.editore_nome).sort().map(n => <option key={n} value={n}>{n}</option>)}
                    </select>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          <div style={{ display: "flex", gap: 8 }}>
            <button style={css.btn("accent")} onClick={() => applicaImport(abbina.risolte, abbina.scelte)}>Conferma e carica</button>
            <button style={css.btn()} onClick={() => setAbbina(null)}>Annulla import</button>
          </div>
        </div>
      )}

      {showCopia && <CopiaPesi key={`${showCopia.da}|${showCopia.a}`} righe={righe} ranking={ranking} nuoviEntranti={nuoviEntranti} iniziale={showCopia} onAnnulla={() => setShowCopia(null)} onCopia={copiaPesi} />}

      {showNuovo && <NuovaRiga righe={righe} ranking={ranking} iniziale={typeof showNuovo === "string" ? showNuovo : ""} onAnnulla={() => setShowNuovo(false)} onAggiungi={aggiungi} />}

      {errore && <div style={{ color: T.red, fontSize: "12px", marginBottom: 10 }}>⚠ {errore}</div>}
      {msg && <div style={{ color: T.green, fontSize: "12px", marginBottom: 10 }}>✓ {msg}</div>}

      {/* Barra salvataggio */}
      {nModifiche > 0 && (
        <div style={{ ...css.card, borderColor: T.amber, display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap", position: "sticky", top: 0, zIndex: 3 }}>
          <span style={{ color: T.amber, fontWeight: 700, fontSize: "12px" }}>{nModifiche} modifiche non salvate</span>
          <span style={{ color: T.textMid, fontSize: "11px" }}>{diff.nuove.length} nuove · {diff.modificate.length} modificate · {diff.eliminate.length} eliminate</span>
          <button style={css.mini} onClick={() => setShowDiff(s => !s)}>{showDiff ? "Nascondi" : "Vedi"} dettaglio</button>
          <div style={{ flex: 1 }} />
          <button style={css.btn()} onClick={() => setRighe(Object.entries(orig).map(([k, pesi]) => { const [e, f] = k.split("|"); return { key: k, editore_nome: e, formato: f, pesi: { ...pesi } }; }))}>Annulla</button>
          <button style={css.btn("accent", saving)} disabled={saving} onClick={salva}>{saving ? "Salvataggio…" : "Salva"}</button>
          {showDiff && (
            <div style={{ width: "100%", fontSize: "11px", color: T.textMid, maxHeight: 180, overflow: "auto" }}>
              {diff.nuove.map(r => <div key={r.key} style={{ color: T.green }}>+ {r.editore_nome} · {r.formato} (somma {somma(r.pesi)}%)</div>)}
              {diff.modificate.map(r => {
                const o = orig[r.key];
                const cambi = [...new Set([...Object.keys(o), ...Object.keys(r.pesi)])].filter(c => String(o[c] ?? "") !== String(r.pesi[c] ?? ""));
                return <div key={r.key}>✎ {r.editore_nome} · {r.formato}: {cambi.map(c => `${c} ${fmt(o[c]) || "—"} → ${fmt(r.pesi[c]) || "—"}`).join(" · ")}</div>;
              })}
              {diff.eliminate.map(r => <div key={r.key} style={{ color: T.red }}>− {r.editore_nome} · {r.formato}</div>)}
            </div>
          )}
        </div>
      )}

      <div style={{ border: `1px solid ${T.border}`, borderRadius: 4, overflow: "auto", maxHeight: "70vh" }}>
        <table style={{ borderCollapse: "collapse", minWidth: "100%" }}>
          <thead>
            <tr>
              <th style={{ ...css.th, left: 0, zIndex: 2 }}>Editore</th>
              <th style={css.th}>Formato</th>
              {canali.map(c => (
                <th key={c} style={{ ...css.th, textAlign: "right" }} title={`${c}${canaliInfo[c]?.gruppo ? ` · gruppo ${canaliInfo[c].gruppo}` : ""}`}>
                  <div style={{ color: T.textDim, fontSize: "9px" }}>{canaliInfo[c]?.gruppo || "—"}</div>
                  {canaliInfo[c]?.nome || c}
                </th>
              ))}
              <th style={{ ...css.th, textAlign: "right" }}>Somma</th>
              <th style={css.th}></th>
            </tr>
          </thead>
          <tbody>
            {filtrate.map(r => {
              const o = orig[r.key];
              const s = somma(r.pesi);
              const ok = sommaOk(s);
              const rk = rankingPer[r.editore_nome];
              return (
                <tr key={r.key} style={{ background: !o ? T.green + "14" : "transparent" }}>
                  <td style={{ ...css.td, fontWeight: 600, position: "sticky", left: 0, background: T.bg, zIndex: 1 }}>
                    {r.editore_nome}
                    {!o && <span style={{ color: T.green, fontSize: "10px", marginLeft: 6 }}>NUOVO</span>}
                    {!rk && <span style={{ color: T.amber, fontSize: "10px", marginLeft: 6 }} title="Nome non presente in Ranking editori: controlla che sia scritto come nei titoli">⚠ non in ranking</span>}
                    {rk?.attivo === false && <span style={{ color: T.textDim, fontSize: "10px", marginLeft: 6 }}>uscito</span>}
                  </td>
                  <td style={{ ...css.td, color: T.textMid }}>{r.formato}</td>
                  {canali.map(c => {
                    const v = c in r.pesi ? r.pesi[c] : "";
                    const cambiato = !o || String(o[c] ?? "") !== String(v);
                    return (
                      <td key={c} style={{ ...css.td, textAlign: "right" }}>
                        <input key={`${r.key}-${c}-${v}`} inputMode="decimal"
                          style={{ ...css.input, width: 54, textAlign: "right", borderColor: cambiato ? T.amber : T.border, color: v === "" ? T.textDim : T.text }}
                          defaultValue={fmt(v)} placeholder="—" title={o && cambiato ? `era ${fmt(o[c]) || "vuoto"}` : undefined}
                          onBlur={e => { if (e.target.value !== fmt(v)) setCella(r.key, c, e.target.value); }}
                          onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); if (e.key === "Escape") { e.currentTarget.value = fmt(v); e.currentTarget.blur(); } }} />
                      </td>
                    );
                  })}
                  <td style={{ ...css.td, textAlign: "right", color: ok ? T.green : T.amber, fontWeight: 700 }}>{s}%</td>
                  <td style={css.td}>
                    {!ok && s > 0 && <><button style={{ ...css.mini, color: T.accent, borderColor: T.accent }} title="Riproporziona i pesi perché la somma faccia 100" onClick={() => normalizza(r.key)}>→100</button>{" "}</>}
                    {o && <><button style={css.mini} title="Ripristina i valori salvati" onClick={() => ripristina(r.key)}>↺</button>{" "}</>}
                    <button style={css.mini} title="Copia i pesi di questo editore su un altro" onClick={() => { setShowNuovo(false); setShowCopia({ da: r.editore_nome, a: "" }); }}>⧉</button>{" "}
                    <button style={{ ...css.mini, color: T.red }} title="Elimina la riga (tutti i pesi di questo editore+formato)" onClick={() => elimina(r.key)}>✕</button>
                  </td>
                </tr>
              );
            })}
            {filtrate.length === 0 && <tr><td colSpan={canali.length + 4} style={{ ...css.td, color: T.textDim, padding: 16 }}>Nessuna riga.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ─── Copia pesi da un editore a un altro ─────────────────────────────────────
function CopiaPesi({ righe, ranking, nuoviEntranti, iniziale, onCopia, onAnnulla }) {
  const [da, setDa] = useState(iniziale.da || "");
  const [a, setA] = useState(iniziale.a || "");
  const [scelti, setScelti] = useState(null); // null = tutti i formati disponibili
  const nomeDa = norm(da), nomeA = norm(a);

  const conPesi = useMemo(() => [...new Set(righe.map(r => r.editore_nome))].sort(), [righe]);
  const destinazioni = useMemo(() => {
    const senza = new Set(conPesi);
    // prima i nuovi entranti e gli editori senza pesi, poi tutti gli altri
    const tutti = [...new Set([...nuoviEntranti, ...ranking.map(r => r.editore_nome)])];
    return [...tutti.filter(n => !senza.has(n)).sort(), ...tutti.filter(n => senza.has(n)).sort()];
  }, [conPesi, ranking, nuoviEntranti]);

  const formatiDa = FORMATI.filter(f => righe.some(r => r.key === chiave(nomeDa, f)));
  const formati = (scelti ?? formatiDa).filter(f => formatiDa.includes(f));
  const sovrascritti = formati.filter(f => righe.some(r => r.key === chiave(nomeA, f)));
  const noto = !nomeA || ranking.some(r => r.editore_nome === nomeA) || nuoviEntranti.includes(nomeA);
  const ok = nomeDa && nomeA && nomeDa !== nomeA && formati.length > 0;

  return (
    <div style={{ ...css.card, borderColor: T.accent }}>
      <div style={{ color: T.accent, fontWeight: 700, fontSize: "12px", marginBottom: 10 }}>⧉ Copia pesi da un editore a un altro</div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))", gap: 10 }}>
        <label style={css.label}>Copia i pesi di *
          <input style={css.input} list="gm-spalm-da" value={da} onChange={e => { setDa(e.target.value); setScelti(null); }} placeholder="editore con pesi" autoFocus={!iniziale.da} />
          <datalist id="gm-spalm-da">{conPesi.map(n => <option key={n} value={n} />)}</datalist>
        </label>
        <label style={css.label}>Sull'editore *
          <input style={css.input} list="gm-spalm-a" value={a} onChange={e => setA(e.target.value)} placeholder="anche nuovo entrante" autoFocus={!!iniziale.da} />
          <datalist id="gm-spalm-a">{destinazioni.map(n => <option key={n} value={n}>{conPesi.includes(n) ? "ha già pesi" : nuoviEntranti.includes(n) ? "nuovo entrante · senza pesi" : "senza pesi"}</option>)}</datalist>
        </label>
        <div style={css.label}>Formati
          <div style={{ display: "flex", gap: 12, alignItems: "center", minHeight: 26 }}>
            {nomeDa && formatiDa.length === 0 && <span style={{ color: T.red }}>questo editore non ha pesi</span>}
            {formatiDa.map(f => (
              <label key={f} style={{ display: "flex", gap: 4, alignItems: "center", color: T.text }}>
                <input type="checkbox" checked={formati.includes(f)} onChange={e => setScelti(e.target.checked ? [...new Set([...formati, f])] : formati.filter(x => x !== f))} /> {f}
              </label>
            ))}
          </div>
        </div>
      </div>
      {nomeDa && nomeDa === nomeA && <div style={{ color: T.red, fontSize: "11px", marginTop: 8 }}>Scegli due editori diversi.</div>}
      {sovrascritti.length > 0 && nomeDa !== nomeA && <div style={{ color: T.amber, fontSize: "11px", marginTop: 8 }}>⚠ {nomeA} ha già pesi {sovrascritti.join(" e ")}: verranno sostituiti (fino a Salva puoi sempre annullare).</div>}
      {!noto && <div style={{ color: T.amber, fontSize: "11px", marginTop: 8 }}>{nomeA} non è in Ranking editori né tra i nuovi editori: i pesi valgono solo se il nome coincide con quello dei titoli.</div>}
      <div style={{ display: "flex", gap: 8, marginTop: 12, alignItems: "center" }}>
        <button style={css.btn("accent", !ok)} disabled={!ok} onClick={() => onCopia({ da: nomeDa, a: nomeA, formati })}>Copia</button>
        <button style={css.btn()} onClick={onAnnulla}>Annulla</button>
        {ok && <span style={{ color: T.textMid, fontSize: "11px" }}>{nomeDa} → {nomeA} · {formati.join(" + ")} — poi premi Salva</span>}
      </div>
    </div>
  );
}

// ─── Form nuova riga editore+formato ─────────────────────────────────────────
function NuovaRiga({ righe, ranking, iniziale, onAggiungi, onAnnulla }) {
  const [editore, setEditore] = useState(iniziale);
  const [formato, setFormato] = useState("Cover");
  const [copiaDa, setCopiaDa] = useState("");
  const nome = norm(editore);
  const esiste = righe.some(r => r.key === chiave(nome, formato));
  const inRanking = ranking.some(r => r.editore_nome === nome);
  const ok = nome && !esiste;
  const opzioniCopia = [...righe].sort((a, b) => a.editore_nome.localeCompare(b.editore_nome) || a.formato.localeCompare(b.formato));

  return (
    <div style={{ ...css.card, borderColor: T.green }}>
      <div style={{ color: T.green, fontWeight: 700, fontSize: "12px", marginBottom: 10 }}>+ Nuovo editore / formato</div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(170px, 1fr))", gap: 10 }}>
        <label style={{ ...css.label, gridColumn: "span 2" }}>Editore *
          <input style={css.input} list="gm-spalm-editori" value={editore} onChange={e => setEditore(e.target.value)} placeholder="come in Ranking editori" autoFocus />
          <datalist id="gm-spalm-editori">{ranking.map(r => <option key={r.editore_nome} value={r.editore_nome} />)}</datalist>
        </label>
        <label style={css.label}>Formato *
          <select style={css.input} value={formato} onChange={e => setFormato(e.target.value)}>{FORMATI.map(f => <option key={f}>{f}</option>)}</select>
        </label>
        <label style={{ ...css.label, gridColumn: "span 2" }}>Parti dai pesi di
          <select style={css.input} value={copiaDa} onChange={e => setCopiaDa(e.target.value)}>
            <option value="">— riga vuota —</option>
            {opzioniCopia.map(r => <option key={r.key} value={r.key}>{r.editore_nome} · {r.formato}</option>)}
          </select>
        </label>
      </div>
      {esiste && <div style={{ color: T.red, fontSize: "11px", marginTop: 8 }}>Questo editore ha già una riga {formato}: modificala direttamente in tabella.</div>}
      {nome && !esiste && !inRanking && <div style={{ color: T.amber, fontSize: "11px", marginTop: 8 }}>Nome non presente in Ranking editori: i pesi valgono solo se il nome coincide con quello dei titoli.</div>}
      <div style={{ display: "flex", gap: 8, marginTop: 12, alignItems: "center" }}>
        <button style={css.btn("green", !ok)} disabled={!ok} onClick={() => onAggiungi({ editore_nome: nome, formato, copiaDa })}>Aggiungi</button>
        <button style={css.btn()} onClick={onAnnulla}>Annulla</button>
        {ok && <span style={{ color: T.textMid, fontSize: "11px" }}>Poi compila i pesi in tabella e premi Salva</span>}
      </div>
    </div>
  );
}

// ─── Correzione pesi con le rese ─────────────────────────────────────────────
// Idea: conta quello che il canale VENDE, non quello che gli spedisci.
//   1. resa del canale "prudente": nei canali piccoli pochi titoli fanno impazzire la percentuale,
//      quindi la resa viene avvicinata alla media dell'editore (più il canale è piccolo, più ci si avvicina).
//   2. fattore = (1 − resa canale) / (1 − resa media editore): sopra 1 il canale trattiene più della media.
//   3. nuovo peso = peso × fattore, con un tetto (es. ±30%), poi la riga torna alla somma di partenza.
// Canali "neutri" (rese non tracciate, es. IBS) o senza dato di resa usano la resa media: restano proporzionali.
// Base comune: voci con peso > 0, resa media editore e resa "prudente" di ogni canale (null = neutro/senza dato)
// regole=true: resa reale del canale (non ammorbidita); i canali sotto `pesoMinimo` sono troppo piccoli per giudicare → neutri
function preparaRese(pesi, { perCanale, totale }, { prudenza, neutri, pesoMinimo }, regole = false) {
  const voci = Object.entries(pesi).map(([c, v]) => [c, Number(v) || 0]).filter(([, v]) => v > 0);
  const sommaPrima = voci.reduce((s, [, v]) => s + v, 0);
  // resa media editore: colonna Totale del file, altrimenti media delle rese pesata sui pesi
  let media = totale === "" || totale == null ? null : Number(totale) / 100;
  if (media == null) {
    const conResa = voci.filter(([c]) => perCanale[c] !== undefined && perCanale[c] !== "");
    const p = conResa.reduce((s, [, v]) => s + v, 0);
    media = p ? conResa.reduce((s, [c, v]) => s + v * Number(perCanale[c]) / 100, 0) / p : 0;
  }
  const k = Math.max(0, Number(prudenza) || 0) / 100;
  const resaDi = {};
  voci.forEach(([c, v]) => {
    const rc = perCanale[c];
    if (neutri.includes(c) || rc === undefined || rc === "") { resaDi[c] = null; return; }
    if (regole) { resaDi[c] = v < (Number(pesoMinimo) || 0) ? null : Number(rc) / 100; return; }
    const quota = v / 100;
    resaDi[c] = (quota * Number(rc) / 100 + k * media) / (quota + k || 1);
  });
  return { voci, sommaPrima, media, resaDi };
}

// Riporta i valori alla somma di partenza, scarto di arrotondamento sul canale più pesante
function chiudiRiga(grezzi, sommaPrima) {
  const sommaGrezzi = grezzi.reduce((s, [, v]) => s + v, 0) || 1;
  const nuovi = Object.fromEntries(grezzi.map(([c, v]) => [c, r2(v * sommaPrima / sommaGrezzi)]));
  const max = Object.keys(nuovi).reduce((a, c) => (nuovi[c] > (nuovi[a] ?? -1) ? c : a), null);
  if (max) nuovi[max] = r2(nuovi[max] + r2(sommaPrima) - somma(nuovi));
  return nuovi;
}

function calcolaCorrezione(pesi, rs, par) {
  const { voci, sommaPrima, media, resaDi } = preparaRese(pesi, rs, par);
  if (media >= 1) return { nuovi: { ...pesi }, info: {} };
  const cap = Math.max(0, Number(par.tetto) || 0) / 100;
  const grezzi = voci.map(([c, v]) => {
    const resaUsata = resaDi[c] == null ? media : resaDi[c];
    const fattore = Math.min(1 + cap, Math.max(1 - cap, (1 - resaUsata) / (1 - media)));
    return [c, v * fattore];
  });
  return { nuovi: chiudiRiga(grezzi, sommaPrima), info: {} };
}

// ─── Eccezione Amazon a peso fisso ────────────────────────────────────────────
// Amazon prende `amazonFisso`% (se l'editore ha già un peso Amazon e non è tra gli esclusi);
// gli altri canali si riproporzionano sul resto mantenendo i loro rapporti.
function applicaAmazonFisso(editore, pesi, { amazonFisso, amazonEsclusi }) {
  const fisso = Number(amazonFisso) || 0;
  const a = Number(pesi.AMAZON) || 0;
  if (!fisso || !a || (amazonEsclusi || []).includes(editore)) return pesi;
  const tot = Object.values(pesi).reduce((s, v) => s + (Number(v) || 0), 0);
  if (tot - a <= 0 || fisso >= tot) return pesi;
  const altri = Object.entries(pesi).filter(([c, v]) => c !== "AMAZON" && Number(v) > 0).map(([c, v]) => [c, Number(v)]);
  const nuovi = chiudiRiga(altri, tot - fisso);
  return { ...pesi, ...nuovi, AMAZON: fisso };
}

// ─── Regole a soglie ─────────────────────────────────────────────────────────
// 1. Ogni canale finisce in una fascia in base alla sua resa (prudente):
//      critica (≥ soglia critica) → perde taglioCritica% del suo peso
//      alta    (≥ soglia alta)    → perde taglioAlta% del suo peso
//      virtuosa (≤ soglia virtuosa) → riceve
//      normale / neutro / senza dato → invariato
// 2. Le quote tolte vanno ai canali virtuosi, in proporzione al loro peso, ciascuno al massimo +aumentoMax%.
//    Quello che avanza (o tutto, se non ci sono virtuosi) va ai canali "normali", con lo stesso tetto;
//    se resta ancora qualcosa torna ai canali tagliati (il taglio si riduce). I neutri non ricevono mai.
//    Se tutti i canali giudicabili sono alta/critica, le quote si redistribuiscono tra loro favorendo chi rende meno.
//    Qui la resa è quella reale; i canali sotto pesoMinimo (es. 2%) sono troppo piccoli per giudicare → invariati.
function calcolaRegole(pesi, rs, par) {
  const { voci, sommaPrima, resaDi } = preparaRese(pesi, rs, par, true);
  const pct = (x) => Math.max(0, Number(x) || 0) / 100;
  const fasce = {};
  voci.forEach(([c]) => {
    const r = resaDi[c];
    fasce[c] = r == null ? "neutro" : r >= pct(par.critica) ? "critica" : r >= pct(par.alta) ? "alta" : r <= pct(par.virtuosa) ? "virtuosa" : "normale";
  });
  const nuovo = Object.fromEntries(voci);
  let pool = 0;
  voci.forEach(([c, v]) => {
    const t = fasce[c] === "critica" ? pct(par.taglioCritica) : fasce[c] === "alta" ? pct(par.taglioAlta) : 0;
    if (t) { nuovo[c] = v * (1 - Math.min(1, t)); pool += v - nuovo[c]; }
  });
  if (pool <= 0) return { nuovi: { ...pesi }, info: { fasce } };

  // distribuisce `quanto` su `canali` in proporzione a `base`, rispettando un tetto per canale
  const distribuisci = (quanto, canali, base, tetto) => {
    let attivi = canali.filter(c => base(c) > 0);
    while (quanto > 1e-9 && attivi.length) {
      const tot = attivi.reduce((s, c) => s + base(c), 0);
      let avanzo = 0;
      const ancora = [];
      attivi.forEach(c => {
        const quota = quanto * base(c) / tot;
        const spazio = tetto(c) - nuovo[c];
        if (quota >= spazio) { nuovo[c] += Math.max(0, spazio); avanzo += quota - Math.max(0, spazio); }
        else { nuovo[c] += quota; ancora.push(c); }
      });
      quanto = avanzo; attivi = ancora;
    }
    return quanto;
  };
  const orig = Object.fromEntries(voci);
  const virtuosi = voci.filter(([c]) => fasce[c] === "virtuosa").map(([c]) => c);
  const normali = voci.filter(([c]) => fasce[c] === "normale").map(([c]) => c);
  const tagliati = voci.filter(([c]) => ["critica", "alta"].includes(fasce[c])).map(([c]) => c);
  const senzaBeneficiari = virtuosi.length === 0;
  const max = c => orig[c] * (1 + pct(par.aumentoMax));
  let resto = distribuisci(pool, virtuosi, c => orig[c], max);
  if (resto > 1e-9) resto = distribuisci(resto, normali, c => orig[c] * (1 - resaDi[c]), max);
  let tagliRidotti = false;
  if (resto > 1e-9) {
    if (!virtuosi.length && !normali.length) {
      // tutti i canali giudicabili sono in fascia alta/critica: le quote tornano a loro, di più a chi rende meno
      distribuisci(resto, tagliati, c => orig[c] * (1 - resaDi[c]) ** 2, () => Infinity);
    } else {
      // chi riceve è già al massimo: quel che avanza torna ai canali tagliati (il taglio si riduce)
      distribuisci(resto, tagliati, c => orig[c] - nuovo[c], c => orig[c]);
      tagliRidotti = true;
    }
  }
  const nuovi = chiudiRiga(voci.map(([c]) => [c, nuovo[c]]), sommaPrima);
  return { nuovi, info: { fasce, tagli: true, senzaBeneficiari, tagliRidotti } };
}

function PannelloResa({ resa, correzione, par, setPar, canaliInfo, onApplica, onAnnulla }) {
  const num = (campo) => (e) => { const v = parseFloat(String(e.target.value).replace(",", ".")); setPar(p => ({ ...p, [campo]: isNaN(v) ? 0 : v })); };
  const toggleNeutro = (c) => setPar(p => ({ ...p, neutri: p.neutri.includes(c) ? p.neutri.filter(x => x !== c) : [...p.neutri, c] }));
  return (
    <div style={{ ...css.card, borderColor: T.accent }}>
      <div style={{ color: T.accent, fontWeight: 700, fontSize: "12px", marginBottom: 6 }}>↺ Correzione pesi con le rese — {resa.nomeFile}</div>
      <div style={{ color: T.textMid, fontSize: "11px", lineHeight: 1.6, marginBottom: 12 }}>
        Il peso di ogni canale viene ricalcolato su quello che <b style={{ color: T.text }}>vende davvero</b> (fornito meno rese), non su quello che gli spedisci.
        Chi rende meno della media dell'editore guadagna peso, chi rende di più ne perde. La somma di ogni riga non cambia.
      </div>
      <div style={{ display: "flex", gap: 6, marginBottom: 12 }}>
        {[["regole", "Regole a soglie"], ["formula", "Formula proporzionale"]].map(([m, l]) => (
          <button key={m} style={css.btn(par.metodo === m ? "accent" : "default")} onClick={() => setPar(p => ({ ...p, metodo: m }))}>{l}</button>
        ))}
      </div>
      {par.metodo === "regole" ? (
        <div style={{ background: T.bg, border: `1px solid ${T.border}`, borderRadius: 4, padding: "10px 12px", marginBottom: 12, fontSize: "12px", color: T.text, lineHeight: 2.2 }}>
          <div><span style={{ color: T.red, fontWeight: 700 }}>● Resa critica</span> — resa ≥ <Num v={par.critica} on={num("critica")} />% → il canale perde il <Num v={par.taglioCritica} on={num("taglioCritica")} />% del suo peso</div>
          <div><span style={{ color: T.amber, fontWeight: 700 }}>● Resa alta</span> — resa ≥ <Num v={par.alta} on={num("alta")} />% → perde il <Num v={par.taglioAlta} on={num("taglioAlta")} />%</div>
          <div><span style={{ color: T.green, fontWeight: 700 }}>● Resa virtuosa</span> — resa ≤ <Num v={par.virtuosa} on={num("virtuosa")} />% → riceve le quote tolte agli altri, al massimo +<Num v={par.aumentoMax} on={num("aumentoMax")} />% del suo peso</div>
          <div style={{ color: T.textMid, fontSize: "11px", lineHeight: 1.5, marginTop: 4 }}>
            Tra le soglie il peso resta invariato. Se i virtuosi sono già al massimo (o mancano), le quote vanno ai canali intermedi; se non c'è spazio, il taglio si riduce. I canali neutri (non affidabili o troppo piccoli) non cambiano.
          </div>
        </div>
      ) : (
        <div style={{ color: T.textMid, fontSize: "11px", marginBottom: 12 }}>Ogni canale sale o scende in proporzione a quanto la sua resa è sotto o sopra la media dell'editore.</div>
      )}
      <div style={{ background: T.bg, border: `1px solid ${T.border}`, borderRadius: 4, padding: "8px 12px", marginBottom: 12, fontSize: "12px", color: T.text, lineHeight: 2 }}>
        <b>Eccezione Amazon</b> — peso fisso <Num v={par.amazonFisso} on={num("amazonFisso")} />% per tutti gli editori (0 = spenta), gli altri canali si riproporzionano.
        <div style={{ color: T.textMid, fontSize: "11px" }}>Esclusi (separati da virgola):{" "}
          <input style={{ ...css.input, width: "min(560px, 100%)", padding: "2px 6px" }} defaultValue={(par.amazonEsclusi || []).join(", ")}
            onBlur={e => { const v = e.target.value.split(",").map(x => x.trim().toUpperCase()).filter(Boolean); setPar(p => ({ ...p, amazonEsclusi: v })); }} />
        </div>
      </div>
      <div style={{ display: "flex", gap: 18, flexWrap: "wrap", marginBottom: 12, alignItems: "flex-end" }}>
        {par.metodo === "formula" && <label style={css.label} title="Di quanto può cambiare al massimo il peso di un canale rispetto a oggi">Variazione massima (±%)
          <input style={{ ...css.input, width: 80 }} defaultValue={par.tetto} onBlur={num("tetto")} />
        </label>}
        {par.metodo === "formula" ? <label style={css.label} title="Nei canali piccoli la resa conta meno e si avvicina alla media dell'editore. 0 = nessuna prudenza">Prudenza canali piccoli
          <input style={{ ...css.input, width: 80 }} defaultValue={par.prudenza} onBlur={num("prudenza")} />
        </label> : <label style={css.label} title="Sotto questo peso pochi titoli fanno impazzire la percentuale di resa: il canale resta invariato">Ignora canali con peso sotto (%)
          <input style={{ ...css.input, width: 80 }} defaultValue={par.pesoMinimo} onBlur={num("pesoMinimo")} />
        </label>}
        <div style={css.label}>Canali con resa non affidabile (restano invariati)
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            {correzione.perCanale.map(({ c }) => (
              <label key={c} style={{ display: "flex", gap: 4, alignItems: "center", color: T.text, fontSize: "11px" }}>
                <input type="checkbox" checked={par.neutri.includes(c)} onChange={() => toggleNeutro(c)} /> {canaliInfo[c]?.nome || c}
              </label>
            ))}
          </div>
        </div>
      </div>
      <table style={{ borderCollapse: "collapse", marginBottom: 12 }}>
        <thead><tr>
          <th style={css.th}>Canale</th><th style={{ ...css.th, textAlign: "right" }}>Resa media</th>
          {par.metodo === "regole" && <th style={{ ...css.th, textAlign: "right" }} title="Numero di editori in cui il canale è in fascia critica / alta / virtuosa">Critica · Alta · Virtuosa</th>}
          <th style={{ ...css.th, textAlign: "right" }}>Peso medio oggi</th><th style={{ ...css.th, textAlign: "right" }}>Peso medio corretto</th>
          <th style={{ ...css.th, textAlign: "right" }}>Differenza</th><th style={{ ...css.th, textAlign: "right" }}>Editori ↑</th><th style={{ ...css.th, textAlign: "right" }}>Editori ↓</th>
        </tr></thead>
        <tbody>
          {correzione.perCanale.map(x => {
            const d = r2(x.dopo - x.prima);
            return (
              <tr key={x.c}>
                <td style={css.td}>{canaliInfo[x.c]?.nome || x.c}{par.neutri.includes(x.c) && <span style={{ color: T.textDim, fontSize: "10px", marginLeft: 6 }}>neutro</span>}</td>
                <td style={{ ...css.td, textAlign: "right", color: x.resaMedia == null ? T.textDim : x.resaMedia >= par.critica ? T.red : x.resaMedia >= par.alta ? T.amber : x.resaMedia <= par.virtuosa ? T.green : T.text }}>{x.resaMedia == null ? "—" : `${x.resaMedia}%`}</td>
                {par.metodo === "regole" && <td style={{ ...css.td, textAlign: "right" }}>
                  <span style={{ color: T.red }}>{x.critica}</span> · <span style={{ color: T.amber }}>{x.alta}</span> · <span style={{ color: T.green }}>{x.virtuosa}</span>
                </td>}
                <td style={{ ...css.td, textAlign: "right" }}>{x.prima}%</td>
                <td style={{ ...css.td, textAlign: "right" }}>{x.dopo}%</td>
                <td style={{ ...css.td, textAlign: "right", fontWeight: 700, color: d > 0.05 ? T.green : d < -0.05 ? T.red : T.textMid }}>{d > 0 ? "+" : ""}{d} pt</td>
                <td style={{ ...css.td, textAlign: "right", color: T.green }}>{x.su || ""}</td>
                <td style={{ ...css.td, textAlign: "right", color: T.red }}>{x.giu || ""}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
      {par.metodo === "regole" && (
        <div style={{ color: T.textMid, fontSize: "11px", marginBottom: 10, lineHeight: 1.6 }}>
          Regole scattate su <b style={{ color: T.text }}>{correzione.conTagli}</b> righe su {correzione.righe.length}.
          {correzione.senzaBeneficiari.length > 0 && <span style={{ color: T.amber }}> Senza canali virtuosi: {correzione.senzaBeneficiari.join(", ")}.</span>}
          {correzione.tagliRidotti > 0 && <span> Su {correzione.tagliRidotti} righe il taglio è stato ridotto perché i canali virtuosi erano già al massimo.</span>}
        </div>
      )}
      {(resa.nonAbbinati.length > 0 || correzione.senzaPesi.length > 0) && (
        <div style={{ color: T.amber, fontSize: "11px", marginBottom: 10, lineHeight: 1.6 }}>
          {resa.nonAbbinati.length > 0 && <div>Non riconosciuti nel Ranking (ignorati): {resa.nonAbbinati.join(", ")}</div>}
          {correzione.senzaPesi.length > 0 && <div>Con resa ma senza pesi in griglia (ignorati): {correzione.senzaPesi.map(k => k.replace("|", " · ")).join(", ")}</div>}
        </div>
      )}
      <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
        <button style={css.btn("accent", !correzione.righe.length)} disabled={!correzione.righe.length} onClick={onApplica}>Applica a {correzione.righe.length} righe</button>
        <button style={css.btn()} onClick={onAnnulla}>Annulla</button>
        <span style={{ color: T.textMid, fontSize: "11px" }}>Poi controlli in griglia e premi Salva</span>
      </div>
    </div>
  );
}

function Num({ v, on }) {
  return <input inputMode="decimal" style={{ ...css.input, width: 46, textAlign: "right", padding: "2px 4px" }} defaultValue={v} onBlur={on}
    onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }} />;
}
