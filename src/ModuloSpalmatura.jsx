import { useState, useEffect, useMemo, useCallback, useRef } from "react";

// ─── Pesi Spalmatura ─────────────────────────────────────────────────────────
// Vista e modifica diretta della tabella spalmatura_obiettivo: per ogni editore+formato, la quota
// (in % 0-100) dell'obiettivo assegnata a ciascun canale. Da qui GiroManager calcola gli obiettivi per
// canale/gruppo (dashboard, cedola direzionale/agenti, Fine Giro) e BookUp gli obiettivi agente.
// • Una riga per editore+formato, una colonna per canale. Cella vuota = nessun peso su quel canale.
// • Le modifiche restano in bozza finché non premi Salva (come Ranking editori).
// • "Importa Excel" carica il template SPALMATURA nella griglia come modifiche non salvate;
//   "Esporta Excel" produce lo stesso template, quindi si può scaricare, modificare e ricaricare.

const SUPABASE_URL = "https://tdflwenlylhctxssatax.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InRkZmx3ZW5seWxoY3R4c3NhdGF4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYzMzgyNzYsImV4cCI6MjA5MTkxNDI3Nn0.l35qEL7LOvyYuI1McQlVqj4vbyTqmlevcmqWbTGYi2Q";

const T = {
  bg: "#1a2140", surface: "#212d54", border: "#2e3d6b", borderHi: "#3d4f82",
  text: "#f0f2f8", textMid: "#8b9cc8", textDim: "#4a5a8a",
  accent: "#7b9fe8", green: "#4caf7d", red: "#e05c5c", amber: "#e0a84c",
};

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
  "FASTBOOK", "CENTROLIBRI", "GROSSISTI", "GDO",
];
const FORMATI = ["Cover", "Tascabile"];

const norm = (s) => String(s ?? "").replace(/\s+/g, " ").trim().toUpperCase();
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
  const fileRef = useRef(null);

  const carica = useCallback(async () => {
    setLoading(true); setErrore("");
    try {
      const [sp, can, rk, ne] = await Promise.all([
        fetchTutto("spalmatura_obiettivo?select=editore_nome,formato,canale_codice,percentuale&order=id", token),
        fetchTutto("canali?select=codice,nome,gruppo", token),
        fetchTutto("ranking_editori?select=editore_nome,ranking,cedola,attivo", token),
        fetchTutto("editori_new_entry?select=nome_editore&attivo=eq.true", token).catch(() => []),
      ]);
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

  const importa = (e) => {
    const f = e.target.files[0];
    e.target.value = "";
    if (!f) return;
    setErrore(""); setMsg("");
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const XLSX = window.XLSX;
        const wb = XLSX.read(evt.target.result, { type: "array" });
        const ws = wb.Sheets["SPALMATURA"];
        if (!ws) throw new Error("Foglio 'SPALMATURA' non trovato nel file.");
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        // Riga 4 = intestazioni (EDITORE, FORMATO, codici canale…), dati da riga 5
        const intest = (data[3] || []).map(h => norm(h).replace(/\s+/g, "_"));
        const colCanali = intest.map((h, i) => [h, i]).filter(([h, i]) => i >= 2 && h && h !== "SOMMA_%" && h !== "SOMMA");
        if (!colCanali.length) throw new Error("Intestazioni dei canali non trovate in riga 4.");
        const errs = [];
        const lette = [];
        data.slice(4).forEach((r, idx) => {
          if (!r.some(v => v !== "")) return;
          const editore = norm(r[0]);
          const formato = String(r[1] ?? "").trim();
          if (!editore || !FORMATI.includes(formato)) { errs.push(`riga ${idx + 5}: editore o formato non valido`); return; }
          const pesi = {};
          colCanali.forEach(([c, i]) => {
            const v = numOVuoto(r[i]);
            if (v === null) errs.push(`riga ${idx + 5}: peso non valido per ${c}`);
            else if (v !== "") pesi[c] = v;
          });
          lette.push({ key: chiave(editore, formato), editore_nome: editore, formato, pesi });
        });
        if (errs.length) throw new Error(`File non caricato, correggi: ${errs.slice(0, 5).join("; ")}${errs.length > 5 ? ` e altri ${errs.length - 5}` : ""}`);
        // Le righe del file sostituiscono quelle esistenti; le altre restano come sono
        setRighe(rs => {
          const m = new Map(rs.map(r => [r.key, r]));
          lette.forEach(r => m.set(r.key, r));
          return [...m.values()];
        });
        setMsg(`Caricate ${lette.length} righe dal file: controlla le modifiche evidenziate e premi Salva.`);
      } catch (err) { setErrore(err.message); }
    };
    reader.readAsArrayBuffer(f);
  };

  // ─── Vista ─────────────────────────────────────────────────────────────────
  const stat = useMemo(() => {
    const fuori = righe.filter(r => !sommaOk(somma(r.pesi))).length;
    const conPesi = new Set(righe.map(r => r.editore_nome));
    const senzaPesi = ranking.filter(r => r.attivo !== false && !conPesi.has(r.editore_nome)).sort((a, b) => a.ranking - b.ranking);
    return { fuori, editori: conPesi.size, senzaPesi };
  }, [righe, ranking]);

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
        <button style={css.btn()} onClick={() => fileRef.current?.click()} title="Carica un file nel formato del template SPALMATURA: le righe del file sostituiscono quelle in griglia (da salvare)">Importa Excel</button>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={importa} />
        <button style={css.btn()} onClick={esporta}>Esporta Excel</button>
      </div>

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
