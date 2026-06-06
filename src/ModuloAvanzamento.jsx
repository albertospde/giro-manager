import { useState, useEffect, useMemo, useCallback, useRef } from "react";

const SUPABASE_URL = "https://tdflwenlylhctxssatax.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InRkZmx3ZW5seWxoY3R4c3NhdGF4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYzMzgyNzYsImV4cCI6MjA5MTkxNDI3Nn0.l35qEL7LOvyYuI1McQlVqj4vbyTqmlevcmqWbTGYi2Q";

const T = {
  bg: "#1a2140", surface: "#212d54", border: "#2e3d6b", borderHi: "#3d4f82",
  text: "#f0f2f8", textMid: "#8b9cc8", textDim: "#4a5a8a",
  accent: "#7b9fe8", green: "#4caf7d", red: "#e05c5c", blue: "#4a5da0",
  orange: "#e08a3c", purple: "#9c6fcf",
};

const css = {
  btn: (v = "default") => ({
    padding: "6px 14px",
    border: `1px solid ${v === "accent" ? T.accent : v === "danger" ? T.red : T.border}`,
    background: v === "accent" ? T.accent : v === "danger" ? T.red + "22" : "transparent",
    color: v === "accent" ? "#000" : v === "danger" ? T.red : T.text,
    cursor: "pointer", fontSize: "12px", fontFamily: "inherit", borderRadius: 3,
    fontWeight: v === "accent" ? "700" : "400", letterSpacing: "0.04em",
  }),
  input: {
    background: T.bg, border: `1px solid ${T.border}`, color: T.text,
    padding: "5px 10px", fontSize: "12px", fontFamily: "inherit", borderRadius: 3, outline: "none",
  },
  th: {
    padding: "8px 12px", textAlign: "left", color: T.textMid, fontWeight: "400",
    fontSize: "11px", letterSpacing: "0.08em", textTransform: "uppercase",
    borderBottom: `1px solid ${T.border}`, whiteSpace: "nowrap",
    position: "sticky", top: 0, background: T.surface, zIndex: 1,
  },
  td: { padding: "7px 12px", borderBottom: `1px solid ${T.border}22`, verticalAlign: "middle", fontSize: "12px" },
};

// ─── Helpers ──────────────────────────────────────────────────────────────────

const sbFetch = async (path, token) => {
  const r = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    headers: {
      "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`,
      "Accept": "application/json", "Range-Unit": "items", "Range": "0-499999",
    },
  });
  return r.json();
};

// Formatta data come YYYY-MM-DD senza problemi di timezone
const toDateStr = (d) => {
  if (!d) return null;
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
};

// Numero da stringa italiana (es. "2.178.600 €" → 2178600)
const parseNum = (v) => {
  if (v == null || v === "") return 0;
  return Number(String(v).replace(/[^0-9,.-]/g, "").replace(/\./g, "").replace(",", ".")) || 0;
};

const MESI = ["Gen","Feb","Mar","Apr","Mag","Giu","Lug","Ago","Set","Ott","Nov","Dic"];

const fmt = (n) => n?.toLocaleString("it-IT", { maximumFractionDigits: 0 }) ?? "0";
const fmtEur = (n) => "€ " + fmt(n);

// ─── KPI Card ─────────────────────────────────────────────────────────────────

function KpiCard({ label, value, sub, color = T.accent, onClick }) {
  return (
    <div
      onClick={onClick}
      style={{
        background: T.surface, border: `1px solid ${T.border}`, borderRadius: 4,
        padding: "14px 18px", minWidth: 150, cursor: onClick ? "pointer" : "default",
      }}
    >
      <div style={{ color: T.textMid, fontSize: "10px", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 6 }}>{label}</div>
      <div style={{ color, fontSize: "22px", fontWeight: "700", lineHeight: 1 }}>{value}</div>
      {sub && <div style={{ color: T.textMid, fontSize: "11px", marginTop: 5 }}>{sub}</div>}
    </div>
  );
}

// ─── Parser CSV Tableau (UTF-16 LE, TSV) ──────────────────────────────────────

function parseTsvTableau(buffer) {
  // Prova UTF-16 LE (BOM), fallback UTF-8
  let text;
  const view = new Uint8Array(buffer);
  if (view[0] === 0xFF && view[1] === 0xFE) {
    text = new TextDecoder("utf-16le").decode(buffer);
  } else {
    text = new TextDecoder("utf-8").decode(buffer);
  }
  const lines = text.split(/\r?\n/).filter(l => l.trim() !== "");
  if (lines.length < 2) return [];

  const headers = lines[0].split("\t").map(h => h.trim().toLowerCase());

  // Rileva colonne per nome header
  const idx = (names) => {
    for (const n of names) {
      const i = headers.findIndex(h => h.includes(n.toLowerCase()));
      if (i >= 0) return i;
    }
    return -1;
  };

  const iEan          = idx(["ean"]);
  const iTitolo       = idx(["titolo"]);
  const iAutore       = idx(["autore"]);
  const iEditore      = idx(["editore"]);
  const iPrezzo       = idx(["prezzo"]);
  const iNumLancio    = idx(["num", "lancio", "n. lancio", "n lancio"]);
  const iCopie        = idx(["copie lanciate", "copie_lanciate", "copie"]);
  const iValore       = idx(["valore lancio", "valore_lancio", "valore"]);
  const iData         = idx(["data messa in vendita", "data_messa_in_vendita", "data vendita", "data"]);

  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    const cols = lines[i].split("\t");
    const ean = iEan >= 0 ? String(cols[iEan] ?? "").replace(/\D/g, "").slice(-13) : "";
    if (!ean || ean.length !== 13) continue;

    // Parsing data: gestisce formati DD/MM/YYYY, YYYY-MM-DD, D/M/YYYY
    let dataMessaInVendita = null;
    if (iData >= 0 && cols[iData]) {
      const raw = String(cols[iData]).trim();
      // Prova DD/MM/YYYY o D/M/YYYY
      const dmY = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (dmY) {
        dataMessaInVendita = toDateStr(new Date(
          parseInt(dmY[3]), parseInt(dmY[2]) - 1, parseInt(dmY[1])
        ));
      } else {
        // Prova YYYY-MM-DD
        const isoMatch = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (isoMatch) {
          dataMessaInVendita = toDateStr(new Date(
            parseInt(isoMatch[1]), parseInt(isoMatch[2]) - 1, parseInt(isoMatch[3])
          ));
        }
      }
    }

    rows.push({
      ean,
      titolo:               iTitolo  >= 0 ? String(cols[iTitolo]  ?? "").trim() : "",
      autore:               iAutore  >= 0 ? String(cols[iAutore]  ?? "").trim() : "",
      editore_nome:         iEditore >= 0 ? String(cols[iEditore] ?? "").trim() : "",
      prezzo:               iPrezzo  >= 0 ? parseNum(cols[iPrezzo])  : 0,
      num_lancio:           iNumLancio >= 0 ? parseInt(cols[iNumLancio]) || null : null,
      copie_lanciate:       iCopie   >= 0 ? parseNum(cols[iCopie])   : 0,
      valore_lancio:        iValore  >= 0 ? parseNum(cols[iValore])  : 0,
      data_messa_in_vendita: dataMessaInVendita,
    });
  }
  return rows;
}

// ─── Modulo principale ────────────────────────────────────────────────────────

export default function ModuloAvanzamento({ token, titoli: titoliGiri, prenotato, ruolo }) {
  // ── Stato DB ────────────────────────────────────────────────────────────────
  const [novitaDB, setNovitaDB]     = useState([]);   // titoli_novita
  const [fatturato, setFatturato]   = useState([]);   // fatturato_lanci_mensile
  const [loading, setLoading]       = useState(true);

  // ── Filtri ──────────────────────────────────────────────────────────────────
  const [filterAnno, setFilterAnno] = useState(new Date().getFullYear());
  const [filterEditore, setFilterEditore] = useState("tutti");
  const [filterLancio, setFilterLancio]   = useState("tutti");
  const [filterCedola, setFilterCedola]   = useState("tutti");
  const [filterSearch, setFilterSearch]   = useState("");

  // ── UI ──────────────────────────────────────────────────────────────────────
  const [showProiezione, setShowProiezione] = useState(false);
  const [uploadMsg, setUploadMsg]           = useState(null);
  const [uploading, setUploading]           = useState(false);
  const fileRef = useRef();

  // ── Load da Supabase ─────────────────────────────────────────────────────────
  const loadNovita = useCallback(async () => {
    if (!token) return;
    const data = await sbFetch("titoli_novita?select=*", token);
    if (Array.isArray(data)) setNovitaDB(data);
  }, [token]);

  const loadFatturato = useCallback(async () => {
    if (!token) return;
    const annoRif  = filterAnno;
    const annoPrev = annoRif - 1;
    const data = await sbFetch(
      `fatturato_lanci_mensile?anno=in.(${annoRif},${annoPrev})&order=anno.asc,mese.asc`,
      token
    );
    if (Array.isArray(data)) setFatturato(data);
  }, [token, filterAnno]);

  useEffect(() => {
    setLoading(true);
    Promise.all([loadNovita(), loadFatturato()]).finally(() => setLoading(false));
  }, [loadNovita, loadFatturato]);

  // ── Join titoli giri + novitaDB ──────────────────────────────────────────────
  // novitaDB è indicizzata per EAN
  const novitaByEan = useMemo(() => {
    const map = {};
    novitaDB.forEach(n => { map[String(n.ean)] = n; });
    return map;
  }, [novitaDB]);

  // Anno di un titolo: estratto da giro_label (es. "5 2026" → 2026)
  const getAnno = (t) => {
    if (!t.giro_label) return null;
    const parts = t.giro_label.split(" ");
    return parts.length >= 2 ? parseInt(parts[parts.length - 1]) : null;
  };

  // Titoli giri dell'anno selezionato, arricchiti con dati titoli_novita
  const novitaArricchite = useMemo(() => {
    return titoliGiri
      .filter(t => getAnno(t) === filterAnno)
      .map(t => {
        const n = novitaByEan[String(t.ean)] ?? {};
        const prenotatoTot = prenotato
          .filter(p => p.titolo_id === t.id)
          .reduce((s, p) => s + p.quantita, 0);
        return {
          ...t,
          // Campi da titoli_novita — assicura che siano numeri
          num_lancio:            n.num_lancio ?? null,
          copie_lanciate:        Number(n.copie_lanciate ?? 0),
          valore_lancio:         Number(n.valore_lancio ?? 0),
          data_messa_in_vendita: n.data_messa_in_vendita ?? null,
          manuale:               n.manuale ?? false,
          stato_vendita:         n.stato_vendita ?? null,
          risposta_editore:      n.risposta_editore ?? null,
          // Calcolato
          prenotato_tot:         prenotatoTot,
        };
      });
  }, [titoliGiri, novitaByEan, prenotato, filterAnno]);

  // ── Filtri applicati ────────────────────────────────────────────────────────
  const novitaFiltrate = useMemo(() => {
    return novitaArricchite
      .filter(n => filterEditore === "tutti" || n.editore_nome === filterEditore)
      .filter(n => filterLancio  === "tutti" || String(n.num_lancio) === filterLancio)
      .filter(n => filterCedola  === "tutti" || n.n_cedola === filterCedola)
      .filter(n => {
        if (!filterSearch) return true;
        const q = filterSearch.toLowerCase();
        return (
          n.titolo?.toLowerCase().includes(q) ||
          n.ean?.includes(q) ||
          n.autore?.toLowerCase().includes(q) ||
          n.editore_nome?.toLowerCase().includes(q)
        );
      });
  }, [novitaArricchite, filterEditore, filterLancio, filterCedola, filterSearch]);

  // ── Opzioni filtri ──────────────────────────────────────────────────────────
  const anni    = useMemo(() => [...new Set(titoliGiri.map(getAnno).filter(Boolean))].sort((a,b)=>b-a), [titoliGiri]);
  const editori = useMemo(() => [...new Set(novitaArricchite.map(n=>n.editore_nome).filter(Boolean))].sort(), [novitaArricchite]);
  const lanci   = useMemo(() => [...new Set(novitaArricchite.map(n=>n.num_lancio).filter(Boolean))].sort((a,b)=>a-b), [novitaArricchite]);
  const cedole  = useMemo(() => [...new Set(novitaArricchite.map(n=>n.n_cedola).filter(Boolean))].sort(), [novitaArricchite]);

  // ── KPI ─────────────────────────────────────────────────────────────────────
  const kpi = useMemo(() => {
    const totTitoli      = novitaFiltrate.length;
    const lanciati       = novitaFiltrate.filter(n => n.copie_lanciate > 0 && !n.manuale);
    const sblRifo        = novitaFiltrate.filter(n => n.copie_lanciate > 0 &&  n.manuale);
    const nonLanciati    = novitaFiltrate.filter(n => n.prenotato_tot > 0 && n.copie_lanciate === 0);

    const valLancio      = lanciati.reduce((s,n) => s + n.valore_lancio, 0);
    const valSbl         = sblRifo.reduce((s,n) => s + n.valore_lancio, 0);
    const valPrenotato   = novitaFiltrate.reduce((s,n) => s + (n.prezzo||0) * n.prenotato_tot, 0);
    const valNonLanciato = nonLanciati.reduce((s,n) => s + (n.prezzo||0) * n.prenotato_tot, 0);
    const copieNonLanc   = nonLanciati.reduce((s,n) => s + n.prenotato_tot, 0);

    const totConCopie    = lanciati.length + sblRifo.length;
    const pctAvanz       = totTitoli > 0 ? Math.round(totConCopie / totTitoli * 100) : 0;

    return {
      totTitoli, lanciati: lanciati.length, sblRifo: sblRifo.length,
      nonLanciati: nonLanciati.length, valLancio, valSbl, valPrenotato,
      valNonLanciato, copieNonLanc, totConCopie, pctAvanz,
    };
  }, [novitaFiltrate]);

  // ── Aggregazione mesi anno corrente ────────────────────────────────────────
  // BUG FIX: usa data_messa_in_vendita come stringa YYYY-MM-DD
  // evita completamente new Date() per estrarre anno/mese e prevenire
  // problemi di timezone (UTC vs CET/CEST)
  const meseOggi = new Date().getMonth() + 1; // 1-12

  const annoCorrentePerMese = useMemo(() => {
    const perMese = {};
    // Usa TUTTI i titoli dell'anno (novitaArricchite, non filtrati)
    // perché la dashboard confronto non deve dipendere dai filtri attivi
    novitaArricchite.forEach(n => {
      if (!n.data_messa_in_vendita) return;
      if (!n.valore_lancio || n.valore_lancio === 0) return;

      // Estrae anno e mese direttamente dalla stringa YYYY-MM-DD
      // senza usare new Date(), eliminando il problema del timezone
      const match = String(n.data_messa_in_vendita).match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (!match) return;

      const annoData = parseInt(match[1]);
      const meseData = parseInt(match[2]); // 1-12

      // Deve essere nell'anno di riferimento e non futuro
      if (annoData !== filterAnno) return;
      if (meseData > meseOggi) return;

      perMese[meseData] = (perMese[meseData] || 0) + n.valore_lancio;
    });
    return perMese;
  }, [novitaArricchite, filterAnno, meseOggi]);

  // ── Fatturato anno precedente da tabella ────────────────────────────────────
  const annoPrecPerMese = useMemo(() => {
    const perMese = {};
    fatturato
      .filter(f => Number(f.anno) === filterAnno - 1)
      .forEach(f => { perMese[Number(f.mese)] = Number(f.fatturato); });
    return perMese;
  }, [fatturato, filterAnno]);

  // ── Salva fatturato mensile anno precedente ─────────────────────────────────
  const saveFatturato = async (anno, mese, valore) => {
    await fetch(`${SUPABASE_URL}/rest/v1/fatturato_lanci_mensile`, {
      method: "POST",
      headers: {
        "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`,
        "Content-Type": "application/json", "Prefer": "resolution=merge-duplicates",
      },
      body: JSON.stringify([{ anno, mese, fatturato: parseFloat(valore) || 0 }]),
    });
    setFatturato(prev => {
      const idx = prev.findIndex(f => Number(f.anno) === anno && Number(f.mese) === mese);
      const updated = { anno, mese, fatturato: parseFloat(valore) || 0 };
      return idx >= 0
        ? prev.map((f,i) => i === idx ? updated : f)
        : [...prev, updated];
    });
  };

  // ── Upload CSV Tableau ──────────────────────────────────────────────────────
  const handleUploadCSV = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setUploading(true);
    setUploadMsg(null);

    try {
      const buffer = await file.arrayBuffer();
      const parsed = parseTsvTableau(buffer);

      if (parsed.length === 0) {
        setUploadMsg({ type: "error", text: "Nessun EAN trovato nel file. Verifica il formato." });
        setUploading(false);
        return;
      }

      // EAN presenti nei giri (per filtrare solo titoli conosciuti)
      const eanGiri = new Set(titoliGiri.map(t => String(t.ean)));

      // Aggiorna solo i record già presenti nei giri e non manuali
      const toUpsert = parsed
        .filter(r => eanGiri.has(r.ean))
        .map(r => ({
          ean: r.ean,
          num_lancio:            r.num_lancio,
          copie_lanciate:        r.copie_lanciate,
          valore_lancio:         r.valore_lancio,
          data_messa_in_vendita: r.data_messa_in_vendita,
          manuale:               false,
        }));

      if (toUpsert.length === 0) {
        setUploadMsg({ type: "warn", text: "Nessun EAN del file corrisponde ai titoli nei giri." });
        setUploading(false);
        return;
      }

      // Upsert a batch su titoli_novita (on conflict: ean)
      const BATCH = 200;
      let ok = 0;
      for (let i = 0; i < toUpsert.length; i += BATCH) {
        const batch = toUpsert.slice(i, i + BATCH);
        const res = await fetch(`${SUPABASE_URL}/rest/v1/titoli_novita`, {
          method: "POST",
          headers: {
            "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`,
            "Content-Type": "application/json",
            "Prefer": "resolution=merge-duplicates",
          },
          body: JSON.stringify(batch),
        });
        if (res.ok || res.status === 201) ok += batch.length;
      }

      await loadNovita();
      setUploadMsg({ type: "ok", text: `Aggiornati ${ok} titoli su ${toUpsert.length} (trovati nel file: ${parsed.length})` });
    } catch (err) {
      setUploadMsg({ type: "error", text: "Errore: " + err.message });
    }

    setUploading(false);
    if (fileRef.current) fileRef.current.value = "";
  };

  // ── Aggiornamento inline titoli_novita ──────────────────────────────────────
  const updateNovita = async (ean, fields) => {
    await fetch(`${SUPABASE_URL}/rest/v1/titoli_novita?ean=eq.${ean}`, {
      method: "PATCH",
      headers: {
        "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`,
        "Content-Type": "application/json", "Prefer": "return=minimal",
      },
      body: JSON.stringify(fields),
    });
    setNovitaDB(prev =>
      prev.map(n => String(n.ean) === String(ean) ? { ...n, ...fields } : n)
    );
  };

  // ── Render ──────────────────────────────────────────────────────────────────
  if (loading) {
    return (
      <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", color: T.textMid }}>
        Caricamento...
      </div>
    );
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%", overflow: "hidden" }}>

      {/* ── Toolbar ── */}
      <div style={{
        padding: "10px 20px", borderBottom: `1px solid ${T.border}`,
        display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap",
        background: T.surface,
      }}>
        {/* Anno */}
        <select style={css.input} value={filterAnno} onChange={e => setFilterAnno(Number(e.target.value))}>
          {anni.map(a => <option key={a} value={a}>{a}</option>)}
        </select>

        {/* Editore */}
        <select style={{ ...css.input, maxWidth: 180 }} value={filterEditore} onChange={e => setFilterEditore(e.target.value)}>
          <option value="tutti">Tutti gli editori</option>
          {editori.map(ed => <option key={ed} value={ed}>{ed}</option>)}
        </select>

        {/* Lancio */}
        <select style={css.input} value={filterLancio} onChange={e => setFilterLancio(e.target.value)}>
          <option value="tutti">Tutti i lanci</option>
          {lanci.map(l => <option key={l} value={String(l)}>Lancio {l}</option>)}
        </select>

        {/* Cedola */}
        <select style={css.input} value={filterCedola} onChange={e => setFilterCedola(e.target.value)}>
          <option value="tutti">Tutte le cedole</option>
          {cedole.map(c => <option key={c} value={c}>{c}</option>)}
        </select>

        {/* Ricerca */}
        <input
          style={{ ...css.input, width: 180 }}
          placeholder="Cerca titolo / EAN..."
          value={filterSearch}
          onChange={e => setFilterSearch(e.target.value)}
        />

        <div style={{ color: T.textMid, fontSize: "11px" }}>
          <span style={{ color: T.text, fontWeight: "700" }}>{novitaFiltrate.length}</span> titoli
        </div>

        {/* Upload CSV */}
        <div style={{ marginLeft: "auto", display: "flex", gap: 8, alignItems: "center" }}>
          {uploading && <span style={{ color: T.textMid, fontSize: "11px" }}>Caricamento...</span>}
          {uploadMsg && (
            <span style={{
              fontSize: "11px",
              color: uploadMsg.type === "ok" ? T.green : uploadMsg.type === "warn" ? T.orange : T.red,
            }}>
              {uploadMsg.text}
            </span>
          )}
          <input ref={fileRef} type="file" accept=".txt,.tsv,.csv" id="csv-novita" style={{ display: "none" }} onChange={handleUploadCSV} />
          <label htmlFor="csv-novita" style={{ ...css.btn("accent"), cursor: "pointer" }}>
            ↑ Aggiorna da Tableau
          </label>
        </div>
      </div>

      {/* ── Contenuto scrollabile ── */}
      <div style={{ flex: 1, overflowY: "auto", padding: "16px 20px" }}>

        {/* ── KPI ── */}
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 20 }}>
          <KpiCard label="Titoli" value={kpi.totTitoli} color={T.text} />
          <KpiCard label="Lanciati" value={kpi.lanciati} sub={fmtEur(kpi.valLancio)} color={T.green} />
          <KpiCard label="SBL / Rifo" value={kpi.sblRifo} sub={fmtEur(kpi.valSbl)} color={T.orange} />
          <KpiCard label="Pren. non lanciato" value={kpi.nonLanciati} sub={`${kpi.copieNonLanc.toLocaleString("it")} copie · ${fmtEur(kpi.valNonLanciato)}`} color={T.red} />

          {/* Barra avanzamento */}
          <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 4, padding: "14px 18px", minWidth: 200 }}>
            <div style={{ color: T.textMid, fontSize: "10px", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 6 }}>Avanzamento lancio</div>
            <div style={{ color: T.accent, fontSize: "22px", fontWeight: "700", marginBottom: 8 }}>{kpi.pctAvanz}%</div>
            <div style={{ height: 6, background: T.borderHi, borderRadius: 3, overflow: "hidden" }}>
              <div style={{
                height: "100%",
                width: `${Math.min(100, (kpi.lanciati / Math.max(kpi.totTitoli,1)) * 100)}%`,
                background: T.green, display: "inline-block",
              }} />
              <div style={{
                height: "100%",
                width: `${Math.min(100, (kpi.sblRifo / Math.max(kpi.totTitoli,1)) * 100)}%`,
                background: T.orange, display: "inline-block",
              }} />
            </div>
            <div style={{ color: T.textMid, fontSize: "10px", marginTop: 5 }}>
              <span style={{ color: T.green }}>● {kpi.lanciati} lanciati</span>
              {" · "}
              <span style={{ color: T.orange }}>● {kpi.sblRifo} sbl/rifo</span>
            </div>
          </div>

          {/* Bottone proiezione */}
          <KpiCard
            label="Confronto & Proiezione"
            value={showProiezione ? "▲ Chiudi" : "▼ Apri"}
            color={T.accent}
            sub={`${filterAnno-1} vs ${filterAnno}`}
            onClick={() => setShowProiezione(v => !v)}
          />
        </div>

        {/* ── Dashboard proiezione ── */}
        {showProiezione && (
          <DashboardProiezione
            filterAnno={filterAnno}
            meseOggi={meseOggi}
            annoCorrentePerMese={annoCorrentePerMese}
            annoPrecPerMese={annoPrecPerMese}
            onSaveFatturato={saveFatturato}
            kpi={kpi}
          />
        )}

        {/* ── Tabella titoli ── */}
        <TabellaAvanzamento
          novita={novitaFiltrate}
          filterAnno={filterAnno}
          token={token}
          ruolo={ruolo}
          onUpdate={updateNovita}
        />
      </div>
    </div>
  );
}

// ─── Dashboard Proiezione ─────────────────────────────────────────────────────

function DashboardProiezione({ filterAnno, meseOggi, annoCorrentePerMese, annoPrecPerMese, onSaveFatturato, kpi }) {
  const [editingCell, setEditingCell] = useState(null);
  const [editVal, setEditVal]         = useState("");

  const totaleCorrAnnuale  = Object.values(annoCorrentePerMese).reduce((s,v)=>s+v,0);
  const totalePrecAnnuale  = Object.values(annoPrecPerMese).reduce((s,v)=>s+v,0);
  const mediaCorrente      = meseOggi > 0 ? totaleCorrAnnuale / meseOggi : 0;
  const mesiRimanenti      = 12 - meseOggi;
  const proiezione         = totaleCorrAnnuale + mediaCorrente * mesiRimanenti;

  const handleSave = async () => {
    if (editingCell == null) return;
    await onSaveFatturato(filterAnno - 1, editingCell, editVal);
    setEditingCell(null);
  };

  return (
    <div style={{
      background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6,
      padding: 20, marginBottom: 20,
    }}>
      <div style={{ color: T.accent, fontWeight: "700", fontSize: "12px", letterSpacing: "0.08em", marginBottom: 16 }}>
        CONFRONTO {filterAnno - 1} vs {filterAnno} — PROIEZIONE
      </div>

      {/* Riepilogo */}
      <div style={{ display: "flex", gap: 16, marginBottom: 16, flexWrap: "wrap" }}>
        {[
          { label: `YTD ${filterAnno} (gen–${MESI[meseOggi-1]})`, val: totaleCorrAnnuale, color: T.green },
          { label: `Totale ${filterAnno-1}`, val: totalePrecAnnuale, color: T.textMid },
          { label: "Media mensile corrente", val: mediaCorrente, color: T.accent },
          { label: `Proiezione fine ${filterAnno}`, val: proiezione, color: T.purple },
        ].map(({ label, val, color }) => (
          <div key={label} style={{ background: T.bg, border: `1px solid ${T.border}`, borderRadius: 4, padding: "10px 14px", minWidth: 160 }}>
            <div style={{ color: T.textMid, fontSize: "10px", textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 4 }}>{label}</div>
            <div style={{ color, fontWeight: "700", fontSize: "16px" }}>{fmtEur(val)}</div>
          </div>
        ))}
      </div>

      {/* Griglia mensile */}
      <div style={{ overflowX: "auto" }}>
        <table style={{ borderCollapse: "collapse", fontSize: "11px" }}>
          <thead>
            <tr>
              <th style={{ ...css.th, width: 80 }}>Mese</th>
              {MESI.map((m, i) => (
                <th key={i} style={{
                  ...css.th, width: 80, textAlign: "right",
                  color: i + 1 === meseOggi ? T.accent : T.textMid,
                }}>
                  {m}
                </th>
              ))}
              <th style={{ ...css.th, width: 90, textAlign: "right" }}>Totale</th>
            </tr>
          </thead>
          <tbody>
            {/* Riga anno precedente — editabile */}
            <tr>
              <td style={{ ...css.td, color: T.textMid, fontWeight: "600", fontSize: "11px" }}>
                {filterAnno - 1}
                <div style={{ color: T.textDim, fontSize: "9px" }}>click per edit</div>
              </td>
              {MESI.map((_, i) => {
                const mese = i + 1;
                const val  = annoPrecPerMese[mese] || 0;
                const isEd = editingCell === mese;
                return (
                  <td key={i} style={{ ...css.td, textAlign: "right", padding: "4px 6px" }}>
                    {isEd ? (
                      <input
                        autoFocus
                        type="number"
                        defaultValue={val}
                        style={{ ...css.input, width: 72, padding: "2px 4px", textAlign: "right", fontSize: "11px" }}
                        onBlur={e => { setEditVal(e.target.value); handleSave(); }}
                        onChange={e => setEditVal(e.target.value)}
                        onKeyDown={e => {
                          if (e.key === "Enter") { setEditVal(e.target.value); handleSave(); }
                          if (e.key === "Escape") setEditingCell(null);
                        }}
                      />
                    ) : (
                      <span
                        style={{ color: val > 0 ? T.text : T.textDim, cursor: "pointer" }}
                        onClick={() => { setEditingCell(mese); setEditVal(String(val)); }}
                      >
                        {val > 0 ? fmtEur(val) : "—"}
                      </span>
                    )}
                  </td>
                );
              })}
              <td style={{ ...css.td, textAlign: "right", color: T.text, fontWeight: "700" }}>
                {fmtEur(totalePrecAnnuale)}
              </td>
            </tr>

            {/* Riga anno corrente — calcolata */}
            <tr style={{ background: T.accent + "0a" }}>
              <td style={{ ...css.td, color: T.accent, fontWeight: "700", fontSize: "11px" }}>
                {filterAnno}
              </td>
              {MESI.map((_, i) => {
                const mese = i + 1;
                const val  = annoCorrentePerMese[mese] || 0;
                const past = mese <= meseOggi;
                return (
                  <td key={i} style={{
                    ...css.td, textAlign: "right",
                    background: mese === meseOggi ? T.accent + "15" : "transparent",
                  }}>
                    <span style={{ color: past ? (val > 0 ? T.green : T.red) : T.textDim, fontWeight: past && val > 0 ? "700" : "400" }}>
                      {val > 0 ? fmtEur(val) : past ? "€ 0" : "—"}
                    </span>
                  </td>
                );
              })}
              <td style={{ ...css.td, textAlign: "right", color: T.green, fontWeight: "700" }}>
                {fmtEur(totaleCorrAnnuale)}
              </td>
            </tr>

            {/* Delta */}
            <tr>
              <td style={{ ...css.td, color: T.textMid, fontSize: "10px" }}>Δ</td>
              {MESI.map((_, i) => {
                const mese = i + 1;
                if (mese > meseOggi) return <td key={i} style={css.td} />;
                const curr = annoCorrentePerMese[mese] || 0;
                const prev = annoPrecPerMese[mese] || 0;
                if (!prev) return <td key={i} style={{ ...css.td, textAlign: "right", color: T.textDim }}>n/d</td>;
                const pct  = Math.round((curr - prev) / prev * 100);
                return (
                  <td key={i} style={{ ...css.td, textAlign: "right" }}>
                    <span style={{ color: pct >= 0 ? T.green : T.red, fontSize: "10px", fontWeight: "700" }}>
                      {pct >= 0 ? "+" : ""}{pct}%
                    </span>
                  </td>
                );
              })}
              <td style={css.td} />
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ─── Tabella Avanzamento ──────────────────────────────────────────────────────

function TabellaAvanzamento({ novita, filterAnno, token, ruolo, onUpdate }) {
  const [sortKey, setSortKey]     = useState("n_cedola");
  const [sortDir, setSortDir]     = useState("asc");
  const [editRow, setEditRow]     = useState(null);
  const [editVals, setEditVals]   = useState({});

  const toggleSort = (key) => {
    if (sortKey === key) setSortDir(d => d === "asc" ? "desc" : "asc");
    else { setSortKey(key); setSortDir("asc"); }
  };

  const sorted = useMemo(() => {
    return [...novita].sort((a, b) => {
      let va = a[sortKey] ?? "";
      let vb = b[sortKey] ?? "";
      if (typeof va === "string") va = va.toLowerCase();
      if (typeof vb === "string") vb = vb.toLowerCase();
      if (va < vb) return sortDir === "asc" ? -1 : 1;
      if (va > vb) return sortDir === "asc" ? 1 : -1;
      return 0;
    });
  }, [novita, sortKey, sortDir]);

  const Th = ({ k, label, right }) => (
    <th
      style={{ ...css.th, cursor: "pointer", textAlign: right ? "right" : "left",
        color: sortKey === k ? T.accent : T.textMid }}
      onClick={() => toggleSort(k)}
    >
      {label}{sortKey === k ? (sortDir === "asc" ? " ↑" : " ↓") : ""}
    </th>
  );

  const startEdit = (n) => {
    setEditRow(n.ean);
    setEditVals({
      copie_lanciate: n.copie_lanciate,
      valore_lancio:  n.valore_lancio,
      data_messa_in_vendita: n.data_messa_in_vendita || "",
      stato_vendita:  n.stato_vendita || "",
      risposta_editore: n.risposta_editore || "",
    });
  };

  const saveEdit = async (ean) => {
    await onUpdate(ean, {
      ...editVals,
      copie_lanciate: Number(editVals.copie_lanciate) || 0,
      valore_lancio:  Number(editVals.valore_lancio)  || 0,
      data_messa_in_vendita: editVals.data_messa_in_vendita || null,
      manuale: true,
    });
    setEditRow(null);
  };

  const SV_COLORS = { A: T.green, B: T.accent, C: T.orange, D: T.red };

  return (
    <div style={{ border: `1px solid ${T.border}`, borderRadius: 4, overflow: "hidden" }}>
      <div style={{ overflowX: "auto", maxHeight: "calc(100vh - 420px)", overflowY: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr>
              <Th k="n_cedola"   label="Cedola" />
              <Th k="ean"        label="EAN" />
              <Th k="titolo"     label="Titolo" />
              <Th k="editore_nome" label="Editore" />
              <Th k="num_lancio" label="Lancio" />
              <Th k="data_messa_in_vendita" label="Data vendita" />
              <Th k="prenotato_tot" label="Prenotato" right />
              <Th k="copie_lanciate" label="Copie" right />
              <Th k="valore_lancio" label="Valore" right />
              <th style={css.th}>SV</th>
              <th style={css.th}>RE</th>
              {ruolo !== "agente" && <th style={css.th}></th>}
            </tr>
          </thead>
          <tbody>
            {sorted.map((n, i) => {
              const isEdit = editRow === n.ean;
              const hasCopie = n.copie_lanciate > 0;
              return (
                <tr key={n.ean} style={{ background: i % 2 === 0 ? "transparent" : T.surface + "66" }}>
                  <td style={{ ...css.td, color: T.textMid, fontSize: "10px" }}>{n.n_cedola}</td>
                  <td style={{ ...css.td, fontFamily: "monospace", fontSize: "11px", color: T.textDim }}>{n.ean}</td>
                  <td style={{ ...css.td, maxWidth: 200 }}>
                    <div style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontWeight: "600" }}>
                      {n.titolo}
                    </div>
                    <div style={{ color: T.textMid, fontSize: "10px" }}>{n.autore}</div>
                  </td>
                  <td style={{ ...css.td, color: T.accent, fontWeight: "600", whiteSpace: "nowrap" }}>{n.editore_nome}</td>
                  <td style={{ ...css.td, color: T.textMid, textAlign: "center" }}>{n.num_lancio ?? "—"}</td>

                  {/* Data vendita */}
                  <td style={{ ...css.td, whiteSpace: "nowrap" }}>
                    {isEdit ? (
                      <input
                        type="date"
                        value={editVals.data_messa_in_vendita}
                        style={{ ...css.input, padding: "2px 6px", fontSize: "11px" }}
                        onChange={e => setEditVals(v => ({ ...v, data_messa_in_vendita: e.target.value }))}
                      />
                    ) : (
                      <span style={{ color: n.data_messa_in_vendita ? T.text : T.textDim }}>
                        {n.data_messa_in_vendita
                          ? (() => {
                              const [y,m,d] = n.data_messa_in_vendita.split("-");
                              return `${d}/${m}/${y}`;
                            })()
                          : "—"}
                      </span>
                    )}
                  </td>

                  {/* Prenotato */}
                  <td style={{ ...css.td, textAlign: "right", color: n.prenotato_tot > 0 ? T.text : T.textDim, fontWeight: "600" }}>
                    {n.prenotato_tot > 0 ? n.prenotato_tot.toLocaleString("it") : "—"}
                  </td>

                  {/* Copie lanciate */}
                  <td style={{ ...css.td, textAlign: "right" }}>
                    {isEdit ? (
                      <input
                        type="number"
                        value={editVals.copie_lanciate}
                        style={{ ...css.input, width: 80, padding: "2px 6px", textAlign: "right", fontSize: "11px" }}
                        onChange={e => setEditVals(v => ({ ...v, copie_lanciate: e.target.value }))}
                      />
                    ) : (
                      <span style={{ color: hasCopie ? (n.manuale ? T.orange : T.green) : T.textDim }}>
                        {hasCopie ? n.copie_lanciate.toLocaleString("it") : "—"}
                        {n.manuale && hasCopie && <span style={{ fontSize: "9px", marginLeft: 3 }}>M</span>}
                      </span>
                    )}
                  </td>

                  {/* Valore lancio */}
                  <td style={{ ...css.td, textAlign: "right" }}>
                    {isEdit ? (
                      <input
                        type="number"
                        value={editVals.valore_lancio}
                        style={{ ...css.input, width: 90, padding: "2px 6px", textAlign: "right", fontSize: "11px" }}
                        onChange={e => setEditVals(v => ({ ...v, valore_lancio: e.target.value }))}
                      />
                    ) : (
                      <span style={{ color: n.valore_lancio > 0 ? T.text : T.textDim }}>
                        {n.valore_lancio > 0 ? fmtEur(n.valore_lancio) : "—"}
                      </span>
                    )}
                  </td>

                  {/* Stato vendita */}
                  <td style={{ ...css.td, textAlign: "center" }}>
                    {isEdit ? (
                      <select
                        value={editVals.stato_vendita || ""}
                        style={{ ...css.input, padding: "2px 4px", fontSize: "11px" }}
                        onChange={e => setEditVals(v => ({ ...v, stato_vendita: e.target.value || null }))}
                      >
                        <option value="">—</option>
                        {["A","B","C","D"].map(v => <option key={v} value={v}>{v}</option>)}
                      </select>
                    ) : (
                      n.stato_vendita
                        ? <span style={{ color: SV_COLORS[n.stato_vendita] || T.textMid, fontWeight: "700" }}>{n.stato_vendita}</span>
                        : <span style={{ color: T.textDim }}>—</span>
                    )}
                  </td>

                  {/* Risposta editore */}
                  <td style={{ ...css.td, maxWidth: 120 }}>
                    {isEdit ? (
                      <input
                        value={editVals.risposta_editore}
                        style={{ ...css.input, width: 110, padding: "2px 6px", fontSize: "11px" }}
                        onChange={e => setEditVals(v => ({ ...v, risposta_editore: e.target.value }))}
                      />
                    ) : (
                      <span style={{ color: T.textMid, fontSize: "11px" }} title={n.risposta_editore}>
                        {n.risposta_editore
                          ? n.risposta_editore.length > 18 ? n.risposta_editore.slice(0,18)+"…" : n.risposta_editore
                          : ""}
                      </span>
                    )}
                  </td>

                  {/* Azioni */}
                  {ruolo !== "agente" && (
                    <td style={css.td}>
                      {isEdit ? (
                        <div style={{ display: "flex", gap: 4 }}>
                          <button style={{ ...css.btn("accent"), padding: "2px 8px", fontSize: "11px" }} onClick={() => saveEdit(n.ean)}>✓</button>
                          <button style={{ ...css.btn(), padding: "2px 8px", fontSize: "11px" }} onClick={() => setEditRow(null)}>✕</button>
                        </div>
                      ) : (
                        <button style={{ ...css.btn(), padding: "2px 8px", fontSize: "11px" }} onClick={() => startEdit(n)}>✎</button>
                      )}
                    </td>
                  )}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}
