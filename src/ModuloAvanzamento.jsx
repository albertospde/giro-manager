import { useState, useEffect, useMemo, useCallback } from "react";

const SUPABASE_URL = "https://tdflwenlylhctxssatax.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InRkZmx3ZW5seWxoY3R4c3NhdGF4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYzMzgyNzYsImV4cCI6MjA5MTkxNDI3Nn0.l35qEL7LOvyYuI1McQlVqj4vbyTqmlevcmqWbTGYi2Q";

const T = {
  bg: "#1a2140", surface: "#212d54", border: "#2e3d6b", borderHi: "#3d4f82",
  text: "#f0f2f8", textMid: "#8b9cc8", textDim: "#4a5a8a",
  accent: "#7b9fe8", green: "#4caf7d", red: "#e05c5c", blue: "#4a5da0", purple: "#9c6fcf",
};

const css = {
  app: { background: T.bg, color: T.text, minHeight: "100vh", fontFamily: "'Inter', 'Segoe UI', Arial, sans-serif", fontSize: "13px" },
  sidebar: { width: 200, background: T.surface, borderRight: `1px solid ${T.border}`, display: "flex", flexDirection: "column", flexShrink: 0 },
  main: { flex: 1, overflow: "hidden", display: "flex", flexDirection: "column" },
  header: { borderBottom: `1px solid ${T.border}`, padding: "12px 20px", display: "flex", alignItems: "center", gap: 12, background: T.surface },
  btn: (v = "default") => ({ padding: "6px 14px", border: `1px solid ${v === "accent" ? T.accent : T.border}`, background: v === "accent" ? T.accent : "transparent", color: v === "accent" ? "#000" : T.text, cursor: "pointer", fontSize: "12px", fontFamily: "inherit", borderRadius: 3, fontWeight: v === "accent" ? "700" : "400", letterSpacing: "0.04em", transition: "all 0.15s" }),
  input: { background: T.bg, border: `1px solid ${T.border}`, color: T.text, padding: "5px 10px", fontSize: "12px", fontFamily: "inherit", borderRadius: 3, outline: "none" },
  tag: (c = T.accent) => ({ display: "inline-block", padding: "2px 7px", background: c + "22", border: `1px solid ${c}44`, color: c, borderRadius: 2, fontSize: "10px", fontWeight: "700", letterSpacing: "0.06em", textTransform: "uppercase" }),
  table: { width: "100%", borderCollapse: "collapse" },
  th: { padding: "8px 12px", textAlign: "left", color: T.textMid, fontWeight: "400", fontSize: "11px", letterSpacing: "0.08em", textTransform: "uppercase", borderBottom: `1px solid ${T.border}`, whiteSpace: "nowrap", position: "sticky", top: 0, background: T.surface, zIndex: 1 },
  td: { padding: "8px 12px", borderBottom: `1px solid ${T.border}22`, verticalAlign: "middle" },
};

const MESI_IT = { gennaio:0, febbraio:1, marzo:2, aprile:3, maggio:4, giugno:5, luglio:6, agosto:7, settembre:8, ottobre:9, novembre:10, dicembre:11 };

const sbFetch = async (path, token) => {
  const r = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`, "Accept": "application/json", "Range-Unit": "items", "Range": "0-499999" },
  });
  return r.json();
}

function KpiCard({ label, value, sub, color = T.accent }) {
  return (
    <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 4, padding: "16px 20px", minWidth: 160 }}>
      <div style={{ color: T.textMid, fontSize: "10px", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 8 }}>{label}</div>
      <div style={{ color, fontSize: "28px", fontWeight: "700", lineHeight: 1 }}>{value}</div>
      {sub && <div style={{ color: T.textMid, fontSize: "11px", marginTop: 6 }}>{sub}</div>}
    </div>
  );
}

function SearchableMultiSelect({ values, onChange, options, placeholder = "Tutti", width = 180, renderOption }) {
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState("");
  const label = (o) => renderOption ? renderOption(o) : o;
  const filtered = options.filter(o => label(o).toLowerCase().includes(search.toLowerCase()));
  return (
    <div style={{ position: "relative" }}>
      <button style={{ ...css.btn(values.length > 0 ? "accent" : "default"), minWidth: width, display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }} onClick={() => setOpen(o => !o)}>
        <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: width - 30 }}>
          {values.length === 0 ? placeholder : renderOption ? values.map(renderOption).join(", ") : `${values.length} selezionati`}
        </span>
        <span>▾</span>
      </button>
      {open && (
        <div style={{ position: "absolute", top: "100%", left: 0, zIndex: 50, background: T.surface, border: `1px solid ${T.borderHi}`, borderRadius: 4, minWidth: width, maxHeight: 320, display: "flex", flexDirection: "column", marginTop: 4, boxShadow: "0 4px 20px #0008" }}>
          <div style={{ padding: "8px 10px", borderBottom: `1px solid ${T.border}`, flexShrink: 0 }}>
            <input style={{ ...css.input, width: "100%", boxSizing: "border-box", marginBottom: 6 }} placeholder="Cerca..." value={search} onChange={e => setSearch(e.target.value)} autoFocus />
            <div style={{ display: "flex", justifyContent: "space-between" }}>
              <span style={{ color: T.textMid, fontSize: "11px" }}>{filtered.length} voci</span>
              <span style={{ color: T.accent, fontSize: "11px", cursor: "pointer" }} onClick={() => onChange([])}>Deseleziona tutti</span>
            </div>
          </div>
          <div style={{ overflowY: "auto", flex: 1 }}>
            {filtered.map((o, i) => (
              <label key={i} style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 12px", cursor: "pointer", borderBottom: `1px solid ${T.border}22`, background: values.includes(o) ? T.accent + "18" : "transparent" }}>
                <input type="checkbox" checked={values.includes(o)} onChange={() => onChange(values.includes(o) ? values.filter(x => x !== o) : [...values, o])} style={{ accentColor: T.accent }} />
                <span style={{ fontSize: "12px", color: values.includes(o) ? T.accent : T.text }}>{label(o)}</span>
              </label>
            ))}
          </div>
          <div style={{ padding: 8, borderTop: `1px solid ${T.border}`, flexShrink: 0 }}>
            <button style={{ ...css.btn("accent"), width: "100%" }} onClick={() => { setOpen(false); setSearch(""); }}>Applica</button>
          </div>
        </div>
      )}
    </div>
  );
}

function parseDataIt(str) {
  if (!str) return null;
  const parts = String(str).trim().split(/\s+/);
  if (parts.length < 2) return null;
  // Formato "giorno mese anno" (es. "2 gennaio 2026") oppure "giorno mese" (senza anno)
  let giorno, mese, anno;
  if (parts.length === 3) {
    [giorno, mese, anno] = parts;
  } else if (parts.length === 2) {
    [mese, anno] = parts; giorno = "1";
  } else return null;
  const m = MESI_IT[mese?.toLowerCase()];
  if (m === undefined) return null;
  const y = parseInt(anno);
  if (isNaN(y)) return null;
  return new Date(y, m, parseInt(giorno) || 1);
}

function fmtDate(d) {
  if (!d) return "—";
  if (typeof d === "string") d = new Date(d);
  if (isNaN(d)) return "—";
  return d.toLocaleDateString("it-IT", { day: "2-digit", month: "short", year: "numeric" });
}

function parseValoreLancio(str) {
  if (typeof str === "number") return str;
  if (!str) return 0;
  // Rimuovi simbolo € (anche \xac), spazi, poi togli tutti i punti (migliaia IT), converti virgola in punto decimale
  const s = String(str).replace(/[€\xac\s]/g, "").replace(/\./g, "").replace(",", ".");
  return parseFloat(s) || 0;
}

function parsePrezzo(str) {
  if (typeof str === "number") return str;
  if (!str) return 0;
  return parseFloat(String(str).replace(",", ".")) || 0;
}

// FIX: parse copie con separatore migliaia italiano (punti) e senza
function parseCopie(str) {
  if (typeof str === "number") return Math.round(str);
  if (!str) return 0;
  // Rimuovi punti (separatori migliaia IT) e spazi, poi parseInt
  const s = String(str).replace(/\./g, "").replace(/\s/g, "").replace(",", ".");
  return parseInt(s) || 0;
}

export default function ModuloAvanzamento({ titoli, prenotato, canali, token, ruolo }) {
  const [novitaDB, setNovitaDB] = useState([]);
  const [loading, setLoading] = useState(true);
  const [uploading, setUploading] = useState(false);
  const [manualOpen, setManualOpen] = useState(false);
  const [manualForm, setManualForm] = useState({ ean: "", titolo: "", autore: "", editore: "", prezzo: "", num_lancio: "", copie_lanciate: "", valore_lancio: "", data_messa_in_vendita: "" });
  const [filterAnno, setFilterAnno] = useState(new Date().getFullYear());
  const [filterEditore, setFilterEditore] = useState("tutti");
  const [filterNumLancio, setFilterNumLancio] = useState([]);
  const [filterCedole, setFilterCedole] = useState([]);
  const [filterGiri, setFilterGiri] = useState([]);
  const [filterMesi, setFilterMesi] = useState([]);
  const [search, setSearch] = useState("");
  const [sortKey, setSortKey] = useState("cedola");
  const [sortDir, setSortDir] = useState("asc");
  const [filterDataVendita, setFilterDataVendita] = useState([]);
  const [filterCopieLanciate, setFilterCopieLanciate] = useState([]);
  const [editingEan, setEditingEan] = useState(null);
  const [editingVal, setEditingVal] = useState("");
  const [toast, setToast] = useState(null);
  const [fatturato, setFatturato] = useState([]);
  const [showProiezione, setShowProiezione] = useState(false);

  const toggleSort = (key) => {
    if (sortKey === key) setSortDir(d => d === "asc" ? "desc" : "asc");
    else { setSortKey(key); setSortDir(key === "prenotato" ? "desc" : "asc"); }
  };
  const sortIcon = (key) => sortKey === key ? (sortDir === "asc" ? " ↑" : " ↓") : "";

  const showToast = (msg, type = "ok") => { setToast({ msg, type }); setTimeout(() => setToast(null), 3000); };

  const loadNovita = useCallback(async () => {
    if (!token) return;
    setLoading(true);
    const data = await sbFetch("titoli_novita?select=*&order=data_messa_in_vendita.desc.nullslast,editore.asc,titolo.asc", token);
    if (Array.isArray(data)) setNovitaDB(data);
    setLoading(false);
  }, [token]);

  useEffect(() => { loadNovita(); }, [loadNovita]);

  const annoRif = filterAnno || new Date().getFullYear();
  const annoPrev = annoRif - 1;
  const loadFatturato = useCallback(async () => {
    if (!token) return;
    const d = await sbFetch(`fatturato_lanci_mensile?anno=in.(${annoRif},${annoPrev})&order=anno.asc,mese.asc`, token);
    if (Array.isArray(d)) setFatturato(d);
  }, [token, annoRif, annoPrev]);
  useEffect(() => { loadFatturato(); }, [loadFatturato]);

  const saveFatturato = async (anno, mese, valore) => {
    await fetch(`${SUPABASE_URL}/rest/v1/fatturato_lanci_mensile?on_conflict=anno,mese`, {
      method: "POST",
      headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`, "Content-Type": "application/json", "Prefer": "resolution=merge-duplicates" },
      body: JSON.stringify([{ anno, mese, fatturato: parseFloat(valore) || 0 }]),
    });
    setFatturato(prev => {
      const existing = prev.find(f => Number(f.anno) === Number(anno) && Number(f.mese) === Number(mese));
      if (existing) return prev.map(f => (Number(f.anno) === Number(anno) && Number(f.mese) === Number(mese)) ? { ...f, fatturato: parseFloat(valore) || 0 } : f);
      return [...prev, { anno: Number(anno), mese: Number(mese), fatturato: parseFloat(valore) || 0 }].sort((a, b) => a.anno - b.anno || a.mese - b.mese);
    });
  };

  const handleFatturatoUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true);
    try {
      const XLSX = window.XLSX;
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

      const MESI = { gen:1,gennaio:1,jan:1,january:1, feb:2,febbraio:2,february:2, mar:3,marzo:3,march:3, apr:4,aprile:4,april:4, mag:5,maggio:5,may:5, giu:6,giugno:6,jun:6,june:6, lug:7,luglio:7,jul:7,july:7, ago:8,agosto:8,aug:8,august:8, set:9,settembre:9,sep:9,september:9, ott:10,ottobre:10,oct:10,october:10, nov:11,novembre:11,november:11, dic:12,dicembre:12,dec:12,december:12 };

      const toNum = (v) => {
        if (v == null || v === "") return 0;
        if (typeof v === "number") return v;
        const s = String(v).trim().replace(/\s/g, "").replace(/[€$£\xac]/g, "");
        if (!s) return 0;
        if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) return parseFloat(s.replace(/\./g, "").replace(",", ".")) || 0;
        if (/^\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) return parseFloat(s.replace(/,/g, "")) || 0;
        if (s.includes(",") && !s.includes(".")) return parseFloat(s.replace(",", ".")) || 0;
        return parseFloat(s) || 0;
      };

      const toStr = (v) => String(v == null ? "" : v).trim().toLowerCase().replace(/[^a-zàèéìòù0-9]/g, " ").trim();

      const getMese = (v) => {
        if (v == null) return null;
        const s = toStr(v);
        for (const [key, num] of Object.entries(MESI)) { if (s === key || s.startsWith(key + " ") || s.endsWith(" " + key)) return num; }
        const n = parseInt(s);
        if (!isNaN(n) && n >= 1 && n <= 12) return n;
        return null;
      };

      const payload = [];
      const found = {};

      if (payload.length < 3) {
        for (let ri = 0; ri < Math.min(rows.length, 10); ri++) {
          const hrow = (rows[ri] || []).map(toStr);
          const colD = hrow.findIndex(h => h.includes("data") || h.includes("pubblica") || h.includes("lancio") || h.includes("vendita"));
          const colF = hrow.findIndex(h => h.includes("fatturato") || h.includes("valore") || h.includes("importo"));
          if (colD >= 0 && colF >= 0) {
            const aggMese = {};
            for (let i = ri + 1; i < rows.length; i++) {
              const r = rows[i] || [];
              const rawData = r[colD];
              const rawVal = r[colF];
              if (!rawData || rawVal == null) continue;
              let mese = null;
              if (rawData instanceof Date && !isNaN(rawData)) {
                mese = rawData.getMonth() + 1;
              } else {
                const s = String(rawData).trim();
                const m1 = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
                const m2 = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
                if (m1) mese = parseInt(m1[2]);
                else if (m2) mese = parseInt(m2[2]);
              }
              const val = toNum(rawVal);
              if (mese >= 1 && mese <= 12 && val > 0) {
                aggMese[mese] = (aggMese[mese] || 0) + val;
              }
            }
            Object.entries(aggMese).forEach(([m, tot]) => {
              payload.push({ anno: annoPrev, mese: parseInt(m), fatturato: tot });
            });
            if (payload.length >= 3) break;
          }
        }
      }

      for (let ri = 0; ri < Math.min(rows.length, 10); ri++) {
        const hrow = (rows[ri] || []).map(toStr);
        const colM = hrow.findIndex(h => h === "mese" || h === "month" || h === "mese anno");
        const colV = hrow.findIndex(h => h.includes("fatturato") || h.includes("valore") || h.includes("importo") || h.includes("totale") || h.includes("revenue") || h.includes("ricavo"));
        const colV2 = colM >= 0 && colV < 0 ? colM + 1 : colV;
        if (colM >= 0 && colV2 >= 0 && colV2 !== colM) {
          for (let i = ri + 1; i < rows.length; i++) {
            const r = rows[i] || [];
            const m = getMese(r[colM]);
            const val = toNum(r[colV2]);
            if (m && val > 0 && !found[m]) { found[m] = true; payload.push({ anno: annoPrev, mese: m, fatturato: val }); }
          }
          if (payload.length >= 3) break;
        }
      }

      if (payload.length < 3) {
        for (const row of rows) {
          if (!row) continue;
          const m = getMese(row[0]);
          if (!m || found[m]) continue;
          for (let ci = 1; ci < Math.min(row.length, 6); ci++) {
            const val = toNum(row[ci]);
            if (val > 100) { found[m] = true; payload.push({ anno: annoPrev, mese: m, fatturato: val }); break; }
          }
        }
      }

      if (payload.length < 3) {
        for (let ri = 0; ri < rows.length - 1; ri++) {
          const row = rows[ri] || [];
          const mesiTrovati = row.map((c, ci) => ({ ci, m: getMese(c) })).filter(x => x.m);
          if (mesiTrovati.length >= 3) {
            const valRow = rows[ri + 1] || [];
            mesiTrovati.forEach(({ ci, m }) => {
              if (found[m]) return;
              const val = toNum(valRow[ci]);
              if (val > 100) { found[m] = true; payload.push({ anno: annoPrev, mese: m, fatturato: val }); }
            });
            if (payload.length >= 3) break;
          }
        }
      }

      if (payload.length < 3) {
        for (const row of rows) {
          if (!row) continue;
          const nums = row.map(toNum).filter(v => v > 100);
          if (nums.length >= 3) {
            nums.forEach((val, idx) => {
              const m = idx + 1;
              if (m <= 12 && !found[m]) { found[m] = true; payload.push({ anno: annoPrev, mese: m, fatturato: val }); }
            });
            break;
          }
        }
      }

      if (payload.length === 0) {
        const preview = rows.slice(0, 5).map(r => (r || []).slice(0, 6).map(c => String(c ?? "")).join(" | ")).join("\n");
        throw new Error(`Nessun dato trovato. Prime righe del file:\n${preview}`);
      }

      payload.sort((a, b) => a.mese - b.mese);

      await fetch(`${SUPABASE_URL}/rest/v1/fatturato_lanci_mensile`, {
        method: "POST",
        headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`, "Content-Type": "application/json", "Prefer": "resolution=merge-duplicates" },
        body: JSON.stringify(payload),
      });
      await loadFatturato();
      const mesiNomi = ["Gen","Feb","Mar","Apr","Mag","Giu","Lug","Ago","Set","Ott","Nov","Dic"];
      const riepilogo = payload.map(p => `${mesiNomi[p.mese-1]}: €${Math.round(p.fatturato).toLocaleString("it")}`).join(" · ");
      showToast(`${annoPrev}: ${payload.length} mesi importati — ${riepilogo}`);
    } catch (err) {
      showToast(err.message, "err");
    }
    setUploading(false);
    e.target.value = "";
  };

  const novitaArricchite = useMemo(() => {
    const prenByEan = {};
    prenotato.forEach(p => {
      const t = titoli.find(t => t.id === p.titolo_id);
      if (!t || !t.ean) return;
      prenByEan[t.ean] = (prenByEan[t.ean] || 0) + p.quantita;
    });
    const objByEan = {};
    titoli.forEach(t => {
      if (!t.ean) return;
      objByEan[t.ean] = (objByEan[t.ean] || 0) + (t.obiettivo_assegnato || 0);
    });
    const novitaByEan = {};
    novitaDB.forEach(n => { if (n.ean) novitaByEan[n.ean] = n; });
    const titoliGiriByEan = {};
    titoli.forEach(t => {
      if (!t.ean || titoliGiriByEan[t.ean]) return;
      titoliGiriByEan[t.ean] = t;
    });
    return Object.values(titoliGiriByEan).map(t => {
      const nov = novitaByEan[t.ean];
      return {
        ean: t.ean,
        titolo: t.titolo,
        autore: t.autore,
        editore: t.editore_nome,
        prezzo: t.prezzo || (nov?.prezzo) || 0,
        nome_cedola: t.n_cedola || "—",
        giro_label: t.giro_label || "",
        prenotato_giri: prenByEan[t.ean] || 0,
        obiettivo_giri: objByEan[t.ean] || 0,
        num_lancio: nov?.num_lancio || null,
        copie_lanciate: nov?.copie_lanciate || 0,
        valore_lancio: nov?.valore_lancio || 0,
        data_messa_in_vendita: nov?.data_messa_in_vendita || null,
        stato_vendita: nov?.stato_vendita || null,
        risposta_editore: nov?.risposta_editore || null,
        manuale: nov?.manuale || false,
        ha_lancio: !!nov,
      };
    });
  }, [novitaDB, prenotato, titoli]);

  const anniDisponibili = useMemo(() => {
    const s = new Set();
    novitaArricchite.forEach(n => {
      // Anno dalla data di messa in vendita (fonte primaria per la proiezione)
      if (n.data_messa_in_vendita) {
        const match = String(n.data_messa_in_vendita).match(/^(\d{4})/);
        if (match) s.add(Number(match[1]));
      }
      // Anno dal giro_label come fallback
      if (n.giro_label) {
        const yr = Number(n.giro_label.split(" ").pop());
        if (yr >= 2020) s.add(yr);
      }
    });
    // Aggiungi sempre anno corrente
    s.add(new Date().getFullYear());
    return [...s].sort((a, b) => b - a);
  }, [novitaArricchite]);

  const getAnnoRecord = (n) => {
    if (n.giro_label && n.giro_label !== "EXTRA") {
      const parts = n.giro_label.split(" ");
      const yr = Number(parts[parts.length - 1]);
      if (yr >= 2020) return yr;
    }
    if (n.giro_label === "EXTRA" || !n.giro_label) {
      if (n.nome_cedola) {
        const m = n.nome_cedola.match(/\b(20\d{2})\b/);
        if (m) return Number(m[1]);
      }
      if (n.data_messa_in_vendita) {
        const m = String(n.data_messa_in_vendita).match(/^(\d{4})/);
        if (m) return Number(m[1]);
      }
      return new Date().getFullYear();
    }
    return null;
  };

  const novitaFiltrate = useMemo(() => {
    const filtered = novitaArricchite
      .filter(n => {
        // Nessun filtro anno sulla tabella: mostra tutti i titoli dei giri
        // Il filtro anno agisce solo sulla proiezione (valoriPerAnnoMese)
        return true;
      })
      .filter(n => filterEditore === "tutti" || n.editore === filterEditore)
      .filter(n => filterNumLancio.length === 0 || filterNumLancio.includes(String(n.num_lancio)))
      .filter(n => filterCedole.length === 0 || filterCedole.includes(n.nome_cedola))
      .filter(n => filterGiri.length === 0 || filterGiri.includes(n.giro_label))
      .filter(n => {
        if (filterMesi.length === 0) return true;
        if (!n.data_messa_in_vendita) return false;
        const match = String(n.data_messa_in_vendita).match(/^\d{4}-(\d{2})/);
        if (!match) return false;
        return filterMesi.includes(parseInt(match[1]));
      })
      .filter(n => {
        if (filterDataVendita.length === 0) return true;
        const haData = !!n.data_messa_in_vendita;
        if (filterDataVendita.includes("con_data") && haData) return true;
        if (filterDataVendita.includes("senza_data") && !haData) return true;
        return false;
      })
      .filter(n => {
        if (filterCopieLanciate.length === 0) return true;
        const haCopie = n.copie_lanciate > 0;
        const haData = !!n.data_messa_in_vendita;
        if (filterCopieLanciate.includes("con_copie") && haCopie) return true;
        if (filterCopieLanciate.includes("senza_copie") && !haCopie) return true;
        if (filterCopieLanciate.includes("da_compilare") && haData && !haCopie) return true;
        return false;
      })
      .filter(n => {
        if (!search) return true;
        const q = search.toLowerCase();
        return n.ean?.toLowerCase().includes(q) || n.titolo?.toLowerCase().includes(q) || n.autore?.toLowerCase().includes(q) || n.editore?.toLowerCase().includes(q);
      });
    return filtered.sort((a, b) => {
      let cmp = 0;
      if (sortKey === "cedola") cmp = (a.nome_cedola || "").localeCompare(b.nome_cedola || "");
      else if (sortKey === "prenotato") cmp = (a.prenotato_giri || 0) - (b.prenotato_giri || 0);
      else if (sortKey === "num_lancio") cmp = (a.num_lancio || 0) - (b.num_lancio || 0);
      else if (sortKey === "data_vendita") {
        const da = a.data_messa_in_vendita || "";
        const db = b.data_messa_in_vendita || "";
        cmp = da.localeCompare(db);
      }
      else if (sortKey === "copie_lanciate") cmp = (a.copie_lanciate || 0) - (b.copie_lanciate || 0);
      else if (sortKey === "stato_vendita") cmp = (a.stato_vendita || "").localeCompare(b.stato_vendita || "");
      else if (sortKey === "risposta_editore") cmp = (a.risposta_editore || "").localeCompare(b.risposta_editore || "");
      return sortDir === "asc" ? cmp : -cmp;
    });
  }, [novitaArricchite, filterAnno, filterEditore, filterNumLancio, filterCedole, filterGiri, filterMesi, filterDataVendita, filterCopieLanciate, search, sortKey, sortDir]);

  const editoriList = useMemo(() => [...new Set(novitaArricchite.map(n => n.editore).filter(Boolean))].sort(), [novitaArricchite]);
  const numLanciList = useMemo(() => [...new Set(novitaArricchite.map(n => n.num_lancio).filter(Boolean))].sort((a, b) => Number(a) - Number(b)), [novitaArricchite]);
  const cedoleList = useMemo(() => [...new Set(novitaArricchite.filter(n => filterGiri.length === 0 || filterGiri.includes(n.giro_label)).map(n => n.nome_cedola).filter(c => c && c !== "—"))].sort(), [novitaArricchite, filterGiri]);
  const giriList = useMemo(() => [...new Set(novitaArricchite.map(n => n.giro_label).filter(Boolean))].sort(), [novitaArricchite]);

  const MESI_NOMI = ["Gen","Feb","Mar","Apr","Mag","Giu","Lug","Ago","Set","Ott","Nov","Dic"];

  const meseOggi = new Date().getMonth() + 1;
  const valoriPerAnnoMese = useMemo(() => {
    // IMPORTANTE: aggrega per anno/mese di data_messa_in_vendita, NON per anno del giro.
    // Un titolo della cedola 2025 lanciato nel 2026 va in 2026, non in 2025.
    const map = {};
    novitaArricchite.forEach(n => {
      if (!n.data_messa_in_vendita || !n.valore_lancio || n.valore_lancio === 0) return;
      const match = String(n.data_messa_in_vendita).match(/^(\d{4})-(\d{2})/);
      if (!match) return;
      const anno = parseInt(match[1]);
      const mese = parseInt(match[2]);
      if (!map[anno]) map[anno] = {};
      map[anno][mese] = (map[anno][mese] || 0) + (n.valore_lancio || 0);
    });
    return map;
  }, [novitaArricchite]);

  const annoCorrentePerMese = useMemo(() => valoriPerAnnoMese[annoRif] || {}, [valoriPerAnnoMese, annoRif]);

  const annoPrecPerMese = useMemo(() => {
    const perMese = {};
    fatturato.filter(f => f.anno === annoPrev && (filterMesi.length === 0 || filterMesi.includes(f.mese))).forEach(f => { perMese[f.mese] = f.fatturato || 0; });
    return perMese;
  }, [fatturato, annoPrev, filterMesi]);

  const kpi = useMemo(() => {
    // KPI calcolati sui titoli dell'anno proiezione selezionato (filterAnno)
    // La tabella sotto mostra tutti, ma i KPI riflettono solo l'anno corrente/selezionato
    const novitaAnno = filterAnno
      ? novitaFiltrate.filter(n => getAnnoRecord(n) === filterAnno)
      : novitaFiltrate;

    const totTitoli = novitaAnno.length;
    const lanciati = novitaAnno.filter(n => n.copie_lanciate > 0 && !n.manuale);
    const numLanciati = lanciati.length;
    const valoreLancio = lanciati.reduce((s, n) => s + (n.valore_lancio || 0), 0);
    const sbloccati = novitaAnno.filter(n => n.copie_lanciate > 0 && n.manuale);
    const numSbloccati = sbloccati.length;
    const valoreSbloccato = sbloccati.reduce((s, n) => s + (n.valore_lancio || 0), 0);
    const totTrasmessi = numLanciati + numSbloccati;
    const pctAvanzamento = totTitoli > 0 ? Math.round(totTrasmessi / totTitoli * 100) : 0;
    const nonTrasmessi = totTitoli - totTrasmessi;
    const valPrenotato = novitaAnno.reduce((s, n) => s + (n.prezzo || 0) * n.prenotato_giri, 0);
    const nonLanciati = novitaAnno.filter(n => n.prenotato_giri > 0 && (!n.copie_lanciate || n.copie_lanciate === 0));
    const copieNonLanciate = nonLanciati.reduce((s, n) => s + n.prenotato_giri, 0);
    const valNonLanciato = nonLanciati.reduce((s, n) => s + (n.prezzo || 0) * n.prenotato_giri, 0);
    const oggi = new Date();
    const meseCorrente = oggi.getMonth() + 1;
    const ytdCorrente = Object.values(annoCorrentePerMese).reduce((s, v) => s + v, 0);
    const totaleAnnoPrev = Object.values(annoPrecPerMese).reduce((s, v) => s + v, 0);
    const ytdPrev = Object.entries(annoPrecPerMese).filter(([m]) => Number(m) <= meseCorrente).reduce((s, [, v]) => s + v, 0);
    const futuriPrev = Object.entries(annoPrecPerMese).filter(([m]) => Number(m) > meseCorrente).reduce((s, [, v]) => s + v, 0);
    const trend = ytdPrev > 0 ? ytdCorrente / ytdPrev : 1;
    const pipeline = valNonLanciato;
    const proiezione = ytdCorrente + pipeline + (futuriPrev * trend);
    const haFatturatoPrec = totaleAnnoPrev > 0;
    return {
      totTitoli, nonTrasmessi, valPrenotato,
      numLanciati, valoreLancio, numSbloccati, valoreSbloccato,
      totTrasmessi, pctAvanzamento,
      copieNonLanciate, valNonLanciato, numNonLanciati: nonLanciati.length,
      ytdCorrente, totaleAnnoPrev, ytdPrev, trend, pipeline, proiezione, haFatturatoPrec, meseCorrente,
    };
  }, [novitaFiltrate, filterAnno, annoCorrentePerMese, annoPrecPerMese]);

  // ─────────────────────────────────────────────────────────────────
  // UPLOAD CSV — FIX mapping colonne con mese e anno separati
  // ─────────────────────────────────────────────────────────────────
  const handleCSVUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true);
    try {
      const buf = await file.arrayBuffer();
      const bytes = new Uint8Array(buf);
      let text;
      if ((bytes[0] === 0xFF && bytes[1] === 0xFE) || (bytes[0] === 0xFE && bytes[1] === 0xFF)) {
        const decoder = new TextDecoder(bytes[0] === 0xFF ? "utf-16le" : "utf-16be");
        text = decoder.decode(buf);
      } else {
        text = new TextDecoder("utf-8").decode(buf);
      }
      text = text.replace(/^\ufeff/, "").replace(/\0/g, "");

      const lines = text.split(/\r?\n/).filter(l => l.trim());
      if (lines.length === 0) throw new Error("File vuoto");

      const sep = lines[0].includes("\t") ? "\t" : lines[0].includes(";") ? ";" : ",";
      const headers = lines[0].split(sep).map(h => h.trim().replace(/^[\ufeff\ufffe"]+|["]+$/g, ""));

      const rows = [];
      for (let i = 1; i < lines.length; i++) {
        const vals = lines[i].split(sep).map(v => v.trim().replace(/^"|"$/g, ""));
        if (vals.length >= headers.length - 1) rows.push(vals);
      }

      // ── MAPPING COLONNE (FIX: gestione "Mese di Data Pubblicaz." e "Anno di Data Pubblicaz.") ──
      const colMap = {};
      headers.forEach((h, i) => {
        const hl = h.toLowerCase().trim();
        if (hl.includes("ean")) colMap.ean = i;
        else if (hl === "titolo" || (hl.includes("titolo") && !hl.includes("sub"))) colMap.titolo = i;
        else if (hl.includes("autore")) colMap.autore = i;
        else if (hl.includes("risposta") && hl.includes("editore")) colMap.risposta_editore = i;
        else if (hl.includes("editore") && !hl.includes("risposta")) colMap.editore = i;
        else if (hl.includes("prezzo")) colMap.prezzo = i;
        else if (hl.includes("num") && hl.includes("lancio")) colMap.num_lancio = i;
        // "Mese, giorno, anno di Data di riferimento" → data completa (formato "2 gennaio 2026")
        else if (hl.includes("giorno") && hl.includes("anno")) colMap.data = i;
        // "Mese di Data Pubblicaz." → colMap.mese
        else if (hl.includes("mese") && (hl.includes("pubbl") || hl.includes("data") || hl.includes("vendita") || hl.includes("riferimento")) && !hl.includes("giorno")) colMap.mese = i;
        else if (hl === "mese" || hl === "month") colMap.mese = i;
        // "Anno di Data Pubblicaz." / "Anno di Data di riferimento" → colMap.anno
        else if (hl.includes("anno") && (hl.includes("pubbl") || hl.includes("data") || hl.includes("vendita") || hl.includes("riferimento")) && !hl.includes("giorno")) colMap.anno = i;
        else if (hl === "anno" || hl === "year") colMap.anno = i;
        // Data completa (es. "Data Vendita": DD/MM/YYYY)
        else if (hl.includes("data") && (hl.includes("pubbl") || hl.includes("vendita")) && !hl.includes("mese") && !hl.includes("anno")) colMap.data = i;
        else if (hl.includes("copie") && hl.includes("lanciate")) colMap.copie = i;
        else if (hl.includes("valore") && hl.includes("lancio")) colMap.valore = i;
        else if (hl.includes("fatturato")) colMap.valore = i;
        else if (hl === "lancio") colMap.copie = i;
        // FIX: "c.Qtà Mov. Uscita" → copie (separatore migliaia IT gestito da parseCopie)
        else if (hl.includes("mov") || hl.includes("uscita")) colMap.copie = i;
        else if (hl.includes("stato") && hl.includes("vendita")) colMap.stato_vendita = i;
      });

      if (colMap.ean === undefined) throw new Error(`Colonna EAN non trovata. Headers: ${headers.join(" | ")}`);

      const payload = rows.map(r => {
        let dataObj = null;

        // ── CASO 1: mese + anno separati (formato Tableau/Messaggerie) ──
        if (colMap.mese !== undefined) {
          const meseVal = String(r[colMap.mese] || "").trim().toLowerCase();
          const meseNum = MESI_IT[meseVal] !== undefined ? MESI_IT[meseVal] + 1 : parseInt(meseVal) || null;
          // FIX: leggi anno dalla colonna dedicata, oppure fallback a filterAnno
          const annoVal = colMap.anno !== undefined
            ? parseInt(String(r[colMap.anno] || "").trim()) || (filterAnno || new Date().getFullYear())
            : (filterAnno || new Date().getFullYear());
          if (meseNum >= 1 && meseNum <= 12) {
            // Costruzione timezone-safe (no toISOString)
            dataObj = { year: annoVal, month: meseNum };
          }
        }

        // ── CASO 2: data completa in una colonna ──
        if (!dataObj && colMap.data !== undefined) {
          const dataStr = r[colMap.data];
          if (dataStr) {
            const parsed = parseDataIt(dataStr);
            if (parsed) {
              dataObj = { year: parsed.getFullYear(), month: parsed.getMonth() + 1 };
            } else {
              const m = String(dataStr).trim().match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
              if (m) dataObj = { year: Number(m[3].length === 2 ? "20" + m[3] : m[3]), month: Number(m[2]) };
              const m2 = String(dataStr).trim().match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
              if (m2) dataObj = { year: Number(m2[1]), month: Number(m2[2]) };
            }
          }
        }

        // Formato stringa YYYY-MM-01 timezone-safe
        const dataStr = dataObj
          ? `${dataObj.year}-${String(dataObj.month).padStart(2, "0")}-01`
          : null;

        // FIX: usa parseCopie per gestire separatori migliaia IT (es. "115.482" → 115482)
        const copie = colMap.copie !== undefined ? parseCopie(r[colMap.copie]) : 0;
        const valore = colMap.valore !== undefined ? parseValoreLancio(r[colMap.valore]) : 0;

        return {
          ean: String(r[colMap.ean] || "").trim(),
          titolo: r[colMap.titolo] || "",
          autore: r[colMap.autore] || "",
          editore: r[colMap.editore] || "",
          prezzo: parsePrezzo(r[colMap.prezzo]),
          num_lancio: colMap.num_lancio !== undefined ? parseInt(String(r[colMap.num_lancio]).replace(/^0+/, "") || "0") || null : null,
          copie_lanciate: copie,
          valore_lancio: valore,
          data_messa_in_vendita: dataStr,
          stato_vendita: colMap.stato_vendita !== undefined ? (String(r[colMap.stato_vendita] || "").trim().split(" ")[0] || null) : null,
          risposta_editore: colMap.risposta_editore !== undefined ? (String(r[colMap.risposta_editore] || "").trim().split(" ")[0] || null) : null,
          manuale: false,
        };
      }).filter(r => r.ean.length >= 10);

      // Aggrega per EAN: somma copie + valore, prende la data minima (prima messa in vendita)
      const byEan = {};
      payload.forEach(r => {
        const existing = byEan[r.ean];
        if (!existing) { byEan[r.ean] = { ...r }; return; }
        // Somma copie e valore
        byEan[r.ean].copie_lanciate = (existing.copie_lanciate || 0) + (r.copie_lanciate || 0);
        byEan[r.ean].valore_lancio = (existing.valore_lancio || 0) + (r.valore_lancio || 0);
        // Data minima (prima messa in vendita)
        if (r.data_messa_in_vendita && (!existing.data_messa_in_vendita || r.data_messa_in_vendita < existing.data_messa_in_vendita)) {
          byEan[r.ean].data_messa_in_vendita = r.data_messa_in_vendita;
        }
        // Prendi num_lancio, stato_vendita, risposta_editore dal primo record non nullo
        if (!existing.num_lancio && r.num_lancio) byEan[r.ean].num_lancio = r.num_lancio;
        if (!existing.stato_vendita && r.stato_vendita) byEan[r.ean].stato_vendita = r.stato_vendita;
        if (!existing.risposta_editore && r.risposta_editore) byEan[r.ean].risposta_editore = r.risposta_editore;
      });
      const deduplicated = Object.values(byEan);

      const eanGiri = new Set(titoli.map(t => t.ean).filter(Boolean));
      const soloGiri = deduplicated.filter(r => eanGiri.has(r.ean));
      // EAN nel CSV ma NON nei giri → report
      const nonNeiGiri = deduplicated.filter(r => !eanGiri.has(r.ean));

      const existingByEan = {};
      novitaDB.forEach(n => { if (n.ean) existingByEan[n.ean] = n; });

      const toInsert = [];
      const toUpdate = [];
      let skipped = 0;

      soloGiri.forEach(row => {
        const existing = existingByEan[row.ean];
        if (!existing) {
          toInsert.push(row);
        } else if (existing.manuale) {
          skipped++;
        } else {
          const changed = row.num_lancio !== existing.num_lancio ||
            row.copie_lanciate !== existing.copie_lanciate ||
            Math.abs((row.valore_lancio || 0) - (existing.valore_lancio || 0)) > 0.01 ||
            row.data_messa_in_vendita !== existing.data_messa_in_vendita ||
            row.stato_vendita !== existing.stato_vendita ||
            row.risposta_editore !== existing.risposta_editore;
          if (changed) toUpdate.push(row);
          else skipped++;
        }
      });

      for (let i = 0; i < toInsert.length; i += 500) {
        const batch = toInsert.slice(i, i + 500);
        const r = await fetch(`${SUPABASE_URL}/rest/v1/titoli_novita`, {
          method: "POST",
          headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`, "Content-Type": "application/json", "Prefer": "return=minimal" },
          body: JSON.stringify(batch),
        });
        if (!r.ok) {
          const err = await r.text();
          throw new Error(`Errore insert: ${err}`);
        }
      }

      for (const row of toUpdate) {
        await fetch(`${SUPABASE_URL}/rest/v1/titoli_novita?ean=eq.${encodeURIComponent(row.ean)}`, {
          method: "PATCH",
          headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`, "Content-Type": "application/json", "Prefer": "return=minimal" },
          body: JSON.stringify({
            num_lancio: row.num_lancio,
            copie_lanciate: row.copie_lanciate,
            valore_lancio: row.valore_lancio,
            data_messa_in_vendita: row.data_messa_in_vendita,
            stato_vendita: row.stato_vendita,
            risposta_editore: row.risposta_editore,
          }),
        });
      }

      const msg = `${toInsert.length} nuovi · ${toUpdate.length} aggiornati · ${skipped} invariati${nonNeiGiri.length > 0 ? ` · ${nonNeiGiri.length} EAN non trovati nei giri` : ""}`;
      showToast(msg);
      await loadNovita();

      // Report EAN non trovati → download Excel automatico
      if (nonNeiGiri.length > 0) {
        const XLSX = window.XLSX;
        const headers = ["EAN", "Num. Lancio", "Data Vendita", "Copie", "Valore €"];
        const rows = nonNeiGiri.map(r => [r.ean, r.num_lancio || "", r.data_messa_in_vendita || "", r.copie_lanciate || 0, r.valore_lancio || 0]);
        const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "EAN non trovati");
        XLSX.writeFile(wb, `EAN_non_trovati_${new Date().toISOString().slice(0,10)}.xlsx`);
      }
    } catch (err) {
      showToast(err.message, "err");
    }
    setUploading(false);
    e.target.value = "";
  };

  const saveManual = async () => {
    if (!manualForm.ean || manualForm.ean.length < 10) { showToast("EAN non valido", "err"); return; }
    const payload = {
      ean: manualForm.ean.trim(),
      titolo: manualForm.titolo,
      autore: manualForm.autore,
      editore: manualForm.editore,
      prezzo: parseFloat(manualForm.prezzo) || 0,
      num_lancio: parseInt(manualForm.num_lancio) || null,
      copie_lanciate: parseInt(manualForm.copie_lanciate) || 0,
      valore_lancio: parseFloat(manualForm.valore_lancio) || 0,
      data_messa_in_vendita: manualForm.data_messa_in_vendita || null,
      manuale: true,
    };
    const r = await fetch(`${SUPABASE_URL}/rest/v1/titoli_novita`, {
      method: "POST",
      headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`, "Content-Type": "application/json", "Prefer": "resolution=merge-duplicates" },
      body: JSON.stringify([payload]),
    });
    if (r.ok) {
      showToast("Titolo aggiunto manualmente");
      setManualForm({ ean: "", titolo: "", autore: "", editore: "", prezzo: "", num_lancio: "", copie_lanciate: "", valore_lancio: "", data_messa_in_vendita: "" });
      setManualOpen(false);
      await loadNovita();
    } else { showToast("Errore salvataggio", "err"); }
  };

  const saveInlineEdit = async (ean, value) => {
    const qta = parseInt(value) || 0;
    const nov = novitaDB.find(n => n.ean === ean);
    const prezzo = nov?.prezzo || novitaArricchite.find(n => n.ean === ean)?.prezzo || 0;
    const titGiro = titoli.find(t => t.ean === ean);
    const payload = {
      ean,
      titolo: nov?.titolo || titGiro?.titolo || "",
      autore: nov?.autore || titGiro?.autore || "",
      editore: nov?.editore || titGiro?.editore_nome || "",
      prezzo: prezzo,
      copie_lanciate: qta,
      valore_lancio: Math.round(prezzo * qta * 100) / 100,
      manuale: true,
    };
    const r = await fetch(`${SUPABASE_URL}/rest/v1/titoli_novita`, {
      method: "POST",
      headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`, "Content-Type": "application/json", "Prefer": "resolution=merge-duplicates" },
      body: JSON.stringify([payload]),
    });
    if (r.ok) {
      showToast("Aggiornato");
      await loadNovita();
    } else { showToast("Errore salvataggio", "err"); }
    setEditingEan(null);
  };

  const exportExcel = () => {
    const XLSX = window.XLSX;
    const headers = ["CEDOLA","EAN","TITOLO","AUTORE","EDITORE","PREZZO","PRENOTATO TOTALE","N. LANCIO","COPIE LANCIATE","VALORE LANCIO","DATA MESSA IN VENDITA","SV","RE"];
    const rows = novitaFiltrate.map(n => [
      n.nome_cedola, n.ean, n.titolo, n.autore, n.editore, n.prezzo,
      n.prenotato_giri, (n.manuale && n.copie_lanciate > 0) ? "SBL/RIFO" : (n.num_lancio || ""), n.copie_lanciate,
      n.valore_lancio, fmtDate(n.data_messa_in_vendita), n.stato_vendita || "", n.risposta_editore || ""
    ]);
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `Novità ${filterAnno || "Tutti"}`);
    XLSX.writeFile(wb, `Avanzamento_Novita_${filterAnno || "Tutti"}.xlsx`);
  };

  if (loading) return <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", color: T.textMid }}>Caricamento titoli novità...</div>;

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%", overflow: "hidden" }}>
      {/* DASHBOARD KPI */}
      <div style={{ padding: "16px 20px", borderBottom: `1px solid ${T.border}` }}>
        <div style={{ display: "flex", gap: 10, alignItems: "center", marginBottom: 14, flexWrap: "wrap" }}>
          <select style={{ ...css.input, fontSize: "14px", fontWeight: "600", color: T.accent }} value={filterAnno || ""} onChange={e => setFilterAnno(e.target.value ? Number(e.target.value) : null)} title="Anno proiezione">
            <option value="">— Anno proiezione —</option>
            {anniDisponibili.map(a => <option key={a} value={a}>{a}</option>)}
          </select>
          <SearchableMultiSelect values={filterEditore === "tutti" ? [] : [filterEditore]} onChange={v => setFilterEditore(v.length > 0 ? v[0] : "tutti")} options={editoriList} placeholder="Tutti gli editori" width={190} />
          <SearchableMultiSelect values={filterNumLancio} onChange={setFilterNumLancio} options={numLanciList.map(String)} placeholder="Tutti i lanci" width={150} />
          <SearchableMultiSelect values={filterGiri} onChange={setFilterGiri} options={giriList} placeholder="Tutti i giri" width={160} />
          <SearchableMultiSelect values={filterCedole} onChange={setFilterCedole} options={cedoleList} placeholder="Tutte le cedole" width={170} />
          <SearchableMultiSelect
            values={filterMesi.map(String)}
            onChange={v => setFilterMesi(v.map(Number))}
            options={["1","2","3","4","5","6","7","8","9","10","11","12"]}
            renderOption={v => ["Gen","Feb","Mar","Apr","Mag","Giu","Lug","Ago","Set","Ott","Nov","Dic"][Number(v)-1]}
            placeholder="Tutti i mesi"
            width={150}
          />
          <input style={{ ...css.input, width: 180 }} placeholder="Cerca EAN / titolo..." value={search} onChange={e => setSearch(e.target.value)} />
          <SearchableMultiSelect
            values={filterDataVendita}
            onChange={setFilterDataVendita}
            options={["con_data","senza_data"]}
            renderOption={v => v === "con_data" ? "Con data vendita" : "Senza data vendita"}
            placeholder="Data: tutte"
            width={160}
          />
          <SearchableMultiSelect
            values={filterCopieLanciate}
            onChange={setFilterCopieLanciate}
            options={["con_copie","senza_copie","da_compilare"]}
            renderOption={v => v === "con_copie" ? "Con copie lanciate" : v === "senza_copie" ? "Senza copie lanciate" : "⚠ Da compilare"}
            placeholder="Copie: tutte"
            width={160}
          />
          <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
            {ruolo !== "agente" && (
              <>
                <label style={{ ...css.btn("accent"), cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 4 }}>
                  {uploading ? "Caricamento..." : "↑ Upload CSV"}
                  <input type="file" accept=".csv,.txt,.tsv" style={{ display: "none" }} onChange={handleCSVUpload} disabled={uploading} />
                </label>
                <label style={{ ...css.btn(), cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 4, borderColor: T.purple, color: T.purple }}>
                  ↑ Fatturato Anno Precedente
                  <input type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={handleFatturatoUpload} disabled={uploading} />
                </label>
              </>
            )}
            <button style={css.btn()} onClick={exportExcel}>↓ Excel</button>
          </div>
        </div>
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
          <KpiCard label={`Titoli novità ${filterAnno || ""}`} value={kpi.totTitoli.toLocaleString("it")} color={T.text} sub={`${kpi.nonTrasmessi} da lanciare/sbloccare`} />
          <KpiCard label="Valore prenotato" value={`€ ${kpi.valPrenotato.toLocaleString("it", { maximumFractionDigits: 0 })}`} color={T.green} />
          <KpiCard label="Valore lancio" value={`€ ${kpi.valoreLancio.toLocaleString("it", { maximumFractionDigits: 0 })}`} color={T.accent} sub={`${kpi.numLanciati} titoli lanciati`} />
          <KpiCard label="Valore SBL/Rifornimento" value={`€ ${kpi.valoreSbloccato.toLocaleString("it", { maximumFractionDigits: 0 })}`} color="#e8a838" sub={`${kpi.numSbloccati} titoli sbloccati`} />
          <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 4, padding: "16px 20px", minWidth: 200 }}>
            <div style={{ color: T.textMid, fontSize: "10px", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 8 }}>Avanzamento</div>
            <div style={{ color: kpi.pctAvanzamento >= 80 ? T.green : kpi.pctAvanzamento >= 50 ? T.accent : T.red, fontSize: "28px", fontWeight: "700", lineHeight: 1, marginBottom: 8 }}>{kpi.pctAvanzamento}%</div>
            <div style={{ height: 8, background: T.borderHi, borderRadius: 4, overflow: "hidden", display: "flex" }}>
              <div style={{ width: `${Math.round(kpi.numLanciati / Math.max(kpi.totTitoli, 1) * 100)}%`, height: "100%", background: T.accent }} />
              <div style={{ width: `${Math.round(kpi.numSbloccati / Math.max(kpi.totTitoli, 1) * 100)}%`, height: "100%", background: "#e8a838" }} />
            </div>
            <div style={{ display: "flex", gap: 12, marginTop: 6, fontSize: "10px" }}>
              <span style={{ color: T.accent }}>● {kpi.numLanciati} lanciati</span>
              <span style={{ color: "#e8a838" }}>● {kpi.numSbloccati} sbloccati</span>
            </div>
          </div>
          <KpiCard label="Pren. non lanciato" value={`€ ${kpi.valNonLanciato.toLocaleString("it", { maximumFractionDigits: 0 })}`} color={T.red} sub={`${kpi.copieNonLanciate.toLocaleString("it")} copie · ${kpi.numNonLanciati} titoli`} />
          <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 4, padding: "16px 20px", minWidth: 200, cursor: "pointer" }} onClick={() => setShowProiezione(p => !p)}>
            <div style={{ color: T.textMid, fontSize: "10px", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 8, display: "flex", justifyContent: "space-between" }}>
              Proiezione anno <span>{showProiezione ? "▲" : "▼"}</span>
            </div>
            <div style={{ color: T.purple, fontSize: "24px", fontWeight: "700", lineHeight: 1, marginBottom: 4 }}>€ {kpi.proiezione.toLocaleString("it", { maximumFractionDigits: 0 })}</div>
            <div style={{ color: T.textMid, fontSize: "10px" }}>
              {kpi.haFatturatoPrec ? `trend ${kpi.trend >= 1 ? "+" : ""}${Math.round((kpi.trend - 1) * 100)}% vs ${annoPrev}` : "carica fatturato anno prec."}
            </div>
          </div>
        </div>

        {/* SEZIONE PROIEZIONE ESPANDIBILE */}
        {showProiezione && (
          <div style={{ marginTop: 14, padding: 16, background: T.bg, border: `1px solid ${T.border}`, borderRadius: 6 }}>
            <div style={{ color: T.purple, fontSize: "11px", fontWeight: "700", letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 12 }}>
              Confronto {annoPrev} vs {annoRif} — Proiezione
            </div>
            <div style={{ overflowX: "auto", marginBottom: 14 }}>
              <table style={{ ...css.table, fontSize: "11px" }}>
                <thead>
                  <tr>
                    <th style={css.th}></th>
                    {MESI_NOMI.map((m, i) => <th key={i} style={{ ...css.th, textAlign: "center", minWidth: 60, color: (i + 1) <= kpi.meseCorrente ? T.text : T.textDim }}>{m}</th>)}
                    <th style={{ ...css.th, textAlign: "center", fontWeight: "700" }}>Totale</th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td style={{ ...css.td, fontWeight: "700", color: T.accent, whiteSpace: "nowrap" }}>{annoRif}</td>
                    {MESI_NOMI.map((_, i) => {
                      const val = annoCorrentePerMese[i + 1] || 0;
                      const prev = annoPrecPerMese[i + 1] || 0;
                      const isFuture = (i + 1) > kpi.meseCorrente;
                      return <td key={i} style={{ ...css.td, textAlign: "right", color: isFuture ? T.textDim : val > prev ? T.green : val > 0 ? T.accent : T.textDim, fontWeight: val > prev ? "700" : "400" }}>
                        {val > 0 ? `€ ${Math.round(val).toLocaleString("it")}` : "—"}
                      </td>;
                    })}
                    <td style={{ ...css.td, textAlign: "right", fontWeight: "700", color: T.accent }}>€ {kpi.ytdCorrente.toLocaleString("it", { maximumFractionDigits: 0 })}</td>
                  </tr>
                  <tr>
                    <td style={{ ...css.td, fontWeight: "700", color: T.textMid, whiteSpace: "nowrap" }}>{annoPrev}</td>
                    {MESI_NOMI.map((_, i) => {
                      const val = annoPrecPerMese[i + 1] || 0;
                      return <td key={i} style={{ ...css.td, textAlign: "right", color: val > 0 ? T.textMid : T.textDim }}>{val > 0 ? `€ ${Math.round(val).toLocaleString("it")}` : "—"}</td>;
                    })}
                    <td style={{ ...css.td, textAlign: "right", fontWeight: "700", color: T.textMid }}>€ {kpi.totaleAnnoPrev.toLocaleString("it", { maximumFractionDigits: 0 })}</td>
                  </tr>
                  {kpi.haFatturatoPrec && (
                    <tr>
                      <td style={{ ...css.td, fontWeight: "600", color: T.textDim, fontSize: "10px" }}>Δ%</td>
                      {MESI_NOMI.map((_, i) => {
                        const cur = annoCorrentePerMese[i + 1] || 0;
                        const prev = annoPrecPerMese[i + 1] || 0;
                        const isFuture = (i + 1) > kpi.meseCorrente;
                        if (isFuture || !prev) return <td key={i} style={{ ...css.td, textAlign: "right", color: T.textDim }}>—</td>;
                        const delta = Math.round((cur / prev - 1) * 100);
                        return <td key={i} style={{ ...css.td, textAlign: "right", fontSize: "10px", fontWeight: "700", color: delta >= 0 ? T.green : T.red }}>{delta >= 0 ? "+" : ""}{delta}%</td>;
                      })}
                      <td style={{ ...css.td, textAlign: "right", fontSize: "10px", fontWeight: "700", color: kpi.trend >= 1 ? T.green : T.red }}>
                        {kpi.ytdPrev > 0 ? `${kpi.trend >= 1 ? "+" : ""}${Math.round((kpi.trend - 1) * 100)}%` : "—"}
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
            <div style={{ display: "flex", gap: 16, fontSize: "11px", color: T.textMid, flexWrap: "wrap", padding: "12px 0", borderTop: `1px solid ${T.border}` }}>
              <div>
                <div style={{ fontSize: "9px", textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 3 }}>YTD {annoRif}</div>
                <div style={{ color: T.accent, fontWeight: "700", fontSize: "16px" }}>€ {kpi.ytdCorrente.toLocaleString("it", { maximumFractionDigits: 0 })}</div>
              </div>
              {kpi.haFatturatoPrec && <div>
                <div style={{ fontSize: "9px", textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 3 }}>YTD {annoPrev}</div>
                <div style={{ color: T.textMid, fontWeight: "700", fontSize: "16px" }}>€ {kpi.ytdPrev.toLocaleString("it", { maximumFractionDigits: 0 })}</div>
              </div>}
              <div>
                <div style={{ fontSize: "9px", textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 3 }}>Pipeline (pren. da lanciare)</div>
                <div style={{ color: "#e8a838", fontWeight: "700", fontSize: "16px" }}>€ {kpi.pipeline.toLocaleString("it", { maximumFractionDigits: 0 })}</div>
              </div>
              {kpi.haFatturatoPrec && <div>
                <div style={{ fontSize: "9px", textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 3 }}>Trend vs {annoPrev}</div>
                <div style={{ color: kpi.trend >= 1 ? T.green : T.red, fontWeight: "700", fontSize: "16px" }}>{kpi.trend >= 1 ? "+" : ""}{Math.round((kpi.trend - 1) * 100)}%</div>
              </div>}
              <div>
                <div style={{ fontSize: "9px", textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 3 }}>🔮 Proiezione {annoRif}</div>
                <div style={{ color: T.purple, fontWeight: "700", fontSize: "20px" }}>€ {kpi.proiezione.toLocaleString("it", { maximumFractionDigits: 0 })}</div>
              </div>
            </div>
            <div style={{ fontSize: "10px", color: T.textDim, fontStyle: "italic", marginTop: 4 }}>
              {kpi.haFatturatoPrec
                ? `Proiezione = YTD ${annoRif} + pipeline prenotato + (mesi futuri ${annoPrev} × trend ${Math.round(kpi.trend * 100)}%). Include stima giro 4 e giro 5.`
                : `Carica il fatturato ${annoPrev} con il bottone "↑ Fatturato ${annoPrev}" per avere una proiezione basata sul confronto anno su anno.`
              }
            </div>
            <div style={{ fontSize: "10px", color: T.textDim, marginTop: 6 }}>
              ⓘ I valori <span style={{ color: T.accent }}>{annoRif}</span> sono calcolati dal campo <b>valore_lancio</b> caricato via CSV (aggregato per mese di messa in vendita). Se un mese risulta -% rispetto all'anno precedente, verifica che il CSV sia aggiornato con tutti i titoli del mese.
            </div>
          </div>
        )}
      </div>

      {/* TABELLA */}
      <div style={{ flex: 1, overflowY: "auto", overflowX: "auto" }}>
        <table style={css.table}>
          <thead>
            <tr>
              <th style={{ ...css.th, cursor: "pointer" }} onClick={() => toggleSort("cedola")}>Cedola{sortIcon("cedola")}</th>
              <th style={css.th}>EAN</th>
              <th style={css.th}>Titolo</th>
              <th style={css.th}>Autore</th>
              <th style={css.th}>Editore</th>
              <th style={css.th}>€</th>
              <th style={{ ...css.th, cursor: "pointer" }} onClick={() => toggleSort("prenotato")}>Pren. Tot.{sortIcon("prenotato")}</th>
              <th style={{ ...css.th, cursor: "pointer" }} onClick={() => toggleSort("num_lancio")}>N. Lancio{sortIcon("num_lancio")}</th>
              <th style={{ ...css.th, cursor: "pointer" }} onClick={() => toggleSort("copie_lanciate")}>Copie Lanc.{sortIcon("copie_lanciate")}</th>
              <th style={css.th}>Val. Lancio</th>
              <th style={{ ...css.th, cursor: "pointer" }} onClick={() => toggleSort("data_vendita")}>Data Vendita{sortIcon("data_vendita")}</th>
              <th style={{ ...css.th, cursor: "pointer" }} onClick={() => toggleSort("stato_vendita")}>SV{sortIcon("stato_vendita")}</th>
              <th style={{ ...css.th, cursor: "pointer" }} onClick={() => toggleSort("risposta_editore")}>RE{sortIcon("risposta_editore")}</th>
            </tr>
          </thead>
          <tbody>
            {novitaFiltrate.map((n, i) => {
              const pct = n.obiettivo_giri > 0 ? Math.round(n.prenotato_giri / n.obiettivo_giri * 100) : 0;
              const isEditingThis = editingEan === n.ean;
              return (
                <tr key={n.ean || i} style={{ background: i % 2 === 0 ? "transparent" : T.surface + "66" }}>
                  <td style={{ ...css.td, color: T.textMid, fontSize: "10px", whiteSpace: "nowrap" }}>{n.nome_cedola}</td>
                  <td style={{ ...css.td, fontFamily: "monospace", fontSize: "11px", color: T.textMid }}>{n.ean}</td>
                  <td style={{ ...css.td, maxWidth: 220 }}><div style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontWeight: "600" }}>{n.titolo}</div></td>
                  <td style={{ ...css.td, color: T.textMid }}>{n.autore}</td>
                  <td style={{ ...css.td, color: T.accent, fontWeight: "600", whiteSpace: "nowrap" }}>{n.editore}</td>
                  <td style={{ ...css.td, whiteSpace: "nowrap" }}>€ {(n.prezzo || 0).toFixed(2)}</td>
                  <td style={css.td}>
                    <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                      <span style={{ color: n.prenotato_giri > 0 ? T.green : T.textDim, fontWeight: "600" }}>{n.prenotato_giri > 0 ? n.prenotato_giri.toLocaleString("it") : "—"}</span>
                      {n.obiettivo_giri > 0 && <span style={{ color: pct >= 80 ? T.green : pct >= 50 ? T.accent : T.red, fontSize: "10px", fontWeight: "700" }}>{pct}%</span>}
                    </div>
                  </td>
                  <td style={{ ...css.td, textAlign: "center" }}>{n.manuale && n.copie_lanciate > 0 ? <span style={css.tag("#e8a838")}>SBL/RIFO</span> : (n.num_lancio || "—")}</td>
                  <td style={css.td}>
                    {isEditingThis ? (
                      <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
                        <input type="number" style={{ ...css.input, width: 70, padding: "2px 6px" }} value={editingVal} autoFocus
                          onChange={e => setEditingVal(e.target.value)}
                          onKeyDown={e => { if (e.key === "Enter") saveInlineEdit(n.ean, editingVal); if (e.key === "Escape") setEditingEan(null); }} />
                        <button style={{ ...css.btn("accent"), padding: "2px 8px", fontSize: "11px" }} onClick={() => saveInlineEdit(n.ean, editingVal)}>✓</button>
                      </div>
                    ) : (
                      <div style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer" }}
                        onClick={() => { setEditingEan(n.ean); setEditingVal(String(n.copie_lanciate || 0)); }}>
                        <span style={{ color: n.copie_lanciate > 0 ? T.text : T.textDim }}>{n.copie_lanciate > 0 ? n.copie_lanciate.toLocaleString("it") : "—"}</span>
                        <span style={{ color: T.accent, fontSize: "11px" }}>✎</span>
                      </div>
                    )}
                  </td>
                  <td style={{ ...css.td, textAlign: "right", whiteSpace: "nowrap" }}>{n.valore_lancio > 0 ? `€ ${n.valore_lancio.toLocaleString("it", { maximumFractionDigits: 0 })}` : "—"}</td>
                  <td style={{ ...css.td, whiteSpace: "nowrap" }}>{fmtDate(n.data_messa_in_vendita)}</td>
                  <td style={{ ...css.td, color: T.textMid, fontSize: "11px" }}>{n.stato_vendita || "—"}</td>
                  <td style={{ ...css.td, color: T.textMid, fontSize: "11px" }}>{n.risposta_editore || "—"}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
        {novitaFiltrate.length === 0 && (
          <div style={{ padding: 40, textAlign: "center", color: T.textDim }}>
            {novitaDB.length === 0 ? "Nessun titolo novità caricato. Usa \"Upload CSV\" per importare il catalogo." : "Nessun titolo corrisponde ai filtri selezionati."}
          </div>
        )}
      </div>

      {/* TOAST */}
      {toast && (
        <div style={{ position: "fixed", bottom: 24, left: "50%", transform: "translateX(-50%)", background: toast.type === "err" ? "#4a1a2a" : "#1a3a2a", border: `1px solid ${toast.type === "err" ? T.red : T.green}`, color: toast.type === "err" ? T.red : T.green, borderRadius: 6, padding: "8px 20px", fontSize: "12px", zIndex: 999, boxShadow: "0 4px 20px #0008" }}>
          {toast.msg}
        </div>
      )}
    </div>
  );
}

// ────────────────────────────────────────────────────────────────
// MODULO LANCI SETTIMANALI
// ────────────────────────────────────────────────────────────────
