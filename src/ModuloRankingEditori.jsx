import { useState, useEffect, useMemo, useCallback, useRef } from "react";

// ─── Ranking Editori ─────────────────────────────────────────────────────────
// Anagrafica e sequenza editori (tabella ranking_editori): è l'ordine con cui gli editori compaiono
// nelle cedole di GiroManager e, via "Pubblica su RPN", anche su RPN.
// • Il numero di ranking è la verità: la lista è ordinata per ranking (a parità, per nome).
// • Spostare un editore (trascinamento o ↑↓) cambia SOLO il suo ranking (valore intermedio tra i vicini);
//   "Rinumera 1…N" ricompatta tutto in interi mantenendo i pari merito.
// • Nuovi editori anche se non ancora su RPN (stato "in arrivo" + data prevista).
// • Salvataggio unico via RPC ranking_editori_salva (solo amministratori PDE), che riallinea anche titoli.ranking_editore.

const SUPABASE_URL = "https://tdflwenlylhctxssatax.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InRkZmx3ZW5seWxoY3R4c3NhdGF4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYzMzgyNzYsImV4cCI6MjA5MTkxNDI3Nn0.l35qEL7LOvyYuI1McQlVqj4vbyTqmlevcmqWbTGYi2Q";

const T = {
  bg: "#1a2140", surface: "#212d54", border: "#2e3d6b", borderHi: "#3d4f82",
  text: "#f0f2f8", textMid: "#8b9cc8", textDim: "#4a5a8a",
  accent: "#7b9fe8", green: "#4caf7d", red: "#e05c5c", amber: "#e0a84c",
};
const CAT_COLORI = { A: "#7b9fe8", B: "#4caf7d", C: "#e0a84c", KIDS: "#d47bd4", SERVICE: "#8b9cc8" };
const CATEGORIE = ["A", "B", "C", "KIDS", "SERVICE"];
const PROMOZIONI = ["PDE Promozione", "PDE Service"];

const css = {
  btn: (v = "default", disabled = false) => ({ padding: "6px 12px", border: `1px solid ${v === "accent" ? T.accent : v === "green" ? T.green : v === "danger" ? T.red : T.border}`, background: v === "accent" ? T.accent : v === "green" ? T.green : "transparent", color: v === "accent" || v === "green" ? "#000" : v === "danger" ? T.red : T.text, cursor: disabled ? "default" : "pointer", fontSize: "12px", fontFamily: "inherit", borderRadius: 3, fontWeight: v === "accent" || v === "green" ? 700 : 400, opacity: disabled ? 0.5 : 1, whiteSpace: "nowrap" }),
  mini: { padding: "2px 6px", border: `1px solid ${T.border}`, background: "transparent", color: T.textMid, cursor: "pointer", fontSize: "11px", fontFamily: "inherit", borderRadius: 3 },
  card: { background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, padding: 16, marginBottom: 14 },
  th: { padding: "7px 8px", textAlign: "left", color: T.textMid, fontWeight: 400, fontSize: "10px", letterSpacing: "0.08em", textTransform: "uppercase", borderBottom: `1px solid ${T.border}`, whiteSpace: "nowrap", position: "sticky", top: 0, background: T.surface, zIndex: 1 },
  td: { padding: "5px 8px", borderBottom: `1px solid ${T.border}55`, fontSize: "12px", color: T.text, whiteSpace: "nowrap", verticalAlign: "middle" },
  input: { background: T.bg, border: `1px solid ${T.borderHi}`, color: T.text, padding: "4px 6px", fontSize: "12px", fontFamily: "inherit", borderRadius: 3 },
  label: { display: "flex", flexDirection: "column", gap: 4, color: T.textMid, fontSize: "11px" },
};

const norm = (s) => String(s ?? "").replace(/\s+/g, " ").trim().toUpperCase();
const r2 = (n) => Math.round(n * 100) / 100;
const fmtRk = (n) => (n == null || n === "" ? "—" : Number.isInteger(Number(n)) ? String(Number(n)) : String(Number(n)));
const CAMPI = ["ranking", "codice_editore", "cedola", "account_editore", "promozione", "stato_rpn", "data_ingresso_rpn", "note", "attivo", "data_uscita"];
const oggi = () => new Date().toISOString().slice(0, 10);

const ordina = (righe) => [...righe].sort((a, b) => (Number(a.ranking) - Number(b.ranking)) || a.editore_nome.localeCompare(b.editore_nome));

export default function ModuloRankingEditori({ token, titoli, onDataChange }) {
  const [orig, setOrig] = useState([]);        // righe come da DB
  const [righe, setRighe] = useState([]);      // editori attivi (la sequenza)
  const [usciti, setUsciti] = useState([]);    // editori usciti: fuori sequenza, storico intatto
  const [eliminati, setEliminati] = useState([]);
  const [loading, setLoading] = useState(true);
  const [errore, setErrore] = useState("");
  const [msg, setMsg] = useState("");
  const [saving, setSaving] = useState(false);
  const [aggTitoli, setAggTitoli] = useState(true);
  const [filtro, setFiltro] = useState("");
  const [filtroCat, setFiltroCat] = useState("");
  const [soloArrivo, setSoloArrivo] = useState(false);
  const [showDiff, setShowDiff] = useState(false);
  const [showNuovo, setShowNuovo] = useState(false);
  const [dragKey, setDragKey] = useState(null);
  const [overKey, setOverKey] = useState(null);
  const tmpId = useRef(0);

  const carica = useCallback(async () => {
    setLoading(true); setErrore("");
    try {
      const r = await fetch(`${SUPABASE_URL}/rest/v1/ranking_editori?select=*`, { headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${token}` } });
      if (!r.ok) throw new Error(`Errore caricamento (${r.status})`);
      const data = (await r.json()).map(x => ({ ...x, _key: `id${x.id}`, ranking: Number(x.ranking), editore_nome: norm(x.editore_nome), cedola: x.cedola ? norm(x.cedola) : "" }));
      data.forEach(x => { x.attivo = x.attivo !== false; });
      setOrig(data); setRighe(ordina(data.filter(x => x.attivo))); setUsciti(ordina(data.filter(x => !x.attivo))); setEliminati([]);
    } catch (e) { setErrore(e.message); }
    setLoading(false);
  }, [token]);
  useEffect(() => { carica(); }, [carica]);

  // titoli in GiroManager per editore (per sapere chi è "in uso" e non eliminabile)
  const titoliPerEditore = useMemo(() => {
    const m = {};
    (titoli || []).forEach(t => { const k = norm(t.editore_nome); if (k) m[k] = (m[k] || 0) + 1; });
    return m;
  }, [titoli]);

  const account = useMemo(() => [...new Set(orig.map(r => r.account_editore).filter(Boolean))].sort(), [orig]);
  const origById = useMemo(() => Object.fromEntries(orig.map(r => [r.id, r])), [orig]);

  // ─── Modifiche ─────────────────────────────────────────────────────────────
  const aggiorna = (key, patch) => setRighe(rs => {
    const out = rs.map(r => (r._key === key ? { ...r, ...patch } : r));
    return "ranking" in patch ? ordina(out) : out;
  });

  // Sposta la riga `key` nella posizione `idx` della lista (senza la riga stessa): assegna un ranking intermedio
  const spostaA = (key, idx) => setRighe(rs => {
    const riga = rs.find(r => r._key === key);
    const altri = rs.filter(r => r._key !== key);
    const p = altri[idx - 1]?.ranking, n = altri[idx]?.ranking;
    let nuovo;
    if (p == null && n == null) nuovo = 1;
    else if (p == null) nuovo = r2(n - 1);
    else if (n == null) nuovo = Math.floor(p) + 1;
    else if (p === n) nuovo = p;                         // dentro un gruppo di pari merito: si unisce al gruppo
    else nuovo = r2((p + n) / 2);
    if (p != null && n != null && p !== n && (nuovo <= p || nuovo >= n)) {
      // spazio esaurito tra i due valori: rinumero tutto e riprovo
      const rin = rinumeraLista([...altri.slice(0, idx), riga, ...altri.slice(idx)]);
      return rin;
    }
    return ordina(altri.concat({ ...riga, ranking: nuovo }).map(r => r));
  });

  const suGiu = (key, dir) => {
    const i = righe.findIndex(r => r._key === key);
    const altri = righe.filter(r => r._key !== key);
    // salta i pari merito: sale sopra il gruppo precedente / scende sotto il gruppo successivo
    const rk = righe[i].ranking;
    if (dir < 0) {
      let j = i - 1;
      while (j >= 0 && righe[j].ranking === rk) j--;
      if (j < 0) return;
      const target = righe[j].ranking;
      let k = j; while (k > 0 && righe[k - 1].ranking === target) k--;
      spostaA(key, altri.findIndex(r => r._key === righe[k]._key));
    } else {
      let j = i + 1;
      while (j < righe.length && righe[j].ranking === rk) j++;
      if (j >= righe.length) return;
      const target = righe[j].ranking;
      let k = j; while (k < righe.length - 1 && righe[k + 1].ranking === target) k++;
      spostaA(key, altri.findIndex(r => r._key === righe[k]._key) + 1);
    }
  };

  const pariAlPrecedente = (key) => {
    const i = righe.findIndex(r => r._key === key);
    if (i <= 0) return;
    if (righe[i].ranking === righe[i - 1].ranking) {          // scollega: va subito dopo il gruppo
      const altri = righe.filter(r => r._key !== key);
      let j = i - 1; const rk = righe[i].ranking;
      while (j < altri.length && altri[j]?.ranking === rk) j++;
      spostaA(key, j);
    } else aggiorna(key, { ranking: righe[i - 1].ranking });
  };

  function rinumeraLista(lista) {
    let n = 0, prev = null;
    return lista.map(r => { if (prev === null || r.ranking !== prev) n++; prev = r.ranking; return { ...r, ranking: n }; });
  }
  const rinumera = () => setRighe(rs => rinumeraLista(rs));

  const elimina = (r) => {
    if (r.id) setEliminati(e => [...e, r.id]);
    setRighe(rs => rs.filter(x => x._key !== r._key));
    setUsciti(us => us.filter(x => x._key !== r._key));
  };
  // Uscita: l'editore esce dalla sequenza ma resta in anagrafica (titoli, prenotato e Fine Giro storici intatti)
  const esci = (r) => {
    setRighe(rs => rs.filter(x => x._key !== r._key));
    setUsciti(us => ordina([...us, { ...r, attivo: false, data_uscita: oggi() }]));
  };
  const riattiva = (r) => {
    setUsciti(us => us.filter(x => x._key !== r._key));
    setRighe(rs => ordina([...rs, { ...r, attivo: true, data_uscita: "" }]));
  };

  // ─── Differenze da salvare ─────────────────────────────────────────────────
  const diff = useMemo(() => {
    const tutte = [...righe, ...usciti];
    const nuovi = tutte.filter(r => !r.id);
    const modificati = tutte.filter(r => r.id && CAMPI.some(c => String(r[c] ?? "") !== String(origById[r.id]?.[c] ?? "")));
    return { nuovi, modificati, eliminati: eliminati.map(id => origById[id]).filter(Boolean) };
  }, [righe, usciti, origById, eliminati]);
  const nModifiche = diff.nuovi.length + diff.modificati.length + diff.eliminati.length;

  const salva = async () => {
    setSaving(true); setErrore(""); setMsg("");
    try {
      const payload = [...diff.nuovi, ...diff.modificati].map(r => {
        const o = { id: r.id ?? null, editore_nome: r.editore_nome };
        CAMPI.forEach(c => { o[c] = r[c] ?? ""; });
        return o;
      });
      const res = await fetch(`${SUPABASE_URL}/rest/v1/rpc/ranking_editori_salva`, {
        method: "POST",
        headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({ p_righe: payload, p_elimina: eliminati, p_aggiorna_titoli: aggTitoli }),
      });
      const j = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(j.message || j.hint || `Errore ${res.status}`);
      // editori nuovi "in arrivo" con codice: li registro anche in Editori New Entry (prenotazioni BookUp)
      const perNewEntry = diff.nuovi.filter(r => r._newEntry && r.codice_editore);
      for (const r of perNewEntry) {
        await fetch(`${SUPABASE_URL}/rest/v1/editori_new_entry?on_conflict=codice_editore`, {
          method: "POST",
          headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${token}`, "Content-Type": "application/json", Prefer: "resolution=merge-duplicates,return=minimal" },
          body: JSON.stringify({ codice_editore: r.codice_editore, nome_editore: r.editore_nome, attivo: true, data_ingresso_rpn: r.data_ingresso_rpn || null }),
        }).catch(() => {});
      }
      setMsg(`Salvato: ${j.inseriti} nuovi · ${j.aggiornati} modificati · ${j.eliminati} eliminati${aggTitoli ? ` · ranking aggiornato su ${j.titoli_aggiornati} titoli` : ""}. Per le cedole già su RPN usa "Pubblica su RPN → Aggiorna" per riordinarle.`);
      setShowDiff(false);
      await carica();
      onDataChange && onDataChange();
    } catch (e) { setErrore(e.message); }
    setSaving(false);
  };

  const esporta = () => {
    const XLSX = window.XLSX;
    if (!XLSX) return;
    const rows = [...righe, ...usciti].map(r => ({ Stato: r.attivo === false ? `Uscito${r.data_uscita ? " " + r.data_uscita : ""}` : "Attivo", Ranking: r.ranking, Editore: r.editore_nome, Codice: r.codice_editore ?? "", Cedola: r.cedola, Account: r.account_editore ?? "", Promozione: r.promozione ?? "", "Stato RPN": r.stato_rpn === "in_arrivo" ? "In arrivo" : "Attivo", "Ingresso RPN": r.data_ingresso_rpn ?? "", "Titoli GM": titoliPerEditore[r.editore_nome] || 0, Note: r.note ?? "" }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "Ranking editori");
    XLSX.writeFile(wb, `ranking_editori_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  // ─── Vista ─────────────────────────────────────────────────────────────────
  const conteggi = useMemo(() => {
    const c = {}; righe.forEach(r => { const k = r.cedola || "—"; c[k] = (c[k] || 0) + 1; }); return c;
  }, [righe]);
  const filtrate = righe.filter(r =>
    (!filtro || r.editore_nome.includes(norm(filtro)) || String(r.codice_editore ?? "").includes(filtro.trim())) &&
    (!filtroCat || r.cedola === filtroCat) && (!soloArrivo || r.stato_rpn === "in_arrivo"));
  const filtroAttivo = !!(filtro || filtroCat || soloArrivo);

  // un editore è "fuori blocco" se la sua categoria è diversa sia dal precedente che dal successivo
  const fuoriBlocco = (i) => {
    const r = righe[i], p = righe[i - 1], n = righe[i + 1];
    return r.cedola && ((p && p.cedola !== r.cedola) && (!n || n.cedola !== r.cedola)) && (p || n) && righe.some((x, j) => j !== i && x.cedola === r.cedola);
  };

  if (loading) return <div style={{ padding: 24, color: T.textMid }}>Caricamento anagrafica editori…</div>;

  return (
    <div style={{ flex: 1, overflowY: "auto", padding: 24 }}>
      <div style={css.card}>
        <div style={{ display: "flex", alignItems: "flex-start", gap: 16, flexWrap: "wrap" }}>
          <div style={{ flex: "1 1 420px" }}>
            <div style={{ color: T.text, fontWeight: 700, fontSize: "13px", marginBottom: 6 }}>🏷 Ranking editori</div>
            <div style={{ color: T.textMid, fontSize: "11px", lineHeight: 1.6 }}>
              La sequenza decide l'ordine degli editori nelle cedole (poi, dentro ogni editore, conta la posizione del titolo).
              Trascina una riga o usa ↑↓ per spostare un editore: cambia solo il suo numero. <b style={{ color: T.text }}>=</b> lo mette a pari merito con quello sopra.
              "Rinumera 1…N" ricompatta la sequenza in numeri interi. Se un editore esce, usa <b style={{ color: T.text }}>esce</b>: lascia la sequenza ma lo storico resta. Nulla viene salvato finché non premi <b style={{ color: T.text }}>Salva</b>.
            </div>
          </div>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
            {CATEGORIE.map(c => (
              <div key={c} style={{ border: `1px solid ${CAT_COLORI[c]}66`, borderRadius: 4, padding: "6px 10px", minWidth: 60 }}>
                <div style={{ color: CAT_COLORI[c], fontSize: "10px", fontWeight: 700 }}>{c}</div>
                <div style={{ color: T.text, fontWeight: 700 }}>{conteggi[c] || 0}</div>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* Barra strumenti */}
      <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", marginBottom: 12 }}>
        <input style={{ ...css.input, width: 200, padding: "6px 8px" }} placeholder="Cerca editore o codice…" value={filtro} onChange={e => setFiltro(e.target.value)} />
        <select style={{ ...css.input, padding: "6px 8px" }} value={filtroCat} onChange={e => setFiltroCat(e.target.value)}>
          <option value="">Tutte le cedole</option>
          {CATEGORIE.map(c => <option key={c}>{c}</option>)}
        </select>
        <label style={{ color: T.textMid, fontSize: "12px", display: "flex", gap: 4, alignItems: "center" }}><input type="checkbox" checked={soloArrivo} onChange={e => setSoloArrivo(e.target.checked)} /> solo in arrivo su RPN</label>
        <div style={{ flex: 1 }} />
        <button style={css.btn("green")} onClick={() => setShowNuovo(s => !s)}>+ Nuovo editore</button>
        <button style={css.btn()} onClick={rinumera} title="Assegna 1, 2, 3… nell'ordine attuale, mantenendo i pari merito">Rinumera 1…N</button>
        <button style={css.btn()} onClick={esporta}>Esporta Excel</button>
      </div>

      {showNuovo && <NuovoEditore righe={righe} usciti={usciti} account={account} onAnnulla={() => setShowNuovo(false)} onAggiungi={(r) => {
        tmpId.current += 1;
        setRighe(rs => ordina([...rs, { ...r, id: null, _key: `new${tmpId.current}` }]));
        setShowNuovo(false);
      }} />}

      {errore && <div style={{ color: T.red, fontSize: "12px", marginBottom: 10 }}>⚠ {errore}</div>}
      {msg && <div style={{ color: T.green, fontSize: "12px", marginBottom: 10 }}>✓ {msg}</div>}

      {/* Barra salvataggio */}
      {nModifiche > 0 && (
        <div style={{ ...css.card, borderColor: T.amber, display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap", position: "sticky", top: 0, zIndex: 3 }}>
          <span style={{ color: T.amber, fontWeight: 700, fontSize: "12px" }}>{nModifiche} modifiche non salvate</span>
          <span style={{ color: T.textMid, fontSize: "11px" }}>{diff.nuovi.length} nuovi · {diff.modificati.length} modificati · {diff.eliminati.length} eliminati</span>
          <button style={css.mini} onClick={() => setShowDiff(s => !s)}>{showDiff ? "Nascondi" : "Vedi"} dettaglio</button>
          <label style={{ color: T.textMid, fontSize: "11px", display: "flex", gap: 4, alignItems: "center" }} title="Allinea il campo ranking_editore dei titoli già caricati">
            <input type="checkbox" checked={aggTitoli} onChange={e => setAggTitoli(e.target.checked)} /> aggiorna anche i titoli in GiroManager
          </label>
          <div style={{ flex: 1 }} />
          <button style={css.btn()} onClick={() => { setRighe(ordina(orig.filter(x => x.attivo))); setUsciti(ordina(orig.filter(x => !x.attivo))); setEliminati([]); }}>Annulla</button>
          <button style={css.btn("accent", saving)} disabled={saving} onClick={salva}>{saving ? "Salvataggio…" : "Salva"}</button>
          {showDiff && (
            <div style={{ width: "100%", fontSize: "11px", color: T.textMid, maxHeight: 180, overflow: "auto" }}>
              {diff.nuovi.map(r => <div key={r._key} style={{ color: T.green }}>+ {r.editore_nome} (rk {fmtRk(r.ranking)}, {r.cedola}{r.stato_rpn === "in_arrivo" ? ", in arrivo" : ""})</div>)}
              {diff.modificati.map(r => {
                const o = origById[r.id];
                const cambi = CAMPI.filter(c => String(r[c] ?? "") !== String(o[c] ?? ""));
                if (cambi.includes("attivo")) return <div key={r._key} style={{ color: r.attivo ? T.green : T.amber }}>{r.attivo ? "↺ RIATTIVATO" : "⏏ USCITO"} {r.editore_nome}{!r.attivo && r.data_uscita ? ` (dal ${r.data_uscita})` : ""}</div>;
                return <div key={r._key}>✎ {r.editore_nome}: {cambi.map(c => `${c.replace("_editore", "")} ${o[c] ?? "—"} → ${r[c] || "—"}`).join(" · ")}</div>;
              })}
              {diff.eliminati.map(r => <div key={r.id} style={{ color: T.red }}>− {r.editore_nome}</div>)}
            </div>
          )}
        </div>
      )}

      {filtroAttivo && <div style={{ color: T.textDim, fontSize: "11px", marginBottom: 6 }}>Filtro attivo: il trascinamento è disattivato (usa ↑↓ o modifica il numero).</div>}

      <div style={{ border: `1px solid ${T.border}`, borderRadius: 4, overflow: "auto", maxHeight: "70vh" }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr>{["", "Rk", "", "Editore", "Codice", "Cedola", "Account", "Promozione", "RPN", "Titoli GM", ""].map((h, i) => <th key={i} style={css.th}>{h}</th>)}</tr></thead>
          <tbody>
            {filtrate.map((r) => {
              const i = righe.findIndex(x => x._key === r._key);
              const o = r.id ? origById[r.id] : null;
              const rkCambiato = o && Number(o.ranking) !== Number(r.ranking);
              const pari = i > 0 && righe[i - 1].ranking === r.ranking;
              const nuovaCat = i === 0 || righe[i - 1].cedola !== r.cedola;
              const nTit = titoliPerEditore[r.editore_nome] || 0;
              const fb = fuoriBlocco(i);
              return [
                nuovaCat && !filtroAttivo && (
                  <tr key={`cat-${r._key}`}><td colSpan={11} style={{ padding: "8px 8px 4px", color: CAT_COLORI[r.cedola] || T.textMid, fontSize: "11px", fontWeight: 700, letterSpacing: "0.1em", borderBottom: `1px solid ${CAT_COLORI[r.cedola] || T.border}55` }}>CEDOLA {r.cedola || "SENZA CATEGORIA"}</td></tr>
                ),
                <tr key={r._key}
                  draggable={!filtroAttivo}
                  onDragStart={() => setDragKey(r._key)}
                  onDragEnd={() => { setDragKey(null); setOverKey(null); }}
                  onDragOver={e => { if (dragKey) { e.preventDefault(); setOverKey(r._key); } }}
                  onDrop={e => { e.preventDefault(); if (dragKey && dragKey !== r._key) { const altri = righe.filter(x => x._key !== dragKey); spostaA(dragKey, altri.findIndex(x => x._key === r._key)); } setDragKey(null); setOverKey(null); }}
                  style={{ background: dragKey === r._key ? T.accent + "22" : !r.id ? T.green + "14" : "transparent", borderTop: overKey === r._key && dragKey !== r._key ? `2px solid ${T.accent}` : undefined, opacity: dragKey === r._key ? 0.5 : 1 }}>
                  <td style={{ ...css.td, color: T.textDim, cursor: filtroAttivo ? "default" : "grab", width: 14 }}>{filtroAttivo ? "" : "⋮⋮"}</td>
                  <td style={css.td}>
                    <input key={`${r._key}-${r.ranking}`} type="number" step="0.1" style={{ ...css.input, width: 58, borderColor: rkCambiato ? T.amber : T.borderHi, color: CAT_COLORI[r.cedola] || T.text, fontWeight: 700 }}
                      defaultValue={r.ranking} title="Modifica il numero e premi Invio (o esci dal campo)"
                      onBlur={e => { const v = Number(e.target.value); if (e.target.value !== "" && !isNaN(v) && v !== r.ranking) aggiorna(r._key, { ranking: v }); }}
                      onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }} />
                    {rkCambiato && <span style={{ color: T.textDim, fontSize: "10px", marginLeft: 4 }}>era {fmtRk(o.ranking)}</span>}
                  </td>
                  <td style={{ ...css.td, whiteSpace: "nowrap" }}>
                    <button style={css.mini} title="Su" onClick={() => suGiu(r._key, -1)}>↑</button>{" "}
                    <button style={css.mini} title="Giù" onClick={() => suGiu(r._key, 1)}>↓</button>{" "}
                    <button style={{ ...css.mini, color: pari ? T.accent : T.textDim, borderColor: pari ? T.accent : T.border }} title={pari ? "Pari merito con l'editore sopra: clicca per separarlo" : "Metti a pari merito con l'editore sopra"} disabled={i === 0} onClick={() => pariAlPrecedente(r._key)}>=</button>
                  </td>
                  <td style={{ ...css.td, fontWeight: 600 }}>
                    {r.editore_nome}
                    {!r.id && <span style={{ color: T.green, fontSize: "10px", marginLeft: 6 }}>NUOVO</span>}
                    {fb && <span style={{ color: T.amber, fontSize: "10px", marginLeft: 6 }} title="La sua categoria è diversa da quella degli editori vicini">⚠ fuori blocco</span>}
                    {r.note && <span style={{ color: T.textDim, fontSize: "10px", marginLeft: 6 }} title={r.note}>✎</span>}
                  </td>
                  <td style={css.td}><input style={{ ...css.input, width: 56 }} value={r.codice_editore ?? ""} placeholder="—" onChange={e => aggiorna(r._key, { codice_editore: e.target.value.trim() })} /></td>
                  <td style={css.td}>
                    <select style={{ ...css.input, color: CAT_COLORI[r.cedola] || T.text }} value={r.cedola || ""} onChange={e => aggiorna(r._key, { cedola: e.target.value })}>
                      <option value="">—</option>{CATEGORIE.map(c => <option key={c}>{c}</option>)}
                    </select>
                  </td>
                  <td style={css.td}>
                    <select style={css.input} value={r.account_editore ?? ""} onChange={e => aggiorna(r._key, { account_editore: e.target.value })}>
                      <option value="">—</option>{[...new Set([...account, r.account_editore].filter(Boolean))].map(a => <option key={a}>{a}</option>)}
                    </select>
                  </td>
                  <td style={css.td}>
                    <select style={css.input} value={r.promozione ?? ""} onChange={e => aggiorna(r._key, { promozione: e.target.value })}>
                      <option value="">—</option>{PROMOZIONI.map(p => <option key={p}>{p}</option>)}
                    </select>
                  </td>
                  <td style={css.td}>
                    {r.stato_rpn === "in_arrivo" ? (
                      <span style={{ display: "flex", gap: 4, alignItems: "center" }}>
                        <button style={{ ...css.mini, color: T.amber, borderColor: T.amber }} title="In arrivo su RPN: clicca per segnarlo attivo" onClick={() => aggiorna(r._key, { stato_rpn: "attivo" })}>in arrivo</button>
                        <input type="date" style={{ ...css.input, width: 118 }} value={r.data_ingresso_rpn ?? ""} onChange={e => aggiorna(r._key, { data_ingresso_rpn: e.target.value })} />
                      </span>
                    ) : (
                      <button style={{ ...css.mini, color: T.green, borderColor: T.green + "66" }} title="Attivo su RPN: clicca per segnarlo in arrivo" onClick={() => aggiorna(r._key, { stato_rpn: "in_arrivo" })}>attivo</button>
                    )}
                  </td>
                  <td style={{ ...css.td, color: nTit ? T.textMid : T.textDim }}>{nTit || "—"}</td>
                  <td style={css.td}>
                    <button style={{ ...css.mini, color: T.amber }} title="Editore uscito: esce dalla sequenza, lo storico resta" onClick={() => esci(r)}>esce</button>{" "}
                    {nTit === 0 && <button style={{ ...css.mini, color: T.red }} title="Elimina (solo editori senza titoli)" onClick={() => elimina(r)}>✕</button>}
                  </td>
                </tr>,
              ];
            })}
          </tbody>
        </table>
      </div>

      {usciti.length > 0 && (
        <div style={{ ...css.card, marginTop: 14, opacity: 0.85 }}>
          <div style={{ color: T.textMid, fontWeight: 700, fontSize: "11px", letterSpacing: "0.1em", marginBottom: 8 }}>EDITORI USCITI ({usciti.length}) — fuori dalla sequenza, esclusi dagli abbinamenti automatici dell'import; titoli, prenotato e Fine Giro storici restano intatti</div>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr>{["Ultimo rk", "Editore", "Codice", "Cedola", "Account", "Uscito il", "Titoli GM", ""].map(h => <th key={h} style={{ ...css.th, position: "static" }}>{h}</th>)}</tr></thead>
            <tbody>
              {usciti.map(r => {
                const nTit = titoliPerEditore[r.editore_nome] || 0;
                return (
                  <tr key={r._key} style={{ color: T.textMid }}>
                    <td style={{ ...css.td, color: T.textDim }}>{fmtRk(r.ranking)}</td>
                    <td style={{ ...css.td, color: T.textMid, textDecoration: "line-through" }}>{r.editore_nome}</td>
                    <td style={{ ...css.td, color: T.textDim }}>{r.codice_editore || "—"}</td>
                    <td style={{ ...css.td, color: T.textDim }}>{r.cedola || "—"}</td>
                    <td style={{ ...css.td, color: T.textDim }}>{r.account_editore || "—"}</td>
                    <td style={css.td}><input type="date" style={css.input} value={r.data_uscita ?? ""} onChange={e => setUsciti(us => us.map(x => x._key === r._key ? { ...x, data_uscita: e.target.value } : x))} /></td>
                    <td style={{ ...css.td, color: T.textDim }}>{nTit || "—"}</td>
                    <td style={css.td}>
                      <button style={{ ...css.mini, color: T.green }} title="Rientra nella sequenza con il suo ultimo ranking" onClick={() => riattiva(r)}>riattiva</button>{" "}
                      {nTit === 0 && <button style={{ ...css.mini, color: T.red }} title="Elimina (solo editori senza titoli)" onClick={() => elimina(r)}>✕</button>}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ─── Form nuovo editore ──────────────────────────────────────────────────────
function NuovoEditore({ righe, usciti = [], account, onAggiungi, onAnnulla }) {
  const [f, setF] = useState({ editore_nome: "", codice_editore: "", cedola: "C", account_editore: "", promozione: "PDE Promozione", stato_rpn: "in_arrivo", data_ingresso_rpn: `${new Date().getFullYear() + 1}-01-01`, note: "", dopo: "__fine_cat__", _newEntry: true });
  const set = (k, v) => setF(x => ({ ...x, [k]: v }));
  const nome = norm(f.editore_nome);
  const simili = nome.length >= 3 ? righe.filter(r => r.editore_nome.includes(nome) || nome.includes(r.editore_nome)).slice(0, 4) : [];
  const esatto = righe.some(r => r.editore_nome === nome);
  const esattoUscito = usciti.some(r => r.editore_nome === nome);

  // ranking proposto: in fondo alla categoria scelta, oppure subito dopo l'editore indicato
  const rankingProposto = () => {
    const lista = ordina(righe);
    let idx;
    if (f.dopo === "__fine_cat__") {
      const ultimi = lista.map((r, i) => [r, i]).filter(([r]) => r.cedola === f.cedola);
      if (ultimi.length) idx = ultimi[ultimi.length - 1][1];
      else {
        // categoria vuota: dopo l'ultimo editore della categoria precedente nell'ordine A, B, C, KIDS, SERVICE
        const ci = CATEGORIE.indexOf(f.cedola);
        const prima = lista.map((r, i) => [r, i]).filter(([r]) => CATEGORIE.indexOf(r.cedola) >= 0 && CATEGORIE.indexOf(r.cedola) < ci);
        idx = prima.length ? prima[prima.length - 1][1] : -1;
      }
    } else idx = lista.findIndex(r => r._key === f.dopo);
    const p = lista[idx]?.ranking, n = lista.slice(idx + 1).find(r => r.ranking !== p)?.ranking;
    if (p == null) return n != null ? r2(n - 1) : 1;
    if (n == null) return Math.floor(p) + 1;
    return r2((p + n) / 2);
  };

  const ok = nome && !esatto && !esattoUscito && f.cedola;
  return (
    <div style={{ ...css.card, borderColor: T.green }}>
      <div style={{ color: T.green, fontWeight: 700, fontSize: "12px", marginBottom: 10 }}>+ Nuovo editore</div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(170px, 1fr))", gap: 10 }}>
        <label style={{ ...css.label, gridColumn: "span 2" }}>Nome editore *
          <input style={css.input} value={f.editore_nome} onChange={e => set("editore_nome", e.target.value)} placeholder="es. NN EDITORE" autoFocus />
        </label>
        <label style={css.label}>Codice editore (MELI)
          <input style={css.input} value={f.codice_editore} onChange={e => set("codice_editore", e.target.value)} placeholder="anche vuoto" />
        </label>
        <label style={css.label}>Cedola *
          <select style={css.input} value={f.cedola} onChange={e => setF(x => ({ ...x, cedola: e.target.value, dopo: "__fine_cat__" }))}>{CATEGORIE.map(c => <option key={c}>{c}</option>)}</select>
        </label>
        <label style={{ ...css.label, gridColumn: "span 2" }}>Posizione nella sequenza
          <select style={css.input} value={f.dopo} onChange={e => set("dopo", e.target.value)}>
            <option value="__fine_cat__">In fondo alla cedola {f.cedola}</option>
            {ordina(righe).filter(r => r.cedola === f.cedola).map(r => <option key={r._key} value={r._key}>Subito dopo {r.editore_nome} (rk {fmtRk(r.ranking)})</option>)}
          </select>
        </label>
        <label style={css.label}>Account
          <input style={css.input} list="gm-account-list" value={f.account_editore} onChange={e => set("account_editore", e.target.value.toUpperCase())} />
          <datalist id="gm-account-list">{account.map(a => <option key={a} value={a} />)}</datalist>
        </label>
        <label style={css.label}>Promozione
          <select style={css.input} value={f.promozione} onChange={e => set("promozione", e.target.value)}>{PROMOZIONI.map(p => <option key={p}>{p}</option>)}</select>
        </label>
        <label style={css.label}>Stato su RPN
          <select style={css.input} value={f.stato_rpn} onChange={e => set("stato_rpn", e.target.value)}>
            <option value="in_arrivo">In arrivo (non ancora su RPN)</option>
            <option value="attivo">Già attivo su RPN</option>
          </select>
        </label>
        {f.stato_rpn === "in_arrivo" && (
          <label style={css.label}>Ingresso previsto su RPN
            <input type="date" style={css.input} value={f.data_ingresso_rpn} onChange={e => set("data_ingresso_rpn", e.target.value)} />
          </label>
        )}
        <label style={{ ...css.label, gridColumn: "span 2" }}>Note
          <input style={css.input} value={f.note} onChange={e => set("note", e.target.value)} placeholder="es. distribuzione da gennaio, marchi collegati…" />
        </label>
      </div>
      {f.stato_rpn === "in_arrivo" && (
        <label style={{ color: T.textMid, fontSize: "11px", display: "flex", gap: 6, alignItems: "center", marginTop: 10 }}>
          <input type="checkbox" checked={f._newEntry} onChange={e => set("_newEntry", e.target.checked)} disabled={!f.codice_editore.trim()} />
          aggiungilo anche a "Editori New Entry" (raccolta prenotazioni in BookUp){!f.codice_editore.trim() && " — serve il codice editore"}
        </label>
      )}
      {esatto && <div style={{ color: T.red, fontSize: "11px", marginTop: 8 }}>Esiste già un editore con questo nome.</div>}
      {esattoUscito && <div style={{ color: T.amber, fontSize: "11px", marginTop: 8 }}>Questo editore è tra gli usciti: usa "riattiva" in fondo alla pagina invece di crearne uno nuovo.</div>}
      {!esatto && simili.length > 0 && <div style={{ color: T.amber, fontSize: "11px", marginTop: 8 }}>Attenzione, nomi simili già presenti: {simili.map(s => s.editore_nome).join(", ")}</div>}
      <div style={{ display: "flex", gap: 8, marginTop: 12, alignItems: "center" }}>
        <button style={css.btn("green", !ok)} disabled={!ok} onClick={() => onAggiungi({
          editore_nome: nome, codice_editore: f.codice_editore.trim(), cedola: f.cedola, account_editore: f.account_editore.trim(), promozione: f.promozione,
          stato_rpn: f.stato_rpn, data_ingresso_rpn: f.stato_rpn === "in_arrivo" ? f.data_ingresso_rpn : "", note: f.note.trim(), ranking: rankingProposto(), _newEntry: f._newEntry && f.stato_rpn === "in_arrivo",
        })}>Aggiungi alla sequenza</button>
        <button style={css.btn()} onClick={onAnnulla}>Annulla</button>
        {ok && <span style={{ color: T.textMid, fontSize: "11px" }}>Ranking proposto: <b style={{ color: T.text }}>{fmtRk(rankingProposto())}</b> — poi premi Salva</span>}
      </div>
    </div>
  );
}
