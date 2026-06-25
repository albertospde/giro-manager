import { useState, useCallback } from "react";

const SUPABASE_URL = "https://tdflwenlylhctxssatax.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InRkZmx3ZW5seWxoY3R4c3NhdGF4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYzMzgyNzYsImV4cCI6MjA5MTkxNDI3Nn0.l35qEL7LOvyYuI1McQlVqj4vbyTqmlevcmqWbTGYi2Q";

const T = {
  bg: "#0f0f0f", surface: "#161616", border: "#252525", borderHi: "#333333",
  text: "#e8e8e8", textMid: "#888888", textDim: "#444444",
  accent: "#c8a96e", green: "#4caf7d", red: "#e05c5c", blue: "#5b8fd4",
};

const css = {
  btn: (v = "default") => ({ padding: "6px 14px", border: `1px solid ${v === "accent" ? T.accent : v === "danger" ? T.red : T.border}`, background: v === "accent" ? T.accent : v === "danger" ? T.red + "22" : "transparent", color: v === "accent" ? "#000" : v === "danger" ? T.red : T.text, cursor: "pointer", fontSize: "12px", fontFamily: "inherit", borderRadius: 3, fontWeight: v === "accent" ? "700" : "400" }),
  input: { background: T.bg, border: `1px solid ${T.border}`, color: T.text, padding: "5px 10px", fontSize: "12px", fontFamily: "inherit", borderRadius: 3, outline: "none" },
  th: { padding: "8px 12px", textAlign: "left", color: T.textMid, fontWeight: "400", fontSize: "11px", letterSpacing: "0.08em", textTransform: "uppercase", borderBottom: `1px solid ${T.border}`, whiteSpace: "nowrap", background: T.surface },
  td: { padding: "7px 12px", borderBottom: `1px solid ${T.border}22`, verticalAlign: "middle", fontSize: "12px" },
};

// Template ridotto — solo i campi essenziali del carico semplice.
// Colonne nel template "CARICO":
// 0=giro_label, 1=ean, 2=titolo, 3=autore, 4=editore_nome, 5=prezzo, 6=uscita,
// 7=obiettivo_assegnato, 8=note, 9=ean_gemello_1, 10=titolo_gemello_1,
// 11=ean_gemello_2, 12=titolo_gemello_2, 13=ean_gemello_3, 14=titolo_gemello_3
const COL_MAP = {
  0: "giro_label", 1: "ean", 2: "titolo", 3: "autore", 4: "editore_nome",
  5: "prezzo", 6: "uscita", 7: "obiettivo_assegnato", 8: "note",
  9: "ean_gemello_1", 10: "titolo_gemello_1",
  11: "ean_gemello_2", 12: "titolo_gemello_2",
  13: "ean_gemello_3", 14: "titolo_gemello_3",
};

const REQUIRED = ["giro_label", "ean", "titolo", "editore_nome", "prezzo"];
const FORMATO_DEFAULT = "Cover";

// ─── Fetch anagrafica editori (codice, ranking, account, promozione) da ranking_editori ──
async function fetchAnagraficaEditori(token) {
  const res = await fetch(
    `${SUPABASE_URL}/rest/v1/ranking_editori?select=editore_nome,codice_editore,ranking,account_editore,promozione`,
    { headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}` } }
  );
  if (!res.ok) {
    console.error("Fetch ranking_editori fallita:", res.status, res.statusText);
    throw new Error(`Errore caricamento anagrafica editori (HTTP ${res.status})`);
  }
  const data = await res.json();
  const map = {};
  data.forEach(({ editore_nome, codice_editore, ranking, account_editore, promozione }) => {
    const nome = String(editore_nome ?? "").trim().toUpperCase();
    if (!nome) return;
    // In caso di doppioni sullo stesso nome, tiene il ranking più basso (più prioritario)
    if (!(nome in map) || ranking < map[nome].ranking) {
      map[nome] = { codice_editore, ranking, account_editore, promozione };
    }
  });
  return map;
}

// ─── Componente principale ───────────────────────────────────────────────────
// titoliEsistenti: l'array `titoli` già caricato in GiroManager, serve per calcolare
// la posizione di partenza per ogni gruppo giro+editore (continua da dove sono arrivati).
export default function ModuloCaricoSemplice({ giriList, titoliEsistenti, token, onClose, onImportDone }) {
  const [giroSel, setGiroSel] = useState(giriList[0]?.id ?? null);
  const [file, setFile] = useState(null);
  const [rows, setRows] = useState([]);
  const [errors, setErrors] = useState([]);
  const [importing, setImporting] = useState(false);
  const [done, setDone] = useState(null);
  const [step, setStep] = useState("upload"); // upload | preview | result
  const [anagraficaMissing, setAnagraficaMissing] = useState([]); // editori non trovati in ranking_editori

  const handleFile = useCallback(async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    setFile(f);

    let anagraficaMap = {};
    try {
      anagraficaMap = await fetchAnagraficaEditori(token);
    } catch (err) {
      alert("Impossibile caricare l'anagrafica editori dal database: " + err.message + "\nRiprova; se l'errore persiste, controlla la sessione/login.");
      return;
    }

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const XLSX = window.XLSX;
        const wb = XLSX.read(evt.target.result, { type: "array" });
        const ws = wb.Sheets["CARICO"] || wb.Sheets[wb.SheetNames[0]];
        if (!ws) { alert("Foglio 'CARICO' non trovato. Stai usando il template corretto?"); return; }
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        // Skip righe 0-2 (titolo, legenda, vuota), header riga 3, dati da riga 4 (0-indexed)
        const dataRows = data.slice(4).filter(r => r.some(v => v !== ""));

        const parsed = [];
        const errs = [];
        const missing = new Set();
        // Contatori di posizione per gruppo "giro_label|editore_nome", inizializzati al
        // massimo già presente in DB per quel gruppo (così si continua la numerazione).
        const contatori = {};

        dataRows.forEach((row, idx) => {
          const obj = {};
          Object.entries(COL_MAP).forEach(([colIdx, field]) => {
            let val = row[parseInt(colIdx)] ?? "";
            if (field === "prezzo") val = parseFloat(String(val).replace(",", ".")) || null;
            if (field === "obiettivo_assegnato") val = parseInt(val) || 0;
            if (field === "giro_label") val = val ? String(val) : null;
            const CAMPI_TESTO = ["titolo","autore","editore_nome","uscita","note","giro_label","titolo_gemello_1","titolo_gemello_2","titolo_gemello_3"];
            if (CAMPI_TESTO.includes(field) && typeof val === "string" && val !== "") val = val.trim().toUpperCase();
            obj[field] = val === "" ? null : val;
          });

          obj.giro_id = null;        // assegnato all'import dal selettore Giro
          obj.n_cedola = null;       // non usato nel carico semplice
          obj.formato = FORMATO_DEFAULT;

          // Risolvi codice_editore / ranking_editore / account_editore / promozione dal DB
          const nome = String(obj.editore_nome ?? "").trim().toUpperCase();
          const anagrafica = nome ? anagraficaMap[nome] : null;
          if (anagrafica) {
            obj.codice_editore = anagrafica.codice_editore ?? null;
            obj.ranking_editore = anagrafica.ranking ?? null;
            obj.account_editore = anagrafica.account_editore ?? null;
            obj.promozione = anagrafica.promozione ?? null;
          } else {
            obj.codice_editore = null;
            obj.ranking_editore = null;
            obj.account_editore = null;
            obj.promozione = null;
            if (nome) missing.add(nome);
          }

          // Posizione: sequenziale nell'ordine esatto delle righe del file, per gruppo giro+editore.
          // Continua dal massimo già presente in DB (titoliEsistenti) per lo stesso giro_label+editore.
          const keyGruppo = `${obj.giro_label ?? ""}|${nome}`;
          if (!(keyGruppo in contatori)) {
            const maxEsistente = (titoliEsistenti || [])
              .filter(t => t.giro_label === obj.giro_label && String(t.editore_nome ?? "").trim().toUpperCase() === nome)
              .reduce((m, t) => Math.max(m, t.posizione || 0), 0);
            contatori[keyGruppo] = maxEsistente;
          }
          contatori[keyGruppo] += 1;
          obj.posizione = contatori[keyGruppo];

          // Validazione campi obbligatori
          const rowErrs = [];
          REQUIRED.forEach(f => { if (!obj[f]) rowErrs.push(f); });
          if (obj.ean && String(obj.ean).replace(/\D/g,"").length !== 13) rowErrs.push("EAN non valido");
          if (rowErrs.length) errs.push({ row: idx + 5, fields: rowErrs });

          parsed.push(obj);
        });

        setRows(parsed);
        setErrors(errs);
        setAnagraficaMissing([...missing]);
        setStep("preview");
      } catch (err) {
        alert("Errore lettura file: " + err.message);
      }
    };
    reader.readAsArrayBuffer(f);
  }, [token, titoliEsistenti]);

  const handleImport = async () => {
    if (!giroSel) return;
    setImporting(true);
    const payload = rows.map(r => ({ ...r, giro_id: giroSel }));
    try {
      const res = await fetch(`${SUPABASE_URL}/rest/v1/rpc/upsert_titoli`, {
        method: "POST",
        headers: { "Content-Type": "application/json", "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}` },
        body: JSON.stringify({ payload }),
      });
      if (res.ok) {
        setDone({ ok: rows.length, err: 0 });
        setStep("result");
        onImportDone && onImportDone();
      } else {
        const err = await res.json();
        alert("Errore import: " + JSON.stringify(err));
      }
    } catch (e) {
      alert("Errore di rete: " + e.message);
    }
    setImporting(false);
  };

  const reset = () => { setFile(null); setRows([]); setErrors([]); setDone(null); setStep("upload"); setAnagraficaMissing([]); };

  return (
    <div style={{ position: "fixed", inset: 0, background: "#000a", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center" }} onClick={onClose}>
      <div style={{ background: T.bg, border: `1px solid ${T.borderHi}`, borderRadius: 6, width: "min(1100px, 95vw)", maxHeight: "90vh", overflowY: "auto", padding: 24 }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <div style={{ color: T.accent, fontWeight: "700", fontSize: "14px" }}>CARICO SEMPLICE</div>
          <button style={css.btn()} onClick={onClose}>✕ Chiudi</button>
        </div>

        {/* Step indicator */}
        <div style={{ display: "flex", gap: 8, marginBottom: 24, alignItems: "center" }}>
          {["upload","preview","result"].map((s, i) => (
            <div key={s} style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <div style={{ width: 24, height: 24, borderRadius: "50%", background: step === s ? T.accent : T.borderHi, color: step === s ? "#000" : T.textMid, display: "flex", alignItems: "center", justifyContent: "center", fontSize: "11px", fontWeight: "700" }}>{i+1}</div>
              <span style={{ color: step === s ? T.accent : T.textMid, fontSize: "11px", textTransform: "uppercase", letterSpacing: "0.08em" }}>{s === "upload" ? "Carica file" : s === "preview" ? "Verifica" : "Completato"}</span>
              {i < 2 && <span style={{ color: T.textDim }}>›</span>}
            </div>
          ))}
        </div>

        {/* STEP 1: UPLOAD */}
        {step === "upload" && (
          <div style={{ maxWidth: 500 }}>
            <div style={{ marginBottom: 16 }}>
              <label style={{ color: T.textMid, fontSize: "10px", textTransform: "uppercase", letterSpacing: "0.08em", display: "block", marginBottom: 4 }}>Giro di destinazione</label>
              <select style={{ ...css.input, width: "100%", boxSizing: "border-box" }} value={giroSel ?? ""} onChange={e => setGiroSel(parseInt(e.target.value))}>
                {giriList.map(g => <option key={g.id} value={g.id}>{g.label ?? `Giro ${g.numero} ${g.anno}`}</option>)}
              </select>
            </div>
            <div style={{ border: `2px dashed ${T.borderHi}`, borderRadius: 6, padding: 40, textAlign: "center", marginBottom: 20 }}>
              <div style={{ fontSize: "32px", marginBottom: 12 }}>📂</div>
              <div style={{ color: T.text, marginBottom: 8 }}>Carica il template compilato</div>
              <div style={{ color: T.textMid, fontSize: "11px", marginBottom: 20 }}>Solo file .xlsx — template "Carico Semplice"</div>
              <input type="file" accept=".xlsx" onChange={handleFile} style={{ display: "none" }} id="file-input-carico" />
              <label htmlFor="file-input-carico" style={{ ...css.btn("accent"), cursor: "pointer", padding: "8px 20px" }}>Scegli file .xlsx</label>
            </div>
          </div>
        )}

        {/* STEP 2: PREVIEW */}
        {step === "preview" && (
          <div>
            <div style={{ display: "flex", gap: 12, marginBottom: 16, alignItems: "center", flexWrap: "wrap" }}>
              <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 4, padding: "10px 16px" }}>
                <span style={{ color: T.textMid, fontSize: "11px" }}>Righe lette: </span>
                <span style={{ color: T.text, fontWeight: "700" }}>{rows.length}</span>
              </div>
              <div style={{ background: T.surface, border: `1px solid ${errors.length > 0 ? T.red : T.green}`, borderRadius: 4, padding: "10px 16px" }}>
                <span style={{ color: T.textMid, fontSize: "11px" }}>Errori: </span>
                <span style={{ color: errors.length > 0 ? T.red : T.green, fontWeight: "700" }}>{errors.length}</span>
              </div>
              {anagraficaMissing.length > 0 && (
                <div style={{ background: T.surface, border: `1px solid ${T.accent}66`, borderRadius: 4, padding: "10px 16px" }}>
                  <span style={{ color: T.textMid, fontSize: "11px" }}>Editori non in anagrafica: </span>
                  <span style={{ color: T.accent, fontWeight: "700" }}>{anagraficaMissing.length}</span>
                </div>
              )}
              <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
                <button style={css.btn()} onClick={reset}>← Ricarica</button>
                <button style={css.btn("accent")} onClick={handleImport} disabled={importing || errors.length > 0}>
                  {importing ? "Import in corso..." : `Importa ${rows.length} titoli`}
                </button>
              </div>
            </div>

            {errors.length > 0 && (
              <div style={{ background: T.red + "11", border: `1px solid ${T.red}44`, borderRadius: 4, padding: 16, marginBottom: 16 }}>
                <div style={{ color: T.red, fontWeight: "700", marginBottom: 8, fontSize: "12px" }}>⚠ Correggi gli errori prima di importare</div>
                {errors.map((e, i) => <div key={i} style={{ color: T.textMid, fontSize: "11px" }}>Riga {e.row}: campi mancanti/errati → {e.fields.join(", ")}</div>)}
              </div>
            )}

            {anagraficaMissing.length > 0 && (
              <div style={{ background: T.accent + "11", border: `1px solid ${T.accent}44`, borderRadius: 4, padding: 16, marginBottom: 16 }}>
                <div style={{ color: T.accent, fontWeight: "700", marginBottom: 8, fontSize: "12px" }}>⚠ Editori non presenti in ranking_editori — verranno importati senza cod. editore / account / promozione / ranking</div>
                {anagraficaMissing.map((e, i) => <div key={i} style={{ color: T.textMid, fontSize: "11px" }}>{e}</div>)}
              </div>
            )}

            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead>
                  <tr>
                    {["Pos.","EAN","Titolo","Autore","Editore","Cod.Ed.","Account","Promo","Prezzo","Uscita","Obj","Gemelli"].map(h => <th key={h} style={css.th}>{h}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {rows.map((r, i) => {
                    const hasErr = errors.find(e => e.row === i + 5);
                    return (
                      <tr key={i} style={{ background: hasErr ? T.red + "11" : i % 2 === 0 ? "transparent" : T.surface + "66" }}>
                        <td style={{ ...css.td, color: T.accent, fontWeight: "700", textAlign: "center" }}>{r.posizione}</td>
                        <td style={{ ...css.td, fontFamily: "monospace", fontSize: "11px", color: T.textMid }}>{r.ean}</td>
                        <td style={{ ...css.td, maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontWeight: "600" }}>{r.titolo}</td>
                        <td style={{ ...css.td, color: T.textMid }}>{r.autore}</td>
                        <td style={{ ...css.td, color: T.accent }}>{r.editore_nome}</td>
                        <td style={{ ...css.td, color: r.codice_editore ? T.text : T.red, fontSize: "11px" }}>{r.codice_editore ?? "—"}</td>
                        <td style={{ ...css.td, color: r.account_editore ? T.text : T.red, fontSize: "11px" }}>{r.account_editore ?? "—"}</td>
                        <td style={{ ...css.td, color: r.promozione ? T.text : T.red, fontSize: "11px" }}>{r.promozione ?? "—"}</td>
                        <td style={css.td}>€ {r.prezzo}</td>
                        <td style={{ ...css.td, color: T.textMid }}>{r.uscita}</td>
                        <td style={css.td}>{r.obiettivo_assegnato}</td>
                        <td style={css.td}>{r.ean_gemello_1 && <div style={{ fontSize: "10px", color: T.textMid }}><div>{r.ean_gemello_1}</div>{r.ean_gemello_2 && <div>{r.ean_gemello_2}</div>}{r.ean_gemello_3 && <div>{r.ean_gemello_3}</div>}</div>}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* STEP 3: RESULT */}
        {step === "result" && done && (
          <div style={{ textAlign: "center", padding: 60 }}>
            <div style={{ fontSize: "48px", marginBottom: 16 }}>✅</div>
            <div style={{ color: T.green, fontSize: "20px", fontWeight: "700", marginBottom: 8 }}>Import completato</div>
            <div style={{ color: T.textMid, marginBottom: 32 }}>{done.ok} titoli importati nel giro selezionato</div>
            <div style={{ display: "flex", gap: 8, justifyContent: "center" }}>
              <button style={css.btn()} onClick={reset}>Nuovo carico</button>
              <button style={css.btn("accent")} onClick={onClose}>Chiudi</button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
