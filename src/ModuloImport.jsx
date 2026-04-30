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

// Mappa colonne template → campi DB
const COL_MAP = {
  0: "n_cedola", 1: "ranking_editore", 2: "ranking_titolo", 3: "ean",
  4: "titolo", 5: "autore", 6: "codice_editore", 7: "editore_nome",
  8: "prezzo", 9: "uscita", 10: "formato", 11: "eta",
  12: "obiettivo_assegnato", 13: "il_triangolo", 14: "top_100",
  15: "account_editore", 16: "promozione", 17: "note_comunicazione", 18: "note",
  19: "ean_gemello_1", 20: "titolo_gemello_1",
  21: "ean_gemello_2", 22: "titolo_gemello_2",
  23: "ean_gemello_3", 24: "titolo_gemello_3",
};

const REQUIRED = ["ean", "titolo", "editore_nome", "prezzo", "formato"];

function parseXlsx(buffer) {
  // Parser XLSX minimale (legge le celle come testo dal file zip)
  // Per parsing completo usare SheetJS nel browser
  return null;
}

export default function ModuloImport({ giriList, token, onImportDone }) {
  const [giroSel, setGiroSel] = useState(giriList[0]?.id ?? null);
  const [file, setFile] = useState(null);
  const [rows, setRows] = useState([]);
  const [errors, setErrors] = useState([]);
  const [importing, setImporting] = useState(false);
  const [done, setDone] = useState(null);
  const [step, setStep] = useState("upload"); // upload | preview | result

  const handleFile = useCallback((e) => {
    const f = e.target.files[0];
    if (!f) return;
    setFile(f);

    // Usa SheetJS (caricato via CDN in index.html)
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const XLSX = window.XLSX;
        const wb = XLSX.read(evt.target.result, { type: "array" });
        const ws = wb.Sheets["CEDOLA"];
        if (!ws) { alert("Foglio 'CEDOLA' non trovato. Stai usando il template corretto?"); return; }
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        // Skip righe 0-2 (titolo, legenda, note), header riga 3, dati da riga 4
        const dataRows = data.slice(4).filter(r => r.some(v => v !== ""));

        const parsed = [];
        const errs = [];
        dataRows.forEach((row, idx) => {
          const obj = {};
          Object.entries(COL_MAP).forEach(([colIdx, field]) => {
            let val = row[parseInt(colIdx)] ?? "";
            if (field === "prezzo") val = parseFloat(val) || null;
            if (field === "obiettivo_assegnato") val = parseInt(val) || 0;
            if (field === "il_triangolo" || field === "top_100") val = String(val).toUpperCase() === "SI";
            if (field === "n_cedola" || field === "ranking_editore" || field === "ranking_titolo") val = parseInt(val) || null;
            obj[field] = val === "" ? null : val;
          });
          obj.giro_id = null; // assegnato all'import

          // Validazione
          const rowErrs = [];
          REQUIRED.forEach(f => { if (!obj[f]) rowErrs.push(f); });
          if (obj.ean && String(obj.ean).replace(/\D/g,"").length !== 13) rowErrs.push("EAN non valido");
          if (rowErrs.length) errs.push({ row: idx + 5, fields: rowErrs });

          parsed.push(obj);
        });

        setRows(parsed);
        setErrors(errs);
        setStep("preview");
      } catch (err) {
        alert("Errore lettura file: " + err.message);
      }
    };
    reader.readAsArrayBuffer(f);
  }, []);

  const handleImport = async () => {
    if (!giroSel) return;
    setImporting(true);
    const payload = rows.map(r => ({ ...r, giro_id: giroSel }));
    try {
      const res = await fetch(`${SUPABASE_URL}/rest/v1/titoli`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "apikey": SUPABASE_KEY,
          "Authorization": `Bearer ${token}`,
          "Prefer": "return=minimal",
        },
        body: JSON.stringify(payload),
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

  const reset = () => { setFile(null); setRows([]); setErrors([]); setDone(null); setStep("upload"); };

  return (
    <div style={{ flex: 1, overflowY: "auto", padding: 24 }}>
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
          <div style={{ marginBottom: 20 }}>
            <div style={{ color: T.textMid, fontSize: "11px", letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 8 }}>Giro di destinazione</div>
            <select style={css.input} value={giroSel ?? ""} onChange={(e) => setGiroSel(Number(e.target.value))}>
              {giriList.map((g) => <option key={g.id} value={g.id}>{g.label}</option>)}
            </select>
          </div>
          <div style={{ border: `2px dashed ${T.borderHi}`, borderRadius: 6, padding: 40, textAlign: "center", marginBottom: 20 }}>
            <div style={{ fontSize: "32px", marginBottom: 12 }}>📂</div>
            <div style={{ color: T.text, marginBottom: 8 }}>Carica il template compilato</div>
            <div style={{ color: T.textMid, fontSize: "11px", marginBottom: 20 }}>Solo file .xlsx — usa il template ufficiale</div>
            <input type="file" accept=".xlsx" onChange={handleFile} style={{ display: "none" }} id="file-input" />
            <label htmlFor="file-input" style={{ ...css.btn("accent"), cursor: "pointer", padding: "8px 20px" }}>Scegli file .xlsx</label>
          </div>
          <div style={{ color: T.textMid, fontSize: "11px" }}>
            Non hai il template? <span style={{ color: T.accent, cursor: "pointer" }} onClick={() => alert("Scarica template_cedola.xlsx dai file del progetto")}>Scaricalo qui</span>
          </div>
        </div>
      )}

      {/* STEP 2: PREVIEW */}
      {step === "preview" && (
        <div>
          <div style={{ display: "flex", gap: 12, marginBottom: 16, alignItems: "center" }}>
            <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 4, padding: "10px 16px" }}>
              <span style={{ color: T.textMid, fontSize: "11px" }}>Righe lette: </span>
              <span style={{ color: T.text, fontWeight: "700" }}>{rows.length}</span>
            </div>
            <div style={{ background: T.surface, border: `1px solid ${errors.length > 0 ? T.red : T.green}`, borderRadius: 4, padding: "10px 16px" }}>
              <span style={{ color: T.textMid, fontSize: "11px" }}>Errori: </span>
              <span style={{ color: errors.length > 0 ? T.red : T.green, fontWeight: "700" }}>{errors.length}</span>
            </div>
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

          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead>
                <tr>
                  {["#","EAN","Titolo","Autore","Editore","Prezzo","Uscita","Formato","Obj","▲","★"].map(h => <th key={h} style={css.th}>{h}</th>)}
                </tr>
              </thead>
              <tbody>
                {rows.map((r, i) => {
                  const hasErr = errors.find(e => e.row === i + 5);
                  return (
                    <tr key={i} style={{ background: hasErr ? T.red + "11" : i % 2 === 0 ? "transparent" : T.surface + "66" }}>
                      <td style={{ ...css.td, color: T.textMid }}>{r.n_cedola}</td>
                      <td style={{ ...css.td, fontFamily: "monospace", fontSize: "11px", color: T.textMid }}>{r.ean}</td>
                      <td style={{ ...css.td, maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontWeight: "600" }}>{r.titolo}</td>
                      <td style={{ ...css.td, color: T.textMid }}>{r.autore}</td>
                      <td style={{ ...css.td, color: T.accent }}>{r.editore_nome}</td>
                      <td style={css.td}>€ {r.prezzo}</td>
                      <td style={{ ...css.td, color: T.textMid }}>{r.uscita}</td>
                      <td style={css.td}>{r.formato}</td>
                      <td style={css.td}>{r.obiettivo_assegnato}</td>
                      <td style={css.td}>{r.il_triangolo ? "▲" : ""}</td>
                      <td style={css.td}>{r.top_100 ? "★" : ""}</td>
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
          <button style={css.btn("accent")} onClick={reset}>Nuovo import</button>
        </div>
      )}
    </div>
  );
}
