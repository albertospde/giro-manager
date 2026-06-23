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
// NOTA: ranking_editore rimosso dal template — viene recuperato da DB via codice_editore
// Le colonne nel template sono ora:
// 0=giro_label, 1=n_cedola, 2=ranking_titolo, 3=ean, 4=titolo, 5=autore,
// 6=codice_editore, 7=editore_nome, 8=prezzo, 9=uscita, 10=formato, 11=eta,
// 12=obiettivo_assegnato, 13=il_triangolo, 14=top_100,
// 15=account_editore, 16=promozione, 17=note_comunicazione, 18=note,
// 19=ean_gemello_1, 20=titolo_gemello_1,
// 21=ean_gemello_2, 22=titolo_gemello_2,
// 23=ean_gemello_3, 24=titolo_gemello_3
const COL_MAP = {
  0: "giro_label", 1: "n_cedola", 2: "ranking_titolo", 3: "ean",
  4: "titolo", 5: "autore", 6: "codice_editore", 7: "editore_nome",
  8: "prezzo", 9: "uscita", 10: "formato", 11: "eta",
  12: "obiettivo_assegnato", 13: "il_triangolo", 14: "top_100",
  15: "account_editore", 16: "promozione", 17: "note_comunicazione", 18: "note",
  19: "ean_gemello_1", 20: "titolo_gemello_1",
  21: "ean_gemello_2", 22: "titolo_gemello_2",
  23: "ean_gemello_3", 24: "titolo_gemello_3",
};

const REQUIRED = ["giro_label", "ean", "titolo", "editore_nome", "prezzo", "formato"];

// ─── Fetch ranking editori da Supabase ───────────────────────────────────────
async function fetchRankingEditori(token) {
  const res = await fetch(
    `${SUPABASE_URL}/rest/v1/ranking_editori?select=editore_nome,ranking&order=ranking.asc`,
    {
      headers: {
        "apikey": SUPABASE_KEY,
        "Authorization": `Bearer ${token}`,
      },
    }
  );
  if (!res.ok) {
    console.error("Fetch ranking_editori fallita:", res.status, res.statusText);
    throw new Error(`Errore caricamento ranking editori (HTTP ${res.status})`);
  }
  const data = await res.json();
  // Mappa editore_nome → ranking minimo (nel caso di doppioni sullo stesso nome, prende il più basso)
  const map = {};
  data.forEach(({ editore_nome, ranking }) => {
    const nome = String(editore_nome ?? "").trim().toUpperCase();
    if (!nome) return;
    if (!(nome in map) || ranking < map[nome]) {
      map[nome] = ranking;
    }
  });
  return map;
}

// ─── Componente principale ───────────────────────────────────────────────────

export default function ModuloImport({ giriList, token, onImportDone }) {
  const [giroSel, setGiroSel] = useState(giriList[0]?.id ?? null);
  const [file, setFile] = useState(null);
  const [rows, setRows] = useState([]);
  const [errors, setErrors] = useState([]);
  const [importing, setImporting] = useState(false);
  const [done, setDone] = useState(null);
  const [step, setStep] = useState("upload"); // upload | preview | result
  const [rankingMissing, setRankingMissing] = useState([]); // editori non trovati nel DB

  const handleFile = useCallback(async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    setFile(f);

    // Carica ranking editori dal DB
    let rankingMap = {};
    try {
      rankingMap = await fetchRankingEditori(token);
    } catch (err) {
      alert("Impossibile caricare il ranking editori dal database: " + err.message + "\nRiprova; se l'errore persiste, controlla la sessione/login.");
      return;
    }

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
        console.log("Prima riga dati:", dataRows[0]);

        const parsed = [];
        const errs = [];
        const missing = new Set();

        dataRows.forEach((row, idx) => {
          const obj = {};
          Object.entries(COL_MAP).forEach(([colIdx, field]) => {
            let val = row[parseInt(colIdx)] ?? "";
            if (field === "prezzo") val = parseFloat(val) || null;
            if (field === "obiettivo_assegnato") val = parseInt(val) || 0;
            if (field === "il_triangolo" || field === "top_100") val = String(val).trim().toUpperCase() === "SI";
            if (field === "ranking_titolo") val = parseInt(val) || null;
            if (field === "n_cedola" || field === "giro_label") val = val ? String(val) : null;
            // Campi testo → stampatello maiuscolo
            const CAMPI_TESTO = ["titolo","autore","editore_nome","codice_editore","uscita","formato","eta","account_editore","promozione","note_comunicazione","note","n_cedola","giro_label","titolo_gemello_1","titolo_gemello_2","titolo_gemello_3"];
            if (CAMPI_TESTO.includes(field) && typeof val === "string" && val !== "") val = val.trim().toUpperCase();
            obj[field] = val === "" ? null : val;
          });
          obj.giro_id = null; // assegnato all'import

          // Risolvi ranking_editore dal DB tramite editore_nome
          const nome = String(obj.editore_nome ?? "").trim().toUpperCase();
          if (nome && rankingMap[nome] !== undefined) {
            obj.ranking_editore = rankingMap[nome];
          } else {
            obj.ranking_editore = null;
            if (nome) missing.add(`${nome}${obj.codice_editore ? ` (cod. ${obj.codice_editore})` : ""}`);
          }

          // Validazione
          const rowErrs = [];
          REQUIRED.forEach(f => { if (!obj[f]) rowErrs.push(f); });
          if (obj.ean && String(obj.ean).replace(/\D/g,"").length !== 13) rowErrs.push("EAN non valido");
          if (rowErrs.length) errs.push({ row: idx + 5, fields: rowErrs });

          parsed.push(obj);
        });

        setRows(parsed);
        setErrors(errs);
        setRankingMissing([...missing]);
        setStep("preview");
      } catch (err) {
        alert("Errore lettura file: " + err.message);
      }
    };
    reader.readAsArrayBuffer(f);
  }, [token]);

  const handleImport = async () => {
    if (!giroSel) return;
    setImporting(true);
    const payload = rows.map(r => ({ ...r, giro_id: giroSel }));
    try {
      const res = await fetch(`${SUPABASE_URL}/rest/v1/rpc/upsert_titoli`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "apikey": SUPABASE_KEY,
          "Authorization": `Bearer ${token}`,
        },
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

  const reset = () => { setFile(null); setRows([]); setErrors([]); setDone(null); setStep("upload"); setRankingMissing([]); };

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
          <div style={{ border: `2px dashed ${T.borderHi}`, borderRadius: 6, padding: 40, textAlign: "center", marginBottom: 20 }}>
            <div style={{ fontSize: "32px", marginBottom: 12 }}>📂</div>
            <div style={{ color: T.text, marginBottom: 8 }}>Carica il template compilato</div>
            <div style={{ color: T.textMid, fontSize: "11px", marginBottom: 20 }}>Solo file .xlsx — usa il template ufficiale</div>
            <input type="file" accept=".xlsx" onChange={handleFile} style={{ display: "none" }} id="file-input" />
            <label htmlFor="file-input" style={{ ...css.btn("accent"), cursor: "pointer", padding: "8px 20px" }}>Scegli file .xlsx</label>
          </div>
          <div style={{ color: T.textMid, fontSize: "11px" }}>
            Non hai il template? <a href="https://albertospde.github.io/giro-manager/template_cedola.xlsx" download style={{ color: T.accent }}>Scaricalo qui</a>
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
            {/* Badge editori senza ranking */}
            {rankingMissing.length > 0 && (
              <div style={{ background: T.surface, border: `1px solid ${T.accent}66`, borderRadius: 4, padding: "10px 16px" }}>
                <span style={{ color: T.textMid, fontSize: "11px" }}>Ranking editore non trovato: </span>
                <span style={{ color: T.accent, fontWeight: "700" }}>{rankingMissing.length}</span>
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

          {/* Warning editori senza ranking — non bloccante */}
          {rankingMissing.length > 0 && (
            <div style={{ background: T.accent + "11", border: `1px solid ${T.accent}44`, borderRadius: 4, padding: 16, marginBottom: 16 }}>
              <div style={{ color: T.accent, fontWeight: "700", marginBottom: 8, fontSize: "12px" }}>⚠ Editori non presenti in ranking_editori (ranking_editore = null)</div>
              {rankingMissing.map((e, i) => <div key={i} style={{ color: T.textMid, fontSize: "11px" }}>{e}</div>)}
            </div>
          )}

          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead>
                <tr>
                  {["#","Rk Ed.","EAN","Titolo","Autore","Editore","Prezzo","Uscita","Formato","Obj","▲","★"].map(h => <th key={h} style={css.th}>{h}</th>)}
                </tr>
              </thead>
              <tbody>
                {rows.map((r, i) => {
                  const hasErr = errors.find(e => e.row === i + 5);
                  return (
                    <tr key={i} style={{ background: hasErr ? T.red + "11" : i % 2 === 0 ? "transparent" : T.surface + "66" }}>
                      <td style={{ ...css.td, color: T.textMid }}>{r.n_cedola}</td>
                      <td style={{ ...css.td, color: r.ranking_editore ? T.accent : T.red, fontWeight: "700", fontSize: "11px" }}>
                        {r.ranking_editore ?? "—"}
                      </td>
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
