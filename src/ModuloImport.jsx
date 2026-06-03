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
  0: "giro_label", 1: "n_cedola", 2: "ranking_editore", 3: "ranking_titolo", 4: "ean",
  5: "titolo", 6: "autore", 7: "codice_editore", 8: "editore_nome",
  9: "prezzo", 10: "uscita", 11: "formato", 12: "eta",
  13: "obiettivo_assegnato", 14: "il_triangolo", 15: "top_100",
  16: "account_editore", 17: "promozione", 18: "note_comunicazione", 19: "note",
  20: "ean_gemello_1", 21: "titolo_gemello_1",
  22: "ean_gemello_2", 23: "titolo_gemello_2",
  24: "ean_gemello_3", 25: "titolo_gemello_3",
};

const REQUIRED = ["giro_label", "ean", "titolo", "editore_nome", "prezzo", "formato"];

// Mappa colonne XL obiettivi → canale_codice DB
const CANALI_MAP = {
  3:  "INDIPENDENTI_ALTRE_CATENE",
  4:  "FELTRINELLI",
  5:  "GIUNTI",
  6:  "MONDADORI",
  7:  "UBIK",
  8:  "AMAZON",
  9:  "IBS",
  10: "ALTRI_ONLINE",
  11: "FASTBOOK",
  12: "GROSSISTI",
  13: "CENTROLIBRI",
  14: "GDO",
};

// Verifica se l'utente loggato è admin leggendo user_profiles
async function checkIsAdmin(token) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/user_profiles?select=ruolo&limit=1`, {
    headers: {
      "apikey": SUPABASE_KEY,
      "Authorization": `Bearer ${token}`,
    },
  });
  if (res.ok) {
    const data = await res.json();
    return data?.[0]?.ruolo === "admin";
  }
  return false;
}

// ─── Componente Import Obiettivi ──────────────────────────────────────────────
function ImportObiettiviSection({ token }) {
  const [adminVerified, setAdminVerified] = useState(false);
  const [checking, setChecking] = useState(false);
  const [adminError, setAdminError] = useState(null);

  const [file, setFile] = useState(null);
  const [preview, setPreview] = useState([]); // [{editore, formato, canale, val}]
  const [parseError, setParseError] = useState(null);
  const [step, setStep] = useState("upload"); // upload | preview | result
  const [importing, setImporting] = useState(false);
  const [result, setResult] = useState(null);
  const [progress, setProgress] = useState(0);

  // Verifica admin al click
  const handleVerifyAdmin = async () => {
    setChecking(true);
    setAdminError(null);
    try {
      const ok = await checkIsAdmin(token);
      if (ok) {
        setAdminVerified(true);
      } else {
        setAdminError("Accesso negato: utente non admin.");
      }
    } catch (e) {
      setAdminError("Errore verifica: " + e.message);
    }
    setChecking(false);
  };

  const handleFile = useCallback((e) => {
    const f = e.target.files[0];
    if (!f) return;
    setFile(f);
    setParseError(null);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const XLSX = window.XLSX;
        const wb = XLSX.read(evt.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        if (!ws) { setParseError("Foglio non trovato nel file."); return; }

        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        // Riga 0: header gruppo, Riga 1: header colonne, Dati da riga 2
        const rows = data.slice(2).filter(r => r[0] && String(r[0]).trim() !== "");

        const parsed = [];
        rows.forEach((row) => {
          const editore = String(row[0]).trim().toUpperCase();
          const formato = String(row[1]).trim();
          if (!editore || !formato || editore === "EDITORE") return;

          Object.entries(CANALI_MAP).forEach(([colIdx, canale]) => {
            const raw = row[parseInt(colIdx)];
            const val = parseFloat(raw);
            parsed.push({
              editore,
              formato,
              canale,
              val: isNaN(val) ? 0 : Math.round(val * 10000) / 10000,
            });
          });
        });

        if (parsed.length === 0) {
          setParseError("Nessun dato trovato. Verifica il formato del file.");
          return;
        }

        setPreview(parsed);
        setStep("preview");
      } catch (err) {
        setParseError("Errore lettura file: " + err.message);
      }
    };
    reader.readAsArrayBuffer(f);
  }, []);

  const handleImport = async () => {
    setImporting(true);
    setProgress(0);
    let ok = 0, err = 0;

    // Upsert a batch di 100
    const BATCH = 100;
    const total = preview.length;

    for (let i = 0; i < total; i += BATCH) {
      const batch = preview.slice(i, i + BATCH).map(r => ({
        editore_nome: r.editore,
        formato: r.formato,
        canale_codice: r.canale,
        percentuale: r.val,
      }));

      const res = await fetch(
        `${SUPABASE_URL}/rest/v1/spalmatura_obiettivo`,
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "apikey": SUPABASE_KEY,
            "Authorization": `Bearer ${token}`,
            "Prefer": "resolution=merge-duplicates",
          },
          body: JSON.stringify(batch),
        }
      );

      if (res.ok || res.status === 201) {
        ok += batch.length;
      } else {
        err += batch.length;
      }

      setProgress(Math.round(((i + BATCH) / total) * 100));
    }

    setResult({ ok, err, total });
    setStep("result");
    setImporting(false);
  };

  const reset = () => {
    setFile(null);
    setPreview([]);
    setParseError(null);
    setStep("upload");
    setResult(null);
    setProgress(0);
  };

  // Raggruppa preview per editore/formato per mostrare tabella leggibile
  const previewGrouped = preview.reduce((acc, r) => {
    const key = `${r.editore}||${r.formato}`;
    if (!acc[key]) acc[key] = { editore: r.editore, formato: r.formato, canali: {} };
    acc[key].canali[r.canale] = r.val;
    return acc;
  }, {});
  const previewRows = Object.values(previewGrouped);
  const canaliCols = Object.values(CANALI_MAP);

  return (
    <div style={{ marginTop: 40, borderTop: `1px solid ${T.border}`, paddingTop: 32 }}>
      {/* Header sezione */}
      <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 20 }}>
        <div style={{ width: 3, height: 20, background: T.accent, borderRadius: 2 }} />
        <span style={{ color: T.text, fontWeight: "600", fontSize: "13px", letterSpacing: "0.05em" }}>
          IMPORT OBIETTIVI
        </span>
        <span style={{ color: T.textDim, fontSize: "11px" }}>spalmatura_obiettivo</span>
        <span style={{
          marginLeft: "auto",
          background: T.red + "22",
          color: T.red,
          border: `1px solid ${T.red}44`,
          borderRadius: 3,
          padding: "2px 8px",
          fontSize: "10px",
          letterSpacing: "0.1em",
          fontWeight: "700",
        }}>
          ADMIN
        </span>
      </div>

      {/* Blocco verifica admin */}
      {!adminVerified ? (
        <div style={{ maxWidth: 420, background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, padding: 24 }}>
          <div style={{ color: T.textMid, fontSize: "12px", marginBottom: 16 }}>
            Questa funzione è riservata agli amministratori.<br />
            Clicca per verificare il tuo accesso.
          </div>
          <button
            style={css.btn("accent")}
            onClick={handleVerifyAdmin}
            disabled={checking}
          >
            {checking ? "Verifica in corso..." : "Verifica accesso admin"}
          </button>
          {adminError && (
            <div style={{ color: T.red, fontSize: "11px", marginTop: 10 }}>{adminError}</div>
          )}
        </div>
      ) : (
        <>
          {/* Step indicator */}
          <div style={{ display: "flex", gap: 8, marginBottom: 20, alignItems: "center" }}>
            {["upload", "preview", "result"].map((s, i) => (
              <div key={s} style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <div style={{
                  width: 22, height: 22, borderRadius: "50%",
                  background: step === s ? T.accent : T.borderHi,
                  color: step === s ? "#000" : T.textMid,
                  display: "flex", alignItems: "center", justifyContent: "center",
                  fontSize: "10px", fontWeight: "700"
                }}>{i + 1}</div>
                <span style={{ color: step === s ? T.accent : T.textMid, fontSize: "11px", textTransform: "uppercase", letterSpacing: "0.08em" }}>
                  {s === "upload" ? "Carica file" : s === "preview" ? "Anteprima" : "Completato"}
                </span>
                {i < 2 && <span style={{ color: T.textDim }}>›</span>}
              </div>
            ))}
          </div>

          {/* STEP 1: UPLOAD */}
          {step === "upload" && (
            <div style={{ maxWidth: 480 }}>
              <div style={{ border: `2px dashed ${T.borderHi}`, borderRadius: 6, padding: 36, textAlign: "center" }}>
                <div style={{ fontSize: "28px", marginBottom: 10 }}>📊</div>
                <div style={{ color: T.text, marginBottom: 6, fontSize: "13px" }}>
                  Carica <strong>obiettivi_per_spalmatura.xlsx</strong>
                </div>
                <div style={{ color: T.textMid, fontSize: "11px", marginBottom: 18 }}>
                  Tutti i valori esistenti verranno sovrascritti
                </div>
                <input type="file" accept=".xlsx" onChange={handleFile} style={{ display: "none" }} id="obj-file-input" />
                <label htmlFor="obj-file-input" style={{ ...css.btn("accent"), cursor: "pointer", padding: "8px 20px" }}>
                  Scegli file .xlsx
                </label>
              </div>
              {parseError && (
                <div style={{ color: T.red, fontSize: "11px", marginTop: 10 }}>{parseError}</div>
              )}
            </div>
          )}

          {/* STEP 2: PREVIEW */}
          {step === "preview" && (
            <div>
              <div style={{ display: "flex", gap: 12, marginBottom: 14, alignItems: "center" }}>
                <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 4, padding: "8px 14px" }}>
                  <span style={{ color: T.textMid, fontSize: "11px" }}>Editore/formato: </span>
                  <span style={{ color: T.text, fontWeight: "700" }}>{previewRows.length}</span>
                </div>
                <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 4, padding: "8px 14px" }}>
                  <span style={{ color: T.textMid, fontSize: "11px" }}>Righe upsert: </span>
                  <span style={{ color: T.accent, fontWeight: "700" }}>{preview.length}</span>
                </div>
                <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
                  <button style={css.btn()} onClick={reset}>← Ricarica</button>
                  <button style={css.btn("accent")} onClick={handleImport} disabled={importing}>
                    {importing ? `Import... ${progress}%` : `Aggiorna ${previewRows.length} editori`}
                  </button>
                </div>
              </div>

              {importing && (
                <div style={{ marginBottom: 14 }}>
                  <div style={{ height: 4, background: T.border, borderRadius: 2, overflow: "hidden" }}>
                    <div style={{ height: "100%", width: `${progress}%`, background: T.accent, transition: "width 0.3s" }} />
                  </div>
                </div>
              )}

              <div style={{ overflowX: "auto", maxHeight: 400, overflowY: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "11px" }}>
                  <thead style={{ position: "sticky", top: 0, zIndex: 1 }}>
                    <tr>
                      <th style={{ ...css.th, minWidth: 180 }}>Editore</th>
                      <th style={css.th}>Formato</th>
                      {canaliCols.map(c => (
                        <th key={c} style={{ ...css.th, minWidth: 70, fontSize: "10px" }}>
                          {c.replace("INDIPENDENTI_ALTRE_CATENE", "INDIPEN.").replace("ALTRI_ONLINE", "ALTRI ON.").replace("CENTROLIBRI", "CENTRO.").replace("GROSSISTI", "GROSS.")}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {previewRows.map((r, i) => (
                      <tr key={i} style={{ background: i % 2 === 0 ? "transparent" : T.surface + "55" }}>
                        <td style={{ ...css.td, color: T.accent, fontWeight: "600" }}>{r.editore}</td>
                        <td style={{ ...css.td, color: T.textMid }}>{r.formato}</td>
                        {canaliCols.map(c => (
                          <td key={c} style={{ ...css.td, textAlign: "right", color: r.canali[c] > 0 ? T.text : T.textDim }}>
                            {r.canali[c] > 0 ? `${(r.canali[c] * 100).toFixed(1)}%` : "—"}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* STEP 3: RESULT */}
          {step === "result" && result && (
            <div style={{ textAlign: "center", padding: 48 }}>
              <div style={{ fontSize: "40px", marginBottom: 14 }}>{result.err === 0 ? "✅" : "⚠️"}</div>
              <div style={{ color: result.err === 0 ? T.green : T.accent, fontSize: "18px", fontWeight: "700", marginBottom: 8 }}>
                {result.err === 0 ? "Import completato" : "Import con errori"}
              </div>
              <div style={{ color: T.textMid, marginBottom: 6, fontSize: "12px" }}>
                {result.ok} righe aggiornate correttamente
              </div>
              {result.err > 0 && (
                <div style={{ color: T.red, fontSize: "12px", marginBottom: 6 }}>
                  {result.err} righe non importate
                </div>
              )}
              <button style={{ ...css.btn("accent"), marginTop: 24 }} onClick={reset}>Nuovo import</button>
            </div>
          )}
        </>
      )}
    </div>
  );
}

// ─── Componente principale ────────────────────────────────────────────────────
export default function ModuloImport({ giriList, token, onImportDone }) {
  const [giroSel, setGiroSel] = useState(giriList[0]?.id ?? null);
  const [file, setFile] = useState(null);
  const [rows, setRows] = useState([]);
  const [errors, setErrors] = useState([]);
  const [importing, setImporting] = useState(false);
  const [done, setDone] = useState(null);
  const [step, setStep] = useState("upload");

  const handleFile = useCallback((e) => {
    const f = e.target.files[0];
    if (!f) return;
    setFile(f);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const XLSX = window.XLSX;
        const wb = XLSX.read(evt.target.result, { type: "array" });
        const ws = wb.Sheets["CEDOLA"];
        if (!ws) { alert("Foglio 'CEDOLA' non trovato. Stai usando il template corretto?"); return; }
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        const dataRows = data.slice(4).filter(r => r.some(v => v !== ""));

        const parsed = [];
        const errs = [];
        dataRows.forEach((row, idx) => {
          const obj = {};
          Object.entries(COL_MAP).forEach(([colIdx, field]) => {
            let val = row[parseInt(colIdx)] ?? "";
            if (field === "prezzo") val = parseFloat(val) || null;
            if (field === "obiettivo_assegnato") val = parseInt(val) || 0;
            if (field === "il_triangolo" || field === "top_100") val = String(val).trim().toUpperCase() === "SI";
            if (field === "ranking_editore" || field === "ranking_titolo") val = parseInt(val) || null;
            if (field === "n_cedola" || field === "giro_label") val = val ? String(val) : null;
            obj[field] = val === "" ? null : val;
          });
          obj.giro_id = null;

          const rowErrs = [];
          REQUIRED.forEach(f => { if (!obj[f]) rowErrs.push(f); });
          if (obj.ean && String(obj.ean).replace(/\D/g, "").length !== 13) rowErrs.push("EAN non valido");
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

  const reset = () => { setFile(null); setRows([]); setErrors([]); setDone(null); setStep("upload"); };

  return (
    <div style={{ flex: 1, overflowY: "auto", padding: 24 }}>
      {/* Step indicator cedola */}
      <div style={{ display: "flex", gap: 8, marginBottom: 24, alignItems: "center" }}>
        {["upload", "preview", "result"].map((s, i) => (
          <div key={s} style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <div style={{ width: 24, height: 24, borderRadius: "50%", background: step === s ? T.accent : T.borderHi, color: step === s ? "#000" : T.textMid, display: "flex", alignItems: "center", justifyContent: "center", fontSize: "11px", fontWeight: "700" }}>{i + 1}</div>
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
                  {["#", "EAN", "Titolo", "Autore", "Editore", "Prezzo", "Uscita", "Formato", "Obj", "▲", "★"].map(h => <th key={h} style={css.th}>{h}</th>)}
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

      {/* ── IMPORT OBIETTIVI (solo admin) ── */}
      <ImportObiettiviSection token={token} />
    </div>
  );
}
