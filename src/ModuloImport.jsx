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
  th: { padding: "8px 12px", textAlign: "left", color: T.textMid, fontWeight: "400", fontSize: "11px", letterSpacing: "0.08em", textTransform: "uppercase", borderBottom: `1px solid ${T.border}`, whiteSpace: "nowrap", background: T.surface },
  td: { padding: "7px 12px", borderBottom: `1px solid ${T.border}22`, verticalAlign: "middle", fontSize: "12px" },
};

// Mappa colonne template → campi DB.
// Template CEDOLA — colonne (riga 4 in poi, foglio "CEDOLA"):
// 0=giro (es. "5 2026"), 1=ean, 2=titolo, 3=autore, 4=editore, 5=prezzo, 6=uscita,
// 7=obiettivo, 8=top 100, 9=promozione, 10=note,
// 11=ean gemello 1, 12=titolo gemello 1, 13=ean gemello 2, 14=titolo gemello 2,
// 15=ean gemello 3, 16=titolo gemello 3
// NON in template: N. CEDOLA (calcolato), RANKING EDITORE/TITOLO (automatico/da ordine riga),
// ACCOUNT EDITORE (automatico da anagrafica), CODICE EDITORE, FORMATO, ETÀ, IL TRIANGOLO, NOTE COMUNICAZIONE.
const COL_MAP = {
  0: "giro_raw", 1: "ean", 2: "titolo", 3: "autore", 4: "editore_nome",
  5: "prezzo", 6: "uscita", 7: "obiettivo_assegnato", 8: "top_100",
  9: "promozione", 10: "note",
  11: "ean_gemello_1", 12: "titolo_gemello_1",
  13: "ean_gemello_2", 14: "titolo_gemello_2",
  15: "ean_gemello_3", 16: "titolo_gemello_3",
};

const REQUIRED = ["giro_raw", "ean", "titolo", "editore_nome", "prezzo"];
const FORMATO_DEFAULT = "Cover";
const GIRO_RE = /^(\d+)\s+(\d{4})$/;

// ─── Fetch anagrafica editori (ranking, account, cedola/categoria) ──────────
async function fetchAnagraficaEditori(token) {
  const res = await fetch(
    `${SUPABASE_URL}/rest/v1/ranking_editori?select=editore_nome,codice_editore,ranking,account_editore,promozione,cedola`,
    { headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}` } }
  );
  if (!res.ok) throw new Error(`Errore caricamento anagrafica editori (HTTP ${res.status})`);
  const data = await res.json();
  const map = {};
  data.forEach(({ editore_nome, codice_editore, ranking, account_editore, promozione, cedola }) => {
    const nome = String(editore_nome ?? "").trim().toUpperCase();
    if (!nome) return;
    if (!(nome in map) || ranking < map[nome].ranking) {
      map[nome] = { codice_editore, ranking, account_editore, promozione, cedola: String(cedola ?? "").trim().toUpperCase() };
    }
  });
  return map;
}

// ─── Risolvi (o crea) il giro_id per ogni combinazione numero+anno+categoria presente ──
// Usa il vincolo UNIQUE(numero, anno, sub_giro) della tabella giri per upsert idempotente.
async function resolveGiri(token, combos) {
  // combos: Set di "numero|anno|categoria"
  const payload = [...combos].map(c => {
    const [numero, anno, categoria] = c.split("|");
    return { numero: parseInt(numero), anno: parseInt(anno), sub_giro: categoria, label: `GIRO ${numero} ${anno} ${categoria}`, attivo: true };
  });
  const res = await fetch(
    `${SUPABASE_URL}/rest/v1/giri?on_conflict=numero,anno,sub_giro`,
    {
      method: "POST",
      headers: {
        "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`,
        "Content-Type": "application/json",
        "Prefer": "resolution=merge-duplicates,return=representation",
      },
      body: JSON.stringify(payload),
    }
  );
  if (!res.ok) throw new Error(`Errore creazione/risoluzione giro (HTTP ${res.status})`);
  const rows = await res.json();
  const map = {};
  rows.forEach(r => { map[`${r.numero}|${r.anno}|${String(r.sub_giro).toUpperCase()}`] = r.id; });
  return map;
}

// ─── Componente principale ───────────────────────────────────────────────────

export default function ModuloImport({ token, onImportDone }) {
  const [file, setFile] = useState(null);
  const [rows, setRows] = useState([]);
  const [errors, setErrors] = useState([]);
  const [importing, setImporting] = useState(false);
  const [loadingFile, setLoadingFile] = useState(false);
  const [done, setDone] = useState(null);
  const [step, setStep] = useState("upload"); // upload | preview | result
  const [anagraficaMissing, setAnagraficaMissing] = useState([]); // editori non trovati in ranking_editori

  const handleFile = useCallback(async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    setFile(f);
    setLoadingFile(true);

    let anagraficaMap = {};
    try {
      anagraficaMap = await fetchAnagraficaEditori(token);
    } catch (err) {
      alert("Impossibile caricare l'anagrafica editori dal database: " + err.message + "\nRiprova; se l'errore persiste, controlla la sessione/login.");
      setLoadingFile(false);
      return;
    }

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const XLSX = window.XLSX;
        const wb = XLSX.read(evt.target.result, { type: "array" });
        const ws = wb.Sheets["CEDOLA"];
        if (!ws) { alert("Foglio 'CEDOLA' non trovato. Stai usando il template corretto?"); setLoadingFile(false); return; }
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        // Skip righe 0-2 (titolo, legenda, note), header riga 3, dati da riga 4
        const dataRows = data.slice(4).filter(r => r.some(v => v !== ""));

        const parsed = [];
        const errs = [];
        const missing = new Set();
        const contatori = {}; // posizione per gruppo giro_label|editore_nome
        const giroCombos = new Set(); // numero|anno|categoria da risolvere su `giri`

        dataRows.forEach((row, idx) => {
          const obj = {};
          Object.entries(COL_MAP).forEach(([colIdx, field]) => {
            let val = row[parseInt(colIdx)] ?? "";
            if (field === "prezzo") val = parseFloat(String(val).replace(",", ".")) || null;
            if (field === "obiettivo_assegnato") val = parseInt(val) || 0;
            if (field === "top_100") val = String(val).trim().toUpperCase() === "SI";
            const CAMPI_TESTO = ["titolo", "autore", "editore_nome", "uscita", "promozione", "note", "titolo_gemello_1", "titolo_gemello_2", "titolo_gemello_3"];
            if (CAMPI_TESTO.includes(field) && typeof val === "string" && val !== "") val = val.trim().toUpperCase();
            obj[field] = val === "" ? null : val;
          });

          const rowErrs = [];

          // Parsing GIRO ("5 2026" → numero=5, anno=2026)
          const giroStr = String(obj.giro_raw ?? "").trim();
          const m = giroStr.match(GIRO_RE);
          let numero = null, anno = null;
          if (m) { numero = parseInt(m[1]); anno = parseInt(m[2]); }
          else rowErrs.push("GIRO non valido (formato atteso: 'N ANNO', es. 5 2026)");
          obj.giro_label = m ? `${numero} ${anno}` : giroStr;
          delete obj.giro_raw;

          // Anagrafica editore: ranking, account, categoria (cedola) — automatica
          const nome = String(obj.editore_nome ?? "").trim().toUpperCase();
          const anagrafica = nome ? anagraficaMap[nome] : null;
          if (anagrafica) {
            obj.codice_editore = anagrafica.codice_editore ?? null;
            obj.ranking_editore = anagrafica.ranking ?? null;
            obj.account_editore = anagrafica.account_editore ?? null;
            if (!obj.promozione) obj.promozione = anagrafica.promozione ?? null; // dal file se presente, altrimenti da anagrafica
            const categoria = anagrafica.cedola;
            if (categoria && m) {
              obj.n_cedola = `GIRO ${numero} ${anno} ${categoria}`;
              giroCombos.add(`${numero}|${anno}|${categoria}`);
              obj._giroKey = `${numero}|${anno}|${categoria}`;
            } else if (!categoria) {
              rowErrs.push("editore senza categoria cedola (A/B/C/KIDS/SERVICE) in anagrafica");
            }
          } else {
            obj.codice_editore = null;
            obj.ranking_editore = null;
            obj.account_editore = null;
            rowErrs.push("editore non trovato in anagrafica");
            if (nome) missing.add(nome);
          }

          obj.formato = FORMATO_DEFAULT;
          obj.eta = null;
          obj.il_triangolo = null;
          obj.note_comunicazione = null;

          // Posizione: sequenziale nell'ordine delle righe del file, per gruppo giro+editore,
          // sempre ripartendo da 1. Un titolo già presente mantiene la stessa posizione se l'ordine
          // nel file non cambia; se inserisci una riga sopra di lui, si sposta di conseguenza
          // (la RPC sovrascrive sempre la posizione con questo valore, vedi upsert_titoli).
          const keyGruppo = `${obj.giro_label ?? ""}|${nome}`;
          if (!(keyGruppo in contatori)) contatori[keyGruppo] = 0;
          contatori[keyGruppo] += 1;
          obj.posizione = contatori[keyGruppo];

          REQUIRED.forEach(f => { if (!obj[f] && f !== "giro_raw") rowErrs.push(`campo mancante: ${f}`); });
          if (obj.ean && String(obj.ean).replace(/\D/g, "").length !== 13) rowErrs.push("EAN non valido");
          if (rowErrs.length) errs.push({ row: idx + 5, fields: rowErrs });

          parsed.push(obj);
        });

        // Risolvi/crea i giro_id per tutte le combinazioni numero+anno+categoria trovate
        let giriMap = {};
        if (giroCombos.size > 0) {
          try {
            giriMap = await resolveGiri(token, giroCombos);
          } catch (err) {
            alert("Errore nella risoluzione dei giri su Supabase: " + err.message);
            setLoadingFile(false);
            return;
          }
        }
        parsed.forEach(obj => {
          if (obj._giroKey) obj.giro_id = giriMap[obj._giroKey] ?? null;
          delete obj._giroKey;
        });

        setRows(parsed);
        setErrors(errs);
        setAnagraficaMissing([...missing]);
        setStep("preview");
      } catch (err) {
        alert("Errore lettura file: " + err.message);
      }
      setLoadingFile(false);
    };
    reader.readAsArrayBuffer(f);
  }, [token]);

  const handleImport = async () => {
    setImporting(true);
    const payload = rows;
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

  const reset = () => { setFile(null); setRows([]); setErrors([]); setDone(null); setStep("upload"); setAnagraficaMissing([]); };

  return (
    <div style={{ flex: 1, overflowY: "auto", padding: 24 }}>
      {/* Step indicator */}
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
            <div style={{ color: T.text, marginBottom: 8 }}>{loadingFile ? "Elaborazione in corso..." : "Carica il template compilato"}</div>
            <div style={{ color: T.textMid, fontSize: "11px", marginBottom: 20 }}>Solo file .xlsx — usa il template ufficiale</div>
            <input type="file" accept=".xlsx" onChange={handleFile} style={{ display: "none" }} id="file-input" disabled={loadingFile} />
            <label htmlFor="file-input" style={{ ...css.btn("accent"), cursor: loadingFile ? "default" : "pointer", padding: "8px 20px", opacity: loadingFile ? 0.6 : 1 }}>Scegli file .xlsx</label>
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
              {errors.map((e, i) => <div key={i} style={{ color: T.textMid, fontSize: "11px" }}>Riga {e.row}: {e.fields.join(", ")}</div>)}
            </div>
          )}

          {anagraficaMissing.length > 0 && (
            <div style={{ background: T.red + "11", border: `1px solid ${T.red}44`, borderRadius: 4, padding: 16, marginBottom: 16 }}>
              <div style={{ color: T.red, fontWeight: "700", marginBottom: 8, fontSize: "12px" }}>⚠ Editori non presenti in anagrafica (ranking_editori) — aggiungili lì prima di importare</div>
              {anagraficaMissing.map((e, i) => <div key={i} style={{ color: T.textMid, fontSize: "11px" }}>{e}</div>)}
            </div>
          )}

          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead>
                <tr>
                  {["N. Cedola", "Pos.", "Rk Ed.", "EAN", "Titolo", "Autore", "Editore", "Acc.", "Prezzo", "Uscita", "Obj", "★"].map(h => <th key={h} style={css.th}>{h}</th>)}
                </tr>
              </thead>
              <tbody>
                {rows.map((r, i) => {
                  const hasErr = errors.find(e => e.row === i + 5);
                  return (
                    <tr key={i} style={{ background: hasErr ? T.red + "11" : i % 2 === 0 ? "transparent" : T.surface + "66" }}>
                      <td style={{ ...css.td, color: r.n_cedola ? T.accent : T.red, fontWeight: "600", fontSize: "11px" }}>{r.n_cedola ?? "—"}</td>
                      <td style={{ ...css.td, color: T.textMid, textAlign: "center" }}>{r.posizione}</td>
                      <td style={{ ...css.td, color: r.ranking_editore ? T.accent : T.red, fontWeight: "700", fontSize: "11px" }}>{r.ranking_editore ?? "—"}</td>
                      <td style={{ ...css.td, fontFamily: "monospace", fontSize: "11px", color: T.textMid }}>{r.ean}</td>
                      <td style={{ ...css.td, maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontWeight: "600" }}>{r.titolo}</td>
                      <td style={{ ...css.td, color: T.textMid }}>{r.autore}</td>
                      <td style={css.td}>{r.editore_nome}</td>
                      <td style={{ ...css.td, color: T.textMid, fontSize: "11px" }}>{r.account_editore ?? "—"}</td>
                      <td style={css.td}>€ {r.prezzo}</td>
                      <td style={{ ...css.td, color: T.textMid }}>{r.uscita}</td>
                      <td style={css.td}>{r.obiettivo_assegnato}</td>
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
          <div style={{ color: T.textMid, marginBottom: 32 }}>{done.ok} titoli importati, suddivisi in automatico per cedola</div>
          <button style={css.btn("accent")} onClick={reset}>Nuovo import</button>
        </div>
      )}
    </div>
  );
}
