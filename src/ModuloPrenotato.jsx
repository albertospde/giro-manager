import { useState, useCallback, useMemo } from "react";

const SUPABASE_URL = "https://tdflwenlylhctxssatax.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InRkZmx3ZW5seWxoY3R4c3NhdGF4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYzMzgyNzYsImV4cCI6MjA5MTkxNDI3Nn0.l35qEL7LOvyYuI1McQlVqj4vbyTqmlevcmqWbTGYi2Q";

const T = {
  bg: "#1a2140", surface: "#212d54", border: "#2e3d6b", borderHi: "#3d4f82",
  text: "#f0f2f8", textMid: "#8b9cc8", textDim: "#4a5a8a",
  accent: "#7b9fe8", green: "#4caf7d", red: "#e05c5c", blue: "#4a5da0",
};

const css = {
  btn: (v = "default") => ({ padding: "6px 14px", border: `1px solid ${v === "accent" ? T.accent : v === "danger" ? T.red : T.border}`, background: v === "accent" ? T.accent : "transparent", color: v === "accent" ? "#000" : T.text, cursor: "pointer", fontSize: "12px", fontFamily: "inherit", borderRadius: 3, fontWeight: v === "accent" ? "700" : "400" }),
  input: { background: T.bg, border: `1px solid ${T.border}`, color: T.text, padding: "5px 10px", fontSize: "12px", fontFamily: "inherit", borderRadius: 3, outline: "none" },
  th: { padding: "8px 12px", textAlign: "left", color: T.textMid, fontWeight: "400", fontSize: "11px", letterSpacing: "0.08em", textTransform: "uppercase", borderBottom: `1px solid ${T.border}`, whiteSpace: "nowrap", background: T.surface, position: "sticky", top: 0, zIndex: 1 },
  td: { padding: "7px 12px", borderBottom: `1px solid ${T.border}22`, verticalAlign: "middle", fontSize: "12px" },
};

const CANALI_LABELS = {
  FELTRINELLI: "Feltrinelli", GIUNTI: "Giunti", MONDADORI: "Mondadori",
  UBIK: "Ubik", LIBRACCIO: "Libraccio", INDIPENDENTI_ALTRE_CATENE: "Indip. & Altre Catene",
  LIB_RELIGIOSE: "Lib. Religiose", ALTRI_ONLINE: "Altri Online", AMAZON: "Amazon",
  IBS: "IBS", FASTBOOK: "Fastbook", GROSSISTI: "Grossisti",
  CENTROLIBRI: "Centrolibri", GDO: "GDO",
};

export default function ModuloPrenotato({ token, titoli }) {
  const [step, setStep] = useState("upload");
  const [file, setFile] = useState(null);
  const [righe, setRighe] = useState([]);
  const [aggregato, setAggregato] = useState([]);
  const [gruppiMap, setGruppiMap] = useState({});
  const [importing, setImporting] = useState(false);
  const [done, setDone] = useState(null);
  const [gruppiCaricati, setGruppiCaricati] = useState(false);

  // Carica mappa gruppi da Supabase
  const caricaGruppi = useCallback(async () => {
    if (gruppiCaricati) return gruppiMap;
    const r = await fetch(`${SUPABASE_URL}/rest/v1/gruppi_clienti?select=codice_gruppo,canale_codice`, {
      headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}` }
    });
    const data = await r.json();
    const map = {};
    data.forEach(g => { map[g.codice_gruppo] = g.canale_codice; });
    setGruppiMap(map);
    setGruppiCaricati(true);
    return map;
  }, [token, gruppiMap, gruppiCaricati]);

  const handleFile = useCallback(async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    setFile(f);

    const map = await caricaGruppi();
    const XLSX = window.XLSX;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array" });
        // Cerca il foglio giusto
        const sheetName = wb.SheetNames.find(n => n.toLowerCase().includes("pianifica")) || wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(ws, { defval: "" });

        setRighe(data);

        // Aggrega per EAN × canale
        const aggMap = {};
        data.forEach(row => {
          const ean = String(row["EAN"] || "").trim();
          const gruppoCliente = row["Gruppo cliente"];
          const qta = parseInt(row["Pren (Qtà)"]) || 0;
          if (!ean || !gruppoCliente || qta === 0) return;

          const canale = (!gruppoCliente || gruppoCliente === "") ? "INDIPENDENTI_ALTRE_CATENE" : (map[parseInt(gruppoCliente)] || "INDIPENDENTI_ALTRE_CATENE");
          const key = `${ean}__${canale}`;
          if (!aggMap[key]) aggMap[key] = { ean, canale, qta: 0 };
          aggMap[key].qta += qta;
        });

        // Arricchisci con dati titolo
        const result = Object.values(aggMap).map(r => {
          const titolo = titoli.find(t => t.ean === r.ean);
          return { ...r, titolo: titolo?.titolo ?? "— non trovato —", found: !!titolo, titolo_id: titolo?.id };
        }).sort((a, b) => a.ean.localeCompare(b.ean));

        setAggregato(result);
        setStep("preview");
      } catch (err) {
        alert("Errore lettura file: " + err.message);
      }
    };
    reader.readAsArrayBuffer(f);
  }, [caricaGruppi, titoli]);

  // Riepilogo per canale
  const riepilogoCanale = useMemo(() => {
    const map = {};
    aggregato.forEach(r => {
      if (!map[r.canale]) map[r.canale] = 0;
      map[r.canale] += r.qta;
    });
    return Object.entries(map).sort((a, b) => b[1] - a[1]);
  }, [aggregato]);

  const handleImport = async () => {
    setImporting(true);
    const validi = aggregato.filter(r => r.found && r.titolo_id);

    // Prima: recupera canali da Supabase
    const rCanali = await fetch(`${SUPABASE_URL}/rest/v1/canali?select=id,codice`, {
      headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}` }
    });
    const canaliDB = await rCanali.json();
    const canaleMap = {};
    canaliDB.forEach(c => { canaleMap[c.codice] = c.id; });

    const payload = validi.map(r => ({
      titolo_id: r.titolo_id,
      canale_id: canaleMap[r.canale] || null,
      quantita: r.qta,
    })).filter(r => r.canale_id !== null);

    // Upsert prenotato
    const res = await fetch(`${SUPABASE_URL}/rest/v1/prenotato`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "apikey": SUPABASE_KEY,
        "Authorization": `Bearer ${token}`,
        "Prefer": "resolution=merge-duplicates,return=minimal",
      },
      body: JSON.stringify(payload),
    });

    if (res.ok) {
      setDone({ ok: payload.length, notFound: aggregato.filter(r => !r.found).length });
      setStep("result");
    } else {
      const err = await res.json();
      alert("Errore import: " + JSON.stringify(err));
    }
    setImporting(false);
  };

  const reset = () => { setFile(null); setRighe([]); setAggregato([]); setDone(null); setStep("upload"); };

  return (
    <div style={{ flex: 1, overflowY: "auto", padding: 24 }}>
      {/* Step indicator */}
      <div style={{ display: "flex", gap: 8, marginBottom: 24, alignItems: "center" }}>
        {["upload", "preview", "result"].map((s, i) => (
          <div key={s} style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <div style={{ width: 24, height: 24, borderRadius: "50%", background: step === s ? T.accent : T.borderHi, color: step === s ? "#000" : T.textMid, display: "flex", alignItems: "center", justifyContent: "center", fontSize: "11px", fontWeight: "700" }}>{i + 1}</div>
            <span style={{ color: step === s ? T.accent : T.textMid, fontSize: "11px", textTransform: "uppercase", letterSpacing: "0.08em" }}>
              {s === "upload" ? "Carica file" : s === "preview" ? "Verifica" : "Completato"}
            </span>
            {i < 2 && <span style={{ color: T.textDim }}>›</span>}
          </div>
        ))}
      </div>

      {/* STEP 1: UPLOAD */}
      {step === "upload" && (
        <div style={{ maxWidth: 500 }}>
          <div style={{ border: `2px dashed ${T.borderHi}`, borderRadius: 6, padding: 40, textAlign: "center", marginBottom: 20 }}>
            <div style={{ fontSize: "32px", marginBottom: 12 }}>📂</div>
            <div style={{ color: T.text, marginBottom: 8 }}>Carica il file "Pianifica Visite"</div>
            <div style={{ color: T.textMid, fontSize: "11px", marginBottom: 20 }}>File .xlsx esportato dal sistema</div>
            <input type="file" accept=".xlsx" onChange={handleFile} style={{ display: "none" }} id="pv-file-input" />
            <label htmlFor="pv-file-input" style={{ ...css.btn("accent"), cursor: "pointer", padding: "8px 20px" }}>Scegli file .xlsx</label>
          </div>
          <div style={{ color: T.textMid, fontSize: "11px" }}>
            L'app leggerà automaticamente le quantità prenotate e le aggregherà per canale.
          </div>
        </div>
      )}

      {/* STEP 2: PREVIEW */}
      {step === "preview" && (
        <div>
          <div style={{ display: "flex", gap: 12, marginBottom: 20, flexWrap: "wrap" }}>
            <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 4, padding: "10px 16px" }}>
              <div style={{ color: T.textMid, fontSize: "10px", textTransform: "uppercase", marginBottom: 4 }}>Righe lette</div>
              <div style={{ color: T.text, fontWeight: "700", fontSize: "20px" }}>{righe.length.toLocaleString("it")}</div>
            </div>
            <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 4, padding: "10px 16px" }}>
              <div style={{ color: T.textMid, fontSize: "10px", textTransform: "uppercase", marginBottom: 4 }}>EAN × Canale</div>
              <div style={{ color: T.accent, fontWeight: "700", fontSize: "20px" }}>{aggregato.length}</div>
            </div>
            <div style={{ background: T.surface, border: `1px solid ${T.green}44`, borderRadius: 4, padding: "10px 16px" }}>
              <div style={{ color: T.textMid, fontSize: "10px", textTransform: "uppercase", marginBottom: 4 }}>Trovati in cedola</div>
              <div style={{ color: T.green, fontWeight: "700", fontSize: "20px" }}>{aggregato.filter(r => r.found).length}</div>
            </div>
            <div style={{ background: T.surface, border: `1px solid ${T.red}44`, borderRadius: 4, padding: "10px 16px" }}>
              <div style={{ color: T.textMid, fontSize: "10px", textTransform: "uppercase", marginBottom: 4 }}>Non trovati</div>
              <div style={{ color: T.red, fontWeight: "700", fontSize: "20px" }}>{aggregato.filter(r => !r.found).length}</div>
            </div>
          </div>

          {/* Riepilogo canale */}
          <div style={{ marginBottom: 20, background: T.surface, border: `1px solid ${T.border}`, borderRadius: 4, padding: 16 }}>
            <div style={{ color: T.textMid, fontSize: "10px", textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: 12 }}>Riepilogo per canale</div>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 12 }}>
              {riepilogoCanale.map(([canale, qta]) => (
                <div key={canale} style={{ display: "flex", gap: 8, alignItems: "baseline" }}>
                  <span style={{ color: T.textMid, fontSize: "11px" }}>{CANALI_LABELS[canale] || canale}:</span>
                  <span style={{ color: T.accent, fontWeight: "700" }}>{qta.toLocaleString("it")}</span>
                </div>
              ))}
            </div>
          </div>

          <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
            <button style={css.btn()} onClick={reset}>← Ricarica</button>
            <button style={css.btn("accent")} onClick={handleImport} disabled={importing}>
              {importing ? "Import in corso..." : `Importa ${aggregato.filter(r => r.found).length} righe`}
            </button>
          </div>

          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead>
                <tr>
                  {["EAN", "Titolo", "Canale", "Qtà", ""].map(h => <th key={h} style={css.th}>{h}</th>)}
                </tr>
              </thead>
              <tbody>
                {aggregato.map((r, i) => (
                  <tr key={i} style={{ background: !r.found ? T.red + "11" : i % 2 === 0 ? "transparent" : T.surface + "66" }}>
                    <td style={{ ...css.td, fontFamily: "monospace", fontSize: "11px", color: T.textMid }}>{r.ean}</td>
                    <td style={{ ...css.td, color: r.found ? T.text : T.red, maxWidth: 280, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.titolo}</td>
                    <td style={{ ...css.td, color: T.textMid }}>{CANALI_LABELS[r.canale] || r.canale}</td>
                    <td style={{ ...css.td, color: T.accent, fontWeight: "700" }}>{r.qta.toLocaleString("it")}</td>
                    <td style={css.td}>
                      <span style={{ display: "inline-block", padding: "2px 7px", background: (r.found ? T.green : T.red) + "22", border: `1px solid ${(r.found ? T.green : T.red)}44`, color: r.found ? T.green : T.red, borderRadius: 2, fontSize: "10px", fontWeight: "700" }}>
                        {r.found ? "OK" : "N/F"}
                      </span>
                    </td>
                  </tr>
                ))}
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
          <div style={{ color: T.textMid, marginBottom: 8 }}>{done.ok} righe importate in Supabase</div>
          {done.notFound > 0 && <div style={{ color: T.red, fontSize: "12px", marginBottom: 24 }}>{done.notFound} EAN non trovati nella cedola — ignorati</div>}
          <button style={css.btn("accent")} onClick={reset}>Nuovo import</button>
        </div>
      )}
    </div>
  );
}
