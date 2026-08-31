import { useState, useCallback, useMemo } from "react";
import { fetchPrenotatoRpn, parseEaggrega, importAggregato, CANALI_LABELS } from "./rpnPrenotatoSync.js";

const T = {
  bg: "#1a2140", surface: "#212d54", border: "#2e3d6b", borderHi: "#3d4f82",
  text: "#f0f2f8", textMid: "#8b9cc8", textDim: "#4a5a8a",
  accent: "#7b9fe8", green: "#4caf7d", red: "#e05c5c",
};

const css = {
  btn: (v = "default") => ({ padding: "6px 14px", border: `1px solid ${v === "accent" ? T.accent : T.border}`, background: v === "accent" ? T.accent : "transparent", color: v === "accent" ? "#000" : T.text, cursor: "pointer", fontSize: "12px", fontFamily: "inherit", borderRadius: 3, fontWeight: v === "accent" ? "700" : "400" }),
  input: { background: T.bg, border: `1px solid ${T.border}`, color: T.text, padding: "5px 10px", fontSize: "12px", fontFamily: "inherit", borderRadius: 3, outline: "none" },
  th: { padding: "8px 12px", textAlign: "left", color: T.textMid, fontWeight: "400", fontSize: "11px", letterSpacing: "0.08em", textTransform: "uppercase", borderBottom: `1px solid ${T.border}`, whiteSpace: "nowrap", background: T.surface, position: "sticky", top: 0, zIndex: 1 },
  td: { padding: "7px 12px", borderBottom: `1px solid ${T.border}22`, verticalAlign: "middle", fontSize: "12px" },
};

export default function ModuloPrenotato({ token, titoli, onImportDone }) {
  const [step, setStep] = useState("upload");
  const [righe, setRighe] = useState([]);
  const [aggregato, setAggregato] = useState([]);
  const [aggregatoClienti, setAggregatoClienti] = useState([]);
  const [importing, setImporting] = useState(false);
  const [done, setDone] = useState(null);

  const applyParsed = useCallback((arrayBuffer) => {
    try {
      const r = parseEaggrega(arrayBuffer, titoli);
      setRighe(r.righe);
      setAggregato(r.aggregato);
      setAggregatoClienti(r.aggregatoClienti);
      setStep("preview");
    } catch (err) {
      alert("Errore lettura file: " + err.message);
    }
  }, [titoli]);

  const handleFile = useCallback((e) => {
    const f = e.target.files[0];
    if (!f) return;
    const reader = new FileReader();
    reader.onload = (evt) => applyParsed(evt.target.result);
    reader.readAsArrayBuffer(f);
  }, [applyParsed]);

  // Giri disponibili (esclude EXTRA), più recente per primo.
  const giriDisponibili = useMemo(() => {
    const set = new Set(titoli.map(t => t.giro_label).filter(l => l && l !== "EXTRA" && !l.startsWith("EXTRA")));
    return [...set].sort((a, b) => {
      const [na, ya] = a.split(" "); const [nb, yb] = b.split(" ");
      return Number(yb) - Number(ya) || Number(nb) - Number(na);
    });
  }, [titoli]);

  const [giroLabelRpn, setGiroLabelRpn] = useState("");
  const [syncingRpn, setSyncingRpn] = useState(false);
  const [rpnError, setRpnError] = useState(null);
  const giroLabelSel = giroLabelRpn || giriDisponibili[0] || "";

  const syncFromRpn = useCallback(async () => {
    if (!giroLabelSel) return;
    setSyncingRpn(true);
    setRpnError(null);
    try {
      const arrayBuffer = await fetchPrenotatoRpn(token, giroLabelSel);
      applyParsed(arrayBuffer);
    } catch (err) {
      setRpnError(err.message);
    } finally {
      setSyncingRpn(false);
    }
  }, [giroLabelSel, token, applyParsed]);

  const riepilogoCanale = useMemo(() => {
    const map = {};
    aggregato.forEach(r => { map[r.canale] = (map[r.canale] || 0) + r.qta; });
    return Object.entries(map).sort((a, b) => b[1] - a[1]);
  }, [aggregato]);

  const totaleAggregato = aggregato.reduce((s, r) => s + r.qta, 0);
  const totaleFound = aggregato.filter(r => r.found).reduce((s, r) => s + r.qta, 0);

  const handleImport = async () => {
    setImporting(true);
    try {
      const result = await importAggregato(token, aggregato, aggregatoClienti);
      setDone(result);
      setStep("result");
      onImportDone && onImportDone();
    } catch (err) {
      alert(err.message);
    }
    setImporting(false);
  };

  const reset = () => { setRighe([]); setAggregato([]); setAggregatoClienti([]); setDone(null); setStep("upload"); };

  return (
    <div style={{ flex: 1, overflowY: "auto", padding: 24 }}>
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

      {step === "upload" && (
        <div style={{ maxWidth: 500 }}>
          <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, padding: 20, marginBottom: 16 }}>
            <div style={{ color: T.textMid, fontSize: "10px", textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: 12 }}>Aggiorna da RPN</div>
            <div style={{ display: "flex", gap: 8, marginBottom: rpnError ? 10 : 0 }}>
              <select
                style={{ ...css.input, flex: 1 }}
                value={giroLabelSel}
                onChange={e => setGiroLabelRpn(e.target.value)}
                disabled={syncingRpn}
              >
                {giriDisponibili.map(g => <option key={g} value={g}>GIRO {g}</option>)}
              </select>
              <button style={css.btn("accent")} onClick={syncFromRpn} disabled={syncingRpn || !giroLabelSel}>
                {syncingRpn ? "Scarico da RPN..." : "📡 Aggiorna da RPN"}
              </button>
            </div>
            {rpnError && (
              <div style={{ color: T.red, fontSize: "11px", marginTop: 4 }}>{rpnError}</div>
            )}
          </div>

          <div style={{ textAlign: "center", color: T.textDim, fontSize: "11px", margin: "12px 0" }}>oppure</div>

          <div style={{ border: `2px dashed ${T.borderHi}`, borderRadius: 6, padding: 40, textAlign: "center", marginBottom: 20 }}>
            <div style={{ fontSize: "32px", marginBottom: 12 }}>📂</div>
            <div style={{ color: T.text, marginBottom: 8 }}>Carica il file "Pianifica Visite"</div>
            <div style={{ color: T.textMid, fontSize: "11px", marginBottom: 20 }}>File .xlsx esportato dal sistema Messaggerie</div>
            <input type="file" accept=".xlsx" onChange={handleFile} style={{ display: "none" }} id="pv-file-input" />
            <label htmlFor="pv-file-input" style={{ ...css.btn("accent"), cursor: "pointer", padding: "8px 20px" }}>Scegli file .xlsx</label>
          </div>
        </div>
      )}

      {step === "preview" && (
        <div>
          <div style={{ display: "flex", gap: 12, marginBottom: 20, flexWrap: "wrap" }}>
            {[["Righe lette", righe.length.toLocaleString("it"), T.text], ["Totale copie", totaleAggregato.toLocaleString("it"), T.accent], ["Trovati in cedola", totaleFound.toLocaleString("it"), T.green], ["Non trovati", aggregato.filter(r => !r.found).length, T.red]].map(([label, val, color]) => (
              <div key={label} style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 4, padding: "10px 16px" }}>
                <div style={{ color: T.textMid, fontSize: "10px", textTransform: "uppercase", marginBottom: 4 }}>{label}</div>
                <div style={{ color, fontWeight: "700", fontSize: "20px" }}>{val}</div>
              </div>
            ))}
          </div>

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
              {importing ? "Import in corso..." : `Importa ${totaleFound.toLocaleString("it")} copie`}
            </button>
          </div>

          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead>
                <tr>{["EAN","Titolo","Canale","Qtà",""].map(h => <th key={h} style={css.th}>{h}</th>)}</tr>
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

      {step === "result" && done && (
        <div style={{ textAlign: "center", padding: 60 }}>
          <div style={{ fontSize: "48px", marginBottom: 16 }}>✅</div>
          <div style={{ color: T.green, fontSize: "20px", fontWeight: "700", marginBottom: 8 }}>Import completato</div>
          <div style={{ color: T.textMid, marginBottom: 24 }}>{done.totQta?.toLocaleString("it")} copie importate</div>
          <button style={css.btn("accent")} onClick={reset}>Nuovo import</button>
        </div>
      )}
    </div>
  );
}
