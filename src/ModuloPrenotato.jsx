import { useState, useCallback, useMemo } from "react";

const SUPABASE_URL = "https://tdflwenlylhctxssatax.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InRkZmx3ZW5seWxoY3R4c3NhdGF4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYzMzgyNzYsImV4cCI6MjA5MTkxNDI3Nn0.l35qEL7LOvyYuI1McQlVqj4vbyTqmlevcmqWbTGYi2Q";

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

const GRUPPI_CANALE = {
  2: "FELTRINELLI", 23: "FELTRINELLI", 59: "FELTRINELLI", 77: "FELTRINELLI",
  8: "MONDADORI", 34: "MONDADORI", 80: "MONDADORI",
  83: "UBIK",
  32: "GIUNTI",
  11: "LIBRACCIO", 56: "LIBRACCIO", 88: "LIBRACCIO", 90: "LIBRACCIO", 91: "LIBRACCIO", 92: "LIBRACCIO",
  6: "LIB_RELIGIOSE", 18: "LIB_RELIGIOSE", 19: "LIB_RELIGIOSE", 21: "LIB_RELIGIOSE", 57: "LIB_RELIGIOSE",
  36: "LIB_COOP",
  4: "INDIPENDENTI_ALTRE_CATENE", 22: "INDIPENDENTI_ALTRE_CATENE", 24: "INDIPENDENTI_ALTRE_CATENE", 60: "INDIPENDENTI_ALTRE_CATENE",
  28: "FASTBOOK",
  63: "CENTROLIBRI",
  25: "GROSSISTI", 30: "GROSSISTI", 94: "GROSSISTI",
  82: "AMAZON",
  58: "IBS",
  33: "ALTRI_ONLINE",
};

const CANALI_LABELS = {
  FELTRINELLI: "Feltrinelli", GIUNTI: "Giunti al Punto", MONDADORI: "Mondadori",
  UBIK: "Ubik", LIBRACCIO: "Libraccio", INDIPENDENTI_ALTRE_CATENE: "Indipendenti",
  LIB_RELIGIOSE: "Librerie Religiose", LIB_COOP: "Librerie Coop", ALTRI_ONLINE: "Librerie On-line",
  AMAZON: "Amazon", IBS: "Stereo Online", FASTBOOK: "Fastbook + GD", GROSSISTI: "Grossisti",
  CENTROLIBRI: "Centro Libri",
};

// Legge un valore da una riga per nome colonna o indice (0-based)
// Prova prima il nome esatto, poi cerca per indice se il file usa header numerici
function leggiCella(row, nomeColonna, indice, headers) {
  // Prima prova il nome diretto
  if (row[nomeColonna] !== undefined && row[nomeColonna] !== "") return row[nomeColonna];
  // Poi prova varianti case-insensitive
  const key = Object.keys(row).find(k => k.trim().toLowerCase() === nomeColonna.toLowerCase());
  if (key && row[key] !== "") return row[key];
  // Fallback: usa l'indice di colonna tramite l'array headers
  if (headers && indice < headers.length) {
    const hKey = headers[indice];
    if (hKey !== undefined && row[hKey] !== undefined) return row[hKey];
  }
  return "";
}

export default function ModuloPrenotato({ token, titoli, onImportDone }) {
  const [step, setStep] = useState("upload");
  const [righe, setRighe] = useState([]);
  const [aggregato, setAggregato] = useState([]);
  const [aggregatoClienti, setAggregatoClienti] = useState([]);
  const [importing, setImporting] = useState(false);
  const [done, setDone] = useState(null);

  const handleFile = useCallback((e) => {
    const f = e.target.files[0];
    if (!f) return;
    const XLSX = window.XLSX;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array" });
        const sheetName = wb.SheetNames.find(n => n.toLowerCase().includes("pianifica")) || wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];

        // Leggi con header:1 per avere sia nomi che posizioni
        const dataRaw = XLSX.utils.sheet_to_json(ws, { defval: "", header: 1 });
        // Prima riga = headers
        const headers = dataRaw[0] ? dataRaw[0].map(h => String(h || "").trim()) : [];
        // Righe dati come oggetti keyed per nome colonna
        const data = XLSX.utils.sheet_to_json(ws, { defval: "" });
        setRighe(data);

        const aggMap = {};
        const aggCliMap = {};

        data.forEach(row => {
          const ean = String(row["EAN"] || "").trim();
          const gruppoRaw = row["Gruppo cliente"];
          const qta = parseInt(row["Pren (Qtà)"]) || 0;
          const codiceCliente = String(row["Codice cliente"] || "").trim();
          const nomeCliente = String(row["Nome Cliente"] || "").trim();
          if (!ean || qta === 0) return;

          // Colonne aggiuntive: F=5 (N. ordine cliente), S=18 (sconto occ.), U=20 (pagamento occ.)
          const numOrdineCliente = String(leggiCella(row, "N. ordine cliente", 5, headers) || leggiCella(row, "Ordine cliente", 5, headers) || "").trim();
          const scontoOccRaw = leggiCella(row, "Sconto occasionale", 18, headers) || leggiCella(row, "Sc. occas.", 18, headers) || "";
          const scontoOcc = parseFloat(String(scontoOccRaw).replace(",", ".")) || 0;
          const pagamentoOcc = String(leggiCella(row, "Pagamento occasionale", 20, headers) || leggiCella(row, "Pag. occas.", 20, headers) || leggiCella(row, "Pag(occ)", 20, headers) || "").trim();

          let canale = "INDIPENDENTI_ALTRE_CATENE";
          if (gruppoRaw !== "" && gruppoRaw !== null && gruppoRaw !== undefined) {
            const gruppoInt = parseInt(parseFloat(gruppoRaw));
            if (!isNaN(gruppoInt) && gruppoInt !== 0) {
              canale = GRUPPI_CANALE[gruppoInt] || "INDIPENDENTI_ALTRE_CATENE";
            }
          }

          const key = `${ean}__${canale}`;
          if (!aggMap[key]) aggMap[key] = { ean, canale, qta: 0 };
          aggMap[key].qta += qta;

          if (codiceCliente) {
            const keyC = `${codiceCliente}__${ean}__${canale}`;
            if (!aggCliMap[keyC]) aggCliMap[keyC] = {
              codice_cliente: codiceCliente,
              nome_cliente: nomeCliente,
              ean, canale, qta: 0,
              sconto_occasionale: scontoOcc > 0 ? scontoOcc : null,
              pagamento_occasionale: pagamentoOcc || null,
              num_ordine_cliente: numOrdineCliente || null,
            };
            aggCliMap[keyC].qta += qta;
            // Se una riga ha sconto/pagamento/ordine, lo salviamo (prende il primo non-null trovato)
            if (scontoOcc > 0 && !aggCliMap[keyC].sconto_occasionale) aggCliMap[keyC].sconto_occasionale = scontoOcc;
            if (pagamentoOcc && !aggCliMap[keyC].pagamento_occasionale) aggCliMap[keyC].pagamento_occasionale = pagamentoOcc;
            if (numOrdineCliente && !aggCliMap[keyC].num_ordine_cliente) aggCliMap[keyC].num_ordine_cliente = numOrdineCliente;
          }
        });

        const result = Object.values(aggMap).map(r => {
          const titolo = titoli.find(t => t.ean === r.ean || t.ean === String(parseInt(r.ean)));
          return { ...r, titolo: titolo?.titolo ?? "— non trovato —", found: !!titolo, titolo_id: titolo?.id };
        }).sort((a, b) => a.ean.localeCompare(b.ean));

        const resultClienti = Object.values(aggCliMap).map(r => {
          const titolo = titoli.find(t => t.ean === r.ean || t.ean === String(parseInt(r.ean)));
          return { ...r, found: !!titolo, titolo_id: titolo?.id };
        });

        setAggregato(result);
        setAggregatoClienti(resultClienti);
        setStep("preview");
      } catch (err) {
        alert("Errore lettura file: " + err.message);
      }
    };
    reader.readAsArrayBuffer(f);
  }, [titoli]);

  const riepilogoCanale = useMemo(() => {
    const map = {};
    aggregato.forEach(r => { map[r.canale] = (map[r.canale] || 0) + r.qta; });
    return Object.entries(map).sort((a, b) => b[1] - a[1]);
  }, [aggregato]);

  const totaleAggregato = aggregato.reduce((s, r) => s + r.qta, 0);
  const totaleFound = aggregato.filter(r => r.found).reduce((s, r) => s + r.qta, 0);

  const handleImport = async () => {
    setImporting(true);

    const rCanali = await fetch(`${SUPABASE_URL}/rest/v1/canali?select=id,codice`, {
      headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}` }
    });
    const canaliDB = await rCanali.json();
    const canaleMap = {};
    canaliDB.forEach(c => { canaleMap[c.codice] = c.id; });

    const validi = aggregato.filter(r => r.found && r.titolo_id);
    const payload = validi.map(r => ({
      titolo_id: r.titolo_id,
      canale_id: canaleMap[r.canale] || null,
      quantita: r.qta,
    })).filter(r => r.canale_id !== null);

    const res1 = await fetch(`${SUPABASE_URL}/rest/v1/rpc/upsert_prenotato`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}` },
      body: JSON.stringify({ payload }),
    });

    const validiClienti = aggregatoClienti.filter(r => r.found && r.titolo_id);
    const payloadClienti = validiClienti.map(r => ({
      codice_cliente: r.codice_cliente,
      nome_cliente: r.nome_cliente,
      canale_id: canaleMap[r.canale] || null,
      titolo_id: r.titolo_id,
      quantita: r.qta,
      sconto_occasionale: r.sconto_occasionale ?? null,
      pagamento_occasionale: r.pagamento_occasionale ?? null,
      num_ordine_cliente: r.num_ordine_cliente ?? null,
    })).filter(r => r.canale_id !== null);

    for (let i = 0; i < payloadClienti.length; i += 500) {
      const batch = payloadClienti.slice(i, i + 500);
      await fetch(`${SUPABASE_URL}/rest/v1/rpc/upsert_prenotato_clienti`, {
        method: "POST",
        headers: { "Content-Type": "application/json", "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}` },
        body: JSON.stringify({ payload: batch }),
      });
    }

    if (res1.ok) {
      setDone({ ok: payload.length, totQta: totaleFound });
      setStep("result");
      onImportDone && onImportDone();
    } else {
      const err = await res1.json();
      alert("Errore import: " + JSON.stringify(err));
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
