import { useState } from "react";
import { tema, cv } from "./tema.js";

// ─── Carica Fatturato ────────────────────────────────────────────────────────
// Fatturato per libreria ed editore da caricare su RPN (pulsante "Carica fatturato" della home di RPN,
// usato da RPN per la spalmatura degli obiettivi). Accetta:
//  • l'export del gestionale (CSV UTF-16 con tabulazioni): Dest. Merci (cod.) | Editore (cod.) | "1.708 €"
//  • il file già nel formato RPN: Cod.Libreria | Editore | Fatturato
// e lo converte nel formato RPN: codice libreria senza zeri iniziali (numero), editore come numero
// oppure come testo se ha lettere o zeri iniziali ("022", "D92"), fatturato numero intero.
// "Carica su RPN" lo invia con l'account RPN dell'utente (edge function rpn-obiettivi, tipo "fatturato";
// su RPN possono caricare solo ADMIN e RESPONSABILE). "Scarica file" per il caricamento a mano.

const SUPABASE_URL = "https://tdflwenlylhctxssatax.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InRkZmx3ZW5seWxoY3R4c3NhdGF4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYzMzgyNzYsImV4cCI6MjA5MTkxNDI3Nn0.l35qEL7LOvyYuI1McQlVqj4vbyTqmlevcmqWbTGYi2Q";

const T = tema({
  bg: "#1a2140", surface: "#212d54", border: "#2e3d6b", borderHi: "#3d4f82",
  text: "#f0f2f8", textMid: "#8b9cc8", textDim: "#4a5a8a",
  accent: "#7b9fe8", green: "#4caf7d", red: "#e05c5c", amber: "#e0a84c",
});
const css = {
  btn: (v = "default", disabled = false) => ({ padding: "7px 14px", border: `1px solid ${v === "accent" ? T.accent : T.border}`, background: v === "accent" ? T.accent : "transparent", color: v === "accent" ? "#000" : T.text, cursor: disabled ? "default" : "pointer", fontSize: "12px", fontFamily: "inherit", borderRadius: 3, fontWeight: v === "accent" ? 700 : 400, opacity: disabled ? 0.5 : 1, whiteSpace: "nowrap" }),
  card: { background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, padding: 18, marginBottom: 16 },
  th: { padding: "6px 10px", textAlign: "left", color: T.textMid, fontWeight: 400, fontSize: "10px", letterSpacing: "0.08em", textTransform: "uppercase", borderBottom: `1px solid ${T.border}` },
  td: { padding: "5px 10px", borderBottom: `1px solid ${T.border}55`, fontSize: "12px", color: T.text },
  kpi: { background: T.bg, border: `1px solid ${T.border}`, borderRadius: 4, padding: "8px 14px", minWidth: 120 },
};

const norm = (s) => String(s ?? "").replace(/\s+/g, " ").trim().toUpperCase();

// Righe del file: CSV (UTF-16/UTF-8/Windows-1252, separatore tab ; o ,) oppure Excel
function leggiRighe(buf, nomeFile) {
  const XLSX = window.XLSX;
  const b = new Uint8Array(buf);
  const utf16 = (b[0] === 0xff && b[1] === 0xfe) ? "utf-16le" : (b[0] === 0xfe && b[1] === 0xff) ? "utf-16be" : null;
  if (utf16 || /\.(csv|txt)$/i.test(nomeFile)) {
    let testo = new TextDecoder(utf16 || "utf-8").decode(b);
    if (!utf16 && testo.includes("�")) testo = new TextDecoder("windows-1252").decode(b);
    testo = testo.replace(/^﻿/, "");
    const inizio = testo.slice(0, 4000);
    const sep = ["\t", ";", ","].sort((x, y) => inizio.split(y).length - inizio.split(x).length)[0];
    return testo.split(/\r?\n/).map(l => l.split(sep));
  }
  const wb = XLSX.read(buf, { type: "array" });
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: "", raw: false });
}

// "1.708 €" → 1708 ; "1.708,50" → 1708.5 ; "471" → 471
const numero = (v) => {
  const s = String(v ?? "").replace(/[^\d.,-]/g, "");
  if (!s) return null;
  const n = parseFloat(s.includes(",") ? s.replace(/\./g, "").replace(",", ".") : /\.\d{3}($|\.)/.test(s) ? s.replace(/\./g, "") : s);
  return isNaN(n) ? null : n;
};

function converti(righe) {
  const hIdx = righe.slice(0, 10).findIndex(r => /DEST|LIBRERIA/.test(norm(r[0])) && /EDITORE/.test(norm(r[1])));
  if (hIdx < 0) throw new Error("Intestazioni non trovate: servono le colonne \"Dest. Merci (cod.)\" (o \"Cod.Libreria\"), \"Editore\" e il fatturato.");
  const out = [], scartate = [];
  righe.slice(hIdx + 1).forEach((r, i) => {
    if (!r.some(v => String(v).trim() !== "")) return;
    const lib = String(r[0] ?? "").trim(), ed = String(r[1] ?? "").trim(), fat = numero(r[2]);
    if (!/^\d+$/.test(lib) || !ed || fat === null) { scartate.push(i + hIdx + 2); return; }
    out.push([Number(lib), /^[1-9]\d*$/.test(ed) ? Number(ed) : ed, Math.round(fat)]);
  });
  if (!out.length) throw new Error("Nessuna riga valida nel file.");
  return { righe: out, scartate };
}

export default function ModuloFatturato({ token }) {
  const [dati, setDati] = useState(null);   // { nome, righe, scartate }
  const [errore, setErrore] = useState("");
  const [busy, setBusy] = useState(false);
  const [esito, setEsito] = useState(null);

  const onFile = (e) => {
    const f = e.target.files[0];
    e.target.value = "";
    if (!f) return;
    setErrore(""); setEsito(null); setDati(null);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try { setDati({ nome: f.name, ...converti(leggiRighe(evt.target.result, f.name)) }); }
      catch (err) { setErrore(err.message); }
    };
    reader.readAsArrayBuffer(f);
  };

  const creaFile = () => {
    const XLSX = window.XLSX;
    const ws = XLSX.utils.aoa_to_sheet([["Cod.Libreria", "Editore", "Fatturato"], ...dati.righe]);
    ws["!cols"] = [{ wch: 14 }, { wch: 10 }, { wch: 12 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Fatturato");
    return wb;
  };
  const nomeFile = () => `fatturato_rpn_${new Date().toISOString().slice(0, 10)}.xlsx`;

  const carica = async () => {
    if (!window.confirm(`Carico su RPN il fatturato: ${dati.righe.length.toLocaleString("it")} righe (librerie × editori)?`)) return;
    setBusy(true); setEsito(null);
    try {
      const XLSX = window.XLSX;
      const r = await fetch(`${SUPABASE_URL}/functions/v1/rpn-obiettivi`, {
        method: "POST",
        headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({ tipo: "fatturato", file_base64: XLSX.write(creaFile(), { type: "base64", bookType: "xlsx" }), nome_file: nomeFile() }),
      });
      const j = await r.json().catch(() => ({}));
      if (!r.ok) throw new Error(j.error || `Errore ${r.status}`);
      const risp = j.risposta && typeof j.risposta === "object" ? j.risposta : { message: String(j.risposta ?? "") };
      const ok = j.status >= 200 && j.status < 300 && risp.success !== false;
      setEsito({
        ok,
        msg: j.status === 401 || j.status === 403 ? "RPN non consente il caricamento con il tuo account: servono i permessi ADMIN o RESPONSABILE."
          : risp.message || risp.error || risp.detail || (ok ? "Fatturato caricato su RPN." : `RPN ha risposto con errore ${j.status}.`),
        dettaglio: ok ? null : JSON.stringify(risp).slice(0, 600),
      });
    } catch (e) { setEsito({ ok: false, msg: e.message }); }
    setBusy(false);
  };

  const stat = dati && {
    librerie: new Set(dati.righe.map(r => r[0])).size,
    editori: new Set(dati.righe.map(r => String(r[1]))).size,
    totale: dati.righe.reduce((s, r) => s + r[2], 0),
  };

  return (
    <div style={{ flex: 1, overflowY: "auto", padding: 24 }}>
      <div style={css.card}>
        <div style={{ color: T.text, fontWeight: 700, fontSize: "13px", marginBottom: 6 }}>€ Carica fatturato su RPN</div>
        <div style={{ color: T.textMid, fontSize: "11px", lineHeight: 1.6 }}>
          Carica l'export del fatturato per libreria ed editore (Dest. Merci · Editore · Fatturato): viene convertito nel formato del pulsante
          "Carica fatturato" di RPN (Cod.Libreria · Editore · Fatturato) e inviato con il tuo account RPN (servono i permessi ADMIN o RESPONSABILE).
          Puoi anche scaricare il file convertito e caricarlo a mano.
        </div>
      </div>

      <div style={{ display: "flex", gap: 10, alignItems: "center", marginBottom: 16 }}>
        <label style={{ ...css.btn("accent"), display: "inline-block" }}>
          Scegli file…
          <input type="file" accept=".csv,.txt,.xlsx,.xls" style={{ display: "none" }} onChange={onFile} />
        </label>
        {dati && <span style={{ color: T.textMid, fontSize: "12px" }}>{dati.nome}</span>}
      </div>

      {errore && <div style={{ color: T.red, fontSize: "12px", marginBottom: 12 }}>⚠ {errore}</div>}

      {dati && (
        <div style={css.card}>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 14 }}>
            <div style={css.kpi}><div style={{ color: T.textMid, fontSize: "10px" }}>RIGHE</div><div style={{ fontWeight: 700 }}>{dati.righe.length.toLocaleString("it")}</div></div>
            <div style={css.kpi}><div style={{ color: T.textMid, fontSize: "10px" }}>LIBRERIE</div><div style={{ fontWeight: 700 }}>{stat.librerie.toLocaleString("it")}</div></div>
            <div style={css.kpi}><div style={{ color: T.textMid, fontSize: "10px" }}>EDITORI</div><div style={{ fontWeight: 700 }}>{stat.editori}</div></div>
            <div style={css.kpi}><div style={{ color: T.textMid, fontSize: "10px" }}>FATTURATO TOTALE</div><div style={{ fontWeight: 700 }}>{stat.totale.toLocaleString("it")} €</div></div>
            {dati.scartate.length > 0 && <div style={{ ...css.kpi, borderColor: T.amber }}><div style={{ color: T.amber, fontSize: "10px" }}>RIGHE SCARTATE</div><div style={{ fontWeight: 700 }} title={`Righe del file: ${dati.scartate.slice(0, 30).join(", ")}`}>{dati.scartate.length}</div></div>}
          </div>

          <div style={{ color: T.textMid, fontSize: "11px", marginBottom: 6 }}>Anteprima del file per RPN (prime 10 righe)</div>
          <table style={{ borderCollapse: "collapse", marginBottom: 16 }}>
            <thead><tr>{["Cod.Libreria", "Editore", "Fatturato"].map(h => <th key={h} style={css.th}>{h}</th>)}</tr></thead>
            <tbody>{dati.righe.slice(0, 10).map((r, i) => (
              <tr key={i}><td style={css.td}>{r[0]}</td><td style={css.td}>{r[1]}</td><td style={{ ...css.td, textAlign: "right" }}>{r[2].toLocaleString("it")}</td></tr>
            ))}</tbody>
          </table>

          <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
            <button style={css.btn()} onClick={() => window.XLSX.writeFile(creaFile(), nomeFile())}>⬇ Scarica file per RPN</button>
            <button style={css.btn("accent", busy)} disabled={busy} onClick={carica}>{busy ? "Caricamento su RPN…" : "⇪ Carica su RPN"}</button>
          </div>

          {esito && (
            <div style={{ marginTop: 12, fontSize: "12px", color: esito.ok ? T.green : T.red }}>
              {esito.ok ? "✓" : "⚠"} {esito.msg}
              {esito.dettaglio && <div style={{ color: T.textDim, marginTop: 6, fontFamily: "monospace", fontSize: "10px", whiteSpace: "pre-wrap" }}>{esito.dettaglio}</div>}
            </div>
          )}
        </div>
      )}
    </div>
  );
}
