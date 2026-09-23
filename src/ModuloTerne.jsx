import { useState, useEffect, useCallback } from "react";

// ─── Aggiorna Terne (Agente – Editore – Libreria) ─────────────────────────────
// RPN è la fonte: le terne vivono nel pannello admin RPN (Terne › Associazione agenti).
//  • "Carica nuovo file": invia il file al pannello RPN (sovrascrive TUTTE le terne RPN),
//    poi riallinea Supabase rileggendo l'export da RPN.
//  • "Riallinea da RPN": scarica l'export RPN e aggiorna agente_cliente_editore
//    (usata da BookUp e GiroManager per la visibilità degli agenti).
// Edge function: rpn-terne-sync  ·  RPC: terne_stato, terne_sync_*

const SUPABASE_URL = "https://tdflwenlylhctxssatax.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InRkZmx3ZW5seWxoY3R4c3NhdGF4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYzMzgyNzYsImV4cCI6MjA5MTkxNDI3Nn0.l35qEL7LOvyYuI1McQlVqj4vbyTqmlevcmqWbTGYi2Q";
const FN_BASE = `${SUPABASE_URL}/functions/v1/rpn-terne-sync`;
const HEADERS_ATTESE = ["codice agente", "codice editore", "codice libreria"];
const CHUNK_CARICA = 10000;
const CHUNK_INSERISCI = 20000;

const T = {
  bg: "#1a2140", surface: "#212d54", border: "#2e3d6b", borderHi: "#3d4f82",
  text: "#f0f2f8", textMid: "#8b9cc8", textDim: "#4a5a8a",
  accent: "#7b9fe8", green: "#4caf7d", red: "#e05c5c", amber: "#e0a84c",
};
const css = {
  btn: (v = "default", disabled = false) => ({ padding: "7px 16px", border: `1px solid ${v === "accent" ? T.accent : v === "danger" ? T.red : T.border}`, background: v === "accent" ? T.accent : v === "danger" ? T.red + "22" : "transparent", color: v === "accent" ? "#000" : v === "danger" ? T.red : T.text, cursor: disabled ? "default" : "pointer", fontSize: "12px", fontFamily: "inherit", borderRadius: 3, fontWeight: v === "accent" ? "700" : "400", opacity: disabled ? 0.5 : 1 }),
  card: { background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, padding: 20, marginBottom: 16 },
  h: { color: T.text, fontSize: "13px", fontWeight: 700, marginBottom: 6, letterSpacing: "0.03em" },
  sub: { color: T.textMid, fontSize: "11px", lineHeight: 1.5, marginBottom: 14 },
  th: { padding: "7px 10px", textAlign: "left", color: T.textMid, fontWeight: "400", fontSize: "10px", letterSpacing: "0.08em", textTransform: "uppercase", borderBottom: `1px solid ${T.border}`, whiteSpace: "nowrap" },
  td: { padding: "7px 10px", borderBottom: `1px solid ${T.border}55`, fontSize: "12px", color: T.text, whiteSpace: "nowrap" },
  kpi: { flex: 1, background: T.bg, border: `1px solid ${T.border}`, borderRadius: 4, padding: "10px 14px" },
};

const fmt = (n) => (n == null ? "—" : Number(n).toLocaleString("it-IT"));
const fmtData = (s) => (s ? new Date(s).toLocaleString("it-IT", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" }) : "—");

async function rpc(fn, token, args = {}) {
  const r = await fetch(`${SUPABASE_URL}/rest/v1/rpc/${fn}`, {
    method: "POST",
    headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify(args),
  });
  const body = await r.json().catch(() => null);
  if (!r.ok) throw new Error(body?.message || body?.hint || `Errore ${r.status} su ${fn}`);
  return body;
}

// Legge l'xlsx e restituisce le righe [agente, editore, libreria] (senza intestazione), validando le colonne.
function leggiTerne(arrayBuffer) {
  const XLSX = window.XLSX;
  if (!XLSX) throw new Error("Libreria Excel non caricata: ricarica la pagina");
  const wb = XLSX.read(arrayBuffer, { type: "array", dense: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" });
  if (!data.length) throw new Error("Il file è vuoto");
  const head = data[0].slice(0, 3).map((h) => String(h).trim().toLowerCase());
  if (HEADERS_ATTESE.some((h, i) => head[i] !== h)) {
    throw new Error(`Intestazioni non valide. Attese: "Codice agente | Codice editore | Codice libreria" (in quest'ordine), trovate: "${data[0].slice(0, 3).join(" | ")}"`);
  }
  const righe = [];
  let scartate = 0;
  for (let i = 1; i < data.length; i++) {
    const a = String(data[i][0] ?? "").trim(), e = String(data[i][1] ?? "").trim(), l = String(data[i][2] ?? "").trim();
    if (!a && !e && !l) continue;
    if (!a || !e || !l) { scartate++; continue; }
    righe.push([a, e, l]);
  }
  const perAgente = {};
  righe.forEach(([a]) => { perAgente[a] = (perAgente[a] || 0) + 1; });
  const distinte = new Set(righe.map((r) => r.join("|"))).size;
  return { righe, scartate, distinte, perAgente };
}

export default function ModuloTerne({ token }) {
  const [stato, setStato] = useState(null);
  const [erroreStato, setErroreStato] = useState(null);
  const [busy, setBusy] = useState(false);
  const [fase, setFase] = useState(null);          // { testo, pct }
  const [esito, setEsito] = useState(null);        // { ok, testo, log }
  const [file, setFile] = useState(null);
  const [analisi, setAnalisi] = useState(null);    // risultato leggiTerne del file scelto
  const [erroreFile, setErroreFile] = useState(null);
  const [conferma, setConferma] = useState(false);

  const caricaStato = useCallback(() => {
    rpc("terne_stato", token).then((s) => { setStato(s); setErroreStato(null); }).catch((e) => setErroreStato(e.message));
  }, [token]);
  useEffect(() => { caricaStato(); }, [caricaStato]);

  // Riallinea agente_cliente_editore con l'export RPN (sempre letto da RPN, anche dopo un upload)
  const riallineaDaRpn = async (origine, fileNome) => {
    setFase({ testo: "Scarico l'export terne da RPN (circa 20 secondi)…", pct: 5 });
    const r = await fetch(`${FN_BASE}/export`, { headers: { Authorization: `Bearer ${token}`, apikey: SUPABASE_KEY } });
    if (!r.ok) { const b = await r.json().catch(() => ({})); throw new Error(b.error || b.message || `Export RPN fallito (${r.status})`); }
    const buf = await r.arrayBuffer();

    setFase({ testo: "Leggo il file…", pct: 15 });
    await new Promise((res) => setTimeout(res, 30));
    const { righe } = leggiTerne(buf);
    if (!righe.length) throw new Error("L'export RPN non contiene terne: riallineamento annullato");

    const sessione = await rpc("terne_sync_inizia", token, { p_origine: origine, p_file_nome: fileNome || null });
    try {
      // 1. staging a blocchi (3 in parallelo)
      let fatti = 0;
      const blocchi = [];
      for (let i = 0; i < righe.length; i += CHUNK_CARICA) blocchi.push(righe.slice(i, i + CHUNK_CARICA));
      let idx = 0;
      const worker = async () => {
        while (idx < blocchi.length) {
          const b = blocchi[idx++];
          await rpc("terne_sync_carica", token, { p_sessione: sessione, p_righe: b });
          fatti += b.length;
          setFase({ testo: `Carico le terne: ${fmt(fatti)} / ${fmt(righe.length)}`, pct: 20 + Math.round((fatti / righe.length) * 40) });
        }
      };
      await Promise.all([worker(), worker(), worker()]);

      // 2. verifica di sicurezza
      setFase({ testo: "Verifico i conteggi…", pct: 62 });
      const v = await rpc("terne_sync_verifica", token, { p_sessione: sessione, p_attese: righe.length });
      if (!v.ok) throw new Error(v.motivo || "Verifica non superata");

      // 3. inserisce le nuove (prima di rimuovere, così nessun agente perde visibilità a metà)
      const range = await rpc("terne_sync_range", token, { p_sessione: sessione });
      for (let da = range.min; da <= range.max; da += CHUNK_INSERISCI) {
        await rpc("terne_sync_inserisci", token, { p_sessione: sessione, p_da: da, p_a: da + CHUNK_INSERISCI - 1 });
        setFase({ testo: "Aggiungo le terne nuove…", pct: 65 + Math.round(((da - range.min) / Math.max(1, range.max - range.min)) * 20) });
      }

      // 4. rimuove quelle non più presenti su RPN
      setFase({ testo: "Rimuovo le terne non più presenti su RPN…", pct: 88 });
      for (let guard = 0; guard < 100; guard++) {
        const n = await rpc("terne_sync_rimuovi", token, { p_sessione: sessione, p_limite: 20000 });
        if (!n) break;
      }

      setFase({ testo: "Chiudo…", pct: 98 });
      return await rpc("terne_sync_chiudi", token, { p_sessione: sessione, p_esito: "ok" });
    } catch (e) {
      await rpc("terne_sync_chiudi", token, { p_sessione: sessione, p_esito: "errore", p_errore: e.message }).catch(() => {});
      throw e;
    }
  };

  const eseguiRiallinea = async () => {
    setBusy(true); setEsito(null);
    try {
      const log = await riallineaDaRpn("rpn_export");
      setEsito({ ok: true, testo: "Terne riallineate con RPN.", log });
    } catch (e) {
      setEsito({ ok: false, testo: e.message });
    } finally { setBusy(false); setFase(null); caricaStato(); }
  };

  const scegliFile = async (e) => {
    const f = e.target.files[0];
    e.target.value = "";
    setFile(null); setAnalisi(null); setErroreFile(null); setConferma(false); setEsito(null);
    if (!f) return;
    if (!/\.xlsx$/i.test(f.name)) { setErroreFile("Il file deve essere in formato .xlsx"); return; }
    try {
      const a = leggiTerne(await f.arrayBuffer());
      if (!a.righe.length) throw new Error("Il file non contiene terne");
      setFile(f); setAnalisi(a);
    } catch (err) { setErroreFile(err.message); }
  };

  const eseguiUpload = async () => {
    if (!file || !analisi || !conferma) return;
    setBusy(true); setEsito(null);
    try {
      setFase({ testo: `Carico "${file.name}" su RPN… (può richiedere qualche minuto)`, pct: 3 });
      const fd = new FormData();
      fd.append("file", file, file.name);
      const r = await fetch(`${FN_BASE}/upload`, { method: "POST", headers: { Authorization: `Bearer ${token}`, apikey: SUPABASE_KEY }, body: fd });
      const b = await r.json().catch(() => ({}));
      if (!r.ok || !b.ok) throw new Error(`RPN ha rifiutato il file: ${b.error || b.message || r.status}`);
      const msgRpn = (b.messages?.success || []).join(" ");
      const log = await riallineaDaRpn("upload_file", file.name);
      setEsito({ ok: true, testo: `File caricato su RPN${msgRpn ? ` (${msgRpn})` : ""} e terne riallineate.`, log });
      setFile(null); setAnalisi(null); setConferma(false);
    } catch (e) {
      setEsito({ ok: false, testo: e.message });
    } finally { setBusy(false); setFase(null); caricaStato(); }
  };

  const attuali = stato?.terne ?? null;
  const calo = analisi && attuali ? (analisi.distinte - attuali) / attuali : 0;

  return (
    <div style={{ flex: 1, overflowY: "auto", padding: 24, maxWidth: 980 }}>
      {/* Stato attuale */}
      <div style={css.card}>
        <div style={css.h}>👥 Terne Agente – Editore – Libreria</div>
        <div style={css.sub}>
          Le terne decidono cosa vede ogni agente in BookUp e GiroManager. La fonte è RPN (pannello admin › Terne › Associazione agenti):
          da qui puoi caricare un nuovo file su RPN oppure riallineare i dati leggendoli da RPN.
        </div>
        {erroreStato && <div style={{ color: T.red, fontSize: 12 }}>⚠ {erroreStato}</div>}
        <div style={{ display: "flex", gap: 10 }}>
          <div style={css.kpi}><div style={{ color: T.textMid, fontSize: 10, textTransform: "uppercase", letterSpacing: "0.08em" }}>Terne attive</div><div style={{ color: T.text, fontSize: 20, fontWeight: 700 }}>{fmt(attuali)}</div></div>
          <div style={css.kpi}><div style={{ color: T.textMid, fontSize: 10, textTransform: "uppercase", letterSpacing: "0.08em" }}>Agenti</div><div style={{ color: T.text, fontSize: 20, fontWeight: 700 }}>{fmt(stato?.agenti)}</div></div>
          <div style={css.kpi}><div style={{ color: T.textMid, fontSize: 10, textTransform: "uppercase", letterSpacing: "0.08em" }}>Ultimo allineamento</div><div style={{ color: T.text, fontSize: 14, fontWeight: 700, marginTop: 4 }}>{fmtData(stato?.storico?.find((s) => s.esito === "ok")?.finished_at)}</div></div>
        </div>
      </div>

      {/* Avanzamento / esito */}
      {fase && (
        <div style={{ ...css.card, borderColor: T.accent }}>
          <div style={{ color: T.text, fontSize: 12, marginBottom: 8 }}><span style={{ display: "inline-block", animation: "gm-spin 1s linear infinite", marginRight: 8 }}>⟳</span>{fase.testo}</div>
          <div style={{ height: 6, background: T.bg, borderRadius: 3 }}><div style={{ height: 6, width: `${fase.pct}%`, background: T.accent, borderRadius: 3, transition: "width .3s" }} /></div>
          <div style={{ color: T.textDim, fontSize: 11, marginTop: 8 }}>Non chiudere la pagina fino al termine.</div>
        </div>
      )}
      {esito && (
        <div style={{ ...css.card, borderColor: esito.ok ? T.green : T.red, background: (esito.ok ? T.green : T.red) + "18", position: "relative" }}>
          <span onClick={() => setEsito(null)} style={{ position: "absolute", top: 10, right: 14, cursor: "pointer", color: T.textMid }}>✕</span>
          <div style={{ color: esito.ok ? T.green : T.red, fontWeight: 700, fontSize: 13, marginBottom: esito.log ? 8 : 0 }}>{esito.ok ? "✓ " : "✕ "}{esito.testo}</div>
          {esito.log && (
            <div style={{ color: T.text, fontSize: 12 }}>
              Terne su RPN: <b>{fmt(esito.log.righe_rpn)}</b> · aggiunte <b style={{ color: T.green }}>+{fmt(esito.log.aggiunte)}</b> · rimosse <b style={{ color: T.red }}>−{fmt(esito.log.rimosse)}</b> · totale ora <b>{fmt(esito.log.righe_dopo)}</b>
            </div>
          )}
        </div>
      )}

      {/* Carica nuovo file */}
      <div style={css.card}>
        <div style={css.h}>📂 Carica nuovo file terne su RPN</div>
        <div style={css.sub}>
          File .xlsx con tre colonne, in quest'ordine: <b>Codice agente · Codice editore · Codice libreria</b> (stesso formato dell'export RPN).
          <br /><span style={{ color: T.amber }}>⚠ Il caricamento SOSTITUISCE tutte le terne presenti su RPN: il file deve contenere l'elenco completo, non solo le modifiche.</span>
        </div>
        <input type="file" accept=".xlsx" id="terne-file" style={{ display: "none" }} onChange={scegliFile} disabled={busy} />
        <label htmlFor="terne-file" style={{ ...css.btn("default", busy), display: "inline-block" }}>Scegli file .xlsx</label>
        {erroreFile && <div style={{ color: T.red, fontSize: 12, marginTop: 10 }}>⚠ {erroreFile}</div>}

        {analisi && (
          <div style={{ marginTop: 16 }}>
            <div style={{ color: T.text, fontSize: 12, marginBottom: 10 }}>
              <b>{file.name}</b> — {fmt(analisi.righe.length)} righe ({fmt(analisi.distinte)} terne distinte), {fmt(Object.keys(analisi.perAgente).length)} agenti
              {analisi.scartate > 0 && <span style={{ color: T.amber }}> · {fmt(analisi.scartate)} righe incomplete scartate</span>}
            </div>
            {attuali != null && (
              <div style={{ fontSize: 12, marginBottom: 10, color: calo < -0.1 ? T.amber : T.textMid }}>
                Oggi le terne sono {fmt(attuali)}: il file ne porta {analisi.distinte >= attuali ? "+" : ""}{fmt(analisi.distinte - attuali)} ({(calo * 100).toFixed(1)}%).
                {calo < -0.1 && " Controlla che il file sia completo."}
                {calo < -0.5 && <span style={{ color: T.red }}> Calo oltre il 50%: il riallineamento verrà bloccato per sicurezza.</span>}
              </div>
            )}
            <div style={{ maxHeight: 180, overflowY: "auto", border: `1px solid ${T.border}`, borderRadius: 4, marginBottom: 12 }}>
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead><tr><th style={css.th}>Codice agente</th><th style={{ ...css.th, textAlign: "right" }}>Terne nel file</th></tr></thead>
                <tbody>
                  {Object.entries(analisi.perAgente).sort((a, b) => b[1] - a[1]).map(([a, n]) => (
                    <tr key={a}><td style={css.td}>{a}</td><td style={{ ...css.td, textAlign: "right" }}>{fmt(n)}</td></tr>
                  ))}
                </tbody>
              </table>
            </div>
            <label style={{ display: "flex", alignItems: "center", gap: 8, color: T.text, fontSize: 12, marginBottom: 12, cursor: "pointer" }}>
              <input type="checkbox" checked={conferma} onChange={(e) => setConferma(e.target.checked)} disabled={busy} />
              Ho capito: il file sostituisce tutte le terne su RPN
            </label>
            <div style={{ display: "flex", gap: 8 }}>
              <button style={css.btn("accent", busy || !conferma)} disabled={busy || !conferma} onClick={eseguiUpload}>⬆ Carica su RPN e aggiorna</button>
              <button style={css.btn("default", busy)} disabled={busy} onClick={() => { setFile(null); setAnalisi(null); setConferma(false); }}>Annulla</button>
            </div>
          </div>
        )}
      </div>

      {/* Riallinea */}
      <div style={css.card}>
        <div style={css.h}>🔄 Riallinea da RPN</div>
        <div style={css.sub}>Legge le terne attuali da RPN e aggiorna BookUp/GiroManager. Usalo se le terne sono state modificate direttamente nel pannello RPN. Non modifica nulla su RPN.</div>
        <button style={css.btn("default", busy)} disabled={busy} onClick={eseguiRiallinea}>🔄 Aggiorna terne da RPN</button>
      </div>

      {/* Storico */}
      {stato?.storico?.length > 0 && (
        <div style={css.card}>
          <div style={css.h}>Storico allineamenti</div>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr>{["Data", "Origine", "Esito", "Terne RPN", "Aggiunte", "Rimosse", "Totale"].map((h) => <th key={h} style={css.th}>{h}</th>)}</tr></thead>
            <tbody>
              {stato.storico.map((s) => (
                <tr key={s.id} title={s.errore || ""}>
                  <td style={css.td}>{fmtData(s.started_at)}</td>
                  <td style={css.td}>{s.origine === "upload_file" ? `File ${s.file_nome || ""}` : "Export RPN"}</td>
                  <td style={{ ...css.td, color: s.esito === "ok" ? T.green : s.esito === "errore" ? T.red : T.amber }}>{s.esito === "ok" ? "✓ ok" : s.esito === "errore" ? "✕ errore" : "in corso"}</td>
                  <td style={css.td}>{fmt(s.righe_rpn)}</td>
                  <td style={{ ...css.td, color: T.green }}>+{fmt(s.aggiunte)}</td>
                  <td style={{ ...css.td, color: T.red }}>−{fmt(s.rimosse)}</td>
                  <td style={css.td}>{fmt(s.righe_dopo)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
