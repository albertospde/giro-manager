import { useState, useEffect, useCallback } from "react";
import { tema } from "./tema.js";

// ─── SFTP Messaggerie ────────────────────────────────────────────────────────
// Impostazioni e stato del download automatico dei file dall'SFTP di Messaggerie.
// • La password si salva con l'RPC sftp_config_salva nel Vault di Supabase: non torna mai al browser.
// • Il download lo fa la edge function sftp-messaggerie: ogni mattina dalle 5:30 (pg_cron) oppure con "Scarica ora".
// • I file finiscono nel bucket privato "sftp-messaggerie", una cartella per giorno.

const SUPABASE_URL = "https://tdflwenlylhctxssatax.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InRkZmx3ZW5seWxoY3R4c3NhdGF4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYzMzgyNzYsImV4cCI6MjA5MTkxNDI3Nn0.l35qEL7LOvyYuI1McQlVqj4vbyTqmlevcmqWbTGYi2Q";
const FN_URL = `${SUPABASE_URL}/functions/v1/sftp-messaggerie`;
const BUCKET = "sftp-messaggerie";

const T = tema({
  bg: "#1a2140", surface: "#212d54", border: "#2e3d6b", borderHi: "#3d4f82",
  text: "#f0f2f8", textMid: "#8b9cc8", textDim: "#4a5a8a",
  accent: "#7b9fe8", green: "#4caf7d", red: "#e05c5c", amber: "#e0a84c",
});
const css = {
  btn: (v = "default", disabled = false) => ({ padding: "7px 16px", border: `1px solid ${v === "accent" ? T.accent : T.border}`, background: v === "accent" ? T.accent : "transparent", color: v === "accent" ? "#000" : T.text, cursor: disabled ? "default" : "pointer", fontSize: "12px", fontFamily: "inherit", borderRadius: 3, fontWeight: v === "accent" ? "700" : "400", opacity: disabled ? 0.5 : 1 }),
  card: { background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, padding: 20, marginBottom: 16 },
  h: { color: T.text, fontSize: "13px", fontWeight: 700, marginBottom: 6, letterSpacing: "0.03em" },
  sub: { color: T.textMid, fontSize: "11px", lineHeight: 1.5, marginBottom: 14 },
  label: { display: "flex", flexDirection: "column", gap: 4, color: T.textMid, fontSize: "11px" },
  input: { background: T.bg, border: `1px solid ${T.borderHi}`, color: T.text, padding: "6px 8px", fontSize: "12px", fontFamily: "inherit", borderRadius: 3 },
  th: { padding: "7px 10px", textAlign: "left", color: T.textMid, fontWeight: "400", fontSize: "10px", letterSpacing: "0.08em", textTransform: "uppercase", borderBottom: `1px solid ${T.border}`, whiteSpace: "nowrap" },
  td: { padding: "7px 10px", borderBottom: `1px solid ${T.border}55`, fontSize: "12px", color: T.text, whiteSpace: "nowrap" },
};

const fmtData = (s) => (s ? new Date(s).toLocaleString("it-IT", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" }) : "—");
const fmtDim = (n) => (n == null ? "—" : n < 1048576 ? `${(n / 1024).toFixed(1)} KB` : `${(n / 1048576).toFixed(1)} MB`);

const headers = (token, extra = {}) => ({ apikey: SUPABASE_KEY, Authorization: `Bearer ${token}`, ...extra });

async function leggi(path, token) {
  const r = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, { headers: headers(token) });
  const body = await r.json().catch(() => null);
  if (!r.ok) throw new Error(body?.message || `Errore ${r.status}`);
  return body;
}

async function chiamaFunzione(azione, token) {
  const r = await fetch(FN_URL, { method: "POST", headers: headers(token, { "Content-Type": "application/json" }), body: JSON.stringify({ azione }) });
  const body = await r.json().catch(() => null);
  if (!r.ok) {
    // 546/504: il runtime ha fermato la funzione (limite di CPU o di tempo) a metà di un file
    const interrotto = !body?.error && (r.status === 546 || r.status === 504);
    const err = new Error(body?.error || body?.message || (interrotto ? "Il download di un file è stato interrotto perché troppo pesante." : `Errore ${r.status}`));
    err.interrotto = interrotto;
    throw err;
  }
  return body;
}

export default function ModuloSftp({ token }) {
  const [form, setForm] = useState(null);
  const [passwordSalvata, setPasswordSalvata] = useState(false);
  const [password, setPassword] = useState("");
  const [impronta, setImpronta] = useState("");
  const [file, setFile] = useState([]);
  const [esecuzioni, setEsecuzioni] = useState([]);
  const [elencoRemoto, setElencoRemoto] = useState(null);
  const [busy, setBusy] = useState(null); // "salva" | "prova" | "scarica"
  const [esito, setEsito] = useState(null); // { ok, testo }
  const [errore, setErrore] = useState(null);

  const carica = useCallback(async () => {
    try {
      const [cfg, f, e] = await Promise.all([
        leggi("sftp_config?id=eq.1&select=*", token),
        leggi("sftp_file?select=*&cancellato_il=is.null&order=scaricato_il.desc&limit=50", token),
        leggi("sftp_esecuzioni?select=*&order=quando.desc&limit=20", token),
      ]);
      if (!cfg?.length) throw new Error("Pagina riservata agli amministratori PDE.");
      const c = cfg[0];
      setForm((prev) => prev || { host: c.host, porta: c.porta, utente: c.utente, cartella_remota: c.cartella_remota, giorni_indietro: c.giorni_indietro, automatico: c.automatico, filtro_nome: c.filtro_nome ?? "" });
      setPasswordSalvata(!!c.password_secret_id);
      setImpronta(c.impronta_server);
      setFile(f);
      setEsecuzioni(e);
      setErrore(null);
    } catch (err) {
      setErrore(err.message);
    }
  }, [token]);

  useEffect(() => { carica(); }, [carica]);

  const set = (k) => (ev) => setForm((f) => ({ ...f, [k]: ev.target.type === "checkbox" ? ev.target.checked : ev.target.value }));

  const salva = async () => {
    setBusy("salva"); setEsito(null);
    try {
      const r = await fetch(`${SUPABASE_URL}/rest/v1/rpc/sftp_config_salva`, {
        method: "POST", headers: headers(token, { "Content-Type": "application/json" }),
        body: JSON.stringify({
          p_host: form.host, p_porta: Number(form.porta) || 22, p_utente: form.utente, p_password: password || null,
          p_cartella_remota: form.cartella_remota, p_giorni_indietro: Number(form.giorni_indietro) || 7, p_automatico: form.automatico, p_filtro_nome: form.filtro_nome,
        }),
      });
      if (!r.ok) { const b = await r.json().catch(() => null); throw new Error(b?.message || `Errore ${r.status}`); }
      setPassword("");
      setEsito({ ok: true, testo: "Impostazioni salvate." });
      await carica();
    } catch (err) {
      setEsito({ ok: false, testo: err.message });
    }
    setBusy(null);
  };

  const prova = async () => {
    setBusy("prova"); setEsito(null); setElencoRemoto(null);
    try {
      const r = await chiamaFunzione("prova", token);
      setElencoRemoto(r);
      setEsito({ ok: true, testo: `Connessione riuscita: ${r.totale} file in ${form.cartella_remota}.` });
    } catch (err) {
      setEsito({ ok: false, testo: err.message });
    }
    await carica();
    setBusy(null);
  };

  // La funzione scarica un file per chiamata: si richiama finché ne restano.
  // Un'interruzione (file troppo pesante) si riprova: dopo 2 la funzione salta quel file da sola.
  const scarica = async () => {
    setBusy("scarica"); setEsito(null);
    const fatti = [];
    let saltati = [], interruzioni = 0, errore = null;
    for (let giro = 0; giro < 60; giro++) {
      try {
        const r = await chiamaFunzione("scarica", token);
        fatti.push(...r.scaricati.map((s) => s.nome));
        saltati = r.saltati || [];
        interruzioni = 0;
        if (!r.scaricati.length || !r.rimanenti) break;
        setEsito({ ok: true, testo: `Scaricati ${fatti.length}, ne restano ${r.rimanenti}…` });
      } catch (err) {
        if (err.interrotto && ++interruzioni <= 2) continue;
        errore = err.message;
        break;
      }
    }
    const nota = saltati.length ? ` Saltati perché troppo pesanti: ${saltati.join(", ")}.` : "";
    setEsito(errore
      ? { ok: false, testo: fatti.length ? `Scaricati ${fatti.length} file, poi: ${errore}` : errore }
      : { ok: true, testo: fatti.length ? `Scaricati ${fatti.length} file: ${fatti.join(", ")}.${nota}` : `Nessun file nuovo da scaricare.${nota}` });
    await carica();
    setBusy(null);
  };

  const apriFile = async (f) => {
    try {
      const r = await fetch(`${SUPABASE_URL}/storage/v1/object/sign/${BUCKET}/${f.percorso_storage.split("/").map(encodeURIComponent).join("/")}`, {
        method: "POST", headers: headers(token, { "Content-Type": "application/json" }), body: JSON.stringify({ expiresIn: 300 }),
      });
      const b = await r.json().catch(() => null);
      if (!r.ok || !b?.signedURL) throw new Error(b?.message || `Errore ${r.status}`);
      const a = document.createElement("a");
      a.href = `${SUPABASE_URL}/storage/v1${b.signedURL}&download=${encodeURIComponent(f.nome_remoto)}`;
      a.click();
    } catch (err) {
      setEsito({ ok: false, testo: `Download di ${f.nome_remoto}: ${err.message}` });
    }
  };

  if (errore && !form) return <div style={{ padding: 24, color: T.red, fontSize: 12 }}>⚠ {errore}</div>;
  if (!form) return <div style={{ padding: 24, color: T.textMid, fontSize: 12 }}>Caricamento…</div>;

  const ultimaAuto = esecuzioni.find((e) => e.origine === "automatico");

  return (
    <div style={{ flex: 1, overflowY: "auto", padding: 24, maxWidth: 980 }}>
      <div style={css.card}>
        <div style={css.h}>🔐 Connessione SFTP Messaggerie</div>
        <div style={css.sub}>
          La password viene salvata cifrata nel Vault di Supabase e non viene più mostrata: lasciala vuota per non cambiarla.
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12 }}>
          <label style={css.label}>Host<input style={css.input} value={form.host} onChange={set("host")} autoComplete="off" /></label>
          <label style={css.label}>Porta<input style={css.input} type="number" value={form.porta} onChange={set("porta")} /></label>
          <label style={css.label}>Nome utente<input style={css.input} value={form.utente} onChange={set("utente")} autoComplete="off" /></label>
          <label style={css.label}>Password
            <input style={css.input} type="password" value={password} onChange={(e) => setPassword(e.target.value)} autoComplete="new-password"
              placeholder={passwordSalvata ? "•••••••• (salvata)" : "non ancora inserita"} />
          </label>
          <label style={css.label}>Cartella sul server<input style={css.input} value={form.cartella_remota} onChange={set("cartella_remota")} /></label>
          <label style={css.label}>Considera file degli ultimi (giorni)<input style={css.input} type="number" min="1" value={form.giorni_indietro} onChange={set("giorni_indietro")} /></label>
          <label style={css.label}>Scarica solo i file che iniziano con<input style={css.input} value={form.filtro_nome} onChange={set("filtro_nome")} placeholder="es. PPSV (vuoto = tutti)" /></label>
        </div>
        <label style={{ display: "flex", alignItems: "center", gap: 8, color: T.text, fontSize: 12, marginTop: 14 }}>
          <input type="checkbox" checked={form.automatico} onChange={set("automatico")} />
          Download automatico ogni mattina dalle 5:30 (se il file non c'è ancora riprova ogni 15 minuti fino alle 6:30)
        </label>
        {impronta && <div style={{ color: T.textDim, fontSize: 11, marginTop: 10 }}>Impronta server: {impronta}</div>}
        <div style={{ display: "flex", gap: 8, marginTop: 16, flexWrap: "wrap" }}>
          <button style={css.btn("accent", !!busy)} disabled={!!busy} onClick={salva}>{busy === "salva" ? "Salvataggio…" : "Salva"}</button>
          <button style={css.btn("default", !!busy)} disabled={!!busy} onClick={prova}>{busy === "prova" ? "Connessione…" : "Prova connessione"}</button>
          <button style={css.btn("default", !!busy)} disabled={!!busy} onClick={scarica}>{busy === "scarica" ? "Download in corso…" : "Scarica ora"}</button>
        </div>
      </div>

      {esito && (
        <div style={{ ...css.card, borderColor: esito.ok ? T.green : T.red, background: (esito.ok ? T.green : T.red) + "18", position: "relative" }}>
          <span onClick={() => setEsito(null)} style={{ position: "absolute", top: 10, right: 14, cursor: "pointer", color: T.textMid }}>✕</span>
          <div style={{ color: esito.ok ? T.green : T.red, fontWeight: 700, fontSize: 12, paddingRight: 20 }}>{esito.ok ? "✓ " : "✕ "}{esito.testo}</div>
        </div>
      )}

      {elencoRemoto && (
        <div style={css.card}>
          <div style={css.h}>📂 File sul server ({elencoRemoto.totale})</div>
          <div style={{ overflowX: "auto", maxHeight: 320, overflowY: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead><tr><th style={css.th}>Nome</th><th style={{ ...css.th, textAlign: "right" }}>Dimensione</th><th style={css.th}>Modificato</th></tr></thead>
              <tbody>{elencoRemoto.file.map((f) => (
                <tr key={f.nome} style={{ opacity: f.filtro ? 1 : 0.45 }}><td style={css.td}>{f.nome}{f.filtro ? "" : " · escluso dal filtro"}</td><td style={{ ...css.td, textAlign: "right" }}>{fmtDim(f.dimensione)}</td><td style={css.td}>{fmtData(f.modificato)}</td></tr>
              ))}</tbody>
            </table>
          </div>
        </div>
      )}

      <div style={css.card}>
        <div style={css.h}>⬇ File scaricati</div>
        <div style={css.sub}>
          Ultimo download automatico: {ultimaAuto ? `${fmtData(ultimaAuto.quando)} · ${ultimaAuto.ok ? ultimaAuto.messaggio : "errore: " + ultimaAuto.messaggio}` : "non ancora eseguito"}
        </div>
        {file.length === 0 ? <div style={{ color: T.textDim, fontSize: 12 }}>Nessun file scaricato finora.</div> : (
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead><tr><th style={css.th}>File</th><th style={{ ...css.th, textAlign: "right" }}>Dimensione</th><th style={css.th}>Sul server dal</th><th style={css.th}>Scaricato il</th><th style={css.th}>Origine</th><th style={css.th}></th></tr></thead>
              <tbody>{file.map((f) => (
                <tr key={f.id}>
                  <td style={css.td}>{f.nome_remoto}</td>
                  <td style={{ ...css.td, textAlign: "right" }}>{fmtDim(f.dimensione)}</td>
                  <td style={css.td}>{fmtData(f.modificato_remoto)}</td>
                  <td style={css.td}>{fmtData(f.scaricato_il)}</td>
                  <td style={css.td}>{f.origine}</td>
                  <td style={css.td}><button style={{ ...css.btn(), padding: "3px 10px" }} onClick={() => apriFile(f)}>↓ Scarica</button></td>
                </tr>
              ))}</tbody>
            </table>
          </div>
        )}
      </div>

      <div style={css.card}>
        <div style={css.h}>🕑 Registro</div>
        {esecuzioni.length === 0 ? <div style={{ color: T.textDim, fontSize: 12 }}>Nessuna esecuzione registrata.</div> : (
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr><th style={css.th}>Quando</th><th style={css.th}>Tipo</th><th style={css.th}>Esito</th></tr></thead>
            <tbody>{esecuzioni.map((e) => (
              <tr key={e.id}>
                <td style={css.td}>{fmtData(e.quando)}</td>
                <td style={css.td}>{e.origine}</td>
                <td style={{ ...css.td, whiteSpace: "normal", color: e.ok ? T.text : T.red }}>{e.ok ? "" : "✕ "}{e.messaggio}</td>
              </tr>
            ))}</tbody>
          </table>
        )}
      </div>
    </div>
  );
}
