import { useState, useEffect, useMemo, useCallback } from "react";

// ─── Pubblica su RPN ─────────────────────────────────────────────────────────
// Crea/aggiorna su RPN le cedole di GiroManager: giro e cedola (se mancano), titoli agganciati
// nell'ordine di "Giri e Cedole", titoli tolti da GiroManager rimossi anche da RPN.
// L'attivazione (visibilità agli agenti) è un passaggio separato.
// I titoli non ancora in anagrafica RPN vengono ritentati ogni notte (job pg_cron → rpn-cedola-sync/cron).
// Edge function: rpn-cedola-sync  ·  Tabella stato: rpn_cedole_sync

const SUPABASE_URL = "https://tdflwenlylhctxssatax.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InRkZmx3ZW5seWxoY3R4c3NhdGF4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYzMzgyNzYsImV4cCI6MjA5MTkxNDI3Nn0.l35qEL7LOvyYuI1McQlVqj4vbyTqmlevcmqWbTGYi2Q";
const FN = `${SUPABASE_URL}/functions/v1/rpn-cedola-sync`;

const T = {
  bg: "#1a2140", surface: "#212d54", border: "#2e3d6b", borderHi: "#3d4f82",
  text: "#f0f2f8", textMid: "#8b9cc8", textDim: "#4a5a8a",
  accent: "#7b9fe8", green: "#4caf7d", red: "#e05c5c", amber: "#e0a84c",
};
const css = {
  btn: (v = "default", disabled = false) => ({ padding: "7px 14px", border: `1px solid ${v === "accent" ? T.accent : v === "green" ? T.green : v === "danger" ? T.red : T.border}`, background: v === "accent" ? T.accent : v === "green" ? T.green : "transparent", color: v === "accent" || v === "green" ? "#000" : v === "danger" ? T.red : T.text, cursor: disabled ? "default" : "pointer", fontSize: "12px", fontFamily: "inherit", borderRadius: 3, fontWeight: v === "accent" || v === "green" ? 700 : 400, opacity: disabled ? 0.5 : 1, whiteSpace: "nowrap" }),
  card: { background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, padding: 18, marginBottom: 16 },
  h: { color: T.text, fontSize: "13px", fontWeight: 700, marginBottom: 6 },
  sub: { color: T.textMid, fontSize: "11px", lineHeight: 1.5 },
  th: { padding: "7px 10px", textAlign: "left", color: T.textMid, fontWeight: 400, fontSize: "10px", letterSpacing: "0.08em", textTransform: "uppercase", borderBottom: `1px solid ${T.border}`, whiteSpace: "nowrap" },
  td: { padding: "7px 10px", borderBottom: `1px solid ${T.border}55`, fontSize: "12px", color: T.text, whiteSpace: "nowrap" },
  input: { background: T.bg, border: `1px solid ${T.borderHi}`, color: T.text, padding: "6px 8px", fontSize: "12px", fontFamily: "inherit", borderRadius: 3 },
  kpi: { flex: 1, minWidth: 120, background: T.bg, border: `1px solid ${T.border}`, borderRadius: 4, padding: "10px 14px" },
};

const STATI = { CRE: ["Creata", T.amber], DAT: ["Da attivare", T.amber], ATT: ["Attiva", T.green], CHI: ["Chiusa", T.textMid] };
const fmtData = (s) => (s ? new Date(s).toLocaleString("it-IT", { day: "2-digit", month: "2-digit", year: "2-digit", hour: "2-digit", minute: "2-digit" }) : "—");

async function chiama(route, token, body) {
  const r = await fetch(`${FN}/${route}`, {
    method: "POST",
    headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok || j.error) throw new Error(j.error || `Errore ${r.status}`);
  return j;
}

function StatoBadge({ stato }) {
  if (!stato) return <span style={{ color: T.textDim, fontSize: "11px" }}>non su RPN</span>;
  const [l, c] = STATI[stato] || [stato, T.textMid];
  return <span style={{ color: c, fontSize: "11px", fontWeight: 700 }}>● {l}</span>;
}

function ListaTitoli({ titolo, righe, colore, aperta }) {
  const [open, setOpen] = useState(!!aperta);
  if (!righe?.length) return null;
  const copia = () => navigator.clipboard?.writeText(righe.map(r => `${r.ean}\t${r.titolo ?? ""}\t${r.editore ?? ""}`).join("\n"));
  return (
    <div style={{ marginTop: 10 }}>
      <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
        <button style={{ ...css.btn(), borderColor: colore, color: colore }} onClick={() => setOpen(o => !o)}>{open ? "▾" : "▸"} {titolo} ({righe.length})</button>
        {open && <button style={css.btn()} onClick={copia}>Copia elenco</button>}
      </div>
      {open && (
        <div style={{ maxHeight: 240, overflow: "auto", marginTop: 6, border: `1px solid ${T.border}`, borderRadius: 4 }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <tbody>{righe.map(r => (
              <tr key={r.ean}><td style={{ ...css.td, fontFamily: "monospace", color: T.textMid }}>{r.ean}</td><td style={{ ...css.td, whiteSpace: "normal" }}>{r.titolo}</td><td style={{ ...css.td, color: T.textMid }}>{r.editore}</td></tr>
            ))}</tbody>
          </table>
        </div>
      )}
    </div>
  );
}

export default function ModuloPubblicaRpn({ token, titoli }) {
  const [syncRows, setSyncRows] = useState({});
  const [giroSel, setGiroSel] = useState("");
  const [sel, setSel] = useState(null);          // n_cedola aperta nel dettaglio
  const [preview, setPreview] = useState(null);
  const [date, setDate] = useState({ data_inizio: "", data_fine: "" });
  const [busy, setBusy] = useState("");
  const [errore, setErrore] = useState("");
  const [esito, setEsito] = useState(null);
  const [batch, setBatch] = useState(null);       // avanzamento "pubblica tutto il giro"

  const caricaStato = useCallback(async () => {
    const r = await fetch(`${SUPABASE_URL}/rest/v1/rpn_cedole_sync?select=*`, { headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${token}` } });
    const rows = r.ok ? await r.json() : [];
    setSyncRows(Object.fromEntries(rows.map(x => [x.n_cedola, x])));
  }, [token]);
  useEffect(() => { caricaStato(); }, [caricaStato]);

  // Cedole di GiroManager raggruppate per giro (giro_label "5 2026" oppure "EXTRA")
  const cedolePerGiro = useMemo(() => {
    const g = {};
    (titoli || []).forEach(t => {
      if (!t.n_cedola) return;
      const giro = t.giro_label || "—";
      g[giro] ??= {};
      g[giro][t.n_cedola] = (g[giro][t.n_cedola] || 0) + 1;
    });
    return g;
  }, [titoli]);
  const giri = useMemo(() => Object.keys(cedolePerGiro).sort((a, b) => {
    if (a === "EXTRA") return 1; if (b === "EXTRA") return -1;
    const [na, ya] = a.split(" ").map(Number), [nb, yb] = b.split(" ").map(Number);
    return (yb - ya) || (nb - na);
  }), [cedolePerGiro]);
  useEffect(() => { if (!giroSel && giri.length) setGiroSel(giri[0]); }, [giri, giroSel]);
  const cedole = useMemo(() => Object.entries(cedolePerGiro[giroSel] || {}).sort((a, b) => a[0].localeCompare(b[0])), [cedolePerGiro, giroSel]);

  const apri = async (nCedola) => {
    setSel(nCedola); setPreview(null); setEsito(null); setErrore(""); setBusy("preview");
    try {
      const p = await chiama("preview", token, { n_cedola: nCedola });
      setPreview(p);
      setDate({ data_inizio: p.date?.data_inizio || "", data_fine: p.date?.data_fine || "" });
    } catch (e) { setErrore(e.message); }
    setBusy("");
  };

  const pubblica = async () => {
    const p = preview;
    const cosa = [
      !p.cedola.esiste && `creare la cedola "${p.n_cedola}"${p.giro && !p.giro.esiste ? ` e il giro "${p.giro.nome}"` : ""}`,
      p.da_aggiungere.length && `agganciare ${p.da_aggiungere.length} titoli`,
      p.da_rimuovere.length && `TOGLIERE ${p.da_rimuovere.length} titoli non più in GiroManager`,
      p.ordine_da_aggiornare && "riordinare i titoli",
    ].filter(Boolean);
    if (!cosa.length) { setEsito({ info: "Già allineata: nulla da fare." }); return; }
    if (!window.confirm(`Su RPN sto per:\n• ${cosa.join("\n• ")}\n\nProcedo?`)) return;
    setBusy("apply"); setErrore("");
    try {
      const r = await chiama("apply", token, { n_cedola: p.n_cedola, ...date });
      setEsito(r.esito); setPreview({ ...r.stato, account: p.account });
      await caricaStato();
    } catch (e) { setErrore(e.message); }
    setBusy("");
  };

  const attiva = async () => {
    if (!window.confirm(`Attivo "${preview.n_cedola}" su RPN? Da questo momento la vedono gli agenti.`)) return;
    setBusy("activate"); setErrore("");
    try {
      const r = await chiama("activate", token, { n_cedola: preview.n_cedola });
      setPreview(p => ({ ...p, cedola: { ...p.cedola, stato: r.stato } }));
      setEsito({ info: `Cedola attivata (stato RPN: ${STATI[r.stato]?.[0] ?? r.stato}).` });
      await caricaStato();
    } catch (e) { setErrore(e.message); }
    setBusy("");
  };

  // Pubblica in sequenza tutte le cedole del giro selezionato (non chiuse)
  const pubblicaGiro = async () => {
    const lista = cedole.map(([n]) => n).filter(n => syncRows[n]?.rpn_stato !== "CHI");
    if (!lista.length) return;
    if (!window.confirm(`Pubblico/aggiorno su RPN ${lista.length} cedole del giro ${giroSel}:\n${lista.join("\n")}\n\nVerranno creati giro e cedole mancanti, agganciati/riordinati i titoli e tolti quelli non più in GiroManager. Procedo?`)) return;
    const out = [];
    setBatch({ fatto: 0, tot: lista.length, out });
    for (const n of lista) {
      try {
        const r = await chiama("apply", token, { n_cedola: n });
        out.push({ n, ok: true, esito: r.esito, mancanti: r.stato.mancanti.length });
      } catch (e) { out.push({ n, ok: false, errore: e.message }); }
      setBatch({ fatto: out.length, tot: lista.length, out: [...out] });
    }
    await caricaStato();
  };

  const p = preview;
  const chiusa = p?.cedola?.stato === "CHI";

  return (
    <div style={{ flex: 1, overflowY: "auto", padding: 24 }}>
      <div style={css.card}>
        <div style={css.h}>⇪ Crea cedola su RPN</div>
        <div style={css.sub}>
          Crea su RPN giro e cedola se mancano, aggancia i titoli nell'ordine di "Giri e Cedole" e toglie quelli non più presenti in GiroManager.
          I titoli che RPN non ha ancora in anagrafica vengono ritentati in automatico ogni notte. L'attivazione per gli agenti è un passaggio separato.
          Note, top 100, obiettivi e gemelli restano solo in GiroManager (RPN non li gestisce).
        </div>
      </div>

      <div style={{ display: "flex", gap: 16, alignItems: "flex-start", flexWrap: "wrap" }}>
        {/* Elenco cedole */}
        <div style={{ ...css.card, flex: "1 1 420px", minWidth: 380 }}>
          <div style={{ display: "flex", gap: 10, alignItems: "center", marginBottom: 12 }}>
            <label style={{ color: T.textMid, fontSize: "12px" }}>Giro&nbsp;
              <select style={css.input} value={giroSel} onChange={e => { setGiroSel(e.target.value); setSel(null); setPreview(null); setBatch(null); }}>
                {giri.map(g => <option key={g} value={g}>{g === "EXTRA" ? "Cedole extra" : `Giro ${g}`}</option>)}
              </select>
            </label>
            {giroSel && giroSel !== "EXTRA" && <button style={{ ...css.btn("accent", !!batch && batch.fatto < batch.tot), marginLeft: "auto" }} disabled={!!batch && batch.fatto < batch.tot} onClick={pubblicaGiro}>Pubblica tutto il giro</button>}
          </div>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr>{["Cedola", "Titoli GM", "Su RPN", "Mancanti", "Ultimo sync", ""].map(h => <th key={h} style={css.th}>{h}</th>)}</tr></thead>
            <tbody>
              {cedole.map(([n, cnt]) => {
                const s = syncRows[n];
                const manc = Array.isArray(s?.mancanti) ? s.mancanti.length : 0;
                return (
                  <tr key={n} style={{ background: sel === n ? T.accent + "18" : "transparent" }}>
                    <td style={{ ...css.td, fontWeight: 600 }}>{n}</td>
                    <td style={{ ...css.td, color: T.textMid }}>{cnt}</td>
                    <td style={css.td}><StatoBadge stato={s?.rpn_stato} />{s?.titoli_rpn != null && <span style={{ color: T.textMid, fontSize: "11px" }}> · {s.titoli_rpn}</span>}</td>
                    <td style={{ ...css.td, color: manc ? T.amber : T.textDim }}>{s ? manc : "—"}</td>
                    <td style={{ ...css.td, color: T.textMid, fontSize: "11px" }}>{fmtData(s?.ultimo_sync)}</td>
                    <td style={css.td}><button style={css.btn("default", busy === "preview")} disabled={busy === "preview"} onClick={() => apri(n)}>Verifica</button></td>
                  </tr>
                );
              })}
              {!cedole.length && <tr><td style={{ ...css.td, color: T.textMid }} colSpan={6}>Nessuna cedola per questo giro</td></tr>}
            </tbody>
          </table>
          {batch && (
            <div style={{ marginTop: 14, fontSize: "12px" }}>
              <div style={{ color: T.textMid, marginBottom: 6 }}>Avanzamento: {batch.fatto}/{batch.tot}</div>
              {batch.out.map(o => (
                <div key={o.n} style={{ color: o.ok ? T.green : T.red, fontSize: "11px", marginBottom: 2 }}>
                  {o.ok ? "✓" : "✗"} {o.n}{o.ok ? ` — +${o.esito.aggiunti} −${o.esito.rimossi}${o.esito.cedola_creata ? " · creata" : ""}${o.mancanti ? ` · ${o.mancanti} non ancora in RPN` : ""}` : ` — ${o.errore}`}
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Dettaglio cedola */}
        <div style={{ ...css.card, flex: "1 1 460px", minWidth: 380 }}>
          {!sel && <div style={css.sub}>Scegli una cedola e premi "Verifica" per vedere cosa cambierebbe su RPN.</div>}
          {sel && busy === "preview" && <div style={css.sub}>Confronto GiroManager ↔ RPN in corso… (qualche secondo)</div>}
          {errore && <div style={{ color: T.red, fontSize: "12px", marginBottom: 10 }}>⚠ {errore}</div>}
          {p && (
            <>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
                <div style={{ ...css.h, marginBottom: 0 }}>{p.n_cedola}</div>
                <StatoBadge stato={p.cedola.esiste ? p.cedola.stato : null} />
                {p.extra && <span style={{ color: T.accent, fontSize: "11px" }}>EXTRA</span>}
              </div>

              <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 12 }}>
                <div style={css.kpi}><div style={css.sub}>Titoli GiroManager</div><div style={{ fontWeight: 700 }}>{p.titoli_gm}</div></div>
                <div style={css.kpi}><div style={css.sub}>Titoli su RPN</div><div style={{ fontWeight: 700 }}>{p.titoli_rpn}</div></div>
                <div style={css.kpi}><div style={css.sub}>Da agganciare</div><div style={{ fontWeight: 700, color: p.da_aggiungere.length ? T.accent : T.textMid }}>{p.da_aggiungere.length}</div></div>
                <div style={css.kpi}><div style={css.sub}>Da togliere</div><div style={{ fontWeight: 700, color: p.da_rimuovere.length ? T.red : T.textMid }}>{p.da_rimuovere.length}</div></div>
                <div style={css.kpi}><div style={css.sub}>Non in anagrafica RPN</div><div style={{ fontWeight: 700, color: p.mancanti.length ? T.amber : T.textMid }}>{p.mancanti.length}</div></div>
              </div>

              <div style={{ ...css.sub, marginBottom: 12 }}>
                {p.giro && <>Giro RPN <b style={{ color: T.text }}>{p.giro.nome}</b>: {p.giro.esiste ? "presente" : <span style={{ color: T.amber }}>verrà creato</span>} · </>}
                Cedola: {p.cedola.esiste ? "presente" : <span style={{ color: T.amber }}>verrà creata</span>} ·
                Ordine: {p.ordine_da_aggiornare ? <span style={{ color: T.amber }}>da aggiornare</span> : "allineato"}
              </div>

              {!p.cedola.esiste && (
                <div style={{ display: "flex", gap: 14, alignItems: "center", marginBottom: 12, flexWrap: "wrap" }}>
                  <label style={{ color: T.textMid, fontSize: "12px" }}>Data inizio&nbsp;<input type="date" style={css.input} value={date.data_inizio} onChange={e => setDate(d => ({ ...d, data_inizio: e.target.value }))} /></label>
                  <label style={{ color: T.textMid, fontSize: "12px" }}>Data fine&nbsp;<input type="date" style={css.input} value={date.data_fine} onChange={e => setDate(d => ({ ...d, data_fine: e.target.value }))} /></label>
                  <span style={{ color: T.textDim, fontSize: "11px" }}>{p.giro?.esiste ? "dal giro RPN" : "dal Calendario Giri"}</span>
                </div>
              )}

              <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                <button style={css.btn("accent", !!busy || chiusa)} disabled={!!busy || chiusa} onClick={pubblica}>
                  {busy === "apply" ? "Aggiornamento RPN…" : p.cedola.esiste ? "Aggiorna su RPN" : "Crea su RPN"}
                </button>
                {p.cedola.esiste && ["CRE", "DAT"].includes(p.cedola.stato) && (
                  <button style={css.btn("green", !!busy)} disabled={!!busy} onClick={attiva}>{busy === "activate" ? "Attivazione…" : "Attiva su RPN"}</button>
                )}
                <button style={css.btn("default", !!busy)} disabled={!!busy} onClick={() => apri(p.n_cedola)}>↻ Ricontrolla</button>
              </div>
              {chiusa && <div style={{ color: T.textMid, fontSize: "11px", marginTop: 8 }}>Cedola chiusa su RPN: sola lettura.</div>}

              {esito && (
                <div style={{ marginTop: 12, padding: 10, border: `1px solid ${T.green}55`, borderRadius: 4, fontSize: "12px", color: T.green }}>
                  {esito.info ?? <>✓ Fatto{esito.giro_creato ? " · giro creato" : ""}{esito.cedola_creata ? " · cedola creata" : ""} · agganciati {esito.aggiunti} · tolti {esito.rimossi} · ordinati {esito.ordinati}</>}
                  {esito.avvisi?.length > 0 && <div style={{ color: T.amber, marginTop: 6 }}>Avvisi RPN: {esito.avvisi.join(" | ")}</div>}
                </div>
              )}

              <ListaTitoli titolo="Da agganciare" righe={p.da_aggiungere} colore={T.accent} />
              <ListaTitoli titolo="Da togliere da RPN" righe={p.da_rimuovere} colore={T.red} aperta />
              <ListaTitoli titolo="Non ancora in anagrafica RPN (ritentati ogni notte)" righe={p.mancanti} colore={T.amber} />
            </>
          )}
        </div>
      </div>
    </div>
  );
}
