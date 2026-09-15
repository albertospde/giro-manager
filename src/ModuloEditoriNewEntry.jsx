import { useState, useEffect, useCallback } from "react";

const SUPABASE_URL = "https://tdflwenlylhctxssatax.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InRkZmx3ZW5seWxoY3R4c3NhdGF4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYzMzgyNzYsImV4cCI6MjA5MTkxNDI3Nn0.l35qEL7LOvyYuI1McQlVqj4vbyTqmlevcmqWbTGYi2Q";
// Stessa edge function proxy RPN usata da BookUp: è una funzione Supabase
// condivisa sul progetto, quindi GiroManager può chiamarla allo stesso modo
// — a patto che l'utente Supabase loggato qui abbia anche lui l'account RPN
// collegato (tabella rpn_credentials), esattamente come su BookUp.
const RPN_PROXY_BASE = "https://tdflwenlylhctxssatax.supabase.co/functions/v1/rpn-sync";
const TEMPLATE_HEADERS = ["Posizione", "Codice cliente", "Nome Cliente", "Gruppo cliente", "Tipo ordine", "N. ordine cliente", "N° cedola", "Stato cedola", "EAN", "Titolo", "Autore", "Editore", "Collana", "Prezzo", "Obt", "Pren (Qtà)", "Trend %", "Sc anagr", "Sconto occasionale", "Pg fisso", "Pag(occ)", "Qtà trasmessa"];

const T = {
  bg: "#1a2140", surface: "#212d54", border: "#2e3d6b", borderHi: "#3d4f82",
  text: "#f0f2f8", textMid: "#8b9cc8", textDim: "#4a5a8a",
  accent: "#7b9fe8", green: "#4caf7d", red: "#e05c5c",
};
const css = {
  btn: (v = "default") => ({ padding: "6px 14px", border: `1px solid ${v === "accent" ? T.accent : T.border}`, background: v === "accent" ? T.accent : "transparent", color: v === "accent" ? "#000" : T.text, cursor: "pointer", fontSize: "12px", fontFamily: "inherit", borderRadius: 3, fontWeight: v === "accent" ? "700" : "400", letterSpacing: "0.04em" }),
  input: { background: T.bg, border: `1px solid ${T.border}`, color: T.text, padding: "5px 10px", fontSize: "12px", fontFamily: "inherit", borderRadius: 3, outline: "none" },
  table: { width: "100%", borderCollapse: "collapse" },
  th: { padding: "8px 12px", textAlign: "left", color: T.textMid, fontWeight: "400", fontSize: "11px", letterSpacing: "0.08em", textTransform: "uppercase", borderBottom: `1px solid ${T.border}`, whiteSpace: "nowrap" },
  td: { padding: "8px 12px", borderBottom: `1px solid ${T.border}22`, verticalAlign: "middle" },
  label: { display: "block", color: T.textDim, fontSize: "10px", letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 4 },
};

async function sbRest(path, token, opts = {}) {
  const r = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    ...opts,
    headers: {
      "apikey": SUPABASE_KEY,
      "Authorization": `Bearer ${token}`,
      "Content-Type": "application/json",
      "Prefer": opts.method && opts.method !== "GET" ? "return=representation" : undefined,
      ...(opts.headers || {}),
    },
  });
  if (!r.ok) {
    const body = await r.json().catch(() => ({}));
    throw new Error(body.message || body.hint || `Errore ${r.status}`);
  }
  if (r.status === 204) return null;
  return r.json();
}

// Editori in ingresso nel perimetro PDE non ancora visibili su RPN, e gli
// EAN inseriti manualmente per loro. Le tabelle (editori_new_entry,
// titoli_manuali) sono le stesse lette da BookUp per raccogliere le
// prenotazioni: qui è solo l'anagrafica, gestita da chi già gestisce
// Titoli/Cedole — coerente col resto di GiroManager.
export default function ModuloEditoriNewEntry({ token, onDataChange }) {
  const [editori, setEditori] = useState([]);
  const [titoli, setTitoli] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  const [edCodice, setEdCodice] = useState("");
  const [edNome, setEdNome] = useState("");
  const [savingEd, setSavingEd] = useState(false);

  const [tEan, setTEan] = useState("");
  const [tTitolo, setTTitolo] = useState("");
  const [tAutore, setTAutore] = useState("");
  const [tEditore, setTEditore] = useState("");
  const [tPrezzo, setTPrezzo] = useState("");
  const [tNCedola, setTNCedola] = useState("");
  const [tCedolaNome, setTCedolaNome] = useState("");
  const [savingT, setSavingT] = useState(false);

  const load = useCallback(async () => {
    setLoading(true); setError("");
    try {
      const [ed, ti] = await Promise.all([
        sbRest("editori_new_entry?select=*&order=nome_editore.asc", token),
        sbRest("titoli_manuali?select=*&order=created_at.desc", token),
      ]);
      setEditori(ed || []);
      setTitoli(ti || []);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }, [token]);

  useEffect(() => { load(); }, [load]);

  const addEditore = async () => {
    if (!edCodice.trim() || !edNome.trim()) { setError("Codice e nome editore sono obbligatori."); return; }
    setSavingEd(true); setError("");
    try {
      await sbRest("editori_new_entry", token, {
        method: "POST",
        body: JSON.stringify({ codice_editore: edCodice.trim(), nome_editore: edNome.trim() }),
      });
      setEdCodice(""); setEdNome("");
      await load();
    } catch (e) {
      setError(e.message);
    } finally {
      setSavingEd(false);
    }
  };

  const toggleEditore = async (codice, nuovoAttivo) => {
    try {
      await sbRest(`editori_new_entry?codice_editore=eq.${encodeURIComponent(codice)}`, token, {
        method: "PATCH",
        body: JSON.stringify({ attivo: nuovoAttivo }),
      });
      await load();
    } catch (e) { setError(e.message); }
  };

  // Ricostruisce le righe { codiceCliente, cedolaNome, ean, qta, sovrasconto,
  // dilazionePagamento, nOrdine } dalle prenotazioni_manuali di un editore —
  // stessa forma usata da BookUp per costruire il file a 22 colonne.
  const raccogliRighePerEditore = async (editore) => {
    const eans = titoli.filter(t => t.editore_codice === editore.codice_editore).map(t => t.ean);
    if (!eans.length) return { rows: [], error: "Nessun titolo inserito per questo editore." };
    const inList = eans.map(e => `"${e}"`).join(",");
    const pren = await sbRest(`prenotazioni_manuali?select=*&ean=in.(${inList})&order=n_cedola.asc,codice_cliente.asc`, token);
    if (!pren || !pren.length) return { rows: [], error: "Nessuna prenotazione raccolta per questo editore." };
    return {
      rows: pren.map(p => ({
        codiceCliente: p.codice_cliente,
        cedolaNome: p.n_cedola,
        ean: p.ean,
        qta: p.quantita,
        sovrasconto: p.sconto_occasionale || "",
        dilazionePagamento: p.pagamento_occasionale || "",
        nOrdine: p.numero_ordine_cliente || "",
      })),
    };
  };

  const buildTemplateFile = (rows) => {
    const aoa = [TEMPLATE_HEADERS];
    rows.forEach(r => {
      const row = TEMPLATE_HEADERS.map(() => "");
      row[TEMPLATE_HEADERS.indexOf("Codice cliente")] = r.codiceCliente;
      row[TEMPLATE_HEADERS.indexOf("N° cedola")] = r.cedolaNome;
      row[TEMPLATE_HEADERS.indexOf("EAN")] = r.ean;
      row[TEMPLATE_HEADERS.indexOf("Pren (Qtà)")] = r.qta;
      if (r.sovrasconto) row[TEMPLATE_HEADERS.indexOf("Sconto occasionale")] = r.sovrasconto;
      if (r.dilazionePagamento) row[TEMPLATE_HEADERS.indexOf("Pag(occ)")] = r.dilazionePagamento;
      if (r.nOrdine) row[TEMPLATE_HEADERS.indexOf("N. ordine cliente")] = r.nOrdine;
      aoa.push(row);
    });
    const ws = window.XLSX.utils.aoa_to_sheet(aoa);
    const wb = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    const buf = window.XLSX.write(wb, { type: "array", bookType: "xlsx" });
    return new File([buf], "prenotazioni.xlsx", { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  };

  // Carica automaticamente su RPN le prenotazioni raccolte per un editore,
  // riusando la stessa edge function proxy di BookUp (upload-prenotazioni).
  // Richiede che l'account RPN sia collegato per l'utente Supabase con cui
  // sei loggato QUI su GiroManager — se non lo è, la funzione risponde 428
  // e te lo segnalo esplicitamente invece di lasciarti un errore muto.
  const [caricandoRpn, setCaricandoRpn] = useState(null); // codice_editore in corso

  const caricaSuRpnAutomatico = async (editore) => {
    setError("");
    setCaricandoRpn(editore.codice_editore);
    try {
      const { rows, error: errRighe } = await raccogliRighePerEditore(editore);
      if (errRighe) { alert(errRighe); return; }
      const file = buildTemplateFile(rows);
      const fd = new FormData();
      fd.append("csv_file", file);
      const res = await fetch(`${RPN_PROXY_BASE}/upload-prenotazioni?confirmed=true`, {
        method: "POST",
        headers: { "Authorization": `Bearer ${token}` },
        body: fd,
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) {
        if (res.status === 428) {
          alert("L'account RPN non risulta collegato per l'utente con cui sei loggato QUI su GiroManager. Collegalo (stessa utenza usata per accedere a BookUp) e riprova, oppure usa \"📥 Esporta per RPN\" e carica il file a mano da BookUp.");
          return;
        }
        throw new Error(data.message || data.error || `Errore ${res.status}`);
      }
      const saved = (data.saved && data.saved.length ? data.saved : data.data) || [];
      const nonConnected = data.nonConnected || [];
      let msg = `✅ ${saved.length} riga/righe caricate su RPN per "${editore.nome_editore}".`;
      if (nonConnected.length) msg += `\n⚠️ ${nonConnected.length} riga/righe NON agganciate su RPN (verifica cedola/cliente): controllale prima di segnare l'editore come entrato.`;
      alert(msg);
    } catch (e) {
      setError("Errore nel caricamento automatico su RPN: " + e.message);
    } finally {
      setCaricandoRpn(null);
    }
  };

  // Backup manuale (nel formato semplice che BookUp legge in "Carica tante
  // prenotazioni insieme"), utile se il caricamento automatico fallisce o
  // se preferisci controllare il file prima di inviarlo.
  const esportaPrenotazioni = async (editore) => {
    setError("");
    try {
      const { rows, error: errRighe } = await raccogliRighePerEditore(editore);
      if (errRighe) { alert(errRighe); return; }
      const aoa = [
        ["Codice cliente", "N° cedola", "EAN", "Copie", "Sconto occasionale", "Pag(occ)", "N. ordine cliente"],
        ...rows.map(r => [r.codiceCliente, r.cedolaNome, r.ean, r.qta, r.sovrasconto, r.dilazionePagamento, r.nOrdine]),
      ];
      const ws = window.XLSX.utils.aoa_to_sheet(aoa);
      const wb = window.XLSX.utils.book_new();
      window.XLSX.utils.book_append_sheet(wb, ws, "Prenotazioni");
      window.XLSX.writeFile(wb, `Prenotazioni_${editore.codice_editore}_per_RPN.xlsx`);
    } catch (e) { setError(e.message); }
  };

  // 2) Dopo aver ricaricato il file su RPN (via BookUp): disattiva editore
  // e titoli manuali, e ripulisce SOLO le righe "ombra" specchiate nel
  // Fine Giro — evita il doppio conteggio una volta che l'import RPN reale
  // porta gli stessi numeri sotto un titolo vero. Lo storico delle
  // prenotazioni raccolte resta su Supabase, non viene toccato.
  const migraEditore = async (editore) => {
    const ok = window.confirm(
      `Confermi di aver già ricaricato su RPN le prenotazioni di "${editore.nome_editore}"?\n\n` +
      `Questa azione:\n` +
      `- disattiva l'editore e i suoi titoli manuali (spariranno dalla griglia di BookUp)\n` +
      `- rimuove i numeri "ombra" di questo editore dal Fine Giro — da qui in poi contano solo i numeri veri importati da RPN\n\n` +
      `Lo storico delle prenotazioni raccolte resta consultabile, non viene cancellato.`
    );
    if (!ok) return;
    setError("");
    try {
      const res = await sbRest(`rpc/migra_editore_new_entry_su_rpn`, token, {
        method: "POST",
        body: JSON.stringify({ p_codice_editore: editore.codice_editore }),
      });
      const r = Array.isArray(res) ? res[0] : res;
      alert(`Fatto: ${r?.titoli_disattivati ?? 0} titoli disattivati, ${r?.righe_finegiro_rimosse ?? 0} righe rimosse dal Fine Giro.`);
      await load();
      if (onDataChange) onDataChange();
    } catch (e) { setError(e.message); }
  };

  const addTitolo = async () => {
    if (!tEan.trim() || !tTitolo.trim() || !tEditore || !tNCedola.trim()) {
      setError("EAN, Titolo, Editore e N° cedola sono obbligatori.");
      return;
    }
    const ed = editori.find(e => e.codice_editore === tEditore);
    setSavingT(true); setError("");
    try {
      await sbRest("titoli_manuali", token, {
        method: "POST",
        body: JSON.stringify({
          ean: tEan.trim(),
          titolo: tTitolo.trim(),
          autore: tAutore.trim() || null,
          editore_codice: tEditore,
          editore_nome: ed ? ed.nome_editore : tEditore,
          prezzo: tPrezzo.trim() ? Number(tPrezzo.trim()) : null,
          n_cedola: tNCedola.trim(),
          cedola_nome: (tCedolaNome.trim() || tNCedola.trim()),
        }),
      });
      setTEan(""); setTTitolo(""); setTAutore(""); setTPrezzo(""); setTNCedola(""); setTCedolaNome("");
      await load();
      if (onDataChange) onDataChange(); // ricarica titoli/prenotato in GiroManager: il trigger DB ha già creato la riga in "titoli"
    } catch (e) {
      setError(e.message);
    } finally {
      setSavingT(false);
    }
  };

  const toggleTitolo = async (id, nuovoAttivo) => {
    try {
      await sbRest(`titoli_manuali?id=eq.${id}`, token, {
        method: "PATCH",
        body: JSON.stringify({ attivo: nuovoAttivo }),
      });
      await load();
      if (onDataChange) onDataChange();
    } catch (e) { setError(e.message); }
  };

  return (
    <div style={{ flex: 1, overflow: "auto", padding: 20 }}>
      <p style={{ color: T.textMid, fontSize: "12px", marginBottom: 4, maxWidth: 720 }}>
        Editori in ingresso nel perimetro PDE non ancora visibili su RPN. Gli EAN inseriti qui vanno agganciati a una cedola RPN <b>già esistente</b> (stesso numero/nome): compariranno mescolati in BookUp → "Inserisci prenotazioni" per quella cedola, e le prenotazioni raccolte lì si salvano su Supabase invece che su RPN — trovi comunque i totali già sommati qui in Fine Giro.
      </p>
      <p style={{ color: T.textDim, fontSize: "11px", marginBottom: 18, maxWidth: 720 }}>
        Quando l'editore entra davvero su RPN: <b>1)</b> "🚀 Carica su RPN" (automatico — richiede che l'account RPN sia collegato per l'utente con cui sei loggato qui); <b>2)</b> controlla l'esito, e solo dopo "✅ Entrato su RPN" per disattivare i titoli manuali ed evitare doppioni/doppio conteggio. "📥 Esporta file" resta come backup se preferisci caricare a mano da BookUp.
      </p>

      {error && <div style={{ color: T.red, fontSize: "12px", marginBottom: 14, background: T.red + "18", border: `1px solid ${T.red}44`, borderRadius: 4, padding: "8px 12px" }}>{error}</div>}
      {loading ? <div style={{ color: T.textMid, fontSize: "12px" }}>Caricamento…</div> : (
        <div style={{ display: "flex", gap: 28, flexWrap: "wrap", alignItems: "flex-start" }}>

          <div style={{ flex: "1 1 320px", minWidth: 320 }}>
            <h3 style={{ fontSize: "13px", color: T.text, marginBottom: 10 }}>Editori</h3>
            <div style={{ display: "flex", gap: 8, marginBottom: 12, alignItems: "flex-end", flexWrap: "wrap" }}>
              <div>
                <label style={css.label}>Codice interno</label>
                <input style={{ ...css.input, width: 130 }} value={edCodice} onChange={e => setEdCodice(e.target.value)} placeholder="es. NUOVOED01" />
              </div>
              <div style={{ flex: 1, minWidth: 160 }}>
                <label style={css.label}>Nome editore</label>
                <input style={{ ...css.input, width: "100%" }} value={edNome} onChange={e => setEdNome(e.target.value)} placeholder="Nome editore" />
              </div>
              <button style={css.btn("accent")} disabled={savingEd} onClick={addEditore}>{savingEd ? "…" : "+ Aggiungi"}</button>
            </div>
            <table style={css.table}>
              <thead><tr><th style={css.th}>Codice</th><th style={css.th}>Nome</th><th style={css.th}>Attivo</th><th style={css.th}></th></tr></thead>
              <tbody>
                {editori.length === 0 && <tr><td style={css.td} colSpan={4}><span style={{ color: T.textDim }}>Nessun editore new entry ancora inserito.</span></td></tr>}
                {editori.map(ed => (
                  <tr key={ed.codice_editore}>
                    <td style={css.td}>{ed.codice_editore}</td>
                    <td style={css.td}>{ed.nome_editore}</td>
                    <td style={css.td}>{ed.attivo ? "✅" : "⛔"}</td>
                    <td style={css.td}>
                      <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                        <button style={{ ...css.btn(), fontSize: "11px", padding: "3px 8px" }} onClick={() => toggleEditore(ed.codice_editore, !ed.attivo)}>{ed.attivo ? "Disattiva" : "Riattiva"}</button>
                        {ed.attivo && <>
                          <button style={{ ...css.btn("accent"), fontSize: "11px", padding: "3px 8px" }} disabled={caricandoRpn === ed.codice_editore} onClick={() => caricaSuRpnAutomatico(ed)} title="Carica direttamente su RPN le prenotazioni raccolte per questo editore">
                            {caricandoRpn === ed.codice_editore ? "⏳ Carico..." : "🚀 Carica su RPN"}
                          </button>
                          <button style={{ ...css.btn(), fontSize: "11px", padding: "3px 8px" }} onClick={() => esportaPrenotazioni(ed)} title="Backup: scarica il file invece di caricarlo in automatico">📥 Esporta file</button>
                          <button style={{ ...css.btn(), fontSize: "11px", padding: "3px 8px", color: T.green, borderColor: T.green }} onClick={() => migraEditore(ed)} title="Da usare SOLO dopo aver caricato le prenotazioni su RPN">✅ Entrato su RPN</button>
                        </>}
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div style={{ flex: "1.6 1 420px", minWidth: 420 }}>
            <h3 style={{ fontSize: "13px", color: T.text, marginBottom: 10 }}>Titoli (EAN)</h3>
            <div style={{ display: "flex", gap: 8, marginBottom: 6, flexWrap: "wrap" }}>
              <div><label style={css.label}>EAN</label><input style={{ ...css.input, width: 130 }} value={tEan} onChange={e => setTEan(e.target.value)} /></div>
              <div style={{ flex: 1, minWidth: 180 }}><label style={css.label}>Titolo</label><input style={{ ...css.input, width: "100%" }} value={tTitolo} onChange={e => setTTitolo(e.target.value)} /></div>
              <div><label style={css.label}>Autore</label><input style={{ ...css.input, width: 140 }} value={tAutore} onChange={e => setTAutore(e.target.value)} /></div>
            </div>
            <div style={{ display: "flex", gap: 8, marginBottom: 6, flexWrap: "wrap", alignItems: "flex-end" }}>
              <div>
                <label style={css.label}>Editore</label>
                <select style={{ ...css.input, width: 180 }} value={tEditore} onChange={e => setTEditore(e.target.value)}>
                  <option value="">— scegli —</option>
                  {editori.map(ed => <option key={ed.codice_editore} value={ed.codice_editore}>{ed.nome_editore}{ed.attivo ? "" : " (disattivato)"}</option>)}
                </select>
              </div>
              <div><label style={css.label}>Prezzo</label><input type="number" step="0.01" style={{ ...css.input, width: 90 }} value={tPrezzo} onChange={e => setTPrezzo(e.target.value)} /></div>
              <div><label style={css.label}>N° cedola RPN</label><input style={{ ...css.input, width: 140 }} value={tNCedola} onChange={e => setTNCedola(e.target.value)} placeholder="es. TEST 2026 A" /></div>
              <div><label style={css.label}>Nome cedola (se diverso)</label><input style={{ ...css.input, width: 140 }} value={tCedolaNome} onChange={e => setTCedolaNome(e.target.value)} /></div>
              <button style={css.btn("accent")} disabled={savingT} onClick={addTitolo}>{savingT ? "…" : "+ Aggiungi"}</button>
            </div>
            <p style={{ color: T.textDim, fontSize: "11px", marginBottom: 10 }}>N° e nome cedola devono combaciare ESATTAMENTE con quelli reali su RPN, altrimenti il titolo non comparirà mescolato nella cedola giusta in BookUp.</p>
            <table style={css.table}>
              <thead><tr><th style={css.th}>EAN</th><th style={css.th}>Titolo</th><th style={css.th}>Editore</th><th style={css.th}>Cedola</th><th style={css.th}></th></tr></thead>
              <tbody>
                {titoli.length === 0 && <tr><td style={css.td} colSpan={5}><span style={{ color: T.textDim }}>Nessun titolo manuale ancora inserito.</span></td></tr>}
                {titoli.map(t => (
                  <tr key={t.id}>
                    <td style={css.td}>{t.ean}</td>
                    <td style={css.td}>{t.titolo}</td>
                    <td style={css.td}>{t.editore_nome}</td>
                    <td style={css.td}>{t.n_cedola}{!t.attivo && " (disattivato)"}</td>
                    <td style={css.td}><button style={{ ...css.btn(), fontSize: "11px", padding: "3px 8px" }} onClick={() => toggleTitolo(t.id, !t.attivo)}>{t.attivo ? "Disattiva" : "Riattiva"}</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

        </div>
      )}
    </div>
  );
}
