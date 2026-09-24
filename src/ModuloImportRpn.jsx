import { useState, useCallback } from "react";
import { fetchAnagraficaEditori, resolveGiri } from "./ModuloImport.jsx";

const SUPABASE_URL = "https://tdflwenlylhctxssatax.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InRkZmx3ZW5seWxoY3R4c3NhdGF4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYzMzgyNzYsImV4cCI6MjA5MTkxNDI3Nn0.l35qEL7LOvyYuI1McQlVqj4vbyTqmlevcmqWbTGYi2Q";

const T = {
  bg: "#0f0f0f", surface: "#161616", border: "#252525", borderHi: "#333333",
  text: "#e8e8e8", textMid: "#888888", textDim: "#444444",
  accent: "#c8a96e", green: "#4caf7d", red: "#e05c5c", blue: "#5b8fd4",
};
const css = {
  btn: (v = "default") => ({ padding: "6px 14px", border: `1px solid ${v === "accent" ? T.accent : v === "danger" ? T.red : T.border}`, background: v === "accent" ? T.accent : v === "danger" ? T.red + "22" : "transparent", color: v === "accent" ? "#000" : v === "danger" ? T.red : T.text, cursor: "pointer", fontSize: "12px", fontFamily: "inherit", borderRadius: 3, fontWeight: v === "accent" ? "700" : "400" }),
  th: { padding: "7px 10px", textAlign: "left", color: T.textMid, fontWeight: "400", fontSize: "11px", letterSpacing: "0.06em", textTransform: "uppercase", borderBottom: `1px solid ${T.border}`, whiteSpace: "nowrap", background: T.surface },
  td: { padding: "6px 10px", borderBottom: `1px solid ${T.border}22`, verticalAlign: "middle", fontSize: "12px" },
};

const FORMATO_DEFAULT = "Cover";
// Parsa il nome giro RPN ("GIRO 5 2026" o "5 2026") in numero+anno.
const GIRO_RPN_RE = /^(?:GIRO\s+)?(\d+)\s+(\d{4})$/i;

// Campi "manuali" che RPN non fornisce mai: se il titolo esiste già, li riportiamo
// invariati per non azzerare lavoro fatto a mano (obiettivi, note, promozioni...).
const CAMPI_MANUALI = [
  "obiettivo_assegnato", "il_triangolo", "top_100", "promozione",
  "note_comunicazione", "note", "uscita",
  "ean_gemello_1", "titolo_gemello_1", "ean_gemello_2", "titolo_gemello_2",
  "ean_gemello_3", "titolo_gemello_3",
];

async function rpnSync(token, path) {
  const res = await fetch(`${SUPABASE_URL}/functions/v1/rpn-sync/${path}`, {
    headers: { Authorization: `Bearer ${token}`, apikey: SUPABASE_KEY },
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    if (data.error === "RPN_NOT_CONNECTED") throw new Error("Account RPN non collegato: ricollegalo in Prenota.");
    throw new Error(data.message || data.error || `Errore ${res.status} su ${path}`);
  }
  return data;
}

// Cedole chiuse su RPN non compaiono più in "titolo-tab" (usato da rpnSync
// per l'elenco titoli): questa funzione dedicata legge invece lo storico
// prenotazioni aggregato per cedola, così i titoli restano recuperabili
// anche dopo la chiusura.
async function rpnTitoliStorico(token, cedolaId) {
  const res = await fetch(`${SUPABASE_URL}/functions/v1/rpn-titoli-storico/${cedolaId}`, {
    headers: { Authorization: `Bearer ${token}`, apikey: SUPABASE_KEY },
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    if (data.error === "RPN_NOT_CONNECTED") throw new Error("Account RPN non collegato: ricollegalo in Prenota.");
    throw new Error(data.message || data.error || `Errore ${res.status} su storico titoli cedola ${cedolaId}`);
  }
  return data;
}

function normSpazi(s) {
  return s.replace(/\s+/g, " ").trim(); // \s in JS copre anche nbsp e simili
}

function normalizeTitolo(t) {
  const ean = String(t.ean ?? t.Ean ?? "").trim();
  const titolo = normSpazi(String(t.titolo ?? t.Titolo ?? "")).toUpperCase();
  const autoreRaw = normSpazi(String(t.autore ?? t.Autore ?? "")).toUpperCase();
  const editore_nome = normSpazi(String(t.editore ?? t.Editore ?? "")).toUpperCase();
  const prezzo = Number(t.prezzo ?? t.Prezzo ?? 0) || null;
  return { ean, titolo, autore: autoreRaw === "NESSUNO" || !autoreRaw ? null : autoreRaw, editore_nome, prezzo };
}

// Rimuove un articolo/preposizione iniziale ("IL SAGGIATORE" → "SAGGIATORE") così che
// il matching RPN↔anagrafica funzioni anche quando una delle due fonti omette l'articolo
// (es. RPN restituisce "SAGGIATORE" mentre in anagrafica è salvato "IL SAGGIATORE").
const PREFISSI_EDITORE = ["IL ", "LO ", "LA ", "GLI ", "LE ", "I ", "L'"];
function stripArticolo(nome) {
  for (const p of PREFISSI_EDITORE) {
    if (nome.startsWith(p)) return nome.slice(p.length);
  }
  return nome;
}

// Parole che in anagrafica possono seguire il nome editore senza cambiarne l'identità.
const SUFFISSI_GENERICI = new Set([
  "EDITORE", "EDITORI", "EDITRICE", "EDIZIONI", "EDITORIALE", "LIBRI",
  "SRL", "SPA", "SAS", "SNC", "SRLS", "AD",
]);

// ─── Import di UNA cedola/giro selezionato ──────────────────────────────────
// Se il nome editore di RPN non matcha esattamente l'anagrafica, prova due fallback:
// 1) RPN a volte restituisce il nome "arricchito" (nome proprio, codice interno, parola
//    EDITORE aggiunta) — es. "SILVANA EDITORIALE 821" invece di "SILVANA EDITORIALE",
//    "SKIRA EDITORE" invece di "SKIRA", "ALLEMANDI UMBERTO" invece di "ALLEMANDI".
// 2) Una delle due fonti omette l'articolo iniziale — es. "SAGGIATORE" (RPN) vs
//    "IL SAGGIATORE" (anagrafica). Si confrontano i nomi anche dopo aver tolto l'articolo.
// Cerca tra le chiavi anagrafica quelle compatibili con il nome RPN secondo queste regole
// (a confine di parola, per evitare match spuri tipo MONDADORI vs MONDADORI EDUCATION).
// Se ne trova esattamente una la usa e lo segnala; se zero o più di una resta un errore.
function risolviAnagrafica(nomeRpn, anagraficaMap) {
  const diretto = anagraficaMap[nomeRpn];
  if (diretto) return { match: diretto, viaFallback: false };

  // 1) Match esatto ignorando l'articolo iniziale (RPN "SAGGIATORE" ↔ anagrafica "IL SAGGIATORE").
  // Confronto stretto (===), non prefix-match, per non confondere editori affini come
  // "IL SAGGIATORE" e "IL SAGGIATORE -TASCABILI" (che dopo lo strip diventano nomi diversi).
  const nomeRpnNorm = stripArticolo(nomeRpn);
  const perArticolo = Object.keys(anagraficaMap).filter(k => stripArticolo(k) === nomeRpnNorm);
  if (perArticolo.length === 1) return { match: anagraficaMap[perArticolo[0]], viaFallback: true, nomeUsato: perArticolo[0] };
  if (perArticolo.length > 1) return { match: null, viaFallback: false, ambiguo: perArticolo };

  // 2) Fallback "nome arricchito": il nome RPN contiene il nome anagrafica come prefisso
  // a confine di parola (es. "SILVANA EDITORIALE 821" → "SILVANA EDITORIALE").
  const perPrefisso = Object.keys(anagraficaMap).filter(k => nomeRpn.startsWith(k + " "));
  if (perPrefisso.length === 1) return { match: anagraficaMap[perPrefisso[0]], viaFallback: true, nomeUsato: perPrefisso[0] };
  if (perPrefisso.length > 1) return { match: null, viaFallback: false, ambiguo: perPrefisso };

  // 3) Fallback "nome ridotto": è RPN ad accorciare il nome, l'anagrafica ha in più solo
  // parole generiche (ragione sociale/suffissi) — es. RPN "GALLUCCI" → "GALLUCCI EDITORE SRL",
  // RPN "GALLUCCI CENTAURIA" → "GALLUCCI CENTAURIA AD". Non si accettano marchi diversi
  // (GALLUCCI BROS, GALLUCCI SPIGA... restano esclusi perché BROS/SPIGA non sono generici).
  const perRiduzione = Object.keys(anagraficaMap).filter(k => {
    if (!k.startsWith(nomeRpn + " ")) return false;
    const resto = k.slice(nomeRpn.length + 1).split(" ").filter(Boolean);
    return resto.length > 0 && resto.every(w => SUFFISSI_GENERICI.has(w.replace(/\./g, "")));
  });
  if (perRiduzione.length === 1) return { match: anagraficaMap[perRiduzione[0]], viaFallback: true, nomeUsato: perRiduzione[0] };
  if (perRiduzione.length > 1) return { match: null, viaFallback: false, ambiguo: perRiduzione };

  return { match: null, viaFallback: false };
}

// ─── Carica l'elenco cedole/giri da RPN ──────────────────────────────────────
// giro-cedola-list / cedola-extra-list: solo le cedole ATTUALMENTE aperte in RPN
// (endpoint "titolo-tab", pensato per la presa titoli corrente).
// giro-cedola-storico-agente / cedola-extra-storico: TUTTO lo storico delle
// prenotazioni RPN (endpoint "le-mie-prenotazione"), quindi include anche le
// cedole già chiuse — necessario per poterle importare/reimportare comunque.
// Le due fonti si sovrappongono sulle cedole ancora aperte: dedup per id+tipo.
async function fetchElencoCedole(token) {
  const [giri, extra, giriStorico, extraStorico] = await Promise.all([
    rpnSync(token, "giro-cedola-list"),
    rpnSync(token, "cedola-extra-list"),
    rpnSync(token, "giro-cedola-storico-agente").catch(() => ({ giri: [] })),
    rpnSync(token, "cedola-extra-storico").catch(() => ({ cedole: [] })),
  ]);
  const statoDi = (c) => {
    const v = c.stato ?? c.Stato ?? c.status ?? c.Status ?? c.statoDescrizione ?? null;
    return v === null || v === undefined || v === "" ? null : String(v).trim();
  };
  const elenco = [];
  const visti = new Set();
  const aggiungiGiro = (g) => {
    (g.cedolaSet || []).forEach(c => {
      const key = `G-${c.id}`;
      if (visti.has(key)) return;
      visti.add(key);
      elenco.push({
        key, cedolaId: c.id, nome: c.nome, tipo: "giro",
        giroId: g.id, giroNome: g.nome, numeroTitoli: c.numeroTitoli ?? 0,
        stato: statoDi(c) ?? statoDi(g),
      });
    });
  };
  const aggiungiExtra = (c) => {
    const key = `X-${c.id}`;
    if (visti.has(key)) return;
    visti.add(key);
    elenco.push({
      key, cedolaId: c.id, nome: c.nome, tipo: "extra",
      giroId: null, giroNome: null, numeroTitoli: c.numeroTitoli ?? 0,
      stato: statoDi(c),
    });
  };
  (Array.isArray(giri) ? giri : []).forEach(aggiungiGiro);
  (Array.isArray(extra) ? extra : []).forEach(aggiungiExtra);
  (giriStorico?.giri || []).forEach(aggiungiGiro);
  (extraStorico?.cedole || []).forEach(aggiungiExtra);
  return elenco;
}

async function importCedola(token, item, anagraficaMap) {
  let { results } = await rpnSync(token, `titoli/${item.cedolaId}`);
  if (!results || !results.length) {
    // Cedola non fra quelle attive in RPN (es. chiusa): ripiega sullo storico
    // prenotazioni per recuperare comunque i titoli da poter importare.
    try {
      const storico = await rpnTitoliStorico(token, item.cedolaId);
      results = storico.results || [];
    } catch (_e) {
      // nessuno storico disponibile per questa cedola: prosegue con lista vuota
    }
  }
  const grezzi = (results || []).map(normalizeTitolo).filter(r => r.ean && r.editore_nome);
  if (!grezzi.length) return { creati: 0, aggiornati: 0, ignorati: (results || []).length, errori: [] };

  const errori = [];
  let giroNumero = null, giroAnno = null;
  if (item.tipo === "giro") {
    const m = item.giroNome.trim().match(GIRO_RPN_RE);
    if (!m) throw new Error(`Nome giro RPN non riconosciuto: "${item.giroNome}"`);
    giroNumero = parseInt(m[1]); giroAnno = parseInt(m[2]);
  }

  // Risolvi anagrafica + n_cedola/giro per ogni titolo
  const giroCombos = new Set();
  const righe = grezzi.map((r, idx) => {
    const { match: anagrafica, viaFallback, nomeUsato, ambiguo } = risolviAnagrafica(r.editore_nome, anagraficaMap);
    const out = {
      ...r,
      codice_editore: anagrafica?.codice_editore ?? null,
      ranking_editore: anagrafica?.ranking ?? null,
      account_editore: anagrafica?.account_editore ?? null,
      formato: FORMATO_DEFAULT,
      posizione: idx + 1,
      giro_id: null, giro_label: item.tipo === "extra" ? "EXTRA" : null, n_cedola: item.nome,
    };
    if (!anagrafica) {
      if (ambiguo) errori.push(`${r.editore_nome}: nome ambiguo, corrisponde a più editori in anagrafica (${ambiguo.join(", ")}) — ean ${r.ean}`);
      else errori.push(`${r.editore_nome}: non trovato in anagrafica (ean ${r.ean})`);
      return null;
    }
    // Salva sempre il nome come in anagrafica (es. "GALLUCCI" → "GALLUCCI EDITORE SRL"):
    // ranking live e raggruppamenti per editore lavorano su editore_nome.
    if (viaFallback && nomeUsato) out.editore_nome = nomeUsato;
    if (viaFallback) errori.push(`ℹ ${r.editore_nome} → abbinato ad anagrafica "${nomeUsato}" (corrispondenza automatica, verifica) — ean ${r.ean}`);
    if (item.tipo === "giro") {
      const categoria = anagrafica.cedola;
      if (!categoria) { errori.push(`${r.editore_nome}: senza categoria cedola in anagrafica`); return null; }
      out.giro_label = `${giroNumero} ${giroAnno}`;
      out.n_cedola = `GIRO ${giroNumero} ${giroAnno} ${categoria}`;
      out._giroKey = `${giroNumero}|${giroAnno}|${categoria}`;
      giroCombos.add(out._giroKey);
    }
    return out;
  }).filter(Boolean);

  if (!righe.length) return { creati: 0, aggiornati: 0, ignorati: grezzi.length, errori };

  if (item.tipo === "giro" && giroCombos.size) {
    const giriMap = await resolveGiri(token, giroCombos);
    righe.forEach(r => { r.giro_id = giriMap[r._giroKey] ?? null; delete r._giroKey; });
  }

  // Recupera i titoli già presenti per riportare i campi manuali e decidere update/insert
  let esistentiMap = {};
  if (item.tipo === "giro") {
    const giroIds = [...new Set(righe.map(r => r.giro_id).filter(Boolean))];
    if (giroIds.length) {
      const res = await fetch(
        `${SUPABASE_URL}/rest/v1/titoli?select=id,ean,giro_id,${CAMPI_MANUALI.join(",")}&giro_id=in.(${giroIds.join(",")})`,
        { headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${token}` } }
      );
      const rows = await res.json();
      rows.forEach(row => { esistentiMap[`${row.ean}__${row.giro_id}`] = row; });
    }
  } else {
    const res = await fetch(
      `${SUPABASE_URL}/rest/v1/titoli?select=id,ean,${CAMPI_MANUALI.join(",")}&n_cedola=eq.${encodeURIComponent(item.nome)}`,
      { headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${token}` } }
    );
    const rows = await res.json();
    rows.forEach(row => { esistentiMap[`${row.ean}__X`] = row; });
  }

  righe.forEach(r => {
    const key = item.tipo === "giro" ? `${r.ean}__${r.giro_id}` : `${r.ean}__X`;
    const esistente = esistentiMap[key];
    r._esistenteId = esistente?.id ?? null;
    CAMPI_MANUALI.forEach(f => { r[f] = esistente ? esistente[f] : (f === "il_triangolo" || f === "top_100" ? false : null); });
  });

  let creati = 0, aggiornati = 0;

  if (item.tipo === "giro") {
    // ON CONFLICT (ean, giro_id) funziona correttamente qui: upsert sicuro in un colpo solo.
    const payload = righe.map(({ _esistenteId, ...r }) => r);
    const res = await fetch(`${SUPABASE_URL}/rest/v1/rpc/upsert_titoli`, {
      method: "POST",
      headers: { "Content-Type": "application/json", apikey: SUPABASE_KEY, Authorization: `Bearer ${token}` },
      body: JSON.stringify({ payload }),
    });
    if (!res.ok) throw new Error("upsert_titoli: " + JSON.stringify(await res.json().catch(() => ({}))));
    creati = righe.filter(r => !r._esistenteId).length;
    aggiornati = righe.filter(r => r._esistenteId).length;
  } else {
    // Cedole extra: giro_id è sempre NULL, il vincolo UNIQUE(ean, giro_id) non le distingue
    // (NULL non genera conflitto in Postgres) — upsert_titoli da sola creerebbe duplicati
    // a ogni risincronizzazione. Split esplicito: update by id per chi esiste già, insert per i nuovi.
    const daAggiornare = righe.filter(r => r._esistenteId);
    const daCreare = righe.filter(r => !r._esistenteId);

    for (const r of daAggiornare) {
      const { _esistenteId, ...body } = r;
      const res = await fetch(`${SUPABASE_URL}/rest/v1/titoli?id=eq.${_esistenteId}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json", apikey: SUPABASE_KEY, Authorization: `Bearer ${token}`, Prefer: "return=minimal" },
        body: JSON.stringify(body),
      });
      if (!res.ok) errori.push(`Errore aggiornamento ean ${r.ean}: ${await res.text()}`);
      else aggiornati++;
    }
    if (daCreare.length) {
      const payload = daCreare.map(({ _esistenteId, ...r }) => r);
      const res = await fetch(`${SUPABASE_URL}/rest/v1/titoli`, {
        method: "POST",
        headers: { "Content-Type": "application/json", apikey: SUPABASE_KEY, Authorization: `Bearer ${token}`, Prefer: "return=minimal" },
        body: JSON.stringify(payload),
      });
      if (!res.ok) errori.push(`Errore creazione: ${await res.text()}`);
      else creati = daCreare.length;
    }
  }

  return { creati, aggiornati, ignorati: grezzi.length - righe.length, errori };
}

// ─── Componente principale ───────────────────────────────────────────────────
export default function ModuloImportRpn({ token, onImportDone }) {
  const [elenco, setElenco] = useState([]);
  const [loadingElenco, setLoadingElenco] = useState(false);
  const [erroreElenco, setErroreElenco] = useState(null);
  const [ricerca, setRicerca] = useState("");
  // Nessun filtro anno di default: molte cedole extra/campagne RPN non hanno l'anno nel nome
  // (es. "CORTINA DIONIGI", "FANDANGO I SWEAR") o hanno un anno diverso da quello corrente
  // (es. "CALENDARIO NERI POZZA 2027") — filtrare per anno corrente le nascondeva silenziosamente.
  const [anno, setAnno] = useState("");
  const [statoFiltro, setStatoFiltro] = useState("");
  const [selezionati, setSelezionati] = useState(new Set());
  const [importando, setImportando] = useState(false);
  const [log, setLog] = useState([]); // [{nome, tipo, stato: pending|ok|errore, dettaglio}]
  const [extraQuery, setExtraQuery] = useState("");
  const [extraOpen, setExtraOpen] = useState(false);

  const caricaElenco = useCallback(async () => {
    setLoadingElenco(true); setErroreElenco(null); setElenco([]); setSelezionati(new Set());
    try {
      const lista = await fetchElencoCedole(token);
      setElenco(lista);
    } catch (e) {
      setErroreElenco(e.message);
    }
    setLoadingElenco(false);
  }, [token]);

  const statiDisponibili = [...new Set(elenco.map(c => c.stato).filter(Boolean))].sort();

  const passaAnnoStato = (c) => {
    if (anno && !(c.nome.includes(anno) || (c.giroNome || "").includes(anno))) return false;
    if (statoFiltro && c.stato !== statoFiltro) return false;
    return true;
  };

  const filtratiGiri = elenco.filter(c => c.tipo === "giro" && passaAnnoStato(c) && (!ricerca || c.nome.toUpperCase().includes(ricerca.toUpperCase())));

  const extraCandidati = elenco
    .filter(c => c.tipo === "extra" && passaAnnoStato(c) && (!extraQuery || c.nome.toUpperCase().includes(extraQuery.toUpperCase())))
    .slice(0, 40); // limite risultati mostrati nella tendina, si affina scrivendo

  const extraSelezionati = elenco.filter(c => c.tipo === "extra" && selezionati.has(c.key));

  const toggle = (key) => {
    setSelezionati(prev => {
      const next = new Set(prev);
      next.has(key) ? next.delete(key) : next.add(key);
      return next;
    });
  };
  const toggleTuttiGiri = () => {
    const tuttiSelezionati = filtratiGiri.length > 0 && filtratiGiri.every(c => selezionati.has(c.key));
    setSelezionati(prev => {
      const next = new Set(prev);
      filtratiGiri.forEach(c => tuttiSelezionati ? next.delete(c.key) : next.add(c.key));
      return next;
    });
  };

  const handleImporta = async () => {
    const daImportare = elenco.filter(c => selezionati.has(c.key)); // ordine = ordine trovato su RPN
    if (!daImportare.length) return;
    setImportando(true);
    setLog(daImportare.map(c => ({ nome: c.nome, tipo: c.tipo, stato: "pending" })));

    let anagraficaMap = {};
    try {
      anagraficaMap = await fetchAnagraficaEditori(token);
    } catch (e) {
      alert("Impossibile caricare l'anagrafica editori: " + e.message);
      setImportando(false);
      return;
    }

    for (let i = 0; i < daImportare.length; i++) {
      const item = daImportare[i];
      try {
        const r = await importCedola(token, item, anagraficaMap);
        setLog(prev => prev.map((l, idx) => idx === i ? { ...l, stato: "ok", dettaglio: r } : l));
      } catch (e) {
        setLog(prev => prev.map((l, idx) => idx === i ? { ...l, stato: "errore", dettaglio: { errori: [e.message] } } : l));
      }
    }
    setImportando(false);
    onImportDone && onImportDone();
  };

  return (
    <div>
      {elenco.length === 0 && !loadingElenco && (
        <div style={{ marginBottom: 16 }}>
          <button style={css.btn("accent")} onClick={caricaElenco}>Carica elenco cedole/giri da RPN</button>
          {erroreElenco && <div style={{ color: T.red, fontSize: "12px", marginTop: 8 }}>⚠ {erroreElenco}</div>}
        </div>
      )}

      {loadingElenco && <div style={{ color: T.textMid, fontSize: "12px" }}>Caricamento elenco da RPN…</div>}

      {elenco.length > 0 && (
        <div>
          <div style={{ display: "flex", gap: 8, marginBottom: 16, alignItems: "center" }}>
            <input
              placeholder="Anno"
              value={anno}
              onChange={e => setAnno(e.target.value)}
              style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 3, padding: "6px 10px", color: T.text, fontSize: "12px", width: 80 }}
            />
            <select
              value={statoFiltro}
              onChange={e => setStatoFiltro(e.target.value)}
              style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 3, padding: "6px 10px", color: T.text, fontSize: "12px" }}
            >
              <option value="">Tutti gli stati</option>
              {statiDisponibili.map(s => <option key={s} value={s}>{s}</option>)}
            </select>
            <button style={css.btn()} onClick={caricaElenco}>↻ Ricarica elenco</button>
            <div style={{ marginLeft: "auto" }}>
              <button style={css.btn("accent")} onClick={handleImporta} disabled={importando || selezionati.size === 0}>
                {importando ? "Import in corso..." : `Importa ${selezionati.size} selezionate`}
              </button>
            </div>
          </div>

          {/* ─── GIRI ─── */}
          <div style={{ color: T.text, fontWeight: "700", fontSize: "12px", marginBottom: 6 }}>Giri</div>
          <input
            placeholder="Cerca per nome giro..."
            value={ricerca}
            onChange={e => setRicerca(e.target.value)}
            style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 3, padding: "6px 10px", color: T.text, fontSize: "12px", width: "100%", marginBottom: 8 }}
          />
          <div style={{ overflowX: "auto", maxHeight: 300, overflowY: "auto", border: `1px solid ${T.border}`, marginBottom: 24 }}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead>
                <tr>
                  <th style={css.th}><input type="checkbox" checked={filtratiGiri.length > 0 && filtratiGiri.every(c => selezionati.has(c.key))} onChange={toggleTuttiGiri} /></th>
                  <th style={css.th}>Nome cedola</th>
                  <th style={css.th}>Giro</th>
                  <th style={css.th}>Stato</th>
                  <th style={css.th}>N. titoli</th>
                </tr>
              </thead>
              <tbody>
                {filtratiGiri.map((c, i) => (
                  <tr key={c.key} style={{ background: i % 2 === 0 ? "transparent" : T.surface + "66", cursor: "pointer" }} onClick={() => toggle(c.key)}>
                    <td style={css.td}><input type="checkbox" checked={selezionati.has(c.key)} onChange={() => toggle(c.key)} onClick={e => e.stopPropagation()} /></td>
                    <td style={{ ...css.td, fontWeight: "600" }}>{c.nome}</td>
                    <td style={{ ...css.td, color: T.textMid }}>{c.giroNome ?? "—"}</td>
                    <td style={{ ...css.td, color: T.textMid }}>{c.stato ?? "—"}</td>
                    <td style={{ ...css.td, color: T.textMid, textAlign: "right" }}>{c.numeroTitoli}</td>
                  </tr>
                ))}
                {filtratiGiri.length === 0 && (
                  <tr><td colSpan={5} style={{ ...css.td, color: T.textDim, textAlign: "center", padding: 16 }}>Nessun giro trovato</td></tr>
                )}
              </tbody>
            </table>
          </div>

          {/* ─── CEDOLE EXTRA: tendina ricercabile invece dell'elenco piatto ─── */}
          <div style={{ color: T.text, fontWeight: "700", fontSize: "12px", marginBottom: 6 }}>Cedole extra</div>
          <div style={{ position: "relative", marginBottom: 10 }}>
            <input
              placeholder="Cerca cedola extra per nome e clicca per aggiungerla..."
              value={extraQuery}
              onChange={e => { setExtraQuery(e.target.value); setExtraOpen(true); }}
              onFocus={() => setExtraOpen(true)}
              onBlur={() => setTimeout(() => setExtraOpen(false), 150)}
              style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 3, padding: "6px 10px", color: T.text, fontSize: "12px", width: "100%" }}
            />
            {extraOpen && (
              <div style={{ position: "absolute", top: "100%", left: 0, right: 0, zIndex: 10, background: T.surface, border: `1px solid ${T.borderHi}`, borderRadius: 3, maxHeight: 260, overflowY: "auto", boxShadow: "0 4px 12px rgba(0,0,0,0.4)" }}>
                {extraCandidati.length === 0 && (
                  <div style={{ padding: "10px 12px", color: T.textDim, fontSize: "12px" }}>Nessuna cedola extra trovata</div>
                )}
                {extraCandidati.map(c => {
                  const sel = selezionati.has(c.key);
                  return (
                    <div
                      key={c.key}
                      onMouseDown={e => e.preventDefault()} // evita che il blur chiuda la tendina prima del click
                      onClick={() => toggle(c.key)}
                      style={{ padding: "7px 12px", fontSize: "12px", cursor: "pointer", display: "flex", justifyContent: "space-between", gap: 8, background: sel ? T.accent + "22" : "transparent", borderBottom: `1px solid ${T.border}22` }}
                    >
                      <span style={{ color: sel ? T.accent : T.text, fontWeight: sel ? "700" : "400" }}>{sel ? "✓ " : ""}{c.nome}</span>
                      <span style={{ color: T.textMid, whiteSpace: "nowrap" }}>{c.stato ? `${c.stato} · ` : ""}{c.numeroTitoli} tit.</span>
                    </div>
                  );
                })}
              </div>
            )}
          </div>

          {extraSelezionati.length > 0 && (
            <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 16 }}>
              {extraSelezionati.map(c => (
                <div key={c.key} style={{ display: "flex", alignItems: "center", gap: 6, background: T.accent + "22", border: `1px solid ${T.accent}66`, borderRadius: 12, padding: "3px 6px 3px 10px", fontSize: "11px", color: T.accent }}>
                  {c.nome}
                  <span onClick={() => toggle(c.key)} style={{ cursor: "pointer", fontWeight: "700", padding: "0 4px" }}>✕</span>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {log.length > 0 && (
        <div style={{ marginTop: 20 }}>
          <div style={{ color: T.text, fontWeight: "700", fontSize: "13px", marginBottom: 8 }}>Risultato import</div>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>{["Cedola", "Tipo", "Stato", "Dettaglio"].map(h => <th key={h} style={css.th}>{h}</th>)}</tr>
            </thead>
            <tbody>
              {log.map((l, i) => (
                <tr key={i}>
                  <td style={{ ...css.td, fontWeight: "600" }}>{l.nome}</td>
                  <td style={css.td}>{l.tipo === "giro" ? "Giro" : "Cedola extra"}</td>
                  <td style={{ ...css.td, color: l.stato === "ok" ? T.green : l.stato === "errore" ? T.red : T.textMid, fontWeight: "700" }}>
                    {l.stato === "pending" ? "…" : l.stato === "ok" ? "✓" : "✗"}
                  </td>
                  <td style={{ ...css.td, color: T.textMid, fontSize: "11px" }}>
                    {l.dettaglio && (() => {
                      const erroriVeri = (l.dettaglio.errori || []).filter(e => !e.startsWith("ℹ"));
                      const base = `${l.dettaglio.creati ?? 0} creati, ${l.dettaglio.aggiornati ?? 0} aggiornati${l.dettaglio.ignorati ? `, ${l.dettaglio.ignorati} ignorati` : ""}`;
                      return erroriVeri.length ? `${base} — ⚠ ${erroriVeri.join("; ")}` : base;
                    })()}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
