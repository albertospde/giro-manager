import { useState, useEffect, useMemo, useCallback } from "react";
import ModuloImport, { ImportSpalmatura } from "./ModuloImport.jsx";
import ModuloCaricoSemplice from "./ModuloCaricoSemplice.jsx";
import ModuloPrenotato from "./ModuloPrenotato.jsx";
import ModuloAvanzamento from "./ModuloAvanzamento.jsx";

const SUPABASE_URL = "https://tdflwenlylhctxssatax.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InRkZmx3ZW5seWxoY3R4c3NhdGF4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYzMzgyNzYsImV4cCI6MjA5MTkxNDI3Nn0.l35qEL7LOvyYuI1McQlVqj4vbyTqmlevcmqWbTGYi2Q";

const normEditoreKey = n => (n || "").trim().toLowerCase();

const ACCOUNT_BY_COD = {
  "103": "CALAVETTA",
  "108": "CALAVETTA",
  "112": "PASSARINI",
  "118": "CALAVETTA",
  "12": "CALAVETTA",
  "134": "PARRELLA",
  "137": "CALAVETTA",
  "138": "PASSARINI",
  "152": "CALAVETTA",
  "22": "PARRELLA",
  "277": "PARRELLA",
  "290": "CALAVETTA",
  "296": "CALAVETTA",
  "299": "CALAVETTA",
  "305": "VENTURA",
  "308": "CALAVETTA",
  "320": "PASSARINI",
  "326": "VENTURA",
  "333": "CALAVETTA",
  "347": "CALAVETTA",
  "366": "CALAVETTA",
  "37": "CALAVETTA",
  "386": "CALAVETTA",
  "422": "SALA",
  "428": "PALLOTTA",
  "437": "CALAVETTA",
  "455": "VENTURA",
  "465": "PASSARINI",
  "468": "PARRELLA",
  "487": "SALA",
  "502": "CALAVETTA",
  "510": "CALAVETTA",
  "522": "BRIZZI",
  "524": "PARRELLA",
  "532": "PASSARINI",
  "534": "CALAVETTA",
  "565": "PALLOTTA",
  "568": "CALAVETTA",
  "58": "CALAVETTA",
  "64": "CALAVETTA",
  "671": "PARRELLA",
  "683": "CALAVETTA",
  "692": "CALAVETTA",
  "710": "PASSARINI",
  "72": "SALA",
  "724": "CALAVETTA",
  "73": "PARRELLA",
  "735": "VENTURA",
  "738": "VENTURA",
  "739": "PASSARINI",
  "747": "PASSARINI",
  "750": "PASSARINI",
  "751": "PASSARINI",
  "753": "PASSARINI",
  "763": "BRIZZI",
  "77": "CALAVETTA",
  "773": "PARRELLA",
  "776": "CALAVETTA",
  "780": "CALAVETTA",
  "795": "PASSARINI",
  "821": "SALA",
  "854": "CALAVETTA",
  "909": "PARRELLA",
  "919": "CALAVETTA",
  "934": "CALAVETTA",
  "935": "PARRELLA",
  "937": "CALAVETTA",
  "941": "PASSARINI",
  "945": "CALAVETTA",
  "952": "CALAVETTA",
  "992": "CALAVETTA",
  "995": "CALAVETTA",
  "A21": "CALAVETTA",
  "A54": "PASSARINI",
  "A58": "CALAVETTA",
  "A59": "BRIZZI",
  "A60": "PASSARINI",
  "A61": "SALA",
  "A77": "CALAVETTA",
  "A86": "PARRELLA",
  "A94": "VALZASINA",
  "A98": "SALA",
  "AC2": "CALAVETTA",
  "AC7": "CALAVETTA",
  "AD6": "SALA",
  "AL5": "VENTURA",
  "AL9": "BRIZZI",
  "B44": "PARRELLA",
  "B58": "PARRELLA",
  "B59": "PARRELLA",
  "B81": "PARRELLA",
  "C48": "CALAVETTA",
  "C52": "CALAVETTA",
  "C72": "PASSARINI",
  "C79": "VENTURA",
  "C85": "CALAVETTA",
  "D52": "SALA",
  "D65": "SALA",
  "D71": "SALA",
  "D80": "PARRELLA",
  "D92": "PASSARINI",
  "D95": "BRIZZI"
};

const sb = {
  auth: {
    signIn: async (email, password) => {
      const r = await fetch(`${SUPABASE_URL}/auth/v1/token?grant_type=password`, {
        method: "POST", headers: { "Content-Type": "application/json", "apikey": SUPABASE_KEY },
        body: JSON.stringify({ email, password }),
      });
      return r.json();
    },
    signOut: async (token) => {
      await fetch(`${SUPABASE_URL}/auth/v1/logout`, { method: "POST", headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}` } });
    },
    getUser: async (token) => {
      const r = await fetch(`${SUPABASE_URL}/auth/v1/user`, { headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}` } });
      return r.json();
    },
  },
};

const sbFetch = async (path, token) => {
  const r = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`, "Accept": "application/json", "Range-Unit": "items", "Range": "0-499999" },
  });
  return r.json();
};

// FIX 5: Funzione per salvare un titolo su Supabase (PATCH)
const sbUpdateTitolo = async (id, updates, token) => {
  const r = await fetch(`${SUPABASE_URL}/rest/v1/titoli?id=eq.${id}`, {
    method: "PATCH",
    headers: {
      "apikey": SUPABASE_KEY,
      "Authorization": `Bearer ${token}`,
      "Content-Type": "application/json",
      "Prefer": "return=minimal",
    },
    body: JSON.stringify(updates),
  });
  return r.ok;
};

// Cancellazione titolo da Giri e Cedole (DELETE)
const sbDeleteTitolo = async (id, token) => {
  const r = await fetch(`${SUPABASE_URL}/rest/v1/titoli?id=eq.${id}`, {
    method: "DELETE",
    headers: {
      "apikey": SUPABASE_KEY,
      "Authorization": `Bearer ${token}`,
      "Prefer": "return=minimal",
    },
  });
  return r.ok;
};

const T = {
  bg: "#1a2140", surface: "#212d54", border: "#2e3d6b", borderHi: "#3d4f82",
  text: "#f0f2f8", textMid: "#8b9cc8", textDim: "#4a5a8a",
  accent: "#7b9fe8", green: "#4caf7d", red: "#e05c5c", blue: "#4a5da0", purple: "#9c6fcf",
};

const css = {
  app: { background: T.bg, color: T.text, minHeight: "100vh", fontFamily: "'Inter', 'Segoe UI', Arial, sans-serif", fontSize: "13px" },
  sidebar: { width: 200, background: T.surface, borderRight: `1px solid ${T.border}`, display: "flex", flexDirection: "column", flexShrink: 0 },
  main: { flex: 1, overflow: "hidden", display: "flex", flexDirection: "column" },
  header: { borderBottom: `1px solid ${T.border}`, padding: "12px 20px", display: "flex", alignItems: "center", gap: 12, background: T.surface },
  btn: (v = "default") => ({ padding: "6px 14px", border: `1px solid ${v === "accent" ? T.accent : T.border}`, background: v === "accent" ? T.accent : "transparent", color: v === "accent" ? "#000" : T.text, cursor: "pointer", fontSize: "12px", fontFamily: "inherit", borderRadius: 3, fontWeight: v === "accent" ? "700" : "400", letterSpacing: "0.04em", transition: "all 0.15s" }),
  input: { background: T.bg, border: `1px solid ${T.border}`, color: T.text, padding: "5px 10px", fontSize: "12px", fontFamily: "inherit", borderRadius: 3, outline: "none" },
  tag: (c = T.accent) => ({ display: "inline-block", padding: "2px 7px", background: c + "22", border: `1px solid ${c}44`, color: c, borderRadius: 2, fontSize: "10px", fontWeight: "700", letterSpacing: "0.06em", textTransform: "uppercase" }),
  table: { width: "100%", borderCollapse: "collapse" },
  th: { padding: "8px 12px", textAlign: "left", color: T.textMid, fontWeight: "400", fontSize: "11px", letterSpacing: "0.08em", textTransform: "uppercase", borderBottom: `1px solid ${T.border}`, whiteSpace: "nowrap", position: "sticky", top: 0, background: T.surface, zIndex: 1 },
  td: { padding: "8px 12px", borderBottom: `1px solid ${T.border}22`, verticalAlign: "middle" },
};

const CANALI_RETE = ["LIBRACCIO", "LIB_RELIGIOSE", "LIB_COOP", "INDIPENDENTI_ALTRE_CATENE"];

// MOD 1: Gruppo 94 (AURORA → DIRETTI_TIPOGRAFIA) ora in GROSSISTI
// MOD 2: "Aurora" rinominata in "Diretti da Tipografia"
const MACROGRUPPI = [
  { id: "RETE", label: "Rete", canali: ["LIBRACCIO", "LIB_RELIGIOSE", "LIB_COOP", "INDIPENDENTI_ALTRE_CATENE"] },
  { id: "CATENE", label: "Catene Centralizzate", canali: ["FELTRINELLI", "MONDADORI", "UBIK", "GIUNTI"] },
  { id: "GROSSISTI", label: "Grossisti", canali: ["FASTBOOK", "CENTROLIBRI", "GROSSISTI", "AURORA"] },
  { id: "ONLINE", label: "Online", canali: ["AMAZON", "IBS", "ALTRI_ONLINE"] },
];

// Label di visualizzazione canali (MOD 2: AURORA → "Diretti da Tipografia")
const CANALE_DISPLAY_NAMES = {
  AURORA: "Diretti da Tipografia",
};
// Helper per ottenere il nome visualizzato di un canale
const getCanaleDisplayName = (canale) => {
  if (!canale) return "";
  return CANALE_DISPLAY_NAMES[canale.codice] || canale.nome;
};

function LoginScreen({ onLogin }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const handleLogin = async () => {
    setLoading(true); setError("");
    const data = await sb.auth.signIn(email, password);
    if (data.access_token) { localStorage.setItem("giro_token", data.access_token); onLogin(data.access_token, data.user); }
    else setError("Email o password errati.");
    setLoading(false);
  };

  return (
    <div style={{ ...css.app, display: "flex", alignItems: "center", justifyContent: "center", height: "100vh" }}>
      <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, padding: 40, width: 340 }}>
        <div style={{ textAlign: "center", marginBottom: 32 }}>
          <div style={{ display: "inline-block", background: "#fff", borderRadius: 12, padding: "10px 16px", marginBottom: 16 }}>
            <img src="https://raw.githubusercontent.com/albertospde/giro-manager/main/.github/logo_pde.png" style={{ height: 48, display: "block" }} alt="PDE" />
          </div>
          <div style={{ color: T.accent, fontSize: "24px", fontWeight: "700", letterSpacing: "0.1em" }}>GIRO</div>
          <div style={{ color: T.textMid, fontSize: "11px", letterSpacing: "0.15em", marginTop: 4 }}>MANAGER</div>
        </div>
        <div style={{ marginBottom: 16 }}>
          <label style={{ color: T.textMid, fontSize: "10px", letterSpacing: "0.08em", textTransform: "uppercase", display: "block", marginBottom: 6 }}>Email</label>
          <input style={{ ...css.input, width: "100%", boxSizing: "border-box" }} type="email" value={email} onChange={e => setEmail(e.target.value)} onKeyDown={e => e.key === "Enter" && handleLogin()} />
        </div>
        <div style={{ marginBottom: 24 }}>
          <label style={{ color: T.textMid, fontSize: "10px", letterSpacing: "0.08em", textTransform: "uppercase", display: "block", marginBottom: 6 }}>Password</label>
          <input style={{ ...css.input, width: "100%", boxSizing: "border-box" }} type="password" value={password} onChange={e => setPassword(e.target.value)} onKeyDown={e => e.key === "Enter" && handleLogin()} />
        </div>
        {error && <div style={{ color: T.red, fontSize: "12px", marginBottom: 16, textAlign: "center" }}>{error}</div>}
        <button style={{ ...css.btn("accent"), width: "100%", padding: "10px" }} onClick={handleLogin} disabled={loading}>{loading ? "Accesso..." : "Accedi"}</button>
        <div style={{ textAlign: "center", marginTop: 20, color: T.textDim, fontSize: "10px", letterSpacing: "0.06em" }}>Powered by PDE</div>
      </div>
    </div>
  );
}

function Badge({ label, color }) { return <span style={css.tag(color)}>{label}</span>; }

function ProgressBar({ value, total, color = T.accent }) {
  const pct = total > 0 ? Math.min(100, Math.round((value / total) * 100)) : 0;
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
      <div style={{ width: 80, height: 4, background: T.borderHi, borderRadius: 2, overflow: "hidden" }}>
        <div style={{ width: `${pct}%`, height: "100%", background: pct >= 100 ? T.green : color }} />
      </div>
      <span style={{ color: T.textMid, fontSize: "11px" }}>{pct}%</span>
    </div>
  );
}

function KpiCard({ label, value, sub, color = T.accent }) {
  return (
    <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 4, padding: "16px 20px", minWidth: 160 }}>
      <div style={{ color: T.textMid, fontSize: "10px", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 8 }}>{label}</div>
      <div style={{ color, fontSize: "28px", fontWeight: "700", lineHeight: 1 }}>{value}</div>
      {sub && <div style={{ color: T.textMid, fontSize: "11px", marginTop: 6 }}>{sub}</div>}
    </div>
  );
}

// FIX 5: EditModal ora riceve anche token e onSaveDB per salvare su Supabase
function EditModal({ titolo, siblings = [], onSave, onClose, onDelete, token }) {
  const [form, setForm] = useState({ ...titolo });
  const [saving, setSaving] = useState(false);
  const [deleting, setDeleting] = useState(false);
  const set = k => e => setForm(f => ({ ...f, [k]: e.target.type === "checkbox" ? e.target.checked : e.target.value }));

  const handleDelete = async () => {
    const conferma = window.confirm(`Cancellare definitivamente "${titolo.titolo || titolo.ean}" dalla cedola? L'operazione non è reversibile.`);
    if (!conferma) return;
    setDeleting(true);
    const ok = await sbDeleteTitolo(titolo.id, token);
    setDeleting(false);
    if (!ok) { alert("Errore nella cancellazione su database."); return; }
    onDelete && onDelete(titolo.id);
    onClose();
  };

  const handleSave = async () => {
    setSaving(true);

    // Sequenza per editore (stessa cedola + stesso editore): se la posizione richiesta cambia, riordina i titoli "fratelli"
    let formFinal = form;
    const posOriginale = titolo.posizione ?? null;
    const posRichiesta = form.posizione === "" || form.posizione == null ? null : parseInt(form.posizione, 10);
    if (posRichiesta != null && posRichiesta !== posOriginale) {
      const gruppo = siblings
        .filter(t => t.id !== form.id && t.giro_label === form.giro_label && t.n_cedola === form.n_cedola && t.editore_nome === form.editore_nome)
        .sort((a, b) => (a.posizione || 0) - (b.posizione || 0));
      const posClamp = Math.max(1, Math.min(posRichiesta, gruppo.length + 1));
      gruppo.splice(posClamp - 1, 0, { id: form.id, posizione: posClamp });
      for (let idx = 0; idx < gruppo.length; idx++) {
        const nuovaPos = idx + 1;
        if (gruppo[idx].id !== form.id && gruppo[idx].posizione !== nuovaPos) {
          const ok = await sbUpdateTitolo(gruppo[idx].id, { posizione: nuovaPos }, token);
          if (ok) onSave({ ...gruppo[idx], posizione: nuovaPos });
        }
      }
      formFinal = { ...form, posizione: String(posClamp) };
    }

    // Prepara l'oggetto con solo i campi modificabili
    const payload = {};
    const editableFields = [
      "titolo","autore","editore_nome","ean","prezzo","uscita","formato","eta",
      "account_editore","promozione","obiettivo_assegnato","obiettivo_raggiunto","posizione","ranking_editore",
      "il_triangolo","top_100","ean_gemello_1","titolo_gemello_1","ean_gemello_2",
      "titolo_gemello_2","ean_gemello_3","titolo_gemello_3","note_comunicazione","note"
    ];
    const numFields = ["prezzo","obiettivo_assegnato","obiettivo_raggiunto","posizione","ranking_editore"];
    editableFields.forEach(k => {
      let valForm = formFinal[k];
      let valOrig = titolo[k];
      // Normalizza numerici per confronto corretto (input restituisce stringhe)
      if (numFields.includes(k)) {
        valForm = valForm === "" || valForm == null ? null : Number(valForm);
        valOrig = valOrig == null ? null : Number(valOrig);
      }
      if (valForm !== valOrig) {
        payload[k] = valForm;
      }
    });

    if (Object.keys(payload).length > 0 && token) {
      const ok = await sbUpdateTitolo(formFinal.id, payload, token);
      if (!ok) { alert("Errore nel salvataggio su database."); setSaving(false); return; }
    }
    // Normalizza i tipi numerici prima di aggiornare lo state React
    const formNorm = { ...formFinal };
    ["prezzo","obiettivo_assegnato","obiettivo_raggiunto","posizione","ranking_editore"].forEach(k => {
      if (formNorm[k] !== "" && formNorm[k] != null) formNorm[k] = Number(formNorm[k]);
    });
    onSave(formNorm);
    setSaving(false);
    onClose();
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "#000a", zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center" }} onClick={onClose}>
      <div style={{ background: T.surface, border: `1px solid ${T.borderHi}`, borderRadius: 6, padding: 24, width: 640, maxHeight: "85vh", overflowY: "auto" }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <span style={{ color: T.accent, fontWeight: "700", fontSize: "13px" }}>MODIFICA TITOLO</span>
          <button style={css.btn()} onClick={onClose}>✕</button>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          {[["titolo","Titolo","full"],["autore","Autore"],["editore_nome","Editore"],["ean","EAN"],["prezzo","Prezzo"],["uscita","Uscita"],["formato","Formato"],["eta","Età"],["posizione","Posizione per editore"],["ranking_editore","Ranking editore"],["account_editore","Account"],["promozione","Promozione"],["obiettivo_assegnato","Obiettivo assegnato"],["obiettivo_raggiunto","Obiettivo raggiunto"]].map(([k, label, span]) => (
            <div key={k} style={span === "full" ? { gridColumn: "1/-1" } : {}}>
              <label style={{ color: T.textMid, fontSize: "10px", letterSpacing: "0.08em", textTransform: "uppercase", display: "block", marginBottom: 4 }}>{label}</label>
              <input style={{ ...css.input, width: "100%", boxSizing: "border-box" }} value={form[k] ?? ""} onChange={set(k)} />
            </div>
          ))}
        </div>
        <div style={{ marginTop: 16, display: "flex", gap: 12 }}>
          {["il_triangolo","top_100"].map(k => (
            <label key={k} style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer", color: T.textMid, fontSize: "12px" }}>
              <input type="checkbox" checked={!!form[k]} onChange={set(k)} style={{ accentColor: T.accent }} />
              {k === "il_triangolo" ? "▲ Il Triangolo" : "★ Top 100"}
            </label>
          ))}
        </div>
        <div style={{ marginTop: 16 }}>
          <div style={{ color: T.textMid, fontSize: "10px", letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 8 }}>GEMELLI</div>
          {[1,2,3].map(n => (
            <div key={n} style={{ display: "grid", gridTemplateColumns: "1fr 2fr", gap: 8, marginBottom: 8 }}>
              <input style={css.input} placeholder={`EAN gemello ${n}`} value={form[`ean_gemello_${n}`] ?? ""} onChange={set(`ean_gemello_${n}`)} />
              <input style={css.input} placeholder={`Titolo gemello ${n}`} value={form[`titolo_gemello_${n}`] ?? ""} onChange={set(`titolo_gemello_${n}`)} />
            </div>
          ))}
        </div>
        <div style={{ marginTop: 16 }}>
          <label style={{ color: T.textMid, fontSize: "10px", letterSpacing: "0.08em", textTransform: "uppercase", display: "block", marginBottom: 4 }}>NOTE BY COMUNICAZIONE</label>
          <textarea style={{ ...css.input, width: "100%", boxSizing: "border-box", height: 60, resize: "vertical" }} value={form.note_comunicazione ?? ""} onChange={set("note_comunicazione")} />
        </div>
        <div style={{ marginTop: 10 }}>
          <label style={{ color: T.textMid, fontSize: "10px", letterSpacing: "0.08em", textTransform: "uppercase", display: "block", marginBottom: 4 }}>NOTE INTERNE</label>
          <textarea style={{ ...css.input, width: "100%", boxSizing: "border-box", height: 60, resize: "vertical" }} value={form.note ?? ""} onChange={set("note")} />
        </div>
        <div style={{ marginTop: 20, display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
          <button style={{ ...css.btn(), color: T.red, borderColor: T.red }} onClick={handleDelete} disabled={saving || deleting}>
            {deleting ? "Cancellazione..." : "🗑 Cancella riga"}
          </button>
          <div style={{ display: "flex", gap: 8 }}>
            <button style={css.btn()} onClick={onClose} disabled={deleting}>Annulla</button>
            <button style={css.btn("accent")} onClick={handleSave} disabled={saving || deleting}>
              {saving ? "Salvataggio..." : "Salva"}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

function ModuloDashboard({ titoli, prenotato, canali, spalmatura, ruolo }) {
  const anniDisp = useMemo(() => {
    const s = new Set();
    titoli.forEach(t => { if (t.giro_label) { const yr = Number(t.giro_label.split(" ")[1]); if (yr >= 2020) s.add(yr); } });
    return [...s].sort((a, b) => b - a);
  }, [titoli]);
  const annoCorrente = new Date().getFullYear();
  const [filterAnno, setFilterAnno] = useState([annoCorrente]);

  const giriLabel = useMemo(() => {
    return [...new Set(titoli.map(t => t.giro_label).filter(Boolean))]
      .filter(g => filterAnno.length === 0 || filterAnno.includes(Number(g.split(" ")[1])))
      .sort((a, b) => {
        const [na, ya] = a.split(" "); const [nb, yb] = b.split(" ");
        return Number(yb || 0) - Number(ya || 0) || Number(nb || 0) - Number(na || 0);
      });
  }, [titoli, filterAnno]);

  const [giriSel, setGiriSel] = useState([]);
  useEffect(() => { if (giriLabel.length > 0 && giriSel.length === 0) setGiriSel([giriLabel[0]]); }, [giriLabel]);

  const titoliGiro = useMemo(() => giriSel.length > 0 ? titoli.filter(t => giriSel.includes(t.giro_label)) : [], [titoli, giriSel]);
  const prenotatoGiro = useMemo(() => { const ids = new Set(titoliGiro.map(t => t.id)); return prenotato.filter(p => ids.has(p.titolo_id)); }, [prenotato, titoliGiro]);
  const totPrenotatoGiro = useMemo(() => prenotatoGiro.reduce((s, p) => s + p.quantita, 0), [prenotatoGiro]);

  const kpiGiro = useMemo(() => {
    const totObj = titoliGiro.reduce((s, t) => s + (t.obiettivo_assegnato || 0), 0);
    const valObj = titoliGiro.reduce((s, t) => s + (t.prezzo || 0) * (t.obiettivo_assegnato || 0), 0);
    const totTriangolo = titoliGiro.filter(t => t.il_triangolo === true).length;
    const totTop100 = titoliGiro.filter(t => t.top_100 === true).length;
    return { totObj, valObj, totTriangolo, totTop100, count: titoliGiro.length, pct: totObj > 0 ? Math.round(totPrenotatoGiro / totObj * 100) : 0 };
  }, [titoliGiro, totPrenotatoGiro]);

  const cedole = useMemo(() => {
    const prenMap = {};
    prenotatoGiro.forEach(p => { const t = titoliGiro.find(t => t.id === p.titolo_id); if (!t) return; const c = t.n_cedola || "—"; prenMap[c] = (prenMap[c] || 0) + p.quantita; });
    const map = {};
    titoliGiro.forEach(t => { const c = t.n_cedola || "—"; if (!map[c]) map[c] = { label: c, totObj: 0, totRag: 0, count: 0 }; map[c].totObj += t.obiettivo_assegnato || 0; map[c].totRag = prenMap[c] || 0; map[c].count++; });
    return Object.values(map).sort((a, b) => a.label.localeCompare(b.label));
  }, [titoliGiro, prenotatoGiro]);

  const prenotatoPerCanale = useMemo(() => {
    const map = {};
    const CANALI_VIS_AGENTE = [...CANALI_RETE, "FASTBOOK", "CENTROLIBRI", "GROSSISTI"];
    prenotatoGiro.forEach(p => { const c = canali.find(c => c.id === p.canale_id); if (!c) return; if (ruolo === "agente" && !CANALI_VIS_AGENTE.includes(c.codice)) return; map[c.codice] = (map[c.codice] || 0) + p.quantita; });
    return map;
  }, [prenotatoGiro, canali, ruolo]);

  const macrogruppiVis = ruolo === "agente" ? MACROGRUPPI.filter(mg => mg.id === "RETE" || mg.id === "GROSSISTI") : MACROGRUPPI;
  const totMacro = useMemo(() => { const map = {}; macrogruppiVis.forEach(mg => { map[mg.id] = mg.canali.reduce((s, cod) => s + (prenotatoPerCanale[cod] || 0), 0); }); return map; }, [prenotatoPerCanale, macrogruppiVis]);
  const maxMacro = Math.max(...Object.values(totMacro), 1);

  // Obiettivo per canale (via spalmatura) integrato nella sezione unica
  const obiPerCanale = useMemo(() => {
    const map = {};
    canali.forEach(c => {
      let assegnato = 0;
      titoliGiro.forEach(t => {
        const spRow = spalmatura.find(s => s.editore_nome === t.editore_nome && s.formato === (t.formato || 'Cover') && s.canale_codice === c.codice);
        if (spRow && t.obiettivo_assegnato) assegnato += Math.round(t.obiettivo_assegnato * spRow.percentuale);
      });
      map[c.codice] = { assegnato };
    });
    return map;
  }, [titoliGiro, canali, spalmatura]);

  return (
    <div style={{ flex: 1, overflowY: "auto", padding: 20 }}>
      <div style={{ display: "flex", gap: 12, alignItems: "center", marginBottom: 20 }}>
        <SearchableMultiSelect values={filterAnno.map(String)} onChange={v => { setFilterAnno(v.map(Number)); setGiriSel([]); }} options={anniDisp.map(String)} renderOption={v => v} placeholder="Anno" width={120} />
        <SearchableMultiSelect values={giriSel} onChange={setGiriSel} options={giriLabel} renderOption={g => `Giro ${g}`} placeholder="Seleziona giro" width={180} />
        <span style={{ color: T.textMid, fontSize: "12px" }}>{kpiGiro.count} titoli · € {kpiGiro.valObj.toLocaleString("it", { maximumFractionDigits: 0 })} valore obiettivo</span>
      </div>
      <div style={{ display: "flex", gap: 12, marginBottom: 24, flexWrap: "wrap" }}>
        <KpiCard label="Titoli" value={kpiGiro.count} color={T.text} />
        <KpiCard label="▲ Triangolo" value={kpiGiro.totTriangolo} color={T.purple} />
        <KpiCard label="★ Top 100" value={kpiGiro.totTop100} color={T.accent} />
        <KpiCard label="Prenotato" value={totPrenotatoGiro.toLocaleString("it")} color={T.green} sub="copie" />
        <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 4, padding: "16px 20px", minWidth: 200 }}>
          <div style={{ color: T.textMid, fontSize: "10px", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 8 }}>Avanzamento obiettivo</div>
          <div style={{ color: T.accent, fontSize: "24px", fontWeight: "700", lineHeight: 1, marginBottom: 8 }}>{kpiGiro.pct}%</div>
          <div style={{ height: 6, background: T.borderHi, borderRadius: 3, overflow: "hidden" }}>
            <div style={{ width: `${kpiGiro.pct}%`, height: "100%", background: kpiGiro.pct >= 80 ? T.green : kpiGiro.pct >= 50 ? T.accent : T.red }} />
          </div>
          <div style={{ color: T.textMid, fontSize: "11px", marginTop: 6 }}>{totPrenotatoGiro.toLocaleString("it")} / {kpiGiro.totObj.toLocaleString("it")}</div>
        </div>
      </div>

      {/* CEDOLE */}
      <div style={{ marginBottom: 24 }}>
        <div style={{ color: T.textMid, fontSize: "11px", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 12 }}>CEDOLE</div>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr>{["Cedola","Titoli","Obj Ass.","Prenotato","Avanz."].map(h => <th key={h} style={{ padding: "8px 12px", textAlign: "left", color: T.textMid, fontWeight: "400", fontSize: "11px", letterSpacing: "0.08em", textTransform: "uppercase", borderBottom: `1px solid ${T.border}`, background: T.surface }}>{h}</th>)}</tr></thead>
          <tbody>
            {cedole.map(({ label, count, totObj, totRag }, i) => {
              const pct = totObj > 0 ? Math.round(totRag / totObj * 100) : 0;
              return (
                <tr key={label} style={{ background: i % 2 === 0 ? "transparent" : T.surface + "66" }}>
                  <td style={{ padding: "8px 12px", color: T.accent, fontWeight: "600", fontSize: "12px" }}>{label}</td>
                  <td style={{ padding: "8px 12px", fontSize: "12px" }}>{count}</td>
                  <td style={{ padding: "8px 12px", fontSize: "12px" }}>{totObj.toLocaleString("it")}</td>
                  <td style={{ padding: "8px 12px", fontSize: "12px", color: T.green }}>{totRag.toLocaleString("it")}</td>
                  <td style={{ padding: "8px 12px" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <div style={{ width: 80, height: 4, background: T.borderHi, borderRadius: 2, overflow: "hidden" }}><div style={{ width: `${pct}%`, height: "100%", background: pct >= 80 ? T.green : pct >= 50 ? T.accent : T.red }} /></div>
                      <span style={{ color: pct >= 80 ? T.green : pct >= 50 ? T.accent : T.red, fontSize: "11px", fontWeight: "700" }}>{pct}%</span>
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* CANALI: prenotato + obiettivi integrati */}
      <div>
        <div style={{ color: T.textMid, fontSize: "11px", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 12 }}>CANALI</div>
        {macrogruppiVis.map(mg => {
          const totMgPren = totMacro[mg.id] || 0;
          const totMgAss = mg.canali.reduce((s, cod) => s + (obiPerCanale[cod]?.assegnato || 0), 0);
          const pctMg = totMgAss > 0 ? Math.round(totMgPren / totMgAss * 100) : 0;
          return (
            <div key={mg.id} style={{ marginBottom: 20 }}>
              {/* Header macrogruppo */}
              <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 4, padding: "10px 16px", background: T.surface, border: `1px solid ${T.borderHi}`, borderRadius: 4 }}>
                <div style={{ fontWeight: "700", color: T.accent, fontSize: "13px", minWidth: 170 }}>{mg.label}</div>
                <div style={{ flex: 1, height: 8, background: T.borderHi, borderRadius: 2, overflow: "hidden" }}>
                  <div style={{ width: `${(totMgPren / maxMacro) * 100}%`, height: "100%", background: T.accent }} />
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                  <div style={{ textAlign: "right" }}>
                    <span style={{ color: T.green, fontWeight: "700", fontSize: "14px" }}>{totMgPren.toLocaleString("it")}</span>
                    {totMgAss > 0 && <span style={{ color: T.textDim, fontSize: "11px" }}> / {totMgAss.toLocaleString("it")}</span>}
                  </div>
                  {totMgAss > 0 && (
                    <span style={{ color: pctMg >= 80 ? T.green : pctMg >= 50 ? T.accent : T.red, fontWeight: "700", fontSize: "12px", minWidth: 36, textAlign: "right" }}>{pctMg}%</span>
                  )}
                </div>
              </div>
              {/* Righe canale */}
              {mg.canali.map(codice => {
                const c = canali.find(c => c.codice === codice); if (!c) return null;
                const qta = prenotatoPerCanale[codice] || 0;
                const obj = obiPerCanale[codice]?.assegnato || 0;
                const pctC = obj > 0 ? Math.round(qta / obj * 100) : 0;
                return (
                  <div key={codice} style={{ display: "flex", alignItems: "center", gap: 12, padding: "5px 16px 5px 32px", borderBottom: `1px solid ${T.border}11` }}>
                    <div style={{ width: 170, fontSize: "12px", color: T.textMid }}>{getCanaleDisplayName(c)}</div>
                    <div style={{ flex: 1, height: 4, background: T.borderHi, borderRadius: 2, overflow: "hidden" }}>
                      <div style={{ width: qta > 0 && totMgPren > 0 ? `${(qta / totMgPren) * 100}%` : "0%", height: "100%", background: T.blue }} />
                    </div>
                    <div style={{ width: 100, textAlign: "right", fontSize: "12px" }}>
                      <span style={{ color: qta > 0 ? T.text : T.textDim, fontWeight: "600" }}>{qta > 0 ? qta.toLocaleString("it") : "—"}</span>
                      {obj > 0 && <span style={{ color: T.textDim, fontSize: "10px" }}> / {obj.toLocaleString("it")}</span>}
                    </div>
                    <div style={{ width: 44, textAlign: "right" }}>
                      {obj > 0 ? (
                        <span style={{ color: pctC >= 80 ? T.green : pctC >= 50 ? T.accent : T.red, fontSize: "11px", fontWeight: "700" }}>{pctC}%</span>
                      ) : (
                        <span style={{ color: T.textMid, fontSize: "11px" }}>{totPrenotatoGiro > 0 && qta > 0 ? `${Math.round(qta / totPrenotatoGiro * 100)}%` : ""}</span>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ===== MODULO CALENDARIO GIRI =====
const MESI_LABELS = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Strenne", "Ottobre", "Novembre", "Dicembre"];

// Ricostruisce l'elenco dei mesi selezionati a partire dal testo salvato (case/separatori variabili)
function parseMesiSelezionati(testo) {
  if (!testo) return [];
  const norm = s => s.toLowerCase().replace(/[^a-zàèéìòù]/g, "");
  const trovati = testo.split(/[,\s]+/).map(norm).filter(Boolean);
  return MESI_LABELS.filter(m => trovati.includes(norm(m)));
}

function fmtDataIt(iso) {
  if (!iso) return "—";
  const [y, m, d] = iso.split("-");
  if (!y || !m || !d) return iso;
  return `${d}/${m}/${y}`;
}

const CAL_GIRI_EMPTY_FORM = { id: null, anno: null, giro: "", mesi: [], consegna_materiali: "", riunioni: "", inizio_giro: "", fine_giro: "", dati_a_editori: "" };

function Modulocalendariogiri({ token, ruolo }) {
  const [righe, setRighe] = useState([]);
  const [loading, setLoading] = useState(true);
  const [anno, setAnno] = useState(new Date().getFullYear());
  const [toast, setToast] = useState(null);
  const [form, setForm] = useState({ ...CAL_GIRI_EMPTY_FORM, anno: new Date().getFullYear() });
  const [saving, setSaving] = useState(false);

  const showToast = (msg, type = "ok") => { setToast({ msg, type }); setTimeout(() => setToast(null), 2500); };

  const load = useCallback(() => {
    setLoading(true);
    sbFetch("calendario_giri?select=*&order=anno.asc,giro.asc", token).then(data => {
      setRighe(Array.isArray(data) ? data : []);
      setLoading(false);
    });
  }, [token]);

  useEffect(() => { load(); }, [load]);

  const anniDisponibili = useMemo(() => {
    const set = new Set([new Date().getFullYear()]);
    righe.forEach(r => { if (r.anno != null) set.add(r.anno); });
    set.add(anno);
    return [...set].sort((a, b) => a - b);
  }, [righe, anno]);

  // Anno assegnato manualmente riga per riga: nessuna logica automatica sulle date,
  // decide tutto Alberto in fase di compilazione.
  const righeAnno = useMemo(() => righe.filter(r => r.anno === anno).sort((a, b) => (a.giro || 0) - (b.giro || 0)), [righe, anno]);

  const resetForm = () => setForm({ ...CAL_GIRI_EMPTY_FORM, anno });
  useEffect(() => { if (form.id === null) setForm(f => ({ ...f, anno })); }, [anno]);

  const startEdit = (riga) => {
    setForm({
      id: riga.id,
      anno: riga.anno ?? anno,
      giro: riga.giro ?? "",
      mesi: parseMesiSelezionati(riga.mese_uscita),
      consegna_materiali: riga.consegna_materiali ?? "",
      riunioni: riga.riunioni ?? "",
      inizio_giro: riga.inizio_giro ?? "",
      fine_giro: riga.fine_giro ?? "",
      dati_a_editori: riga.dati_a_editori ?? "",
    });
    window.scrollTo({ top: 0, behavior: "smooth" });
  };

  const toggleMese = (mese) => setForm(f => ({ ...f, mesi: f.mesi.includes(mese) ? f.mesi.filter(m => m !== mese) : [...f.mesi, mese] }));

  const salvaForm = async () => {
    if (!form.anno) { showToast("Anno obbligatorio", "err"); return; }
    setSaving(true);
    const payload = {
      anno: Number(form.anno),
      giro: form.giro === "" ? null : Number(form.giro),
      mese_uscita: form.mesi.length > 0 ? form.mesi.join(", ") : null,
      consegna_materiali: form.consegna_materiali || null,
      riunioni: form.riunioni || null,
      inizio_giro: form.inizio_giro || null,
      fine_giro: form.fine_giro || null,
      dati_a_editori: form.dati_a_editori || null,
      updated_at: new Date().toISOString(),
    };
    if (form.id) {
      const r = await fetch(`${SUPABASE_URL}/rest/v1/calendario_giri?id=eq.${form.id}`, {
        method: "PATCH",
        headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`, "Content-Type": "application/json", "Prefer": "return=representation" },
        body: JSON.stringify(payload),
      });
      const data = await r.json();
      if (r.ok && data?.[0]) { setRighe(prev => prev.map(x => x.id === form.id ? data[0] : x)); showToast("Riga aggiornata"); resetForm(); }
      else showToast("Errore salvataggio", "err");
    } else {
      const r = await fetch(`${SUPABASE_URL}/rest/v1/calendario_giri`, {
        method: "POST",
        headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`, "Content-Type": "application/json", "Prefer": "return=representation" },
        body: JSON.stringify({ ...payload, ordine: righe.length }),
      });
      const data = await r.json();
      if (r.ok && data?.[0]) { setRighe(prev => [...prev, data[0]]); showToast("Giro aggiunto"); resetForm(); }
      else showToast("Errore creazione riga", "err");
    }
    setSaving(false);
  };

  const eliminaRiga = async (id) => {
    if (!confirm("Eliminare questa riga del calendario?")) return;
    const r = await fetch(`${SUPABASE_URL}/rest/v1/calendario_giri?id=eq.${id}`, {
      method: "DELETE",
      headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}` },
    });
    if (r.ok) { setRighe(prev => prev.filter(x => x.id !== id)); if (form.id === id) resetForm(); showToast("Riga eliminata"); }
    else showToast("Errore eliminazione", "err");
  };

  if (loading) return <div style={{ padding: 40, color: T.textMid, fontSize: "13px" }}>Caricamento calendario...</div>;

  const campo = (label, children) => (
    <div>
      <label style={{ color: T.textMid, fontSize: "10px", letterSpacing: "0.08em", textTransform: "uppercase", display: "block", marginBottom: 4 }}>{label}</label>
      {children}
    </div>
  );

  return (
    <div style={{ flex: 1, overflow: "auto", padding: 20 }}>
      {/* FORM DI INSERIMENTO / MODIFICA */}
      <div style={{ background: T.surface, border: `1px solid ${form.id ? T.accent : T.border}`, borderRadius: 6, padding: 20, marginBottom: 24 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16 }}>
          <span style={{ color: T.accent, fontWeight: "700", fontSize: "13px" }}>{form.id ? "✎ MODIFICA GIRO" : "+ NUOVO GIRO"}</span>
          {form.id && <button style={{ ...css.btn(), fontSize: "11px", padding: "3px 10px" }} onClick={resetForm}>Annulla modifica</button>}
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "80px 80px 1fr 1fr 1fr", gap: 12, marginBottom: 14 }}>
          {campo("Anno", <input type="number" style={{ ...css.input, width: "100%", boxSizing: "border-box" }} value={form.anno ?? ""} onChange={e => setForm(f => ({ ...f, anno: e.target.value === "" ? null : Number(e.target.value) }))} />)}
          {campo("Giro", <input type="number" style={{ ...css.input, width: "100%", boxSizing: "border-box" }} value={form.giro} onChange={e => setForm(f => ({ ...f, giro: e.target.value }))} placeholder="es. 1" />)}
          {campo("Consegna materiali", <input type="date" style={{ ...css.input, width: "100%", boxSizing: "border-box" }} value={form.consegna_materiali} onChange={e => setForm(f => ({ ...f, consegna_materiali: e.target.value }))} />)}
          {campo("Inizio giro", <input type="date" style={{ ...css.input, width: "100%", boxSizing: "border-box" }} value={form.inizio_giro} onChange={e => setForm(f => ({ ...f, inizio_giro: e.target.value }))} />)}
          {campo("Fine giro", <input type="date" style={{ ...css.input, width: "100%", boxSizing: "border-box" }} value={form.fine_giro} onChange={e => setForm(f => ({ ...f, fine_giro: e.target.value }))} />)}
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr", gap: 12, marginBottom: 14 }}>
          {campo("Riunioni", <input style={{ ...css.input, width: "100%", boxSizing: "border-box" }} value={form.riunioni} onChange={e => setForm(f => ({ ...f, riunioni: e.target.value }))} placeholder="es. 29-30 settembre + 1-3 ottobre" />)}
          {campo("Dati a editori", <input type="date" style={{ ...css.input, width: "100%", boxSizing: "border-box" }} value={form.dati_a_editori} onChange={e => setForm(f => ({ ...f, dati_a_editori: e.target.value }))} />)}
        </div>
        <div style={{ marginBottom: 16 }}>
          <label style={{ color: T.textMid, fontSize: "10px", letterSpacing: "0.08em", textTransform: "uppercase", display: "block", marginBottom: 6 }}>Mese uscita</label>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
            {MESI_LABELS.map(m => (
              <label key={m} style={{ display: "flex", alignItems: "center", gap: 5, cursor: "pointer", padding: "4px 10px", borderRadius: 4, border: `1px solid ${form.mesi.includes(m) ? T.accent : T.border}`, background: form.mesi.includes(m) ? T.accent + "18" : "transparent" }}>
                <input type="checkbox" checked={form.mesi.includes(m)} onChange={() => toggleMese(m)} style={{ accentColor: T.accent }} />
                <span style={{ fontSize: "12px", color: form.mesi.includes(m) ? T.accent : T.textMid }}>{m}</span>
              </label>
            ))}
          </div>
        </div>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
          <button style={css.btn("accent")} onClick={salvaForm} disabled={saving}>{saving ? "Salvataggio..." : form.id ? "💾 Salva modifiche" : "+ Aggiungi Giro"}</button>
        </div>
      </div>

      {/* TABELLA DI CONSULTAZIONE */}
      <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 12 }}>
        <span style={{ color: T.textMid, fontSize: "11px", letterSpacing: "0.08em", textTransform: "uppercase" }}>Anno</span>
        <select style={{ ...css.input, width: 110 }} value={anno} onChange={e => setAnno(Number(e.target.value))}>
          {anniDisponibili.map(y => <option key={y} value={y}>{y}</option>)}
        </select>
        <button style={css.btn()} onClick={() => setAnno(a => a - 1)}>← {anno - 1}</button>
        <button style={css.btn()} onClick={() => setAnno(a => a + 1)}>{anno + 1} →</button>
        <div style={{ flex: 1 }} />
        <span style={{ color: T.textDim, fontSize: "11px" }}>{righeAnno.length} giri</span>
      </div>

      <table style={css.table}>
        <thead>
          <tr>
            <th style={{ ...css.th, width: 60 }}>Giro</th>
            <th style={css.th}>Mese uscita</th>
            <th style={css.th}>Consegna materiali</th>
            <th style={css.th}>Riunioni</th>
            <th style={css.th}>Inizio giro</th>
            <th style={css.th}>Fine giro</th>
            <th style={css.th}>Dati a editori</th>
            <th style={{ ...css.th, width: 60 }}></th>
          </tr>
        </thead>
        <tbody>
          {righeAnno.map(riga => (
            <tr key={riga.id}>
              <td style={{ ...css.td, fontWeight: "600", color: T.accent }}>{riga.giro ?? "—"}</td>
              <td style={css.td}>{riga.mese_uscita || "—"}</td>
              <td style={css.td}>{fmtDataIt(riga.consegna_materiali)}</td>
              <td style={css.td}>{riga.riunioni || "—"}</td>
              <td style={css.td}>{fmtDataIt(riga.inizio_giro)}</td>
              <td style={css.td}>{fmtDataIt(riga.fine_giro)}</td>
              <td style={css.td}>{fmtDataIt(riga.dati_a_editori)}</td>
              <td style={{ ...css.td, textAlign: "center", whiteSpace: "nowrap" }}>
                <button title="Modifica riga" style={{ ...css.btn(), padding: "2px 6px", fontSize: "11px", marginRight: 4 }} onClick={() => startEdit(riga)}>✎</button>
                <button title="Elimina riga" style={{ ...css.btn(), padding: "2px 6px", fontSize: "11px", color: T.red, borderColor: T.red }} onClick={() => eliminaRiga(riga.id)}>✕</button>
              </td>
            </tr>
          ))}
          {righeAnno.length === 0 && (
            <tr><td colSpan={8} style={{ ...css.td, textAlign: "center", color: T.textMid, padding: 30 }}>Nessun giro per il {anno}. Usa il form sopra per aggiungerne uno.</td></tr>
          )}
        </tbody>
      </table>

      {toast && (
        <div style={{ position: "fixed", bottom: 24, left: "50%", transform: "translateX(-50%)", background: toast.type === "err" ? "#4a1a2a" : "#1a3a2a", border: `1px solid ${toast.type === "err" ? T.red : T.green}`, color: toast.type === "err" ? T.red : T.green, borderRadius: 6, padding: "8px 20px", fontSize: "12px", zIndex: 999, boxShadow: "0 4px 20px #0008" }}>
          {toast.msg}
        </div>
      )}
    </div>
  );
}

// Dropdown con ricerca testuale — selezione singola
function SearchableSelect({ value, onChange, options, placeholder = "Tutti", labelKey = null, valueKey = null, width = 180 }) {
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState("");
  const filtered = options.filter(o => {
    const label = labelKey ? o[labelKey] : o;
    return String(label).toLowerCase().includes(search.toLowerCase());
  });
  const currentLabel = value === "tutti" || value === null
    ? placeholder
    : (labelKey ? (options.find(o => o[valueKey] === value)?.[labelKey] || value) : value);
  return (
    <div style={{ position: "relative" }}>
      <button style={{ ...css.btn(value !== "tutti" && value !== null ? "accent" : "default"), minWidth: width, display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }} onClick={() => setOpen(o => !o)}>
        <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: width - 30 }}>{currentLabel}</span>
        <span>▾</span>
      </button>
      {open && (
        <div style={{ position: "absolute", top: "100%", left: 0, zIndex: 50, background: T.surface, border: `1px solid ${T.borderHi}`, borderRadius: 4, minWidth: width, maxHeight: 300, display: "flex", flexDirection: "column", marginTop: 4, boxShadow: "0 4px 20px #0008" }}>
          <div style={{ padding: "8px 10px", borderBottom: `1px solid ${T.border}`, flexShrink: 0 }}>
            <input style={{ ...css.input, width: "100%", boxSizing: "border-box" }} placeholder="Cerca..." value={search} onChange={e => setSearch(e.target.value)} autoFocus />
          </div>
          <div style={{ overflowY: "auto", flex: 1 }}>
            <div style={{ padding: "7px 12px", cursor: "pointer", fontSize: "12px", color: value === "tutti" || value === null ? T.accent : T.textMid, borderBottom: `1px solid ${T.border}22` }}
              onClick={() => { onChange("tutti"); setOpen(false); setSearch(""); }}>{placeholder}</div>
            {filtered.map((o, i) => {
              const val = valueKey ? o[valueKey] : o;
              const label = labelKey ? o[labelKey] : o;
              const selected = value === val;
              return (
                <div key={i} style={{ padding: "7px 12px", cursor: "pointer", fontSize: "12px", color: selected ? T.accent : T.text, background: selected ? T.accent + "18" : "transparent", borderBottom: `1px solid ${T.border}22` }}
                  onClick={() => { onChange(val); setOpen(false); setSearch(""); }}>{label}</div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}

// Dropdown con ricerca testuale — selezione multipla
function SearchableMultiSelect({ values, onChange, options, placeholder = "Tutti", width = 180, renderOption }) {
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState("");
  const label = (o) => renderOption ? renderOption(o) : o;
  const filtered = options.filter(o => label(o).toLowerCase().includes(search.toLowerCase()));
  return (
    <div style={{ position: "relative" }}>
      <button style={{ ...css.btn(values.length > 0 ? "accent" : "default"), minWidth: width, display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }} onClick={() => setOpen(o => !o)}>
        <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: width - 30 }}>
          {values.length === 0 ? placeholder : renderOption ? values.map(renderOption).join(", ") : `${values.length} selezionati`}
        </span>
        <span>▾</span>
      </button>
      {open && (
        <div style={{ position: "absolute", top: "100%", left: 0, zIndex: 50, background: T.surface, border: `1px solid ${T.borderHi}`, borderRadius: 4, minWidth: width, maxHeight: 320, display: "flex", flexDirection: "column", marginTop: 4, boxShadow: "0 4px 20px #0008" }}>
          <div style={{ padding: "8px 10px", borderBottom: `1px solid ${T.border}`, flexShrink: 0 }}>
            <input style={{ ...css.input, width: "100%", boxSizing: "border-box", marginBottom: 6 }} placeholder="Cerca..." value={search} onChange={e => setSearch(e.target.value)} autoFocus />
            <div style={{ display: "flex", justifyContent: "space-between" }}>
              <span style={{ color: T.textMid, fontSize: "11px" }}>{filtered.length} voci</span>
              <span style={{ color: T.accent, fontSize: "11px", cursor: "pointer" }} onClick={() => onChange([])}>Deseleziona tutti</span>
            </div>
          </div>
          <div style={{ overflowY: "auto", flex: 1 }}>
            {filtered.map((o, i) => (
              <label key={i} style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 12px", cursor: "pointer", borderBottom: `1px solid ${T.border}22`, background: values.includes(o) ? T.accent + "18" : "transparent" }}>
                <input type="checkbox" checked={values.includes(o)} onChange={() => onChange(values.includes(o) ? values.filter(x => x !== o) : [...values, o])} style={{ accentColor: T.accent }} />
                <span style={{ fontSize: "12px", color: values.includes(o) ? T.accent : T.text }}>{label(o)}</span>
              </label>
            ))}
          </div>
          <div style={{ padding: 8, borderTop: `1px solid ${T.border}`, flexShrink: 0 }}>
            <button style={{ ...css.btn("accent"), width: "100%" }} onClick={() => { setOpen(false); setSearch(""); }}>Applica</button>
          </div>
        </div>
      )}
    </div>
  );
}

function ModuloCedola({ titoli, giriList, onUpdateTitolo, onDeleteTitolo, spalmatura, prenotato, ruolo, token, onTitoliChange, userAccount }) {
  const [giroLabelSel, setGiroLabelSel] = useState([]);
  const [extraSel, setExtraSel] = useState([]);
  const [giroSel, setGiroSel] = useState([]);
  const [search, setSearch] = useState("");
  const [filterFlag, setFilterFlag] = useState("tutti");
  const [filterEditori, setFilterEditori] = useState([]);
  const [filterAccount, setFilterAccount] = useState([]);
  useEffect(() => { if (ruolo === 'agente' && userAccount) { setFilterAccount(prev => prev.length > 0 ? prev : [userAccount]); } }, [ruolo, userAccount]);
  const [editingId, setEditingId] = useState(null);
  const [sortKey, setSortKey] = useState("n_cedola");
  const [showNuovoGiro, setShowNuovoGiro] = useState(false);
  const [showNuovoTitolo, setShowNuovoTitolo] = useState(false);
  const [showCaricoSemplice, setShowCaricoSemplice] = useState(false);
  const [toastCedola, setToastCedola] = useState(null);
  const showToastCedola = (msg, type = "ok") => { setToastCedola({ msg, type }); setTimeout(() => setToastCedola(null), 3000); };

  // Edit inline obiettivo assegnato (matitina in tabella, senza apertura modale completa)
  const [editingObjId, setEditingObjId] = useState(null);
  const [objInputValue, setObjInputValue] = useState("");
  const [savingObjId, setSavingObjId] = useState(null);
  const [savingTop100Id, setSavingTop100Id] = useState(null);

  const toggleTop100 = async (t) => {
    setSavingTop100Id(t.id);
    const nuovoValore = !t.top_100;
    const ok = await sbUpdateTitolo(t.id, { top_100: nuovoValore }, token);
    setSavingTop100Id(null);
    if (!ok) { showToastCedola("Errore nell'aggiornamento Top 100", "err"); return; }
    onUpdateTitolo({ ...t, top_100: nuovoValore });
    showToastCedola(nuovoValore ? "Aggiunto a Top 100" : "Rimosso da Top 100");
  };

  const startEditObj = (t) => { setEditingObjId(t.id); setObjInputValue(t.obiettivo_assegnato ?? ""); };
  const cancelEditObj = () => { setEditingObjId(null); setObjInputValue(""); };

  const saveEditObj = async (t) => {
    const nuovoValore = objInputValue === "" ? null : parseInt(objInputValue, 10);
    if (objInputValue !== "" && (isNaN(nuovoValore) || nuovoValore < 0)) {
      showToastCedola("Valore obiettivo non valido", "err");
      return;
    }
    if (nuovoValore === (t.obiettivo_assegnato ?? null)) { cancelEditObj(); return; }
    setSavingObjId(t.id);
    const ok = await sbUpdateTitolo(t.id, { obiettivo_assegnato: nuovoValore }, token);
    setSavingObjId(null);
    if (!ok) { showToastCedola("Errore nel salvataggio dell'obiettivo", "err"); return; }
    onUpdateTitolo({ ...t, obiettivo_assegnato: nuovoValore });
    showToastCedola("Obiettivo aggiornato");
    setEditingObjId(null);
    setObjInputValue("");
  };


  // Form nuovo giro
  const emptyGiro = { numero: "", anno: new Date().getFullYear(), descrizione: "" };
  const [formGiro, setFormGiro] = useState(emptyGiro);
  const [savingGiro, setSavingGiro] = useState(false);

  const saveNuovoGiro = async () => {
    if (!formGiro.numero || !formGiro.anno) { showToastCedola("Numero e anno obbligatori", "err"); return; }
    setSavingGiro(true);
    const r = await fetch(`${SUPABASE_URL}/rest/v1/giri`, {
      method: "POST",
      headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`, "Content-Type": "application/json", "Prefer": "return=minimal" },
      body: JSON.stringify([{ numero: parseInt(formGiro.numero), anno: parseInt(formGiro.anno), descrizione: formGiro.descrizione || null }]),
    });
    if (r.ok) {
      showToastCedola(`Giro ${formGiro.numero} ${formGiro.anno} creato`);
      setShowNuovoGiro(false);
      setFormGiro(emptyGiro);
      if (onTitoliChange) onTitoliChange();
    } else {
      const err = await r.text();
      showToastCedola(`Errore: ${err}`, "err");
    }
    setSavingGiro(false);
  };

  // Form nuovo titolo manuale
  const emptyTitolo = { ean: "", titolo: "", autore: "", editore_nome: "", codice_editore: "", prezzo: "", uscita: "", formato: "Cover", n_cedola: "", giro_label: "", posizione: "", obiettivo_assegnato: "", account_editore: "", note_comunicazione: "", note: "" };
  const [formTitolo, setFormTitolo] = useState(emptyTitolo);
  const [savingTitolo, setSavingTitolo] = useState(false);
  const giriLabelAll = useMemo(() => [...new Set(titoli.map(t => t.giro_label).filter(Boolean))].sort((a, b) => { const [na, ya] = a.split(" "); const [nb, yb] = b.split(" "); return Number(yb) - Number(ya) || Number(nb) - Number(na); }), [titoli]);

  const saveNuovoTitolo = async () => {
    const mancanti = [];
    if (!formTitolo.ean?.trim()) mancanti.push("EAN");
    if (!formTitolo.titolo?.trim()) mancanti.push("Titolo");
    if (!formTitolo.n_cedola?.trim()) mancanti.push("N° Cedola");
    if (!formTitolo.giro_label) mancanti.push("Giro");
    if (mancanti.length > 0) {
      showToastCedola(`Campo obbligatorio mancante: ${mancanti.join(", ")}`, "err"); return;
    }
    setSavingTitolo(true);

    // Eredita il ranking_editore dai titoli già esistenti dello stesso editore.
    // ATTENZIONE: lo stesso codice_editore può essere condiviso da imprint diversi con ranking diversi
    // (es. "A94" = sia ADELPHI che ADELPHI RAGAZZI). Quindi il nome editore è più affidabile del codice
    // per identificare l'imprint corretto: 1) codice+nome insieme, 2) nome da solo, 3) codice da solo.
    let candidati = titoli.filter(t => t.ranking_editore != null && formTitolo.codice_editore && formTitolo.editore_nome && t.codice_editore === formTitolo.codice_editore && t.editore_nome === formTitolo.editore_nome);
    if (candidati.length === 0 && formTitolo.editore_nome) {
      candidati = titoli.filter(t => t.ranking_editore != null && t.editore_nome === formTitolo.editore_nome);
    }
    if (candidati.length === 0 && formTitolo.codice_editore) {
      candidati = titoli.filter(t => t.ranking_editore != null && t.codice_editore === formTitolo.codice_editore);
    }
    let rankingEditoreEredita = null;
    if (candidati.length > 0) {
      const conteggio = {};
      candidati.forEach(t => { conteggio[t.ranking_editore] = (conteggio[t.ranking_editore] || 0) + 1; });
      const [valorePiuFrequente] = Object.entries(conteggio).sort((a, b) => b[1] - a[1] || Number(a[0]) - Number(b[0]))[0];
      rankingEditoreEredita = Number(valorePiuFrequente);
    }

    // Sequenza per editore: i titoli dello stesso editore nella stessa cedola condividono la numerazione "posizione" (es. 1..15 Adelphi, 1..20 Harper).
    const gruppoCedola = titoli.filter(t => t.giro_label === formTitolo.giro_label && t.n_cedola === formTitolo.n_cedola && t.editore_nome === formTitolo.editore_nome);
    const maxPos = gruppoCedola.reduce((m, t) => Math.max(m, t.posizione || 0), 0);
    const posRichiesta = formTitolo.posizione ? parseInt(formTitolo.posizione, 10) : null;
    // Se non indicata o fuori range, il titolo va in fondo alla cedola
    const posizioneFinale = (posRichiesta && posRichiesta >= 1 && posRichiesta <= maxPos + 1) ? posRichiesta : maxPos + 1;
    if (posizioneFinale <= maxPos) {
      // Spinge giù di un posto tutti i titoli della cedola da quella posizione in poi
      const daSpostare = gruppoCedola.filter(t => (t.posizione || 0) >= posizioneFinale);
      await Promise.all(daSpostare.map(async t => {
        const nuovaPos = (t.posizione || 0) + 1;
        const ok = await sbUpdateTitolo(t.id, { posizione: nuovaPos }, token);
        if (ok) onUpdateTitolo({ ...t, posizione: nuovaPos });
      }));
    }

    const payload = {
      ean: formTitolo.ean.trim(),
      titolo: formTitolo.titolo,
      autore: formTitolo.autore || null,
      editore_nome: formTitolo.editore_nome || null,
      codice_editore: formTitolo.codice_editore || null,
      ranking_editore: rankingEditoreEredita,
      prezzo: parseFloat(String(formTitolo.prezzo).replace(",", ".")) || null,
      uscita: formTitolo.uscita || null,
      formato: formTitolo.formato || "Cover",
      n_cedola: formTitolo.n_cedola,
      giro_label: formTitolo.giro_label,
      posizione: posizioneFinale,
      obiettivo_assegnato: parseInt(formTitolo.obiettivo_assegnato) || null,
      account_editore: formTitolo.account_editore || null,
      note_comunicazione: formTitolo.note_comunicazione || null,
      note: formTitolo.note || null,
    };
    const r = await fetch(`${SUPABASE_URL}/rest/v1/titoli`, {
      method: "POST",
      headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`, "Content-Type": "application/json", "Prefer": "return=representation" },
      body: JSON.stringify([payload]),
    });
    if (r.ok) {
      const data = await r.json();
      showToastCedola(`"${formTitolo.titolo}" aggiunto in posizione ${posizioneFinale}`);
      setShowNuovoTitolo(false);
      setFormTitolo(emptyTitolo);
      if (onTitoliChange) onTitoliChange();
      if (data?.[0]) onUpdateTitolo(data[0]);
    } else {
      const err = await r.text();
      showToastCedola(`Errore: ${err}`, "err");
    }
    setSavingTitolo(false);
  };

  const anniDispCedola = useMemo(() => {
    const s = new Set();
    titoli.forEach(t => { if (t.giro_label) { const yr = Number(t.giro_label.split(" ")[1]); if (yr >= 2020) s.add(yr); } });
    return [...s].sort((a, b) => b - a);
  }, [titoli]);
  const [filterAnnoCedola, setFilterAnnoCedola] = useState([new Date().getFullYear()]);

  const giriLabel = useMemo(() => [...new Set(titoli.filter(t => t.giro_label !== "EXTRA").map(t => t.giro_label).filter(Boolean))]
    .filter(g => filterAnnoCedola.length === 0 || filterAnnoCedola.includes(Number(g.split(" ")[1])))
    .sort((a, b) => { const [na, ya] = a.split(" "); const [nb, yb] = b.split(" "); return Number(yb || 0) - Number(ya || 0) || Number(nb || 0) - Number(na || 0); }), [titoli, filterAnnoCedola]);

  // Anno dedicato alla colonna Extragiri (indipendente dall'anno dei Giri Vendita)
  const anniDispExtraCedola = useMemo(() => {
    const s = new Set([new Date().getFullYear()]);
    titoli.forEach(t => { if (t.giro_label === "EXTRA" && t.uscita) { const yr = new Date(t.uscita).getFullYear(); if (yr >= 2020 && !isNaN(yr)) s.add(yr); } });
    return [...s].sort((a, b) => b - a);
  }, [titoli]);
  const [filterAnnoExtraCedola, setFilterAnnoExtraCedola] = useState([new Date().getFullYear()]);

  // Cedole extra in ordine alfabetico, filtrate per anno (in base alla data di uscita del titolo).
  // Le cedole senza data di uscita restano sempre visibili, per non perderle.
  const cedoleExtra = useMemo(() => {
    const anniSel = new Set(filterAnnoExtraCedola);
    const cods = new Set();
    titoli.forEach(t => {
      if (t.giro_label !== "EXTRA" || !t.n_cedola) return;
      if (anniSel.size === 0) { cods.add(t.n_cedola); return; }
      const yr = t.uscita ? new Date(t.uscita).getFullYear() : null;
      if (yr === null || isNaN(yr) || anniSel.has(yr)) cods.add(t.n_cedola);
    });
    return [...cods].sort((a, b) => a.localeCompare(b, "it", { sensitivity: "base" }));
  }, [titoli, filterAnnoExtraCedola]);

  const cedole = useMemo(() => { const t = giroLabelSel.length === 0 ? titoli.filter(t => filterAnnoCedola.length === 0 || filterAnnoCedola.includes(Number((t.giro_label||"").split(" ")[1]))) : titoli.filter(t => giroLabelSel.includes(t.giro_label)); return [...new Set(t.map(t => t.n_cedola).filter(Boolean))].sort(); }, [titoli, giroLabelSel, filterAnnoCedola]);
  const accounts = useMemo(() => { const t = giroSel.length === 0 ? (giroLabelSel.length === 0 ? titoli : titoli.filter(t => giroLabelSel.includes(t.giro_label))) : titoli.filter(t => giroSel.includes(t.n_cedola)); return [...new Set(t.map(t => t.account_editore).filter(Boolean))].sort(); }, [titoli, giroLabelSel, giroSel]);
  const editori = useMemo(() => { const t = giroSel.length === 0 ? (giroLabelSel.length === 0 ? titoli : titoli.filter(t => giroLabelSel.includes(t.giro_label))) : titoli.filter(t => giroSel.includes(t.n_cedola)); return [...new Set(t.map(t => t.editore_nome).filter(Boolean))].sort(); }, [titoli, giroLabelSel, giroSel]);

  const resetSelezione = () => { setGiroLabelSel([]); setExtraSel([]); setGiroSel([]); setFilterEditori([]); setFilterAccount([]); setSearch(""); setFilterFlag("tutti"); };

  const filtered = useMemo(() => titoli
    .filter(t => extraSel.length > 0 ? (t.giro_label === "EXTRA" && extraSel.includes(t.n_cedola)) : (giroLabelSel.length === 0 || giroLabelSel.includes(t.giro_label)))
    .filter(t => giroSel.length === 0 || giroSel.includes(t.n_cedola))
    .filter(t => filterEditori.length === 0 || filterEditori.includes(t.editore_nome))
    .filter(t => filterAccount.length === 0 || filterAccount.includes(t.account_editore))
    .filter(t => { if (!search) return true; const q = search.toLowerCase(); return t.titolo?.toLowerCase().includes(q) || t.autore?.toLowerCase().includes(q) || t.editore_nome?.toLowerCase().includes(q) || t.ean?.includes(q); })
    .filter(t => { if (filterFlag === "triangolo") return t.il_triangolo; if (filterFlag === "top100") return t.top_100; if (filterFlag === "gemelli") return t.ean_gemello_1; return true; })
    .sort((a, b) => { if (sortKey === "n_cedola") return (a.n_cedola ?? "").localeCompare(b.n_cedola ?? "") || ((a.ranking_editore ?? 99) - (b.ranking_editore ?? 99)) || (a.editore_nome ?? "").localeCompare(b.editore_nome ?? "") || (a.posizione ?? 0) - (b.posizione ?? 0); if (sortKey === "editore") return ((a.ranking_editore ?? 99) - (b.ranking_editore ?? 99)) || (a.editore_nome ?? "").localeCompare(b.editore_nome ?? "") || (a.posizione ?? 0) - (b.posizione ?? 0); if (sortKey === "prezzo") return (b.prezzo ?? 0) - (a.prezzo ?? 0); return 0; }),
  [titoli, giroLabelSel, extraSel, giroSel, search, filterFlag, filterEditori, filterAccount, sortKey]);

  const editingTitolo = titoli.find(t => t.id === editingId);

  const getObjCanale = (titolo, canale_codice) => {
    const spRow = spalmatura.find(s => s.editore_nome === titolo.editore_nome && s.formato === (titolo.formato || 'Cover') && s.canale_codice === canale_codice);
    if (!spRow || !titolo.obiettivo_assegnato) return 0;
    return Math.round(titolo.obiettivo_assegnato * spRow.percentuale);
  };

  // Costruisce un foglio nello stesso formato del template ufficiale di import:
  // riga 1 = titolo (merge), riga 2 = legenda (merge), riga 3 = intestazioni, dati da riga 4.
  // Replica ESATTAMENTE la struttura del template ufficiale di import (4 righe prima dei dati:
  // titolo, legenda, intestazioni, riga vuota di separazione) — ModuloImport legge i dati da riga 5
  // (data.slice(4)), quindi questa riga vuota è obbligatoria o la prima riga di dati viene scartata.
  const buildTemplateSheet = (XLSX, title, legend, headers, rows) => {
    const blankRow = headers.map(() => "");
    const ws = XLSX.utils.aoa_to_sheet([[title], [legend], headers, blankRow, ...rows]);
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: headers.length - 1 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: headers.length - 1 } },
    ];
    ws["!cols"] = headers.map(h => ({ wch: Math.max(12, Math.min(32, h.length + 4)) }));
    ws["!rows"] = [{ hpt: 22 }, { hpt: 16 }, { hpt: 26 }];
    return ws;
  };

  // Numera editori e titoli in sequenza seguendo l'ordine già presente nell'elenco (rispetta
  // ranking_editore/posizione già impostati quando presenti). Non lascia mai vuoto: se il ranking
  // editore non è valorizzato in anagrafica, lo si deduce dall'ordine di comparsa in questo export.
  const computeRanking = (lista) => {
    const rankEditoreMap = new Map();
    const rankTitoloMap = new Map();
    const contatoriTitolo = {};
    let prossimoRankEditore = 1;
    const editoreToRank = {};
    lista.forEach(t => {
      const chiave = t.editore_nome || t.codice_editore || "—";
      if (!(chiave in editoreToRank)) {
        editoreToRank[chiave] = t.ranking_editore ?? prossimoRankEditore;
        prossimoRankEditore = Math.max(prossimoRankEditore, editoreToRank[chiave] + 1);
      }
      rankEditoreMap.set(t.id, editoreToRank[chiave]);
      contatoriTitolo[chiave] = (contatoriTitolo[chiave] || 0) + 1;
      rankTitoloMap.set(t.id, contatoriTitolo[chiave]);
    });
    return { rankEditoreMap, rankTitoloMap };
  };

  const exportAgenti = () => {
    const XLSX = window.XLSX;
    const { rankEditoreMap, rankTitoloMap } = computeRanking(filtered);
    const headers = ["N° CEDOLA","RANK.EDITORE","RANK.TITOLO","EAN","TITOLO","AUTORE","COD.EDITORE","EDITORE","TOP 100","PREZZO","USCITA","NOTE","EAN GEM 1","TITOLO GEM 1","EAN GEM 2","TITOLO GEM 2","EAN GEM 3","TITOLO GEM 3","OBJ INDIPENDENTI & ALTRE CATENE"];
    const rows = filtered.map(t => [t.n_cedola, rankEditoreMap.get(t.id), rankTitoloMap.get(t.id), t.ean, t.titolo, t.autore, t.codice_editore, t.editore_nome, t.top_100 ? "SI" : "", t.prezzo, t.uscita, t.note_comunicazione || t.note, t.ean_gemello_1, t.titolo_gemello_1, t.ean_gemello_2, t.titolo_gemello_2, t.ean_gemello_3, t.titolo_gemello_3, getObjCanale(t, 'INDIPENDENTI_ALTRE_CATENE')]);
    const ws = buildTemplateSheet(XLSX, "CEDOLA AGENTI — GIRO " + (giroLabelSel.length === 0 ? "TUTTI" : giroLabelSel.join(", ")), "Esportazione dati cedola per la rete agenti", headers, rows);
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "CEDOLA AGENTI");
    XLSX.writeFile(wb, `CEDOLA_AGENTI_${giroLabelSel.length === 0 ? "TUTTI" : giroLabelSel.join("-")}.xlsx`);
  };

  const exportDirezionale = () => {
    const XLSX = window.XLSX;
    const canaliDir = [{ codice: 'FELTRINELLI', label: 'Feltrinelli' },{ codice: 'GIUNTI', label: 'Giunti' },{ codice: 'MONDADORI', label: 'Mondadori' },{ codice: 'UBIK', label: 'Ubik' },{ codice: 'INDIPENDENTI_ALTRE_CATENE', label: 'Indip. & Altre Catene' },{ codice: 'AMAZON', label: 'Amazon' },{ codice: 'IBS', label: 'IBS' },{ codice: 'ALTRI_ONLINE', label: 'Altri Online' },{ codice: 'FASTBOOK', label: 'Fastbook' },{ codice: 'GROSSISTI', label: 'Grossisti' },{ codice: 'CENTROLIBRI', label: 'Centrolibri' },{ codice: 'GDO', label: 'GDO' }];
    const giroLabelStr = giroLabelSel.length === 0 ? "TUTTI" : giroLabelSel.join(", ");
    const { rankEditoreMap, rankTitoloMap } = computeRanking(filtered);
    const headersCedola = ["N° CEDOLA","RANK.EDITORE","RANK.TITOLO","EAN","TITOLO","AUTORE","COD.EDITORE","EDITORE","TOP 100","PREZZO","OBJ TOTALE","NOTE","EAN GEM 1","TITOLO GEM 1","EAN GEM 2","TITOLO GEM 2","EAN GEM 3","TITOLO GEM 3"];
    const rowsCedola = filtered.map(t => [t.n_cedola, rankEditoreMap.get(t.id), rankTitoloMap.get(t.id), t.ean, t.titolo, t.autore, t.codice_editore, t.editore_nome, t.top_100 ? "SI" : "", t.prezzo, t.obiettivo_assegnato || 0, t.note_comunicazione || t.note, t.ean_gemello_1, t.titolo_gemello_1, t.ean_gemello_2, t.titolo_gemello_2, t.ean_gemello_3, t.titolo_gemello_3]);
    const wsCedola = buildTemplateSheet(XLSX, "CEDOLA DIREZIONALE — GIRO " + giroLabelStr, "Esportazione dati cedola per la direzione", headersCedola, rowsCedola);
    const headersObj = ["N° CEDOLA","EAN","TITOLO","AUTORE","COD.EDITORE","EDITORE","PREZZO","OBJ TOTALE",...canaliDir.map(c => c.label)];
    const rowsObj = filtered.map(t => [t.n_cedola, t.ean, t.titolo, t.autore, t.codice_editore, t.editore_nome, t.prezzo, t.obiettivo_assegnato || 0, ...canaliDir.map(c => getObjCanale(t, c.codice))]);
    const wsObj = buildTemplateSheet(XLSX, "OBIETTIVI PER CANALE — GIRO " + giroLabelStr, "Ripartizione obiettivi per canale di vendita", headersObj, rowsObj);
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, wsCedola, "CEDOLA"); XLSX.utils.book_append_sheet(wb, wsObj, "OBIETTIVI");
    XLSX.writeFile(wb, `CEDOLA_DIREZIONALE_${giroLabelSel.length === 0 ? "TUTTI" : giroLabelSel.join("-")}.xlsx`);
  };

  // Export nello stesso identico formato del template di import (foglio "CEDOLA", colonne nell'ordine
  // atteso da ModuloImport/COL_MAP), così il file scaricato può essere modificato e ricaricato direttamente.
  const exportTemplateRicaricabile = () => {
    const XLSX = window.XLSX;
    const headers = ["GIRO","EAN","TITOLO","AUTORE","EDITORE","PREZZO","USCITA","OBIETTIVO","TOP 100","PROMOZIONE","NOTE","EAN GEMELLO 1","TITOLO GEMELLO 1","EAN GEMELLO 2","TITOLO GEMELLO 2","EAN GEMELLO 3","TITOLO GEMELLO 3"];
    const rows = filtered.map(t => [
      t.giro_label === "EXTRA" ? "" : t.giro_label,
      t.ean, t.titolo, t.autore, t.editore_nome, t.prezzo, t.uscita,
      t.obiettivo_assegnato || "", t.top_100 ? "SI" : "", t.promozione || "", t.note_comunicazione || t.note || "",
      t.ean_gemello_1, t.titolo_gemello_1, t.ean_gemello_2, t.titolo_gemello_2, t.ean_gemello_3, t.titolo_gemello_3,
    ]);
    const ws = buildTemplateSheet(XLSX, "TEMPLATE CEDOLA GIRO — NON MODIFICARE LE INTESTAZIONI", "Colonne obbligatorie: GIRO, EAN, TITOLO, EDITORE, PREZZO • file ricaricabile da \"Carica con Template\"", headers, rows);
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "CEDOLA");
    XLSX.writeFile(wb, `TEMPLATE_CEDOLA_${giroLabelSel.length === 0 ? "TUTTI" : giroLabelSel.join("-")}.xlsx`);
  };

  if (giroLabelSel.length === 0 && extraSel.length === 0) {
    return (
      <div style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "flex-start", gap: 28, padding: "48px 20px", overflowY: "auto" }}>
        <div style={{ color: T.textMid, fontSize: "14px" }}>Seleziona un giro o una cedola extra</div>
        <div style={{ display: "flex", gap: 48, alignItems: "flex-start", justifyContent: "center", flexWrap: "wrap" }}>

          {/* Colonna GIRI VENDITA */}
          <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 14, width: 240 }}>
            <div style={{ color: T.text, fontSize: "12px", fontWeight: 700, letterSpacing: "0.5px", textTransform: "uppercase" }}>Giri Vendita</div>
            <SearchableMultiSelect values={filterAnnoCedola.map(String)} onChange={v => setFilterAnnoCedola(v.map(Number))} options={anniDispCedola.map(String)} renderOption={v => v} placeholder="Anno" width={140} />
            <div style={{ display: "flex", flexDirection: "column", gap: 8, width: "100%" }}>
              {giriLabel.length === 0 && <div style={{ color: T.textMid, fontSize: "12px", textAlign: "center", padding: "10px 0" }}>Nessun giro per l'anno selezionato</div>}
              {giriLabel.map(g => (
                <button key={g} style={{ ...css.btn("accent"), padding: "10px 16px", fontSize: "13px", width: "100%" }} onClick={() => setGiroLabelSel([g])}>Giro {g}</button>
              ))}
            </div>
          </div>

          {/* Colonna EXTRAGIRI */}
          <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 14, width: 240 }}>
            <div style={{ color: T.text, fontSize: "12px", fontWeight: 700, letterSpacing: "0.5px", textTransform: "uppercase" }}>Extragiri</div>
            <SearchableMultiSelect values={filterAnnoExtraCedola.map(String)} onChange={v => setFilterAnnoExtraCedola(v.map(Number))} options={anniDispExtraCedola.map(String)} renderOption={v => v} placeholder="Anno" width={140} />
            <div style={{ display: "flex", flexDirection: "column", gap: 8, width: "100%" }}>
              {cedoleExtra.length === 0 && <div style={{ color: T.textMid, fontSize: "12px", textAlign: "center", padding: "10px 0" }}>Nessuna cedola extra per l'anno selezionato</div>}
              {cedoleExtra.map(c => (
                <button key={c} style={{ ...css.btn(), padding: "10px 16px", fontSize: "13px", borderColor: T.accent, color: T.accent, width: "100%" }} onClick={() => setExtraSel([c])}>{c}</button>
              ))}
            </div>
          </div>

        </div>
      </div>
    );
  }


  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%", overflow: "hidden" }}>
      {editingTitolo && (
        <EditModal
          titolo={editingTitolo}
          siblings={titoli}
          onSave={onUpdateTitolo}
          onDelete={id => { onDeleteTitolo && onDeleteTitolo(id); showToastCedola("Riga cancellata"); if (onTitoliChange) onTitoliChange(); }}
          onClose={() => setEditingId(null)}
          token={token}
        />
      )}

      {/* MODAL NUOVO GIRO */}
      {showNuovoGiro && (
        <div style={{ position: "fixed", inset: 0, background: "#000a", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center" }} onClick={() => setShowNuovoGiro(false)}>
          <div style={{ background: T.surface, border: `1px solid ${T.borderHi}`, borderRadius: 6, padding: 28, width: 380 }} onClick={e => e.stopPropagation()}>
            <div style={{ color: T.accent, fontWeight: "700", fontSize: "13px", marginBottom: 20 }}>NUOVO GIRO</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 12 }}>
              <div>
                <label style={{ color: T.textMid, fontSize: "10px", textTransform: "uppercase", letterSpacing: "0.08em", display: "block", marginBottom: 4 }}>Numero *</label>
                <input style={{ ...css.input, width: "100%", boxSizing: "border-box" }} type="number" value={formGiro.numero} onChange={e => setFormGiro(f => ({ ...f, numero: e.target.value }))} placeholder="es. 1" />
              </div>
              <div>
                <label style={{ color: T.textMid, fontSize: "10px", textTransform: "uppercase", letterSpacing: "0.08em", display: "block", marginBottom: 4 }}>Anno *</label>
                <input style={{ ...css.input, width: "100%", boxSizing: "border-box" }} type="number" value={formGiro.anno} onChange={e => setFormGiro(f => ({ ...f, anno: e.target.value }))} />
              </div>
            </div>
            <div style={{ marginBottom: 20 }}>
              <label style={{ color: T.textMid, fontSize: "10px", textTransform: "uppercase", letterSpacing: "0.08em", display: "block", marginBottom: 4 }}>Descrizione</label>
              <input style={{ ...css.input, width: "100%", boxSizing: "border-box" }} value={formGiro.descrizione} onChange={e => setFormGiro(f => ({ ...f, descrizione: e.target.value }))} placeholder="opzionale" />
            </div>
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
              <button style={css.btn()} onClick={() => setShowNuovoGiro(false)}>Annulla</button>
              <button style={css.btn("accent")} onClick={saveNuovoGiro} disabled={savingGiro}>{savingGiro ? "Salvataggio..." : "Crea Giro"}</button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL NUOVO TITOLO */}
      {showNuovoTitolo && (
        <div style={{ position: "fixed", inset: 0, background: "#000a", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center" }} onClick={() => setShowNuovoTitolo(false)}>
          <div style={{ background: T.surface, border: `1px solid ${T.borderHi}`, borderRadius: 6, padding: 28, width: 640, maxHeight: "85vh", overflowY: "auto" }} onClick={e => e.stopPropagation()}>
            <div style={{ color: T.accent, fontWeight: "700", fontSize: "13px", marginBottom: 20 }}>NUOVO TITOLO</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
              {[
                ["ean", "EAN *", "full"], ["titolo", "Titolo *", "full"],
                ["autore", "Autore"], ["editore_nome", "Editore *"],
                ["codice_editore", "Cod. Editore"], ["prezzo", "Prezzo"],
                ["uscita", "Uscita (es. 15/04/2026)"], ["formato", "Formato"],
                ["account_editore", "Account editore"], ["obiettivo_assegnato", "Obiettivo assegnato"],
              ].map(([k, label, span]) => (
                <div key={k} style={span === "full" ? { gridColumn: "1/-1" } : {}}>
                  <label style={{ color: T.textMid, fontSize: "10px", textTransform: "uppercase", letterSpacing: "0.08em", display: "block", marginBottom: 4 }}>{label}</label>
                  <input style={{ ...css.input, width: "100%", boxSizing: "border-box" }} value={formTitolo[k]} onChange={e => setFormTitolo(f => ({ ...f, [k]: e.target.value }))} />
                </div>
              ))}
              <div>
                <label style={{ color: T.textMid, fontSize: "10px", textTransform: "uppercase", letterSpacing: "0.08em", display: "block", marginBottom: 4 }}>Giro *</label>
                <select style={{ ...css.input, width: "100%", boxSizing: "border-box" }} value={formTitolo.giro_label} onChange={e => setFormTitolo(f => ({ ...f, giro_label: e.target.value }))}>
                  <option value="">— Seleziona —</option>
                  {giriLabelAll.map(g => <option key={g} value={g}>Giro {g}</option>)}
                </select>
              </div>
              <div>
                <label style={{ color: T.textMid, fontSize: "10px", textTransform: "uppercase", letterSpacing: "0.08em", display: "block", marginBottom: 4 }}>N° Cedola *</label>
                <input style={{ ...css.input, width: "100%", boxSizing: "border-box" }} value={formTitolo.n_cedola} onChange={e => setFormTitolo(f => ({ ...f, n_cedola: e.target.value }))} placeholder="es. 1A 2026" />
              </div>
              <div>
                <label style={{ color: T.textMid, fontSize: "10px", textTransform: "uppercase", letterSpacing: "0.08em", display: "block", marginBottom: 4 }}>Posizione per editore</label>
                <input style={{ ...css.input, width: "100%", boxSizing: "border-box" }} type="number" min="1" value={formTitolo.posizione} onChange={e => setFormTitolo(f => ({ ...f, posizione: e.target.value }))} placeholder={(() => {
                  const n = titoli.filter(t => t.giro_label === formTitolo.giro_label && t.n_cedola === formTitolo.n_cedola && t.editore_nome === formTitolo.editore_nome).length;
                  return n > 0 ? `vuoto = in fondo (pos. ${n + 1})` : "vuoto = pos. 1";
                })()} />
              </div>
            </div>
            <div style={{ marginTop: 12 }}>
              <label style={{ color: T.textMid, fontSize: "10px", textTransform: "uppercase", letterSpacing: "0.08em", display: "block", marginBottom: 4 }}>Note comunicazione</label>
              <textarea style={{ ...css.input, width: "100%", boxSizing: "border-box", height: 50, resize: "vertical" }} value={formTitolo.note_comunicazione} onChange={e => setFormTitolo(f => ({ ...f, note_comunicazione: e.target.value }))} />
            </div>
            <div style={{ marginTop: 10 }}>
              <label style={{ color: T.textMid, fontSize: "10px", textTransform: "uppercase", letterSpacing: "0.08em", display: "block", marginBottom: 4 }}>Note interne</label>
              <textarea style={{ ...css.input, width: "100%", boxSizing: "border-box", height: 50, resize: "vertical" }} value={formTitolo.note} onChange={e => setFormTitolo(f => ({ ...f, note: e.target.value }))} />
            </div>
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 20 }}>
              <button style={css.btn()} onClick={() => setShowNuovoTitolo(false)}>Annulla</button>
              <button style={css.btn("accent")} onClick={saveNuovoTitolo} disabled={savingTitolo}>{savingTitolo ? "Salvataggio..." : "Aggiungi Titolo"}</button>
            </div>
          </div>
        </div>
      )}

      {/* MODALE CARICO SEMPLICE */}
      {showCaricoSemplice && (
        <ModuloCaricoSemplice
          giriList={giriList}
          titoliEsistenti={titoli}
          token={token}
          onClose={() => setShowCaricoSemplice(false)}
          onImportDone={() => { onTitoliChange && onTitoliChange(); showToastCedola("Carico completato"); }}
        />
      )}

      {/* TOAST LOCALE */}
      {toastCedola && (
        <div style={{ position: "fixed", bottom: 24, left: "50%", transform: "translateX(-50%)", background: toastCedola.type === "err" ? "#4a1a2a" : "#1a3a2a", border: `1px solid ${toastCedola.type === "err" ? T.red : T.green}`, color: toastCedola.type === "err" ? T.red : T.green, borderRadius: 6, padding: "8px 20px", fontSize: "12px", zIndex: 999, boxShadow: "0 4px 20px #0008" }}>
          {toastCedola.msg}
        </div>
      )}
      <div style={{ padding: "12px 20px", borderBottom: `1px solid ${T.border}`, display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
        <SearchableMultiSelect values={filterAnnoCedola.map(String)} onChange={v => { setFilterAnnoCedola(v.map(Number)); setGiroLabelSel([]); setExtraSel([]); setGiroSel([]); setFilterEditori([]); }} options={anniDispCedola.map(String)} renderOption={v => v} placeholder="Anno" width={110} />
        <SearchableMultiSelect values={giroLabelSel} onChange={v => { setGiroLabelSel(v); setExtraSel([]); setGiroSel([]); setFilterEditori([]); }} options={giriLabel} renderOption={g => `Giro ${g}`} placeholder="Tutti i giri" width={170} />
        {cedoleExtra.length > 0 && (
          <SearchableMultiSelect values={extraSel} onChange={v => { setExtraSel(v); setGiroLabelSel([]); setGiroSel([]); setFilterEditori([]); }} options={cedoleExtra} placeholder="Cedole Extra" width={170} />
        )}
        <SearchableMultiSelect values={giroSel} onChange={v => { setGiroSel(v); setFilterEditori([]); }} options={cedole} placeholder="Tutte le cedole" width={160} />
        <SearchableMultiSelect values={filterEditori} onChange={setFilterEditori} options={editori} placeholder="Tutti gli editori" width={190} />
        <SearchableMultiSelect values={filterAccount} onChange={setFilterAccount} options={accounts} placeholder="Tutti gli account" width={160} />
        <input style={{ ...css.input, width: 180 }} placeholder="Cerca..." value={search} onChange={e => setSearch(e.target.value)} />
        {["tutti","triangolo","top100","gemelli"].map(f => (
          <button key={f} style={{ ...css.btn(filterFlag === f ? "accent" : "default"), padding: "5px 10px" }} onClick={() => setFilterFlag(f)}>
            {f === "tutti" ? "Tutti" : f === "triangolo" ? "▲" : f === "top100" ? "★" : "Gem."}
          </button>
        ))}
        <button style={css.btn()} onClick={resetSelezione}>↺ Cambia giro</button>
        <div style={{ marginLeft: "auto", color: T.textMid, fontSize: "11px" }}><span style={{ color: T.text }}>{filtered.length}</span> titoli</div>
        {ruolo !== "agente" && <>
          <button style={{ ...css.btn(), borderColor: T.green, color: T.green }} onClick={() => setShowNuovoGiro(true)}>+ Giro</button>
          <button style={{ ...css.btn(), borderColor: T.green, color: T.green }} onClick={() => setShowNuovoTitolo(true)}>+ Titolo</button>
          <button style={{ ...css.btn(), borderColor: T.blue, color: T.blue }} onClick={() => setShowCaricoSemplice(true)} title="Carica più titoli da file Excel — supporta anche cedole EXTRA">+ Carico</button>
        </>}
        <button style={css.btn("accent")} onClick={exportAgenti}>↓ Download Agenti</button>
        {ruolo !== "agente" && <button style={css.btn("accent")} onClick={exportDirezionale}>↓ Download Direzionali</button>}
        {ruolo !== "agente" && <button style={{ ...css.btn(), borderColor: T.blue, color: T.blue }} onClick={exportTemplateRicaricabile} title="Genera un file nello stesso formato del template di import, ricaricabile direttamente">↓ Template ricaricabile</button>}
        <a href="https://lafeltrinelli.sharepoint.com/:f:/s/PDE/IgD7OJj1nZrhTKrDAVfuqOc3AQQaTrH4gMuXZT7Ob6Fbm2w?e=mEDMdm" target="_blank" rel="noopener noreferrer" style={{ ...css.btn(), borderColor: "#9c6fcf", color: "#9c6fcf", textDecoration: "none" }}>📄 PDF & Materiali</a>
      </div>
      <div style={{ flex: 1, overflowY: "auto" }}>
        <table style={css.table}>
          <thead>
            <tr>
              <th style={css.th} onClick={() => setSortKey("n_cedola")}>Cedola{sortKey === "n_cedola" ? " ↓" : ""}</th>
              <th style={{ ...css.th, textAlign: "center" }}>Pos.</th>
              <th style={css.th}>EAN</th><th style={css.th}>Titolo</th><th style={css.th}>Autore</th><th style={css.th}>Cod.Ed.</th>
              <th style={css.th} onClick={() => setSortKey("editore")}>Editore{sortKey === "editore" ? " ↓" : ""}</th>
              <th style={css.th} onClick={() => setSortKey("prezzo")}>€{sortKey === "prezzo" ? " ↓" : ""}</th>
              <th style={css.th}>Obj</th><th style={css.th}>Top 100</th><th style={css.th}>Gemelli</th><th style={css.th}>Note</th><th style={css.th}></th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((t, i) => (
              <tr key={t.id} style={{ background: i % 2 === 0 ? "transparent" : T.surface + "66" }}>
                <td style={{ ...css.td, color: T.textMid, fontSize: "10px", whiteSpace: "nowrap" }}>{t.n_cedola}</td>
                <td style={{ ...css.td, color: T.textMid, fontSize: "11px", textAlign: "center" }}>{t.posizione ?? "—"}</td>
                <td style={{ ...css.td, color: T.textDim, fontFamily: "monospace", fontSize: "11px" }}>{t.ean}</td>
                <td style={{ ...css.td, maxWidth: 260 }}><div style={{ fontWeight: "600", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{t.titolo}</div></td>
                <td style={{ ...css.td, color: T.textMid }}>{t.autore}</td>
                <td style={{ ...css.td, color: T.textMid, fontSize: "11px" }}>{t.codice_editore}</td>
                <td style={{ ...css.td, color: T.accent, fontWeight: "600", whiteSpace: "nowrap" }}>{t.editore_nome}</td>
                <td style={{ ...css.td, whiteSpace: "nowrap" }}>€ {t.prezzo?.toFixed(2)}</td>
                <td style={{ ...css.td, whiteSpace: "nowrap", minWidth: 130 }}>
                  {(() => {
                    const pren = prenotato.filter(p => p.titolo_id === t.id).reduce((s,p) => s+p.quantita, 0);
                    if (editingObjId === t.id) {
                      return (
                        <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                          <input
                            autoFocus
                            type="number"
                            min="0"
                            style={{ ...css.input, width: 64, padding: "3px 6px", fontSize: "11px" }}
                            value={objInputValue}
                            onChange={e => setObjInputValue(e.target.value)}
                            onKeyDown={e => { if (e.key === "Enter") saveEditObj(t); if (e.key === "Escape") cancelEditObj(); }}
                            disabled={savingObjId === t.id}
                          />
                          <button title="Salva" style={{ ...css.btn(), padding: "2px 6px", fontSize: "11px", color: T.green, borderColor: T.green }} onClick={() => saveEditObj(t)} disabled={savingObjId === t.id}>✓</button>
                          <button title="Annulla" style={{ ...css.btn(), padding: "2px 6px", fontSize: "11px" }} onClick={cancelEditObj} disabled={savingObjId === t.id}>✕</button>
                        </div>
                      );
                    }
                    return (
                      <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                        <div style={{ flex: 1 }}>
                          <div style={{ fontSize: "11px", marginBottom: 4 }}>{pren.toLocaleString("it")} / {(t.obiettivo_assegnato||0).toLocaleString("it")}</div>
                          <ProgressBar value={pren} total={t.obiettivo_assegnato} />
                        </div>
                        {ruolo !== "agente" && (
                          <button title="Modifica obiettivo" style={{ background: "none", border: "none", cursor: "pointer", color: T.textMid, fontSize: "12px", padding: 2 }} onClick={() => startEditObj(t)}>✎</button>
                        )}
                      </div>
                    );
                  })()}
                </td>

                <td style={css.td}>
                  <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    {t.il_triangolo && <Badge label="▲" color={T.purple} />}
                    {ruolo !== "agente" ? (
                      <button
                        title={t.top_100 ? "Rimuovi da Top 100" : "Aggiungi a Top 100"}
                        onClick={() => toggleTop100(t)}
                        disabled={savingTop100Id === t.id}
                        style={{
                          background: t.top_100 ? T.accent + "22" : "transparent",
                          border: `1px solid ${t.top_100 ? T.accent : T.border}`,
                          borderRadius: 4,
                          color: t.top_100 ? T.accent : T.textMid,
                          cursor: savingTop100Id === t.id ? "default" : "pointer",
                          fontSize: "11px",
                          padding: "2px 7px",
                          display: "flex",
                          alignItems: "center",
                          gap: 4,
                          opacity: savingTop100Id === t.id ? 0.6 : 1,
                        }}
                      >
                        {t.top_100 ? "★" : "☆"} {savingTop100Id === t.id ? "..." : ""}
                      </button>
                    ) : (
                      t.top_100 && <Badge label="★" color={T.accent} />
                    )}
                  </div>
                </td>
                <td style={css.td}>{t.ean_gemello_1 && <div style={{ fontSize: "10px", color: T.textMid }}><div>{t.ean_gemello_1}</div>{t.ean_gemello_2 && <div>{t.ean_gemello_2}</div>}{t.ean_gemello_3 && <div>{t.ean_gemello_3}</div>}</div>}</td>
                <td style={{ ...css.td, maxWidth: 160 }}>{(t.note || t.note_comunicazione) && <div style={{ fontSize: "11px", color: T.textMid, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 150 }} title={t.note_comunicazione || t.note}>{t.note_comunicazione || t.note}</div>}</td>
                <td style={css.td}><button style={{ ...css.btn(), padding: "3px 9px", fontSize: "11px" }} onClick={() => setEditingId(t.id)}>Edit</button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function ModuloFineGiro({ titoli, prenotato, canali, token, ruolo, spalmatura, userAccount }) {
  const anniDispFineGiro = useMemo(() => {
    const s = new Set();
    titoli.forEach(t => { if (t.giro_label && t.giro_label !== "EXTRA") { const yr = Number(t.giro_label.split(" ")[1]); if (yr >= 2020) s.add(yr); } });
    return [...s].sort((a, b) => b - a);
  }, [titoli]);
  const [filterAnnoFineGiro, setFilterAnnoFineGiro] = useState([new Date().getFullYear()]);

  const giriLabel = useMemo(() => [...new Set(titoli.filter(t => t.giro_label !== "EXTRA").map(t => t.giro_label).filter(Boolean))]
    .filter(g => filterAnnoFineGiro.length === 0 || filterAnnoFineGiro.includes(Number(g.split(" ")[1])))
    .sort((a, b) => {
    const [na, ya] = a.split(" "); const [nb, yb] = b.split(" ");
    return Number(yb || 0) - Number(ya || 0) || Number(nb || 0) - Number(na || 0);
  }), [titoli]);

  // Anno dedicato alla colonna Extragiri (indipendente dall'anno dei Giri Vendita)
  const anniDispExtraFineGiro = useMemo(() => {
    const s = new Set([new Date().getFullYear()]);
    titoli.forEach(t => { if (t.giro_label === "EXTRA" && t.uscita) { const yr = new Date(t.uscita).getFullYear(); if (yr >= 2020 && !isNaN(yr)) s.add(yr); } });
    return [...s].sort((a, b) => b - a);
  }, [titoli]);
  const [filterAnnoExtraFineGiro, setFilterAnnoExtraFineGiro] = useState([new Date().getFullYear()]);

  // Cedole extra in ordine alfabetico, filtrate per anno (in base alla data di uscita del titolo).
  // Le cedole senza data di uscita restano sempre visibili, per non perderle.
  const cedoleExtra = useMemo(() => {
    const anniSel = new Set(filterAnnoExtraFineGiro);
    const cods = new Set();
    titoli.forEach(t => {
      if (t.giro_label !== "EXTRA" || !t.n_cedola) return;
      if (anniSel.size === 0) { cods.add(t.n_cedola); return; }
      const yr = t.uscita ? new Date(t.uscita).getFullYear() : null;
      if (yr === null || isNaN(yr) || anniSel.has(yr)) cods.add(t.n_cedola);
    });
    return [...cods].sort((a, b) => a.localeCompare(b, "it", { sensitivity: "base" }));
  }, [titoli, filterAnnoExtraFineGiro]);

  const [giroLabelSel, setGiroLabelSel] = useState([]);
  const [extraSel, setExtraSel] = useState([]);
  const [cedolaSel, setCedolaSel] = useState([]);
  const [filterEditori, setFilterEditori] = useState([]);
  const [filterAccount, setFilterAccount] = useState([]);
  useEffect(() => { if (ruolo === 'agente' && userAccount) { setFilterAccount(prev => prev.length > 0 ? prev : [userAccount]); } }, [ruolo, userAccount]);
  const [filterPromozione, setFilterPromozione] = useState([]);
  const [filterCanale, setFilterCanale] = useState([]);
  const [search, setSearch] = useState("");
  const [sortPren, setSortPren] = useState(null); // null | "asc" | "desc"
  const [soloPrenotati, setSoloPrenotati] = useState(false);
  const [clienteSearch, setClienteSearch] = useState("");
  const [clienteSel, setClienteSel] = useState(null);
  const [clientiList, setClientiList] = useState([]);
  const [clienteDropdownOpen, setClienteDropdownOpen] = useState(false);
  const [prenotatoCliente, setPrenotatoCliente] = useState([]);
  const [auroraEdit, setAuroraEdit] = useState({});
  const [auroraEditing, setAuroraEditing] = useState(null);

  const resetFiltri = () => { setGiroLabelSel([]); setExtraSel([]); setCedolaSel([]); setFilterEditori([]); setFilterAccount([]); setFilterPromozione([]); setFilterCanale([]); setSearch(""); setClienteSel(null); setSoloPrenotati(false); };

  useEffect(() => {
    if (!token) return;
    fetch(`${SUPABASE_URL}/rest/v1/prenotato_clienti?select=codice_cliente,nome_cliente&limit=10000`, {
      headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}` }
    }).then(r => r.json()).then(data => {
      if (!Array.isArray(data)) return;
      const map = {};
      data.forEach(r => { map[r.codice_cliente] = r.nome_cliente; });
      setClientiList(Object.entries(map).map(([cod, nome]) => ({ cod, nome })).sort((a, b) => a.nome.localeCompare(b.nome)));
    });
  }, [token]);

  useEffect(() => {
    if (!clienteSel || !token) { setPrenotatoCliente([]); return; }
    fetch(`${SUPABASE_URL}/rest/v1/prenotato_clienti?codice_cliente=eq.${encodeURIComponent(clienteSel)}&select=titolo_id,canale_id,quantita&limit=100000`, {
      headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}` }
    }).then(r => r.json()).then(data => { if (Array.isArray(data)) setPrenotatoCliente(data); });
  }, [clienteSel, token]);

  useEffect(() => {
    if (!token) return;
    const auroraCanale = canali.find(c => c.codice === "AURORA");
    if (!auroraCanale) return;
    fetch(`${SUPABASE_URL}/rest/v1/prenotato?canale_id=eq.${auroraCanale.id}&select=titolo_id,quantita&limit=10000`, {
      headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}` }
    }).then(r => r.json()).then(data => {
      if (!Array.isArray(data)) return;
      const map = {}; data.forEach(r => { map[r.titolo_id] = r.quantita; }); setAuroraEdit(map);
    });
  }, [token, canali]);

  const saveAurora = async (titoloId, qta) => {
    const auroraCanale = canali.find(c => c.codice === "AURORA");
    if (!auroraCanale) return;
    await fetch(`${SUPABASE_URL}/rest/v1/rpc/upsert_prenotato`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}` },
      body: JSON.stringify({ payload: [{ titolo_id: titoloId, canale_id: auroraCanale.id, quantita: parseInt(qta) || 0 }] }),
    });
    setAuroraEdit(prev => ({ ...prev, [titoloId]: parseInt(qta) || 0 }));
    setAuroraEditing(null);
  };

  const cedole = useMemo(() => { if (giroLabelSel.length === 0) return []; return [...new Set(titoli.filter(t => giroLabelSel.includes(t.giro_label)).map(t => t.n_cedola).filter(Boolean))].sort(); }, [titoli, giroLabelSel]);
  const editori = useMemo(() => { const t = giroLabelSel.length > 0 ? titoli.filter(t => giroLabelSel.includes(t.giro_label)) : extraSel.length > 0 ? titoli.filter(t => t.giro_label === "EXTRA" && extraSel.includes(t.n_cedola)) : []; return [...new Set(t.map(t => t.editore_nome).filter(Boolean))].sort(); }, [titoli, giroLabelSel, extraSel]);
  const accounts = useMemo(() => { const t = giroLabelSel.length > 0 ? titoli.filter(t => giroLabelSel.includes(t.giro_label)) : extraSel.length > 0 ? titoli.filter(t => t.giro_label === "EXTRA" && extraSel.includes(t.n_cedola)) : []; return [...new Set(t.map(t => t.account_editore).filter(Boolean))].sort(); }, [titoli, giroLabelSel, extraSel]);
  const normPromo = p => p ? p.trim().toUpperCase() : p;
  const promozioni = useMemo(() => { const t = giroLabelSel.length > 0 ? titoli.filter(t => giroLabelSel.includes(t.giro_label)) : extraSel.length > 0 ? titoli.filter(t => t.giro_label === "EXTRA" && extraSel.includes(t.n_cedola)) : []; return [...new Set(t.map(t => normPromo(t.promozione)).filter(Boolean))].sort(); }, [titoli, giroLabelSel, extraSel]);

  const titoliSel = useMemo(() => {
    if (giroLabelSel.length === 0 && extraSel.length === 0) return [];
    return titoli
      .filter(t => extraSel.length > 0 ? (t.giro_label === "EXTRA" && extraSel.includes(t.n_cedola)) : giroLabelSel.includes(t.giro_label))
      .filter(t => cedolaSel.length === 0 || cedolaSel.includes(t.n_cedola))
      .filter(t => filterEditori.length === 0 || filterEditori.includes(t.editore_nome))
      .filter(t => filterAccount.length === 0 || filterAccount.includes(t.account_editore))
      .filter(t => filterPromozione.length === 0 || filterPromozione.includes(normPromo(t.promozione)))
      .filter(t => { if (!search) return true; const q = search.toLowerCase(); return t.titolo?.toLowerCase().includes(q) || t.ean?.includes(q); })
      .sort((a, b) => (a.ranking_editore ?? 99) - (b.ranking_editore ?? 99) || (a.ranking_titolo ?? 99) - (b.ranking_titolo ?? 99));
  }, [titoli, giroLabelSel, extraSel, cedolaSel, filterEditori, filterAccount, filterPromozione, search]);

  const macrogruppiVis = ruolo === "agente" ? MACROGRUPPI.filter(mg => mg.id === "RETE" || mg.id === "GROSSISTI") : MACROGRUPPI;

  const byCanaleCliente = useMemo(() => {
    if (!clienteSel || prenotatoCliente.length === 0) return null;
    const map = {};
    prenotatoCliente.forEach(p => { const c = canali.find(c => c.id === p.canale_id); if (c) map[c.codice] = (map[c.codice] || 0) + p.quantita; });
    return map;
  }, [clienteSel, prenotatoCliente, canali]);

  const canaliTabella = useMemo(() => {
    const CANALI_AGENTE = [...CANALI_RETE, "FASTBOOK", "CENTROLIBRI", "GROSSISTI"];
    let base = ruolo === "agente" ? canali.filter(c => CANALI_AGENTE.includes(c.codice)) : canali.filter(c => c.codice !== "AURORA" && c.codice !== "GDO");
    // Se c'è un cliente selezionato, mostra solo la colonna del suo canale
    if (clienteSel && byCanaleCliente) {
      const codiciCliente = Object.keys(byCanaleCliente);
      if (codiciCliente.length > 0) base = base.filter(c => codiciCliente.includes(c.codice));
    } else if (filterCanale.length > 0) {
      base = base.filter(c => filterCanale.includes(c.codice));
    }
    // Ordina seguendo l'ordine dei MACROGRUPPI (Rete → Catene → Grossisti → Online)
    const ordine = MACROGRUPPI.flatMap(mg => mg.canali);
    base = [...base].sort((a, b) => {
      const ia = ordine.indexOf(a.codice);
      const ib = ordine.indexOf(b.codice);
      return (ia === -1 ? 999 : ia) - (ib === -1 ? 999 : ib);
    });
    return base;
  }, [canali, ruolo, filterCanale, clienteSel, byCanaleCliente]);

  const prenotatoClienteByTitolo = useMemo(() => {
    if (!clienteSel || prenotatoCliente.length === 0) return null;
    const map = {};
    prenotatoCliente.forEach(p => {
      if (!p.titolo_id) return;
      if (!map[p.titolo_id]) map[p.titolo_id] = {};
      const c = canali.find(c => c.id === p.canale_id);
      if (c) map[p.titolo_id][c.codice] = (map[p.titolo_id][c.codice] || 0) + p.quantita;
    });
    return map;
  }, [clienteSel, prenotatoCliente, canali]);

  // Helper: calcola obiettivo per canale di un singolo titolo via spalmatura
  const getObjCanalePerTitolo = useCallback((t, canale_codice) => {
    const spRow = spalmatura.find(s => s.editore_nome === t.editore_nome && s.formato === (t.formato || 'Cover') && s.canale_codice === canale_codice);
    if (!spRow || !t.obiettivo_assegnato) return 0;
    return Math.round(t.obiettivo_assegnato * spRow.percentuale);
  }, [spalmatura]);

  const righe = useMemo(() => titoliSel.map(t => {
    // Calcola obiettivo per canale per questo titolo
    const byCanaleObj = {};
    canali.forEach(c => { byCanaleObj[c.codice] = getObjCanalePerTitolo(t, c.codice); });

    if (clienteSel && prenotatoClienteByTitolo) {
      const byCanale = prenotatoClienteByTitolo[t.id] || {};
      const totPren = Object.values(byCanale).reduce((s, v) => s + v, 0);
      const dirStampatore = ruolo !== "agente" ? (auroraEdit[t.id] || 0) : 0;
      return { titolo: t, totPren, byCanale, byCanaleObj, dirStampatore };
    }
    const pren = prenotato.filter(p => p.titolo_id === t.id);
    const totPrenBase = pren.reduce((s, p) => s + p.quantita, 0);
    const byCanale = {};
    // AURORA da upload va in byCanale normalmente (dentro Grossisti)
    pren.forEach(p => { const c = canali.find(c => c.id === p.canale_id); if (c) byCanale[c.codice] = p.quantita; });
    // Dir.Stamp./Fuori lancio: valore editabile manuale, separato
    const dirStampatore = ruolo !== "agente" ? (auroraEdit[t.id] || 0) : 0;
    if (ruolo === "agente") {
      const totRete = CANALI_RETE.reduce((s, cod) => s + (byCanale[cod] || 0), 0);
      return { titolo: t, totPren: totRete, byCanale, byCanaleObj, dirStampatore: 0 };
    }
    return { titolo: t, totPren: totPrenBase, byCanale, byCanaleObj, dirStampatore };
  }), [titoliSel, prenotato, canali, auroraEdit, ruolo, clienteSel, prenotatoClienteByTitolo, getObjCanalePerTitolo]);

  const [sortCanale, setSortCanale] = useState(null); // { codice, dir: "asc"|"desc" } o null

  const righeFiltrate = useMemo(() => {
    let result = [...righe];
    // Se soloPrenotati è ON e c'è un filtro canale o cliente attivo, mostra solo titoli con prenotato > 0
    if (soloPrenotati && filterCanale.length > 0) {
      result = result.filter(({ byCanale }) => filterCanale.some(cod => (byCanale[cod] || 0) > 0));
    }
    if (soloPrenotati && clienteSel) {
      result = result.filter(({ totPren }) => totPren > 0);
    }
    if (sortPren) {
      result = [...result].sort((a, b) => sortPren === "asc" ? a.totPren - b.totPren : b.totPren - a.totPren);
    }
    // Sort per colonna canale specifica
    if (sortCanale) {
      result = [...result].sort((a, b) => {
        const va = a.byCanale[sortCanale.codice] || 0;
        const vb = b.byCanale[sortCanale.codice] || 0;
        return sortCanale.dir === "asc" ? va - vb : vb - va;
      });
    }
    return result;
  }, [righe, filterCanale, sortPren, soloPrenotati, clienteSel, sortCanale]);

  const totObj = righeFiltrate.reduce((s, r) => s + (r.titolo.obiettivo_assegnato || 0), 0);
  const totPrenotato = righeFiltrate.reduce((s, r) => s + r.totPren, 0);
  const pctTot = totObj > 0 ? Math.round(totPrenotato / totObj * 100) : 0;

  const prenotatoPerCanale = useMemo(() => {
    const map = {};
    righeFiltrate.forEach(({ byCanale }) => { Object.entries(byCanale).forEach(([cod, qta]) => { map[cod] = (map[cod] || 0) + (qta || 0); }); });
    return map;
  }, [righeFiltrate]);

  const totMacro = useMemo(() => { const map = {}; macrogruppiVis.forEach(mg => { map[mg.id] = mg.canali.reduce((s, cod) => s + (prenotatoPerCanale[cod] || 0), 0); }); return map; }, [prenotatoPerCanale, macrogruppiVis]);

  // MOD 4: Calcolo obiettivi per canale nello scarico (Fine Giro)
  const obiPerCanaleFinGiro = useMemo(() => {
    const map = {};
    canali.forEach(c => {
      let assegnato = 0;
      righeFiltrate.forEach(({ titolo: t }) => {
        const spRow = spalmatura.find(s => s.editore_nome === t.editore_nome && s.formato === (t.formato || 'Cover') && s.canale_codice === c.codice);
        if (spRow && t.obiettivo_assegnato) {
          assegnato += Math.round(t.obiettivo_assegnato * spRow.percentuale);
        }
      });
      const raggiunto = prenotatoPerCanale[c.codice] || 0;
      map[c.codice] = { assegnato, raggiunto };
    });
    return map;
  }, [righeFiltrate, canali, spalmatura, prenotatoPerCanale]);

  const clientiFiltrati = useMemo(() => { if (!clienteSearch) return clientiList; const q = clienteSearch.toLowerCase(); return clientiList.filter(c => c.nome.toLowerCase().includes(q) || c.cod.includes(q)); }, [clientiList, clienteSearch]);

  const exportExcel = async () => {
    const XLSX = window.XLSX;
    const colCanali = canaliTabella;
    const dirStampHeaders = ruolo !== "agente" ? ["Dir.Stamp./Fuori lancio"] : [];
    const headers = ["N° CEDOLA","EAN","TITOLO","AUTORE","COD.EDITORE","EDITORE","PREZZO","OBJ ASS.","PRENOTATO","AVANZ %",...colCanali.map(c => getCanaleDisplayName(c)),...dirStampHeaders];
    const rows = righeFiltrate.map(({ titolo: t, totPren, byCanale, dirStampatore }) => {
      const pct = t.obiettivo_assegnato > 0 ? Math.round(totPren / t.obiettivo_assegnato * 100) : 0;
      const base = [t.n_cedola, t.ean, t.titolo, t.autore, t.codice_editore, t.editore_nome, t.prezzo, t.obiettivo_assegnato, totPren, `${pct}%`, ...colCanali.map(c => byCanale[c.codice] ?? 0)];
      if (ruolo !== "agente") base.push(dirStampatore || 0);
      return base;
    });
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    // Foglio RIEPILOGO con obiettivi per canale
    const riepilogoRows = [["MACROGRUPPO","CANALE","OBJ ASSEGNATO","PRENOTATO","AVANZ %"]];
    macrogruppiVis.forEach(mg => {
      const totMgAss = mg.canali.reduce((s, cod) => s + (obiPerCanaleFinGiro[cod]?.assegnato || 0), 0);
      const totMgRag = mg.canali.reduce((s, cod) => s + (obiPerCanaleFinGiro[cod]?.raggiunto || 0), 0);
      const pctMg = totMgAss > 0 ? Math.round(totMgRag / totMgAss * 100) : 0;
      riepilogoRows.push([mg.label, "TOTALE", totMgAss, totMgRag, `${pctMg}%`]);
      mg.canali.forEach(cod => {
        const c = canali.find(c => c.codice === cod);
        if (c) {
          const { assegnato, raggiunto } = obiPerCanaleFinGiro[cod] || { assegnato: 0, raggiunto: 0 };
          const pctC = assegnato > 0 ? Math.round(raggiunto / assegnato * 100) : 0;
          riepilogoRows.push(["", getCanaleDisplayName(c), assegnato, raggiunto, assegnato > 0 ? `${pctC}%` : "—"]);
        }
      });
      riepilogoRows.push(["", "", "", "", ""]);
    });
    const wsRiepilogo = XLSX.utils.aoa_to_sheet(riepilogoRows);
    wsRiepilogo["!cols"] = [{ wch: 25 }, { wch: 30 }, { wch: 14 }, { wch: 12 }, { wch: 10 }];

    // Foglio DETTAGLIO PRENOTATO: dettaglio per punto vendita (come il file caricato)
    const titoloIds = righeFiltrate.map(({ titolo: t }) => t.id);
    let dettaglioRows = [["LIBRERIA", "COD. CLIENTE", "EAN", "TITOLO", "CANALE", "QUANTITÀ"]];
    if (titoloIds.length > 0) {
      try {
        const res = await fetch(
          `${SUPABASE_URL}/rest/v1/prenotato_clienti?titolo_id=in.(${titoloIds.join(",")})&select=codice_cliente,nome_cliente,titolo_id,canale_id,quantita&limit=100000`,
          { headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}` } }
        );
        const data = await res.json();
        if (Array.isArray(data)) {
          const titoloById = {};
          righeFiltrate.forEach(({ titolo: t }) => { titoloById[t.id] = t; });
          const canaleById = {};
          canali.forEach(c => { canaleById[c.id] = getCanaleDisplayName(c); });
          const righe = data
            .filter(r => r.quantita > 0)
            .map(r => {
              const t = titoloById[r.titolo_id];
              return [r.nome_cliente || "—", r.codice_cliente, t?.ean || "", t?.titolo || "", canaleById[r.canale_id] || "—", r.quantita];
            })
            .sort((a, b) => a[0].localeCompare(b[0], "it", { sensitivity: "base" }));
          const totQta = righe.reduce((s, r) => s + r[5], 0);
          dettaglioRows = dettaglioRows.concat(righe);
          dettaglioRows.push(["", "", "", "", "TOTALE", totQta]);
        }
      } catch (e) {
        console.error("Errore fetch dettaglio prenotato:", e);
      }
    }
    const wsDettaglio = XLSX.utils.aoa_to_sheet(dettaglioRows);
    wsDettaglio["!cols"] = [{ wch: 35 }, { wch: 14 }, { wch: 15 }, { wch: 40 }, { wch: 20 }, { wch: 12 }];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "FINE GIRO");
    XLSX.utils.book_append_sheet(wb, wsRiepilogo, "RIEPILOGO");
    XLSX.utils.book_append_sheet(wb, wsDettaglio, "DETTAGLIO PRENOTATO");
    XLSX.writeFile(wb, `FINE_GIRO_${giroLabelSel.length > 0 ? giroLabelSel.join("-") : extraSel.length > 0 ? extraSel.join("-") : "EXPORT"}.xlsx`);
  };

  if (giroLabelSel.length === 0 && extraSel.length === 0) {
    return (
      <div style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "flex-start", gap: 28, padding: "48px 20px", overflowY: "auto" }}>
        <div style={{ color: T.textMid, fontSize: "14px" }}>Seleziona un giro o una cedola extra</div>
        <div style={{ display: "flex", gap: 48, alignItems: "flex-start", justifyContent: "center", flexWrap: "wrap" }}>

          {/* Colonna GIRI VENDITA */}
          <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 14, width: 240 }}>
            <div style={{ color: T.text, fontSize: "12px", fontWeight: 700, letterSpacing: "0.5px", textTransform: "uppercase" }}>Giri Vendita</div>
            <SearchableMultiSelect values={filterAnnoFineGiro.map(String)} onChange={v => setFilterAnnoFineGiro(v.map(Number))} options={anniDispFineGiro.map(String)} renderOption={v => v} placeholder="Anno" width={140} />
            <div style={{ display: "flex", flexDirection: "column", gap: 8, width: "100%" }}>
              {giriLabel.length === 0 && <div style={{ color: T.textMid, fontSize: "12px", textAlign: "center", padding: "10px 0" }}>Nessun giro per l'anno selezionato</div>}
              {giriLabel.map(g => (
                <button key={g} style={{ ...css.btn("accent"), padding: "10px 16px", fontSize: "13px", width: "100%" }} onClick={() => setGiroLabelSel([g])}>Giro {g}</button>
              ))}
            </div>
          </div>

          {/* Colonna EXTRAGIRI */}
          <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 14, width: 240 }}>
            <div style={{ color: T.text, fontSize: "12px", fontWeight: 700, letterSpacing: "0.5px", textTransform: "uppercase" }}>Extragiri</div>
            <SearchableMultiSelect values={filterAnnoExtraFineGiro.map(String)} onChange={v => setFilterAnnoExtraFineGiro(v.map(Number))} options={anniDispExtraFineGiro.map(String)} renderOption={v => v} placeholder="Anno" width={140} />
            <div style={{ display: "flex", flexDirection: "column", gap: 8, width: "100%" }}>
              {cedoleExtra.length === 0 && <div style={{ color: T.textMid, fontSize: "12px", textAlign: "center", padding: "10px 0" }}>Nessuna cedola extra per l'anno selezionato</div>}
              {cedoleExtra.map(c => (
                <button key={c} style={{ ...css.btn(), padding: "10px 16px", fontSize: "13px", borderColor: T.accent, color: T.accent, width: "100%" }} onClick={() => setExtraSel([c])}>{c}</button>
              ))}
            </div>
          </div>

        </div>
      </div>
    );
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%", overflow: "hidden" }}>
      {/* Toolbar filtri */}
      <div style={{ padding: "12px 20px", borderBottom: `1px solid ${T.border}`, display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
        <SearchableMultiSelect values={filterAnnoFineGiro.map(String)} onChange={v => { setFilterAnnoFineGiro(v.map(Number)); setGiroLabelSel([]); setExtraSel([]); setCedolaSel([]); setFilterEditori([]); setFilterAccount([]); setFilterPromozione([]); setFilterCanale([]); setSearch(""); setClienteSel(null); setSoloPrenotati(false); }} options={anniDispFineGiro.map(String)} renderOption={v => v} placeholder="Anno" width={110} />
        <SearchableMultiSelect values={giroLabelSel} onChange={v => { setGiroLabelSel(v); setExtraSel([]); setCedolaSel([]); setFilterEditori([]); setFilterAccount([]); setFilterPromozione([]); setFilterCanale([]); setSearch(""); setClienteSel(null); setSoloPrenotati(false); }} options={giriLabel} placeholder="— Giro —" width={160} renderOption={g => `Giro ${g}`} />
        {cedoleExtra.length > 0 && (
          <SearchableMultiSelect values={extraSel} onChange={v => { setExtraSel(v); setGiroLabelSel([]); setCedolaSel([]); setFilterEditori([]); setFilterAccount([]); setFilterPromozione([]); setFilterCanale([]); setSearch(""); setClienteSel(null); setSoloPrenotati(false); }} options={cedoleExtra} placeholder="Cedole Extra" width={170} />
        )}
        {giroLabelSel.length > 0 && (
          <SearchableMultiSelect values={cedolaSel} onChange={setCedolaSel} options={cedole} placeholder="Tutte le cedole" width={160} />
        )}
        <SearchableMultiSelect values={filterEditori} onChange={setFilterEditori} options={editori} placeholder="Tutti gli editori" width={190} />
        <SearchableMultiSelect values={filterAccount} onChange={setFilterAccount} options={accounts} placeholder="Tutti gli account" width={160} />
        <SearchableMultiSelect values={filterPromozione} onChange={setFilterPromozione} options={promozioni} renderOption={p => p === "PDE SERVICE" ? "PDE Service" : p === "0" ? "Nessuna" : "PDE Promozione"} placeholder="Tutte le promozioni" width={170} />
        <SearchableMultiSelect values={filterCanale} onChange={setFilterCanale} options={canali.filter(c => c.codice !== "AURORA" && c.codice !== "GDO").map(c => c.codice)} renderOption={cod => { const c = canali.find(x => x.codice === cod); return c?.nome || cod; }} placeholder="Tutti i canali" width={170} />
        <input style={{ ...css.input, width: 160 }} placeholder="Cerca EAN / titolo..." value={search} onChange={e => setSearch(e.target.value)} />
        {(filterCanale.length > 0 || clienteSel) && (
          <label style={{ display: "flex", alignItems: "center", gap: 5, cursor: "pointer", fontSize: "11px", color: soloPrenotati ? T.accent : T.textMid, userSelect: "none" }}>
            <input type="checkbox" checked={soloPrenotati} onChange={e => setSoloPrenotati(e.target.checked)} style={{ accentColor: T.accent }} />
            Solo prenotati
          </label>
        )}
        {ruolo !== "agente" && (
          <div style={{ position: "relative" }}>
            <button style={{ ...css.btn(clienteSel ? "accent" : "default"), minWidth: 160, display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }} onClick={() => setClienteDropdownOpen(o => !o)}>
              <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 130 }}>{clienteSel ? clientiList.find(c => c.cod === clienteSel)?.nome || clienteSel : "Tutti i clienti"}</span>
              <span>▾</span>
            </button>
            {clienteDropdownOpen && (
              <div style={{ position: "absolute", top: "100%", left: 0, zIndex: 50, background: T.surface, border: `1px solid ${T.borderHi}`, borderRadius: 4, minWidth: 260, maxHeight: 300, overflowY: "auto", marginTop: 4, boxShadow: "0 4px 20px #0008" }}>
                <div style={{ padding: "8px 12px", borderBottom: `1px solid ${T.border}`, position: "sticky", top: 0, background: T.surface }}>
                  <input style={{ ...css.input, width: "100%", boxSizing: "border-box" }} placeholder="Cerca cliente..." value={clienteSearch} onChange={e => setClienteSearch(e.target.value)} autoFocus />
                </div>
                <div style={{ padding: "6px 12px", cursor: "pointer", color: !clienteSel ? T.accent : T.textMid, fontSize: "12px", borderBottom: `1px solid ${T.border}22` }}
                  onClick={() => { setClienteSel(null); setClienteDropdownOpen(false); setClienteSearch(""); }}>Tutti i clienti</div>
                {clientiFiltrati.map(c => (
                  <div key={c.cod} style={{ padding: "6px 12px", cursor: "pointer", fontSize: "12px", color: clienteSel === c.cod ? T.accent : T.text, background: clienteSel === c.cod ? T.accent + "18" : "transparent", borderBottom: `1px solid ${T.border}22` }}
                    onClick={() => { setClienteSel(c.cod); setClienteDropdownOpen(false); setClienteSearch(""); }}>
                    <div style={{ fontWeight: "600" }}>{c.nome}</div>
                    <div style={{ fontSize: "10px", color: T.textMid }}>{c.cod}</div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
        <div style={{ color: T.textMid, fontSize: "12px" }}>
          Obj: <span style={{ color: T.accent, fontWeight: "700" }}>{totObj.toLocaleString("it")}</span>
          &nbsp;·&nbsp; Pren: <span style={{ color: T.green, fontWeight: "700" }}>{totPrenotato.toLocaleString("it")}</span>
          &nbsp;·&nbsp; <span style={{ color: pctTot >= 80 ? T.green : pctTot >= 50 ? T.accent : T.red, fontWeight: "700" }}>{pctTot}%</span>
        </div>
        <button style={css.btn()} onClick={resetFiltri}>↺ Reset</button>
        <button style={{ ...css.btn("accent"), marginLeft: "auto" }} onClick={exportExcel}>↓ Export Excel</button>
      </div>

      {/* Riepilogo cliente */}
      {clienteSel && byCanaleCliente && (
        <div style={{ padding: "10px 20px", borderBottom: `1px solid ${T.border}`, background: T.accent + "11" }}>
          <div style={{ color: T.accent, fontWeight: "700", fontSize: "11px", marginBottom: 6 }}>
            CLIENTE: {clientiList.find(c => c.cod === clienteSel)?.nome || clienteSel}
          </div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 16 }}>
            {Object.entries(byCanaleCliente).sort((a,b) => b[1]-a[1]).map(([cod, qta]) => {
              const c = canali.find(c => c.codice === cod);
              return c ? <div key={cod} style={{ fontSize: "12px" }}><span style={{ color: T.textMid }}>{getCanaleDisplayName(c)}: </span><span style={{ color: T.green, fontWeight: "700" }}>{qta.toLocaleString("it")}</span></div> : null;
            })}
            <div style={{ fontSize: "12px", marginLeft: "auto" }}>
              <span style={{ color: T.textMid }}>Totale: </span>
              <span style={{ color: T.accent, fontWeight: "700" }}>{Object.values(byCanaleCliente).reduce((s,v)=>s+v,0).toLocaleString("it")}</span>
            </div>
          </div>
        </div>
      )}

      {/* Macrogruppi — box compatti */}
      <div style={{ padding: "12px 20px", borderBottom: `1px solid ${T.border}`, display: "flex", gap: 12, flexWrap: "wrap" }}>
        {macrogruppiVis.map(mg => {
          const totMgAss = mg.canali.reduce((s, cod) => s + (obiPerCanaleFinGiro[cod]?.assegnato || 0), 0);
          const totMgRag = mg.canali.reduce((s, cod) => s + (obiPerCanaleFinGiro[cod]?.raggiunto || 0), 0);
          const pctMg = totMgAss > 0 ? Math.round(totMgRag / totMgAss * 100) : 0;
          return (
            <div key={mg.id} style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 4, padding: "10px 14px", minWidth: 180 }}>
              <div style={{ color: T.accent, fontWeight: "700", fontSize: "11px", letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 4 }}>{mg.label}</div>
              <div style={{ display: "flex", gap: 10, alignItems: "center", marginBottom: 8, paddingBottom: 6, borderBottom: `1px solid ${T.border}44` }}>
                <div style={{ fontSize: "10px", color: T.textMid }}>Ass: <span style={{ color: T.text, fontWeight: "600" }}>{totMgAss.toLocaleString("it")}</span></div>
                <div style={{ fontSize: "10px", color: T.textMid }}>Rag: <span style={{ color: T.green, fontWeight: "600" }}>{totMgRag.toLocaleString("it")}</span></div>
                {totMgAss > 0 && <span style={{ color: pctMg >= 80 ? T.green : pctMg >= 50 ? T.accent : T.red, fontSize: "11px", fontWeight: "700" }}>{pctMg}%</span>}
              </div>
              {mg.canali.map(cod => {
                const c = canali.find(c => c.codice === cod); if (!c) return null;
                const { assegnato, raggiunto } = obiPerCanaleFinGiro[cod] || { assegnato: 0, raggiunto: 0 };
                const qta = prenotatoPerCanale[cod] || 0;
                const pctC = assegnato > 0 ? Math.round(raggiunto / assegnato * 100) : 0;
                return (
                  <div key={cod} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: "11px", padding: "3px 0", borderTop: `1px solid ${T.border}22` }}>
                    <span style={{ color: T.textMid, flex: 1 }}>{getCanaleDisplayName(c)}</span>
                    <span style={{ color: T.textDim, fontSize: "10px", width: 50, textAlign: "right" }}>{assegnato > 0 ? assegnato.toLocaleString("it") : "—"}</span>
                    <span style={{ color: qta > 0 ? T.green : T.textDim, fontWeight: qta > 0 ? "600" : "400", width: 50, textAlign: "right" }}>{qta > 0 ? qta.toLocaleString("it") : "—"}</span>
                    <span style={{ color: pctC >= 80 ? T.green : pctC >= 50 ? T.accent : T.red, fontWeight: "700", width: 36, textAlign: "right", fontSize: "10px" }}>{assegnato > 0 ? `${pctC}%` : ""}</span>
                  </div>
                );
              })}
            </div>
          );
        })}
      </div>

      {/* Tabella */}
      <div style={{ flex: 1, overflowY: "auto", overflowX: "auto" }}>
        <table style={css.table}>
          <thead>
            <tr>
              <th style={css.th}>Cedola</th><th style={css.th}>EAN</th><th style={css.th}>Titolo</th>
              <th style={css.th}>Autore</th><th style={css.th}>Cod.Ed.</th><th style={css.th}>Editore</th>
              <th style={css.th}>€</th><th style={css.th}>Obj</th><th style={{ ...css.th, cursor: "pointer" }} onClick={() => { setSortPren(s => s === "desc" ? "asc" : s === "asc" ? null : "desc"); setSortCanale(null); }}>Pren.{sortPren === "desc" ? " ↓" : sortPren === "asc" ? " ↑" : ""}</th><th style={css.th}>%</th>
              {canaliTabella.map(c => <th key={c.id} style={{ ...css.th, whiteSpace: "normal", maxWidth: 70, lineHeight: 1.2, cursor: "pointer" }} onClick={() => { setSortCanale(prev => prev?.codice === c.codice ? (prev.dir === "desc" ? { codice: c.codice, dir: "asc" } : null) : { codice: c.codice, dir: "desc" }); setSortPren(null); }}>{getCanaleDisplayName(c)}{sortCanale?.codice === c.codice ? (sortCanale.dir === "desc" ? " ↓" : " ↑") : ""}</th>)}
              {ruolo !== "agente" && <th style={{ ...css.th, color: T.accent }}>Dir.Stamp./Fuori lancio ✎</th>}
            </tr>
          </thead>
          <tbody>
            {righeFiltrate.map(({ titolo: t, totPren, byCanale, byCanaleObj, dirStampatore }, i) => {
              const pct = t.obiettivo_assegnato > 0 ? Math.round(totPren / t.obiettivo_assegnato * 100) : 0;
              const isEditingAurora = auroraEditing === t.id;
              return (
                <tr key={t.id} style={{ background: i % 2 === 0 ? "transparent" : T.surface + "66" }}>
                  <td style={{ ...css.td, color: T.textMid, fontSize: "10px", whiteSpace: "nowrap" }}>{t.n_cedola}</td>
                  <td style={{ ...css.td, fontFamily: "monospace", fontSize: "11px", color: T.textMid }}>{t.ean}</td>
                  <td style={{ ...css.td, maxWidth: 200 }}><div style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{t.titolo}</div></td>
                  <td style={{ ...css.td, color: T.textMid }}>{t.autore}</td>
                  <td style={{ ...css.td, color: T.textMid, fontSize: "11px" }}>{t.codice_editore}</td>
                  <td style={{ ...css.td, color: T.accent, fontWeight: "600", whiteSpace: "nowrap" }}>{t.editore_nome}</td>
                  <td style={{ ...css.td, whiteSpace: "nowrap" }}>€ {t.prezzo?.toFixed(2)}</td>
                  <td style={css.td}>{t.obiettivo_assegnato?.toLocaleString("it")}</td>
                  <td style={{ ...css.td, color: T.green, fontWeight: "600" }}>{totPren > 0 ? totPren.toLocaleString("it") : "—"}</td>
                  <td style={css.td}><span style={{ color: pct >= 80 ? T.green : pct >= 50 ? T.accent : T.red, fontWeight: "700" }}>{pct}%</span></td>
                  {canaliTabella.map(c => {
                    const qta = byCanale[c.codice] || 0;
                    const obj = byCanaleObj?.[c.codice] || 0;
                    return (
                      <td key={c.id} style={{ ...css.td, whiteSpace: "nowrap", minWidth: 60 }}>
                        {qta > 0 || obj > 0 ? (
                          <div>
                            <span style={{ color: qta > 0 ? T.text : T.textDim, fontWeight: "600", fontSize: "12px" }}>{qta > 0 ? qta.toLocaleString("it") : "0"}</span>
                            {obj > 0 && <span style={{ color: T.textDim, fontSize: "10px" }}> / {obj.toLocaleString("it")}</span>}
                          </div>
                        ) : <span style={{ color: T.textDim }}>—</span>}
                      </td>
                    );
                  })}
                  {ruolo !== "agente" && (
                    <td style={css.td}>
                      {isEditingAurora ? (
                        <div style={{ display: "flex", gap: 4 }}>
                          <input type="number" defaultValue={auroraEdit[t.id] || 0} style={{ ...css.input, width: 70, padding: "2px 6px" }} id={`aurora-${t.id}`} autoFocus
                            onKeyDown={e => { if (e.key === "Enter") saveAurora(t.id, e.target.value); if (e.key === "Escape") setAuroraEditing(null); }} />
                          <button style={{ ...css.btn("accent"), padding: "2px 8px", fontSize: "11px" }} onClick={() => saveAurora(t.id, document.getElementById(`aurora-${t.id}`).value)}>✓</button>
                        </div>
                      ) : (
                        <div style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer" }} onClick={() => setAuroraEditing(t.id)}>
                          <span style={{ color: auroraEdit[t.id] ? T.text : T.textDim }}>{auroraEdit[t.id]?.toLocaleString("it") || "—"}</span>
                          <span style={{ color: T.accent, fontSize: "11px" }}>✎</span>
                        </div>
                      )}
                    </td>
                  )}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ────────────────────────────────────────────────────────────────
// MODULO AVANZAMENTO TITOLI NOVITÀ
// ────────────────────────────────────────────────────────────────

// Mesi italiani per parsing date dal CSV Tableau ("10 ottobre 2025")
const MESI_IT = { gennaio:0, febbraio:1, marzo:2, aprile:3, maggio:4, giugno:5, luglio:6, agosto:7, settembre:8, ottobre:9, novembre:10, dicembre:11 };

function parseDataIt(str) {
  if (!str) return null;
  const parts = String(str).trim().split(/\s+/);
  if (parts.length !== 3) return null;
  const [giorno, mese, anno] = parts;
  const m = MESI_IT[mese?.toLowerCase()];
  if (m === undefined) return null;
  return new Date(Number(anno), m, Number(giorno));
}

function fmtDate(d) {
  if (!d) return "—";
  if (typeof d === "string") d = new Date(d);
  if (isNaN(d)) return "—";
  return d.toLocaleDateString("it-IT", { day: "2-digit", month: "short", year: "numeric" });
}

function parseValoreLancio(str) {
  if (typeof str === "number") return str;
  if (!str) return 0;
  // Formato italiano: 2.178.600 € oppure 2.178.600,50 €
  // Rimuovi simbolo € e spazi, poi togli tutti i punti (migliaia), converti virgola in punto decimale
  const s = String(str).replace(/[€\s]/g, "").replace(/\./g, "").replace(",", ".");
  return parseFloat(s) || 0;
}

function parsePrezzo(str) {
  if (typeof str === "number") return str;
  if (!str) return 0;
  return parseFloat(String(str).replace(",", ".")) || 0;
}



function ModuloLanciSettimanali({ token, titoli, prenotato, canali, ruolo, userAccount, onNavigateAnticipi }) {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [uploading, setUploading] = useState(false);
  const [uploadMode, setUploadMode] = useState(null);
  const [toast, setToast] = useState(null);
  const [filterAnno, setFilterAnno] = useState(null);
  const [filterLancio, setFilterLancio] = useState([]);
  const [filterAccount, setFilterAccount] = useState([]);
  const [filterGiorno, setFilterGiorno] = useState("tutti");
  const [search, setSearch] = useState("");
  const [sortKey, setSortKey] = useState(null);
  const [sortDir, setSortDir] = useState("desc");

  const [editCell, setEditCell] = useState(null); // { id, field, value }
  const [anticipiPopup, setAnticipiPopup] = useState(null); // array di righe novita_fuori_lancio notificate da questo upload

  // Dopo un upload, controlla se qualche EAN caricato sblocca un anticipo lancio "da_gestire" → "notificato"
  const checkAnticipiNotificati = async (eans) => {
    if (!eans || eans.length === 0 || !token) return;
    const trovati = [];
    for (let i = 0; i < eans.length; i += 150) {
      const batch = eans.slice(i, i + 150);
      const list = batch.map(e => `"${e}"`).join(",");
      const res = await sbFetch(`novita_fuori_lancio?select=*&ean=in.(${list})&stato=eq.notificato`, token);
      if (Array.isArray(res)) trovati.push(...res);
    }
    if (trovati.length > 0) setAnticipiPopup(trovati);
  };

  const showToast = (msg, type = "ok") => { setToast({ msg, type }); setTimeout(() => setToast(null), 4000); };

  const loadData = useCallback(async () => {
    if (!token) return;
    setLoading(true);
    const d = await sbFetch("lanci_settimanali?select=*&order=anno_lancio.desc,num_lancio.desc,editore.asc,titolo.asc", token);
    if (Array.isArray(d)) setData(d);
    setLoading(false);
  }, [token]);

  useEffect(() => { loadData(); }, [loadData]);

  // Anni e lanci
  const anniDisponibili = useMemo(() => [...new Set(data.map(r => r.anno_lancio))].sort((a, b) => b - a), [data]);
  const lanciPerAnno = useMemo(() => {
    const anno = filterAnno || anniDisponibili[0];
    if (!anno) return [];
    return [...new Set(data.filter(r => r.anno_lancio === anno).map(r => r.num_lancio))].sort((a, b) => b - a);
  }, [data, filterAnno, anniDisponibili]);

  useEffect(() => { if (!filterAnno && anniDisponibili.length > 0) setFilterAnno(anniDisponibili[0]); }, [anniDisponibili]);
  useEffect(() => { if (lanciPerAnno.length > 0 && filterLancio.length === 0) setFilterLancio([lanciPerAnno[0]]); }, [lanciPerAnno]);

  // Mappa prenotato per EAN (da fine giro) — per canale
  const prenByEanCanale = useMemo(() => {
    const map = {};
    prenotato.forEach(p => {
      const t = titoli.find(t => t.id === p.titolo_id);
      if (!t || !t.ean) return;
      if (!map[t.ean]) map[t.ean] = {};
      const c = canali.find(c => c.id === p.canale_id);
      if (c) map[t.ean][c.codice] = (map[t.ean][c.codice] || 0) + p.quantita;
    });
    return map;
  }, [prenotato, titoli, canali]);

  // Mappa cedole per EAN (un titolo può avere più cedole)
  const cedoleByEan = useMemo(() => {
    const map = {};
    titoli.forEach(t => {
      if (!t.ean || !t.n_cedola) return;
      if (!map[t.ean]) map[t.ean] = new Set();
      map[t.ean].add(t.n_cedola);
    });
    const result = {};
    Object.entries(map).forEach(([ean, s]) => { result[ean] = [...s].sort(); });
    return result;
  }, [titoli]);

  // Canale Amazon ID
  const amazonCodice = "AMAZON";

  // Mappa EAN → account_editore e editore_nome → account_editore (da titoli)
  const accountMaps = useMemo(() => {
    const byEan = {};
    const byEditore = {};
    titoli.forEach(t => {
      if (t.account_editore) {
        if (t.ean) byEan[t.ean] = t.account_editore;
        if (t.editore_nome) byEditore[t.editore_nome.trim().toLowerCase()] = t.account_editore;
      }
    });
    return { byEan, byEditore };
  }, [titoli]);

  // Lista account presenti nel lancio corrente
  const accountsDisponibili = useMemo(() => {
    const anno = filterAnno || anniDisponibili[0];
    const rows = data.filter(r => r.anno_lancio === anno && (filterLancio.length === 0 || filterLancio.includes(r.num_lancio)));
    return [...new Set(rows.map(r =>
      accountMaps.byEan[r.ean] || accountMaps.byEditore[r.editore?.trim().toLowerCase()] || null
    ).filter(Boolean))].sort();
  }, [data, filterAnno, filterLancio, anniDisponibili, accountMaps]);

  // Auto-selezione account per agente al cambio lancio
  useEffect(() => {
    if (ruolo === "agente" && userAccount && accountsDisponibili.includes(userAccount)) {
      setFilterAccount([userAccount]);
    }
  }, [ruolo, userAccount, accountsDisponibili]);

  // Arricchisci dati lancio con dati fine giro (live DB → fallback manuale)
  const dataArricchita = useMemo(() => {
    const anno = filterAnno || anniDisponibili[0];
    return data
      .filter(r => r.anno_lancio === anno && (filterLancio.length === 0 || filterLancio.includes(r.num_lancio)))
      .map(r => {
        const prenCanali = prenByEanCanale[r.ean] || {};
        const prenFineGiroLive = Object.entries(prenCanali).reduce((s, [cod, q]) => s + q, 0);
        const prenAmazonLive = prenCanali[amazonCodice] || 0;
        const cedoleLive = cedoleByEan[r.ean] || [];

        // Usa dati live se disponibili, altrimenti fallback manuale
        const haLiveCedole = cedoleLive.length > 0;
        const haLiveFG = prenFineGiroLive > 0;
        const haLiveAmazon = prenAmazonLive > 0;

        const cedole = haLiveCedole ? cedoleLive : (r.cedole_manual ? r.cedole_manual.split(",").map(s => s.trim()).filter(Boolean) : []);
        const prenFineGiro = haLiveFG ? prenFineGiroLive : (r.pren_fine_giro_manual || 0);
        const prenAmazon = haLiveAmazon ? prenAmazonLive : (r.pren_amazon_manual || 0);

        const prenSenzaAmazon = prenFineGiro - prenAmazon;
        // Teorico = prenotato trasmesso (o iscritto) + amazon
        const teorico = (r.prenotato_trasmesso ?? r.prenotato_iscrizione ?? 0) + prenAmazon;
        const giornoCalcolato = teorico >= 2000 ? "martedì" : "venerdì";
        const giornoUscita = r.giorno_uscita_override || giornoCalcolato;
        const isOverride = !!r.giorno_uscita_override;

        const deltaPortale = teorico - prenFineGiro;

        return {
          ...r,
          cedole,
          pren_fine_giro: prenFineGiro,
          pren_amazon: prenAmazon,
          pren_senza_amazon: prenSenzaAmazon,
          teorico,
          giorno_calcolato: giornoCalcolato,
          giorno_uscita: giornoUscita,
          is_override: isOverride,
          delta_portale: deltaPortale,
          account_editore: accountMaps.byEan[r.ean] || accountMaps.byEditore[r.editore?.trim().toLowerCase()] || null,
          // Flag per sapere se i dati vengono dal DB o sono manuali/vuoti
          is_live_cedole: haLiveCedole,
          is_live_fg: haLiveFG,
          is_live_amazon: haLiveAmazon,
          is_manual_cedole: !haLiveCedole && !!r.cedole_manual,
          is_manual_fg: !haLiveFG && r.pren_fine_giro_manual > 0,
          is_manual_amazon: !haLiveAmazon && r.pren_amazon_manual > 0,
        };
      });
  }, [data, filterAnno, filterLancio, anniDisponibili, prenByEanCanale, cedoleByEan, accountMaps]);

  // Filtro e sort
  const dataFiltrata = useMemo(() => {
    let result = [...dataArricchita];
    if (search) {
      const q = search.toLowerCase();
      result = result.filter(r => r.ean?.toLowerCase().includes(q) || r.titolo?.toLowerCase().includes(q) || r.autore?.toLowerCase().includes(q) || r.editore?.toLowerCase().includes(q));
    }
    if (filterAccount.length > 0) {
      result = result.filter(r => filterAccount.includes(r.account_editore));
    }
    if (filterGiorno !== "tutti") {
      result = result.filter(r => r.giorno_uscita === filterGiorno);
    }
    if (sortKey) {
      result = [...result].sort((a, b) => {
        const va = a[sortKey] ?? 0, vb = b[sortKey] ?? 0;
        return sortDir === "asc" ? (typeof va === "string" ? String(va).localeCompare(String(vb)) : va - vb) : (typeof va === "string" ? String(vb).localeCompare(String(va)) : vb - va);
      });
    }
    return result;
  }, [dataArricchita, search, sortKey, sortDir, filterAccount, filterGiorno]);

  // KPI
  const kpi = useMemo(() => {
    const d = dataFiltrata;
    const tot = d.length;
    const copieLanciate = d.reduce((s, r) => s + (r.prenotato_iscrizione || 0), 0);
    const copieTrasmesse = d.reduce((s, r) => s + (r.prenotato_trasmesso ?? 0), 0);
    const valoreLanciate = d.reduce((s, r) => s + (r.prezzo || 0) * (r.prenotato_iscrizione || 0), 0);
    const valoreTrasmesso = d.reduce((s, r) => s + (r.prezzo || 0) * (r.prenotato_trasmesso ?? 0), 0);
    const haTrasmesso = d.some(r => r.prenotato_trasmesso !== null);
    const totFineGiro = d.reduce((s, r) => s + (r.pren_fine_giro || 0), 0);
    const totAmazon = d.reduce((s, r) => s + (r.pren_amazon || 0), 0);
    const totTeorico = d.reduce((s, r) => s + (r.teorico || 0), 0);
    const valoreFineGiro = d.reduce((s, r) => s + (r.prezzo || 0) * (r.pren_fine_giro || 0), 0);
    const valoreAmazon = d.reduce((s, r) => s + (r.prezzo || 0) * (r.pren_amazon || 0), 0);
    const valoreTeorico = d.reduce((s, r) => s + (r.prezzo || 0) * (r.teorico || 0), 0);
    const martedi = d.filter(r => r.giorno_uscita === "martedì").length;
    const venerdi = d.filter(r => r.giorno_uscita === "venerdì").length;
    const pctTrasmesso = copieLanciate > 0 ? Math.round(copieTrasmesse / copieLanciate * 100) : 0;
    // Per editore
    const byEditore = {};
    d.forEach(r => {
      if (!byEditore[r.editore]) byEditore[r.editore] = { titoli: 0, lanciate: 0, trasmesse: 0, fineGiro: 0, amazon: 0, valore: 0 };
      byEditore[r.editore].titoli++;
      byEditore[r.editore].lanciate += r.prenotato_iscrizione || 0;
      byEditore[r.editore].trasmesse += r.prenotato_trasmesso ?? 0;
      byEditore[r.editore].fineGiro += r.pren_fine_giro || 0;
      byEditore[r.editore].amazon += r.pren_amazon || 0;
      byEditore[r.editore].valore += (r.prezzo || 0) * (r.prenotato_iscrizione || 0);
    });
    return { tot, copieLanciate, copieTrasmesse, valoreLanciate, valoreTrasmesso, haTrasmesso, totFineGiro, totAmazon, totTeorico, valoreFineGiro, valoreAmazon, valoreTeorico, martedi, venerdi, pctTrasmesso, byEditore };
  }, [dataFiltrata]);

  const toggleSort = (key) => {
    if (sortKey === key) setSortDir(d => d === "asc" ? "desc" : "asc");
    else { setSortKey(key); setSortDir("desc"); }
  };
  const sortIcon = (key) => sortKey === key ? (sortDir === "asc" ? " ↑" : " ↓") : "";

  // Salva giorno uscita override
  const saveGiornoUscita = async (id, value) => {
    const payload = { giorno_uscita_override: value || null };
    await fetch(`${SUPABASE_URL}/rest/v1/lanci_settimanali?id=eq.${id}`, {
      method: "PATCH",
      headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`, "Content-Type": "application/json", "Prefer": "return=minimal" },
      body: JSON.stringify(payload),
    });
    setData(prev => prev.map(r => r.id === id ? { ...r, giorno_uscita_override: value || null } : r));
  };

  // Salva valore manuale (cedole, fine giro, amazon) su DB
  const saveManualCell = async (id, field, value) => {
    const dbField = field === "cedole" ? "cedole_manual" : field === "fg" ? "pren_fine_giro_manual" : "pren_amazon_manual";
    const dbValue = field === "cedole" ? (value || null) : (parseInt(value) || null);
    await fetch(`${SUPABASE_URL}/rest/v1/lanci_settimanali?id=eq.${id}`, {
      method: "PATCH",
      headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`, "Content-Type": "application/json", "Prefer": "return=minimal" },
      body: JSON.stringify({ [dbField]: dbValue }),
    });
    setData(prev => prev.map(r => r.id === id ? { ...r, [dbField]: dbValue } : r));
    setEditCell(null);
  };

  // Upload handler
  const handleUpload = async (e, mode) => {
    const file = e.target.files?.[0];
    if (!file || !mode) return;
    setUploading(true);
    try {
      const XLSX = window.XLSX;
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
      let headerIdx = 0;
      for (let i = 0; i < Math.min(rows.length, 10); i++) {
        if (rows[i] && rows[i].some(c => String(c).toLowerCase().includes("ean"))) { headerIdx = i; break; }
      }
      const headers = rows[headerIdx].map(h => String(h || "").trim().toLowerCase());
      const dataRows = rows.slice(headerIdx + 1).filter(r => r && r.length > 2);

      const col = {};
      headers.forEach((h, i) => {
        if (h.includes("anno")) col.anno = i;
        else if (h === "n. lancio" || (h.includes("n.") && h.includes("lancio"))) col.num = i;
        else if (h.includes("ean")) col.ean = i;
        else if (h === "cod. editore" || (h.includes("cod") && h.includes("edit"))) col.cod_ed = i;
        else if ((h.includes("descr") && h.includes("edit")) || h === "editore") col.editore = i;
        else if (h === "titolo") col.titolo = i;
        else if (h === "autore") col.autore = i;
        else if (h === "prezzo") col.prezzo = i;
        else if (h.includes("prenotato")) col.prenotato = i;
      });
      if (col.ean === undefined) throw new Error("Colonna EAN non trovata");

      if (mode === "iscrizione") {
        const payload = dataRows.map(r => {
          const ean = String(r[col.ean] || "").replace(/\.0$/, "").trim();
          if (ean.length < 10) return null;
          return {
            anno_lancio: parseInt(r[col.anno]) || new Date().getFullYear(),
            num_lancio: parseInt(r[col.num]) || 0,
            ean,
            codice_editore: String(r[col.cod_ed] || "").trim(),
            editore: String(r[col.editore] || "").trim(),
            titolo: String(r[col.titolo] || "").trim(),
            autore: String(r[col.autore] || "").trim(),
            prezzo: parseFloat(String(r[col.prezzo] || "0").replace(",", ".")) || 0,
            prenotato_iscrizione: parseInt(String(r[col.prenotato] || "0").replace(/\./g, "")) || 0,
            prenotato_trasmesso: parseInt(String(r[col.prenotato] || "0").replace(/\./g, "")) || 0,
          };
        }).filter(Boolean);
        if (payload.length === 0) throw new Error("Nessun EAN valido");

        // Se un titolo (EAN) era già presente in un lancio precedente, lo "spostiamo":
        // viene eliminato dal lancio vecchio e resta solo nell'ultimo lancio caricato.
        const gruppiPerLancio = {};
        payload.forEach(p => {
          const key = `${p.anno_lancio}|${p.num_lancio}`;
          if (!gruppiPerLancio[key]) gruppiPerLancio[key] = { anno: p.anno_lancio, num: p.num_lancio, eans: [] };
          gruppiPerLancio[key].eans.push(p.ean);
        });
        for (const { anno, num, eans } of Object.values(gruppiPerLancio)) {
          for (let i = 0; i < eans.length; i += 200) {
            const eanBatch = eans.slice(i, i + 200);
            const eanList = eanBatch.join(",");
            const delResp = await fetch(
              `${SUPABASE_URL}/rest/v1/lanci_settimanali?ean=in.(${eanList})&not.and=(anno_lancio.eq.${anno},num_lancio.eq.${num})`,
              {
                method: "DELETE",
                headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`, "Content-Type": "application/json", "Prefer": "return=minimal" },
              }
            );
            if (!delResp.ok) console.warn("Pulizia lanci precedenti non riuscita:", await delResp.text());
          }

          // Se il file ricaricato per questo stesso lancio non contiene più
          // alcuni EAN presenti in precedenza, quei titoli sono usciti dal
          // lancio: li rimuoviamo (sovrascrittura completa per quel lancio).
          const cleanupResp = await fetch(`${SUPABASE_URL}/rest/v1/rpc/cleanup_lancio_orphans`, {
            method: "POST",
            headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`, "Content-Type": "application/json" },
            body: JSON.stringify({ p_anno: anno, p_num: num, p_eans: eans }),
          });
          if (!cleanupResp.ok) console.warn("Pulizia titoli usciti dal lancio non riuscita:", await cleanupResp.text());
        }

        for (let i = 0; i < payload.length; i += 500) {
          const batch = payload.slice(i, i + 500);
          const r = await fetch(`${SUPABASE_URL}/rest/v1/rpc/upsert_lanci`, {
  method: "POST",
  headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`, "Content-Type": "application/json" },
  body: JSON.stringify({ payload: batch }),
});
if (!r.ok) throw new Error(await r.text());
        }
        showToast(`Lancio: ${payload.length} titoli caricati (lancio ${payload[0]?.num_lancio})`);
        await checkAnticipiNotificati(payload.map(p => p.ean));
      } else {
        let updated = 0;
        for (const r of dataRows) {
          const ean = String(r[col.ean] || "").replace(/\.0$/, "").trim();
          const qta = parseInt(String(r[col.prenotato] || "0").replace(/\./g, "")) || 0;
          const anno = parseInt(r[col.anno]) || filterAnno;
          const num = parseInt(r[col.num]) || filterLancio[0];
          if (ean.length < 10) continue;
          const resp = await fetch(`${SUPABASE_URL}/rest/v1/lanci_settimanali?anno_lancio=eq.${anno}&num_lancio=eq.${num}&ean=eq.${encodeURIComponent(ean)}`, {
            method: "PATCH",
            headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`, "Content-Type": "application/json", "Prefer": "return=minimal" },
            body: JSON.stringify({ prenotato_trasmesso: qta, trasmesso_confermato: true }),
          });
          if (resp.ok) updated++;
        }
        showToast(`Trasmesso: ${updated} titoli aggiornati`);
      }
      await loadData();
    } catch (err) {
      showToast(err.message, "err");
    }
    setUploading(false);
    setUploadMode(null);
    e.target.value = "";
  };

  // Export Excel
  const exportExcel = () => {
    const XLSX = window.XLSX;
    const headers = ["LANCIO","CEDOLA","EAN","TITOLO","AUTORE","COD.EDITORE","EDITORE","PREZZO","F.G.","P.O. MELI","AMAZON","TOT.TEORICO","FG VS TOT","GIORNO USCITA"];
    const rows = dataFiltrata.map(r => [
      r.num_lancio, r.cedole.join(", "), r.ean, r.titolo, r.autore, r.codice_editore, r.editore, r.prezzo,
      r.pren_fine_giro, r.prenotato_trasmesso ?? "", r.pren_amazon, r.teorico, r.delta_portale, r.giorno_uscita
    ]);
    // Riepilogo editori
    const rH = ["EDITORE","TITOLI","LANCIATE","TRASMESSE","FINE GIRO","AMAZON","VALORE"];
    const rR = Object.entries(kpi.byEditore).sort((a, b) => b[1].lanciate - a[1].lanciate).map(([ed, v]) => [
      ed, v.titoli, v.lanciate, v.trasmesse, v.fineGiro, v.amazon, Math.round(v.valore)
    ]);
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    const wsR = XLSX.utils.aoa_to_sheet([rH, ...rR]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `Lanci ${filterLancio.join("-")}`);
    XLSX.utils.book_append_sheet(wb, wsR, "Riepilogo Editori");
    XLSX.writeFile(wb, `Lancio_${filterLancio.join("-")}_${filterAnno}.xlsx`);
  };

  if (loading) return <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", color: T.textMid }}>Caricamento lanci...</div>;

  if (data.length === 0) return (
    <div style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 20, padding: 40 }}>
      <div style={{ fontSize: "48px", opacity: 0.3 }}>🚀</div>
      <div style={{ color: T.textMid, fontSize: "14px", textAlign: "center", maxWidth: 400 }}>Nessun lancio caricato. Carica il file del lancio.</div>
      <label style={{ ...css.btn("accent"), cursor: "pointer", padding: "10px 24px", fontSize: "13px" }}>
        ↑ Carica primo lancio
        <input type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={(e) => handleUpload(e, "iscrizione")} />
      </label>
    </div>
  );

  const editoriTop = Object.entries(kpi.byEditore).sort((a, b) => b[1].lanciate - a[1].lanciate).slice(0, 8);
  const maxLanciate = editoriTop.length > 0 ? editoriTop[0][1].lanciate : 1;

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%", overflow: "hidden" }}>
      {/* TOOLBAR */}
      <div style={{ padding: "14px 20px", borderBottom: `1px solid ${T.border}`, display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
        <select style={{ ...css.input, fontSize: "13px", fontWeight: "600", color: T.accent }} value={filterAnno || ""} onChange={e => { setFilterAnno(Number(e.target.value)); setFilterLancio([]); }}>
          {anniDisponibili.map(a => <option key={a} value={a}>{a}</option>)}
        </select>
        <SearchableMultiSelect values={filterLancio.map(String)} onChange={v => setFilterLancio(v.map(Number))} options={lanciPerAnno.map(String)} renderOption={v => `Lancio ${v}`} placeholder="Seleziona lancio" width={170} />
        <input style={{ ...css.input, width: 180 }} placeholder="Cerca EAN / titolo..." value={search} onChange={e => setSearch(e.target.value)} />
        <SearchableMultiSelect values={filterAccount} onChange={setFilterAccount} options={accountsDisponibili} placeholder="Tutti gli account" width={160} />
        {/* Riepilogo giorni */}
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <span style={{ ...css.tag(T.accent), cursor: "pointer", opacity: filterGiorno === "venerdì" ? 0.4 : 1 }} onClick={() => setFilterGiorno(g => g === "martedì" ? "tutti" : "martedì")}>MAR {kpi.martedi}</span>
          <span style={{ ...css.tag(T.green), cursor: "pointer", opacity: filterGiorno === "martedì" ? 0.4 : 1 }} onClick={() => setFilterGiorno(g => g === "venerdì" ? "tutti" : "venerdì")}>VEN {kpi.venerdi}</span>
          {filterGiorno !== "tutti" && (
            <button style={{ ...css.btn(), fontSize: "11px", padding: "2px 8px" }} onClick={() => setFilterGiorno("tutti")}>✕</button>
          )}
        </div>
        <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
          <label style={{ ...css.btn("accent"), cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 4 }}>
            {uploading && uploadMode === "iscrizione" ? "..." : "Upload File Lancio"}
            <input type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={(e) => handleUpload(e, "iscrizione")} disabled={uploading} />
          </label>
          <button style={css.btn()} onClick={exportExcel}>Download Excel</button>
        </div>
      </div>

      {/* KPI */}
      <div style={{ padding: "16px 20px", borderBottom: `1px solid ${T.border}` }}>
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 16 }}>
          <KpiCard label="Titoli al lancio" value={kpi.tot} color={T.text} />
          <KpiCard label="Copie lanciate" value={kpi.copieLanciate.toLocaleString("it")} color={T.accent} sub={`€ ${kpi.valoreLanciate.toLocaleString("it", { maximumFractionDigits: 0 })}`} />

          <KpiCard label="Fine Giro" value={kpi.totFineGiro.toLocaleString("it")} color={T.purple} sub={`€ ${kpi.valoreFineGiro.toLocaleString("it", { maximumFractionDigits: 0 })}`} />
          <KpiCard label="Amazon" value={kpi.totAmazon.toLocaleString("it")} color="#e8a838" sub={`€ ${kpi.valoreAmazon.toLocaleString("it", { maximumFractionDigits: 0 })}`} />
          <KpiCard label="Teorico totale" value={kpi.totTeorico.toLocaleString("it")} color={T.text} sub={`€ ${kpi.valoreTeorico.toLocaleString("it", { maximumFractionDigits: 0 })}`} />
        </div>

        {/* Mini chart editori */}
        <div style={{ color: T.textMid, fontSize: "10px", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 10 }}>Top editori</div>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {editoriTop.map(([ed, v]) => (
            <div key={ed} style={{ background: T.bg, border: `1px solid ${T.border}`, borderRadius: 4, padding: "8px 12px", minWidth: 150, flex: "1 1 150px", maxWidth: 220 }}>
              <div style={{ fontSize: "11px", fontWeight: "600", color: T.text, marginBottom: 4, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{ed}</div>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 3 }}>
                <div style={{ flex: 1, height: 4, background: T.borderHi, borderRadius: 2, overflow: "hidden" }}>
                  <div style={{ width: `${Math.round(v.lanciate / maxLanciate * 100)}%`, height: "100%", background: T.accent }} />
                </div>
                <span style={{ fontSize: "11px", fontWeight: "700", color: T.accent, minWidth: 40, textAlign: "right" }}>{v.lanciate.toLocaleString("it")}</span>
              </div>
              {v.amazon > 0 && (
                <div style={{ fontSize: "9px", color: "#e8a838" }}>🅰 {v.amazon.toLocaleString("it")}</div>
              )}
              <div style={{ fontSize: "9px", color: T.textDim, marginTop: 2 }}>{v.titoli} titoli · € {Math.round(v.valore).toLocaleString("it")}</div>
            </div>
          ))}
        </div>
      </div>

      {/* TABELLA */}
      <div style={{ flex: 1, overflowY: "auto", overflowX: "auto" }}>
        <table style={css.table}>
          <thead>
            <tr>
              <th style={css.th}>EAN</th>
              <th style={css.th}>Cod.Ed.</th>
              <th style={{ ...css.th, cursor: "pointer" }} onClick={() => toggleSort("editore")}>Editore{sortIcon("editore")}</th>
              <th style={css.th}>Titolo</th>
              <th style={css.th}>Autore</th>
              <th style={css.th}>€</th>
              <th style={css.th}>Cedole</th>
              <th style={{ ...css.th, cursor: "pointer" }} onClick={() => toggleSort("pren_fine_giro")}>F.G.{sortIcon("pren_fine_giro")}</th>
              <th style={{ ...css.th, cursor: "pointer" }} onClick={() => toggleSort("prenotato_trasmesso")}>P.O. Meli{sortIcon("prenotato_trasmesso")}</th>
              <th style={{ ...css.th, cursor: "pointer", color: "#e8a838" }} onClick={() => toggleSort("pren_amazon")}>Amazon{sortIcon("pren_amazon")}</th>
              <th style={{ ...css.th, cursor: "pointer" }} onClick={() => toggleSort("teorico")}>Tot. Teorico{sortIcon("teorico")}</th>
              <th style={{ ...css.th, cursor: "pointer" }} onClick={() => toggleSort("delta_portale")} title="Fine Giro vs Tot. Teorico">FG vs Tot.Teorico{sortIcon("delta_portale")}</th>
              <th style={css.th}>Uscita</th>
            </tr>
          </thead>
          <tbody>
            {dataFiltrata.map((r, i) => (
              <tr key={r.id || i} style={{
                background: r.prenotato_trasmesso === 0
                  ? T.red + "22"
                  : (i % 2 === 0 ? "transparent" : T.surface + "66"),
                ...(r.prenotato_trasmesso === 0 ? { boxShadow: `inset 3px 0 0 ${T.red}` } : {})
              }}>
                <td style={{ ...css.td, fontFamily: "monospace", fontSize: "11px", color: T.textMid }}>{r.ean}</td>
                <td style={{ ...css.td, color: T.textMid, fontSize: "11px" }}>{r.codice_editore}</td>
                <td style={{ ...css.td, color: T.accent, fontWeight: "600", whiteSpace: "nowrap", maxWidth: 140, overflow: "hidden", textOverflow: "ellipsis" }}>{r.editore}</td>
                <td style={{ ...css.td, maxWidth: 200 }}><div style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontWeight: "600" }}>{r.titolo}</div></td>
                <td style={{ ...css.td, color: T.textMid, maxWidth: 130, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.autore}</td>
                <td style={{ ...css.td, whiteSpace: "nowrap" }}>€ {(r.prezzo || 0).toFixed(2)}</td>
                <td style={{ ...css.td, fontSize: "10px", color: T.textMid, maxWidth: 120 }}>
                  {r.is_live_cedole ? (
                    r.cedole.map((c, ci) => <span key={ci}>{ci > 0 && <br />}{c}</span>)
                  ) : editCell?.id === r.id && editCell?.field === "cedole" ? (
                    <div style={{ display: "flex", gap: 3, alignItems: "center" }}>
                      <input style={{ ...css.input, width: 90, padding: "2px 5px", fontSize: "10px" }} value={editCell.value} autoFocus
                        onChange={e => setEditCell(p => ({ ...p, value: e.target.value }))}
                        onKeyDown={e => { if (e.key === "Enter") saveManualCell(r.id, "cedole", editCell.value); if (e.key === "Escape") setEditCell(null); }}
                        placeholder="es. 26_1_1A" />
                      <button style={{ ...css.btn("accent"), padding: "1px 5px", fontSize: "10px" }} onClick={() => saveManualCell(r.id, "cedole", editCell.value)}>✓</button>
                    </div>
                  ) : (
                    <div style={{ cursor: "pointer", display: "flex", alignItems: "center", gap: 4 }}
                      onClick={() => setEditCell({ id: r.id, field: "cedole", value: r.cedole_manual || r.cedole.join(", ") || "" })}>
                      {r.cedole.length > 0 ? r.cedole.map((c, ci) => <span key={ci} style={{ color: r.is_manual_cedole ? "#e8a838" : T.textMid }}>{ci > 0 && ", "}{c}</span>) : <span style={{ color: T.textDim }}>—</span>}
                      <span style={{ color: T.accent, fontSize: "10px" }}>✎</span>
                    </div>
                  )}
                </td>
                <td style={{ ...css.td, fontWeight: "600" }}>
                  {r.is_live_fg ? (
                    <span style={{ color: T.purple }}>{r.pren_fine_giro.toLocaleString("it")}</span>
                  ) : editCell?.id === r.id && editCell?.field === "fg" ? (
                    <div style={{ display: "flex", gap: 3, alignItems: "center" }}>
                      <input type="number" style={{ ...css.input, width: 70, padding: "2px 5px", fontSize: "11px" }} value={editCell.value} autoFocus
                        onChange={e => setEditCell(p => ({ ...p, value: e.target.value }))}
                        onKeyDown={e => { if (e.key === "Enter") saveManualCell(r.id, "fg", editCell.value); if (e.key === "Escape") setEditCell(null); }} />
                      <button style={{ ...css.btn("accent"), padding: "1px 5px", fontSize: "10px" }} onClick={() => saveManualCell(r.id, "fg", editCell.value)}>✓</button>
                    </div>
                  ) : (
                    <div style={{ cursor: "pointer", display: "flex", alignItems: "center", gap: 4 }}
                      onClick={() => setEditCell({ id: r.id, field: "fg", value: String(r.pren_fine_giro_manual || r.pren_fine_giro || "") })}>
                      <span style={{ color: r.pren_fine_giro > 0 ? (r.is_manual_fg ? "#e8a838" : T.purple) : T.textDim }}>{r.pren_fine_giro > 0 ? r.pren_fine_giro.toLocaleString("it") : "—"}</span>
                      <span style={{ color: T.accent, fontSize: "10px" }}>✎</span>
                    </div>
                  )}
                </td>
                <td style={{ ...css.td, fontWeight: "600", color: r.prenotato_trasmesso === 0 ? T.red : "inherit" }}>
                  {r.prenotato_trasmesso != null ? r.prenotato_trasmesso.toLocaleString("it") : "—"}
                </td>
                <td style={{ ...css.td, fontWeight: "600" }}>
                  {r.is_live_amazon ? (
                    <span style={{ color: "#e8a838" }}>{r.pren_amazon.toLocaleString("it")}</span>
                  ) : editCell?.id === r.id && editCell?.field === "amazon" ? (
                    <div style={{ display: "flex", gap: 3, alignItems: "center" }}>
                      <input type="number" style={{ ...css.input, width: 70, padding: "2px 5px", fontSize: "11px" }} value={editCell.value} autoFocus
                        onChange={e => setEditCell(p => ({ ...p, value: e.target.value }))}
                        onKeyDown={e => { if (e.key === "Enter") saveManualCell(r.id, "amazon", editCell.value); if (e.key === "Escape") setEditCell(null); }} />
                      <button style={{ ...css.btn("accent"), padding: "1px 5px", fontSize: "10px" }} onClick={() => saveManualCell(r.id, "amazon", editCell.value)}>✓</button>
                    </div>
                  ) : (
                    <div style={{ cursor: "pointer", display: "flex", alignItems: "center", gap: 4 }}
                      onClick={() => setEditCell({ id: r.id, field: "amazon", value: String(r.pren_amazon_manual || r.pren_amazon || "") })}>
                      <span style={{ color: r.pren_amazon > 0 ? (r.is_manual_amazon ? "#e8a838" : "#e8a838") : T.textDim }}>{r.pren_amazon > 0 ? r.pren_amazon.toLocaleString("it") : "—"}</span>
                      <span style={{ color: T.accent, fontSize: "10px" }}>✎</span>
                    </div>
                  )}
                </td>
                <td style={{ ...css.td, fontWeight: "700", color: r.teorico >= 2000 ? T.green : T.text }}>
                  {r.teorico > 0 ? r.teorico.toLocaleString("it") : "—"}
                </td>
                <td style={{ ...css.td, fontWeight: "700", color: r.delta_portale > 0 ? T.green : r.delta_portale < 0 ? T.red : T.textMid, textAlign: "right" }}>
                  {r.prenotato_trasmesso != null || r.pren_fine_giro > 0
                    ? (r.delta_portale > 0 ? "+" : "") + r.delta_portale.toLocaleString("it")
                    : "—"}
                </td>
                <td style={css.td}>
                  <select
                    value={r.giorno_uscita}
                    onChange={e => saveGiornoUscita(r.id, e.target.value === r.giorno_calcolato ? null : e.target.value)}
                    style={{
                      ...css.input,
                      width: 90,
                      padding: "3px 6px",
                      fontSize: "11px",
                      fontWeight: "700",
                      color: r.giorno_uscita === "martedì" ? T.accent : T.green,
                      borderColor: r.is_override ? "#e8a838" : T.border,
                    }}
                  >
                    <option value="martedì">Martedì</option>
                    <option value="venerdì">Venerdì</option>
                  </select>
                  {r.is_override && <span style={{ color: "#e8a838", fontSize: "9px", marginLeft: 4 }}>✎</span>}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        {dataFiltrata.length === 0 && (
          <div style={{ padding: 40, textAlign: "center", color: T.textDim }}>Nessun titolo per questo lancio.</div>
        )}
      </div>

      {toast && (
        <div style={{ position: "fixed", bottom: 24, left: "50%", transform: "translateX(-50%)", background: toast.type === "err" ? "#4a1a2a" : "#1a3a2a", border: `1px solid ${toast.type === "err" ? T.red : T.green}`, color: toast.type === "err" ? T.red : T.green, borderRadius: 6, padding: "8px 20px", fontSize: "12px", zIndex: 999, boxShadow: "0 4px 20px #0008" }}>
          {toast.msg}
        </div>
      )}

      {/* POPUP ANTICIPI LANCIO NOTIFICATI */}
      {anticipiPopup && (
        <div style={{ position: "fixed", inset: 0, background: "#000a", zIndex: 300, display: "flex", alignItems: "center", justifyContent: "center" }} onClick={() => setAnticipiPopup(null)}>
          <div style={{ background: T.surface, border: `1px solid #e8a838`, borderRadius: 6, padding: 28, width: 520, maxHeight: "80vh", overflowY: "auto" }} onClick={e => e.stopPropagation()}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 6 }}>
              <span style={{ fontSize: "22px" }}>🔔</span>
              <span style={{ color: "#e8a838", fontWeight: "700", fontSize: "14px" }}>ANTICIPI LANCIO SBLOCCATI</span>
            </div>
            <div style={{ color: T.textMid, fontSize: "12px", marginBottom: 16 }}>
              {anticipiPopup.length === 1
                ? "1 anticipo lancio è ora collegato a un titolo di questo lancio ed è pronto per essere gestito."
                : `${anticipiPopup.length} anticipi lancio sono ora collegati a titoli di questo lancio e pronti per essere gestiti.`}
            </div>
            <div style={{ border: `1px solid ${T.border}`, borderRadius: 4, overflow: "hidden", marginBottom: 20 }}>
              <table style={css.table}>
                <thead>
                  <tr>
                    <th style={css.th}>Cliente</th>
                    <th style={css.th}>EAN</th>
                    <th style={css.th}>Qtà</th>
                  </tr>
                </thead>
                <tbody>
                  {anticipiPopup.map((r, i) => (
                    <tr key={r.id} style={{ background: i % 2 === 0 ? "transparent" : T.bg }}>
                      <td style={{ ...css.td, fontWeight: "600" }}>{r.codice_cliente}</td>
                      <td style={{ ...css.td, fontFamily: "monospace", fontSize: "11px", color: T.textMid }}>{r.ean}</td>
                      <td style={css.td}>{r.quantita}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
              <button style={css.btn()} onClick={() => setAnticipiPopup(null)}>Chiudi</button>
              {onNavigateAnticipi && (
                <button style={{ ...css.btn("accent") }} onClick={() => { setAnticipiPopup(null); onNavigateAnticipi(); }}>Vai ad Anticipi Lancio →</button>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── MODULO VERIFICA LANCI AMAZON ──────────────────────────────────────────
function ModuloVerificaLanciAmazon({ token, titoli, prenotato, canali }) {
  const [data, setData] = useState([]);
  const [verificaData, setVerificaData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [toast, setToast] = useState(null);
  const [filterAnno, setFilterAnno] = useState(null);
  const [filterLancio, setFilterLancio] = useState(null); // singolo lancio, sempre
  const [search, setSearch] = useState("");
  const [editCellVA, setEditCellVA] = useState(null); // { ean, field, value }
  const [showRiepilogoEditori, setShowRiepilogoEditori] = useState(false);

  const showToast = (msg, type = "ok") => { setToast({ msg, type }); setTimeout(() => setToast(null), 4000); };

  const loadData = useCallback(async () => {
    if (!token) return;
    setLoading(true);
    const d = await sbFetch("lanci_settimanali?select=*&order=anno_lancio.desc,num_lancio.desc,editore.asc,titolo.asc", token);
    if (Array.isArray(d)) setData(d);
    setLoading(false);
  }, [token]);
  useEffect(() => { loadData(); }, [loadData]);

  const loadVerifica = useCallback(async () => {
    if (!token) return;
    const d = await sbFetch("verifica_amazon?select=*", token);
    if (Array.isArray(d)) setVerificaData(d);
  }, [token]);
  useEffect(() => { loadVerifica(); }, [loadVerifica]);

  // Anni e lanci disponibili (si aggiornano da soli se i titoli slittano da un lancio all'altro)
  const anniDisponibili = useMemo(() => [...new Set(data.map(r => r.anno_lancio))].sort((a, b) => b - a), [data]);
  const lanciPerAnno = useMemo(() => {
    const anno = filterAnno || anniDisponibili[0];
    if (!anno) return [];
    return [...new Set(data.filter(r => r.anno_lancio === anno).map(r => r.num_lancio))].sort((a, b) => b - a);
  }, [data, filterAnno, anniDisponibili]);

  useEffect(() => { if (!filterAnno && anniDisponibili.length > 0) setFilterAnno(anniDisponibili[0]); }, [anniDisponibili]);
  useEffect(() => { if (lanciPerAnno.length > 0 && (filterLancio == null || !lanciPerAnno.includes(filterLancio))) setFilterLancio(lanciPerAnno[0]); }, [lanciPerAnno]);

  const amazonCodice = "AMAZON";

  // Mappa prenotato per EAN (da fine giro) — per canale, per calcolare Amazon Cedola live
  const prenByEanCanale = useMemo(() => {
    const map = {};
    prenotato.forEach(p => {
      const t = titoli.find(t => t.id === p.titolo_id);
      if (!t || !t.ean) return;
      if (!map[t.ean]) map[t.ean] = {};
      const c = canali.find(c => c.id === p.canale_id);
      if (c) map[t.ean][c.codice] = (map[t.ean][c.codice] || 0) + p.quantita;
    });
    return map;
  }, [prenotato, titoli, canali]);

  // Dati del lancio selezionato, arricchiti con Amazon Cedola/Teorico live
  const dataArricchita = useMemo(() => {
    const anno = filterAnno || anniDisponibili[0];
    return data
      .filter(r => r.anno_lancio === anno && r.num_lancio === filterLancio)
      .map(r => {
        const prenCanali = prenByEanCanale[r.ean] || {};
        const prenAmazonLive = prenCanali[amazonCodice] || 0;
        const haLiveAmazon = prenAmazonLive > 0;
        const prenAmazon = haLiveAmazon ? prenAmazonLive : (r.pren_amazon_manual || 0);
        const teorico = (r.prenotato_trasmesso ?? r.prenotato_iscrizione ?? 0) + prenAmazon;
        return { ...r, pren_amazon: prenAmazon, teorico };
      });
  }, [data, filterAnno, filterLancio, anniDisponibili, prenByEanCanale]);

  const dataFiltrata = useMemo(() => {
    let result = [...dataArricchita];
    if (search) {
      const q = search.toLowerCase();
      result = result.filter(r => r.ean?.toLowerCase().includes(q) || r.titolo?.toLowerCase().includes(q) || r.autore?.toLowerCase().includes(q) || r.editore?.toLowerCase().includes(q));
    }
    return result;
  }, [dataArricchita, search]);

  // ── Verifica Amazon: mappa record per EAN (lancio/anno correnti) + calcolo formule ──
  const verificaByEan = useMemo(() => {
    const anno = filterAnno || anniDisponibili[0];
    const map = {};
    verificaData.forEach(v => {
      if (v.anno_lancio === anno && v.num_lancio === filterLancio) {
        map[v.ean] = v;
      }
    });
    return map;
  }, [verificaData, filterAnno, filterLancio, anniDisponibili]);

  const dataVerifica = useMemo(() => {
    return dataFiltrata.map(r => {
      const v = verificaByEan[r.ean] || {};
      const F = r.pren_amazon || 0;          // AMAZON IN CEDOLA (attuale, da fine giro)
      const E = r.teorico || 0;              // TOTALE teorico
      const G = v.proposta_amaz ?? null;      // Proposta fatta ad Amazon
      const I = v.proposta_pde ?? null;       // Proposta PDE — resta vuota finché non caricata manualmente/da upload dedicato
      const J = v.copie ?? null;              // Copie effettivamente inserite da Amazon
      const K = v.evaso ?? null;
      const L = v.inevaso ?? null;
      const Pr = v.prenotato ?? null;
      const Im = v.impegnato ?? null;
      const U = v.preorder ?? 0;
      const H = (G != null && F) ? (G - F) / F : null;
      const M = (J != null && L != null) ? J - L : (J != null ? J : null); // Netto confermato
      const N = M != null ? M - F : null;
      const O = (M != null && G != null) ? M - G : null;
      const P = (M != null && I != null) ? M - I : null;
      const Q = (N != null && F) ? N / F : null;
      const R = (O != null && G) ? O / G : null;
      const S = (P != null && I) ? P / I : null;
      const V = M != null ? M - U : null;
      const W = (V != null && M) ? U / M : null;
      return { ...r, va: v, vF: F, vE: E, vG: G, vH: H, vI: I, vJ: J, vK: K, vL: L, vPren: Pr, vImp: Im, vM: M, vN: N, vO: O, vP: P, vQ: Q, vR: R, vS: S, vNote: v.note || "", vU: U, vV: V, vW: W, vX: !!v.richiesta_rifornimento, vY: !!v.rottura_stock, haConferma: M != null, vAsin: v.asin || "", vPropostaVendor: v.proposta_vendor ?? "", vFiltroReti: v.filtro_reti || "", vPreordiniVendor: v.preordini_vendor ?? "" };
    });
  }, [dataFiltrata, verificaByEan]);

  const kpiVerifica = useMemo(() => {
    const totProposto = dataVerifica.reduce((s, r) => s + (r.vF || 0), 0);
    const totConfermato = dataVerifica.reduce((s, r) => s + (r.vM || 0), 0);
    const valoreProposto = dataVerifica.reduce((s, r) => s + (r.prezzo || 0) * (r.vF || 0), 0);
    const valoreConfermato = dataVerifica.reduce((s, r) => s + (r.prezzo || 0) * (r.vM || 0), 0);
    const nConfermati = dataVerifica.filter(r => r.haConferma).length;
    const nInAttesa = dataVerifica.length - nConfermati;
    const scostamento = totConfermato - totProposto;
    const valoreScostamento = valoreConfermato - valoreProposto;
    return { totProposto, totConfermato, valoreProposto, valoreConfermato, nConfermati, nInAttesa, scostamento, valoreScostamento };
  }, [dataVerifica]);

  // Riepilogo Verifica Amazon per editore (proposto/confermato/scostamento)
  const riepilogoEditoriVerifica = useMemo(() => {
    const byEd = {};
    dataVerifica.forEach(r => {
      if (!byEd[r.editore]) byEd[r.editore] = { editore: r.editore, titoli: 0, proposto: 0, confermato: 0, confermati: 0, inAttesa: 0, valore: 0, valoreProposto: 0 };
      const e = byEd[r.editore];
      e.titoli++;
      e.proposto += r.vF || 0;
      e.confermato += r.vM || 0;
      e.valore += (r.prezzo || 0) * (r.vM || 0);
      e.valoreProposto += (r.prezzo || 0) * (r.vF || 0);
      if (r.haConferma) e.confermati++; else e.inAttesa++;
    });
    return Object.values(byEd)
      .map(e => ({ ...e, scostamento: e.confermato - e.proposto }))
      .sort((a, b) => b.confermato - a.confermato);
  }, [dataVerifica]);

  // ── Upload mail "Differenza prenotazioni Amazon" (Messaggerie) ──
  const [showMailUpload, setShowMailUpload] = useState(false);
  const [mailPasteText, setMailPasteText] = useState("");
  const [mailParseResult, setMailParseResult] = useState(null);
  const [mailProcessing, setMailProcessing] = useState(false);
  const [mailLancioInfo, setMailLancioInfo] = useState(null);
  const [mailDragOver, setMailDragOver] = useState(false);

  // Estrae righe EAN/Titolo/Prezzo/Proposto/Restituito dal testo della mail Messaggerie
  const parseMessaggerieMail = (text) => {
    const pattern = /(\d{9,14})\s*(?:\r?\n\s*)+([^\r\n\d][^\r\n]*?)\s*(?:\r?\n\s*)+(\d{1,4}[.,]\d{2})\s*€\s*(?:\r?\n\s*)+(-?\d+)\s*(?:\r?\n\s*)+(-?\d+)/g;
    const rows = [];
    let m;
    while ((m = pattern.exec(text)) !== null) {
      rows.push({
        ean: m[1].trim(),
        titolo: m[2].trim(),
        prezzo: parseFloat(m[3].replace(",", ".")),
        proposto: parseInt(m[4], 10),
        restituito: parseInt(m[5], 10),
      });
    }
    return rows;
  };

  // Estrae il numero di lancio (es. "LANCIO NR. 29.0 2026" / "lancio n. 29.0") dal testo mail
  const parseLancioFromMail = (text) => {
    const m = text.match(/lancio\s*(?:nr\.?|n\.)?\s*(\d{1,3})[.,](\d)(?:\D{0,10}(\d{4}))?/i);
    if (!m) return null;
    const numLancio = parseInt(m[1], 10) * 10 + parseInt(m[2], 10);
    const anno = m[3] ? parseInt(m[3], 10) : null;
    return { numLancio, anno };
  };

  const runMailParse = (text) => {
    let rows = parseMessaggerieMail(text);
    if (rows.length === 0) showToast("Nessuna riga riconosciuta nel testo/file caricato", "err");
    setMailLancioInfo(parseLancioFromMail(text));
    setMailParseResult(rows);
  };

  const processMailFile = async (file) => {
    if (!file) return;
    const buf = await file.arrayBuffer();
    let text = "";
    if (/\.msg$/i.test(file.name)) {
      // I file .msg (Outlook, formato OLE) contengono il corpo testo in UTF-16LE
      text = new TextDecoder("utf-16le").decode(buf);
      if (parseMessaggerieMail(text).length === 0) text = new TextDecoder("latin1").decode(buf);
    } else {
      text = new TextDecoder("utf-8").decode(buf);
    }
    runMailParse(text);
  };

  const handleMailFile = async (e) => {
    const file = e.target.files?.[0];
    await processMailFile(file);
    e.target.value = "";
  };

  const handleMailDrop = async (e) => {
    e.preventDefault();
    e.stopPropagation();
    setMailDragOver(false);
    const file = e.dataTransfer.files?.[0];
    await processMailFile(file);
  };

  const mailByEan = useMemo(() => {
    const map = {};
    (mailParseResult || []).forEach(r => { map[r.ean] = r; });
    return map;
  }, [mailParseResult]);

  const mailPreview = useMemo(() => {
    if (!mailParseResult) return null;
    const daMail = mailParseResult.length;
    const autoConfermati = dataVerifica.filter(r => !mailByEan[r.ean]).length;
    return { daMail, autoConfermati };
  }, [mailParseResult, mailByEan, dataVerifica]);

  // Blocco di sicurezza: la mail deve riferirsi esattamente al lancio/anno selezionati in app
  const mailLancioMismatch = useMemo(() => {
    if (!mailParseResult) return null;
    const annoSel = filterAnno || anniDisponibili[0];
    if (!mailLancioInfo) {
      return { tipo: "non_rilevato", msg: "Non sono riuscito a capire a quale lancio si riferisce la mail. Verifica manualmente prima di procedere." };
    }
    if (filterLancio !== mailLancioInfo.numLancio || (mailLancioInfo.anno && annoSel !== mailLancioInfo.anno)) {
      return { tipo: "mismatch", msg: `La mail si riferisce al lancio ${mailLancioInfo.numLancio}${mailLancioInfo.anno ? "/" + mailLancioInfo.anno : ""}, ma qui hai selezionato il lancio ${filterLancio}/${annoSel}. Seleziona il lancio corretto prima di continuare.` };
    }
    return null;
  }, [mailParseResult, mailLancioInfo, filterLancio, filterAnno, anniDisponibili]);

  // ── Upload "Ordini Amazon" (file Dettaglio cliente: aggrega Copie/Prenotato/Impegnato/Evaso per EAN) ──
  const [showOrdiniUpload, setShowOrdiniUpload] = useState(false);
  const [ordiniParseResult, setOrdiniParseResult] = useState(null);
  const [ordiniProcessing, setOrdiniProcessing] = useState(false);
  const [ordiniDragOver, setOrdiniDragOver] = useState(false);

  const parseNumIt = (s) => {
    if (s == null || s === "") return 0;
    const n = parseInt(String(s).trim().replace(/\./g, ""), 10);
    return isNaN(n) ? 0 : n;
  };

  // Estrae e aggrega per EAN le colonne Copie/Prenotato/Impegnato/Evaso dal file "Dettaglio cliente" (righe = righe ordine, un EAN può ripetersi su più righe/ordini)
  const parseOrdiniAmazon = (text) => {
    const lines = text.replace(/^\uFEFF/, "").split(/\r?\n/).filter(l => l.trim() !== "");
    if (lines.length < 2) return [];
    const agg = {};
    for (let i = 1; i < lines.length; i++) {
      const cols = lines[i].split("\t");
      if (cols.length < 13) continue;
      const ordine = (cols[0] || "").trim();
      if (!ordine || /^totale/i.test(ordine)) continue; // salta riga "Totale complessivo"
      const ean = (cols[2] || "").trim();
      if (!/^\d{9,14}$/.test(ean)) continue;
      const titolo = (cols[3] || "").trim();
      if (!agg[ean]) agg[ean] = { ean, titolo, copie: 0, prenotato: 0, impegnato: 0, evaso: 0 };
      agg[ean].copie += parseNumIt(cols[9]);
      agg[ean].prenotato += parseNumIt(cols[10]);
      agg[ean].impegnato += parseNumIt(cols[11]);
      agg[ean].evaso += parseNumIt(cols[12]);
      if (titolo) agg[ean].titolo = titolo;
    }
    return Object.values(agg);
  };

  const processOrdiniFile = async (file) => {
    if (!file) return;
    const buf = await file.arrayBuffer();
    let text = new TextDecoder("utf-16le").decode(buf);
    let rows = parseOrdiniAmazon(text);
    if (rows.length === 0) {
      text = new TextDecoder("utf-8").decode(buf);
      rows = parseOrdiniAmazon(text);
    }
    if (rows.length === 0) showToast("Nessuna riga riconosciuta nel file caricato", "err");
    setOrdiniParseResult(rows);
  };

  const handleOrdiniFile = async (e) => {
    const file = e.target.files?.[0];
    await processOrdiniFile(file);
    e.target.value = "";
  };

  const handleOrdiniDrop = async (e) => {
    e.preventDefault();
    e.stopPropagation();
    setOrdiniDragOver(false);
    const file = e.dataTransfer.files?.[0];
    await processOrdiniFile(file);
  };

  const ordiniByEan = useMemo(() => {
    const map = {};
    (ordiniParseResult || []).forEach(r => { map[r.ean] = r; });
    return map;
  }, [ordiniParseResult]);

  const ordiniPreview = useMemo(() => {
    if (!ordiniParseResult) return null;
    const matched = dataVerifica.filter(r => ordiniByEan[r.ean]).length;
    return { totRighe: ordiniParseResult.length, matched, nonMatchedNelFile: ordiniParseResult.length - matched };
  }, [ordiniParseResult, ordiniByEan, dataVerifica]);

  const confirmApplyOrdini = async () => {
    if (!ordiniParseResult) return;
    setOrdiniProcessing(true);
    let done = 0;
    for (const r of dataVerifica) {
      const o = ordiniByEan[r.ean];
      if (o) {
        await saveVerificaAmazon(r.anno_lancio, r.num_lancio, r.ean, { copie: o.copie, prenotato: o.prenotato, impegnato: o.impegnato, evaso: o.evaso });
        done++;
      }
    }
    showToast(`Ordini applicati: ${done} titoli del lancio aggiornati (Copie/Prenotato/Impegnato/Evaso)`);
    setOrdiniParseResult(null);
    setShowOrdiniUpload(false);
    setOrdiniProcessing(false);
  };

  // ── Upload "Prima Proposta" (file Excel: Proposta Amazon + Proposta PDE = Amazon Cedola) ──
  const [showProposteUpload, setShowProposteUpload] = useState(false);
  const [proposteDragOver, setProposteDragOver] = useState(false);
  const [proposteParseResult, setProposteParseResult] = useState(null);
  const [proposteProcessing, setProposteProcessing] = useState(false);

  const processProposteFile = async (file) => {
    if (!file) return;
    try {
      const XLSX = window.XLSX;
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
      let headerIdx = 0;
      for (let i = 0; i < Math.min(rows.length, 10); i++) {
        if (rows[i] && rows[i].some(c => String(c).toLowerCase().includes("ean"))) { headerIdx = i; break; }
      }
      const headers = rows[headerIdx].map(h => String(h || "").trim().toLowerCase());
      const col = {};
      headers.forEach((h, i) => {
        if (h === "ean") col.ean = i;
        else if (h.includes("proposta") && h.includes("amazon")) col.propostaAmaz = i;
        else if (h === "asin") col.asin = i;
        else if (h.includes("proposta") && h.includes("vendor")) col.propostaVendor = i;
        else if (h.includes("filtro")) col.filtroReti = i;
        else if (h.includes("preordini") || h.includes("preordine")) col.preordini = i;
      });
      if (col.ean === undefined) throw new Error("Colonna EAN non trovata nel file");
      if (col.propostaAmaz === undefined) throw new Error("Colonna 'Proposta Amazon' non trovata nel file");
      const dataRows = rows.slice(headerIdx + 1).filter(r => r && r[col.ean]);
      const parsed = dataRows.map(r => {
        const ean = String(r[col.ean]).trim();
        const val = r[col.propostaAmaz];
        const propostaAmaz = val === "" || val == null ? null : parseInt(val, 10);
        const numOrNull = v => (v === "" || v == null ? null : (isNaN(parseInt(v, 10)) ? null : parseInt(v, 10)));
        return {
          ean,
          propostaAmaz: isNaN(propostaAmaz) ? null : propostaAmaz,
          asin: col.asin !== undefined ? String(r[col.asin] ?? "").trim() || null : null,
          propostaVendor: col.propostaVendor !== undefined ? numOrNull(r[col.propostaVendor]) : null,
          filtroReti: col.filtroReti !== undefined ? String(r[col.filtroReti] ?? "").trim() || null : null,
          preordini: col.preordini !== undefined ? numOrNull(r[col.preordini]) : null,
        };
      }).filter(r => /^\d{9,14}$/.test(r.ean));
      setProposteParseResult(parsed);
    } catch (err) {
      showToast("Errore lettura file: " + err.message, "err");
    }
  };

  const handleProposteFile = async (e) => {
    const file = e.target.files?.[0];
    await processProposteFile(file);
    e.target.value = "";
  };

  const handleProposteDrop = async (e) => {
    e.preventDefault();
    e.stopPropagation();
    setProposteDragOver(false);
    const file = e.dataTransfer.files?.[0];
    await processProposteFile(file);
  };

  const proposteByEan = useMemo(() => {
    const map = {};
    (proposteParseResult || []).forEach(r => { map[r.ean] = r; });
    return map;
  }, [proposteParseResult]);

  const proposteApplyPreview = useMemo(() => {
    if (!proposteParseResult) return null;
    const matched = dataVerifica.filter(r => proposteByEan[r.ean]);
    const conQuantita = matched.filter(r => proposteByEan[r.ean].propostaAmaz != null && proposteByEan[r.ean].propostaAmaz > 0);
    return { matched: matched.length, conQuantita: conQuantita.length };
  }, [proposteParseResult, proposteByEan, dataVerifica]);

  const confirmApplyProposte = async () => {
    if (!proposteParseResult) return;
    setProposteProcessing(true);
    let done = 0;
    for (const r of dataVerifica) {
      const p = proposteByEan[r.ean];
      if (!p) continue;
      const payload = { proposta_pde: r.vF }; // Proposta PDE = sempre Amazon Cedola
      if (p.propostaAmaz != null && p.propostaAmaz > 0) payload.proposta_amaz = p.propostaAmaz; // Proposta Amaz = solo se c'è quantità
      if (p.asin != null) payload.asin = p.asin;
      if (p.propostaVendor != null) payload.proposta_vendor = p.propostaVendor;
      if (p.filtroReti != null) payload.filtro_reti = p.filtroReti;
      if (p.preordini != null) payload.preordini_vendor = p.preordini;
      await saveVerificaAmazon(r.anno_lancio, r.num_lancio, r.ean, payload);
      done++;
    }
    showToast(`Prima proposta applicata: ${done} titoli aggiornati`);
    setProposteParseResult(null);
    setShowProposteUpload(false);
    setProposteProcessing(false);
  };

  // ── Upload "Preorder" (file .txt tab-delimited: colonna ordini_aperti = preorder) ──
  const [showPreorderUpload, setShowPreorderUpload] = useState(false);
  const [preorderParseResult, setPreorderParseResult] = useState(null);
  const [preorderProcessing, setPreorderProcessing] = useState(false);
  const [preorderDragOver, setPreorderDragOver] = useState(false);

  const parsePreorderFile = (text) => {
    const lines = text.replace(/^\uFEFF/, "").split(/\r?\n/).filter(l => l.trim() !== "");
    if (lines.length < 2) return [];
    const headers = lines[0].split("\t").map(h => h.trim().toLowerCase());
    const idxEan = headers.indexOf("ean");
    const idxPreorder = headers.indexOf("storico_ordini");
    if (idxEan === -1 || idxPreorder === -1) return [];
    const out = [];
    for (let i = 1; i < lines.length; i++) {
      const cols = lines[i].split("\t");
      const ean = (cols[idxEan] || "").trim();
      if (!/^\d{9,14}$/.test(ean)) continue;
      const val = parseInt(cols[idxPreorder], 10);
      out.push({ ean, preorder: isNaN(val) ? 0 : val });
    }
    return out;
  };

  const processPreorderFile = async (file) => {
    if (!file) return;
    const buf = await file.arrayBuffer();
    let text = new TextDecoder("utf-8").decode(buf);
    let rows = parsePreorderFile(text);
    if (rows.length === 0) {
      text = new TextDecoder("utf-16le").decode(buf);
      rows = parsePreorderFile(text);
    }
    if (rows.length === 0) showToast("Nessuna riga riconosciuta nel file caricato (colonne 'ean' e 'storico_ordini' attese)", "err");
    setPreorderParseResult(rows);
  };

  const handlePreorderFile = async (e) => {
    const file = e.target.files?.[0];
    await processPreorderFile(file);
    e.target.value = "";
  };

  const handlePreorderDrop = async (e) => {
    e.preventDefault();
    e.stopPropagation();
    setPreorderDragOver(false);
    const file = e.dataTransfer.files?.[0];
    await processPreorderFile(file);
  };

  const preorderByEan = useMemo(() => {
    const map = {};
    (preorderParseResult || []).forEach(r => { map[r.ean] = r; });
    return map;
  }, [preorderParseResult]);

  const preorderApplyPreview = useMemo(() => {
    if (!preorderParseResult) return null;
    const matched = dataVerifica.filter(r => preorderByEan[r.ean]);
    const conQuantita = matched.filter(r => preorderByEan[r.ean].preorder > 0);
    return { matched: matched.length, conQuantita: conQuantita.length };
  }, [preorderParseResult, preorderByEan, dataVerifica]);

  const confirmApplyPreorder = async () => {
    if (!preorderParseResult) return;
    setPreorderProcessing(true);
    let done = 0;
    for (const r of dataVerifica) {
      const p = preorderByEan[r.ean];
      if (p && p.preorder > 0) {
        await saveVerificaAmazon(r.anno_lancio, r.num_lancio, r.ean, { preorder: p.preorder });
        done++;
      }
    }
    showToast(`Preorder applicati: ${done} titoli aggiornati`);
    setPreorderParseResult(null);
    setShowPreorderUpload(false);
    setPreorderProcessing(false);
  };

  // Upsert su tabella verifica_amazon (chiave: anno_lancio + num_lancio + ean)
  const saveVerificaAmazon = async (annoLancio, numLancio, ean, fields) => {
    const body = { anno_lancio: annoLancio, num_lancio: numLancio, ean, ...fields };
    try {
      const resp = await fetch(`${SUPABASE_URL}/rest/v1/verifica_amazon?on_conflict=anno_lancio,num_lancio,ean`, {
        method: "POST",
        headers: {
          "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`, "Content-Type": "application/json",
          "Prefer": "resolution=merge-duplicates,return=representation",
        },
        body: JSON.stringify(body),
      });
      if (!resp.ok) throw new Error(await resp.text());
      const saved = (await resp.json())[0];
      setVerificaData(prev => {
        const idx = prev.findIndex(x => x.anno_lancio === annoLancio && x.num_lancio === numLancio && x.ean === ean);
        if (idx >= 0) { const copy = [...prev]; copy[idx] = saved; return copy; }
        return [...prev, saved];
      });
    } catch (err) {
      showToast("Errore salvataggio Verifica Amazon: " + err.message, "err");
    }
    setEditCellVA(null);
  };

  const confirmApplyMail = async () => {
    if (!mailParseResult || mailLancioMismatch) return;
    setMailProcessing(true);
    let doneMail = 0, doneAuto = 0;
    for (const r of dataVerifica) {
      const mailRow = mailByEan[r.ean];
      if (mailRow) {
        // Sovrascrittura: solo Copie prende "restituito Amazon" dalla mail, zero compreso (0 = Amazon non ha prenotato). Proposta Amaz non viene toccata.
        await saveVerificaAmazon(r.anno_lancio, r.num_lancio, r.ean, { copie: mailRow.restituito });
        doneMail++;
      } else {
        // Nessuna differenza segnalata da Messaggerie: Amazon ha confermato la proposta iniziale di PDE (Amazon Cedola). Proposta Amaz non viene toccata (verrà popolata da un upload dedicato).
        await saveVerificaAmazon(r.anno_lancio, r.num_lancio, r.ean, { copie: r.vF });
        doneAuto++;
      }
    }
    showToast(`Mail applicata: ${doneMail} titoli da mail, ${doneAuto} confermati sulla proposta iniziale PDE`);
    setMailParseResult(null);
    setMailLancioInfo(null);
    setMailPasteText("");
    setShowMailUpload(false);
    setMailProcessing(false);
  };

  // ── Esporta Controproposta: stessa struttura E formattazione del file "prima proposta" ──
  const exportControproposta = () => {
    const XS = window.XLSXStyle;
    const cHeaders = ["ASIN", "EAN", "Proposta Vendor", "Editore", "Filtro Reti", "Proposta Amazon", "Preordini", "Controproposta Vendor", "Note"];
    const cRows = dataArricchita.map(r => {
      const v = verificaByEan[r.ean] || {};
      return [
        v.asin || "",
        r.ean,
        v.proposta_vendor ?? "",
        r.editore,
        v.filtro_reti || "",
        v.proposta_amaz ?? "",
        v.preordini_vendor ?? "",
        v.proposta_pde ?? "",
        v.note || "",
      ];
    });

    const FONT = { name: "Aptos Narrow", sz: 11 };
    const THIN = { style: "thin", color: { rgb: "FF000000" } };
    const BORDER_ALL = { top: THIN, bottom: THIN, left: THIN, right: THIN };
    const headerStyle = {
      font: { ...FONT, bold: true, color: { rgb: "FFFFFFFF" } },
      fill: { patternType: "solid", fgColor: { rgb: "FF000000" } },
      alignment: { horizontal: "center", vertical: "center" },
      border: BORDER_ALL,
    };
    const cellStyle = { font: FONT, border: BORDER_ALL };
    const cellStyleNum = { ...cellStyle, numFmt: "0" };

    const aoa = [cHeaders, ...cRows];
    const ws = XS.utils.aoa_to_sheet(aoa);
    cHeaders.forEach((_, c) => {
      const addr = XS.utils.encode_cell({ r: 0, c });
      if (ws[addr]) ws[addr].s = headerStyle;
    });
    cRows.forEach((row, rIdx) => {
      row.forEach((_, c) => {
        const addr = XS.utils.encode_cell({ r: rIdx + 1, c });
        if (ws[addr]) ws[addr].s = (c === 0 || c === 1) ? cellStyleNum : cellStyle; // ASIN, EAN senza notazione scientifica
      });
    });
    ws["!cols"] = [
      { wch: 13 }, { wch: 14.5 }, { wch: 15 }, { wch: 22 }, { wch: 10 }, { wch: 16 }, { wch: 10 }, { wch: 21 }, { wch: 30 },
    ];

    const wbC = XS.utils.book_new();
    XS.utils.book_append_sheet(wbC, ws, "Controproposta");
    XS.writeFile(wbC, `Controproposta_Lancio${filterLancio}_${filterAnno}.xlsx`);
  };

  const exportExcel = () => {
    const XLSX = window.XLSX;
    const vHeaders = [
      "EAN", "Titolo", "Autore", "Editore", "Prezzo", "Totale", "Amazon cedola",
      "Proposta Amaz", "Taglio %", "Proposta PDE", "Copie", "Prenotato", "Impegnato", "Evaso", "Inevaso", "Netto",
      "Diff FG", "Diff Amz", "Diff PDE", "% Diff FG", "% Diff Amz", "% Diff PDE",
      "Note", "Preorder", "Residuo", "% Preorder/Netto", "Rifornim.", "Rottura stock"
    ];
    const vRows = dataVerifica.map(r => [
      r.ean, r.titolo, r.autore, r.editore, r.prezzo, r.vE, r.vF,
      r.vG ?? "", r.vH ?? "", r.vI ?? "", r.vJ ?? "", r.vPren ?? "", r.vImp ?? "", r.vK ?? "", r.vL ?? "", r.vM ?? "",
      r.vN ?? "", r.vO ?? "", r.vP ?? "", r.vQ ?? "", r.vR ?? "", r.vS ?? "",
      r.vNote, r.vU, r.vV ?? "", r.vW ?? "", r.vX ? "SI" : "", r.vY ? "SI" : ""
    ]);
    const wsV = XLSX.utils.aoa_to_sheet([vHeaders, ...vRows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, wsV, "Verifica Amazon");
    XLSX.writeFile(wb, `VerificaAmazon_${filterLancio}_${filterAnno}.xlsx`);
  };

  if (loading) return <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", color: T.textMid }}>Caricamento...</div>;

  if (data.length === 0) return (
    <div style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 20, padding: 40 }}>
      <div style={{ fontSize: "48px", opacity: 0.3 }}>📦</div>
      <div style={{ color: T.textMid, fontSize: "14px", textAlign: "center", maxWidth: 400 }}>Nessun lancio caricato. Carica prima un lancio nella sezione Lanci Settimanali.</div>
    </div>
  );

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%", overflow: "hidden" }}>
      {/* TOOLBAR */}
      <div style={{ padding: "14px 20px", borderBottom: `1px solid ${T.border}`, display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
        <select style={{ ...css.input, fontSize: "13px", fontWeight: "600", color: T.accent }} value={filterAnno || ""} onChange={e => { setFilterAnno(Number(e.target.value)); setFilterLancio(null); }}>
          {anniDisponibili.map(a => <option key={a} value={a}>{a}</option>)}
        </select>
        <select style={{ ...css.input, fontSize: "13px", fontWeight: "600", color: T.accent }} value={filterLancio || ""} onChange={e => setFilterLancio(Number(e.target.value))}>
          {lanciPerAnno.map(n => <option key={n} value={n}>Lancio {n}</option>)}
        </select>
        <input style={{ ...css.input, width: 180 }} placeholder="Cerca EAN / titolo..." value={search} onChange={e => setSearch(e.target.value)} />
        <div style={{ marginLeft: "auto", display: "flex", gap: 10 }}>
          <button style={{ ...css.btn(showProposteUpload ? "accent" : undefined) }} onClick={() => setShowProposteUpload(s => !s)}>
            🎯 Carica prima proposta
          </button>
          <button style={css.btn()} onClick={exportControproposta}>
            🔁 Esporta controproposta
          </button>
          <button style={{ ...css.btn(showMailUpload ? "accent" : undefined) }} onClick={() => setShowMailUpload(s => !s)}>
            📧 Carica mail Messaggerie
          </button>
          <button style={{ ...css.btn(showOrdiniUpload ? "accent" : undefined) }} onClick={() => setShowOrdiniUpload(s => !s)}>
            📥 Carica ordini Amazon
          </button>
          <button style={{ ...css.btn(showPreorderUpload ? "accent" : undefined) }} onClick={() => setShowPreorderUpload(s => !s)}>
            📖 Carica preorder
          </button>
          <button style={css.btn()} onClick={exportExcel}>Download Excel</button>
        </div>
      </div>

      {/* KPI */}
      <div style={{ padding: "16px 20px", borderBottom: `1px solid ${T.border}` }}>
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
          <KpiCard label="Proposto ad Amazon" value={kpiVerifica.totProposto.toLocaleString("it")} color="#e8a838" sub={`€ ${kpiVerifica.valoreProposto.toLocaleString("it", { maximumFractionDigits: 0 })}`} />
          <KpiCard label="Confermato (Netto)" value={kpiVerifica.totConfermato.toLocaleString("it")} color={T.green} sub={`€ ${kpiVerifica.valoreConfermato.toLocaleString("it", { maximumFractionDigits: 0 })}`} />
          <KpiCard label="Scostamento vs proposta" value={(kpiVerifica.scostamento > 0 ? "+" : "") + kpiVerifica.scostamento.toLocaleString("it")} color={kpiVerifica.scostamento < 0 ? T.red : T.green} sub={`${kpiVerifica.valoreScostamento > 0 ? "+" : ""}€ ${kpiVerifica.valoreScostamento.toLocaleString("it", { maximumFractionDigits: 0 })}`} />
          <KpiCard label="Titoli confermati" value={kpiVerifica.nConfermati} color={T.text} />
          <KpiCard label="In attesa di conferma" value={kpiVerifica.nInAttesa} color={kpiVerifica.nInAttesa > 0 ? "#e8a838" : T.textMid} />
        </div>
      </div>

      {/* UPLOAD MAIL MESSAGGERIE */}
      {showMailUpload && (
        <div style={{ padding: "16px 20px", borderBottom: `1px solid ${T.border}`, background: T.bg }}>
          {!mailParseResult ? (
            <>
              <div style={{ fontSize: "12px", color: T.textMid, marginBottom: 10 }}>
                Carica il file <b>.msg</b> della mail "DIFFERENZA PRENOTAZIONI AMAZON" di Messaggerie, oppure incolla qui sotto il testo del corpo mail.
                I titoli presenti nella mail aggiornano <b>Copie</b> con il valore "restituito Amazon" (anche se è zero: significa che Amazon non ha prenotato). Tutti gli altri titoli del lancio (senza differenza segnalata) vengono confermati in automatico con la <b>proposta iniziale PDE</b> (Amazon Cedola) in <b>Copie</b>. La colonna <b>Proposta Amaz</b> non viene mai toccata da questo caricamento. Ogni caricamento sovrascrive i dati esistenti.
              </div>
              <div style={{ display: "flex", gap: 10, alignItems: "flex-start", flexWrap: "wrap" }}>
                <label
                  style={{
                    ...css.btn(mailDragOver ? "accent" : undefined),
                    cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center",
                    width: 200, minHeight: 70, border: `2px dashed ${mailDragOver ? T.accent : T.borderHi}`,
                    background: mailDragOver ? T.accent + "18" : T.surface, textAlign: "center", padding: "8px",
                  }}
                  onDragOver={e => { e.preventDefault(); e.stopPropagation(); setMailDragOver(true); }}
                  onDragLeave={e => { e.preventDefault(); e.stopPropagation(); setMailDragOver(false); }}
                  onDrop={handleMailDrop}
                >
                  <span>↑ Trascina qui il file .msg</span>
                  <span style={{ fontSize: "10px", color: T.textDim, marginTop: 3 }}>oppure clicca per sceglierlo</span>
                  <input type="file" accept=".msg,.txt,.eml" style={{ display: "none" }} onChange={handleMailFile} />
                </label>
                <div style={{ display: "flex", flexDirection: "column", gap: 6, flex: 1, minWidth: 280 }}>
                  <textarea
                    style={{ ...css.input, minHeight: 70, fontFamily: "inherit", fontSize: "11px", resize: "vertical" }}
                    placeholder="...oppure incolla qui il testo della mail Messaggerie"
                    value={mailPasteText}
                    onChange={e => setMailPasteText(e.target.value)}
                  />
                  <button style={{ ...css.btn(), alignSelf: "flex-start" }} disabled={!mailPasteText.trim()} onClick={() => runMailParse(mailPasteText)}>Estrai dati dal testo incollato</button>
                </div>
                <button style={{ ...css.btn(), alignSelf: "flex-start" }} onClick={() => setShowMailUpload(false)}>Annulla</button>
              </div>
            </>
          ) : (
            <>
              {mailLancioMismatch ? (
                <div style={{ background: T.red + "18", border: `1px solid ${T.red}`, borderRadius: 4, padding: "10px 14px", marginBottom: 12, color: T.red, fontSize: "12px", fontWeight: "600" }}>
                  🚫 {mailLancioMismatch.msg}
                </div>
              ) : mailLancioInfo && (
                <div style={{ background: T.green + "18", border: `1px solid ${T.green}`, borderRadius: 4, padding: "8px 14px", marginBottom: 12, color: T.green, fontSize: "12px" }}>
                  ✓ La mail si riferisce al lancio {mailLancioInfo.numLancio}{mailLancioInfo.anno ? "/" + mailLancioInfo.anno : ""}, corrispondente al lancio selezionato.
                </div>
              )}
              <div style={{ fontSize: "12px", color: T.text, marginBottom: 10 }}>
                Trovati <b style={{ color: "#e8a838" }}>{mailParseResult.length}</b> titoli nella mail. Gli altri <b style={{ color: T.green }}>{mailPreview?.autoConfermati ?? 0}</b> titoli del lancio (senza differenza) verranno confermati in automatico sulla proposta iniziale PDE (Amazon Cedola). Tutti i valori esistenti verranno sovrascritti.
              </div>
              {mailParseResult.length > 0 && (
                <div style={{ maxHeight: 260, overflowY: "auto", marginBottom: 12, border: `1px solid ${T.border}`, borderRadius: 4 }}>
                <table style={css.table}>
                  <thead>
                    <tr>
                      <th style={css.th}>EAN</th>
                      <th style={css.th}>Titolo</th>
                      <th style={css.th}>Prezzo</th>
                      <th style={css.th}>Proposto</th>
                      <th style={css.th}>Restituito</th>
                    </tr>
                  </thead>
                  <tbody>
                    {mailParseResult.map((r, i) => (
                      <tr key={i}>
                        <td style={{ ...css.td, fontFamily: "monospace", fontSize: "11px", color: T.textMid }}>{r.ean}</td>
                        <td style={css.td}>{r.titolo}</td>
                        <td style={css.td}>€ {r.prezzo.toFixed(2)}</td>
                        <td style={css.td}>{r.proposto}</td>
                        <td style={{ ...css.td, color: r.restituito < r.proposto ? T.red : T.green, fontWeight: "700" }}>{r.restituito}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                </div>
              )}
              <div style={{ display: "flex", gap: 10 }}>
                <button style={css.btn("accent")} disabled={mailProcessing || !!mailLancioMismatch} title={mailLancioMismatch ? "Seleziona il lancio corretto per procedere" : ""} onClick={confirmApplyMail}>
                  {mailProcessing ? "Applico..." : "✓ Conferma e applica"}
                </button>
                <button style={css.btn()} disabled={mailProcessing} onClick={() => { setMailParseResult(null); setMailPasteText(""); }}>Annulla</button>
              </div>
            </>
          )}
        </div>
      )}

      {/* UPLOAD ORDINI AMAZON */}
      {showOrdiniUpload && (
        <div style={{ padding: "16px 20px", borderBottom: `1px solid ${T.border}`, background: T.bg }}>
          {!ordiniParseResult ? (
            <>
              <div style={{ fontSize: "12px", color: T.textMid, marginBottom: 10 }}>
                Carica il file <b>Dettaglio cliente</b> (export ordini Amazon). Le righe vengono aggregate per EAN e aggiornano <b>Copie</b>, <b>Prenotato</b>, <b>Impegnato</b> ed <b>Evaso</b>. Solo i titoli presenti nel file e nel lancio selezionato vengono toccati; gli altri restano invariati. Ogni caricamento sovrascrive i valori esistenti per i titoli coinvolti.
              </div>
              <div style={{ display: "flex", gap: 10, alignItems: "flex-start", flexWrap: "wrap" }}>
                <label
                  style={{
                    ...css.btn(ordiniDragOver ? "accent" : undefined),
                    cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center",
                    width: 200, minHeight: 70, border: `2px dashed ${ordiniDragOver ? T.accent : T.borderHi}`,
                    background: ordiniDragOver ? T.accent + "18" : T.surface, textAlign: "center", padding: "8px",
                  }}
                  onDragOver={e => { e.preventDefault(); e.stopPropagation(); setOrdiniDragOver(true); }}
                  onDragLeave={e => { e.preventDefault(); e.stopPropagation(); setOrdiniDragOver(false); }}
                  onDrop={handleOrdiniDrop}
                >
                  <span>↑ Trascina qui il file</span>
                  <span style={{ fontSize: "10px", color: T.textDim, marginTop: 3 }}>oppure clicca per sceglierlo (.csv)</span>
                  <input type="file" accept=".csv,.txt" style={{ display: "none" }} onChange={handleOrdiniFile} />
                </label>
                <button style={{ ...css.btn(), alignSelf: "flex-start" }} onClick={() => setShowOrdiniUpload(false)}>Annulla</button>
              </div>
            </>
          ) : (
            <>
              <div style={{ fontSize: "12px", color: T.text, marginBottom: 10 }}>
                Trovati <b style={{ color: "#e8a838" }}>{ordiniPreview?.totRighe ?? 0}</b> EAN nel file, di cui <b style={{ color: T.green }}>{ordiniPreview?.matched ?? 0}</b> corrispondono a titoli del lancio {filterLancio}/{filterAnno} selezionato. Verranno aggiornati solo questi.
              </div>
              {ordiniParseResult.length > 0 && (
                <div style={{ maxHeight: 260, overflowY: "auto", marginBottom: 12, border: `1px solid ${T.border}`, borderRadius: 4 }}>
                <table style={css.table}>
                  <thead>
                    <tr>
                      <th style={css.th}>EAN</th>
                      <th style={css.th}>Titolo</th>
                      <th style={css.th}>Copie</th>
                      <th style={css.th}>Prenotato</th>
                      <th style={css.th}>Impegnato</th>
                      <th style={css.th}>Evaso</th>
                      <th style={css.th}>Nel lancio</th>
                    </tr>
                  </thead>
                  <tbody>
                    {ordiniParseResult.map((r, i) => (
                      <tr key={i} style={{ opacity: dataVerifica.some(d => d.ean === r.ean) ? 1 : 0.5 }}>
                        <td style={{ ...css.td, fontFamily: "monospace", fontSize: "11px", color: T.textMid }}>{r.ean}</td>
                        <td style={css.td}>{r.titolo}</td>
                        <td style={css.td}>{r.copie.toLocaleString("it")}</td>
                        <td style={css.td}>{r.prenotato.toLocaleString("it")}</td>
                        <td style={css.td}>{r.impegnato.toLocaleString("it")}</td>
                        <td style={css.td}>{r.evaso.toLocaleString("it")}</td>
                        <td style={{ ...css.td, textAlign: "center" }}>{dataVerifica.some(d => d.ean === r.ean) ? <span style={{ color: T.green }}>✓</span> : <span style={{ color: T.textDim }}>—</span>}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                </div>
              )}
              <div style={{ display: "flex", gap: 10 }}>
                <button style={css.btn("accent")} disabled={ordiniProcessing} onClick={confirmApplyOrdini}>
                  {ordiniProcessing ? "Applico..." : "✓ Conferma e applica"}
                </button>
                <button style={css.btn()} disabled={ordiniProcessing} onClick={() => setOrdiniParseResult(null)}>Annulla</button>
              </div>
            </>
          )}
        </div>
      )}

      {/* UPLOAD PRIMA PROPOSTA */}
      {showProposteUpload && (
        <div style={{ padding: "16px 20px", borderBottom: `1px solid ${T.border}`, background: T.bg }}>
          {!proposteParseResult ? (
            <>
              <div style={{ fontSize: "12px", color: T.textMid, marginBottom: 10 }}>
                Carica il file Excel della <b>prima proposta</b> (colonne EAN e Proposta Amazon). Aggiorna <b>Proposta Amaz</b> solo sui titoli con quantità nel file. <b>Proposta PDE</b> viene sempre riportata a Amazon Cedola per i titoli presenti nel file; se la Proposta Amaz differisce da Amazon Cedola, il numero in Proposta PDE viene evidenziato in rosso. Ogni caricamento sovrascrive i dati esistenti.
              </div>
              <div style={{ display: "flex", gap: 10, alignItems: "flex-start", flexWrap: "wrap" }}>
                <label
                  style={{
                    ...css.btn(proposteDragOver ? "accent" : undefined),
                    cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center",
                    width: 200, minHeight: 70, border: `2px dashed ${proposteDragOver ? T.accent : T.borderHi}`,
                    background: proposteDragOver ? T.accent + "18" : T.surface, textAlign: "center", padding: "8px",
                  }}
                  onDragOver={e => { e.preventDefault(); e.stopPropagation(); setProposteDragOver(true); }}
                  onDragLeave={e => { e.preventDefault(); e.stopPropagation(); setProposteDragOver(false); }}
                  onDrop={handleProposteDrop}
                >
                  <span>↑ Trascina qui il file Excel</span>
                  <span style={{ fontSize: "10px", color: T.textDim, marginTop: 3 }}>oppure clicca per sceglierlo (.xlsx)</span>
                  <input type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleProposteFile} />
                </label>
                <button style={{ ...css.btn(), alignSelf: "flex-start" }} onClick={() => setShowProposteUpload(false)}>Annulla</button>
              </div>
            </>
          ) : (
            <>
              <div style={{ fontSize: "12px", color: T.text, marginBottom: 10 }}>
                Trovati <b style={{ color: "#e8a838" }}>{proposteParseResult.length}</b> EAN nel file, di cui <b style={{ color: T.green }}>{proposteApplyPreview?.matched ?? 0}</b> nel lancio {filterLancio}/{filterAnno} selezionato ({proposteApplyPreview?.conQuantita ?? 0} con quantità in Proposta Amazon).
              </div>
              <div style={{ display: "flex", gap: 10 }}>
                <button style={css.btn("accent")} disabled={proposteProcessing} onClick={confirmApplyProposte}>
                  {proposteProcessing ? "Applico..." : "✓ Conferma e applica"}
                </button>
                <button style={css.btn()} disabled={proposteProcessing} onClick={() => setProposteParseResult(null)}>Annulla</button>
              </div>
            </>
          )}
        </div>
      )}

      {/* UPLOAD PREORDER */}
      {showPreorderUpload && (
        <div style={{ padding: "16px 20px", borderBottom: `1px solid ${T.border}`, background: T.bg }}>
          {!preorderParseResult ? (
            <>
              <div style={{ fontSize: "12px", color: T.textMid, marginBottom: 10 }}>
                Carica il file <b>.txt</b> dei preorder (colonne <code>ean</code> e <code>storico_ordini</code>). Aggiorna <b>Preorder</b> solo sui titoli con quantità &gt; 0.
              </div>
              <div style={{ display: "flex", gap: 10, alignItems: "flex-start", flexWrap: "wrap" }}>
                <label
                  style={{
                    ...css.btn(preorderDragOver ? "accent" : undefined),
                    cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center",
                    width: 200, minHeight: 70, border: `2px dashed ${preorderDragOver ? T.accent : T.borderHi}`,
                    background: preorderDragOver ? T.accent + "18" : T.surface, textAlign: "center", padding: "8px",
                  }}
                  onDragOver={e => { e.preventDefault(); e.stopPropagation(); setPreorderDragOver(true); }}
                  onDragLeave={e => { e.preventDefault(); e.stopPropagation(); setPreorderDragOver(false); }}
                  onDrop={handlePreorderDrop}
                >
                  <span>↑ Trascina qui il file .txt</span>
                  <span style={{ fontSize: "10px", color: T.textDim, marginTop: 3 }}>oppure clicca per sceglierlo (tab-delimited)</span>
                  <input type="file" accept=".txt,.csv" style={{ display: "none" }} onChange={handlePreorderFile} />
                </label>
                <button style={{ ...css.btn(), alignSelf: "flex-start" }} onClick={() => setShowPreorderUpload(false)}>Annulla</button>
              </div>
            </>
          ) : (
            <>
              <div style={{ fontSize: "12px", color: T.text, marginBottom: 10 }}>
                Trovati <b style={{ color: "#e8a838" }}>{preorderParseResult.length}</b> EAN nel file, di cui <b style={{ color: T.green }}>{preorderApplyPreview?.matched ?? 0}</b> nel lancio {filterLancio}/{filterAnno} selezionato ({preorderApplyPreview?.conQuantita ?? 0} con quantità preorder &gt; 0).
              </div>
              <div style={{ display: "flex", gap: 10 }}>
                <button style={css.btn("accent")} disabled={preorderProcessing} onClick={confirmApplyPreorder}>
                  {preorderProcessing ? "Applico..." : "✓ Conferma e applica"}
                </button>
                <button style={css.btn()} disabled={preorderProcessing} onClick={() => setPreorderParseResult(null)}>Annulla</button>
              </div>
            </>
          )}
        </div>
      )}

      {/* RIEPILOGO PER EDITORE */}
      <div style={{ padding: "10px 20px", borderBottom: `1px solid ${T.border}` }}>
        <div
          style={{ color: T.textMid, fontSize: "10px", letterSpacing: "0.1em", textTransform: "uppercase", cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 6, userSelect: "none" }}
          onClick={() => setShowRiepilogoEditori(s => !s)}
        >
          <span style={{ transform: showRiepilogoEditori ? "rotate(90deg)" : "none", transition: "transform 0.15s", display: "inline-block", fontSize: "9px" }}>▶</span>
          Riepilogo per editore ({riepilogoEditoriVerifica.length})
        </div>
        {showRiepilogoEditori && (
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginTop: 10 }}>
          {riepilogoEditoriVerifica.map(e => (
            <div key={e.editore} style={{ background: T.bg, border: `1px solid ${T.border}`, borderRadius: 4, padding: "8px 12px", minWidth: 150, flex: "1 1 150px", maxWidth: 220 }}>
              <div style={{ fontSize: "11px", fontWeight: "600", color: T.text, marginBottom: 4, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{e.editore}</div>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 3 }}>
                <div style={{ flex: 1, height: 4, background: T.borderHi, borderRadius: 2, overflow: "hidden" }}>
                  <div style={{ width: `${Math.round(e.confermato / (riepilogoEditoriVerifica[0]?.confermato || 1) * 100)}%`, height: "100%", background: T.green }} />
                </div>
                <span style={{ fontSize: "11px", fontWeight: "700", color: T.green, minWidth: 40, textAlign: "right" }}>{e.confermato.toLocaleString("it")}</span>
              </div>
              {e.proposto > 0 && (
                <div style={{ fontSize: "9px", color: "#e8a838" }}>🎯 {e.proposto.toLocaleString("it")} · € {Math.round(e.valoreProposto).toLocaleString("it")}</div>
              )}
              <div style={{ fontSize: "9px", color: T.textDim, marginTop: 2 }}>{e.titoli} titoli · € {Math.round(e.valore).toLocaleString("it")}</div>
            </div>
          ))}
        </div>
        )}
      </div>

      {/* TABELLA */}
      <div style={{ flex: 1, overflowY: "auto", overflowX: "auto" }}>
        <table style={css.table}>
          <thead>
            <tr>
              <th style={css.th}>EAN</th>
              <th style={css.th}>Titolo</th>
              <th style={css.th}>Autore</th>
              <th style={css.th}>Editore</th>
              <th style={css.th}>Prezzo</th>
              <th style={css.th} title="Totale teorico (F.G. trasmesso + Amazon)">Totale</th>
              <th style={{ ...css.th, color: "#e8a838" }} title="Amazon in cedola (attuale)">Amazon cedola</th>
              <th style={css.th} title="Quantità proposta ad Amazon">Proposta Amaz</th>
              <th style={css.th} title="Taglio prenotazione = (Proposta-Cedola)/Cedola">Taglio %</th>
              <th style={css.th} title="Proposta interna PDE">Proposta PDE</th>
              <th style={css.th} title="Copie effettivamente inserite da Amazon">Copie</th>
              <th style={css.th}>Prenotato</th>
              <th style={css.th}>Impegnato</th>
              <th style={css.th}>Evaso</th>
              <th style={css.th}>Inevaso</th>
              <th style={css.th} title="Netto = Copie - Inevaso">Netto</th>
              <th style={css.th} title="Netto - Amazon cedola">Diff FG</th>
              <th style={css.th} title="Netto - Proposta Amaz">Diff Amz</th>
              <th style={css.th} title="Netto - Proposta PDE">Diff PDE</th>
              <th style={css.th}>Note</th>
              <th style={css.th}>Preorder</th>
              <th style={css.th} title="Netto - Preorder">Residuo</th>
              <th style={css.th}>Rifornim.</th>
              <th style={css.th}>Rottura stock</th>
            </tr>
          </thead>
          <tbody>
            {dataVerifica.map((r, i) => {
              const editVACell = (field, type = "number", forceRed = false) => {
                const isEditing = editCellVA?.ean === r.ean && editCellVA?.field === field;
                const rawVal = { proposta_amaz: r.vG, proposta_pde: r.vI, copie: r.vJ, prenotato: r.vPren, impegnato: r.vImp, evaso: r.vK, inevaso: r.vL, note: r.va.note, preorder: r.vU }[field];
                if (isEditing) {
                  return (
                    <div style={{ display: "flex", gap: 3, alignItems: "center" }}>
                      <input
                        type={type} style={{ ...css.input, width: type === "text" ? 120 : 65, padding: "2px 5px", fontSize: "11px" }}
                        value={editCellVA.value} autoFocus
                        onChange={e => setEditCellVA(p => ({ ...p, value: e.target.value }))}
                        onKeyDown={e => {
                          if (e.key === "Enter") saveVerificaAmazon(r.anno_lancio, r.num_lancio, r.ean, { [field]: type === "number" ? (editCellVA.value === "" ? null : parseInt(editCellVA.value)) : editCellVA.value });
                          if (e.key === "Escape") setEditCellVA(null);
                        }}
                      />
                      <button style={{ ...css.btn("accent"), padding: "1px 5px", fontSize: "10px" }}
                        onClick={() => saveVerificaAmazon(r.anno_lancio, r.num_lancio, r.ean, { [field]: type === "number" ? (editCellVA.value === "" ? null : parseInt(editCellVA.value)) : editCellVA.value })}>✓</button>
                    </div>
                  );
                }
                return (
                  <div style={{ cursor: "pointer", display: "flex", alignItems: "center", gap: 4 }}
                    onClick={() => setEditCellVA({ ean: r.ean, field, value: rawVal != null ? String(rawVal) : "" })}>
                    <span style={{ color: rawVal == null || rawVal === "" ? T.textDim : forceRed ? T.red : T.text }}>{rawVal != null && rawVal !== "" ? rawVal : "—"}</span>
                    <span style={{ color: T.accent, fontSize: "10px" }}>✎</span>
                  </div>
                );
              };
              return (
                <tr key={r.ean} style={{ background: r.haConferma ? T.green + "14" : (i % 2 === 0 ? "transparent" : T.surface + "66") }}>
                  <td style={{ ...css.td, fontFamily: "monospace", fontSize: "11px", color: T.textMid }}>{r.ean}</td>
                  <td style={{ ...css.td, maxWidth: 200 }}><div style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontWeight: "600" }}>{r.titolo}</div></td>
                  <td style={{ ...css.td, color: T.textMid, maxWidth: 130, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.autore}</td>
                  <td style={{ ...css.td, color: T.accent, fontWeight: "600", whiteSpace: "nowrap", maxWidth: 130, overflow: "hidden", textOverflow: "ellipsis" }}>{r.editore}</td>
                  <td style={{ ...css.td, whiteSpace: "nowrap" }}>€ {(r.prezzo || 0).toFixed(2)}</td>
                  <td style={css.td}>{r.vE.toLocaleString("it")}</td>
                  <td style={{ ...css.td, color: "#e8a838", fontWeight: "600" }}>{r.vF.toLocaleString("it")}</td>
                  <td style={css.td}>{editVACell("proposta_amaz")}</td>
                  <td style={{ ...css.td, color: r.vH == null ? T.textDim : r.vH < 0 ? T.red : T.green }}>{r.vH != null ? (r.vH * 100).toFixed(1) + "%" : "—"}</td>
                  <td style={css.td}>{editVACell("proposta_pde", "number", r.vG != null && r.vG !== r.vF)}</td>
                  <td style={css.td}>{editVACell("copie")}</td>
                  <td style={css.td}>{editVACell("prenotato")}</td>
                  <td style={css.td}>{editVACell("impegnato")}</td>
                  <td style={css.td}>{editVACell("evaso")}</td>
                  <td style={css.td}>{editVACell("inevaso")}</td>
                  <td style={{ ...css.td, fontWeight: "700", color: r.haConferma ? T.green : T.textDim }}>{r.vM != null ? r.vM.toLocaleString("it") : "—"}</td>
                  <td style={{ ...css.td, color: r.vN == null ? T.textDim : r.vN < 0 ? T.red : T.green }}>{r.vN != null ? (r.vN > 0 ? "+" : "") + r.vN.toLocaleString("it") : "—"}</td>
                  <td style={{ ...css.td, color: r.vO == null ? T.textDim : r.vO < 0 ? T.red : T.green }}>{r.vO != null ? (r.vO > 0 ? "+" : "") + r.vO.toLocaleString("it") : "—"}</td>
                  <td style={{ ...css.td, color: r.vP == null ? T.textDim : r.vP < 0 ? T.red : T.green }}>{r.vP != null ? (r.vP > 0 ? "+" : "") + r.vP.toLocaleString("it") : "—"}</td>
                  <td style={{ ...css.td, fontSize: "11px", maxWidth: 140 }}>{editVACell("note", "text")}</td>
                  <td style={css.td}>{editVACell("preorder")}</td>
                  <td style={css.td}>{r.vV != null ? r.vV.toLocaleString("it") : "—"}</td>
                  <td style={{ ...css.td, textAlign: "center", cursor: "pointer" }} onClick={() => saveVerificaAmazon(r.anno_lancio, r.num_lancio, r.ean, { richiesta_rifornimento: !r.vX })}>
                    {r.vX ? <span style={{ color: "#e8a838" }}>⚠️</span> : <span style={{ color: T.textDim }}>—</span>}
                  </td>
                  <td style={{ ...css.td, textAlign: "center", cursor: "pointer" }} onClick={() => saveVerificaAmazon(r.anno_lancio, r.num_lancio, r.ean, { rottura_stock: !r.vY })}>
                    {r.vY ? <span style={{ color: T.red }}>🔴</span> : <span style={{ color: T.textDim }}>—</span>}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
        {dataFiltrata.length === 0 && (
          <div style={{ padding: 40, textAlign: "center", color: T.textDim }}>Nessun titolo per questo lancio.</div>
        )}
      </div>

      {toast && (
        <div style={{ position: "fixed", bottom: 24, left: "50%", transform: "translateX(-50%)", background: toast.type === "err" ? "#4a1a2a" : "#1a3a2a", border: `1px solid ${toast.type === "err" ? T.red : T.green}`, color: toast.type === "err" ? T.red : T.green, borderRadius: 6, padding: "8px 20px", fontSize: "12px", zIndex: 999, boxShadow: "0 4px 20px #0008" }}>
          {toast.msg}
        </div>
      )}
    </div>
  );
}
// ─── MODULO ANTICIPI LANCIO ────────────────────────────────────────────────
function ModuloAnticipiLancio({ token, userEmail }) {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [template, setTemplate] = useState(null);
  const [filterStato, setFilterStato] = useState("attivi"); // attivi | da_gestire | notificato | gestito | tutti
  const [search, setSearch] = useState("");
  const [selected, setSelected] = useState(new Set());
  const [showForm, setShowForm] = useState(false);
  const [saving, setSaving] = useState(false);
  const [toast, setToast] = useState(null);
  const emptyForm = { codice_cliente: "", ean: "", quantita: "", sconto_occasionale: "", pagamento_occasionale: "", numero_ordine: "", note: "" };
  const [form, setForm] = useState(emptyForm);

  const showToast = (msg, type = "ok") => { setToast({ msg, type }); setTimeout(() => setToast(null), 3500); };

  const loadData = useCallback(() => {
    if (!token) return;
    setLoading(true);
    sbFetch("novita_fuori_lancio?select=*,lanci_settimanali(num_lancio,anno_lancio,titolo,editore)&order=created_at.desc", token).then(d => {
      setData(Array.isArray(d) ? d : []);
      setLoading(false);
    });
  }, [token]);
  useEffect(() => { loadData(); }, [loadData]);

  useEffect(() => {
    if (!token) return;
    sbFetch("mail_templates?select=*&key=eq.anticipo_lancio", token).then(d => { if (Array.isArray(d) && d[0]) setTemplate(d[0]); });
  }, [token]);

  const salvaNuovo = async () => {
    if (!form.codice_cliente.trim() || !form.ean.trim() || !form.quantita) {
      showToast("Codice cliente, EAN e quantità sono obbligatori", "err"); return;
    }
    setSaving(true);
    const payload = {
      codice_cliente: form.codice_cliente.trim(),
      ean: form.ean.trim(),
      quantita: parseInt(form.quantita, 10) || 0,
      sconto_occasionale: form.sconto_occasionale ? parseFloat(String(form.sconto_occasionale).replace(",", ".")) : null,
      pagamento_occasionale: form.pagamento_occasionale.trim() || null,
      numero_ordine: form.numero_ordine.trim() || null,
      note: form.note.trim() || null,
      creato_da: userEmail || null,
    };
    const r = await fetch(`${SUPABASE_URL}/rest/v1/novita_fuori_lancio`, {
      method: "POST",
      headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`, "Content-Type": "application/json", "Prefer": "return=representation" },
      body: JSON.stringify([payload]),
    });
    if (r.ok) {
      showToast("Anticipo lancio aggiunto");
      setShowForm(false);
      setForm(emptyForm);
      loadData();
    } else {
      showToast("Errore nel salvataggio: " + await r.text(), "err");
    }
    setSaving(false);
  };

  const segnaGestito = async (id) => {
    const r = await fetch(`${SUPABASE_URL}/rest/v1/novita_fuori_lancio?id=eq.${id}`, {
      method: "PATCH",
      headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`, "Content-Type": "application/json", "Prefer": "return=minimal" },
      body: JSON.stringify({ stato: "gestito", gestito_at: new Date().toISOString() }),
    });
    if (r.ok) {
      setData(prev => prev.map(x => x.id === id ? { ...x, stato: "gestito", gestito_at: new Date().toISOString() } : x));
      showToast("Ordine segnato come gestito");
    } else showToast("Errore aggiornamento stato", "err");
  };

  const riapri = async (id) => {
    const r = await fetch(`${SUPABASE_URL}/rest/v1/novita_fuori_lancio?id=eq.${id}`, {
      method: "PATCH",
      headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`, "Content-Type": "application/json", "Prefer": "return=minimal" },
      body: JSON.stringify({ stato: "da_gestire", gestito_at: null }),
    });
    if (r.ok) { setData(prev => prev.map(x => x.id === id ? { ...x, stato: "da_gestire", gestito_at: null } : x)); showToast("Riaperto"); }
  };

  const eliminaRiga = async (id) => {
    if (!confirm("Eliminare questa riga?")) return;
    const r = await fetch(`${SUPABASE_URL}/rest/v1/novita_fuori_lancio?id=eq.${id}`, {
      method: "DELETE",
      headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}` },
    });
    if (r.ok) { setData(prev => prev.filter(x => x.id !== id)); showToast("Riga eliminata"); }
  };

  const dataFiltrata = useMemo(() => {
    let result = [...data];
    if (filterStato === "attivi") result = result.filter(r => r.stato !== "gestito");
    else if (filterStato !== "tutti") result = result.filter(r => r.stato === filterStato);
    if (search) {
      const q = search.toLowerCase();
      result = result.filter(r => r.codice_cliente?.toLowerCase().includes(q) || r.ean?.includes(q) || r.numero_ordine?.toLowerCase().includes(q));
    }
    return result;
  }, [data, filterStato, search]);

  const counts = useMemo(() => ({
    da_gestire: data.filter(r => r.stato === "da_gestire").length,
    notificato: data.filter(r => r.stato === "notificato").length,
    gestito: data.filter(r => r.stato === "gestito").length,
  }), [data]);

  const toggleSelect = (id) => setSelected(prev => { const n = new Set(prev); n.has(id) ? n.delete(id) : n.add(id); return n; });

  const creaMail = (row) => {
    if (!template) { showToast("Template mail non disponibile", "err"); return; }
    const lancioLabel = row.lanci_settimanali ? `${row.lanci_settimanali.num_lancio}/${row.lanci_settimanali.anno_lancio}` : "(non ancora lanciato)";
    const corpo = (template.corpo || "")
      .replaceAll("{cliente}", row.codice_cliente)
      .replaceAll("{lancio}", lancioLabel)
      .replaceAll("{ean}", row.ean);
    const subject = encodeURIComponent(template.oggetto || "Anticipo lancio");
    const body = encodeURIComponent(corpo);
    window.location.href = `mailto:?subject=${subject}&body=${body}`;
  };

  const exportMessaggerie = () => {
    const righe = selected.size > 0 ? dataFiltrata.filter(r => selected.has(r.id)) : dataFiltrata.filter(r => r.stato !== "gestito");
    if (righe.length === 0) { showToast("Nessuna riga da esportare", "err"); return; }
    const XLSX = window.XLSX;
    const header = [
      "Cod. cliente \n(max 10 caratteri)",
      "Tipo ordine \n(vd elenco)                                      ",
      "Cod. campagna\n(max 8 caratteri)",
      "EAN",
      "Copie ",
      "Sovrasc. occas.",
      "Pag. occas.",
      "N. ord. Cliente \n(max 30 caratteri)",
      "Tenuta prenotazioni S/N",
      "Data consegna tassativa \n(GG/MM/AAAA)",
      "motivo tassativa\n(vd elenco)",
      "Nota di testata da riportare in bolla (max 30 caratteri)",
      "Nuovo destinatario - Ragione sociale",
      "Nuovo destinatario - Indirizzo",
      "Nuovo destinatario - Località",
      "Nuovo destinatario - CAP",
      "Nuovo destinatario - Provincia",
      "Nuovo destinatario - Nazione"
    ];
    const rows = righe.map(r => [
      r.codice_cliente, "Anticipo lancio", "", r.ean, r.quantita,
      r.sconto_occasionale > 0 ? r.sconto_occasionale : "",
      r.pagamento_occasionale || "",
      r.numero_ordine || "",
      "", "", "", "", "", "", "", "", "", ""
    ]);
    const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
    ws["!cols"] = [10,15,10,14,8,12,12,20,10,20,15,30,25,25,20,8,8,8].map(w => ({ wch: w }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Ordini");
    XLSX.writeFile(wb, `Anticipi_Lancio_${new Date().toISOString().slice(0,10)}.xlsx`);
  };

  if (loading) return <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", color: T.textMid }}>Caricamento...</div>;

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%", overflow: "hidden" }}>
      {/* MODAL NUOVO ANTICIPO */}
      {showForm && (
        <div style={{ position: "fixed", inset: 0, background: "#000a", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center" }} onClick={() => setShowForm(false)}>
          <div style={{ background: T.surface, border: `1px solid ${T.borderHi}`, borderRadius: 6, padding: 28, width: 460 }} onClick={e => e.stopPropagation()}>
            <div style={{ color: T.accent, fontWeight: "700", fontSize: "13px", marginBottom: 20 }}>NUOVO ANTICIPO LANCIO</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
              {[
                ["codice_cliente", "Codice/i cliente *", "full"],
                ["ean", "EAN *"], ["quantita", "Quantità *"],
                ["sconto_occasionale", "Sconto occasionale (%)"], ["pagamento_occasionale", "Pagamento occasionale"],
                ["numero_ordine", "N° ordine", "full"],
              ].map(([k, label, span]) => (
                <div key={k} style={span === "full" ? { gridColumn: "1/-1" } : {}}>
                  <label style={{ color: T.textMid, fontSize: "10px", textTransform: "uppercase", letterSpacing: "0.08em", display: "block", marginBottom: 4 }}>{label}</label>
                  <input style={{ ...css.input, width: "100%", boxSizing: "border-box" }} value={form[k]} onChange={e => setForm(f => ({ ...f, [k]: e.target.value }))} placeholder={k === "codice_cliente" ? "es. 12345, 67890" : ""} />
                </div>
              ))}
            </div>
            <div style={{ marginTop: 12 }}>
              <label style={{ color: T.textMid, fontSize: "10px", textTransform: "uppercase", letterSpacing: "0.08em", display: "block", marginBottom: 4 }}>Note</label>
              <textarea style={{ ...css.input, width: "100%", boxSizing: "border-box", height: 60, resize: "vertical" }} value={form.note} onChange={e => setForm(f => ({ ...f, note: e.target.value }))} />
            </div>
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 20 }}>
              <button style={css.btn()} onClick={() => setShowForm(false)}>Annulla</button>
              <button style={css.btn("accent")} onClick={salvaNuovo} disabled={saving}>{saving ? "Salvataggio..." : "Aggiungi"}</button>
            </div>
          </div>
        </div>
      )}

      {/* TOOLBAR */}
      <div style={{ padding: "14px 20px", borderBottom: `1px solid ${T.border}`, display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
        {[["attivi","Attivi"],["da_gestire","Da gestire"],["notificato","Notificato"],["gestito","Gestiti"],["tutti","Tutti"]].map(([k, label]) => (
          <button key={k} style={css.btn(filterStato === k ? "accent" : "default")} onClick={() => setFilterStato(k)}>{label}</button>
        ))}
        <input style={{ ...css.input, width: 200 }} placeholder="Cerca cliente / EAN / ordine..." value={search} onChange={e => setSearch(e.target.value)} />
        <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
          <button style={css.btn()} onClick={exportMessaggerie}>↓ Export Messaggerie{selected.size > 0 ? ` (${selected.size})` : ""}</button>
          <button style={{ ...css.btn(), borderColor: T.green, color: T.green }} onClick={() => setShowForm(true)}>+ Nuovo</button>
        </div>
      </div>

      {/* KPI */}
      <div style={{ padding: "14px 20px", borderBottom: `1px solid ${T.border}`, display: "flex", gap: 12 }}>
        <KpiCard label="Da gestire" value={counts.da_gestire} color={T.textMid} />
        <KpiCard label="🔔 Notificati" value={counts.notificato} color="#e8a838" />
        <KpiCard label="Gestiti" value={counts.gestito} color={T.green} />
      </div>

      {/* TABELLA */}
      <div style={{ flex: 1, overflowY: "auto" }}>
        <table style={css.table}>
          <thead>
            <tr>
              <th style={{ ...css.th, width: 30 }}></th>
              <th style={css.th}>Stato</th>
              <th style={css.th}>Cliente/i</th>
              <th style={css.th}>EAN</th>
              <th style={css.th}>Titolo (se lanciato)</th>
              <th style={css.th}>Qtà</th>
              <th style={css.th}>Sconto occ.</th>
              <th style={css.th}>Pag. occ.</th>
              <th style={css.th}>N° ordine</th>
              <th style={css.th}>Note</th>
              <th style={css.th}></th>
            </tr>
          </thead>
          <tbody>
            {dataFiltrata.map((r, i) => (
              <tr key={r.id} style={{ background: r.stato === "notificato" ? "#e8a83822" : (i % 2 === 0 ? "transparent" : T.surface + "66") }}>
                <td style={css.td}><input type="checkbox" checked={selected.has(r.id)} onChange={() => toggleSelect(r.id)} style={{ accentColor: T.accent }} /></td>
                <td style={css.td}>
                  {r.stato === "gestito" && <Badge label="Gestito" color={T.green} />}
                  {r.stato === "notificato" && <Badge label="🔔 Notificato" color="#e8a838" />}
                  {r.stato === "da_gestire" && <Badge label="Da gestire" color={T.textMid} />}
                </td>
                <td style={{ ...css.td, fontWeight: "600" }}>{r.codice_cliente}</td>
                <td style={{ ...css.td, fontFamily: "monospace", fontSize: "11px", color: T.textMid }}>{r.ean}</td>
                <td style={{ ...css.td, color: T.accent, maxWidth: 200 }}>{r.lanci_settimanali ? `${r.lanci_settimanali.titolo} (lancio ${r.lanci_settimanali.num_lancio}/${r.lanci_settimanali.anno_lancio})` : <span style={{ color: T.textDim }}>—</span>}</td>
                <td style={css.td}>{r.quantita}</td>
                <td style={css.td}>{r.sconto_occasionale != null ? `${r.sconto_occasionale}%` : "—"}</td>
                <td style={css.td}>{r.pagamento_occasionale || "—"}</td>
                <td style={css.td}>{r.numero_ordine || "—"}</td>
                <td style={{ ...css.td, maxWidth: 160, fontSize: "11px", color: T.textMid, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={r.note}>{r.note || ""}</td>
                <td style={{ ...css.td, whiteSpace: "nowrap" }}>
                  <button title="Crea mail" style={{ ...css.btn(), padding: "2px 6px", fontSize: "11px", marginRight: 4 }} onClick={() => creaMail(r)}>✉️</button>
                  {r.stato !== "gestito" ? (
                    <button title="Segna gestito" style={{ ...css.btn(), padding: "2px 6px", fontSize: "11px", color: T.green, borderColor: T.green, marginRight: 4 }} onClick={() => segnaGestito(r.id)}>✓</button>
                  ) : (
                    <button title="Riapri" style={{ ...css.btn(), padding: "2px 6px", fontSize: "11px", marginRight: 4 }} onClick={() => riapri(r.id)}>↺</button>
                  )}
                  <button title="Elimina" style={{ ...css.btn(), padding: "2px 6px", fontSize: "11px", color: T.red, borderColor: T.red }} onClick={() => eliminaRiga(r.id)}>✕</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        {dataFiltrata.length === 0 && (
          <div style={{ padding: 40, textAlign: "center", color: T.textDim }}>Nessun anticipo lancio in questa vista.</div>
        )}
      </div>

      {toast && (
        <div style={{ position: "fixed", bottom: 24, left: "50%", transform: "translateX(-50%)", background: toast.type === "err" ? "#4a1a2a" : "#1a3a2a", border: `1px solid ${toast.type === "err" ? T.red : T.green}`, color: toast.type === "err" ? T.red : T.green, borderRadius: 6, padding: "8px 20px", fontSize: "12px", zIndex: 999, boxShadow: "0 4px 20px #0008" }}>
          {toast.msg}
        </div>
      )}
    </div>
  );
}

const MODULES = [
  { id: "dashboard", label: "Dashboard", icon: "◈" },
  { id: "calendariogiri", label: "Calendario Giri", icon: "📅" },
  { id: "cedola", label: "Giri e Cedole", icon: "≡" },
  { id: "finegiro", label: "Fine Giro", icon: "⊞" },
  { id: "avanzamento", label: "Avanzamento Novità", icon: "▣" },
  { id: "lanci", label: "Lanci Settimanali", icon: "🚀" },
  { id: "verificalanci", label: "Verifica Lanci Amazon", icon: "📦" },
  { id: "anticipilancio", label: "Anticipi Lancio", icon: "⏱" },
];

const MODULES_IMPORT = [
  { id: "import", label: "Import Cedola", icon: "↑" },
  { id: "prenotato", label: "Import Prenotato", icon: "↳" },
  { id: "spalmatura", label: "Import Pesi Spalmatura", icon: "⚖" },
];

const style = document.createElement('style');
style.textContent = `button:hover { filter: brightness(1.3); }`;
document.head.appendChild(style);

export default function App() {
  const [session, setSession] = useState(null);
  const [checkingAuth, setCheckingAuth] = useState(true);
  const [activeModule, setActiveModule] = useState("dashboard");
  const [titoli, setTitoli] = useState([]);
  const [prenotato, setPrenotato] = useState([]);
  const [spalmatura, setSpalmatura] = useState([]);
  const [canali, setCanali] = useState([]);
  const [giriDB, setGiriDB] = useState([]);
  const [ruolo, setRuolo] = useState("admin");
  const [userAccount, setUserAccount] = useState(null);

  // Applica alle righe titoli il ranking_editore AGGIORNATO (da ranking_editori), invece del
  // valore congelato al momento dell'import. Evita il disallineamento che causa titoli di
  // editori diversi ma con stesso ranking "vecchio" a intrecciarsi nell'ordinamento.
  const applyRankingLive = useCallback((rows, rankingMap) => rows.map(t => {
    const rLive = rankingMap[normEditoreKey(t.editore_nome)];
    return { ...t, account_editore: ACCOUNT_BY_COD[t.codice_editore] || t.account_editore || null, ranking_editore: rLive != null ? rLive : t.ranking_editore };
  }), []);

  useEffect(() => {
    if (!session) return;
    sbFetch("giri?select=*&order=anno.desc,numero.desc", session.token).then(setGiriDB);
    sbFetch("ranking_editori?select=editore_nome,ranking", session.token).then(reData => {
      const rankingMap = {};
      if (Array.isArray(reData)) reData.forEach(r => { rankingMap[normEditoreKey(r.editore_nome)] = Math.round(Number(r.ranking)); });
      sbFetch("titoli?select=*&order=ranking_editore.asc,ranking_titolo.asc", session.token).then(data => setTitoli(Array.isArray(data) ? applyRankingLive(data, rankingMap) : data));
    });
    sbFetch("prenotato?select=*&limit=100000", session.token).then(setPrenotato);
    sbFetch("spalmatura_obiettivo?select=*", session.token).then(setSpalmatura);
    sbFetch("canali?select=*&order=nome.asc", session.token).then(setCanali);
    sbFetch(`user_profiles?id=eq.${session.user.id}&select=ruolo,account_editore`, session.token).then(data => { if (Array.isArray(data) && data[0]) { setRuolo(data[0].ruolo); if (data[0].account_editore) setUserAccount(data[0].account_editore); } });
  }, [session]);

  useEffect(() => {
    const token = localStorage.getItem("giro_token");
    if (token) {
      sb.auth.getUser(token).then(user => {
        if (user?.email) setSession({ token, user });
        else localStorage.removeItem("giro_token");
        setCheckingAuth(false);
      });
    } else { setCheckingAuth(false); }
  }, []);

  const handleLogin = (token, user) => setSession({ token, user });
  const handleLogout = async () => { await sb.auth.signOut(session.token); localStorage.removeItem("giro_token"); setSession(null); };
  const updateTitolo = useCallback(updated => setTitoli(prev => prev.map(t => t.id === updated.id ? { ...t, ...updated } : t)), []);
  const deleteTitolo = useCallback(id => setTitoli(prev => prev.filter(t => t.id !== id)), []);
  const refreshDati = useCallback(() => {
    if (!session) return;
    sbFetch("prenotato?select=*&limit=100000", session.token).then(data => { if (Array.isArray(data)) setPrenotato(data); });
    sbFetch("ranking_editori?select=editore_nome,ranking", session.token).then(reData => {
      const rankingMap = {};
      if (Array.isArray(reData)) reData.forEach(r => { rankingMap[normEditoreKey(r.editore_nome)] = Math.round(Number(r.ranking)); });
      sbFetch("titoli?select=*&order=ranking_editore.asc,ranking_titolo.asc", session.token).then(data => { if (Array.isArray(data)) setTitoli(applyRankingLive(data, rankingMap)); });
    });
  }, [session, applyRankingLive]);

  if (checkingAuth) return <div style={{ ...css.app, display: "flex", alignItems: "center", justifyContent: "center", height: "100vh", color: T.textMid }}>Caricamento...</div>;
  if (!session) return <LoginScreen onLogin={handleLogin} />;

  const moduliVis = ruolo === "agente" ? MODULES.filter(m => ["dashboard","cedola","finegiro"].includes(m.id)) : MODULES;
  const moduliImport = ruolo !== "agente" ? MODULES_IMPORT : [];

  return (
    <div style={{ ...css.app, display: "flex", height: "100vh" }}>
      <div style={css.sidebar}>
        <div style={{ padding: "20px 16px 16px", borderBottom: `1px solid ${T.border}` }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
            <img src="https://raw.githubusercontent.com/albertospde/giro-manager/main/.github/logo_pde.png" style={{ height: 36, borderRadius: 6 }} alt="PDE" />
          </div>
          <div style={{ color: T.accent, fontSize: "13px", fontWeight: "700", letterSpacing: "0.06em" }}>GIRO MANAGER</div>
        </div>
        <nav style={{ flex: 1, padding: "8px 0", display: "flex", flexDirection: "column" }}>
          <div style={{ flex: 1 }}>
            {moduliVis.map(m => (
              <button key={m.id} style={{ width: "100%", textAlign: "left", padding: "10px 16px", border: "none", background: activeModule === m.id ? T.accent + "18" : "transparent", color: activeModule === m.id ? T.accent : T.textMid, cursor: "pointer", fontFamily: "inherit", fontSize: "12px", borderLeft: `2px solid ${activeModule === m.id ? T.accent : "transparent"}`, display: "flex", alignItems: "center", gap: 10, letterSpacing: "0.04em" }}
                onClick={() => setActiveModule(m.id)}>
                <span style={{ fontSize: "14px" }}>{m.icon}</span> {m.label}
              </button>
            ))}
          </div>
          {moduliImport.length > 0 && (
            <div>
              <div style={{ borderTop: `1px solid ${T.border}`, margin: "8px 0", padding: "6px 16px 2px" }}>
                <span style={{ color: T.textDim, fontSize: "9px", letterSpacing: "0.1em", textTransform: "uppercase" }}>Import</span>
              </div>
              {moduliImport.map(m => (
                <button key={m.id} style={{ width: "100%", textAlign: "left", padding: "8px 16px", border: "none", background: activeModule === m.id ? T.accent + "18" : "transparent", color: activeModule === m.id ? T.accent : T.textDim, cursor: "pointer", fontFamily: "inherit", fontSize: "11px", borderLeft: `2px solid ${activeModule === m.id ? T.accent : "transparent"}`, display: "flex", alignItems: "center", gap: 10, letterSpacing: "0.04em" }}
                  onClick={() => setActiveModule(m.id)}>
                  <span style={{ fontSize: "13px" }}>{m.icon}</span> {m.label}
                </button>
              ))}
            </div>
          )}
        </nav>
        <div style={{ padding: "12px 16px", borderTop: `1px solid ${T.border}` }}>
          <div style={{ color: T.textDim, fontSize: "10px", marginBottom: 2 }}>{session.user?.email}</div>
          <div style={{ color: T.textDim, fontSize: "10px", marginBottom: 6, textTransform: "uppercase", letterSpacing: "0.06em" }}>{ruolo}</div>
          <button style={{ ...css.btn(), fontSize: "11px", padding: "4px 10px", width: "100%", marginBottom: 4 }} onClick={async () => {
            const nuova = prompt("Nuova password (min. 6 caratteri):"); if (!nuova || nuova.length < 6) { alert("Password troppo corta."); return; }
            const conferma = prompt("Conferma nuova password:"); if (nuova !== conferma) { alert("Le password non coincidono."); return; }
            const r = await fetch(`${SUPABASE_URL}/auth/v1/user`, { method: "PUT", headers: { "Content-Type": "application/json", "apikey": SUPABASE_KEY, "Authorization": `Bearer ${session.token}` }, body: JSON.stringify({ password: nuova }) });
            if (r.ok) alert("Password aggiornata."); else alert("Errore aggiornamento password.");
          }}>Modifica password</button>
          <button style={{ ...css.btn(), fontSize: "11px", padding: "4px 10px", width: "100%" }} onClick={handleLogout}>Esci</button>
        </div>
      </div>
      <div style={css.main}>
        <div style={css.header}>
          <span style={{ color: T.accent, fontSize: "11px", letterSpacing: "0.12em", textTransform: "uppercase" }}>{[...MODULES, ...MODULES_IMPORT].find(m => m.id === activeModule)?.label}</span>
          <span style={{ color: T.borderHi }}>·</span>
          <span style={{ color: T.textMid, fontSize: "11px" }}>{titoli.length} titoli · {[...new Set(titoli.map(t => t.n_cedola).filter(Boolean))].length} cedole</span>
          {ruolo !== "agente" && <button style={{ ...css.btn(), marginLeft: "auto", fontSize: "11px", padding: "4px 10px" }} onClick={refreshDati}>↺ Aggiorna</button>}
        </div>
        <div style={{ flex: 1, overflow: "hidden", display: "flex", flexDirection: "column" }}>
          {/* MOD 3: Passato spalmatura alla Dashboard */}
          {activeModule === "dashboard" && <ModuloDashboard titoli={titoli} prenotato={prenotato} canali={canali} spalmatura={spalmatura} ruolo={ruolo} />}
          {activeModule === "calendariogiri" && <Modulocalendariogiri token={session.token} ruolo={ruolo} />}
          {/* FIX 5: Passato token a ModuloCedola */}
          {activeModule === "cedola" && <ModuloCedola titoli={titoli} giriList={giriDB} onUpdateTitolo={t => { updateTitolo(t); setTitoli(prev => prev.some(x => x.id === t.id) ? prev.map(x => x.id === t.id ? t : x) : [...prev, t]); }} onDeleteTitolo={deleteTitolo} spalmatura={spalmatura} prenotato={prenotato} ruolo={ruolo} token={session.token} onTitoliChange={refreshDati} userAccount={userAccount} />}
          {activeModule === "prenotato" && <ModuloPrenotato token={session.token} titoli={titoli} onImportDone={() => sbFetch("prenotato?select=*&limit=100000", session.token).then(setPrenotato)} />}
          {/* MOD 4: Passato spalmatura a ModuloFineGiro */}
          {activeModule === "finegiro" && <ModuloFineGiro titoli={titoli} prenotato={prenotato} canali={canali} token={session.token} ruolo={ruolo} spalmatura={spalmatura} userAccount={userAccount} />}
          {activeModule === "avanzamento" && <ModuloAvanzamento token={session.token} titoli={titoli} prenotato={prenotato} ruolo={ruolo} userAccount={userAccount} />}
          {activeModule === "lanci" && <ModuloLanciSettimanali token={session.token} titoli={titoli} prenotato={prenotato} canali={canali} ruolo={ruolo} userAccount={userAccount} onNavigateAnticipi={() => setActiveModule("anticipilancio")} />}
          {activeModule === "verificalanci" && <ModuloVerificaLanciAmazon token={session.token} titoli={titoli} prenotato={prenotato} canali={canali} />}
          {activeModule === "anticipilancio" && <ModuloAnticipiLancio token={session.token} userEmail={session.user?.email} />}
          {activeModule === "import" && <ModuloImport giriList={giriDB} token={session.token} />}
          {activeModule === "spalmatura" && <ImportSpalmatura token={session.token} onImportDone={() => sbFetch("spalmatura_obiettivo?select=*", session.token).then(setSpalmatura)} />}
        </div>
      </div>
    </div>
  );
}
