import { useState, useEffect, useMemo, useCallback } from "react";
import ModuloImport from "./ModuloImport.jsx";
import ModuloPrenotato from "./ModuloPrenotato.jsx";
import ModuloAvanzamento from "./ModuloAvanzamento.jsx";

const SUPABASE_URL = "https://tdflwenlylhctxssatax.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InRkZmx3ZW5seWxoY3R4c3NhdGF4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYzMzgyNzYsImV4cCI6MjA5MTkxNDI3Nn0.l35qEL7LOvyYuI1McQlVqj4vbyTqmlevcmqWbTGYi2Q";

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
function EditModal({ titolo, onSave, onClose, token }) {
  const [form, setForm] = useState({ ...titolo });
  const [saving, setSaving] = useState(false);
  const set = k => e => setForm(f => ({ ...f, [k]: e.target.type === "checkbox" ? e.target.checked : e.target.value }));

  const handleSave = async () => {
    setSaving(true);
    // Prepara l'oggetto con solo i campi modificabili
    const payload = {};
    const editableFields = [
      "titolo","autore","editore_nome","ean","prezzo","uscita","formato","eta",
      "account_editore","promozione","obiettivo_assegnato","obiettivo_raggiunto",
      "il_triangolo","top_100","ean_gemello_1","titolo_gemello_1","ean_gemello_2",
      "titolo_gemello_2","ean_gemello_3","titolo_gemello_3","note_comunicazione","note"
    ];
    const numFields = ["prezzo","obiettivo_assegnato","obiettivo_raggiunto"];
    editableFields.forEach(k => {
      let valForm = form[k];
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
      const ok = await sbUpdateTitolo(form.id, payload, token);
      if (!ok) { alert("Errore nel salvataggio su database."); setSaving(false); return; }
    }
    // Normalizza i tipi numerici prima di aggiornare lo state React
    const formNorm = { ...form };
    ["prezzo","obiettivo_assegnato","obiettivo_raggiunto"].forEach(k => {
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
          {[["titolo","Titolo","full"],["autore","Autore"],["editore_nome","Editore"],["ean","EAN"],["prezzo","Prezzo"],["uscita","Uscita"],["formato","Formato"],["eta","Età"],["account_editore","Account"],["promozione","Promozione"],["obiettivo_assegnato","Obiettivo assegnato"],["obiettivo_raggiunto","Obiettivo raggiunto"]].map(([k, label, span]) => (
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
        <div style={{ marginTop: 20, display: "flex", justifyContent: "flex-end", gap: 8 }}>
          <button style={css.btn()} onClick={onClose}>Annulla</button>
          <button style={css.btn("accent")} onClick={handleSave} disabled={saving}>
            {saving ? "Salvataggio..." : "Salva"}
          </button>
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

// FIX 5: ModuloCedola ora riceve token per passarlo alla EditModal

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

function ModuloCedola({ titoli, giriList, onUpdateTitolo, spalmatura, prenotato, ruolo, token, onTitoliChange }) {
  const [giroLabelSel, setGiroLabelSel] = useState([]);
  const [giroSel, setGiroSel] = useState([]);
  const [search, setSearch] = useState("");
  const [filterFlag, setFilterFlag] = useState("tutti");
  const [filterEditori, setFilterEditori] = useState([]);
  const [filterAccount, setFilterAccount] = useState([]);
  const [editingId, setEditingId] = useState(null);
  const [sortKey, setSortKey] = useState("n_cedola");
  const [showNuovoGiro, setShowNuovoGiro] = useState(false);
  const [showNuovoTitolo, setShowNuovoTitolo] = useState(false);
  const [toastCedola, setToastCedola] = useState(null);
  const showToastCedola = (msg, type = "ok") => { setToastCedola({ msg, type }); setTimeout(() => setToastCedola(null), 3000); };

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
  const emptyTitolo = { ean: "", titolo: "", autore: "", editore_nome: "", codice_editore: "", prezzo: "", uscita: "", formato: "Cover", n_cedola: "", giro_label: "", obiettivo_assegnato: "", account_editore: "", note_comunicazione: "", note: "" };
  const [formTitolo, setFormTitolo] = useState(emptyTitolo);
  const [savingTitolo, setSavingTitolo] = useState(false);
  const giriLabelAll = useMemo(() => [...new Set(titoli.map(t => t.giro_label).filter(Boolean))].sort((a, b) => { const [na, ya] = a.split(" "); const [nb, yb] = b.split(" "); return Number(yb) - Number(ya) || Number(nb) - Number(na); }), [titoli]);

  const saveNuovoTitolo = async () => {
    if (!formTitolo.ean || !formTitolo.titolo || !formTitolo.n_cedola || !formTitolo.giro_label) {
      showToastCedola("EAN, titolo, cedola e giro sono obbligatori", "err"); return;
    }
    setSavingTitolo(true);
    const payload = {
      ean: formTitolo.ean.trim(),
      titolo: formTitolo.titolo,
      autore: formTitolo.autore || null,
      editore_nome: formTitolo.editore_nome || null,
      codice_editore: formTitolo.codice_editore || null,
      prezzo: parseFloat(String(formTitolo.prezzo).replace(",", ".")) || null,
      uscita: formTitolo.uscita || null,
      formato: formTitolo.formato || "Cover",
      n_cedola: formTitolo.n_cedola,
      giro_label: formTitolo.giro_label,
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
      showToastCedola(`"${formTitolo.titolo}" aggiunto`);
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

  const giriLabel = useMemo(() => [...new Set(titoli.map(t => t.giro_label).filter(Boolean))]
    .filter(g => filterAnnoCedola.length === 0 || filterAnnoCedola.includes(Number(g.split(" ")[1])))
    .sort((a, b) => { const [na, ya] = a.split(" "); const [nb, yb] = b.split(" "); return Number(yb) - Number(ya) || Number(nb) - Number(na); }), [titoli, filterAnnoCedola]);
  const cedole = useMemo(() => { const t = giroLabelSel.length === 0 ? titoli.filter(t => filterAnnoCedola.length === 0 || filterAnnoCedola.includes(Number((t.giro_label||"").split(" ")[1]))) : titoli.filter(t => giroLabelSel.includes(t.giro_label)); return [...new Set(t.map(t => t.n_cedola).filter(Boolean))].sort(); }, [titoli, giroLabelSel, filterAnnoCedola]);
  const accounts = useMemo(() => { const t = giroSel.length === 0 ? (giroLabelSel.length === 0 ? titoli : titoli.filter(t => giroLabelSel.includes(t.giro_label))) : titoli.filter(t => giroSel.includes(t.n_cedola)); return [...new Set(t.map(t => t.account_editore).filter(Boolean))].sort(); }, [titoli, giroLabelSel, giroSel]);
  const editori = useMemo(() => { const t = giroSel.length === 0 ? (giroLabelSel.length === 0 ? titoli : titoli.filter(t => giroLabelSel.includes(t.giro_label))) : titoli.filter(t => giroSel.includes(t.n_cedola)); return [...new Set(t.map(t => t.editore_nome).filter(Boolean))].sort(); }, [titoli, giroLabelSel, giroSel]);

  const filtered = useMemo(() => titoli
    .filter(t => giroLabelSel.length === 0 || giroLabelSel.includes(t.giro_label))
    .filter(t => giroSel.length === 0 || giroSel.includes(t.n_cedola))
    .filter(t => filterEditori.length === 0 || filterEditori.includes(t.editore_nome))
    .filter(t => filterAccount.length === 0 || filterAccount.includes(t.account_editore))
    .filter(t => { if (!search) return true; const q = search.toLowerCase(); return t.titolo?.toLowerCase().includes(q) || t.autore?.toLowerCase().includes(q) || t.editore_nome?.toLowerCase().includes(q) || t.ean?.includes(q); })
    .filter(t => { if (filterFlag === "triangolo") return t.il_triangolo; if (filterFlag === "top100") return t.top_100; if (filterFlag === "gemelli") return t.ean_gemello_1; return true; })
    .sort((a, b) => { if (sortKey === "n_cedola") return (a.n_cedola ?? "").localeCompare(b.n_cedola ?? ""); if (sortKey === "editore") return (a.editore_nome ?? "").localeCompare(b.editore_nome ?? ""); if (sortKey === "prezzo") return (b.prezzo ?? 0) - (a.prezzo ?? 0); return 0; }),
  [titoli, giroLabelSel, giroSel, search, filterFlag, filterEditori, filterAccount, sortKey]);

  const editingTitolo = titoli.find(t => t.id === editingId);

  const getObjCanale = (titolo, canale_codice) => {
    const spRow = spalmatura.find(s => s.editore_nome === titolo.editore_nome && s.formato === (titolo.formato || 'Cover') && s.canale_codice === canale_codice);
    if (!spRow || !titolo.obiettivo_assegnato) return 0;
    return Math.round(titolo.obiettivo_assegnato * spRow.percentuale);
  };

  const exportAgenti = () => {
    const XLSX = window.XLSX;
    const headers = ["N° CEDOLA","EAN","TITOLO","AUTORE","COD.EDITORE","EDITORE","PREZZO","USCITA","NOTE","EAN GEM 1","TITOLO GEM 1","EAN GEM 2","TITOLO GEM 2","EAN GEM 3","TITOLO GEM 3","OBJ INDIPENDENTI & ALTRE CATENE"];
    const rows = filtered.map(t => [t.n_cedola, t.ean, t.titolo, t.autore, t.codice_editore, t.editore_nome, t.prezzo, t.uscita, t.note_comunicazione || t.note, t.ean_gemello_1, t.titolo_gemello_1, t.ean_gemello_2, t.titolo_gemello_2, t.ean_gemello_3, t.titolo_gemello_3, getObjCanale(t, 'INDIPENDENTI_ALTRE_CATENE')]);
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "CEDOLA AGENTI");
    XLSX.writeFile(wb, `CEDOLA_AGENTI_${giroLabelSel.length === 0 ? "TUTTI" : giroLabelSel.join("-")}.xlsx`);
  };

  const exportDirezionale = () => {
    const pwd = prompt("Password direzionale:"); if (pwd !== "nuovaluce") { alert("Password errata."); return; }
    const XLSX = window.XLSX;
    const canaliDir = [{ codice: 'FELTRINELLI', label: 'Feltrinelli' },{ codice: 'GIUNTI', label: 'Giunti' },{ codice: 'MONDADORI', label: 'Mondadori' },{ codice: 'UBIK', label: 'Ubik' },{ codice: 'INDIPENDENTI_ALTRE_CATENE', label: 'Indip. & Altre Catene' },{ codice: 'AMAZON', label: 'Amazon' },{ codice: 'IBS', label: 'IBS' },{ codice: 'ALTRI_ONLINE', label: 'Altri Online' },{ codice: 'FASTBOOK', label: 'Fastbook' },{ codice: 'GROSSISTI', label: 'Grossisti' },{ codice: 'CENTROLIBRI', label: 'Centrolibri' },{ codice: 'GDO', label: 'GDO' }];
    const headersCedola = ["N° CEDOLA","EAN","TITOLO","AUTORE","COD.EDITORE","EDITORE","PREZZO","OBJ TOTALE","NOTE","EAN GEM 1","TITOLO GEM 1","EAN GEM 2","TITOLO GEM 2","EAN GEM 3","TITOLO GEM 3"];
    const rowsCedola = filtered.map(t => [t.n_cedola, t.ean, t.titolo, t.autore, t.codice_editore, t.editore_nome, t.prezzo, t.obiettivo_assegnato || 0, t.note_comunicazione || t.note, t.ean_gemello_1, t.titolo_gemello_1, t.ean_gemello_2, t.titolo_gemello_2, t.ean_gemello_3, t.titolo_gemello_3]);
    const wsCedola = XLSX.utils.aoa_to_sheet([headersCedola, ...rowsCedola]);
    const headersObj = ["N° CEDOLA","EAN","TITOLO","AUTORE","COD.EDITORE","EDITORE","PREZZO","OBJ TOTALE",...canaliDir.map(c => c.label)];
    const rowsObj = filtered.map(t => [t.n_cedola, t.ean, t.titolo, t.autore, t.codice_editore, t.editore_nome, t.prezzo, t.obiettivo_assegnato || 0, ...canaliDir.map(c => getObjCanale(t, c.codice))]);
    const wsObj = XLSX.utils.aoa_to_sheet([headersObj, ...rowsObj]);
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, wsCedola, "CEDOLA"); XLSX.utils.book_append_sheet(wb, wsObj, "OBIETTIVI");
    XLSX.writeFile(wb, `CEDOLA_DIREZIONALE_${giroLabelSel.length === 0 ? "TUTTI" : giroLabelSel.join("-")}.xlsx`);
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%", overflow: "hidden" }}>
      {editingTitolo && <EditModal titolo={editingTitolo} onSave={onUpdateTitolo} onClose={() => setEditingId(null)} token={token} />}

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

      {/* TOAST LOCALE */}
      {toastCedola && (
        <div style={{ position: "fixed", bottom: 24, left: "50%", transform: "translateX(-50%)", background: toastCedola.type === "err" ? "#4a1a2a" : "#1a3a2a", border: `1px solid ${toastCedola.type === "err" ? T.red : T.green}`, color: toastCedola.type === "err" ? T.red : T.green, borderRadius: 6, padding: "8px 20px", fontSize: "12px", zIndex: 999, boxShadow: "0 4px 20px #0008" }}>
          {toastCedola.msg}
        </div>
      )}
      <div style={{ padding: "12px 20px", borderBottom: `1px solid ${T.border}`, display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
        <SearchableMultiSelect values={filterAnnoCedola.map(String)} onChange={v => { setFilterAnnoCedola(v.map(Number)); setGiroLabelSel([]); setGiroSel([]); setFilterEditori([]); }} options={anniDispCedola.map(String)} renderOption={v => v} placeholder="Anno" width={110} />
        <SearchableMultiSelect values={giroLabelSel} onChange={v => { setGiroLabelSel(v); setGiroSel([]); setFilterEditori([]); }} options={giriLabel} renderOption={g => `Giro ${g}`} placeholder="Tutti i giri" width={170} />
        <SearchableMultiSelect values={giroSel} onChange={v => { setGiroSel(v); setFilterEditori([]); }} options={cedole} placeholder="Tutte le cedole" width={160} />
        <SearchableMultiSelect values={filterEditori} onChange={setFilterEditori} options={editori} placeholder="Tutti gli editori" width={190} />
        <SearchableMultiSelect values={filterAccount} onChange={setFilterAccount} options={accounts} placeholder="Tutti gli account" width={160} />
        <input style={{ ...css.input, width: 180 }} placeholder="Cerca..." value={search} onChange={e => setSearch(e.target.value)} />
        {["tutti","triangolo","top100","gemelli"].map(f => (
          <button key={f} style={{ ...css.btn(filterFlag === f ? "accent" : "default"), padding: "5px 10px" }} onClick={() => setFilterFlag(f)}>
            {f === "tutti" ? "Tutti" : f === "triangolo" ? "▲" : f === "top100" ? "★" : "Gem."}
          </button>
        ))}
        <div style={{ marginLeft: "auto", color: T.textMid, fontSize: "11px" }}><span style={{ color: T.text }}>{filtered.length}</span> titoli</div>
        {ruolo !== "agente" && <>
          <button style={{ ...css.btn(), borderColor: T.green, color: T.green }} onClick={() => setShowNuovoGiro(true)}>+ Giro</button>
          <button style={{ ...css.btn(), borderColor: T.green, color: T.green }} onClick={() => setShowNuovoTitolo(true)}>+ Titolo</button>
        </>}
        <button style={css.btn("accent")} onClick={exportAgenti}>↓ Download Agenti</button>
        {ruolo !== "agente" && <button style={css.btn("accent")} onClick={exportDirezionale}>↓ Download Direzionali</button>}
      </div>
      <div style={{ flex: 1, overflowY: "auto" }}>
        <table style={css.table}>
          <thead>
            <tr>
              <th style={css.th} onClick={() => setSortKey("n_cedola")}>Cedola{sortKey === "n_cedola" ? " ↓" : ""}</th>
              <th style={css.th}>EAN</th><th style={css.th}>Titolo</th><th style={css.th}>Autore</th><th style={css.th}>Cod.Ed.</th>
              <th style={css.th} onClick={() => setSortKey("editore")}>Editore{sortKey === "editore" ? " ↓" : ""}</th>
              <th style={css.th} onClick={() => setSortKey("prezzo")}>€{sortKey === "prezzo" ? " ↓" : ""}</th>
              <th style={css.th}>Obj</th><th style={css.th}>Flag</th><th style={css.th}>Gemelli</th><th style={css.th}>Note</th><th style={css.th}></th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((t, i) => (
              <tr key={t.id} style={{ background: i % 2 === 0 ? "transparent" : T.surface + "66" }}>
                <td style={{ ...css.td, color: T.textMid, fontSize: "10px", whiteSpace: "nowrap" }}>{t.n_cedola}</td>
                <td style={{ ...css.td, color: T.textDim, fontFamily: "monospace", fontSize: "11px" }}>{t.ean}</td>
                <td style={{ ...css.td, maxWidth: 260 }}><div style={{ fontWeight: "600", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{t.titolo}</div></td>
                <td style={{ ...css.td, color: T.textMid }}>{t.autore}</td>
                <td style={{ ...css.td, color: T.textMid, fontSize: "11px" }}>{t.codice_editore}</td>
                <td style={{ ...css.td, color: T.accent, fontWeight: "600", whiteSpace: "nowrap" }}>{t.editore_nome}</td>
                <td style={{ ...css.td, whiteSpace: "nowrap" }}>€ {t.prezzo?.toFixed(2)}</td>
                <td style={{ ...css.td, whiteSpace: "nowrap", minWidth: 120 }}>
                  {(() => { const pren = prenotato.filter(p => p.titolo_id === t.id).reduce((s,p) => s+p.quantita, 0); return (<><div style={{ fontSize: "11px", marginBottom: 4 }}>{pren.toLocaleString("it")} / {(t.obiettivo_assegnato||0).toLocaleString("it")}</div><ProgressBar value={pren} total={t.obiettivo_assegnato} /></>); })()}
                </td>
                <td style={css.td}><div style={{ display: "flex", gap: 4 }}>{t.il_triangolo && <Badge label="▲" color={T.purple} />}{t.top_100 && <Badge label="★" color={T.accent} />}</div></td>
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

function ModuloFineGiro({ titoli, prenotato, canali, token, ruolo, spalmatura }) {
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

  const cedoleExtra = useMemo(() => [...new Set(titoli.filter(t => t.giro_label === "EXTRA").map(t => t.n_cedola).filter(Boolean))].sort(), [titoli]);

  const [giroLabelSel, setGiroLabelSel] = useState([]);
  const [extraSel, setExtraSel] = useState([]);
  const [cedolaSel, setCedolaSel] = useState([]);
  const [filterEditori, setFilterEditori] = useState([]);
  const [filterAccount, setFilterAccount] = useState([]);
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
  const [clienteRegioneMap, setClienteRegioneMap] = useState({});
  const [filterRegioni, setFilterRegioni] = useState([]);

  // Auto-seleziona il giro più recente all'avvio (giriLabel è ordinato DESC → [0] è il più recente)
  useEffect(() => {
    if (giriLabel.length > 0 && giroLabelSel.length === 0 && extraSel.length === 0) {
      setGiroLabelSel([giriLabel[0]]);
    }
  }, [giriLabel]);

  const resetFiltri = () => { setGiroLabelSel([]); setExtraSel([]); setCedolaSel([]); setFilterEditori([]); setFilterAccount([]); setFilterCanale([]); setFilterRegioni([]); setSearch(""); setClienteSel(null); setSoloPrenotati(false); };

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
    if (!token) return;
    fetch(`${SUPABASE_URL}/rest/v1/clienti?select=codice_cliente,regione&limit=100000`, {
      headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`, "Range": "0-499999" }
    }).then(r => r.json()).then(data => {
      if (!Array.isArray(data)) return;
      const map = {};
      data.forEach(r => { map[r.codice_cliente] = r.regione; });
      setClienteRegioneMap(map);
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

  const regioniDisponibili = useMemo(() => {
    return [...new Set(Object.values(clienteRegioneMap))].filter(Boolean).sort();
  }, [clienteRegioneMap]);

  const titoliSel = useMemo(() => {
    if (giroLabelSel.length === 0 && extraSel.length === 0) return [];
    return titoli
      .filter(t => extraSel.length > 0 ? (t.giro_label === "EXTRA" && extraSel.includes(t.n_cedola)) : giroLabelSel.includes(t.giro_label))
      .filter(t => cedolaSel.length === 0 || cedolaSel.includes(t.n_cedola))
      .filter(t => filterEditori.length === 0 || filterEditori.includes(t.editore_nome))
      .filter(t => filterAccount.length === 0 || filterAccount.includes(t.account_editore))
      .filter(t => { if (!search) return true; const q = search.toLowerCase(); return t.titolo?.toLowerCase().includes(q) || t.ean?.includes(q); })
      .sort((a, b) => (a.ranking_editore ?? 99) - (b.ranking_editore ?? 99) || (a.ranking_titolo ?? 99) - (b.ranking_titolo ?? 99));
  }, [titoli, giroLabelSel, extraSel, cedolaSel, filterEditori, filterAccount, search]);

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
    // Dir. Stampatore: valore editabile manuale, separato
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
    if (filterRegioni.length > 0 && clienteSel) {
      const regioneCliente = clienteRegioneMap[clienteSel];
      if (!regioneCliente || !filterRegioni.includes(regioneCliente)) return [];
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

  const exportExcel = () => {
    const XLSX = window.XLSX;
    const colCanali = canaliTabella;
    const dirStampHeaders = ruolo !== "agente" ? ["DIR. STAMPATORE"] : [];
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
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "FINE GIRO");
    XLSX.utils.book_append_sheet(wb, wsRiepilogo, "RIEPILOGO");
    XLSX.writeFile(wb, `FINE_GIRO_${giroLabelSel.length > 0 ? giroLabelSel.join("-") : extraSel.length > 0 ? extraSel.join("-") : "EXPORT"}.xlsx`);
  };

  if (giroLabelSel.length === 0 && extraSel.length === 0) {
    return (
      <div style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 24 }}>
        <div style={{ color: T.textMid, fontSize: "14px" }}>Seleziona un giro o una cedola extra</div>
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap", justifyContent: "center" }}>
          <div style={{ marginBottom: 16, display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
            <SearchableMultiSelect values={filterAnnoFineGiro.map(String)} onChange={v => { setFilterAnnoFineGiro(v.map(Number)); }} options={anniDispFineGiro.map(String)} renderOption={v => v} placeholder="Anno" width={110} />
          </div>
          {giriLabel.map(g => <button key={g} style={{ ...css.btn("accent"), padding: "10px 20px", fontSize: "13px" }} onClick={() => setGiroLabelSel([g])}>Giro {g}</button>)}
          {cedoleExtra.map(c => <button key={c} style={{ ...css.btn(), padding: "10px 20px", fontSize: "13px", borderColor: T.accent, color: T.accent }} onClick={() => setExtraSel([c])}>{c}</button>)}
        </div>
      </div>
    );
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%", overflow: "hidden" }}>
      {/* Toolbar filtri */}
      <div style={{ padding: "12px 20px", borderBottom: `1px solid ${T.border}`, display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
        <SearchableMultiSelect values={filterAnnoFineGiro.map(String)} onChange={v => { setFilterAnnoFineGiro(v.map(Number)); setGiroLabelSel([]); setExtraSel([]); setCedolaSel([]); setFilterEditori([]); setFilterAccount([]); setFilterCanale([]); setSearch(""); setClienteSel(null); setSoloPrenotati(false); }} options={anniDispFineGiro.map(String)} renderOption={v => v} placeholder="Anno" width={110} />
        <SearchableMultiSelect values={giroLabelSel} onChange={v => { setGiroLabelSel(v); setExtraSel([]); setCedolaSel([]); setFilterEditori([]); setFilterAccount([]); setFilterCanale([]); setSearch(""); setClienteSel(null); setSoloPrenotati(false); }} options={giriLabel} placeholder="— Giro —" width={160} renderOption={g => `Giro ${g}`} />
        {cedoleExtra.length > 0 && (
          <SearchableMultiSelect values={extraSel} onChange={v => { setExtraSel(v); setGiroLabelSel([]); setCedolaSel([]); setFilterEditori([]); setFilterAccount([]); setFilterCanale([]); setSearch(""); setClienteSel(null); setSoloPrenotati(false); }} options={cedoleExtra} placeholder="Cedole Extra" width={170} />
        )}
        {giroLabelSel.length > 0 && (
          <SearchableMultiSelect values={cedolaSel} onChange={setCedolaSel} options={cedole} placeholder="Tutte le cedole" width={160} />
        )}
        <SearchableMultiSelect values={filterEditori} onChange={setFilterEditori} options={editori} placeholder="Tutti gli editori" width={190} />
        <SearchableMultiSelect values={filterAccount} onChange={setFilterAccount} options={accounts} placeholder="Tutti gli account" width={160} />
        <SearchableMultiSelect values={filterCanale} onChange={setFilterCanale} options={canali.filter(c => c.codice !== "AURORA" && c.codice !== "GDO").map(c => c.codice)} renderOption={cod => { const c = canali.find(x => x.codice === cod); return c?.nome || cod; }} placeholder="Tutti i canali" width={170} />
        <SearchableMultiSelect values={filterRegioni} onChange={setFilterRegioni} options={regioniDisponibili} placeholder="Tutte le regioni" width={170} />
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
                    <div style={{ fontSize: "10px", color: T.textMid }}>{c.cod}{clienteRegioneMap[c.cod] ? ` · ${clienteRegioneMap[c.cod]}` : ""}</div>
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
              {ruolo !== "agente" && <th style={{ ...css.th, color: T.accent }}>Dir. Stampatore ✎</th>}
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



function ModuloLanciSettimanali({ token, titoli, prenotato, canali, ruolo, userAccount }) {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [uploading, setUploading] = useState(false);
  const [uploadMode, setUploadMode] = useState(null);
  const [toast, setToast] = useState(null);
  const [filterAnno, setFilterAnno] = useState(null);
  const [filterLancio, setFilterLancio] = useState([]);
  const [filterAccount, setFilterAccount] = useState([]);
  const [search, setSearch] = useState("");
  const [sortKey, setSortKey] = useState(null);
  const [sortDir, setSortDir] = useState("desc");

  const [editCell, setEditCell] = useState(null); // { id, field, value }

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

        const deltaPortale = (r.prenotato_trasmesso ?? r.prenotato_iscrizione ?? 0) - prenSenzaAmazon;

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
    if (sortKey) {
      result = [...result].sort((a, b) => {
        const va = a[sortKey] ?? 0, vb = b[sortKey] ?? 0;
        return sortDir === "asc" ? (typeof va === "string" ? String(va).localeCompare(String(vb)) : va - vb) : (typeof va === "string" ? String(vb).localeCompare(String(va)) : vb - va);
      });
    }
    return result;
  }, [dataArricchita, search, sortKey, sortDir, filterAccount]);

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
          };
        }).filter(Boolean);
        if (payload.length === 0) throw new Error("Nessun EAN valido");
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
            body: JSON.stringify({ prenotato_trasmesso: qta }),
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
    const headers = ["EAN","COD.ED.","EDITORE","ACCOUNT","TITOLO","AUTORE","PREZZO","CEDOLE","LANCIATE","TRASMESSO","FINE GIRO","AMAZON","FG NO AMAZON","TEORICO","Δ PORTALE","GIORNO USCITA"];
    const rows = dataFiltrata.map(r => [
      r.ean, r.codice_editore, r.editore, r.account_editore || "", r.titolo, r.autore, r.prezzo,
      r.cedole.join(", "), r.prenotato_iscrizione, r.prenotato_trasmesso ?? "",
      r.pren_fine_giro, r.pren_amazon, r.pren_senza_amazon, r.teorico, r.delta_portale, r.giorno_uscita
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
          <span style={css.tag(T.accent)}>MAR {kpi.martedi}</span>
          <span style={css.tag(T.green)}>VEN {kpi.venerdi}</span>
        </div>
        <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
          <label style={{ ...css.btn("accent"), cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 4 }}>
            {uploading && uploadMode === "iscrizione" ? "..." : "↑ Lancio"}
            <input type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={(e) => handleUpload(e, "iscrizione")} disabled={uploading} />
          </label>
          <label style={{ ...css.btn(), cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 4, borderColor: T.green, color: T.green }}>
            {uploading && uploadMode === "trasmesso" ? "..." : "↑ Trasmesso"}
            <input type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={(e) => handleUpload(e, "trasmesso")} disabled={uploading} />
          </label>
          <button style={css.btn()} onClick={exportExcel}>↓ Excel</button>
        </div>
      </div>

      {/* KPI */}
      <div style={{ padding: "16px 20px", borderBottom: `1px solid ${T.border}` }}>
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 16 }}>
          <KpiCard label="Titoli al lancio" value={kpi.tot} color={T.text} />
          <KpiCard label="Copie lanciate" value={kpi.copieLanciate.toLocaleString("it")} color={T.accent} sub={`€ ${kpi.valoreLanciate.toLocaleString("it", { maximumFractionDigits: 0 })}`} />
          {kpi.haTrasmesso && (
            <>
              <KpiCard label="Copie trasmesse" value={kpi.copieTrasmesse.toLocaleString("it")} color={T.green} sub={`€ ${kpi.valoreTrasmesso.toLocaleString("it", { maximumFractionDigits: 0 })}`} />
              <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 4, padding: "16px 20px", minWidth: 180 }}>
                <div style={{ color: T.textMid, fontSize: "10px", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 8 }}>Copertura trasmesso</div>
                <div style={{ color: kpi.pctTrasmesso >= 90 ? T.green : kpi.pctTrasmesso >= 70 ? T.accent : T.red, fontSize: "28px", fontWeight: "700", lineHeight: 1, marginBottom: 8 }}>{kpi.pctTrasmesso}%</div>
                <div style={{ height: 6, background: T.borderHi, borderRadius: 3, overflow: "hidden" }}>
                  <div style={{ width: `${Math.min(kpi.pctTrasmesso, 100)}%`, height: "100%", background: kpi.pctTrasmesso >= 90 ? T.green : kpi.pctTrasmesso >= 70 ? T.accent : T.red }} />
                </div>
              </div>
            </>
          )}
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
              <th style={{ ...css.th, cursor: "pointer" }} onClick={() => toggleSort("prenotato_iscrizione")}>Lanciate{sortIcon("prenotato_iscrizione")}</th>
              {kpi.haTrasmesso && <th style={{ ...css.th, cursor: "pointer" }} onClick={() => toggleSort("prenotato_trasmesso")}>Trasm.{sortIcon("prenotato_trasmesso")}</th>}
              <th style={{ ...css.th, cursor: "pointer" }} onClick={() => toggleSort("pren_fine_giro")}>Fine Giro{sortIcon("pren_fine_giro")}</th>
              <th style={{ ...css.th, cursor: "pointer", color: "#e8a838" }} onClick={() => toggleSort("pren_amazon")}>Amazon{sortIcon("pren_amazon")}</th>
              <th style={{ ...css.th, cursor: "pointer" }} onClick={() => toggleSort("teorico")}>Teorico{sortIcon("teorico")}</th>
              <th style={{ ...css.th, cursor: "pointer" }} onClick={() => toggleSort("delta_portale")} title="Fine Giro (no Amazon) − Lanciate">Δ Portale{sortIcon("delta_portale")}</th>
              <th style={css.th}>Uscita</th>
            </tr>
          </thead>
          <tbody>
            {dataFiltrata.map((r, i) => (
              <tr key={r.id || i} style={{ background: i % 2 === 0 ? "transparent" : T.surface + "66" }}>
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
                <td style={{ ...css.td, fontWeight: "600" }}>{(r.prenotato_iscrizione || 0).toLocaleString("it")}</td>
                {kpi.haTrasmesso && (
                  <td style={{ ...css.td, color: r.prenotato_trasmesso > 0 ? T.green : T.textDim, fontWeight: "600" }}>
                    {r.prenotato_trasmesso != null ? r.prenotato_trasmesso.toLocaleString("it") : "—"}
                  </td>
                )}
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
                  {r.pren_senza_amazon > 0 || (r.prenotato_trasmesso ?? r.prenotato_iscrizione) > 0
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
    </div>
  );
}

const MODULES = [
  { id: "dashboard", label: "Dashboard", icon: "◈" },
  { id: "cedola", label: "Giri e Cedole", icon: "≡" },
  { id: "finegiro", label: "Fine Giro", icon: "⊞" },
  { id: "avanzamento", label: "Avanzamento Novità", icon: "▣" },
  { id: "lanci", label: "Lanci Settimanali", icon: "🚀" },
];

const MODULES_IMPORT = [
  { id: "import", label: "Import Cedola", icon: "↑" },
  { id: "prenotato", label: "Import Prenotato", icon: "↳" },
  { id: "importobiettivi", label: "Import Obiettivi", icon: "◎" },
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

  useEffect(() => {
    if (!session) return;
    sbFetch("giri?select=*&order=anno.desc,numero.desc", session.token).then(setGiriDB);
    sbFetch("titoli?select=*&order=ranking_editore.asc,ranking_titolo.asc", session.token).then(setTitoli);
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
  const refreshDati = useCallback(() => {
    if (!session) return;
    sbFetch("prenotato?select=*&limit=100000", session.token).then(data => { if (Array.isArray(data)) setPrenotato(data); });
    sbFetch("titoli?select=*&order=ranking_editore.asc,ranking_titolo.asc", session.token).then(data => { if (Array.isArray(data)) setTitoli(data); });
  }, [session]);

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
          {/* FIX 5: Passato token a ModuloCedola */}
          {activeModule === "cedola" && <ModuloCedola titoli={titoli} giriList={giriDB} onUpdateTitolo={t => { updateTitolo(t); setTitoli(prev => prev.some(x => x.id === t.id) ? prev.map(x => x.id === t.id ? t : x) : [...prev, t]); }} spalmatura={spalmatura} prenotato={prenotato} ruolo={ruolo} token={session.token} onTitoliChange={refreshDati} />}
          {activeModule === "prenotato" && <ModuloPrenotato token={session.token} titoli={titoli} onImportDone={() => sbFetch("prenotato?select=*&limit=100000", session.token).then(setPrenotato)} />}
          {/* MOD 4: Passato spalmatura a ModuloFineGiro */}
          {activeModule === "finegiro" && <ModuloFineGiro titoli={titoli} prenotato={prenotato} canali={canali} token={session.token} ruolo={ruolo} spalmatura={spalmatura} />}
          {activeModule === "avanzamento" && <ModuloAvanzamento token={session.token} titoli={titoli} prenotato={prenotato} ruolo={ruolo} />}
          {activeModule === "lanci" && <ModuloLanciSettimanali token={session.token} titoli={titoli} prenotato={prenotato} canali={canali} ruolo={ruolo} userAccount={userAccount} />}
          {activeModule === "import" && <ModuloImport giriList={giriDB} token={session.token} />}
          {activeModule === "importobiettivi" && <ModuloImport giriList={giriDB} token={session.token} onlyObiettivi={true} />}
        </div>
      </div>
    </div>
  );
}
