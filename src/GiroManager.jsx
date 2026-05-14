import { useState, useEffect, useMemo, useCallback } from "react";
import ModuloImport from "./ModuloImport.jsx";
import ModuloPrenotato from "./ModuloPrenotato.jsx";

const SUPABASE_URL = "https://tdflwenlylhctxssatax.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InRkZmx3ZW5seWxoY3R4c3NhdGF4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYzMzgyNzYsImV4cCI6MjA5MTkxNDI3Nn0.l35qEL7LOvyYuI1McQlVqj4vbyTqmlevcmqWbTGYi2Q";

const sb = {
  auth: {
    signIn: async (email, password) => {
      const r = await fetch(`${SUPABASE_URL}/auth/v1/token?grant_type=password`, {
        method: "POST",
        headers: { "Content-Type": "application/json", "apikey": SUPABASE_KEY },
        body: JSON.stringify({ email, password }),
      });
      return r.json();
    },
    signOut: async (token) => {
      await fetch(`${SUPABASE_URL}/auth/v1/logout`, {
        method: "POST",
        headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}` },
      });
    },
    getUser: async (token) => {
      const r = await fetch(`${SUPABASE_URL}/auth/v1/user`, {
        headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}` },
      });
      return r.json();
    },
  },
};

const sbFetch = async (path, token) => {
  const r = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    headers: { 
      "apikey": SUPABASE_KEY, 
      "Authorization": `Bearer ${token}`, 
      "Accept": "application/json",
      "Range-Unit": "items",
      "Range": "0-9999",
    },
  });
  return r.json();
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

// ─── LOGIN ────────────────────────────────────────────────────────────────────
function LoginScreen({ onLogin }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const handleLogin = async () => {
    setLoading(true);
    setError("");
    const data = await sb.auth.signIn(email, password);
    if (data.access_token) {
      localStorage.setItem("giro_token", data.access_token);
      onLogin(data.access_token, data.user);
    } else {
      setError("Email o password errati.");
    }
    setLoading(false);
  };

  return (
    <div style={{ ...css.app, display: "flex", alignItems: "center", justifyContent: "center", height: "100vh" }}>
      <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, padding: 40, width: 340 }}>
        <div style={{ textAlign: "center", marginBottom: 32 }}>
          <div style={{ color: T.accent, fontSize: "24px", fontWeight: "700", letterSpacing: "0.1em" }}>GIRO</div>
          <div style={{ color: T.textMid, fontSize: "11px", letterSpacing: "0.15em", marginTop: 4 }}>MANAGER</div>
        </div>
        <div style={{ marginBottom: 16 }}>
          <label style={{ color: T.textMid, fontSize: "10px", letterSpacing: "0.08em", textTransform: "uppercase", display: "block", marginBottom: 6 }}>Email</label>
          <input style={{ ...css.input, width: "100%", boxSizing: "border-box" }} type="email" value={email} onChange={(e) => setEmail(e.target.value)} onKeyDown={(e) => e.key === "Enter" && handleLogin()} />
        </div>
        <div style={{ marginBottom: 24 }}>
          <label style={{ color: T.textMid, fontSize: "10px", letterSpacing: "0.08em", textTransform: "uppercase", display: "block", marginBottom: 6 }}>Password</label>
          <input style={{ ...css.input, width: "100%", boxSizing: "border-box" }} type="password" value={password} onChange={(e) => setPassword(e.target.value)} onKeyDown={(e) => e.key === "Enter" && handleLogin()} />
        </div>
        {error && <div style={{ color: T.red, fontSize: "12px", marginBottom: 16, textAlign: "center" }}>{error}</div>}
        <button style={{ ...css.btn("accent"), width: "100%", padding: "10px" }} onClick={handleLogin} disabled={loading}>
          {loading ? "Accesso..." : "Accedi"}
        </button>
      </div>
    </div>
  );
}

// ─── UI COMPONENTS ────────────────────────────────────────────────────────────
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

function EditModal({ titolo, onSave, onClose }) {
  const [form, setForm] = useState({ ...titolo });
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.type === "checkbox" ? e.target.checked : e.target.value }));
  return (
    <div style={{ position: "fixed", inset: 0, background: "#000a", zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center" }} onClick={onClose}>
      <div style={{ background: T.surface, border: `1px solid ${T.borderHi}`, borderRadius: 6, padding: 24, width: 640, maxHeight: "85vh", overflowY: "auto" }} onClick={(e) => e.stopPropagation()}>
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
          {["il_triangolo","top_100"].map((k) => (
            <label key={k} style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer", color: T.textMid, fontSize: "12px" }}>
              <input type="checkbox" checked={!!form[k]} onChange={set(k)} style={{ accentColor: T.accent }} />
              {k === "il_triangolo" ? "▲ Il Triangolo" : "★ Top 100"}
            </label>
          ))}
        </div>
        <div style={{ marginTop: 16 }}>
          <div style={{ color: T.textMid, fontSize: "10px", letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 8 }}>GEMELLI</div>
          {[1,2,3].map((n) => (
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
          <button style={css.btn("accent")} onClick={() => { onSave(form); onClose(); }}>Salva</button>
        </div>
      </div>
    </div>
  );
}

// ─── MODULO: DASHBOARD ────────────────────────────────────────────────────────
function ModuloDashboard({ titoli, prenotato, canali }) {
  const giriLabel = useMemo(() => {
    return [...new Set(titoli.map(t => t.giro_label).filter(Boolean))].sort((a, b) => {
      const [na, ya] = a.split(" "); const [nb, yb] = b.split(" ");
      return Number(yb || 0) - Number(ya || 0) || Number(nb || 0) - Number(na || 0);
    });
  }, [titoli]);

  const [giroSel, setGiroSel] = useState(null);
  useEffect(() => { if (giriLabel.length > 0 && !giroSel) setGiroSel(giriLabel[0]); }, [giriLabel]);

  const titoliGiro = useMemo(() => giroSel ? titoli.filter(t => t.giro_label === giroSel) : [], [titoli, giroSel]);

  const prenotatoGiro = useMemo(() => {
    const titoliIds = new Set(titoliGiro.map(t => t.id));
    return prenotato.filter(p => titoliIds.has(p.titolo_id));
  }, [prenotato, titoliGiro]);

  const totPrenotatoGiro = useMemo(() => prenotatoGiro.reduce((s, p) => s + p.quantita, 0), [prenotatoGiro]);

  const kpiGiro = useMemo(() => {
    const totObj = titoliGiro.reduce((s, t) => s + (t.obiettivo_assegnato || 0), 0);
    const totRag = totPrenotatoGiro;
    const valObj = titoliGiro.reduce((s, t) => s + (t.prezzo || 0) * (t.obiettivo_assegnato || 0), 0);
    const totTriangolo = titoliGiro.filter(t => t.il_triangolo === true).length;
    const totTop100 = titoliGiro.filter(t => t.top_100 === true).length;
    return { totObj, totRag, valObj, totTriangolo, totTop100, count: titoliGiro.length, pct: totObj > 0 ? Math.round(totRag / totObj * 100) : 0 };
  }, [titoliGiro, totPrenotatoGiro]);

  const cedole = useMemo(() => {
    const prenotatoMap = {};
    prenotatoGiro.forEach(p => {
      const t = titoliGiro.find(t => t.id === p.titolo_id);
      if (!t) return;
      const c = t.n_cedola || "—";
      prenotatoMap[c] = (prenotatoMap[c] || 0) + p.quantita;
    });
    const map = {};
    titoliGiro.forEach(t => {
      const c = t.n_cedola || "—";
      if (!map[c]) map[c] = { label: c, totObj: 0, totRag: 0, count: 0 };
      map[c].totObj += t.obiettivo_assegnato || 0;
      map[c].totRag = prenotatoMap[c] || 0;
      map[c].count++;
    });
    return Object.values(map).sort((a, b) => a.label.localeCompare(b.label));
  }, [titoliGiro, prenotatoGiro]);

  const prenotatoPerCanale = useMemo(() => {
    const map = {};
    prenotatoGiro.forEach(p => {
      const c = canali.find(c => c.id === p.canale_id);
      if (!c) return;
      map[c.nome] = (map[c.nome] || 0) + p.quantita;
    });
    return Object.entries(map).sort((a, b) => b[1] - a[1]);
  }, [prenotatoGiro, canali]);
  const maxCanale = prenotatoPerCanale[0]?.[1] || 1;

  return (
    <div style={{ flex: 1, overflowY: "auto", padding: 20 }}>
      <div style={{ display: "flex", gap: 12, alignItems: "center", marginBottom: 20 }}>
        <select style={{ ...css.input, fontSize: "14px", fontWeight: "600", color: T.accent }} value={giroSel || ""} onChange={e => setGiroSel(e.target.value)}>
          {giriLabel.map(g => <option key={g} value={g}>Giro {g}</option>)}
        </select>
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
          <div style={{ color: T.textMid, fontSize: "11px", marginTop: 6 }}>{kpiGiro.totRag.toLocaleString("it")} / {kpiGiro.totObj.toLocaleString("it")}</div>
        </div>
      </div>

      <div style={{ marginBottom: 24 }}>
        <div style={{ color: T.textMid, fontSize: "11px", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 12 }}>CEDOLE</div>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr>{["Cedola","Titoli","Obj Ass.","Prenotato","Avanz."].map(h => <th key={h} style={{ padding: "8px 12px", textAlign: "left", color: T.textMid, fontWeight: "400", fontSize: "11px", letterSpacing: "0.08em", textTransform: "uppercase", borderBottom: `1px solid ${T.border}`, background: T.surface }}>{h}</th>)}</tr>
          </thead>
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
                      <div style={{ width: 80, height: 4, background: T.borderHi, borderRadius: 2, overflow: "hidden" }}>
                        <div style={{ width: `${pct}%`, height: "100%", background: pct >= 80 ? T.green : pct >= 50 ? T.accent : T.red }} />
                      </div>
                      <span style={{ color: pct >= 80 ? T.green : pct >= 50 ? T.accent : T.red, fontSize: "11px", fontWeight: "700" }}>{pct}%</span>
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {prenotatoPerCanale.length > 0 && (
        <div>
          <div style={{ color: T.textMid, fontSize: "11px", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 12 }}>PRENOTATO PER CANALE</div>
          {prenotatoPerCanale.map(([nome, qta]) => (
            <div key={nome} style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 6 }}>
              <div style={{ width: 200, fontSize: "12px" }}>{nome}</div>
              <div style={{ flex: 1, height: 8, background: T.borderHi, borderRadius: 2, overflow: "hidden" }}>
                <div style={{ width: `${(qta / maxCanale) * 100}%`, height: "100%", background: T.blue }} />
              </div>
              <div style={{ width: 70, textAlign: "right", color: T.accent, fontSize: "12px", fontWeight: "600" }}>{qta.toLocaleString("it")}</div>
            </div>
          ))}
          <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 8, paddingTop: 8, borderTop: `1px solid ${T.border}` }}>
            <span style={{ color: T.textMid, fontSize: "12px" }}>Totale: <span style={{ color: T.green, fontWeight: "700" }}>{totPrenotatoGiro.toLocaleString("it")}</span></span>
          </div>
        </div>
      )}
    </div>
  );
}



// ─── MODULO: CEDOLA ───────────────────────────────────────────────────────────
function ModuloCedola({ titoli, giriList, onUpdateTitolo, spalmatura, prenotato }) {
  const [giroLabelSel, setGiroLabelSel] = useState("tutti");
  const [giroSel, setGiroSel] = useState("tutti");
  const [search, setSearch] = useState("");
  const [filterFlag, setFilterFlag] = useState("tutti");
  const [filterEditori, setFilterEditori] = useState([]);
  const [filterAccount, setFilterAccount] = useState("tutti");
  const [editoreDropdownOpen, setEditoreDropdownOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [sortKey, setSortKey] = useState("n_cedola");

  const giriLabel = useMemo(() => {
    return [...new Set(titoli.map(t => t.giro_label).filter(Boolean))].sort((a, b) => {
      const [na, ya] = a.split(" "); const [nb, yb] = b.split(" ");
      return Number(yb) - Number(ya) || Number(nb) - Number(na);
    });
  }, [titoli]);

  const cedole = useMemo(() => {
    const t = giroLabelSel === "tutti" ? titoli : titoli.filter(t => t.giro_label === giroLabelSel);
    return [...new Set(t.map(t => t.n_cedola).filter(Boolean))].sort();
  }, [titoli, giroLabelSel]);
const accounts = useMemo(() => {
    const t = giroSel === "tutti"
      ? (giroLabelSel === "tutti" ? titoli : titoli.filter(t => t.giro_label === giroLabelSel))
      : titoli.filter(t => t.n_cedola === giroSel);
    return [...new Set(t.map(t => t.account_editore).filter(Boolean))].sort();
  }, [titoli, giroLabelSel, giroSel]);
  const editori = useMemo(() => {
    const t = giroSel === "tutti"
      ? (giroLabelSel === "tutti" ? titoli : titoli.filter(t => t.giro_label === giroLabelSel))
      : titoli.filter(t => t.n_cedola === giroSel);
    return [...new Set(t.map(t => t.editore_nome).filter(Boolean))].sort();
  }, [titoli, giroLabelSel, giroSel]);

  const filtered = useMemo(() => {
    return titoli
      .filter(t => giroLabelSel === "tutti" || t.giro_label === giroLabelSel)
      .filter(t => giroSel === "tutti" || t.n_cedola === giroSel)
      .filter(t => filterEditori.length === 0 || filterEditori.includes(t.editore_nome))
      .filter(t => filterAccount === "tutti" || t.account_editore === filterAccount)
      .filter(t => {
        if (!search) return true;
        const q = search.toLowerCase();
        return t.titolo?.toLowerCase().includes(q) || t.autore?.toLowerCase().includes(q) || t.editore_nome?.toLowerCase().includes(q) || t.ean?.includes(q);
      })
      .filter(t => {
        if (filterFlag === "triangolo") return t.il_triangolo;
        if (filterFlag === "top100") return t.top_100;
        if (filterFlag === "gemelli") return t.ean_gemello_1;
        return true;
      })
      .sort((a, b) => {
        if (sortKey === "n_cedola") return (a.n_cedola ?? "").localeCompare(b.n_cedola ?? "");
        if (sortKey === "editore") return (a.editore_nome ?? "").localeCompare(b.editore_nome ?? "");
        if (sortKey === "prezzo") return (b.prezzo ?? 0) - (a.prezzo ?? 0);
        return 0;
      });
  }, [titoli, giroLabelSel, giroSel, search, filterFlag, filterEditori, filterAccount, sortKey]);

  const editingTitolo = titoli.find(t => t.id === editingId);
  const avanzamento = useMemo(() => {
    const totObj = filtered.reduce((s, t) => s + (t.obiettivo_assegnato || 0), 0);
    const totRag = filtered.reduce((s, t) => s + (t.obiettivo_raggiunto || 0), 0);
    return { totObj, totRag, pct: totObj > 0 ? Math.round((totRag / totObj) * 100) : 0, count: filtered.length };
  }, [filtered]);

  const getObjCanale = (titolo, canale_codice) => {
    const spRow = spalmatura.find(s => s.editore_nome === titolo.editore_nome && s.formato === (titolo.formato || 'Cover') && s.canale_codice === canale_codice);
    if (!spRow || !titolo.obiettivo_assegnato) return 0;
    return Math.round(titolo.obiettivo_assegnato * spRow.percentuale);
  };

  const exportAgenti = () => {
    const XLSX = window.XLSX;
    const headers = ["N° CEDOLA","EAN","TITOLO","AUTORE","EDITORE","PREZZO","USCITA","NOTE","EAN GEM 1","TITOLO GEM 1","EAN GEM 2","TITOLO GEM 2","EAN GEM 3","TITOLO GEM 3","OBJ INDIPENDENTI & ALTRE CATENE"];
    const rows = filtered.map(t => [t.n_cedola, t.ean, t.titolo, t.autore, t.editore_nome, t.prezzo, t.uscita, t.note_comunicazione || t.note, t.ean_gemello_1, t.titolo_gemello_1, t.ean_gemello_2, t.titolo_gemello_2, t.ean_gemello_3, t.titolo_gemello_3, getObjCanale(t, 'INDIPENDENTI_ALTRE_CATENE')]);
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    ws["!cols"] = [14,16,40,25,20,8,10,30,16,30,16,30,16,30,14].map(w => ({ wch: w }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "CEDOLA AGENTI");
    XLSX.writeFile(wb, `CEDOLA_AGENTI_${giroLabelSel === "tutti" ? "TUTTI" : giroLabelSel}.xlsx`);
  };

  const exportDirezionale = () => {
    const pwd = prompt("Password direzionale:");
    if (pwd !== "nuovaluce") { alert("Password errata."); return; }
    const XLSX = window.XLSX;
    const canaliDir = [
      { codice: 'FELTRINELLI', label: 'Feltrinelli' }, { codice: 'GIUNTI', label: 'Giunti' },
      { codice: 'MONDADORI', label: 'Mondadori' }, { codice: 'UBIK', label: 'Ubik' },
      { codice: 'INDIPENDENTI_ALTRE_CATENE', label: 'Indip. & Altre Catene' },
      { codice: 'AMAZON', label: 'Amazon' }, { codice: 'IBS', label: 'IBS' },
      { codice: 'ALTRI_ONLINE', label: 'Altri Online' }, { codice: 'FASTBOOK', label: 'Fastbook' },
      { codice: 'GROSSISTI', label: 'Grossisti' }, { codice: 'CENTROLIBRI', label: 'Centrolibri' }, { codice: 'GDO', label: 'GDO' },
    ];
    const headersCedola = ["N° CEDOLA","EAN","TITOLO","AUTORE","EDITORE","PREZZO","OBJ TOTALE","NOTE","EAN GEM 1","TITOLO GEM 1","EAN GEM 2","TITOLO GEM 2","EAN GEM 3","TITOLO GEM 3"];
    const rowsCedola = filtered.map(t => [t.n_cedola, t.ean, t.titolo, t.autore, t.editore_nome, t.prezzo, t.obiettivo_assegnato || 0, t.note_comunicazione || t.note, t.ean_gemello_1, t.titolo_gemello_1, t.ean_gemello_2, t.titolo_gemello_2, t.ean_gemello_3, t.titolo_gemello_3]);
    const wsCedola = XLSX.utils.aoa_to_sheet([headersCedola, ...rowsCedola]);
    wsCedola["!cols"] = [14,16,40,25,20,8,10,30,16,30,16,30,16,30].map(w => ({ wch: w }));
    const headersObj = ["N° CEDOLA","EAN","TITOLO","AUTORE","EDITORE","PREZZO","OBJ TOTALE",...canaliDir.map(c => c.label)];
    const rowsObj = filtered.map(t => { const objPerCanale = canaliDir.map(c => getObjCanale(t, c.codice)); return [t.n_cedola, t.ean, t.titolo, t.autore, t.editore_nome, t.prezzo, t.obiettivo_assegnato || 0, ...objPerCanale]; });
    const wsObj = XLSX.utils.aoa_to_sheet([headersObj, ...rowsObj]);
    wsObj["!cols"] = [14,16,40,25,20,8,10,...canaliDir.map(() => 14)].map(w => ({ wch: w }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, wsCedola, "CEDOLA");
    XLSX.utils.book_append_sheet(wb, wsObj, "OBIETTIVI");
    XLSX.writeFile(wb, `CEDOLA_DIREZIONALE_${giroLabelSel === "tutti" ? "TUTTI" : giroLabelSel}.xlsx`);
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%", overflow: "hidden" }}>
      {editingTitolo && <EditModal titolo={editingTitolo} onSave={onUpdateTitolo} onClose={() => setEditingId(null)} />}
      <div style={{ padding: "12px 20px", borderBottom: `1px solid ${T.border}`, display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
        <select style={css.input} value={giroLabelSel} onChange={(e) => { setGiroLabelSel(e.target.value); setGiroSel("tutti"); setFilterEditori([]); }}>
          <option value="tutti">Tutti i giri</option>
          {giriLabel.map(g => <option key={g} value={g}>Giro {g}</option>)}
        </select>
        <select style={css.input} value={giroSel} onChange={(e) => { setGiroSel(e.target.value); setFilterEditori([]); }}>
          <option value="tutti">Tutte le cedole</option>
          {cedole.map(c => <option key={c} value={c}>{c}</option>)}
        </select>
        <div style={{ position: "relative" }}>
          <button style={{ ...css.btn(), minWidth: 180, display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }} onClick={() => setEditoreDropdownOpen(o => !o)}>
            <span>{filterEditori.length === 0 ? "Tutti gli editori" : `${filterEditori.length} selezionati`}</span>
            <span>▾</span>
          </button>
          {editoreDropdownOpen && (
            <div style={{ position: "absolute", top: "100%", left: 0, zIndex: 50, background: T.surface, border: `1px solid ${T.borderHi}`, borderRadius: 4, minWidth: 220, maxHeight: 300, overflowY: "auto", marginTop: 4, boxShadow: "0 4px 20px #0008" }}>
              <div style={{ padding: "8px 12px", borderBottom: `1px solid ${T.border}`, display: "flex", justifyContent: "space-between" }}>
                <span style={{ color: T.textMid, fontSize: "11px" }}>{editori.length} editori</span>
                <span style={{ color: T.accent, fontSize: "11px", cursor: "pointer" }} onClick={() => setFilterEditori([])}>Deseleziona tutti</span>
              </div>
              {editori.map(e => (
                <label key={e} style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 12px", cursor: "pointer", borderBottom: `1px solid ${T.border}22`, background: filterEditori.includes(e) ? T.accent + "18" : "transparent" }}>
                  <input type="checkbox" checked={filterEditori.includes(e)} onChange={() => setFilterEditori(prev => prev.includes(e) ? prev.filter(x => x !== e) : [...prev, e])} style={{ accentColor: T.accent }} />
                  <span style={{ fontSize: "12px", color: filterEditori.includes(e) ? T.accent : T.text }}>{e}</span>
                </label>
              ))}
            <div style={{ padding: 8, borderTop: `1px solid ${T.border}` }}>
                <button style={{ ...css.btn("accent"), width: "100%" }} onClick={() => setEditoreDropdownOpen(false)}>Applica</button>
              </div>
            </div>
          )}
        </div>
        <select style={css.input} value={filterAccount} onChange={(e) => setFilterAccount(e.target.value)}>
          <option value="tutti">Tutti gli account</option>
          {accounts.map(a => <option key={a} value={a}>{a}</option>)}
        </select>
        <input style={{ ...css.input, width: 180 }} placeholder="Cerca..." value={search} onChange={(e) => setSearch(e.target.value)} />
        {["tutti","triangolo","top100","gemelli"].map(f => (
          <button key={f} style={{ ...css.btn(filterFlag === f ? "accent" : "default"), padding: "5px 10px" }} onClick={() => setFilterFlag(f)}>
            {f === "tutti" ? "Tutti" : f === "triangolo" ? "▲" : f === "top100" ? "★" : "Gem."}
          </button>
        ))}
        <div style={{ marginLeft: "auto", color: T.textMid, fontSize: "11px" }}>
          <span style={{ color: T.text }}>{avanzamento.count}</span> titoli &nbsp;·&nbsp; Obj: <span style={{ color: T.accent }}>{avanzamento.pct}%</span>
        </div>
        <button style={css.btn("accent")} onClick={exportAgenti}>↓ Download Agenti</button>
        <button style={css.btn("accent")} onClick={exportDirezionale}>↓ Download Direzionali</button>
      </div>
      <div style={{ flex: 1, overflowY: "auto" }}>
        <table style={css.table}>
          <thead>
            <tr>
              {[["n_cedola","Cedola"],["","EAN"],["editore","Editore"],["","Titolo"],["","Autore"],["prezzo","€"],["","Uscita"],["","Obj"],["","Flag"],["","Gemelli"],["","Note"],["",""]].map(([key, label]) => (
                <th key={label} style={css.th} onClick={() => key && setSortKey(key)}>{label}{key && sortKey === key ? " ↓" : ""}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.map((t, i) => (
              <tr key={t.id} style={{ background: i % 2 === 0 ? "transparent" : T.surface + "66" }}>
                <td style={{ ...css.td, color: T.textMid, fontSize: "10px", whiteSpace: "nowrap" }}>{t.n_cedola}</td>
                <td style={{ ...css.td, color: T.textDim, fontFamily: "monospace", fontSize: "11px" }}>{t.ean}</td>
                <td style={{ ...css.td, color: T.accent, fontWeight: "600", whiteSpace: "nowrap" }}>{t.editore_nome}</td>
                <td style={{ ...css.td, maxWidth: 260 }}>
                  <div style={{ fontWeight: "600", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{t.titolo}</div>
                </td>
                <td style={{ ...css.td, color: T.textMid }}>{t.autore}</td>
                <td style={css.td}>€ {t.prezzo?.toFixed(2)}</td>
                <td style={{ ...css.td, color: T.textMid }}>{t.uscita}</td>
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

function ModuloFineGiro({ titoli, prenotato, canali }) {
  const giriLabel = useMemo(() => [...new Set(titoli.map(t => t.giro_label).filter(Boolean))].sort((a, b) => {
    const [na, ya] = a.split(" "); const [nb, yb] = b.split(" ");
    return Number(yb || 0) - Number(ya || 0) || Number(nb || 0) - Number(na || 0);
  }), [titoli]);

  const [giroLabelSel, setGiroLabelSel] = useState("tutti");
  const [cedolaSel, setCedolaSel] = useState("tutti");
  const [filterEditore, setFilterEditore] = useState("tutti");
  const [filterAccount, setFilterAccount] = useState("tutti");

  const cedole = useMemo(() => {
    const t = giroLabelSel === "tutti" ? titoli : titoli.filter(t => t.giro_label === giroLabelSel);
    return [...new Set(t.map(t => t.n_cedola).filter(Boolean))].sort();
  }, [titoli, giroLabelSel]);

  const editori = useMemo(() => {
    const t = giroLabelSel === "tutti" ? titoli : titoli.filter(t => t.giro_label === giroLabelSel);
    return [...new Set(t.map(t => t.editore_nome).filter(Boolean))].sort();
  }, [titoli, giroLabelSel]);

  const accounts = useMemo(() => {
    const t = giroLabelSel === "tutti" ? titoli : titoli.filter(t => t.giro_label === giroLabelSel);
    return [...new Set(t.map(t => t.account_editore).filter(Boolean))].sort();
  }, [titoli, giroLabelSel]);

  const titoliSel = useMemo(() => titoli
    .filter(t => giroLabelSel === "tutti" || t.giro_label === giroLabelSel)
    .filter(t => cedolaSel === "tutti" || t.n_cedola === cedolaSel)
    .filter(t => filterEditore === "tutti" || t.editore_nome === filterEditore)
    .filter(t => filterAccount === "tutti" || t.account_editore === filterAccount)
    .sort((a, b) => (a.ranking_editore ?? 99) - (b.ranking_editore ?? 99) || (a.ranking_titolo ?? 99) - (b.ranking_titolo ?? 99)),
  [titoli, giroLabelSel, cedolaSel, filterEditore, filterAccount]);

  const totObj = titoliSel.reduce((s, t) => s + (t.obiettivo_assegnato || 0), 0);

  const righe = useMemo(() => titoliSel.map(t => {
    const pren = prenotato.filter(p => p.titolo_id === t.id);
    const totPren = pren.reduce((s, p) => s + p.quantita, 0);
    const byCanale = {};
    pren.forEach(p => { const c = canali.find(c => c.id === p.canale_id); if (c) byCanale[c.codice] = p.quantita; });
    return { titolo: t, totPren, byCanale };
  }), [titoliSel, prenotato, canali]);

  const totPrenotato = righe.reduce((s, r) => s + r.totPren, 0);
  const pctTot = totObj > 0 ? Math.round(totPrenotato / totObj * 100) : 0;
  const canaliPrincipali = canali;

  const exportExcel = () => {
    const XLSX = window.XLSX;
    const headers = ["N° CEDOLA","EAN","EDITORE","TITOLO","OBJ ASS.","PRENOTATO","AVANZ %", ...canaliPrincipali.map(c => c.nome)];
    const rows = righe.map(({ titolo: t, totPren, byCanale }) => {
      const pct = t.obiettivo_assegnato > 0 ? Math.round(totPren / t.obiettivo_assegnato * 100) : 0;
      return [t.n_cedola, t.ean, t.editore_nome, t.titolo, t.obiettivo_assegnato, totPren, `${pct}%`, ...canaliPrincipali.map(c => byCanale[c.codice] ?? 0)];
    });
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    ws["!cols"] = [14,16,20,40,10,10,8,...canaliPrincipali.map(() => 12)].map(w => ({ wch: w }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "FINE GIRO");
    const label = giroLabelSel === "tutti" ? "TUTTI" : giroLabelSel;
    XLSX.writeFile(wb, `FINE_GIRO_${label}.xlsx`);
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%", overflow: "hidden" }}>
      <div style={{ padding: "12px 20px", borderBottom: `1px solid ${T.border}`, display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
        <select style={css.input} value={giroLabelSel} onChange={(e) => { setGiroLabelSel(e.target.value); setCedolaSel("tutti"); setFilterEditore("tutti"); setFilterAccount("tutti"); }}>
          <option value="tutti">Tutti i giri</option>
          {giriLabel.map(g => <option key={g} value={g}>Giro {g}</option>)}
        </select>
        <select style={css.input} value={cedolaSel} onChange={(e) => setCedolaSel(e.target.value)}>
          <option value="tutti">Tutte le cedole</option>
          {cedole.map(c => <option key={c} value={c}>{c}</option>)}
        </select>
        <select style={css.input} value={filterEditore} onChange={(e) => setFilterEditore(e.target.value)}>
          <option value="tutti">Tutti gli editori</option>
          {editori.map(e => <option key={e} value={e}>{e}</option>)}
        </select>
        <select style={css.input} value={filterAccount} onChange={(e) => setFilterAccount(e.target.value)}>
          <option value="tutti">Tutti gli account</option>
          {accounts.map(a => <option key={a} value={a}>{a}</option>)}
        </select>
        <div style={{ color: T.textMid, fontSize: "12px" }}>
          Obj: <span style={{ color: T.accent, fontWeight: "700" }}>{totObj.toLocaleString("it")}</span>
          &nbsp;·&nbsp; Prenotato: <span style={{ color: T.green, fontWeight: "700" }}>{totPrenotato.toLocaleString("it")}</span>
          &nbsp;·&nbsp; <span style={{ color: pctTot >= 80 ? T.green : pctTot >= 50 ? T.accent : T.red, fontWeight: "700" }}>{pctTot}%</span>
        </div>
        <button style={{ ...css.btn("accent"), marginLeft: "auto" }} onClick={exportExcel}>↓ Export Excel</button>
      </div>
      <div style={{ flex: 1, overflowY: "auto" }}>
        <table style={css.table}>
          <thead>
            <tr>
              <th style={css.th}>Cedola</th>
              <th style={css.th}>Editore</th>
              <th style={css.th}>Titolo</th>
              <th style={css.th}>EAN</th>
              <th style={css.th}>Obj. Ass.</th>
              <th style={css.th}>Prenotato</th>
              <th style={css.th}>Avanz.</th>
              {canaliPrincipali.map(c => <th key={c.id} style={{ ...css.th, writingMode: "vertical-rl", transform: "rotate(180deg)", height: 80, whiteSpace: "nowrap", verticalAlign: "bottom", padding: "8px 4px" }}>{c.nome}</th>)}
            </tr>
          </thead>
          <tbody>
            {righe.map(({ titolo: t, totPren, byCanale }, i) => {
              const pct = t.obiettivo_assegnato > 0 ? Math.round(totPren / t.obiettivo_assegnato * 100) : 0;
              return (
                <tr key={t.id} style={{ background: i % 2 === 0 ? "transparent" : T.surface + "66" }}>
                  <td style={{ ...css.td, color: T.textMid, fontSize: "10px", whiteSpace: "nowrap" }}>{t.n_cedola}</td>
                  <td style={{ ...css.td, color: T.accent, fontWeight: "600", whiteSpace: "nowrap" }}>{t.editore_nome}</td>
                  <td style={{ ...css.td, maxWidth: 220 }}><div style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{t.titolo}</div></td>
                  <td style={{ ...css.td, fontFamily: "monospace", fontSize: "11px", color: T.textMid }}>{t.ean}</td>
                  <td style={css.td}>{t.obiettivo_assegnato?.toLocaleString("it")}</td>
                  <td style={{ ...css.td, color: T.green, fontWeight: "600" }}>{totPren > 0 ? totPren.toLocaleString("it") : "—"}</td>
                  <td style={css.td}><span style={{ color: pct >= 80 ? T.green : pct >= 50 ? T.accent : T.red, fontWeight: "700" }}>{pct}%</span></td>
                  {canaliPrincipali.map(c => <td key={c.id} style={{ ...css.td, color: byCanale[c.codice] ? T.text : T.textDim }}>{byCanale[c.codice]?.toLocaleString("it") ?? "—"}</td>)}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}


// ─── APP ROOT ─────────────────────────────────────────────────────────────────
const MODULES = [
  { id: "dashboard", label: "Dashboard", icon: "◈" },
  { id: "cedola", label: "Giri e Cedole", icon: "≡" },
  { id: "import", label: "Import Cedola", icon: "↑" },
  { id: "finegiro", label: "Fine Giro", icon: "⊞" },
  { id: "prenotato", label: "Import Prenotato", icon: "↳" },
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

  useEffect(() => {
    if (!session) return;
    sbFetch("giri?select=*&order=anno.desc,numero.desc", session.token).then(setGiriDB);
    sbFetch("titoli?select=*&order=ranking_editore.asc,ranking_titolo.asc", session.token).then(setTitoli);
   sbFetch("prenotato?select=*&limit=100000", session.token).then(setPrenotato);
    sbFetch("spalmatura_obiettivo?select=*", session.token).then(setSpalmatura);
    sbFetch("canali?select=*&order=nome.asc", session.token).then(setCanali);
  }, [session]);

  useEffect(() => {
    const token = localStorage.getItem("giro_token");
    if (token) {
      sb.auth.getUser(token).then((user) => {
        if (user?.email) setSession({ token, user });
        else localStorage.removeItem("giro_token");
        setCheckingAuth(false);
      });
    } else {
      setCheckingAuth(false);
    }
  }, []);

  const handleLogin = (token, user) => setSession({ token, user });
  const handleLogout = async () => {
    await sb.auth.signOut(session.token);
    localStorage.removeItem("giro_token");
    setSession(null);
  };

  const updateTitolo = useCallback((updated) => {
    setTitoli(prev => prev.map(t => t.id === updated.id ? { ...t, ...updated } : t));
  }, []);
  const refreshDati = useCallback(() => {
    if (!session) return;
    sbFetch("prenotato?select=*", session.token).then(data => { if (Array.isArray(data)) setPrenotato(data); });
    sbFetch("titoli?select=*&order=ranking_editore.asc,ranking_titolo.asc", session.token).then(data => { if (Array.isArray(data)) setTitoli(data); });
  }, [session]);

  if (checkingAuth) return <div style={{ ...css.app, display: "flex", alignItems: "center", justifyContent: "center", height: "100vh", color: T.textMid }}>Caricamento...</div>;
  if (!session) return <LoginScreen onLogin={handleLogin} />;

  return (
    <div style={{ ...css.app, display: "flex", height: "100vh" }}>
      <div style={css.sidebar}>
        <div style={{ padding: "20px 16px 16px", borderBottom: `1px solid ${T.border}` }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
            <img src="https://raw.githubusercontent.com/albertospde/giro-manager/main/.github/logo_pde.png" style={{ height: 36, borderRadius: 6 }} alt="PDE" />
          </div>
          <div style={{ color: T.accent, fontSize: "13px", fontWeight: "700", letterSpacing: "0.06em" }}>GIRO MANAGER</div>
        </div>
        <nav style={{ flex: 1, padding: "8px 0" }}>
          {MODULES.map(m => (
            <button key={m.id} style={{ width: "100%", textAlign: "left", padding: "10px 16px", border: "none", background: activeModule === m.id ? T.accent + "18" : "transparent", color: activeModule === m.id ? T.accent : T.textMid, cursor: "pointer", fontFamily: "inherit", fontSize: "12px", borderLeft: `2px solid ${activeModule === m.id ? T.accent : "transparent"}`, display: "flex", alignItems: "center", gap: 10, letterSpacing: "0.04em" }}
              onClick={() => setActiveModule(m.id)}>
              <span style={{ fontSize: "14px" }}>{m.icon}</span> {m.label}
            </button>
          ))}
        </nav>
        <div style={{ padding: "12px 16px", borderTop: `1px solid ${T.border}` }}>
          <div style={{ color: T.textDim, fontSize: "10px", marginBottom: 6 }}>{session.user?.email}</div>
          <button style={{ ...css.btn(), fontSize: "11px", padding: "4px 10px", width: "100%" }} onClick={handleLogout}>Esci</button>
        </div>
      </div>
      <div style={css.main}>
        <div style={css.header}>
          <span style={{ color: T.accent, fontSize: "11px", letterSpacing: "0.12em", textTransform: "uppercase" }}>{MODULES.find(m => m.id === activeModule)?.label}</span>
          <span style={{ color: T.borderHi }}>·</span>
          <span style={{ color: T.textMid, fontSize: "11px" }}>{titoli.length} titoli · {[...new Set(titoli.map(t => t.n_cedola).filter(Boolean))].length} cedole</span>
          <button style={{ ...css.btn(), marginLeft: "auto", fontSize: "11px", padding: "4px 10px" }} onClick={refreshDati}>↺ Aggiorna</button>
        </div>
        <div style={{ flex: 1, overflow: "hidden", display: "flex", flexDirection: "column" }}>
          {activeModule === "dashboard" && <ModuloDashboard titoli={titoli} prenotato={prenotato} canali={canali} />}
          {activeModule === "cedola" && <ModuloCedola titoli={titoli} giriList={giriDB} onUpdateTitolo={updateTitolo} spalmatura={spalmatura} prenotato={prenotato} />}
          {activeModule === "prenotato" && <ModuloPrenotato token={session.token} titoli={titoli} onImportDone={() => sbFetch("prenotato?select=*", session.token).then(setPrenotato)} />}
          {activeModule === "finegiro" && <ModuloFineGiro titoli={titoli} prenotato={prenotato} canali={canali} />}
          {activeModule === "import" && <ModuloImport giriList={giriDB} token={session.token} />}
        </div>
      </div>
    </div>
  );
}
