import ModuloPrenotato from "./ModuloPrenotato.jsx";
import { useState, useEffect, useMemo, useCallback } from "react";
import ModuloImport from "./ModuloImport.jsx";

// ─── SUPABASE CONFIG ──────────────────────────────────────────────────────────
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
    headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${token}`, "Accept": "application/json" },
  });
  return r.json();
};
// ─── MOCK DATA ────────────────────────────────────────────────────────────────
const giriDB = [
  { id: 1, numero: 4, anno: 2026, label: "GIRO 4 2026 A", sub_giro: "A", attivo: true },
  { id: 2, numero: 4, anno: 2026, label: "GIRO 4 2026 B", sub_giro: "B", attivo: true },
  { id: 3, numero: 4, anno: 2026, label: "GIRO 4 2026 C", sub_giro: "C", attivo: true },
  { id: 4, numero: 4, anno: 2026, label: "GIRO 4 2026 KIDS", sub_giro: "KIDS", attivo: true },
];

const MOCK_TITOLI = [
  { id: 1, giro_id: 1, ean: "9788845938047", titolo: "IL NOME DELLA ROSA", autore: "ECO UMBERTO", editore_nome: "BOMPIANI", prezzo: 22, uscita: "MAGGIO", formato: "Cover", ranking_editore: 1, ranking_titolo: 1, n_cedola: 1, obiettivo_assegnato: 120, obiettivo_raggiunto: 87, il_triangolo: true, top_100: true, note: "", ean_gemello_1: "", titolo_gemello_1: "", ean_gemello_2: "", titolo_gemello_2: "", ean_gemello_3: "", titolo_gemello_3: "", account_editore: "ROSSI", promozione: "PDE Promozione" },
  { id: 2, giro_id: 1, ean: "9788807034053", titolo: "SE UNA NOTTE D'INVERNO UN VIAGGIATORE", autore: "CALVINO ITALO", editore_nome: "FELTRINELLI", prezzo: 14, uscita: "GIUGNO", formato: "Tascabile", ranking_editore: 1, ranking_titolo: 2, n_cedola: 2, obiettivo_assegnato: 80, obiettivo_raggiunto: 0, il_triangolo: false, top_100: true, note: "Tit. molto atteso", ean_gemello_1: "9788807034060", titolo_gemello_1: "EDIZIONE SPECIALE", ean_gemello_2: "", titolo_gemello_2: "", ean_gemello_3: "", titolo_gemello_3: "", account_editore: "BIANCHI", promozione: "PDE Promozione" },
  { id: 3, giro_id: 1, ean: "9788845936128", titolo: "LA COGNIZIONE DEL DOLORE", autore: "GADDA CARLO EMILIO", editore_nome: "EINAUDI", prezzo: 18, uscita: "GIUGNO", formato: "Cover", ranking_editore: 2, ranking_titolo: 1, n_cedola: 3, obiettivo_assegnato: 60, obiettivo_raggiunto: 0, il_triangolo: true, top_100: false, note: "", ean_gemello_1: "", titolo_gemello_1: "", ean_gemello_2: "", titolo_gemello_2: "", ean_gemello_3: "", titolo_gemello_3: "", account_editore: "VERDI", promozione: "PDE Promozione" },
  { id: 4, giro_id: 1, ean: "9788806204075", titolo: "STORIA DELLA COLONNA INFAME", autore: "MANZONI ALESSANDRO", editore_nome: "EINAUDI", prezzo: 12, uscita: "LUGLIO", formato: "Tascabile", ranking_editore: 2, ranking_titolo: 2, n_cedola: 4, obiettivo_assegnato: 45, obiettivo_raggiunto: 0, il_triangolo: false, top_100: false, note: "", ean_gemello_1: "", titolo_gemello_1: "", ean_gemello_2: "", titolo_gemello_2: "", ean_gemello_3: "", titolo_gemello_3: "", account_editore: "VERDI", promozione: "PDE Promozione" },
  { id: 5, giro_id: 2, ean: "9788852073229", titolo: "CENTOMILA MILIARDI DI POESIE", autore: "QUENEAU RAYMOND", editore_nome: "EINAUDI", prezzo: 20, uscita: "MAGGIO", formato: "Cover", ranking_editore: 3, ranking_titolo: 1, n_cedola: 1, obiettivo_assegnato: 70, obiettivo_raggiunto: 0, il_triangolo: false, top_100: true, note: "", ean_gemello_1: "", titolo_gemello_1: "", ean_gemello_2: "", titolo_gemello_2: "", ean_gemello_3: "", titolo_gemello_3: "", account_editore: "NERI", promozione: "PDE Promozione" },
];

const MOCK_CANALI = [
  { id: 1, codice: "FELTRINELLI", nome: "Feltrinelli", gruppo: "CATENE" },
  { id: 2, codice: "MONDADORI", nome: "Mondadori", gruppo: "CATENE" },
  { id: 3, codice: "UBIK", nome: "Ubik", gruppo: "CATENE" },
  { id: 4, codice: "GIUNTI", nome: "Giunti al Punto", gruppo: "CATENE" },
  { id: 5, codice: "LIBRACCIO", nome: "Libraccio", gruppo: "CATENE" },
  { id: 6, codice: "FASTBOOK", nome: "Fastbook + GD", gruppo: "GROSSISTI" },
  { id: 7, codice: "CENTROLIBRI", nome: "Centrolibri", gruppo: "GROSSISTI" },
  { id: 8, codice: "AMAZON", nome: "Amazon", gruppo: "ONLINE" },
  { id: 9, codice: "IBS", nome: "IBS", gruppo: "ONLINE" },
  { id: 10, codice: "ALTRI_ONLINE", nome: "Altri Online", gruppo: "ONLINE" },
  { id: 11, codice: "INDIPENDENTI", nome: "Indipendenti", gruppo: "INDIE" },
  { id: 12, codice: "LIB_RELIGIOSE", nome: "Lib. Religiose", gruppo: "INDIE" },
];

const MOCK_PRENOTATO = [
  { id: 1, titolo_id: 1, canale_id: 1, quantita: 25 },
  { id: 2, titolo_id: 1, canale_id: 2, quantita: 18 },
  { id: 3, titolo_id: 1, canale_id: 8, quantita: 30 },
  { id: 4, titolo_id: 2, canale_id: 1, quantita: 12 },
  { id: 5, titolo_id: 2, canale_id: 9, quantita: 8 },
];

// ─── COLORI TEMA ──────────────────────────────────────────────────────────────
const T = {
  bg: "#1a2140",
  surface: "#212d54",
  border: "#2e3d6b",
  borderHi: "#3d4f82",
  text: "#f0f2f8",
  textMid: "#8b9cc8",
  textDim: "#4a5a8a",
  accent: "#7b9fe8",
  green: "#4caf7d",
  red: "#e05c5c",
  blue: "#4a5da0",
  purple: "#9c6fcf",
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
          <input style={{ ...css.input, width: "100%", boxSizing: "border-box" }} type="email" value={email}
            onChange={(e) => setEmail(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && handleLogin()} />
        </div>
        <div style={{ marginBottom: 24 }}>
          <label style={{ color: T.textMid, fontSize: "10px", letterSpacing: "0.08em", textTransform: "uppercase", display: "block", marginBottom: 6 }}>Password</label>
          <input style={{ ...css.input, width: "100%", boxSizing: "border-box" }} type="password" value={password}
            onChange={(e) => setPassword(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && handleLogin()} />
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

// ─── MODULO: CEDOLA ───────────────────────────────────────────────────────────
function ModuloDashboard({ titoli, prenotato, canali }) {
  // Raggruppa per giro_label
  const giri = useMemo(() => {
    const map = {};
    titoli.forEach(t => {
      const g = t.giro_label || "—";
      if (!map[g]) map[g] = { label: g, titoli: [], totObj: 0, totRag: 0, valObj: 0 };
      map[g].titoli.push(t);
      map[g].totObj += t.obiettivo_assegnato || 0;
      map[g].totRag += t.obiettivo_raggiunto || 0;
      map[g].valObj += (t.prezzo || 0) * (t.obiettivo_assegnato || 0);
    });
    return Object.values(map).sort((a, b) => {
      const [na, ya] = a.label.split(" "); const [nb, yb] = b.label.split(" ");
      return Number(yb || 0) - Number(ya || 0) || Number(nb || 0) - Number(na || 0);
    });
  }, [titoli]);

  // Raggruppa per cedola (n_cedola)
  const cedole = useMemo(() => {
    const map = {};
    titoli.forEach(t => {
      const c = t.n_cedola || "—";
      if (!map[c]) map[c] = { label: c, giro: t.giro_label, titoli: [], totObj: 0, totRag: 0 };
      map[c].titoli.push(t);
      map[c].totObj += t.obiettivo_assegnato || 0;
      map[c].totRag += t.obiettivo_raggiunto || 0;
    });
    return Object.values(map).sort((a, b) => (a.giro || "").localeCompare(b.giro || "") || (a.label || "").localeCompare(b.label || ""));
  }, [titoli]);

  const totPrenotato = prenotato.reduce((s, p) => s + p.quantita, 0);
  const totTriangolo = titoli.filter(t => t.il_triangolo === true).length;
  const totTop100 = titoli.filter(t => t.top_100 === true).length;

  const prenotatoPerCanale = useMemo(() => {
    const map = {};
    prenotato.forEach(p => {
      const c = canali.find(c => c.id === p.canale_id);
      if (!c) return;
      map[c.nome] = (map[c.nome] || 0) + p.quantita;
    });
    return Object.entries(map).sort((a, b) => b[1] - a[1]);
  }, [prenotato, canali]);
  const maxCanale = prenotatoPerCanale[0]?.[1] || 1;

  return (
    <div style={{ flex: 1, overflowY: "auto", padding: 20 }}>
      {/* KPI */}
      <div style={{ display: "flex", gap: 12, marginBottom: 24, flexWrap: "wrap" }}>
        <KpiCard label="Titoli totali" value={titoli.length} color={T.text} />
        <KpiCard label="▲ Triangolo" value={totTriangolo} color={T.purple} />
        <KpiCard label="★ Top 100" value={totTop100} color={T.accent} />
        <KpiCard label="Prenotato tot." value={totPrenotato.toLocaleString("it")} color={T.green} sub="copie" />
      </div>

      {/* Avanzamento per GIRO */}
      <div style={{ marginBottom: 24 }}>
        <div style={{ color: T.textMid, fontSize: "11px", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 12 }}>AVANZAMENTO PER GIRO</div>
        {giri.map(({ label, titoli: tList, totObj, totRag, valObj }) => {
          const pct = totObj > 0 ? Math.round((totRag / totObj) * 100) : 0;
          return (
            <div key={label} style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 4, padding: "12px 16px", display: "flex", alignItems: "center", gap: 16, marginBottom: 8 }}>
              <div style={{ width: 100, color: T.accent, fontWeight: "600", fontSize: "12px" }}>Giro {label}</div>
              <div style={{ color: T.textMid, fontSize: "11px", width: 80 }}>{tList.length} titoli</div>
              <div style={{ flex: 1 }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                  <span style={{ color: T.textMid, fontSize: "11px" }}>{totRag.toLocaleString("it")} / {totObj.toLocaleString("it")}</span>
                  <span style={{ color: pct >= 80 ? T.green : pct >= 50 ? T.accent : T.red, fontSize: "11px", fontWeight: "700" }}>{pct}%</span>
                </div>
                <div style={{ height: 6, background: T.borderHi, borderRadius: 3, overflow: "hidden" }}>
                  <div style={{ width: `${pct}%`, height: "100%", background: pct >= 80 ? T.green : pct >= 50 ? T.accent : T.red }} />
                </div>
              </div>
              <div style={{ color: T.textMid, fontSize: "11px", width: 110, textAlign: "right" }}>€ {valObj.toLocaleString("it", { maximumFractionDigits: 0 })}</div>
            </div>
          );
        })}
      </div>

      {/* Avanzamento per CEDOLA */}
      <div style={{ marginBottom: 24 }}>
        <div style={{ color: T.textMid, fontSize: "11px", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 12 }}>AVANZAMENTO PER CEDOLA</div>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                {["Cedola", "Giro", "Titoli", "Obj Ass.", "Obj Rag.", "Avanz."].map(h => (
                  <th key={h} style={{ padding: "8px 12px", textAlign: "left", color: T.textMid, fontWeight: "400", fontSize: "11px", letterSpacing: "0.08em", textTransform: "uppercase", borderBottom: `1px solid ${T.border}`, background: T.surface }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {cedole.map(({ label, giro, titoli: tList, totObj, totRag }, i) => {
                const pct = totObj > 0 ? Math.round((totRag / totObj) * 100) : 0;
                return (
                  <tr key={label} style={{ background: i % 2 === 0 ? "transparent" : T.surface + "66" }}>
                    <td style={{ padding: "8px 12px", color: T.accent, fontWeight: "600", fontSize: "12px" }}>{label}</td>
                    <td style={{ padding: "8px 12px", color: T.textMid, fontSize: "12px" }}>{giro}</td>
                    <td style={{ padding: "8px 12px", fontSize: "12px" }}>{tList.length}</td>
                    <td style={{ padding: "8px 12px", fontSize: "12px" }}>{totObj.toLocaleString("it")}</td>
                    <td style={{ padding: "8px 12px", fontSize: "12px" }}>{totRag.toLocaleString("it")}</td>
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
      </div>

      {/* Prenotato per canale */}
      {prenotatoPerCanale.length > 0 && (
        <div>
          <div style={{ color: T.textMid, fontSize: "11px", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 12 }}>PRENOTATO PER CANALE</div>
          {prenotatoPerCanale.map(([nome, qta]) => (
            <div key={nome} style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 6 }}>
              <div style={{ width: 160, fontSize: "12px" }}>{nome}</div>
              <div style={{ flex: 1, height: 8, background: T.borderHi, borderRadius: 2, overflow: "hidden" }}>
                <div style={{ width: `${(qta / maxCanale) * 100}%`, height: "100%", background: T.blue }} />
              </div>
              <div style={{ width: 70, textAlign: "right", color: T.accent, fontSize: "12px", fontWeight: "600" }}>{qta.toLocaleString("it")}</div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function ModuloDashboard({ titoli, prenotato, canali }) {
  // Raggruppa per giro_label
  const giri = useMemo(() => {
    const map = {};
    titoli.forEach(t => {
      const g = t.giro_label || "—";
      if (!map[g]) map[g] = { label: g, titoli: [], totObj: 0, totRag: 0, valObj: 0 };
      map[g].titoli.push(t);
      map[g].totObj += t.obiettivo_assegnato || 0;
      map[g].totRag += t.obiettivo_raggiunto || 0;
      map[g].valObj += (t.prezzo || 0) * (t.obiettivo_assegnato || 0);
    });
    return Object.values(map).sort((a, b) => {
      const [na, ya] = a.label.split(" "); const [nb, yb] = b.label.split(" ");
      return Number(yb || 0) - Number(ya || 0) || Number(nb || 0) - Number(na || 0);
    });
  }, [titoli]);

  // Raggruppa per cedola (n_cedola)
  const cedole = useMemo(() => {
    const map = {};
    titoli.forEach(t => {
      const c = t.n_cedola || "—";
      if (!map[c]) map[c] = { label: c, giro: t.giro_label, titoli: [], totObj: 0, totRag: 0 };
      map[c].titoli.push(t);
      map[c].totObj += t.obiettivo_assegnato || 0;
      map[c].totRag += t.obiettivo_raggiunto || 0;
    });
    return Object.values(map).sort((a, b) => (a.giro || "").localeCompare(b.giro || "") || (a.label || "").localeCompare(b.label || ""));
  }, [titoli]);

  const totPrenotato = prenotato.reduce((s, p) => s + p.quantita, 0);
  const totTriangolo = titoli.filter(t => t.il_triangolo === true).length;
  const totTop100 = titoli.filter(t => t.top_100 === true).length;

  const prenotatoPerCanale = useMemo(() => {
    const map = {};
    prenotato.forEach(p => {
      const c = canali.find(c => c.id === p.canale_id);
      if (!c) return;
      map[c.nome] = (map[c.nome] || 0) + p.quantita;
    });
    return Object.entries(map).sort((a, b) => b[1] - a[1]);
  }, [prenotato, canali]);
  const maxCanale = prenotatoPerCanale[0]?.[1] || 1;

  return (
    <div style={{ flex: 1, overflowY: "auto", padding: 20 }}>
      {/* KPI */}
      <div style={{ display: "flex", gap: 12, marginBottom: 24, flexWrap: "wrap" }}>
        <KpiCard label="Titoli totali" value={titoli.length} color={T.text} />
        <KpiCard label="▲ Triangolo" value={totTriangolo} color={T.purple} />
        <KpiCard label="★ Top 100" value={totTop100} color={T.accent} />
        <KpiCard label="Prenotato tot." value={totPrenotato.toLocaleString("it")} color={T.green} sub="copie" />
      </div>

      {/* Avanzamento per GIRO */}
      <div style={{ marginBottom: 24 }}>
        <div style={{ color: T.textMid, fontSize: "11px", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 12 }}>AVANZAMENTO PER GIRO</div>
        {giri.map(({ label, titoli: tList, totObj, totRag, valObj }) => {
          const pct = totObj > 0 ? Math.round((totRag / totObj) * 100) : 0;
          return (
            <div key={label} style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 4, padding: "12px 16px", display: "flex", alignItems: "center", gap: 16, marginBottom: 8 }}>
              <div style={{ width: 100, color: T.accent, fontWeight: "600", fontSize: "12px" }}>Giro {label}</div>
              <div style={{ color: T.textMid, fontSize: "11px", width: 80 }}>{tList.length} titoli</div>
              <div style={{ flex: 1 }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                  <span style={{ color: T.textMid, fontSize: "11px" }}>{totRag.toLocaleString("it")} / {totObj.toLocaleString("it")}</span>
                  <span style={{ color: pct >= 80 ? T.green : pct >= 50 ? T.accent : T.red, fontSize: "11px", fontWeight: "700" }}>{pct}%</span>
                </div>
                <div style={{ height: 6, background: T.borderHi, borderRadius: 3, overflow: "hidden" }}>
                  <div style={{ width: `${pct}%`, height: "100%", background: pct >= 80 ? T.green : pct >= 50 ? T.accent : T.red }} />
                </div>
              </div>
              <div style={{ color: T.textMid, fontSize: "11px", width: 110, textAlign: "right" }}>€ {valObj.toLocaleString("it", { maximumFractionDigits: 0 })}</div>
            </div>
          );
        })}
      </div>

      {/* Avanzamento per CEDOLA */}
      <div style={{ marginBottom: 24 }}>
        <div style={{ color: T.textMid, fontSize: "11px", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 12 }}>AVANZAMENTO PER CEDOLA</div>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                {["Cedola", "Giro", "Titoli", "Obj Ass.", "Obj Rag.", "Avanz."].map(h => (
                  <th key={h} style={{ padding: "8px 12px", textAlign: "left", color: T.textMid, fontWeight: "400", fontSize: "11px", letterSpacing: "0.08em", textTransform: "uppercase", borderBottom: `1px solid ${T.border}`, background: T.surface }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {cedole.map(({ label, giro, titoli: tList, totObj, totRag }, i) => {
                const pct = totObj > 0 ? Math.round((totRag / totObj) * 100) : 0;
                return (
                  <tr key={label} style={{ background: i % 2 === 0 ? "transparent" : T.surface + "66" }}>
                    <td style={{ padding: "8px 12px", color: T.accent, fontWeight: "600", fontSize: "12px" }}>{label}</td>
                    <td style={{ padding: "8px 12px", color: T.textMid, fontSize: "12px" }}>{giro}</td>
                    <td style={{ padding: "8px 12px", fontSize: "12px" }}>{tList.length}</td>
                    <td style={{ padding: "8px 12px", fontSize: "12px" }}>{totObj.toLocaleString("it")}</td>
                    <td style={{ padding: "8px 12px", fontSize: "12px" }}>{totRag.toLocaleString("it")}</td>
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
      </div>

      {/* Prenotato per canale */}
      {prenotatoPerCanale.length > 0 && (
        <div>
          <div style={{ color: T.textMid, fontSize: "11px", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 12 }}>PRENOTATO PER CANALE</div>
          {prenotatoPerCanale.map(([nome, qta]) => (
            <div key={nome} style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 6 }}>
              <div style={{ width: 160, fontSize: "12px" }}>{nome}</div>
              <div style={{ flex: 1, height: 8, background: T.borderHi, borderRadius: 2, overflow: "hidden" }}>
                <div style={{ width: `${(qta / maxCanale) * 100}%`, height: "100%", background: T.blue }} />
              </div>
              <div style={{ width: 70, textAlign: "right", color: T.accent, fontSize: "12px", fontWeight: "600" }}>{qta.toLocaleString("it")}</div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}


// ─── MODULO: FINE GIRO ────────────────────────────────────────────────────────
function ModuloFineGiro({ titoli, giriList, prenotato, canali }) {
  const [giroSel, setGiroSel] = useState(giriList[0]?.id ?? null);
  const titoliGiro = titoli.filter((t) => t.giro_id === giroSel);
  const totObj = titoliGiro.reduce((s, t) => s + (t.obiettivo_assegnato || 0), 0);
  const totRag = titoliGiro.reduce((s, t) => s + (t.obiettivo_raggiunto || 0), 0);
  const canaliPrincipali = canali.slice(0, 5);

  const righe = useMemo(() => titoliGiro.map((t) => {
    const pren = prenotato.filter((p) => p.titolo_id === t.id);
    const totPren = pren.reduce((s, p) => s + p.quantita, 0);
    const byCanale = {};
    pren.forEach((p) => { const c = canali.find((c) => c.id === p.canale_id); if (c) byCanale[c.codice] = p.quantita; });
    return { titolo: t, totPren, byCanale };
  }).sort((a, b) => (a.titolo.ranking_editore ?? 99) - (b.titolo.ranking_editore ?? 99) || (a.titolo.ranking_titolo ?? 99) - (b.titolo.ranking_titolo ?? 99)), [titoliGiro, prenotato, canali]);

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%", overflow: "hidden" }}>
      <div style={{ padding: "12px 20px", borderBottom: `1px solid ${T.border}`, display: "flex", gap: 10, alignItems: "center" }}>
        <select style={css.input} value={giroSel ?? ""} onChange={(e) => setGiroSel(Number(e.target.value))}>
          {giriList.map((g) => <option key={g.id} value={g.id}>{g.label}</option>)}
        </select>
        <div style={{ color: T.textMid, fontSize: "12px" }}>
          Obj: <span style={{ color: T.accent, fontWeight: "700" }}>{totObj}</span> &nbsp;·&nbsp;
          Raggiunto: <span style={{ color: T.green, fontWeight: "700" }}>{totRag}</span> &nbsp;·&nbsp;
          <span style={{ color: T.accent, fontWeight: "700" }}>{totObj > 0 ? Math.round(totRag / totObj * 100) : 0}%</span>
        </div>
      </div>
      <div style={{ flex: 1, overflowY: "auto" }}>
        <table style={css.table}>
          <thead>
            <tr>
              <th style={css.th}>Editore</th><th style={css.th}>Titolo</th><th style={css.th}>EAN</th>
              <th style={css.th}>Obj. Ass.</th><th style={css.th}>Obj. Rag.</th><th style={css.th}>Avanz.</th>
              {canaliPrincipali.map((c) => <th key={c.id} style={css.th}>{c.nome}</th>)}
              <th style={css.th}>Tot. Pren.</th>
            </tr>
          </thead>
          <tbody>
            {righe.map(({ titolo: t, totPren, byCanale }, i) => {
              const pct = t.obiettivo_assegnato > 0 ? Math.round(t.obiettivo_raggiunto / t.obiettivo_assegnato * 100) : 0;
              return (
                <tr key={t.id} style={{ background: i % 2 === 0 ? "transparent" : T.surface + "66" }}>
                  <td style={{ ...css.td, color: T.accent, fontWeight: "600" }}>{t.editore_nome}</td>
                  <td style={{ ...css.td, maxWidth: 200 }}><div style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{t.titolo}</div></td>
                  <td style={{ ...css.td, fontFamily: "monospace", fontSize: "11px", color: T.textMid }}>{t.ean}</td>
                  <td style={css.td}>{t.obiettivo_assegnato}</td>
                  <td style={css.td}>{t.obiettivo_raggiunto}</td>
                  <td style={css.td}><span style={{ color: pct >= 80 ? T.green : pct >= 50 ? T.accent : T.red, fontWeight: "700" }}>{pct}%</span></td>
                  {canaliPrincipali.map((c) => (<td key={c.id} style={{ ...css.td, color: byCanale[c.codice] ? T.text : T.textDim }}>{byCanale[c.codice] ?? "—"}</td>))}
                  <td style={{ ...css.td, color: totPren > 0 ? T.green : T.textDim, fontWeight: "700" }}>{totPren || "—"}</td>
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
  { id: "cedola", label: "Cedola", icon: "≡" },
  { id: "prenotato", label: "Prenotato", icon: "↳" },
  { id: "finegiro", label: "Fine Giro", icon: "⊞" },
  { id: "import", label: "Import Cedola", icon: "↑" },
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
  const [giriDB, setGiriDB] = useState([]);
  
  useEffect(() => {
    if (!session) return;
    sbFetch("giri?select=*&order=anno.desc,numero.desc", session.token).then(setGiriDB);
    sbFetch("titoli?select=*&order=ranking_editore.asc,ranking_titolo.asc", session.token).then(setTitoli);
    sbFetch("prenotato?select=*", session.token).then(setPrenotato);
    sbFetch("spalmatura_obiettivo?select=*", session.token).then(setSpalmatura);
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
    setTitoli((prev) => prev.map((t) => (t.id === updated.id ? { ...t, ...updated } : t)));
  }, []);

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
          {MODULES.map((m) => (
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
          <span style={{ color: T.accent, fontSize: "11px", letterSpacing: "0.12em", textTransform: "uppercase" }}>{MODULES.find((m) => m.id === activeModule)?.label}</span>
          <span style={{ color: T.borderHi }}>·</span>
          <span style={{ color: T.textMid, fontSize: "11px" }}>{titoli.length} titoli · {giriDB.length} cedole</span>
        </div>
        <div style={{ flex: 1, overflow: "hidden", display: "flex", flexDirection: "column" }}>
          {activeModule === "import" && <ModuloImport giriList={giriDB} token={session.token} />}
          {activeModule === "dashboard" && <ModuloDashboard titoli={titoli} giriList={giriDB} prenotato={prenotato} canali={MOCK_CANALI} />}
          {activeModule === "cedola" && <ModuloCedola titoli={titoli} giriList={giriDB} onUpdateTitolo={updateTitolo} spalmatura={spalmatura} />}
          {activeModule === "prenotato" && <ModuloPrenotato token={session.token} titoli={titoli} />}
          {activeModule === "finegiro" && <ModuloFineGiro titoli={titoli} giriList={giriDB} prenotato={prenotato} canali={MOCK_CANALI} />}
        </div>
      </div>
    </div>
  );
}
