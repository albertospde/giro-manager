// Tema chiaro / scuro, condiviso con il PDE Hub e le altre app PDE: la scelta sta in localStorage
// con la stessa chiave del Hub ("pde-hub-theme": "light" | "dark"); le app sono tutte su
// albertospde.github.io (stessa origine), quindi una scelta fatta nel Hub vale anche qui e viceversa.
// index.html imposta html[data-theme] e lo sfondo prima del caricamento (niente lampeggio).
//
// Grafica allineata a BookUp: stessa tavolozza (chiara e scura), font DM Sans / Bree Serif,
// bottoni chiari con bordo, card arrotondate, tabelle con intestazioni maiuscole.
// I colori dell'app sono negli oggetti T dei vari moduli: tema(T) sostituisce i colori base
// con la tavolozza BookUp del tema attivo (le chiavi in più del modulo restano).
// Cambiare tema ricarica la pagina: gli stili sono calcolati quando i moduli vengono caricati.

export const CHIAVE_TEMA = "pde-hub-theme";

const leggi = () => { try { return localStorage.getItem(CHIAVE_TEMA) === "light" ? "light" : "dark"; } catch { return "dark"; } };

export const temaAttivo = leggi();
export const CHIARO = temaAttivo === "light";

// Tavolozze di BookUp (variabili CSS --bg, --card, --card-2, --cyan, --text, --text-dim, --border…).
// In esadecimale a 6 cifre perché il codice aggiunge la trasparenza in coda al colore (es. T.accent + "18").
const TAVOLOZZA_CHIARA = {
  bg: "#f4f7fb", surface: "#ffffff", surface2: "#eef4fa", border: "#d7e0ea", borderHi: "#b9c7d6",
  text: "#132033", textMid: "#5b6b7d", textDim: "#8796a8",
  accent: "#0090c8", green: "#1a7f37", red: "#c62828", amber: "#b06a00", orange: "#c2620a",
  blue: "#3f5fb8", purple: "#7e22ce",
};
const TAVOLOZZA_SCURA = {
  bg: "#0d1334", surface: "#161f47", surface2: "#1c2754", border: "#2f3860", borderHi: "#434d78",
  text: "#e8ecf8", textMid: "#97a3c9", textDim: "#6873a0",
  accent: "#33c3f0", green: "#4fd18b", red: "#ff7a7a", amber: "#f0b84c", orange: "#f0a050",
  blue: "#6e8cff", purple: "#b18cf5",
};
const P = CHIARO ? TAVOLOZZA_CHIARA : TAVOLOZZA_SCURA;

// Oggetto colori di un modulo: i colori base diventano quelli BookUp del tema attivo
export const tema = (scuro) => ({ ...scuro, ...P });

// Per i pochi colori scritti fissi nel codice: valore per il tema scuro e per quello chiaro
export const cv = (scuro, chiaro) => (CHIARO ? chiaro : scuro);

// ---------- Stili condivisi (come i componenti base di BookUp) ----------
export const FONT = "'DM Sans', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif";
export const FONT_TITOLI = "'Bree Serif', Georgia, serif";
// Testo sui bottoni pieni (accent / green): bianco come BookUp; nel tema scuro l'azzurro è chiaro e serve il blu notte
export const SU_ACCENT = cv("#0d1334", "#ffffff");
export const OMBRA_CARD = cv("0 4px 18px rgba(0,0,0,.35)", "0 2px 10px rgba(20,45,90,0.05)");
export const OMBRA_POPUP = cv("0 14px 36px rgba(0,0,0,.5)", "0 14px 36px rgba(19,32,51,.18)");

// Bottone: "default" = chiaro con bordo (btn-secondary), "attivo" = selezionato (btn-secondary.active),
// "accent"/"green" = pieno (btn-primary), "danger" = rosso tenue
export const stileBtn = (v = "default", disabled = false) => {
  const pieno = v === "accent" || v === "green";
  const attivo = v === "attivo";
  const col = v === "green" ? P.green : v === "danger" ? P.red : P.accent;
  return {
    padding: "6px 14px", borderRadius: 10, fontSize: "13px", fontFamily: "inherit", lineHeight: 1.3,
    border: `2px solid ${pieno || attivo ? col : v === "danger" ? P.red + "88" : P.border}`,
    background: pieno ? col : attivo ? P.accent + "1c" : v === "danger" ? P.red + "14" : P.surface,
    color: pieno ? SU_ACCENT : attivo ? P.accent : v === "danger" ? P.red : P.text,
    fontWeight: pieno || attivo ? 700 : 600, cursor: disabled ? "not-allowed" : "pointer", opacity: disabled ? 0.45 : 1,
    transition: "border-color .15s, color .15s, background .15s, filter .15s",
  };
};
// Bottoncino nelle righe delle tabelle
export const stileMini = { padding: "2px 8px", border: `1px solid ${P.border}`, background: P.surface, color: P.textMid, cursor: "pointer", fontSize: "11px", fontFamily: "inherit", borderRadius: 7, fontWeight: 600 };
export const stileInput = { background: P.surface, border: `2px solid ${P.border}`, color: P.text, padding: "5px 10px", fontSize: "13px", fontFamily: "inherit", borderRadius: 8, outline: "none" };
export const stileTh = { padding: "9px 10px", textAlign: "left", color: P.textMid, fontWeight: 700, fontSize: "11px", letterSpacing: "0.03em", textTransform: "uppercase", borderBottom: `2px solid ${P.border}`, whiteSpace: "nowrap" };
export const stileTd = { padding: "8px 10px", borderBottom: `1px solid ${P.border}`, verticalAlign: "middle" };
export const stileCard = { background: P.surface, border: `1px solid ${P.border}`, borderRadius: 14, boxShadow: OMBRA_CARD };
export const stileKpi = { background: P.surface2, border: `1px solid ${P.border}`, borderRadius: 9 };
export const stileTag = (c = P.accent) => ({ display: "inline-block", padding: "2px 8px", background: c + "1c", border: `1px solid ${c}44`, color: c, borderRadius: 20, fontSize: "10.5px", fontWeight: 700, letterSpacing: "0.02em", textTransform: "uppercase", whiteSpace: "nowrap" });
export const stileLabel = { color: P.textMid, fontSize: "12px", fontWeight: 600 };

export function impostaTema(t) {
  try { localStorage.setItem(CHIAVE_TEMA, t === "light" ? "light" : "dark"); } catch { /* storage non disponibile */ }
  if ((t === "light" ? "light" : "dark") !== temaAttivo) location.reload();
}

// Tema cambiato da un'altra scheda o dal Hub che contiene l'app: si allinea
window.addEventListener("storage", (e) => {
  if (e.key === CHIAVE_TEMA && (e.newValue === "light" ? "light" : "dark") !== temaAttivo) location.reload();
});
