// Tema chiaro / scuro, condiviso con il PDE Hub e le altre app PDE: la scelta sta in localStorage
// con la stessa chiave del Hub ("pde-hub-theme": "light" | "dark"); le app sono tutte su
// albertospde.github.io (stessa origine), quindi una scelta fatta nel Hub vale anche qui e viceversa.
// index.html imposta html[data-theme] e lo sfondo prima del caricamento (niente lampeggio).
//
// I colori dell'app sono negli oggetti T dei vari moduli: in tema scuro restano quelli di sempre
// (tema(T) li restituisce invariati), in tema chiaro vengono sostituiti dalla tavolozza chiara.
// Cambiare tema ricarica la pagina: gli stili sono calcolati quando i moduli vengono caricati.

export const CHIAVE_TEMA = "pde-hub-theme";

const leggi = () => { try { return localStorage.getItem(CHIAVE_TEMA) === "light" ? "light" : "dark"; } catch { return "dark"; } };

export const temaAttivo = leggi();
export const CHIARO = temaAttivo === "light";

// Stessi valori del tema chiaro del Hub; in esadecimale perché il codice aggiunge la trasparenza
// in coda al colore (es. T.accent + "18").
const TAVOLOZZA_CHIARA = {
  bg: "#eef2f9", surface: "#ffffff", border: "#d3d9e6", borderHi: "#b9c3d8",
  text: "#16204a", textMid: "#5a6388", textDim: "#8a92b2",
  accent: "#0092c2", green: "#2e9e62", red: "#d14343", amber: "#b7791f", orange: "#b7791f",
  blue: "#3f5fb8", purple: "#7d55c7",
};

// Oggetto colori di un modulo: invariato in tema scuro, tavolozza chiara in tema chiaro
export const tema = (scuro) => (CHIARO ? { ...scuro, ...TAVOLOZZA_CHIARA } : scuro);

// Per i pochi colori scritti fissi nel codice: valore per il tema scuro e per quello chiaro
export const cv = (scuro, chiaro) => (CHIARO ? chiaro : scuro);

export function impostaTema(t) {
  try { localStorage.setItem(CHIAVE_TEMA, t === "light" ? "light" : "dark"); } catch { /* storage non disponibile */ }
  if ((t === "light" ? "light" : "dark") !== temaAttivo) location.reload();
}

// Tema cambiato da un'altra scheda o dal Hub che contiene l'app: si allinea
window.addEventListener("storage", (e) => {
  if (e.key === CHIAVE_TEMA && (e.newValue === "light" ? "light" : "dark") !== temaAttivo) location.reload();
});
