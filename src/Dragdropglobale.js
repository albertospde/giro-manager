// ═══════════════════════════════════════════════════════════════════════════
// TRASCINAMENTO FILE GLOBALE
// Rende "trascinabili" tutti i caricamenti file di GiroManager senza toccare i
// singoli moduli: il file lasciato cadere viene passato al campo file più vicino
// al punto di rilascio (o all'unico visibile a schermo), come se fosse stato
// scelto con "Sfoglia". I moduli che gestiscono già il drop da soli
// (es. Import da file editore) hanno la precedenza.
// ═══════════════════════════════════════════════════════════════════════════

const CLASSE = "gm-drop-target";

function iniettaStile() {
  if (document.getElementById("gm-drop-style")) return;
  const st = document.createElement("style");
  st.id = "gm-drop-style";
  st.textContent = `.${CLASSE}{outline:2px dashed #c8a96e!important;outline-offset:4px;border-radius:4px;background-color:#c8a96e14!important;transition:all .15s}`;
  document.head.appendChild(st);
}

const haFile = (e) => Array.from(e.dataTransfer?.types || []).includes("Files");
const visibile = (el) => !!el && el.getClientRects().length > 0;
// un input nascosto (display:none) conta come visibile se lo è il suo contenitore / label
const inputUsabile = (inp) => !inp.disabled && (visibile(inp) || visibile(inp.parentElement) || (inp.id && visibile(document.querySelector(`label[for="${CSS.escape(inp.id)}"]`))));

function trovaInput(x, y) {
  // 1) il campo file più vicino risalendo dal punto di rilascio
  let el = document.elementFromPoint(x, y);
  while (el && el !== document.body) {
    const inp = Array.from(el.querySelectorAll?.('input[type="file"]') || []).filter(inputUsabile);
    if (inp.length) return inp[0];
    el = el.parentElement;
  }
  // 2) altrimenti l'unico campo file presente a schermo
  const tutti = Array.from(document.querySelectorAll('input[type="file"]')).filter(inputUsabile);
  return tutti.length === 1 ? tutti[0] : null;
}

function accetta(inp, file) {
  const acc = (inp.getAttribute("accept") || "").split(",").map(s => s.trim().toLowerCase()).filter(Boolean);
  if (!acc.length) return true;
  const nome = file.name.toLowerCase(), tipo = (file.type || "").toLowerCase();
  return acc.some(a => a.startsWith(".") ? nome.endsWith(a) : a.endsWith("/*") ? tipo.startsWith(a.slice(0, -1)) : tipo === a);
}

const contenitore = (inp) => (inp.id && document.querySelector(`label[for="${CSS.escape(inp.id)}"]`)?.parentElement) || inp.parentElement;

export function attivaDragDropGlobale() {
  if (window.__gmDragDrop) return;
  window.__gmDragDrop = true;
  iniettaStile();
  let evidenziato = null;
  const evidenzia = (el) => {
    if (evidenziato === el) return;
    evidenziato?.classList.remove(CLASSE);
    evidenziato = el;
    el?.classList.add(CLASSE);
  };

  window.addEventListener("dragover", (e) => {
    if (!haFile(e)) return;
    e.preventDefault(); // evita che il browser apra il file al posto dell'app
    const inp = trovaInput(e.clientX, e.clientY);
    e.dataTransfer.dropEffect = inp ? "copy" : "none";
    evidenzia(inp ? contenitore(inp) : null);
  });
  window.addEventListener("dragleave", (e) => { if (!e.relatedTarget) evidenzia(null); });
  window.addEventListener("drop", (e) => {
    if (!haFile(e)) return;
    const giaGestito = e.defaultPrevented; // un modulo con drop proprio l'ha già preso
    e.preventDefault();
    evidenzia(null);
    if (giaGestito) return;
    const inp = trovaInput(e.clientX, e.clientY);
    if (!inp) return;
    const files = Array.from(e.dataTransfer.files || []).filter(f => accetta(inp, f));
    if (!files.length) { alert(`Formato non supportato. File accettati: ${inp.getAttribute("accept") || "qualsiasi"}`); return; }
    const dt = new DataTransfer();
    (inp.multiple ? files : files.slice(0, 1)).forEach(f => dt.items.add(f));
    inp.files = dt.files;
    inp.dispatchEvent(new Event("change", { bubbles: true }));
  });
}
