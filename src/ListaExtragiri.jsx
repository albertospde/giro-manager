import { useState } from "react";

// Raggruppa le cedole extra: quelle con "campagna"/"campagne" nel nome
// finiscono in un gruppo a parte, le altre restano nell'elenco piatto.
const RE_CAMPAGNA = /campagn[ae]/i;

export function raggruppaExtragiri(cedole, campoNome = "n_cedola") {
  const campagne = cedole.filter((c) => RE_CAMPAGNA.test(c[campoNome] ?? ""));
  const altre = cedole.filter((c) => !RE_CAMPAGNA.test(c[campoNome] ?? ""));
  return { campagne, altre };
}

// Bottone stile "pill" coerente con lo screenshot (bordo blu, testo blu)
function CedolaButton({ label, onClick }) {
  return (
    <button
      onClick={onClick}
      style={{
        width: "100%",
        padding: "10px 12px",
        marginBottom: 8,
        background: "transparent",
        border: "1px solid #5b7fd4",
        borderRadius: 6,
        color: "#8fb0ff",
        fontSize: 12,
        fontWeight: 600,
        letterSpacing: 0.3,
        cursor: "pointer",
        textAlign: "center",
      }}
    >
      {label}
    </button>
  );
}

export default function ListaExtragiri({ cedole, campoNome = "n_cedola", onSelect }) {
  const [aperto, setAperto] = useState(false);
  const { campagne, altre } = raggruppaExtragiri(cedole, campoNome);

  return (
    <div>
      {altre.map((c) => (
        <CedolaButton key={c.id} label={c[campoNome]} onClick={() => onSelect(c)} />
      ))}

      {campagne.length > 0 && (
        <>
          <button
            onClick={() => setAperto((v) => !v)}
            style={{
              width: "100%",
              padding: "10px 12px",
              marginBottom: 8,
              background: "#5b7fd4",
              border: "none",
              borderRadius: 6,
              color: "#0d1330",
              fontSize: 12,
              fontWeight: 700,
              cursor: "pointer",
              textAlign: "center",
            }}
          >
            {aperto ? "▾" : "▸"} CAMPAGNE ({campagne.length})
          </button>

          {aperto &&
            campagne.map((c) => (
              <CedolaButton key={c.id} label={c[campoNome]} onClick={() => onSelect(c)} />
            ))}
        </>
      )}
    </div>
  );
}
