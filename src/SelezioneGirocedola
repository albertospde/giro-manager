// Divide le cedole extra in due liste: quelle con "campagna"/"campagne" nel nome
// vanno nella colonna CAMPAGNE, le altre restano in EXTRAGIRI.
const RE_CAMPAGNA = /campagn[ae]/i;

export function dividiExtragiriECampagne(cedoleExtra, campoNome = "n_cedola") {
  const campagne = cedoleExtra.filter((c) => RE_CAMPAGNA.test(c[campoNome] ?? ""));
  const extragiri = cedoleExtra.filter((c) => !RE_CAMPAGNA.test(c[campoNome] ?? ""));
  return { extragiri, campagne };
}

function Bottone({ label, variante = "extra", onClick }) {
  const stili = {
    giro: { background: "#8fa8e8", border: "none", color: "#12183a" },
    extra: { background: "transparent", border: "1px solid #5b7fd4", color: "#8fb0ff" },
    campagna: { background: "transparent", border: "1px solid #7a5bd4", color: "#b08fff" },
  };
  return (
    <button
      onClick={onClick}
      style={{
        width: "100%",
        padding: "10px 12px",
        marginBottom: 8,
        borderRadius: 6,
        fontSize: 12,
        fontWeight: 600,
        letterSpacing: 0.3,
        cursor: "pointer",
        textAlign: "center",
        ...stili[variante],
      }}
    >
      {label}
    </button>
  );
}

function Colonna({ titolo, anno, onCambiaAnno, children }) {
  return (
    <div style={{ minWidth: 220 }}>
      <div style={{ textAlign: "center", fontWeight: 700, color: "#fff", marginBottom: 12, fontSize: 13 }}>
        {titolo}
      </div>
      {onCambiaAnno && (
        <select
          value={anno}
          onChange={(e) => onCambiaAnno(e.target.value)}
          style={{
            display: "block",
            margin: "0 auto 16px",
            background: "#8fa8e8",
            border: "none",
            borderRadius: 4,
            padding: "6px 10px",
            fontWeight: 600,
          }}
        >
          <option value={anno}>{anno}</option>
        </select>
      )}
      {children}
    </div>
  );
}

export default function SelezioneGiroCedola({
  giri,               // array giri vendita, es. [{id, label:"Giro 5 2026"}]
  cedoleExtra,         // array cedole extra grezze (campo nome = n_cedola)
  anno,
  onCambiaAnno,
  onSelectGiro,
  onSelectCedola,
  campoNome = "n_cedola",
}) {
  const { extragiri, campagne } = dividiExtragiriECampagne(cedoleExtra, campoNome);

  return (
    <div>
      <div style={{ textAlign: "center", color: "#8fb0ff", marginBottom: 24 }}>
        Seleziona un giro o una cedola extra
      </div>
      <div style={{ display: "flex", gap: 48, justifyContent: "center" }}>
        <Colonna titolo="GIRI VENDITA" anno={anno} onCambiaAnno={onCambiaAnno}>
          {giri.map((g) => (
            <Bottone key={g.id} label={g.label} variante="giro" onClick={() => onSelectGiro(g)} />
          ))}
        </Colonna>

        <Colonna titolo="EXTRAGIRI" anno={anno} onCambiaAnno={onCambiaAnno}>
          {extragiri.map((c) => (
            <Bottone key={c.id} label={c[campoNome]} variante="extra" onClick={() => onSelectCedola(c)} />
          ))}
        </Colonna>

        <Colonna titolo="CAMPAGNE" anno={anno} onCambiaAnno={onCambiaAnno}>
          {campagne.map((c) => (
            <Bottone key={c.id} label={c[campoNome]} variante="campagna" onClick={() => onSelectCedola(c)} />
          ))}
        </Colonna>
      </div>
    </div>
  );
}
