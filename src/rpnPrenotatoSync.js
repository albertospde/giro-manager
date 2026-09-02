// Logica condivisa per l'aggiornamento automatico del "prenotato" da RPN.
// Usata sia da ModuloPrenotato (import manuale + automatico, con Verifica dettagliata)
// sia da ModuloFineGiro (bottone rapido "Aggiorna da RPN" dalla pagina Fine Giro).
// Tenerla in un solo posto evita che le due schermate divergano nel calcolo.

const SUPABASE_URL = "https://tdflwenlylhctxssatax.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InRkZmx3ZW5seWxoY3R4c3NhdGF4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYzMzgyNzYsImV4cCI6MjA5MTkxNDI3Nn0.l35qEL7LOvyYuI1McQlVqj4vbyTqmlevcmqWbTGYi2Q";

export const GRUPPI_CANALE = {
  2: "FELTRINELLI", 23: "FELTRINELLI", 59: "FELTRINELLI", 77: "FELTRINELLI",
  8: "MONDADORI", 34: "MONDADORI", 80: "MONDADORI",
  83: "UBIK",
  32: "GIUNTI",
  11: "LIBRACCIO", 56: "LIBRACCIO", 88: "LIBRACCIO", 90: "LIBRACCIO", 91: "LIBRACCIO", 92: "LIBRACCIO",
  6: "LIB_RELIGIOSE", 18: "LIB_RELIGIOSE", 19: "LIB_RELIGIOSE", 21: "LIB_RELIGIOSE", 57: "LIB_RELIGIOSE",
  36: "LIB_COOP",
  4: "INDIPENDENTI_ALTRE_CATENE", 22: "INDIPENDENTI_ALTRE_CATENE", 24: "INDIPENDENTI_ALTRE_CATENE", 60: "INDIPENDENTI_ALTRE_CATENE",
  28: "FASTBOOK",
  63: "CENTROLIBRI",
  25: "GROSSISTI", 30: "GROSSISTI", 94: "GROSSISTI",
  82: "AMAZON",
  58: "IBS",
  33: "ALTRI_ONLINE",
};

export const CANALI_LABELS = {
  FELTRINELLI: "Feltrinelli", GIUNTI: "Giunti al Punto", MONDADORI: "Mondadori",
  UBIK: "Ubik", LIBRACCIO: "Libraccio", INDIPENDENTI_ALTRE_CATENE: "Indipendenti",
  LIB_RELIGIOSE: "Librerie Religiose", LIB_COOP: "Librerie Coop", ALTRI_ONLINE: "Librerie On-line",
  AMAZON: "Amazon", IBS: "Stereo Online", FASTBOOK: "Fastbook + GD", GROSSISTI: "Grossisti",
  CENTROLIBRI: "Centro Libri",
};

function leggiCella(row, nomeColonna, indice, headers) {
  if (row[nomeColonna] !== undefined && row[nomeColonna] !== "") return row[nomeColonna];
  const key = Object.keys(row).find(k => k.trim().toLowerCase() === nomeColonna.toLowerCase());
  if (key && row[key] !== "") return row[key];
  if (headers && indice < headers.length) {
    const hKey = headers[indice];
    if (hKey !== undefined && row[hKey] !== undefined) return row[hKey];
  }
  return "";
}

// Sceglie, tra i titoli con lo stesso EAN, quello del giro più recente
// (ristampe/relanci su giri diversi: preferiamo sempre l'ultimo).
function trovaTitoloPerEan(titoli, ean) {
  const candidati = titoli.filter(t => t.ean === ean || t.ean === String(parseInt(ean)));
  if (candidati.length <= 1) return candidati[0];
  return [...candidati].sort((a, b) => {
    const annoA = Number((a.giro_label || "").split(" ")[1]) || 0;
    const annoB = Number((b.giro_label || "").split(" ")[1]) || 0;
    if (annoB !== annoA) return annoB - annoA;
    const numA = Number((a.giro_label || "").split(" ")[0]) || 0;
    const numB = Number((b.giro_label || "").split(" ")[0]) || 0;
    if (numB !== numA) return numB - numA;
    return (b.id || 0) - (a.id || 0);
  })[0];
}

// Scarica il file "Pianifica Visite" da RPN tramite la edge function
// (login + download RPN gestiti lato server, riusa le credenziali già collegate in BookUp).
export async function fetchPrenotatoRpn(token, giroLabel) {
  return fetchPrenotatoRpnByParam(token, "giro_label", giroLabel);
}

// Cedole extra (non associate a nessun giro su RPN): si selezionano per nome cedola
// invece che per giro, con un endpoint diverso lato edge function.
export async function fetchPrenotatoRpnCedola(token, cedolaNome) {
  return fetchPrenotatoRpnByParam(token, "cedola_nome", cedolaNome);
}

async function fetchPrenotatoRpnByParam(token, paramName, paramValue) {
  const res = await fetch(
    `${SUPABASE_URL}/functions/v1/giro-prenotato-sync?${paramName}=${encodeURIComponent(paramValue)}`,
    { headers: { Authorization: `Bearer ${token}`, apikey: SUPABASE_KEY } }
  );
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    if (body.error === "RPN_NOT_CONNECTED") {
      throw new Error("Account RPN non collegato: ricollega il tuo account RPN in BookUp.");
    }
    throw new Error(body.message || body.error || `Errore ${res.status}`);
  }
  return res.arrayBuffer();
}

// Parsing + aggregazione (per ean+canale e per cliente+ean+canale), identico sia che
// il file arrivi da upload manuale sia da download automatico RPN.
export function parseEaggrega(arrayBuffer, titoli) {
  const XLSX = window.XLSX;
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = wb.SheetNames.find(n => n.toLowerCase().includes("pianifica")) || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];

  const dataRaw = XLSX.utils.sheet_to_json(ws, { defval: "", header: 1 });
  const headers = dataRaw[0] ? dataRaw[0].map(h => String(h || "").trim()) : [];
  const data = XLSX.utils.sheet_to_json(ws, { defval: "" });

  const aggMap = {};
  const aggCliMap = {};

  data.forEach(row => {
    const ean = String(row["EAN"] || "").trim();
    const gruppoRaw = row["Gruppo cliente"];
    const qta = parseInt(row["Pren (Qtà)"]) || 0;
    const codiceCliente = String(row["Codice cliente"] || "").trim();
    const nomeCliente = String(row["Nome Cliente"] || "").trim();
    if (!ean || qta === 0) return;

    const numOrdineCliente = String(leggiCella(row, "N. ordine cliente", 5, headers) || leggiCella(row, "Ordine cliente", 5, headers) || "").trim();
    const scontoOccRaw = leggiCella(row, "Sconto occasionale", 18, headers) || leggiCella(row, "Sc. occas.", 18, headers) || "";
    const scontoOcc = parseFloat(String(scontoOccRaw).replace(",", ".")) || 0;
    const pagamentoOcc = String(leggiCella(row, "Pagamento occasionale", 20, headers) || leggiCella(row, "Pag. occas.", 20, headers) || leggiCella(row, "Pag(occ)", 20, headers) || "").trim();

    let canale = "INDIPENDENTI_ALTRE_CATENE";
    if (gruppoRaw !== "" && gruppoRaw !== null && gruppoRaw !== undefined) {
      const gruppoInt = parseInt(parseFloat(gruppoRaw));
      if (!isNaN(gruppoInt) && gruppoInt !== 0) canale = GRUPPI_CANALE[gruppoInt] || "INDIPENDENTI_ALTRE_CATENE";
    }

    const key = `${ean}__${canale}`;
    if (!aggMap[key]) aggMap[key] = { ean, canale, qta: 0 };
    aggMap[key].qta += qta;

    if (codiceCliente) {
      const keyC = `${codiceCliente}__${ean}__${canale}`;
      if (!aggCliMap[keyC]) aggCliMap[keyC] = {
        codice_cliente: codiceCliente, nome_cliente: nomeCliente, ean, canale, qta: 0,
        sconto_occasionale: scontoOcc > 0 ? scontoOcc : null,
        pagamento_occasionale: pagamentoOcc || null,
        num_ordine_cliente: numOrdineCliente || null,
      };
      aggCliMap[keyC].qta += qta;
      if (scontoOcc > 0 && !aggCliMap[keyC].sconto_occasionale) aggCliMap[keyC].sconto_occasionale = scontoOcc;
      if (pagamentoOcc && !aggCliMap[keyC].pagamento_occasionale) aggCliMap[keyC].pagamento_occasionale = pagamentoOcc;
      if (numOrdineCliente && !aggCliMap[keyC].num_ordine_cliente) aggCliMap[keyC].num_ordine_cliente = numOrdineCliente;
    }
  });

  const aggregato = Object.values(aggMap).map(r => {
    const titolo = trovaTitoloPerEan(titoli, r.ean);
    return { ...r, titolo: titolo?.titolo ?? "— non trovato —", found: !!titolo, titolo_id: titolo?.id };
  }).sort((a, b) => a.ean.localeCompare(b.ean));

  const aggregatoClienti = Object.values(aggCliMap).map(r => {
    const titolo = trovaTitoloPerEan(titoli, r.ean);
    return { ...r, found: !!titolo, titolo_id: titolo?.id };
  });

  const totaleAggregato = aggregato.reduce((s, r) => s + r.qta, 0);
  const totaleFound = aggregato.filter(r => r.found).reduce((s, r) => s + r.qta, 0);
  const riepilogoCanale = Object.entries(
    aggregato.reduce((m, r) => { m[r.canale] = (m[r.canale] || 0) + r.qta; return m; }, {})
  ).sort((a, b) => b[1] - a[1]);

  return { righe: data, aggregato, aggregatoClienti, totaleAggregato, totaleFound, riepilogoCanale };
}

// Scrive aggregato/aggregatoClienti su Supabase tramite le stesse RPC dell'import manuale.
export async function importAggregato(token, aggregato, aggregatoClienti) {
  const rCanali = await fetch(`${SUPABASE_URL}/rest/v1/canali?select=id,codice`, {
    headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${token}` },
  });
  const canaliDB = await rCanali.json();
  const canaleMap = {};
  canaliDB.forEach(c => { canaleMap[c.codice] = c.id; });

  const validi = aggregato.filter(r => r.found && r.titolo_id);
  const payload = validi.map(r => ({
    titolo_id: r.titolo_id, canale_id: canaleMap[r.canale] || null, quantita: r.qta,
  })).filter(r => r.canale_id !== null);

  const res1 = await fetch(`${SUPABASE_URL}/rest/v1/rpc/upsert_prenotato`, {
    method: "POST",
    headers: { "Content-Type": "application/json", apikey: SUPABASE_KEY, Authorization: `Bearer ${token}` },
    body: JSON.stringify({ payload }),
  });
  if (!res1.ok) throw new Error("Errore import prenotato: " + JSON.stringify(await res1.json().catch(() => ({}))));

  const validiClienti = aggregatoClienti.filter(r => r.found && r.titolo_id);
  const payloadClienti = validiClienti.map(r => ({
    codice_cliente: r.codice_cliente, nome_cliente: r.nome_cliente, canale_id: canaleMap[r.canale] || null,
    titolo_id: r.titolo_id, quantita: r.qta,
    sconto_occasionale: r.sconto_occasionale ?? null, pagamento_occasionale: r.pagamento_occasionale ?? null,
    num_ordine_cliente: r.num_ordine_cliente ?? null,
  })).filter(r => r.canale_id !== null);

  for (let i = 0; i < payloadClienti.length; i += 500) {
    const batch = payloadClienti.slice(i, i + 500);
    await fetch(`${SUPABASE_URL}/rest/v1/rpc/upsert_prenotato_clienti`, {
      method: "POST",
      headers: { "Content-Type": "application/json", apikey: SUPABASE_KEY, Authorization: `Bearer ${token}` },
      body: JSON.stringify({ payload: batch }),
    });
  }

  const totQta = validi.reduce((s, r) => s + r.qta, 0);
  return { ok: payload.length, totQta };
}
