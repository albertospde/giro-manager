// Aggiornamento automatico del mattino del prenotato (Fine Giro) da RPN.
// Eseguito da GitHub Actions (.github/workflows/prenotato-notturno.yml) alle 5:00 ora italiana.
//
// Per ogni giro / cedola extra NON chiuso (giro_cedola_stato) non ancora aggiornato oggi:
//   1. chiede alla edge function giro-prenotato-sync l'URL dell'export RPN e un token temporaneo
//   2. scarica il file "Pianifica Visite" (xlsx, ~2 minuti per generarlo lato RPN)
//   3. lo legge con parseEaggrega di src/rpnPrenotatoSync.js: stesso identico calcolo del
//      bottone "Aggiorna da RPN" di Fine Giro, limitato ai titoli di quel giro/cedola
//   4. salva prenotato e prenotato clienti tramite la edge function e scrive l'esito in prenotato_sync_log
//
// Variabili d'ambiente: GIRO_PRENOTATO_CRON_KEY (segreto GitHub = app_secrets.giro_prenotato_cron),
// NODE_EXTRA_CA_CERTS=scripts/sectigo-rpn.pem (catena certificati incompleta lato RPN),
// FORZA=1 per ignorare il controllo dell'orario (avvio manuale).
// Nei log stampa solo conteggi: il repo può essere pubblico.

import * as XLSX from "xlsx";

globalThis.window = { XLSX };
const { parseEaggrega } = await import("../src/rpnPrenotatoSync.js");

const FN_URL = "https://tdflwenlylhctxssatax.supabase.co/functions/v1/giro-prenotato-sync";
const ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InRkZmx3ZW5seWxoY3R4c3NhdGF4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYzMzgyNzYsImV4cCI6MjA5MTkxNDI3Nn0.l35qEL7LOvyYuI1McQlVqj4vbyTqmlevcmqWbTGYi2Q";
const CRON_KEY = process.env.GIRO_PRENOTATO_CRON_KEY;
const ORA_INIZIO = 5; // ora italiana

async function job(azione, dati = {}) {
  const res = await fetch(FN_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${ANON_KEY}`, "x-cron-key": CRON_KEY },
    body: JSON.stringify({ azione, ...dati }),
  });
  const j = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(`${azione}: ${j.error || res.status}`);
  return j;
}

function oraItaliana() {
  return Number(new Intl.DateTimeFormat("it-IT", { timeZone: "Europe/Rome", hour: "2-digit", hourCycle: "h23" }).format(new Date()));
}

async function aggiorna(target, canaleId) {
  const { tipo, chiave } = target;
  const t0 = Date.now();
  const { id: logId } = await job("log", { tipo, chiave, origine: "github", stato: "in_corso" });
  try {
    const { url, token } = await job("export", { tipo, chiave });
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) throw new Error(`Export RPN ${res.status}`);
    const buf = new Uint8Array(await res.arrayBuffer());
    const { titoli } = await job("titoli", { tipo, chiave });

    const { aggregato, aggregatoClienti, totaleAggregato, totaleFound, righe } = parseEaggrega(buf, titoli);

    // stesso payload di importAggregato (rpnPrenotatoSync.js)
    const prenotato = aggregato.filter(r => r.found && r.titolo_id && canaleId[r.canale])
      .map(r => ({ titolo_id: r.titolo_id, canale_id: canaleId[r.canale], quantita: r.qta }));
    const clienti = aggregatoClienti.filter(r => r.found && r.titolo_id && canaleId[r.canale])
      .map(r => ({
        codice_cliente: r.codice_cliente, nome_cliente: r.nome_cliente, canale_id: canaleId[r.canale],
        titolo_id: r.titolo_id, quantita: r.qta,
        sconto_occasionale: r.sconto_occasionale ?? null, pagamento_occasionale: r.pagamento_occasionale ?? null,
        num_ordine_cliente: r.num_ordine_cliente ?? null,
      }));

    for (let i = 0; i < prenotato.length; i += 2000) await job("salva", { tipo, chiave, prenotato: prenotato.slice(i, i + 2000) });
    for (let i = 0; i < clienti.length; i += 2000) await job("salva", { tipo, chiave, clienti: clienti.slice(i, i + 2000) });

    const esito = {
      righe_file: righe.length, qta_file: totaleAggregato, qta_abbinata: totaleFound,
      prenotato_righe: prenotato.length, clienti_righe: clienti.length,
      ean_non_trovati: new Set(aggregato.filter(r => !r.found).map(r => r.ean)).size, formato_file: `xlsx ${Math.round(buf.length / 1024)}KB`,
    };
    await job("log", { id: logId, stato: "ok", finito_at: new Date().toISOString(), durata_ms: Date.now() - t0, ...esito });
    console.log(`OK   ${tipo} ${chiave}: ${esito.qta_abbinata}/${esito.qta_file} copie abbinate, ${esito.prenotato_righe} righe prenotato, ${esito.clienti_righe} righe clienti (${Math.round((Date.now() - t0) / 1000)}s)`);
    return true;
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    await job("log", { id: logId, stato: "errore", finito_at: new Date().toISOString(), durata_ms: Date.now() - t0, messaggio: msg.slice(0, 2000) }).catch(() => {});
    console.log(`ERR  ${tipo} ${chiave}: ${msg.slice(0, 300)}`);
    return false;
  }
}

if (!CRON_KEY) { console.error("Manca il segreto GIRO_PRENOTATO_CRON_KEY"); process.exit(1); }
const ora = oraItaliana();
if (!process.env.FORZA && ora < ORA_INIZIO) {
  console.log(`In Italia sono le ${ora}: si parte dalle ${ORA_INIZIO}:00, esco.`);
  process.exit(0);
}

const { targets, canali } = await job("piano");
const canaleId = Object.fromEntries(canali.map(c => [c.codice, c.id]));
console.log(`Da aggiornare: ${targets.length} (${targets.map(t => t.chiave).join(", ") || "nessuno"})`);

let errori = 0;
for (const t of targets) if (!(await aggiorna(t, canaleId))) errori++;
console.log(`Fine: ${targets.length - errori} aggiornati, ${errori} con errore`);
process.exit(errori ? 1 : 0);
