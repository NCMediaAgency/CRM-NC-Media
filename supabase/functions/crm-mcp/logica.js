/* Connettore del CRM per Claude: strumenti, ricerca dei clienti, riassunti dei dati e
   lettura dei numeri degli annunci. Funzioni pure, provate con node (tests/crm-mcp.test.mjs). */

export const nomeChiave = (s) => String(s || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/\s+/g, " ").trim();

/* Valori predefiniti delle winner ads: gli stessi del CRM (src/lib/winner.js) */
export const BENCH_BASE = { cpl: 9, app: 20, cliente: 100, minClienti: 5, minSpesa: 0, soloAttive: true, pubblico: "", raggruppa: true };

const NUM = { type: "number" };
export const STRUMENTI = [
  {
    name: "elenco_clienti",
    description: "Elenca i clienti del CRM NC Media (id, nome, cliente principale). Usalo per trovare il nome esatto prima degli altri strumenti.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false },
    annotations: { readOnlyHint: true },
  },
  {
    name: "dati_cliente",
    description: "Dati del CRM su un cliente: campagne Meta con spesa, lead e conversioni, ciclo di vendita degli ultimi mesi (lead, appuntamenti, chiusure, contratti, incassato), task aperti e benchmark delle winner ads impostati nel CRM.",
    inputSchema: { type: "object", properties: { cliente: { type: "string", description: "Nome o id del cliente" } }, required: ["cliente"], additionalProperties: false },
    annotations: { readOnlyHint: true },
  },
  {
    name: "leggi_benchmark",
    description: "Benchmark delle winner ads impostati nel CRM per un cliente: costo per lead, per appuntamento e per cliente chiuso di riferimento, minimo di clienti e di spesa, solo attive. Usali per giudicare gli annunci come li giudica il CRM.",
    inputSchema: { type: "object", properties: { cliente: { type: "string" } }, required: ["cliente"], additionalProperties: false },
    annotations: { readOnlyHint: true },
  },
  {
    name: "salva_dati_ads",
    description: "Salva nel CRM i numeri di ogni annuncio di un cliente (per esempio dal report di attribuzione di Hyros). Ogni invio è la fotografia di una data: rimandare la stessa data la sostituisce. Il CRM calcola da qui la classifica delle winner ads, separata per traffico freddo e retargeting, sommando le copie della stessa creatività (anche in pausa). Manda tutti gli annunci con spesa nel periodo, attivi e in pausa, ciascuno con pubblico e creativa.",
    inputSchema: {
      type: "object",
      properties: {
        cliente: { type: "string", description: "Nome o id del cliente nel CRM" },
        fonte: { type: "string", description: "Da dove vengono i numeri, es. Hyros" },
        data: { type: "string", description: "Data della fotografia AAAA-MM-GG (predefinita: oggi)" },
        periodo: { type: "string", description: "A che periodo si riferiscono i numeri, es. 'ultimi 30 giorni' o '2026-09-01 / 2026-10-16'" },
        ads: {
          type: "array",
          items: {
            type: "object",
            properties: {
              id: { type: "string", description: "Id dell'annuncio (quello di Meta se c'è)" },
              nome: { type: "string" },
              creativa: { type: "string", description: "Nome o codice della creatività, uguale per tutte le sue copie (es. '71' per '71 - tempo perso' in più campagne o ad set): il CRM somma le copie, anche quelle in pausa" },
              campagna: { type: "string" }, adset: { type: "string" },
              pubblico: { type: "string", enum: ["freddo", "retargeting", "altro"], description: "Traffico freddo o retargeting: il CRM fa una classifica per ciascuno" },
              attiva: { type: "boolean" },
              spesa: NUM, lead: NUM, appuntamenti: NUM, clienti: NUM, incassato: NUM,
              note: { type: "string" },
            },
            required: ["id", "nome", "spesa", "lead", "appuntamenti", "clienti"],
          },
        },
      },
      required: ["cliente", "ads"],
      additionalProperties: false,
    },
  },
  {
    name: "salva_analisi",
    description: "Salva nel CRM un'analisi scritta su un cliente (winner ads, creatività, andamento, consigli). Compare nella sezione Winner ads del cliente, con la data. Scrivi in italiano, in markdown semplice.",
    inputSchema: {
      type: "object",
      properties: {
        cliente: { type: "string" },
        tipo: { type: "string", enum: ["winner_ads", "creativita", "andamento", "altro"] },
        titolo: { type: "string" },
        testo: { type: "string" },
      },
      required: ["cliente", "titolo", "testo"],
      additionalProperties: false,
    },
  },
  {
    name: "leggi_analisi",
    description: "Le ultime analisi salvate nel CRM per un cliente, per confrontarle con quella nuova.",
    inputSchema: { type: "object", properties: { cliente: { type: "string" }, limite: { type: "number", description: "Quante (predefinito 3, massimo 10)" } }, required: ["cliente"], additionalProperties: false },
    annotations: { readOnlyHint: true },
  },
  {
    name: "crea_task",
    description: "Crea un task nel CRM per un cliente (per esempio: rifare una creatività, spegnere un annuncio). Il CRM lo inserisce entro un minuto.",
    inputSchema: {
      type: "object",
      properties: {
        cliente: { type: "string" }, titolo: { type: "string" }, note: { type: "string" },
        scadenza: { type: "string", description: "AAAA-MM-GG" },
        priorita: { type: "string", enum: ["urgente", "alta", "media", "bassa"] },
      },
      required: ["cliente", "titolo"],
      additionalProperties: false,
    },
  },
];

const SOCIETA = new Set(["llc", "srl", "srls", "spa", "sas", "snc", "ltd", "inc", "sa", "gmbh", "sl"]);
const sigla = (nome) => nomeChiave(nome).replace(/[^a-z0-9 ]/g, " ").split(" ").filter((w) => w && !SOCIETA.has(w)).map((w) => w[0]).join("");

/* Il cliente dal nome scritto da Claude: id, nome esatto, iniziali, poi nome contenuto */
export function trovaCliente(clienti, testo) {
  const l = (clienti || []).filter((z) => z.archiviato !== "si");
  const t = String(testo || "").trim();
  if (!t) return { errore: "Manca il nome del cliente." };
  const perId = l.find((z) => z.id === t);
  if (perId) return { cliente: perId };
  const k = nomeChiave(t);
  const esatti = l.filter((z) => nomeChiave(z.nome) === k);
  if (esatti.length === 1) return { cliente: esatti[0] };
  // Iniziali: "BDF" per Boutique Del Fitness LLC
  const perSigla = l.filter((z) => k.length >= 2 && !k.includes(" ") && sigla(z.nome) === k);
  if (perSigla.length === 1) return { cliente: perSigla[0] };
  const simili = l.filter((z) => nomeChiave(z.nome).includes(k) || k.includes(nomeChiave(z.nome)));
  if (simili.length === 1) return { cliente: simili[0] };
  if (simili.length > 1) {
    const principali = simili.filter((z) => !z.padreId);
    if (principali.length === 1) return { cliente: principali[0] };
    return { errore: `Più clienti corrispondono a "${t}": ${simili.slice(0, 10).map((z) => `${z.nome} (id ${z.id})`).join(", ")}. Usa l'id.` };
  }
  return { errore: `Nessun cliente "${t}" nel CRM. Usa elenco_clienti.` };
}

export const benchmark = (cliente) => ({ ...BENCH_BASE, ...((cliente && cliente.winner) || {}) });

const r2 = (x) => Math.round((Number(x) || 0) * 100) / 100;

/* Quello che il CRM sa di un cliente (e dei suoi sottoclienti), in poco spazio */
export function riassuntoCliente(data, cliente) {
  const clienti = data.clienti || [];
  const figli = clienti.filter((z) => z.padreId === cliente.id && z.archiviato !== "si");
  const ids = [cliente.id, ...figli.map((z) => z.id)];
  const nome = (id) => (clienti.find((z) => z.id === id) || {}).nome || "";
  const campagne = (data.campagne || []).filter((k) => ids.includes(k.clienteId) && k.archiviata !== "si").slice(0, 80).map((k) => ({
    nome: k.nome, cliente: nome(k.clienteId), stato: k.stato, obiettivo: k.obiettivo, inizio: k.inizio || "",
    spesa: r2(k.spesa), lead: k.lead || 0, conversioni: k.conversioni || 0, ricavi: r2(k.ricavi),
    cpl: k.lead ? r2(k.spesa / k.lead) : null, budgetGiorno: k.budgetGiorno || 0,
    ...(k.causaErrore ? { problema: k.causaErrore } : {}),
  }));
  const vendite = (data.vendite || []).filter((v) => ids.includes(v.clienteId))
    .sort((a, b) => (a.periodo < b.periodo ? 1 : -1)).slice(0, 12).map((v) => ({
      mese: v.periodo, cliente: nome(v.clienteId), spesa: r2(v.spesa), lead: v.lead || 0, chiamate: v.chiamate || 0,
      appuntamentiFissati: (v.discoveryFissate || 0) + (v.demoFissate || 0), appuntamentiSvolti: (v.discoverySvolte || 0) + (v.demoSvolte || 0),
      demoSvolte: v.demoSvolte || 0, chiusure: v.chiusure || 0, contratti: r2(v.valoreContratto), incassato: r2(v.incassato),
    }));
  const task = (data.task || []).filter((t) => ids.includes(t.clienteId) && t.stato !== "fatto").slice(0, 30)
    .map((t) => ({ titolo: t.titolo, scadenza: t.scadenza || "", priorita: t.priorita, stato: t.stato }));
  return {
    cliente: { id: cliente.id, nome: cliente.nome, settore: cliente.settore || "", principale: cliente.padreId ? nome(cliente.padreId) : "" },
    sottoclienti: figli.map((z) => z.nome),
    benchmarkWinnerAds: benchmark(cliente),
    campagne, cicloDiVendita: vendite, taskAperti: task,
  };
}

const numero = (x) => { const n = Number(String(x ?? "").replace(",", ".")); return Number.isFinite(n) && n > 0 ? n : 0; };
const testo = (x, max) => String(x ?? "").trim().slice(0, max);

/* Le righe di analisi_ads da quello che manda Claude */
export function righeAds(args, { ws, clienteId, oggi }) {
  const data = /^\d{4}-\d{2}-\d{2}$/.test(String(args.data || "")) ? args.data : oggi;
  const fonte = testo(args.fonte || "Hyros", 40) || "Hyros";
  const viste = new Set();
  const righe = [];
  for (const a of Array.isArray(args.ads) ? args.ads.slice(0, 1000) : []) {
    const id = testo(a && a.id, 120);
    if (!id || viste.has(id)) continue;
    viste.add(id);
    righe.push({
      workspace_id: ws, cliente_id: clienteId, fonte, data, ad_id: id,
      nome: testo(a.nome, 300) || id, creativa: testo(a.creativa, 120), campagna: testo(a.campagna, 300), adset: testo(a.adset, 300),
      pubblico: ["freddo", "retargeting", "altro"].includes(a.pubblico) ? a.pubblico : "",
      attiva: a.attiva !== false,
      spesa: numero(a.spesa), lead: numero(a.lead), appuntamenti: numero(a.appuntamenti), clienti: numero(a.clienti), incassato: numero(a.incassato),
      periodo: testo(args.periodo, 120), note: testo(a.note, 1000),
    });
  }
  return { data, fonte, righe };
}
