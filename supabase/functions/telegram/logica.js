/* Logica del bot Telegram, senza dipendenze: date di Roma, contesto per Claude,
   strumenti, riepilogo del mattino e promemoria. Funzioni pure, provate con node
   (tests/telegram.test.mjs); index.ts le usa per parlare con Telegram e con Claude. */

export const MESI = ["gen", "feb", "mar", "apr", "mag", "giu", "lug", "ago", "set", "ott", "nov", "dic"];
const GIORNI = ["domenica", "lunedì", "martedì", "mercoledì", "giovedì", "venerdì", "sabato"];
const PRIORITA = ["urgente", "alta", "media", "bassa"];
const TIPI_NOTA = ["budget", "creativita", "pubblico", "offerta", "tracciamento", "esterno", "altro"];

/* Data e ora nel fuso dello spazio di lavoro (predefinito Roma): { data: "AAAA-MM-GG", hhmm: "HH:MM", dow: 0-6 } */
export function oraRoma(adesso = new Date(), fuso = "Europe/Rome") {
  const parti = (tz) => new Intl.DateTimeFormat("en-CA", {
    timeZone: tz, year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hourCycle: "h23", weekday: "short",
  }).formatToParts(adesso);
  let l;
  try { l = parti(fuso || "Europe/Rome"); } catch { l = parti("Europe/Rome"); }
  const p = Object.fromEntries(l.map((x) => [x.type, x.value]));
  const dow = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"].indexOf(p.weekday);
  return { data: `${p.year}-${p.month}-${p.day}`, hhmm: `${p.hour}:${p.minute}`, dow };
}
export const sommaGiorni = (d, n) => { const x = new Date(d + "T12:00:00Z"); x.setUTCDate(x.getUTCDate() + n); return x.toISOString().slice(0, 10); };
export const giornoSettimana = (d) => GIORNI[new Date(d + "T12:00:00Z").getUTCDay()];
export const dataBreve = (d) => (d ? `${giornoSettimana(d).slice(0, 3)} ${Number(d.slice(8, 10))} ${MESI[Number(d.slice(5, 7)) - 1]}` : "senza data");
/* Prossimo lunedì (se oggi è lunedì, quello dopo) */
export const prossimoLunedi = (d) => { const dow = new Date(d + "T12:00:00Z").getUTCDay(); return sommaGiorni(d, ((8 - dow) % 7) || 7); };
const minuti = (hhmm) => { const m = /^(\d{1,2}):(\d{2})$/.exec(hhmm || ""); return m ? Number(m[1]) * 60 + Number(m[2]) : null; };

const aperto = (t) => t && t.stato !== "fatto";
const nomeCliDi = (data) => { const m = new Map((data.clienti || []).map((z) => [z.id, z.nome])); return (id) => m.get(id) || ""; };

/* Quello che Claude deve sapere del CRM, in poco spazio: clienti, aree, task aperti, campagne
   ferme, ultimi numeri commerciali. I comandi ancora in coda contano come già fatti. */
export function contestoCRM(data, oggi, inCoda = []) {
  const nome = nomeCliDi(data);
  const clienti = (data.clienti || []).filter((z) => z.archiviato !== "si" && z.stato !== "chiuso")
    .map((z) => ({ id: z.id, nome: z.nome, ...(z.padreId ? { sottoclienteDi: nome(z.padreId) } : {}) }));
  let task = (data.task || []).filter(aperto).map((t) => ({
    id: t.id, titolo: t.titolo, cliente: nome(t.clienteId) || null, scadenza: t.scadenza || null, ora: t.ora || null, priorita: t.priorita || "media", stato: t.stato,
  }));
  // I task creati o cambiati da Telegram e non ancora applicati dal CRM
  inCoda.forEach((c) => {
    if (c.tipo === "crea_task") task.push({ id: c.dati.id, titolo: c.dati.titolo, cliente: nome(c.dati.clienteId) || null, scadenza: c.dati.scadenza || null, ora: c.dati.ora || null, priorita: c.dati.priorita || "media", stato: "dafare", inArrivo: true });
    if (c.tipo === "modifica_task") task = task.map((t) => (t.id === c.dati.taskId ? { ...t, ...Object.fromEntries(Object.entries(c.dati).filter(([k, v]) => k !== "taskId" && v !== null)) } : t)).filter(aperto);
  });
  task = task.sort((a, b) => (a.scadenza || "9999").localeCompare(b.scadenza || "9999")).slice(0, 120);
  const camp = (data.campagne || []).filter((k) => k.archiviata !== "si");
  const ferme = camp.filter((k) => k.stato === "errore").map((k) => ({ cliente: nome(k.clienteId), campagna: k.nome, causa: k.causaErrore || "pagamento" }));
  const perCli = {};
  camp.filter((k) => ["attiva", "test", "errore"].includes(k.stato)).forEach((k) => {
    const g = perCli[k.clienteId] || (perCli[k.clienteId] = { cliente: nome(k.clienteId), campagneAccese: 0, spesa: 0, lead: 0, conversioni: 0 });
    g.campagneAccese++; g.spesa += k.spesa || 0; g.lead += k.lead || 0; g.conversioni += k.conversioni || 0;
  });
  const vendite = {};
  (data.vendite || []).slice().sort((a, b) => (a.periodo < b.periodo ? 1 : -1)).forEach((v) => {
    const l = vendite[nome(v.clienteId)] || (vendite[nome(v.clienteId)] = []);
    if (l.length < 3) l.push({ mese: v.periodo, lead: v.lead || 0, demoSvolte: v.demoSvolte || 0, chiusure: v.chiusure || 0, contratti: Math.round(v.valoreContratto || 0), incassato: Math.round(v.incassato || 0), spesa: Math.round(v.spesa || 0) });
  });
  return {
    oggi, giorno: giornoSettimana(oggi),
    clienti, aree: (data.aree || []).map((a) => ({ id: a.id, nome: a.nome })),
    taskAperti: task,
    campagneAccesePerCliente: Object.values(perCli).map((g) => ({ ...g, spesa: Math.round(g.spesa) })),
    campagneFerme: ferme,
    cicloDiVendita: vendite,
  };
}

/* Istruzioni per Claude: stabili (vanno in cache); il contesto del giorno arriva nel messaggio */
export const ISTRUZIONI = `Sei l'assistente del CRM di un'agenzia di media buying, e parli con il titolare su Telegram, in italiano.
Puoi fare tre cose sul CRM, solo con gli strumenti: creare un task (crea_task), scrivere una nota nel diario (crea_nota), modificare o chiudere un task esistente (modifica_task).
Rispondi anche alle domande sui dati usando solo il contesto CRM che ricevi: se un dato non c'è, dillo, non inventare.

Regole:
- Le date vanno sempre nel formato AAAA-MM-GG, calcolate a partire da "oggi" del contesto ("giovedì" = il prossimo giovedì, "fra due settimane", "fine mese" = ultimo giorno del mese).
- Il cliente si indica con il suo id preso dall'elenco clienti del contesto. Accetta abbreviazioni e iniziali (BDF = Boutique Del Fitness). Se più clienti sono plausibili, non creare nulla e chiedi quale.
- Un task senza cliente è un task dell'agenzia: clienteId null.
- Per spostare, chiudere o correggere un task usa modifica_task con l'id del task preso da taskAperti (anche quelli inArrivo, appena creati). Per chiuderlo: stato "fatto".
- Più richieste nello stesso messaggio: usa più strumenti nella stessa risposta.
- Se il messaggio non chiede nulla di concreto, rispondi e basta, senza strumenti.
- Risposte brevi, testo semplice senza markdown. Le conferme dei task e delle note le aggiunge il CRM: non ripeterle.`;

/* Strumenti: strict, tutti i campi presenti (null quando non servono) */
const nullabile = (t, extra = {}) => ({ type: [t, "null"], ...extra });
export const STRUMENTI = [
  {
    name: "crea_task", strict: true,
    description: "Crea un task nel CRM. Usalo quando il titolare chiede di ricordare, fare o programmare qualcosa.",
    input_schema: {
      type: "object", additionalProperties: false,
      required: ["titolo", "clienteId", "scadenza", "ora", "priorita", "note", "areaId"],
      properties: {
        titolo: { type: "string", description: "Cosa fare, breve e all'infinito" },
        clienteId: nullabile("string", { description: "id del cliente dal contesto, null se è dell'agenzia" }),
        scadenza: nullabile("string", { description: "AAAA-MM-GG" }),
        ora: nullabile("string", { description: "HH:MM se il titolare indica un orario" }),
        priorita: { type: "string", enum: PRIORITA },
        note: nullabile("string"),
        areaId: nullabile("string", { description: "id dell'area di lavoro, se nominata" }),
      },
    },
  },
  {
    name: "crea_nota", strict: true,
    description: "Scrive una nota nel diario del CRM: una decisione presa, un cambiamento fatto, un fatto da ricordare.",
    input_schema: {
      type: "object", additionalProperties: false,
      required: ["testo", "clienteId", "data", "tipo"],
      properties: {
        testo: { type: "string" },
        clienteId: nullabile("string"),
        data: nullabile("string", { description: "AAAA-MM-GG, se diversa da oggi" }),
        tipo: { type: "string", enum: TIPI_NOTA },
      },
    },
  },
  {
    name: "modifica_task", strict: true,
    description: "Sposta, chiude o corregge un task esistente. Metti null nei campi che non cambiano.",
    input_schema: {
      type: "object", additionalProperties: false,
      required: ["taskId", "titolo", "scadenza", "ora", "priorita", "stato"],
      properties: {
        taskId: { type: "string" },
        titolo: nullabile("string"),
        scadenza: nullabile("string", { description: "AAAA-MM-GG" }),
        ora: nullabile("string", { description: "HH:MM" }),
        priorita: nullabile("string", { enum: [...PRIORITA, null] }),
        stato: nullabile("string", { enum: ["dafare", "incorso", "attesa", "fatto", null] }),
      },
    },
  },
];

const dataValida = (s) => (typeof s === "string" && /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : null);
const oraValida = (s) => (typeof s === "string" && minuti(s) !== null ? s.padStart(5, "0") : null);

/* Da una chiamata di strumento al comando in coda per il CRM, con la conferma da mandare */
export function comandoDaStrumento(nome, input, { data, oggi, nuovoId }) {
  const nomeCli = nomeCliDi(data);
  const cliOk = (id) => (id && (data.clienti || []).some((z) => z.id === id) ? id : null);
  if (nome === "crea_task") {
    const dati = {
      id: "tg" + nuovoId(), titolo: String(input.titolo || "").trim().slice(0, 200) || "Task da Telegram",
      clienteId: cliOk(input.clienteId), scadenza: dataValida(input.scadenza), ora: oraValida(input.ora),
      priorita: PRIORITA.includes(input.priorita) ? input.priorita : "media", note: input.note || "",
      areaId: (data.aree || []).some((a) => a.id === input.areaId) ? input.areaId : null,
    };
    return { tipo: "crea_task", dati,
      conferma: `✅ Task creato · ${dati.titolo}${dati.clienteId ? ` · ${nomeCli(dati.clienteId)}` : ""} · ${dataBreve(dati.scadenza)}${dati.ora ? ` alle ${dati.ora}` : ""}${dati.priorita !== "media" ? ` · priorità ${dati.priorita}` : ""}` };
  }
  if (nome === "crea_nota") {
    const dati = { id: "tg" + nuovoId(), testo: String(input.testo || "").trim().slice(0, 2000), clienteId: cliOk(input.clienteId),
      data: dataValida(input.data) || oggi, tipo: TIPI_NOTA.includes(input.tipo) ? input.tipo : "altro" };
    if (!dati.testo) return null;
    return { tipo: "crea_nota", dati, conferma: `📝 Nota nel diario · ${dati.clienteId ? `${nomeCli(dati.clienteId)} · ` : ""}${dataBreve(dati.data)} · "${dati.testo.slice(0, 120)}${dati.testo.length > 120 ? "…" : ""}"` };
  }
  if (nome === "modifica_task") {
    const dati = { taskId: String(input.taskId || ""), titolo: input.titolo || null, scadenza: dataValida(input.scadenza), ora: oraValida(input.ora),
      priorita: PRIORITA.includes(input.priorita) ? input.priorita : null, stato: ["dafare", "incorso", "attesa", "fatto"].includes(input.stato) ? input.stato : null };
    if (!dati.taskId) return null;
    const t = (data.task || []).find((x) => x.id === dati.taskId);
    const cosa = dati.stato === "fatto" ? "chiuso" : dati.scadenza ? `spostato a ${dataBreve(dati.scadenza)}${dati.ora ? ` alle ${dati.ora}` : ""}` : "aggiornato";
    return { tipo: "modifica_task", dati, conferma: `${dati.stato === "fatto" ? "✅" : "🔁"} Task ${cosa}${t ? ` · ${t.titolo}` : ""}` };
  }
  return null;
}

/* Riepilogo del mattino: in ritardo, oggi, domani urgenti/alti, campagne ferme */
export function riepilogo(data, oggi, inCoda = []) {
  const nome = nomeCliDi(data);
  const chiusi = new Set(inCoda.filter((c) => c.tipo === "modifica_task" && c.dati.stato === "fatto").map((c) => c.dati.taskId));
  const spostati = new Map(inCoda.filter((c) => c.tipo === "modifica_task" && c.dati.scadenza).map((c) => [c.dati.taskId, c.dati.scadenza]));
  const tutti = (data.task || []).filter((t) => aperto(t) && !chiusi.has(t.id)).map((t) => ({ ...t, scadenza: spostati.get(t.id) || t.scadenza }));
  const domani = sommaGiorni(oggi, 1);
  const ritardo = tutti.filter((t) => t.scadenza && t.scadenza < oggi);
  const diOggi = tutti.filter((t) => t.scadenza === oggi);
  const diDomani = tutti.filter((t) => t.scadenza === domani && ["urgente", "alta"].includes(t.priorita));
  const ferme = (data.campagne || []).filter((k) => k.archiviata !== "si" && k.stato === "errore");
  const riga = (t) => `• ${t.titolo}${t.clienteId ? ` · ${nome(t.clienteId)}` : ""}${t.ora ? ` · ${t.ora}` : ""}${["urgente", "alta"].includes(t.priorita) ? ` · ${t.priorita}` : ""}`;
  const parti = [`☀️ Buongiorno, ecco ${giornoSettimana(oggi)} ${dataBreve(oggi).split(" ").slice(1).join(" ")}`];
  if (ritardo.length) parti.push(`\n⏰ In ritardo (${ritardo.length})\n${ritardo.slice(0, 10).map((t) => `${riga(t)} · era ${dataBreve(t.scadenza)}`).join("\n")}`);
  if (diOggi.length) parti.push(`\n📌 Oggi (${diOggi.length})\n${diOggi.slice(0, 12).map(riga).join("\n")}`);
  if (diDomani.length) parti.push(`\n🔜 Domani, priorità alta (${diDomani.length})\n${diDomani.slice(0, 8).map(riga).join("\n")}`);
  if (ferme.length) {
    const perCli = {};
    ferme.forEach((k) => { const c = nome(k.clienteId); perCli[c] = (perCli[c] || 0) + 1; });
    parti.push(`\n🛑 Campagne ferme\n${Object.entries(perCli).map(([c, n]) => `• ${c}: ${n} ${n === 1 ? "campagna" : "campagne"}`).join("\n")}`);
  }
  if (parti.length === 1) parti.push("\nNiente in scadenza oggi e nessuna campagna ferma. 👌");
  // Pulsanti per i task da chiudere o spostare oggi (in ritardo e di oggi)
  const daGestire = [...ritardo, ...diOggi].slice(0, 8);
  const bottoni = daGestire.map((t) => [
    { text: `✅ ${t.titolo.slice(0, 28)}`, callback_data: `f:${t.id}` },
    { text: "⏰ Domani", callback_data: `d:${t.id}` },
  ]);
  return { testo: parti.join("\n"), bottoni, nTask: daGestire.length };
}

/* Task con orario da ricordare adesso: oggi, fra (ora - preavviso) e l'ora stessa */
export function taskDaRicordare(data, oggi, hhmm, preavviso = 30) {
  const ora = minuti(hhmm);
  return (data.task || []).filter((t) => aperto(t) && t.scadenza === oggi && minuti(t.ora) !== null)
    .filter((t) => { const m = minuti(t.ora); return ora >= m - preavviso && ora <= m; });
}

export const bottoniTask = (id) => [[
  { text: "✅ Fatto", callback_data: `f:${id}` },
  { text: "⏰ Domani", callback_data: `d:${id}` },
  { text: "📅 Lunedì", callback_data: `l:${id}` },
]];

/* Pulsante premuto → comando: f = fatto, d = domani, l = lunedì */
export function comandoDaPulsante(dato, oggi) {
  const [az, taskId] = String(dato || "").split(":");
  if (!taskId) return null;
  if (az === "f") return { tipo: "modifica_task", dati: { taskId, stato: "fatto" }, testo: "✅ Fatto" };
  if (az === "d") { const d = sommaGiorni(oggi, 1); return { tipo: "modifica_task", dati: { taskId, scadenza: d }, testo: `⏰ Spostato a ${dataBreve(d)}` }; }
  if (az === "l") { const d = prossimoLunedi(oggi); return { tipo: "modifica_task", dati: { taskId, scadenza: d }, testo: `📅 Spostato a ${dataBreve(d)}` }; }
  return null;
}

/* È ora del riepilogo? Ora passata, giorno giusto (weekend solo se scelto) */
export const oraDelRiepilogo = (adesso, chat) => {
  if (!chat.weekend && (adesso.dow === 0 || adesso.dow === 6)) return false;
  const m = minuti(chat.ora_riepilogo || "09:10");
  return m !== null && minuti(adesso.hhmm) >= m;
};
