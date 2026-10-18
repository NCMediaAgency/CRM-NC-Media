// Prova del bot Telegram senza rete: logica (supabase/functions/telegram/logica.js) e
// applicazione dei comandi nel CRM (src/lib/comandi.js). Uso: node tests/telegram.test.mjs
import assert from "node:assert/strict";
import {
  oraRoma, sommaGiorni, prossimoLunedi, dataBreve, contestoCRM, comandoDaStrumento, riepilogo,
  taskDaRicordare, comandoDaPulsante, oraDelRiepilogo, STRUMENTI,
} from "../supabase/functions/telegram/logica.js";
import { applicaComandi } from "../src/lib/comandi.js";

// Date di Roma: 14 ottobre 2026 alle 7:30 UTC sono le 9:30 a Roma (ora legale)
const r = oraRoma(new Date("2026-10-14T07:30:00Z"));
assert.deepEqual(r, { data: "2026-10-14", hhmm: "09:30", dow: 3 });
assert.equal(oraRoma(new Date("2026-12-31T23:30:00Z")).data, "2027-01-01", "dopo mezzanotte a Roma è già domani");
assert.equal(sommaGiorni("2026-10-31", 1), "2026-11-01");
assert.equal(prossimoLunedi("2026-10-14"), "2026-10-19"); assert.equal(prossimoLunedi("2026-10-19"), "2026-10-26");
assert.equal(dataBreve("2026-10-23"), "ven 23 ott");

const data = {
  clienti: [{ id: "z1", nome: "Boutique Del Fitness LLC" }, { id: "z2", nome: "Oliva Consulting" }, { id: "z3", nome: "Vecchio", archiviato: "si" }],
  aree: [{ id: "ar1", nome: "Report" }],
  task: [
    { id: "t1", titolo: "Report settembre", clienteId: "z1", scadenza: "2026-10-10", priorita: "alta", stato: "dafare", ricorrenza: "30", inizio: "2026-10-01" },
    { id: "t2", titolo: "Call Oliva", clienteId: "z2", scadenza: "2026-10-14", ora: "15:00", priorita: "media", stato: "dafare", ricorrenza: "no" },
    { id: "t3", titolo: "Fatture", clienteId: "", scadenza: "2026-10-15", priorita: "urgente", stato: "dafare", ricorrenza: "no" },
    { id: "t4", titolo: "Già fatto", clienteId: "z1", scadenza: "2026-10-14", stato: "fatto" },
  ],
  campagne: [{ id: "k1", clienteId: "z2", nome: "Rottamazione", stato: "errore", causaErrore: "disattivato", spesa: 100, lead: 5 }],
  vendite: [{ clienteId: "z1", periodo: "2026-09", lead: 4758, demoSvolte: 4390, chiusure: 1612, valoreContratto: 1103534.32, incassato: 473879, spesa: 175506 }],
  annotazioni: [],
};

// Contesto per Claude: niente clienti archiviati, niente task fatti, i comandi in coda contano
const ctx = contestoCRM(data, "2026-10-14", [{ tipo: "crea_task", dati: { id: "tgx", titolo: "Nuovo", clienteId: "z1", scadenza: "2026-10-16" } }, { tipo: "modifica_task", dati: { taskId: "t3", stato: "fatto" } }]);
assert.deepEqual(ctx.clienti.map((z) => z.id), ["z1", "z2"]);
assert.ok(!ctx.taskAperti.some((t) => t.id === "t4" || t.id === "t3"), "fatti e chiusi in coda esclusi");
assert.ok(ctx.taskAperti.some((t) => t.id === "tgx" && t.inArrivo), "il task appena creato si può già correggere");
assert.equal(ctx.campagneFerme[0].causa, "disattivato");
assert.equal(ctx.cicloDiVendita["Boutique Del Fitness LLC"][0].chiusure, 1612);
assert.equal(ctx.giorno, "mercoledì");

// Strumenti: schema rigoroso, ogni campo richiesto
for (const t of STRUMENTI) assert.deepEqual([...t.input_schema.required].sort(), Object.keys(t.input_schema.properties).sort(), t.name);

// Da strumento a comando: validazioni e conferme
let n = 0; const nuovoId = () => String(++n);
let c = comandoDaStrumento("crea_task", { titolo: "Preparare il report di ottobre", clienteId: "z1", scadenza: "2026-10-23", ora: "9:30", priorita: "alta", note: null, areaId: "ar9" }, { data, oggi: "2026-10-14", nuovoId });
assert.equal(c.tipo, "crea_task"); assert.equal(c.dati.ora, "09:30"); assert.equal(c.dati.areaId, null, "area inesistente scartata");
assert.match(c.conferma, /Preparare il report di ottobre · Boutique Del Fitness LLC · ven 23 ott alle 09:30 · priorità alta/);
c = comandoDaStrumento("crea_task", { titolo: "X", clienteId: "inventato", scadenza: "giovedì", ora: null, priorita: "boh", note: null, areaId: null }, { data, oggi: "2026-10-14", nuovoId });
assert.equal(c.dati.clienteId, null); assert.equal(c.dati.scadenza, null); assert.equal(c.dati.priorita, "media");
c = comandoDaStrumento("crea_nota", { testo: "Spente le campagne Rottamazione", clienteId: "z2", data: null, tipo: "esterno" }, { data, oggi: "2026-10-14", nuovoId });
assert.equal(c.dati.data, "2026-10-14"); assert.match(c.conferma, /Nota nel diario · Oliva Consulting/);
c = comandoDaStrumento("modifica_task", { taskId: "t2", titolo: null, scadenza: "2026-10-24", ora: null, priorita: null, stato: null }, { data, oggi: "2026-10-14", nuovoId });
assert.match(c.conferma, /spostato a sab 24 ott · Call Oliva/);

// Riepilogo del mattino
const rp = riepilogo(data, "2026-10-14");
assert.match(rp.testo, /In ritardo \(1\)\n• Report settembre · Boutique Del Fitness LLC · alta · era sab 10 ott/);
assert.match(rp.testo, /Oggi \(1\)\n• Call Oliva · Oliva Consulting · 15:00/);
assert.match(rp.testo, /Domani, priorità alta \(1\)\n• Fatture · urgente/);
assert.match(rp.testo, /Campagne ferme\n• Oliva Consulting: 1 campagna/);
assert.deepEqual(rp.bottoni.map((r) => r[0].callback_data), ["f:t1", "f:t2"]);
assert.equal(riepilogo(data, "2026-10-14", [{ tipo: "modifica_task", dati: { taskId: "t1", stato: "fatto" } }]).bottoni.length, 1, "un task chiuso dal pulsante sparisce dal riepilogo");

// Promemoria dei task con orario
assert.deepEqual(taskDaRicordare(data, "2026-10-14", "14:40", 30).map((t) => t.id), ["t2"]);
assert.deepEqual(taskDaRicordare(data, "2026-10-14", "14:20", 30), [], "troppo presto");
assert.deepEqual(taskDaRicordare(data, "2026-10-14", "15:05", 30), [], "ormai passato");

// Ora del riepilogo e weekend
assert.equal(oraDelRiepilogo({ hhmm: "09:12", dow: 3 }, { ora_riepilogo: "09:10", weekend: false }), true);
assert.equal(oraDelRiepilogo({ hhmm: "09:05", dow: 3 }, { ora_riepilogo: "09:10" }), false);
assert.equal(oraDelRiepilogo({ hhmm: "10:00", dow: 6 }, { ora_riepilogo: "09:10", weekend: false }), false);
assert.equal(oraDelRiepilogo({ hhmm: "10:00", dow: 6 }, { ora_riepilogo: "09:10", weekend: true }), true);

// Pulsanti
assert.deepEqual(comandoDaPulsante("f:t2", "2026-10-14").dati, { taskId: "t2", stato: "fatto" });
assert.equal(comandoDaPulsante("d:t2", "2026-10-14").dati.scadenza, "2026-10-15");
assert.equal(comandoDaPulsante("l:t2", "2026-10-14").dati.scadenza, "2026-10-19");
assert.equal(comandoDaPulsante("ok:", "2026-10-14"), null);

// Applicazione nel CRM
let k = 0; const uid = () => "u" + ++k;
const ap = applicaComandi(data, [
  { id: "c1", tipo: "crea_task", dati: { id: "tg1", titolo: "Report ottobre", clienteId: "z1", scadenza: "2026-10-23", ora: "09:30", priorita: "alta" } },
  { id: "c2", tipo: "crea_nota", dati: { id: "tg2", testo: "Spente le Rottamazione", clienteId: "z2", data: "2026-10-14", tipo: "esterno" } },
  { id: "c3", tipo: "modifica_task", dati: { taskId: "t1", stato: "fatto" } },
  { id: "c4", tipo: "modifica_task", dati: { taskId: "t2", scadenza: "2026-10-15" } },
  { id: "c5", tipo: "modifica_task", dati: { taskId: "non-c-e", stato: "fatto" } },
  { id: "c6", tipo: "crea_task", dati: { id: "tg1", titolo: "doppione" } },
], { uid, oggi: "2026-10-14", giorniRicorrenza: (v) => (v === "30" ? 30 : 0) });
const T = (id) => ap.data.task.find((t) => t.id === id);
assert.equal(T("tg1").ora, "09:30"); assert.equal(T("tg1").stato, "dafare"); assert.equal(T("tg1").origine, "telegram");
assert.equal(ap.data.annotazioni[0].testo, "Spente le Rottamazione"); assert.equal(ap.data.annotazioni[0].clienteId, "z2");
assert.equal(T("t1").stato, "fatto");
assert.ok(ap.data.task.some((t) => t.titolo === "Report settembre" && t.stato === "dafare" && t.scadenza === "2026-11-09"), "task ricorrente: nasce il successivo");
assert.equal(T("t2").scadenza, "2026-10-15");
assert.deepEqual(ap.esiti.map((e) => e.esito), ["task creato", "nota creata", "task aggiornato", "task aggiornato", "task non trovato", "già presente"]);
assert.equal(ap.data.task.filter((t) => t.id === "tg1").length, 1, "niente doppioni");

// Fuso dello spazio di lavoro: New York sei ore indietro; un fuso sbagliato vale come Roma
assert.deepEqual(oraRoma(new Date("2026-10-18T14:30:00Z"), "America/New_York").hhmm, "10:30");
assert.deepEqual(oraRoma(new Date("2026-10-18T02:30:00Z"), "America/Los_Angeles").data, "2026-10-17");
assert.deepEqual(oraRoma(new Date("2026-10-18T14:30:00Z"), "Fuso/Inesistente").hhmm, "16:30");
assert.deepEqual(oraRoma(new Date("2026-10-18T14:30:00Z")).hhmm, "16:30");
console.log("telegram: tutte le prove passate ✔");
