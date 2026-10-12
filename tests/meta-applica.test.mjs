// Prova di src/lib/meta.js: come i dati di Meta entrano nel CRM.
// Uso: node tests/meta-applica.test.mjs
import assert from "node:assert/strict";
import { applicaMeta, suggerisciAccount, statoCampagnaDaMeta, obiettivoDaMeta } from "../src/lib/meta.js";

let n = 0;
const opz = { oggi: "2026-10-12", uid: () => "id" + ++n, bmkDi: (o) => ({ cplMax: o === "lead" ? 35 : 0 }), valutaAgenzia: "EUR" };

const data = {
  clienti: [{ id: "z1", nome: "Boutique del Fitness", valuta: "" }, { id: "z2", nome: "Cantine", valuta: "" }],
  campagne: [
    // scritta a mano prima del collegamento: va agganciata per nome
    { id: "k1", clienteId: "z1", nome: "Lead fitness", stato: "test", spesa: 1, esitoManuale: "si", conversioni: 4, ricavi: 0, storicoBudget: [], budgetGiorno: 0, bmk: {} },
    // altro cliente con lo stesso nome: non va toccata
    { id: "k9", clienteId: "z2", nome: "Vendite corsi", stato: "attiva", spesa: 5, storicoBudget: [], bmk: {} },
  ],
  creativi: [{ id: "d1", clienteId: "z1", campagnaId: "k1", tipo: "adset", nome: "Donne 25-45", stato: "attiva", stadio: "2", note: "nota mia", spesa: 0 }],
  periodi: [
    // finestra manuale che si sovrappone a quelle di Meta: sparisce
    { id: "p1", clienteId: "z1", campagnaId: "k1", tipo: "attiva", inizio: "2026-09-24", fine: "2026-09-30", spesa: 99, valoreContratto: 0, incassato: 0, note: "" },
    // stessa data di inizio di una finestra Meta: resta, con i numeri di Meta e la nota mia
    { id: "p2", clienteId: "z1", campagnaId: "k1", tipo: "test", inizio: "2026-09-22", fine: "2026-09-24", spesa: 5, valoreContratto: 1000, incassato: 0, note: "tengo" },
  ],
};
const met = (o) => ({ spesa: 0, impression: 0, click: 0, click_link: 0, lead: 0, conversioni: 0, ricavi: 0, frequenza: 0, ...o });
const meta = {
  collegamenti: [{ account_id: "555", cliente_id: "z1" }],
  account: [{ account_id: "555", nome: "BDF 1", valuta: "USD", stato: 1 }],
  oggetti: [
    met({ livello: "campaign", id: "c1", account_id: "555", nome: "Lead fitness", stato: "ACTIVE", obiettivo: "OUTCOME_LEADS", budget_giorno: 50, inizio: "2026-09-22", fine: null, spesa: 300, lead: 25, conversioni: 0, click_link: 30 }),
    met({ livello: "campaign", id: "c2", account_id: "555", nome: "Vendite corsi", stato: "PAUSED", obiettivo: "OUTCOME_SALES", budget_giorno: 0, inizio: "2026-09-01", fine: null, spesa: 500, conversioni: 7, ricavi: 980 }),
    met({ livello: "adset", id: "s1", account_id: "555", campagna_id: "c1", nome: "Donne 25-45", stato: "ACTIVE", spesa: 300, lead: 25 }),
    met({ livello: "ad", id: "a1", account_id: "555", campagna_id: "c1", adset_id: "s1", nome: "Video 1", stato: "PAUSED", spesa: 200, lead: 18 }),
    // account non collegato: ignorato
    met({ livello: "campaign", id: "x1", account_id: "999", nome: "Estranea", stato: "ACTIVE", spesa: 1 }),
  ],
  finestre: [
    met({ campagna_id: "c1", inizio: "2026-09-22", fine: "2026-09-24", tipo: "test", spesa: 40, lead: 3 }),
    met({ campagna_id: "c1", inizio: "2026-09-25", fine: "2026-10-01", tipo: "matura", spesa: 100, lead: 8 }),
    met({ campagna_id: "c1", inizio: "2026-10-02", fine: "2026-10-08", tipo: "attiva", spesa: 120, lead: 10 }),
    met({ campagna_id: "c1", inizio: "2026-10-09", fine: "2026-10-11", tipo: "attiva", spesa: 40, lead: 4 }),
    met({ campagna_id: "c2", inizio: "2026-09-01", fine: "2026-09-03", tipo: "test", spesa: 300 }),
    met({ campagna_id: "c2", inizio: "2026-09-04", fine: "2026-09-10", tipo: "matura", spesa: 200 }),
    met({ campagna_id: "c2", inizio: "2026-09-11", fine: "2026-09-17", tipo: "attiva", spesa: 0 }),
  ],
};

const r = applicaMeta(data, meta, opz);
assert.ok(r.cambiato);
const d = r.data;
const k1 = d.campagne.find((k) => k.id === "k1");
assert.equal(k1.metaId, "c1", "agganciata per nome");
assert.equal(k1.stato, "test", "lo stato 'In test' scelto da te resta");
assert.equal(k1.spesa, 300); assert.equal(k1.lead, 25); assert.equal(k1.clickLink, 30);
assert.equal(k1.conversioni, 4, "esito scritto a mano: Meta non lo sovrascrive");
assert.equal(k1.budgetGiorno, 50); assert.equal(k1.storicoBudget.length, 1); assert.equal(k1.storicoBudget[0].data, "2026-09-22");
assert.equal(k1.fine, "", "attiva: nessuna data di fine");
const k2 = d.campagne.find((k) => k.metaId === "c2");
assert.ok(k2 && k2.id !== "k9", "nuova campagna, non quella omonima di un altro cliente");
assert.equal(k2.clienteId, "z1"); assert.equal(k2.obiettivo, "vendite"); assert.equal(k2.piattaforma, "meta");
assert.equal(k2.stato, "pausa"); assert.equal(k2.conversioni, 7); assert.equal(k2.ricavi, 980);
assert.equal(k2.fine, "2026-09-10", "in pausa: chiusa sull'ultima finestra con spesa");
assert.equal(d.campagne.find((k) => k.id === "k9").spesa, 5, "l'altro cliente non si tocca");
assert.ok(!d.campagne.some((k) => k.metaId === "x1"), "account non collegato ignorato");
const s1 = d.creativi.find((c) => c.id === "d1");
assert.equal(s1.metaId, "s1"); assert.equal(s1.stadio, "2"); assert.equal(s1.note, "nota mia"); assert.equal(s1.spesa, 300);
const a1 = d.creativi.find((c) => c.metaId === "a1");
assert.equal(a1.tipo, "annuncio"); assert.equal(a1.adsetId, "d1", "annuncio agganciato al suo ad set"); assert.equal(a1.stato, "pausa");
const pk1 = d.periodi.filter((p) => p.campagnaId === "k1").sort((a, b) => (a.inizio < b.inizio ? -1 : 1));
assert.equal(pk1.length, 4, "la finestra manuale sovrapposta è sparita");
assert.ok(!pk1.some((p) => p.id === "p1"));
const p2 = pk1.find((p) => p.id === "p2");
assert.equal(p2.spesa, 40); assert.equal(p2.note, "tengo"); assert.equal(p2.valoreContratto, 1000); assert.equal(p2.origine, "meta");
assert.equal(d.clienti.find((z) => z.id === "z1").valuta, "USD", "valuta presa dall'account");
assert.equal(d.clienti.find((z) => z.id === "z2").valuta, "", "altro cliente invariato");

// Riapplicare gli stessi dati non cambia niente (niente salvataggi a vuoto)
const r2 = applicaMeta(d, meta, opz);
assert.equal(r2.cambiato, false, JSON.stringify(r2.conta));

// Cambio di budget: entra nello storico con la data di oggi
const meta3 = { ...meta, oggetti: meta.oggetti.map((o) => (o.id === "c1" ? { ...o, budget_giorno: 80 } : o)) };
const k1b = applicaMeta(d, meta3, opz).data.campagne.find((k) => k.id === "k1");
assert.equal(k1b.budgetGiorno, 80); assert.equal(k1b.storicoBudget.length, 2); assert.equal(k1b.storicoBudget[1].data, "2026-10-12");

// Account con pagamento in sospeso: le attive diventano "errore", con la data
const meta4 = { ...meta, account: [{ ...meta.account[0], stato: 3 }] };
const k1c = applicaMeta({ ...d, campagne: d.campagne.map((k) => (k.id === "k1" ? { ...k, stato: "attiva" } : k)) }, meta4, opz).data.campagne.find((k) => k.id === "k1");
assert.equal(k1c.stato, "errore"); assert.equal(k1c.erroreDal, "2026-10-12");

// Stati e obiettivi
assert.equal(statoCampagnaDaMeta({ stato: "ARCHIVED" }, null, "", "2026-10-12"), "conclusa");
assert.equal(statoCampagnaDaMeta({ stato: "ACTIVE", fine: "2026-10-01" }, null, "", "2026-10-12"), "conclusa");
assert.equal(obiettivoDaMeta("OUTCOME_TRAFFIC"), "traffico");
assert.equal(obiettivoDaMeta("OUTCOME_AWARENESS"), "awareness");

// Suggerimenti di collegamento
const conti = [
  { account_id: "1", nome: "BDF 1", business: "Boutique Del Fitness" },
  { account_id: "2", nome: "Studio Legale Campilongo", business: "Studio Legale Campilongo" },
  { account_id: "3", nome: "F3 motors €", business: "Gruppo Formula 3" },
  { account_id: "4", nome: "Studio Legale Loprieno", business: "Studio Legale Loprieno" },
];
assert.equal(suggerisciAccount({ nome: "Boutique del Fitness" }, conti)[0].account_id, "1");
assert.equal(suggerisciAccount({ nome: "Avv. Campilongo" }, conti)[0].account_id, "2");
assert.equal(suggerisciAccount({ nome: "Gruppo Formula 3" }, conti)[0].account_id, "3");
assert.ok(!suggerisciAccount({ nome: "Studio Legale Rossi" }, conti).some((a) => a.account_id === "4"), "\"studio legale\" da solo non basta");

console.log("meta-applica: tutte le prove passate ✔", r.conta);
