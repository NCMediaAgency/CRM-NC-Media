// Prova della lettura dei file KPI (src/lib/kpi.js) con due fogli d'esempio nei formati più comuni.
// Uso: node tests/kpi.test.mjs
import assert from "node:assert/strict";
import { trovaTabelleKpi, applicaKpi, numeroCella, leggiMese } from "../src/lib/kpi.js";

assert.equal(numeroCella("€ 47,872.79"), 47872.79);
assert.equal(numeroCella("6,096"), 6096);
assert.equal(numeroCella("1.350,00"), 1350);
assert.equal(numeroCella("6.096"), 6096);
assert.equal(numeroCella("/"), null);
assert.equal(numeroCella("#DIV/0!"), null);
assert.equal(numeroCella("8/1/12"), null, "una data finita in una cella di importo non è un numero");
assert.equal(numeroCella("35.12%"), null);
assert.deepEqual(leggiMese("GENNAIO 2026"), { m: 1, anno: 2026 });
assert.deepEqual(leggiMese("Ottobre"), { m: 10, anno: null });
assert.equal(leggiMese("Aprile (da 15)"), null);

// Foglio con un blocco per mese e una sottocolonna per funnel (riga in mezzo fra mesi e funnel)
const analisi = [
  ["", "GENNAIO 2026", "", "", "", "FEBBRAIO 2026", "", ""],
  ["FATTURATO MESE", 1000, "", "", "", 2000, "", ""],
  ["", "FUNNEL OT", "BOOST REELS", "TOTALE", "", "FUNNEL OT", "BOOST REELS", "TOTALE"],
  ["Spesa", 100, 50, 150, "", 200, 60, 260],
  ["Numero Leads", 400, "/", 400, "", 500, "/", 500],
  ["% Lead/Appuntamenti", 0.3, "/", 0.3, "", 0.4, "/", 0.4],
  ["Appuntamenti Fissati ", 120, 80, 200, "", 130, 90, 220],
  ["Appuntamenti Effettivi", 70, 60, 130, "", 80, 70, 150],
  ["NUOVI CLIENTI", 20, 25, 45, "", 22, 28, 50],
  ["FATTURATO NUOVI CLIENTI", 9000, 8000, 17000, "", 9500, 8500, 18000],
  ["INCASSATO NUOVI CLIENTI", 6000, 5000, 11000, "", 6500, 5500, 12000],
  ["TOTALE INCASSATO", 7000, 6000, 13000, "", 7500, 6500, 14000],
];
// Foglio tipo dashboard: un blocco per funnel, anno nel nome del foglio, canali senza totale
const dashboard = [
  ["FUNNEL BOOST REEL", "OTTOBRE", "NOVEMBRE", "DICEMBRE"],
  ["SPESA", "€ 1,000.00", "€ 2,000.00", ""],
  ["Appuntamenti Fissati IG", 10, 20, ""],
  ["Appuntamenti Fissati WP", 5, 6, ""],
  ["Appuntamenti Fissati Totali", 15, 26, ""],
  ["CLIENTI WP", 3, 4, ""],
  ["CLIENTI IG", 2, 2, ""],
  ["CLIENTI TOTALI", 5, 6, ""],
  ["INCASSATO NUOVI CLIENTI WP", 300, 400, ""],
  ["INCASSATO NUOVI CLIENTI IG", 200, 100, ""],
];
const t = trovaTabelleKpi([{ nome: "ANALISI TEAM", righe: analisi }, { nome: "DASHBOARD 2025", righe: dashboard }]);
const di = (g) => t.find((x) => x.gruppo === g);
assert.ok(di("TOTALE") && di("FUNNEL OT") && di("BOOST REELS") && di("FUNNEL BOOST REEL"), t.map((x) => x.gruppo).join());
assert.deepEqual(di("TOTALE").mesi.map((x) => x.periodo), ["2026-01", "2026-02"]);
assert.deepEqual(di("TOTALE").mesi[0].valori, { spesa: 150, lead: 400, demoFissate: 200, demoSvolte: 130, chiusure: 45, valoreContratto: 17000, incassato: 11000 });
assert.equal(di("BOOST REELS").mesi[0].valori.lead, undefined, "\"/\" non è zero: il campo resta fuori");
const boost = di("FUNNEL BOOST REEL");
assert.deepEqual(boost.mesi.map((x) => x.periodo), ["2025-10", "2025-11"], "dicembre vuoto non si importa");
assert.equal(boost.mesi[0].valori.spesa, 1000);
assert.equal(boost.mesi[0].valori.demoFissate, 15, "vince la riga dei totali");
assert.equal(boost.mesi[0].valori.chiusure, 5);
assert.equal(boost.mesi[0].valori.incassato, 500, "canali senza totale: WP + IG");

// Scrittura: aggiorna i campi presenti, crea i mesi mancanti, ricaricare non somma
let n = 0;
const prima = [{ id: "v1", clienteId: "z1", periodo: "2026-01", lead: 1, chiamate: 99, chiusure: 1, slot: 40 }];
let r = applicaKpi(prima, "z1", di("TOTALE").mesi, { uid: () => "n" + ++n, fonte: "kpi.xlsx" });
assert.equal(r.nuovi, 1); assert.equal(r.aggiornati, 1);
const gen = r.vendite.find((v) => v.periodo === "2026-01");
assert.equal(gen.lead, 400); assert.equal(gen.chiamate, 99, "i campi non presenti nel file restano"); assert.equal(gen.slot, 40);
r = applicaKpi(r.vendite, "z1", di("TOTALE").mesi, { uid: () => "n" + ++n });
assert.equal(r.vendite.filter((v) => v.clienteId === "z1").length, 2, "ricaricare non crea doppioni");
assert.equal(r.vendite.find((v) => v.periodo === "2026-02").chiusure, 50);

console.log("kpi: tutte le prove passate ✔", t.map((x) => `${x.foglio}·${x.gruppo}:${x.mesi.length}`).join("  "));
