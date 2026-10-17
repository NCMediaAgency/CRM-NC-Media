// Prove delle winner ads (classifica) e del connettore per Claude (funzione crm-mcp con database simulato).
// Uso: node --no-warnings tests/analisi.test.mjs
import { readFileSync } from "node:fs";
import { stripTypeScriptTypes } from "node:module";
import assert from "node:assert/strict";
import { classifica, valutaAd, fotografie, conAndamento, benchmarkDi } from "../src/lib/winner.js";
import { trovaCliente, riassuntoCliente, righeAds, STRUMENTI } from "../supabase/functions/crm-mcp/logica.js";
import { applicaComandi } from "../src/lib/comandi.js";

const vicino = (a, b) => assert.ok(Math.abs(a - b) < 0.01, `${a} ≠ ${b}`);

// ---------- classifica ----------
// Come l'ad 71 dell'analisi di BDF: CPL 9,6 · appuntamento 17,6 · cliente 46 → (1,067 + 0,88 + 0,46) / 3
const ad71 = { ad_id: "71", nome: "71 - tempo perso", spesa: 736, lead: 736 / 9.6, appuntamenti: 736 / 17.6, clienti: 16, attiva: true, pubblico: "freddo" };
vicino(valutaAd(ad71, {}).punteggio, (9.6 / 9 + 17.6 / 20 + 46 / 100) / 3);
vicino(valutaAd(ad71, {}).chiusura, 16 / (736 / 17.6));
const ads = [
  ad71,
  { ad_id: "303", nome: "analisi evento", spesa: 945, lead: 89, appuntamenti: 55, clienti: 15, attiva: true, pubblico: "retargeting" },
  { ad_id: "9", nome: "pochi clienti", spesa: 100, lead: 20, appuntamenti: 10, clienti: 2, attiva: true },
  { ad_id: "8", nome: "spenta", spesa: 900, lead: 120, appuntamenti: 60, clienti: 20, attiva: false },
  { ad_id: "7", nome: "costosa", spesa: 3000, lead: 100, appuntamenti: 40, clienti: 10, attiva: true, pubblico: "freddo" },
];
let c = classifica(ads, {});
assert.deepEqual(c.righe.map((a) => a.ad_id), ["71", "303", "7"]);
assert.equal(c.esclusi, 2);
assert.deepEqual(c.righe.map((a) => a.pos), [1, 2, 3]);
assert.deepEqual(classifica(ads, { soloAttive: false }).righe.map((a) => a.ad_id), ["8", "71", "303", "7"]);
assert.deepEqual(classifica(ads, { pubblico: "freddo" }).righe.map((a) => a.ad_id), ["71", "7"]);
assert.deepEqual(classifica(ads, { minClienti: 0 }).righe.length, 4);
assert.equal(benchmarkDi({ winner: { cpl: 12 } }).cpl, 12);
assert.equal(benchmarkDi({ winner: { cpl: 12 } }).app, 20);
const foto = fotografie([{ data: "2026-10-01", fonte: "Hyros" }, { data: "2026-10-16", fonte: "Hyros" }, { data: "2026-10-16", fonte: "Hyros" }]);
assert.deepEqual(foto.map((f) => [f.data, f.n]), [["2026-10-16", 2], ["2026-10-01", 1]]);
const and = conAndamento([{ ad_id: "a", pos: 1 }, { ad_id: "b", pos: 2 }, { ad_id: "n", pos: 3 }], [{ ad_id: "b", pos: 1 }, { ad_id: "a", pos: 4 }]);
assert.deepEqual(and.map((x) => x.salita), [3, -1, null]);

// ---------- logica del connettore ----------
const clienti = [{ id: "z1", nome: "Boutique Del Fitness LLC", winner: { minClienti: 10 } }, { id: "s1", nome: "BDF Milano", padreId: "z1" }, { id: "z2", nome: "Oliva Consulting" }, { id: "z3", nome: "Vecchio", archiviato: "si" }];
assert.equal(trovaCliente(clienti, "boutique del fitness").cliente.id, "z1");
assert.equal(trovaCliente(clienti, "z2").cliente.id, "z2");
assert.match(trovaCliente(clienti, "Vecchio").errore, /Nessun cliente/);
assert.equal(trovaCliente([...clienti, { id: "z4", nome: "Oliva Studio" }], "oliva").errore.includes("Più clienti"), true);
assert.equal(trovaCliente(clienti, "BDF Milano").cliente.id, "s1");
const r = riassuntoCliente({ clienti, campagne: [{ id: "k", clienteId: "s1", nome: "Lead", stato: "attiva", spesa: 100, lead: 10 }], vendite: [{ clienteId: "z1", periodo: "2026-09", demoSvolte: 4, chiusure: 2 }], task: [] }, clienti[0]);
assert.equal(r.campagne[0].cpl, 10);
assert.equal(r.campagne[0].cliente, "BDF Milano");
assert.equal(r.benchmarkWinnerAds.minClienti, 10);
assert.equal(r.cicloDiVendita[0].chiusure, 2);
const ra = righeAds({ fonte: "Hyros", ads: [{ id: "1", nome: "A", spesa: "120,5", lead: 10, appuntamenti: 3, clienti: 1, pubblico: "freddo" }, { id: "1", nome: "doppio" }, { nome: "senza id" }] }, { ws: "ws1", clienteId: "z1", oggi: "2026-10-17" });
assert.equal(ra.righe.length, 1);
assert.equal(ra.righe[0].spesa, 120.5);
assert.equal(ra.data, "2026-10-17");
assert.ok(STRUMENTI.every((s) => s.name && s.description && s.inputSchema.type === "object"));

// ---------- funzione crm-mcp ----------
const ENVV = { SUPABASE_URL: "https://db.test", SUPABASE_SECRET_KEYS: '{"default":"sb_secret_x"}' };
let gestore = null;
globalThis.Deno = { env: { get: (k) => ENVV[k] }, serve: (f) => { gestore = f; } };
let src = stripTypeScriptTypes(readFileSync(new URL("../supabase/functions/crm-mcp/index.ts", import.meta.url), "utf8"));
src = src.replace(/from "\.\/logica\.js"/, `from "${new URL("../supabase/functions/crm-mcp/logica.js", import.meta.url).href}"`);
await import("data:text/javascript;base64," + Buffer.from(src).toString("base64"));
assert.ok(gestore);

const impronta = async (s) => [...new Uint8Array(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(s)))].map((b) => b.toString(16).padStart(2, "0")).join("");
const CHIAVE = "crmnc_" + "a".repeat(48), CHIAVE_TASK = "crmnc_" + "b".repeat(48);
const T = {
  integrazioni: [
    { id: "i1", workspace_id: "ws1", tipo: "mcp", attiva: true, chiave_hash: await impronta(CHIAVE) },
    { id: "i2", workspace_id: "ws1", tipo: "task", attiva: true, chiave_hash: await impronta(CHIAVE_TASK) },
  ],
  kv_store: [{ workspace_id: "ws1", key: "crm:data:v6", value: JSON.stringify({ clienti, campagne: [], vendite: [], task: [] }) }],
  analisi_ads: [], analisi_testi: [], crm_comandi: [],
};
let seq = 0;
const filtri = (q) => [...q.entries()].filter(([k]) => !["select", "order", "limit"].includes(k)).map(([k, v]) => {
  const [op, ...resto] = v.split("."); const val = resto.join(".");
  if (op === "eq") return (r) => String(r[k]) === val;
  throw new Error("filtro " + v);
});
function db(url, init) {
  const u = new URL(url); const p = u.pathname.replace("/rest/v1/", ""); const metodo = init.method || "GET";
  const corpo = init.body ? JSON.parse(init.body) : undefined;
  const t = T[p]; if (!t) throw new Error("tabella " + p);
  const f = filtri(u.searchParams); const passa = (r) => f.every((g) => g(r));
  if (metodo === "GET") return t.filter(passa).slice(0, Number(u.searchParams.get("limit") || 1e9));
  if (metodo === "PATCH") { t.filter(passa).forEach((r) => Object.assign(r, corpo)); return null; }
  if (metodo === "DELETE") { T[p] = t.filter((r) => !passa(r)); return null; }
  for (const r of Array.isArray(corpo) ? corpo : [corpo]) T[p].push({ id: ++seq, creato: "2026-10-17T07:00:00Z", ...r });
  return null;
}
globalThis.fetch = async (url, init = {}) => {
  const s = String(url);
  const json = (c, st = 200) => ({ ok: st < 400, status: st, json: async () => c, text: async () => (c === null ? "" : JSON.stringify(c)) });
  if (s.startsWith("https://db.test/rest/v1/")) return json(db(s, init));
  throw new Error("rete non simulata: " + s);
};
let nId = 0;
const rpc = async (method, params, chiave = CHIAVE) => {
  const res = await gestore(new Request(`https://db.test/functions/v1/crm-mcp/${chiave}`, { method: "POST", headers: { "Content-Type": "application/json", Accept: "application/json, text/event-stream" }, body: JSON.stringify({ jsonrpc: "2.0", id: ++nId, method, params }) }));
  return { stato: res.status, corpo: res.status === 202 ? null : await res.json() };
};
const usa = async (name, args) => { const r = await rpc("tools/call", { name, arguments: args }); return { testo: r.corpo.result.content[0].text, errore: !!r.corpo.result.isError }; };

assert.equal((await rpc("initialize", {}, "crmnc_" + "c".repeat(48))).stato, 401);
assert.equal((await rpc("initialize", {}, CHIAVE_TASK)).stato, 401);   // la chiave di Boschetto qui non vale
const init = await rpc("initialize", { protocolVersion: "2025-03-26", capabilities: {}, clientInfo: { name: "claude-ai", version: "1" } });
assert.equal(init.corpo.result.protocolVersion, "2025-03-26");
assert.ok(init.corpo.result.capabilities.tools);
const notifica = await gestore(new Request(`https://db.test/functions/v1/crm-mcp/${CHIAVE}`, { method: "POST", body: JSON.stringify({ jsonrpc: "2.0", method: "notifications/initialized" }) }));
assert.equal(notifica.status, 202);
assert.equal((await rpc("tools/list")).corpo.result.tools.length, STRUMENTI.length);
// chiave nell'intestazione Authorization e indirizzo senza chiave (come la conserva claude.ai)
const conIntestazione = await gestore(new Request("https://db.test/functions/v1/crm-mcp", { method: "POST", headers: { Authorization: `Bearer ${CHIAVE}`, "Content-Type": "application/json" }, body: JSON.stringify({ jsonrpc: "2.0", id: 99, method: "tools/list" }) }));
assert.equal(conIntestazione.status, 200);
assert.equal((await gestore(new Request("https://db.test/functions/v1/crm-mcp", { method: "POST", body: "{}" }))).status, 401);
assert.equal((await rpc("metodo/strano")).corpo.error.code, -32601);
assert.match((await usa("elenco_clienti", {})).testo, /BDF Milano/);
assert.ok(!(await usa("elenco_clienti", {})).testo.includes("Vecchio"));
assert.equal(JSON.parse((await usa("leggi_benchmark", { cliente: "Boutique del fitness" })).testo).minClienti, 10);
assert.equal((await usa("dati_cliente", { cliente: "Inesistente" })).errore, true);
let sv = await usa("salva_dati_ads", { cliente: "Boutique del Fitness", fonte: "Hyros", data: "2026-10-16", periodo: "ultimi 30 giorni", ads: [{ id: "71", nome: "tempo perso", spesa: 736, lead: 77, appuntamenti: 42, clienti: 16 }, { id: "303", nome: "analisi evento", spesa: 945, lead: 89, appuntamenti: 55, clienti: 15 }] });
assert.match(sv.testo, /Salvati 2 annunci/);
// stessa data: la fotografia si sostituisce
await usa("salva_dati_ads", { cliente: "z1", fonte: "Hyros", data: "2026-10-16", ads: [{ id: "71", nome: "tempo perso", spesa: 800, lead: 80, appuntamenti: 44, clienti: 17 }] });
assert.equal(T.analisi_ads.length, 1);
assert.equal(T.analisi_ads[0].spesa, 800);
assert.equal(T.analisi_ads[0].workspace_id, "ws1");
assert.equal(trovaCliente(clienti, "BDF").cliente.id, "z1");   // le iniziali vincono sul sottocliente che contiene "BDF"
sv = await usa("salva_analisi", { cliente: "BDF", tipo: "winner_ads", titolo: "Winner ads 16 ottobre", testo: "## Da scalare\n- **71** tempo perso" });
assert.equal(sv.errore, false);
assert.equal(T.analisi_testi[0].cliente_id, "z1");
assert.equal(T.analisi_testi[0].tipo, "winner_ads");
assert.equal((await usa("salva_analisi", { cliente: "z1", titolo: "vuota", testo: "  " })).errore, true);
const lette = JSON.parse((await usa("leggi_analisi", { cliente: "z1" })).testo);
assert.equal(lette[0].titolo, "Winner ads 16 ottobre");
// task in coda, applicati dal CRM con origine claude
sv = await usa("crea_task", { cliente: "Boutique del Fitness", titolo: "Rifare la creatività 284", priorita: "alta", scadenza: "2026-10-20" });
assert.match(sv.testo, /entro un minuto/);
assert.equal(T.crm_comandi[0].tipo, "crea_task");
const applicato = applicaComandi({ clienti, task: [], aree: [] }, T.crm_comandi, { uid: () => "u1", oggi: "2026-10-17" }).data.task[0];
assert.equal(applicato.clienteId, "z1");
assert.equal(applicato.origine, "claude");
assert.equal(applicato.priorita, "alta");
// in pausa: rifiutato
T.integrazioni[0].attiva = false;
assert.equal((await rpc("tools/list")).stato, 401);

console.log("analisi: tutte le prove passate");
