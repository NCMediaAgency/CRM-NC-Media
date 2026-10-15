// Prove dell'integrazione con i software esterni (Boschetto): lettura delle task,
// applicazione nel CRM, canale di ritorno e funzione integrazioni con database simulato.
// Uso: node --no-warnings tests/integrazioni.test.mjs
import { readFileSync } from "node:fs";
import { stripTypeScriptTypes } from "node:module";
import assert from "node:assert/strict";
import { leggiTaskEsterna, eventiDaCorpo } from "../supabase/functions/integrazioni/logica.js";
import { applicaComandi, statoEsterno, nomeChiave } from "../src/lib/comandi.js";

// ---------- lettura delle task ----------
const t1 = leggiTaskEsterna({ id: 42, title: "Grafiche ottobre", due: "2026-10-23T15:30:00", priority: "High", status: "In Progress", project: "Ristorante Da Mario", url: "https://boschetto.app/t/42" });
assert.deepEqual(t1, { esternoId: "42", titolo: "Grafiche ottobre", descrizione: "", scadenza: "2026-10-23", ora: "15:30", priorita: "alta", stato: "incorso", cliente: "Ristorante Da Mario", url: "https://boschetto.app/t/42", assegnatoA: "" });
assert.equal(leggiTaskEsterna({ titolo: "senza id" }), null);
assert.equal(leggiTaskEsterna({ id: "a", scadenza: "23/10/2026", stato: "da fare", url: "javascript:alert(1)" }).scadenza, "2026-10-23");
assert.equal(leggiTaskEsterna({ id: "a", url: "javascript:alert(1)" }).url, "");
assert.equal(eventiDaCorpo({ evento: "task.eliminata", task: { id: "x" } })[0].eliminata, true);
assert.equal(eventiDaCorpo({ tasks: [{ id: "a" }, { nulla: 1 }, { id: "b", deleted: true }] }).length, 2);
assert.equal(eventiDaCorpo({ tasks: [{ id: "b", deleted: true }] })[0].eliminata, true);
assert.equal(nomeChiave("  Ristorante  DA Mário "), "ristorante da mario");

// ---------- applicazione nel CRM ----------
let n = 0; const uid = () => "n" + ++n;
const base = { clienti: [{ id: "k", nome: "Kesia" }, { id: "s1", nome: "Da Mario", padreId: "k" }, { id: "s2", nome: "Bar Sport", padreId: "k" }], task: [], annotazioni: [],
  integrazioni: [{ id: "i1", nome: "Boschetto", clienteBase: "k", abbinamenti: { "ristorante da mario": "s1" }, nomiVisti: {} }] };
const cmd = (dati) => ({ id: "c" + ++n, tipo: "importa_task", dati: { integrazioneId: "i1", fonte: "Boschetto", ...dati } });
const opz = { uid, oggi: "2026-10-15" };
let r = applicaComandi(base, [cmd(t1), cmd({ esternoId: "43", titolo: "Post Bar", cliente: "Bar Sport", stato: "dafare", priorita: "media" })], opz);
let d = r.data;
assert.equal(d.task.length, 2);
assert.equal(d.task[0].clienteId, "s1");          // abbinato
assert.equal(d.task[1].clienteId, "k");           // senza abbinamento: cliente principale
assert.equal(d.task[0].origine, "Boschetto");
assert.deepEqual(d.task[0].esterno.ultimo, statoEsterno(d.task[0]));
assert.deepEqual(d.integrazioni[0].nomiVisti, { "ristorante da mario": "Ristorante Da Mario", "bar sport": "Bar Sport" });
// stesso id → aggiorna, non duplica; il cliente spostato a mano resta
d = { ...d, task: d.task.map((t) => (t.esterno.id === "43" ? { ...t, clienteId: "s2" } : t)) };
r = applicaComandi(d, [cmd({ esternoId: "43", titolo: "Post Bar (rivisto)", cliente: "Bar Sport", stato: "incorso" })], opz);
d = r.data;
assert.equal(d.task.length, 2);
assert.equal(d.task[1].titolo, "Post Bar (rivisto)");
assert.equal(d.task[1].clienteId, "s2");
assert.equal(r.esiti[0].esito, "task aggiornata");
// eliminata alla fonte → chiusa e fuori dal ritorno
r = applicaComandi(d, [cmd({ esternoId: "42", eliminata: true })], opz);
assert.equal(r.data.task[0].stato, "fatto");
assert.equal(r.data.task[0].esterno.eliminata, true);
// integrazione sconosciuta ai dati: si crea la sua scheda
r = applicaComandi({ ...base, integrazioni: [] }, [cmd({ esternoId: "1", titolo: "x" })], opz);
assert.equal(r.data.integrazioni[0].nome, "Boschetto");
assert.equal(r.data.task[0].clienteId, "");

// ---------- funzione integrazioni ----------
const ENVV = { SUPABASE_URL: "https://db.test", SUPABASE_SECRET_KEYS: '{"default":"sb_secret_x"}', SUPABASE_PUBLISHABLE_KEYS: '{"default":"sb_publishable_y"}' };
let gestore = null;
globalThis.Deno = { env: { get: (k) => ENVV[k] }, serve: (f) => { gestore = f; } };
let src = stripTypeScriptTypes(readFileSync(new URL("../supabase/functions/integrazioni/index.ts", import.meta.url), "utf8"));
src = src.replace(/from "\.\/logica\.js"/, `from "${new URL("../supabase/functions/integrazioni/logica.js", import.meta.url).href}"`);
await import("data:text/javascript;base64," + Buffer.from(src).toString("base64"));
assert.ok(gestore);

const T = { integrazioni: [], integrazione_uscite: [], crm_comandi: [] };
let seq = 0;
const filtri = (q) => [...q.entries()].filter(([k]) => !["select", "order", "limit"].includes(k)).map(([k, v]) => {
  const [op, ...resto] = v.split("."); const val = resto.join(".");
  if (op === "eq") return (r) => String(r[k]) === val;
  if (op === "gt") return (r) => Number(r[k]) > Number(val);
  if (op === "in") { const l = val.replace(/^\(|\)$/g, "").split(","); return (r) => l.includes(String(r[k])); }
  throw new Error("filtro " + v);
});
function db(url, init) {
  const u = new URL(url); const p = u.pathname.replace("/rest/v1/", ""); const metodo = init.method || "GET";
  const corpo = init.body ? JSON.parse(init.body) : undefined;
  if (p === "rpc/my_workspace") return "ws1";
  const t = T[p]; if (!t) throw new Error("tabella " + p);
  const f = filtri(u.searchParams); const passa = (r) => f.every((g) => g(r));
  if (metodo === "GET") {
    const sel = u.searchParams.get("select"), l = t.filter(passa).slice(0, Number(u.searchParams.get("limit") || 1e9));
    return !sel || sel === "*" ? l : l.map((r) => Object.fromEntries(sel.split(",").map((k) => [k, r[k]])));
  }
  if (metodo === "PATCH") { t.filter(passa).forEach((r) => Object.assign(r, corpo)); return null; }
  if (metodo === "DELETE") { T[p] = t.filter((r) => !passa(r)); return null; }
  const nuove = (Array.isArray(corpo) ? corpo : [corpo]).map((r) => ({
    ...(p === "integrazioni" ? { id: "int-" + ++seq, url_ritorno: "", chiave_ritorno: "", attiva: true, ultimo_contatto: null } : {}),
    ...(p === "integrazione_uscite" ? { id: ++seq, inviato: null, esito: "" } : {}),
    creato: "2026-10-15T10:00:00Z", ...r }));
  t.push(...nuove);
  return /representation/.test(init.headers.Prefer || "") ? nuove : null;
}
const ritorni = [];
globalThis.fetch = async (url, init = {}) => {
  const s = String(url);
  const json = (c, st = 200) => ({ ok: st < 400, status: st, json: async () => c, text: async () => (c === null ? "" : JSON.stringify(c)) });
  if (s.startsWith("https://db.test/rest/v1/")) return json(db(s, init));
  if (s === "https://db.test/auth/v1/user") return json({ id: "u1" }, init.headers.Authorization === "Bearer utente" ? 200 : 401);
  if (s === "https://boschetto.test/crm") { ritorni.push({ auth: init.headers.Authorization, ...JSON.parse(init.body) }); return json({ ok: true }); }
  throw new Error("rete non simulata: " + s);
};
const chiama = async (corpo, auth = "Bearer utente") => {
  const res = await gestore(new Request("https://db.test/functions/v1/integrazioni", { method: "POST", headers: { Authorization: auth, "Content-Type": "application/json" }, body: JSON.stringify(corpo) }));
  return { stato: res.status, corpo: await res.json() };
};

assert.equal((await chiama({ azione: "stato" }, "Bearer altro")).stato, 401);
const creata = await chiama({ azione: "crea", nome: "Boschetto" });
assert.equal(creata.stato, 200);
assert.match(creata.corpo.chiave, /^crmnc_[0-9a-f]{48}$/);
assert.equal(creata.corpo.indirizzo, "https://db.test/functions/v1/integrazioni");
assert.ok(!JSON.stringify(T.integrazioni).includes(creata.corpo.chiave));   // in archivio solo l'impronta
const id = creata.corpo.id, chiave = "Bearer " + creata.corpo.chiave;
const st = await chiama({ azione: "stato" });
assert.equal(st.corpo.integrazioni.length, 1);
assert.equal(st.corpo.integrazioni[0].chiave_hash, undefined);

// andata
assert.equal((await chiama({ task: { id: "1" } }, "Bearer crmnc_sbagliata")).stato, 401);
assert.equal((await chiama({ task: { titolo: "senza id" } }, chiave)).stato, 400);
const inv = await chiama({ evento: "task.creata", tasks: [{ id: "42", title: "Grafiche", project: "Da Mario" }, { id: "43", titolo: "Post" }] }, chiave);
assert.deepEqual(inv.corpo, { ok: true, ricevute: 2 });
assert.equal(T.crm_comandi.length, 2);
assert.equal(T.crm_comandi[0].tipo, "importa_task");
assert.equal(T.crm_comandi[0].workspace_id, "ws1");
assert.equal(T.crm_comandi[0].dati.integrazioneId, id);
assert.equal(T.crm_comandi[0].dati.cliente, "Da Mario");
assert.ok(T.integrazioni[0].ultimo_contatto);
// i comandi della coda si applicano nel CRM
r = applicaComandi({ clienti: [], task: [], integrazioni: [] }, T.crm_comandi, opz);
assert.equal(r.data.task.length, 2);

// ritorno a lettura
await chiama({ azione: "notifica", id, modifiche: [{ esternoId: "42", stato: "fatto", cambiati: ["stato"] }, { stato: "senza id" }] });
const lette = await chiama({ azione: "modifiche", dal: 0 }, chiave);
assert.equal(lette.corpo.modifiche.length, 1);
assert.equal(lette.corpo.modifiche[0].esternoId, "42");
assert.equal(lette.corpo.modifiche[0].stato, "fatto");
assert.equal((await chiama({ azione: "modifiche", dal: lette.corpo.ultimo }, chiave)).corpo.modifiche.length, 0);
// ritorno a spinta
assert.equal((await chiama({ azione: "ritorno", id, url: "http://boschetto.test/crm" })).stato, 400);
await chiama({ azione: "ritorno", id, url: "https://boschetto.test/crm", chiave: "segreta-boschetto" });
await chiama({ azione: "notifica", id, modifiche: [{ esternoId: "43", scadenza: "2026-10-30", cambiati: ["scadenza"] }] });
assert.equal(ritorni.length, 1);
assert.equal(ritorni[0].auth, "Bearer segreta-boschetto");
assert.equal(ritorni[0].evento, "task.modificata_nel_crm");
assert.equal(ritorni[0].modifiche[0].esternoId, "43");
assert.equal(T.integrazione_uscite.at(-1).esito, "ok");
// pausa, nuova chiave, altro spazio di lavoro
await chiama({ azione: "attiva", id, attiva: false });
assert.equal((await chiama({ task: { id: "9" } }, chiave)).stato, 401);
await chiama({ azione: "attiva", id, attiva: true });
const nuova = await chiama({ azione: "rigenera", id });
assert.equal((await chiama({ task: { id: "9" } }, chiave)).stato, 401);
assert.equal((await chiama({ task: { id: "9" } }, "Bearer " + nuova.corpo.chiave)).stato, 200);
T.integrazioni.push({ id: "altrui", workspace_id: "ws2", nome: "X", chiave_hash: "h", attiva: true });
assert.equal((await chiama({ azione: "elimina", id: "altrui" })).stato, 400);
assert.equal(T.integrazioni.length, 2);
await chiama({ azione: "elimina", id });
assert.equal(T.integrazioni.length, 1);

console.log("integrazioni: tutte le prove passate");
