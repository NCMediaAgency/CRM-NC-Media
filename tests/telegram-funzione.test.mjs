// Prova della funzione telegram senza rete: Telegram, Claude e il database sono simulati.
// Uso: node tests/telegram-funzione.test.mjs
import { readFileSync } from "node:fs";
import { stripTypeScriptTypes } from "node:module";
import assert from "node:assert/strict";

// ---------- Claude finto: legge il messaggio e risponde con strumenti ----------
globalThis.__claude = [];
const stubClaude = `export default class Anthropic {
  constructor(o) { this.o = o; }
  get beta() { return { messages: { create: async (req) => { globalThis.__claude.push(req); return globalThis.__rispostaClaude(req); } } }; }
}`;
globalThis.__rispostaClaude = (req) => {
  const ultimo = req.messages.at(-1).content.split("Messaggio del titolare:").pop();
  if (/report/i.test(ultimo)) return { stop_reason: "tool_use", content: [
    { type: "thinking", thinking: "" },
    { type: "tool_use", id: "x1", name: "crea_task", input: { titolo: "Preparare il report di ottobre", clienteId: "z1", scadenza: "2026-10-23", ora: null, priorita: "alta", note: null, areaId: null } },
    { type: "tool_use", id: "x2", name: "crea_nota", input: { testo: "Spente le campagne Rottamazione", clienteId: "z2", data: null, tipo: "esterno" } },
  ] };
  return { stop_reason: "end_turn", content: [{ type: "text", text: "Domani hai 1 task: Fatture." }] };
};

// ---------- carico la funzione, senza tipi e con gli import sostituiti ----------
const ENVV = { SUPABASE_URL: "https://db.test", SUPABASE_SECRET_KEYS: '{"default":"sb_secret_x"}', SUPABASE_PUBLISHABLE_KEYS: '{"default":"sb_publishable_y"}', TELEGRAM_TOKEN: "123:abc", ANTHROPIC_API_KEY: "sk-test" };
let gestore = null; const inSottofondo = [];
globalThis.Deno = { env: { get: (k) => ENVV[k] }, serve: (f) => { gestore = f; } };
globalThis.EdgeRuntime = { waitUntil: (p) => inSottofondo.push(p) };
let src = stripTypeScriptTypes(readFileSync(new URL("../supabase/functions/telegram/index.ts", import.meta.url), "utf8"));
src = src.replace(/import Anthropic from "npm:@anthropic-ai\/sdk";/, `import Anthropic from "data:text/javascript;base64,${Buffer.from(stubClaude).toString("base64")}";`)
  .replace(/from "\.\/logica\.js"/, `from "${new URL("../supabase/functions/telegram/logica.js", import.meta.url).href}"`);
await import("data:text/javascript;base64," + Buffer.from(src).toString("base64"));
assert.ok(gestore);

// ---------- database finto ----------
const T = { telegram_chat: [], telegram_codici: [], telegram_storico: [], telegram_inviati: [], crm_comandi: [], kv_store: [] };
const CHIAVI = { telegram_chat: ["chat_id"], telegram_inviati: ["workspace_id", "chiave"] };
const datiCRM = {
  clienti: [{ id: "z1", nome: "Boutique Del Fitness LLC" }, { id: "z2", nome: "Oliva Consulting" }],
  task: [{ id: "t2", titolo: "Call Oliva", clienteId: "z2", scadenza: "2000-01-01", priorita: "alta", stato: "dafare" }],
  campagne: [], vendite: [], aree: [],
};
T.kv_store.push({ workspace_id: "ws1", key: "crm:data:v6", value: JSON.stringify(datiCRM) });
const filtri = (q) => [...q.entries()].filter(([k]) => !["select", "on_conflict", "order", "limit"].includes(k)).map(([k, v]) => {
  const [op, ...resto] = v.split("."); const val = resto.join(".");
  if (op === "eq") return (r) => String(r[k]) === val;
  if (op === "lt") return (r) => String(r[k]) < val;
  if (op === "is") return (r) => r[k] === null || r[k] === undefined;
  throw new Error("filtro " + v);
});
function db(url, init) {
  const u = new URL(url); const p = u.pathname.replace("/rest/v1/", ""); const metodo = init.method || "GET";
  const corpo = init.body ? JSON.parse(init.body) : undefined;
  if (p === "rpc/meta_cron_chiave") return "chiave-cron";
  if (p === "rpc/my_workspace") return "ws1";
  const t = T[p]; if (!t) throw new Error("tabella " + p);
  const f = filtri(u.searchParams); const passa = (r) => f.every((g) => g(r));
  if (metodo === "GET") { let l = t.filter(passa); if (u.searchParams.get("order") === "id.desc") l = [...l].reverse(); return l.slice(0, Number(u.searchParams.get("limit") || 1e9)); }
  if (metodo === "PATCH") { t.filter(passa).forEach((r) => Object.assign(r, corpo)); return null; }
  if (metodo === "DELETE") { T[p] = t.filter((r) => !passa(r)); return null; }
  for (const r of Array.isArray(corpo) ? corpo : [corpo]) {
    const ch = CHIAVI[p]; const e = ch && t.find((x) => ch.every((k) => String(x[k]) === String(r[k])));
    if (e) { if (!/ignore/.test(init.headers.Prefer || "")) Object.assign(e, r); } else t.push({ creato: new Date().toISOString(), applicato: null, ...r });
  }
  return null;
}
const tgChiamate = [];
globalThis.fetch = async (url, init = {}) => {
  const s = String(url);
  const json = (c, st = 200) => ({ ok: st < 400, status: st, json: async () => c, text: async () => (c === null ? "" : JSON.stringify(c)) });
  if (s.startsWith("https://db.test/rest/v1/")) return json(db(s, init));
  if (s === "https://db.test/auth/v1/user") return json({ id: "u1" }, init.headers.Authorization === "Bearer utente" ? 200 : 401);
  const m = s.match(/^https:\/\/api\.telegram\.org\/bot123:abc\/(\w+)$/);
  if (m) {
    const corpo = JSON.parse(init.body); tgChiamate.push({ metodo: m[1], ...corpo });
    return json({ ok: true, result: m[1] === "getMe" ? { username: "crm_nc_bot" } : { message_id: 1 } });
  }
  throw new Error("rete non simulata: " + s);
};
const segreto = async () => {
  const h = await crypto.subtle.digest("SHA-256", new TextEncoder().encode("crm-nc-media:123:abc"));
  return [...new Uint8Array(h)].map((b) => b.toString(16).padStart(2, "0")).join("").slice(0, 48);
};
const chiama = async (corpo, headers) => {
  const r = await gestore(new Request("https://db.test/functions/v1/telegram", { method: "POST", headers: { "Content-Type": "application/json", ...headers }, body: JSON.stringify(corpo) }));
  while (inSottofondo.length) await inSottofondo.shift();
  return { stato: r.status, corpo: await r.json() };
};
const daTelegram = async (u) => chiama(u, { "X-Telegram-Bot-Api-Secret-Token": await segreto() });
const ultimoMsg = () => tgChiamate.filter((c) => c.metodo === "sendMessage").at(-1);

// 1. webhook con il segreto sbagliato: rifiutato
assert.equal((await chiama({ message: {} }, { "X-Telegram-Bot-Api-Secret-Token": "no" })).stato, 401);

// 2. dal CRM: collega → webhook impostato, link con codice
let r = await chiama({ azione: "collega" }, { Authorization: "Bearer utente" });
assert.equal(r.stato, 200, JSON.stringify(r.corpo));
assert.match(r.corpo.link, /^https:\/\/t\.me\/crm_nc_bot\?start=\w{24}$/);
const wh = tgChiamate.find((c) => c.metodo === "setWebhook");
assert.equal(wh.url, "https://db.test/functions/v1/telegram"); assert.equal(wh.secret_token, await segreto());
const codice = r.corpo.link.split("start=")[1];
assert.equal((await chiama({ azione: "collega" }, { Authorization: "Bearer altro" })).stato, 401);

// 3. chat sconosciuta: niente Claude
await daTelegram({ message: { chat: { id: 99 }, text: "ciao" } });
assert.match(ultimoMsg().text, /CRM privato/); assert.equal(globalThis.__claude.length, 0);

// 4. /start CODICE: chat collegata, codice consumato
await daTelegram({ message: { chat: { id: 42 }, from: { first_name: "Nicola" }, text: "/start " + codice } });
assert.equal(T.telegram_chat[0].workspace_id, "ws1"); assert.equal(T.telegram_chat[0].utente, "Nicola");
assert.match(ultimoMsg().text, /Collegato al CRM/); assert.ok(!T.telegram_codici.length);

// 5. messaggio: Claude crea un task e una nota → due comandi in coda e la conferma
await daTelegram({ message: { chat: { id: 42 }, text: "Task per BDF venerdì 23: report di ottobre, alta. E nel diario: spente le Rottamazione" } });
const req = globalThis.__claude.at(-1);
assert.equal(req.model, "claude-opus-5-5"); assert.deepEqual(req.betas, ["server-side-fallback-2026-07-01"]); assert.equal(req.fallbacks, "default");
assert.equal(req.tool_choice.type, "auto"); assert.equal(req.output_config.effort, "low");
assert.match(req.messages.at(-1).content, /"Boutique Del Fitness LLC"/, "il contesto CRM arriva a Claude");
assert.deepEqual(T.crm_comandi.map((c) => c.tipo), ["crea_task", "crea_nota"]);
assert.equal(T.crm_comandi[0].dati.clienteId, "z1"); assert.equal(T.crm_comandi[0].workspace_id, "ws1");
const conf = ultimoMsg();
assert.match(conf.text, /Task creato · Preparare il report di ottobre · Boutique Del Fitness LLC · ven 23 ott · priorità alta/);
assert.match(conf.text, /Nota nel diario · Oliva Consulting/);
assert.equal(T.telegram_storico.length, 2);

// 6. domanda: risposta in testo, nessun comando, la storia arriva a Claude
await daTelegram({ message: { chat: { id: 42 }, text: "Che scadenze ho domani?" } });
assert.equal(ultimoMsg().text, "Domani hai 1 task: Fatture."); assert.equal(T.crm_comandi.length, 2);
assert.equal(globalThis.__claude.at(-1).messages[0].role, "user"); assert.ok(globalThis.__claude.at(-1).messages.length >= 3, "ricorda lo scambio precedente");

// 7. pulsante Fatto
await daTelegram({ callback_query: { id: "q1", data: "f:t2", message: { message_id: 7, chat: { id: 42 }, reply_markup: { inline_keyboard: [[{ text: "✅ Call", callback_data: "f:t2" }, { text: "⏰", callback_data: "d:t2" }]] } } } });
assert.deepEqual(T.crm_comandi.at(-1).dati, { taskId: "t2", stato: "fatto" });
assert.equal(tgChiamate.find((c) => c.metodo === "answerCallbackQuery").text, "✅ Fatto");

// 8. cron: riepilogo una volta sola al giorno
T.telegram_chat[0].ora_riepilogo = "00:00"; T.telegram_chat[0].weekend = true; T.crm_comandi = T.crm_comandi.filter((c) => c.tipo !== "modifica_task");
const primaInvii = tgChiamate.filter((c) => c.metodo === "sendMessage").length;
r = await chiama({ azione: "promemoria" }, { "x-cron-key": "chiave-cron" });
assert.equal(r.corpo.chat, 1);
assert.match(ultimoMsg().text, /Buongiorno/); assert.match(ultimoMsg().text, /In ritardo \(1\)/);
assert.equal(ultimoMsg().reply_markup.inline_keyboard[0][0].callback_data, "f:t2");
await chiama({ azione: "promemoria" }, { "x-cron-key": "chiave-cron" });
assert.equal(tgChiamate.filter((c) => c.metodo === "sendMessage").length, primaInvii + 1, "niente doppione");
assert.equal((await chiama({ azione: "promemoria" }, { "x-cron-key": "sbagliata" })).stato, 401);

// 9. preferenze dal CRM
r = await chiama({ azione: "preferenze", chat: 42, ora_riepilogo: "8:30", weekend: false, preavviso: 15 }, { Authorization: "Bearer utente" });
assert.equal(T.telegram_chat[0].ora_riepilogo, "08:30"); assert.equal(T.telegram_chat[0].preavviso, 15); assert.equal(T.telegram_chat[0].weekend, false);

console.log("telegram (funzione): tutte le prove passate ✔", { messaggi: tgChiamate.filter((c) => c.metodo === "sendMessage").length });
