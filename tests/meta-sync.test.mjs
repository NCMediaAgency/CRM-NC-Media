// Prova della funzione meta-sync senza rete: Meta e il database sono simulati.
// Uso: node tests/meta-sync.test.mjs
import { readFileSync } from "node:fs";
import assert from "node:assert/strict";

// ---------- carico la funzione come modulo, con Deno ed EdgeRuntime finti ----------
// Chiavi nuove di Supabase (dizionari JSON); quelle vecchie restano vuote
const ENVV = { SUPABASE_URL: "https://db.test", SUPABASE_SECRET_KEYS: '{"default":"sb_secret_x"}', SUPABASE_PUBLISHABLE_KEYS: '{"default":"sb_publishable_y"}', META_TOKEN: "tok", META_APP_ID: "app1", META_APP_SECRET: "segreto" };
let gestore = null;
const inSottofondo = [];
globalThis.Deno = { env: { get: (k) => ENVV[k] }, serve: (f) => { gestore = f; } };
globalThis.EdgeRuntime = { waitUntil: (p) => inSottofondo.push(p) };
const sorgente = readFileSync(new URL("../supabase/functions/meta-sync/index.ts", import.meta.url), "utf8");
await import("data:text/javascript;base64," + Buffer.from(sorgente).toString("base64"));
assert.ok(gestore, "la funzione registra il gestore");

// ---------- database finto (solo ciò che usa la funzione) ----------
const tabelle = { meta_account: [], meta_collegamenti: [], meta_oggetti: [], meta_finestre: [], meta_sync: [], meta_token: [], meta_login: [] };
const CHIAVI = { meta_account: ["workspace_id", "account_id"], meta_oggetti: ["workspace_id", "livello", "id"], meta_finestre: ["workspace_id", "campagna_id", "inizio"], meta_token: ["workspace_id"], meta_login: ["stato"] };
let prossimoId = 1;
const filtri = (q) => [...q.entries()].filter(([k]) => !["select", "on_conflict"].includes(k)).map(([k, v]) => {
  const [op, ...resto] = v.split("."); const val = resto.join(".");
  if (op === "eq") return (r) => String(r[k]) === val;
  if (op === "gt") return (r) => String(r[k]) > val;
  if (op === "lt") return (r) => String(r[k]) < val;
  if (op === "in") { const l = val.slice(1, -1).split(",").map((x) => x.replace(/"/g, "")); return (r) => l.includes(String(r[k])); }
  throw new Error("filtro " + v);
});
function dbFinto(url, init) {
  const u = new URL(url); const percorso = u.pathname.replace("/rest/v1/", ""); const metodo = init.method || "GET";
  const corpo = init.body ? JSON.parse(init.body) : undefined;
  if (percorso === "rpc/meta_cron_chiave") return "chiave-cron";
  if (percorso === "rpc/my_workspace") return "ws1";
  const t = tabelle[percorso]; if (!t) throw new Error("tabella " + percorso);
  const f = filtri(u.searchParams); const passa = (r) => f.every((g) => g(r));
  if (metodo === "GET") return t.filter(passa);
  if (metodo === "PATCH") { t.filter(passa).forEach((r) => Object.assign(r, corpo)); return null; }
  if (metodo === "DELETE") { tabelle[percorso] = t.filter((r) => !passa(r)); return null; }
  if (metodo === "POST") {
    const righe = Array.isArray(corpo) ? corpo : [corpo];
    const out = [];
    for (const r of righe) {
      const ch = CHIAVI[percorso];
      const esist = ch && t.find((x) => ch.every((k) => String(x[k]) === String(r[k])));
      if (esist) Object.assign(esist, r);
      else { const n = percorso === "meta_sync" ? { id: prossimoId++, avvio: new Date().toISOString(), esito: "in_corso", dettagli: {}, messaggio: "", ...r } : percorso === "meta_login" ? { creato: new Date().toISOString(), ...r } : { ...r }; t.push(n); out.push(n); }
    }
    return out;
  }
}

// ---------- Meta finto: un account in USD, 3 campagne ----------
// c1 attiva da 20 giorni (CBO 50 USD), c2 in pausa con spesa recente (ABO), c3 vecchia: va ignorata
const giorni = (n) => new Date(Date.now() - n * 864e5).toISOString().slice(0, 10);
const azioni = (lead, acquisti) => [{ action_type: "lead", value: String(lead) }, { action_type: "omni_purchase", value: String(acquisti) }, { action_type: "link_click", value: "999" }];
const ins = (spend, lead = 0, acq = 0, valore = 0) => ({ spend: String(spend), impressions: "1000", clicks: "50", inline_link_clicks: "30", reach: "800", frequency: "1.25", actions: azioni(lead, acq), action_values: [{ action_type: "omni_purchase", value: String(valore) }] });
const chiamate = [];
let tokenAtteso = "tok";
function metaFinto(url) {
  const u = new URL(url); const p = u.pathname.replace(/^\/v[\d.]+\//, ""); const q = Object.fromEntries(u.searchParams);
  chiamate.push(p + (q.level ? `?level=${q.level}` : "") + (q.date_preset ? `&${q.date_preset}` : ""));
  if (p === "oauth/access_token") {
    assert.equal(q.client_secret, "segreto"); assert.ok(!q.access_token, "niente token vuoto");
    if (q.code) { assert.equal(q.code, "codice-fb"); assert.equal(q.redirect_uri, "https://db.test/functions/v1/meta-sync"); return { access_token: "breve" }; }
    assert.equal(q.fb_exchange_token, "breve"); return { access_token: "tok-lungo", expires_in: 5184000 };
  }
  if (p === "me") { assert.equal(q.access_token, "tok-lungo"); return { id: "9", name: "Nicola Ciappa" }; }
  assert.equal(q.access_token, tokenAtteso);
  if (p === "me/adaccounts") return { data: [{ account_id: "555", name: "BDF 1", currency: "USD", account_status: 1, timezone_name: "America/New_York", business: { name: "Boutique Del Fitness" } }, { account_id: "777", name: "Altro", currency: "EUR", account_status: 101, timezone_name: "Europe/Rome" }] };
  if (p === "act_555/campaigns") return { data: [
    { id: "c1", name: "Lead fitness", effective_status: "ACTIVE", objective: "OUTCOME_LEADS", daily_budget: "5000", start_time: giorni(20) + "T10:00:00-0400" },
    { id: "c2", name: "Vendite corsi", effective_status: "PAUSED", objective: "OUTCOME_SALES", start_time: giorni(40) + "T10:00:00-0400" },
    { id: "c3", name: "Vecchia", effective_status: "PAUSED", objective: "OUTCOME_TRAFFIC", start_time: "2023-01-01T10:00:00-0400" },
  ], paging: {} };
  if (p === "act_555/insights" && q.date_preset === "last_90d") return { data: [{ campaign_id: "c1", campaign_name: "Lead fitness", spend: "300" }, { campaign_id: "c2", campaign_name: "Vendite corsi", spend: "120" }, { campaign_id: "c9", campaign_name: "Archiviata", spend: "40" }] };
  if (p === "act_555/insights" && q.date_preset === "maximum") {
    const ids = JSON.parse(q.filtering)[0].value;
    assert.ok(!ids.includes("c3"), "la campagna vecchia non viene scaricata");
    if (q.level === "campaign") return { data: [{ campaign_id: "c1", campaign_name: "Lead fitness", ...ins(300, 25) }, { campaign_id: "c2", campaign_name: "Vendite corsi", ...ins(500, 0, 7, 980) }, { campaign_id: "c9", campaign_name: "Archiviata", ...ins(40) }] };
    if (q.level === "adset") return { data: [{ campaign_id: "c1", adset_id: "s1", adset_name: "Donne 25-45", ...ins(300, 25) }, { campaign_id: "c2", adset_id: "s2", adset_name: "Retargeting", ...ins(500, 0, 7, 980) }] };
    if (q.level === "ad") return { data: [{ campaign_id: "c1", adset_id: "s1", ad_id: "a1", ad_name: "Video 1", ...ins(200, 18) }, { campaign_id: "c1", adset_id: "s1", ad_id: "a2", ad_name: "Video 2", ...ins(100, 7) }] };
  }
  if (p === "act_555/adsets") return { data: [{ id: "s1", name: "Donne 25-45", campaign_id: "c1", effective_status: "ACTIVE" }, { id: "s2", name: "Retargeting", campaign_id: "c2", effective_status: "CAMPAIGN_PAUSED", daily_budget: "2000" }, { id: "s3", name: "Mai partito", campaign_id: "c2", effective_status: "PAUSED" }] };
  if (p === "act_555/ads") return { data: [{ id: "a1", name: "Video 1", campaign_id: "c1", adset_id: "s1", effective_status: "ACTIVE" }, { id: "a2", name: "Video 2", campaign_id: "c1", adset_id: "s1", effective_status: "PAUSED" }, { id: "a3", name: "Bozza", campaign_id: "c1", adset_id: "s1", effective_status: "PAUSED" }] };
  const m = p.match(/^(c\d)\/insights$/);
  if (m) { const r = JSON.parse(q.time_ranges); return { data: r.map((x, i) => ({ date_start: x.since, date_stop: x.until, ...ins(10 + i, i, 0, 0) })) }; }
  throw new Error("Meta non simulato: " + p);
}

globalThis.fetch = async (url, init = {}) => {
  const s = String(url);
  const json = (corpo, stato = 200) => ({ ok: stato < 400, status: stato, json: async () => corpo, text: async () => (corpo === null ? "" : JSON.stringify(corpo)) });
  if (s.startsWith("https://graph.facebook.com/")) return json(metaFinto(s));
  if (s.startsWith("https://db.test/rest/v1/")) return json(dbFinto(s, init));
  if (s.startsWith("https://db.test/rest/v1/")) assert.equal(init.headers.apikey, init.headers.Authorization ? "sb_publishable_y" : "sb_secret_x");
  if (s === "https://db.test/auth/v1/user") { assert.equal(init.headers.apikey, "sb_publishable_y"); return json({ id: "u1" }, init.headers.Authorization === "Bearer utente" ? 200 : 401); }
  if (s === "https://db.test/functions/v1/meta-sync") { // la funzione che si richiama
    const r = await gestore(new Request(s, { method: "POST", headers: init.headers, body: init.body }));
    return json(await r.json(), r.status);
  }
  throw new Error("rete non simulata: " + s);
};
const chiama = async (corpo, headers) => { const r = await gestore(new Request("https://db.test/functions/v1/meta-sync", { method: "POST", headers: { "Content-Type": "application/json", ...headers }, body: JSON.stringify(corpo) })); return { stato: r.status, corpo: await r.json() }; };
const finisci = async () => { while (inSottofondo.length) await inSottofondo.shift(); };

// ---------- 1. senza accesso: rifiutato ----------
assert.equal((await chiama({}, { Authorization: "Bearer falso" })).stato, 401);
assert.equal((await chiama({}, { "x-cron-key": "sbagliata" })).stato, 401);

// ---------- 2. elenco account dal CRM ----------
let r = await chiama({ azione: "account" }, { Authorization: "Bearer utente" });
assert.equal(r.corpo.account, 2);
assert.equal(tabelle.meta_account.find((a) => a.account_id === "555").valuta, "USD");

// ---------- 3. Aggiorna ora senza collegamenti ----------
r = await chiama({}, { Authorization: "Bearer utente" }); await finisci();
assert.equal(tabelle.meta_sync.at(-1).esito, "ok");
assert.match(tabelle.meta_sync.at(-1).messaggio, /Nessun account collegato/);

// ---------- 4. collego BDF 1 e aggiorno ----------
tabelle.meta_collegamenti.push({ workspace_id: "ws1", account_id: "555", cliente_id: "z1" });
r = await chiama({}, { Authorization: "Bearer utente" }); await finisci();
const sync = tabelle.meta_sync.at(-1);
assert.equal(sync.esito, "ok", sync.messaggio);
const camp = (id) => tabelle.meta_oggetti.find((o) => o.livello === "campaign" && o.id === id);
assert.ok(camp("c1") && camp("c2") && camp("c9"), "attiva, in pausa con spesa e archiviata con spesa");
assert.ok(!camp("c3"), "la campagna vecchia resta fuori");
assert.equal(camp("c1").budget_giorno, 50, "CBO: 5000 centesimi = 50 USD");
assert.equal(camp("c2").budget_giorno, 0, "ABO con ad set non attivo: budget giornaliero 0");
assert.equal(camp("c1").lead, 25); assert.equal(camp("c1").conversioni, 0, "il lead non è una conversione");
assert.equal(camp("c2").conversioni, 7); assert.equal(camp("c2").ricavi, 980);
assert.equal(camp("c1").click_link, 30); assert.equal(camp("c1").frequenza, 1.25);
assert.equal(camp("c1").inizio, giorni(20));
assert.ok(tabelle.meta_oggetti.find((o) => o.id === "s1") && !tabelle.meta_oggetti.find((o) => o.id === "s3"), "ad set mai partito escluso");
assert.ok(tabelle.meta_oggetti.find((o) => o.id === "a2") && !tabelle.meta_oggetti.find((o) => o.id === "a3"), "annuncio in bozza escluso, in pausa con spesa incluso");
const fin = (c) => tabelle.meta_finestre.filter((f) => f.campagna_id === c).sort((a, b) => (a.inizio < b.inizio ? -1 : 1));
const f1 = fin("c1");
assert.equal(f1[0].tipo, "test"); assert.equal(f1[0].inizio, giorni(20)); assert.equal(f1[0].fine, giorni(18));
assert.equal(f1[1].tipo, "matura"); assert.equal(f1[1].inizio, giorni(17)); assert.equal(f1[1].fine, giorni(11));
assert.ok(f1.slice(2).every((f) => f.tipo === "attiva"));
const ieriNY = new Intl.DateTimeFormat("en-CA", { timeZone: "America/New_York" }).format(new Date(Date.now() - 864e5));
assert.equal(f1.at(-1).fine, ieriNY, "l'ultima finestra si ferma a ieri nel fuso dell'account");
assert.ok(f1[0].spesa > 0);
assert.equal(tabelle.meta_account.find((a) => a.account_id === "555").ultimo_errore, "");

// ---------- 5. secondo giro: riscarica solo le finestre recenti ----------
const prima = chiamate.length;
r = await chiama({}, { Authorization: "Bearer utente" }); await finisci();
const nuove = chiamate.slice(prima).filter((c) => /^c\d\/insights/.test(c)).length;
assert.ok(nuove <= 2, `finestre: ${nuove} chiamate, ne bastano una per campagna`);
assert.equal(tabelle.meta_sync.at(-1).esito, "ok");

// ---------- 6. cron: fuori dalle 9 non fa niente, con forza lavora ----------
r = await chiama({ origine: "programmato" }, { "x-cron-key": "chiave-cron" });
const ora = Number(new Intl.DateTimeFormat("en-GB", { timeZone: "Europe/Rome", hour: "2-digit", hourCycle: "h23" }).format(new Date()));
if (ora !== 9) assert.equal(r.corpo.saltato, "non sono le 9 a Roma");
tabelle.meta_account.forEach((a) => { a.ultimo_sync = null; });
r = await chiama({ origine: "programmato", forza: true }, { "x-cron-key": "chiave-cron" }); await finisci();
assert.equal(r.stato, 202); assert.equal(tabelle.meta_sync.at(-1).origine, "programmato"); assert.equal(tabelle.meta_sync.at(-1).esito, "ok");

// ---------- 7. errore di Meta su un account: registrato, il giro chiude ----------
const metaVero = metaFinto;
tabelle.meta_collegamenti.push({ workspace_id: "ws1", account_id: "777", cliente_id: "z2" });
r = await chiama({}, { Authorization: "Bearer utente" }); await finisci();
const ultimo = tabelle.meta_sync.at(-1);
assert.equal(ultimo.esito, "parziale", ultimo.messaggio);
assert.match(tabelle.meta_account.find((a) => a.account_id === "777").ultimo_errore, /Meta non simulato/);

// ---------- 8. tanti account: a fine tempo si richiama e finisce il lavoro ----------
// Ogni account aggiornato fa scorrere l'orologio di 60 secondi: ne entrano due per giro
const dataVera = Date.now; let salto = 0; Date.now = () => dataVera() + salto;
const metaPrima = metaFinto;
tabelle.meta_collegamenti = tabelle.meta_collegamenti.filter((c) => c.account_id !== "777");
for (const n of ["556", "557", "558", "559"]) {
  tabelle.meta_account.push({ workspace_id: "ws1", account_id: n, nome: "Acc " + n, valuta: "EUR", fuso: "Europe/Rome", stato: 1 });
  tabelle.meta_collegamenti.push({ workspace_id: "ws1", account_id: n, cliente_id: "z" + n });
}
const fetchPrima = globalThis.fetch;
const perAccount = {};
globalThis.fetch = async (url, init = {}) => {
  const s = String(url);
  const m = s.match(/graph\.facebook\.com\/[^/]+\/act_(\d+)\/campaigns/);
  if (m && m[1] !== "555") { perAccount[m[1]] = (perAccount[m[1]] || 0) + 1; salto += 60000; return { ok: true, status: 200, json: async () => ({ data: [] }) }; }
  if (/act_55[6-9]\/insights/.test(s)) return { ok: true, status: 200, json: async () => ({ data: [] }) };
  if (s.includes("me/adaccounts")) return { ok: true, status: 200, json: async () => ({ data: [] }) };
  return fetchPrima(url, init);
};
r = await chiama({}, { Authorization: "Bearer utente" });
await finisci();
const s8 = tabelle.meta_sync.at(-1);
assert.equal(s8.esito, "ok", s8.messaggio);
assert.ok(s8.dettagli.giri >= 2, `si è richiamata (${s8.dettagli.giri} giri)`);
assert.deepEqual(Object.keys(perAccount).sort(), ["556", "557", "558", "559"], "tutti gli account passati");
assert.ok(Object.values(perAccount).every((n) => n === 1), "nessun account ripetuto");
globalThis.fetch = fetchPrima; Date.now = dataVera;

// ---------- 9. Collega Facebook: finestra, ritorno, token lungo salvato e usato ----------
r = await chiama({ azione: "login", ritorno: "https://evil.example.com/" }, { Authorization: "Bearer utente" });
assert.equal(r.stato, 400, "ritorno verso un altro sito: rifiutato");
r = await chiama({ azione: "login", ritorno: "https://crm-nc-media.attaccante.workers.dev/" }, { Authorization: "Bearer utente" });
assert.equal(r.stato, 400, "stesso nome su un altro account Cloudflare: rifiutato");
r = await chiama({ azione: "login", ritorno: "https://claude-x-crm-nc-media.nciappa9.workers.dev/" }, { Authorization: "Bearer utente" });
assert.equal(r.stato, 200, JSON.stringify(r.corpo));
const dialog = new URL(r.corpo.url);
assert.equal(dialog.hostname, "www.facebook.com"); assert.equal(dialog.searchParams.get("client_id"), "app1");
assert.equal(dialog.searchParams.get("scope"), "ads_read,business_management");
const stato = dialog.searchParams.get("state");
assert.ok(stato.length >= 32 && tabelle.meta_login.some((x) => x.stato === stato && x.workspace_id === "ws1"));
// state inventato: rifiutato
let cb = await gestore(new Request("https://db.test/functions/v1/meta-sync?code=x&state=falso"));
assert.equal(cb.status, 400);
// annullato dall'utente
r = await chiama({ azione: "login", ritorno: "http://localhost:5173/" }, { Authorization: "Bearer utente" });
const stato2 = new URL(r.corpo.url).searchParams.get("state");
cb = await gestore(new Request(`https://db.test/functions/v1/meta-sync?error=access_denied&error_description=Annullato&state=${stato2}`));
assert.equal(cb.status, 302); assert.match(cb.headers.get("Location"), /meta=errore/);
// ritorno buono
tokenAtteso = "tok-lungo";
cb = await gestore(new Request(`https://db.test/functions/v1/meta-sync?code=codice-fb&state=${stato}`));
assert.equal(cb.status, 302, await cb.clone().text());
assert.equal(cb.headers.get("Location"), "https://claude-x-crm-nc-media.nciappa9.workers.dev/?meta=collegato");
const tk = tabelle.meta_token.find((x) => x.workspace_id === "ws1");
assert.equal(tk.token, "tok-lungo"); assert.equal(tk.utente_meta, "Nicola Ciappa");
assert.ok(new Date(tk.scadenza) - Date.now() > 59 * 864e5, "scade fra 60 giorni");
assert.ok(!tabelle.meta_login.some((x) => x.stato === stato), "lo state si usa una volta sola");
cb = await gestore(new Request(`https://db.test/functions/v1/meta-sync?code=codice-fb&state=${stato}`));
assert.equal(cb.status, 400, "riusare lo stesso ritorno non funziona");
// l'aggiornamento ora usa il token dello spazio di lavoro (metaFinto controlla access_token)
tabelle.meta_collegamenti = [{ workspace_id: "ws1", account_id: "555", cliente_id: "z1" }];
r = await chiama({}, { Authorization: "Bearer utente" }); await finisci();
assert.equal(tabelle.meta_sync.at(-1).esito, "ok", tabelle.meta_sync.at(-1).messaggio);
// token scaduto: messaggio chiaro
const fetchOk = globalThis.fetch;
globalThis.fetch = async (url, init) => (String(url).includes("graph.facebook.com") && !String(url).includes("oauth")
  ? { ok: false, status: 400, json: async () => ({ error: { message: "Error validating access token", code: 190 } }) } : fetchOk(url, init));
r = await chiama({}, { Authorization: "Bearer utente" }); await finisci();
assert.equal(tabelle.meta_sync.at(-1).esito, "errore"); assert.match(tabelle.meta_sync.at(-1).messaggio, /scaduto/);
globalThis.fetch = fetchOk;
// scollega
r = await chiama({ azione: "scollega" }, { Authorization: "Bearer utente" });
assert.ok(!tabelle.meta_token.length);

// ---------- 10. proseguimento bloccato (Verify JWT acceso), ripresa e "continua" dal CRM ----------
tabelle.meta_token.push({ workspace_id: "ws1", token: "tok-lungo" });
tokenAtteso = "tok-lungo";
const dataVera2 = Date.now; let salto2 = 0; Date.now = () => dataVera2() + salto2;
tabelle.meta_collegamenti = [];
tabelle.meta_account = tabelle.meta_account.filter((a) => !/^6\d\d$/.test(a.account_id));
for (const n of ["601", "602", "603", "604", "605"]) {
  tabelle.meta_account.push({ workspace_id: "ws1", account_id: n, nome: "Acc " + n, valuta: "EUR", fuso: "Europe/Rome", stato: 1, ultimo_sync: null });
  tabelle.meta_collegamenti.push({ workspace_id: "ws1", account_id: n, cliente_id: "z" + n });
}
const giri10 = {};
const fetch10 = globalThis.fetch;
let bloccaRichiamo = true;
globalThis.fetch = async (url, init = {}) => {
  const s = String(url);
  const m = s.match(/act_(6\d\d)\/campaigns/);
  if (m) { giri10[m[1]] = (giri10[m[1]] || 0) + 1; salto2 += 60000; return { ok: true, status: 200, json: async () => ({ data: [] }) }; }
  if (/act_6\d\d\/insights/.test(s) || s.includes("me/adaccounts")) return { ok: true, status: 200, json: async () => ({ data: [] }) };
  if (s === "https://db.test/functions/v1/meta-sync" && bloccaRichiamo) return { ok: false, status: 401, json: async () => ({}), text: async () => "" };
  return fetch10(url, init);
};
r = await chiama({}, { Authorization: "Bearer utente" }); await finisci();
let s10 = tabelle.meta_sync.at(-1);
assert.equal(s10.esito, "parziale", s10.messaggio);
assert.match(s10.messaggio, /Verify JWT/);
const primi = Object.keys(giri10);
assert.equal(primi.length, 2, "nel primo giro ne entrano due");
// nuovo Aggiorna ora: riparte dagli account mai aggiornati, non dagli stessi
salto2 += 11 * 60e3;
r = await chiama({}, { Authorization: "Bearer utente" }); await finisci();
const secondi = Object.keys(giri10).filter((k) => !primi.includes(k));
assert.equal(secondi.length, 2, "il secondo giro prende due account nuovi: " + JSON.stringify(giri10));
// "continua" dal CRM su un aggiornamento fermo
bloccaRichiamo = false;
const fermo = { id: 999, workspace_id: "ws1", origine: "manuale", avvio: new Date(Date.now()).toISOString(), esito: "in_corso", messaggio: "In corso", dettagli: { account: {}, giri: 1, battito: new Date(Date.now() - 120e3).toISOString() } };
tabelle.meta_sync.push(fermo);
r = await chiama({ azione: "continua", sync: 999 }, { Authorization: "Bearer utente" });
assert.equal(r.stato, 202, JSON.stringify(r.corpo)); await finisci();
assert.equal(fermo.esito, "ok", fermo.messaggio);
// un account rimasto "in lavorazione" (esecuzione interrotta lì) viene saltato con un errore, gli altri proseguono
const interrotto = { id: 998, workspace_id: "ws1", origine: "manuale", avvio: new Date(Date.now()).toISOString(), esito: "in_corso", messaggio: "",
  dettagli: { giri: 1, battito: new Date(Date.now() - 120e3).toISOString(), account: { "601": { nome: "Acc 601", inLavorazione: true } } } };
tabelle.meta_sync.push(interrotto);
const prima601 = giri10["601"] || 0;
r = await chiama({ azione: "continua", sync: 998 }, { Authorization: "Bearer utente" }); await finisci();
assert.equal(giri10["601"] || 0, prima601, "l'account interrotto non si riprova nello stesso aggiornamento");
assert.match(interrotto.dettagli.account["601"].errore, /pesante/);
assert.equal(interrotto.esito, "parziale", interrotto.messaggio);
assert.ok(["602", "603", "604", "605"].every((k) => interrotto.dettagli.account[k] && !interrotto.dettagli.account[k].inLavorazione), "gli altri completati");
// con battito recente non riparte (il lavoro sta già andando)
const vivo = { id: 1000, workspace_id: "ws1", origine: "manuale", avvio: new Date(Date.now()).toISOString(), esito: "in_corso", messaggio: "", dettagli: { battito: new Date(Date.now()).toISOString() } };
tabelle.meta_sync.push(vivo);
r = await chiama({ azione: "continua", sync: 1000 }, { Authorization: "Bearer utente" });
assert.equal(r.corpo.inCorso, true);
globalThis.fetch = fetch10; Date.now = dataVera2;

console.log("meta-sync: tutte le prove passate ✔", { chiamateMeta: chiamate.length, oggetti: tabelle.meta_oggetti.length, finestre: tabelle.meta_finestre.length, ultimo: ultimo.messaggio });
