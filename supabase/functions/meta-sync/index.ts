// meta-sync — scarica da Meta Ads i dati degli account collegati ai clienti del CRM.
//
// Tre modi di avvio:
//  · dal cron delle 9 (intestazione x-cron-key, vedi migrazione 0002): tutti gli spazi di lavoro;
//  · dal CRM (utente collegato): "Collega Facebook", "Aggiorna ora", "Aggiorna l'elenco degli account";
//  · da Facebook (GET con code e state): il ritorno dalla finestra di accesso.
// Risponde subito e lavora in sottofondo. Se gli account sono tanti, a fine tempo si
// richiama da sola e prosegue da dove era arrivata.
//
// Scrive in meta_account, meta_oggetti (totali di campagne, ad set, annunci),
// meta_finestre (Test 3 giorni, Matura 7, poi settimane) e meta_sync (registro).
// Accesso a Meta: il token del profilo Facebook di chi preme "Collega Facebook" (tabella meta_token,
// migrazione 0003), rinnovato ogni volta che si ricollega. Segreti delle Edge Functions:
// META_APP_ID, META_APP_SECRET e, se l'app usa Facebook Login for Business, META_CONFIG_ID.
// META_TOKEN resta come riserva (per esempio un token di utente di sistema).
//
// Il file è un unico modulo senza dipendenze, così si può incollare nell'editor di Supabase.

const VERSIONE = Deno.env.get("META_API_VERSION") || "v24.0";
const GRAPH = `https://graph.facebook.com/${VERSIONE}`;
const TEST_1 = 3, TEST_2 = 10, FINESTRA = 7; // come nel CRM
const TEMPO_MAX = 110_000; // ms di lavoro per invocazione, poi si richiama
const GIORNI_RICALCOLO = 8; // Meta ritocca i giorni recenti: le finestre che li toccano si riscaricano
const VALUTE_SENZA_CENTESIMI = ["CLP", "COP", "CRC", "HUF", "ISK", "IDR", "JPY", "KRW", "PYG", "TWD", "VND"];

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const risposta = (corpo, stato = 200) =>
  new Response(JSON.stringify(corpo), { status: stato, headers: { ...CORS, "Content-Type": "application/json" } });

/* Chiavi di Supabase: le nuove (SUPABASE_SECRET_KEYS / SUPABASE_PUBLISHABLE_KEYS, dizionari JSON)
   o, se presenti, quelle vecchie ormai deprecate */
const primaChiave = (nome) => {
  try { const d = JSON.parse(Deno.env.get(nome) || "{}"); return d.default || Object.values(d)[0] || ""; } catch { return ""; }
};
const ENV = () => ({
  url: Deno.env.get("SUPABASE_URL") || "",
  servizio: primaChiave("SUPABASE_SECRET_KEYS") || Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") || "",
  anon: primaChiave("SUPABASE_PUBLISHABLE_KEYS") || Deno.env.get("SUPABASE_ANON_KEY") || "",
  token: Deno.env.get("META_TOKEN") || "",
  appId: Deno.env.get("META_APP_ID") || "",
  appSecret: Deno.env.get("META_APP_SECRET") || "",
  configId: Deno.env.get("META_CONFIG_ID") || "",
  origini: (Deno.env.get("CRM_ORIGINI") || "").split(",").map((x) => x.trim()).filter(Boolean),
});

// ---------- database (PostgREST con la chiave di servizio) ----------
const intestazioniServizio = () => {
  const k = ENV().servizio;
  // Le chiavi nuove (sb_secret_…) vanno solo in apikey; quelle JWT anche in Authorization
  return k.startsWith("sb_") ? { apikey: k } : { apikey: k, Authorization: `Bearer ${k}` };
};
async function db(percorso, { metodo = "GET", corpo, prefer, auth } = {}) {
  const h = { "Content-Type": "application/json", ...(auth || intestazioniServizio()) };
  if (prefer) h.Prefer = prefer;
  const r = await fetch(`${ENV().url}/rest/v1/${percorso}`, { method: metodo, headers: h, body: corpo === undefined ? undefined : JSON.stringify(corpo) });
  const testo = await r.text();
  if (!r.ok) throw new Error(`database ${r.status}: ${testo.slice(0, 300)}`);
  return testo ? JSON.parse(testo) : null;
}
const upsert = async (tabella, conflitto, righe) => {
  for (let i = 0; i < righe.length; i += 500) {
    await db(`${tabella}?on_conflict=${conflitto}`, { metodo: "POST", corpo: righe.slice(i, i + 500), prefer: "resolution=merge-duplicates,return=minimal" });
  }
};
const lista = (v) => `(${v.map((x) => `"${String(x).replace(/"/g, "")}"`).join(",")})`;

// ---------- Meta ----------
class ErroreMeta extends Error {
  codice;
  constructor(messaggio, codice) { super(messaggio); this.codice = codice; }
  get limite() { return [4, 17, 32, 613].includes(this.codice) || (this.codice >= 80000 && this.codice <= 80014); }
  get scaduto() { return this.codice === 190 || this.codice === 102; }
}
async function graph(percorso, parametri = {}, token) {
  const tutti = [];
  const q = new URLSearchParams(parametri);
  if (token) q.set("access_token", token);
  let url = `${GRAPH}/${percorso}?${q}`;
  for (let pagine = 0; url && pagine < 200; pagine++) {
    const r = await fetch(url);
    const j = await r.json().catch(() => ({}));
    if (!r.ok || j.error) {
      const e = j.error || {};
      throw new ErroreMeta(e.error_user_msg || e.message || `Meta ha risposto ${r.status}`, e.code);
    }
    if (!Array.isArray(j.data)) return j; // oggetto singolo
    tutti.push(...j.data);
    url = j.paging && j.paging.next ? j.paging.next : "";
  }
  return tutti;
}

// ---------- date ----------
const dataNelFuso = (fuso, spostamento = 0) => {
  const d = new Date(Date.now() + spostamento * 864e5);
  try { return new Intl.DateTimeFormat("en-CA", { timeZone: fuso || "Europe/Rome", year: "numeric", month: "2-digit", day: "2-digit" }).format(d); }
  catch { return new Intl.DateTimeFormat("en-CA", { timeZone: "Europe/Rome", year: "numeric", month: "2-digit", day: "2-digit" }).format(d); }
};
const oraRoma = () => Number(new Intl.DateTimeFormat("en-GB", { timeZone: "Europe/Rome", hour: "2-digit", hourCycle: "h23" }).format(new Date()));
const sommaGiorni = (s, n) => { const d = new Date(s + "T00:00:00Z"); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); };
const minimo = (a, b) => (a < b ? a : b);

/* Le finestre del CRM: 3 giorni di Test, 7 di Matura, poi settimane. L'ultima si ferma a ieri. */
function finestre(inizio, ieri) {
  if (!inizio || inizio > ieri) return [];
  const out = [];
  const agg = (tipo, dal, al) => { if (dal <= ieri) out.push({ tipo, inizio: dal, fine: minimo(al, ieri) }); };
  agg("test", inizio, sommaGiorni(inizio, TEST_1 - 1));
  agg("matura", sommaGiorni(inizio, TEST_1), sommaGiorni(inizio, TEST_2 - 1));
  let dal = sommaGiorni(inizio, TEST_2);
  for (let g = 0; dal <= ieri && g < 400; g++) { agg("attiva", dal, sommaGiorni(dal, FINESTRA - 1)); dal = sommaGiorni(dal, FINESTRA); }
  return out;
}

// ---------- numeri ----------
const num = (x) => { const n = parseFloat(x); return isFinite(n) ? n : 0; };
const azione = (elenco, tipi) => {
  const m = {};
  (elenco || []).forEach((a) => { m[a.action_type] = num(a.value); });
  for (const t of tipi) if (m[t] !== undefined) return m[t];
  return 0;
};
/* Lead: "lead" è già la somma di modulo e sito; se manca sommo le voci separate.
   Conversioni e ricavi: gli acquisti. Il lead non è una conversione (regola del CRM). */
function metriche(r) {
  const az = r.actions || [];
  let lead = azione(az, ["lead"]);
  if (!lead) lead = azione(az, ["onsite_conversion.lead_grouped"]) + azione(az, ["offsite_conversion.fb_pixel_lead"]);
  return {
    spesa: num(r.spend), impression: Math.round(num(r.impressions)), click: Math.round(num(r.clicks)),
    click_link: Math.round(num(r.inline_link_clicks)), reach: Math.round(num(r.reach)), frequenza: num(r.frequency),
    lead, conversioni: azione(az, ["omni_purchase", "purchase", "offsite_conversion.fb_pixel_purchase"]),
    ricavi: azione(r.action_values, ["omni_purchase", "purchase", "offsite_conversion.fb_pixel_purchase"]),
  };
}
const CAMPI_INSIGHTS = "spend,impressions,clicks,inline_link_clicks,reach,frequency,actions,action_values";
const aBudget = (x, valuta) => num(x) / (VALUTE_SENZA_CENTESIMI.includes(valuta) ? 1 : 100);
const aData = (s) => (s ? String(s).slice(0, 10) : null);
const pezzi = (a, n) => { const o = []; for (let i = 0; i < a.length; i += n) o.push(a.slice(i, i + n)); return o; };
const filtroCampagne = (ids) => JSON.stringify([{ field: "campaign.id", operator: "IN", value: ids }]);

// ---------- elenco account ----------
async function aggiornaAccount(ws, token) {
  const acc = await graph("me/adaccounts", { fields: "account_id,name,currency,account_status,timezone_name,business{name}", limit: "200" }, token);
  await upsert("meta_account", "workspace_id,account_id", acc.map((a) => ({
    workspace_id: ws, account_id: String(a.account_id), nome: a.name || "", valuta: a.currency || "",
    stato: a.account_status ?? null, fuso: a.timezone_name || "", business: (a.business && a.business.name) || "",
    aggiornato: new Date().toISOString(),
  })));
  return acc.length;
}

// ---------- un account ----------
async function sincronizzaAccount(ws, acc, token) {
  const g = (percorso, parametri) => graph(percorso, parametri, token);
  const id = acc.account_id, act = `act_${id}`, valuta = acc.valuta;
  const ieri = dataNelFuso(acc.fuso, -1);
  const ora = new Date().toISOString();

  // Campagne da seguire: attive, o con spesa negli ultimi 90 giorni. Le vecchie restano fuori.
  const campagne = await g(`${act}/campaigns`, { fields: "id,name,effective_status,objective,daily_budget,lifetime_budget,start_time,stop_time,created_time", limit: "500" });
  const recenti = await g(`${act}/insights`, { level: "campaign", date_preset: "last_90d", fields: "campaign_id,campaign_name,spend", limit: "500" });
  const conSpesa = new Map(recenti.filter((r) => num(r.spend) > 0).map((r) => [r.campaign_id, r.campaign_name]));
  const seguite = campagne.filter((c) => c.effective_status === "ACTIVE" || conSpesa.has(c.id));
  // Archiviate con spesa recente: non escono dall'elenco delle campagne, le ricostruisco dagli insights
  conSpesa.forEach((nome, cid) => { if (!seguite.some((c) => c.id === cid)) seguite.push({ id: cid, name: nome, effective_status: "ARCHIVED" }); });
  if (!seguite.length) return { campagne: 0, adset: 0, annunci: 0, finestre: 0 };
  const ids = seguite.map((c) => c.id);

  // Totali dall'inizio e stati, per campagna, ad set e annuncio
  const totali = { campaign: new Map(), adset: new Map(), ad: new Map() };
  const adsets = [], ads = [];
  for (const blocco of pezzi(ids, 50)) {
    const filtro = filtroCampagne(blocco);
    for (const [livello, chiave] of [["campaign", "campaign_id"], ["adset", "adset_id"], ["ad", "ad_id"]]) {
      const righe = await g(`${act}/insights`, { level: livello, date_preset: "maximum", filtering: filtro, limit: "500",
        fields: `campaign_id,campaign_name,adset_id,adset_name,ad_id,ad_name,${CAMPI_INSIGHTS}` });
      righe.forEach((r) => totali[livello].set(r[chiave], r));
    }
    adsets.push(...await g(`${act}/adsets`, { fields: "id,name,campaign_id,effective_status,daily_budget,lifetime_budget,start_time,end_time", filtering: filtro, limit: "500" }));
    ads.push(...await g(`${act}/ads`, { fields: "id,name,campaign_id,adset_id,effective_status", filtering: filtro, limit: "500" }));
  }

  const vuote = { spesa: 0, impression: 0, click: 0, click_link: 0, reach: 0, frequenza: 0, lead: 0, conversioni: 0, ricavi: 0 };
  const riga = (livello, o, extra) => ({
    workspace_id: ws, livello, id: o.id, account_id: id, campagna_id: "", adset_id: "", nome: o.name || "",
    stato: o.effective_status || "", obiettivo: "", budget_giorno: 0, budget_totale: 0, inizio: null, fine: null,
    ...vuote, ...(totali[livello].has(o.id) ? metriche(totali[livello].get(o.id)) : {}), aggiornato: ora, ...extra,
  });
  const attivi = (x) => x.effective_status === "ACTIVE";
  const righe = [];
  for (const c of seguite) {
    // Budget di campagna (CBO) o somma degli ad set attivi (ABO)
    const suoi = adsets.filter((a) => a.campaign_id === c.id);
    const giorno = c.daily_budget ? aBudget(c.daily_budget, valuta) : suoi.filter(attivi).reduce((s, a) => s + aBudget(a.daily_budget, valuta), 0);
    const totale = c.lifetime_budget ? aBudget(c.lifetime_budget, valuta) : suoi.filter(attivi).reduce((s, a) => s + aBudget(a.lifetime_budget, valuta), 0);
    righe.push(riga("campaign", c, { campagna_id: c.id, obiettivo: c.objective || "", budget_giorno: giorno, budget_totale: totale,
      inizio: aData(c.start_time || c.created_time), fine: aData(c.stop_time) }));
  }
  // Ad set e annunci: solo quelli che hanno speso o sono accesi
  adsets.filter((a) => attivi(a) || totali.adset.has(a.id)).forEach((a) => righe.push(riga("adset", a, {
    campagna_id: a.campaign_id, budget_giorno: aBudget(a.daily_budget, valuta), budget_totale: aBudget(a.lifetime_budget, valuta),
    inizio: aData(a.start_time), fine: aData(a.end_time) })));
  ads.filter((a) => attivi(a) || totali.ad.has(a.id)).forEach((a) => righe.push(riga("ad", a, { campagna_id: a.campaign_id, adset_id: a.adset_id })));
  await upsert("meta_oggetti", "workspace_id,livello,id", righe);

  // Finestre: le nuove, quelle che toccano gli ultimi giorni e quelle cambiate di confine
  const salvate = await db(`meta_finestre?select=campagna_id,inizio,fine&workspace_id=eq.${ws}&campagna_id=in.${lista(ids)}`);
  const giaFatte = new Map(salvate.map((f) => [`${f.campagna_id}|${f.inizio}`, f.fine]));
  const soglia = sommaGiorni(ieri, -GIORNI_RICALCOLO);
  let nFinestre = 0;
  for (const c of righe.filter((r) => r.livello === "campaign")) {
    const tutte = finestre(c.inizio, minimo(ieri, c.fine || ieri));
    const valide = new Set(tutte.map((f) => f.inizio));
    const vecchie = salvate.filter((f) => f.campagna_id === c.id && !valide.has(f.inizio)).map((f) => f.inizio);
    if (vecchie.length) await db(`meta_finestre?workspace_id=eq.${ws}&campagna_id=eq.${c.id}&inizio=in.${lista(vecchie)}`, { metodo: "DELETE" });
    const daFare = tutte.filter((f) => giaFatte.get(`${c.id}|${f.inizio}`) !== f.fine || f.fine >= soglia);
    for (const blocco of pezzi(daFare, 25)) {
      const r = await g(`${c.id}/insights`, { fields: CAMPI_INSIGHTS, limit: "100",
        time_ranges: JSON.stringify(blocco.map((f) => ({ since: f.inizio, until: f.fine }))) });
      const perInizio = new Map(r.map((x) => [x.date_start, x]));
      await upsert("meta_finestre", "workspace_id,campagna_id,inizio", blocco.map((f) => ({
        workspace_id: ws, campagna_id: c.id, account_id: id, tipo: f.tipo, inizio: f.inizio, fine: f.fine,
        ...vuote, ...(perInizio.has(f.inizio) ? metriche(perInizio.get(f.inizio)) : {}), aggiornato: ora,
      })));
      nFinestre += blocco.length;
    }
  }
  return { campagne: seguite.length, adset: righe.filter((r) => r.livello === "adset").length, annunci: righe.filter((r) => r.livello === "ad").length, finestre: nFinestre };
}

// ---------- uno spazio di lavoro, a tempo ----------
async function richiamati(corpo) {
  const chiave = await db("rpc/meta_cron_chiave", { metodo: "POST", corpo: {} });
  await fetch(`${ENV().url}/functions/v1/meta-sync`, {
    method: "POST", headers: { "Content-Type": "application/json", "x-cron-key": chiave }, body: JSON.stringify(corpo),
  });
}

/* Il token dello spazio di lavoro (collegato con "Collega Facebook"), o quello di riserva */
async function tokenDi(ws) {
  const r = await db(`meta_token?select=token,scadenza&workspace_id=eq.${ws}`);
  return (r[0] && r[0].token) || ENV().token;
}

async function lavora(ws, syncId, inizioMs) {
  const [sync] = await db(`meta_sync?id=eq.${syncId}&select=dettagli`);
  const dettagli = sync.dettagli || {};
  dettagli.account = dettagli.account || {};
  dettagli.giri = (dettagli.giri || 0) + 1;
  const salvaSync = (campi) => db(`meta_sync?id=eq.${syncId}`, { metodo: "PATCH", corpo: { dettagli, ...campi } });
  try {
    const token = await tokenDi(ws);
    if (!token) {
      await salvaSync({ esito: "errore", messaggio: "Facebook non è collegato: premi \"Collega Facebook\" nelle Impostazioni del CRM.", fine: new Date().toISOString() });
      return;
    }
    if (dettagli.giri === 1) dettagli.accountVisti = await aggiornaAccount(ws, token);
    const coll = await db(`meta_collegamenti?select=account_id&workspace_id=eq.${ws}`);
    const accounts = coll.length
      ? await db(`meta_account?select=*&workspace_id=eq.${ws}&account_id=in.${lista(coll.map((c) => c.account_id))}`)
      : [];
    // Da fare in questo giro: i collegati non ancora passati in questo aggiornamento (il registro lo ricorda fra un giro e l'altro)
    const restanti = accounts.filter((a) => !dettagli.account[a.account_id]);
    let limite = false;
    for (const a of restanti) {
      if (Date.now() - inizioMs > TEMPO_MAX) break;
      try {
        dettagli.account[a.account_id] = { nome: a.nome, ...(await sincronizzaAccount(ws, a, token)) };
        await db(`meta_account?workspace_id=eq.${ws}&account_id=eq.${a.account_id}`, { metodo: "PATCH", corpo: { ultimo_sync: new Date().toISOString(), ultimo_errore: "" } });
      } catch (e) {
        dettagli.account[a.account_id] = { nome: a.nome, errore: e.message };
        // Segno il passaggio anche se è andato male: l'errore resta visibile, il giro non si blocca
        await db(`meta_account?workspace_id=eq.${ws}&account_id=eq.${a.account_id}`, { metodo: "PATCH", corpo: { ultimo_sync: new Date().toISOString(), ultimo_errore: e.message.slice(0, 500) } });
        if (e instanceof ErroreMeta && e.limite) { limite = true; break; }
        if (e instanceof ErroreMeta && e.scaduto) throw new Error("L'accesso a Facebook è scaduto o è stato revocato: ricollegalo dalle Impostazioni del CRM.");
      }
    }
    const fatti = accounts.filter((a) => dettagli.account[a.account_id]).length;
    const mancano = accounts.length - fatti;
    if (mancano > 0 && !limite && dettagli.giri < 30) {
      await salvaSync({ messaggio: `In corso: ${fatti} di ${accounts.length} account` });
      await richiamati({ continua: syncId, workspace_id: ws });
      return;
    }
    const errori = Object.values(dettagli.account).filter((x) => x.errore).length;
    const esito = !accounts.length ? "ok" : errori === 0 && !mancano ? "ok" : errori + mancano < accounts.length ? "parziale" : "errore";
    const messaggio = !accounts.length
      ? `Nessun account collegato. Account visibili con il token: ${dettagli.accountVisti ?? "—"}.`
      : `${fatti - errori} account aggiornati${errori ? `, ${errori} con errori` : ""}${limite ? ". Meta ha limitato le richieste: riprovo al prossimo aggiornamento" : ""}${mancano && !limite ? `, ${mancano} non raggiunti` : ""}.`;
    await salvaSync({ esito, messaggio, fine: new Date().toISOString() });
  } catch (e) {
    const msg = e instanceof ErroreMeta && e.scaduto ? "L'accesso a Facebook è scaduto o è stato revocato: ricollegalo dalle Impostazioni del CRM." : e.message;
    await salvaSync({ esito: "errore", messaggio: msg.slice(0, 500), fine: new Date().toISOString() });
  }
}

// ---------- accesso con Facebook ----------
const ritornoValido = (u) => {
  try {
    const x = new URL(u);
    if (ENV().origini.length) return ENV().origini.includes(x.origin);
    // Solo il CRM pubblicato su questo account Cloudflare (versione definitiva e anteprime) o la prova locale;
    // per un dominio proprio si elenca l'origine nel segreto CRM_ORIGINI
    return (x.protocol === "https:" && /^([a-z0-9-]+-)?crm-nc-media\.nciappa9\.workers\.dev$/.test(x.hostname)) || ["localhost", "127.0.0.1"].includes(x.hostname);
  } catch { return false; }
};
const indirizzoRitorno = () => `${ENV().url}/functions/v1/meta-sync`;

/* Prepara la finestra di Facebook: lo state lega il ritorno a questo spazio di lavoro */
async function avviaLogin(ws, ritorno) {
  if (!ENV().appId || !ENV().appSecret) throw new Error("Mancano i segreti META_APP_ID e META_APP_SECRET nelle impostazioni delle Edge Functions.");
  if (!ritornoValido(ritorno)) throw new Error("Indirizzo di ritorno non consentito.");
  const stato = crypto.randomUUID().replace(/-/g, "") + crypto.randomUUID().replace(/-/g, "");
  await db(`meta_login?creato=lt.${new Date(Date.now() - 3600e3).toISOString()}`, { metodo: "DELETE" });
  await db("meta_login", { metodo: "POST", corpo: { stato, workspace_id: ws, ritorno } });
  const q = new URLSearchParams({ client_id: ENV().appId, redirect_uri: indirizzoRitorno(), state: stato, response_type: "code" });
  if (ENV().configId) q.set("config_id", ENV().configId);
  else q.set("scope", "ads_read,business_management");
  return `https://www.facebook.com/${VERSIONE}/dialog/oauth?${q}`;
}

/* Ritorno da Facebook: codice → token breve → token lungo (60 giorni), salvato per lo spazio di lavoro */
async function chiudiLogin(url) {
  const stato = url.searchParams.get("state") || "";
  const righe = stato ? await db(`meta_login?stato=eq.${encodeURIComponent(stato)}&select=*`) : [];
  const voce = righe[0];
  if (!voce) return new Response("Richiesta di accesso scaduta o non valida: riprova dal CRM.", { status: 400, headers: { "Content-Type": "text/plain; charset=utf-8" } });
  await db(`meta_login?stato=eq.${encodeURIComponent(stato)}`, { metodo: "DELETE" });
  const torna = (esito, motivo = "") => {
    const r = new URL(voce.ritorno);
    r.searchParams.set("meta", esito);
    if (motivo) r.searchParams.set("motivo", motivo.slice(0, 200));
    return new Response(null, { status: 302, headers: { Location: r.toString() } });
  };
  if (Date.now() - new Date(voce.creato).getTime() > 15 * 60e3) return torna("errore", "Richiesta scaduta: riprova.");
  const code = url.searchParams.get("code");
  if (!code) return torna("errore", url.searchParams.get("error_description") || "Accesso annullato.");
  try {
    const { appId, appSecret } = ENV();
    const breve = await graph("oauth/access_token", { client_id: appId, client_secret: appSecret, redirect_uri: indirizzoRitorno(), code }, "");
    const lungo = await graph("oauth/access_token", { grant_type: "fb_exchange_token", client_id: appId, client_secret: appSecret, fb_exchange_token: breve.access_token }, "");
    const token = lungo.access_token || breve.access_token;
    const io = await graph("me", { fields: "id,name" }, token);
    const durata = Number(lungo.expires_in || breve.expires_in || 0);
    await upsert("meta_token", "workspace_id", [{
      workspace_id: voce.workspace_id, token, utente_meta: io.name || "",
      scadenza: durata ? new Date(Date.now() + durata * 1000).toISOString() : null, aggiornato: new Date().toISOString(),
    }]);
    // Subito l'elenco degli account, così il CRM può proporre i collegamenti
    try { await aggiornaAccount(voce.workspace_id, token); } catch { /* lo si rifà dal CRM */ }
    return torna("collegato");
  } catch (e) {
    return torna("errore", e.message);
  }
}

async function nuovoSync(ws, origine) {
  // Un aggiornamento alla volta per spazio: se uno è partito da meno di 10 minuti, non ne apro un altro
  const dieciMin = new Date(Date.now() - 10 * 60e3).toISOString();
  const attivi = await db(`meta_sync?select=id&workspace_id=eq.${ws}&esito=eq.in_corso&avvio=gt.${dieciMin}`);
  if (attivi.length) return { id: attivi[0].id, giaInCorso: true };
  const [s] = await db("meta_sync", { metodo: "POST", corpo: { workspace_id: ws, origine }, prefer: "return=representation" });
  return { id: s.id, giaInCorso: false };
}

const sottofondo = (p) => {
  const lavoro = p.catch((e) => console.error("meta-sync", e));
  // @ts-ignore EdgeRuntime esiste solo su Supabase
  if (typeof EdgeRuntime !== "undefined" && EdgeRuntime.waitUntil) EdgeRuntime.waitUntil(lavoro);
  return lavoro;
};

// ---------- ingresso ----------
async function gestisci(req) {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  const inizioMs = Date.now();
  const url = new URL(req.url);
  if (req.method === "GET" && url.searchParams.has("state")) return chiudiLogin(url);
  const corpo = await req.json().catch(() => ({}));
  const chiaveCron = req.headers.get("x-cron-key");

  if (chiaveCron) {
    const giusta = await db("rpc/meta_cron_chiave", { metodo: "POST", corpo: {} });
    if (!giusta || chiaveCron !== giusta) return risposta({ errore: "chiave non valida" }, 401);
    if (corpo.continua) { sottofondo(lavora(corpo.workspace_id, corpo.continua, inizioMs)); return risposta({ ok: true, continua: corpo.continua }, 202); }
    // Cron delle 7 e delle 8 UTC: lavora solo quello che cade alle 9 di Roma
    if (corpo.origine === "programmato" && oraRoma() !== 9 && !corpo.forza) return risposta({ ok: true, saltato: "non sono le 9 a Roma" });
    const spazi = [...new Set((await db("meta_collegamenti?select=workspace_id")).map((c) => c.workspace_id))];
    for (const ws of spazi) {
      const s = await nuovoSync(ws, "programmato");
      if (!s.giaInCorso) await richiamati({ continua: s.id, workspace_id: ws });
    }
    return risposta({ ok: true, spazi: spazi.length }, 202);
  }

  // Dal CRM: utente collegato
  const auth = req.headers.get("Authorization") || "";
  const utente = await fetch(`${ENV().url}/auth/v1/user`, { headers: { apikey: ENV().anon, Authorization: auth } });
  if (!auth || !utente.ok) return risposta({ errore: "accesso richiesto" }, 401);
  const ws = await db("rpc/my_workspace", { metodo: "POST", corpo: {}, auth: { apikey: ENV().anon, Authorization: auth } });
  if (corpo.azione === "login") {
    try { return risposta({ ok: true, url: await avviaLogin(ws, corpo.ritorno || "") }); }
    catch (e) { return risposta({ errore: e.message }, 400); }
  }
  if (corpo.azione === "scollega") {
    await db(`meta_token?workspace_id=eq.${ws}`, { metodo: "DELETE" });
    return risposta({ ok: true });
  }
  if (corpo.azione === "account") {
    const token = await tokenDi(ws);
    if (!token) return risposta({ errore: "Facebook non è collegato: premi \"Collega Facebook\"." }, 400);
    try { return risposta({ ok: true, account: await aggiornaAccount(ws, token) }); }
    catch (e) { return risposta({ errore: e instanceof ErroreMeta && e.scaduto ? "L'accesso a Facebook è scaduto: ricollegalo." : e.message }, 502); }
  }
  const s = await nuovoSync(ws, "manuale");
  if (!s.giaInCorso) sottofondo(lavora(ws, s.id, inizioMs));
  return risposta({ ok: true, sync: s.id, giaInCorso: s.giaInCorso }, 202);
}

Deno.serve(gestisci);
