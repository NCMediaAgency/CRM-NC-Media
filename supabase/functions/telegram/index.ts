// Bot Telegram del CRM (Supabase Edge Function, "Verify JWT" spento: la chiamano Telegram e il cron).
//  · da Telegram (webhook, intestazione X-Telegram-Bot-Api-Secret-Token): messaggi e pulsanti;
//  · dal cron ogni 5 minuti (x-cron-key, migrazione 0006): riepilogo del mattino e promemoria;
//  · dal CRM (utente collegato): collega una chat, preferenze, scollega, messaggio di prova.
// Il bot non scrive crm:data:v6: mette i comandi in crm_comandi e il CRM aperto li applica.
// Segreti: TELEGRAM_TOKEN (da @BotFather), ANTHROPIC_API_KEY; facoltativo CRM_URL.
import Anthropic from "npm:@anthropic-ai/sdk";
import {
  oraRoma, contestoCRM, ISTRUZIONI, STRUMENTI, comandoDaStrumento, riepilogo, taskDaRicordare,
  bottoniTask, comandoDaPulsante, oraDelRiepilogo,
} from "./logica.js";

const MODELLO = "claude-opus-5-5";
const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const risposta = (corpo: unknown, stato = 200) =>
  new Response(JSON.stringify(corpo), { status: stato, headers: { ...CORS, "Content-Type": "application/json" } });

const primaChiave = (nome: string) => {
  try { const d = JSON.parse(Deno.env.get(nome) || "{}"); return d.default || Object.values(d)[0] || ""; } catch { return ""; }
};
const ENV = () => ({
  url: Deno.env.get("SUPABASE_URL") || "",
  servizio: primaChiave("SUPABASE_SECRET_KEYS") || Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") || "",
  anon: primaChiave("SUPABASE_PUBLISHABLE_KEYS") || Deno.env.get("SUPABASE_ANON_KEY") || "",
  telegram: Deno.env.get("TELEGRAM_TOKEN") || "",
  anthropic: Deno.env.get("ANTHROPIC_API_KEY") || "",
  crm: Deno.env.get("CRM_URL") || "https://crm-nc-media.nciappa9.workers.dev",
});

// ---------- database (PostgREST con la chiave di servizio) ----------
const intestazioniServizio = () => {
  const k = ENV().servizio;
  return k.startsWith("sb_") ? { apikey: k } : { apikey: k, Authorization: `Bearer ${k}` };
};
async function db(percorso: string, { metodo = "GET", corpo, prefer, auth }: { metodo?: string; corpo?: unknown; prefer?: string; auth?: Record<string, string> } = {}) {
  const h: Record<string, string> = { "Content-Type": "application/json", ...(auth || intestazioniServizio()) };
  if (prefer) h.Prefer = prefer;
  const r = await fetch(`${ENV().url}/rest/v1/${percorso}`, { method: metodo, headers: h, body: corpo === undefined ? undefined : JSON.stringify(corpo) });
  const testo = await r.text();
  if (!r.ok) throw new Error(`database ${r.status}: ${testo.slice(0, 300)}`);
  return testo ? JSON.parse(testo) : null;
}

/* Dati del CRM dello spazio di lavoro (sola lettura) */
async function datiCRM(ws: string) {
  const r = await db(`kv_store?select=value&workspace_id=eq.${ws}&key=eq.${encodeURIComponent("crm:data:v6")}`);
  try { return r && r[0] ? JSON.parse(r[0].value) : {}; } catch { return {}; }
}
const comandiInCoda = (ws: string) => db(`crm_comandi?select=tipo,dati&workspace_id=eq.${ws}&applicato=is.null&order=creato.asc`);
async function accoda(ws: string, tipo: string, dati: Record<string, unknown>) {
  await db("crm_comandi", { metodo: "POST", corpo: { id: crypto.randomUUID(), workspace_id: ws, tipo, dati, origine: "telegram" }, prefer: "return=minimal" });
}

// ---------- Telegram ----------
async function tg(metodo: string, corpo: Record<string, unknown>) {
  const r = await fetch(`https://api.telegram.org/bot${ENV().telegram}/${metodo}`, {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(corpo),
  });
  const j = await r.json().catch(() => ({}));
  if (!j.ok) throw new Error(`Telegram ${metodo}: ${j.description || r.status}`);
  return j.result;
}
const manda = (chat: number, testo: string, bottoni?: unknown[]) =>
  tg("sendMessage", { chat_id: chat, text: testo.slice(0, 4000), ...(bottoni && bottoni.length ? { reply_markup: { inline_keyboard: bottoni } } : {}) });

/* Segreto del webhook ricavato dal token: Telegram lo rimanda in ogni chiamata, nessun segreto in più */
async function segretoWebhook() {
  const h = await crypto.subtle.digest("SHA-256", new TextEncoder().encode("crm-nc-media:" + ENV().telegram));
  return [...new Uint8Array(h)].map((b) => b.toString(16).padStart(2, "0")).join("").slice(0, 48);
}

/* Fuso orario dello spazio di lavoro (migrazione 0010; senza impostazioni: Roma) */
const fusi = new Map<string, string>();
async function fusoDi(ws: string) {
  if (!fusi.has(ws)) {
    const r = await db(`workspace_impostazioni?select=fuso&workspace_id=eq.${ws}`).catch(() => []);
    fusi.set(ws, (r[0] && r[0].fuso) || "Europe/Rome");
  }
  return fusi.get(ws) as string;
}

// ---------- Claude ----------
type Storico = { ruolo: string; testo: string };
async function chiediAClaude(ws: string, chat: number, testo: string) {
  const [data, inCoda, storico] = await Promise.all([
    datiCRM(ws), comandiInCoda(ws),
    db(`telegram_storico?select=ruolo,testo&chat_id=eq.${chat}&order=id.desc&limit=10`) as Promise<Storico[]>,
  ]);
  const adesso = oraRoma(new Date(), await fusoDi(ws));
  const contesto = contestoCRM(data, adesso.data, inCoda);
  const client = new Anthropic({ apiKey: ENV().anthropic });
  // Storia come testo semplice (niente blocchi di ragionamento da rimandare); il contesto va solo nell'ultimo messaggio
  const messaggi: Anthropic.Beta.BetaMessageParam[] = storico.reverse().map((m) => ({ role: m.ruolo as "user" | "assistant", content: m.testo }));
  while (messaggi.length && messaggi[0].role !== "user") messaggi.shift();
  messaggi.push({ role: "user", content: `Contesto CRM (oggi ${adesso.data}, ore ${adesso.hhmm}):\n${JSON.stringify(contesto)}\n\nMessaggio del titolare:\n${testo}` });

  const r = await client.beta.messages.create({
    model: MODELLO,
    max_tokens: 8000,
    // Se il modello declina, la richiesta riparte su un altro modello nella stessa chiamata
    betas: ["server-side-fallback-2026-07-01"],
    fallbacks: "default",
    output_config: { effort: "low" },
    system: [{ type: "text", text: ISTRUZIONI, cache_control: { type: "ephemeral" } }],
    tools: STRUMENTI,
    tool_choice: { type: "auto" },
    messages: messaggi,
  } as unknown as Anthropic.Beta.MessageCreateParamsNonStreaming) as Anthropic.Beta.BetaMessage;

  if (r.stop_reason === "refusal") return { testo: "Non riesco a gestire questa richiesta: prova a riformularla.", comandi: [] };
  const testi: string[] = [];
  const comandi: { tipo: string; dati: Record<string, unknown>; conferma: string }[] = [];
  for (const b of r.content) {
    if (b.type === "text" && b.text.trim()) testi.push(b.text.trim());
    if (b.type === "tool_use") {
      const c = comandoDaStrumento(b.name, b.input as Record<string, unknown>, { data, oggi: adesso.data, nuovoId: () => crypto.randomUUID().slice(0, 8) });
      if (c) comandi.push(c);
    }
  }
  return { testo: testi.join("\n\n"), comandi };
}

// ---------- messaggi in arrivo ----------
async function gestisciAggiornamento(u: Record<string, any>) {
  // Pulsanti dei promemoria
  if (u.callback_query) {
    const q = u.callback_query, chat = q.message?.chat?.id;
    const [c] = await db(`telegram_chat?select=workspace_id&chat_id=eq.${chat}&attivo=eq.true`);
    if (!c) { await tg("answerCallbackQuery", { callback_query_id: q.id, text: "Chat non collegata al CRM" }); return; }
    const cmd = comandoDaPulsante(q.data, oraRoma(new Date(), await fusoDi(c.workspace_id)).data);
    if (!cmd) { await tg("answerCallbackQuery", { callback_query_id: q.id }); return; }
    await accoda(c.workspace_id, cmd.tipo, cmd.dati);
    await tg("answerCallbackQuery", { callback_query_id: q.id, text: cmd.testo });
    // Il pulsante premuto diventa la conferma, gli altri restano
    const tastiera = (q.message?.reply_markup?.inline_keyboard || [])
      .map((riga: any[]) => riga.some((b) => String(b.callback_data || "").endsWith(":" + cmd.dati.taskId)) ? [{ text: `${cmd.testo}`, callback_data: "ok:" }] : riga);
    await tg("editMessageReplyMarkup", { chat_id: chat, message_id: q.message.message_id, reply_markup: { inline_keyboard: tastiera } }).catch(() => {});
    return;
  }
  const m = u.message;
  if (!m || !m.chat) return;
  const chat = m.chat.id as number;
  const testo = String(m.text || "").trim();

  // Collegamento: /start CODICE dal link generato nel CRM
  const avvio = /^\/start(?:\s+(\S+))?/.exec(testo);
  if (avvio) {
    const codice = avvio[1] || "";
    const [k] = codice ? await db(`telegram_codici?select=workspace_id,creato&codice=eq.${encodeURIComponent(codice)}`) : [];
    if (!k || Date.now() - new Date(k.creato).getTime() > 30 * 60e3) {
      const [gia] = await db(`telegram_chat?select=workspace_id&chat_id=eq.${chat}`);
      await manda(chat, gia ? "Sei già collegato al CRM: scrivimi pure." : "Per collegarti apri il link dalle Impostazioni del CRM (sezione Telegram): il codice vale 30 minuti.");
      return;
    }
    await db(`telegram_codici?codice=eq.${encodeURIComponent(codice)}`, { metodo: "DELETE" });
    const utente = [m.from?.first_name, m.from?.last_name].filter(Boolean).join(" ") || m.from?.username || "";
    await db("telegram_chat?on_conflict=chat_id", { metodo: "POST", prefer: "resolution=merge-duplicates,return=minimal",
      corpo: { workspace_id: k.workspace_id, chat_id: chat, utente, attivo: true } });
    await manda(chat, `Collegato al CRM ✅\n\nScrivimi cosa fare, per esempio:\n• "Crea un task per BDF per giovedì: report di ottobre, priorità alta"\n• "Segna nel diario che oggi abbiamo spento le campagne Rottamazione"\n• "Che scadenze ho domani?"\n\nOgni mattina ti mando il riepilogo dei task, con i pulsanti per chiuderli o spostarli.`);
    return;
  }

  const [c] = await db(`telegram_chat?select=workspace_id&chat_id=eq.${chat}&attivo=eq.true`);
  if (!c) { await manda(chat, "Questo bot è collegato a un CRM privato. Se sei il titolare, collegati dalle Impostazioni del CRM."); return; }
  if (m.voice || m.audio) { await manda(chat, "I vocali arrivano a breve: per ora scrivimi il messaggio."); return; }
  if (!testo) return;
  if (!ENV().anthropic) { await manda(chat, "Manca la chiave ANTHROPIC_API_KEY tra i segreti di Supabase: senza non posso capire i messaggi."); return; }

  await tg("sendChatAction", { chat_id: chat, action: "typing" }).catch(() => {});
  let esito;
  try { esito = await chiediAClaude(c.workspace_id, chat, testo); }
  catch (e) { await manda(chat, `Non sono riuscito a rispondere: ${(e as Error).message.slice(0, 200)}`); return; }
  for (const cmd of esito.comandi) await accoda(c.workspace_id, cmd.tipo, cmd.dati);
  const conferme = esito.comandi.map((x) => x.conferma);
  const finale = [esito.testo, ...conferme].filter(Boolean).join("\n\n") || "Ok.";
  // Un task con data appena creato porta già i suoi pulsanti
  const unico = esito.comandi.length === 1 && esito.comandi[0].tipo === "crea_task" ? bottoniTask(String(esito.comandi[0].dati.id)) : undefined;
  await manda(chat, finale + (esito.comandi.length ? "\n\nNel CRM compare entro un minuto se è aperto, altrimenti alla prossima apertura." : ""), unico);
  await db("telegram_storico", { metodo: "POST", prefer: "return=minimal", corpo: [
    { workspace_id: c.workspace_id, chat_id: chat, ruolo: "user", testo },
    { workspace_id: c.workspace_id, chat_id: chat, ruolo: "assistant", testo: finale },
  ] });
  // La memoria resta corta: bastano gli ultimi scambi per le correzioni
  await db(`telegram_storico?chat_id=eq.${chat}&creato=lt.${new Date(Date.now() - 2 * 864e5).toISOString()}`, { metodo: "DELETE" });
}

// ---------- promemoria (cron ogni 5 minuti) ----------
const giaInviato = async (ws: string, chiave: string) =>
  (await db(`telegram_inviati?select=chiave&workspace_id=eq.${ws}&chiave=eq.${encodeURIComponent(chiave)}`)).length > 0;
const segnaInviato = (ws: string, chiave: string) =>
  db("telegram_inviati?on_conflict=workspace_id,chiave", { metodo: "POST", prefer: "resolution=ignore-duplicates,return=minimal", corpo: { workspace_id: ws, chiave } });

async function promemoria() {
  const chats = await db("telegram_chat?select=*&attivo=eq.true");
  const perSpazio = new Map<string, any>();
  for (const chat of chats) {
    if (!perSpazio.has(chat.workspace_id)) perSpazio.set(chat.workspace_id, { data: await datiCRM(chat.workspace_id), coda: await comandiInCoda(chat.workspace_id) });
    const { data, coda } = perSpazio.get(chat.workspace_id);
    // Riepilogo e promemoria all'ora locale dello spazio di lavoro
    const adesso = oraRoma(new Date(), await fusoDi(chat.workspace_id));
    try {
      const chiaveR = `riepilogo:${chat.chat_id}:${adesso.data}`;
      if (oraDelRiepilogo(adesso, chat) && !(await giaInviato(chat.workspace_id, chiaveR))) {
        await segnaInviato(chat.workspace_id, chiaveR);
        const r = riepilogo(data, adesso.data, coda);
        await manda(chat.chat_id, r.testo, r.bottoni);
      }
      for (const t of taskDaRicordare(data, adesso.data, adesso.hhmm, chat.preavviso ?? 30)) {
        const chiaveT = `task:${chat.chat_id}:${t.id}:${adesso.data}:${t.ora}`;
        if (await giaInviato(chat.workspace_id, chiaveT)) continue;
        await segnaInviato(chat.workspace_id, chiaveT);
        await manda(chat.chat_id, `🔔 Alle ${t.ora}: ${t.titolo}${t.clienteId ? ` · ${((data.clienti || []).find((z: any) => z.id === t.clienteId) || {}).nome || ""}` : ""}`, bottoniTask(t.id));
      }
    } catch (e) { console.error("promemoria", chat.chat_id, e); }
  }
  // Pulizia: promemoria e codici vecchi
  await db(`telegram_inviati?inviato=lt.${new Date(Date.now() - 7 * 864e5).toISOString()}`, { metodo: "DELETE" });
  await db(`telegram_codici?creato=lt.${new Date(Date.now() - 864e5).toISOString()}`, { metodo: "DELETE" });
  return chats.length;
}

const sottofondo = (p: Promise<unknown>) => {
  const lavoro = p.catch((e) => console.error("telegram", e));
  // @ts-ignore EdgeRuntime esiste solo su Supabase
  if (typeof EdgeRuntime !== "undefined" && EdgeRuntime.waitUntil) EdgeRuntime.waitUntil(lavoro);
  return lavoro;
};

// ---------- ingresso ----------
async function gestisci(req: Request) {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  if (!ENV().telegram) return risposta({ errore: "Manca il segreto TELEGRAM_TOKEN nelle Edge Functions di Supabase." }, 400);

  // Da Telegram: rispondo subito e lavoro in sottofondo (Claude può metterci qualche secondo)
  const segreto = req.headers.get("x-telegram-bot-api-secret-token");
  if (segreto !== null) {
    if (segreto !== await segretoWebhook()) return risposta({ errore: "non autorizzato" }, 401);
    const u = await req.json().catch(() => null);
    if (u) sottofondo(gestisciAggiornamento(u));
    return risposta({ ok: true });
  }

  const corpo = await req.json().catch(() => ({}));
  // Dal cron: promemoria
  const chiaveCron = req.headers.get("x-cron-key");
  if (chiaveCron) {
    const giusta = await db("rpc/meta_cron_chiave", { metodo: "POST", corpo: {} });
    if (!giusta || chiaveCron !== giusta) return risposta({ errore: "chiave non valida" }, 401);
    return risposta({ ok: true, chat: await promemoria() });
  }

  // Dal CRM: utente collegato
  const auth = req.headers.get("Authorization") || "";
  const utente = await fetch(`${ENV().url}/auth/v1/user`, { headers: { apikey: ENV().anon, Authorization: auth } });
  if (!auth || !utente.ok) return risposta({ errore: "accesso richiesto" }, 401);
  const ws = await db("rpc/my_workspace", { metodo: "POST", corpo: {}, auth: { apikey: ENV().anon, Authorization: auth } });
  try {
    if (corpo.azione === "collega") {
      // Webhook (idempotente) e link di collegamento con un codice usa e getta
      await tg("setWebhook", { url: `${ENV().url}/functions/v1/telegram`, secret_token: await segretoWebhook(), allowed_updates: ["message", "callback_query"] });
      const io = await tg("getMe", {});
      const codice = crypto.randomUUID().replace(/-/g, "").slice(0, 24);
      await db("telegram_codici", { metodo: "POST", prefer: "return=minimal", corpo: { codice, workspace_id: ws } });
      return risposta({ ok: true, bot: io.username, link: `https://t.me/${io.username}?start=${codice}`, anthropic: !!ENV().anthropic });
    }
    if (corpo.azione === "preferenze" && corpo.chat) {
      const campi: Record<string, unknown> = {};
      if (/^\d{1,2}:\d{2}$/.test(corpo.ora_riepilogo || "")) campi.ora_riepilogo = corpo.ora_riepilogo.padStart(5, "0");
      if (typeof corpo.weekend === "boolean") campi.weekend = corpo.weekend;
      if (Number.isFinite(corpo.preavviso)) campi.preavviso = Math.max(0, Math.min(240, Math.round(corpo.preavviso)));
      if (typeof corpo.attivo === "boolean") campi.attivo = corpo.attivo;
      await db(`telegram_chat?workspace_id=eq.${ws}&chat_id=eq.${Number(corpo.chat)}`, { metodo: "PATCH", corpo: campi });
      return risposta({ ok: true });
    }
    if (corpo.azione === "scollega" && corpo.chat) {
      await db(`telegram_chat?workspace_id=eq.${ws}&chat_id=eq.${Number(corpo.chat)}`, { metodo: "DELETE" });
      return risposta({ ok: true });
    }
    if (corpo.azione === "prova") {
      const chats = await db(`telegram_chat?select=chat_id&workspace_id=eq.${ws}&attivo=eq.true`);
      const data = await datiCRM(ws);
      const r = riepilogo(data, oraRoma(new Date(), await fusoDi(ws)).data, await comandiInCoda(ws));
      for (const c of chats) await manda(c.chat_id, `Messaggio di prova dal CRM 👋\n\n${r.testo}`, r.bottoni);
      return risposta({ ok: true, inviati: chats.length });
    }
    return risposta({ errore: "azione sconosciuta" }, 400);
  } catch (e) {
    return risposta({ errore: (e as Error).message }, 502);
  }
}

Deno.serve(gestisci);
