// Integrazioni con software esterni (Supabase Edge Function, "Verify JWT" spento: la chiama il software esterno).
//  · dal software esterno, con la sua chiave (Authorization: Bearer crmnc_…):
//      POST { evento: "task.creata" | "task.aggiornata" | "task.eliminata", task: {…} }  oppure { tasks: [ … ] }
//      POST { azione: "modifiche", dal: <id> }  → i cambi fatti nel CRM sulle sue task (canale di ritorno)
//  · dal CRM (utente collegato): crea, stato, ritorno, rigenera, elimina, notifica.
// Le task in arrivo vanno in crm_comandi (tipo importa_task): il CRM aperto le inserisce.
import { eventiDaCorpo } from "./logica.js";

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
});
const intestazioniServizio = () => {
  const k = ENV().servizio;
  return k.startsWith("sb_") ? { apikey: k } : { apikey: k, Authorization: `Bearer ${k}` };
};
async function db(percorso: string, { metodo = "GET", corpo, prefer, auth }: { metodo?: string; corpo?: unknown; prefer?: string; auth?: Record<string, string> } = {}) {
  const h: Record<string, string> = { "Content-Type": "application/json", ...(auth || intestazioniServizio()) };
  if (prefer) h.Prefer = prefer;
  const r = await fetch(`${ENV().url}/rest/v1/${percorso}`, { method: metodo, headers: h, body: corpo === undefined ? undefined : JSON.stringify(corpo) });
  const t = await r.text();
  if (!r.ok) throw new Error(`database ${r.status}: ${t.slice(0, 300)}`);
  return t ? JSON.parse(t) : null;
}
const impronta = async (s: string) => {
  const h = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(s));
  return [...new Uint8Array(h)].map((b) => b.toString(16).padStart(2, "0")).join("");
};
const nuovaChiave = () => "crmnc_" + crypto.randomUUID().replace(/-/g, "") + crypto.randomUUID().replace(/-/g, "").slice(0, 16);
const indirizzo = () => `${ENV().url}/functions/v1/integrazioni`;

/* Manda i cambi all'indirizzo di ritorno del software esterno, se c'è */
async function consegna(integ: Record<string, any>, righe: Record<string, any>[]) {
  if (!integ.url_ritorno || !righe.length) return;
  let esito = "";
  try {
    const r = await fetch(integ.url_ritorno, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...(integ.chiave_ritorno ? { Authorization: `Bearer ${integ.chiave_ritorno}` } : {}) },
      body: JSON.stringify({ evento: "task.modificata_nel_crm", modifiche: righe.map((x) => ({ id: x.id, esternoId: x.esterno_id, ...x.modifiche, quando: x.creato })) }),
    });
    esito = r.ok ? "ok" : `risposta ${r.status}`;
  } catch (e) { esito = `non raggiungibile: ${(e as Error).message.slice(0, 120)}`; }
  await db(`integrazione_uscite?id=in.(${righe.map((x) => x.id).join(",")})`, { metodo: "PATCH",
    corpo: { esito, ...(esito === "ok" ? { inviato: new Date().toISOString() } : {}) } });
}

async function gestisci(req: Request) {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  const auth = req.headers.get("Authorization") || "";
  const corpo = await req.json().catch(() => ({}));

  // ---------- dal software esterno ----------
  if (/^Bearer crmnc_/.test(auth)) {
    const [integ] = await db(`integrazioni?select=*&chiave_hash=eq.${await impronta(auth.slice(7).trim())}&attiva=eq.true`);
    if (!integ) return risposta({ errore: "chiave non valida o integrazione disattivata" }, 401);
    await db(`integrazioni?id=eq.${integ.id}`, { metodo: "PATCH", corpo: { ultimo_contatto: new Date().toISOString() } });
    // Canale di ritorno a lettura: i cambi fatti nel CRM dopo un certo id
    if (corpo.azione === "modifiche") {
      const dal = Number(corpo.dal) || 0;
      const righe = await db(`integrazione_uscite?select=id,esterno_id,modifiche,creato&integrazione_id=eq.${integ.id}&id=gt.${dal}&order=id.asc&limit=200`);
      return risposta({ ok: true, modifiche: righe.map((x: any) => ({ id: x.id, esternoId: x.esterno_id, ...x.modifiche, quando: x.creato })),
        ultimo: righe.length ? righe[righe.length - 1].id : dal });
    }
    const eventi = eventiDaCorpo(corpo);
    if (!eventi.length) return risposta({ errore: "nessuna task valida: serve almeno { task: { id, titolo } }" }, 400);
    await db("crm_comandi", { metodo: "POST", prefer: "return=minimal", corpo: eventi.map((e) => ({
      id: crypto.randomUUID(), workspace_id: integ.workspace_id, tipo: "importa_task", origine: integ.nome,
      dati: { ...e, integrazioneId: integ.id, fonte: integ.nome },
    })) });
    return risposta({ ok: true, ricevute: eventi.length });
  }

  // ---------- dal CRM ----------
  const utente = await fetch(`${ENV().url}/auth/v1/user`, { headers: { apikey: ENV().anon, Authorization: auth } });
  if (!auth || !utente.ok) return risposta({ errore: "accesso richiesto" }, 401);
  const ws = await db("rpc/my_workspace", { metodo: "POST", corpo: {}, auth: { apikey: ENV().anon, Authorization: auth } });
  const mia = async (id: string) => {
    const [x] = await db(`integrazioni?select=*&id=eq.${encodeURIComponent(String(id || ""))}&workspace_id=eq.${ws}`);
    if (!x) throw new Error("integrazione non trovata");
    return x;
  };
  try {
    if (corpo.azione === "stato") {
      const l = await db(`integrazioni?select=id,nome,chiave_fine,url_ritorno,attiva,creato,ultimo_contatto&workspace_id=eq.${ws}&order=creato.asc`);
      return risposta({ ok: true, indirizzo: indirizzo(), integrazioni: l.map((x: any) => ({ ...x, ritorno: !!x.url_ritorno })) });
    }
    if (corpo.azione === "crea") {
      const chiave = nuovaChiave();
      const [x] = await db("integrazioni", { metodo: "POST", prefer: "return=representation",
        corpo: { workspace_id: ws, nome: String(corpo.nome || "Integrazione").slice(0, 60), chiave_hash: await impronta(chiave), chiave_fine: chiave.slice(-4) } });
      return risposta({ ok: true, id: x.id, chiave, indirizzo: indirizzo() });
    }
    if (corpo.azione === "rigenera") {
      const x = await mia(corpo.id);
      const chiave = nuovaChiave();
      await db(`integrazioni?id=eq.${x.id}`, { metodo: "PATCH", corpo: { chiave_hash: await impronta(chiave), chiave_fine: chiave.slice(-4) } });
      return risposta({ ok: true, chiave, indirizzo: indirizzo() });
    }
    if (corpo.azione === "ritorno") {
      const x = await mia(corpo.id);
      const url = String(corpo.url || "").trim();
      if (url && !/^https:\/\//.test(url)) return risposta({ errore: "L'indirizzo di ritorno deve iniziare con https://" }, 400);
      await db(`integrazioni?id=eq.${x.id}`, { metodo: "PATCH", corpo: { url_ritorno: url, ...(corpo.chiave !== undefined ? { chiave_ritorno: String(corpo.chiave || "") } : {}) } });
      return risposta({ ok: true });
    }
    if (corpo.azione === "attiva") {
      const x = await mia(corpo.id);
      await db(`integrazioni?id=eq.${x.id}`, { metodo: "PATCH", corpo: { attiva: !!corpo.attiva } });
      return risposta({ ok: true });
    }
    if (corpo.azione === "elimina") {
      const x = await mia(corpo.id);
      await db(`integrazioni?id=eq.${x.id}`, { metodo: "DELETE" });
      return risposta({ ok: true });
    }
    if (corpo.azione === "notifica") {
      const x = await mia(corpo.id);
      const mod = (Array.isArray(corpo.modifiche) ? corpo.modifiche : []).slice(0, 200)
        .filter((m: any) => m && m.esternoId).map((m: any) => {
          const { esternoId, ...resto } = m;
          return { workspace_id: ws, integrazione_id: x.id, esterno_id: String(esternoId), modifiche: resto };
        });
      if (!mod.length) return risposta({ ok: true, registrate: 0 });
      const righe = await db("integrazione_uscite", { metodo: "POST", prefer: "return=representation", corpo: mod });
      await consegna(x, righe);
      return risposta({ ok: true, registrate: righe.length });
    }
    return risposta({ errore: "azione sconosciuta" }, 400);
  } catch (e) {
    return risposta({ errore: (e as Error).message }, 400);
  }
}

Deno.serve(gestisci);
