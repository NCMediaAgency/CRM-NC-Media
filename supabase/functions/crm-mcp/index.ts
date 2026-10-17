// Connettore del CRM per Claude (server MCP remoto, Supabase Edge Function, "Verify JWT" spento).
// In claude.ai: Impostazioni → Connettori → Aggiungi connettore personalizzato, con l'indirizzo
//   https://<progetto>.supabase.co/functions/v1/crm-mcp/<chiave>
// La chiave (crmnc_…) si crea nel CRM, Impostazioni → Connettore per Claude; nel database ne resta
// solo l'impronta (tabella integrazioni, tipo 'mcp'). Il Claude del cliente legge i dati del CRM,
// salva i numeri degli annunci (analisi_ads) e le analisi scritte (analisi_testi), crea task (crm_comandi).
import { STRUMENTI, trovaCliente, riassuntoCliente, benchmark, righeAds } from "./logica.js";

const VERSIONI = ["2025-06-18", "2025-03-26", "2024-11-05"];
const primaChiave = (nome: string) => {
  try { const d = JSON.parse(Deno.env.get(nome) || "{}"); return d.default || Object.values(d)[0] || ""; } catch { return ""; }
};
const ENV = () => ({
  url: Deno.env.get("SUPABASE_URL") || "",
  servizio: primaChiave("SUPABASE_SECRET_KEYS") || Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") || "",
});
const intestazioniServizio = () => {
  const k = ENV().servizio;
  return k.startsWith("sb_") ? { apikey: k } : { apikey: k, Authorization: `Bearer ${k}` };
};
async function db(percorso: string, { metodo = "GET", corpo, prefer }: { metodo?: string; corpo?: unknown; prefer?: string } = {}) {
  const h: Record<string, string> = { "Content-Type": "application/json", ...intestazioniServizio() };
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
const oggiRoma = () => new Intl.DateTimeFormat("en-CA", { timeZone: "Europe/Rome" }).format(new Date());
const json = (corpo: unknown, stato = 200) => new Response(JSON.stringify(corpo), { status: stato, headers: { "Content-Type": "application/json" } });

async function datiCRM(ws: string) {
  const [r] = await db(`kv_store?select=value&workspace_id=eq.${ws}&key=eq.${encodeURIComponent("crm:data:v6")}`);
  if (!r) return { clienti: [] };
  return typeof r.value === "string" ? JSON.parse(r.value) : r.value;
}

/* Esegue uno strumento: restituisce il testo per Claude */
async function strumento(ws: string, nome: string, args: Record<string, any>) {
  const data = await datiCRM(ws);
  if (nome === "elenco_clienti") {
    const l = (data.clienti || []).filter((z: any) => z.archiviato !== "si");
    const nomeDi = (id: string) => (l.find((z: any) => z.id === id) || {}).nome || "";
    return l.map((z: any) => ({ id: z.id, nome: z.nome, ...(z.padreId ? { principale: nomeDi(z.padreId) } : {}) }));
  }
  const t = trovaCliente(data.clienti, args.cliente);
  if (t.errore) throw new Error(t.errore);
  const z = t.cliente;
  if (nome === "dati_cliente") return riassuntoCliente(data, z);
  if (nome === "leggi_benchmark") return { cliente: z.nome, ...benchmark(z) };
  if (nome === "salva_dati_ads") {
    const { data: giorno, fonte, righe } = righeAds(args, { ws, clienteId: z.id, oggi: oggiRoma() });
    if (!righe.length) throw new Error("Nessun annuncio valido: ogni annuncio deve avere almeno id e nome.");
    await db(`analisi_ads?workspace_id=eq.${ws}&cliente_id=eq.${encodeURIComponent(z.id)}&fonte=eq.${encodeURIComponent(fonte)}&data=eq.${giorno}`, { metodo: "DELETE" });
    for (let i = 0; i < righe.length; i += 500) await db("analisi_ads", { metodo: "POST", prefer: "return=minimal", corpo: righe.slice(i, i + 500) });
    return `Salvati ${righe.length} annunci di ${z.nome} (${fonte}, ${giorno}). La classifica è nella sezione Winner ads del cliente nel CRM.`;
  }
  if (nome === "salva_analisi") {
    const testo = String(args.testo || "").trim();
    if (!testo) throw new Error("Il testo dell'analisi è vuoto.");
    await db("analisi_testi", { metodo: "POST", prefer: "return=minimal", corpo: {
      workspace_id: ws, cliente_id: z.id, tipo: ["winner_ads", "creativita", "andamento", "altro"].includes(args.tipo) ? args.tipo : "altro",
      titolo: String(args.titolo || "Analisi").slice(0, 200), testo: testo.slice(0, 60000), fonte: "Claude" } });
    return `Analisi "${String(args.titolo || "Analisi").slice(0, 200)}" salvata nel CRM per ${z.nome}.`;
  }
  if (nome === "leggi_analisi") {
    const n = Math.min(10, Math.max(1, Number(args.limite) || 3));
    const l = await db(`analisi_testi?select=tipo,titolo,testo,creato&workspace_id=eq.${ws}&cliente_id=eq.${encodeURIComponent(z.id)}&order=creato.desc&limit=${n}`);
    return l.length ? l : `Nessuna analisi salvata per ${z.nome}.`;
  }
  if (nome === "crea_task") {
    const titolo = String(args.titolo || "").trim();
    if (!titolo) throw new Error("Manca il titolo del task.");
    await db("crm_comandi", { metodo: "POST", prefer: "return=minimal", corpo: {
      id: crypto.randomUUID(), workspace_id: ws, tipo: "crea_task", origine: "claude",
      dati: { id: crypto.randomUUID(), titolo: titolo.slice(0, 200), clienteId: z.id, note: String(args.note || "").slice(0, 4000),
        scadenza: /^\d{4}-\d{2}-\d{2}$/.test(String(args.scadenza || "")) ? args.scadenza : "", priorita: args.priorita || "media", origine: "claude" } } });
    return `Task "${titolo}" creato per ${z.nome}: compare nel CRM entro un minuto.`;
  }
  throw new Error(`Strumento sconosciuto: ${nome}`);
}

async function rispondi(ws: string, m: any) {
  const ok = (result: unknown) => ({ jsonrpc: "2.0", id: m.id, result });
  const ko = (code: number, message: string) => ({ jsonrpc: "2.0", id: m.id ?? null, error: { code, message } });
  if (!m || m.jsonrpc !== "2.0" || typeof m.method !== "string") return ko(-32600, "Richiesta non valida");
  if (m.id === undefined || m.id === null) return null;   // notifica: nessuna risposta
  if (m.method === "initialize") {
    const chiesta = m.params && m.params.protocolVersion;
    return ok({
      protocolVersion: VERSIONI.includes(chiesta) ? chiesta : VERSIONI[0],
      capabilities: { tools: { listChanged: false } },
      serverInfo: { name: "crm-nc-media", title: "CRM NC Media", version: "1.0.0" },
      instructions: "Connettore del CRM NC Media. Per le winner ads: leggi_benchmark, poi salva_dati_ads con i numeri attribuiti a ogni annuncio (per esempio da Hyros), poi salva_analisi con il commento. Usa elenco_clienti se non conosci il nome esatto del cliente.",
    });
  }
  if (m.method === "ping") return ok({});
  if (m.method === "tools/list") return ok({ tools: STRUMENTI });
  if (m.method === "tools/call") {
    const p = m.params || {};
    try {
      const r = await strumento(ws, String(p.name || ""), p.arguments || {});
      return ok({ content: [{ type: "text", text: typeof r === "string" ? r : JSON.stringify(r) }] });
    } catch (e) {
      return ok({ content: [{ type: "text", text: (e as Error).message }], isError: true });
    }
  }
  if (m.method === "resources/list") return ok({ resources: [] });
  if (m.method === "prompts/list") return ok({ prompts: [] });
  return ko(-32601, `Metodo non supportato: ${m.method}`);
}

async function gestisci(req: Request) {
  // La chiave arriva nell'indirizzo (/crm-mcp/crmnc_…) o come Bearer
  const dalPercorso = new URL(req.url).pathname.split("/").find((x) => x.startsWith("crmnc_")) || "";
  const auth = req.headers.get("Authorization") || "";
  const chiave = dalPercorso || (/^Bearer crmnc_/.test(auth) ? auth.slice(7).trim() : "");
  if (!chiave) return json({ errore: "chiave mancante" }, 401);
  const [integ] = await db(`integrazioni?select=id,workspace_id,tipo&chiave_hash=eq.${await impronta(chiave)}&attiva=eq.true`);
  if (!integ || integ.tipo !== "mcp") return json({ errore: "chiave non valida o connettore disattivato" }, 401);
  if (req.method === "GET") return new Response("Usa POST (MCP Streamable HTTP, senza flusso SSE)", { status: 405, headers: { Allow: "POST, DELETE" } });
  if (req.method === "DELETE") return new Response(null, { status: 204 });
  if (req.method !== "POST") return new Response(null, { status: 405 });
  let corpo: any;
  try { corpo = await req.json(); } catch { return json({ jsonrpc: "2.0", id: null, error: { code: -32700, message: "JSON non valido" } }, 400); }
  db(`integrazioni?id=eq.${integ.id}`, { metodo: "PATCH", corpo: { ultimo_contatto: new Date().toISOString() } }).catch(() => {});
  if (Array.isArray(corpo)) {
    const r = (await Promise.all(corpo.map((m) => rispondi(integ.workspace_id, m)))).filter(Boolean);
    return r.length ? json(r) : new Response(null, { status: 202 });
  }
  const r = await rispondi(integ.workspace_id, corpo);
  return r ? json(r) : new Response(null, { status: 202 });
}

Deno.serve(gestisci);
