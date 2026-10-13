import { supabase } from "./supabase";

/* Il CRM parla con Meta solo attraverso queste funzioni: legge le tabelle meta_*
   (scritte dalla funzione meta-sync), decide i collegamenti e chiede gli aggiornamenti. */

/* PostgREST restituisce al massimo 1000 righe per richiesta: leggo a pagine */
async function tutte(tabella, ws, colonne = "*") {
  const out = [];
  for (let da = 0; ; da += 1000) {
    const { data, error } = await supabase.from(tabella).select(colonne).eq("workspace_id", ws).range(da, da + 999);
    if (error) throw error;
    out.push(...(data || []));
    if (!data || data.length < 1000) return out;
  }
}

export function apiMetaCloud(ws) {
  return {
    async stato() {
      const [account, collegamenti, sync, conn] = await Promise.all([
        tutte("meta_account", ws),
        tutte("meta_collegamenti", ws),
        supabase.from("meta_sync").select("*").eq("workspace_id", ws).order("avvio", { ascending: false }).limit(1),
        supabase.rpc("meta_connessione"),
      ]);
      if (sync.error) throw sync.error;
      /* Prima della migrazione 0003 la funzione non esiste: il collegamento risulta semplicemente assente */
      const facebook = (!conn.error && (conn.data || [])[0]) || null;
      return { account, collegamenti, ultimaSync: (sync.data || [])[0] || null, facebook };
    },
    async dati() {
      const [oggetti, finestre] = await Promise.all([tutte("meta_oggetti", ws), tutte("meta_finestre", ws)]);
      return { oggetti, finestre };
    },
    async collega(accountId, clienteId) {
      const { error } = await supabase.from("meta_collegamenti")
        .upsert({ workspace_id: ws, account_id: String(accountId), cliente_id: clienteId }, { onConflict: "workspace_id,account_id" });
      if (error) throw error;
    },
    async scollega(accountId) {
      const { error } = await supabase.from("meta_collegamenti").delete().eq("workspace_id", ws).eq("account_id", String(accountId));
      if (error) throw error;
    },
    /* Apre la finestra di Facebook; al ritorno il CRM riceve ?meta=collegato (o ?meta=errore&motivo=…) */
    async collegaFacebook() {
      const ritorno = window.location.origin + window.location.pathname;
      const r = await this.sincronizza("login", { ritorno });
      window.location.href = r.url;
    },
    async scollegaFacebook() { await this.sincronizza("scollega"); },
    /* azione "account": solo l'elenco degli account; senza azione l'aggiornamento
       (di tutti i collegati, o solo di extra.account: gli account di un cliente) */
    async sincronizza(azione, extra = {}) {
      const { data, error } = await supabase.functions.invoke("meta-sync", { body: azione ? { azione, ...extra } : { ...extra } });
      if (error) {
        let dettaglio = "";
        try { dettaglio = (await error.context.json()).errore || ""; } catch { /* risposta senza corpo */ }
        throw new Error(dettaglio || (error.message === "Failed to send a request to the Edge Function"
          ? "La funzione meta-sync non risponde: controlla che sia pubblicata in Supabase."
          : error.message));
      }
      return data;
    },
  };
}
