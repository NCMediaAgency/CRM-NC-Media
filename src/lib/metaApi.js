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
    /* seguita: l'aggiornamento partito da questa pagina (può non essere l'ultimo, se un altro è in corso) */
    async stato(seguita = null) {
      const [account, collegamenti, sync, conn, mio] = await Promise.all([
        tutte("meta_account", ws),
        tutte("meta_collegamenti", ws),
        supabase.from("meta_sync").select("*").eq("workspace_id", ws).order("avvio", { ascending: false }).limit(1),
        supabase.rpc("meta_connessione"),
        seguita ? supabase.from("meta_sync").select("*").eq("workspace_id", ws).eq("id", seguita).limit(1) : Promise.resolve({ data: [] }),
      ]);
      if (sync.error) throw sync.error;
      /* Prima della migrazione 0003 la funzione non esiste: il collegamento risulta semplicemente assente */
      const facebook = (!conn.error && (conn.data || [])[0]) || null;
      const ultimaSync = (sync.data || [])[0] || null;
      return { account, collegamenti, ultimaSync, facebook, seguita: ((mio && mio.data) || [])[0] || ultimaSync };
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
    /* Regole dell'account (migrazione 0004): campagne_modo, campagne_scelte, automatico */
    async regola(accountId, campi) {
      const { error } = await supabase.from("meta_collegamenti").update(campi).eq("workspace_id", ws).eq("account_id", String(accountId));
      if (error) throw new Error(/campagne_modo|automatico|campagne_scelte/.test(error.message || "")
        ? "Manca la migrazione 0004 in Supabase: eseguila dallo SQL Editor." : error.message);
    },
    /* Le campagne di un account, lette da Meta al momento, per scegliere quali scaricare */
    async campagne(accountId) { return (await this.sincronizza("campagne", { account: String(accountId) })).campagne || []; },
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
