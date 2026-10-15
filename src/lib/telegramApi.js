import { supabase } from "./supabase";

/* Il CRM e il bot Telegram: collegamento e preferenze passano dalla funzione telegram,
   i comandi in coda (crm_comandi) si leggono e si segnano direttamente. */
export function apiTelegramCloud(ws) {
  const funzione = async (corpo) => {
    const { data, error } = await supabase.functions.invoke("telegram", { body: corpo });
    if (error) {
      let dettaglio = "";
      try { dettaglio = (await error.context.json()).errore || ""; } catch { /* risposta senza corpo */ }
      throw new Error(dettaglio || (error.message === "Failed to send a request to the Edge Function"
        ? "La funzione telegram non risponde: controlla che sia pubblicata in Supabase." : error.message));
    }
    return data;
  };
  return {
    /* Chat collegate (prima della migrazione 0006 la tabella non esiste: nessuna chat) */
    async chat() {
      const { data, error } = await supabase.from("telegram_chat").select("*").eq("workspace_id", ws);
      return error ? [] : data || [];
    },
    collega: () => funzione({ azione: "collega" }),
    preferenze: (chat, campi) => funzione({ azione: "preferenze", chat, ...campi }),
    scollega: (chat) => funzione({ azione: "scollega", chat }),
    prova: () => funzione({ azione: "prova" }),
    /* Prende i comandi in attesa segnandoli come applicati in un colpo solo: se il CRM è aperto
       in due schede, ogni comando lo prende una sola delle due */
    async prendiComandi() {
      const { data: attesa, error } = await supabase.from("crm_comandi").select("id").eq("workspace_id", ws).is("applicato", null).order("creato").limit(50);
      if (error || !attesa || !attesa.length) return [];
      const { data, error: e2 } = await supabase.from("crm_comandi").update({ applicato: new Date().toISOString() })
        .in("id", attesa.map((c) => c.id)).is("applicato", null).select("id,tipo,dati,creato");
      if (e2) return [];
      return (data || []).sort((a, b) => (a.creato < b.creato ? -1 : 1));
    },
  };
}
