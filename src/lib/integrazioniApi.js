import { supabase } from "./supabase";

/* Integrazioni con software esterni (Boschetto): creazione, chiave, indirizzo di ritorno
   e invio dei cambi fatti nel CRM. Tutto passa dalla funzione integrazioni: le chiavi
   non sono leggibili dal browser. */
export function apiIntegrazioniCloud() {
  const funzione = async (corpo) => {
    const { data, error } = await supabase.functions.invoke("integrazioni", { body: corpo });
    if (error) {
      let dettaglio = "";
      try { dettaglio = (await error.context.json()).errore || ""; } catch { /* risposta senza corpo */ }
      throw new Error(dettaglio || (error.message === "Failed to send a request to the Edge Function"
        ? "La funzione integrazioni non risponde: controlla che sia pubblicata in Supabase." : error.message));
    }
    return data;
  };
  return {
    stato: () => funzione({ azione: "stato" }),
    crea: (nome) => funzione({ azione: "crea", nome }),
    rigenera: (id) => funzione({ azione: "rigenera", id }),
    ritorno: (id, url, chiave) => funzione({ azione: "ritorno", id, url, ...(chiave !== undefined ? { chiave } : {}) }),
    attiva: (id, attiva) => funzione({ azione: "attiva", id, attiva }),
    elimina: (id) => funzione({ azione: "elimina", id }),
    notifica: (id, modifiche) => funzione({ azione: "notifica", id, modifiche }),
  };
}

/* Solo per la prova locale (VITE_LOCALE=1). Una task "in arrivo" si simula scrivendo in
   localStorage "crm-locale:comandi" un comando importa_task; i cambi rimandati indietro
   finiscono in "crm-locale:uscite". */
const K = "crm-locale:integrazioni", KU = "crm-locale:uscite";
const leggi = (k, d) => { try { return JSON.parse(localStorage.getItem(k)) || d; } catch { return d; } };
const scrivi = (k, v) => localStorage.setItem(k, JSON.stringify(v));
export function apiIntegrazioniFinte() {
  const indirizzo = "https://esempio.supabase.co/functions/v1/integrazioni";
  const cambia = (id, campi) => scrivi(K, leggi(K, []).map((x) => (x.id === id ? { ...x, ...campi } : x)));
  return {
    async stato() { return { ok: true, indirizzo, integrazioni: leggi(K, []) }; },
    async crea(nome) {
      const id = "int-" + Math.random().toString(36).slice(2, 10), chiave = "crmnc_prova" + Math.random().toString(36).slice(2, 14);
      scrivi(K, [...leggi(K, []), { id, nome, chiave_fine: chiave.slice(-4), url_ritorno: "", ritorno: false, attiva: true, creato: new Date().toISOString(), ultimo_contatto: null }]);
      return { ok: true, id, chiave, indirizzo };
    },
    async rigenera(id) { const chiave = "crmnc_prova" + Math.random().toString(36).slice(2, 14); cambia(id, { chiave_fine: chiave.slice(-4) }); return { ok: true, chiave, indirizzo }; },
    async ritorno(id, url) { cambia(id, { url_ritorno: url, ritorno: !!url }); return { ok: true }; },
    async attiva(id, attiva) { cambia(id, { attiva }); return { ok: true }; },
    async elimina(id) { scrivi(K, leggi(K, []).filter((x) => x.id !== id)); return { ok: true }; },
    async notifica(id, modifiche) { scrivi(KU, [...leggi(KU, []), ...modifiche.map((m) => ({ integrazione: id, ...m }))]); return { ok: true, registrate: modifiche.length }; },
  };
}
