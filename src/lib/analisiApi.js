import { supabase } from "./supabase";

/* Analisi degli annunci salvate dal connettore per Claude (migrazione 0008): il CRM le legge soltanto */
export function apiAnalisiCloud(ws) {
  return {
    async ads(clienti) {
      if (!clienti.length) return [];
      const { data, error } = await supabase.from("analisi_ads").select("*").eq("workspace_id", ws).in("cliente_id", clienti)
        .order("data", { ascending: false }).limit(5000);
      return error ? [] : data || [];
    },
    async testi(clienti) {
      if (!clienti.length) return [];
      const { data, error } = await supabase.from("analisi_testi").select("*").eq("workspace_id", ws).in("cliente_id", clienti)
        .order("creato", { ascending: false }).limit(50);
      return error ? [] : data || [];
    },
    async eliminaTesto(id) {
      const { error } = await supabase.from("analisi_testi").delete().eq("id", id);
      if (error) throw error;
    },
  };
}

/* Prova locale: le righe si mettono in localStorage "crm-locale:analisi_ads" e "crm-locale:analisi_testi" */
const leggi = (k) => { try { return JSON.parse(localStorage.getItem(k)) || []; } catch { return []; } };
export function apiAnalisiFinte() {
  return {
    async ads(clienti) { return leggi("crm-locale:analisi_ads").filter((r) => clienti.includes(r.cliente_id)); },
    async testi(clienti) { return leggi("crm-locale:analisi_testi").filter((r) => clienti.includes(r.cliente_id)).sort((a, b) => (a.creato < b.creato ? 1 : -1)); },
    async eliminaTesto(id) { localStorage.setItem("crm-locale:analisi_testi", JSON.stringify(leggi("crm-locale:analisi_testi").filter((r) => r.id !== id))); },
  };
}
