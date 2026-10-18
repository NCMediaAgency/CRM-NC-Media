import { supabase } from "./supabase";

/* Impostazioni dello spazio di lavoro che servono anche al server (migrazione 0010):
   fuso orario e ora dell'aggiornamento di Meta. Senza riga valgono Roma e le 9. */
export const fusoDelBrowser = () => { try { return Intl.DateTimeFormat().resolvedOptions().timeZone || "Europe/Rome"; } catch { return "Europe/Rome"; } };

export function apiSpazioCloud(ws) {
  return {
    async leggi() {
      const { data, error } = await supabase.from("workspace_impostazioni").select("fuso,ora_meta").eq("workspace_id", ws).maybeSingle();
      if (error) throw error;
      return data;   // null = mai impostate
    },
    async salva(campi) {
      const { error } = await supabase.from("workspace_impostazioni")
        .upsert({ workspace_id: ws, ...campi, aggiornato: new Date().toISOString() }, { onConflict: "workspace_id" });
      if (error) throw error;
    },
  };
}

/* Prova locale: nel browser */
const K = "crm-locale:spazio";
export function apiSpazioFinto() {
  return {
    async leggi() { try { return JSON.parse(localStorage.getItem(K)); } catch { return null; } },
    async salva(campi) { let v = {}; try { v = JSON.parse(localStorage.getItem(K)) || {}; } catch { /* vuoto */ } localStorage.setItem(K, JSON.stringify({ ...v, ...campi })); },
  };
}
