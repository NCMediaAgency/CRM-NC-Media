import { supabase } from "./supabase";

/* Il CRM è nato dentro Claude e salva con window.storage (get/set/list/delete).
   Qui la stessa interfaccia scrive sulla tabella kv_store del database,
   dentro lo spazio di lavoro dell'utente: il codice del CRM non cambia. */
export function archivioCloud(workspaceId) {
  const tab = () => supabase.from("kv_store");
  return {
    async get(key) {
      const { data, error } = await tab().select("value").eq("workspace_id", workspaceId).eq("key", key).maybeSingle();
      if (error) throw error;
      return data ? { key, value: data.value } : null;
    },
    async set(key, value) {
      const { error } = await tab().upsert(
        { workspace_id: workspaceId, key, value, updated_at: new Date().toISOString() },
        { onConflict: "workspace_id,key" },
      );
      if (error) throw error;
      return { key, value };
    },
    async delete(key) {
      const { error } = await tab().delete().eq("workspace_id", workspaceId).eq("key", key);
      if (error) throw error;
      return { key, deleted: true };
    },
    async list(prefix = "") {
      const filtro = prefix.replace(/[\\%_]/g, (c) => "\\" + c) + "%";
      const { data, error } = await tab().select("key").eq("workspace_id", workspaceId).like("key", filtro);
      if (error) throw error;
      return { keys: (data || []).map((r) => r.key), prefix };
    },
  };
}

/* Solo in sviluppo: archivio nel browser, per provare l'interfaccia senza database */
export function archivioLocale() {
  const P = "crm-locale:";
  return {
    async get(key) { const v = localStorage.getItem(P + key); return v === null ? null : { key, value: v }; },
    async set(key, value) { localStorage.setItem(P + key, value); return { key, value }; },
    async delete(key) { localStorage.removeItem(P + key); return { key, deleted: true }; },
    async list(prefix = "") {
      const keys = Object.keys(localStorage).filter((k) => k.startsWith(P + prefix)).map((k) => k.slice(P.length));
      return { keys, prefix };
    },
  };
}
