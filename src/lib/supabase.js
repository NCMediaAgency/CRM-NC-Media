import { createClient } from "@supabase/supabase-js";

/* URL e chiave "publishable" sono pubbliche per costruzione: stanno nel sito.
   I dati sono protetti dalle regole di accesso del database (RLS), non da queste. */
const URL = import.meta.env.VITE_SUPABASE_URL || "https://cggxgzazjyeayfhynhhd.supabase.co";
const CHIAVE = import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY || "sb_publishable_Qv2YEupaxLL5cpjoRY0dXA_cFe0IbPN";

export const supabase = createClient(URL, CHIAVE, {
  auth: { persistSession: true, autoRefreshToken: true },
});
