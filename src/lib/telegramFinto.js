/* Solo per la prova locale (VITE_LOCALE=1): simula il bot Telegram nel browser.
   I comandi si mettono in coda scrivendo in localStorage "crm-locale:comandi". */
const K = "crm-locale:telegram", KC = "crm-locale:comandi";
const leggi = (k, d) => { try { return JSON.parse(localStorage.getItem(k)) || d; } catch { return d; } };
export function apiTelegramFinto() {
  return {
    async chat() { return leggi(K, []); },
    async collega() {
      localStorage.setItem(K, JSON.stringify([{ chat_id: 1, utente: "Prova", attivo: true, ora_riepilogo: "09:10", weekend: false, preavviso: 30 }]));
      return { ok: true, bot: "crm_prova_bot", link: "https://t.me/crm_prova_bot?start=prova", anthropic: true };
    },
    async preferenze(chat, campi) { localStorage.setItem(K, JSON.stringify(leggi(K, []).map((c) => (c.chat_id === chat ? { ...c, ...campi } : c)))); return { ok: true }; },
    async scollega(chat) { localStorage.setItem(K, JSON.stringify(leggi(K, []).filter((c) => c.chat_id !== chat))); return { ok: true }; },
    async prova() { return { ok: true, inviati: leggi(K, []).length }; },
    async prendiComandi() { const c = leggi(KC, []); localStorage.setItem(KC, "[]"); return c; },
  };
}
