/* Solo per la prova locale (VITE_LOCALE=1): simula Meta e le tabelle meta_*
   nel browser, così l'interfaccia si prova senza token e senza database. */
const K = "crm-locale:meta";
const leggi = () => { try { return JSON.parse(localStorage.getItem(K)) || {}; } catch { return {}; } };
const scrivi = (s) => localStorage.setItem(K, JSON.stringify(s));
const giorno = (n) => new Date(Date.now() + n * 864e5).toISOString().slice(0, 10);
const somma = (s, n) => { const d = new Date(s + "T00:00:00Z"); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); };

const ACCOUNT = [
  { account_id: "569452278716206", nome: "BDF 1", valuta: "USD", stato: 1, business: "Boutique Del Fitness" },
  { account_id: "111", nome: "Cantine del Chienti — Shop", valuta: "EUR", stato: 1, business: "Cantine del Chienti" },
  { account_id: "222", nome: "Tecnoplast ADV", valuta: "EUR", stato: 3, business: "Tecnoplast Marche" },
  { account_id: "333", nome: "Studio Legale Campilongo", valuta: "EUR", stato: 1, business: "Studio Legale Campilongo" },
  { account_id: "444", nome: "Vecchio account", valuta: "EUR", stato: 101, business: "" },
];

function generaDati(collegamenti) {
  const oggetti = [], finestre = [];
  collegamenti.forEach(({ account_id }, j) => {
    const c = `c${account_id}`;
    const inizio = giorno(-24 - j * 3);
    oggetti.push({ livello: "campaign", id: c, account_id, nome: `Lead · ${account_id.slice(-3)}`, stato: "ACTIVE", obiettivo: "OUTCOME_LEADS",
      budget_giorno: 60, budget_totale: 0, inizio, fine: null, spesa: 1380.5, impression: 98000, click: 2100, click_link: 1300, lead: 92, conversioni: 0, ricavi: 0, frequenza: 2.1 });
    oggetti.push({ livello: "adset", id: `s${account_id}`, account_id, campagna_id: c, nome: "Broad 25-55", stato: "ACTIVE", spesa: 1380.5, impression: 98000, click: 2100, click_link: 1300, lead: 92, frequenza: 2.1 });
    oggetti.push({ livello: "ad", id: `a${account_id}`, account_id, campagna_id: c, adset_id: `s${account_id}`, nome: "Video 1", stato: "ACTIVE", spesa: 900, impression: 61000, click: 1300, click_link: 800, lead: 60, frequenza: 1.9 });
    const ieri = giorno(-1);
    const agg = (tipo, dal, al) => { if (dal <= ieri) finestre.push({ campagna_id: c, account_id, tipo, inizio: dal, fine: al < ieri ? al : ieri, spesa: 410.2, impression: 28000, click: 600, click_link: 380, lead: 27, conversioni: 0, ricavi: 0, frequenza: 1.4 }); };
    agg("test", inizio, somma(inizio, 2)); agg("matura", somma(inizio, 3), somma(inizio, 9));
    for (let dal = somma(inizio, 10); dal <= ieri; dal = somma(dal, 7)) agg("attiva", dal, somma(dal, 6));
  });
  return { oggetti, finestre };
}

export function apiMetaFinta() {
  return {
    async stato() {
      const s = leggi();
      return { account: s.account || [], collegamenti: s.collegamenti || [], ultimaSync: s.ultimaSync || null, facebook: s.facebook || null };
    },
    async collegaFacebook() {
      const s = leggi(); s.facebook = { collegato: true, utente_meta: "Profilo di prova", scadenza: new Date(Date.now() + 60 * 864e5).toISOString() }; scrivi(s);
      window.location.href = window.location.pathname + "?meta=collegato";
    },
    async scollegaFacebook() { const s = leggi(); delete s.facebook; scrivi(s); },
    async dati() { const s = leggi(); return { oggetti: s.oggetti || [], finestre: s.finestre || [] }; },
    async collega(accountId, clienteId) {
      const s = leggi(); s.collegamenti = (s.collegamenti || []).filter((c) => c.account_id !== accountId).concat({ account_id: accountId, cliente_id: clienteId }); scrivi(s);
    },
    async scollega(accountId) { const s = leggi(); s.collegamenti = (s.collegamenti || []).filter((c) => c.account_id !== accountId); scrivi(s); },
    async sincronizza(azione, extra = {}) {
      const s = leggi();
      if (!s.facebook) throw new Error("Facebook non è collegato: premi \"Collega Facebook\".");
      const prima = new Map((s.account || []).map((a) => [a.account_id, a]));
      const tocca = (a) => !azione && (!extra.account || extra.account.includes(a.account_id));
      s.account = ACCOUNT.map((a) => ({ ...a,
        ultimo_sync: tocca(a) ? new Date().toISOString() : (prima.get(a.account_id) || {}).ultimo_sync || null,
        ultimo_errore: tocca(a) && a.stato === 3 ? "Pagamento in sospeso sull'account" : (prima.get(a.account_id) || {}).ultimo_errore || "" }));
      if (!azione) {
        Object.assign(s, generaDati(s.collegamenti || []));
        s.ultimaSync = { esito: "in_corso", avvio: new Date().toISOString(), origine: "manuale", messaggio: "" };
        scrivi(s);
        setTimeout(() => { const t = leggi(); t.ultimaSync = { ...t.ultimaSync, esito: "ok", fine: new Date().toISOString(), messaggio: `${extra.account ? extra.account.length : (t.collegamenti || []).length} account aggiornati.` }; scrivi(t); }, 1500);
        return { ok: true };
      }
      scrivi(s);
      return { ok: true, account: ACCOUNT.length };
    },
  };
}
