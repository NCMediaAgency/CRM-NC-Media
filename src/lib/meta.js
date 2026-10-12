/* Dati di Meta Ads dentro il CRM.
   La funzione meta-sync scarica da Meta e scrive nelle tabelle meta_*; qui quei numeri
   diventano campagne, ad set, annunci e finestre del CRM, con le stesse regole
   dell'importazione da file. È una funzione pura: riceve i dati, restituisce i dati nuovi. */

const norm = (s) => String(s || "").trim().toLowerCase();

/* Obiettivo di Meta → obiettivo del CRM */
export const obiettivoDaMeta = (o) => {
  const x = String(o || "").toUpperCase();
  if (/LEAD/.test(x)) return "lead";
  if (/SALES|CONVERSION|CATALOG|PURCHASE/.test(x)) return "vendite";
  if (/TRAFFIC|LINK_CLICK|APP/.test(x)) return "traffico";
  if (/AWARENESS|REACH|ENGAGEMENT|VIDEO|POST|PAGE_LIKES|BRAND/.test(x)) return "awareness";
  return "vendite";
};

/* Account con problemi di pagamento: le campagne accese non stanno davvero erogando */
const ACCOUNT_IN_ERRORE = [3, 9];
export const ACCOUNT_STATI = {
  1: "attivo", 2: "disattivato", 3: "pagamento in sospeso", 7: "in revisione", 8: "in attesa di saldo",
  9: "periodo di tolleranza", 100: "in chiusura", 101: "chiuso", 201: "attivo", 202: "chiuso",
};
export const accountUtilizzabile = (a) => !a || [1, 3, 9, 201].includes(Number(a.stato));

export const statoCampagnaDaMeta = (o, account, attuale, oggi) => {
  const s = String(o.stato || "").toUpperCase();
  if (["ARCHIVED", "DELETED"].includes(s)) return "conclusa";
  if (o.fine && o.fine < oggi) return "conclusa";
  if (s === "ACTIVE" || s === "IN_PROCESS" || s === "WITH_ISSUES") {
    if (account && ACCOUNT_IN_ERRORE.includes(Number(account.stato))) return "errore";
    /* "In test" è una scelta tua: Meta la vede solo come attiva */
    return attuale === "test" ? "test" : "attiva";
  }
  return "pausa";
};
const statoCreativo = (o) => (String(o.stato || "").toUpperCase() === "ACTIVE" ? "attiva" : "pausa");

const numeri = (o) => ({
  spesa: Number(o.spesa) || 0, impression: Number(o.impression) || 0, click: Number(o.click) || 0,
  clickLink: Number(o.click_link) || 0, lead: Number(o.lead) || 0, frequenza: Number(o.frequenza) || 0,
});
const esito = (o) => ({ conversioni: Number(o.conversioni) || 0, ricavi: Number(o.ricavi) || 0 });

/* Suggerimenti per collegare un account a un cliente: parole in comune, nome contenuto, iniziali (BDF = Boutique Del Fitness) */
const VUOTE = new Set(["studio", "legale", "avv", "avvocato", "avvocata", "srl", "srls", "spa", "snc", "sas", "ads", "ad", "account", "di", "del", "della", "dei", "de", "e", "the", "and", "s", "il", "la", "lo", "le"]);
const parole = (s) => norm(s).normalize("NFD").replace(/[̀-ͯ]/g, "").split(/[^a-z0-9]+/).filter(Boolean);
export const somiglianza = (cliente, account) => {
  const pc = parole(cliente), pa = parole(account);
  const sc = pc.filter((p) => !VUOTE.has(p) && p.length > 2), sa = new Set(pa.filter((p) => !VUOTE.has(p) && p.length > 2));
  if (!pc.length || !pa.length) return 0;
  let punti = 0;
  const comuni = sc.filter((p) => sa.has(p) || [...sa].some((q) => q.length > 3 && p.length > 3 && (q.startsWith(p) || p.startsWith(q))));
  if (sc.length) punti += comuni.length / sc.length;
  const iniziali = pc.map((p) => p[0]).join("");
  if (iniziali.length >= 2 && pa.some((p) => p.replace(/\d+$/, "") === iniziali)) punti += 0.9;
  const a = parole(account).join(" "), c = parole(cliente).join(" ");
  if (c.length > 3 && (a.includes(c) || c.includes(a))) punti += 0.5;
  return Math.min(1.5, punti);
};
export const suggerisciAccount = (cliente, accounts, minimo = 0.34) => accounts
  .map((a) => ({ a, p: Math.max(somiglianza(cliente.nome, a.nome), somiglianza(cliente.nome, a.business || "") * 0.9) }))
  .filter((x) => x.p >= minimo)
  .sort((x, y) => y.p - x.p)
  .map((x) => x.a);

/* ---------- applicazione ----------
   meta: { collegamenti:[{account_id, cliente_id}], account:[...], oggetti:[...], finestre:[...] }
   opz:  { oggi: "AAAA-MM-GG", uid: () => id, bmkDi: (obiettivo) => benchmark, valutaAgenzia } */
export function applicaMeta(data, meta, opz) {
  const { oggi, uid, bmkDi, valutaAgenzia = "EUR" } = opz;
  const cliDiAccount = new Map((meta.collegamenti || []).map((c) => [String(c.account_id), c.cliente_id]));
  const idClienti = new Set((data.clienti || []).map((z) => z.id));
  const accDi = new Map((meta.account || []).map((a) => [String(a.account_id), a]));
  const conta = { campagneNuove: 0, campagneAggiornate: 0, creativi: 0, finestre: 0 };

  const ogg = (meta.oggetti || []).filter((o) => idClienti.has(cliDiAccount.get(String(o.account_id))));
  if (!ogg.length) return { data, conta, cambiato: false };

  let campagne = [...data.campagne];
  let creativi = [...(data.creativi || [])];
  let periodi = [...(data.periodi || [])];
  let clienti = data.clienti;
  const finestreDi = new Map();
  (meta.finestre || []).forEach((f) => { const l = finestreDi.get(f.campagna_id) || []; l.push(f); finestreDi.set(f.campagna_id, l); });

  /* Valuta del cliente: se non l'hai indicata e l'account di Meta è in un'altra valuta, prendo quella */
  const valutePerCli = new Map();
  (meta.collegamenti || []).forEach((c) => {
    const a = accDi.get(String(c.account_id));
    if (!a || !a.valuta) return;
    const s = valutePerCli.get(c.cliente_id) || new Set(); s.add(a.valuta); valutePerCli.set(c.cliente_id, s);
  });
  clienti = clienti.map((z) => {
    const v = valutePerCli.get(z.id);
    if (z.valuta || !v || v.size !== 1) return z;
    const unica = [...v][0];
    return unica !== valutaAgenzia ? { ...z, valuta: unica } : z;
  });

  /* Campagne */
  const crmDiCampagna = new Map();
  ogg.filter((o) => o.livello === "campaign").forEach((o) => {
    const cid = cliDiAccount.get(String(o.account_id));
    const acc = accDi.get(String(o.account_id));
    let i = campagne.findIndex((k) => k.metaId === o.id);
    if (i < 0) i = campagne.findIndex((k) => !k.metaId && k.clienteId === cid && norm(k.nome) === norm(o.nome));
    const vecchia = i > -1 ? campagne[i] : null;
    const stato = statoCampagnaDaMeta(o, acc, vecchia && vecchia.stato, oggi);
    const fin = (finestreDi.get(o.id) || []).slice().sort((a, b) => (a.inizio < b.inizio ? -1 : 1));
    /* Spenta senza data di fine: la chiudo sull'ultima finestra con spesa, così i giorni non crescono a vuoto */
    let fine = o.fine || "";
    if (!fine && stato !== "attiva" && stato !== "test" && stato !== "errore") {
      const ultima = [...fin].reverse().find((f) => Number(f.spesa) > 0);
      fine = ultima ? ultima.fine : (vecchia && vecchia.fine) || "";
    }
    const manuale = vecchia && vecchia.esitoManuale === "si";
    const campi = {
      metaId: o.id, metaAccount: String(o.account_id), nome: o.nome || (vecchia && vecchia.nome) || "Campagna Meta",
      piattaforma: "meta", stato, inizio: o.inizio || (vecchia && vecchia.inizio) || oggi, fine,
      ...numeri(o), ...(manuale ? {} : esito(o)),
    };
    if (Number(o.budget_totale) > 0) campi.budget = Number(o.budget_totale);
    /* Budget giornaliero: ogni cambio entra nello storico, come quando lo cambi dalla scheda */
    const bg = Number(o.budget_giorno) || 0;
    const storico = [...((vecchia && vecchia.storicoBudget) || [])];
    if (bg > 0) {
      const ultimo = storico.length ? storico[storico.length - 1].importo : (vecchia && vecchia.budgetGiorno) || 0;
      if (!storico.length) storico.push({ id: uid(), data: campi.inizio, importo: bg, nota: "budget iniziale (da Meta)" });
      else if (ultimo !== bg) {
        const j = storico.findIndex((b) => b.data === oggi);
        if (j > -1) storico[j] = { ...storico[j], importo: bg };
        else storico.push({ id: uid(), data: oggi, importo: bg, nota: "aggiornato da Meta" });
      }
      campi.storicoBudget = storico;
      campi.budgetGiorno = bg;
    }
    /* Traccia del fermo per pagamento, come il salvataggio della scheda */
    if (stato === "errore" && !(vecchia && vecchia.erroreDal)) campi.erroreDal = oggi;
    if (stato !== "errore") campi.erroreDal = "";

    if (vecchia) {
      const nuova = { ...vecchia, ...campi };
      if (JSON.stringify(nuova) !== JSON.stringify(vecchia)) { campagne[i] = nuova; conta.campagneAggiornate++; }
      crmDiCampagna.set(o.id, vecchia);
    } else {
      const obiettivo = obiettivoDaMeta(o.obiettivo);
      const k = {
        id: uid(), clienteId: cid, linea: "", obiettivo, budgetGiorno: 0, budget: 0,
        valoreContratto: 0, incassato: 0, archiviata: "no", erroreDal: "", modificaVista: "no", esitoManuale: "",
        storicoBudget: [], conversioni: 0, ricavi: 0, bmk: { ...bmkDi(obiettivo) }, ...campi,
      };
      campagne = [k, ...campagne];
      crmDiCampagna.set(o.id, k);
      conta.campagneNuove++;
    }
  });

  /* Ad set e annunci: prima gli ad set, così gli annunci trovano il padre */
  const crmDiCreativo = new Map();
  ["adset", "ad"].forEach((livello) => {
    ogg.filter((o) => o.livello === livello).forEach((o) => {
      const k = crmDiCampagna.get(o.campagna_id);
      if (!k) return;
      const tipo = livello === "adset" ? "adset" : "annuncio";
      let i = creativi.findIndex((c) => c.metaId === o.id);
      if (i < 0) i = creativi.findIndex((c) => !c.metaId && c.campagnaId === k.id && c.tipo === tipo && norm(c.nome) === norm(o.nome));
      const padre = livello === "ad" ? crmDiCreativo.get(o.adset_id) : null;
      const campi = { metaId: o.id, tipo, nome: o.nome, stato: statoCreativo(o), ...numeri(o), ...esito(o), ...(padre ? { adsetId: padre.id } : {}) };
      if (i > -1) {
        const nuovo = { ...creativi[i], ...campi };
        if (JSON.stringify(nuovo) !== JSON.stringify(creativi[i])) { creativi[i] = nuovo; conta.creativi++; }
        crmDiCreativo.set(o.id, creativi[i]);
      } else {
        const c = { id: uid(), clienteId: k.clienteId, campagnaId: k.id, adsetId: "", stadio: "", note: "", valoreContratto: 0, incassato: 0, ...campi };
        creativi = [c, ...creativi];
        crmDiCreativo.set(o.id, c);
        conta.creativi++;
      }
    });
  });

  /* Finestre: quelle di Meta prendono il posto delle finestre scritte a mano che si sovrappongono */
  crmDiCampagna.forEach((k, metaId) => {
    const fin = finestreDi.get(metaId) || [];
    if (!fin.length) return;
    const da = fin.reduce((m, f) => (f.inizio < m ? f.inizio : m), fin[0].inizio);
    const a = fin.reduce((m, f) => (f.fine > m ? f.fine : m), fin[0].fine);
    const inizi = new Set(fin.map((f) => f.inizio));
    const prima = periodi.length;
    periodi = periodi.filter((pz) => pz.campagnaId !== k.id || inizi.has(pz.inizio) || pz.fine < da || pz.inizio > a);
    if (periodi.length !== prima) conta.finestre += prima - periodi.length;
    fin.forEach((f) => {
      const i = periodi.findIndex((pz) => pz.campagnaId === k.id && pz.inizio === f.inizio);
      const campi = { tipo: f.tipo, inizio: f.inizio, fine: f.fine, origine: "meta", ...numeri(f), ...(k.esitoManuale === "si" ? {} : esito(f)) };
      if (i > -1) {
        const nuovo = { ...periodi[i], ...campi };
        if (JSON.stringify(nuovo) !== JSON.stringify(periodi[i])) { periodi[i] = nuovo; conta.finestre++; }
      } else {
        periodi.push({ id: uid(), clienteId: k.clienteId, campagnaId: k.id, valoreContratto: 0, incassato: 0, note: "", ...campi });
        conta.finestre++;
      }
    });
  });

  const cambiato = conta.campagneNuove + conta.campagneAggiornate + conta.creativi + conta.finestre > 0 || clienti !== data.clienti && clienti.some((z, i) => z !== data.clienti[i]);
  if (!cambiato) return { data, conta, cambiato: false };
  return { data: { ...data, clienti, campagne, creativi, periodi }, conta, cambiato: true };
}
