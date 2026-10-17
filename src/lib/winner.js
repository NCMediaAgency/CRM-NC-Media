/* Winner ads: la classifica degli annunci dai numeri salvati dal connettore (tabella analisi_ads).
   Punteggio = media di costo per lead / benchmark, costo per appuntamento / benchmark e
   costo per cliente / benchmark: sotto 1 l'annuncio fa meglio dei riferimenti, più è basso meglio è.
   Funzioni pure, provate con node (tests/winner.test.mjs). */

export const BENCH_BASE = { cpl: 9, app: 20, cliente: 100, minClienti: 5, minSpesa: 0, soloAttive: true, pubblico: "" };
export const benchmarkDi = (cliente) => ({ ...BENCH_BASE, ...((cliente && cliente.winner) || {}) });

const costo = (spesa, n) => (n > 0 ? spesa / n : 0);

/* I numeri di un annuncio con costi, rapporti e punteggio */
export function valutaAd(a, bm) {
  const b = { ...BENCH_BASE, ...bm };
  const spesa = Number(a.spesa) || 0, lead = Number(a.lead) || 0, app = Number(a.appuntamenti) || 0, cli = Number(a.clienti) || 0;
  const cpl = costo(spesa, lead), cpa = costo(spesa, app), cpc = costo(spesa, cli);
  const rapporti = [[cpl, b.cpl], [cpa, b.app], [cpc, b.cliente]].filter(([c, r]) => c > 0 && r > 0).map(([c, r]) => c / r);
  return {
    ...a, spesa, lead, appuntamenti: app, clienti: cli, cpl, cpa, cpc,
    rCpl: cpl && b.cpl ? cpl / b.cpl : 0, rApp: cpa && b.app ? cpa / b.app : 0, rCli: cpc && b.cliente ? cpc / b.cliente : 0,
    chiusura: app ? cli / app : 0,
    punteggio: rapporti.length ? rapporti.reduce((x, y) => x + y, 0) / rapporti.length : 0,
  };
}

/* Gli annunci che rispettano i filtri, dal migliore; esclusi = quanti non li rispettano */
export function classifica(ads, bm) {
  const b = { ...BENCH_BASE, ...bm };
  const tutti = (ads || []).map((a) => valutaAd(a, b));
  const dentro = tutti.filter((a) => (!b.soloAttive || a.attiva !== false) && (!b.pubblico || a.pubblico === b.pubblico)
    && a.clienti >= (Number(b.minClienti) || 0) && a.spesa >= (Number(b.minSpesa) || 0) && a.punteggio > 0);
  dentro.sort((x, y) => x.punteggio - y.punteggio || y.clienti - x.clienti);
  return { righe: dentro.map((a, i) => ({ ...a, pos: i + 1 })), esclusi: tutti.length - dentro.length };
}

/* Le fotografie disponibili (data e fonte), dalla più recente */
export function fotografie(righe) {
  const m = new Map();
  (righe || []).forEach((r) => { const k = `${r.data}|${r.fonte}`; if (!m.has(k)) m.set(k, { data: r.data, fonte: r.fonte, periodo: r.periodo || "", n: 0 }); m.get(k).n++; });
  return [...m.values()].sort((a, b) => (a.data < b.data ? 1 : a.data > b.data ? -1 : a.fonte.localeCompare(b.fonte)));
}

/* Posizione nella fotografia precedente: + vuol dire salita in classifica */
export function conAndamento(attuale, precedente) {
  const prima = new Map((precedente || []).map((a) => [a.ad_id, a.pos]));
  return attuale.map((a) => ({ ...a, primaPos: prima.get(a.ad_id) || 0, salita: prima.has(a.ad_id) ? prima.get(a.ad_id) - a.pos : null }));
}
