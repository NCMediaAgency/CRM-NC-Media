/* Lettura dei file KPI dei clienti (Excel) per il Ciclo di vendita.
   Riconosce il formato più diffuso nei fogli di riepilogo: le voci in colonna
   (Spesa, Lead, Appuntamenti fissati, Nuovi clienti, Fatturato…) e i mesi in riga,
   eventualmente con una sottocolonna per funnel sotto ogni mese
   (es. GENNAIO 2026 → FUNNEL OT | BOOST REELS | TOTALE).
   Funzioni pure: nessun accesso a React o all'archivio, si provano con node. */

const MESI = ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio", "agosto", "settembre", "ottobre", "novembre", "dicembre"];
const MESI_EN = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];

const norm = (s) => String(s ?? "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "")
  .replace(/[^a-z0-9%+/ ]/g, " ").replace(/\s+/g, " ").trim();

/* "GENNAIO 2026", "Febbraio", "MAR 2025" → { m: 1..12, anno|null }; altrimenti null */
export function leggiMese(cella) {
  const t = norm(cella);
  if (!t || t.length > 24) return null;
  const parole = t.split(" ");
  const p = parole[0];
  let m = MESI.indexOf(p) + 1 || MESI_EN.indexOf(p) + 1;
  if (!m && p.length >= 3) m = MESI.findIndex((x) => x.startsWith(p) && p.length >= 3) + 1;
  if (!m) return null;
  // Dopo il mese al massimo un anno: "aprile (da 15)" o frasi lunghe non sono intestazioni di mese
  const resto = parole.slice(1);
  if (resto.length > 1 || (resto.length === 1 && !/^(19|20)\d\d$/.test(resto[0]))) return null;
  return { m, anno: resto.length ? Number(resto[0]) : null };
}

/* Numero da una cella: numeri veri di Excel, oppure testi come "€ 47,872.79", "6.096", "1.350,00".
   "/", "-", errori (#DIV/0!) e date restano vuoti. */
export function numeroCella(v) {
  if (typeof v === "number") return isFinite(v) ? v : null;
  if (v instanceof Date || v === null || v === undefined) return null;
  let t = String(v).trim();
  if (!t || /^#|^[/\-–—]$/.test(t) || /\d+\/\d+\/\d+/.test(t)) return null;
  t = t.replace(/[\s '’€$£]/g, "");
  if (/%$/.test(t)) return null;
  if (!/^-?[\d.,]+$/.test(t)) return null;
  const v1 = t.lastIndexOf(","), p1 = t.lastIndexOf(".");
  if (v1 > -1 && p1 > -1) t = v1 > p1 ? t.replace(/\./g, "").replace(",", ".") : t.replace(/,/g, "");
  else if (v1 > -1) t = /,\d{3}$/.test(t) && t.split(",").length > 1 && !/,\d{1,2}$/.test(t) ? t.replace(/,/g, "") : t.replace(",", ".");
  else if (p1 > -1 && /^\d{1,3}(\.\d{3})+$/.test(t)) t = t.replace(/\./g, "");
  const n = parseFloat(t);
  return isFinite(n) ? n : null;
}

/* Voci riconosciute → campi del Ciclo di vendita. Fra più righe candidate nello stesso blocco
   vince quella con "totale"; senza, se sono varianti per canale (IG, WP…) si sommano. */
export const VOCI_KPI = [
  { k: "spesa", b: "Spesa", l: "Spesa", re: /^spesa( tot(ale)?)?( .*)?$/, no: /solo|conv/ },
  { k: "lead", b: "Lead", l: "Lead", re: /^(numero )?leads?( .*)?$/, no: /effettiv|fake|target/ },
  { k: "demoFissate", b: "App. fissati", l: "Appuntamenti fissati", re: /^(appuntamenti|apt|app) fissati/ },
  { k: "demoSvolte", b: "App. svolti", l: "Appuntamenti svolti", re: /^(appuntamenti|apt|app) (effettivi|svolti|effettivamente svolti)/ },
  { k: "chiusure", b: "Chiusure", l: "Chiusure (nuovi clienti)", re: /^(nuovi )?clienti( .*)?$/ },
  { k: "valoreContratto", b: "Contratti", l: "Valore contratti (fatturato nuovi clienti)", re: /^(fatturato nuovi clienti|cash ?\+ ?rate totale)/ },
  { k: "incassato", b: "Incassato", l: "Incassato (nuovi clienti)", re: /^incassato nuovi clienti/ },
];

const etichetta = (riga, finoA) => {
  for (let j = 0; j < finoA; j++) { const t = String(riga[j] ?? "").trim(); if (t) return t; }
  return "";
};
const mesiDellaRiga = (riga) => riga.map((c, j) => ({ j, x: leggiMese(c) })).filter((x) => x.x);

/* fogli: [{ nome, righe: matrice di celle (sheet_to_json con header:1, raw:true) }]
   → tabelle: [{ id, foglio, gruppo, mesi: [{ periodo|null, m, anno, valori: { campo: numero } }], annoMancante }] */
export function trovaTabelleKpi(fogli) {
  const tabelle = new Map();
  for (const { nome, righe } of fogli) {
    const annoFoglio = (String(nome).match(/\b(19|20)\d\d\b/) || [])[0];
    let i = 0;
    while (i < righe.length) {
      const testata = mesiDellaRiga(righe[i] || []);
      if (testata.length < 2) { i++; continue; }
      // Colonne di ogni mese: dalla sua fino a quella del mese successivo
      const c0 = testata[0].j;
      const titolo = etichetta(righe[i], c0);
      // Riga dei gruppi (funnel) sotto i mesi: senza etichetta a sinistra e con testi sotto i mesi.
      // Può non essere subito sotto (es. una riga "FATTURATO MESE" in mezzo): la cerco nelle tre righe seguenti
      const eGruppi = (rr) => !etichetta(rr, c0) && testata.some((t, n) => {
        const fine = n + 1 < testata.length ? testata[n + 1].j : t.j + 8;
        return rr.slice(t.j, fine).some((c) => typeof c === "string" && c.trim() && numeroCella(c) === null && !leggiMese(c));
      });
      let posGruppi = -1;
      for (let k = 1; k <= 3 && i + k < righe.length; k++) {
        if (mesiDellaRiga(righe[i + k] || []).length >= 2) break;
        if (eGruppi(righe[i + k] || [])) { posGruppi = i + k; break; }
      }
      const conGruppi = posGruppi > -1;
      const sotto = conGruppi ? righe[posGruppi] : [];
      const colonne = []; // { j, mese, gruppo }
      testata.forEach((t, n) => {
        const fine = n + 1 < testata.length ? testata[n + 1].j : (conGruppi ? Math.max(t.j + 1, (sotto.length || t.j + 1)) : t.j + 1);
        if (!conGruppi) { colonne.push({ j: t.j, mese: t.x, gruppo: titolo || nome }); return; }
        for (let j = t.j; j < fine; j++) {
          const g = String(sotto[j] ?? "").trim();
          if (g && numeroCella(g) === null) colonne.push({ j, mese: t.x, gruppo: g });
        }
      });
      // Righe dei dati: fino alla prossima riga di mesi
      const dati = [];
      let r = conGruppi ? posGruppi + 1 : i + 1;
      for (; r < righe.length; r++) {
        if (mesiDellaRiga(righe[r] || []).length >= 2) break;
        const lab = etichetta(righe[r] || [], c0);
        if (lab) dati.push({ lab: norm(lab), riga: righe[r] });
      }
      for (const col of colonne) {
        const anno = col.mese.anno || (annoFoglio ? Number(annoFoglio) : null);
        const valori = {};
        for (const v of VOCI_KPI) {
          const cand = dati.filter((d) => !d.lab.startsWith("%") && v.re.test(d.lab) && !(v.no && v.no.test(d.lab)))
            .map((d) => ({ lab: d.lab, n: numeroCella(d.riga[col.j]) })).filter((d) => d.n !== null);
          if (!cand.length) continue;
          const tot = cand.find((d) => /\btotal[ei]\b/.test(d.lab));
          // Varianti per canale senza riga di totale (es. INCASSATO NUOVI CLIENTI WP + IG): si sommano
          const canali = cand.filter((d) => /\b(ig|wp|instagram|whatsapp)\b/.test(d.lab));
          valori[v.k] = tot ? tot.n : canali.length > 1 ? canali.reduce((s, d) => s + d.n, 0) : cand[0].n;
        }
        // Mesi ancora vuoti (tutti zero) non si importano
        if (!Object.values(valori).some((n) => n)) continue;
        const gruppo = col.gruppo.replace(/\s+/g, " ");
        const id = `${nome}|${norm(gruppo)}`;
        if (!tabelle.has(id)) tabelle.set(id, { id, foglio: nome, gruppo, mesi: [] });
        const t = tabelle.get(id);
        const periodo = anno ? `${anno}-${String(col.mese.m).padStart(2, "0")}` : null;
        const gia = t.mesi.find((x) => x.m === col.mese.m && x.anno === anno);
        if (gia) Object.assign(gia.valori, valori); else t.mesi.push({ periodo, m: col.mese.m, anno, valori });
      }
      i = r;
    }
  }
  return [...tabelle.values()].map((t) => ({
    ...t,
    mesi: t.mesi.sort((a, b) => (a.anno || 0) - (b.anno || 0) || a.m - b.m),
    annoMancante: t.mesi.some((x) => !x.anno),
  })).sort((a, b) => b.mesi.length - a.mesi.length || (/totale/i.test(b.gruppo) ? 1 : 0) - (/totale/i.test(a.gruppo) ? 1 : 0));
}

/* Scrive i mesi nel Ciclo di vendita del cliente: aggiorna solo i campi presenti nel file,
   crea il periodo se manca. Restituisce { vendite, nuovi, aggiornati }. */
export function applicaKpi(vendite, clienteId, mesi, { uid, annoPredefinito, fonte = "" }) {
  const out = [...vendite];
  let nuovi = 0, aggiornati = 0;
  for (const x of mesi) {
    const periodo = x.periodo || (annoPredefinito ? `${annoPredefinito}-${String(x.m).padStart(2, "0")}` : null);
    if (!periodo) continue;
    const i = out.findIndex((v) => v.clienteId === clienteId && v.periodo === periodo);
    if (i > -1) { out[i] = { ...out[i], ...x.valori, fonte }; aggiornati++; }
    else {
      out.unshift({ id: uid(), clienteId, periodo, lead: 0, chiamate: 0, discoveryFissate: 0, discoverySvolte: 0, demoFissate: 0,
        demoSvolte: 0, offerte: 0, chiusure: 0, valoreContratto: 0, incassato: 0, spesa: 0, slot: 0, oreContatto: 0, convPiattaforma: 0,
        ...x.valori, fonte });
      nuovi++;
    }
  }
  return { vendite: out, nuovi, aggiornati };
}
