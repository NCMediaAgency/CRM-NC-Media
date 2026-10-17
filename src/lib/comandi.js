/* Comandi in arrivo dal bot Telegram (tabella crm_comandi, migrazione 0006).
   Il bot non scrive i dati del CRM: mette in coda, e il CRM aperto applica qui,
   con una funzione pura come applicaMeta. */

const PRIORITA = ["urgente", "alta", "media", "bassa"];
/* Nomi dei clienti di un software esterno: confronto senza maiuscole, accenti e spazi doppi */
export const nomeChiave = (s) => String(s || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, " ").trim();
/* Campi di una task che viaggiano verso il software esterno (canale di ritorno) */
export const CAMPI_RITORNO = ["titolo", "scadenza", "ora", "priorita", "stato"];
export const statoEsterno = (t) => Object.fromEntries(CAMPI_RITORNO.map((k) => [k, t[k] || ""]));
const STATI = ["dafare", "incorso", "attesa", "fatto"];

const somma = (d, n) => { const x = new Date(d + "T12:00:00Z"); x.setUTCDate(x.getUTCDate() + n); return x.toISOString().slice(0, 10); };

/* giorniRicorrenza(v): ogni quanti giorni si ripete un task ("no" = 0), come nel CRM */
export function applicaComandi(data, comandi, { uid, oggi, giorniRicorrenza = () => 0 }) {
  let task = [...(data.task || [])];
  let annotazioni = [...(data.annotazioni || [])];
  let integrazioni = [...(data.integrazioni || [])];
  const esiti = [];
  const cliOk = (id) => (id && (data.clienti || []).some((z) => z.id === id) ? id : "");
  for (const c of comandi) {
    const d = c.dati || {};
    if (c.tipo === "crea_task") {
      if (task.some((t) => t.id === d.id)) { esiti.push({ id: c.id, esito: "già presente" }); continue; }
      task.push({
        id: d.id || uid(), clienteId: cliOk(d.clienteId), areaId: (data.aree || []).some((a) => a.id === d.areaId) ? d.areaId : "",
        titolo: String(d.titolo || "Task da Telegram"), note: d.note || "", stato: "dafare",
        priorita: PRIORITA.includes(d.priorita) ? d.priorita : "media",
        inizio: oggi, scadenza: d.scadenza || "", ora: d.ora || "", stima: 0, ricorrenza: "no", checklist: [], creato: oggi, origine: d.origine || "telegram",
      });
      esiti.push({ id: c.id, esito: "task creato" });
    } else if (c.tipo === "crea_nota") {
      if (annotazioni.some((n) => n.id === d.id)) { esiti.push({ id: c.id, esito: "già presente" }); continue; }
      annotazioni = [{ id: d.id || uid(), clienteId: cliOk(d.clienteId), campagnaId: "", data: d.data || oggi, tipo: d.tipo || "altro", testo: String(d.testo || ""), origine: "telegram" }, ...annotazioni];
      esiti.push({ id: c.id, esito: "nota creata" });
    } else if (c.tipo === "modifica_task") {
      const i = task.findIndex((t) => t.id === d.taskId);
      if (i < 0) { esiti.push({ id: c.id, esito: "task non trovato" }); continue; }
      const campi = {};
      if (d.titolo) campi.titolo = String(d.titolo);
      if (d.scadenza) campi.scadenza = d.scadenza;
      if (d.ora) campi.ora = d.ora;
      if (PRIORITA.includes(d.priorita)) campi.priorita = d.priorita;
      if (STATI.includes(d.stato)) campi.stato = d.stato;
      const t = task[i];
      task[i] = { ...t, ...campi };
      // Chiudere un task che si ripete fa nascere il successivo, come la spunta nel CRM
      const g = giorniRicorrenza(t.ricorrenza);
      if (campi.stato === "fatto" && t.stato !== "fatto" && g) {
        task.unshift({ ...t, id: uid(), stato: "dafare", inizio: somma(t.inizio || oggi, g), scadenza: somma(t.scadenza || oggi, g),
          checklist: (t.checklist || []).map((r) => ({ ...r, fatto: false })), creato: oggi });
      }
      esiti.push({ id: c.id, esito: "task aggiornato" });
    } else if (c.tipo === "importa_task") {
      /* Task da un software esterno (Boschetto): una sola task per id esterno, aggiornata a ogni invio */
      let ii = integrazioni.findIndex((x) => x.id === d.integrazioneId);
      if (ii < 0) { integrazioni.push({ id: d.integrazioneId, nome: d.fonte || "Integrazione", clienteBase: "", abbinamenti: {}, nomiVisti: {} }); ii = integrazioni.length - 1; }
      const integ = integrazioni[ii];
      const chiave = nomeChiave(d.cliente);
      if (chiave && !(integ.nomiVisti || {})[chiave]) integrazioni[ii] = { ...integ, nomiVisti: { ...(integ.nomiVisti || {}), [chiave]: d.cliente } };
      const abbinato = chiave && (integ.abbinamenti || {})[chiave];
      const cliente = cliOk(abbinato) || cliOk(integ.clienteBase) || "";
      const i = task.findIndex((t) => t.esterno && t.esterno.integrazioneId === d.integrazioneId && t.esterno.id === d.esternoId);
      const campi = {
        titolo: String(d.titolo || "Task"), note: d.descrizione || "", scadenza: d.scadenza || "", ora: d.ora || "",
        priorita: PRIORITA.includes(d.priorita) ? d.priorita : "media", stato: STATI.includes(d.stato) ? d.stato : "dafare",
      };
      const esterno = { integrazioneId: d.integrazioneId, fonte: d.fonte || integ.nome, id: d.esternoId, cliente: d.cliente || "", url: d.url || "",
        ultimo: statoEsterno(campi), eliminata: !!d.eliminata };
      if (d.eliminata) {
        if (i > -1) { task[i] = { ...task[i], stato: "fatto", esterno: { ...task[i].esterno, eliminata: true } }; esiti.push({ id: c.id, esito: "task eliminata alla fonte" }); }
        else esiti.push({ id: c.id, esito: "eliminata, mai arrivata" });
        continue;
      }
      if (i > -1) {
        // Il cliente si aggiorna solo se non l'hai spostata tu su un altro sottocliente
        const vecchio = task[i].clienteId;
        const tuo = vecchio && vecchio !== (integ.clienteBase || "") && vecchio !== cliOk((integ.abbinamenti || {})[nomeChiave(task[i].esterno.cliente)]);
        task[i] = { ...task[i], ...campi, clienteId: tuo ? vecchio : cliente, esterno };
        esiti.push({ id: c.id, esito: "task aggiornata" });
      } else {
        task.push({ id: uid(), clienteId: cliente, areaId: "", ...campi, inizio: oggi, stima: 0, ricorrenza: "no", checklist: [], creato: oggi, origine: esterno.fonte, esterno });
        esiti.push({ id: c.id, esito: "task importata" });
      }
    } else esiti.push({ id: c.id, esito: "sconosciuto" });
  }
  return { data: { ...data, task, annotazioni, integrazioni }, esiti };
}
