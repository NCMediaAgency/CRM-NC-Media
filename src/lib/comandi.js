/* Comandi in arrivo dal bot Telegram (tabella crm_comandi, migrazione 0006).
   Il bot non scrive i dati del CRM: mette in coda, e il CRM aperto applica qui,
   con una funzione pura come applicaMeta. */

const PRIORITA = ["urgente", "alta", "media", "bassa"];
const STATI = ["dafare", "incorso", "attesa", "fatto"];

const somma = (d, n) => { const x = new Date(d + "T12:00:00Z"); x.setUTCDate(x.getUTCDate() + n); return x.toISOString().slice(0, 10); };

/* giorniRicorrenza(v): ogni quanti giorni si ripete un task ("no" = 0), come nel CRM */
export function applicaComandi(data, comandi, { uid, oggi, giorniRicorrenza = () => 0 }) {
  let task = [...(data.task || [])];
  let annotazioni = [...(data.annotazioni || [])];
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
        inizio: oggi, scadenza: d.scadenza || "", ora: d.ora || "", stima: 0, ricorrenza: "no", checklist: [], creato: oggi, origine: "telegram",
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
    } else esiti.push({ id: c.id, esito: "sconosciuto" });
  }
  return { data: { ...data, task, annotazioni }, esiti };
}
