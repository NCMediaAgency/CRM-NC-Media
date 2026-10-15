/* Integrazioni: lettura delle task che arrivano da un software esterno (Boschetto)
   Funzioni pure, provate con node (tests/integrazioni.test.mjs). */

const PRIORITA = { urgente: "urgente", urgent: "urgente", critica: "urgente", alta: "alta", high: "alta", media: "media", medium: "media", normale: "media", normal: "media", bassa: "bassa", low: "bassa" };
const STATI = {
  dafare: "dafare", "da fare": "dafare", todo: "dafare", "to do": "dafare", aperta: "dafare", open: "dafare", nuova: "dafare", new: "dafare",
  incorso: "incorso", "in corso": "incorso", in_progress: "incorso", "in progress": "incorso", doing: "incorso",
  attesa: "attesa", "in attesa": "attesa", waiting: "attesa", blocked: "attesa", bloccata: "attesa",
  fatto: "fatto", fatta: "fatto", completata: "fatto", completato: "fatto", done: "fatto", completed: "fatto", chiusa: "fatto", closed: "fatto",
};
const giorno = (s) => {
  if (!s) return "";
  const t = String(s).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(t)) return t.slice(0, 10);
  const m = /^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$/.exec(t);
  return m ? `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}` : "";
};
const ora = (s, data) => {
  const t = String(s || "").trim();
  if (/^\d{1,2}:\d{2}/.test(t)) return t.slice(0, 5).padStart(5, "0");
  // Data con orario: 2026-10-23T15:30
  const m = /T(\d{2}:\d{2})/.exec(String(data || ""));
  return m && m[1] !== "00:00" ? m[1] : "";
};
const testo = (x, max) => String(x ?? "").trim().slice(0, max);

/* Una task come la manda Boschetto → i campi del CRM. Accetta nomi italiani o inglesi. */
export function leggiTaskEsterna(t) {
  if (!t || typeof t !== "object") return null;
  const id = testo(t.id ?? t.esternoId ?? t.uuid, 120);
  if (!id) return null;
  const scad = t.scadenza ?? t.due ?? t.due_date ?? t.dueDate ?? "";
  return {
    esternoId: id,
    titolo: testo(t.titolo ?? t.title ?? t.nome ?? t.name, 200) || "Task senza titolo",
    descrizione: testo(t.descrizione ?? t.description ?? t.note ?? t.notes, 4000),
    scadenza: giorno(scad),
    ora: ora(t.ora ?? t.time, scad),
    priorita: PRIORITA[String(t.priorita ?? t.priority ?? "").toLowerCase().trim()] || "media",
    stato: STATI[String(t.stato ?? t.status ?? "").toLowerCase().trim()] || "dafare",
    cliente: testo(t.cliente ?? t.client ?? t.progetto ?? t.project, 200),
    url: /^https?:\/\//.test(String(t.url ?? t.link ?? "")) ? testo(t.url ?? t.link, 500) : "",
    assegnatoA: testo(t.assegnatoA ?? t.assignee ?? "", 200),
  };
}

/* Il corpo che arriva: { evento, task } oppure { task: [...] } oppure { tasks: [...] } */
export function eventiDaCorpo(corpo) {
  if (!corpo || typeof corpo !== "object") return [];
  const evento = String(corpo.evento ?? corpo.event ?? "task.aggiornata");
  const elimina = /elimin|delet|remov/i.test(evento);
  const lista = Array.isArray(corpo.tasks) ? corpo.tasks : Array.isArray(corpo.task) ? corpo.task : corpo.task ? [corpo.task] : [];
  return lista.slice(0, 200).map((t) => {
    const x = leggiTaskEsterna(t);
    if (!x) return null;
    return { ...x, eliminata: elimina || t.eliminata === true || t.deleted === true };
  }).filter(Boolean);
}

