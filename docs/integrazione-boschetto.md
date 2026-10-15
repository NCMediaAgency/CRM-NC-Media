# Collegamento Boschetto → CRM NC Media

Documento da dare a chi sviluppa Boschetto (anche al suo Claude). L'indirizzo e la chiave
si copiano dal CRM: Impostazioni → Software esterni → «Copia istruzioni con la chiave».
La chiave va tenuta sul server di Boschetto, mai nel codice che gira nel browser.

## Andata: le task da Boschetto al CRM

Ogni volta che una task assegnata a NC Media viene creata, modificata, completata o eliminata:

```
POST https://cggxgzazjyeayfhynhhd.supabase.co/functions/v1/integrazioni
Authorization: Bearer crmnc_…
Content-Type: application/json

{
  "evento": "task.creata",            // "task.aggiornata", "task.eliminata"
  "task": {
    "id": "id stabile della task in Boschetto",
    "titolo": "Testo della task",
    "descrizione": "facoltativa",
    "cliente": "Nome del cliente/progetto in Boschetto",
    "scadenza": "2026-10-23",         // AAAA-MM-GG o con orario 2026-10-23T15:30; anche 23/10/2026
    "ora": "15:30",                   // facoltativa
    "priorita": "alta",               // urgente | alta | media | bassa (anche urgent/high/medium/low)
    "stato": "da fare",               // da fare | in corso | in attesa | fatto (anche todo/in_progress/done)
    "url": "https://… link alla task" // facoltativo, diventa cliccabile nel CRM
  }
}
```

- Più task insieme: `{ "tasks": [ {…}, {…} ] }` (fino a 200 per richiesta). Al primo collegamento conviene mandare tutte le task aperte.
- Stesso `id` = stessa task: rimandarla la aggiorna. `task.eliminata` (o `"deleted": true` nella task) la chiude nel CRM.
- Sono accettati anche i nomi inglesi dei campi: `title`, `description`, `project`/`client`, `due`, `priority`, `status`, `link`.
- Risposte: `200 { "ok": true, "ricevute": N }`; `401` chiave errata o collegamento in pausa; `400` nessuna task valida (manca `id`).

## Ritorno: i cambi fatti nel CRM

Quando NC Media cambia una di queste task nel CRM (stato, scadenza, ora, priorità, titolo) il CRM registra il cambio. Due modi per riceverlo:

**1. Lettura periodica** (più semplice, per esempio ogni 5 minuti):

```
POST …/functions/v1/integrazioni        (stessa chiave)
{ "azione": "modifiche", "dal": 0 }
→ { "ok": true, "ultimo": 12, "modifiche": [
     { "id": 12, "esternoId": "id della task in Boschetto", "titolo": "…", "scadenza": "2026-10-30",
       "ora": "", "priorita": "alta", "stato": "fatto", "cambiati": ["stato"], "quando": "2026-10-15T10:00:00Z" } ] }
```

Salvare `ultimo` e usarlo come `dal` alla richiesta successiva. Gli stati arrivano come `dafare`, `incorso`, `attesa`, `fatto`.

**2. Avviso immediato**: Boschetto espone un suo indirizzo https (e, se vuole, una chiave). NC Media lo inserisce nel CRM
(Software esterni → Ritorno) e il CRM manda subito:

```
POST <indirizzo di Boschetto>
Authorization: Bearer <chiave di Boschetto>
{ "evento": "task.modificata_nel_crm", "modifiche": [ … come sopra … ] }
```

Rispondere 2xx. Nel CRM i cambi restano comunque leggibili con il modo 1.

Attenzione ai rimbalzi: quando Boschetto applica un cambio arrivato dal CRM e poi rimanda la task, il CRM la riconosce
uguale e non genera un nuovo cambio.
