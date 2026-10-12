# CRM NC Media

CRM per agenzie di media buying (clienti, campagne, funnel, vendite, task). Oggi lo usa una sola agenzia; l'obiettivo è venderlo ad altri media buyer, quindi ogni scelta deve reggere più spazi di lavoro separati.

## Com'è fatto

- **App**: React 18 + Vite. Tutto il CRM è in `src/CRM.jsx` (un unico file, ~11.000 righe, CSS incluso nella costante `CSS`). È nato come artifact in una chat di Claude: modificarlo con interventi mirati, senza riscritture.
- **Login e avvio**: `src/App.jsx` (accesso email/password con Supabase, poi apre lo spazio di lavoro con la funzione `my_workspace()` e monta il CRM).
- **Archivio**: il CRM salva con `window.storage` (`get/set/list/delete`), l'interfaccia degli artifact di Claude. `src/lib/storage.js` la reimplementa sulla tabella `kv_store` di Supabase, filtrata per `workspace_id`. Chiavi:
  - `crm:data:v6`: tutti i dati del CRM in un unico JSON;
  - `crm:snap:AAAA-MM-GG`: copie automatiche giornaliere (ne tiene 6).
- **Database**: Supabase (progetto `cggxgzazjyeayfhynhhd`, regione UE). Lo schema è in `supabase/migrations/`, numerate in ordine; si applicano a mano dallo SQL Editor di Supabase. Ogni tabella con dati ha `workspace_id` e la RLS che passa da `public.is_member()`.
- **Pubblicazione**: Cloudflare Workers (progetto `crm-nc-media`, configurazione in `wrangler.jsonc`, sito statico servito da `dist`). Build `npm run build`; deploy `npx wrangler deploy` per `main` (versione definitiva), `npx wrangler versions upload` per gli altri branch (link di anteprima).

## Regole di lavoro

- Il proprietario verifica ogni modifica sull'anteprima prima di pubblicarla: si lavora su un branch e si porta su `main` solo quando lo chiede.
- Mai segreti nel repository: la chiave `service_role`/secret di Supabase, la password del database e i token di Meta vanno solo nelle impostazioni segrete di Supabase o di Cloudflare. URL e chiave `publishable` sono pubbliche per costruzione (in `src/lib/supabase.js`).
- Se cambia la struttura dei dati in `crm:data:v6`, aggiornare `normalizza()` in `CRM.jsx`, così gli archivi esistenti si adattano al caricamento.
- Prova locale senza database: `VITE_LOCALE=1 npm run dev` (salta il login e salva nel browser).
- Testi dell'interfaccia e commenti in italiano, come nel resto del codice.

## Prossimi passi previsti

- **Fase 2, sincronizzazione Meta**: una funzione di Supabase avviata ogni giorno alle 9:00 (Europe/Rome) scarica per ogni cliente con `accountAdv` (`act_…`) gli ultimi 7 giorni, giorno per giorno, a livello di campagna, ad set e annuncio. Va salvata in tabelle dedicate, non dentro `crm:data:v6`: il browser riscrive quel JSON intero a ogni modifica e cancellerebbe i dati arrivati dal server.
- **Fase 3, prodotto vendibile**: collegamento Facebook per ogni utente (OAuth e verifica dell'app da parte di Meta), abbonamenti con Stripe, registrazione pubblica, chiave dell'assistente AI spostata sul server (oggi `aiKey` sta nei dati e la chiamata ad Anthropic parte dal browser), dati normalizzati in tabelle al posto del JSON unico.
- `xlsx` 0.18.5 da npm ha vulnerabilità note sui file non fidati: passare alla versione distribuita da cdn.sheetjs.com.
