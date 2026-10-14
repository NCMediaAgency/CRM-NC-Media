# CRM NC Media

CRM per agenzie di media buying (clienti, campagne, funnel, vendite, task). Oggi lo usa una sola agenzia; l'obiettivo è venderlo ad altri media buyer, quindi ogni scelta deve reggere più spazi di lavoro separati.

## Com'è fatto

- **App**: React 18 + Vite. Tutto il CRM è in `src/CRM.jsx` (un unico file, ~11.000 righe, CSS incluso nella costante `CSS`). È nato come artifact in una chat di Claude: modificarlo con interventi mirati, senza riscritture.
- **Login e avvio**: `src/App.jsx` (accesso email/password con Supabase, poi apre lo spazio di lavoro con la funzione `my_workspace()` e monta il CRM).
- **Archivio**: il CRM salva con `window.storage` (`get/set/list/delete`), l'interfaccia degli artifact di Claude. `src/lib/storage.js` la reimplementa sulla tabella `kv_store` di Supabase, filtrata per `workspace_id`. Chiavi:
  - `crm:data:v6`: tutti i dati del CRM in un unico JSON;
  - `crm:snap:AAAA-MM-GG`: copie automatiche giornaliere (ne tiene 6).
- **Database**: Supabase (progetto `cggxgzazjyeayfhynhhd`, regione UE). Lo schema è in `supabase/migrations/`, numerate in ordine e scritte per poter essere rieseguite. Ogni tabella con dati ha `workspace_id` e la RLS che passa da `public.is_member()`.
- **Meta Ads**: la Edge Function `supabase/functions/meta-sync/index.ts` (file unico senza dipendenze, si incolla nell'editor di Supabase; "Verify JWT" spento perché la chiama anche il cron) scarica dagli account collegati campagne, ad set e annunci attivi o con spesa negli ultimi 90 giorni (totali da sempre) e le finestre Test/Matura/settimane via `time_ranges`. Scrive in `meta_account`, `meta_oggetti`, `meta_finestre`, `meta_sync` (migrazione 0002). Il cron delle 7 e 8 UTC la avvia, lavora solo alle 9 di Roma; risponde subito e lavora in sottofondo, a fine tempo si richiama. Accesso a Meta: "Collega Facebook" (OAuth con l'app del proprietario; segreti `META_APP_ID`, `META_APP_SECRET`, facoltativo `META_CONFIG_ID` per Facebook Login for Business) salva un token utente da 60 giorni per spazio di lavoro in `meta_token` (migrazione 0003, invisibile dal browser; il CRM vede nome e scadenza con `meta_connessione()`). Il ritorno da Facebook arriva alla stessa funzione (GET con `state`) e rimanda al CRM con `?meta=collegato`. Indirizzi di ritorno ammessi: `*crm-nc-media.nciappa9.workers.dev` e localhost, altri con il segreto `CRM_ORIGINI`. `META_TOKEN` resta come riserva.
- **Dati Meta nel CRM**: il CRM legge le tabelle all'apertura e dopo "Aggiorna ora" (`src/lib/metaApi.js`) e le applica con `applicaMeta` (`src/lib/meta.js`, funzione pura): campagne per `metaId` o per nome alla prima volta, creativi, finestre (`origine: "meta"`), storico del budget, stato. Il browser resta l'unico a scrivere `crm:data:v6`. I collegamenti account → cliente stanno in `meta_collegamenti` e si gestiscono da Impostazioni e Scheda cliente. Ogni collegamento ha anche le sue regole (migrazione 0004, si cambiano dalla Scheda cliente): `campagne_modo` (`recenti` = attive o con spesa in 90 giorni, `attive`, `scelte` con gli id in `campagne_scelte`; con le ultime due la funzione toglie da `meta_oggetti`/`meta_finestre` le campagne rimaste fuori) e `automatico` (false = fuori dal cron e da "Aggiorna ora", si aggiorna solo con il pulsante del cliente). L'azione `campagne` della funzione elenca le campagne di un account per la scelta. In prova locale `src/lib/metaFinta.js` simula Meta.
- **Pubblicazione su Supabase**: automatica. `.github/workflows/supabase.yml` parte a ogni invio che tocca `supabase/` (branch di prova e `main`, o a mano da Actions): esegue le migrazioni non ancora fatte (`supabase db push`, indirizzo dal pooler) e pubblica `meta-sync` con Verify JWT spento. Segreti del repository: `SUPABASE_ACCESS_TOKEN` (token con Project Settings, Database, Migrations, Connection Pooling, Edge Functions; niente chiavi API, quindi niente `supabase link`) e `SUPABASE_DB_PASSWORD`. C'è un solo progetto Supabase per anteprima e definitiva: le modifiche a database e funzione valgono subito per entrambe, quindi vanno sempre compatibili con la versione pubblicata.
- **File KPI → Ciclo di vendita**: `src/lib/kpi.js` (funzioni pure) trova nei fogli Excel le tabelle con le voci in colonna e i mesi in riga (anche con sottocolonne per funnel) e le scrive nei periodi `vendite` del cliente; il riquadro è nella scheda File e nel Ciclo di vendita del cliente.
- **Prove**: `npm test` (applicazione dei dati Meta e funzione meta-sync con Meta e database simulati).
- **Pubblicazione**: Cloudflare Workers (progetto `crm-nc-media`, configurazione in `wrangler.jsonc`, sito statico servito da `dist`). Build `npm run build`; deploy `npx wrangler deploy` per `main` (versione definitiva), `npx wrangler preview` per gli altri branch (link di anteprima).

## Regole di lavoro

- Il proprietario verifica ogni modifica sull'anteprima prima di pubblicarla: si lavora su un branch e si porta su `main` solo quando lo chiede.
- Mai segreti nel repository: la chiave `service_role`/secret di Supabase, la password del database e i token di Meta vanno solo nelle impostazioni segrete di Supabase o di Cloudflare. URL e chiave `publishable` sono pubbliche per costruzione (in `src/lib/supabase.js`).
- Se cambia la struttura dei dati in `crm:data:v6`, aggiornare `normalizza()` in `CRM.jsx`, così gli archivi esistenti si adattano al caricamento.
- Prova locale senza database: `VITE_LOCALE=1 npm run dev` (salta il login e salva nel browser).
- Testi dell'interfaccia e commenti in italiano, come nel resto del codice.

## Prossimi passi previsti

- **Fase 2, Meta**: fatta (vedi sopra). Da rifinire con l'uso: lead dai moduli (serve `leads_retrieval`), ritmo di spesa giornaliero dai dati per giorno.
- **Fase 3, prodotto vendibile**: collegamento Facebook per ogni utente (OAuth e verifica dell'app da parte di Meta), abbonamenti con Stripe, registrazione pubblica, chiave dell'assistente AI spostata sul server (oggi `aiKey` sta nei dati e la chiamata ad Anthropic parte dal browser), dati normalizzati in tabelle al posto del JSON unico.
- `xlsx` 0.18.5 da npm ha vulnerabilità note sui file non fidati: passare alla versione distribuita da cdn.sheetjs.com.
