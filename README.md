# CRM NC Media

CRM per agenzie di media buying. React + Vite, dati su Supabase, pubblicato su Cloudflare Workers.

## Avvio in locale

```bash
npm install
npm run dev               # con login e database Supabase
VITE_LOCALE=1 npm run dev # senza database: salta il login e salva nel browser
```

## Database

Le migrazioni sono in `supabase/migrations/`. Si eseguono in ordine dallo SQL Editor di Supabase.

## Pubblicazione

Cloudflare Workers, configurazione in `wrangler.jsonc`. Build `npm run build`, deploy `npx wrangler deploy`; per i branch diversi da `main`, `npx wrangler versions upload`, che crea un link di anteprima. Il branch `main` è la versione definitiva.

Dettagli sull'architettura e sulle regole di lavoro in `CLAUDE.md`.
