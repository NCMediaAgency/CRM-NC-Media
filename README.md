# CRM NC Media

CRM per agenzie di media buying. React + Vite, dati su Supabase, pubblicato su Cloudflare Pages.

## Avvio in locale

```bash
npm install
npm run dev               # con login e database Supabase
VITE_LOCALE=1 npm run dev # senza database: salta il login e salva nel browser
```

## Database

Le migrazioni sono in `supabase/migrations/`. Si eseguono in ordine dallo SQL Editor di Supabase.

## Pubblicazione

Cloudflare Pages con build `npm run build` e cartella di output `dist`. Il branch `main` è la versione definitiva; ogni altro branch ha il suo link di anteprima.

Dettagli sull'architettura e sulle regole di lavoro in `CLAUDE.md`.
