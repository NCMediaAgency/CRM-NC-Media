-- Winner ads: la stessa creatività vive in più copie (una per campagna o ad set, spesso in pausa).
-- "creativa" è il nome o il codice comune a tutte le copie: il CRM somma i loro numeri.
-- Vuoto = il CRM lo ricava dal nome (numero iniziale, oppure nome senza "copia").
alter table public.analisi_ads add column if not exists creativa text not null default '';
