-- Regole per account: quali campagne scaricare e se l'account entra nell'aggiornamento di tutti.
-- Le decide l'utente dal CRM (Scheda cliente o Impostazioni), la funzione meta-sync le rispetta.
-- Da eseguire una volta in Supabase: SQL Editor → New query → incolla → Run. Si può rieseguire.

alter table public.meta_collegamenti
  -- 'recenti': attive o con spesa negli ultimi 90 giorni (come finora)
  -- 'attive':  solo quelle accese (più quelle che erano accese all'aggiornamento precedente, per registrarne lo stop)
  -- 'scelte':  solo le campagne elencate in campagne_scelte
  add column if not exists campagne_modo text not null default 'recenti',
  add column if not exists campagne_scelte text[] not null default '{}',
  -- false: l'account resta fuori dall'aggiornamento delle 9 e da "Aggiorna ora";
  -- si aggiorna solo con il pulsante del suo cliente
  add column if not exists automatico boolean not null default true;

alter table public.meta_collegamenti drop constraint if exists meta_collegamenti_campagne_modo_check;
alter table public.meta_collegamenti add constraint meta_collegamenti_campagne_modo_check
  check (campagne_modo in ('recenti', 'attive', 'scelte'));
