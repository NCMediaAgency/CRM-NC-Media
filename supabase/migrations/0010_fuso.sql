-- Fuso orario e ora dell'aggiornamento di Meta per ogni spazio di lavoro.
-- Prima tutto era sull'ora di Roma (aggiornamento alle 9 italiane): un'agenzia americana
-- riceveva i dati nel cuore della notte. Ora ogni spazio ha il suo fuso; il CRM lo propone
-- dal computer di chi lo crea e si cambia dalle Impostazioni.
-- Si applica da solo con l'automazione di GitHub; si può rieseguire.

create table if not exists public.workspace_impostazioni (
  workspace_id uuid primary key references public.workspaces (id) on delete cascade,
  fuso text not null default 'Europe/Rome',          -- nome IANA, es. America/New_York
  ora_meta integer not null default 9 check (ora_meta between 0 and 23),   -- ora locale dell'aggiornamento di Meta
  ultimo_meta date,                                  -- ultimo giorno (locale) in cui è partito l'aggiornamento programmato
  aggiornato timestamptz not null default now()
);

alter table public.workspace_impostazioni enable row level security;
drop policy if exists "membri leggono" on public.workspace_impostazioni;
create policy "membri leggono" on public.workspace_impostazioni for select to authenticated using (public.is_member(workspace_id));
drop policy if exists "membri creano" on public.workspace_impostazioni;
create policy "membri creano" on public.workspace_impostazioni for insert to authenticated with check (public.is_member(workspace_id));
drop policy if exists "membri cambiano" on public.workspace_impostazioni;
create policy "membri cambiano" on public.workspace_impostazioni for update to authenticated using (public.is_member(workspace_id)) with check (public.is_member(workspace_id));
grant select, insert, update (fuso, ora_meta, aggiornato) on public.workspace_impostazioni to authenticated;
grant all on public.workspace_impostazioni to service_role;

-- Il cron di Meta ora parte ogni ora: la funzione aggiorna gli spazi per cui è l'ora scelta nel loro fuso
select cron.schedule(
  'meta-sync-mattina',
  '0 * * * *',
  $cron$
  select net.http_post(
    url := 'https://cggxgzazjyeayfhynhhd.supabase.co/functions/v1/meta-sync',
    headers := jsonb_build_object('Content-Type', 'application/json', 'x-cron-key', public.meta_cron_chiave()),
    body := '{"origine":"programmato"}'::jsonb,
    timeout_milliseconds := 10000
  );
  $cron$
);
