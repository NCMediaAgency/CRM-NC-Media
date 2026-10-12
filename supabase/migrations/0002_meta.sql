-- Fase 2: dati di Meta Ads.
-- La funzione meta-sync (supabase/functions/meta-sync) scrive qui con la chiave di servizio;
-- il CRM legge queste tabelle e applica i numeri a campagne, ad set, annunci e finestre.
-- L'unica tabella che il CRM scrive è meta_collegamenti: quale account appartiene a quale cliente.
-- Da eseguire una volta in Supabase: SQL Editor → New query → incolla → Run. Si può rieseguire.

-- Account pubblicitari che il token vede, con l'ultimo aggiornamento riuscito
create table if not exists public.meta_account (
  workspace_id uuid not null references public.workspaces (id) on delete cascade,
  account_id text not null,               -- numerico, senza "act_"
  nome text not null default '',
  valuta text not null default '',
  stato int,                              -- account_status di Meta: 1 attivo, 2 disattivato, 3 pagamento in sospeso, 101 chiuso…
  fuso text not null default '',
  business text not null default '',
  ultimo_sync timestamptz,
  ultimo_errore text not null default '',
  aggiornato timestamptz not null default now(),
  primary key (workspace_id, account_id)
);

-- Quale account appartiene a quale cliente del CRM (un account, un cliente; un cliente, più account)
create table if not exists public.meta_collegamenti (
  workspace_id uuid not null references public.workspaces (id) on delete cascade,
  account_id text not null,
  cliente_id text not null,               -- id del cliente dentro crm:data:v6
  creato timestamptz not null default now(),
  primary key (workspace_id, account_id)
);

-- Campagne, ad set e annunci con i totali dall'inizio
create table if not exists public.meta_oggetti (
  workspace_id uuid not null references public.workspaces (id) on delete cascade,
  livello text not null check (livello in ('campaign', 'adset', 'ad')),
  id text not null,
  account_id text not null,
  campagna_id text not null default '',
  adset_id text not null default '',
  nome text not null default '',
  stato text not null default '',         -- effective_status
  obiettivo text not null default '',
  budget_giorno numeric not null default 0,
  budget_totale numeric not null default 0,
  inizio date,
  fine date,
  spesa numeric not null default 0,
  impression bigint not null default 0,
  click bigint not null default 0,
  click_link bigint not null default 0,
  lead numeric not null default 0,
  conversioni numeric not null default 0,
  ricavi numeric not null default 0,
  reach bigint not null default 0,
  frequenza numeric not null default 0,
  aggiornato timestamptz not null default now(),
  primary key (workspace_id, livello, id)
);
create index if not exists meta_oggetti_account_idx on public.meta_oggetti (workspace_id, account_id);

-- Finestre di valutazione di ogni campagna (Test 3 giorni, Matura 7, poi settimane), già sommate da Meta
create table if not exists public.meta_finestre (
  workspace_id uuid not null references public.workspaces (id) on delete cascade,
  campagna_id text not null,
  inizio date not null,
  fine date not null,
  tipo text not null check (tipo in ('test', 'matura', 'attiva')),
  account_id text not null,
  spesa numeric not null default 0,
  impression bigint not null default 0,
  click bigint not null default 0,
  click_link bigint not null default 0,
  lead numeric not null default 0,
  conversioni numeric not null default 0,
  ricavi numeric not null default 0,
  reach bigint not null default 0,
  frequenza numeric not null default 0,
  aggiornato timestamptz not null default now(),
  primary key (workspace_id, campagna_id, inizio)
);

-- Registro degli aggiornamenti
create table if not exists public.meta_sync (
  id bigint generated always as identity primary key,
  workspace_id uuid references public.workspaces (id) on delete cascade,
  origine text not null default 'programmato',
  avvio timestamptz not null default now(),
  fine timestamptz,
  esito text not null default 'in_corso' check (esito in ('in_corso', 'ok', 'parziale', 'errore')),
  messaggio text not null default '',
  dettagli jsonb not null default '{}'::jsonb
);
create index if not exists meta_sync_ws_idx on public.meta_sync (workspace_id, avvio desc);

alter table public.meta_account enable row level security;
alter table public.meta_collegamenti enable row level security;
alter table public.meta_oggetti enable row level security;
alter table public.meta_finestre enable row level security;
alter table public.meta_sync enable row level security;

-- I membri leggono tutto del proprio spazio; scrive solo la funzione (chiave di servizio, fuori dalla RLS).
drop policy if exists "membri leggono gli account" on public.meta_account;
create policy "membri leggono gli account" on public.meta_account
  for select to authenticated using (public.is_member(workspace_id));
drop policy if exists "membri leggono gli oggetti" on public.meta_oggetti;
create policy "membri leggono gli oggetti" on public.meta_oggetti
  for select to authenticated using (public.is_member(workspace_id));
drop policy if exists "membri leggono le finestre" on public.meta_finestre;
create policy "membri leggono le finestre" on public.meta_finestre
  for select to authenticated using (public.is_member(workspace_id));
drop policy if exists "membri leggono il registro" on public.meta_sync;
create policy "membri leggono il registro" on public.meta_sync
  for select to authenticated using (public.is_member(workspace_id));
-- I collegamenti li decide l'utente dal CRM
drop policy if exists "membri gestiscono i collegamenti" on public.meta_collegamenti;
create policy "membri gestiscono i collegamenti" on public.meta_collegamenti
  for all to authenticated
  using (public.is_member(workspace_id))
  with check (public.is_member(workspace_id));

grant select on public.meta_account, public.meta_oggetti, public.meta_finestre, public.meta_sync to authenticated;
grant select, insert, update, delete on public.meta_collegamenti to authenticated;
grant all on public.meta_account, public.meta_collegamenti, public.meta_oggetti, public.meta_finestre, public.meta_sync to service_role;

-- ---------- Avvio programmato: ogni mattina alle 9 ora italiana ----------
-- pg_cron ragiona in UTC: lo lancio alle 7 e alle 8 UTC e la funzione lavora solo
-- quando a Roma sono le 9 (7 UTC d'estate, 8 UTC d'inverno).
create extension if not exists pg_cron;
create extension if not exists pg_net;

-- Chiave condivisa fra il cron e la funzione, generata qui: nessuno deve copiarla a mano
do $$
begin
  if not exists (select 1 from vault.secrets where name = 'meta_sync_cron') then
    perform vault.create_secret(
      replace(gen_random_uuid()::text || gen_random_uuid()::text, '-', ''),
      'meta_sync_cron',
      'Chiave con cui il cron avvia la funzione meta-sync'
    );
  end if;
end $$;

-- La funzione la legge per riconoscere il cron e per richiamarsi quando gli account sono tanti
create or replace function public.meta_cron_chiave()
returns text
language sql
stable
security definer
set search_path = ''
as $$
  select decrypted_secret from vault.decrypted_secrets where name = 'meta_sync_cron' limit 1;
$$;
revoke all on function public.meta_cron_chiave() from public, anon, authenticated;
grant execute on function public.meta_cron_chiave() to service_role;

select cron.schedule(
  'meta-sync-mattina',
  '0 7,8 * * *',
  $cron$
  select net.http_post(
    url := 'https://cggxgzazjyeayfhynhhd.supabase.co/functions/v1/meta-sync',
    headers := jsonb_build_object('Content-Type', 'application/json', 'x-cron-key', public.meta_cron_chiave()),
    body := '{"origine":"programmato"}'::jsonb,
    timeout_milliseconds := 10000
  );
  $cron$
);
