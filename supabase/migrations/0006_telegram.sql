-- Bot Telegram del CRM: chat collegate, memoria breve della conversazione, promemoria inviati
-- e coda dei comandi che il CRM aperto applica (task, note nel diario, modifiche ai task).
-- Il browser resta l'unico a scrivere crm:data:v6: il bot non lo tocca, mette in coda.
-- Si applica da solo con l'automazione di GitHub; si può rieseguire.

-- Chat collegate a uno spazio di lavoro, con le preferenze dei promemoria
create table if not exists public.telegram_chat (
  workspace_id uuid not null references public.workspaces (id) on delete cascade,
  chat_id bigint not null,
  utente text not null default '',
  attivo boolean not null default true,
  ora_riepilogo text not null default '09:10',   -- HH:MM, ora di Roma
  weekend boolean not null default false,        -- riepilogo anche sabato e domenica
  preavviso integer not null default 30,         -- minuti prima dei task con orario
  collegato timestamptz not null default now(),
  primary key (workspace_id, chat_id)
);
create unique index if not exists telegram_chat_unica on public.telegram_chat (chat_id);

-- Codici usa e getta per collegare una chat (link t.me/…?start=CODICE), validi 30 minuti
create table if not exists public.telegram_codici (
  codice text primary key,
  workspace_id uuid not null references public.workspaces (id) on delete cascade,
  creato timestamptz not null default now()
);

-- Ultimi messaggi della conversazione: servono al bot per capire "no, spostalo al 24"
create table if not exists public.telegram_storico (
  id bigserial primary key,
  workspace_id uuid not null references public.workspaces (id) on delete cascade,
  chat_id bigint not null,
  ruolo text not null check (ruolo in ('user', 'assistant')),
  testo text not null,
  creato timestamptz not null default now()
);
create index if not exists telegram_storico_chat on public.telegram_storico (chat_id, id desc);

-- Promemoria già mandati (riepilogo del giorno, avviso di un task con orario): niente doppioni
create table if not exists public.telegram_inviati (
  workspace_id uuid not null references public.workspaces (id) on delete cascade,
  chiave text not null,
  inviato timestamptz not null default now(),
  primary key (workspace_id, chiave)
);

-- Coda dei comandi per il CRM: li applica il browser e li segna come applicati
create table if not exists public.crm_comandi (
  id text primary key,
  workspace_id uuid not null references public.workspaces (id) on delete cascade,
  tipo text not null,                 -- crea_task | crea_nota | modifica_task
  dati jsonb not null default '{}'::jsonb,
  origine text not null default 'telegram',
  creato timestamptz not null default now(),
  applicato timestamptz
);
create index if not exists crm_comandi_attesa on public.crm_comandi (workspace_id, creato) where applicato is null;

alter table public.telegram_chat enable row level security;
alter table public.telegram_codici enable row level security;
alter table public.telegram_storico enable row level security;
alter table public.telegram_inviati enable row level security;
alter table public.crm_comandi enable row level security;

-- Il CRM vede le chat collegate (per mostrarle) e la coda; preferenze e collegamento passano dalla funzione
drop policy if exists "membri vedono le chat" on public.telegram_chat;
create policy "membri vedono le chat" on public.telegram_chat
  for select to authenticated using (public.is_member(workspace_id));
drop policy if exists "membri leggono i comandi" on public.crm_comandi;
create policy "membri leggono i comandi" on public.crm_comandi
  for select to authenticated using (public.is_member(workspace_id));
drop policy if exists "membri segnano i comandi" on public.crm_comandi;
create policy "membri segnano i comandi" on public.crm_comandi
  for update to authenticated using (public.is_member(workspace_id)) with check (public.is_member(workspace_id));

revoke all on public.telegram_chat, public.telegram_codici, public.telegram_storico, public.telegram_inviati, public.crm_comandi from anon, authenticated;
grant select on public.telegram_chat to authenticated;
grant select, update (applicato) on public.crm_comandi to authenticated;
grant all on public.telegram_chat, public.telegram_codici, public.telegram_storico, public.telegram_inviati, public.crm_comandi to service_role;
grant usage, select on sequence public.telegram_storico_id_seq to service_role;

-- Promemoria: ogni 5 minuti la funzione guarda se c'è qualcosa da mandare (riepilogo, task con orario).
-- Stessa chiave del cron di meta-sync (migrazione 0002).
do $$
begin
  if exists (select 1 from cron.job where jobname = 'telegram-promemoria') then
    perform cron.unschedule('telegram-promemoria');
  end if;
end $$;
select cron.schedule(
  'telegram-promemoria',
  '*/5 * * * *',
  $cron$
  select net.http_post(
    url := 'https://cggxgzazjyeayfhynhhd.supabase.co/functions/v1/telegram',
    headers := jsonb_build_object('Content-Type', 'application/json', 'x-cron-key', public.meta_cron_chiave()),
    body := '{"azione":"promemoria"}'::jsonb,
    timeout_milliseconds := 10000
  );
  $cron$
);
