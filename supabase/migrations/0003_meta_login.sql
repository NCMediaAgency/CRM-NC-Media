-- Collegamento a Facebook con il profilo dell'utente (al posto dell'utente di sistema).
-- Il token resta solo qui, letto dalla funzione meta-sync con la chiave di servizio:
-- il CRM non può leggerlo, vede solo a nome di chi è il collegamento e quando scade.
-- Da eseguire una volta in Supabase: SQL Editor → New query → incolla → Run. Si può rieseguire.

create table if not exists public.meta_token (
  workspace_id uuid primary key references public.workspaces (id) on delete cascade,
  token text not null,
  scadenza timestamptz,
  utente_meta text not null default '',
  aggiornato timestamptz not null default now()
);

-- Tentativi di accesso in corso: il "state" della finestra di Facebook, valido pochi minuti
create table if not exists public.meta_login (
  stato text primary key,
  workspace_id uuid not null references public.workspaces (id) on delete cascade,
  ritorno text not null,
  creato timestamptz not null default now()
);

-- RLS attiva e nessuna regola per gli utenti: le tabelle sono invisibili dal browser
alter table public.meta_token enable row level security;
alter table public.meta_login enable row level security;
revoke all on public.meta_token, public.meta_login from anon, authenticated;
grant all on public.meta_token, public.meta_login to service_role;

-- Quello che il CRM può sapere del collegamento: nome e scadenza, mai il token
create or replace function public.meta_connessione()
returns table (collegato boolean, utente_meta text, scadenza timestamptz)
language sql
stable
security definer
set search_path = ''
as $$
  select true, t.utente_meta, t.scadenza
  from public.meta_token t
  where public.is_member(t.workspace_id)
  limit 1;
$$;
revoke all on function public.meta_connessione() from public, anon;
grant execute on function public.meta_connessione() to authenticated;
