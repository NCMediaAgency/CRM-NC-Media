-- Fase 1: spazi di lavoro e archivio del CRM.
-- Ogni dato appartiene a uno spazio di lavoro; un utente vede solo gli spazi di cui è membro.
-- Da eseguire una volta in Supabase: SQL Editor → New query → incolla → Run.

create table if not exists public.workspaces (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  created_at timestamptz not null default now()
);

create table if not exists public.workspace_members (
  workspace_id uuid not null references public.workspaces (id) on delete cascade,
  user_id uuid not null references auth.users (id) on delete cascade,
  role text not null default 'owner' check (role in ('owner', 'member')),
  created_at timestamptz not null default now(),
  primary key (workspace_id, user_id)
);
create index if not exists workspace_members_user_idx on public.workspace_members (user_id);

-- Archivio chiave/valore: è quello che il CRM chiamava window.storage.
-- crm:data:v6 contiene tutti i dati; crm:snap:AAAA-MM-GG le copie automatiche giornaliere.
create table if not exists public.kv_store (
  workspace_id uuid not null references public.workspaces (id) on delete cascade,
  key text not null,
  value text not null,
  updated_at timestamptz not null default now(),
  primary key (workspace_id, key)
);

alter table public.workspaces enable row level security;
alter table public.workspace_members enable row level security;
alter table public.kv_store enable row level security;

-- Vero se l'utente collegato è membro dello spazio. security definer: evita il giro
-- ricorsivo delle regole su workspace_members.
create or replace function public.is_member(ws uuid)
returns boolean
language sql
stable
security definer
set search_path = ''
as $$
  select exists (
    select 1 from public.workspace_members
    where workspace_id = ws and user_id = (select auth.uid())
  );
$$;

drop policy if exists "membri leggono lo spazio" on public.workspaces;
create policy "membri leggono lo spazio" on public.workspaces
  for select to authenticated using (public.is_member(id));

drop policy if exists "utente legge le sue appartenenze" on public.workspace_members;
create policy "utente legge le sue appartenenze" on public.workspace_members
  for select to authenticated using (user_id = (select auth.uid()));

drop policy if exists "membri gestiscono l'archivio" on public.kv_store;
create policy "membri gestiscono l'archivio" on public.kv_store
  for all to authenticated
  using (public.is_member(workspace_id))
  with check (public.is_member(workspace_id));

-- Restituisce lo spazio dell'utente; al primo accesso lo crea e lo nomina proprietario.
create or replace function public.my_workspace()
returns uuid
language plpgsql
security definer
set search_path = ''
as $$
declare
  uid uuid := auth.uid();
  ws uuid;
begin
  if uid is null then
    raise exception 'non autenticato';
  end if;
  select workspace_id into ws
    from public.workspace_members
    where user_id = uid
    order by created_at
    limit 1;
  if ws is null then
    insert into public.workspaces (name) values ('Il mio spazio') returning id into ws;
    insert into public.workspace_members (workspace_id, user_id, role) values (ws, uid, 'owner');
  end if;
  return ws;
end;
$$;

-- Permessi espliciti: i progetti nuovi possono non esporre le tabelle in automatico.
revoke all on function public.my_workspace() from public, anon;
grant execute on function public.my_workspace() to authenticated;
revoke all on function public.is_member(uuid) from public, anon;
grant execute on function public.is_member(uuid) to authenticated;

grant select on public.workspaces to authenticated;
grant select on public.workspace_members to authenticated;
grant select, insert, update, delete on public.kv_store to authenticated;
