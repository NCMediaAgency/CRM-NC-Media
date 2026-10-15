-- Integrazioni con software esterni (il primo: Boschetto, il gestionale di Kesia).
-- Andata: il software esterno manda le task alla funzione integrazioni con la sua chiave;
-- la funzione le mette in crm_comandi (migrazione 0006) e il CRM aperto le inserisce.
-- Ritorno: quando una task arrivata da fuori cambia nel CRM, il CRM lo comunica alla funzione,
-- che lo registra qui (il software esterno può leggerlo) e, se c'è, lo manda al suo indirizzo.
-- Si applica da solo con l'automazione di GitHub; si può rieseguire.

create table if not exists public.integrazioni (
  id uuid primary key default gen_random_uuid(),
  workspace_id uuid not null references public.workspaces (id) on delete cascade,
  nome text not null,
  chiave_hash text not null,               -- SHA-256 della chiave: la chiave in chiaro si vede una volta sola
  chiave_fine text not null default '',    -- ultimi caratteri, per riconoscerla
  url_ritorno text not null default '',    -- facoltativo: dove avvisare il software esterno dei cambi fatti nel CRM
  chiave_ritorno text not null default '', -- facoltativa: la chiave che il software esterno vuole ricevere
  attiva boolean not null default true,
  creato timestamptz not null default now(),
  ultimo_contatto timestamptz
);
create unique index if not exists integrazioni_chiave on public.integrazioni (chiave_hash);

-- Cambi fatti nel CRM sulle task arrivate da fuori: il canale di ritorno
create table if not exists public.integrazione_uscite (
  id bigserial primary key,
  workspace_id uuid not null references public.workspaces (id) on delete cascade,
  integrazione_id uuid not null references public.integrazioni (id) on delete cascade,
  esterno_id text not null,
  modifiche jsonb not null,
  creato timestamptz not null default now(),
  inviato timestamptz,                     -- consegnato all'indirizzo di ritorno
  esito text not null default ''
);
create index if not exists integrazione_uscite_lettura on public.integrazione_uscite (integrazione_id, id);

-- Invisibili dal browser: chiavi e indirizzi passano solo dalla funzione
alter table public.integrazioni enable row level security;
alter table public.integrazione_uscite enable row level security;
revoke all on public.integrazioni, public.integrazione_uscite from anon, authenticated;
grant all on public.integrazioni, public.integrazione_uscite to service_role;
grant usage, select on sequence public.integrazione_uscite_id_seq to service_role;
