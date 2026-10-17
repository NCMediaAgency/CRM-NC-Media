-- Analisi delle inserzioni (il primo cliente: Boutique del Fitness).
-- Le scrive il connettore del CRM per Claude (funzione crm-mcp): il Claude del cliente legge Hyros
-- e salva qui i numeri di ogni annuncio e le sue analisi. Il CRM le legge e calcola la classifica.
-- Si applica da solo con l'automazione di GitHub; si può rieseguire.

-- Le chiavi delle integrazioni ora hanno un tipo: 'task' (software esterni) o 'mcp' (connettore per Claude)
alter table public.integrazioni add column if not exists tipo text not null default 'task';

-- Numeri di ogni annuncio in una data: una fotografia per cliente, fonte e giorno
create table if not exists public.analisi_ads (
  workspace_id uuid not null references public.workspaces (id) on delete cascade,
  cliente_id text not null,
  fonte text not null default 'Hyros',
  data date not null,
  ad_id text not null,
  nome text not null default '',
  campagna text not null default '',
  adset text not null default '',
  pubblico text not null default '',     -- freddo | retargeting | altro
  attiva boolean not null default true,
  spesa numeric not null default 0,
  lead numeric not null default 0,
  appuntamenti numeric not null default 0,
  clienti numeric not null default 0,
  incassato numeric not null default 0,
  periodo text not null default '',      -- a cosa si riferiscono i numeri, es. "ultimi 90 giorni"
  note text not null default '',
  creato timestamptz not null default now(),
  primary key (workspace_id, cliente_id, fonte, data, ad_id)
);

-- Analisi scritte (winner ads, creatività, andamento…)
create table if not exists public.analisi_testi (
  id bigserial primary key,
  workspace_id uuid not null references public.workspaces (id) on delete cascade,
  cliente_id text not null,
  tipo text not null default 'altro',
  titolo text not null default '',
  testo text not null default '',
  fonte text not null default 'Claude',
  creato timestamptz not null default now()
);
create index if not exists analisi_testi_cliente on public.analisi_testi (workspace_id, cliente_id, creato desc);

alter table public.analisi_ads enable row level security;
alter table public.analisi_testi enable row level security;
drop policy if exists "membri leggono" on public.analisi_ads;
create policy "membri leggono" on public.analisi_ads for select to authenticated using (public.is_member(workspace_id));
drop policy if exists "membri leggono" on public.analisi_testi;
create policy "membri leggono" on public.analisi_testi for select to authenticated using (public.is_member(workspace_id));
drop policy if exists "membri cancellano" on public.analisi_testi;
create policy "membri cancellano" on public.analisi_testi for delete to authenticated using (public.is_member(workspace_id));
grant select on public.analisi_ads to authenticated;
grant select, delete on public.analisi_testi to authenticated;
grant all on public.analisi_ads, public.analisi_testi to service_role;
grant usage, select on sequence public.analisi_testi_id_seq to service_role;
