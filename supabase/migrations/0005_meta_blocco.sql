-- Motivo per cui Meta ha disattivato un account pubblicitario (campo disable_reason dell'API).
-- Il CRM lo mostra accanto alle campagne dell'account, come l'errore di pagamento.
-- Si applica da solo con l'automazione di GitHub; si può rieseguire.

alter table public.meta_account
  add column if not exists motivo_blocco integer;
