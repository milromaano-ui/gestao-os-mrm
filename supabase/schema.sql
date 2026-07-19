-- Rode este script no SQL Editor do seu projeto Supabase
-- (Supabase > seu projeto > SQL Editor > New query > colar e Run)

create table if not exists kv_store (
  key text primary key,
  value text not null,
  updated_at timestamptz not null default now()
);

-- Habilita Row Level Security
alter table kv_store enable row level security;

-- Como este é um app interno (só o Milton usa, com a chave anônima),
-- liberamos leitura e escrita para quem tiver a anon key do projeto.
-- A anon key não é secreta no sentido de "senha de admin", mas não
-- fica pública em lugar nenhum a não ser no seu próprio app.
create policy "Permitir leitura" on kv_store
  for select using (true);

create policy "Permitir escrita" on kv_store
  for insert with check (true);

create policy "Permitir atualização" on kv_store
  for update using (true);
