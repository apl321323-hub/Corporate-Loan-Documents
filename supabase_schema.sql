create table if not exists public.app_data (
  key text primary key,
  value jsonb not null,
  updated_at timestamptz not null default now()
);

create or replace function public.set_app_data_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists trg_app_data_updated_at on public.app_data;

create trigger trg_app_data_updated_at
before update on public.app_data
for each row
execute function public.set_app_data_updated_at();
