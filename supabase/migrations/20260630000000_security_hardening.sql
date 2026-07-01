-- Security hardening: relational ownership checks and persistent server-side rate limits.

drop policy if exists "transactions_insert_own" on public.transactions;
drop policy if exists "transactions_update_own" on public.transactions;

create policy "transactions_insert_own" on public.transactions
  for insert with check (
    auth.uid() = user_id
    and (category_id is null or exists (
      select 1 from public.categories
      where categories.id = transactions.category_id
        and categories.user_id = auth.uid()
    ))
    and (account_id is null or exists (
      select 1 from public.accounts
      where accounts.id = transactions.account_id
        and accounts.user_id = auth.uid()
    ))
  );

create policy "transactions_update_own" on public.transactions
  for update using (auth.uid() = user_id)
  with check (
    auth.uid() = user_id
    and (category_id is null or exists (
      select 1 from public.categories
      where categories.id = transactions.category_id
        and categories.user_id = auth.uid()
    ))
    and (account_id is null or exists (
      select 1 from public.accounts
      where accounts.id = transactions.account_id
        and accounts.user_id = auth.uid()
    ))
  );

drop policy if exists "goal_movements_insert_own" on public.goal_movements;
drop policy if exists "goal_movements_update_own" on public.goal_movements;

create policy "goal_movements_insert_own" on public.goal_movements
  for insert with check (
    auth.uid() = user_id
    and exists (
      select 1 from public.goals
      where goals.id = goal_movements.goal_id
        and goals.user_id = auth.uid()
    )
  );

create policy "goal_movements_update_own" on public.goal_movements
  for update using (auth.uid() = user_id)
  with check (
    auth.uid() = user_id
    and exists (
      select 1 from public.goals
      where goals.id = goal_movements.goal_id
        and goals.user_id = auth.uid()
    )
  );

create table if not exists public.security_rate_limits (
  key text primary key,
  window_started_at timestamptz not null default now(),
  request_count integer not null default 0 check (request_count >= 0),
  updated_at timestamptz not null default now()
);

alter table public.security_rate_limits enable row level security;
revoke all on table public.security_rate_limits from public, anon, authenticated;
create index if not exists security_rate_limits_updated_at_idx
  on public.security_rate_limits (updated_at);

create or replace function public.consume_rate_limit(
  p_key text,
  p_limit integer,
  p_window_seconds integer
)
returns boolean
language plpgsql
security definer
set search_path = public
as $$
declare
  current_count integer;
  window_size interval;
begin
  if length(p_key) <> 64 or p_limit < 1 or p_limit > 10000
     or p_window_seconds < 1 or p_window_seconds > 86400 then
    raise exception 'invalid rate limit parameters';
  end if;

  window_size := make_interval(secs => p_window_seconds);

  insert into public.security_rate_limits as limits (key, window_started_at, request_count, updated_at)
  values (p_key, now(), 1, now())
  on conflict (key) do update set
    request_count = case
      when limits.window_started_at <= now() - window_size then 1
      else limits.request_count + 1
    end,
    window_started_at = case
      when limits.window_started_at <= now() - window_size then now()
      else limits.window_started_at
    end,
    updated_at = now()
  returning request_count into current_count;

  delete from public.security_rate_limits
  where updated_at < now() - interval '2 days';

  return current_count <= p_limit;
end;
$$;

revoke all on function public.consume_rate_limit(text, integer, integer) from public, anon, authenticated;
grant execute on function public.consume_rate_limit(text, integer, integer) to service_role;

do $$
begin
  if not exists (select 1 from pg_constraint where conname = 'accounts_name_length') then
    alter table public.accounts add constraint accounts_name_length check (char_length(name) between 1 and 160) not valid;
  end if;
  if not exists (select 1 from pg_constraint where conname = 'categories_name_length') then
    alter table public.categories add constraint categories_name_length check (char_length(name) between 1 and 160) not valid;
  end if;
  if not exists (select 1 from pg_constraint where conname = 'transactions_text_lengths') then
    alter table public.transactions add constraint transactions_text_lengths check (
      char_length(description) between 1 and 200
      and char_length(category) between 1 and 160
      and (notes is null or char_length(notes) <= 2000)
    ) not valid;
  end if;
  if not exists (select 1 from pg_constraint where conname = 'goals_name_length') then
    alter table public.goals add constraint goals_name_length check (char_length(name) between 1 and 160) not valid;
  end if;
  if not exists (select 1 from pg_constraint where conname = 'goal_movements_note_length') then
    alter table public.goal_movements add constraint goal_movements_note_length check (note is null or char_length(note) <= 2000) not valid;
  end if;
end $$;
