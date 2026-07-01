-- Push notifications for RubyLife PWA

create table if not exists public.push_devices (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  endpoint text not null unique,
  auth_key text not null,
  p256dh_key text not null,
  platform text,
  browser text,
  active boolean not null default true,
  created_at timestamptz not null default now(),
  last_used_at timestamptz not null default now()
);

create index if not exists push_devices_user_active_idx on public.push_devices(user_id, active);

create table if not exists public.notification_preferences (
  user_id uuid primary key references auth.users(id) on delete cascade,
  preferences jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default now()
);

create table if not exists public.notification_history (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  title text not null,
  message text not null,
  type text not null,
  target_url text,
  read boolean not null default false,
  created_at timestamptz not null default now()
);

create index if not exists notification_history_user_created_idx on public.notification_history(user_id, created_at desc);

alter table public.push_devices enable row level security;
alter table public.notification_preferences enable row level security;
alter table public.notification_history enable row level security;

drop policy if exists "Users manage own push devices" on public.push_devices;
create policy "Users manage own push devices"
  on public.push_devices
  for all
  using (auth.uid() = user_id)
  with check (auth.uid() = user_id);

drop policy if exists "Users manage own notification preferences" on public.notification_preferences;
create policy "Users manage own notification preferences"
  on public.notification_preferences
  for all
  using (auth.uid() = user_id)
  with check (auth.uid() = user_id);

drop policy if exists "Users read own notification history" on public.notification_history;
create policy "Users read own notification history"
  on public.notification_history
  for select
  using (auth.uid() = user_id);

-- Inserts are performed by the service role inside /api/push when push delivery is attempted.
