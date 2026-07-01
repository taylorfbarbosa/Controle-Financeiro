alter table public.profiles
  add column if not exists phone text;

do $$
begin
  if not exists (select 1 from pg_constraint where conname = 'profiles_full_name_length') then
    alter table public.profiles
      add constraint profiles_full_name_length
      check (full_name is null or char_length(full_name) between 2 and 160) not valid;
  end if;

  if not exists (select 1 from pg_constraint where conname = 'profiles_phone_format') then
    alter table public.profiles
      add constraint profiles_phone_format
      check (phone is null or (char_length(phone) <= 30 and phone ~ '^[0-9()+ -]*$')) not valid;
  end if;
end $$;
