-- Remove the discontinued push-notification module and its stored data.
drop table if exists public.notification_history cascade;
drop table if exists public.notification_preferences cascade;
drop table if exists public.push_devices cascade;
