-- DamCerti Supabase schema
-- Run these in Supabase SQL Editor, in order, on a fresh project.
-- This file is a reference copy of what has already been run in your project.

-- 1. Profiles table (roles + subscription status)
create table profiles (
  id uuid references auth.users on delete cascade primary key,
  email text unique not null,
  role text default 'user' check (role in ('super_admin', 'sub_admin', 'user')),
  subscription_status text default 'inactive' check (subscription_status in ('active', 'inactive')),
  created_at timestamp with time zone default now(),
  last_sign_in timestamp with time zone
);

-- 2. Row Level Security
alter table profiles enable row level security;

create policy "Users can view own profile"
on profiles for select
using (auth.uid() = id);

create policy "Users can update own profile"
on profiles for update
using (auth.uid() = id);

create policy "Super admin can view all profiles"
on profiles for select
using (
  exists (
    select 1 from profiles
    where id = auth.uid() and role = 'super_admin'
  )
);

-- 3. Auto-create profile row on signup, auto-assign super_admin to your email
create function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.profiles (id, email, role, subscription_status)
  values (
    new.id,
    new.email,
    case
      when new.email = 'damilsajjad@gmail.com' then 'super_admin'
      else 'user'
    end,
    'inactive'
  );
  return new;
end;
$$;

create trigger on_auth_user_created
  after insert on auth.users
  for each row execute function public.handle_new_user();

-- 4. Free trial credit tracking (15 free certificates per non-subscribed user)
alter table profiles
add column free_certificates_used integer default 0 not null;
