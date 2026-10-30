-- CAI Photo Upload Platform : database schema
-- Run this once in the Supabase SQL editor.

create extension if not exists "pgcrypto";

create table if not exists app_users (
  id uuid primary key default gen_random_uuid(),
  login_id text not null unique,
  display_name text not null,
  role text not null default 'admin',
  created_at bigint not null default (extract(epoch from now()) * 1000)::bigint,
  last_seen bigint not null default (extract(epoch from now()) * 1000)::bigint
);

create table if not exists sites (
  id uuid primary key default gen_random_uuid(),
  code text not null unique,
  name text not null default '',
  state text not null default '',
  region text not null default 'east',
  address text not null default '',
  latlong text not null default '',
  layout text not null default '',
  subfolders jsonb not null default '[]'::jsonb,
  created_at bigint not null default (extract(epoch from now()) * 1000)::bigint,
  updated_at bigint not null default (extract(epoch from now()) * 1000)::bigint
);

create index if not exists sites_region_idx on sites (region);

create table if not exists photos (
  id uuid primary key default gen_random_uuid(),
  site_id uuid not null references sites (id) on delete cascade,
  folder_id text not null,
  file_id text not null,
  file_name text not null default '',
  mime text not null default 'image/jpeg',
  size_bytes bigint not null default 0,
  uploaded_at bigint not null default (extract(epoch from now()) * 1000)::bigint,
  uploaded_by text not null default '',
  caption text not null default ''
);

create index if not exists photos_site_idx on photos (site_id);
create index if not exists photos_folder_idx on photos (site_id, folder_id);

-- Private bucket that holds every uploaded photo.
insert into storage.buckets (id, name, public)
values ('photos', 'photos', false)
on conflict (id) do nothing;

-- Row level security stays on. The server talks to Supabase with the
-- service role key, which bypasses RLS, so no public access is granted.
alter table app_users enable row level security;
alter table sites enable row level security;
alter table photos enable row level security;
