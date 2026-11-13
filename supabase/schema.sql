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

-- ---------------------------------------------------------------------------
-- Individual logins, PPT document history and the access log
-- ---------------------------------------------------------------------------

-- PPT team members sign in with their own login ID and password (role 'ppt').
-- Everyone else uses the shared photo-upload access (role 'uploader').
alter table app_users add column if not exists pass_hash text;
alter table app_users add column if not exists aliases text[] not null default '{}';
alter table app_users add column if not exists is_owner boolean not null default false;
alter table app_users add column if not exists active boolean not null default true;
alter table app_users alter column role set default 'uploader';

-- Page 2 of the PPT (Document Owners / Change History) is filled from these.
alter table sites add column if not exists dpr_start text not null default '';
alter table sites add column if not exists ppt_created_by text not null default '';
alter table sites add column if not exists ppt_created_at bigint not null default 0;
alter table sites add column if not exists ppt_updated_by text not null default '';
alter table sites add column if not exists ppt_updated_at bigint not null default 0;

-- Who opened the website, signed in, generated a PPT, and when.
create table if not exists access_log (
  id bigserial primary key,
  login_id text not null default '',
  login_key text not null default '',
  role text not null default '',
  action text not null default '',
  detail text not null default '',
  ip text not null default '',
  device text not null default '',
  at bigint not null default (extract(epoch from now()) * 1000)::bigint
);
create index if not exists access_log_at_idx on access_log (at desc);
create index if not exists access_log_key_idx on access_log (login_key, at desc);
alter table access_log enable row level security;
