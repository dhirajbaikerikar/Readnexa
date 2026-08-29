-- ============================================================
-- READNEXA — Supabase setup script
-- Run this in Supabase: left sidebar > SQL Editor > New query > Run
-- ============================================================

-- If you already ran the earlier, simpler version of this script,
-- just run the two ALTER TABLE lines near the bottom instead.

create table if not exists folders (
  id text primary key,
  name text not null,
  created_at timestamp with time zone default now()
);

create table if not exists books (
  id text primary key,
  folder_id text references folders(id) on delete cascade,
  name text not null,
  file_name text,
  file_type text,
  size bigint,
  total_pages int not null default 1,
  last_read_page int default 0,
  uploaded_at timestamp with time zone default now(),
  last_read_at timestamp with time zone,
  file_url text,
  storage_path text,
  edited_html text,
  edited_at timestamp with time zone
);

-- Seed the two default shelves (safe to re-run)
insert into folders (id, name) values
  ('lean-six-sigma', 'Lean Six Sigma'),
  ('pmp', 'PMP')
on conflict (id) do nothing;

-- ------------------------------------------------------------
-- If your tables already existed from an earlier version,
-- just run these two lines to add the missing columns:
-- ------------------------------------------------------------
-- alter table books add column if not exists edited_html text;
-- alter table books add column if not exists edited_at timestamp with time zone;
-- alter table books add column if not exists storage_path text;

-- ------------------------------------------------------------
-- Row Level Security: keep this OFF for a simple solo project.
-- (Anon key is designed to be public; RLS is what protects data.
--  Since this is a personal/simple app with no login, we keep
--  RLS disabled so the anon key can read/write freely.)
-- ------------------------------------------------------------
alter table folders disable row level security;
alter table books disable row level security;
