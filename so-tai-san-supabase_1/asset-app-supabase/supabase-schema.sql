-- ============================================================================
-- SỔ TÀI SẢN — Supabase schema
-- Dán toàn bộ file này vào Supabase Dashboard → SQL Editor → New query → Run
-- ============================================================================

-- 1) PROFILES — vai trò của từng người dùng, gắn với tài khoản trong Supabase Auth
create table if not exists public.profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  email text not null default '',
  name text not null default 'Người dùng mới',
  role text not null default 'user' check (role in ('admin', 'user')),
  created_at timestamptz not null default now()
);

alter table public.profiles enable row level security;

-- Ai cũng xem được danh sách người dùng (để hiển thị tên trong "người dùng" của tài sản)
create policy "profiles_select_all" on public.profiles
  for select using (auth.role() = 'authenticated');

-- Mỗi người chỉ tự sửa tên của chính mình
create policy "profiles_update_self" on public.profiles
  for update using (auth.uid() = id);

-- Chỉ admin được đổi role của người khác
create policy "profiles_update_role_by_admin" on public.profiles
  for update using (
    exists (select 1 from public.profiles p where p.id = auth.uid() and p.role = 'admin')
  );

-- Tự động tạo hồ sơ (role = user) mỗi khi có tài khoản đăng ký mới
create or replace function public.handle_new_user()
returns trigger as $$
begin
  insert into public.profiles (id, email, name, role)
  values (new.id, new.email, coalesce(new.raw_user_meta_data->>'name', split_part(new.email, '@', 1)), 'user');
  return new;
end;
$$ language plpgsql security definer;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
  after insert on auth.users
  for each row execute procedure public.handle_new_user();


-- 2) APP_DATA — toàn bộ dữ liệu nghiệp vụ (tài sản, nhân sự, giao dịch, cài đặt...)
--    lưu dạng 1 dòng JSONB duy nhất, giữ cấu trúc tương thích với app hiện tại.
create table if not exists public.app_data (
  id int primary key default 1,
  data jsonb not null,
  updated_at timestamptz not null default now(),
  constraint singleton check (id = 1)
);

alter table public.app_data enable row level security;

create policy "app_data_select" on public.app_data
  for select using (auth.role() = 'authenticated');

create policy "app_data_upsert" on public.app_data
  for all using (auth.role() = 'authenticated')
  with check (auth.role() = 'authenticated');


-- 3) BACKUPS — các bản chụp dữ liệu theo mốc thời gian
create table if not exists public.backups (
  id uuid primary key default gen_random_uuid(),
  label text not null default 'Thủ công',
  created_by text not null default '',
  data jsonb not null,
  created_at timestamptz not null default now()
);

alter table public.backups enable row level security;

create policy "backups_select" on public.backups
  for select using (auth.role() = 'authenticated');

create policy "backups_insert_admin" on public.backups
  for insert with check (
    exists (select 1 from public.profiles p where p.id = auth.uid() and p.role = 'admin')
  );

create policy "backups_delete_admin" on public.backups
  for delete using (
    exists (select 1 from public.profiles p where p.id = auth.uid() and p.role = 'admin')
  );


-- 4) Dữ liệu khởi tạo — dòng app_data rỗng ban đầu (app sẽ tự điền dữ liệu mẫu lần đầu chạy)
insert into public.app_data (id, data)
values (1, '{}'::jsonb)
on conflict (id) do nothing;


-- ============================================================================
-- SAU KHI CHẠY XONG FILE NÀY:
-- 1. Vào Authentication → Users → Add user, tạo tài khoản đầu tiên (email + mật khẩu)
--    Đây sẽ là quản trị viên (Toàn quyền) đầu tiên của bạn.
-- 2. Vào Table Editor → profiles, tìm dòng vừa được tự tạo (trigger ở trên tự thêm),
--    sửa cột "role" từ "user" thành "admin", và đặt "name" là tên hiển thị của bạn.
-- 3. Lấy Project URL và anon public key ở Settings → API, dán vào file .env (xem HUONG_DAN.md)
-- ============================================================================
