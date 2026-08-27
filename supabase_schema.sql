-- ============================================================
-- SKEMA DATABASE SUPABASE UNTUK LOGIN MULTI-GURU
-- Jalankan seluruh script ini di: Supabase Dashboard > SQL Editor > New Query > Run
-- ============================================================

-- Tabel Jurnal Mengajar
create table if not exists jurnal_mengajar (
    id bigint generated always as identity primary key,
    user_id uuid references auth.users(id) on delete cascade not null,
    tanggal text,
    kelas text,
    mapel text,
    jp text,
    materi text,
    kegiatan text,
    hadir text,
    catatan text,
    created_at timestamptz default now()
);

-- Tabel Data Siswa
create table if not exists data_siswa (
    id bigint generated always as identity primary key,
    user_id uuid references auth.users(id) on delete cascade not null,
    nama text,
    jk text,
    nisn text,
    created_at timestamptz default now()
);

-- Aktifkan Row Level Security (RLS) - WAJIB agar data tiap guru terpisah/privat
alter table jurnal_mengajar enable row level security;
alter table data_siswa enable row level security;

-- Kebijakan: setiap guru hanya boleh melihat & mengubah datanya sendiri
create policy "guru hanya akses jurnal miliknya sendiri"
    on jurnal_mengajar for all
    using (auth.uid() = user_id)
    with check (auth.uid() = user_id);

create policy "guru hanya akses data siswa miliknya sendiri"
    on data_siswa for all
    using (auth.uid() = user_id)
    with check (auth.uid() = user_id);

-- ============================================================
-- TAMBAHAN: SISTEM KODE LISENSI (AKSES BERBAYAR PERMANEN PER GURU)
-- Jalankan bagian ini di SQL Editor SETELAH skema di atas (boleh sekali jalan bersamaan).
-- ============================================================

create table if not exists kode_lisensi (
    id bigint generated always as identity primary key,
    kode text unique not null,
    status text not null default 'belum_terpakai',  -- belum_terpakai | terpakai
    dipakai_oleh_user_id uuid references auth.users(id) on delete set null,
    dipakai_oleh_email text,
    catatan text,
    dibuat_tanggal timestamptz default now(),
    digunakan_tanggal timestamptz
);

alter table kode_lisensi enable row level security;
-- Sengaja TIDAK ada policy untuk anon/authenticated di sini -> tabel ini TERKUNCI TOTAL
-- dari akses langsung lewat anon key. Akses hanya lewat:
--   1) tiga fungsi RPC di bawah (dipakai alur pendaftaran guru dengan anon key), atau
--   2) SUPABASE_SERVICE_KEY (dipakai KHUSUS di panel admin Anda, tidak pernah dibagikan).

-- 1) Kunci ("reserve") kode secara atomik SEBELUM akun guru dibuat, mencegah 1 kode
--    terpakai dua kali walau ada percobaan bersamaan (race condition).
create or replace function reserve_kode_lisensi(p_kode text)
returns boolean
language plpgsql
security definer
set search_path = public
as $$
declare
    v_updated int;
begin
    update kode_lisensi
    set status = 'terpakai'
    where kode = p_kode and status = 'belum_terpakai';
    get diagnostics v_updated = row_count;
    return v_updated > 0;
end;
$$;

-- 2) Batalkan reservasi jika pendaftaran akun Supabase Auth gagal setelah kode terlanjur dikunci.
create or replace function batalkan_kode_lisensi(p_kode text)
returns void
language plpgsql
security definer
set search_path = public
as $$
begin
    update kode_lisensi
    set status = 'belum_terpakai'
    where kode = p_kode and dipakai_oleh_user_id is null;
end;
$$;

-- 3) Tautkan kode ke akun guru yang baru berhasil dibuat (permanen, tanpa masa berlaku).
create or replace function tautkan_kode_lisensi(p_kode text, p_user_id uuid, p_email text)
returns void
language plpgsql
security definer
set search_path = public
as $$
begin
    update kode_lisensi
    set dipakai_oleh_user_id = p_user_id,
        dipakai_oleh_email = p_email,
        digunakan_tanggal = now()
    where kode = p_kode;
end;
$$;

grant execute on function reserve_kode_lisensi(text) to anon, authenticated;
grant execute on function batalkan_kode_lisensi(text) to anon, authenticated;
grant execute on function tautkan_kode_lisensi(text, uuid, text) to anon, authenticated;

-- ============================================================
-- TAMBAHAN: ABSENSI BARCODE (data ringkas, tidak menyimpan foto -> hemat storage)
-- Jalankan bagian ini di SQL Editor SETELAH skema-skema di atas.
-- ============================================================

create table if not exists absensi_barcode (
    id bigint generated always as identity primary key,
    user_id uuid references auth.users(id) on delete cascade not null,
    nisn text not null,
    nama text,
    kelas text,
    tanggal date not null default current_date,
    waktu timestamptz default now(),
    status text not null default 'Hadir',  -- Hadir | Sakit | Izin | Alpa
    unique (user_id, nisn, tanggal)
);

alter table absensi_barcode enable row level security;

create policy "guru hanya akses absensi barcode miliknya sendiri"
    on absensi_barcode for all
    using (auth.uid() = user_id)
    with check (auth.uid() = user_id);

create index if not exists idx_absensi_barcode_user_tanggal
    on absensi_barcode (user_id, tanggal);
