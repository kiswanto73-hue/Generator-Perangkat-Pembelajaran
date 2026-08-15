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
