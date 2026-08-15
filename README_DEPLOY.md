# Panduan Membuat Aplikasi Ini Online

## 1. Cara Tercepat & Gratis — Streamlit Community Cloud
1. Buat akun GitHub (github.com), buat repository baru, unggah `app.py` dan `requirements.txt`.
2. Buka https://share.streamlit.io, login dengan akun GitHub Anda.
3. Klik "New app", pilih repository dan file `app.py`, lalu klik Deploy.
4. Dalam beberapa menit aplikasi punya alamat publik, contoh:
   `https://nama-aplikasi-anda.streamlit.app`
5. Bagikan tautan tersebut ke rekan guru lain — mereka tinggal buka lewat browser HP/laptop,
   tanpa perlu instal apa pun.

Catatan: setiap guru tetap memasukkan API Key Groq miliknya sendiri di sidebar (gratis di
console.groq.com), jadi Anda tidak perlu menanggung biaya pemakaian AI orang lain.

## 2. Alternatif Berbayar/Lebih Fleksibel
Jika ingin domain sendiri, kontrol server penuh, atau menyimpan data (jurnal/absen) secara
permanen di database, pertimbangkan:
- **Railway.app** atau **Render.com** — hosting Streamlit dengan paket gratis terbatas & berbayar.
- **VPS (misal DigitalOcean/Niagahoster)** — jalankan `streamlit run app.py` dengan `nginx` +
  domain sendiri, cocok jika ingin skala lebih besar atau menjual sebagai layanan berlangganan.

## 3. Login Multi-Guru & Data Tersimpan Permanen (SUDAH TERSEDIA di app.py)
Fitur ini sudah ditambahkan ke aplikasi menggunakan **Supabase** (database gratis dengan sistem
login bawaan). Setiap guru daftar akun sendiri (email + kata sandi), dan data Jurnal Mengajar +
Data Siswa miliknya tersimpan permanen serta **privat** (guru lain tidak bisa melihatnya).

### Langkah Setup (sekali saja, ±10 menit)
1. **Buat akun & proyek Supabase**
   - Buka https://supabase.com → Sign up (gratis) → "New Project".
   - Beri nama proyek bebas, buat kata sandi database (simpan baik-baik), pilih region terdekat
     (misal Singapore), lalu klik "Create new project". Tunggu ±2 menit sampai siap.

2. **Jalankan skema database**
   - Di dashboard proyek, buka menu **SQL Editor** → **New query**.
   - Salin seluruh isi berkas `supabase_schema.sql` (disertakan bersama app.py ini), tempel, lalu
     klik **Run**. Ini akan membuat tabel `jurnal_mengajar` dan `data_siswa` beserta aturan
     keamanan (RLS) agar data tiap guru terpisah otomatis.

3. **Ambil URL & API Key proyek**
   - Buka menu **Settings** (⚙️) → **API**.
   - Salin **Project URL** dan **anon public key** (bukan `service_role`, cukup `anon public`).

4. **Masukkan ke aplikasi**
   - **Jika di komputer sendiri (lokal)**: buat folder `.streamlit` di sebelah `app.py`, lalu buat
     berkas `.streamlit/secrets.toml` isinya:
     ```toml
     SUPABASE_URL = "https://xxxxxxxx.supabase.co"
     SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9....(anon public key)"
     ```
   - **Jika di Streamlit Community Cloud**: buka pengaturan aplikasi Anda → **Settings** →
     **Secrets**, tempel dua baris di atas, lalu **Save**. Aplikasi otomatis restart.

5. **Selesai!** Buka aplikasi lagi — akan muncul halaman **Masuk / Daftar Akun Baru**. Guru lain
   tinggal daftar akun sendiri lewat halaman itu, data mereka otomatis terpisah dan tersimpan
   permanen di database Anda.

> **Catatan:** Jika `SUPABASE_URL`/`SUPABASE_KEY` belum diisi, aplikasi tetap berjalan seperti
> biasa TANPA login (data hanya tersimpan selama sesi browser terbuka) — jadi aman dicoba dulu
> tanpa setup database.

> **Verifikasi email (opsional):** Secara default Supabase mengirim email konfirmasi saat
> pendaftaran. Untuk mempermudah guru (tanpa perlu cek email), Anda bisa menonaktifkannya di
> **Authentication → Providers → Email → matikan "Confirm email"**.

### Opsi Tambahan yang Bisa Dibuatkan Berikutnya
- **Kode akses (access code)** tambahan di halaman daftar — supaya hanya guru yang punya kode
  dari Anda yang bisa membuat akun (cocok untuk versi berbayar/terbatas).
- **Watermark/branding** nama Anda sebagai pembuat aplikasi di setiap dokumen hasil unduhan.
- **Dashboard admin** untuk melihat daftar guru yang terdaftar & memantau pemakaian.

## 4. Yang Perlu Diperhatikan
- Jangan menaruh API Key Groq pribadi Anda langsung di kode — biarkan tiap pengguna mengisi
  sendiri di sidebar seperti sekarang, agar aman dan biaya pemakaian tidak dibebankan ke Anda.
- `SUPABASE_KEY` (anon public key) aman dimasukkan ke `secrets.toml`/Secrets — kunci ini memang
  didesain untuk dipakai di sisi aplikasi, karena akses data tetap dibatasi oleh RLS per akun.
  Jangan pernah memakai/membagikan `service_role key` (kunci itu punya akses penuh tanpa batas).

