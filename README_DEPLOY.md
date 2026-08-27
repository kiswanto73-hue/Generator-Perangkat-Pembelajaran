Panduan Membuat Aplikasi Ini Online
1. Cara Tercepat & Gratis — Streamlit Community Cloud
Buat akun GitHub (github.com), buat repository baru, unggah `app.py` dan `requirements.txt`.
Buka https://share.streamlit.io, login dengan akun GitHub Anda.
Klik "New app", pilih repository dan file `app.py`, lalu klik Deploy.
Dalam beberapa menit aplikasi punya alamat publik, contoh:
`https://nama-aplikasi-anda.streamlit.app`
Bagikan tautan tersebut ke rekan guru lain — mereka tinggal buka lewat browser HP/laptop,
tanpa perlu instal apa pun.
Catatan: setiap guru tetap memasukkan API Key Groq miliknya sendiri di sidebar (gratis di
console.groq.com), jadi Anda tidak perlu menanggung biaya pemakaian AI orang lain.
2. Alternatif Berbayar/Lebih Fleksibel
Jika ingin domain sendiri, kontrol server penuh, atau menyimpan data (jurnal/absen) secara
permanen di database, pertimbangkan:
Railway.app atau Render.com — hosting Streamlit dengan paket gratis terbatas & berbayar.
VPS (misal DigitalOcean/Niagahoster) — jalankan `streamlit run app.py` dengan `nginx` +
domain sendiri, cocok jika ingin skala lebih besar atau menjual sebagai layanan berlangganan.
3. Login Multi-Guru & Data Tersimpan Permanen (SUDAH TERSEDIA di app.py)
Fitur ini sudah ditambahkan ke aplikasi menggunakan Supabase (database gratis dengan sistem
login bawaan). Setiap guru daftar akun sendiri (email + kata sandi), dan data Jurnal Mengajar +
Data Siswa miliknya tersimpan permanen serta privat (guru lain tidak bisa melihatnya).
Langkah Setup (sekali saja, ±10 menit)
Buat akun & proyek Supabase
Buka https://supabase.com → Sign up (gratis) → "New Project".
Beri nama proyek bebas, buat kata sandi database (simpan baik-baik), pilih region terdekat
(misal Singapore), lalu klik "Create new project". Tunggu ±2 menit sampai siap.
Jalankan skema database
Di dashboard proyek, buka menu SQL Editor → New query.
Salin seluruh isi berkas `supabase_schema.sql` (disertakan bersama app.py ini), tempel, lalu
klik Run. Ini akan membuat tabel `jurnal_mengajar` dan `data_siswa` beserta aturan
keamanan (RLS) agar data tiap guru terpisah otomatis.
Ambil URL & API Key proyek
Buka menu Settings (⚙️) → API.
Salin Project URL dan anon public key (bukan `service_role`, cukup `anon public`).
Masukkan ke aplikasi
Jika di komputer sendiri (lokal): buat folder `.streamlit` di sebelah `app.py`, lalu buat
berkas `.streamlit/secrets.toml` isinya:
```toml
     SUPABASE_URL = "https://xxxxxxxx.supabase.co"
     SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9....(anon public key)"
     ```
Jika di Streamlit Community Cloud: buka pengaturan aplikasi Anda → Settings →
Secrets, tempel dua baris di atas, lalu Save. Aplikasi otomatis restart.
Selesai! Buka aplikasi lagi — akan muncul halaman Masuk / Daftar Akun Baru. Guru lain
tinggal daftar akun sendiri lewat halaman itu, data mereka otomatis terpisah dan tersimpan
permanen di database Anda.
> **Catatan:** Jika `SUPABASE_URL`/`SUPABASE_KEY` belum diisi, aplikasi tetap berjalan seperti
> biasa TANPA login (data hanya tersimpan selama sesi browser terbuka) — jadi aman dicoba dulu
> tanpa setup database.
> **Verifikasi email (opsional):** Secara default Supabase mengirim email konfirmasi saat
> pendaftaran. Untuk mempermudah guru (tanpa perlu cek email), Anda bisa menonaktifkannya di
> **Authentication → Providers → Email → matikan "Confirm email"**.
Opsi Tambahan yang Bisa Dibuatkan Berikutnya
Watermark/branding nama Anda sebagai pembuat aplikasi di setiap dokumen hasil unduhan.
5. Kode Lisensi (Akses Berbayar) & Panel Admin — SUDAH TERSEDIA
Sekarang pendaftaran akun WAJIB memakai kode aktivasi dari Anda. Satu kode = satu akun guru,
berlaku selamanya (tanpa masa kadaluarsa). Ada juga panel admin khusus Anda untuk membuat
dan memantau kode-kode tersebut.
Langkah Setup (sekali saja, ±5 menit, dilakukan SETELAH langkah bagian 3 di atas)
Jalankan tambahan skema database
Buka lagi SQL Editor di Supabase → New query.
Salin bagian "TAMBAHAN: SISTEM KODE LISENSI" di berkas `supabase_schema.sql` (paling
bawah), tempel, lalu Run. Ini membuat tabel `kode_lisensi` beserta 3 fungsi keamanan
yang mengatur pemakaian kode secara aman (anti-dipakai-dua-kali).
Ambil Service Role Key (kunci rahasia khusus admin — JANGAN PERNAH dibagikan ke siapa pun)
Di Supabase: Settings → API → salin key di bagian `service_role` (bukan `anon`).
Tentukan Kata Sandi Panel Admin
Bebas Anda tentukan sendiri, contoh: `admin-kkg-2026-rahasia`.
Tambahkan ke Secrets (lokal di `.streamlit/secrets.toml`, atau di Streamlit Cloud →
Settings → Secrets), tambahkan 2 baris baru di bawah `SUPABASE_URL`/`SUPABASE_KEY` yang sudah ada:
```toml
   SUPABASE_SERVICE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9....(service_role key)"
   ADMIN_PASSWORD = "admin-kkg-2026-rahasia"
   ```
Buka Panel Admin
Tambahkan `?admin=1` di akhir alamat aplikasi Anda, contoh:
`https://nama-aplikasi-anda.streamlit.app/?admin=1`
Masukkan kata sandi admin → generate kode sebanyak yang dibutuhkan → salin kode → kirim ke
guru yang sudah membeli/berlangganan.
Halaman ini tidak muncul di navigasi biasa, jadi guru lain tidak akan melihatnya.
Alur guru baru: buka link biasa (tanpa `?admin=1`) → tab "📝 Daftar Akun Baru" → isi
nama, email, kata sandi, + kode aktivasi dari Anda → akun langsung aktif permanen.
> **Keamanan:** `SUPABASE_SERVICE_KEY` hanya dipakai di sisi server (Streamlit Secrets), tidak
> pernah terkirim ke browser guru — aman selama Anda tidak membagikan isi Secrets ke orang lain.
> Tabel `kode_lisensi` sendiri terkunci total dari akses langsung (Row Level Security tanpa
> policy), hanya bisa diubah lewat fungsi khusus atau Service Key Anda.
6. Notifikasi WhatsApp Otomatis Saat Ada Pendaftar Baru — SUDAH TERSEDIA
Setiap kali ada guru berhasil mendaftar, Anda otomatis dapat pesan WA berisi nama, email, dan
kode aktivasi yang dipakai — jadi tidak perlu bolak-balik cek Panel Admin.
Langkah Setup (±5 menit)
Daftar akun Fonnte di https://fonnte.com (ada paket gratis terbatas untuk mulai mencoba).
Hubungkan nomor WA Anda lewat scan QR di dashboard Fonnte, lalu salin Token perangkat
yang muncul di dashboard.
Tambahkan ke Secrets (bersama secrets lain yang sudah ada):
```toml
   TOKEN_FONNTE = "isi_token_dari_dashboard_fonnte"
   NOMOR_WA_ADMIN = "6281234567890"
   ```
Nomor WA ditulis format internasional (awali `62`, tanpa tanda `+` atau `0` di depan).
Selesai — deploy ulang aplikasi, lalu coba daftar dengan akun uji untuk memastikan notifikasi masuk.
> Jika `TOKEN_FONNTE`/`NOMOR_WA_ADMIN` belum diisi, fitur ini otomatis nonaktif dan pendaftaran
> guru tetap berjalan normal seperti biasa (tidak ada error yang muncul ke guru).
7. Sistem Absen Kartu Barcode — SUDAH TERSEDIA
Ada di tab "Data Siswa & Lembar Absensi", sekarang terbagi jadi 4 sub-menu:
📋 Data Siswa & Lembar Absen — seperti sebelumnya (input/impor data siswa).
🖨️ Cetak Kartu Barcode — pilih siswa (harus sudah ada NISN-nya di data siswa), lalu unduh
PDF berisi kartu (8 kartu per halaman A4): nama, kelas, NISN, barcode, dan kotak putus-putus
untuk tempel foto manual (seperti kartu pelajar biasa — foto tidak diunggah ke sistem,
jadi tidak menambah beban database sama sekali).
📷 Scan Absen — dua cara discan, keduanya tersedia sekaligus:
Scanner fisik (USB/Bluetooth): klik kotak input, tembak barcode kartu, otomatis tercatat.
Kamera HP: arahkan kamera langsung ke barcode di kartu lewat browser.
Ada juga "Tandai Manual" untuk Sakit/Izin/Alpa (siswa yang tidak discan).
📊 Rekap Bulanan — pilih bulan, lihat rekap Hadir/Sakit/Izin/Alpa per siswa + persentase
kehadiran, bisa diunduh sebagai Excel.
Langkah Setup
Jalankan bagian SQL "TAMBAHAN: ABSENSI BARCODE" di `supabase_schema.sql` (paling bawah)
lewat SQL Editor Supabase.
Pastikan data siswa yang mau dicetak kartunya sudah diisi NISN-nya (kolom NISN saat
tambah/impor siswa).
Fitur Scan Absen & Rekap Bulanan butuh guru sudah login (kehadiran tersimpan per akun).
Catatan tentang kamera HP
Scan pakai kamera memakai library `html5-qrcode` yang jalan langsung di browser (tidak perlu
instalasi apa pun oleh guru), lalu memicu Streamlit memproses hasilnya. Sebagian browser (terutama
di HP tertentu) kadang perlu izin akses kamera diklik dulu, atau ada batasan keamanan browser yang
membuatnya tidak langsung jalan mulus di semua perangkat — kalau kamera bermasalah, scanner fisik
selalu jadi cadangan yang pasti berfungsi. Disarankan uji coba dulu di beberapa HP setelah deploy.
Kenapa fitur ini tidak bikin database bengkak
Satu baris absensi cuma berisi NISN, nama, kelas, tanggal, status — beberapa ratus byte saja.
1.000 siswa × 200 hari absen setahun ≈ 200.000 baris ≈ masih di bawah 50 MB, jauh dari batas
500 MB paket gratis Supabase. Yang perlu dipantau justru pertumbuhan jangka panjang (jurnal +
absensi bertahun-tahun) — kalau nanti sudah banyak pelanggan, beri tahu saya, saya bisa buatkan
fitur arsip otomatis per tahun ajaran agar database tetap ringan.
4. Yang Perlu Diperhatikan
Jangan menaruh API Key Groq pribadi Anda langsung di kode — biarkan tiap pengguna mengisi
sendiri di sidebar seperti sekarang, agar aman dan biaya pemakaian tidak dibebankan ke Anda.
`SUPABASE_KEY` (anon public key) aman dimasukkan ke `secrets.toml`/Secrets — kunci ini memang
didesain untuk dipakai di sisi aplikasi, karena akses data tetap dibatasi oleh RLS per akun.
Jangan pernah memakai/membagikan `service_role key` (kunci itu punya akses penuh tanpa batas).
