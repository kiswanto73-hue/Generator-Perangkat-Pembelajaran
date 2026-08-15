import streamlit as st
from groq import Groq
import re
import pandas as pd
from datetime import datetime, date, timedelta
from io import BytesIO

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, Image as RLImage
)

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

try:
    from pypdf import PdfReader
except Exception:
    PdfReader = None

try:
    from supabase import create_client
except Exception:
    create_client = None

# ============================================================
# KONFIGURASI HALAMAN
# ============================================================
st.set_page_config(page_title="Generator Perangkat Ajar KKG", layout="wide", page_icon="📚")

MODEL_GROQ = "llama-3.3-70b-versatile"

COLOR_PRIMARY = colors.HexColor("#1E4D8C")
COLOR_ACCENT = colors.HexColor("#F2A900")
COLOR_LIGHT = colors.HexColor("#EEF3FA")
DOCX_PRIMARY_RGB = RGBColor(0x1E, 0x4D, 0x8C)
DOCX_GREY_RGB = RGBColor(0x55, 0x55, 0x55)

# ============================================================
# TEMA VISUAL - CSS ELEGAN (diterapkan ke semua halaman/tab)
# ============================================================
st.markdown("""
<style>
    .stApp { background: linear-gradient(180deg, #F7F9FC 0%, #FFFFFF 100%); }
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1E4D8C 0%, #163a6b 100%);
    }
    section[data-testid="stSidebar"] * { color: #F2F5FA !important; }
    section[data-testid="stSidebar"] input, section[data-testid="stSidebar"] textarea {
        color: #16324F !important;
    }
    h1, h2, h3 { color: #1E4D8C; font-family: "Source Sans Pro", sans-serif; }
    div.stButton > button {
        background: linear-gradient(135deg, #1E4D8C 0%, #2E6BB8 100%);
        color: white; border: none; border-radius: 8px; font-weight: 600;
        padding: 0.5rem 1.1rem; transition: 0.15s ease-in-out;
    }
    div.stButton > button:hover { transform: translateY(-1px); box-shadow: 0 4px 10px rgba(30,77,140,0.35); }
    div[data-testid="stExpander"], div[data-baseweb="tab-panel"] {
        background: #FFFFFF; border-radius: 12px;
    }
    .stTabs [data-baseweb="tab-list"] { gap: 6px; }
    .stTabs [data-baseweb="tab"] {
        background-color: #EEF3FA; border-radius: 8px 8px 0 0; padding: 8px 16px; font-weight: 600;
    }
    .stTabs [aria-selected="true"] { background-color: #1E4D8C !important; color: white !important; }
    div[data-testid="stMetric"], .kkg-card {
        background: #FFFFFF; border: 1px solid #E3E9F2; border-radius: 12px;
        padding: 14px 18px; box-shadow: 0 2px 8px rgba(30,77,140,0.06);
    }
</style>
""", unsafe_allow_html=True)

# ============================================================
# PEDOMAN PEMBELAJARAN MENDALAM (DEEP LEARNING)
# Permendikdasmen No. 13 Tahun 2025 - penyempurnaan Permendikbudristek No. 12/2024
# ============================================================
DEEP_LEARNING_GUIDE = """
PEDOMAN WAJIB - PARADIGMA PEMBELAJARAN MENDALAM (DEEP LEARNING)
Sesuai Permendikdasmen No. 13 Tahun 2025, seluruh rancangan pembelajaran harus
menghadirkan tiga prinsip berikut secara EKSPLISIT (beri label di setiap langkah kegiatan inti):

1. MINDFUL (Berkesadaran): peserta didik diajak fokus penuh, sadar akan tujuan
   belajarnya, serta diberi ruang refleksi diri di awal/akhir kegiatan.
2. MEANINGFUL (Bermakna): materi dikaitkan secara kontekstual dengan pengalaman
   nyata dan kehidupan sehari-hari murid, bukan sekadar hafalan.
3. JOYFUL (Menyenangkan): suasana belajar positif, melibatkan unsur kolaborasi,
   permainan, eksplorasi, atau apresiasi karya, sehingga murid termotivasi.

Setiap langkah pada bagian Kegiatan Inti WAJIB diberi penanda, contoh:
"Guru mengajak murid duduk tenang sejenak dan menyebutkan tujuan belajar hari ini. (Mindful)"
"""

SUBJECT_OPTIONS = ["PAI & BP", "Bahasa Indonesia", "Matematika", "IPAS", "PJOK",
                    "Pendidikan Pancasila", "Seni Rupa", "Bahasa Inggris", "Lainnya"]

# Input Kelas 1-6 (setiap kelas otomatis dipasangkan dengan Fase Kurikulum Merdeka)
KELAS_FASE_MAP = {
    "Kelas 1": "Fase A", "Kelas 2": "Fase A",
    "Kelas 3": "Fase B", "Kelas 4": "Fase B",
    "Kelas 5": "Fase C", "Kelas 6": "Fase C",
}
KELAS_OPTIONS = list(KELAS_FASE_MAP.keys())


def label_kelas_fase(kelas: str) -> str:
    return f"{kelas} ({KELAS_FASE_MAP.get(kelas, '')})"


# ============================================================
# STATE AWAL
# ============================================================
defaults = {
    "sekolah": "SDN 165/V Tanjung Jabung Barat",
    "penyusun": "",
    "guru_nip": "",
    "kepsek_nama": "",
    "kepsek_nip": "",
    "tahun_ajaran": "2025/2026",
    "hasil_tp_atp": "", "meta_tp_atp": {},
    "hasil_modul": "", "meta_modul": {},
    "hasil_lkpd": "", "meta_lkpd": {}, "lkpd_gambar_bytes": None,
    "hasil_prota_promes": "", "meta_prota_promes": {}, "judul_prota_promes": "PROGRAM TAHUNAN (PROTA)",
    "hasil_minggu_efektif": "", "meta_minggu_efektif": {}, "tabel_minggu_efektif": None,
    "jurnal_rows": [],
    "siswa_rows": [],
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v
if "user" not in st.session_state:
    st.session_state.user = None
if "data_loaded" not in st.session_state:
    st.session_state.data_loaded = False


# ============================================================
# AUTENTIKASI & PENYIMPANAN PERMANEN (SUPABASE) - OPSIONAL
# Aktif otomatis jika SUPABASE_URL & SUPABASE_KEY diisi di .streamlit/secrets.toml.
# Jika tidak diisi, aplikasi tetap berjalan seperti biasa (data hanya tersimpan
# selama sesi browser terbuka, tanpa login) - lihat README_DEPLOY.md.
# ============================================================
SUPABASE_URL = st.secrets.get("SUPABASE_URL", "") if hasattr(st, "secrets") else ""
SUPABASE_KEY = st.secrets.get("SUPABASE_KEY", "") if hasattr(st, "secrets") else ""
AUTH_AKTIF = bool(SUPABASE_URL and SUPABASE_KEY and create_client is not None)


@st.cache_resource
def get_supabase():
    return create_client(SUPABASE_URL, SUPABASE_KEY)


def sb_daftar(email: str, password: str, nama: str):
    sb = get_supabase()
    res = sb.auth.sign_up({"email": email, "password": password,
                            "options": {"data": {"nama_lengkap": nama}}})
    return res


def sb_masuk(email: str, password: str):
    sb = get_supabase()
    res = sb.auth.sign_in_with_password({"email": email, "password": password})
    return res


def sb_keluar():
    try:
        get_supabase().auth.sign_out()
    except Exception:
        pass
    st.session_state.user = None
    st.session_state.data_loaded = False
    st.session_state.jurnal_rows = []
    st.session_state.siswa_rows = []


def db_muat_data(user_id: str):
    """Muat jurnal & data siswa milik guru yang sedang login dari Supabase."""
    sb = get_supabase()
    try:
        j = sb.table("jurnal_mengajar").select("*").eq("user_id", user_id).order("id").execute()
        st.session_state.jurnal_rows = [
            {"id": r["id"], "Tanggal": r["tanggal"], "Kelas": r["kelas"], "Mapel": r["mapel"],
             "JP": r["jp"], "Materi": r["materi"], "Kegiatan": r["kegiatan"],
             "Hadir": r["hadir"], "Catatan": r["catatan"]}
            for r in (j.data or [])
        ]
    except Exception as e:
        st.warning(f"⚠️ Gagal memuat data jurnal dari database: {e}")
    try:
        s = sb.table("data_siswa").select("*").eq("user_id", user_id).order("id").execute()
        st.session_state.siswa_rows = [
            {"id": r["id"], "Nama": r["nama"], "L/P": r["jk"], "NISN": r["nisn"]}
            for r in (s.data or [])
        ]
    except Exception as e:
        st.warning(f"⚠️ Gagal memuat data siswa dari database: {e}")


def db_tambah_jurnal(user_id: str, row: dict):
    sb = get_supabase()
    payload = {"user_id": user_id, "tanggal": row["Tanggal"], "kelas": row["Kelas"],
               "mapel": row["Mapel"], "jp": row["JP"], "materi": row["Materi"],
               "kegiatan": row["Kegiatan"], "hadir": row["Hadir"], "catatan": row["Catatan"]}
    try:
        res = sb.table("jurnal_mengajar").insert(payload).execute()
        if res.data:
            row["id"] = res.data[0]["id"]
    except Exception as e:
        st.warning(f"⚠️ Gagal menyimpan jurnal ke database: {e}")
    return row


def db_hapus_jurnal(row_id):
    if row_id is None:
        return
    try:
        get_supabase().table("jurnal_mengajar").delete().eq("id", row_id).execute()
    except Exception as e:
        st.warning(f"⚠️ Gagal menghapus data di database: {e}")


def db_kosongkan_jurnal(user_id: str):
    try:
        get_supabase().table("jurnal_mengajar").delete().eq("user_id", user_id).execute()
    except Exception as e:
        st.warning(f"⚠️ Gagal mengosongkan data di database: {e}")


def db_tambah_siswa(user_id: str, row: dict):
    sb = get_supabase()
    payload = {"user_id": user_id, "nama": row["Nama"], "jk": row["L/P"], "nisn": row.get("NISN", "")}
    try:
        res = sb.table("data_siswa").insert(payload).execute()
        if res.data:
            row["id"] = res.data[0]["id"]
    except Exception as e:
        st.warning(f"⚠️ Gagal menyimpan data siswa ke database: {e}")
    return row


def db_hapus_siswa(row_id):
    if row_id is None:
        return
    try:
        get_supabase().table("data_siswa").delete().eq("id", row_id).execute()
    except Exception as e:
        st.warning(f"⚠️ Gagal menghapus data di database: {e}")


def db_kosongkan_siswa(user_id: str):
    try:
        get_supabase().table("data_siswa").delete().eq("user_id", user_id).execute()
    except Exception as e:
        st.warning(f"⚠️ Gagal mengosongkan data di database: {e}")


def tambah_baris_jurnal(row: dict):
    if AUTH_AKTIF and st.session_state.user:
        row = db_tambah_jurnal(st.session_state.user["id"], row)
    st.session_state.jurnal_rows.append(row)


def hapus_jurnal_terakhir():
    if st.session_state.jurnal_rows:
        row = st.session_state.jurnal_rows.pop()
        if AUTH_AKTIF and st.session_state.user:
            db_hapus_jurnal(row.get("id"))


def kosongkan_jurnal():
    if AUTH_AKTIF and st.session_state.user:
        db_kosongkan_jurnal(st.session_state.user["id"])
    st.session_state.jurnal_rows = []


def tambah_baris_siswa(row: dict):
    if AUTH_AKTIF and st.session_state.user:
        row = db_tambah_siswa(st.session_state.user["id"], row)
    st.session_state.siswa_rows.append(row)


def hapus_siswa_terakhir():
    if st.session_state.siswa_rows:
        row = st.session_state.siswa_rows.pop()
        if AUTH_AKTIF and st.session_state.user:
            db_hapus_siswa(row.get("id"))


def kosongkan_siswa():
    if AUTH_AKTIF and st.session_state.user:
        db_kosongkan_siswa(st.session_state.user["id"])
    st.session_state.siswa_rows = []


def tampilkan_gerbang_login():
    """Tampilkan formulir Masuk/Daftar. Menghentikan eksekusi (st.stop()) sampai berhasil login."""
    st.title("📚 Generator Perangkat Ajar Kurikulum Merdeka")
    st.caption("Silakan masuk atau daftar akun guru untuk mulai menggunakan aplikasi. "
               "Data jurnal & absen Anda akan tersimpan permanen dan hanya bisa diakses oleh akun Anda sendiri.")
    tab_masuk, tab_daftar = st.tabs(["🔑 Masuk", "📝 Daftar Akun Baru"])

    with tab_masuk:
        with st.form("form_masuk"):
            email_masuk = st.text_input("Email", key="email_masuk")
            pw_masuk = st.text_input("Kata Sandi", type="password", key="pw_masuk")
            submit_masuk = st.form_submit_button("🔑 Masuk", use_container_width=True)
        if submit_masuk:
            if not email_masuk or not pw_masuk:
                st.error("⚠️ Mohon isi email dan kata sandi.")
            else:
                try:
                    res = sb_masuk(email_masuk, pw_masuk)
                    if res.user:
                        st.session_state.user = {"id": res.user.id, "email": res.user.email}
                        st.success("✅ Berhasil masuk!")
                        st.rerun()
                    else:
                        st.error("❌ Email atau kata sandi salah.")
                except Exception as e:
                    st.error(f"❌ Gagal masuk: {e}")

    with tab_daftar:
        with st.form("form_daftar"):
            nama_daftar = st.text_input("Nama Lengkap", key="nama_daftar")
            email_daftar = st.text_input("Email", key="email_daftar")
            pw_daftar = st.text_input("Kata Sandi (minimal 6 karakter)", type="password", key="pw_daftar")
            submit_daftar = st.form_submit_button("📝 Daftar", use_container_width=True)
        if submit_daftar:
            if not nama_daftar or not email_daftar or not pw_daftar:
                st.error("⚠️ Mohon lengkapi semua isian.")
            elif len(pw_daftar) < 6:
                st.error("⚠️ Kata sandi minimal 6 karakter.")
            else:
                try:
                    res = sb_daftar(email_daftar, pw_daftar, nama_daftar)
                    if res.user:
                        st.success("✅ Pendaftaran berhasil! Jika verifikasi email diaktifkan, "
                                    "silakan cek email Anda terlebih dahulu, lalu masuk lewat tab 'Masuk'.")
                    else:
                        st.error("❌ Pendaftaran gagal, coba lagi.")
                except Exception as e:
                    st.error(f"❌ Gagal mendaftar: {e}")
    st.stop()


if AUTH_AKTIF and st.session_state.user is None:
    tampilkan_gerbang_login()

if AUTH_AKTIF and st.session_state.user is not None and not st.session_state.data_loaded:
    db_muat_data(st.session_state.user["id"])
    st.session_state.data_loaded = True


# ============================================================
# FUNGSI EKSTRAKSI FILE REFERENSI (Buku Guru PDF/DOCX/TXT/XLSX/CSV)
# ============================================================
def ekstrak_teks_referensi(uploaded_file, max_chars: int = 8000) -> tuple:
    """Ambil cuplikan teks dari file referensi (PDF/DOCX/TXT/XLSX/CSV) untuk dijadikan
    konteks tambahan bagi AI. Mengembalikan (teks, pesan_status)."""
    if uploaded_file is None:
        return "", ""
    nama = uploaded_file.name.lower()
    try:
        data = uploaded_file.getvalue()  # ambil bytes secara utuh agar tidak bentrok posisi pointer
    except Exception as e:
        return "", f"❌ Gagal membaca berkas: {e}"

    try:
        if nama.endswith(".pdf"):
            if PdfReader is None:
                return "", "❌ Modul pembaca PDF tidak tersedia."
            reader = PdfReader(BytesIO(data))
            n_halaman = len(reader.pages)
            potongan = []
            for page in reader.pages[:30]:
                try:
                    t = page.extract_text() or ""
                except Exception:
                    t = ""
                if t.strip():
                    potongan.append(t)
            teks = "\n".join(potongan).strip()
            if not teks:
                return "", (f"⚠️ Berkas PDF ({n_halaman} halaman) terbaca tapi TIDAK ADA teks yang bisa "
                             f"diekstrak — kemungkinan ini adalah hasil SCAN/FOTO (gambar), bukan PDF teks. "
                             f"Silakan gunakan PDF asli (bukan hasil scan) atau salin manual isinya ke kolom teks.")
            return teks[:max_chars], f"✅ Berhasil membaca {len(teks)} karakter dari {n_halaman} halaman PDF."

        if nama.endswith(".docx"):
            doc = Document(BytesIO(data))
            bagian = [p.text for p in doc.paragraphs if p.text.strip()]
            for tabel in doc.tables:
                for row in tabel.rows:
                    bagian.append(" | ".join(c.text for c in row.cells))
            teks = "\n".join(bagian).strip()
            if not teks:
                return "", "⚠️ Berkas DOCX terbaca tapi tidak ada teks di dalamnya (kemungkinan hanya berisi gambar)."
            return teks[:max_chars], f"✅ Berhasil membaca {len(teks)} karakter dari berkas DOCX."

        if nama.endswith(".doc"):
            return "", ("❌ Format .doc (Word lama) belum didukung. Mohon simpan ulang berkas sebagai "
                        ".docx terlebih dahulu (Save As > Word Document .docx) lalu unggah kembali.")

        if nama.endswith(".txt"):
            teks = data.decode("utf-8", errors="ignore").strip()
            if not teks:
                return "", "⚠️ Berkas TXT kosong."
            return teks[:max_chars], f"✅ Berhasil membaca {len(teks)} karakter dari berkas TXT."

        if nama.endswith(".csv") or nama.endswith(".xlsx") or nama.endswith(".xls"):
            import pandas as pd
            if nama.endswith(".csv"):
                df = pd.read_csv(BytesIO(data))
            else:
                df = pd.read_excel(BytesIO(data))
            if df.empty:
                return "", "⚠️ Berkas tabel terbaca tapi tidak ada data di dalamnya."
            teks = df.to_string(index=False)
            return teks[:max_chars], f"✅ Berhasil membaca {len(df)} baris data dari berkas tabel."

        if nama.endswith((".jpg", ".jpeg", ".png")):
            return "", ("⚠️ Berkas berupa gambar. AI teks tidak dapat membaca isi gambar secara otomatis — "
                        "gambar tetap bisa dipakai sebagai ilustrasi (jika fitur ini menyediakan unggah gambar), "
                        "namun isinya tidak dijadikan konteks teks.")
    except Exception as e:
        return "", f"❌ Gagal mengekstrak isi berkas ({uploaded_file.name}): {e}"

    return "", "⚠️ Format berkas tidak didukung."


def uploader_referensi(key: str, label: str = "📎 Unggah Referensi Tambahan (Buku Guru / Materi - PDF, DOCX, TXT, opsional)",
                        tipe: list = None):
    """Widget unggah file referensi yang dipakai untuk memperkaya konteks prompt AI.
    Menampilkan status pembacaan secara jelas + cuplikan isi agar guru bisa memastikan berkas terbaca."""
    tipe = tipe or ["pdf", "docx", "txt"]
    f = st.file_uploader(label, type=tipe, key=key)
    if f is None:
        return ""
    with st.spinner(f"Membaca isi berkas {f.name}..."):
        teks, status = ekstrak_teks_referensi(f)
    if teks:
        st.success(status)
        with st.expander(f"👀 Pratinjau isi yang terbaca dari '{f.name}' (klik untuk lihat)"):
            st.text(teks[:2000] + ("..." if len(teks) > 2000 else ""))
        return teks
    else:
        st.warning(status or f"⚠️ Tidak dapat membaca isi berkas '{f.name}'.")
        return ""


def baca_tabel_upload(uploaded_file):
    """Baca berkas CSV/XLSX menjadi list of dict (baris), dipakai untuk impor massal
    data siswa / jurnal / kaldik. Mengembalikan (list_baris, pesan_status)."""
    if uploaded_file is None:
        return [], ""
    nama = uploaded_file.name.lower()
    try:
        import pandas as pd
        data = uploaded_file.getvalue()
        if nama.endswith(".csv"):
            df = pd.read_csv(BytesIO(data))
        elif nama.endswith(".xlsx") or nama.endswith(".xls"):
            df = pd.read_excel(BytesIO(data))
        elif nama.endswith(".txt"):
            teks = data.decode("utf-8", errors="ignore")
            baris = [{"Nama": b.strip()} for b in teks.split("\n") if b.strip()]
            return baris, f"✅ {len(baris)} baris terbaca dari TXT."
        else:
            return [], "⚠️ Format tidak didukung untuk impor tabel (gunakan CSV/XLSX/TXT)."
        df = df.fillna("")
        baris = df.to_dict(orient="records")
        return baris, f"✅ {len(baris)} baris data berhasil dibaca dari '{uploaded_file.name}'."
    except Exception as e:
        return [], f"❌ Gagal membaca berkas '{uploaded_file.name}': {e}"


# ============================================================
# FUNGSI PEMANGGILAN GROQ
# ============================================================
def call_groq(api_key: str, prompt: str) -> str:
    client = Groq(api_key=api_key)
    completion = client.chat.completions.create(
        model=MODEL_GROQ,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.6,
        max_completion_tokens=8192,
    )
    return completion.choices[0].message.content


# ============================================================
# PARSER MARKDOWN BERSAMA (dipakai untuk PDF & DOCX)
# ============================================================
def parse_markdown_blocks(md_text: str) -> list:
    """Mengubah teks markdown dari LLM menjadi daftar blok terstruktur:
    ('h1'|'h2'|'h3', text) / ('bullet'|'numbered'|'para', text) / ('table', rows)"""
    blocks = []
    lines = md_text.replace("\r\n", "\n").split("\n")
    i, n = 0, len(lines)

    while i < n:
        line = lines[i].strip()

        if not line:
            i += 1
            continue

        if line.startswith("|"):
            table_rows = []
            while i < n and lines[i].strip().startswith("|"):
                row_line = lines[i].strip()
                if not table_rows and i + 1 < n and re.match(r"^\|?[\s:\-|]+\|?$", lines[i + 1].strip()):
                    i += 2
                    table_rows.append([c.strip() for c in row_line.strip("|").split("|")])
                    continue
                table_rows.append([c.strip() for c in row_line.strip("|").split("|")])
                i += 1
            if table_rows:
                blocks.append(("table", table_rows))
            continue

        if line.startswith("# "):
            blocks.append(("h1", line[2:])); i += 1; continue
        if line.startswith("## "):
            blocks.append(("h2", line[3:])); i += 1; continue
        if line.startswith("### "):
            blocks.append(("h3", line[4:])); i += 1; continue
        if re.match(r"^[-*]\s+", line):
            blocks.append(("bullet", re.sub(r"^[-*]\s+", "", line))); i += 1; continue
        if re.match(r"^\d+[.)]\s+", line):
            blocks.append(("numbered", re.sub(r"^\d+[.)]\s+", "", line))); i += 1; continue

        blocks.append(("para", line)); i += 1

    return blocks


# ---------- Rendering ke PDF (reportlab) ----------
def _inline_format_pdf(text: str) -> str:
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", text)
    text = re.sub(r"(?<!\*)\*(?!\*)(.+?)\*(?!\*)", r"<i>\1</i>", text)
    return text


def build_pdf_styles():
    base = getSampleStyleSheet()
    return {
        "h1": ParagraphStyle("H1", parent=base["Heading1"], fontSize=15,
                              textColor=COLOR_PRIMARY, spaceAfter=4, alignment=TA_LEFT),
        "h2": ParagraphStyle("H2", parent=base["Heading2"], fontSize=12.5,
                              textColor=COLOR_PRIMARY, spaceAfter=4, alignment=TA_LEFT),
        "h3": ParagraphStyle("H3", parent=base["Heading3"], fontSize=11,
                              textColor=colors.HexColor("#333333"), spaceAfter=3),
        "body": ParagraphStyle("Body", parent=base["Normal"], fontSize=10.3,
                                leading=15, alignment=TA_JUSTIFY, spaceAfter=5),
        "bullet": ParagraphStyle("Bullet", parent=base["Normal"], fontSize=10.3,
                                  leading=14, leftIndent=14, spaceAfter=3, alignment=TA_JUSTIFY),
        "cell": ParagraphStyle("Cell", parent=base["Normal"], fontSize=9.3, leading=12),
        "meta_label": ParagraphStyle("MetaLabel", parent=base["Normal"], fontSize=9.5,
                                      textColor=colors.HexColor("#555555"), alignment=TA_RIGHT),
        "meta_value": ParagraphStyle("MetaValue", parent=base["Normal"], fontSize=10,
                                      fontName="Helvetica-Bold"),
        "title": ParagraphStyle("DocTitle", parent=base["Heading1"], fontSize=18,
                                 alignment=TA_CENTER, textColor=COLOR_PRIMARY, spaceAfter=2),
        "subtitle": ParagraphStyle("DocSubtitle", parent=base["Normal"], fontSize=10.5,
                                    alignment=TA_CENTER, textColor=colors.HexColor("#555555"),
                                    spaceAfter=10),
    }


def blocks_to_pdf_flowables(blocks: list, styles: dict) -> list:
    flowables = []
    for kind, content in blocks:
        if kind == "table":
            para_rows = [[Paragraph(_inline_format_pdf(c), styles["cell"]) for c in row] for row in content]
            tbl = Table(para_rows, repeatRows=1, hAlign="LEFT")
            tbl.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), COLOR_PRIMARY),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.6, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, COLOR_LIGHT]),
                ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]))
            flowables.append(tbl)
            flowables.append(Spacer(1, 10))
        elif kind == "h1":
            flowables.append(Spacer(1, 6))
            flowables.append(Paragraph(_inline_format_pdf(content), styles["h1"]))
            flowables.append(HRFlowable(width="100%", thickness=1.2, color=COLOR_PRIMARY, spaceAfter=8))
        elif kind == "h2":
            flowables.append(Spacer(1, 8))
            flowables.append(Paragraph(_inline_format_pdf(content), styles["h2"]))
        elif kind == "h3":
            flowables.append(Spacer(1, 4))
            flowables.append(Paragraph(_inline_format_pdf(content), styles["h3"]))
        elif kind == "bullet":
            flowables.append(Paragraph("• " + _inline_format_pdf(content), styles["bullet"]))
        elif kind == "numbered":
            flowables.append(Paragraph(_inline_format_pdf(content), styles["bullet"]))
        else:
            flowables.append(Paragraph(_inline_format_pdf(content), styles["body"]))
    return flowables


def _pdf_footer(canvas, doc):
    canvas.saveState()
    canvas.setStrokeColor(COLOR_ACCENT)
    canvas.setLineWidth(2)
    canvas.line(2 * cm, 1.6 * cm, A4[0] - 2 * cm, 1.6 * cm)
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.HexColor("#777777"))
    canvas.drawString(2 * cm, 1.2 * cm,
                       f"Dibuat dengan Generator Perangkat Ajar KKG - {datetime.now().strftime('%d %B %Y')}")
    canvas.drawRightString(A4[0] - 2 * cm, 1.2 * cm, f"Halaman {doc.page}")
    canvas.restoreState()


def _meta_pairs(meta: dict):
    rows, pair = [], []
    for k, v in meta.items():
        if not v:
            continue
        pair.append((k, v))
        if len(pair) == 2:
            rows.append(pair); pair = []
    if pair:
        rows.append(pair)
    return rows


def generate_pdf(doc_title: str, meta: dict, markdown_text: str, gambar_bytes: bytes = None) -> bytes:
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        topMargin=1.8 * cm, bottomMargin=2.1 * cm, leftMargin=2 * cm, rightMargin=2 * cm,
        title=doc_title,
    )
    styles = build_pdf_styles()
    story = [
        Paragraph(doc_title, styles["title"]),
        Paragraph("Kurikulum Merdeka &bull; Pembelajaran Mendalam (Deep Learning)", styles["subtitle"]),
    ]

    rows = _meta_pairs(meta)
    if rows:
        table_data = []
        for pair in rows:
            row = []
            for k, v in pair:
                row.append(Paragraph(f"<b>{k}</b>", styles["meta_label"]))
                row.append(Paragraph(f": {v}", styles["meta_value"]))
            if len(pair) == 1:
                row += ["", ""]
            table_data.append(row)
        # Label rata-kanan pada kolom lebar tetap -> tanda titik dua otomatis lurus,
        # tanpa perlu kolom sempit terpisah (yang rawan error karena lebih kecil dari padding sel).
        # Total HARUS <= lebar halaman - margin kiri/kanan (A4 21cm - 2cm - 2cm = 17cm).
        meta_table = Table(table_data, colWidths=[3.4 * cm, 5.4 * cm, 3.0 * cm, 5.0 * cm])
        meta_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), COLOR_LIGHT),
            ("BOX", (0, 0), (-1, -1), 0.7, COLOR_PRIMARY),
            ("INNERGRID", (0, 0), (-1, -1), 0.4, colors.white),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.append(meta_table)
        story.append(Spacer(1, 12))

    if gambar_bytes:
        try:
            img_buf = BytesIO(gambar_bytes)
            rl_img = RLImage(img_buf, width=8 * cm, height=6 * cm, kind="proportional")
            rl_img.hAlign = "CENTER"
            story.append(rl_img)
            story.append(Paragraph("Gambar Ilustrasi Stimulus Soal", styles["subtitle"]))
            story.append(Spacer(1, 8))
        except Exception:
            pass

    story.extend(blocks_to_pdf_flowables(parse_markdown_blocks(markdown_text), styles))
    doc.build(story, onFirstPage=_pdf_footer, onLaterPages=_pdf_footer)
    buffer.seek(0)
    return buffer.getvalue()


# ---------- Rendering ke DOCX (python-docx, bisa diedit) ----------
def _add_formatted_runs_docx(paragraph, text: str):
    tokens = re.split(r"(\*\*.+?\*\*|\*.+?\*)", text)
    for tok in tokens:
        if not tok:
            continue
        if tok.startswith("**") and tok.endswith("**"):
            run = paragraph.add_run(tok[2:-2]); run.bold = True
        elif tok.startswith("*") and tok.endswith("*"):
            run = paragraph.add_run(tok[1:-1]); run.italic = True
        else:
            paragraph.add_run(tok)


def generate_docx(doc_title: str, meta: dict, markdown_text: str, gambar_bytes: bytes = None) -> bytes:
    document = Document()

    section = document.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)

    p_title = document.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run(doc_title)
    r_title.bold = True
    r_title.font.size = Pt(16)
    r_title.font.color.rgb = DOCX_PRIMARY_RGB

    p_sub = document.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Kurikulum Merdeka • Pembelajaran Mendalam (Deep Learning)")
    r_sub.italic = True
    r_sub.font.size = Pt(10)
    r_sub.font.color.rgb = DOCX_GREY_RGB

    rows = _meta_pairs(meta)
    if rows:
        table = document.add_table(rows=0, cols=4)
        table.style = "Light Grid Accent 1"
        table.autofit = False
        table.allow_autofit = False
        # Label rata-kanan pada kolom lebar tetap -> titik dua otomatis lurus di setiap baris.
        # Total <= lebar halaman - margin kiri/kanan (21cm - 2cm - 2cm = 17cm)
        lebar_kolom = [Cm(3.4), Cm(5.4), Cm(3.0), Cm(5.0)]
        for pair in rows:
            cells = table.add_row().cells
            for c, w in zip(cells, lebar_kolom):
                c.width = w
            idx = 0
            for k, v in pair:
                cells[idx].text = ""
                run = cells[idx].paragraphs[0].add_run(k)
                run.bold = True
                cells[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                cells[idx + 1].text = f": {v}"
                idx += 2
            if len(pair) == 1:
                for extra_idx in range(2, 4):
                    cells[extra_idx].text = ""
        document.add_paragraph("")

    if gambar_bytes:
        try:
            document.add_picture(BytesIO(gambar_bytes), width=Inches(3.2))
            document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cap = document.add_paragraph("Gambar Ilustrasi Stimulus Soal")
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cap.runs[0].italic = True
            cap.runs[0].font.size = Pt(9)
        except Exception:
            pass

    for kind, content in parse_markdown_blocks(markdown_text):
        if kind == "h1":
            h = document.add_heading(level=1)
            _add_formatted_runs_docx(h, content)
        elif kind == "h2":
            h = document.add_heading(level=2)
            _add_formatted_runs_docx(h, content)
        elif kind == "h3":
            h = document.add_heading(level=3)
            _add_formatted_runs_docx(h, content)
        elif kind == "bullet":
            p = document.add_paragraph(style="List Bullet")
            _add_formatted_runs_docx(p, content)
        elif kind == "numbered":
            p = document.add_paragraph(style="List Number")
            _add_formatted_runs_docx(p, content)
        elif kind == "table":
            ncols = max(len(r) for r in content)
            tbl = document.add_table(rows=0, cols=ncols)
            tbl.style = "Light Grid Accent 1"
            for ridx, row in enumerate(content):
                cells = tbl.add_row().cells
                for cidx in range(ncols):
                    text = row[cidx] if cidx < len(row) else ""
                    cells[cidx].text = ""
                    _add_formatted_runs_docx(cells[cidx].paragraphs[0], text)
                    if ridx == 0:
                        for run in cells[cidx].paragraphs[0].runs:
                            run.bold = True
            document.add_paragraph("")
        else:
            p = document.add_paragraph()
            _add_formatted_runs_docx(p, content)

    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def download_row(base_filename: str, pdf_bytes: bytes, docx_bytes: bytes, key_prefix: str):
    c1, c2 = st.columns(2)
    with c1:
        st.download_button("⬇️ Download PDF (A4, cetak)", data=pdf_bytes,
                            file_name=f"{base_filename}.pdf", mime="application/pdf",
                            key=f"{key_prefix}_pdf")
    with c2:
        st.download_button(
            "⬇️ Download DOC (bisa diedit)", data=docx_bytes,
            file_name=f"{base_filename}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"{key_prefix}_docx",
        )


# ============================================================
# SIDEBAR - API KEY & IDENTITAS SEKOLAH (dipakai lintas tab)
# ============================================================
with st.sidebar:
    if AUTH_AKTIF and st.session_state.user:
        st.success(f"👤 Masuk sebagai: **{st.session_state.user['email']}**")
        if st.button("🚪 Keluar", key="btn_keluar", use_container_width=True):
            sb_keluar()
            st.rerun()
        st.divider()
    st.header("⚙️ Pengaturan")
    groq_api_key = st.text_input(
        "🔑 API Key Groq", type="password",
        help="Dapatkan API key gratis di console.groq.com. Key TIDAK disimpan di server, "
             "hanya dipakai selama sesi ini.",
    )
    st.caption(
        "⚠️ Jangan pernah menempelkan API key langsung di dalam kode. Jika key Anda pernah "
        "tertulis di source code / dibagikan publik, segera regenerate di console.groq.com."
    )
    st.divider()
    st.subheader("🏫 Identitas Sekolah")
    st.session_state.sekolah = st.text_input("Nama Sekolah", st.session_state.sekolah)
    st.session_state.tahun_ajaran = st.text_input("Tahun Ajaran", st.session_state.tahun_ajaran)
    st.divider()
    st.subheader("👩‍🏫 Identitas Guru Penyusun")
    st.session_state.penyusun = st.text_input("Nama Guru Penyusun", st.session_state.penyusun)
    st.session_state.guru_nip = st.text_input("NIP Guru Penyusun", st.session_state.guru_nip,
                                               placeholder="Contoh: 199001012019031001")
    st.divider()
    st.subheader("🧑‍💼 Identitas Kepala Sekolah")
    st.session_state.kepsek_nama = st.text_input("Nama Kepala Sekolah", st.session_state.kepsek_nama)
    st.session_state.kepsek_nip = st.text_input("NIP Kepala Sekolah", st.session_state.kepsek_nip,
                                                 placeholder="Contoh: 198005052008011003")
    st.divider()
    st.caption("📘 Seluruh materi dihasilkan berpedoman pada Permendikdasmen No. 13 Tahun 2025 "
               "tentang Pembelajaran Mendalam (Mindful, Meaningful, Joyful).")
    st.divider()
    st.caption("🔗 Ingin aplikasi ini online & bisa dibagikan/dijual ke rekan guru lain? "
               "Lihat panduan **README_DEPLOY.md** yang disertakan.")


st.title("📚 Generator Perangkat Ajar Kurikulum Merdeka")
st.caption(f"Aplikasi Pembantu Guru {st.session_state.sekolah} — Powered by Groq / Llama-3.3 — "
           f"Berpedoman Pembelajaran Mendalam (Deep Learning)")

tab_tp, tab_modul, tab_lkpd, tab_prota, tab_minggu, tab_jurnal, tab_absen = st.tabs(
    ["🎯 TP & ATP", "📖 Modul Ajar", "📝 LKPD", "📅 PROTA & PROMES",
     "🗓️ Minggu Efektif", "📔 Jurnal Mengajar", "🧑‍🤝‍🧑 Absensi Siswa"]
)

# ============================================================
# TAB 1 - TUJUAN PEMBELAJARAN (TP) & ALUR TUJUAN PEMBELAJARAN (ATP)
# ============================================================
with tab_tp:
    st.subheader("Rumuskan TP & Susun ATP")
    c1, c2 = st.columns(2)
    with c1:
        tp_mapel = st.selectbox("Mata Pelajaran", SUBJECT_OPTIONS, key="tp_mapel")
        tp_kelas = st.selectbox("Kelas (1-6)", KELAS_OPTIONS, key="tp_kelas")
        tp_semester = st.selectbox("Semester", ["Ganjil", "Genap"], key="tp_semester")
    with c2:
        tp_cp = st.text_area(
            "Capaian Pembelajaran (CP) / Elemen",
            placeholder="Tempel CP resmi dari Kemendikbudristek untuk elemen ini, "
                        "atau tuliskan ringkasan elemen materi.",
            height=120, key="tp_cp",
        )
        tp_jumlah_tp = st.number_input("Perkiraan Jumlah TP yang diinginkan", 3, 20, 6)

    tp_referensi = uploader_referensi("tp_referensi_upl")

    if st.button("🚀 Buat TP & ATP", key="btn_tp"):
        if not groq_api_key:
            st.error("⚠️ Mohon masukkan API Key Groq di sidebar terlebih dahulu.")
        elif not tp_cp:
            st.error("⚠️ Mohon isi Capaian Pembelajaran (CP) / Elemen.")
        else:
            try:
                with st.spinner("Menyusun Tujuan Pembelajaran dan Alur Tujuan Pembelajaran..."):
                    prompt = f"""
Anda adalah Pakar Pengembang Kurikulum Merdeka dan Pengawas Sekolah Senior di Indonesia.
Susun TUJUAN PEMBELAJARAN (TP) dan ALUR TUJUAN PEMBELAJARAN (ATP) yang siap pakai.

{DEEP_LEARNING_GUIDE}

DATA INPUT:
- Satuan Pendidikan: {st.session_state.sekolah}
- Mata Pelajaran: {tp_mapel}
- Kelas: {label_kelas_fase(tp_kelas)}
- Semester: {tp_semester}
- Tahun Ajaran: {st.session_state.tahun_ajaran}
- Capaian Pembelajaran (CP) / Elemen: {tp_cp}
- Target jumlah TP: sekitar {tp_jumlah_tp} tujuan
{f"- REFERENSI TAMBAHAN dari guru (jadikan acuan konteks/istilah bila relevan): {tp_referensi}" if tp_referensi else ""}

FORMAT OUTPUT (Markdown, WAJIB rapi, gunakan tabel markdown dengan simbol | untuk ATP):

# TUJUAN PEMBELAJARAN & ALUR TUJUAN PEMBELAJARAN

## A. IDENTITAS
Sebutkan satuan pendidikan, mapel, kelas/fase, semester, tahun ajaran.

## B. CAPAIAN PEMBELAJARAN (CP)
Tuliskan ulang CP secara ringkas dan jelas.

## C. TUJUAN PEMBELAJARAN (TP)
Buat daftar TP dengan kode (TP 1, TP 2, dst), masing-masing memuat KOMPETENSI + LINGKUP MATERI,
disusun terukur dan operasional (gunakan kata kerja operasional Taksonomi Bloom revisi).

## D. ALUR TUJUAN PEMBELAJARAN (ATP)
Sajikan dalam bentuk TABEL markdown dengan kolom:
| No | Kode TP | Tujuan Pembelajaran | Alokasi Waktu (JP) | Dimensi Profil Lulusan Terkait | Prinsip Deep Learning (Mindful/Meaningful/Joyful) |
Urutkan dari kompetensi prasyarat termudah menuju paling kompleks (logis & berjenjang).

## E. CATATAN PENGEMBANGAN
Berikan catatan singkat strategi asesmen dan diferensiasi yang disarankan sepanjang alur ini.

Gunakan Bahasa Indonesia resmi, komunikatif, dan terstruktur rapi.
"""
                    hasil = call_groq(groq_api_key, prompt)
                    st.session_state.hasil_tp_atp = hasil
                    st.session_state.meta_tp_atp = {
                        "Satuan Pendidikan": st.session_state.sekolah,
                        "Mata Pelajaran": tp_mapel,
                        "Kelas": label_kelas_fase(tp_kelas),
                        "Semester": tp_semester,
                        "Tahun Ajaran": st.session_state.tahun_ajaran,
                        "Guru Penyusun": st.session_state.penyusun,
                        "NIP Guru": st.session_state.guru_nip,
                        "Kepala Sekolah": st.session_state.kepsek_nama,
                        "NIP Kepala Sekolah": st.session_state.kepsek_nip,
                    }
                st.success("✨ TP & ATP berhasil disusun!")
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan API Groq: {e}")

    if st.session_state.hasil_tp_atp:
        st.markdown(st.session_state.hasil_tp_atp)
        pdf_bytes = generate_pdf("TUJUAN PEMBELAJARAN & ALUR TUJUAN PEMBELAJARAN",
                                  st.session_state.meta_tp_atp, st.session_state.hasil_tp_atp)
        docx_bytes = generate_docx("TUJUAN PEMBELAJARAN & ALUR TUJUAN PEMBELAJARAN",
                                    st.session_state.meta_tp_atp, st.session_state.hasil_tp_atp)
        download_row(f"TP_ATP_{tp_mapel.replace(' ', '_')}", pdf_bytes, docx_bytes, "dl_tp")

# ============================================================
# TAB 2 - MODUL AJAR
# ============================================================
with tab_modul:
    st.subheader("Susun Modul Ajar Lengkap")
    c1, c2 = st.columns(2)
    with c1:
        m_mapel = st.selectbox("Mata Pelajaran", SUBJECT_OPTIONS, key="m_mapel")
        m_kelas = st.selectbox("Kelas (1-6)", KELAS_OPTIONS, key="m_kelas")
        m_alokasi = st.text_input("Alokasi Waktu", "2 x 35 Menit (1 Pertemuan)", key="m_alokasi")
    with c2:
        m_topik = st.text_input("Topik / Materi Utama", key="m_topik")
        m_tp_acuan = st.text_area(
            "Tujuan Pembelajaran acuan (opsional, bisa salin dari tab TP & ATP)",
            height=90, key="m_tp_acuan",
        )

    m_referensi = uploader_referensi("m_referensi_upl")

    if st.button("🚀 Buat Modul Ajar Sekarang", key="btn_modul"):
        if not groq_api_key:
            st.error("⚠️ Mohon masukkan API Key Groq di sidebar terlebih dahulu.")
        elif not m_topik:
            st.error("⚠️ Mohon isi Topik / Materi Utama!")
        else:
            try:
                with st.spinner("Groq AI sedang menyusun Modul Ajar presisi... Mohon tunggu..."):
                    prompt = f"""
Anda adalah Pakar Pengembang Kurikulum Merdeka dan Pengawas Sekolah Senior di Indonesia.
Tugas Anda adalah membuatkan MODUL AJAR LENGKAP & SIAP PAKAI sesuai format resmi Kemendikbudristek.

{DEEP_LEARNING_GUIDE}

DATA INPUT:
- Satuan Pendidikan: {st.session_state.sekolah}
- Guru Penyusun: {st.session_state.penyusun or '(diisi oleh guru)'}
- Mata Pelajaran: {m_mapel}
- Kelas: {label_kelas_fase(m_kelas)}
- Topik Utama: {m_topik}
- Alokasi Waktu: {m_alokasi}
- Tujuan Pembelajaran acuan (jika ada): {m_tp_acuan or 'Rumuskan sendiri berdasarkan topik.'}
{f"- REFERENSI TAMBAHAN dari Buku Guru/materi yang diunggah (jadikan acuan isi & istilah bila relevan): {m_referensi}" if m_referensi else ""}

PETUNJUK FORMAT OUTPUT (Wajib Rapi dengan Markdown):

# MODUL AJAR KURIKULUM MERDEKA

## A. INFORMASI UMUM
1. **Identitas Modul**: Nama Penyusun, Sekolah, Jenjang SD, Mata Pelajaran, Alokasi Waktu.
2. **Kompetensi Awal**: Pengetahuan dasar yang harus dimiliki murid sebelum materi ini.
3. **Profil Lulusan / Profil Pelajar Pancasila**: Sebutkan dimensi yang relevan.
4. **Sarana & Prasarana**: Alat dan bahan yang dibutuhkan.
5. **Target Peserta Didik**: Reguler / Tipikal.
6. **Model Pembelajaran**: Tatap Muka / Problem Based Learning / Diferensiasi (sesuaikan Deep Learning).

## B. KOMPONEN INTI
1. **Tujuan Pembelajaran (TP)**: Jabarkan secara rinci dan terukur.
2. **Pemahaman Bermakna**: Manfaat materi dalam kehidupan sehari-hari murid.
3. **Pertanyaan Pemantik**: 2-3 pertanyaan pemicu diskusi di awal kelas.
4. **Kegiatan Pembelajaran**:
   - **Pendahuluan (10 Menit)**: Orientasi, Apersepsi, Motivasi. Beri label prinsip Deep Learning.
   - **Kegiatan Inti (50 Menit)**: Langkah konkret pembelajaran diferensiasi/aktif, SETIAP langkah
     diberi label (Mindful)/(Meaningful)/(Joyful) sesuai pedoman di atas.
   - **Penutup (10 Menit)**: Evaluasi, Refleksi (Mindful), Rencana Pertemuan Berikutnya.
5. **Asesmen**: Asesmen Formatif lengkap dengan Rubrik Penilaian dalam TABEL markdown (Skor 1-4).

## C. LAMPIRAN
1. **Lembar Kerja Peserta Didik (LKPD) ringkas**: Buatkan soal/tugas interaktif siap cetak.
2. **Bahan Bacaan Guru & Murid**: Ringkasan materi singkat.

Gunakan Bahasa Indonesia resmi, komunikatif, dan terstruktur rapi.
"""
                    hasil = call_groq(groq_api_key, prompt)
                    st.session_state.hasil_modul = hasil
                    st.session_state.meta_modul = {
                        "Satuan Pendidikan": st.session_state.sekolah,
                        "Mata Pelajaran": m_mapel,
                        "Kelas": label_kelas_fase(m_kelas),
                        "Topik": m_topik,
                        "Alokasi Waktu": m_alokasi,
                        "Guru Penyusun": st.session_state.penyusun,
                        "NIP Guru": st.session_state.guru_nip,
                        "Kepala Sekolah": st.session_state.kepsek_nama,
                        "NIP Kepala Sekolah": st.session_state.kepsek_nip,
                    }
                st.success("✨ Modul Ajar berhasil disusun!")
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan API Groq: {e}")

    if st.session_state.hasil_modul:
        st.markdown(st.session_state.hasil_modul)
        pdf_bytes = generate_pdf("MODUL AJAR KURIKULUM MERDEKA",
                                  st.session_state.meta_modul, st.session_state.hasil_modul)
        docx_bytes = generate_docx("MODUL AJAR KURIKULUM MERDEKA",
                                    st.session_state.meta_modul, st.session_state.hasil_modul)
        download_row(f"Modul_Ajar_{m_mapel.replace(' ', '_') if m_mapel else 'Modul'}",
                     pdf_bytes, docx_bytes, "dl_modul")

# ============================================================
# TAB 3 - LEMBAR KERJA PESERTA DIDIK (LKPD)
# ============================================================
with tab_lkpd:
    st.subheader("Susun LKPD Siap Cetak")
    c1, c2 = st.columns(2)
    with c1:
        l_mapel = st.selectbox("Mata Pelajaran", SUBJECT_OPTIONS, key="l_mapel")
        l_kelas = st.selectbox("Kelas (1-6)", KELAS_OPTIONS, key="l_kelas")
        l_jumlah_soal = st.slider("Jumlah Soal", 3, 15, 8, key="l_jumlah_soal")
    with c2:
        l_topik = st.text_input("Topik / Materi", key="l_topik")
        l_tingkat = st.select_slider(
            "Tingkat Kesulitan",
            options=["Mudah (C1-C2)", "Sedang (C3-C4)", "Menantang / HOTS (C5-C6)"],
            value="Sedang (C3-C4)", key="l_tingkat",
        )
        l_jenis_soal = st.multiselect(
            "Jenis Soal", ["Pilihan Ganda", "Isian Singkat", "Uraian / Essay", "Praktik/Unjuk Kerja"],
            default=["Pilihan Ganda", "Uraian / Essay"], key="l_jenis_soal",
        )

    st.markdown("**🎨 Stimulus Soal** — bantu murid memahami konteks soal agar tidak bingung")
    c3, c4 = st.columns(2)
    with c3:
        l_stimulus = st.selectbox(
            "Jenis Stimulus pada beberapa soal",
            ["Tanpa stimulus", "Cerita singkat / narasi kontekstual",
             "Ilustrasi gambar (deskripsi/petunjuk gambar)", "Cerita + Ilustrasi gambar"],
            key="l_stimulus",
        )
    with c4:
        l_gambar_upl = st.file_uploader(
            "🖼️ Unggah Gambar Ilustrasi (opsional, akan disisipkan di LKPD)",
            type=["png", "jpg", "jpeg"], key="l_gambar_upl",
        )
        st.session_state.lkpd_gambar_bytes = l_gambar_upl.getvalue() if l_gambar_upl is not None else None

    l_referensi = uploader_referensi("l_referensi_upl", "📎 Unggah Referensi Materi/Buku Guru (opsional)")

    if st.button("🚀 Buat LKPD Sekarang", key="btn_lkpd"):
        if not groq_api_key:
            st.error("⚠️ Mohon masukkan API Key Groq di sidebar terlebih dahulu.")
        elif not l_topik:
            st.error("⚠️ Mohon isi Topik / Materi.")
        else:
            try:
                with st.spinner("Menyusun LKPD siap cetak..."):
                    prompt = f"""
Anda adalah Pakar Pengembang Kurikulum Merdeka di Indonesia.
Susun LEMBAR KERJA PESERTA DIDIK (LKPD) yang siap dicetak untuk murid Sekolah Dasar.

{DEEP_LEARNING_GUIDE}

DATA INPUT:
- Satuan Pendidikan: {st.session_state.sekolah}
- Mata Pelajaran: {l_mapel}
- Kelas: {label_kelas_fase(l_kelas)}
- Topik / Materi: {l_topik}
- Jumlah Soal: {l_jumlah_soal}
- Tingkat Kesulitan: {l_tingkat}
- Jenis Soal yang diminta: {', '.join(l_jenis_soal) if l_jenis_soal else 'campuran sesuai topik'}
- Kebutuhan Stimulus Soal: {l_stimulus}
{f"- REFERENSI TAMBAHAN dari guru (jadikan acuan isi bila relevan): {l_referensi}" if l_referensi else ""}

INSTRUKSI STIMULUS (WAJIB diikuti sesuai kebutuhan stimulus di atas):
- Jika stimulus berupa "Cerita singkat / narasi kontekstual": sisipkan 1 cerita pendek (3-5 kalimat,
  dekat dengan kehidupan anak SD) sebelum sekelompok soal yang berkaitan, agar anak memahami konteks
  sebelum menjawab.
- Jika stimulus berupa "Ilustrasi gambar (deskripsi/petunjuk gambar)": untuk soal yang relevan, tuliskan
  baris "*(Petunjuk ilustrasi untuk guru: [deskripsi singkat gambar yang perlu ditempel/digambar di sini])*"
  tepat sebelum soal tersebut, agar guru tahu gambar apa yang perlu dilampirkan secara manual.
- Jika "Cerita + Ilustrasi gambar": gabungkan keduanya secara proporsional.
- Jika "Tanpa stimulus": langsung ke soal seperti biasa.
Stimulus TIDAK perlu diberikan di setiap nomor, cukup pada beberapa soal yang paling membutuhkan
konteks agar anak tidak bingung.

FORMAT OUTPUT (Markdown, WAJIB rapi):

# LEMBAR KERJA PESERTA DIDIK (LKPD)

## A. IDENTITAS
Mapel, Kelas/Fase, Topik. Sertakan kolom isian: Nama: __________  Kelas: ______  Tanggal: ______

## B. PETUNJUK PENGERJAAN
Instruksi singkat, jelas, ramah anak, sisipkan kalimat pemantik semangat (Joyful).

## C. TUJUAN
1-2 kalimat tujuan belajar dari LKPD ini (Meaningful - kaitkan dengan kehidupan sehari-hari).

## D. KEGIATAN PEMANASAN (Mindful)
1 kegiatan singkat sebelum mengerjakan soal, misal refleksi/pertanyaan pemantik.

## E. SOAL-SOAL
Nomori setiap soal, sesuai jenis soal yang diminta, tingkat kesulitan {l_tingkat}, sediakan
ruang jawaban (garis titik-titik) untuk isian/uraian. Untuk pilihan ganda beri opsi A-D.

## F. KUNCI JAWABAN & PEDOMAN PENSKORAN
Pisahkan bagian ini dengan jelas (halaman guru), sertakan tabel skor per nomor.

Gunakan Bahasa Indonesia yang komunikatif dan ramah anak sesuai jenjang {label_kelas_fase(l_kelas)}.
"""
                    hasil = call_groq(groq_api_key, prompt)
                    st.session_state.hasil_lkpd = hasil
                    st.session_state.meta_lkpd = {
                        "Satuan Pendidikan": st.session_state.sekolah,
                        "Mata Pelajaran": l_mapel,
                        "Kelas": label_kelas_fase(l_kelas),
                        "Topik": l_topik,
                        "Tingkat": l_tingkat,
                        "Guru Penyusun": st.session_state.penyusun,
                        "NIP Guru": st.session_state.guru_nip,
                    }
                st.success("✨ LKPD berhasil disusun!")
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan API Groq: {e}")

    if st.session_state.hasil_lkpd:
        st.markdown(st.session_state.hasil_lkpd)
        if st.session_state.lkpd_gambar_bytes:
            st.image(st.session_state.lkpd_gambar_bytes, caption="Gambar Ilustrasi Stimulus Soal", width=300)
        pdf_bytes = generate_pdf("LEMBAR KERJA PESERTA DIDIK (LKPD)",
                                  st.session_state.meta_lkpd, st.session_state.hasil_lkpd,
                                  gambar_bytes=st.session_state.lkpd_gambar_bytes)
        docx_bytes = generate_docx("LEMBAR KERJA PESERTA DIDIK (LKPD)",
                                    st.session_state.meta_lkpd, st.session_state.hasil_lkpd,
                                    gambar_bytes=st.session_state.lkpd_gambar_bytes)
        download_row(f"LKPD_{l_mapel.replace(' ', '_') if l_mapel else 'LKPD'}",
                     pdf_bytes, docx_bytes, "dl_lkpd")

# ============================================================
# TAB 4 - PROGRAM TAHUNAN (PROTA) & PROGRAM SEMESTER (PROMES)
# ============================================================
with tab_prota:
    st.subheader("Susun Program Tahunan (PROTA) & Program Semester (PROMES)")
    c1, c2 = st.columns(2)
    with c1:
        pp_mapel = st.selectbox("Mata Pelajaran", SUBJECT_OPTIONS, key="pp_mapel")
        pp_kelas = st.selectbox("Kelas (1-6)", KELAS_OPTIONS, key="pp_kelas")
    with c2:
        pp_minggu_ganjil = st.number_input("Minggu Efektif Semester Ganjil", 8, 24, 18, key="pp_mg")
        pp_minggu_genap = st.number_input("Minggu Efektif Semester Genap", 8, 24, 18, key="pp_mgg")

    pp_materi = st.text_area(
        "Daftar Materi Pokok / TP satu tahun (bisa salin dari tab TP & ATP)",
        height=150, key="pp_materi",
        placeholder="Contoh:\nSemester Ganjil: Tata Cara Wudhu, Sholat Fardhu Berjamaah, Adab kepada Orang Tua\n"
                    "Semester Genap: Puasa Ramadhan, Zakat Fitrah, Kisah Nabi",
    )
    pp_referensi = uploader_referensi("pp_referensi_upl",
                                       "📎 Unggah Referensi Tambahan (Buku Guru/Materi/ATP, opsional)")

    colA, colB = st.columns(2)
    with colA:
        gen_prota = st.button("🚀 Buat PROTA", key="btn_prota", use_container_width=True)
    with colB:
        pp_semester_promes = st.selectbox("Semester untuk PROMES", ["Ganjil", "Genap"], key="pp_semester_promes")
        gen_promes = st.button("🚀 Buat PROMES", key="btn_promes", use_container_width=True)

    if gen_prota or gen_promes:
        if not groq_api_key:
            st.error("⚠️ Mohon masukkan API Key Groq di sidebar terlebih dahulu.")
        elif not pp_materi:
            st.error("⚠️ Mohon isi Daftar Materi Pokok / TP satu tahun.")
        else:
            try:
                if gen_prota:
                    judul = "PROGRAM TAHUNAN (PROTA)"
                    with st.spinner("Menyusun Program Tahunan (PROTA)..."):
                        prompt = f"""
Anda adalah Pakar Pengembang Kurikulum Merdeka di Indonesia.
Susun PROGRAM TAHUNAN (PROTA) yang siap pakai untuk satu tahun ajaran penuh (2 semester).

DATA INPUT:
- Satuan Pendidikan: {st.session_state.sekolah}
- Mata Pelajaran: {pp_mapel}
- Kelas: {label_kelas_fase(pp_kelas)}
- Tahun Ajaran: {st.session_state.tahun_ajaran}
- Minggu Efektif Semester Ganjil: {pp_minggu_ganjil} minggu
- Minggu Efektif Semester Genap: {pp_minggu_genap} minggu
- Daftar Materi Pokok / TP setahun: {pp_materi}
{f"- REFERENSI TAMBAHAN dari guru (jadikan acuan bila relevan): {pp_referensi}" if pp_referensi else ""}

FORMAT OUTPUT (Markdown, WAJIB rapi, gunakan tabel markdown dengan simbol |):

# PROGRAM TAHUNAN (PROTA)

## A. IDENTITAS
Satuan pendidikan, mapel, kelas, tahun ajaran.

## B. TABEL PROGRAM TAHUNAN
Sajikan dalam TABEL markdown dengan kolom:
| No | Semester | Materi Pokok / Tujuan Pembelajaran | Alokasi Waktu (JP) | Keterangan |
Kelompokkan baris per semester (Ganjil dahulu, lalu Genap), alokasi waktu total harus
proporsional dengan minggu efektif yang diberikan (asumsikan sekitar 2-4 JP per minggu per materi,
sesuaikan agar realistis).

## C. REKAPITULASI
Tabel ringkas: total JP semester ganjil, total JP semester genap, total JP satu tahun.

Gunakan Bahasa Indonesia resmi dan terstruktur rapi.
"""
                else:
                    judul = f"PROGRAM SEMESTER (PROMES) - SEMESTER {pp_semester_promes.upper()}"
                    minggu_pakai = pp_minggu_ganjil if pp_semester_promes == "Ganjil" else pp_minggu_genap
                    bulan_range = "Juli - Desember" if pp_semester_promes == "Ganjil" else "Januari - Juni"
                    with st.spinner(f"Menyusun Program Semester ({pp_semester_promes})..."):
                        prompt = f"""
Anda adalah Pakar Pengembang Kurikulum Merdeka di Indonesia.
Susun PROGRAM SEMESTER (PROMES) yang siap pakai untuk semester {pp_semester_promes}.

DATA INPUT:
- Satuan Pendidikan: {st.session_state.sekolah}
- Mata Pelajaran: {pp_mapel}
- Kelas: {label_kelas_fase(pp_kelas)}
- Tahun Ajaran: {st.session_state.tahun_ajaran}
- Semester: {pp_semester_promes} (rentang bulan {bulan_range})
- Minggu Efektif: {minggu_pakai} minggu
- Daftar Materi Pokok / TP (ambil bagian yang relevan untuk semester {pp_semester_promes}): {pp_materi}
{f"- REFERENSI TAMBAHAN dari guru (jadikan acuan bila relevan): {pp_referensi}" if pp_referensi else ""}

FORMAT OUTPUT (Markdown, WAJIB rapi, gunakan tabel markdown dengan simbol |):

# PROGRAM SEMESTER (PROMES) - SEMESTER {pp_semester_promes.upper()}

## A. IDENTITAS
Satuan pendidikan, mapel, kelas, tahun ajaran, semester, rentang bulan {bulan_range}.

## B. TABEL PROGRAM SEMESTER
Sajikan dalam TABEL markdown dengan kolom:
| No | Materi Pokok / TP | Alokasi Waktu (JP) | Bulan & Minggu Pelaksanaan | Keterangan |
Sebar materi secara realistis ke {minggu_pakai} minggu efektif dalam rentang {bulan_range},
sisakan alokasi untuk Penilaian Tengah Semester (PTS) dan Penilaian Akhir Semester (PAS)
sebagai baris tersendiri di tabel.

## C. CATATAN
Catatan singkat tentang hari libur/jeda semester yang perlu diperhitungkan guru.

Gunakan Bahasa Indonesia resmi dan terstruktur rapi.
"""
                hasil = call_groq(groq_api_key, prompt)
                st.session_state.hasil_prota_promes = hasil
                st.session_state.judul_prota_promes = judul
                st.session_state.meta_prota_promes = {
                    "Satuan Pendidikan": st.session_state.sekolah,
                    "Mata Pelajaran": pp_mapel,
                    "Kelas": label_kelas_fase(pp_kelas),
                    "Tahun Ajaran": st.session_state.tahun_ajaran,
                    "Guru Penyusun": st.session_state.penyusun,
                    "NIP Guru": st.session_state.guru_nip,
                    "Kepala Sekolah": st.session_state.kepsek_nama,
                    "NIP Kepala Sekolah": st.session_state.kepsek_nip,
                }
                st.success(f"✨ {judul} berhasil disusun!")
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan API Groq: {e}")

    if st.session_state.hasil_prota_promes:
        st.markdown(st.session_state.hasil_prota_promes)
        pdf_bytes = generate_pdf(st.session_state.judul_prota_promes,
                                  st.session_state.meta_prota_promes, st.session_state.hasil_prota_promes)
        docx_bytes = generate_docx(st.session_state.judul_prota_promes,
                                    st.session_state.meta_prota_promes, st.session_state.hasil_prota_promes)
        nama_file = st.session_state.judul_prota_promes.split(" - ")[0].replace(" ", "_")
        download_row(f"{nama_file}_{pp_mapel.replace(' ', '_') if pp_mapel else 'Program'}",
                     pdf_bytes, docx_bytes, "dl_prota")

# ============================================================
# TAB 5 - ANALISIS MINGGU EFEKTIF (berbasis Kalender Pendidikan/KALDIK)
# ============================================================
with tab_minggu:
    st.subheader("Analisis Minggu Efektif Berdasarkan Kalender Pendidikan (KALDIK)")
    st.caption("Hitung otomatis jumlah minggu efektif dalam satu semester berdasarkan rentang semester "
               "dan daftar hari/minggu tidak efektif (libur, PTS/PAS, jeda semester, dll) dari KALDIK sekolah.")

    c1, c2 = st.columns(2)
    with c1:
        me_semester = st.selectbox("Semester", ["Ganjil", "Genap"], key="me_semester")
        me_tanggal_mulai = st.date_input("Tanggal Mulai Semester", value=date(2025, 7, 14), key="me_mulai")
    with c2:
        me_mapel = st.selectbox("Mata Pelajaran (opsional, untuk judul dokumen)", SUBJECT_OPTIONS, key="me_mapel")
        me_tanggal_selesai = st.date_input("Tanggal Selesai Semester", value=date(2025, 12, 20), key="me_selesai")

    st.markdown("**📋 Daftar Minggu/Periode Tidak Efektif (dari KALDIK)** — satu baris per periode, format:\n"
                "`tanggal_awal - tanggal_akhir | keterangan` (format tanggal: YYYY-MM-DD)")

    me_file_kaldik = st.file_uploader(
        "📎 Atau unggah berkas Kaldik (CSV/XLSX dengan kolom tanggal_awal, tanggal_akhir, keterangan — "
        "atau PDF/DOCX/TXT berisi teks kaldik untuk dibaca sebagai referensi)",
        type=["csv", "xlsx", "xls", "pdf", "docx", "txt"], key="me_file_kaldik",
    )
    if me_file_kaldik is not None:
        nama_kaldik = me_file_kaldik.name.lower()
        if nama_kaldik.endswith((".csv", ".xlsx", ".xls")):
            baris_kaldik, status_kaldik = baca_tabel_upload(me_file_kaldik)
            if baris_kaldik:
                st.success(status_kaldik)
                with st.expander("👀 Pratinjau data kaldik yang terbaca"):
                    st.dataframe(baris_kaldik, use_container_width=True, hide_index=True)
                if st.button("➕ Tambahkan Baris di Atas ke Kolom Kaldik", key="btn_tambah_kaldik_upload"):
                    baris_baru = []
                    for r in baris_kaldik:
                        keys_lower = {str(k).strip().lower(): k for k in r.keys()}
                        k_awal = next((keys_lower[k] for k in keys_lower if "awal" in k or "mulai" in k), None)
                        k_akhir = next((keys_lower[k] for k in keys_lower if "akhir" in k or "selesai" in k), None)
                        k_ket = next((keys_lower[k] for k in keys_lower if "ket" in k), None)
                        if k_awal and k_akhir:
                            try:
                                ta = pd.to_datetime(r[k_awal]).strftime("%Y-%m-%d")
                                tb = pd.to_datetime(r[k_akhir]).strftime("%Y-%m-%d")
                                ket = str(r.get(k_ket, "Tidak Efektif")) if k_ket else "Tidak Efektif"
                                baris_baru.append(f"{ta} - {tb} | {ket}")
                            except Exception:
                                continue
                    if baris_baru:
                        gabungan = (st.session_state.me_kaldik_text + "\n" if st.session_state.me_kaldik_text else "")
                        st.session_state.me_kaldik_text = gabungan + "\n".join(baris_baru)
                        st.success(f"✅ {len(baris_baru)} periode ditambahkan ke kolom kaldik di bawah.")
                        st.rerun()
                    else:
                        st.warning("⚠️ Tidak menemukan kolom tanggal_awal/tanggal_akhir yang bisa dikenali di berkas ini.")
            else:
                st.warning(status_kaldik)
        else:
            teks_kaldik, status_kaldik = ekstrak_teks_referensi(me_file_kaldik)
            if teks_kaldik:
                st.success(status_kaldik)
                with st.expander(f"👀 Pratinjau isi '{me_file_kaldik.name}' — salin manual periode liburnya ke kolom di bawah"):
                    st.text(teks_kaldik[:3000])
            else:
                st.warning(status_kaldik)

    me_kaldik_text = st.text_area(
        "Input Kaldik (libur, PTS, PAS, jeda semester, dll)",
        height=150, key="me_kaldik_text",
        placeholder="2025-07-14 - 2025-07-19 | Libur Akhir Semester Genap Sebelumnya\n"
                    "2025-08-17 - 2025-08-17 | Libur HUT RI\n"
                    "2025-10-06 - 2025-10-11 | Penilaian Tengah Semester (PTS)\n"
                    "2025-12-08 - 2025-12-13 | Penilaian Akhir Semester (PAS)\n"
                    "2025-12-15 - 2025-12-20 | Libur Akhir Semester",
    )

    hitung_me = st.button("🧮 Hitung Minggu Efektif", key="btn_hitung_me", use_container_width=True)

    if hitung_me:
        if me_tanggal_selesai <= me_tanggal_mulai:
            st.error("⚠️ Tanggal selesai harus setelah tanggal mulai semester.")
        else:
            # Parsing baris kaldik
            periode_libur = []
            for baris in me_kaldik_text.strip().split("\n"):
                baris = baris.strip()
                if not baris:
                    continue
                try:
                    bagian_tanggal, keterangan = baris.split("|", 1)
                except ValueError:
                    bagian_tanggal, keterangan = baris, "Tidak Efektif"
                m = re.match(r"\s*(\d{4}-\d{2}-\d{2})\s*-\s*(\d{4}-\d{2}-\d{2})\s*", bagian_tanggal)
                if not m:
                    continue
                try:
                    tgl_awal = datetime.strptime(m.group(1), "%Y-%m-%d").date()
                    tgl_akhir = datetime.strptime(m.group(2), "%Y-%m-%d").date()
                    periode_libur.append((tgl_awal, tgl_akhir, keterangan.strip()))
                except Exception:
                    continue

            # Bangun daftar minggu (Senin-Minggu) yang beririsan rentang semester
            hasil_minggu = []
            senin = me_tanggal_mulai - timedelta(days=me_tanggal_mulai.weekday())
            no_minggu = 1
            while senin <= me_tanggal_selesai:
                minggu_akhir = senin + timedelta(days=6)
                keterangan_tidak_efektif = []
                for (la, lb, ket) in periode_libur:
                    if la <= minggu_akhir and lb >= senin:
                        keterangan_tidak_efektif.append(ket)
                status = "Tidak Efektif" if keterangan_tidak_efektif else "Efektif"
                hasil_minggu.append({
                    "No": no_minggu,
                    "Minggu Ke": f"{senin.strftime('%d %b')} - {minggu_akhir.strftime('%d %b %Y')}",
                    "Status": status,
                    "Keterangan": "; ".join(sorted(set(keterangan_tidak_efektif))) if keterangan_tidak_efektif else "-",
                })
                no_minggu += 1
                senin += timedelta(days=7)

            total_minggu = len(hasil_minggu)
            minggu_efektif = sum(1 for m_ in hasil_minggu if m_["Status"] == "Efektif")
            minggu_tidak_efektif = total_minggu - minggu_efektif

            st.session_state.tabel_minggu_efektif = hasil_minggu
            st.session_state.meta_minggu_efektif = {
                "Satuan Pendidikan": st.session_state.sekolah,
                "Semester": me_semester,
                "Tahun Ajaran": st.session_state.tahun_ajaran,
                "Rentang Semester": f"{me_tanggal_mulai.strftime('%d %b %Y')} - {me_tanggal_selesai.strftime('%d %b %Y')}",
                "Guru Penyusun": st.session_state.penyusun,
                "NIP Guru": st.session_state.guru_nip,
            }

            md = f"# ANALISIS MINGGU EFEKTIF SEMESTER {me_semester.upper()}\n\n"
            md += "## A. REKAPITULASI\n\n"
            md += "| Keterangan | Jumlah |\n|---|---|\n"
            md += f"| Total Minggu dalam Semester | {total_minggu} minggu |\n"
            md += f"| Minggu Efektif | {minggu_efektif} minggu |\n"
            md += f"| Minggu Tidak Efektif | {minggu_tidak_efektif} minggu |\n\n"
            md += "## B. RINCIAN PER MINGGU\n\n"
            md += "| No | Minggu Ke | Status | Keterangan |\n|---|---|---|---|\n"
            for m_ in hasil_minggu:
                md += f"| {m_['No']} | {m_['Minggu Ke']} | {m_['Status']} | {m_['Keterangan']} |\n"
            st.session_state.hasil_minggu_efektif = md
            st.success(f"✨ Minggu efektif berhasil dihitung: **{minggu_efektif} dari {total_minggu} minggu** efektif.")

    if st.session_state.hasil_minggu_efektif:
        colm1, colm2, colm3 = st.columns(3)
        tabel = st.session_state.tabel_minggu_efektif or []
        total_ = len(tabel)
        efektif_ = sum(1 for m_ in tabel if m_["Status"] == "Efektif")
        colm1.metric("Total Minggu", total_)
        colm2.metric("Minggu Efektif", efektif_)
        colm3.metric("Minggu Tidak Efektif", total_ - efektif_)
        st.markdown(st.session_state.hasil_minggu_efektif)
        pdf_bytes = generate_pdf(f"ANALISIS MINGGU EFEKTIF - SEMESTER {me_semester.upper()}",
                                  st.session_state.meta_minggu_efektif, st.session_state.hasil_minggu_efektif)
        docx_bytes = generate_docx(f"ANALISIS MINGGU EFEKTIF - SEMESTER {me_semester.upper()}",
                                    st.session_state.meta_minggu_efektif, st.session_state.hasil_minggu_efektif)
        download_row("Analisis_Minggu_Efektif", pdf_bytes, docx_bytes, "dl_me")

# ============================================================
# TAB 6 - JURNAL MENGAJAR HARIAN
# ============================================================
with tab_jurnal:
    st.subheader("Jurnal Mengajar Harian")
    st.caption("Catat kegiatan mengajar setiap hari, lalu unduh rekap jurnal dalam PDF/DOCX kapan saja.")

    with st.form("form_jurnal", clear_on_submit=True):
        jc1, jc2, jc3 = st.columns(3)
        with jc1:
            j_tanggal = st.date_input("Tanggal", value=date.today(), key="j_tanggal")
            j_kelas = st.selectbox("Kelas", KELAS_OPTIONS, key="j_kelas")
        with jc2:
            j_mapel = st.selectbox("Mata Pelajaran", SUBJECT_OPTIONS, key="j_mapel")
            j_jampel = st.text_input("Jam Pelajaran (JP)", "1-2", key="j_jampel")
        with jc3:
            j_hadir = st.text_input("Jumlah Siswa Hadir", "", key="j_hadir")
        j_materi = st.text_input("Materi / Kompetensi yang Diajarkan", key="j_materi")
        j_kegiatan = st.text_area("Ringkasan Kegiatan Pembelajaran", height=80, key="j_kegiatan")
        j_catatan = st.text_area("Catatan / Kendala / Tindak Lanjut (opsional)", height=60, key="j_catatan")
        submit_jurnal = st.form_submit_button("➕ Tambahkan ke Jurnal")

    if submit_jurnal:
        if not j_materi:
            st.error("⚠️ Mohon isi Materi / Kompetensi yang diajarkan.")
        else:
            tambah_baris_jurnal({
                "Tanggal": j_tanggal.strftime("%d-%m-%Y"),
                "Kelas": j_kelas,
                "Mapel": j_mapel,
                "JP": j_jampel,
                "Materi": j_materi,
                "Kegiatan": j_kegiatan,
                "Hadir": j_hadir,
                "Catatan": j_catatan,
            })
            st.success("✅ Catatan jurnal ditambahkan.")

    with st.expander("📎 Atau impor banyak catatan jurnal sekaligus dari berkas CSV/XLSX"):
        st.caption("Kolom yang dikenali: Tanggal, Kelas, Mapel, JP, Materi, Kegiatan, Hadir, Catatan "
                   "(kolom yang tidak ada akan dikosongkan).")
        j_file_upl = st.file_uploader("Unggah berkas jurnal", type=["csv", "xlsx", "xls"], key="j_file_upl")
        if j_file_upl is not None:
            baris_jurnal, status_jurnal = baca_tabel_upload(j_file_upl)
            if baris_jurnal:
                st.success(status_jurnal)
                st.dataframe(baris_jurnal, use_container_width=True, hide_index=True)
                if st.button("➕ Tambahkan Semua ke Jurnal", key="btn_impor_jurnal"):
                    for r in baris_jurnal:
                        keys_lower = {str(k).strip().lower(): k for k in r.keys()}
                        def ambil(*alias):
                            for a in alias:
                                if a in keys_lower:
                                    return str(r[keys_lower[a]])
                            return ""
                        tambah_baris_jurnal({
                            "Tanggal": ambil("tanggal") or date.today().strftime("%d-%m-%Y"),
                            "Kelas": ambil("kelas"),
                            "Mapel": ambil("mapel", "mata pelajaran"),
                            "JP": ambil("jp", "jam pelajaran"),
                            "Materi": ambil("materi"),
                            "Kegiatan": ambil("kegiatan"),
                            "Hadir": ambil("hadir"),
                            "Catatan": ambil("catatan"),
                        })
                    st.success(f"✅ {len(baris_jurnal)} catatan jurnal berhasil diimpor.")
                    st.rerun()
            else:
                st.warning(status_jurnal)

    if st.session_state.jurnal_rows:
        st.markdown(f"**Total catatan jurnal: {len(st.session_state.jurnal_rows)}**")
        st.dataframe([{k: v for k, v in r.items() if k != "id"} for r in st.session_state.jurnal_rows],
                     use_container_width=True, hide_index=True)

        cja, cjb = st.columns(2)
        with cja:
            if st.button("🗑️ Hapus Catatan Terakhir", key="btn_hapus_jurnal"):
                hapus_jurnal_terakhir()
                st.rerun()
        with cjb:
            if st.button("🧹 Kosongkan Semua Jurnal", key="btn_kosong_jurnal"):
                kosongkan_jurnal()
                st.rerun()

        md_jurnal = "# JURNAL MENGAJAR HARIAN\n\n"
        md_jurnal += "| No | Tanggal | Kelas | Mapel | JP | Materi | Hadir | Catatan |\n"
        md_jurnal += "|---|---|---|---|---|---|---|---|\n"
        for i, r in enumerate(st.session_state.jurnal_rows, 1):
            md_jurnal += (f"| {i} | {r['Tanggal']} | {r['Kelas']} | {r['Mapel']} | {r['JP']} | "
                          f"{r['Materi']} | {r['Hadir']} | {r['Catatan'] or '-'} |\n")

        meta_jurnal = {
            "Satuan Pendidikan": st.session_state.sekolah,
            "Tahun Ajaran": st.session_state.tahun_ajaran,
            "Guru Penyusun": st.session_state.penyusun,
            "NIP Guru": st.session_state.guru_nip,
        }
        pdf_bytes = generate_pdf("JURNAL MENGAJAR HARIAN", meta_jurnal, md_jurnal)
        docx_bytes = generate_docx("JURNAL MENGAJAR HARIAN", meta_jurnal, md_jurnal)
        download_row("Jurnal_Mengajar", pdf_bytes, docx_bytes, "dl_jurnal")
    else:
        st.info("Belum ada catatan jurnal. Isi formulir di atas untuk menambahkan.")

# ============================================================
# TAB 7 - DATA SISWA & LEMBAR ABSENSI
# ============================================================
with tab_absen:
    st.subheader("Data Peserta Didik & Lembar Absensi")
    st.caption("Input data siswa satu kali, lalu buat lembar absen siap cetak kapan saja.")

    ac1, ac2 = st.columns(2)
    with ac1:
        a_kelas = st.selectbox("Kelas", KELAS_OPTIONS, key="a_kelas")
    with ac2:
        a_jumlah_pertemuan = st.number_input("Jumlah Kolom Pertemuan/Tanggal pada Lembar Absen", 1, 31, 10, key="a_jp")

    st.markdown("**➕ Tambah Data Siswa**")
    with st.form("form_siswa", clear_on_submit=True):
        sc1, sc2, sc3 = st.columns([3, 1, 1])
        with sc1:
            s_nama = st.text_input("Nama Siswa", key="s_nama")
        with sc2:
            s_jk = st.selectbox("L/P", ["L", "P"], key="s_jk")
        with sc3:
            s_nisn = st.text_input("NISN (opsional)", key="s_nisn")
        submit_siswa = st.form_submit_button("➕ Tambahkan Siswa")

    if submit_siswa:
        if not s_nama:
            st.error("⚠️ Mohon isi nama siswa.")
        else:
            tambah_baris_siswa({"Nama": s_nama, "L/P": s_jk, "NISN": s_nisn})
            st.success(f"✅ {s_nama} ditambahkan ke data siswa {a_kelas}.")

    st.markdown("**📋 Atau tempel banyak nama sekaligus (satu nama per baris)**")
    tempel_siswa = st.text_area("Tempel Daftar Nama", height=100, key="tempel_siswa")
    if st.button("📥 Impor dari Daftar Tempel", key="btn_impor_siswa"):
        nama_list = [n.strip() for n in tempel_siswa.split("\n") if n.strip()]
        for n in nama_list:
            tambah_baris_siswa({"Nama": n, "L/P": "-", "NISN": ""})
        st.success(f"✅ {len(nama_list)} siswa diimpor.")
        st.rerun()

    with st.expander("📎 Atau impor data siswa dari berkas CSV/XLSX/TXT (data dari Dapodik, dll)"):
        st.caption("Kolom yang dikenali: Nama, L/P (atau Jenis Kelamin), NISN. Untuk TXT, satu nama per baris.")
        a_file_upl = st.file_uploader("Unggah berkas data siswa", type=["csv", "xlsx", "xls", "txt"], key="a_file_upl")
        if a_file_upl is not None:
            baris_siswa, status_siswa = baca_tabel_upload(a_file_upl)
            if baris_siswa:
                st.success(status_siswa)
                st.dataframe(baris_siswa, use_container_width=True, hide_index=True)
                if st.button("➕ Tambahkan Semua ke Data Siswa", key="btn_impor_siswa_file"):
                    for r in baris_siswa:
                        keys_lower = {str(k).strip().lower(): k for k in r.keys()}
                        def ambil_s(*alias):
                            for a in alias:
                                if a in keys_lower:
                                    return str(r[keys_lower[a]])
                            return ""
                        nama_ = ambil_s("nama", "nama siswa")
                        if not nama_:
                            continue
                        tambah_baris_siswa({
                            "Nama": nama_,
                            "L/P": ambil_s("l/p", "jk", "jenis kelamin") or "-",
                            "NISN": ambil_s("nisn"),
                        })
                    st.success(f"✅ {len(baris_siswa)} siswa berhasil diimpor.")
                    st.rerun()
            else:
                st.warning(status_siswa)

    if st.session_state.siswa_rows:
        st.markdown(f"**Total siswa terdaftar: {len(st.session_state.siswa_rows)}**")
        st.dataframe([{k: v for k, v in r.items() if k != "id"} for r in st.session_state.siswa_rows],
                     use_container_width=True, hide_index=True)

        sca, scb = st.columns(2)
        with sca:
            if st.button("🗑️ Hapus Siswa Terakhir", key="btn_hapus_siswa"):
                hapus_siswa_terakhir()
                st.rerun()
        with scb:
            if st.button("🧹 Kosongkan Semua Data Siswa", key="btn_kosong_siswa"):
                kosongkan_siswa()
                st.rerun()

        if st.button("🖨️ Buat Lembar Absensi", key="btn_buat_absen", use_container_width=True):
            kolom_tanggal = [f"P{i}" for i in range(1, a_jumlah_pertemuan + 1)]
            header = ["No", "Nama Siswa", "L/P"] + kolom_tanggal
            rows_tabel = [header]
            for i, s in enumerate(st.session_state.siswa_rows, 1):
                rows_tabel.append([str(i), s["Nama"], s["L/P"]] + ["" for _ in kolom_tanggal])

            meta_absen = {
                "Satuan Pendidikan": st.session_state.sekolah,
                "Kelas": a_kelas,
                "Tahun Ajaran": st.session_state.tahun_ajaran,
                "Guru Penyusun": st.session_state.penyusun,
                "NIP Guru": st.session_state.guru_nip,
            }

            # PDF lembar absen (orientasi landscape agar kolom pertemuan muat)
            buffer = BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=landscape(A4),
                                     topMargin=1.5 * cm, bottomMargin=1.8 * cm,
                                     leftMargin=1.5 * cm, rightMargin=1.5 * cm,
                                     title="LEMBAR ABSENSI SISWA")
            styles = build_pdf_styles()
            story = [
                Paragraph("LEMBAR ABSENSI SISWA", styles["title"]),
                Paragraph(f"Kelas {a_kelas} &bull; {st.session_state.sekolah}", styles["subtitle"]),
            ]
            para_rows = [[Paragraph(f"<b>{c}</b>" if r == 0 else c, styles["cell"]) for c in row]
                         for r, row in enumerate(rows_tabel)]
            # Lebar kolom pertemuan menyesuaikan otomatis agar tabel selalu pas di halaman
            # landscape A4 (29.7cm - margin kiri/kanan 1.5cm masing-masing = 26.7cm tersedia),
            # sehingga tidak overflow walau jumlah kolom pertemuan banyak.
            lebar_tersedia = landscape(A4)[0] / cm - 1.5 - 1.5  # cm
            lebar_no, lebar_nama, lebar_lp = 1.0, 5.0, 1.0
            sisa = max(lebar_tersedia - (lebar_no + lebar_nama + lebar_lp), 4.0)
            lebar_per_pertemuan = min(1.0, sisa / max(len(kolom_tanggal), 1))
            col_widths = ([lebar_no * cm, lebar_nama * cm, lebar_lp * cm]
                          + [lebar_per_pertemuan * cm] * len(kolom_tanggal))
            tbl = Table(para_rows, colWidths=col_widths, repeatRows=1)
            tbl.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), COLOR_PRIMARY),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.6, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, COLOR_LIGHT]),
            ]))
            story.append(tbl)
            doc.build(story, onFirstPage=_pdf_footer, onLaterPages=_pdf_footer)
            buffer.seek(0)
            pdf_absen = buffer.getvalue()

            # DOCX lembar absen
            document = Document()
            section = document.sections[0]
            section.orientation = 1  # landscape
            section.page_width, section.page_height = Cm(29.7), Cm(21.0)
            section.top_margin = section.bottom_margin = Cm(1.5)
            section.left_margin = section.right_margin = Cm(1.5)
            p_title = document.add_paragraph()
            p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_title = p_title.add_run("LEMBAR ABSENSI SISWA")
            r_title.bold = True; r_title.font.size = Pt(15); r_title.font.color.rgb = DOCX_PRIMARY_RGB
            p_sub = document.add_paragraph()
            p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_sub.add_run(f"Kelas {a_kelas} • {st.session_state.sekolah}").italic = True

            tbl_docx = document.add_table(rows=0, cols=len(header))
            tbl_docx.style = "Light Grid Accent 1"
            tbl_docx.alignment = WD_TABLE_ALIGNMENT.CENTER
            hdr_cells = tbl_docx.add_row().cells
            for idx, h in enumerate(header):
                hdr_cells[idx].text = ""
                run = hdr_cells[idx].paragraphs[0].add_run(h)
                run.bold = True
            for row in rows_tabel[1:]:
                cells = tbl_docx.add_row().cells
                for idx, v in enumerate(row):
                    cells[idx].text = str(v)

            buffer2 = BytesIO()
            document.save(buffer2)
            buffer2.seek(0)
            docx_absen = buffer2.getvalue()

            st.success("✨ Lembar absensi berhasil dibuat!")
            download_row(f"Absensi_{a_kelas.replace(' ', '_')}", pdf_absen, docx_absen, "dl_absen")
    else:
        st.info("Belum ada data siswa. Tambahkan lewat formulir atau impor daftar tempel di atas.")
