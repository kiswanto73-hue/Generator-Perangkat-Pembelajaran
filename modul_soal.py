# -*- coding: utf-8 -*-
"""
modul_soal.py  -  Modul Generator Soal (UH / PTS / PAS) untuk SD kelas 1-6.

Bisa dipakai dua cara:
  1) Mandiri        : streamlit run modul_soal.py
  2) Tab di app KKG : import modul_soal  ->  modul_soal.render_tab_soal()

Tahap 1-2: template soal PG bergaya stimulus + ekspor PDF/Word,
           gambar stimulus dari kode (matematika) dan dari pustaka gambar.

Kebutuhan (requirements.txt): streamlit, groq, matplotlib, pillow, reportlab, python-docx
"""
import ast
import io
import json
import math
import operator
import os
import re
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle, Polygon, Circle
from PIL import Image

BASE_DIR = Path(__file__).parent
PUSTAKA_DIR = BASE_DIR / "pustaka_gambar"
MODEL_DEFAULT = "llama-3.3-70b-versatile"

# =====================================================================
# 1. GAMBAR DARI KODE  (presisi: angka pada gambar = angka pada soal)
# =====================================================================

def _n(x):
    """Format angka gaya Indonesia: 7.5 -> '7,5'."""
    x = float(x)
    return str(int(x)) if x == int(x) else str(round(x, 2)).replace(".", ",")


def _fig(w=3.6, h=2.6):
    fig, ax = plt.subplots(figsize=(w, h), dpi=200)
    ax.set_aspect("equal")
    ax.axis("off")
    return fig, ax


def _png(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", pad_inches=0.1, facecolor="white")
    plt.close(fig)
    return buf.getvalue()


_LBL = dict(fontsize=11, fontweight="bold")


def gambar_persegi_panjang(panjang, lebar, satuan="m", tiang=False):
    p, l = float(panjang), float(lebar)
    W = 4.0
    H = min(max(W * l / p, 1.4), 3.4)
    fig, ax = _fig()
    ax.add_patch(Rectangle((0, 0), W, H, fill=False, lw=1.4, ec="black"))
    if tiang:  # kotak kecil di sudut (tiang pagar), seperti pada contoh
        s = 0.2
        for x, y in [(0, 0), (W - s, 0), (0, H - s), (W - s, H - s)]:
            ax.add_patch(Rectangle((x, y), s, s, fc="white", ec="black", lw=1.2))
    ax.text(W / 2, H + 0.15, f"{_n(p)} {satuan}", ha="center", va="bottom", **_LBL)
    ax.text(W + 0.15, H / 2, f"{_n(l)} {satuan}", ha="left", va="center", **_LBL)
    ax.set_xlim(-0.2, W + 1.4)
    ax.set_ylim(-0.2, H + 0.7)
    return _png(fig)


def gambar_persegi(sisi, satuan="cm"):
    s = float(sisi)
    fig, ax = _fig(3, 3)
    ax.add_patch(Rectangle((0, 0), 3, 3, fill=False, lw=1.4, ec="black"))
    ax.text(1.5, -0.15, f"{_n(s)} {satuan}", ha="center", va="top", **_LBL)
    ax.text(3.15, 1.5, f"{_n(s)} {satuan}", ha="left", va="center", **_LBL)
    ax.set_xlim(-0.2, 4.6)
    ax.set_ylim(-0.7, 3.2)
    return _png(fig)


def gambar_segitiga(alas, tinggi, satuan="cm", bentuk="siku-siku"):
    a, t = float(alas), float(tinggi)
    W = 4.0
    H = min(max(W * t / a, 1.4), 3.4)
    siku = str(bentuk).lower().startswith("siku")
    ax_ = 0.0 if siku else W * 0.4
    fig, ax = _fig()
    ax.add_patch(Polygon([(0, 0), (W, 0), (ax_, H)], closed=True, fill=False, lw=1.4, ec="black"))
    if not siku:
        ax.plot([ax_, ax_], [0, H], ls="--", lw=1, c="black")
    m = 0.22
    ax.plot([ax_ + m, ax_ + m, ax_], [0, m, m], lw=1, c="black")  # tanda siku-siku
    ax.text(W / 2, -0.15, f"{_n(a)} {satuan}", ha="center", va="top", **_LBL)
    if siku:
        ax.text(-0.15, H / 2, f"{_n(t)} {satuan}", ha="right", va="center", **_LBL)
    else:
        ax.text(ax_ + 0.12, H / 2, f"{_n(t)} {satuan}", ha="left", va="center", **_LBL)
    ax.set_xlim(-1.6, W + 0.4)
    ax.set_ylim(-0.7, H + 0.3)
    return _png(fig)


def gambar_lingkaran(jari_jari=None, diameter=None, satuan="cm"):
    fig, ax = _fig(3, 3)
    R = 1.5
    ax.add_patch(Circle((0, 0), R, fill=False, lw=1.4, ec="black"))
    ax.plot([0], [0], "ko", ms=3)
    if diameter is not None and jari_jari is None:
        ax.plot([-R, R], [0, 0], lw=1.2, c="black")
        ax.text(0, 0.1, f"d = {_n(diameter)} {satuan}", ha="center", va="bottom", **_LBL)
    else:
        ax.plot([0, R], [0, 0], lw=1.2, c="black")
        ax.text(R / 2, 0.1, f"r = {_n(jari_jari)} {satuan}", ha="center", va="bottom", **_LBL)
    ax.set_xlim(-R - 0.2, R + 0.2)
    ax.set_ylim(-R - 0.2, R + 0.2)
    return _png(fig)


def gambar_trapesium(sisi_atas, sisi_bawah, tinggi, satuan="cm"):
    a, b, t = float(sisi_atas), float(sisi_bawah), float(tinggi)
    B = 4.0
    top = B * a / b
    off = (B - top) / 2
    H = min(max(B * t / b, 1.2), 2.8)
    fig, ax = _fig()
    ax.add_patch(Polygon([(0, 0), (B, 0), (B - off, H), (off, H)], closed=True,
                         fill=False, lw=1.4, ec="black"))
    ax.plot([off, off], [0, H], ls="--", lw=1, c="black")
    m = 0.18
    ax.plot([off + m, off + m, off], [0, m, m], lw=1, c="black")
    ax.text(B / 2, H + 0.12, f"{_n(a)} {satuan}", ha="center", va="bottom", **_LBL)
    ax.text(B / 2, -0.15, f"{_n(b)} {satuan}", ha="center", va="top", **_LBL)
    ax.text(off + 0.12, H / 2, f"{_n(t)} {satuan}", ha="left", va="center", **_LBL)
    ax.set_xlim(-0.3, B + 0.3)
    ax.set_ylim(-0.7, H + 0.7)
    return _png(fig)


def gambar_diagram_batang(judul, label, nilai, satuan=""):
    if len(label) != len(nilai) or not label:
        raise ValueError("label dan nilai harus sama panjang")
    nilai = [float(v) for v in nilai]
    fig, ax = plt.subplots(figsize=(4.2, 2.8), dpi=200)
    ax.bar([str(x) for x in label], nilai, color="#b8bcc2", edgecolor="black", lw=1, zorder=3)
    ax.set_title(str(judul), fontsize=10, fontweight="bold")
    if satuan:
        ax.set_ylabel(satuan, fontsize=9)
    ax.grid(axis="y", lw=0.5, alpha=0.6, zorder=0)
    ax.tick_params(labelsize=9)
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    ax.set_ylim(0, max(nilai) * 1.15)
    return _png(fig)


def gambar_garis_bilangan(awal, akhir, langkah=1, titik=None):
    awal, akhir, langkah = float(awal), float(akhir), float(langkah)
    n = int(round((akhir - awal) / langkah))
    if n < 1 or n > 25:
        raise ValueError("rentang garis bilangan tidak wajar")
    fig, ax = plt.subplots(figsize=(5.2, 1.4), dpi=200)
    ax.axis("off")
    ax.annotate("", xy=(akhir + langkah * 0.6, 0), xytext=(awal - langkah * 0.6, 0),
                arrowprops=dict(arrowstyle="<|-|>", lw=1.2, color="black"))
    for i in range(n + 1):
        x = awal + i * langkah
        ax.plot([x, x], [-0.08, 0.08], c="black", lw=1)
        ax.text(x, -0.16, _n(x), ha="center", va="top", fontsize=9)
    for k, p in enumerate(titik or []):
        if isinstance(p, dict):
            val, lab = float(p["nilai"]), str(p.get("label", ""))
        else:
            val, lab = float(p), ""
        ax.plot([val], [0], "o", c="black", ms=6)
        if lab:
            ax.text(val, 0.18, lab, ha="center", va="bottom", **_LBL)
    ax.set_xlim(awal - langkah, akhir + langkah)
    ax.set_ylim(-0.6, 0.6)
    return _png(fig)


def gambar_jam(jam, menit=0):
    jam, menit = int(jam), int(menit)
    fig, ax = _fig(2.6, 2.6)
    ax.add_patch(Circle((0, 0), 1, fill=False, lw=2, ec="black"))
    for i in range(60):
        a = math.radians(90 - i * 6)
        r1 = 0.92 if i % 5 == 0 else 0.96
        ax.plot([r1 * math.cos(a), math.cos(a)], [r1 * math.sin(a), math.sin(a)], c="black",
                lw=1.2 if i % 5 == 0 else 0.5)
    for h in range(1, 13):
        a = math.radians(90 - h * 30)
        ax.text(0.78 * math.cos(a), 0.78 * math.sin(a), str(h), ha="center", va="center",
                fontsize=11, fontweight="bold")
    ah = math.radians(90 - ((jam % 12) + menit / 60) * 30)
    am = math.radians(90 - menit * 6)
    ax.plot([0, 0.4 * math.cos(ah)], [0, 0.4 * math.sin(ah)], c="black", lw=3.5, solid_capstyle="round")
    ax.plot([0, 0.62 * math.cos(am)], [0, 0.62 * math.sin(am)], c="black", lw=2, solid_capstyle="round")
    ax.plot([0], [0], "ko", ms=4)
    ax.set_xlim(-1.1, 1.1)
    ax.set_ylim(-1.1, 1.1)
    return _png(fig)


GAMBAR_KODE = {
    "persegi_panjang": gambar_persegi_panjang,
    "persegi": gambar_persegi,
    "segitiga": gambar_segitiga,
    "lingkaran": gambar_lingkaran,
    "trapesium": gambar_trapesium,
    "diagram_batang": gambar_diagram_batang,
    "garis_bilangan": gambar_garis_bilangan,
    "jam": gambar_jam,
}

GAMBAR_DOKUMEN = """Pilihan isi field "gambar" (pilih SATU, atau null bila gambar tidak membantu):
- {"jenis":"persegi_panjang","param":{"panjang":20,"lebar":15,"satuan":"m","tiang":false}}
- {"jenis":"persegi","param":{"sisi":12,"satuan":"cm"}}
- {"jenis":"segitiga","param":{"alas":10,"tinggi":8,"satuan":"cm","bentuk":"siku-siku"}}   (bentuk: "siku-siku" atau "sembarang")
- {"jenis":"lingkaran","param":{"jari_jari":7,"satuan":"cm"}}   (atau "diameter":14)
- {"jenis":"trapesium","param":{"sisi_atas":6,"sisi_bawah":10,"tinggi":4,"satuan":"cm"}}
- {"jenis":"diagram_batang","param":{"judul":"Hasil Panen","label":["Senin","Selasa"],"nilai":[12,18],"satuan":"kg"}}
- {"jenis":"garis_bilangan","param":{"awal":0,"akhir":10,"langkah":1,"titik":[3,7]}}
- {"jenis":"jam","param":{"jam":3,"menit":30}}
- {"jenis":"pustaka","objek":["bus","halte"]}   (benda sehari-hari: dicarikan dari pustaka gambar)
Angka pada param gambar HARUS sama dengan angka pada stimulus. Jangan menulis teks apa pun di dalam gambar."""

# =====================================================================
# 2. PUSTAKA GAMBAR  (pustaka_gambar/index.json + file PNG)
# =====================================================================

def muat_indeks():
    f = PUSTAKA_DIR / "index.json"
    if not f.exists():
        return []
    try:
        return json.loads(f.read_text(encoding="utf-8"))
    except Exception:
        return []


def cari_gambar_pustaka(objek):
    """Kembalikan daftar bytes gambar (satu per objek yang ketemu, tanpa duplikat)."""
    indeks = muat_indeks()
    hasil, dipakai = [], set()
    for o in objek or []:
        o = str(o).lower().strip()
        terbaik, skor_max = None, 0
        for item in indeks:
            f = PUSTAKA_DIR / item.get("file", "")
            if not f.is_file():
                continue
            kata = [k.lower() for k in item.get("kata_kunci", [])] + [item.get("nama", "").lower()]
            skor = sum(2 if o == k else 1 for k in kata if k and (o in k or k in o))
            if skor > skor_max:
                terbaik, skor_max = item, skor
        if terbaik and terbaik["file"] not in dipakai:
            dipakai.add(terbaik["file"])
            hasil.append((PUSTAKA_DIR / terbaik["file"]).read_bytes())
    return hasil


def gabung_horizontal(daftar, tinggi=320, jarak=24):
    if len(daftar) == 1:
        return daftar[0]
    ims = []
    for b in daftar:
        im = Image.open(io.BytesIO(b)).convert("RGBA")
        w = int(im.width * tinggi / im.height)
        ims.append(im.resize((w, tinggi)))
    total = sum(i.width for i in ims) + jarak * (len(ims) - 1)
    kanvas = Image.new("RGBA", (total, tinggi), (255, 255, 255, 255))
    x = 0
    for im in ims:
        kanvas.alpha_composite(im, (x, 0))
        x += im.width + jarak
    out = io.BytesIO()
    kanvas.convert("RGB").save(out, format="PNG")
    return out.getvalue()


def buat_gambar(spec, upload=None):
    """Spec gambar -> bytes PNG, atau None bila tidak ada / gagal (soal tetap jalan tanpa gambar)."""
    if upload:
        return upload
    if not spec or not isinstance(spec, dict):
        return None
    try:
        jenis = spec.get("jenis")
        if jenis == "pustaka":
            ketemu = cari_gambar_pustaka(spec.get("objek", []))
            return gabung_horizontal(ketemu) if ketemu else None
        fungsi = GAMBAR_KODE.get(jenis)
        return fungsi(**(spec.get("param") or {})) if fungsi else None
    except Exception:
        return None


def gambar_untuk(s):
    """Gambar sebuah soal (dengan cache sederhana di dalam dict soal)."""
    if s.get("_hapus_gambar"):
        return None
    if s.get("_upload"):
        return s["_upload"]
    kunci = json.dumps(s.get("gambar"), sort_keys=True, default=str)
    if s.get("_ck") != kunci:
        s["_cb"] = buat_gambar(s.get("gambar"))
        s["_ck"] = kunci
    return s.get("_cb")

# =====================================================================
# 3. VERIFIKASI KUNCI JAWABAN
# =====================================================================

_OPS = {ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul,
        ast.Div: operator.truediv}


def hitung_aman(ekspresi):
    """Hitung ekspresi aritmetika sederhana tanpa eval()."""
    def ev(n):
        if isinstance(n, ast.Expression):
            return ev(n.body)
        if isinstance(n, ast.Constant) and isinstance(n.value, (int, float)):
            return n.value
        if isinstance(n, ast.BinOp) and type(n.op) in _OPS:
            return _OPS[type(n.op)](ev(n.left), ev(n.right))
        if isinstance(n, ast.UnaryOp) and isinstance(n.op, (ast.USub, ast.UAdd)):
            v = ev(n.operand)
            return -v if isinstance(n.op, ast.USub) else v
        raise ValueError("ekspresi tidak diizinkan")
    return float(ev(ast.parse(str(ekspresi).strip(), mode="eval")))


_RE_ANGKA = re.compile(r"\d{1,3}(?:\.\d{3})+(?:,\d+)?|\d+(?:,\d+)?")


def parse_angka(teks):
    """'Rp1.050.000' -> 1050000.0 ; '3,5 jam' -> 3.5 ; tidak ada angka -> None."""
    m = _RE_ANGKA.search(str(teks))
    if not m:
        return None
    return float(m.group(0).replace(".", "").replace(",", "."))


def verifikasi(s):
    """Periksa struktur + kunci. Mengisi s['_status'] dan s['_catatan']; boleh memperbaiki kunci."""
    opsi = s.get("opsi", {})
    if sorted(opsi.keys()) != ["A", "B", "C", "D"] or any(not str(v).strip() for v in opsi.values()):
        s["_status"], s["_catatan"] = "gagal", "Pilihan jawaban harus lengkap A-D."
        return s
    if len({str(v).strip().lower() for v in opsi.values()}) < 4:
        s["_status"], s["_catatan"] = "gagal", "Ada pilihan jawaban yang kembar."
        return s
    if s.get("kunci") not in opsi:
        s["_status"], s["_catatan"] = "gagal", "Kunci jawaban tidak valid."
        return s
    ek = s.get("ekspresi_kunci")
    if not ek:
        s["_status"], s["_catatan"] = "tanpa_cek", "Bukan soal hitungan: periksa kunci secara manual."
        return s
    try:
        hasil = hitung_aman(ek)
    except Exception:
        s["_status"], s["_catatan"] = "tanpa_cek", "Ekspresi hitung tidak bisa dibaca: periksa manual."
        return s
    nilai = {h: parse_angka(v) for h, v in opsi.items()}
    cocok = [h for h, v in nilai.items() if v is not None and math.isclose(v, hasil, rel_tol=1e-6, abs_tol=1e-9)]
    if s["kunci"] in cocok:
        s["_status"], s["_catatan"] = "ok", f"Terverifikasi: hasil hitung = {_n(hasil)}."
    elif len(cocok) == 1:
        lama = s["kunci"]
        s["kunci"] = cocok[0]
        s["_status"] = "diperbaiki"
        s["_catatan"] = f"Kunci diperbaiki otomatis {lama} -> {cocok[0]} (hasil hitung = {_n(hasil)})."
    else:
        s["_status"] = "gagal"
        s["_catatan"] = (f"Hasil hitung = {_n(hasil)} tetapi tidak ada pilihan yang cocok. "
                         "Perbaiki soal atau pilihan jawabannya.")
    return s


def normalisasi(s):
    opsi = s.get("opsi") or {}
    out = {
        "stimulus": str(s.get("stimulus") or "").strip(),
        "pertanyaan": str(s.get("pertanyaan") or "").strip(),
        "opsi": {h: str(opsi.get(h, "")).strip() for h in "ABCD"},
        "kunci": str(s.get("kunci") or "").strip().upper()[:1],
        "pembahasan": str(s.get("pembahasan") or "").strip(),
        "ekspresi_kunci": s.get("ekspresi_kunci") or None,
        "gambar": s.get("gambar") if isinstance(s.get("gambar"), dict) else None,
    }
    return verifikasi(out)

# =====================================================================
# 4. GENERATE SOAL DENGAN AI (Groq)
# =====================================================================

def buat_prompt(p):
    return (
        "Buatkan soal pilihan ganda untuk siswa SD sesuai Kurikulum Merdeka.\n"
        f"- Jenis ujian : {p['jenis_ujian']}\n- Kelas       : {p['kelas']}\n"
        f"- Mapel       : {p['mapel']}\n- Materi/Bab  : {p['materi']}\n"
        f"- Jumlah soal : {p['jumlah']}\n- Tingkat     : {p['tingkat']}\n\n"
        "ATURAN:\n"
        "1. Bahasa Indonesia baku, panjang kalimat sesuai kelas (kelas 1-2: kalimat pendek dan sederhana).\n"
        "2. 'stimulus' = cerita/data singkat yang memuat SEMUA angka dan satuan yang dibutuhkan. "
        "'pertanyaan' = satu kalimat tanya.\n"
        "3. Empat pilihan A-D, hanya satu benar. Pengecoh berupa kesalahan umum siswa. Tulis angka gaya Indonesia "
        "(titik ribuan, koma desimal), pilihan angka diurutkan naik, panjang pilihan seimbang.\n"
        "4. 'ekspresi_kunci': untuk soal hitungan, tulis ekspresi Python (hanya angka dengan TITIK desimal, "
        "operator + - * / dan kurung) yang hasilnya sama dengan nilai pilihan yang benar, dalam satuan yang sama "
        "dengan pilihan tersebut. Untuk soal bukan hitungan isi null.\n"
        "5. 'pembahasan' = langkah penyelesaian singkat.\n"
        "6. " + GAMBAR_DOKUMEN + "\n\n"
        "Balas HANYA JSON valid dengan bentuk:\n"
        '{"soal":[{"stimulus":"...","pertanyaan":"...","opsi":{"A":"...","B":"...","C":"...","D":"..."},'
        '"kunci":"B","pembahasan":"...","ekspresi_kunci":null,"gambar":null}]}'
    )


def _ambil_json(teks):
    try:
        return json.loads(teks)
    except Exception:
        m = re.search(r"\{.*\}", teks, re.S)
        if not m:
            raise ValueError("Jawaban AI bukan JSON.")
        return json.loads(m.group(0))


def generate_soal_ai(client, p, model=MODEL_DEFAULT):
    resp = client.chat.completions.create(
        model=model, temperature=0.4,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": "Kamu guru SD berpengalaman dan penyusun soal asesmen. "
                                          "Keluarkan hanya JSON valid, tanpa teks lain."},
            {"role": "user", "content": buat_prompt(p)},
        ],
    )
    data = _ambil_json(resp.choices[0].message.content)
    return [normalisasi(s) for s in data.get("soal", [])]


CONTOH_SOAL = [
    {"stimulus": "Sebuah bus kota berangkat dari terminal pukul 07.15 dan tiba di terminal tujuan pukul 10.45. "
                 "Sepanjang perjalanan, bus berhenti di tiga halte masing-masing selama 10 menit. "
                 "Jarak yang ditempuh bus adalah 180 km.",
     "pertanyaan": "Berapa kecepatan rata-rata bus tersebut saat bergerak (tanpa waktu berhenti)?",
     "opsi": {"A": "50 km/jam", "B": "55 km/jam", "C": "60 km/jam", "D": "65 km/jam"}, "kunci": "C",
     "pembahasan": "Waktu total 3 jam 30 menit. Waktu berhenti 3 x 10 = 30 menit. Waktu bergerak = 3 jam. "
                   "Kecepatan = 180 : 3 = 60 km/jam.",
     "ekspresi_kunci": "180/((10*60+45-(7*60+15)-3*10)/60)",
     "gambar": {"jenis": "pustaka", "objek": ["bus", "halte"]}},
    {"stimulus": "Pak Budi memiliki sepetak tanah berbentuk persegi panjang dengan panjang 20 m dan lebar 15 m. "
                 "Ia ingin membuat pagar di sekeliling tanah tersebut. Setiap meter pagar membutuhkan 2 batang "
                 "bambu. Harga satu batang bambu Rp7.500.",
     "pertanyaan": "Berapa total biaya yang dibutuhkan Pak Budi untuk membeli bambu?",
     "opsi": {"A": "Rp1.050.000", "B": "Rp1.150.000", "C": "Rp1.250.000", "D": "Rp1.350.000"}, "kunci": "A",
     "pembahasan": "Keliling = 2 x (20 + 15) = 70 m. Bambu = 70 x 2 = 140 batang. Biaya = 140 x Rp7.500 = Rp1.050.000.",
     "ekspresi_kunci": "2*(20+15)*2*7500",
     "gambar": {"jenis": "persegi_panjang", "param": {"panjang": 20, "lebar": 15, "satuan": "m", "tiang": True}}},
    {"stimulus": "Toko Pak Ahmad memberikan harga khusus untuk pembelian grosir. Satu lusin buku tulis dijual "
                 "Rp36.000. Jika membeli 3 lusin atau lebih, pembeli mendapat potongan 10% dari total harga. "
                 "Bu Sari membeli 5 lusin buku tulis.",
     "pertanyaan": "Berapa yang harus dibayar Bu Sari?",
     "opsi": {"A": "Rp150.000", "B": "Rp162.000", "C": "Rp170.000", "D": "Rp180.000"}, "kunci": "B",
     "pembahasan": "Total = 5 x Rp36.000 = Rp180.000. Potongan 10% = Rp18.000. Bayar = Rp162.000.",
     "ekspresi_kunci": "5*36000*(1-10/100)",
     "gambar": {"jenis": "pustaka", "objek": ["buku tulis"]}},
]

# =====================================================================
# 5. EKSPOR PDF  (reportlab)
# =====================================================================

def _esc(t):
    t = str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return re.sub(r"\^(\d+)", r"<super>\1</super>", t)


def buat_pdf(soal, kop, sertakan_kunci=True):
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import (HRFlowable, Image as RLImage, PageBreak, Paragraph,
                                    SimpleDocTemplate, Spacer, Table, TableStyle)

    M = 1.8 * cm
    LEBAR = A4[0] - 2 * M
    base = ParagraphStyle("b", fontName="Helvetica", fontSize=10, leading=13.5)
    st_judul = ParagraphStyle("j", parent=base, fontName="Helvetica-Bold", fontSize=11, spaceAfter=3)
    st_label = ParagraphStyle("l", parent=base, fontName="Helvetica-Bold", spaceBefore=4, spaceAfter=1)
    st_opsi = ParagraphStyle("o", parent=base, leftIndent=14, spaceBefore=1)
    st_kop1 = ParagraphStyle("k1", parent=base, fontName="Helvetica-Bold", fontSize=13, alignment=TA_CENTER, leading=16)
    st_kop2 = ParagraphStyle("k2", parent=base, fontName="Helvetica-Bold", fontSize=11, alignment=TA_CENTER, leading=14)

    cerita = []
    cerita.append(Paragraph(_esc((kop.get("sekolah") or "NAMA SEKOLAH").upper()), st_kop1))
    judul = kop.get("jenis_ujian", "")
    if kop.get("tahun_ajaran"):
        judul += f" - TAHUN AJARAN {kop['tahun_ajaran']}"
    cerita += [Paragraph(_esc(judul.upper()), st_kop2), Spacer(1, 4),
               HRFlowable(width="100%", thickness=1.2, color=colors.black), Spacer(1, 4)]
    info = [
        [Paragraph("Mata Pelajaran", base), Paragraph(": " + _esc(kop.get("mapel", "")), base),
         Paragraph("Hari/Tanggal", base), Paragraph(": ..............................", base)],
        [Paragraph("Kelas", base), Paragraph(": " + _esc(kop.get("kelas", "")), base),
         Paragraph("Waktu", base), Paragraph(": " + _esc(kop.get("waktu", "") or "........ menit"), base)],
        [Paragraph("Nama", base), Paragraph(": ..............................", base),
         Paragraph("No. Absen", base), Paragraph(": ..........", base)],
    ]
    t = Table(info, colWidths=[3.2 * cm, 5.5 * cm, 3 * cm, LEBAR - 11.7 * cm])
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                           ("TOPPADDING", (0, 0), (-1, -1), 1), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    cerita += [t, Spacer(1, 3), HRFlowable(width="100%", thickness=1.2, color=colors.black), Spacer(1, 10)]

    GAMBAR_W, GAMBAR_H = 5.4 * cm, 4.6 * cm
    for no, s in enumerate(soal, 1):
        kiri = [Paragraph(f"{no}.&nbsp;&nbsp;SOAL {no} - PG".replace("-", "\u2014"), st_judul)]
        if s["stimulus"]:
            kiri += [Paragraph("Stimulus:", st_label), Paragraph(_esc(s["stimulus"]), base)]
        kiri += [Paragraph("Pertanyaan:", st_label), Paragraph(_esc(s["pertanyaan"]), base), Spacer(1, 2)]
        for h in "ABCD":
            kiri.append(Paragraph(f"{h}.&nbsp;&nbsp;&nbsp;{_esc(s['opsi'][h])}", st_opsi))
        gb = gambar_untuk(s)
        if gb:
            im = Image.open(io.BytesIO(gb))
            skala = min(GAMBAR_W / im.width, GAMBAR_H / im.height)
            img = RLImage(io.BytesIO(gb), width=im.width * skala, height=im.height * skala)
            blok = Table([[kiri, img]], colWidths=[LEBAR - GAMBAR_W - 0.4 * cm, GAMBAR_W + 0.4 * cm])
            blok.setStyle(TableStyle([("VALIGN", (0, 0), (0, 0), "TOP"), ("VALIGN", (1, 0), (1, 0), "MIDDLE"),
                                      ("ALIGN", (1, 0), (1, 0), "CENTER"),
                                      ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
        else:
            blok = Table([[kiri]], colWidths=[LEBAR])
            blok.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0)]))
        cerita += [blok, Spacer(1, 14)]

    if sertakan_kunci:
        cerita.append(PageBreak())
        cerita += [Paragraph("KUNCI JAWABAN DAN PEMBAHASAN", st_kop2), Spacer(1, 8)]
        baris = [[Paragraph("<b>No</b>", base), Paragraph("<b>Kunci</b>", base), Paragraph("<b>Pembahasan</b>", base)]]
        for no, s in enumerate(soal, 1):
            baris.append([Paragraph(str(no), base), Paragraph(s["kunci"], base),
                          Paragraph(_esc(s["pembahasan"]) or "-", base)])
        tk = Table(baris, colWidths=[1.2 * cm, 1.6 * cm, LEBAR - 2.8 * cm], repeatRows=1)
        tk.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.6, colors.black),
                                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e6e6e6")),
                                ("VALIGN", (0, 0), (-1, -1), "TOP")]))
        cerita.append(tk)

    def footer(c, d):
        c.saveState()
        c.setFont("Helvetica", 8)
        c.drawCentredString(A4[0] / 2, 1.0 * cm, f"Halaman {d.page}")
        c.restoreState()

    buf = io.BytesIO()
    SimpleDocTemplate(buf, pagesize=A4, leftMargin=M, rightMargin=M, topMargin=1.6 * cm,
                      bottomMargin=1.8 * cm, title=kop.get("jenis_ujian", "Soal")).build(
        cerita, onFirstPage=footer, onLaterPages=footer)
    return buf.getvalue()

# =====================================================================
# 6. EKSPOR WORD  (python-docx)
# =====================================================================

def _teks_docx(par, teks, bold=False, size=10.5):
    """Tulis teks ke paragraf; ^2 menjadi pangkat (superscript)."""
    from docx.shared import Pt
    for i, bagian in enumerate(re.split(r"\^(\d+)", str(teks))):
        if bagian == "":
            continue
        r = par.add_run(bagian)
        r.bold = bold
        r.font.size = Pt(size)
        if i % 2 == 1:
            r.font.superscript = True


def _tanpa_pecah(row):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:cantSplit")
    trPr.append(el)


def _atur_lebar(tabel, lebar_cm):
    """Set lebar kolom (cm) agar dipatuhi Word maupun LibreOffice."""
    from docx.shared import Cm
    tabel.autofit = False
    for kol, w in zip(tabel.columns, lebar_cm):
        kol.width = Cm(w)
        for c in kol.cells:
            c.width = Cm(w)


def buat_docx(soal, kop, sertakan_kunci=True):
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(1.8)
    sec.top_margin, sec.bottom_margin = Cm(1.6), Cm(1.8)
    stl = doc.styles["Normal"]
    stl.font.name = "Arial"
    stl.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    stl.font.size = Pt(10.5)
    stl.paragraph_format.space_after = Pt(2)
    LEBAR = 17.4

    def garis(par):
        pPr = par._p.get_or_add_pPr()
        b = OxmlElement("w:pBdr")
        bt = OxmlElement("w:bottom")
        for k, v in (("w:val", "single"), ("w:sz", "12"), ("w:space", "1"), ("w:color", "000000")):
            bt.set(qn(k), v)
        b.append(bt)
        pPr.append(b)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _teks_docx(p, (kop.get("sekolah") or "NAMA SEKOLAH").upper(), True, 13)
    judul = kop.get("jenis_ujian", "")
    if kop.get("tahun_ajaran"):
        judul += f" - TAHUN AJARAN {kop['tahun_ajaran']}"
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _teks_docx(p, judul.upper(), True, 11)
    garis(p)

    info = [("Mata Pelajaran", kop.get("mapel", ""), "Hari/Tanggal", ".............................."),
            ("Kelas", kop.get("kelas", ""), "Waktu", kop.get("waktu", "") or "........ menit"),
            ("Nama", "..............................", "No. Absen", "..........")]
    t = doc.add_table(rows=3, cols=4)
    for i, baris in enumerate(info):
        for j, (isi, w) in enumerate(zip(baris, (3.2, 5.5, 3.0, 5.7))):
            c = t.cell(i, j)
            c.width = Cm(w)
            _teks_docx(c.paragraphs[0], (": " if j % 2 else "") + str(isi))
    _atur_lebar(t, (3.2, 5.5, 3.0, 5.7))
    p = doc.add_paragraph()
    garis(p)

    GW = 5.2
    for no, s in enumerate(soal, 1):
        gb = gambar_untuk(s)
        lebar_kiri = LEBAR - GW - 0.3 if gb else LEBAR
        tb = doc.add_table(rows=1, cols=2 if gb else 1)
        _tanpa_pecah(tb.rows[0])
        kiri = tb.cell(0, 0)
        kiri.width = Cm(lebar_kiri)
        _teks_docx(kiri.paragraphs[0], f"{no}.  SOAL {no} \u2014 PG", True, 11)
        if s["stimulus"]:
            _teks_docx(kiri.add_paragraph(), "Stimulus:", True)
            _teks_docx(kiri.add_paragraph(), s["stimulus"])
        _teks_docx(kiri.add_paragraph(), "Pertanyaan:", True)
        _teks_docx(kiri.add_paragraph(), s["pertanyaan"])
        for h in "ABCD":
            po = kiri.add_paragraph()
            po.paragraph_format.left_indent = Cm(0.5)
            _teks_docx(po, f"{h}.   {s['opsi'][h]}")
        if gb:
            kanan = tb.cell(0, 1)
            kanan.width = Cm(GW + 0.3)
            im = Image.open(io.BytesIO(gb))
            lebar = min(GW - 0.5, 4.6 * im.width / im.height)
            kanan.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            kanan.paragraphs[0].add_run().add_picture(io.BytesIO(gb), width=Cm(lebar))
            from docx.enum.table import WD_ALIGN_VERTICAL
            kanan.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _atur_lebar(tb, (lebar_kiri, GW + 0.3) if gb else (LEBAR,))
        doc.add_paragraph()

    if sertakan_kunci:
        doc.add_page_break()
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _teks_docx(p, "KUNCI JAWABAN DAN PEMBAHASAN", True, 11)
        tk = doc.add_table(rows=1, cols=3)
        tk.style = "Table Grid"
        for c, teks, w in zip(tk.rows[0].cells, ("No", "Kunci", "Pembahasan"), (1.2, 1.6, LEBAR - 2.8)):
            c.width = Cm(w)
            _teks_docx(c.paragraphs[0], teks, True)
        for no, s in enumerate(soal, 1):
            r = tk.add_row().cells
            for c, teks, w in zip(r, (str(no), s["kunci"], s["pembahasan"] or "-"), (1.2, 1.6, LEBAR - 2.8)):
                c.width = Cm(w)
                _teks_docx(c.paragraphs[0], teks)
        _atur_lebar(tk, (1.2, 1.6, LEBAR - 2.8))

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# =====================================================================
# 7. TAMPILAN STREAMLIT
# =====================================================================

MAPEL = ["Matematika", "Bahasa Indonesia", "IPAS", "Pendidikan Pancasila", "Bahasa Inggris",
         "Pendidikan Agama Islam", "PJOK", "Seni Budaya (SBdP)", "Bahasa Daerah", "Lainnya"]
JENIS_UJIAN = ["Ulangan Harian (UH)", "Penilaian Tengah Semester (PTS)", "Penilaian Akhir Semester (PAS)"]
LABEL_STATUS = {"ok": "\u2705", "diperbaiki": "\U0001F527", "gagal": "\u26A0\uFE0F", "tanpa_cek": "\u2139\uFE0F"}


def _ambil_kunci_groq(st):
    try:
        k = st.secrets.get("GROQ_API_KEY")
        if k:
            return k
    except Exception:
        pass
    return os.environ.get("GROQ_API_KEY")


def render_tab_soal():
    import streamlit as st

    st.subheader("\U0001F4DD Generator Soal (UH / PTS / PAS)")
    st.session_state.setdefault("soal_list", [])
    st.session_state.setdefault("soal_uid", 0)

    c1, c2, c3 = st.columns(3)
    jenis = c1.selectbox("Jenis ujian", JENIS_UJIAN, key="sj_jenis")
    kelas = c2.selectbox("Kelas", [1, 2, 3, 4, 5, 6], index=3, key="sj_kelas")
    mapel = c3.selectbox("Mata pelajaran", MAPEL, key="sj_mapel")
    if mapel == "Lainnya":
        mapel = st.text_input("Nama mata pelajaran", key="sj_mapel_lain")
    materi = st.text_area("Materi / Bab (boleh lebih dari satu)", height=80, key="sj_materi",
                          placeholder="Contoh: Kecepatan, jarak, dan waktu; Keliling bangun datar; Perbandingan")
    c4, c5 = st.columns(2)
    jumlah = c4.number_input("Jumlah soal", 1, 40, 5, key="sj_jumlah")
    tingkat = c5.selectbox("Tingkat kesulitan", ["Campuran", "Mudah", "Sedang", "HOTS"], key="sj_tingkat")

    with st.expander("Kop soal & pengaturan"):
        k1, k2 = st.columns(2)
        sekolah = k1.text_input("Nama sekolah", key="sj_sekolah")
        tahun = k2.text_input("Tahun ajaran", placeholder="2026/2027", key="sj_tahun")
        waktu = k1.text_input("Waktu", placeholder="90 menit", key="sj_waktu")
        model = k2.text_input("Model Groq", MODEL_DEFAULT, key="sj_model")

    b1, b2 = st.columns([2, 1])
    if b1.button("\u2728 Generate soal", type="primary", use_container_width=True):
        kunci_api = _ambil_kunci_groq(st)
        if not materi.strip():
            st.warning("Isi materi/bab dulu.")
        elif not kunci_api:
            st.error("GROQ_API_KEY belum diatur (st.secrets atau environment).")
        else:
            try:
                from groq import Groq
                with st.spinner("AI sedang menyusun dan memeriksa soal..."):
                    hasil = generate_soal_ai(
                        Groq(api_key=kunci_api),
                        dict(jenis_ujian=jenis, kelas=kelas, mapel=mapel, materi=materi.strip(),
                             jumlah=int(jumlah), tingkat=tingkat), model=model)
                st.session_state["soal_list"] = hasil
                st.session_state["soal_uid"] += 1
            except Exception as e:
                st.error(f"Gagal generate: {e}")
    if b2.button("Coba contoh (tanpa AI)", use_container_width=True):
        st.session_state["soal_list"] = [normalisasi(dict(s)) for s in CONTOH_SOAL]
        st.session_state["soal_uid"] += 1

    soal = st.session_state["soal_list"]
    if not soal:
        return
    uid = st.session_state["soal_uid"]
    st.divider()
    st.markdown(f"**{len(soal)} soal** \u2014 periksa dan edit bila perlu.")

    for i, s in enumerate(soal):
        k = f"{uid}_{i}"
        ikon = LABEL_STATUS.get(s.get("_status"), "")
        with st.expander(f"{ikon} Soal {i + 1}", expanded=(s.get("_status") == "gagal")):
            if s.get("_status") == "gagal":
                st.error(s["_catatan"])
            elif s.get("_catatan"):
                st.caption(s["_catatan"])
            kiri, kanan = st.columns([3, 1])
            with kiri:
                s["stimulus"] = st.text_area("Stimulus", s["stimulus"], key=f"st_{k}", height=110)
                s["pertanyaan"] = st.text_input("Pertanyaan", s["pertanyaan"], key=f"pt_{k}")
                for h in "ABCD":
                    s["opsi"][h] = st.text_input(f"Pilihan {h}", s["opsi"][h], key=f"o{h}_{k}")
                s["kunci"] = st.selectbox("Kunci jawaban", list("ABCD"),
                                          index="ABCD".index(s["kunci"]) if s["kunci"] in "ABCD" and s["kunci"] else 0,
                                          key=f"kc_{k}")
                s["pembahasan"] = st.text_area("Pembahasan", s["pembahasan"], key=f"pb_{k}", height=80)
            with kanan:
                gb = gambar_untuk(s)
                if gb:
                    st.image(gb, use_container_width=True)
                else:
                    st.caption("Tanpa gambar")
                up = st.file_uploader("Ganti gambar", type=["png", "jpg", "jpeg"], key=f"up_{k}")
                if up is not None:
                    s["_upload"] = up.getvalue()
                s["_hapus_gambar"] = st.checkbox("Tanpa gambar", s.get("_hapus_gambar", False), key=f"hg_{k}")
            # kunci diedit manual -> kunci dianggap pilihan guru; status dihitung ulang kecuali yang diperbaiki
            verifikasi(s)

    st.divider()
    kop = dict(sekolah=sekolah, jenis_ujian=jenis, tahun_ajaran=tahun, mapel=mapel,
               kelas=f"{kelas} (SD)", waktu=waktu)
    sertakan = st.checkbox("Sertakan halaman kunci jawaban & pembahasan", True)
    d1, d2 = st.columns(2)
    nama = re.sub(r"[^A-Za-z0-9_-]+", "_", f"soal_{mapel}_kelas{kelas}")
    try:
        d1.download_button("\u2B07\uFE0F Unduh PDF", buat_pdf(soal, kop, sertakan), f"{nama}.pdf",
                           "application/pdf", use_container_width=True)
        d2.download_button("\u2B07\uFE0F Unduh Word (.docx)", buat_docx(soal, kop, sertakan), f"{nama}.docx",
                           "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           use_container_width=True)
    except Exception as e:
        st.error(f"Gagal membuat file: {e}")


if __name__ == "__main__":
    import streamlit as st
    st.set_page_config(page_title="Generator Soal", page_icon="\U0001F4DD", layout="wide")
    render_tab_soal()
