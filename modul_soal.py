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
import base64
import io
import json
import math
import operator
import os
import re
import uuid
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle, Polygon, Circle
from PIL import Image

BASE_DIR = Path(__file__).parent
PUSTAKA_DIR = BASE_DIR / "pustaka_gambar"
MODEL_DEFAULT = "openai/gpt-oss-120b"

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


class PustakaSheet:
    """Pustaka gambar milik guru di Google Sheet (lewat Apps Script).
    Dibuat per sesi pengguna (bukan variabel global) supaya data antar guru tidak tercampur."""

    def __init__(self, meta=None, ambil=None, versi=0):
        self.meta = meta if meta is not None else []
        self.ambil = ambil
        self.versi = versi


def _kata_kunci_list(item):
    k = item.get("kata_kunci", [])
    if isinstance(k, str):
        k = k.split(",")
    return [str(x).strip().lower() for x in k if str(x).strip()]


def _skor(objek, item):
    kata = _kata_kunci_list(item) + [str(item.get("nama", "")).strip().lower()]
    return sum(2 if objek == k else 1 for k in kata if k and (objek in k or k in objek))


def cari_gambar_pustaka(objek, pustaka=None):
    """Kembalikan daftar bytes gambar (satu per objek yang ketemu, tanpa duplikat).
    Gambar milik guru di Google Sheet diprioritaskan, lalu pustaka bawaan di folder pustaka_gambar/."""
    kandidat = []
    if pustaka is not None and pustaka.ambil:
        for item in pustaka.meta:
            if item.get("id"):
                kandidat.append(("sheet:" + str(item["id"]), item, lambda it=item: pustaka.ambil(it["id"])))
    for item in muat_indeks():
        f = PUSTAKA_DIR / item.get("file", "")
        if f.is_file():
            kandidat.append(("file:" + item["file"], item, lambda ff=f: ff.read_bytes()))
    hasil, dipakai = [], set()
    for o in objek or []:
        o = str(o).lower().strip()
        if not o:
            continue
        terbaik, skor_max = None, 0
        for kunci, item, ambil in kandidat:
            sk = _skor(o, item)
            if sk > skor_max:
                terbaik, skor_max = (kunci, ambil), sk
        if terbaik and terbaik[0] not in dipakai:
            try:
                data = terbaik[1]()
            except Exception:
                continue
            dipakai.add(terbaik[0])
            hasil.append(data)
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


def buat_gambar(spec, upload=None, pustaka=None):
    """Spec gambar -> bytes PNG, atau None bila tidak ada / gagal (soal tetap jalan tanpa gambar)."""
    if upload:
        return upload
    if not spec or not isinstance(spec, dict):
        return None
    try:
        jenis = spec.get("jenis")
        if jenis == "pustaka":
            ketemu = cari_gambar_pustaka(spec.get("objek", []), pustaka)
            return gabung_horizontal(ketemu) if ketemu else None
        fungsi = GAMBAR_KODE.get(jenis)
        return fungsi(**(spec.get("param") or {})) if fungsi else None
    except Exception:
        return None


def gambar_untuk(s, pustaka=None):
    """Gambar sebuah soal (dengan cache sederhana di dalam dict soal)."""
    if s.get("_hapus_gambar"):
        return None
    if s.get("_upload"):
        return s["_upload"]
    kunci = json.dumps(s.get("gambar"), sort_keys=True, default=str) + f"|v{getattr(pustaka, 'versi', 0)}"
    if s.get("_ck") != kunci:
        s["_cb"] = buat_gambar(s.get("gambar"), pustaka=pustaka)
        s["_ck"] = kunci
    return s.get("_cb")


# ---------- Pustaka gambar di Google Sheet (disimpan sebagai teks base64, 1 gambar = beberapa baris) ----------
UKURAN_BAGIAN = 45000    # 1 sel Google Sheet maksimal 50.000 karakter
MAKS_KARAKTER = 135000   # maksimal 3 bagian per gambar


def kompres_gambar(raw, maks_sisi=800):
    """Kecilkan gambar (JPEG, latar putih) sampai muat di sel Google Sheet."""
    im = Image.open(io.BytesIO(raw))
    if im.mode in ("RGBA", "LA", "P"):
        im = im.convert("RGBA")
        latar = Image.new("RGB", im.size, "white")
        latar.paste(im, mask=im.split()[-1])
        im = latar
    else:
        im = im.convert("RGB")
    sisi = maks_sisi
    for _ in range(6):
        c = im.copy()
        c.thumbnail((sisi, sisi))
        for q in (85, 75, 65, 55, 45):
            buf = io.BytesIO()
            c.save(buf, "JPEG", quality=q, optimize=True)
            if len(buf.getvalue()) * 4 // 3 + 4 <= MAKS_KARAKTER:
                return buf.getvalue()
        sisi = int(sisi * 0.8)
    raise ValueError("Gambar terlalu rumit untuk dikompres. Gunakan gambar yang lebih sederhana.")


def _pesan_gas(m):
    m = str(m)
    if "Aksi tidak dikenal" in m:
        return ("Apps Script di Google Sheet Anda masih versi lama. Tempel kode terbaru dari panel "
                "'Panduan Google Sheet', lalu Deploy > Kelola deployment > Versi baru.")
    return m


def sheet_daftar_gambar(gas):
    ok, j = gas("daftar_gambar")
    if not ok:
        raise RuntimeError(_pesan_gas(j))
    return list(j.get("gambar", []))


def sheet_simpan_gambar(gas, nama, kata_kunci, raw):
    nama = (nama or "").strip()
    if not nama:
        raise ValueError("Nama gambar harus diisi.")
    jpg = kompres_gambar(raw)
    b64 = base64.b64encode(jpg).decode("ascii")
    potong = [b64[i:i + UKURAN_BAGIAN] for i in range(0, len(b64), UKURAN_BAGIAN)]
    gid = "g" + uuid.uuid4().hex[:10]
    kk = (kata_kunci or "").strip()
    ok, j = gas("simpan_gambar", {"id": gid, "nama": nama, "kata_kunci": kk, "potong": potong})
    if not ok:
        raise RuntimeError(_pesan_gas(j))
    return {"id": gid, "nama": nama, "kata_kunci": kk}, jpg


def sheet_ambil_gambar(gas, gid, cache):
    if gid in cache:
        return cache[gid]
    ok, j = gas("muat_gambar", {"id": gid})
    if not ok:
        raise RuntimeError(_pesan_gas(j))
    cache[gid] = base64.b64decode(j["data"])
    return cache[gid]


def sheet_hapus_gambar(gas, gid):
    ok, j = gas("hapus_gambar", {"id": gid})
    if not ok:
        raise RuntimeError(_pesan_gas(j))

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


def _verifikasi_pg(s, perbaiki=True):
    """Periksa struktur + kunci. Mengisi s['_status'] dan s['_catatan'].
    perbaiki=True: kunci yang salah diganti otomatis; False: hanya diberi peringatan (dipakai saat guru mengedit)."""
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
    elif len(cocok) == 1 and not perbaiki:
        s["_status"] = "gagal"
        s["_catatan"] = (f"Kunci {s['kunci']} tidak cocok dengan hasil hitung ({_n(hasil)}). "
                         f"Pilihan yang cocok: {cocok[0]}.")
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


NAMA_TIPE = {"PG": "Pilihan Ganda", "ISIAN": "Isian Singkat", "ESAI": "Esai (Uraian)"}
LABEL_TIPE = {"PG": "PG", "ISIAN": "ISIAN", "ESAI": "ESAI"}
URUT_TIPE = ("PG", "ISIAN", "ESAI")
BATCH_TIPE = {"PG": 5, "ISIAN": 5, "ESAI": 3}  # soal per panggilan AI
BAGIAN = {
    "PG": ("PILIHAN GANDA", "Pilihlah jawaban yang paling tepat dengan memberi tanda silang (X) pada huruf A, B, C, atau D!"),
    "ISIAN": ("ISIAN SINGKAT", "Isilah titik-titik berikut dengan jawaban yang benar!"),
    "ESAI": ("ESAI", "Jawablah pertanyaan berikut dengan jelas dan benar!"),
}


def _angka_semua(teks):
    return [float(x.replace(".", "").replace(",", ".")) for x in _RE_ANGKA.findall(str(teks))]


def _verifikasi_isian(s):
    jw = str(s.get("jawaban", "")).strip()
    if not jw:
        s["_status"], s["_catatan"] = "gagal", "Kunci jawaban isian belum diisi."
        return s
    ek = s.get("ekspresi_kunci")
    if not ek:
        s["_status"], s["_catatan"] = "tanpa_cek", "Isian bukan hitungan: periksa kunci secara manual."
        return s
    try:
        hasil = hitung_aman(ek)
    except Exception:
        s["_status"], s["_catatan"] = "tanpa_cek", "Ekspresi hitung tidak bisa dibaca: periksa manual."
        return s
    ang = parse_angka(jw)
    if ang is None:
        s["_status"], s["_catatan"] = "tanpa_cek", "Jawaban bukan angka: periksa manual."
    elif math.isclose(ang, hasil, rel_tol=1e-6, abs_tol=1e-9):
        s["_status"], s["_catatan"] = "ok", f"Terverifikasi: hasil hitung = {_n(hasil)}."
    else:
        s["_status"] = "gagal"
        s["_catatan"] = f"Hasil hitung = {_n(hasil)} tetapi kunci tertulis '{jw}'. Perbaiki kunci jawaban."
    return s


def _verifikasi_esai(s):
    if not str(s.get("jawaban", "")).strip() or not str(s.get("pedoman", "")).strip():
        s["_status"], s["_catatan"] = "gagal", "Kunci jawaban dan pedoman penskoran esai harus diisi."
        return s
    ek = s.get("ekspresi_kunci")
    if ek:
        try:
            hasil = hitung_aman(ek)
            if any(math.isclose(a, hasil, rel_tol=1e-6, abs_tol=1e-9) for a in _angka_semua(s["jawaban"])):
                s["_status"], s["_catatan"] = "ok", f"Terverifikasi: hasil akhir {_n(hasil)} ada pada kunci jawaban."
            else:
                s["_status"] = "gagal"
                s["_catatan"] = f"Hasil hitung = {_n(hasil)} tidak ditemukan pada kunci jawaban. Periksa kunci."
            return s
        except Exception:
            pass
    s["_status"], s["_catatan"] = "tanpa_cek", "Esai: periksa kunci jawaban dan pedoman penskoran secara manual."
    return s


def verifikasi(s, perbaiki=True):
    tipe = s.get("tipe", "PG")
    if tipe == "ISIAN":
        return _verifikasi_isian(s)
    if tipe == "ESAI":
        return _verifikasi_esai(s)
    return _verifikasi_pg(s, perbaiki)


def normalisasi(s, tipe=None):
    tipe = str(tipe or s.get("tipe") or "PG").upper()
    if tipe not in LABEL_TIPE:
        tipe = "PG"
    out = {
        "tipe": tipe,
        "stimulus": str(s.get("stimulus") or "").strip(),
        "pertanyaan": str(s.get("pertanyaan") or "").strip(),
        "pembahasan": str(s.get("pembahasan") or "").strip(),
        "ekspresi_kunci": s.get("ekspresi_kunci") or None,
        "gambar": s.get("gambar") if isinstance(s.get("gambar"), dict) else None,
    }
    if tipe == "PG":
        opsi = s.get("opsi") or {}
        out["opsi"] = {h: str(opsi.get(h, "")).strip() for h in "ABCD"}
        out["kunci"] = str(s.get("kunci") or "").strip().upper()[:1]
    else:
        out["jawaban"] = str(s.get("jawaban") or "").strip()
        if tipe == "ESAI":
            out["pedoman"] = str(s.get("pedoman") or "").strip()
            try:
                out["skor"] = max(1, min(20, int(float(s.get("skor") or 4))))
            except (TypeError, ValueError):
                out["skor"] = 4
    return verifikasi(out)

# =====================================================================
# 4. GENERATE SOAL DENGAN AI (Groq)
# =====================================================================

def _mapel_matematika(mapel):
    return "matematika" in str(mapel).lower()


PETUNJUK_MAPEL = [
    ("ipas", "Soal IPAS menguji PEMAHAMAN KONSEP: fenomena alam/sosial, sebab-akibat, ciri dan klasifikasi, "
             "pengamatan/percobaan sederhana, kegiatan manusia dan lingkungan. Stimulus berupa teks pengamatan, "
             "cerita fenomena, atau tabel/data singkat. BUKAN soal hitungan."),
    ("indonesia", "Soal Bahasa Indonesia berbasis teks bacaan/stimulus pendek: ide pokok, isi teks, kosakata, "
                  "ejaan dan tanda baca, jenis kalimat/paragraf. BUKAN soal hitungan."),
    ("inggris", "Soal Bahasa Inggris sesuai level SD: vocabulary, short text/dialogue, simple grammar. "
                "BUKAN soal hitungan."),
    ("agama", "Soal sesuai materi agama: pemahaman nilai, kisah, tata cara ibadah, akhlak. BUKAN soal hitungan."),
    ("pancasila", "Soal Pendidikan Pancasila: nilai Pancasila, norma, hak dan kewajiban, kasus sederhana "
                  "kehidupan sehari-hari. BUKAN soal hitungan."),
]


def _petunjuk_mapel(mapel):
    m = str(mapel).lower()
    for kunci, teks in PETUNJUK_MAPEL:
        if kunci in m:
            return teks
    return "Soal harus murni materi mata pelajaran ini dan BUKAN soal hitungan Matematika."


_BENTUK = {
    "PG": '{"soal":[{"stimulus":"...","pertanyaan":"...","opsi":{"A":"...","B":"...","C":"...","D":"..."},'
          '"kunci":"B","pembahasan":"...","ekspresi_kunci":null,"gambar":null}]}',
    "ISIAN": '{"soal":[{"stimulus":"...","pertanyaan":"... ..... ...","jawaban":"...","pembahasan":"...",'
             '"ekspresi_kunci":null,"gambar":null}]}',
    "ESAI": '{"soal":[{"stimulus":"...","pertanyaan":"...","jawaban":"...","pedoman":"...","skor":4,'
            '"ekspresi_kunci":null,"gambar":null}]}',
}


def buat_prompt(p, tipe="PG"):
    mat = _mapel_matematika(p["mapel"])
    kepala = (
        f"Buatkan soal {NAMA_TIPE[tipe].lower()} untuk siswa SD sesuai Kurikulum Merdeka.\n"
        f"- Jenis ujian : {p['jenis_ujian']}\n- Jenis soal  : {NAMA_TIPE[tipe]}\n- Kelas       : {p['kelas']}\n"
        f"- Mapel       : {p['mapel']}\n- Materi/Bab  : {p['materi']}\n"
        f"- Jumlah soal : {p['jumlah']}\n- Tingkat     : {p['tingkat']}\n\n"
        f"WAJIB: SEMUA soal harus membahas mata pelajaran {p['mapel']} untuk {p['kelas']} dengan materi: "
        f"{p['materi']}. DILARANG membuat soal mata pelajaran lain"
        + ("" if mat else " (terutama soal hitungan Matematika seperti harga, keliling, luas, kecepatan)")
        + ".\n"
    )
    if not mat:
        kepala += "Panduan mapel: " + _petunjuk_mapel(p["mapel"]) + "\n"
    angka_id = " Tulis angka gaya Indonesia (titik ribuan, koma desimal)." if mat else ""
    aturan = ["Bahasa Indonesia baku, panjang kalimat sesuai kelas (kelas 1-2: kalimat pendek dan sederhana)."]
    if tipe == "PG":
        if mat:
            aturan += ["'stimulus' = cerita/data singkat yang memuat SEMUA angka dan satuan yang dibutuhkan. "
                       "'pertanyaan' = satu kalimat tanya.",
                       "Empat pilihan A-D, hanya satu benar. Pengecoh berupa kesalahan umum siswa." + angka_id +
                       " Pilihan angka diurutkan naik, panjang pilihan seimbang."]
        else:
            aturan += ["'stimulus' = teks/pengamatan/data singkat yang memuat informasi yang dibutuhkan untuk menjawab. "
                       "'pertanyaan' = satu kalimat tanya.",
                       "Empat pilihan A-D, hanya satu benar. Pengecoh berupa miskonsepsi yang masuk akal, "
                       "panjang pilihan seimbang."]
        aturan.append("'pembahasan' = " + ("langkah penyelesaian singkat." if mat else "alasan jawaban benar, singkat."))
    elif tipe == "ISIAN":
        aturan += ["'stimulus' = " + ("cerita/data singkat yang memuat SEMUA angka dan satuan yang dibutuhkan."
                                      if mat else "teks/pengamatan/data singkat yang memuat informasi yang dibutuhkan."),
                   "'pertanyaan' memuat titik-titik (.....) sebagai tempat jawaban, atau kalimat tanya yang "
                   "jawabannya singkat. Jangan membuat pilihan A-D.",
                   "'jawaban' = jawaban singkat saja (satu kata, angka, atau frasa pendek, maksimal 4 kata). "
                   "Hanya SATU jawaban yang benar dan tidak ambigu." + angka_id,
                   "'pembahasan' = alasan atau langkah singkat."]
    else:
        aturan += ["'stimulus' = " + ("cerita/data singkat yang memuat SEMUA angka dan satuan yang dibutuhkan."
                                      if mat else "teks/pengamatan/data singkat yang memuat informasi yang dibutuhkan."),
                   "'pertanyaan' menuntut penjelasan atau penalaran (mengapa, jelaskan, bagaimana)"
                   + (" atau penyelesaian bertahap (hitunglah dengan langkah-langkahnya)" if mat else "")
                   + ". Kelas 1-2: cukup jawaban 1-2 kalimat.",
                   "'jawaban' = kunci jawaban lengkap." + angka_id,
                   "'pedoman' = pedoman penskoran rinci, skor per komponen, contoh: "
                   "'Menyebutkan penyebab (1); memberi alasan (2); kesimpulan benar (1)'. "
                   "Jumlah skor komponen HARUS sama dengan 'skor'.",
                   "'skor' = bilangan bulat 2 sampai 6 sesuai bobot soal."]
    if mat:
        aturan.append("'ekspresi_kunci': untuk soal hitungan, tulis ekspresi Python (hanya angka dengan TITIK desimal, "
                      "operator + - * / dan kurung) yang hasilnya sama dengan nilai jawaban akhir "
                      + {"PG": "pada pilihan yang benar", "ISIAN": "pada 'jawaban'",
                         "ESAI": "yang tertulis pada 'jawaban'"}[tipe]
                      + ", dalam satuan yang sama. Untuk soal bukan hitungan isi null.")
        aturan.append("Gambar dari kode (bangun datar, diagram, jam, dst.) HANYA dipakai bila materinya memang "
                      "membutuhkannya. " + GAMBAR_DOKUMEN)
    else:
        aturan.append("'ekspresi_kunci' WAJIB null (ini bukan soal hitungan).")
        aturan.append("Gambar: pakai {\"jenis\":\"pustaka\",\"objek\":[...]} untuk benda/hewan/tumbuhan/tempat bila "
                      "membantu, {\"jenis\":\"diagram_batang\",...} bila ada data tabel/grafik, selain itu null. "
                      "Jangan memaksakan gambar. Jangan menulis teks apa pun di dalam gambar.")
    teks_aturan = "".join(f"{i}. {a}\n" for i, a in enumerate(aturan, 1))
    return (kepala + "\nATURAN:\n" + teks_aturan + "\nBalas HANYA JSON valid dengan bentuk:\n" + _BENTUK[tipe])


_RE_HITUNG = re.compile(r"keliling|luas (persegi|segitiga|lingkaran|trapesium|bangun|tanah|lahan)|kecepatan rata|"
                        r"kembalian|diskon|potongan harga|berapa total|berapa biaya|berapa harga", re.I)
_RE_OPSI_ANGKA = re.compile(r"^\s*(rp\s*)?[\d\.,]+\s*[a-z/%\u00b2\u00b3 ]{0,12}$", re.I)


def kemungkinan_salah_mapel(s, mapel):
    """True bila soal non-Matematika tampak seperti soal hitungan Matematika."""
    if _mapel_matematika(mapel):
        return False
    if s.get("ekspresi_kunci"):
        return True
    tipe = s.get("tipe", "PG")
    if tipe == "PG":
        semua_angka = all(_RE_OPSI_ANGKA.match(str(v)) for v in s.get("opsi", {}).values())
    elif tipe == "ISIAN":
        semua_angka = bool(_RE_OPSI_ANGKA.match(str(s.get("jawaban", ""))))
    else:
        return False
    teks = f"{s.get('stimulus', '')} {s.get('pertanyaan', '')}"
    return bool(semua_angka and _RE_HITUNG.search(teks))


def _ambil_json(teks):
    try:
        return json.loads(teks)
    except Exception:
        m = re.search(r"\{.*\}", teks, re.S)
        if not m:
            raise ValueError("Jawaban AI bukan JSON.")
        return json.loads(m.group(0))


BATCH_SOAL = 5  # soal per panggilan: menjaga jawaban tidak terpotong dan hemat token per menit


def _panggil_groq(client, model, pesan, max_token=5000):
    """Satu panggilan Groq. Coba lagi bila kena batas token/menit atau format JSON ditolak model."""
    import time
    pakai_json = True
    coba_max = max_token
    err = None
    for percobaan in range(4):
        try:
            kw = dict(model=model, messages=pesan, temperature=0.4, max_completion_tokens=coba_max)
            if pakai_json:
                kw["response_format"] = {"type": "json_object"}
            resp = client.chat.completions.create(**kw)
            isi = resp.choices[0].message.content or ""
            if not isi.strip():
                raise ValueError("Jawaban AI kosong (kemungkinan token habis).")
            return isi
        except Exception as e:
            err = e
            teks = str(e)
            if "response_format" in teks or "json_validate_failed" in teks or "json_object" in teks:
                pakai_json = False
                continue
            if "rate_limit" in teks or "tokens per minute" in teks or "Request too large" in teks:
                coba_max = max(1500, coba_max // 2)
                time.sleep(12)
                continue
            raise
    raise RuntimeError(f"Groq menolak permintaan berulang kali. Detail: {err}")


def _minta_soal(client, model, p, n, sudah, tipe="PG", tegas=False):
    pp = dict(p, jumlah=n)
    prompt = buat_prompt(pp, tipe)
    if sudah:
        prompt += ("\n\nSoal yang SUDAH dibuat (jangan diulang; gunakan sub-topik atau konteks berbeda, "
                   f"tetap pada mapel {p['mapel']}):\n- " + "\n- ".join(h["pertanyaan"][:90] for h in sudah))
    if tegas:
        prompt += (f"\n\nPERHATIAN: percobaan sebelumnya salah mapel (berupa soal hitungan Matematika). "
                   f"Kali ini SEMUA soal harus murni {p['mapel']}, materi: {p['materi']}. ekspresi_kunci harus null.")
    pesan = [
        {"role": "system", "content": "Kamu guru SD berpengalaman dan penyusun soal asesmen. "
                                      "Keluarkan hanya JSON valid, tanpa teks lain."},
        {"role": "user", "content": prompt},
    ]
    data = _ambil_json(_panggil_groq(client, model, pesan))
    return [normalisasi(x, tipe) for x in data.get("soal", [])][:n]


def generate_soal_ai(client, p, model=MODEL_DEFAULT):
    """Membuat soal per jenis (PG, isian, esai) dan per batch, memeriksa kesesuaian mapel, lalu menggabungkan.
    p['jumlah_pg'], p['jumlah_isian'], p['jumlah_esai'] = jumlah tiap jenis (p['jumlah'] dianggap PG bila ketiganya kosong)."""
    komposisi = {"PG": int(p.get("jumlah_pg", p.get("jumlah", 0)) or 0),
                 "ISIAN": int(p.get("jumlah_isian", 0) or 0),
                 "ESAI": int(p.get("jumlah_esai", 0) or 0)}
    hasil = []
    for tipe in URUT_TIPE:
        total, selesai = komposisi[tipe], 0
        while selesai < total:
            n = min(BATCH_TIPE[tipe], total - selesai)
            baru = _minta_soal(client, model, p, n, hasil, tipe)
            if not baru:
                raise ValueError(f"AI tidak menghasilkan soal {NAMA_TIPE[tipe]}. Coba lagi atau persingkat materi.")
            # soal non-Matematika yang tampak seperti hitungan Matematika: minta diganti (maks 2 kali)
            for _ in range(2):
                salah = [i for i, x in enumerate(baru) if kemungkinan_salah_mapel(x, p["mapel"])]
                if not salah:
                    break
                ganti_ = _minta_soal(client, model, p, len(salah), hasil + baru, tipe, tegas=True)
                for i, g in zip(salah, ganti_):
                    if not kemungkinan_salah_mapel(g, p["mapel"]):
                        baru[i] = g
            for x in baru:
                if kemungkinan_salah_mapel(x, p["mapel"]):
                    x["_salah_mapel"] = True
            hasil.extend(baru)
            selesai += len(baru)
    if not hasil:
        raise ValueError("Jumlah soal masih 0. Isi jumlah minimal satu jenis soal.")
    return hasil


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
    {"tipe": "ISIAN",
     "stimulus": "Sebuah taman berbentuk persegi memiliki panjang sisi 12 m.",
     "pertanyaan": "Keliling taman tersebut adalah ..... m.",
     "jawaban": "48", "pembahasan": "Keliling persegi = 4 x sisi = 4 x 12 = 48 m.",
     "ekspresi_kunci": "4*12",
     "gambar": {"jenis": "persegi", "param": {"sisi": 12, "satuan": "m"}}},
    {"tipe": "ESAI",
     "stimulus": "Ani membeli 3 pensil seharga Rp2.500 per buah dan 2 buku tulis seharga Rp6.000 per buah. "
                 "Ani membayar dengan uang Rp50.000.",
     "pertanyaan": "Hitunglah uang kembalian Ani! Tuliskan langkah pengerjaannya.",
     "jawaban": "Harga pensil = 3 x 2.500 = Rp7.500. Harga buku = 2 x 6.000 = Rp12.000. "
                "Total belanja = Rp19.500. Kembalian = 50.000 - 19.500 = Rp30.500.",
     "pedoman": "Menghitung harga pensil (1); menghitung harga buku (1); menjumlahkan total belanja (1); "
                "menghitung kembalian dengan jawaban akhir Rp30.500 (1)",
     "skor": 4, "ekspresi_kunci": "50000-(3*2500+2*6000)", "gambar": None},
]

# =====================================================================
# 5. EKSPOR PDF  (reportlab)
# =====================================================================

def _esc(t):
    t = str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return re.sub(r"\^(\d+)", r"<super>\1</super>", t)


def _baris_kunci(s):
    """(jenis, kunci/jawaban, pembahasan/pedoman) untuk halaman kunci."""
    t = s.get("tipe", "PG")
    if t == "PG":
        return "PG", s.get("kunci", ""), s.get("pembahasan") or "-"
    if t == "ISIAN":
        return "Isian", s.get("jawaban") or "-", s.get("pembahasan") or "-"
    return "Esai", s.get("jawaban") or "-", f"{s.get('pedoman') or '-'} (Skor maksimal: {s.get('skor', 4)})"


def buat_pdf(soal, kop, sertakan_kunci=True, pustaka=None):
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
    st_bagian = ParagraphStyle("bg", parent=base, fontName="Helvetica-Bold", fontSize=11, spaceBefore=4,
                               keepWithNext=1)
    st_petunjuk = ParagraphStyle("pt", parent=base, fontName="Helvetica-Oblique", fontSize=9.5, spaceAfter=8,
                                 keepWithNext=1)
    jenis_ada = {x.get("tipe", "PG") for x in soal}
    tipe_lalu, urut_bagian = None, 0
    for no, s in enumerate(soal, 1):
        tipe = s.get("tipe", "PG")
        if len(jenis_ada) > 1 and tipe != tipe_lalu:
            urut_bagian += 1
            judul_b, petunjuk = BAGIAN[tipe]
            cerita += [Paragraph(f"{chr(64 + urut_bagian)}.&nbsp;&nbsp;{judul_b}", st_bagian),
                       Paragraph(_esc(petunjuk), st_petunjuk)]
        tipe_lalu = tipe
        kiri = [Paragraph(f"{no}.&nbsp;&nbsp;SOAL {no} \u2014 {LABEL_TIPE[tipe]}", st_judul)]
        if s["stimulus"]:
            kiri += [Paragraph("Stimulus:", st_label), Paragraph(_esc(s["stimulus"]), base)]
        kiri += [Paragraph("Pertanyaan:", st_label), Paragraph(_esc(s["pertanyaan"]), base), Spacer(1, 2)]
        if tipe == "PG":
            for h in "ABCD":
                kiri.append(Paragraph(f"{h}.&nbsp;&nbsp;&nbsp;{_esc(s['opsi'][h])}", st_opsi))
        elif tipe == "ISIAN":
            kiri.append(Paragraph("Jawaban: " + "." * 38, st_opsi))
        else:
            kiri.append(Paragraph(f"<i>(Skor maksimal: {s.get('skor', 4)})</i>", st_opsi))
            for _ in range(4):
                kiri.append(HRFlowable(width="100%", thickness=0.6, dash=[1, 2], color=colors.black,
                                       spaceBefore=12, spaceAfter=0))
        gb = gambar_untuk(s, pustaka)
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
        baris = [[Paragraph("<b>No</b>", base), Paragraph("<b>Jenis</b>", base),
                  Paragraph("<b>Kunci / Jawaban</b>", base), Paragraph("<b>Pembahasan / Pedoman Penskoran</b>", base)]]
        for no, s in enumerate(soal, 1):
            jenis_k, kunci_k, bahas_k = _baris_kunci(s)
            baris.append([Paragraph(str(no), base), Paragraph(jenis_k, base),
                          Paragraph(_esc(kunci_k), base), Paragraph(_esc(bahas_k), base)])
        tk = Table(baris, colWidths=[1.1 * cm, 1.6 * cm, 5.2 * cm, LEBAR - 7.9 * cm], repeatRows=1)
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


def _baris_jawaban(par, lebar_cm):
    """Baris jawaban berupa titik-titik sepanjang lebar_cm (titik Arial 10,5 pt = sekitar 0,103 cm)."""
    _teks_docx(par, "." * max(10, int(lebar_cm * 9.0)))


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


def buat_docx(soal, kop, sertakan_kunci=True, pustaka=None):
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
    jenis_ada = {x.get("tipe", "PG") for x in soal}
    tipe_lalu, urut_bagian = None, 0
    for no, s in enumerate(soal, 1):
        tipe = s.get("tipe", "PG")
        if len(jenis_ada) > 1 and tipe != tipe_lalu:
            urut_bagian += 1
            judul_b, petunjuk = BAGIAN[tipe]
            ph = doc.add_paragraph()
            ph.paragraph_format.keep_with_next = True
            _teks_docx(ph, f"{chr(64 + urut_bagian)}.  {judul_b}", True, 11)
            pp_ = doc.add_paragraph()
            pp_.paragraph_format.keep_with_next = True
            rr = pp_.add_run(petunjuk)
            rr.italic = True
            rr.font.size = Pt(9.5)
        tipe_lalu = tipe
        gb = gambar_untuk(s, pustaka)
        lebar_kiri = LEBAR - GW - 0.3 if gb else LEBAR
        tb = doc.add_table(rows=1, cols=2 if gb else 1)
        _tanpa_pecah(tb.rows[0])
        kiri = tb.cell(0, 0)
        kiri.width = Cm(lebar_kiri)
        _teks_docx(kiri.paragraphs[0], f"{no}.  SOAL {no} \u2014 {LABEL_TIPE[tipe]}", True, 11)
        if s["stimulus"]:
            _teks_docx(kiri.add_paragraph(), "Stimulus:", True)
            _teks_docx(kiri.add_paragraph(), s["stimulus"])
        _teks_docx(kiri.add_paragraph(), "Pertanyaan:", True)
        _teks_docx(kiri.add_paragraph(), s["pertanyaan"])
        if tipe == "PG":
            for h in "ABCD":
                po = kiri.add_paragraph()
                po.paragraph_format.left_indent = Cm(0.5)
                _teks_docx(po, f"{h}.   {s['opsi'][h]}")
        elif tipe == "ISIAN":
            po = kiri.add_paragraph()
            po.paragraph_format.left_indent = Cm(0.5)
            _teks_docx(po, "Jawaban: " + "." * 38)
        else:
            po = kiri.add_paragraph()
            ri = po.add_run(f"(Skor maksimal: {s.get('skor', 4)})")
            ri.italic = True
            ri.font.size = Pt(9.5)
            for _ in range(4):
                pl = kiri.add_paragraph()
                pl.paragraph_format.space_before = Pt(10)
                _baris_jawaban(pl, lebar_kiri - 0.5)
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
        lk = (1.1, 1.6, 5.2, LEBAR - 7.9)
        tk = doc.add_table(rows=1, cols=4)
        tk.style = "Table Grid"
        for c, teks, w in zip(tk.rows[0].cells, ("No", "Jenis", "Kunci / Jawaban", "Pembahasan / Pedoman Penskoran"), lk):
            c.width = Cm(w)
            _teks_docx(c.paragraphs[0], teks, True)
        for no, s in enumerate(soal, 1):
            r = tk.add_row().cells
            jenis_k, kunci_k, bahas_k = _baris_kunci(s)
            for c, teks, w in zip(r, (str(no), jenis_k, kunci_k, bahas_k), lk):
                c.width = Cm(w)
                _teks_docx(c.paragraphs[0], teks)
        _atur_lebar(tk, lk)

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


def render_tab_soal(api_key=None, sekolah="", tahun_ajaran="", daftar_mapel=None, daftar_kelas=None, model=None, gas_kirim=None, sumber_awal=None):
    """api_key/sekolah/tahun_ajaran/daftar_mapel/daftar_kelas/model/gas_kirim boleh diisi dari app induk (app KKG).
    gas_kirim(aksi, payload) -> (ok, respons) : fungsi pengirim ke Apps Script milik guru.
    sumber_awal() -> dict(mapel, kelas, materi) : isian tab lain yang bisa disalin ke sini."""
    import streamlit as st
    ss = st.session_state
    ss.setdefault("sg_pg_meta", [])
    ss.setdefault("sg_pg_cache", {})
    ss.setdefault("sg_pg_versi", 0)
    daftar_mapel = list(daftar_mapel or MAPEL)
    if "Lainnya" not in daftar_mapel:
        daftar_mapel.append("Lainnya")
    daftar_kelas = list(daftar_kelas or [f"Kelas {i}" for i in range(1, 7)])

    st.subheader("\U0001F4DD Generator Soal (UH / PTS / PAS)")
    st.session_state.setdefault("sg_soal_list", [])
    st.session_state.setdefault("sg_soal_uid", 0)

    with st.expander("\U0001F4DA Pustaka gambar saya (Google Sheet)"):
        if not gas_kirim:
            st.caption("Fitur ini membutuhkan sambungan Google Sheet dari app KKG (URL Apps Script di sidebar).")
        else:
            st.caption("Gambar yang Anda unggah disimpan di Google Sheet milik Anda sendiri dan bisa dipakai sebagai "
                       "gambar stimulus. Bila kata kuncinya cocok dengan soal, gambar dipilih otomatis.")
            if st.button("\U0001F504 Muat daftar gambar dari Google Sheet", key="sg_pg_muat"):
                try:
                    with st.spinner("Mengambil daftar gambar..."):
                        ss["sg_pg_meta"][:] = sheet_daftar_gambar(gas_kirim)
                    ss["sg_pg_versi"] += 1
                    st.success(f"{len(ss['sg_pg_meta'])} gambar ditemukan.")
                except Exception as e:
                    st.error(f"Gagal memuat daftar: {e}")
            for item in list(ss["sg_pg_meta"]):
                ca, cb = st.columns([4, 1])
                ca.markdown(f"**{item.get('nama', '')}**  \n<small>{item.get('kata_kunci', '')}</small>",
                            unsafe_allow_html=True)
                if cb.button("\U0001F5D1\uFE0F Hapus", key=f"sg_pg_del_{item['id']}"):
                    try:
                        sheet_hapus_gambar(gas_kirim, item["id"])
                        ss["sg_pg_meta"][:] = [m for m in ss["sg_pg_meta"] if m["id"] != item["id"]]
                        ss["sg_pg_cache"].pop(item["id"], None)
                        ss["sg_pg_versi"] += 1
                        st.rerun()
                    except Exception as e:
                        st.error(f"Gagal menghapus: {e}")
            st.markdown("**Tambah gambar baru**")
            f_baru = st.file_uploader("Pilih gambar (PNG/JPG)", type=["png", "jpg", "jpeg", "webp"], key="sg_pg_upl")
            n_baru = st.text_input("Nama gambar", key="sg_pg_nama", placeholder="Contoh: Bus kota")
            k_baru = st.text_input("Kata kunci (pisahkan dengan koma)", key="sg_pg_kk",
                                   placeholder="bus, angkutan, terminal")
            if st.button("\U0001F4BE Simpan ke Google Sheet", key="sg_pg_simpan"):
                if f_baru is None:
                    st.warning("Pilih gambar dulu.")
                else:
                    try:
                        with st.spinner("Mengompres dan menyimpan ke Google Sheet..."):
                            item, jpg = sheet_simpan_gambar(gas_kirim, n_baru, k_baru, f_baru.getvalue())
                        ss["sg_pg_cache"][item["id"]] = jpg
                        ss["sg_pg_meta"].append(item)
                        ss["sg_pg_versi"] += 1
                        st.success(f"Gambar '{item['nama']}' tersimpan di Google Sheet.")
                    except Exception as e:
                        st.error(f"Gagal menyimpan: {e}")
    pustaka = None
    if gas_kirim:
        pustaka = PustakaSheet(ss["sg_pg_meta"],
                               lambda gid: sheet_ambil_gambar(gas_kirim, gid, ss["sg_pg_cache"]),
                               ss["sg_pg_versi"])

    if sumber_awal:
        def _salin():
            v = sumber_awal() or {}
            if v.get("mapel") in daftar_mapel:
                ss["sj_mapel"] = v["mapel"]
            if v.get("kelas") in daftar_kelas:
                ss["sj_kelas"] = v["kelas"]
            if str(v.get("materi") or "").strip():
                ss["sj_materi"] = v["materi"]
        st.button("\u21AA\uFE0F Ambil mapel, kelas & materi dari tab Kisi-kisi & Soal", key="sg_salin", on_click=_salin)

    c1, c2, c3 = st.columns(3)
    jenis = c1.selectbox("Jenis ujian", JENIS_UJIAN, key="sj_jenis")
    kelas = c2.selectbox("Kelas", daftar_kelas, index=min(3, len(daftar_kelas) - 1), key="sj_kelas")
    mapel = c3.selectbox("Mata pelajaran", daftar_mapel, key="sj_mapel")
    if mapel == "Lainnya":
        mapel = st.text_input("Nama mata pelajaran", key="sj_mapel_lain")
    materi = st.text_area("Materi / Bab (boleh lebih dari satu)", height=80, key="sj_materi",
                          placeholder="Contoh: Kecepatan, jarak, dan waktu; Keliling bangun datar; Perbandingan")
    st.markdown("**Jenis dan jumlah soal**")
    n1, n2, n3, n4 = st.columns(4)
    jml_pg = int(n1.number_input("Pilihan ganda", 0, 40, 5, key="sj_jml_pg"))
    jml_isian = int(n2.number_input("Isian singkat", 0, 20, 0, key="sj_jml_isian"))
    jml_esai = int(n3.number_input("Esai", 0, 10, 0, key="sj_jml_esai"))
    tingkat = n4.selectbox("Tingkat kesulitan", ["Campuran", "Mudah", "Sedang", "HOTS"], key="sj_tingkat")
    total_soal = jml_pg + jml_isian + jml_esai
    st.caption(f"Total {total_soal} soal: {jml_pg} pilihan ganda, {jml_isian} isian singkat, {jml_esai} esai.")

    with st.expander("Kop soal & pengaturan"):
        k1, k2 = st.columns(2)
        if not sekolah:
            sekolah = k1.text_input("Nama sekolah", key="sj_sekolah")
        else:
            k1.caption(f"Sekolah: {sekolah} (dari pengaturan sidebar)")
        if not tahun_ajaran:
            tahun_ajaran = k2.text_input("Tahun ajaran", placeholder="2026/2027", key="sj_tahun")
        else:
            k2.caption(f"Tahun ajaran: {tahun_ajaran} (dari pengaturan sidebar)")
        waktu = k1.text_input("Waktu", placeholder="90 menit", key="sj_waktu")
        if not model:
            model = k2.text_input("Model Groq", MODEL_DEFAULT, key="sj_model")
        else:
            k2.caption(f"Model AI: {model} (sama dengan tab lain)")

    b1, b2 = st.columns([2, 1])
    if b1.button("\u2728 Generate soal", type="primary", use_container_width=True):
        kunci_api = api_key or _ambil_kunci_groq(st)
        if total_soal < 1:
            st.warning("Isi jumlah minimal satu jenis soal.")
        elif not materi.strip():
            st.warning("Isi materi/bab dulu.")
        elif not kunci_api:
            st.error("API Key Groq belum diisi (isi di sidebar).")
        else:
            try:
                from groq import Groq
                with st.spinner("AI sedang menyusun dan memeriksa soal..."):
                    hasil = generate_soal_ai(
                        Groq(api_key=kunci_api),
                        dict(jenis_ujian=jenis, kelas=kelas, mapel=mapel, materi=materi.strip(),
                             jumlah=total_soal, jumlah_pg=jml_pg, jumlah_isian=jml_isian,
                             jumlah_esai=jml_esai, tingkat=tingkat), model=model)
                st.session_state["sg_soal_list"] = hasil
                st.session_state["sg_soal_uid"] += 1
                st.session_state["sg_meta"] = {"mapel": mapel, "kelas": str(kelas), "jenis": jenis, "materi": materi.strip(),
                                              "rincian": f"{jml_pg} PG, {jml_isian} isian, {jml_esai} esai"}
            except Exception as e:
                st.error(f"Gagal generate: {e}")
    if b2.button("Contoh Matematika (tanpa AI)", use_container_width=True):
        st.session_state["sg_soal_list"] = [normalisasi(dict(s)) for s in CONTOH_SOAL]
        st.session_state["sg_soal_uid"] += 1
        st.session_state["sg_meta"] = {"contoh": True}

    soal = st.session_state["sg_soal_list"]
    if not soal:
        return
    uid = st.session_state["sg_soal_uid"]
    st.divider()
    meta_gen = st.session_state.get("sg_meta") or {}
    if meta_gen.get("contoh"):
        st.info("Ini hanya CONTOH soal Matematika untuk mencoba tampilan, bukan hasil AI sesuai pilihan Anda. "
                "Klik Generate soal untuk membuat soal sesuai mapel dan materi.")
    elif meta_gen:
        st.caption(f"Dibuat untuk: {meta_gen['mapel']} \u2022 {meta_gen['kelas']} \u2022 {meta_gen['jenis']} "
                   f"\u2022 {meta_gen.get('rincian', '')} \u2022 Materi: {meta_gen['materi'][:100]}")
    st.markdown(f"**{len(soal)} soal** \u2014 periksa dan edit bila perlu.")

    for i, s in enumerate(soal):
        k = f"{uid}_{i}"
        ikon = LABEL_STATUS.get(s.get("_status"), "")
        tipe = s.get("tipe", "PG")
        with st.expander(f"{ikon} Soal {i + 1} \u2014 {LABEL_TIPE.get(tipe, 'PG')}", expanded=(s.get("_status") == "gagal")):
            if s.get("_status") == "gagal":
                st.error(s["_catatan"])
            elif s.get("_catatan"):
                st.caption(s["_catatan"])
            kiri, kanan = st.columns([3, 1])
            with kiri:
                s["stimulus"] = st.text_area("Stimulus", s["stimulus"], key=f"sg_st_{k}", height=110)
                s["pertanyaan"] = st.text_input("Pertanyaan", s["pertanyaan"], key=f"sg_pt_{k}")
                if tipe == "PG":
                    for h in "ABCD":
                        s["opsi"][h] = st.text_input(f"Pilihan {h}", s["opsi"][h], key=f"sg_o{h}_{k}")
                    s["kunci"] = st.selectbox("Kunci jawaban", list("ABCD"),
                                              index="ABCD".index(s["kunci"]) if s["kunci"] in "ABCD" and s["kunci"] else 0,
                                              key=f"sg_kc_{k}")
                    s["pembahasan"] = st.text_area("Pembahasan", s["pembahasan"], key=f"sg_pb_{k}", height=80)
                elif tipe == "ISIAN":
                    s["jawaban"] = st.text_input("Kunci jawaban (isian singkat)", s.get("jawaban", ""), key=f"sg_jw_{k}")
                    s["pembahasan"] = st.text_area("Pembahasan", s["pembahasan"], key=f"sg_pb_{k}", height=80)
                else:
                    s["jawaban"] = st.text_area("Kunci jawaban", s.get("jawaban", ""), key=f"sg_jw_{k}", height=100)
                    s["pedoman"] = st.text_area("Pedoman penskoran", s.get("pedoman", ""), key=f"sg_pd_{k}", height=80)
                    s["skor"] = int(st.number_input("Skor maksimal", 1, 20, int(s.get("skor", 4)), key=f"sg_sk_{k}"))
            # ---- gambar stimulus: unggah sendiri / pilih dari pustaka Google Sheet / otomatis ----
            ga, gb2, gc = st.columns([2, 2, 1])
            up = ga.file_uploader("Unggah gambar sendiri", type=["png", "jpg", "jpeg"], key=f"sg_up_{k}")
            if up is not None:
                s["_upload"] = up.getvalue()
                s["_sumber"] = "upload"
                s.pop("_pustaka_id", None)
            elif s.get("_sumber") == "upload":  # file dikosongkan lagi
                s.pop("_upload", None)
                s.pop("_sumber", None)
            if pustaka is not None and pustaka.meta:
                nama_map = {m["id"]: m.get("nama", m["id"]) for m in pustaka.meta}
                pilih = gb2.selectbox("Atau pilih dari pustaka Google Sheet", ["-"] + list(nama_map),
                                      format_func=lambda x: "(otomatis)" if x == "-" else nama_map.get(x, x),
                                      key=f"sg_pl_{k}")
                if s.get("_sumber") != "upload":
                    if pilih != "-":
                        if s.get("_pustaka_id") != pilih:
                            try:
                                s["_upload"] = pustaka.ambil(pilih)
                                s["_sumber"] = "pustaka"
                                s["_pustaka_id"] = pilih
                            except Exception as e:
                                st.error(f"Gagal mengambil gambar: {e}")
                    elif s.get("_sumber") == "pustaka":
                        for kk_ in ("_upload", "_sumber", "_pustaka_id"):
                            s.pop(kk_, None)
            s["_hapus_gambar"] = gc.checkbox("Tanpa gambar", s.get("_hapus_gambar", False), key=f"sg_hg_{k}")
            if up is not None and pustaka is not None:
                st.markdown("**\U0001F4BE Simpan gambar ini ke pustaka Google Sheet**")
                sa, sb2, sc = st.columns([2, 2, 1])
                nm = sa.text_input("Nama gambar", key=f"sg_nm_{k}")
                kw = sb2.text_input("Kata kunci (pisahkan koma)", key=f"sg_kw_{k}")
                if sc.button("Simpan", key=f"sg_sv_{k}"):
                    try:
                        with st.spinner("Menyimpan ke Google Sheet..."):
                            item, jpg = sheet_simpan_gambar(gas_kirim, nm, kw, up.getvalue())
                        ss["sg_pg_cache"][item["id"]] = jpg
                        ss["sg_pg_meta"].append(item)
                        ss["sg_pg_versi"] += 1
                        st.success(f"Tersimpan sebagai '{item['nama']}'.")
                    except Exception as e:
                        st.error(f"Gagal menyimpan: {e}")
            with kanan:
                gb = gambar_untuk(s, pustaka)
                if gb:
                    st.image(gb, use_container_width=True)
                else:
                    st.caption("Tanpa gambar")
            # kunci diedit manual -> kunci dianggap pilihan guru; status dihitung ulang kecuali yang diperbaiki
            verifikasi(s, perbaiki=False)
            if s.get("_salah_mapel") and not _mapel_matematika(mapel):
                s["_status"] = "gagal"
                s["_catatan"] = (f"Soal ini tampak seperti soal hitungan Matematika, bukan {mapel}. "
                                 "Edit soalnya atau buang dengan menggenerate ulang.")

    st.divider()
    kop = dict(sekolah=sekolah, jenis_ujian=jenis, tahun_ajaran=tahun_ajaran, mapel=mapel,
               kelas=str(kelas), waktu=waktu)
    sertakan = st.checkbox("Sertakan halaman kunci jawaban & pembahasan", True)
    d1, d2 = st.columns(2)
    nama = re.sub(r"[^A-Za-z0-9_-]+", "_", f"soal_{mapel}_{kelas}")
    try:
        d1.download_button("\u2B07\uFE0F Unduh PDF", buat_pdf(soal, kop, sertakan, pustaka), f"{nama}.pdf",
                           "application/pdf", use_container_width=True)
        d2.download_button("\u2B07\uFE0F Unduh Word (.docx)", buat_docx(soal, kop, sertakan, pustaka), f"{nama}.docx",
                           "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           use_container_width=True)
    except Exception as e:
        st.error(f"Gagal membuat file: {e}")


if __name__ == "__main__":
    import streamlit as st
    st.set_page_config(page_title="Generator Soal", page_icon="\U0001F4DD", layout="wide")
    render_tab_soal()
