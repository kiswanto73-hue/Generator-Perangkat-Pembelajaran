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


def verifikasi(s, perbaiki=True):
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
        f"WAJIB: SEMUA soal harus membahas mata pelajaran {p['mapel']} untuk {p['kelas']} dengan materi: "
        f"{p['materi']}. DILARANG membuat soal mata pelajaran lain (misalnya soal hitungan Matematika untuk "
        "mapel non-Matematika) kecuali materi tersebut memang memuat hitungan.\n\n"
        "ATURAN:\n"
        "1. Bahasa Indonesia baku, panjang kalimat sesuai kelas (kelas 1-2: kalimat pendek dan sederhana).\n"
        "2. 'stimulus' = cerita/data singkat yang memuat SEMUA angka dan satuan yang dibutuhkan. "
        "'pertanyaan' = satu kalimat tanya.\n"
        "3. Empat pilihan A-D, hanya satu benar. Pengecoh berupa kesalahan umum siswa. Tulis angka gaya Indonesia "
        "(titik ribuan, koma desimal), pilihan angka diurutkan naik, panjang pilihan seimbang.\n"
        "4. 'ekspresi_kunci': untuk soal hitungan, tulis ekspresi Python (hanya angka dengan TITIK desimal, "
        "operator + - * / dan kurung) yang hasilnya sama dengan nilai pilihan yang benar, dalam satuan yang sama "
        "dengan pilihan tersebut. Untuk soal bukan hitungan (termasuk hampir semua soal IPAS, bahasa, agama) isi null.\n"
        "5. 'pembahasan' = langkah penyelesaian singkat.\n"
        "6. Gambar dari kode (bangun datar, diagram, jam, dst.) HANYA dipakai bila materinya memang membutuhkannya; "
        "untuk mapel lain pakai \"pustaka\" bila perlu, atau null. " + GAMBAR_DOKUMEN + "\n\n"
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


def generate_soal_ai(client, p, model=MODEL_DEFAULT):
    """Membuat soal per batch (maks BATCH_SOAL) lalu menggabungkan dan memverifikasi."""
    total = int(p["jumlah"])
    hasil = []
    while len(hasil) < total:
        n = min(BATCH_SOAL, total - len(hasil))
        pp = dict(p, jumlah=n)
        prompt = buat_prompt(pp)
        if hasil:
            prompt += ("\n\nSoal yang SUDAH dibuat (jangan diulang, gunakan konteks dan angka berbeda):\n- "
                       + "\n- ".join(h["pertanyaan"][:90] for h in hasil))
        pesan = [
            {"role": "system", "content": "Kamu guru SD berpengalaman dan penyusun soal asesmen. "
                                          "Keluarkan hanya JSON valid, tanpa teks lain."},
            {"role": "user", "content": prompt},
        ]
        data = _ambil_json(_panggil_groq(client, model, pesan))
        baru = [normalisasi(s) for s in data.get("soal", [])][:n]
        if not baru:
            raise ValueError("AI tidak menghasilkan soal. Coba lagi atau persingkat materi.")
        hasil.extend(baru)
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
]

# =====================================================================
# 5. EKSPOR PDF  (reportlab)
# =====================================================================

def _esc(t):
    t = str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return re.sub(r"\^(\d+)", r"<super>\1</super>", t)


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
    for no, s in enumerate(soal, 1):
        kiri = [Paragraph(f"{no}.&nbsp;&nbsp;SOAL {no} - PG".replace("-", "\u2014"), st_judul)]
        if s["stimulus"]:
            kiri += [Paragraph("Stimulus:", st_label), Paragraph(_esc(s["stimulus"]), base)]
        kiri += [Paragraph("Pertanyaan:", st_label), Paragraph(_esc(s["pertanyaan"]), base), Spacer(1, 2)]
        for h in "ABCD":
            kiri.append(Paragraph(f"{h}.&nbsp;&nbsp;&nbsp;{_esc(s['opsi'][h])}", st_opsi))
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
    for no, s in enumerate(soal, 1):
        gb = gambar_untuk(s, pustaka)
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
    c4, c5 = st.columns(2)
    jumlah = c4.number_input("Jumlah soal", 1, 40, 5, key="sj_jumlah")
    tingkat = c5.selectbox("Tingkat kesulitan", ["Campuran", "Mudah", "Sedang", "HOTS"], key="sj_tingkat")

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
        if not materi.strip():
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
                             jumlah=int(jumlah), tingkat=tingkat), model=model)
                st.session_state["sg_soal_list"] = hasil
                st.session_state["sg_soal_uid"] += 1
                st.session_state["sg_meta"] = {"mapel": mapel, "kelas": str(kelas), "jenis": jenis, "materi": materi.strip()}
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
                   f"\u2022 Materi: {meta_gen['materi'][:120]}")
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
                s["stimulus"] = st.text_area("Stimulus", s["stimulus"], key=f"sg_st_{k}", height=110)
                s["pertanyaan"] = st.text_input("Pertanyaan", s["pertanyaan"], key=f"sg_pt_{k}")
                for h in "ABCD":
                    s["opsi"][h] = st.text_input(f"Pilihan {h}", s["opsi"][h], key=f"sg_o{h}_{k}")
                s["kunci"] = st.selectbox("Kunci jawaban", list("ABCD"),
                                          index="ABCD".index(s["kunci"]) if s["kunci"] in "ABCD" and s["kunci"] else 0,
                                          key=f"sg_kc_{k}")
                s["pembahasan"] = st.text_area("Pembahasan", s["pembahasan"], key=f"sg_pb_{k}", height=80)
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
