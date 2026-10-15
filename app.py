"""
Dashboard Utilization Review PLKK  —  versi Streamlit
=====================================================
Jalankan di VS Code (terminal):

    pip install -r requirements.txt
    streamlit run app.py

Lalu upload file Excel dengan format template (sheet BMIV-01 s.d. BMIV-04,
contoh: Data_Dummy.xlsx) lewat sidebar.
"""
from __future__ import annotations

import io
import re
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ─────────────────────────────────────────────────────────────────────────────
# 1. KONSTANTA & SPESIFIKASI TEMPLATE
# ─────────────────────────────────────────────────────────────────────────────
BS = ["RJTP", "RJTL", "RANAP", "Khusus"]
BN = ["BMIV-01 · RJTP", "BMIV-02 · RJTL", "BMIV-03 · RANAP", "BMIV-04 · Khusus"]
BC = ["#2a9d8f", "#3b82f6", "#e9a23b", "#a855f7"]
BD = [
    "Rawat Jalan Tingkat Pertama — dokter umum, gigi, obat, penunjang diagnostik, tindakan medis.",
    "Rawat Jalan Tingkat Lanjutan — dokter spesialis, laboratorium, radiologi, fisioterapi, ODC, emergensi.",
    "Rawat Inap — akomodasi, operasi, ICU, transfusi, implan ortopedi, penunjang.",
    "Pelayanan khusus & alat bantu — kaca mata, gigi palsu, protesis, ambulans, program RTW.",
]
MON = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]
ACC = "#0f766e"

# Metadata kolom tiap sheet BMIV  (index = kode BMIV 0..3)
SHEETS = [
    dict(key="BMIV-01", date="TGL PELAYANAN", plkk="NAMA PLKK TK I",
         total="TOTAL BIAYA DISETUJUI", visit="KUNJUNGAN KE-"),
    dict(key="BMIV-02", date="TGL PELAYANAN", plkk="NAMA PLKK TK II / RS",
         total="TOTAL BIAYA DISETUJUI", visit="KUNJUNGAN KE-", poli="POLI / SPESIALISASI"),
    dict(key="BMIV-03", date="TGL MASUK", plkk="NAMA PLKK TK II / RS",
         total="TOTAL BIAYA DISETUJUI", los="JML HR RWT", kelas="KELAS RAWAT"),
    dict(key="BMIV-04", date="TGL PELAYANAN", plkk="TEMPAT PELAYANAN / PENYEDIA",
         total="JUMLAH JAMINAN (DISETUJUI)"),
]
COMMON_COLS = ["TGL KECELAKAAN", "NAMA TK / PASIEN", "NO KPJ", "NAMA PERUSAHAAN",
               "KODE CAB", "NO. KASUS KK", "JENIS KASUS", "DIAGNOSA", "KODE ICD-10", "KANWIL"]

# Komponen layanan: (kode, nama, satuan, [kolom volume], [kolom biaya])
def _pairs(rows, prefix, start=1):
    out = []
    for i, (nama, sat, vol, biaya) in enumerate(rows, start):
        out.append((f"{prefix}-{i:02d}", nama, sat, [vol], [biaya]))
    return out


SPEC = {
    0: [
        ("RJTP-01", "Dokter Umum", "kunjungan", ["JASA MEDIS DR UMUM | JML KUNJ"], ["BIAYA (RP)"]),
        ("RJTP-02", "Tindakan Gigi", "tindakan",
         ["JASA MEDIS DR GIGI (TRAUMA GIGI) | PX + KONSUL", "TS", "TP", "ExGT", "ExKOMP", "SPLINT / REPOSISI"],
         ["BIAYA (RP) [2]"]),
        ("RJTP-03", "Obat/Resep", "resep", ["OBAT | KS | JML", "LB | JML", "RX | JML"],
         ["BIAYA (RP) [3]", "BIAYA (RP) [4]", "BIAYA (RP) [5]"]),
        ("RJTP-04", "Penunjang Diagnostik", "kunjungan",
         ["PENUNJANG DIAGNOSTIK SEDERHANA | DR", "UR", "RO EKSTREMITAS", "RO THORAX / LAIN", "RO GIGI"],
         ["BIAYA (RP) [6]"]),
        ("RJTP-05", "Tindakan Medis Dokter Umum", "tindakan",
         ["TIND MEDIS DR UMUM | JAHIT LUKA", "BERSIH LUKA", "GANTI BALUT", "INSISI", "EKSISI", "EKSPLORASI",
          "DEBRIDEMENT", "EKSTRAKSI BENDA ASING", "BIDAI / SPALK", "REPOSISI SEDERHANA", "LUKA BAKAR"],
         ["BIAYA (RP) [7]"]),
        ("RJTP-06", "Tindakan Medis Dokter Gigi", "tindakan",
         ["TIND MEDIS DR GIGI | ODONTEKTOMI", "ALVEOLEKTOMI", "FLAP OPERASI", "INSISI [2]", "EKSISI [2]",
          "REPLANTASI / REPOSISI"],
         ["BIAYA (RP) [8]"]),
    ],
    1: [
        ("RJTL-01", "Jasa Dokter Spesialis", "kunjungan", ["JASA DOKTER | JML KUNJ"], ["BIAYA (RP)"]),
        ("RJTL-02", "Obat KS", "kasus", ["OBAT / RESEP | KS | JML"], ["BIAYA (RP) [2]"]),
        ("RJTL-03", "Obat LB", "kasus", ["LB | JML"], ["BIAYA (RP) [3]"]),
        ("RJTL-04", "Obat RX", "resep", ["RX | JML"], ["BIAYA (RP) [4]"]),
        ("RJTL-05", "Laboratorium", "kunjungan", ["LAB | JML KUNJ"], ["BIAYA (RP) [5]"]),
        ("RJTL-06", "Radiologi (Rontgen)", "kunjungan", ["RONTGEN | JML KUNJ"], ["BIAYA (RP) [6]"]),
        ("RJTL-07", "EKG", "kunjungan", ["ELEKTROMEDIK / PEMERIKSAAN KHUSUS | EKG | JML KUNJ"], ["BIAYA (RP) [7]"]),
        ("RJTL-08", "CT Scan", "kunjungan", ["CT SCAN | JML KUNJ"], ["BIAYA (RP) [8]"]),
        ("RJTL-09", "MRI", "kunjungan", ["MRI | JML KUNJ"], ["BIAYA (RP) [9]"]),
        ("RJTL-10", "USG", "kunjungan", ["USG | JML KUNJ"], ["BIAYA (RP) [10]"]),
        ("RJTL-11", "EMG/NCV", "kunjungan", ["EMG / NCV | JML KUNJ"], ["BIAYA (RP) [11]"]),
        ("RJTL-12", "Elektromedik Lain", "kunjungan", ["LAIN-LAIN | JML KUNJ"], ["BIAYA (RP) [12]"]),
        ("RJTL-13", "Tindakan Spesialis", "kunjungan",
         ["TINDAKAN MEDIS DOKTER SPESIALIS | TINDAKAN | JML KUNJ"], ["BIAYA (RP) [13]"]),
        ("RJTL-14", "ODC", "kunjungan", ["ODC (ONE DAY CARE) | JML KUNJ"], ["BIAYA (RP) [14]"]),
        ("RJTL-15", "Fisioterapi", "kasus", ["REHABILITASI MEDIK | FISIOTERAPI | KS"], ["BIAYA (RP) [15]"]),
        ("RJTL-16", "Rehab Medik Lain", "kasus", ["REHABILITASI MEDIK LAIN | KS"], ["BIAYA (RP) [16]"]),
        ("RJTL-17", "Emergensi", "kunjungan", ["EMERGENSI | JML KUNJ"], ["BIAYA (RP) [17]"]),
        ("RJTL-18", "Ambulans", "unit", ["AMBULANS | JML"], ["BIAYA (RP) [18]"]),
        ("RJTL-19", "Homecare", "kunjungan", ["HOMECARE | JML KUNJ."], ["BIAYA (RP) [19]"]),
    ],
    2: [
        ("RI-01", "Akomodasi Rawat Inap", "hari", ["JML HR RWT"], ["BIAYA DISETUJUI (RP)"]),
        ("RI-02", "Obat/Resep", "resep", ["OBAT / RESEP | LB", "KS", "RX"], ["BIAYA DISETUJUI (RP) [2]"]),
        ("RI-03", "Tindakan Non Operatif", "kunjungan", ["TINDAKAN NON OPERATIF | JML KUNJ"], ["BIAYA (RP)"]),
        ("RI-04", "Operasi Kecil", "kunjungan", ["TINDAKAN OPERATIF | KECIL | JML KUNJ"], ["BIAYA (RP) [2]"]),
        ("RI-05", "Operasi Sedang", "kunjungan", ["SEDANG | JML KUNJ"], ["BIAYA (RP) [3]"]),
        ("RI-06", "Operasi Besar", "kunjungan", ["BESAR | JML KUNJ"], ["BIAYA (RP) [4]"]),
        ("RI-07", "Operasi Khusus", "kunjungan", ["KHUSUS | JML KUNJ"], ["BIAYA (RP) [5]"]),
        ("RI-08", "ICU/ICCU/HCU/Burn Unit", "hari", ["JML HR RWT [2]"], ["BIAYA DISETUJUI (RP) [4]"]),
        ("RI-09", "Lab", "kunjungan", ["LAB | JML KUNJ"], ["BIAYA (RP) [6]"]),
        ("RI-10", "Rontgen", "kunjungan", ["RONTGEN | JML KUNJ"], ["BIAYA (RP) [7]"]),
        ("RI-11", "EKG", "kunjungan", ["ELEKTROMEDIK / PEMERIKSAAN KHUSUS | EKG | JML KUNJ"], ["BIAYA (RP) [8]"]),
        ("RI-12", "CT Scan", "kunjungan", ["CT SCAN | JML KUNJ"], ["BIAYA (RP) [9]"]),
        ("RI-13", "MRI", "kunjungan", ["MRI | JML KUNJ"], ["BIAYA (RP) [10]"]),
        ("RI-14", "USG", "kunjungan", ["USG | JML KUNJ"], ["BIAYA (RP) [11]"]),
        ("RI-15", "EMG/NCV", "kunjungan", ["EMG / NCV | JML KUNJ"], ["BIAYA (RP) [12]"]),
        ("RI-16", "Elektromedik Lain", "kunjungan", ["LAIN-LAIN | JML KUNJ"], ["BIAYA (RP) [13]"]),
        ("RI-17", "Pelayanan Darah", "kasus", ["PELAYANAN DARAH | KS"], ["BIAYA (RP) [14]"]),
        ("RI-18", "Transfusi Darah", "kasus", ["TRANSFUSI DARAH | KS"], ["BIAYA (RP) [15]"]),
        ("RI-19", "Implan Ortopedi", "kasus",
         ["ALAT KESEHATAN IMPLAN ORTOPEDI (PIN, PLATE, SCREW, DLL) | KS"], ["BIAYA (RP) [16]"]),
        ("RI-20", "Fisioterapi", "kasus", ["REHABILITASI MEDIK | FISIOTERAPI | KS"], ["BIAYA (RP) [17]"]),
        ("RI-21", "Rehab Medik Lain", "kasus", ["REHABILITASI MEDIK LAIN | KS"], ["BIAYA (RP) [18]"]),
        ("RI-22", "Ambulans", "unit", ["AMBULANS | JML"], ["BIAYA (RP) [19]"]),
        ("RI-23", "Homecare", "kunjungan", ["HOMECARE | JML KUNJ."], ["BIAYA (RP) [20]"]),
    ],
    3: [
        ("AB-01", "Kacamata", "kasus", ["KACAMATA | KS"], ["BIAYA (RP)"]),
        ("AB-02", "Gigi Palsu", "kasus", ["GIGI PALSU | KS"], ["BIAYA (RP) [2]"]),
        ("AB-03", "Protesis Anggota Gerak", "kasus", ["PROTESIS ANGGOTA GERAK | KS"], ["BIAYA (RP) [3]"]),
        ("AB-04", "Ortosis", "kasus", ["ORTOSIS | KS"], ["BIAYA (RP) [4]"]),
        ("AB-05", "Alat Bantu Jalan", "kasus", ["ALAT BANTU JALAN | KS"], ["BIAYA (RP) [5]"]),
        ("AB-06", "Alat Bantu Dengar", "kasus", ["ALAT BANTU DENGAR | KS"], ["BIAYA (RP) [6]"]),
        ("AB-07", "Mata Palsu", "kasus", ["MATA PALSU | KS"], ["BIAYA (RP) [7]"]),
        ("AB-08", "Ambulans/Transportasi", "kasus", ["AMBULANS / TRANSPORTASI | KS"], ["BIAYA (RP) [8]"]),
        ("AB-09", "Program Kembali Bekerja (RTW)", "kasus", ["PROGRAM KEMBALI BEKERJA (RTW) | KS"],
         ["BIAYA (RP) [9]"]),
    ],
}

# Nama cabang per KODE CAB (di template raw hanya ada kodenya)
CABANG = {
    "K00": "Jakarta Raya (cabang pelapor)", "K01": "Tangerang", "K02": "Bekasi-Karawang", "K03": "Bandung",
    "K04": "Cirebon", "K05": "Serang-Cilegon", "K06": "Semarang", "K07": "Yogyakarta", "K08": "Surakarta",
    "K09": "Surabaya", "K10": "Sidoarjo", "K11": "Malang", "K12": "Denpasar", "K13": "Mataram",
    "K14": "Kupang", "K15": "Medan", "K16": "Banda Aceh", "K17": "Batam", "K18": "Pekanbaru",
    "K19": "Padang", "K20": "Palembang", "K21": "Bandar Lampung", "K22": "Jambi", "K23": "Pontianak",
    "K24": "Banjarmasin", "K25": "Balikpapan-Samarinda", "K26": "Makassar", "K27": "Manado",
    "K28": "Kendari", "K29": "Jayapura-Timika",
}

# Sektor usaha diturunkan dari kata kunci pada NAMA PERUSAHAAN (urutan = prioritas)
SEKTOR_RULES = [
    ("Jasa Pendidikan", ["pendidikan"]),
    ("Jasa Penunjang", ["cleaning", "sekuriti"]),
    ("Perhotelan", ["hotel"]),
    ("Keuangan", ["finansial"]),
    ("Migas & Pertambangan", ["migas"]),
    ("Teknologi Informasi", ["teknologi"]),
    ("Pertanian & Perkebunan", ["sawit", "pertanian"]),
    ("Konstruksi", ["infrastruktur", "teknik sipil", "konstruksi", "pondasi", "struktur"]),
    ("Industri Manufaktur", ["kertas", "plastik", "kimia", "makanan", "logam"]),
    ("Perdagangan & Ritel", ["perdagangan", "niaga", "mart", "retail"]),
    ("Transportasi & Logistik", ["angkutan", "logistik", "terminal", "ekspedisi", "kargo", "pergudangan"]),
]


def sektor_of(nama: str) -> str:
    s = str(nama).lower()
    for sektor, kws in SEKTOR_RULES:
        if any(k in s for k in kws):
            return sektor
    return "Lainnya"


# ─────────────────────────────────────────────────────────────────────────────
# 2. LOADER: baca template Excel → tabel episode + tabel komponen
# ─────────────────────────────────────────────────────────────────────────────
def _norm(s) -> str:
    return re.sub(r"\s+", " ", str(s)).strip().upper()


class TemplateError(Exception):
    pass


def _find_sheet(xl: pd.ExcelFile, key: str) -> str | None:
    """Cari sheet 'BMIV-0x' (boleh ada akhiran seperti _Dummy)."""
    for name in xl.sheet_names:
        if _norm(name).replace("_", " ").startswith(key):
            return name
    return None


def required_columns(b: int) -> list[str]:
    m = SHEETS[b]
    cols = list(COMMON_COLS) + [m["date"], m["plkk"], m["total"]]
    for k in ("visit", "poli", "los", "kelas"):
        if k in m:
            cols.append(m[k])
    cols += ["JUMLAH KUNJUNGAN / UNIT"] if b == 3 else ["TOTAL BIAYA DIAJUKAN", "TOTAL SELISIH"]
    for _, _, _, vols, biaya in SPEC[b]:
        cols += vols + biaya
    seen, out = set(), []
    for c in cols:
        if c not in seen:
            seen.add(c)
            out.append(c)
    return out


def read_template(file) -> tuple[dict, list[str]]:
    """Baca & validasi 4 sheet BMIV. Return (dict b→DataFrame, daftar peringatan)."""
    try:
        xl = pd.ExcelFile(file)
    except Exception as e:  # noqa: BLE001
        raise TemplateError(f"File tidak bisa dibaca sebagai Excel (.xlsx): {e}")
    frames, warns, errors = {}, [], []
    for b, m in enumerate(SHEETS):
        sname = _find_sheet(xl, m["key"])
        if sname is None:
            errors.append(f"Sheet **{m['key']}** tidak ditemukan.")
            continue
        df = xl.parse(sname)
        df.columns = [_norm(c) for c in df.columns]
        miss = [c for c in required_columns(b) if _norm(c) not in df.columns]
        if miss:
            errors.append(f"Sheet **{sname}** kurang kolom: " + ", ".join(f"`{c}`" for c in miss[:12])
                          + (" …" if len(miss) > 12 else ""))
            continue
        frames[b] = df
    if errors:
        raise TemplateError("\n\n".join(errors))
    return frames, warns


def _num(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    return df[[_norm(c) for c in cols]].apply(pd.to_numeric, errors="coerce").fillna(0.0)


# Kolom penutup (total) & kolom subtotal per sheet — dipakai untuk deteksi pergeseran kolom
FOOTER_COL = {0: "RUJUKAN | R. JALAN", 1: "TOTAL BIAYA DIAJUKAN", 2: "TOTAL BIAYA DIAJUKAN",
              3: "JUMLAH KUNJUNGAN / UNIT"}
SUBTOTAL_COLS = {
    0: [],
    1: ["SUB TOTAL PENUNJANG"],
    2: ["BIAYA DISETUJUI (RP) [3]", "SUB TOTAL PENUNJANG", "SUB TOTAL LAIN-LAIN"],
    3: [],
}


def _plausible(vol: np.ndarray, cost: np.ndarray) -> np.ndarray:
    """Volume = bilangan bulat kecil, biaya = 0 atau nominal rupiah (>= 100)."""
    return (vol >= 0) & (vol <= 200) & (vol == np.floor(vol)) & ((cost == 0) | (cost >= 100))


def read_components(df: pd.DataFrame, b: int, ref: np.ndarray | None = None):
    """
    Baca volume & biaya tiap komponen layanan.
    Header template dan isi data kadang tidak sejajar (mis. file dummy: pasangan [volume, biaya]
    bergeser 1 kolom). Untuk tiap komponen dicoba offset 0 / +1 / -1 dan dipilih yang paling
    masuk akal (volume = bilangan bulat kecil, biaya = nominal rupiah).
    Return: (list komponen, daftar nama komponen yang dikoreksi).
    """
    cols = list(df.columns)
    pos = {c: i for i, c in enumerate(cols)}
    M = df.apply(pd.to_numeric, errors="coerce").fillna(0.0).to_numpy(dtype=float)
    limit = pos.get(_norm(FOOTER_COL[b]), len(cols))
    # kolom subtotal (header subtotal + 1 kolom sesudahnya yang berisi nilai subtotal) tidak dibaca sebagai komponen
    blocked = set()
    for c in SUBTOTAL_COLS[b]:
        if _norm(c) in pos:
            p = pos[_norm(c)]
            blocked.add(p)
            if _norm(c).startswith("SUB TOTAL"):
                blocked.add(p + 1)
    prev, out, shifted = 0, [], []
    for kode, nama, sat, vcols, bcols in SPEC[b]:
        vi = [pos[_norm(c)] for c in vcols]
        bi = [pos[_norm(c)] for c in bcols]
        best = None
        for o in dict.fromkeys([prev, 0, 1]):
            vv = [i + o for i in vi if 0 <= i + o < limit and (i + o) not in blocked]
            bb = [i + o for i in bi if 0 <= i + o < limit and (i + o) not in blocked]
            vol = M[:, vv].sum(axis=1) if vv else np.zeros(len(M))
            cost = M[:, bb].sum(axis=1) if bb else np.zeros(len(M))
            has = (vol != 0) | (cost != 0)
            ok = _plausible(vol, cost)
            net = int((has & ok).sum()) - int((has & ~ok).sum())
            if best is None or net > best[0]:
                best = (net, o, vol, cost)
        net, o, vol, cost = best
        if net > 0:
            prev = o
            if o != 0:
                shifted.append(nama)
        out.append([kode, nama, sat, vol, cost])
    # Komponen terakhir (Homecare / RTW): kolom biayanya bertabrakan dengan kolom total pada file yang
    # bergeser → biaya dihitung sebagai sisa (total - biaya komponen lain)
    if ref is not None and out and b in (1, 3):
        last = out[-1]
        others = sum(x[4] for x in out[:-1])
        need = (last[3] > 0) & (last[4] == 0)
        last[4] = np.where(need, np.maximum(ref - others, 0), last[4])
    # BMIV-03: biaya 'Operasi Khusus' bertabrakan dengan kolom subtotal operatif → hitung dari selisih subtotal
    if b == 2 and _norm("BIAYA DISETUJUI (RP) [3]") in pos:
        st = M[:, pos[_norm("BIAYA DISETUJUI (RP) [3]")]]
        names = [x[1] for x in out]
        if "Operasi Khusus" in names:
            k = names.index("Operasi Khusus")
            others = sum(x[4] for x in out if x[1] in ("Operasi Kecil", "Operasi Sedang", "Operasi Besar"))
            need = (out[k][3] > 0) & (out[k][4] == 0)
            out[k][4] = np.where(need, np.maximum(st - others, 0), out[k][4])
    return [tuple(x) for x in out], shifted


def read_total(df: pd.DataFrame, b: int) -> tuple[np.ndarray, np.ndarray, str | None]:
    """Total biaya disetujui per episode, dengan deteksi kolom total yang tertukar/bergeser."""
    cols = list(df.columns)
    pos = {c: i for i, c in enumerate(cols)}
    M = df.apply(pd.to_numeric, errors="coerce").fillna(0.0).to_numpy(dtype=float)
    m = SHEETS[b]
    if b == 3:  # BMIV-04: ambil kolom yang berisi nominal rupiah (bukan jumlah unit)
        c1, c2 = pos[_norm("JUMLAH KUNJUNGAN / UNIT")], pos[_norm(m["total"])]
        use = c1 if M[:, c1].sum() > M[:, c2].sum() else c2
        note = None if use == c2 else (
            "BMIV-04: kolom 'JUMLAH KUNJUNGAN / UNIT' dan 'JUMLAH JAMINAN (DISETUJUI)' tampak tertukar; "
            "kolom berisi nominal rupiah dipakai sebagai total disetujui.")
        return M[:, use], M[:, use], note
    d, a, sl = (pos[_norm(x)] for x in ("TOTAL BIAYA DIAJUKAN", m["total"], "TOTAL SELISIH"))
    best = None
    for o in (0, 1):
        if sl + o >= len(cols):
            continue
        ident = (M[:, d + o] > 0) & (np.abs(M[:, d + o] - M[:, a + o] - M[:, sl + o]) <= 1)
        score = int(ident.sum())
        if best is None or score > best[0]:
            best = (score, o)
    o = best[1]
    note = None if o == 0 else (
        f"{m['key']}: kolom TOTAL (diajukan/disetujui/selisih) tampak bergeser +1 kolom terhadap header; "
        "nilai disetujui dibaca dari posisi yang sesuai.")
    return M[:, a + o], M[:, d + o], note


def build_dataset(frames: dict) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    """
    Return:
      eps  — 1 baris = 1 episode pelayanan (1 baris sheet BMIV)
      comp — 1 baris = 1 komponen layanan pada 1 episode (biaya sudah dialokasikan
             proporsional terhadap TOTAL BIAYA DISETUJUI episode)
      warns — catatan kualitas data
    """
    ep_parts, comp_parts, warns = [], [], []
    uid = 0
    for b in range(4):
        df, m = frames[b], SHEETS[b]
        n = len(df)
        ids = np.arange(uid, uid + n)
        uid += n
        ep = pd.DataFrame({"ep": ids, "b": b})
        ep["tgl"] = pd.to_datetime(df[_norm(m["date"])], errors="coerce")
        ep["kpj"] = df["NO KPJ"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
        ep["peserta"] = df["NAMA TK / PASIEN"].astype(str)
        ep["kasus"] = df["NO. KASUS KK"].astype(str).str.strip()
        ep["kanwil"] = df["KANWIL"].astype(str).str.strip()
        ep["kode_cab"] = df["KODE CAB"].astype(str).str.strip()
        ep["cabang"] = ep["kode_cab"].map(CABANG).fillna(ep["kode_cab"])
        ep["plkk"] = df[_norm(m["plkk"])].astype(str).str.strip()
        ep["jenis_kasus"] = df["JENIS KASUS"].astype(str).str.strip()
        ep["perusahaan"] = df["NAMA PERUSAHAAN"].astype(str).str.strip()
        ep["sektor"] = ep["perusahaan"].map(sektor_of)
        ep["diagnosa"] = df["DIAGNOSA"].astype(str).str.strip()
        tot_arr, diaj_arr, tnote = read_total(df, b)
        ep["total"] = tot_arr
        if tnote:
            warns.append(tnote)
        ep["visit"] = pd.to_numeric(df[_norm(m["visit"])], errors="coerce") if "visit" in m else np.nan
        ep["los"] = pd.to_numeric(df[_norm(m["los"])], errors="coerce").fillna(0) if "los" in m else np.nan
        ep["kelas"] = df[_norm(m["kelas"])].astype(str) if "kelas" in m else ""
        ep["poli"] = df[_norm(m["poli"])].astype(str) if "poli" in m else ""

        # ── komponen layanan ──
        comps, shifted = read_components(df, b, diaj_arr)
        if shifted:
            warns.append(f"{m['key']}: komponen {', '.join(shifted)} terbaca dengan koreksi geser kolom "
                         "(isi data tidak sejajar dengan header template).")
        raw_sum = np.zeros(n)
        pieces = []
        for kode, nama, sat, vol, biaya in comps:
            raw_sum += biaya
            keep = (biaya > 0) | (vol > 0)
            if keep.any():
                pieces.append(pd.DataFrame({
                    "ep": ids[keep], "kode": kode, "komponen": nama, "satuan": sat, "b": b,
                    "volume": vol[keep], "biaya_raw": biaya[keep], "_i": np.flatnonzero(keep)}))
        if pieces:
            c = pd.concat(pieces, ignore_index=True)
            tot = ep["total"].to_numpy()
            scale = np.where(raw_sum > 0, tot / np.where(raw_sum > 0, raw_sum, 1), 0.0)
            c["biaya"] = c["biaya_raw"] * scale[c["_i"].to_numpy()]
            comp_parts.append(c.drop(columns=["_i"]))
        ep_parts.append(ep)

    eps = pd.concat(ep_parts, ignore_index=True)
    comp = pd.concat(comp_parts, ignore_index=True) if comp_parts else pd.DataFrame(
        columns=["ep", "kode", "komponen", "satuan", "b", "volume", "biaya_raw", "biaya"])

    bad_date = int(eps["tgl"].isna().sum())
    if bad_date:
        warns.append(f"{bad_date} baris dibuang karena tanggal pelayanan/masuk kosong atau tidak valid.")
        eps = eps[eps["tgl"].notna()]
        comp = comp[comp["ep"].isin(eps["ep"])]
    eps["periode"] = eps["tgl"].dt.strftime("%Y-%m")
    eps["tgl_s"] = eps["tgl"].dt.strftime("%Y-%m-%d")
    comp = comp.merge(eps[["ep", "periode", "kpj", "kasus", "kanwil", "cabang", "plkk",
                           "jenis_kasus", "sektor", "tgl"]], on="ep", how="left")
    return eps.reset_index(drop=True), comp.reset_index(drop=True), warns


@st.cache_data(show_spinner="Membaca & menormalisasi data…")
def load_data(content: bytes):
    frames, w1 = read_template(io.BytesIO(content))
    eps, comp, w2 = build_dataset(frames)
    return eps, comp, w1 + w2


# ─────────────────────────────────────────────────────────────────────────────
# 3. FORMATTER (format Indonesia)
# ─────────────────────────────────────────────────────────────────────────────
def fN(n) -> str:
    return f"{round(float(n)):,}".replace(",", ".")


def fF(n) -> str:
    return "Rp " + fN(n)


def fR(n) -> str:
    n = float(n)
    a = abs(n)
    if a >= 1e9:
        return "Rp " + f"{n / 1e9:.2f}".replace(".", ",") + " M"
    if a >= 1e6:
        return "Rp " + f"{n / 1e6:.1f}".replace(".", ",") + " jt"
    return fF(n)


def f1(n) -> str:
    r = round(float(n), 1)
    if r == int(r):
        return fN(r)
    return f"{r:,.1f}".replace(",", "§").replace(".", ",").replace("§", ".")


def fP(n) -> str:
    return f1(float(n) * 100) + "%"


def ml(p: str) -> str:
    return f"{MON[int(p[5:7]) - 1]} {p[2:4]}"


# ─────────────────────────────────────────────────────────────────────────────
# 4. AGREGASI
# ─────────────────────────────────────────────────────────────────────────────
def agg(L: pd.DataFrame, m: int) -> dict:
    k, c, n, t = L["kpj"].nunique(), L["kasus"].nunique(), len(L), float(L["total"].sum())
    return dict(pes=k, kas=c, kun=n, bi=t, uc=t / n if n else 0.0, pm=t / (k * m) if k else 0.0,
                cpc=t / c if c else 0.0, ut=n / k * 1000 if k else 0.0)


def cagg(C: pd.DataFrame) -> pd.DataFrame:
    if C.empty:
        return pd.DataFrame(columns=["kode", "komponen", "b", "vol", "bi"])
    return (C.groupby(["kode", "komponen", "b"], as_index=False)
            .agg(vol=("volume", "sum"), bi=("biaya", "sum"))
            .sort_values("bi", ascending=False).reset_index(drop=True))


# ─────────────────────────────────────────────────────────────────────────────
# 5. HELPER UI
# ─────────────────────────────────────────────────────────────────────────────
_KEY = [0]


def _nk() -> str:
    _KEY[0] += 1
    return f"plt{_KEY[0]}"


def plot(fig: go.Figure, height: int = 360):
    fig.update_layout(height=height, margin=dict(l=10, r=10, t=30, b=10), separators=",.",
                      legend=dict(orientation="h", yanchor="top", y=-0.15, x=0))
    try:
        st.plotly_chart(fig, width="stretch", key=_nk())
    except TypeError:  # Streamlit versi lama
        st.plotly_chart(fig, use_container_width=True, key=_nk())


def table(df: pd.DataFrame, fmts: dict | None = None, height: int | None = None, total: list | None = None):
    sty = df.style.format(fmts or {}, na_rep="–")
    kw = dict(hide_index=True)
    if height:
        kw["height"] = height
    try:
        st.dataframe(sty, width="stretch", **kw)
    except TypeError:
        st.dataframe(sty, use_container_width=True, **kw)
    if total is not None:
        t = pd.DataFrame([total], columns=df.columns)
        tf = {k: v for k, v in (fmts or {}).items()}
        try:
            st.dataframe(t.style.format(tf, na_rep=""), width="stretch", hide_index=True)
        except TypeError:
            st.dataframe(t.style.format(tf, na_rep=""), use_container_width=True, hide_index=True)


def kpis(items: list[tuple]):
    cols = st.columns(len(items))
    for col, it in zip(cols, items):
        title, val = it[0], it[1]
        sub = it[2] if len(it) > 2 else ""
        with col.container(border=True):
            st.metric(title, val)
            if sub:
                st.caption(sub)


def money_axis(fig: go.Figure, vmax: float, axis: str = "x"):
    """Tick uang ringkas (jt / M) pada sumbu."""
    if vmax <= 0:
        return
    ticks = np.linspace(0, vmax, 6)
    txt = [fR(v).replace("Rp ", "") if v else "0" for v in ticks]
    (fig.update_xaxes if axis == "x" else fig.update_yaxes)(tickvals=ticks, ticktext=txt)


def hbar(labels, values, color, fmt=fF, title=None, height=None, axis_money=True):
    labels, values = list(labels), [float(v) for v in values]
    fig = go.Figure(go.Bar(x=values, y=labels, orientation="h", marker_color=color,
                           text=[fmt(v) for v in values], textposition="outside", cliponaxis=False,
                           hovertemplate="%{y}<br>%{text}<extra></extra>"))
    fig.update_yaxes(autorange="reversed", automargin=True)
    vmax = max(values) if values else 0
    fig.update_xaxes(range=[0, vmax * 1.28 if vmax else 1])
    if axis_money:
        money_axis(fig, vmax * 1.28)
    if title:
        fig.update_layout(title=title)
    plot(fig, height or max(260, 28 * len(labels) + 70))


def empty():
    st.info("Tidak ada data untuk kombinasi filter ini. Ubah atau reset filter.")


# ─────────────────────────────────────────────────────────────────────────────
# 6. TAB-TAB DASHBOARD
# ─────────────────────────────────────────────────────────────────────────────
def tab_exec(L, C, m):
    a = agg(L, m)
    kpis([("Total Peserta", fN(a["pes"]), "NO KPJ unik"), ("Kasus", fN(a["kas"]), "No. Kasus KK unik"),
          ("Kunjungan", fN(a["kun"]), "episode layanan"), ("Total Biaya", fR(a["bi"]), fF(a["bi"])),
          ("Unit Cost", fF(a["uc"]), "biaya / kunjungan"),
          ("PMPM", fF(a["pm"]), f"biaya / peserta / bulan ({m} bln)")])
    c1, c2 = st.columns(2)
    with c1.container(border=True):
        st.markdown("**Komposisi Biaya per Kelompok Pelayanan (BMIV)**")
        st.caption("Porsi biaya disetujui")
        bb = [float(L.loc[L.b == b, "total"].sum()) for b in range(4)]
        fig = go.Figure(go.Pie(labels=BN, values=bb, hole=0.6, marker=dict(colors=BC), sort=False,
                               texttemplate="%{percent:.1%}", hovertemplate="%{label}<br>Rp %{value:,.0f}<extra></extra>"))
        plot(fig, 340)
    with c2.container(border=True):
        st.markdown("**Top 10 Komponen Layanan**")
        st.caption("Penyerapan biaya tertinggi")
        if not C.empty:
            t = C.assign(lbl=C["komponen"] + " · " + C["b"].map(dict(enumerate(BS)))) \
                 .groupby("lbl")["biaya"].sum().sort_values(ascending=False).head(10)
            hbar(t.index, t.values, ACC, height=340)
    st.caption("PMPM memakai peserta unik yang melapor klaim × jumlah bulan pada rentang filter (proxy; jumlah tenaga "
               "kerja terdaftar tidak tersedia pada data). Biaya per komponen dialokasikan proporsional terhadap "
               "Total Biaya Disetujui tiap baris.")


def tab_trend(L, periods, m):
    sel = st.selectbox("Metrik", ["Biaya", "Volume (Kunjungan)", "Utilisasi (kunjungan per 1.000 peserta)"], key="tm")
    rows = []
    for p in periods:
        X = L[L.periode == p]
        a = agg(X, 1)
        rows.append(dict(p=p, X=X, pes=a["pes"], tot=a))
    labels = [ml(p) for p in periods]

    def val(o, b):
        X = o["X"] if b < 0 else o["X"][o["X"].b == b]
        if sel == "Biaya":
            return float(X["total"].sum())
        if sel.startswith("Volume"):
            return float(len(X))
        return len(X) / o["pes"] * 1000 if o["pes"] else 0.0

    with st.container(border=True):
        st.markdown("**Tren Bulanan**")
        st.caption("Seluruh bulan dalam rentang filter, urut kronologis")
        fig = go.Figure()
        if sel.startswith("Utilisasi"):
            fig.add_scatter(x=labels, y=[val(o, -1) for o in rows], name="Total", mode="lines+markers",
                            line=dict(color="#1c2430", width=3))
            for b in range(4):
                fig.add_scatter(x=labels, y=[val(o, b) for o in rows], name=BN[b], mode="lines+markers",
                                line=dict(color=BC[b]))
            fig.update_yaxes(title="kunjungan / 1.000 peserta")
        else:
            for b in range(4):
                fig.add_bar(x=labels, y=[val(o, b) for o in rows], name=BN[b], marker_color=BC[b])
            fig.update_layout(barmode="stack")
            if sel == "Biaya":
                vmax = max([sum(val(o, b) for b in range(4)) for o in rows] or [0])
                money_axis(fig, vmax * 1.05, "y")
        fig.update_xaxes(type="category")
        plot(fig, 380)
    df = pd.DataFrame(dict(
        Bulan=[ml(o["p"]) for o in rows], Peserta=[o["tot"]["pes"] for o in rows],
        Kasus=[o["tot"]["kas"] for o in rows], Kunjungan=[o["tot"]["kun"] for o in rows],
        Biaya=[o["tot"]["bi"] for o in rows], **{"Unit Cost": [o["tot"]["uc"] for o in rows]},
        **{"Util /1.000": [o["tot"]["ut"] for o in rows]}))
    bi = df["Biaya"].to_numpy()
    df["Δ Biaya MoM"] = [np.nan] + [(bi[j] / bi[j - 1] - 1) if bi[j - 1] else np.nan for j in range(1, len(bi))]
    with st.container(border=True):
        st.markdown("**Rekap Bulanan**")
        table(df, {"Peserta": fN, "Kasus": fN, "Kunjungan": fN, "Biaya": fF, "Unit Cost": fF, "Util /1.000": f1,
                   "Δ Biaya MoM": lambda v: ("+" if v >= 0 else "") + f1(v * 100) + "%"})


def tab_bmiv(L, C, b, m):
    L = L[L.b == b]
    if L.empty:
        return empty()
    C = C[C.b == b]
    a = agg(L, m)
    cg = cagg(C)
    st.markdown(f"**{BN[b]}** — {BD[b]}")
    items = [("Total Biaya", fR(a["bi"]), fF(a["bi"])), ("Kunjungan / Episode", fN(a["kun"])),
             ("Kasus", fN(a["kas"])), ("Peserta", fN(a["pes"])), ("Unit Cost", fF(a["uc"]), "biaya / kunjungan")]
    if b == 2:
        h = float(L["los"].sum())
        items.append(("ALOS", f1(h / len(L)) + " hari", f"total {fN(h)} hari rawat"))
    kpis(items)
    c1, c2 = st.columns(2)
    with c1.container(border=True):
        st.markdown("**Biaya per Komponen**")
        top = cg.head(15)
        hbar(top["komponen"], top["bi"], BC[b])
    with c2.container(border=True):
        if b == 2:
            st.markdown("**Distribusi ALOS (hari rawat)**")
            bk = [("1–3", 1, 3), ("4–7", 4, 7), ("8–10", 8, 10), ("11–14", 11, 14), ("15+", 15, 999)]
            fig = go.Figure(go.Bar(x=[x[0] + " hr" for x in bk],
                                   y=[int(((L.los >= lo) & (L.los <= hi)).sum()) for _, lo, hi in bk],
                                   marker_color=BC[2], name="Kasus"))
            plot(fig, 340)
        elif b == 1:
            st.markdown("**Kunjungan per Poli/Spesialisasi**")
            t = L["poli"].value_counts().head(10)
            hbar(t.index, t.values, BC[1], fmt=fN, axis_money=False)
        else:
            st.markdown("**Top Diagnosa (biaya)**")
            t = L.groupby("diagnosa")["total"].sum().sort_values(ascending=False).head(10)
            hbar(t.index, t.values, BC[b])
    with st.container(border=True):
        st.markdown("**Rincian Komponen**")
        df = pd.DataFrame(dict(Kode=cg["kode"], Komponen=cg["komponen"], Volume=cg["vol"], Biaya=cg["bi"],
                               **{"% Biaya": cg["bi"] / a["bi"] if a["bi"] else 0.0},
                               **{"Unit Cost": np.where(cg["vol"] > 0, cg["bi"] / cg["vol"].where(cg["vol"] > 0, 1), 0.0)}))
        table(df, {"Volume": fN, "Biaya": fF, "% Biaya": fP, "Unit Cost": fF},
              total=["Total", "", cg["vol"].sum(), a["bi"], 1.0, np.nan])
    if b == 2:
        with st.container(border=True):
            st.markdown("**ALOS per Kelas Rawat**")
            g = L.groupby("kelas").agg(Kasus=("ep", "size"), Hari=("los", "sum"), Biaya=("total", "sum")).reset_index()
            g["ALOS"] = g["Hari"] / g["Kasus"]
            g["Biaya / Kasus"] = g["Biaya"] / g["Kasus"]
            g = g.rename(columns={"kelas": "Kelas", "Hari": "Hari Rawat", "ALOS": "ALOS (hari)"})
            g = g[["Kelas", "Kasus", "Hari Rawat", "ALOS (hari)", "Biaya", "Biaya / Kasus"]].sort_values("Kelas")
            table(g, {"Kasus": fN, "Hari Rawat": fN, "ALOS (hari)": f1, "Biaya": fF, "Biaya / Kasus": fF})


def tab_plkk(L, m):
    c1, c2 = st.columns(2)
    bs = c1.selectbox("Basis perbandingan", ["Semua layanan"] + BN, key="pb")
    mt = c2.selectbox("Metrik", ["Total Biaya", "Kunjungan (utilisasi)", "Unit Cost", "Cost per Case"], key="pm")
    X = L if bs == "Semua layanan" else L[L.b == BN.index(bs)]
    if X.empty:
        return empty()
    g = X.groupby("plkk").agg(pes=("kpj", "nunique"), kas=("kasus", "nunique"), kun=("ep", "size"),
                              bi=("total", "sum")).reset_index()
    g["uc"] = g["bi"] / g["kun"]
    g["cpc"] = g["bi"] / g["kas"]
    g["ut"] = g["kun"] / g["pes"] * 1000
    col, fmt = {"Total Biaya": ("bi", fF), "Kunjungan (utilisasi)": ("kun", fN), "Unit Cost": ("uc", fF),
                "Cost per Case": ("cpc", fF)}[mt]
    T = g[g["kun"] >= 5] if mt in ("Unit Cost", "Cost per Case") else g
    T = T.sort_values(col, ascending=False).head(15)
    with st.container(border=True):
        st.markdown("**Peringkat PLKK (Top 15)**")
        st.caption("Bandingkan PLKK pada level layanan yang sama (apple-to-apple) dengan memilih basis BMIV. "
                   "Metrik rasio hanya untuk PLKK dengan ≥5 kunjungan.")
        hbar(T["plkk"], T[col], ACC, fmt=fmt, axis_money=col != "kun")
    with st.container(border=True):
        st.markdown(f"**Rekapitulasi PLKK** — {len(g)} PLKK")
        df = g.rename(columns=dict(plkk="PLKK", pes="Peserta", kas="Kasus", kun="Kunjungan", bi="Total Biaya",
                                   uc="Unit Cost", cpc="Cost per Case", ut="Util /1.000")) \
              .sort_values("Total Biaya", ascending=False)
        table(df, {"Peserta": fN, "Kasus": fN, "Kunjungan": fN, "Total Biaya": fF, "Unit Cost": fF,
                   "Cost per Case": fF, "Util /1.000": f1}, height=460)


def tab_lbst(L, C, periods):
    sel = st.selectbox("Tampilkan", ["Biaya (Rp)", "Volume"], key="ls")
    val = "biaya" if sel.startswith("Biaya") else "volume"
    X = C[C.periode.isin(periods)]
    with st.container(border=True):
        st.markdown("**LB-ST — Rekapitulasi Layanan Aktual**")
        st.caption("Rekap layanan aktual per komponen × bulan dari BMIV-01 s.d. BMIV-04.")
        if X.empty:
            return empty()
        pv = X.pivot_table(index=["b", "kode", "komponen"], columns="periode", values=val, aggfunc="sum",
                           fill_value=0.0).reindex(columns=periods, fill_value=0.0)
        pv["Total"] = pv.sum(axis=1)
        pv = pv.reset_index().sort_values(["b", "kode"])
        pv["Layanan"] = pv["b"].map(dict(enumerate(BS)))
        pv = pv.rename(columns={**{p: ml(p) for p in periods}, "komponen": "Komponen"})
        cols = ["Layanan", "Komponen"] + [ml(p) for p in periods] + ["Total"]
        pv = pv[cols]
        num = cols[2:]
        table(pv, {c: fN for c in num}, height=520, total=["Total", ""] + [pv[c].sum() for c in num])


def tab_uc(L, C, m):
    bs = st.selectbox("Kelompok", ["Semua BMIV"] + BN, key="ub")
    if bs != "Semua BMIV":
        b = BN.index(bs)
        L, C = L[L.b == b], C[C.b == b]
    if L.empty:
        return empty()
    a = agg(L, m)
    cg = cagg(C)
    kpis([("Unit Cost", fF(a["uc"]), "biaya / kunjungan"), ("Cost per Case", fF(a["cpc"])),
          ("PMPM", fF(a["pm"]), f"{m} bulan"), ("Utilisasi /1.000 peserta", f1(a["ut"]), "kunjungan per 1.000")])
    cg["lbl"] = cg["komponen"] + " (" + cg["b"].map(dict(enumerate(BS))) + ")"
    cg["pmpm"] = cg["bi"] / (a["pes"] * m) if a["pes"] else 0.0
    cg["ucc"] = np.where(cg["vol"] > 0, cg["bi"] / cg["vol"].where(cg["vol"] > 0, 1), 0.0)
    c1, c2 = st.columns(2)
    with c1.container(border=True):
        st.markdown("**Top 10 PMPM per Komponen**")
        t = cg.sort_values("pmpm", ascending=False).head(10)
        hbar(t["lbl"], t["pmpm"], ACC)
    with c2.container(border=True):
        st.markdown("**Top 10 Unit Cost per Komponen**")
        st.caption("Komponen dengan volume ≥3")
        t = cg[cg["vol"] >= 3].sort_values("ucc", ascending=False).head(10)
        hbar(t["lbl"], t["ucc"], "#a855f7")
    with st.container(border=True):
        st.markdown("**Unit Cost, Utilisasi & PMPM per Komponen**")
        st.caption("Proporsi = % terhadap total biaya kelompok terpilih")
        df = pd.DataFrame(dict(
            Layanan=cg["b"].map(dict(enumerate(BS))), Komponen=cg["komponen"], Volume=cg["vol"], Biaya=cg["bi"],
            **{"% Proporsi": cg["bi"] / a["bi"] if a["bi"] else 0.0, "Unit Cost": cg["ucc"],
               "Util /1.000": cg["vol"] / a["pes"] * 1000 if a["pes"] else 0.0, "PMPM": cg["pmpm"]}))
        table(df, {"Volume": fN, "Biaya": fF, "% Proporsi": fP, "Unit Cost": fF, "Util /1.000": f1, "PMPM": fF},
              height=460, total=["Total", "", cg["vol"].sum(), a["bi"], 1.0, a["uc"], a["ut"], a["pm"]])


def tab_case(eps, comp):
    st.caption("Case Explorer menelusuri seluruh data (tidak terpengaruh filter sidebar).")
    top = eps.groupby("kasus")["total"].sum().sort_values(ascending=False).head(12)
    cases = sorted(eps["kasus"].unique())

    def pick(c):
        st.session_state["case_sel"] = c

    st.selectbox("Case ID (ketik untuk mencari)", [""] + cases, key="case_sel",
                 format_func=lambda x: "— pilih No. Kasus —" if x == "" else x)
    cur = st.session_state.get("case_sel", "")
    if not cur:
        with st.container(border=True):
            st.markdown("**Pilih Case ID**")
            st.caption("Pilih No. Kasus di atas, atau klik salah satu kasus berbiaya tertinggi:")
            cols = st.columns(4)
            for i, (c, v) in enumerate(top.items()):
                cols[i % 4].button(f"{c} · {fR(v)}", key=f"chip{i}", on_click=pick, args=(c,), width="stretch") \
                    if _has_width_param() else cols[i % 4].button(f"{c} · {fR(v)}", key=f"chip{i}", on_click=pick, args=(c,))
        return
    E = eps[eps.kasus == cur].sort_values(["tgl_s", "b"])
    cc = comp[comp.ep.isin(E.ep)]
    names = {}
    for ep, g in cc.groupby("ep"):
        names[ep] = ", ".join(f"{r.komponen}{' ×' + fN(r.volume) if r.volume > 1 else ''}" for r in g.itertuples())
    t = float(E["total"].sum())
    days = (E["tgl"].max() - E["tgl"].min()).days
    first = E.iloc[0]
    kpis([("Total Biaya", fF(t)), ("Layanan / Episode", str(len(E))), ("PLKK Terlibat", str(E["plkk"].nunique())),
          ("Rentang", f"{days} hari", f"{E.tgl_s.iloc[0]} → {E.tgl_s.iloc[-1]}"),
          ("Jenis Kasus", first["jenis_kasus"], first["sektor"]), ("Diagnosa", first["diagnosa"], "KPJ " + first["kpj"])])
    with st.container(border=True):
        st.markdown("**Timeline / Journey Layanan**")
        st.caption("Urut tanggal pelayanan / masuk RS")
        html = []
        for e in E.itertuples():
            extra = f" · {int(e.los)} hari rawat" if e.b == 2 else ""
            html.append(
                f"<div style='border-left:4px solid {BC[e.b]};padding:2px 0 2px 12px;margin:0 0 12px 4px'>"
                f"<span style='background:{BC[e.b]};color:#fff;border-radius:99px;padding:1px 9px;font-size:12px'>{BS[e.b]}</span> "
                f"<b>{e.tgl_s}</b> · {e.plkk}<br><span style='opacity:.7'>{names.get(e.ep, '–')}{extra}</span><br>{fF(e.total)}</div>")
        st.markdown("".join(html), unsafe_allow_html=True)
    with st.container(border=True):
        st.markdown("**Detail Baris Data**")
        df = pd.DataFrame(dict(Tanggal=E["tgl_s"], Layanan=E["b"].map(dict(enumerate(BN))), PLKK=E["plkk"],
                               **{"Kunjungan ke": E["visit"], "Komponen": E["ep"].map(names).fillna("–"),
                                  "Biaya Disetujui": E["total"]}))
        table(df, {"Kunjungan ke": lambda v: "–" if pd.isna(v) or not v else fN(v), "Biaya Disetujui": fF})


def _has_width_param() -> bool:
    try:
        import inspect
        return "width" in inspect.signature(st.button).parameters
    except Exception:  # noqa: BLE001
        return False


# ─────────────────────────────────────────────────────────────────────────────
# 7. APLIKASI UTAMA
# ─────────────────────────────────────────────────────────────────────────────
FILTER_KEYS = ["f_kw", "f_cb", "f_pl", "f_b", "f_jk", "f_sk"]


def reset_filters(dmin, dmax):
    for k in FILTER_KEYS:
        st.session_state[k] = []
    st.session_state["f_date"] = (dmin, dmax)
    st.session_state["case_sel"] = ""


def main():
    st.set_page_config(page_title="Dashboard Utilization Review PLKK", page_icon="🏥", layout="wide")
    _KEY[0] = 0
    sample = Path(__file__).with_name("Data_Dummy.xlsx")

    sb = st.sidebar
    sb.title("Utilization Review PLKK")
    up = sb.file_uploader("Upload file Excel (format template BMIV-01 s.d. BMIV-04)", type=["xlsx", "xlsm"])
    if up is not None:
        content, src = up.getvalue(), up.name
    elif sample.exists():
        content, src = sample.read_bytes(), f"{sample.name} (contoh)"
        sb.info("Memakai file contoh **Data_Dummy.xlsx**. Upload file Anda di atas untuk menggantinya.")
    else:
        st.title("Dashboard Utilization Review PLKK")
        st.info("Upload file Excel dengan format template (sheet **BMIV-01**, **BMIV-02**, **BMIV-03**, **BMIV-04**) "
                "melalui sidebar untuk memulai.")
        st.stop()

    try:
        eps, comp, warns = load_data(content)
    except TemplateError as e:
        st.error("File tidak sesuai template:")
        st.markdown(str(e))
        st.caption("Pastikan workbook memiliki sheet BMIV-01, BMIV-02, BMIV-03, BMIV-04 (boleh berakhiran "
                   "mis. _Dummy) dengan nama kolom persis seperti template.")
        st.stop()
    if eps.empty:
        st.error("Tidak ada baris data yang valid pada file.")
        st.stop()

    all_periods = sorted(eps["periode"].unique())
    dmin, dmax = eps["tgl"].min().date(), eps["tgl"].max().date()
    sig = (src, len(content))
    if st.session_state.get("sig") != sig:
        reset_filters(dmin, dmax)
        st.session_state["sig"] = sig

    # ── Filter ──
    sb.caption(f"Sumber: **{src}** · {ml(all_periods[0])} – {ml(all_periods[-1])}")
    sb.date_input("Periode Tanggal", min_value=dmin, max_value=dmax, key="f_date")
    sb.multiselect("Kanwil", sorted(eps["kanwil"].unique()), key="f_kw", placeholder="Semua Kanwil")
    sb.multiselect("Cabang", sorted(eps["cabang"].unique()), key="f_cb", placeholder="Semua Cabang")
    sb.multiselect("PLKK", sorted(eps["plkk"].unique()), key="f_pl", placeholder="Semua PLKK")
    sb.multiselect("Jenis Layanan", BN, key="f_b", placeholder="Semua Layanan")
    sb.multiselect("Jenis Kasus", sorted(eps["jenis_kasus"].unique()), key="f_jk", placeholder="Semua Jenis Kasus")
    sb.multiselect("Sektor Usaha", sorted(eps["sektor"].unique()), key="f_sk", placeholder="Semua Sektor")
    sb.button("Reset filter", on_click=reset_filters, args=(dmin, dmax), width="stretch") if _has_width_param() \
        else sb.button("Reset filter", on_click=reset_filters, args=(dmin, dmax))

    dr = st.session_state.get("f_date", (dmin, dmax))
    d1 = dr[0] if len(dr) >= 1 else dmin
    d2 = dr[1] if len(dr) >= 2 else dmax
    mask = (eps["tgl_s"] >= d1.isoformat()) & (eps["tgl_s"] <= d2.isoformat())
    for key, col in [("f_kw", "kanwil"), ("f_cb", "cabang"), ("f_pl", "plkk"), ("f_jk", "jenis_kasus"),
                     ("f_sk", "sektor")]:
        sel = st.session_state.get(key) or []
        if sel:
            mask &= eps[col].isin(sel)
    sel_b = [BN.index(x) for x in (st.session_state.get("f_b") or [])]
    if sel_b:
        mask &= eps["b"].isin(sel_b)
    L = eps[mask]
    C = comp[comp["ep"].isin(set(L["ep"]))]
    periods = [p for p in all_periods if d1.isoformat()[:7] <= p <= d2.isoformat()[:7]]
    m = max(1, len(periods))
    sb.caption(f"{fN(len(L))} dari {fN(len(eps))} baris terfilter")

    # ── Utilitas sidebar ──
    with sb.expander("Catatan kualitas data", expanded=bool(warns)):
        if warns:
            for w in warns:
                st.markdown(f"- {w}")
        else:
            st.markdown("Tidak ada catatan. Struktur kolom sesuai template.")
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        eps.drop(columns=["tgl_s"]).to_excel(xw, sheet_name="Episode", index=False)
        comp.drop(columns=["tgl"]).to_excel(xw, sheet_name="Komponen", index=False)
    sb.download_button("Unduh data ternormalisasi (.xlsx)", buf.getvalue(), "data_ternormalisasi_plkk.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if sample.exists():
        sb.download_button("Unduh file contoh / template", sample.read_bytes(), sample.name,
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ── Konten ──
    st.title("Dashboard Utilization Review PLKK")
    tabs = st.tabs(["Executive Summary", "Tren Bulanan", "BMIV-01 RJTP", "BMIV-02 RJTL", "BMIV-03 RANAP",
                    "BMIV-04 Khusus", "PLKK Performance", "LB-ST", "Unit Cost & Per Kapita", "Case Explorer"])
    with tabs[9]:
        tab_case(eps, comp)
    if L.empty:
        for t in tabs[:9]:
            with t:
                empty()
        return
    with tabs[0]:
        tab_exec(L, C, m)
    with tabs[1]:
        tab_trend(L, periods, m)
    for b in range(4):
        with tabs[2 + b]:
            tab_bmiv(L, C, b, m)
    with tabs[6]:
        tab_plkk(L, m)
    with tabs[7]:
        tab_lbst(L, C, periods)
    with tabs[8]:
        tab_uc(L, C, m)


if __name__ == "__main__":
    main()
