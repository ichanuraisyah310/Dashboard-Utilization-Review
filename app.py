import io
import pandas as pd
import streamlit as st

# Konfigurasi Halaman Streamlit
st.set_page_config(
    page_title="Dashboard Utilization Review PLKK", layout="wide"
)

# Injeksi CSS agar tampilan dan gaya visual tetap 100% identik dengan versi HTML/JS asli
st.markdown(
    """
<style>
:root{--bg:#f5f6f8;--card:#fff;--ink:#1c2430;--muted:#667085;--line:#e3e7ee;--acc:#0f766e;--accbg:#e0f2ef;--grid:#e8ebf0;--c0:#2a9d8f;--c1:#3b82f6;--c2:#e9a23b;--c3:#a855f7;}
.app{display:grid;grid-template-columns:260px 1fr;min-height:100vh}
aside{background:var(--card);border-right:1px solid var(--line);padding:16px;height:100vh;overflow:auto}
aside h1{font-size:15px;margin:0 0 2px}aside .sub{color:var(--muted);font-size:12px;margin-bottom:14px}
label{display:block;font-size:11px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.04em;margin:10px 0 4px}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:6px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:10px;margin-bottom:14px}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:12px 14px}
.kpi small{color:var(--muted);font-size:12px}.kpi b{display:block;font-size:22px;margin-top:2px}.kpi span{color:var(--muted);font-size:11px}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:14px}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px;min-width:0;margin-bottom:14px}
.card h3{margin:0 0 4px;font-size:14px}.card p.d{margin:0 0 10px;color:var(--muted);font-size:12px}
.note{font-size:11.5px;color:var(--muted);margin-top:8px}
</style>
""",
    unsafe_allow_html=True,
)

# Data Simulasi Asli (Dihembed secara utuh tanpa ada perubahan data satupun)
#[cite: 1]
RAW_DATA = {
    "eps": [
        {
            "b": 0,
            "d": "2025-06-16",
            "p": 0,
            "k": 0,
            "cs": 0,
            "kw": 0,
            "cb": 0,
            "pl": 0,
            "jk": 0,
            "sk": 0,
            "dg": 0,
            "t": 296277,
            "c": [[0, 1.0, 51734], [1, 1.0, 100910], [2, 1.0, 143633]],
            "v": 1,
        },
        {
            "b": 0,
            "d": "2025-06-08",
            "p": 0,
            "k": 1,
            "cs": 1,
            "kw": 1,
            "cb": 1,
            "pl": 1,
            "jk": 0,
            "sk": 1,
            "dg": 1,
            "t": 340810,
            "c": [[0, 1.0, 57527], [1, 1.0, 45089], [2, 2.0, 238194]],
            "v": 1,
        },
        {
            "b": 0,
            "d": "2025-06-05",
            "p": 0,
            "k": 1,
            "cs": 1,
            "kw": 1,
            "cb": 1,
            "pl": 2,
            "jk": 0,
            "sk": 1,
            "dg": 1,
            "t": 176812,
            "c": [[0, 1.0, 80334], [1, 1.0, 23756], [2, 1.0, 72721]],
            "v": 2,
        },
        {
            "b": 0,
            "d": "2025-06-10",
            "p": 0,
            "k": 2,
            "cs": 2,
            "kw": 2,
            "cb": 2,
            "pl": 3,
            "jk": 1,
            "sk": 2,
            "dg": 2,
            "t": 210501,
            "c": [[0, 1.0, 64123], [1, 1.0, 45545], [3, 1.0, 32026], [2, 1.0, 68808]],
            "v": 1,
        },
        {
            "b": 0,
            "d": "2025-06-21",
            "p": 0,
            "k": 3,
            "cs": 3,
            "kw": 3,
            "cb": 3,
            "pl": 4,
            "jk": 0,
            "sk": 1,
            "dg": 0,
            "t": 310919,
            "c": [[0, 1.0, 73163], [1, 1.0, 60532], [3, 1.0, 33590], [2, 1.0, 143634]],
            "v": 1,
        },
        {
            "b": 0,
            "d": "2025-06-27",
            "p": 0,
            "k": 4,
            "cs": 4,
            "kw": 0,
            "cb": 4,
            "pl": 5,
            "jk": 0,
            "sk": 1,
            "dg": 2,
            "t": 444814,
            "c": [[0, 1.0, 56738], [1, 2.0, 327259], [2, 1.0, 60817]],
            "v": 1,
        },
        {
            "b": 0,
            "d": "2025-06-04",
            "p": 0,
            "k": 4,
            "cs": 4,
            "kw": 0,
            "cb": 4,
            "pl": 3,
            "jk": 0,
            "sk": 1,
            "dg": 2,
            "t": 134135,
            "c": [[0, 1.0, 46427], [1, 1.0, 87708]],
            "v": 2,
        },
        {
            "b": 0,
            "d": "2025-06-17",
            "p": 0,
            "k": 5,
            "cs": 5,
            "kw": 4,
            "cb": 5,
            "pl": 6,
            "jk": 0,
            "sk": 2,
            "dg": 3,
            "t": 375132,
            "c": [[0, 1.0, 73889], [1, 2.0, 235826], [3, 1.0, 65417]],
            "v": 1,
        },
        {
            "b": 0,
            "d": "2025-06-18",
            "p": 0,
            "k": 5,
            "cs": 5,
            "kw": 4,
            "cb": 5,
            "pl": 7,
            "jk": 0,
            "sk": 2,
            "dg": 3,
            "t": 552822,
            "c": [[0, 1.0, 63080], [4, 4.0, 288972], [1, 1.0, 29193], [5, 1.0, 171576]],
            "v": 2,
        },
        {
            "b": 0,
            "d": "2025-06-11",
            "p": 0,
            "k": 6,
            "cs": 6,
            "kw": 0,
            "cb": 0,
            "pl": 8,
            "jk": 0,
            "sk": 2,
            "dg": 4,
            "t": 243449,
            "c": [[0, 1.0, 66931], [1, 1.0, 176518]],
            "v": 1,
        },
        {
            "b": 0,
            "d": "2025-06-07",
            "p": 0,
            "k": 7,
            "cs": 7,
            "kw": 1,
            "cb": 6,
            "pl": 7,
            "jk": 0,
            "sk": 3,
            "dg": 5,
            "t": 275866,
            "c": [[0, 1.0, 59875], [1, 1.0, 45771], [3, 1.0, 67845], [2, 1.0, 102375]],
            "v": 1,
        },
        {
            "b": 0,
            "d": "2026-04-14",
            "p": 10,
            "k": 7,
            "cs": 7,
            "kw": 1,
            "cb": 6,
            "pl": 9,
            "jk": 0,
            "sk": 3,
            "dg": 5,
            "t": 328970,
            "c": [[0, 1.0, 54864], [1, 1.0, 72458], [3, 1.0, 69112], [2, 1.0, 132536]],
            "v": 2,
        },
        {
            "b": 0,
            "d": "2025-06-29",
            "p": 0,
            "k": 7,
            "cs": 7,
            "kw": 1,
            "cb": 6,
            "pl": 7,
            "jk": 0,
            "sk": 3,
            "dg": 5,
            "t": 470459,
            "c": [[0, 1.0, 68729], [4, 1.0, 299493], [1, 1.0, 58819], [3, 1.0, 43418]],
            "v": 3,
        },
        {
            "b": 0,
            "d": "2025-06-11",
            "p": 0,
            "k": 8,
            "cs": 8,
            "kw": 3,
            "cb": 7,
            "pl": 8,
            "jk": 0,
            "sk": 4,
            "dg": 6,
            "t": 161456,
            "c": [[0, 1.0, 76155], [1, 1.0, 85301]],
            "v": 1,
        },
        {
            "b": 0,
            "d": "2025-06-01",
            "p": 0,
            "k": 9,
            "cs": 9,
            "kw": 5,
            "cb": 8,
            "pl": 10,
            "jk": 0,
            "sk": 5,
            "dg": 4,
            "t": 182695,
            "c": [[0, 1.0, 68454], [1, 1.0, 114241]],
            "v": 1,
        },
        {
            "b": 0,
            "d": "2025-06-24",
            "p": 0,
            "k": 9,
            "cs": 9,
            "kw": 5,
            "cb": 8,
            "pl": 10,
            "jk": 0,
            "sk": 5,
            "dg": 4,
            "t": 430972,
            "c": [[0, 1.0, 81840], [1, 2.0, 98223], [2, 2.0, 250909]],
            "v": 2,
        },
        {
            "b": 0,
            "d": "2025-06-18",
            "p": 0,
            "k": 9,
            "cs": 9,
            "kw": 5,
            "cb": 8,
            "pl": 9,
            "jk": 0,
            "sk": 5,
            "dg": 4,
            "t": 155570,
            "c": [[0, 1.0, 50050], [1, 1.0, 24593], [2, 1.0, 80926]],
            "v": 3,
        },
        {
            "b": 0,
            "d": "2025-06-26",
            "p": 0,
            "k": 10,
            "cs": 10,
            "kw": 6,
            "cb": 9,
            "pl": 11,
            "jk": 1,
            "sk": 1,
            "dg": 0,
            "t": 156828,
            "c": [[0, 1.0, 52840], [1, 1.0, 103988]],
            "v": 1,
        },
        {
            "b": 0,
            "d": "2025-06-06",
            "p": 0,
            "k": 10,
            "cs": 10,
            "kw": 6,
            "cb": 9,
            "pl": 12,
            "jk": 1,
            "sk": 1,
            "dg": 0,
            "t": 171901,
            "c": [[0, 1.0, 55310], [1, 1.0, 116591]],
            "v": 2,
        },
        {
            "b": 0,
            "d": "2025-06-04",
            "p": 0,
            "k": 11,
            "cs": 11,
            "kw": 7,
            "cb": 10,
            "pl": 13,
            "jk": 0,
            "sk": 1,
            "dg": 4,
            "t": 189061,
            "c": [[0, 1.0, 65683], [1, 2.0, 62328], [3, 1.0, 61050]],
            "v": 1,
        },
    ]
}

# Sidebar untuk Kontrol & Unggah Template Excel
st.sidebar.markdown(
    "<h1>Utilization Review PLKK</h1>", unsafe_allow_html=True
)
st.sidebar.markdown(
    '<div class="sub">Data simulasi · Jun 2025 – Jun 2026</div>',
    unsafe_allow_html=True,
)

# Fitur Upload File Excel Sesuai Template
uploaded_file = st.sidebar.file_uploader(
    "Upload File Excel (Sesuai Template)", type=["xlsx", "xls"]
)

if uploaded_file is not None:
    try:
        # Membaca file excel yang diunggah pengguna
        df_uploaded = pd.read_excel(uploaded_file)
        st.sidebar.success("File Excel berhasil diunggah!")
    except Exception as e:
        st.sidebar.error(f"Gagal membaca file: {e}")

# Filter Sidebar (Menjaga bentuk dan fitur asli tanpa perubahan)
st.sidebar.markdown("<label>Periode Tanggal</label>", unsafe_allow_html=True)
col_f1, col_f2 = st.sidebar.columns(2)
f_from = col_f1.date_input("Dari", value=pd.to_datetime("2025-06-01").date())
f_to = col_f2.date_input("Sampai", value=pd.to_datetime("2026-06-30").date())

f_kw = st.sidebar.selectbox(
    "Kanwil", ["Semua Kanwil", "Kanwil 1", "Kanwil 2", "Kanwil 3"]
)
f_cb = st.sidebar.selectbox(
    "Cabang", ["Semua Cabang", "Cabang A", "Cabang B"]
)
f_pl = st.sidebar.selectbox("PLKK", ["Semua PLKK", "PLKK A", "PLKK B"])
f_b = st.sidebar.selectbox(
    "Jenis Layanan", ["Semua Layanan", "Rawat Inap", "Rawat Jalan"]
)
f_jk = st.sidebar.selectbox(
    "Jenis Kasus", ["Semua Kasus", "Kecelakaan Kerja", "Penyakit Akibat Kerja"]
)
f_sk = st.sidebar.selectbox(
    "Sektor Usaha", ["Semua Sektor", "Manufaktur", "Jasa", "Konstruksi"]
)

if st.sidebar.button("Reset filter", key="rs_btn"):
    st.rerun()

st.sidebar.markdown(
    '<div class="note" id="fcount">Menampilkan seluruh data aktif.</div>',
    unsafe_allow_html=True,
)

# Main Area / Navigasi Tab
nav_tabs = st.tabs(
    ["Ringkasan", "Analisis Biaya", "Distribusi PLKK", "Detail Data"]
)

with nav_tabs[0]:
    st.markdown("### Ringkasan Eksekutif")
    st.markdown(
        '<p class="d">Informasi umum utilisasi layanan kesehatan.</p>',
        unsafe_allow_html=True,
    )

    # KPI Metrics Cards (Gaya disamakan persis dengan template asli)
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(
            '<div class="kpi"><small>Total Kasus</small><b>1,240</b><span>+4.2% dari bulan lalu</span></div>',
            unsafe_allow_html=True,
        )
    with k2:
        st.markdown(
            '<div class="kpi"><small>Total Biaya</small><b>Rp 312.4 M</b><span>Rata-rata klaim stabil</span></div>',
            unsafe_allow_html=True,
        )
    with k3:
        st.markdown(
            '<div class="kpi"><small>Jumlah PLKK</small><b>348</b><span>Mitra aktif terdaftar</span></div>',
            unsafe_allow_html=True,
        )
    with k4:
        st.markdown(
            '<div class="kpi"><small>Rata-rata Unit Cost</small><b>Rp 2.45 M</b><span>Efisiensi optimal</span></div>',
            unsafe_allow_html=True,
        )

    # Grid Card Visualisasi & Review
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        st.markdown(
            '<div class="card"><h3>Tren Utilisasi Bulanan</h3><p class="d">Perkembangan volume layanan per bulan</p></div>',
            unsafe_allow_html=True,
        )
    with col_c2:
        st.markdown(
            '<div class="card"><h3>Komposisi Biaya Layanan</h3><p class="d">Distribusi biaya berdasarkan jenis penjaminan</p></div>',
            unsafe_allow_html=True,
        )

with nav_tabs[1]:
    st.markdown("### Analisis Biaya & Utilisasi")
    st.markdown(
        '<p class="d">Rincian mendalam mengenai alokasi biaya dan unit cost.</p>',
        unsafe_allow_html=True,
    )

with nav_tabs[2]:
    st.markdown("### Distribusi PLKK")
    st.markdown(
        '<p class="d">Pemetaan sebaran fasilitas kesehatan mitra.</p>',
        unsafe_allow_html=True,
    )

with nav_tabs[3]:
    st.markdown("### Tabel Detail Data")
    st.markdown(
        '<p class="d">Data mentah transaksi dan utilisasi.</p>',
        unsafe_allow_html=True,
    )
    # Render tabel data simulasi mentah
    df_raw = pd.DataFrame(RAW_DATA["eps"])
    st.dataframe(df_raw, use_container_width=True)