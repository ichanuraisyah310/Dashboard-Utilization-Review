import io
from pathlib import Path
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# ============================================================
# CONFIG
# ============================================================
st.set_page_config(
    page_title="UR Dashboard • PLKK",
    page_icon="📊",
    layout="wide",
)

st.markdown("""
<style>
.block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
.kpi {
    padding: 0.9rem 1rem;
    border: 1px solid #e5e7eb;
    border-radius: 12px;
    background: #ffffff;
    min-height: 95px;
}
.kpi .v {font-size: 1.35rem; font-weight: 750; line-height: 1.2;}
.kpi .l {font-size: .78rem; color: #667085; margin-top: .3rem;}
.section-title {font-size: 1.05rem; font-weight: 700; margin: .5rem 0 .7rem;}
.small-note {font-size: .78rem; color: #667085;}
</style>
""", unsafe_allow_html=True)

# ============================================================
# SERVICE MAP
# Posisi kolom mengikuti template BMIV yang diberikan.
# Index kolom adalah 0-based karena pandas membaca header=None.
# ============================================================
SERVICE_MAP = {
    "RJTP": [
        ("Dokter Umum", "kunjungan", "BMIV-01", [23], 24),
        ("Dokter Gigi (trauma)", "tindakan", "BMIV-01", [25], 31),
        ("Obat / Resep", "resep", "BMIV-01", [32, 33], 34),
        ("Penunjang Diagnostik Sederhana", "pemeriksaan", "BMIV-01", list(range(35, 40)), 40),
        ("Tindakan Medis Dokter Umum", "tindakan", "BMIV-01", list(range(41, 52)), 52),
        ("Tindakan Medis Dokter Gigi", "tindakan", "BMIV-01", list(range(53, 59)), 59),
        ("Vaksin & Profilaksis", "dosis", "BMIV-01", list(range(60, 64)), 64),
        ("Rujukan", "rujukan", "BMIV-01", list(range(65, 68)), None),
    ],
    "RJTL": [
        ("Dokter Spesialis", "kunjungan", "BMIV-02", [25], 26),
        ("Obat / Resep", "resep", "BMIV-02", [29], 30),
        ("Radiologi / Rontgen", "pemeriksaan", "BMIV-02", [33, 35], [34, 36]),
        ("Laboratorium", "pemeriksaan", "BMIV-02", [31], 32),
        ("Pemeriksaan Khusus / Elektromedik", "pemeriksaan", "BMIV-02", [37, 39, 41, 43, 45, 47, 49], [38, 40, 42, 44, 46, 48, 50]),
        ("Patologi Anatomi", "pemeriksaan", "BMIV-02", [51], 52),
        ("Tindakan Medis Spesialis", "tindakan", "BMIV-02", [54, 56], [55, 57]),
        ("Fisioterapi", "kunjungan", "BMIV-02", [58], 59),
        ("Rehabilitasi Medik Lain", "kunjungan", "BMIV-02", [60], 61),
        ("Emergensi", "kunjungan", "BMIV-02", [62], 63),
    ],
    "RANAP": [
        ("Akomodasi Rawat Inap", "kasus", "BMIV-03", [25], 27),
        ("Obat / Resep", "resep", "BMIV-03", [30], 32),
        ("Operasi", "tindakan", "BMIV-03", [33, 34, 35, 36], 37),
        ("Perawatan Khusus (ICU/HCU/Burn)", "kasus", "BMIV-03", [38], 40),
        ("Radiologi / Rontgen", "pemeriksaan", "BMIV-03", [43, 45], [44, 46]),
        ("Laboratorium", "pemeriksaan", "BMIV-03", [41], 42),
        ("Pemeriksaan Khusus / Elektromedik", "pemeriksaan", "BMIV-03", [47, 49, 51, 53, 55, 57, 59], [48, 50, 52, 54, 56, 58, 60]),
        ("Patologi Anatomi", "pemeriksaan", "BMIV-03", [61], 62),
        ("Tindakan Medis Spesialis", "tindakan", "BMIV-03", [64, 65], 66),
        ("Labu Darah", "labu", "BMIV-03", [67], 68),
        ("Transfusi Darah", "tindakan", "BMIV-03", [69], 70),
        ("Implan Ortopedi (Pin, Plate, Screw)", "kasus", "BMIV-03", [71], 72),
        ("Fisioterapi", "tindakan", "BMIV-03", [74], 75),
        ("Rehabilitasi Medik Lain", "tindakan", "BMIV-03", [76], 77),
    ],
    "KHUSUS": [
        ("Kaca Mata", "unit", "BMIV-04", [22], 23),
        ("Gigi Palsu", "unit", "BMIV-04", [24], 25),
        ("Protesis Anggota Gerak", "unit", "BMIV-04", [26], 27),
        ("Ortosis", "unit", "BMIV-04", [28], 29),
        ("Alat Bantu Jalan", "unit", "BMIV-04", [30], 31),
        ("Alat Bantu Dengar", "unit", "BMIV-04", [32], 33),
        ("Mata Palsu", "unit", "BMIV-04", [34], 35),
        ("Ambulans / Transportasi", "perjalanan", "BMIV-04", [36], 37),
        ("Program Kembali Bekerja (RTW)", "kasus", "BMIV-04", [38], 39),
    ],
}

BMIV_LABEL = {
    "RJTP": "BMIV-01",
    "RJTL": "BMIV-02",
    "RANAP": "BMIV-03",
    "KHUSUS": "BMIV-04",
}
GROUP_LABEL = {
    "RJTP": "BMIV-01 · RJTP",
    "RJTL": "BMIV-02 · RJTL",
    "RANAP": "BMIV-03 · RANAP",
    "KHUSUS": "BMIV-04 · Khusus",
}

# Kolom fallback dari template lama
BRANCH_COL = {"BMIV-01": 8, "BMIV-02": 8, "BMIV-03": 9, "BMIV-04": 6}
PLKK_COL = {"BMIV-01": 4, "BMIV-02": 4, "BMIV-03": 4, "BMIV-04": 11}
DATE_COL = {"BMIV-01": 2, "BMIV-02": 2, "BMIV-03": 2, "BMIV-04": 2}

MONTHS = {
    "Januari": 1, "Februari": 2, "Maret": 3, "April": 4,
    "Mei": 5, "Juni": 6, "Juli": 7, "Agustus": 8,
    "September": 9, "Oktober": 10, "November": 11, "Desember": 12,
}

def num(x):
    if pd.isna(x):
        return 0.0
    try:
        return float(str(x).replace(",", "").replace("Rp", "").strip())
    except Exception:
        return 0.0

# ============================================================
# STREAMLIT CACHE COMPATIBILITY
# ============================================================
_CACHE = getattr(st, "cache_data", None)
if _CACHE is None:
    _CACHE = getattr(st, "cache", None)

def cache_compat(func):
    if _CACHE is None:
        return func
    try:
        return _CACHE(show_spinner=False)(func)
    except TypeError:
        return _CACHE(func)

@cache_compat
def load_excel_data(file_bytes):
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    return {s: xl.parse(s, header=None) for s in xl.sheet_names}

# ============================================================
# HELPERS UNTUK MENCARI SHEET BMIV & PLOTLY AXIS AMAN
# ============================================================
def find_sheet(sheets, prefix):
    exact = [s for s in sheets if str(s).strip().upper() == prefix.upper()]
    if exact:
        return exact[0]
    candidates = [s for s in sheets if str(s).strip().upper().startswith(prefix.upper())]
    return candidates[0] if candidates else None

def locate_bmiv_sheets(sheets):
    found = {}
    missing = []
    for bmiv in ["BMIV-01", "BMIV-02", "BMIV-03", "BMIV-04"]:
        s = find_sheet(sheets, bmiv)
        if s is None:
            missing.append(bmiv)
        else:
            found[bmiv] = s
    return found, missing

def detect_header(df, max_rows=20):
    best_row = 0
    best_score = -1
    keywords = ["tanggal", "date", "plkk", "faskes", "biaya", "cost", "kanwil", "cabang", "branch", "kpj", "case"]
    for r in range(min(max_rows, len(df))):
        vals = " ".join(str(v).lower() for v in df.iloc[r].tolist() if pd.notna(v))
        score = sum(k in vals for k in keywords)
        if score > best_score:
            best_score = score
            best_row = r
    return best_row

def clean_sheet(df):
    h = detect_header(df)
    out = df.iloc[h + 1:].copy()
    out.columns = [str(x).strip() if pd.notna(x) else f"COL_{i}" for i, x in enumerate(df.iloc[h].tolist())]
    out = out.reset_index(drop=True)
    return out, h

def find_col(df, names):
    norm = {str(c).strip().lower(): c for c in df.columns}
    for name in names:
        if name.lower() in norm:
            return norm[name.lower()]
    for c in df.columns:
        lc = str(c).lower()
        if any(name.lower() in lc for name in names):
            return c
    return None

def get_series(df, col, default=""):
    if col is None or col not in df.columns:
        return pd.Series([default] * len(df), index=df.index)
    return df[col]

def axis_title(fig, axis_name):
    """Pemeriksaan aman untuk menghindari AttributeError pada layout Plotly."""
    if hasattr(fig, "layout") and hasattr(fig.layout, axis_name):
        ax_obj = getattr(fig.layout, axis_name)
        if ax_obj is not None:
            return getattr(ax_obj, "title", None)
    return None

def format_figure_units(fig):
    _ = axis_title(fig, "yaxis2")
    return fig

def prepare_sheet(df, bmiv):
    header_row = detect_header(df)
    header_values = [str(x).strip() if pd.notna(x) else f"COL_{i}"
                     for i, x in enumerate(df.iloc[header_row].tolist())]
    data = df.iloc[header_row + 1:].copy().reset_index(drop=True)
    data.columns = list(range(len(data.columns)))

    def header_col(names):
        for i, h in enumerate(header_values):
            lh = h.lower()
            if any(n.lower() == lh or n.lower() in lh for n in names):
                return i
        return None

    date_col = header_col(["tanggal", "tgl pelayanan", "date"])
    plkk_col = header_col(["plkk", "faskes", "nama plkk", "provider"])
    branch_col = header_col(["cabang", "branch", "kantor cabang", "wilayah"])
    kanwil_col = header_col(["kanwil", "kantor wilayah"])
    case_col = header_col(["case id", "case_id", "kasus", "id kasus"])
    kpj_col = header_col(["kpj", "peserta", "participant", "member"])
    case_type_col = header_col(["case type", "jenis kasus"])
    sector_col = header_col(["sektor usaha", "sector", "sektor"])
    service_type_col = header_col(["service type", "jenis layanan", "jenis pelayanan"])

    if date_col is None:
        date_col = DATE_COL[bmiv]
    if plkk_col is None:
        plkk_col = PLKK_COL[bmiv]
    if branch_col is None:
        branch_col = BRANCH_COL[bmiv]

    def col_series(idx):
        if idx is None or idx not in data.columns:
            return pd.Series([""] * len(data), index=data.index)
        return data[idx]

    data["_date"] = pd.to_datetime(col_series(date_col), errors="coerce")
    data["_plkk"] = col_series(plkk_col).astype(str).str.strip()
    data["_branch"] = col_series(branch_col).astype(str).str.strip()
    data["_kanwil"] = col_series(kanwil_col).astype(str).str.strip()
    data["_case"] = col_series(case_col).astype(str).str.strip()
    data["_kpj"] = col_series(kpj_col).astype(str).str.strip()
    data["_case_type"] = col_series(case_type_col).astype(str).str.strip()
    data["_sector"] = col_series(sector_col).astype(str).str.strip()
    data["_service_type"] = col_series(service_type_col).astype(str).str.strip()
    data["_bmiv"] = bmiv

    service_cols = set()
    for _, _, _, volcols, costcols in SERVICE_MAP[{
        "BMIV-01":"RJTP", "BMIV-02":"RJTL",
        "BMIV-03":"RANAP", "BMIV-04":"KHUSUS"
    }[bmiv]]:
        service_cols.update(volcols or [])
        if costcols is not None:
            service_cols.update(costcols if isinstance(costcols, list) else [costcols])
    for c in sorted(service_cols):
        if c in data.columns:
            data[f"__num_{c}"] = pd.to_numeric(data[c], errors="coerce").fillna(0.0)

    data["__service_volume_total"] = 0.0
    data["__service_cost_total"] = 0.0
    for _, _, _, volcols, costcols in SERVICE_MAP[{
        "BMIV-01":"RJTP", "BMIV-02":"RJTL",
        "BMIV-03":"RANAP", "BMIV-04":"KHUSUS"
    }[bmiv]]:
        vcols = [f"__num_{c}" for c in (volcols or []) if f"__num_{c}" in data.columns]
        if vcols:
            data["__service_volume_total"] += data[vcols].sum(axis=1)
        if costcols is not None:
            ccols = [f"__num_{c}" for c in (costcols if isinstance(costcols, list) else [costcols]) if f"__num_{c}" in data.columns]
            if ccols:
                data["__service_cost_total"] += data[ccols].sum(axis=1)

    data.attrs["header_values"] = header_values
    data.attrs["header_row"] = header_row
    return data, header_row

def value_sum_positional(df, cols):
    valid = [f"__num_{c}" for c in cols if f"__num_{c}" in df.columns]
    if valid:
        return float(df[valid].sum(axis=1).sum())
    raw_valid = [c for c in cols if c in df.columns]
    if not raw_valid:
        return 0.0
    return float(df[raw_valid].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1).sum())

def calc_service_table(prepared, group_filter="Semua"):
    rows = []
    groups = SERVICE_MAP if group_filter == "Semua" else {group_filter: SERVICE_MAP[group_filter]}
    for group, items in groups.items():
        sheet_name = prepared["sheet_names"].get(BMIV_LABEL[group])
        raw = prepared["raw"].get(BMIV_LABEL[group])
        if raw is None:
            continue
        raw = raw.copy()
        for name, unit, bmiv, volcols, costcols in items:
            vol = value_sum_positional(raw, volcols)
            cost = 0.0 if costcols is None else value_sum_positional(raw, costcols if isinstance(costcols, list) else [costcols])
            rows.append({
                "Kelompok": group,
                "BMIV": bmiv,
                "Komponen Layanan": name,
                "Satuan": unit,
                "Volume": vol,
                "Biaya (Rp)": cost,
                "Unit Cost (Rp)": cost / vol if vol > 0 else np.nan,
            })
    return pd.DataFrame(rows)

def filter_raw(prepared, bmiv_filter="Semua", plkk="Semua PLKK",
               kanwil="Semua Kanwil", branch="Semua Cabang",
               case_type="Semua Jenis Kasus", sector="Semua Sektor",
               service_type="Semua Jenis Layanan", start_date=None, end_date=None):
    frames = []
    bmivs = list(prepared["raw"].keys()) if bmiv_filter == "Semua BMIV" else [bmiv_filter]
    for bmiv in bmivs:
        d = prepared["raw"].get(bmiv)
        if d is None or d.empty:
            continue
        x = d
        if start_date is not None:
            x = x[x["_date"] >= pd.Timestamp(start_date)]
        if end_date is not None:
            x = x[x["_date"] <= pd.Timestamp(end_date) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)]
        if plkk != "Semua PLKK":
            x = x[x["_plkk"].eq(plkk)]
        if kanwil != "Semua Kanwil":
            x = x[x["_kanwil"].eq(kanwil)]
        if branch != "Semua Cabang":
            x = x[x["_branch"].eq(branch)]
        if case_type != "Semua Jenis Kasus":
            x = x[x["_case_type"].eq(case_type)]
        if sector != "Semua Sektor":
            x = x[x["_sector"].eq(sector)]
        if service_type != "Semua Jenis Layanan":
            x = x[x["_service_type"].eq(service_type)]
        frames.append(x)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)

def filtered_services(prepared, filters):
    rows = []
    groups = SERVICE_MAP if filters["group"] == "Semua" else {filters["group"]: SERVICE_MAP[filters["group"]]}
    for group, items in groups.items():
        bmiv = BMIV_LABEL[group]
        raw = prepared["raw"].get(bmiv)
        if raw is None:
            continue
        x = raw
        if filters["start_date"] is not None:
            x = x[x["_date"] >= pd.Timestamp(filters["start_date"])]
        if filters["end_date"] is not None:
            x = x[x["_date"] <= pd.Timestamp(filters["end_date"]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)]
        for colname, val in [
            ("_plkk", filters["plkk"]),
            ("_kanwil", filters["kanwil"]),
            ("_branch", filters["branch"]),
            ("_case_type", filters["case_type"]),
            ("_sector", filters["sector"]),
            ("_service_type", filters["service_type"]),
        ]:
            if val != "Semua " + {
                "_plkk":"PLKK", "_kanwil":"Kanwil", "_branch":"Cabang",
                "_case_type":"Jenis Kasus", "_sector":"Sektor", "_service_type":"Jenis Layanan"
            }[colname]:
                x = x[x[colname].eq(val)]
        for name, unit, _, volcols, costcols in items:
            vol = value_sum_positional(x, volcols)
            cost = 0.0 if costcols is None else value_sum_positional(x, costcols if isinstance(costcols, list) else [costcols])
            rows.append({
                "Kelompok": group, "BMIV": bmiv, "Komponen Layanan": name,
                "Satuan": unit, "Volume": vol, "Biaya (Rp)": cost,
                "Unit Cost (Rp)": cost / vol if vol > 0 else np.nan,
            })
    return pd.DataFrame(rows)

def build_service_by_month(filtered_data, start_date, end_date):
    months = pd.date_range(
        pd.Timestamp(start_date).replace(day=1),
        pd.Timestamp(end_date).replace(day=1),
        freq="MS",
    )
    if filtered_data is None or filtered_data.empty:
        return pd.DataFrame({
            "Bulan": months,
            "Biaya": 0.0,
            "Volume": 0.0,
            "Utilisasi": 0.0,
        })

    x = filtered_data.copy()
    x = x[x["_date"].notna()].copy()
    if x.empty:
        return pd.DataFrame({
            "Bulan": months,
            "Biaya": 0.0,
            "Volume": 0.0,
            "Utilisasi": 0.0,
        })

    x["Bulan"] = x["_date"].dt.to_period("M").dt.to_timestamp()
    agg = x.groupby("Bulan", as_index=False).agg(
        Biaya=("__service_cost_total", "sum"),
        Volume=("__service_volume_total", "sum"),
    )
    agg["Utilisasi"] = agg["Volume"]

    out = pd.DataFrame({"Bulan": months}).merge(agg, on="Bulan", how="left")
    for c in ["Biaya", "Volume", "Utilisasi"]:
        out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0.0)
    return out

def distinct_count(d, col):
    if d.empty or col not in d.columns:
        return 0
    s = d[col].replace({"": np.nan, "nan": np.nan}).dropna()
    return int(s.nunique())

def case_metrics(d):
    peserta = distinct_count(d, "_kpj")
    kasus = distinct_count(d, "_case")
    kunjungan = len(d)
    return peserta, kasus, kunjungan

def rupiah(x, decimals=0):
    if pd.isna(x):
        return "-"
    if decimals:
        return f"Rp {x:,.{decimals}f}"
    return f"Rp {x:,.0f}"

def kpi(col, label, value):
    col.markdown(
        f'<div class="kpi"><div class="v">{value}</div><div class="l">{label}</div></div>',
        unsafe_allow_html=True
    )

def calculate_plkk_performance(prepared, bmiv_rank, filters):
    x = prepared["raw"].get(bmiv_rank)
    if x is None or x.empty:
        return pd.DataFrame()

    mask = pd.Series(True, index=x.index)
    if filters["start_date"] is not None:
        mask &= x["_date"] >= pd.Timestamp(filters["start_date"])
    if filters["end_date"] is not None:
        mask &= x["_date"] <= pd.Timestamp(filters["end_date"]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    for colname, val, default in [
        ("_plkk", filters["plkk"], "Semua PLKK"),
        ("_kanwil", filters["kanwil"], "Semua Kanwil"),
        ("_branch", filters["branch"], "Semua Cabang"),
        ("_case_type", filters["case_type"], "Semua Jenis Kasus"),
        ("_sector", filters["sector"], "Semua Sektor"),
        ("_service_type", filters["service_type"], "Semua Jenis Layanan"),
    ]:
        if val != default:
            mask &= x[colname].eq(val)
    x = x.loc[mask]
    if x.empty:
        return pd.DataFrame()

    def nunique_valid(z):
        return z.replace({"": np.nan, "nan": np.nan}).nunique()

    base = x.groupby("_plkk", sort=False).agg(
        Kunjungan=("_plkk", "size"),
        Peserta=("_kpj", nunique_valid),
        Kasus=("_case", nunique_valid),
        Biaya=("__service_cost_total", "sum"),
        Utilisasi=("__service_volume_total", "sum"),
    ).reset_index().rename(columns={"_plkk": "PLKK"})
    base["Unit Cost"] = np.where(base["Utilisasi"] > 0, base["Biaya"] / base["Utilisasi"], np.nan)
    base["Cost per Case"] = np.where(base["Kasus"] > 0, base["Biaya"] / base["Kasus"], np.nan)
    return base

# ============================================================
# HEADER + UPLOAD
# ============================================================
st.title("UR Dashboard • PLKK")
st.caption("Monitoring biaya, utilisasi, unit cost, PMPM, kinerja PLKK, dan case-level analysis.")

with st.sidebar:
    st.header("📁 Upload Data")
    uploaded = st.file_uploader(
        "Upload Excel dengan format BMIV yang sama",
        type=["xlsx"],
        help="Nama file dan suffix sheet boleh berbeda. Struktur kolom BMIV harus tetap konsisten."
    )

if uploaded is None:
    st.info("Upload workbook Excel untuk mulai membuat dashboard.")
    st.stop()

try:
    file_bytes = uploaded.getvalue()
    dataset_key = hash(file_bytes)

    if st.session_state.get("_dataset_key") != dataset_key:
        sheets = load_excel_data(file_bytes)
        sheet_map, missing = locate_bmiv_sheets(sheets)
        if missing:
            st.error("Sheet BMIV berikut tidak ditemukan: " + ", ".join(missing))
            st.caption("Contoh yang diterima: BMIV-01, BMIV-01_Dummy, BMIV-01_Juni2026, dst.")
            st.stop()

        raw = {}
        header_info = {}
        for bmiv, sheet_name in sheet_map.items():
            raw[bmiv], header_info[bmiv] = prepare_sheet(sheets[sheet_name], bmiv)

        prepared = {"raw": raw, "sheet_names": sheet_map, "header_info": header_info}
        st.session_state["_dataset_key"] = dataset_key
        st.session_state["_prepared"] = prepared
        st.session_state["_sheets"] = sheets
        for k in ["_filter_cache", "_trend_cache", "_plkk_cache", "_all_data", "_all_data_key"]:
            st.session_state.pop(k, None)
    else:
        prepared = st.session_state["_prepared"]
        sheets = st.session_state["_sheets"]

except Exception as e:
    st.error(f"Gagal membaca workbook: {e}")
    st.exception(e)
    st.stop()

raw = prepared["raw"]

# ============================================================
# GLOBAL FILTER OPTIONS
# ============================================================
if st.session_state.get("_all_data_key") != st.session_state.get("_dataset_key"):
    st.session_state["_all_data"] = pd.concat(raw.values(), ignore_index=True)
    st.session_state["_all_data_key"] = st.session_state.get("_dataset_key")
all_data = st.session_state["_all_data"]
valid_dates = all_data["_date"].dropna()

if valid_dates.empty:
    min_date = pd.Timestamp.today().replace(day=1)
    max_date = min_date
else:
    min_date = valid_dates.min().normalize()
    max_date = valid_dates.max().normalize()

with st.sidebar:
    st.markdown("---")
    st.header("🎛️ Filter Dashboard")

    date_range = st.date_input(
        "Periode",
        value=(min_date.date(), max_date.date()),
        min_value=min_date.date(),
        max_value=max_date.date(),
    )
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
    else:
        start_date = end_date = date_range

    bmiv_options = ["Semua BMIV"] + ["BMIV-01", "BMIV-02", "BMIV-03", "BMIV-04"]
    selected_bmiv = st.selectbox("BMIV", bmiv_options)

    kanwil_values = sorted([x for x in all_data["_kanwil"].dropna().astype(str).unique() if x and x.lower() != "nan"])
    branch_values = sorted([x for x in all_data["_branch"].dropna().astype(str).unique() if x and x.lower() != "nan"])
    plkk_values = sorted([x for x in all_data["_plkk"].dropna().astype(str).unique() if x and x.lower() != "nan"])
    case_values = sorted([x for x in all_data["_case_type"].dropna().astype(str).unique() if x and x.lower() != "nan"])
    sector_values = sorted([x for x in all_data["_sector"].dropna().astype(str).unique() if x and x.lower() != "nan"])
    service_values = sorted([x for x in all_data["_service_type"].dropna().astype(str).unique() if x and x.lower() != "nan"])

    selected_kanwil = st.selectbox("Kanwil", ["Semua Kanwil"] + kanwil_values)
    selected_branch = st.selectbox("Cabang", ["Semua Cabang"] + branch_values)
    selected_plkk = st.selectbox("PLKK", ["Semua PLKK"] + plkk_values)
    selected_service_type = st.selectbox("Jenis Layanan", ["Semua Jenis Layanan"] + service_values)
    selected_case_type = st.selectbox("Jenis Kasus", ["Semua Jenis Kasus"] + case_values)
    selected_sector = st.selectbox("Sektor Usaha", ["Semua Sektor"] + sector_values)

filters = {
    "bmiv": selected_bmiv,
    "plkk": selected_plkk,
    "kanwil": selected_kanwil,
    "branch": selected_branch,
    "case_type": selected_case_type,
    "sector": selected_sector,
    "service_type": selected_service_type,
    "start_date": start_date,
    "end_date": end_date,
    "group": "Semua",
}

filter_key = (
    selected_bmiv, selected_plkk, selected_kanwil, selected_branch,
    selected_case_type, selected_sector, selected_service_type,
    str(start_date), str(end_date)
)

_filter_cache = st.session_state.setdefault("_filter_cache", {})
if filter_key in _filter_cache:
    filtered, svc = _filter_cache[filter_key]
else:
    filtered = filter_raw(
        prepared,
        bmiv_filter=selected_bmiv,
        plkk=selected_plkk,
        kanwil=selected_kanwil,
        branch=selected_branch,
        case_type=selected_case_type,
        sector=selected_sector,
        service_type=selected_service_type,
        start_date=start_date,
        end_date=end_date,
    )
    svc = filtered_services(prepared, filters)
    _filter_cache[filter_key] = (filtered, svc)
    while len(_filter_cache) > 12:
        _filter_cache.pop(next(iter(_filter_cache)))

peserta, kasus, kunjungan = case_metrics(filtered)
total_cost = float(svc["Biaya (Rp)"].sum()) if not svc.empty else 0.0
total_volume = float(svc["Volume"].sum()) if not svc.empty else 0.0
unit_cost = total_cost / total_volume if total_volume else np.nan

tk = max(peserta, 1)
pmpm = total_cost / tk / max((pd.Timestamp(end_date).to_period("M") - pd.Timestamp(start_date).to_period("M")).n + 1, 1)

# ============================================================
# NAVIGASI CEPAT — TAB KOTAK
# ============================================================
PAGES = [
    "1. Executive Summary",
    "2. Trend & Monitoring",
    "3. BMIV-01",
    "4. BMIV-02",
    "5. BMIV-03",
    "6. BMIV-04",
    "7. PLKK Performance",
    "8. LB-ST",
    "9. Unit Cost & Per Kapita",
    "10. Case Explorer",
]

if st.session_state.get("_active_page") not in PAGES:
    st.session_state["_active_page"] = PAGES[0]

st.markdown("""
<style>
div[data-testid="stButton"] > button {
    border-radius: 4px !important;
    min-height: 42px !important;
    padding: 0.35rem 0.45rem !important;
    font-size: 0.78rem !important;
    font-weight: 650 !important;
    white-space: normal !important;
    line-height: 1.15 !important;
}
</style>
""", unsafe_allow_html=True)

for row_start in range(0, len(PAGES), 5):
    row_pages = PAGES[row_start:row_start + 5]
    cols = st.columns(5)
    for i, page in enumerate(row_pages):
        with cols[i]:
            label = page.split(". ", 1)[1] if ". " in page else page
            if st.button(label, key=f"nav_{row_start}_{i}"):
                st.session_state["_active_page"] = page
                rerun_fn = getattr(st, "rerun", None) or getattr(st, "experimental_rerun", None)
                if rerun_fn is not None:
                    rerun_fn()

active_page = st.session_state["_active_page"]
st.markdown("---")

# ============================================================
# 1 EXECUTIVE SUMMARY
# ============================================================
if active_page == PAGES[0]:
    st.subheader("Executive Summary")
    st.caption(f"Periode {pd.Timestamp(start_date).strftime('%d %b %Y')} – {pd.Timestamp(end_date).strftime('%d %b %Y')}")

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    kpi(k1, "Peserta", f"{peserta:,.0f}")
    kpi(k2, "Kasus", f"{kasus:,.0f}")
    kpi(k3, "Kunjungan", f"{kunjungan:,.0f}")
    kpi(k4, "Biaya", rupiah(total_cost))
    kpi(k5, "Unit Cost", rupiah(unit_cost))
    kpi(k6, "PMPM", rupiah(pmpm, 2))

    st.markdown("")
    left, right = st.columns(2)

    with left:
        st.markdown('<div class="section-title">Komposisi Biaya per BMIV</div>', unsafe_allow_html=True)
        if svc.empty or svc["Biaya (Rp)"].sum() <= 0:
            st.info("Tidak ada data biaya pada filter.")
        else:
            cost_bmiv = svc.groupby("BMIV", as_index=False)["Biaya (Rp)"].sum()
            fig = px.pie(cost_bmiv, names="BMIV", values="Biaya (Rp)", hole=.48)
            fig.update_layout(height=380, margin=dict(l=10, r=10, t=20, b=10))
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with right:
        st.markdown('<div class="section-title">Top 10 Komponen berdasarkan Biaya</div>', unsafe_allow_html=True)
        top_cost = svc.sort_values("Biaya (Rp)", ascending=False).head(10)
        fig = px.bar(top_cost.sort_values("Biaya (Rp)"), x="Biaya (Rp)", y="Komponen Layanan", orientation="h")
        fig.update_layout(height=380, margin=dict(l=10, r=10, t=20, b=10), yaxis_title="")
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

# ============================================================
# 2 TREND & MONITORING
# ============================================================
if active_page == PAGES[1]:
    st.subheader("Trend & Monitoring")
    metric = st.radio("Tampilkan", ["Biaya", "Volume", "Utilisasi"], horizontal=True)
    trend_key = (filter_key, "trend")
    _trend_cache = st.session_state.setdefault("_trend_cache", {})
    if trend_key in _trend_cache:
        trend = _trend_cache[trend_key]
    else:
        trend = build_service_by_month(filtered, start_date, end_date)
        _trend_cache[trend_key] = trend
        while len(_trend_cache) > 12:
            _trend_cache.pop(next(iter(_trend_cache)))

    if trend.empty:
        st.info("Tidak ada data pada periode/filter yang dipilih.")
    else:
        ycol = "Biaya" if metric == "Biaya" else ("Volume" if metric == "Volume" else "Utilisasi")
        month_vals = trend["Bulan"].tolist()
        month_labels = [pd.Timestamp(x).strftime("%b %Y") for x in month_vals]

        fig = px.line(trend, x="Bulan", y=ycol, markers=True)
        fig.update_xaxes(
            tickmode="array",
            tickvals=month_vals,
            ticktext=month_labels,
            tickangle=-45,
            rangeslider_visible=False,
        )
        fig.update_layout(
            height=450,
            xaxis_title="Bulan",
            yaxis_title=metric,
            margin=dict(l=45, r=20, t=45, b=85),
        )
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

        fig2 = go.Figure()
        fig2.add_trace(go.Bar(
            x=month_vals, y=trend["Biaya"], name="Biaya", yaxis="y"
        ))
        fig2.add_trace(go.Scatter(
            x=month_vals, y=trend["Volume"], name="Volume/Utilisasi",
            mode="lines+markers", yaxis="y2"
        ))
        fig2.update_xaxes(
            tickmode="array",
            tickvals=month_vals,
            ticktext=month_labels,
            tickangle=-45,
        )
        fig2.update_layout(
            height=430,
            title="Biaya + Volume/Utilisasi per Bulan",
            xaxis_title="Bulan",
            yaxis=dict(title="Biaya (Rp)"),
            yaxis2=dict(title="Volume", overlaying="y", side="right"),
            legend=dict(orientation="h"),
            margin=dict(l=45, r=55, t=55, b=85),
        )
        st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar": False})

# ============================================================
# BMIV TABS
# ============================================================
def render_bmiv_tab(page_name, group, title, inpatient=False):
    if active_page != page_name:
        return
    st.subheader(title)
    sub = svc[svc["Kelompok"].eq(group)].copy()

    if sub.empty:
        st.info("Tidak ada data pada filter yang dipilih.")
        return

    c1, c2 = st.columns(2)

    with c1:
        util = sub.groupby("Komponen Layanan", as_index=False)["Volume"].sum().sort_values("Volume", ascending=False)
        fig = px.pie(util, names="Komponen Layanan", values="Volume", hole=.45,
                     title="Komposisi Utilisasi per Layanan")
        fig.update_layout(height=380)
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with c2:
        top = sub.sort_values("Volume", ascending=False).head(10).sort_values("Volume")
        fig = px.bar(top, x="Volume", y="Komponen Layanan", orientation="h",
                     title="Top 10 Layanan berdasarkan Volume")
        fig.update_layout(height=380, yaxis_title="")
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    c3, c4 = st.columns(2)

    with c3:
        cost = sub.groupby("Komponen Layanan", as_index=False)["Biaya (Rp)"].sum().sort_values("Biaya (Rp)", ascending=False)
        fig = px.bar(cost, x="Biaya (Rp)", y="Komponen Layanan", orientation="h",
                     title="Komposisi Biaya" + (" Rawat Inap" if inpatient else " per Layanan"))
        fig.update_layout(height=400, yaxis_title="")
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with c4:
        if inpatient:
            d = filtered[filtered["_bmiv"].eq("BMIV-03")].copy()
            alos_col = find_col(d, ["alos", "lama rawat", "length of stay", "los"])
            if alos_col is not None and not d.empty:
                alos = pd.to_numeric(d[alos_col], errors="coerce").dropna()
                if not alos.empty:
                    fig = px.histogram(alos, x=alos_col, nbins=15, title="Distribusi ALOS")
                    fig.update_layout(height=400)
                    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
                else:
                    st.info("Kolom ALOS tersedia tetapi tidak berisi angka.")
            else:
                st.info("Kolom ALOS/Lama Rawat tidak tersedia pada workbook ini.")

            cost_case = total_cost / kasus if kasus else np.nan
            st.metric("Cost per Case", rupiah(cost_case))
        else:
            scatter = sub[sub["Volume"] > 0].copy()
            fig = px.scatter(
                scatter, x="Volume", y="Unit Cost (Rp)",
                size="Biaya (Rp)", hover_name="Komponen Layanan",
                title="Volume vs Unit Cost"
            )
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

render_bmiv_tab(PAGES[2], "RJTP", "BMIV-01 · Rawat Jalan Tingkat Pertama")
render_bmiv_tab(PAGES[3], "RJTL", "BMIV-02 · Rawat Jalan Tingkat Lanjutan")
render_bmiv_tab(PAGES[4], "RANAP", "BMIV-03 · Rawat Inap", inpatient=True)
render_bmiv_tab(PAGES[5], "KHUSUS", "BMIV-04 · Pelayanan Khusus & Alat Bantu")

# ============================================================
# 7 PLKK PERFORMANCE
# ============================================================
if active_page == PAGES[6]:
    st.subheader("PLKK Performance")
    st.caption("Ranking PLKK mengikuti BMIV yang dipilih. Pilih satu BMIV untuk perbandingan PLKK yang apple-to-apple.")

    bmiv_rank = st.selectbox(
        "BMIV untuk Ranking PLKK",
        ["BMIV-01", "BMIV-02", "BMIV-03", "BMIV-04"],
        index=0 if selected_bmiv == "Semua BMIV" else ["BMIV-01","BMIV-02","BMIV-03","BMIV-04"].index(selected_bmiv),
    )
    plkk_key = (filter_key, bmiv_rank, "plkk")
    _plkk_cache = st.session_state.setdefault("_plkk_cache", {})
    if plkk_key in _plkk_cache:
        base = _plkk_cache[plkk_key]
    else:
        base = calculate_plkk_performance(prepared, bmiv_rank, filters)
        _plkk_cache[plkk_key] = base
        while len(_plkk_cache) > 12:
            _plkk_cache.pop(next(iter(_plkk_cache)))

    if base.empty:
        st.info("Tidak ada data PLKK pada filter.")
    else:
        a, b, c = st.columns(3)
        with a:
            top = base.nlargest(10, "Biaya").sort_values("Biaya")
            st.plotly_chart(px.bar(top, x="Biaya", y="PLKK", orientation="h", title="Top 10 PLKK berdasarkan Total Biaya"), use_container_width=True)
        with b:
            top = base.nlargest(10, "Utilisasi").sort_values("Utilisasi")
            st.plotly_chart(px.bar(top, x="Utilisasi", y="PLKK", orientation="h", title="Top 10 PLKK berdasarkan Utilisasi"), use_container_width=True)
        with c:
            top = base[base["Unit Cost"].notna()].nlargest(10, "Unit Cost").sort_values("Unit Cost")
            st.plotly_chart(px.bar(top, x="Unit Cost", y="PLKK", orientation="h", title="Top 10 PLKK berdasarkan Unit Cost"), use_container_width=True)

        topcase = base.nlargest(10, "Cost per Case").sort_values("Cost per Case")
        st.plotly_chart(px.bar(topcase, x="Cost per Case", y="PLKK", orientation="h", title="Cost per Case per PLKK"), use_container_width=True)

        st.dataframe(
            base.sort_values("Biaya", ascending=False).style.format({
                "Kunjungan": "{:,.0f}", "Utilisasi": "{:,.0f}", "Peserta": "{:,.0f}", "Kasus": "{:,.0f}",
                "Biaya": "Rp {:,.0f}", "Unit Cost": "Rp {:,.0f}", "Cost per Case": "Rp {:,.0f}"
            }),
            use_container_width=True, height=500
        )

# ============================================================
# 8 LB-ST
# ============================================================
if active_page == PAGES[7]:
    st.subheader("LB-ST")
    lbst_sheet = find_sheet(sheets, "LB-ST")
    if lbst_sheet:
        st.caption(f"Membaca sheet: {lbst_sheet}")
        st.dataframe(sheets[lbst_sheet], use_container_width=True, height=600)
    else:
        st.info("Sheet LB-ST tidak ditemukan. Ditampilkan rekap layanan aktual sebagai fallback.")
        st.dataframe(
            svc[["BMIV","Kelompok","Komponen Layanan","Satuan","Volume","Biaya (Rp)","Unit Cost (Rp)"]]
            .style.format({"Volume":"{:,.0f}", "Biaya (Rp)":"Rp {:,.0f}", "Unit Cost (Rp)":"Rp {:,.0f}"}),
            use_container_width=True, height=600
        )

# ============================================================
# 9 UNIT COST & PER KAPITA
# ============================================================
if active_page == PAGES[8]:
    st.subheader("Unit Cost & Per Kapita")
    uc = svc.copy()
    uc["Utilisasi / 1.000 TK"] = np.where(tk > 0, uc["Volume"] / tk * 1000, np.nan)
    uc["PMPM"] = np.where(tk > 0, uc["Biaya (Rp)"] / tk / max((pd.Timestamp(end_date).to_period("M") - pd.Timestamp(start_date).to_period("M")).n + 1,1), np.nan)
    uc["% Biaya"] = np.where(total_cost > 0, uc["Biaya (Rp)"] / total_cost, np.nan)
    uc.insert(0, "No", range(1, len(uc)+1))

    st.dataframe(
        uc[["No","BMIV","Kelompok","Komponen Layanan","Satuan","Volume","Biaya (Rp)","Unit Cost (Rp)","Utilisasi / 1.000 TK","PMPM","% Biaya"]]
        .style.format({
            "Volume":"{:,.0f}", "Biaya (Rp)":"Rp {:,.0f}", "Unit Cost (Rp)":"Rp {:,.0f}",
            "Utilisasi / 1.000 TK":"{:,.2f}", "PMPM":"Rp {:,.2f}", "% Biaya":"{:.1%}"
        }),
        use_container_width=True, height=550
    )

    c1, c2 = st.columns(2)
    with c1:
        topuc = uc[uc["Unit Cost (Rp)"].notna()].nlargest(10, "Unit Cost (Rp)").sort_values("Unit Cost (Rp)")
        st.plotly_chart(px.bar(topuc, x="Unit Cost (Rp)", y="Komponen Layanan", orientation="h", title="Unit Cost per Komponen"), use_container_width=True)
    with c2:
        scatter = uc[uc["Unit Cost (Rp)"].notna() & (uc["Volume"] > 0)]
        st.plotly_chart(px.scatter(scatter, x="Volume", y="Unit Cost (Rp)", size="Biaya (Rp)", hover_name="Komponen Layanan", title="Volume vs Unit Cost"), use_container_width=True)

# ============================================================
# 10 CASE EXPLORER
# ============================================================
if active_page == PAGES[9]:
    st.subheader("Case Explorer")
    if filtered.empty:
        st.info("Tidak ada case pada filter yang dipilih.")
    else:
        case_ids = sorted([x for x in filtered["_case"].dropna().astype(str).unique() if x and x.lower() != "nan"])
        if not case_ids:
            st.info("Kolom Case ID tidak tersedia/berisi kosong pada workbook ini.")
        else:
            selected_case = st.selectbox("Case ID", case_ids)
            case = filtered[filtered["_case"].eq(selected_case)].copy().sort_values("_date")

            case_cost = float(case["__service_cost_total"].sum()) if "__service_cost_total" in case.columns else 0.0

            c1,c2,c3,c4 = st.columns(4)
            kpi(c1, "Total Cost", rupiah(case_cost))
            kpi(c2, "Jumlah Layanan", f"{len(case):,.0f}")
            kpi(c3, "Peserta", f"{distinct_count(case,'_kpj'):,.0f}")
            if case["_date"].notna().any():
                alos_days = (case["_date"].max() - case["_date"].min()).days + 1
            else:
                alos_days = np.nan
            kpi(c4, "ALOS / Rentang Hari", f"{alos_days:,.0f}" if pd.notna(alos_days) else "-")

            st.markdown("### Timeline / Journey Layanan")
            timeline = case[["_date","_bmiv","_plkk"]].copy()
            timeline["Tanggal"] = timeline["_date"].dt.strftime("%d %b %Y")
            timeline = timeline.rename(columns={"_bmiv":"BMIV","_plkk":"PLKK"})
            st.dataframe(timeline, use_container_width=True, height=280)

            st.markdown("### Detail Record Case")
            display_cols = [c for c in case.columns if not str(c).startswith("_")]
            if not display_cols:
                display_cols = list(case.columns)
            st.dataframe(case[display_cols], use_container_width=True, height=400)