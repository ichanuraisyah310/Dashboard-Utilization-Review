"""
UR Dashboard • PLKK  (Streamlit)
Dashboard kosong sampai workbook Excel berformat BMIV diunggah.
Jalankan:  streamlit run app.py
"""
import io

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="UR Dashboard • PLKK", page_icon="📊", layout="wide")

# ----------------------------------------------------------------------------
# Konstanta
# ----------------------------------------------------------------------------
BS = ["RJTP", "RJTL", "RANAP", "Khusus"]
BN = ["BMIV-01 · RJTP", "BMIV-02 · RJTL", "BMIV-03 · RANAP", "BMIV-04 · Khusus"]
BC = ["#2a9d8f", "#3b82f6", "#e9a23b", "#a855f7"]
BD = [
    "Rawat Jalan Tingkat Pertama — dokter umum, gigi, obat, penunjang diagnostik, tindakan medis.",
    "Rawat Jalan Tingkat Lanjutan — dokter spesialis, laboratorium, radiologi, fisioterapi, ODC, emergensi.",
    "Rawat Inap — akomodasi, operasi, ICU, transfusi, implan ortopedi, penunjang, lengkap dengan distribusi ALOS.",
    "Pelayanan khusus & alat bantu — kaca mata, gigi palsu, protesis, ambulans, program RTW.",
]
MON = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]

# (kode, nama, kolom biaya, kolom volume) — posisi kolom (0-based) pada template BMIV.
# Data pada template bergeser dari header di beberapa blok, sehingga dipetakan per posisi data.
S1 = [("RJTP-01", "Dokter Umum", [23], [22]), ("RJTP-02", "Tindakan Gigi", [30], range(24, 30)),
      ("RJTP-03", "Obat/Resep", [32, 34, 36], [31, 33, 35]), ("RJTP-04", "Penunjang Diagnostik", [42], range(37, 42)),
      ("RJTP-05", "Tindakan Medis Dokter Umum", [54], range(43, 54)),
      ("RJTP-06", "Tindakan Medis Dokter Gigi", [61], range(55, 61))]
S2 = [("RJTL-01", "Jasa Dokter Spesialis", [25], [24]), ("RJTL-02", "Obat KS", [27], [26]),
      ("RJTL-03", "Obat LB", [29], [28]), ("RJTL-04", "Obat RX", [31], [30]),
      ("RJTL-05", "Laboratorium", [33], [32]), ("RJTL-06", "Radiologi (Rontgen)", [35], [34]),
      ("RJTL-07", "EKG", [37], [36]), ("RJTL-08", "CT Scan", [39], [38]), ("RJTL-09", "MRI", [41], [40]),
      ("RJTL-10", "USG", [43], [42]), ("RJTL-11", "EMG/NCV", [45], [44]), ("RJTL-12", "Elektromedik Lain", [47], [46]),
      ("RJTL-13", "Tindakan Spesialis", [51], [50]), ("RJTL-14", "ODC", [53], [52]),
      ("RJTL-15", "Fisioterapi", [55], [54]), ("RJTL-16", "Rehab Medik Lain", [57], [56]),
      ("RJTL-17", "Emergensi", [59], [58]), ("RJTL-18", "Ambulans", [61], [60]), ("RJTL-19", "Homecare", [63], [62])]
S3 = [("RI-01", "Akomodasi Rawat Inap", [26], [25]), ("RI-02", "Obat/Resep", [31], [27, 28, 29]),
      ("RI-03", "Tindakan Non Operatif", [33], [32]), ("RI-04", "Operasi Kecil", [36], [35]),
      ("RI-05", "Operasi Sedang", [38], [37]), ("RI-06", "Operasi Besar", [40], [39]),
      ("RI-07", "Operasi Khusus", [42], [41]), ("RI-08", "ICU/ICCU/HCU/Burn Unit", [45], [44]),
      ("RI-09", "Lab", [47], [46]), ("RI-10", "Rontgen", [49], [48]), ("RI-11", "EKG", [51], [50]),
      ("RI-12", "CT Scan", [53], [52]), ("RI-13", "MRI", [55], [54]), ("RI-14", "USG", [57], [56]),
      ("RI-15", "EMG/NCV", [59], [58]), ("RI-16", "Elektromedik Lain", [61], [60]),
      ("RI-17", "Pelayanan Darah", [65], [64]), ("RI-18", "Transfusi Darah", [67], [66]),
      ("RI-19", "Implan Ortopedi", [69], [68]), ("RI-20", "Fisioterapi", [72], [71]),
      ("RI-21", "Rehab Medik Lain", [74], [73]), ("RI-22", "Ambulans", [76], [75]), ("RI-23", "Homecare", [78], [77])]
S4 = [("AB-01", "Kacamata"), ("AB-02", "Gigi Palsu"), ("AB-03", "Protesis Anggota Gerak"), ("AB-04", "Ortosis"),
      ("AB-05", "Alat Bantu Jalan"), ("AB-06", "Alat Bantu Dengar"), ("AB-07", "Mata Palsu"),
      ("AB-08", "Ambulans/Transportasi"), ("AB-09", "Program Kembali Bekerja (RTW)")]

SECTOR_KEYS = [
    ("pendidikan", "Jasa Pendidikan"), ("perdagangan", "Perdagangan & Ritel"), ("niaga", "Perdagangan & Ritel"),
    ("mart", "Perdagangan & Ritel"), ("retail", "Perdagangan & Ritel"), ("angkutan", "Transportasi & Logistik"),
    ("logistik", "Transportasi & Logistik"), ("ekspedisi", "Transportasi & Logistik"),
    ("kargo", "Transportasi & Logistik"), ("terminal", "Transportasi & Logistik"),
    ("pergudangan", "Transportasi & Logistik"), ("infrastruktur", "Konstruksi"), ("konstruksi", "Konstruksi"),
    ("teknik sipil", "Konstruksi"), ("pondasi", "Konstruksi"), ("struktur", "Konstruksi"),
    ("kertas", "Industri Manufaktur"), ("plastik", "Industri Manufaktur"), ("kimia", "Industri Manufaktur"),
    ("logam", "Industri Manufaktur"), ("makanan", "Industri Manufaktur"), ("cleaning", "Jasa Penunjang"),
    ("sekuriti", "Jasa Penunjang"), ("sawit", "Pertanian & Perkebunan"), ("pertanian", "Pertanian & Perkebunan"),
    ("hotel", "Perhotelan"), ("finansial", "Keuangan"), ("migas", "Migas & Pertambangan"),
    ("teknologi", "Teknologi Informasi")]


# ----------------------------------------------------------------------------
# Helper format (gaya Indonesia)
# ----------------------------------------------------------------------------
def _id(s: str) -> str:
    return s.replace(",", "§").replace(".", ",").replace("§", ".")


def fN(n):
    return _id(f"{n:,.0f}")


def f1(n):
    return _id(f"{n:,.1f}")


def fF(n):
    return "Rp " + fN(n)


def fR(n):
    a = abs(n)
    if a >= 1e9:
        return "Rp " + _id(f"{n / 1e9:.2f}") + " M"
    if a >= 1e6:
        return "Rp " + _id(f"{n / 1e6:.1f}") + " jt"
    return fF(n)


def fP(x):
    return f1(x * 100) + "%"


def ml(p):
    return f"{MON[int(p[5:7]) - 1]} {p[2:4]}"


def sector(name: str) -> str:
    n = str(name).lower()
    for k, v in SECTOR_KEYS:
        if k in n:
            return v
    return "Lainnya"


def N(v):
    try:
        return 0.0 if pd.isna(v) else float(v)
    except Exception:
        return 0.0


# ----------------------------------------------------------------------------
# ETL: workbook BMIV -> tabel episode + tabel komponen
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load(file_bytes: bytes):
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    names = xl.sheet_names
    warns = []
    cab = {}
    norm = next((s for s in names if s.lower().startswith("dashboard_raw")), None)
    if norm:
        d = pd.read_excel(xl, norm, usecols=["KODE CAB", "Kantor Wilayah"]).drop_duplicates("KODE CAB")
        cab = d.set_index("KODE CAB")["Kantor Wilayah"].to_dict()

    eps, comps = [], []
    for b in range(4):
        sh = next((s for s in names if s.upper().startswith(f"BMIV-0{b + 1}")), None)
        if sh is None:
            warns.append(f"Sheet BMIV-0{b + 1} tidak ditemukan — dilewati.")
            continue
        r = pd.read_excel(xl, sh)
        cols = list(r.columns)
        A = r.values
        g = lambda c: cols.index(c) if c in cols else None  # noqa: E731
        pc = 10 if b == 3 else 4
        tc = {0: 66, 1: 64, 2: 80, 3: 39}[b]
        for i in range(len(r)):
            row = A[i]
            if pd.isna(row[2]):
                continue
            c = []
            if b < 3:
                spec = (S1, S2, S3)[b]
                for code, nm, cc, vc in spec:
                    cost = sum(N(row[j]) for j in cc)
                    vol = sum(N(row[j]) for j in vc)
                    if b == 1 and code == "RJTL-19":
                        others = sum(N(row[j]) for s in S2[:-1] for j in s[2])
                        cost = max(0.0, N(row[63]) - others) if vol > 0 else 0.0
                    if cost > 0 or vol > 0:
                        c.append([code, nm, vol, cost])
            else:
                tot, oth = N(row[39]), 0.0
                for k, (code, nm) in enumerate(S4[:8]):
                    v, cost = N(row[22 + 2 * k]), N(row[23 + 2 * k])
                    oth += cost
                    if v > 0 or cost > 0:
                        c.append([code, nm, v, cost])
                if N(row[38]) > 0:
                    c.append(["AB-09", S4[8][1], N(row[38]), max(0.0, tot - oth)])
            approved = N(row[tc])
            csum = sum(z[3] for z in c)
            if b < 3 and csum > 0:
                f = approved / csum
                for z in c:
                    z[3] = round(z[3] * f)
            elif b < 3 and approved > 0 and c:
                for z in c:
                    z[3] = round(approved / len(c))
            eid = len(eps)
            gv = lambda name, default="": row[g(name)] if g(name) is not None else default  # noqa: E731
            tgl = pd.Timestamp(row[2])
            kode = gv("KODE CAB")
            eps.append(dict(
                id=eid, b=b, tgl=tgl, periode=tgl.strftime("%Y-%m"),
                kpj=str(int(N(gv("NO KPJ")))), case=str(gv("NO. KASUS KK")),
                kanwil=str(gv("KANWIL")), cabang=str(cab.get(kode, kode)), plkk=str(row[pc]),
                jenis_kasus=str(gv("JENIS KASUS")), sektor=sector(gv("NAMA PERUSAHAAN")),
                diagnosa=str(gv("DIAGNOSA")), total=round(approved),
                los=int(N(row[25])) if b == 2 else 0, kelas=str(gv("KELAS RAWAT")) if b == 2 else "",
                poli=str(gv("POLI / SPESIALISASI")) if b == 1 else "",
                kunj_ke=int(N(gv("KUNJUNGAN KE-", 0))) if b in (0, 1) else 0))
            for code, nm, v, cost in c:
                comps.append((eid, b, code, nm, v, cost))
    if not eps:
        return None, None, warns + ["Tidak ada data BMIV yang terbaca. Pastikan format workbook sama dengan template."]
    E = pd.DataFrame(eps)
    C = pd.DataFrame(comps, columns=["id", "b", "kode", "komponen", "vol", "biaya"])
    C["label"] = C.komponen + " · " + C.b.map(lambda i: BS[i])
    return E, C, warns


# ----------------------------------------------------------------------------
# Agregasi
# ----------------------------------------------------------------------------
def agg(L, m):
    pes, kas, kun, bi = L.kpj.nunique(), L["case"].nunique(), len(L), float(L.total.sum())
    return dict(pes=pes, kas=kas, kun=kun, bi=bi, uc=bi / kun if kun else 0, pm=bi / (pes * m) if pes else 0,
                cpc=bi / kas if kas else 0, ut=kun / pes * 1000 if pes else 0)


def hbar(df, x, y, color, title=None, fmt="rp"):
    d = df.iloc[::-1]
    fig = go.Figure(go.Bar(x=d[x], y=d[y], orientation="h", marker_color=color,
                           hovertemplate="%{y}<br>" + ("Rp %{x:,.0f}" if fmt == "rp" else "%{x:,.0f}") + "<extra></extra>"))
    fig.update_layout(height=max(300, 30 * len(d) + 80), margin=dict(l=0, r=10, t=30 if title else 10, b=0),
                      title=title, yaxis=dict(automargin=True))
    st.plotly_chart(fig, width="stretch")


NUM = st.column_config.NumberColumn
RP = lambda label: NUM(label, format="localized")  # noqa: E731


# ----------------------------------------------------------------------------
# Views
# ----------------------------------------------------------------------------
def view_exec(L, C, m):
    a = agg(L, m)
    cols = st.columns(6)
    cols[0].metric("Total Peserta", fN(a["pes"]), help="NO KPJ unik")
    cols[1].metric("Kasus", fN(a["kas"]), help="No. Kasus KK unik")
    cols[2].metric("Kunjungan", fN(a["kun"]), help="Jumlah episode layanan")
    cols[3].metric("Total Biaya", fR(a["bi"]), help=fF(a["bi"]))
    cols[4].metric("Unit Cost", fF(a["uc"]), help="Biaya / kunjungan")
    cols[5].metric("PMPM", fF(a["pm"]), help=f"Biaya / peserta / bulan ({m} bulan)")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Komposisi Biaya per Kelompok Pelayanan (BMIV)")
        bb = L.groupby("b").total.sum().reindex(range(4), fill_value=0)
        fig = go.Figure(go.Pie(labels=BN, values=bb.values, hole=.6, marker_colors=BC, sort=False,
                               hovertemplate="%{label}<br>Rp %{value:,.0f} (%{percent})<extra></extra>"))
        fig.update_layout(height=400, margin=dict(t=10, b=0), legend=dict(orientation="h", y=-0.05))
        st.plotly_chart(fig, width="stretch")
    with c2:
        st.subheader("Top 10 Komponen Layanan (penyerapan biaya)")
        t = C.groupby("label").biaya.sum().sort_values(ascending=False).head(10).reset_index()
        hbar(t, "biaya", "label", "#0f766e")
    st.caption("PMPM = total biaya ÷ (peserta unik yang melapor klaim × jumlah bulan pada rentang filter) — proxy, "
               "karena jumlah tenaga kerja terdaftar tidak tersedia. Biaya komponen dialokasikan proporsional "
               "terhadap Total Disetujui tiap baris.")


def view_trend(L, pers):
    sel = st.radio("Metrik", ["Biaya", "Volume (Kunjungan)", "Utilisasi (kunjungan per 1.000 peserta)"], horizontal=True)
    rows = []
    for p in pers:
        X = L[L.periode == p]
        pes = X.kpj.nunique()
        d = dict(periode=p, Bulan=ml(p), pes=pes, kas=X["case"].nunique(), kun=len(X), bi=float(X.total.sum()))
        d["ut"] = d["kun"] / pes * 1000 if pes else 0
        for b in range(4):
            Xb = X[X.b == b]
            d[f"bi{b}"], d[f"kun{b}"] = float(Xb.total.sum()), len(Xb)
            d[f"ut{b}"] = len(Xb) / pes * 1000 if pes else 0
        rows.append(d)
    T = pd.DataFrame(rows)
    labs = list(T.Bulan)
    fig = go.Figure()
    if sel.startswith("Utilisasi"):
        fig.add_trace(go.Scatter(x=labs, y=T.ut, name="Total", line=dict(width=4, color="#64748b")))
        for b in range(4):
            fig.add_trace(go.Scatter(x=labs, y=T[f"ut{b}"], name=BN[b], line=dict(color=BC[b])))
        fig.update_yaxes(title="kunjungan / 1.000 peserta")
    else:
        key = "bi" if sel == "Biaya" else "kun"
        for b in range(4):
            fig.add_trace(go.Bar(x=labs, y=T[f"{key}{b}"], name=BN[b], marker_color=BC[b]))
        fig.update_layout(barmode="stack")
    fig.update_xaxes(categoryorder="array", categoryarray=labs)
    fig.update_layout(height=420, margin=dict(t=10), legend=dict(orientation="h", y=-0.15))
    st.plotly_chart(fig, width="stretch")
    T["mom"] = T.bi.pct_change().replace([np.inf, -np.inf], np.nan)
    T["uc"] = (T.bi / T.kun.replace(0, np.nan)).fillna(0)
    st.subheader("Rekap Bulanan")
    st.dataframe(T[["Bulan", "pes", "kas", "kun", "bi", "uc", "ut", "mom"]], hide_index=True, width="stretch",
                 column_config={"pes": RP("Peserta"), "kas": RP("Kasus"), "kun": RP("Kunjungan"), "bi": RP("Biaya"),
                                "uc": RP("Unit Cost"), "ut": NUM("Util /1.000", format="%.1f"),
                                "mom": NUM("Δ Biaya MoM", format="percent")})


def view_bmiv(L, C, b, m):
    st.markdown(f"**{BN[b]}** — {BD[b]}")
    Lb = L[L.b == b]
    if Lb.empty:
        st.info("Tidak ada data untuk kombinasi filter ini.")
        return
    Cb = C[C.b == b]
    a = agg(Lb, m)
    cols = st.columns(6 if b == 2 else 5)
    cols[0].metric("Total Biaya", fR(a["bi"]), help=fF(a["bi"]))
    cols[1].metric("Kunjungan / Episode", fN(a["kun"]))
    cols[2].metric("Kasus", fN(a["kas"]))
    cols[3].metric("Peserta", fN(a["pes"]))
    cols[4].metric("Unit Cost", fF(a["uc"]))
    if b == 2:
        cols[5].metric("ALOS", f"{f1(Lb.los.mean())} hari", help=f"Total {fN(Lb.los.sum())} hari rawat")
    G = Cb.groupby(["kode", "komponen"], as_index=False).agg(vol=("vol", "sum"), biaya=("biaya", "sum"))
    G = G.sort_values("biaya", ascending=False)
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Biaya per Komponen")
        hbar(G.head(15), "biaya", "komponen", BC[b])
    with c2:
        if b == 2:
            st.subheader("Distribusi ALOS (hari rawat)")
            bins = [0, 3, 7, 10, 14, 10_000]
            lab = ["1–3 hr", "4–7 hr", "8–10 hr", "11–14 hr", "15+ hr"]
            cnt = pd.cut(Lb.los, bins, labels=lab).value_counts().reindex(lab, fill_value=0)
            fig = go.Figure(go.Bar(x=lab, y=cnt.values, marker_color=BC[2]))
            fig.update_layout(height=380, margin=dict(t=10), yaxis_title="Kasus")
            st.plotly_chart(fig, width="stretch")
        elif b == 1:
            st.subheader("Kunjungan per Poli/Spesialisasi")
            t = Lb.poli.value_counts().head(10).rename_axis("poli").reset_index(name="n")
            hbar(t, "n", "poli", BC[1], fmt="n")
        else:
            st.subheader("Top Diagnosa (biaya)")
            t = Lb.groupby("diagnosa").total.sum().sort_values(ascending=False).head(10).reset_index()
            hbar(t, "total", "diagnosa", BC[b])
    st.subheader("Rincian Komponen")
    G["pct"] = G.biaya / G.biaya.sum()
    G["uc"] = (G.biaya / G.vol.replace(0, np.nan)).fillna(0)
    st.dataframe(G.rename(columns={"kode": "Kode", "komponen": "Komponen"}), hide_index=True, width="stretch",
                 column_config={"vol": RP("Volume"), "biaya": RP("Biaya"), "pct": NUM("% Biaya", format="percent"),
                                "uc": RP("Unit Cost")})
    if b == 2:
        st.subheader("ALOS per Kelas Rawat")
        K = Lb.groupby("kelas").agg(kasus=("id", "count"), hari=("los", "sum"), biaya=("total", "sum")).reset_index()
        K["alos"] = K.hari / K.kasus
        K["bpk"] = K.biaya / K.kasus
        st.dataframe(K, hide_index=True, width="stretch",
                     column_config={"kelas": "Kelas", "kasus": RP("Kasus"), "hari": RP("Hari Rawat"),
                                    "alos": NUM("ALOS (hari)", format="%.1f"), "biaya": RP("Biaya"),
                                    "bpk": RP("Biaya / Kasus")})


def view_plkk(L, m):
    c1, c2 = st.columns(2)
    basis = c1.selectbox("Basis perbandingan (apple-to-apple)", ["Semua layanan"] + BN)
    metric = c2.selectbox("Metrik peringkat", ["Total Biaya", "Kunjungan (utilisasi)", "Unit Cost", "Cost per Case"])
    X = L if basis == "Semua layanan" else L[L.b == BN.index(basis)]
    if X.empty:
        st.info("Tidak ada data.")
        return
    R = X.groupby("plkk").agg(pes=("kpj", "nunique"), kas=("case", "nunique"), kun=("id", "count"),
                              bi=("total", "sum")).reset_index()
    R["uc"], R["cpc"] = R.bi / R.kun, R.bi / R.kas
    R["ut"] = R.kun / R.pes * 1000
    key = {"Total Biaya": "bi", "Kunjungan (utilisasi)": "kun", "Unit Cost": "uc", "Cost per Case": "cpc"}[metric]
    T = R[R.kun >= 5] if key in ("uc", "cpc") else R
    T = T.sort_values(key, ascending=False).head(15)
    st.subheader(f"Peringkat PLKK (Top 15) — {metric}")
    st.caption("Bandingkan PLKK pada level layanan yang sama. Metrik rasio hanya untuk PLKK dengan ≥5 kunjungan.")
    hbar(T, key, "plkk", "#0f766e", fmt="n" if key == "kun" else "rp")
    st.subheader(f"Rekapitulasi PLKK ({len(R)} PLKK)")
    st.dataframe(R.sort_values("bi", ascending=False), hide_index=True, width="stretch",
                 column_config={"plkk": "PLKK", "pes": RP("Peserta"), "kas": RP("Kasus"), "kun": RP("Kunjungan"),
                                "bi": RP("Total Biaya"), "uc": RP("Unit Cost"), "cpc": RP("Cost per Case"),
                                "ut": NUM("Util /1.000", format="%.1f")})


def view_lbst(L, C, pers):
    st.caption("Sheet LB-ST tidak tersedia pada berkas; ditampilkan rekap layanan aktual per komponen × bulan "
               "dari BMIV-01 s.d. BMIV-04 (fallback).")
    mode = st.radio("Tampilkan", ["Biaya (Rp)", "Volume"], horizontal=True)
    val = "biaya" if mode.startswith("Biaya") else "vol"
    X = C.merge(L[["id", "periode"]], on="id")
    P = X.pivot_table(index=["b", "komponen"], columns="periode", values=val, aggfunc="sum", fill_value=0)
    P = P.reindex(columns=pers, fill_value=0)
    P["Total"] = P.sum(axis=1)
    P = P.reset_index()
    P.insert(0, "Layanan", P.b.map(lambda i: BS[i]))
    P = P.drop(columns="b").rename(columns={"komponen": "Komponen", **{p: ml(p) for p in pers}})
    tot = P.drop(columns=["Layanan", "Komponen"]).sum()
    P = pd.concat([P, pd.DataFrame([{"Layanan": "Total", "Komponen": "", **tot.to_dict()}])], ignore_index=True)
    st.dataframe(P, hide_index=True, width="stretch",
                 column_config={c: NUM(c, format="localized") for c in P.columns[2:]})


def view_uc(L, C, m):
    basis = st.selectbox("Kelompok", ["Semua BMIV"] + BN)
    X = L if basis == "Semua BMIV" else L[L.b == BN.index(basis)]
    Cx = C[C.id.isin(X.id)]
    a = agg(X, m)
    k = st.columns(4)
    k[0].metric("Unit Cost", fF(a["uc"]), help="Biaya / kunjungan")
    k[1].metric("Cost per Case", fF(a["cpc"]))
    k[2].metric("PMPM", fF(a["pm"]), help=f"{m} bulan")
    k[3].metric("Utilisasi /1.000 peserta", f1(a["ut"]), help="Kunjungan per 1.000 peserta")
    G = Cx.groupby(["b", "komponen", "label"], as_index=False).agg(vol=("vol", "sum"), biaya=("biaya", "sum"))
    G["pct"] = G.biaya / G.biaya.sum() if G.biaya.sum() else 0
    G["uc"] = (G.biaya / G.vol.replace(0, np.nan)).fillna(0)
    G["ut"] = G.vol / a["pes"] * 1000 if a["pes"] else 0
    G["pm"] = G.biaya / (a["pes"] * m) if a["pes"] else 0
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Top 10 PMPM per Komponen")
        hbar(G.sort_values("pm", ascending=False).head(10), "pm", "label", "#0f766e")
    with c2:
        st.subheader("Top 10 Unit Cost per Komponen")
        st.caption("Komponen dengan volume ≥3")
        hbar(G[G.vol >= 3].sort_values("uc", ascending=False).head(10), "uc", "label", "#a855f7")
    st.subheader("Unit Cost, Utilisasi & PMPM per Komponen")
    out = G.sort_values("biaya", ascending=False).copy()
    out["Layanan"] = out.b.map(lambda i: BS[i])
    st.dataframe(out[["Layanan", "komponen", "vol", "biaya", "pct", "uc", "ut", "pm"]], hide_index=True,
                 width="stretch",
                 column_config={"komponen": "Komponen", "vol": RP("Volume"), "biaya": RP("Biaya"),
                                "pct": NUM("% Proporsi", format="percent"), "uc": RP("Unit Cost"),
                                "ut": NUM("Util /1.000", format="%.1f"), "pm": RP("PMPM")})


def view_case(E, C):
    st.caption("Case Explorer menelusuri seluruh data (tidak terpengaruh filter sidebar).")
    top = E.groupby("case").total.sum().sort_values(ascending=False).head(12)
    cases = sorted(E["case"].unique())
    sel = st.selectbox("Case ID (ketik untuk mencari)", cases, index=None, placeholder="mis. " + cases[0])
    if sel is None:
        st.markdown("**Kasus berbiaya tertinggi:**")
        st.dataframe(top.rename("Total Biaya").reset_index().rename(columns={"case": "Case ID"}), hide_index=True,
                     column_config={"Total Biaya": RP("Total Biaya")})
        return
    X = E[E["case"] == sel].sort_values(["tgl", "b"])
    cx = C[C.id.isin(X.id)]
    days = (X.tgl.max() - X.tgl.min()).days
    k = st.columns(6)
    k[0].metric("Total Biaya", fF(X.total.sum()))
    k[1].metric("Layanan / Episode", len(X))
    k[2].metric("PLKK Terlibat", X.plkk.nunique())
    k[3].metric("Rentang", f"{days} hari", help=f"{X.tgl.min():%Y-%m-%d} → {X.tgl.max():%Y-%m-%d}")
    k[4].metric("Jenis Kasus", X.jenis_kasus.iloc[0], help=X.sektor.iloc[0])
    k[5].metric("Diagnosa", X.diagnosa.iloc[0], help="KPJ " + X.kpj.iloc[0])

    def cn(eid):
        z = cx[cx.id == eid]
        return ", ".join(f"{r.komponen}{' ×' + fN(r.vol) if r.vol > 1 else ''}" for r in z.itertuples())

    st.subheader("Timeline / Journey Layanan")
    html = ""
    for e in X.itertuples():
        extra = f" · {e.los} hari rawat" if e.b == 2 else ""
        html += (f"<div style='border-left:4px solid {BC[e.b]};padding:2px 0 2px 12px;margin:0 0 12px 6px'>"
                 f"<span style='background:{BC[e.b]};color:#fff;border-radius:99px;padding:1px 9px;font-size:12px'>{BS[e.b]}</span> "
                 f"<b>{e.tgl:%Y-%m-%d}</b> · {e.plkk}<br><span style='opacity:.7'>{cn(e.id) or '–'}{extra}</span><br>"
                 f"<b>{fF(e.total)}</b></div>")
    st.markdown(html, unsafe_allow_html=True)
    st.subheader("Detail Baris Data")
    D = pd.DataFrame({"Tanggal": X.tgl.dt.strftime("%Y-%m-%d"), "Layanan": X.b.map(lambda i: BN[i]), "PLKK": X.plkk,
                      "Kunjungan ke": X.kunj_ke.replace(0, np.nan), "Komponen": [cn(i) for i in X.id],
                      "Biaya Disetujui": X.total})
    st.dataframe(D, hide_index=True, width="stretch", column_config={"Biaya Disetujui": RP("Biaya Disetujui")})


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------
st.title("UR Dashboard • PLKK")
st.caption("Monitoring biaya, utilisasi, unit cost, PMPM, kinerja PLKK, dan case-level analysis.")

st.sidebar.markdown("## 📁 Upload Data")
up = st.sidebar.file_uploader("Upload Excel dengan format BMIV yang sama", type=["xlsx"])
if up is None:
    st.info("Upload workbook Excel untuk mulai membuat dashboard.")
    st.stop()

with st.spinner("Membaca & memetakan workbook BMIV…"):
    E, C, warns = load(up.getvalue())
for w in warns:
    st.warning(w)
if E is None:
    st.stop()

st.sidebar.markdown("---\n## 🔎 Filter Global")
dmin, dmax = E.tgl.min().date(), E.tgl.max().date()
rng = st.sidebar.date_input("Periode Tanggal", (dmin, dmax), min_value=dmin, max_value=dmax)
if isinstance(rng, (tuple, list)) and len(rng) == 2:
    d0, d1 = rng
else:
    d0 = d1 = rng[0] if isinstance(rng, (tuple, list)) else rng
mask = (E.tgl.dt.date >= d0) & (E.tgl.dt.date <= d1)
for label, col, opts in [("Kanwil", "kanwil", None), ("Cabang", "cabang", None), ("PLKK", "plkk", None),
                         ("Jenis Layanan", "b", BN), ("Jenis Kasus", "jenis_kasus", None),
                         ("Sektor Usaha", "sektor", None)]:
    if opts:
        pick = st.sidebar.multiselect(label, range(4), format_func=lambda i: BN[i])
    else:
        pick = st.sidebar.multiselect(label, sorted(E[col].unique()))
    if pick:
        mask &= E[col].isin(pick)
L = E[mask]
Cf = C[C.id.isin(L.id)]
st.sidebar.caption(f"{fN(len(L))} dari {fN(len(E))} baris terfilter")

fm, tm = f"{d0:%Y-%m}", f"{d1:%Y-%m}"
pers = [p for p in sorted(E.periode.unique()) if fm <= p <= tm]
m = max(1, len(pers))

tabs = st.tabs(["Executive Summary", "Tren Bulanan", "BMIV-01 RJTP", "BMIV-02 RJTL", "BMIV-03 RANAP",
                "BMIV-04 Khusus", "PLKK Performance", "LB-ST", "Unit Cost & Per Kapita", "Case Explorer"])
with tabs[9]:
    view_case(E, C)
if L.empty:
    for i in range(9):
        with tabs[i]:
            st.info("Tidak ada data untuk kombinasi filter ini. Ubah atau reset filter.")
    st.stop()
with tabs[0]:
    view_exec(L, Cf, m)
with tabs[1]:
    view_trend(L, pers)
for b in range(4):
    with tabs[2 + b]:
        view_bmiv(L, Cf, b, m)
with tabs[6]:
    view_plkk(L, m)
with tabs[7]:
    view_lbst(L, Cf, pers)
with tabs[8]:
    view_uc(L, Cf, m)
