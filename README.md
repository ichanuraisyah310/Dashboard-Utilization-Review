# Dashboard JKK PLKK — Python 3.7 Compatible

Project ini sengaja memakai versi library yang masih kompatibel dengan Python 3.7.x.

## Jalankan di VS Code

Buka Terminal di folder project, lalu:

```powershell
python -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
streamlit run app.py
```

Jika PowerShell memblokir aktivasi environment, gunakan langsung:

```powershell
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe -m streamlit run app.py
```

Upload `DATA TIRUAN.xlsx` pada dashboard.

## Output
- LB-ST dihitung dari BMIV-01 s.d. BMIV-04
- Unit Cost setempat dan gabungan
- Utilisasi per 1.000 TK per bulan
- PMPM / unit cost per kapita
- Grafik dan tabel detail
- Download hasil Excel

Kode cabang setempat default: `K00`.
