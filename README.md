# Dashboard Utilization Review PLKK (Streamlit)

## Menjalankan di VS Code
```bash
python -m venv .venv
.venv\Scripts\activate          # Windows   (macOS/Linux: source .venv/bin/activate)
pip install -r requirements.txt
streamlit run app.py
```
Browser terbuka di http://localhost:8501. Upload file Excel lewat sidebar.

## Format file yang diterima
Workbook `.xlsx` dengan sheet **BMIV-01, BMIV-02, BMIV-03, BMIV-04** (boleh berakhiran, mis. `_Dummy`)
dan nama kolom sama persis dengan `Data_Dummy.xlsx`. Sheet lain (README, Column_Mapping, dst.) diabaikan.
Jika tidak ada file yang di-upload, app memakai `Data_Dummy.xlsx` di folder yang sama.
