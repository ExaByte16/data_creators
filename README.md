## Oxnya Balance General – Streamlit App

This app reproduces the provided notebook logic with a UI to upload the SIIGO Excel export, enter MES/ESTADO/AÑO/CENTRO DE COSTOS, and download two Excel outputs:
- datos_balance_general.xlsx
- datos_estado_resultados.xlsx

### Run locally

1. Create a virtual environment (optional but recommended)
```bash
python3 -m venv .venv && source .venv/bin/activate
```

2. Install dependencies
```bash
pip install -r requirements.txt
```

3. Start the app
```bash
streamlit run streamlit_app.py
```

### Usage
- Upload the SIIGO Excel export (.xlsx). The header must be on row 8 (header=7).
- Fill MES, ESTADO, AÑO, and CENTRO DE COSTOS.
- Click "Procesar y Generar Archivos".
- Download both Excel files.

### Notes
- If the column `Nombre tercero` is not present in the export, the app creates it blank so the `TERCERO` column exists in outputs, matching the notebook note that sometimes the name cannot be determined.
