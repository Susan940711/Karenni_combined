# Karenni CHDN + KNA Combiner

This package combines CHDN and KNA report workbooks for these sheets only:
- indicators
- Td_ALOD
- ALOD_cummu
- IDP
- Td2_indicator

For the `indicators` summary sheet, it keeps township-level rows only (removes clinic rows) and does not append Karenni Total rows.

For other target sheets, it appends new rows in `Organization` as `Karenni Total` by summing reported numeric values from CHDN and KNA.

## Files
- `combine_karenni_reports.py` - command-line combiner
- `streamlit_karenni_combiner.py` - Streamlit web app
- `requirements.txt` - Python dependencies

## Setup
```powershell
pip install -r requirements.txt
```

## Run Streamlit App
```powershell
streamlit run streamlit_karenni_combiner.py
```

## Run CLI
```powershell
python combine_karenni_reports.py --chdn CHDN_report.xlsx --kna KNA_report.xlsx --output Karenni_combined.xlsx
```

## Notes
- Close source Excel files in Microsoft Excel before running.
- The Streamlit app allows upload and download directly, and can also save output to the current folder.
