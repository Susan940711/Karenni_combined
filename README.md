# Karenni CHDN + KNA Combiner

This package combines CHDN and KNA report workbooks for these sheets only:
- Summary
- indicators
- Td_ALOD
- ALOD_cummu
- IDP
- Td2_indicator
- semester report

For the `Summary` sheet, it combines clinic-level rows into township totals, converts quarterly periods to S1/S2/Annual, removes the clinic column, and does not append Karenni Total rows.

For `indicators` and other target sheets, it applies the standard combine logic and appends Karenni Total rows.

For `semester report`, it is generated from the combined `indicators` sheet by rolling up quarterly periods into S1, S2, and Annual columns:
- S1 Target, S1 Male, S1 Female, S1 Total
- S2 Target, S2 Male, S2 Female, S2 Total
- Annual Target, Annual Male, Annual Female, Annual Total

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
