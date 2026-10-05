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
- `combine_karenni_reports.py` - complete Streamlit app and workbook processing logic
- `requirements.txt` - Python dependencies

## Setup
```powershell
pip install -r requirements.txt
```

## Run Locally
```powershell
streamlit run combine_karenni_reports.py
```

For Streamlit Community Cloud, set **Main file path** to `combine_karenni_reports.py`.

## Notes
- Upload both `.xlsx` source workbooks, then select **Generate combined workbook**.
- The app processes uploads in memory and downloads the result directly; it does not write generated files to the deployment filesystem.
