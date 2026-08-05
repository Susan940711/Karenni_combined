from __future__ import annotations

from datetime import datetime
from io import BytesIO
from pathlib import Path
import tempfile

import pandas as pd
import streamlit as st

from combine_karenni_reports import (
    AGE_SEMESTER_SHEET_NAME,
    AT_LEAST_ONE_SEMESTER_SHEET_NAME,
    IDP_SEMESTER_SHEET_NAME,
    SEMESTER_REPORT_SHEET_NAME,
    TARGET_SHEETS,
    build_age_semester_from_sheet_map,
    build_at_least_one_semester_from_alod,
    build_idp_semester_from_sheet_map,
    build_semester_report_from_sheet_map,
    combine_sheet,
    write_sheet_with_aliases,
)


BASE_DIR = Path(__file__).resolve().parent


def build_combined_workbook(chdn_path: Path, kna_path: Path) -> tuple[bytes, dict[str, pd.DataFrame]]:
    sheet_map: dict[str, pd.DataFrame] = {}
    for canonical, aliases in TARGET_SHEETS.items():
        sheet_map[canonical] = combine_sheet(chdn_path, kna_path, canonical, aliases)

    if "indicators" in sheet_map:
        sheet_map[SEMESTER_REPORT_SHEET_NAME] = build_semester_report_from_sheet_map(sheet_map)
        sheet_map[AGE_SEMESTER_SHEET_NAME] = build_age_semester_from_sheet_map(sheet_map)
    if "IDP" in sheet_map:
        sheet_map[IDP_SEMESTER_SHEET_NAME] = build_idp_semester_from_sheet_map(sheet_map)
    if "ALOD_cummu" in sheet_map:
        sheet_map[AT_LEAST_ONE_SEMESTER_SHEET_NAME] = build_at_least_one_semester_from_alod(sheet_map["ALOD_cummu"])

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, df in sheet_map.items():
            write_sheet_with_aliases(writer, sheet_name, df)

    return buffer.getvalue(), sheet_map


def save_uploaded_file(uploaded_file, prefix: str) -> Path:
    suffix = Path(uploaded_file.name).suffix if uploaded_file.name else ".xlsx"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix, prefix=prefix) as tmp:
        tmp.write(uploaded_file.getbuffer())
        return Path(tmp.name)


def main() -> None:
    st.set_page_config(page_title="Karenni CHDN + KNA Combiner", layout="wide")

    st.title("Karenni Combination Builder")
    st.caption(
        "Combine CHDN and KNA report sheets. Summary uses township rollup with S1/S2/Annual periods; "
        "indicators and the other sheets use standard combine logic with appended Karenni Total rows; "
        "semester report is built from indicators quarterly values."
    )

    with st.expander("Target sheets", expanded=True):
        st.write("This app combines and outputs these sheets only:")
        st.write(", ".join(TARGET_SHEETS.keys()))

    left, right = st.columns(2)
    with left:
        chdn_upload = st.file_uploader("Upload CHDN report workbook", type=["xlsx"], key="chdn")
    with right:
        kna_upload = st.file_uploader("Upload KNA long workbook", type=["xlsx"], key="kna")

    default_output_name = f"Karenni_combined_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    output_name = st.text_input("Output filename", value=default_output_name)

    run_clicked = st.button("Generate Combined Workbook", type="primary", use_container_width=True)

    if not run_clicked:
        return

    if chdn_upload is None or kna_upload is None:
        st.error("Please upload both CHDN and KNA Excel files.")
        return

    temp_paths: list[Path] = []
    try:
        chdn_path = save_uploaded_file(chdn_upload, "chdn_")
        kna_path = save_uploaded_file(kna_upload, "kna_")
        temp_paths.extend([chdn_path, kna_path])

        workbook_bytes, sheets = build_combined_workbook(chdn_path, kna_path)

        save_to_folder = st.checkbox(
            "Also save output to Karenni_combination folder",
            value=True,
            help="If enabled, this writes a copy beside the app/script files.",
        )

        safe_name = output_name.strip() or default_output_name
        if not safe_name.lower().endswith(".xlsx"):
            safe_name = f"{safe_name}.xlsx"

        if save_to_folder:
            output_path = BASE_DIR / safe_name
            with open(output_path, "wb") as f:
                f.write(workbook_bytes)
            st.success(f"Workbook saved to: {output_path}")
        else:
            st.success("Workbook generated successfully.")

        st.download_button(
            label="Download Combined Workbook",
            data=workbook_bytes,
            file_name=safe_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        st.subheader("Preview")
        preview_sheet = st.selectbox("Choose a sheet to preview", list(sheets.keys()))
        st.dataframe(sheets[preview_sheet], use_container_width=True, height=420)

    except Exception as exc:  # noqa: BLE001
        st.error(f"Failed to combine files: {exc}")
        st.info("If the source files are open in Excel, close them and try again.")
    finally:
        for path in temp_paths:
            try:
                path.unlink(missing_ok=True)
            except Exception:
                pass


if __name__ == "__main__":
    main()
