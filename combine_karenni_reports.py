from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path
import re

import pandas as pd


TARGET_SHEETS: dict[str, list[str]] = {
    "Summary": ["Summary", "summary"],
    "indicators": ["indicators", "indicator"],
    "Td_ALOD": ["Td_ALOD", "Td ALOD", "Td_alod", "TD_ALOD"],
    "ALOD_cummu": ["ALOD_cummu", "ALOD cummu"],
    "IDP": ["IDP", "idp"],
    "Td2_indicator": ["Td2_indicator", "TD2_indicator", "Td2 indicator"],
}

SEMESTER_REPORT_SHEET_NAME = "semester report"


def normalize_name(value: str) -> str:
    return "".join(ch for ch in str(value).lower() if ch.isalnum())


def resolve_sheet_name(workbook: pd.ExcelFile, aliases: list[str]) -> str:
    wanted = {normalize_name(name) for name in aliases}
    for sheet in workbook.sheet_names:
        if normalize_name(sheet) in wanted:
            return sheet
    raise KeyError(
        f"Could not find sheet matching any of {aliases} in {Path(workbook.io).name}. "
        f"Available: {workbook.sheet_names}"
    )


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename_map: dict[str, str] = {}
    for col in df.columns:
        if isinstance(col, str):
            cleaned = re.sub(r"\s+", " ", col.strip())
            rename_map[col] = cleaned
    return df.rename(columns=rename_map)


def align_columns(primary: pd.DataFrame, secondary: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    ordered = list(primary.columns)
    for col in secondary.columns:
        if col not in ordered:
            ordered.append(col)
    return primary.reindex(columns=ordered), secondary.reindex(columns=ordered)


def is_empty_value(value: object) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def canonicalize_alod_label(value: object) -> str:
    text = "" if is_empty_value(value) else str(value).strip()
    norm = normalize_name(text)

    if not norm:
        return ""

    # Treat CHDN variants like "ALOD cummulative" as the same indicator label.
    if "alod" in norm and ("cummu" in norm or "cumu" in norm or "cumul" in norm):
        return "At least one dose under 5-yr-old"

    if "atleastonedoseunder5" in norm:
        return "At least one dose under 5-yr-old"

    return text


def harmonize_alod_cummu_indicator_columns(df: pd.DataFrame) -> pd.DataFrame:
    has_indicator_upper = "Indicator" in df.columns
    has_indicator_lower = "indicator" in df.columns

    if not has_indicator_upper and not has_indicator_lower:
        return df

    if not has_indicator_upper:
        df["Indicator"] = ""
    if not has_indicator_lower:
        df["indicator"] = ""

    def resolve_row_label(row: pd.Series) -> str:
        upper_val = row.get("Indicator", "")
        lower_val = row.get("indicator", "")

        if not is_empty_value(upper_val):
            return canonicalize_alod_label(upper_val)
        if not is_empty_value(lower_val):
            return canonicalize_alod_label(lower_val)
        return ""

    merged_label = df.apply(resolve_row_label, axis=1)
    df["Indicator"] = merged_label
    df["indicator"] = merged_label
    return df


def cleaned_numeric(series: pd.Series) -> pd.Series:
    as_str = series.astype(str).str.replace(",", "", regex=False).str.strip()
    as_str = as_str.where(~series.isna(), other="")
    as_str = as_str.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
    return pd.to_numeric(as_str, errors="coerce")


def find_column_by_token(df: pd.DataFrame, token: str) -> str | None:
    for col in df.columns:
        if token in normalize_name(col):
            return col
    return None


def find_metric_column(df: pd.DataFrame, tokens: list[str], banned_tokens: list[str] | None = None) -> str | None:
    banned_tokens = banned_tokens or []
    for col in df.columns:
        norm = normalize_name(col)
        if any(token in norm for token in tokens) and not any(bad in norm for bad in banned_tokens):
            return col
    return None


def canonical_quarter_label(value: object) -> str | None:
    if is_empty_value(value):
        return None

    norm = normalize_name(str(value))
    if not norm:
        return None

    q1_tokens = ["q1", "quarter1", "qtr1", "firstquarter", "1stquarter", "janmar"]
    q2_tokens = ["q2", "quarter2", "qtr2", "secondquarter", "2ndquarter", "aprjun"]
    q3_tokens = ["q3", "quarter3", "qtr3", "thirdquarter", "3rdquarter", "julsep"]
    q4_tokens = ["q4", "quarter4", "qtr4", "fourthquarter", "4thquarter", "octdec"]

    if norm in {"1", "01"} or any(token in norm for token in q1_tokens):
        return "Q1"
    if norm in {"2", "02"} or any(token in norm for token in q2_tokens):
        return "Q2"
    if norm in {"3", "03"} or any(token in norm for token in q3_tokens):
        return "Q3"
    if norm in {"4", "04"} or any(token in norm for token in q4_tokens):
        return "Q4"
    return None


def combine_summary_clinic_rows_to_township(df: pd.DataFrame) -> pd.DataFrame:
    township_col = find_column_by_token(df, "township")
    clinic_col = find_column_by_token(df, "clinic")
    period_col = find_column_by_token(df, "period")
    organization_col = next((c for c in df.columns if normalize_name(c) == "organization"), None)

    working = df.copy()
    if working.empty:
        return working.drop(columns=[clinic_col], errors="ignore")

    dimension_cols, numeric_cols = detect_dimension_columns(working)
    # Merge CHDN and KNA into one township-period result by excluding organization.
    excluded_cols = {clinic_col, period_col, organization_col}
    group_cols = [col for col in dimension_cols if col not in excluded_cols]

    if township_col is not None and township_col not in group_cols:
        group_cols.append(township_col)

    if clinic_col is not None:
        working = working.drop(columns=[clinic_col])

    if not numeric_cols:
        return working.reindex(columns=[col for col in df.columns if col != clinic_col]).reset_index(drop=True)

    for col in numeric_cols:
        working[col] = cleaned_numeric(working[col])

    def grouped_sum(frame: pd.DataFrame, keys: list[str]) -> pd.DataFrame:
        if not keys:
            sums = frame[numeric_cols].sum(min_count=1).fillna(0)
            return pd.DataFrame([sums.to_dict()])
        return (
            frame.groupby(keys, dropna=False, as_index=False, sort=False)[numeric_cols]
            .sum(min_count=1)
            .fillna(0)
        )

    if period_col is None:
        grouped = grouped_sum(working, group_cols)
        if organization_col is not None and organization_col in df.columns:
            grouped[organization_col] = "Karenni Total"
        output_columns = [col for col in df.columns if col != clinic_col]
        return grouped.reindex(columns=output_columns).reset_index(drop=True)

    working["__quarter"] = working[period_col].apply(canonical_quarter_label)
    quarterly = working.loc[working["__quarter"].notna()].copy()

    if quarterly.empty:
        output_columns = [col for col in df.columns if col != clinic_col]
        return working.drop(columns=["__quarter"], errors="ignore").reindex(columns=output_columns).reset_index(drop=True)

    def rollup_period(periods: set[str], output_period: str) -> pd.DataFrame:
        subset = quarterly.loc[quarterly["__quarter"].isin(periods)]
        if subset.empty:
            return pd.DataFrame(columns=group_cols + numeric_cols + [period_col])

        out = grouped_sum(subset, group_cols)
        out[period_col] = output_period
        return out

    summary_frames = [
        rollup_period({"Q1", "Q2"}, "S1"),
        rollup_period({"Q3", "Q4"}, "S2"),
        rollup_period({"Q1", "Q2", "Q3", "Q4"}, "Annual"),
    ]
    grouped = pd.concat(summary_frames, ignore_index=True)
    if organization_col is not None and organization_col in df.columns:
        grouped[organization_col] = "Karenni Total"

    twp_mimu_col = find_column_by_token(grouped, "twpmimu")
    year_col = find_column_by_token(grouped, "year")
    if twp_mimu_col is not None and period_col in grouped.columns:
        dedupe_keys = [twp_mimu_col, period_col]
        if year_col is not None and year_col not in dedupe_keys:
            dedupe_keys.append(year_col)

        numeric_in_grouped = [c for c in numeric_cols if c in grouped.columns]
        non_numeric_cols = [c for c in grouped.columns if c not in dedupe_keys + numeric_in_grouped]

        agg_map: dict[str, str] = {col: "sum" for col in numeric_in_grouped}
        for col in non_numeric_cols:
            agg_map[col] = "first"

        grouped = grouped.groupby(dedupe_keys, dropna=False, as_index=False, sort=False).agg(agg_map)
        if organization_col is not None and organization_col in grouped.columns:
            grouped[organization_col] = "Karenni Total"

    output_columns = [col for col in df.columns if col != clinic_col]
    return grouped.reindex(columns=output_columns).reset_index(drop=True)


def detect_dimension_columns(df: pd.DataFrame) -> tuple[list[str], list[str]]:
    organization_col = next((c for c in df.columns if normalize_name(c) == "organization"), "Organization")
    protected_tokens = [
        "year",
        "period",
        "organization",
        "project",
        "indicator",
        "district",
        "township",
        "clinic",
        "camp",
        "site",
        "state",
        "region",
    ]

    dimension_cols: list[str] = []
    numeric_cols: list[str] = []

    for col in df.columns:
        if col == organization_col:
            continue

        norm = normalize_name(col)
        if any(token in norm for token in protected_tokens):
            dimension_cols.append(col)
            continue

        series = df[col]
        if pd.api.types.is_numeric_dtype(series):
            numeric_cols.append(col)
            continue

        numeric_try = cleaned_numeric(series)
        non_empty = series.dropna().astype(str).str.strip()
        non_empty = non_empty[non_empty != ""]

        # Treat a column as numeric only if every non-empty value can be parsed.
        if len(non_empty) > 0 and int(numeric_try.notna().sum()) == len(non_empty):
            numeric_cols.append(col)
        else:
            dimension_cols.append(col)

    return dimension_cols, numeric_cols


def append_karenni_total_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    organization_col = next((c for c in df.columns if normalize_name(c) == "organization"), "Organization")
    if organization_col not in df.columns:
        raise KeyError("Organization column not found in source sheet.")

    dimension_cols, numeric_cols = detect_dimension_columns(df)
    if not numeric_cols:
        return df

    working = df.copy()
    for col in numeric_cols:
        working[col] = cleaned_numeric(working[col])

    grouped = (
        working.groupby(dimension_cols, dropna=False, as_index=False)[numeric_cols]
        .sum(min_count=1)
        .fillna(0)
    )
    grouped[organization_col] = "Karenni Total"
    grouped = grouped.reindex(columns=df.columns)

    return pd.concat([df, grouped], ignore_index=True)


def aggregate_by_keys(df: pd.DataFrame, keys: list[str], value_cols: list[str]) -> pd.DataFrame:
    if not value_cols:
        return pd.DataFrame(columns=keys)

    if not keys:
        sums = df[value_cols].sum(min_count=1).fillna(0)
        return pd.DataFrame([sums.to_dict()])

    return (
        df.groupby(keys, dropna=False, as_index=False, sort=False)[value_cols]
        .sum(min_count=1)
        .fillna(0)
    )


def build_semester_report_from_indicators(indicators_df: pd.DataFrame) -> pd.DataFrame:
    if indicators_df.empty:
        return indicators_df.copy()

    period_col = find_column_by_token(indicators_df, "period")
    if period_col is None:
        raise KeyError("Period column not found in indicators sheet.")

    dimension_cols, numeric_cols = detect_dimension_columns(indicators_df)
    group_cols = [col for col in dimension_cols if col != period_col]

    target_col = find_metric_column(indicators_df, ["target"])
    male_col = find_metric_column(indicators_df, ["male"], banned_tokens=["female"])
    female_col = find_metric_column(indicators_df, ["female"])
    total_col = find_metric_column(indicators_df, ["total"], banned_tokens=["subtotal", "grandtotal"])

    metric_map = {
        "Target": target_col,
        "Male": male_col,
        "Female": female_col,
        "Total": total_col,
    }
    value_cols = [col for col in metric_map.values() if col is not None and col in numeric_cols]

    working = indicators_df.copy()
    for col in value_cols:
        working[col] = cleaned_numeric(working[col])

    working["__quarter"] = working[period_col].apply(canonical_quarter_label)
    quarterly = working.loc[working["__quarter"].notna()].copy()

    if quarterly.empty:
        base_cols = [col for col in [period_col, *group_cols] if col in indicators_df.columns]
        output = working.drop(columns=["__quarter"], errors="ignore").reindex(columns=base_cols).drop_duplicates()
        output[period_col] = "Semester"
        for label in [
            "S1 Target", "S1 Male", "S1 Female", "S1 Total",
            "S2 Target", "S2 Male", "S2 Female", "S2 Total",
            "Annual Target", "Annual Male", "Annual Female", "Annual Total",
        ]:
            output[label] = 0
        return output.reset_index(drop=True)

    def period_rollup(periods: set[str], prefix: str) -> pd.DataFrame:
        subset = quarterly.loc[quarterly["__quarter"].isin(periods)]
        aggregated = aggregate_by_keys(subset, group_cols, value_cols)
        rename_map = {
            metric_map["Target"]: f"{prefix} Target",
            metric_map["Male"]: f"{prefix} Male",
            metric_map["Female"]: f"{prefix} Female",
            metric_map["Total"]: f"{prefix} Total",
        }
        rename_map = {k: v for k, v in rename_map.items() if k is not None and k in aggregated.columns}
        aggregated = aggregated.rename(columns=rename_map)
        return aggregated

    s1 = period_rollup({"Q1", "Q2"}, "S1")
    s2 = period_rollup({"Q3", "Q4"}, "S2")
    annual = period_rollup({"Q1", "Q2", "Q3", "Q4"}, "Annual")

    base = quarterly[group_cols].drop_duplicates() if group_cols else pd.DataFrame([{}])
    merged = base.merge(s1, on=group_cols, how="left") if group_cols else s1.copy()
    merged = merged.merge(s2, on=group_cols, how="left") if group_cols else merged.join(s2, rsuffix="_s2")
    merged = merged.merge(annual, on=group_cols, how="left") if group_cols else merged.join(annual, rsuffix="_annual")

    for label in [
        "S1 Target", "S1 Male", "S1 Female", "S1 Total",
        "S2 Target", "S2 Male", "S2 Female", "S2 Total",
        "Annual Target", "Annual Male", "Annual Female", "Annual Total",
    ]:
        if label not in merged.columns:
            merged[label] = 0
        merged[label] = merged[label].fillna(0)

    merged[period_col] = "Semester"

    output_columns = [col for col in indicators_df.columns if col in dimension_cols]
    if period_col not in output_columns:
        output_columns = [period_col, *output_columns]
    output_columns = [col for col in output_columns if col != period_col]
    output_columns = [period_col, *output_columns]

    metric_columns = [
        "S1 Target", "S1 Male", "S1 Female", "S1 Total",
        "S2 Target", "S2 Male", "S2 Female", "S2 Total",
        "Annual Target", "Annual Male", "Annual Female", "Annual Total",
    ]

    return merged.reindex(columns=output_columns + metric_columns).reset_index(drop=True)


def read_target_sheet(path: Path, canonical_sheet: str, aliases: list[str]) -> pd.DataFrame:
    workbook = pd.ExcelFile(path)
    sheet_name = resolve_sheet_name(workbook, aliases)
    df = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")
    df = normalize_columns(df)

    # Use canonical naming for output consistency.
    df.attrs["sheet_name"] = canonical_sheet
    return df


def combine_sheet(chdn_path: Path, kna_path: Path, canonical_sheet: str, aliases: list[str]) -> pd.DataFrame:
    chdn_df = read_target_sheet(chdn_path, canonical_sheet, aliases)
    kna_df = read_target_sheet(kna_path, canonical_sheet, aliases)

    if canonical_sheet == "ALOD_cummu":
        chdn_df = harmonize_alod_cummu_indicator_columns(chdn_df)
        kna_df = harmonize_alod_cummu_indicator_columns(kna_df)

    chdn_df, kna_df = align_columns(chdn_df, kna_df)
    combined = pd.concat([chdn_df, kna_df], ignore_index=True)

    # Summary sheet should include only township-level rows and no appended totals.
    if canonical_sheet == "Summary":
        return combine_summary_clinic_rows_to_township(combined)

    return append_karenni_total_rows(combined)


def write_sheet_with_aliases(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame) -> None:
    df.to_excel(writer, sheet_name=sheet_name, index=False)


def parse_args() -> argparse.Namespace:
    base_dir = Path(__file__).resolve().parent
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    parser = argparse.ArgumentParser(
        description=(
            "Combine CHDN and KNA reports; keep township-level rows only in Summary "
            "and append Karenni Total rows in other target sheets."
        )
    )
    parser.add_argument(
        "--chdn",
        type=Path,
        default=base_dir / "CHDN_report_20260804_034801.xlsx",
        help="Path to CHDN report workbook.",
    )
    parser.add_argument(
        "--kna",
        type=Path,
        default=base_dir / "KNA_EPI_long_20260804_040627.xlsx",
        help="Path to KNA long workbook.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=base_dir / f"Karenni_combined_report_{timestamp}.xlsx",
        help="Output workbook path.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    chdn_path: Path = args.chdn
    kna_path: Path = args.kna
    output_path: Path = args.output

    missing = [p for p in [chdn_path, kna_path] if not p.exists()]
    if missing:
        lines = "\n".join(f"- {item}" for item in missing)
        raise FileNotFoundError(f"Missing source workbook(s):\n{lines}")

    sheet_map: dict[str, pd.DataFrame] = {}
    try:
        for canonical, aliases in TARGET_SHEETS.items():
            sheet_map[canonical] = combine_sheet(chdn_path, kna_path, canonical, aliases)
        if "indicators" in sheet_map:
            sheet_map[SEMESTER_REPORT_SHEET_NAME] = build_semester_report_from_indicators(sheet_map["indicators"])
    except PermissionError as exc:
        raise PermissionError(
            "Cannot read one or more source workbooks. Close CHDN/KNA files in Excel and run again."
        ) from exc

    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for sheet_name, df in sheet_map.items():
                write_sheet_with_aliases(writer, sheet_name, df)
    except PermissionError as exc:
        raise PermissionError(
            "Cannot write output file. Close the workbook in Excel and run again."
        ) from exc

    print(f"Wrote combined report: {output_path}")


if __name__ == "__main__":
    main()
