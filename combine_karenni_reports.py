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


def find_metric_columns(
    df: pd.DataFrame,
    required_tokens: list[str],
    any_tokens: list[str] | None = None,
    banned_tokens: list[str] | None = None,
) -> list[str]:
    any_tokens = any_tokens or []
    banned_tokens = banned_tokens or []

    matched: list[str] = []
    for col in df.columns:
        norm = normalize_name(col)
        if any(bad in norm for bad in banned_tokens):
            continue
        if not all(token in norm for token in required_tokens):
            continue
        if any_tokens and not any(token in norm for token in any_tokens):
            continue
        matched.append(col)
    return matched


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

    def quarter_metric_column(quarter: int, required_tokens: list[str], banned_tokens: list[str] | None = None) -> str | None:
        banned_tokens = banned_tokens or []
        quarter_tokens = [f"q{quarter}", f"quarter{quarter}", f"qtr{quarter}"]

        for col in indicators_df.columns:
            norm = normalize_name(col)
            if not any(token in norm for token in quarter_tokens):
                continue
            if any(token not in norm for token in required_tokens):
                continue
            if any(bad in norm for bad in banned_tokens):
                continue
            return col
        return None

    q_cols: dict[int, dict[str, str | None]] = {}
    for q in [1, 2, 3, 4]:
        q_cols[q] = {
            "target": quarter_metric_column(q, ["target"]),
            "u1_male": quarter_metric_column(q, ["u1", "male"], banned_tokens=["female"]),
            "u1_female": quarter_metric_column(q, ["u1", "female"]),
            "one5_male": quarter_metric_column(q, ["15", "male"], banned_tokens=["female"]),
            "one5_female": quarter_metric_column(q, ["15", "female"]),
            "total": quarter_metric_column(q, ["total"], banned_tokens=["subtotal", "grandtotal"]),
        }

    quarter_value_cols = [
        col
        for q in [1, 2, 3, 4]
        for col in q_cols[q].values()
        if col is not None
    ]
    quarter_value_cols = list(dict.fromkeys(quarter_value_cols))

    dimension_cols, _ = detect_dimension_columns(indicators_df)
    organization_col = next((c for c in indicators_df.columns if normalize_name(c) == "organization"), None)
    group_cols = [col for col in dimension_cols if col != period_col]
    if organization_col is not None and organization_col in indicators_df.columns and organization_col not in group_cols:
        group_cols.append(organization_col)

    if not quarter_value_cols:
        base = indicators_df[group_cols].drop_duplicates() if group_cols else pd.DataFrame([{}])
        base[period_col] = "Semester"
        for label in [
            "S1 Target", "S1 Male", "S1 Female", "S1 Total",
            "S2 Target", "S2 Male", "S2 Female", "S2 Total",
            "Annual Target", "Annual Male", "Annual Female", "Annual Total",
        ]:
            base[label] = 0
        output_columns = [period_col, *[c for c in dimension_cols if c != period_col]]
        return base.reindex(columns=output_columns + [
            "S1 Target", "S1 Male", "S1 Female", "S1 Total",
            "S2 Target", "S2 Male", "S2 Female", "S2 Total",
            "Annual Target", "Annual Male", "Annual Female", "Annual Total",
        ]).reset_index(drop=True)

    working = indicators_df.copy()
    for col in quarter_value_cols:
        working[col] = cleaned_numeric(working[col])

    aggregated = aggregate_by_keys(working, group_cols, quarter_value_cols)

    def colsum(frame: pd.DataFrame, names: list[str | None]) -> pd.Series:
        present = [name for name in names if name is not None and name in frame.columns]
        if not present:
            return pd.Series(0, index=frame.index, dtype="float64")
        return frame[present].sum(axis=1, min_count=1).fillna(0)

    merged = aggregated.copy()

    merged["S1 Target"] = colsum(merged, [q_cols[1]["target"], q_cols[2]["target"]])
    merged["S1 Male"] = colsum(merged, [q_cols[1]["u1_male"], q_cols[1]["one5_male"], q_cols[2]["u1_male"], q_cols[2]["one5_male"]])
    merged["S1 Female"] = colsum(merged, [q_cols[1]["u1_female"], q_cols[1]["one5_female"], q_cols[2]["u1_female"], q_cols[2]["one5_female"]])
    s1_total_cols = [q_cols[1]["total"], q_cols[2]["total"]]
    s1_total_from_quarter_total = colsum(merged, s1_total_cols)
    if any(col is not None for col in s1_total_cols):
        merged["S1 Total"] = s1_total_from_quarter_total
    else:
        merged["S1 Total"] = merged["S1 Male"] + merged["S1 Female"]

    merged["S2 Target"] = colsum(merged, [q_cols[3]["target"], q_cols[4]["target"]])
    merged["S2 Male"] = colsum(merged, [q_cols[3]["u1_male"], q_cols[3]["one5_male"], q_cols[4]["u1_male"], q_cols[4]["one5_male"]])
    merged["S2 Female"] = colsum(merged, [q_cols[3]["u1_female"], q_cols[3]["one5_female"], q_cols[4]["u1_female"], q_cols[4]["one5_female"]])
    s2_total_cols = [q_cols[3]["total"], q_cols[4]["total"]]
    s2_total_from_quarter_total = colsum(merged, s2_total_cols)
    if any(col is not None for col in s2_total_cols):
        merged["S2 Total"] = s2_total_from_quarter_total
    else:
        merged["S2 Total"] = merged["S2 Male"] + merged["S2 Female"]

    merged["Annual Target"] = colsum(merged, [q_cols[1]["target"], q_cols[2]["target"], q_cols[3]["target"], q_cols[4]["target"]])
    merged["Annual Male"] = colsum(merged, [q_cols[1]["u1_male"], q_cols[1]["one5_male"], q_cols[2]["u1_male"], q_cols[2]["one5_male"], q_cols[3]["u1_male"], q_cols[3]["one5_male"], q_cols[4]["u1_male"], q_cols[4]["one5_male"]])
    merged["Annual Female"] = colsum(merged, [q_cols[1]["u1_female"], q_cols[1]["one5_female"], q_cols[2]["u1_female"], q_cols[2]["one5_female"], q_cols[3]["u1_female"], q_cols[3]["one5_female"], q_cols[4]["u1_female"], q_cols[4]["one5_female"]])
    annual_total_cols = [q_cols[1]["total"], q_cols[2]["total"], q_cols[3]["total"], q_cols[4]["total"]]
    annual_total_from_quarter_total = colsum(merged, annual_total_cols)
    if any(col is not None for col in annual_total_cols):
        merged["Annual Total"] = annual_total_from_quarter_total
    else:
        merged["Annual Total"] = merged["Annual Male"] + merged["Annual Female"]

    for label in [
        "S1 Target", "S1 Male", "S1 Female", "S1 Total",
        "S2 Target", "S2 Male", "S2 Female", "S2 Total",
        "Annual Target", "Annual Male", "Annual Female", "Annual Total",
    ]:
        if label not in merged.columns:
            merged[label] = 0
        merged[label] = merged[label].fillna(0)

    # Keep period value from indicators source (first non-empty value per group).
    period_source = indicators_df.copy()
    period_values = period_source[period_col].astype("string").fillna("").str.strip()
    period_source = period_source.loc[period_values != ""]

    if group_cols and not period_source.empty:
        period_lookup = period_source[group_cols + [period_col]].drop_duplicates(subset=group_cols, keep="first")
        merged = merged.merge(period_lookup, on=group_cols, how="left")
    elif not period_source.empty:
        merged[period_col] = period_source[period_col].iloc[0]
    else:
        merged[period_col] = ""

    output_columns = [col for col in indicators_df.columns if col in dimension_cols or col == organization_col]
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
