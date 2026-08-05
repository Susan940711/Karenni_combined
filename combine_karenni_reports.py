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
SEMESTER_ALOD_OVERRIDE_LABEL = "At least one dose under 5-yr-old"


def normalize_name(value: str) -> str:
    return "".join(ch for ch in str(value).lower() if ch.isalnum())


def normalize_key_value(value: object) -> str:
    if is_empty_value(value):
        return ""
    text = str(value).strip()
    # Align keys like 2026 and 2026.0 across sheets.
    try:
        num = float(text.replace(",", ""))
        if num.is_integer():
            return str(int(num))
    except Exception:
        pass
    return text


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


def build_semester_metric_frame(source_df: pd.DataFrame) -> pd.DataFrame:
    if source_df.empty:
        return source_df.copy()

    def quarter_metric_columns(
        quarter: int,
        required_tokens: list[str],
        banned_tokens: list[str] | None = None,
    ) -> list[str]:
        banned_tokens = banned_tokens or []
        quarter_tokens = [f"q{quarter}", f"quarter{quarter}", f"qtr{quarter}"]

        matches: list[str] = []
        for col in source_df.columns:
            norm = normalize_name(col)
            if not any(token in norm for token in quarter_tokens):
                continue
            if any(token not in norm for token in required_tokens):
                continue
            if any(bad in norm for bad in banned_tokens):
                continue
            matches.append(col)
        return matches

    q_cols: dict[int, dict[str, list[str]]] = {}
    for q in [1, 2, 3, 4]:
        q_cols[q] = {
            "target": quarter_metric_columns(q, ["target"]),
            "u1_male": quarter_metric_columns(q, ["u1", "male"], banned_tokens=["female"]),
            "u1_female": quarter_metric_columns(q, ["u1", "female"]),
            "one5_male": quarter_metric_columns(q, ["15", "male"], banned_tokens=["female"]),
            "one5_female": quarter_metric_columns(q, ["15", "female"]),
            "total": quarter_metric_columns(q, ["total"], banned_tokens=["subtotal", "grandtotal"]),
        }

    quarter_value_cols = list(
        dict.fromkeys(
            col
            for q in [1, 2, 3, 4]
            for cols in q_cols[q].values()
            for col in cols
        )
    )

    working = source_df.copy()
    for col in quarter_value_cols:
        working[col] = cleaned_numeric(working[col])

    def row_sum(frame: pd.DataFrame, names: list[str]) -> pd.Series:
        present = [name for name in names if name in frame.columns]
        if not present:
            return pd.Series(0, index=frame.index, dtype="float64")
        return frame[present].sum(axis=1, min_count=1).fillna(0)

    working["S1 Target"] = row_sum(working, q_cols[1]["target"] + q_cols[2]["target"])
    working["S1 Male"] = row_sum(working, q_cols[1]["u1_male"] + q_cols[1]["one5_male"] + q_cols[2]["u1_male"] + q_cols[2]["one5_male"])
    working["S1 Female"] = row_sum(working, q_cols[1]["u1_female"] + q_cols[1]["one5_female"] + q_cols[2]["u1_female"] + q_cols[2]["one5_female"])
    s1_total_cols = q_cols[1]["total"] + q_cols[2]["total"]
    working["S1 Total"] = row_sum(working, s1_total_cols) if s1_total_cols else working["S1 Male"] + working["S1 Female"]

    working["S2 Target"] = row_sum(working, q_cols[3]["target"] + q_cols[4]["target"])
    working["S2 Male"] = row_sum(working, q_cols[3]["u1_male"] + q_cols[3]["one5_male"] + q_cols[4]["u1_male"] + q_cols[4]["one5_male"])
    working["S2 Female"] = row_sum(working, q_cols[3]["u1_female"] + q_cols[3]["one5_female"] + q_cols[4]["u1_female"] + q_cols[4]["one5_female"])
    s2_total_cols = q_cols[3]["total"] + q_cols[4]["total"]
    working["S2 Total"] = row_sum(working, s2_total_cols) if s2_total_cols else working["S2 Male"] + working["S2 Female"]

    working["Annual Target"] = row_sum(working, q_cols[1]["target"] + q_cols[2]["target"] + q_cols[3]["target"] + q_cols[4]["target"])
    working["Annual Male"] = row_sum(working, q_cols[1]["u1_male"] + q_cols[1]["one5_male"] + q_cols[2]["u1_male"] + q_cols[2]["one5_male"] + q_cols[3]["u1_male"] + q_cols[3]["one5_male"] + q_cols[4]["u1_male"] + q_cols[4]["one5_male"])
    working["Annual Female"] = row_sum(working, q_cols[1]["u1_female"] + q_cols[1]["one5_female"] + q_cols[2]["u1_female"] + q_cols[2]["one5_female"] + q_cols[3]["u1_female"] + q_cols[3]["one5_female"] + q_cols[4]["u1_female"] + q_cols[4]["one5_female"])
    annual_total_cols = q_cols[1]["total"] + q_cols[2]["total"] + q_cols[3]["total"] + q_cols[4]["total"]
    working["Annual Total"] = row_sum(working, annual_total_cols) if annual_total_cols else working["Annual Male"] + working["Annual Female"]

    metric_cols = [
        "S1 Target", "S1 Male", "S1 Female", "S1 Total",
        "S2 Target", "S2 Male", "S2 Female", "S2 Total",
        "Annual Target", "Annual Male", "Annual Female", "Annual Total",
    ]
    for label in metric_cols:
        working[label] = working[label].fillna(0)

    return working


def build_alod_override_metric_frame(source_df: pd.DataFrame) -> pd.DataFrame:
    if source_df.empty:
        return source_df.copy()

    def period_metric_columns(
        period_prefix: str,
        required_tokens: list[str],
        banned_tokens: list[str] | None = None,
    ) -> list[str]:
        banned_tokens = banned_tokens or []
        period_tokens = [period_prefix]

        matches: list[str] = []
        for col in source_df.columns:
            norm = normalize_name(col)
            if not any(token in norm for token in period_tokens):
                continue
            if any(token not in norm for token in required_tokens):
                continue
            if any(bad in norm for bad in banned_tokens):
                continue
            matches.append(col)
        return matches

    p_cols: dict[str, dict[str, list[str]]] = {}
    for prefix in ["S1", "S2", "Annual"]:
        p_cols[prefix] = {
            "target": period_metric_columns(prefix, ["target"]),
            "u1_male": period_metric_columns(prefix, ["u1", "male"], banned_tokens=["female"]),
            "u1_female": period_metric_columns(prefix, ["u1", "female"]),
            "one5_male": period_metric_columns(prefix, ["15", "male"], banned_tokens=["female"]),
            "one5_female": period_metric_columns(prefix, ["15", "female"]),
            "total": period_metric_columns(prefix, ["total"], banned_tokens=["subtotal", "grandtotal"]),
        }

    period_value_cols = list(
        dict.fromkeys(
            col
            for prefix in ["S1", "S2", "Annual"]
            for cols in p_cols[prefix].values()
            for col in cols
        )
    )

    working = source_df.copy()
    for col in period_value_cols:
        working[col] = cleaned_numeric(working[col])

    def row_sum(frame: pd.DataFrame, names: list[str]) -> pd.Series:
        present = [name for name in names if name in frame.columns]
        if not present:
            return pd.Series(0, index=frame.index, dtype="float64")
        return frame[present].sum(axis=1, min_count=1).fillna(0)

    working["S1 Target"] = row_sum(working, p_cols["S1"]["target"])
    working["S1 Male"] = row_sum(working, p_cols["S1"]["u1_male"] + p_cols["S1"]["one5_male"])
    working["S1 Female"] = row_sum(working, p_cols["S1"]["u1_female"] + p_cols["S1"]["one5_female"])
    s1_total_cols = p_cols["S1"]["total"]
    working["S1 Total"] = row_sum(working, s1_total_cols) if s1_total_cols else working["S1 Male"] + working["S1 Female"]

    working["S2 Target"] = row_sum(working, p_cols["S2"]["target"])
    working["S2 Male"] = row_sum(working, p_cols["S2"]["u1_male"] + p_cols["S2"]["one5_male"])
    working["S2 Female"] = row_sum(working, p_cols["S2"]["u1_female"] + p_cols["S2"]["one5_female"])
    s2_total_cols = p_cols["S2"]["total"]
    working["S2 Total"] = row_sum(working, s2_total_cols) if s2_total_cols else working["S2 Male"] + working["S2 Female"]

    working["Annual Target"] = row_sum(working, p_cols["Annual"]["target"])
    working["Annual Male"] = row_sum(working, p_cols["Annual"]["u1_male"] + p_cols["Annual"]["one5_male"])
    working["Annual Female"] = row_sum(working, p_cols["Annual"]["u1_female"] + p_cols["Annual"]["one5_female"])
    annual_total_cols = p_cols["Annual"]["total"]
    working["Annual Total"] = row_sum(working, annual_total_cols) if annual_total_cols else working["Annual Male"] + working["Annual Female"]

    metric_cols = [
        "S1 Target", "S1 Male", "S1 Female", "S1 Total",
        "S2 Target", "S2 Male", "S2 Female", "S2 Total",
        "Annual Target", "Annual Male", "Annual Female", "Annual Total",
    ]
    for label in metric_cols:
        working[label] = working[label].fillna(0)

    return working


def build_alod_semester_lookup(source_df: pd.DataFrame) -> dict[tuple[str, str, str, str], pd.Series]:
    if source_df.empty:
        return {}

    frame = build_alod_override_metric_frame(source_df)
    indicator_col = next((c for c in frame.columns if normalize_name(c) == "indicator"), None)
    period_col = next((c for c in frame.columns if normalize_name(c) == "period"), None)
    organization_col = next((c for c in frame.columns if normalize_name(c) == "organization"), None)
    project_col = next((c for c in frame.columns if "project" in normalize_name(c)), None)

    if indicator_col is None:
        return {}

    lookup: dict[tuple[str, str, str, str], pd.Series] = {}
    metric_columns = [
        "S1 Target", "S1 Male", "S1 Female", "S1 Total",
        "S2 Target", "S2 Male", "S2 Female", "S2 Total",
        "Annual Target", "Annual Male", "Annual Female", "Annual Total",
    ]

    for _, row in frame.iterrows():
        label = canonicalize_alod_label(row.get(indicator_col, ""))
        if label != SEMESTER_ALOD_OVERRIDE_LABEL:
            continue

        period_value = normalize_key_value(row.get(period_col, "")) if period_col is not None else ""
        organization_value = normalize_key_value(row.get(organization_col, "")) if organization_col is not None else ""
        project_value = normalize_key_value(row.get(project_col, "")) if project_col is not None else ""

        key = (period_value, organization_value, project_value, label)
        score = float(pd.Series([row.get(metric, 0) for metric in metric_columns]).sum())

        existing = lookup.get(key)
        existing_score = -1.0 if existing is None else float(pd.Series([existing.get(metric, 0) for metric in metric_columns]).sum())

        if existing is None or score > existing_score:
            lookup[key] = row

    return lookup


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


def build_semester_report_from_indicators(
    indicators_df: pd.DataFrame,
    alod_lookup: dict[tuple[str, str, str, str], pd.Series] | None = None,
) -> pd.DataFrame:
    if indicators_df.empty:
        return indicators_df.copy()

    period_col = find_column_by_token(indicators_df, "period")
    if period_col is None:
        raise KeyError("Period column not found in indicators sheet.")
    metric_frame = build_semester_metric_frame(indicators_df)
    dimension_cols, _ = detect_dimension_columns(metric_frame)
    organization_col = next((c for c in metric_frame.columns if normalize_name(c) == "organization"), None)

    metric_columns = [
        "S1 Target", "S1 Male", "S1 Female", "S1 Total",
        "S2 Target", "S2 Male", "S2 Female", "S2 Total",
        "Annual Target", "Annual Male", "Annual Female", "Annual Total",
    ]

    output_columns = [col for col in metric_frame.columns if col in dimension_cols or col == organization_col]
    if period_col not in output_columns:
        output_columns = [period_col, *output_columns]
    output_columns = [col for col in output_columns if col != period_col]
    output_columns = [period_col, *output_columns]

    semester_df = metric_frame.reindex(columns=output_columns + metric_columns).copy()

    if alod_lookup:
        semester_indicator_col = next((c for c in semester_df.columns if normalize_name(c) == "indicator"), None)
        semester_organization_col = next((c for c in semester_df.columns if normalize_name(c) == "organization"), None)
        project_col = next((c for c in semester_df.columns if "project" in normalize_name(c)), None)

        if semester_indicator_col is not None:
            target_rows = semester_df.index[
                semester_df[semester_indicator_col].astype("string").fillna("").str.strip().map(canonicalize_alod_label)
                == SEMESTER_ALOD_OVERRIDE_LABEL
            ].tolist()

            for row_index in target_rows:
                period_value = normalize_key_value(semester_df.at[row_index, period_col])
                org_value = normalize_key_value(semester_df.at[row_index, semester_organization_col]) if semester_organization_col in semester_df.columns else ""
                project_value = normalize_key_value(semester_df.at[row_index, project_col]) if project_col in semester_df.columns else ""
                key_candidates = [
                    (period_value, org_value, project_value, SEMESTER_ALOD_OVERRIDE_LABEL),
                    (period_value, "", project_value, SEMESTER_ALOD_OVERRIDE_LABEL),
                    (period_value, org_value, "", SEMESTER_ALOD_OVERRIDE_LABEL),
                    (period_value, "", "", SEMESTER_ALOD_OVERRIDE_LABEL),
                    ("", "", "", SEMESTER_ALOD_OVERRIDE_LABEL),
                ]

                source_row = None
                for key in key_candidates:
                    if key in alod_lookup:
                        source_row = alod_lookup[key]
                        break

                if source_row is None:
                    continue

                for metric_col in metric_columns:
                    if metric_col in source_row:
                        semester_df.at[row_index, metric_col] = source_row[metric_col]

    return semester_df.reset_index(drop=True)


def build_semester_report_from_sheet_map(sheet_map: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Build the semester sheet from the already combined output sheet data."""
    alod_lookup = build_alod_semester_lookup(sheet_map.get("ALOD_cummu", pd.DataFrame()))
    return build_semester_report_from_indicators(
        sheet_map["indicators"],
        alod_lookup,
    )


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
            sheet_map[SEMESTER_REPORT_SHEET_NAME] = build_semester_report_from_sheet_map(sheet_map)
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
