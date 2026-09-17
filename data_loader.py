"""
Loads the four CLMRS datasets (Farmer List, Household, Inspection, Follow Up)
from user-selected Excel/CSV files, applies the programme's column mapping,
and normalises IDs/dates so downstream analysis never has to guess a dtype.

RAW DATA principle: the original uploaded file is never modified. Loading
produces a new in-memory (and optionally cached) dataframe; the source path
is remembered for traceability but nothing is written back to it.
"""
from __future__ import annotations
import os
import pandas as pd
from .column_mapping import map_headers, apply_mapping


class LoadResult:
    def __init__(self, df: pd.DataFrame, source_path: str, dataset_type: str,
                 mapping: dict, unmapped: list, warnings: list):
        self.df = df
        self.source_path = source_path
        self.dataset_type = dataset_type  # "farmer_list" | "household" | "inspection" | "follow_up"
        self.mapping = mapping
        self.unmapped = unmapped
        self.warnings = warnings


def _read_any(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        return pd.read_csv(path, dtype=str, keep_default_na=True)
    # Read as-is first (mixed types); IDs get stringified explicitly after.
    return pd.read_excel(path, dtype=str)


def _clean_id_series(s: pd.Series) -> pd.Series:
    """IDs must compare reliably: strip whitespace, drop a trailing '.0' that
    creeps in when Excel/pandas briefly treats a numeric-looking ID as a float."""
    s = s.astype(str).str.strip()
    s = s.str.replace(r"\.0$", "", regex=True)
    s = s.replace({"nan": None, "None": None, "": None})
    return s


def _clean_dates(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce")


DATE_FIELDS = {"DATE_REGISTERED", "NO_CHILDREN_DATE", "SURVEY_DATE"}
ID_FIELDS = {"FARMER_ID", "CHILD_ID"}


def load_survey_file(path: str, dataset_type: str, header_map: dict) -> LoadResult:
    """dataset_type is one of 'household', 'inspection', 'follow_up' — used only
    for labelling; all three share the same header vocabulary."""
    raw = _read_any(path)
    raw.columns = [str(c).strip() for c in raw.columns]
    mapping, unmapped = map_headers(list(raw.columns), header_map)
    df = apply_mapping(raw, mapping)

    warnings = []
    for f in ID_FIELDS:
        if f in df.columns:
            df[f] = _clean_id_series(df[f])
        else:
            warnings.append(f"Missing expected field: {f}")
    for f in DATE_FIELDS:
        if f in df.columns:
            n_before = df[f].notna().sum()
            df[f] = _clean_dates(df[f])
            n_after = df[f].notna().sum()
            if n_before and n_after < n_before:
                warnings.append(
                    f"{n_before - n_after} value(s) in {f} could not be parsed as dates"
                )

    df["SOURCE_FILE"] = os.path.basename(path)
    df["SOURCE_ROW"] = range(2, len(df) + 2)  # +2: header row + 1-based Excel rows

    return LoadResult(df, path, dataset_type, mapping, unmapped, warnings)


def load_farmer_list(path: str, header_map: dict) -> LoadResult:
    raw = _read_any(path)
    raw.columns = [str(c).strip() for c in raw.columns]
    mapping, unmapped = map_headers(list(raw.columns), header_map)
    df = apply_mapping(raw, mapping)

    warnings = []
    if "FARMER_ID" in df.columns:
        df["FARMER_ID"] = _clean_id_series(df["FARMER_ID"])
    else:
        warnings.append("Farmer List is missing a FARMER_ID column")

    df["SOURCE_FILE"] = os.path.basename(path)
    return LoadResult(df, path, "farmer_list", mapping, unmapped, warnings)
