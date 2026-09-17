"""
Quarter classification and Case Identified Date calculation.

Case Identified Date uses TRUE calendar-month subtraction (not a fixed
90/180-day offset), matching month-end correctly: 30 Sep minus 6 months =
30 Mar; 31 Jan minus 1 month = 31 Dec (clamped to the shorter month where
the day doesn't exist, e.g. 31 Mar minus 1 month = 28/29 Feb).
"""
from __future__ import annotations
import pandas as pd
from dateutil.relativedelta import relativedelta


def quarter_for_date(dt, quarter_map: dict):
    if pd.isna(dt):
        return None
    return quarter_map.get(dt.month)


def add_quarter_column(df: pd.DataFrame, date_field: str, quarter_map: dict,
                        out_field: str = "QUARTER"):
    df[out_field] = df[date_field].apply(lambda d: quarter_for_date(d, quarter_map))
    return df


def case_identified_date(follow_up_date, status: str, lag_months_map: dict):
    """status is matched case-insensitively against the programme's
    case_id_lag_months config; unrecognised statuses return None rather than
    guessing a lag."""
    if pd.isna(follow_up_date) or not status:
        return None
    key = str(status).strip().lower()
    months = lag_months_map.get(key)
    if months is None:
        return None
    return follow_up_date - relativedelta(months=months)


def add_case_identified_date(df: pd.DataFrame, date_field: str, status_field: str,
                              lag_months_map: dict, out_field: str = "CASE_IDENTIFIED_DATE"):
    df[out_field] = df.apply(
        lambda r: case_identified_date(r.get(date_field), r.get(status_field), lag_months_map),
        axis=1,
    )
    return df
