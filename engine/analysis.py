"""
Core CLMRS analysis engine. Every function takes already-loaded, mapped
dataframes (see data_loader.py) and a ProgrammeConfig, and returns a
dataframe or dict of dataframes ready for Excel export. Nothing here writes
files; excel_export.py owns that.

Design principle (traceability): every output row carries FARMER_ID / CHILD_ID
/ SOURCE_FILE / SOURCE_ROW where applicable, so every number in the client
report can be traced back to a specific source record.
"""
from __future__ import annotations
import pandas as pd
from .quarters import add_quarter_column, add_case_identified_date

QUARTER_ORDER = ["Q1", "Q2", "Q3", "Q4"]


def _quarter_sort_key(q):
    try:
        return QUARTER_ORDER.index(q)
    except ValueError:
        return 99


def _join_unique(series, sort=True):
    vals = [str(v) for v in series.dropna().unique()]
    if sort:
        vals = sorted(vals)
    return ", ".join(vals)


# ---------------------------------------------------------- Farmer/Child detail
# One row per unique entity, raw fields + analytical fields combined, split
# into on-Annex-list / not-on-list — mirrors Philemon's standing reporting
# style for these recurring analyses.

def farmer_detail_table(df: pd.DataFrame, farmer_list_ids: set, farmer_gender_map: dict = None) -> dict:
    """Returns {'on_list': df, 'not_on_list': df}. farmer_gender_map (optional):
    {FARMER_ID: gender} sourced from the Farmer List/Annex, the only dataset
    that carries the farmer's own gender."""
    farmer_gender_map = farmer_gender_map or {}
    farmers = df.dropna(subset=["FARMER_ID"]).copy()
    if farmers.empty:
        empty = pd.DataFrame()
        return {"on_list": empty, "not_on_list": empty}

    if "SURVEY_DATE" in farmers.columns:
        farmers = farmers.sort_values("SURVEY_DATE")
        # QUARTER is already computed upstream (pipeline.py) from SURVEY_DATE

    enum_id_col = "SURVEYED_BY_ID" if "SURVEYED_BY_ID" in farmers.columns else None
    enum_name_col = "SURVEYED_BY_NAME" if "SURVEYED_BY_NAME" in farmers.columns else None
    visit_id_col = "VISITED_BY_ID" if "VISITED_BY_ID" in farmers.columns else None
    visit_name_col = "VISITED_BY_NAME" if "VISITED_BY_NAME" in farmers.columns else None
    enum_gender_col = "ENUM_GENDER" if "ENUM_GENDER" in farmers.columns else None
    enum_gender_code_col = "ENUM_GENDER_CODE" if "ENUM_GENDER_CODE" in farmers.columns else None

    rows = []
    for farmer_id, sub in farmers.groupby("FARMER_ID", sort=False):
        sub_sorted = sub.sort_values("SURVEY_DATE") if "SURVEY_DATE" in sub.columns else sub
        first_row = sub_sorted.iloc[0]
        children_profiled = sub["CHILD_ID"].nunique() if "CHILD_ID" in sub.columns else 0

        # Enumerator: prefer "surveyed" id/name; fall back to "visited" (used
        # when the farmer has no child, per Butter's standing convention)
        e_ids = sub[enum_id_col].dropna() if enum_id_col else pd.Series(dtype=str)
        e_names = sub[enum_name_col].dropna() if enum_name_col else pd.Series(dtype=str)
        if e_ids.empty and visit_id_col:
            e_ids = sub[visit_id_col].dropna()
        if e_names.empty and visit_name_col:
            e_names = sub[visit_name_col].dropna()
        e_genders = sub[enum_gender_col].dropna() if enum_gender_col else pd.Series(dtype=str)
        e_gender_codes = sub[enum_gender_code_col].dropna() if enum_gender_code_col else pd.Series(dtype=str)

        quarters_visited = sorted(sub["QUARTER"].dropna().unique(), key=_quarter_sort_key) if "QUARTER" in sub.columns else []
        survey_dates = sorted(d.strftime("%Y-%m-%d") for d in sub["SURVEY_DATE"].dropna().unique()) if "SURVEY_DATE" in sub.columns else []

        rows.append({
            "Quarter": quarters_visited[0] if quarters_visited else None,
            "Farmer ID": farmer_id,
            "Farmer Name": first_row.get("FARMER_NAME"),
            "Farmer Gender": farmer_gender_map.get(farmer_id),
            "Farmer Group": first_row.get("FARMER_GROUP"),
            "Children Profiled": int(children_profiled),
            "Survey Date(s)": ", ".join(survey_dates),
            "Quarters Visited": ", ".join(quarters_visited),
            "Enumerator ID": _join_unique(e_ids, sort=False),
            "Enumerator Name": _join_unique(e_names, sort=False),
            "Enumerator Gender": _join_unique(e_genders, sort=False),
            "Enumerator Gender Code": _join_unique(e_gender_codes, sort=False),
            "Farmer List Match": "Present" if farmer_id in farmer_list_ids else "Absent",
        })

    detail = pd.DataFrame(rows)
    on_list = detail[detail["Farmer List Match"] == "Present"].reset_index(drop=True)
    not_on_list = detail[detail["Farmer List Match"] == "Absent"].reset_index(drop=True)
    return {"on_list": on_list, "not_on_list": not_on_list}


def child_detail_table(df: pd.DataFrame, farmer_list_ids: set,
                        gender_codes: dict, school_codes: dict) -> dict:
    """Returns {'on_list': df, 'not_on_list': df}. One row per unique Child ID
    (latest record if a child has more than one)."""
    children = df.dropna(subset=["CHILD_ID"]).copy()
    if children.empty:
        empty = pd.DataFrame()
        return {"on_list": empty, "not_on_list": empty}

    if "SURVEY_DATE" in children.columns:
        children = children.sort_values("SURVEY_DATE")
    unique_children = children.drop_duplicates(subset=["CHILD_ID"], keep="last").copy()

    def _code_lookup(code, code_map):
        if pd.isna(code):
            return None
        key = str(code).strip()
        if key.endswith(".0"):
            key = key[:-2]
        return code_map.get(key)

    if "SCHOOL_ATTENDANCE" in unique_children.columns:
        school_col = unique_children["SCHOOL_ATTENDANCE"]
    elif "SCHOOL_ATTENDANCE_CODE" in unique_children.columns:
        school_col = unique_children["SCHOOL_ATTENDANCE_CODE"].apply(lambda c: _code_lookup(c, school_codes))
    else:
        school_col = None

    if "GENDER" in unique_children.columns:
        gender_col = unique_children["GENDER"]
    elif "GENDER_CODE" in unique_children.columns:
        gender_col = unique_children["GENDER_CODE"].apply(lambda c: _code_lookup(c, gender_codes))
    else:
        gender_col = None

    rows = pd.DataFrame({
        "Quarter": unique_children.get("QUARTER"),
        "Child ID": unique_children["CHILD_ID"],
        "Child Name": unique_children.get("CHILD_NAME"),
        "Farmer ID": unique_children.get("FARMER_ID"),
        "Farmer Name": unique_children.get("FARMER_NAME"),
        "Farmer Group": unique_children.get("FARMER_GROUP"),
        "Survey Date": unique_children.get("SURVEY_DATE"),
        "Status": unique_children.get("STATUS"),
        "Gender": gender_col,
        "Gender Code": unique_children.get("GENDER_CODE"),
        "School Attendance": school_col,
        "School Attendance Code": unique_children.get("SCHOOL_ATTENDANCE_CODE"),
        "Education Level": unique_children.get("EDUCATION_LEVEL"),
        "Farmer List Match": unique_children.get("FARMER_ID").apply(
            lambda fid: "Present" if fid in farmer_list_ids else "Absent"
        ),
    })

    on_list = rows[rows["Farmer List Match"] == "Present"].reset_index(drop=True)
    not_on_list = rows[rows["Farmer List Match"] == "Absent"].reset_index(drop=True)
    return {"on_list": on_list, "not_on_list": not_on_list}


def child_gender_summary(child_on: pd.DataFrame) -> pd.DataFrame:
    """Pivot of child Gender by quarter, for Annex-matched children — same
    shape as the School Attendance pivot."""
    if child_on.empty or "Gender" not in child_on.columns:
        return pd.DataFrame()
    working = child_on.copy()
    working["Gender"] = working["Gender"].fillna("Missing")
    pivot = pd.pivot_table(
        working, index="Gender", columns="Quarter", values="Child ID",
        aggfunc="count", fill_value=0, observed=True,
    )
    for q in QUARTER_ORDER:
        if q not in pivot.columns:
            pivot[q] = 0
    pivot = pivot[[q for q in QUARTER_ORDER if q in pivot.columns]]
    pivot["Total"] = pivot.sum(axis=1)
    pivot = pivot.reset_index()
    total_row = {"Gender": "Total"}
    for col in pivot.columns:
        if col != "Gender":
            total_row[col] = pivot[col].sum()
    return pd.concat([pivot, pd.DataFrame([total_row])], ignore_index=True)


def farmer_gender_summary(farmer_on: pd.DataFrame) -> pd.DataFrame:
    """Simple count/percentage of Annex-matched farmers by their own gender
    (sourced from the Farmer List, since Household/Inspection records don't
    carry the farmer's own gender)."""
    if farmer_on.empty or "Farmer Gender" not in farmer_on.columns:
        return pd.DataFrame()
    working = farmer_on.copy()
    working["Farmer Gender"] = working["Farmer Gender"].fillna("Not on Farmer List / Missing")
    counts = working["Farmer Gender"].value_counts().reset_index()
    counts.columns = ["Farmer Gender", "Count"]
    total = counts["Count"].sum()
    counts["% of Farmers"] = (counts["Count"] / total * 100).round(1) if total else 0
    total_row = {"Farmer Gender": "Total", "Count": total, "% of Farmers": 100.0 if total else 0}
    return pd.concat([counts, pd.DataFrame([total_row])], ignore_index=True)


def enumerator_gender_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Count of survey records by enumerator gender (falls back to the
    visiting enumerator's gender where no child was surveyed)."""
    working = df.copy()
    gender_col = None
    if "ENUM_GENDER" in working.columns:
        gender_col = working["ENUM_GENDER"]
    if gender_col is None or gender_col.isna().all():
        return pd.DataFrame()
    working["_gender"] = gender_col.fillna("Missing")
    counts = working["_gender"].value_counts().reset_index()
    counts.columns = ["Enumerator Gender", "Records"]
    total = counts["Records"].sum()
    counts["% of Records"] = (counts["Records"] / total * 100).round(1) if total else 0
    total_row = {"Enumerator Gender": "Total", "Records": total, "% of Records": 100.0 if total else 0}
    return pd.concat([counts, pd.DataFrame([total_row])], ignore_index=True)


def quarterly_summary_table(farmer_on: pd.DataFrame, child_on: pd.DataFrame,
                             farmer_off: pd.DataFrame, child_off: pd.DataFrame) -> dict:
    """Matches Philemon's standing Quarterly Summary style: per-quarter unique
    counts (Annex-matched only), avg children per farmer, repeat-visit count,
    and excluded (not-on-list) totals reported separately."""
    rows = []
    for q in QUARTER_ORDER:
        f_q = farmer_on[farmer_on["Quarter"] == q] if not farmer_on.empty else pd.DataFrame()
        c_q = child_on[child_on["Quarter"] == q] if not child_on.empty else pd.DataFrame()
        if f_q.empty and c_q.empty:
            continue
        farmers_with_children = f_q[f_q["Children Profiled"] > 0] if not f_q.empty else pd.DataFrame()
        avg = round(farmers_with_children["Children Profiled"].mean(), 2) if len(farmers_with_children) else 0
        rows.append({
            "Quarter": q,
            "Unique Farmers Profiled": len(f_q),
            "Unique Children Profiled": len(c_q),
            "Avg Children per Farmer (with children)": avg,
        })
    total_row = {
        "Quarter": "Total",
        "Unique Farmers Profiled": len(farmer_on),
        "Unique Children Profiled": len(child_on),
        "Avg Children per Farmer (with children)": None,
    }
    summary = pd.DataFrame(rows + [total_row])

    repeat_visit_farmers = 0
    if not farmer_on.empty:
        repeat_visit_farmers = (farmer_on["Quarters Visited"].str.count(",") > 0).sum()

    excluded = pd.DataFrame([{
        "Farmers not on Annex list": len(farmer_off),
        "Children not on Annex list": len(child_off),
    }])

    return {"summary": summary, "repeat_visit_farmers": int(repeat_visit_farmers), "excluded": excluded}


def status_summary_table(child_on: pd.DataFrame, farmer_on: pd.DataFrame) -> dict:
    """Pivot of CLMRS case Status by quarter, for Annex-matched children only,
    plus the count of Annex-matched farmer households with no child under 18."""
    if child_on.empty or "Status" not in child_on.columns:
        return {"pivot": pd.DataFrame(), "no_child_farmers": 0}

    pivot = pd.pivot_table(
        child_on, index="Status", columns="Quarter", values="Child ID",
        aggfunc="count", fill_value=0, observed=True,
    )
    for q in QUARTER_ORDER:
        if q not in pivot.columns:
            pivot[q] = 0
    pivot = pivot[[q for q in QUARTER_ORDER if q in pivot.columns]]
    pivot["Total"] = pivot.sum(axis=1)
    pivot = pivot.sort_values("Total", ascending=False).reset_index()

    total_row = {"Status": "Total"}
    for col in pivot.columns:
        if col != "Status":
            total_row[col] = pivot[col].sum()
    pivot = pd.concat([pivot, pd.DataFrame([total_row])], ignore_index=True)

    no_child_farmers = int((farmer_on["Children Profiled"] == 0).sum()) if not farmer_on.empty else 0
    return {"pivot": pivot, "no_child_farmers": no_child_farmers}


def school_education_summary(child_on: pd.DataFrame) -> dict:
    if child_on.empty:
        return {"attendance": pd.DataFrame(), "education": pd.DataFrame()}

    attendance = pd.DataFrame()
    if "School Attendance" in child_on.columns:
        pivot = pd.pivot_table(
            child_on, index="School Attendance", columns="Quarter", values="Child ID",
            aggfunc="count", fill_value=0, observed=True,
        )
        for q in QUARTER_ORDER:
            if q not in pivot.columns:
                pivot[q] = 0
        pivot = pivot[[q for q in QUARTER_ORDER if q in pivot.columns]]
        pivot["Total"] = pivot.sum(axis=1)
        pivot = pivot.reset_index()
        total_row = {"School Attendance": "Total"}
        for col in pivot.columns:
            if col != "School Attendance":
                total_row[col] = pivot[col].sum()
        attendance = pd.concat([pivot, pd.DataFrame([total_row])], ignore_index=True)

    education = pd.DataFrame()
    if "Education Level" in child_on.columns:
        counts = child_on["Education Level"].dropna().value_counts().reset_index()
        counts.columns = ["Education Level", "Number of Children"]
        education = counts

    return {"attendance": attendance, "education": education}


def annex_match_by_group(farmer_on: pd.DataFrame, farmer_off: pd.DataFrame) -> pd.DataFrame:
    """Farmer List (Annex) match rate broken down by Farmer Group — surfaces
    uneven match rates that a single topline percentage would hide (e.g. one
    group's Annex list being out of date while others are near-complete)."""
    if farmer_on.empty and farmer_off.empty:
        return pd.DataFrame()
    on = farmer_on.copy()
    off = farmer_off.copy()
    on["_match"] = True
    off["_match"] = False
    combined = pd.concat([on, off], ignore_index=True)
    if "Farmer Group" not in combined.columns:
        return pd.DataFrame()
    grp = combined.groupby("Farmer Group")["_match"].agg(Present="sum", Total="count")
    grp["Absent"] = grp["Total"] - grp["Present"]
    grp["% Present"] = (grp["Present"] / grp["Total"] * 100).round(1)
    grp = grp.reset_index().sort_values("% Present")
    total_row = {
        "Farmer Group": "Total", "Present": grp["Present"].sum(), "Total": grp["Total"].sum(),
        "Absent": grp["Absent"].sum(),
        "% Present": round(grp["Present"].sum() / grp["Total"].sum() * 100, 1) if grp["Total"].sum() else 0,
    }
    return pd.concat([grp, pd.DataFrame([total_row])], ignore_index=True)


# ---------------------------------------------------------------- Household

def household_farmer_analysis(hh: pd.DataFrame) -> dict:
    """Unique farmer counts, with/without children, per sections 5-6."""
    farmers = hh.dropna(subset=["FARMER_ID"]).copy()
    no_child_mask = farmers["CHILD_ID"].isna()

    unique_farmers = farmers["FARMER_ID"].nunique()
    farmers_with_children = farmers.loc[~no_child_mask, "FARMER_ID"].nunique()
    # a farmer counts as "without children" only if EVERY one of their rows has no child
    farmer_has_any_child = farmers.groupby("FARMER_ID")["CHILD_ID"].apply(lambda s: s.notna().any())
    farmers_without_children_ids = farmer_has_any_child[~farmer_has_any_child].index
    farmers_without_children = len(farmers_without_children_ids)

    summary = pd.DataFrame([{
        "Total Household Records": len(farmers),
        "Unique Farmers Profiled": unique_farmers,
        "Farmers With Children": farmers_with_children,
        "Farmers Without Children": farmers_without_children,
        "% With Children": round(farmers_with_children / unique_farmers * 100, 1) if unique_farmers else 0,
        "% Without Children": round(farmers_without_children / unique_farmers * 100, 1) if unique_farmers else 0,
    }])

    fwc_rows = farmers[farmers["FARMER_ID"].isin(farmers_without_children_ids)].copy()
    fwc_cols = [c for c in [
        "FARMER_ID", "FARMER_NAME", "NO_CHILDREN_DATE", "SURVEYED_BY_NAME",
        "SURVEYED_BY_ID", "QUARTER", "FARMER_GROUP", "DISTRICT", "PLACE",
        "SOURCE_FILE", "SOURCE_ROW",
    ] if c in fwc_rows.columns]
    farmers_without_children_sheet = fwc_rows[fwc_cols].drop_duplicates(subset=["FARMER_ID"])

    return {
        "summary": summary,
        "farmers_without_children": farmers_without_children_sheet,
        "farmers_without_children_ids": set(farmers_without_children_ids),
    }


def unique_child_analysis(df: pd.DataFrame, gender_codes: dict, school_codes: dict) -> dict:
    """Unique child counts, gender, school attendance, per section 7-9.
    Works on Household or Inspection (same schema)."""
    children = df.dropna(subset=["CHILD_ID"]).copy()
    # one row per unique child: take the latest record by SURVEY_DATE if present
    if "SURVEY_DATE" in children.columns:
        children = children.sort_values("SURVEY_DATE")
    unique_children = children.drop_duplicates(subset=["CHILD_ID"], keep="last").copy()

    if "GENDER_CODE" in unique_children.columns:
        unique_children["GENDER_READABLE"] = unique_children["GENDER_CODE"].map(
            lambda c: gender_codes.get(str(c).strip().rstrip(".0") if pd.notna(c) else None, None)
        )
    if "SCHOOL_ATTENDANCE_CODE" in unique_children.columns:
        def _school_lookup(c):
            if pd.isna(c):
                return None
            key = str(c).strip()
            if key.endswith(".0"):
                key = key[:-2]
            return school_codes.get(key, None)
        unique_children["SCHOOL_READABLE"] = unique_children["SCHOOL_ATTENDANCE_CODE"].apply(_school_lookup)

    male = (unique_children.get("GENDER_READABLE") == "Male").sum() if "GENDER_READABLE" in unique_children else 0
    female = (unique_children.get("GENDER_READABLE") == "Female").sum() if "GENDER_READABLE" in unique_children else 0
    missing_gender = unique_children["GENDER_READABLE"].isna().sum() if "GENDER_READABLE" in unique_children else len(unique_children)
    attending = (unique_children.get("SCHOOL_READABLE") == "Attending School").sum() if "SCHOOL_READABLE" in unique_children else 0
    dropped = (unique_children.get("SCHOOL_READABLE") == "Dropped Out").sum() if "SCHOOL_READABLE" in unique_children else 0
    missing_school = unique_children["SCHOOL_READABLE"].isna().sum() if "SCHOOL_READABLE" in unique_children else len(unique_children)

    summary = pd.DataFrame([{
        "Total Child Records": len(children),
        "Unique Children Profiled": len(unique_children),
        "Unique Male Children": int(male),
        "Unique Female Children": int(female),
        "Children Attending School": int(attending),
        "Children Dropped Out": int(dropped),
        "Children Missing Gender": int(missing_gender),
        "Children Missing School Status": int(missing_school),
    }])

    detail_cols = [c for c in [
        "CHILD_ID", "CHILD_NAME", "FARMER_ID", "FARMER_NAME", "AGE",
        "GENDER_CODE", "GENDER_READABLE", "SCHOOL_ATTENDANCE_CODE", "SCHOOL_READABLE",
        "STATUS", "SURVEY_DATE", "QUARTER", "SURVEYED_BY_NAME", "FARMER_GROUP",
        "DISTRICT", "PLACE", "SOURCE_FILE", "SOURCE_ROW",
    ] if c in unique_children.columns]

    return {"summary": summary, "unique_children": unique_children[detail_cols]}


# ---------------------------------------------------------------- Enumerator

def enumerator_analysis(hh: pd.DataFrame, insp: pd.DataFrame, fu: pd.DataFrame,
                         farmers_without_children_ids: set) -> pd.DataFrame:
    rows = []

    def enum_col(df):
        return "SURVEYED_BY_NAME" if "SURVEYED_BY_NAME" in df.columns else None

    hh_col = enum_col(hh)
    if hh_col:
        hh_valid = hh.dropna(subset=["FARMER_ID"])
        g = hh_valid.groupby(hh_col)
        for name, sub in g:
            farmers_profiled = sub["FARMER_ID"].nunique()
            fwc = sub.loc[sub["FARMER_ID"].isin(farmers_without_children_ids), "FARMER_ID"].nunique()
            children_profiled = sub["CHILD_ID"].nunique() if "CHILD_ID" in sub.columns else 0
            gender = sub["ENUM_GENDER"].dropna().iloc[0] if "ENUM_GENDER" in sub.columns and sub["ENUM_GENDER"].notna().any() else None
            gender_code = sub["ENUM_GENDER_CODE"].dropna().iloc[0] if "ENUM_GENDER_CODE" in sub.columns and sub["ENUM_GENDER_CODE"].notna().any() else None
            rows.append({
                "Enumerator": name, "Enumerator Gender": gender, "Enumerator Gender Code": gender_code,
                "Farmers Profiled": farmers_profiled,
                "Farmers With Children": farmers_profiled - fwc,
                "Farmers Without Children": fwc,
                "Children Profiled": children_profiled,
                "Inspection Records": 0, "Follow Up Records": 0,
            })

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    insp_col = enum_col(insp)
    if insp_col and not insp.empty:
        insp_counts = insp.groupby(insp_col).size()
        df["Inspection Records"] = df["Enumerator"].map(insp_counts).fillna(0).astype(int)

    fu_col = enum_col(fu)
    if fu_col and not fu.empty:
        fu_counts = fu.groupby(fu_col).size()
        df["Follow Up Records"] = df["Enumerator"].map(fu_counts).fillna(0).astype(int)

    return df.sort_values("Farmers Profiled", ascending=False).reset_index(drop=True)


# ---------------------------------------------------------------- Farmer List matching

def farmer_list_matching(source_dfs: dict, farmer_list_ids: set) -> dict:
    """source_dfs: {'Household': df, 'Inspection': df, ...}. Every unique
    Farmer ID from each source is checked against the Annex farmer list."""
    rows = []
    for source_name, df in source_dfs.items():
        if df.empty or "FARMER_ID" not in df.columns:
            continue
        sub = df.dropna(subset=["FARMER_ID"]).drop_duplicates(subset=["FARMER_ID"])
        for _, r in sub.iterrows():
            rows.append({
                "Farmer ID": r["FARMER_ID"],
                "Farmer Name": r.get("FARMER_NAME"),
                "Source": source_name,
                "Tab P Status": "Present" if r["FARMER_ID"] in farmer_list_ids else "Absent",
                "Survey Date": r.get("SURVEY_DATE"),
                "Quarter": r.get("QUARTER"),
                "Farmer Group": r.get("FARMER_GROUP"),
            })
    detail = pd.DataFrame(rows)
    if detail.empty:
        summary = pd.DataFrame([{
            "Unique Farmers Analysed": 0, "Present in Farmer List": 0,
            "Absent from Farmer List": 0, "% Present": 0, "% Absent": 0,
        }])
        return {"detail": detail, "summary": summary}

    all_unique = detail.drop_duplicates(subset=["Farmer ID"])
    present = (all_unique["Tab P Status"] == "Present").sum()
    absent = (all_unique["Tab P Status"] == "Absent").sum()
    total = len(all_unique)
    summary = pd.DataFrame([{
        "Unique Farmers Analysed": total,
        "Present in Farmer List": int(present),
        "Absent from Farmer List": int(absent),
        "% Present": round(present / total * 100, 1) if total else 0,
        "% Absent": round(absent / total * 100, 1) if total else 0,
    }])
    return {"detail": detail, "summary": summary}


# ---------------------------------------------------------------- Inspection

def inspection_analysis(insp: pd.DataFrame) -> dict:
    farmers = insp.dropna(subset=["FARMER_ID"])
    children = insp.dropna(subset=["CHILD_ID"])
    summary = pd.DataFrame([{
        "Unique Farmers Inspected": farmers["FARMER_ID"].nunique(),
        "Unique Children Inspected": children["CHILD_ID"].nunique(),
        "Total Inspection Records": len(insp),
    }])

    by_quarter = pd.DataFrame()
    if "QUARTER" in insp.columns:
        by_quarter = insp.groupby("QUARTER").agg(
            Farmers_Inspected=("FARMER_ID", "nunique"),
            Children_Inspected=("CHILD_ID", "nunique"),
        ).reset_index().rename(columns={"QUARTER": "Quarter"})

    return {"summary": summary, "by_quarter": by_quarter}


# ---------------------------------------------------------------- HH vs Inspection

def household_vs_inspection(hh: pd.DataFrame, insp: pd.DataFrame) -> dict:
    hh_farmers = set(hh.dropna(subset=["FARMER_ID"])["FARMER_ID"])
    insp_farmers = set(insp.dropna(subset=["FARMER_ID"])["FARMER_ID"])
    hh_children = set(hh.dropna(subset=["CHILD_ID"])["CHILD_ID"])
    insp_children = set(insp.dropna(subset=["CHILD_ID"])["CHILD_ID"])

    farmer_summary = pd.DataFrame([
        {"Category": "Profiled and Inspected", "Count": len(hh_farmers & insp_farmers)},
        {"Category": "Profiled but Not Inspected", "Count": len(hh_farmers - insp_farmers)},
        {"Category": "Inspected but Not Profiled", "Count": len(insp_farmers - hh_farmers)},
    ])
    child_summary = pd.DataFrame([
        {"Category": "Profiled and Inspected", "Count": len(hh_children & insp_children)},
        {"Category": "Profiled but Not Inspected", "Count": len(hh_children - insp_children)},
        {"Category": "Inspected but Not Profiled", "Count": len(insp_children - hh_children)},
    ])
    return {"farmers": farmer_summary, "children": child_summary}


# ---------------------------------------------------------------- Follow Up

def follow_up_analysis(fu: pd.DataFrame, lag_months_map: dict) -> dict:
    events = fu.dropna(subset=["CHILD_ID"]).copy()
    if events.empty:
        empty = pd.DataFrame()
        return {"summary": empty, "cases": empty, "events": empty}

    events = add_case_identified_date(events, "SURVEY_DATE", "STATUS", lag_months_map)
    if "SURVEY_DATE" in events.columns:
        events = events.sort_values("SURVEY_DATE")

    # Case-level: LATEST record per child for status; FIRST record for quarter placement
    latest = events.drop_duplicates(subset=["CHILD_ID"], keep="last").set_index("CHILD_ID")
    first = events.drop_duplicates(subset=["CHILD_ID"], keep="first").set_index("CHILD_ID")

    cases = latest.copy()
    cases["FIRST_FOLLOWUP_DATE"] = first["SURVEY_DATE"]
    cases["FIRST_QUARTER"] = first["QUARTER"] if "QUARTER" in first.columns else None
    cases = cases.reset_index()

    status_norm = cases["STATUS"].astype(str).str.strip().str.lower()
    n_pass = (status_norm == "follow up pass").sum()
    n_fail = (status_norm == "follow up fail").sum()
    n_remediated = (status_norm == "fully remediated").sum()
    n_other = len(cases) - n_pass - n_fail - n_remediated

    summary = pd.DataFrame([{
        "Total Follow Up Records": len(events),
        "Unique Child IDs": events["CHILD_ID"].nunique(),
        "Unique Child Labour Cases": len(cases),
        "Follow Up Pass": int(n_pass),
        "Follow Up Fail": int(n_fail),
        "Fully Remediated": int(n_remediated),
        "Other/Unresolved": int(n_other),
    }])

    case_cols = [c for c in [
        "CHILD_ID", "CHILD_NAME", "FARMER_ID", "FARMER_NAME", "GENDER", "GENDER_CODE", "STATUS",
        "SURVEY_DATE", "CASE_IDENTIFIED_DATE", "QUARTER", "FIRST_FOLLOWUP_DATE",
        "FIRST_QUARTER", "SURVEYED_BY_NAME", "FARMER_GROUP", "SOURCE_FILE", "SOURCE_ROW",
    ] if c in cases.columns]

    return {"summary": summary, "cases": cases[case_cols], "events": events}


def follow_up_visit_timeline(fu: pd.DataFrame) -> dict:
    """For children with more than one Follow Up record: an event-level visit
    log (one row per visit, with days-since-previous-visit) plus a per-child
    summary (number of visits, gaps between consecutive visits, and \u2014 for
    children who reach Fully Remediated \u2014 the time from their FIRST follow-up
    to the remediation date)."""
    events = fu.dropna(subset=["CHILD_ID"]).copy()
    if events.empty or "SURVEY_DATE" not in events.columns:
        empty = pd.DataFrame()
        return {"events": empty, "summary": empty}

    events = events.sort_values(["CHILD_ID", "SURVEY_DATE"])
    events["VISIT_NUMBER"] = events.groupby("CHILD_ID").cumcount() + 1
    events["DAYS_SINCE_PREVIOUS"] = events.groupby("CHILD_ID")["SURVEY_DATE"].diff().dt.days

    event_cols = [c for c in [
        "CHILD_ID", "CHILD_NAME", "FARMER_ID", "FARMER_NAME", "VISIT_NUMBER",
        "SURVEY_DATE", "DAYS_SINCE_PREVIOUS", "STATUS", "SURVEYED_BY_NAME",
    ] if c in events.columns]
    event_log = events[event_cols].rename(columns={
        "CHILD_ID": "Child ID", "CHILD_NAME": "Child Name", "FARMER_ID": "Farmer ID",
        "FARMER_NAME": "Farmer Name", "VISIT_NUMBER": "Visit Number",
        "SURVEY_DATE": "Follow-Up Date", "DAYS_SINCE_PREVIOUS": "Days Since Previous Visit",
        "STATUS": "Status", "SURVEYED_BY_NAME": "Enumerator",
    })

    # Per-child summary
    summary_rows = []
    for child_id, sub in events.groupby("CHILD_ID", sort=False):
        sub = sub.sort_values("SURVEY_DATE")
        n_visits = len(sub)
        dates = list(sub["SURVEY_DATE"])
        statuses = list(sub["STATUS"])
        row = {
            "Child ID": child_id,
            "Child Name": sub["CHILD_NAME"].iloc[0] if "CHILD_NAME" in sub.columns else None,
            "Farmer ID": sub["FARMER_ID"].iloc[0] if "FARMER_ID" in sub.columns else None,
            "Farmer Name": sub["FARMER_NAME"].iloc[0] if "FARMER_NAME" in sub.columns else None,
            "Number of Follow-Ups": n_visits,
            "1st Follow-Up Date": dates[0] if n_visits >= 1 else None,
            "2nd Follow-Up Date": dates[1] if n_visits >= 2 else None,
            "Days: 1st to 2nd": (dates[1] - dates[0]).days if n_visits >= 2 else None,
            "3rd Follow-Up Date": dates[2] if n_visits >= 3 else None,
            "Days: 2nd to 3rd": (dates[2] - dates[1]).days if n_visits >= 3 else None,
            "Latest Status": statuses[-1],
        }
        remediated_idx = [i for i, s in enumerate(statuses) if str(s).strip().lower() == "fully remediated"]
        if remediated_idx:
            remediation_date = dates[remediated_idx[0]]
            row["Remediation Date"] = remediation_date
            row["Days: 1st Follow-Up to Remediation"] = (remediation_date - dates[0]).days
        else:
            row["Remediation Date"] = None
            row["Days: 1st Follow-Up to Remediation"] = None
        summary_rows.append(row)

    per_child_summary = pd.DataFrame(summary_rows).sort_values("Number of Follow-Ups", ascending=False).reset_index(drop=True)
    return {"events": event_log, "summary": per_child_summary}


def follow_up_alerts(fu_cases: pd.DataFrame, lag_months_map: dict, as_of=None) -> pd.DataFrame:
    """Children whose latest Follow Up status is NOT a closing status (Pass or
    Fully Remediated) \u2014 i.e. still need another follow-up. Next Due Date is
    the latest follow-up date plus the programme's fail-lag (default 3 months);
    Days Overdue is relative to `as_of` (defaults to today when the report is
    generated, so this stays current every time it's re-run)."""
    from dateutil.relativedelta import relativedelta
    import datetime as _dt

    if as_of is None:
        as_of = pd.Timestamp(_dt.date.today())
    if fu_cases.empty or "STATUS" not in fu_cases.columns:
        return pd.DataFrame()

    closing_statuses = {"fully remediated", "follow up pass"}
    status_norm = fu_cases["STATUS"].astype(str).str.strip().str.lower()
    needs_followup = fu_cases[~status_norm.isin(closing_statuses)].copy()
    if needs_followup.empty:
        return pd.DataFrame()

    fail_lag_months = lag_months_map.get("follow up fail", 3)
    date_col = "SURVEY_DATE" if "SURVEY_DATE" in needs_followup.columns else None
    needs_followup["Next Follow-Up Due"] = needs_followup[date_col].apply(
        lambda d: d + relativedelta(months=fail_lag_months) if pd.notna(d) else None
    ) if date_col else None
    needs_followup["Days Overdue"] = needs_followup["Next Follow-Up Due"].apply(
        lambda d: (as_of - d).days if pd.notna(d) else None
    )
    needs_followup["Alert"] = needs_followup["Days Overdue"].apply(
        lambda d: "OVERDUE" if pd.notna(d) and d > 0 else ("Due Soon" if pd.notna(d) else "")
    )

    cols = [c for c in [
        "CHILD_ID", "CHILD_NAME", "FARMER_ID", "FARMER_NAME", "GENDER", "STATUS",
        "SURVEY_DATE", "Next Follow-Up Due", "Days Overdue", "Alert",
        "SURVEYED_BY_NAME", "FARMER_GROUP",
    ] if c in needs_followup.columns]
    result = needs_followup[cols].rename(columns={
        "CHILD_ID": "Child ID", "CHILD_NAME": "Child Name", "FARMER_ID": "Farmer ID",
        "FARMER_NAME": "Farmer Name", "GENDER": "Gender", "STATUS": "Current Status",
        "SURVEY_DATE": "Last Follow-Up Date", "SURVEYED_BY_NAME": "Enumerator",
        "FARMER_GROUP": "Farmer Group",
    })
    return result.sort_values("Days Overdue", ascending=False, na_position="last").reset_index(drop=True)


def combined_farmers_without_children(hh_farmer_on: pd.DataFrame, hh_farmer_off: pd.DataFrame,
                                       insp_farmer_on: pd.DataFrame, insp_farmer_off: pd.DataFrame) -> dict:
    """Farmers with zero children profiled, combined across Household AND
    Inspection, each tagged by source, checked against the (updated) Annex
    farmer list. A farmer appearing in both datasets shows once per source
    (so the same farmer can appear twice, tagged Household and Inspection),
    plus a de-duplicated unique-farmer summary."""
    frames = []
    for source, df in [("Household", hh_farmer_on), ("Household", hh_farmer_off),
                        ("Inspection", insp_farmer_on), ("Inspection", insp_farmer_off)]:
        if df.empty or "Children Profiled" not in df.columns:
            continue
        sub = df[df["Children Profiled"] == 0].copy()
        if sub.empty:
            continue
        sub["Source"] = source
        frames.append(sub)

    if not frames:
        return {"detail": pd.DataFrame(), "unique_summary": pd.DataFrame()}

    combined = pd.concat(frames, ignore_index=True)
    cols = [c for c in [
        "Source", "Farmer ID", "Farmer Name", "Farmer Gender", "Farmer Group",
        "Quarter", "Enumerator Name", "Farmer List Match",
    ] if c in combined.columns]
    detail = combined[cols].sort_values(["Farmer ID", "Source"]).reset_index(drop=True)

    unique_farmers = combined.drop_duplicates(subset=["Farmer ID"])
    present = (unique_farmers["Farmer List Match"] == "Present").sum()
    total = len(unique_farmers)
    unique_summary = pd.DataFrame([{
        "Unique Farmers with No Child (Household + Inspection combined)": total,
        "Present on Annex Farmer List": int(present),
        "Absent from Annex Farmer List": int(total - present),
        "% Present": round(present / total * 100, 1) if total else 0,
    }])
    return {"detail": detail, "unique_summary": unique_summary}


# ---------------------------------------------------------------- Quarter analysis

def quarter_analysis(hh_fwc_ids: set, hh: pd.DataFrame, insp: pd.DataFrame,
                      fu_cases: pd.DataFrame) -> pd.DataFrame:
    quarters = ["Q1", "Q2", "Q3", "Q4"]
    rows = []
    for q in quarters:
        hh_q = hh[hh.get("QUARTER") == q] if "QUARTER" in hh.columns else hh.iloc[0:0]
        insp_q = insp[insp.get("QUARTER") == q] if "QUARTER" in insp.columns else insp.iloc[0:0]
        fu_q = fu_cases[fu_cases.get("FIRST_QUARTER") == q] if "FIRST_QUARTER" in fu_cases.columns else fu_cases.iloc[0:0]

        farmers_wo_children_q = hh_q.loc[hh_q["FARMER_ID"].isin(hh_fwc_ids), "FARMER_ID"].nunique()
        status_norm = fu_q["STATUS"].astype(str).str.strip().str.lower() if "STATUS" in fu_q.columns else pd.Series(dtype=str)

        rows.append({
            "Quarter": q,
            "Unique Farmers Profiled": hh_q["FARMER_ID"].nunique() if "FARMER_ID" in hh_q.columns else 0,
            "Unique Children Profiled": hh_q["CHILD_ID"].nunique() if "CHILD_ID" in hh_q.columns else 0,
            "Farmers Without Children": farmers_wo_children_q,
            "Unique Farmers Inspected": insp_q["FARMER_ID"].nunique() if "FARMER_ID" in insp_q.columns else 0,
            "Unique Children Inspected": insp_q["CHILD_ID"].nunique() if "CHILD_ID" in insp_q.columns else 0,
            "Follow Up Cases": len(fu_q),
            "Follow Up Pass": int((status_norm == "follow up pass").sum()),
            "Follow Up Fail": int((status_norm == "follow up fail").sum()),
            "Fully Remediated": int((status_norm == "fully remediated").sum()),
        })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------- Data quality

def data_quality_checks(hh: pd.DataFrame, insp: pd.DataFrame, fu: pd.DataFrame,
                         farmer_list_ids: set, gender_codes: dict, school_codes: dict) -> pd.DataFrame:
    checks = []

    def add(dataset, issue, count, sample_ids=None):
        checks.append({
            "Dataset": dataset, "Issue": issue, "Count": count,
            "Sample IDs": ", ".join(map(str, list(sample_ids)[:5])) if sample_ids else "",
        })

    for name, df in [("Household", hh), ("Inspection", insp), ("Follow Up", fu)]:
        if df.empty:
            continue
        if "FARMER_ID" in df.columns:
            missing = df["FARMER_ID"].isna().sum()
            if missing:
                add(name, "Missing Farmer ID", missing)
            dup = df.dropna(subset=["FARMER_ID"])["FARMER_ID"].duplicated().sum()
            # duplicates are expected in survey data (multiple children per farmer);
            # only flag true full-row duplicates
            full_dup = df.duplicated().sum()
            if full_dup:
                add(name, "Fully duplicate records", full_dup)
        if "CHILD_ID" in df.columns:
            dup_child = df.dropna(subset=["CHILD_ID"])["CHILD_ID"].duplicated().sum()
            # not inherently an error for follow-up (repeat visits expected)
        # Child-specific fields (gender code, school code) only apply to rows
        # that actually have a child on them — a farmer-with-no-child visit
        # legitimately has these blank, so scope the check to CHILD_ID rows.
        child_rows = df[df["CHILD_ID"].notna()] if "CHILD_ID" in df.columns else df
        if "GENDER_CODE" in df.columns:
            valid = set(gender_codes.keys())
            bad = child_rows["GENDER_CODE"].dropna()
            bad = bad[bad.astype(str).str.strip() != ""]
            bad = bad[~bad.astype(str).str.replace(r"\.0$", "", regex=True).isin(valid)]
            missing = child_rows["GENDER_CODE"].isna().sum() + (child_rows["GENDER_CODE"].astype(str).str.strip() == "").sum()
            if len(bad):
                add(name, "Invalid gender code (on child records)", len(bad))
            if missing:
                add(name, "Missing gender code (on child records)", missing)
        if "SCHOOL_ATTENDANCE_CODE" in df.columns:
            valid = set(school_codes.keys())
            bad = child_rows["SCHOOL_ATTENDANCE_CODE"].dropna()
            bad = bad[bad.astype(str).str.strip() != ""]
            bad = bad[~bad.astype(str).str.replace(r"\.0$", "", regex=True).isin(valid)]
            if len(bad):
                add(name, "Invalid school attendance code (on child records)", len(bad))
        if "SURVEY_DATE" in df.columns:
            missing_dates = child_rows["SURVEY_DATE"].isna().sum()
            if missing_dates:
                add(name, "Missing/unparseable survey date (on child records)", missing_dates)
        if "FARMER_ID" in df.columns and farmer_list_ids:
            ids = set(df["FARMER_ID"].dropna())
            absent = ids - farmer_list_ids
            if absent:
                add(name, "Farmer IDs absent from Farmer List", len(absent), absent)

    if not fu.empty and not hh.empty and "CHILD_ID" in fu.columns and "CHILD_ID" in hh.columns:
        fu_children = set(fu["CHILD_ID"].dropna())
        hh_children = set(hh["CHILD_ID"].dropna())
        insp_children = set(insp["CHILD_ID"].dropna()) if "CHILD_ID" in insp.columns else set()
        orphan = fu_children - hh_children - insp_children
        if orphan:
            add("Follow Up", "Child IDs not found in Household or Inspection", len(orphan), orphan)

    return pd.DataFrame(checks)
