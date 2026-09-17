"""
Orchestrates the full analysis: load -> quarter-tag -> run every analysis
module -> package results for Excel export. This is what the GUI's
"RUN COMPLETE ANALYSIS" button calls.

Report style follows Philemon's standing household-profiling reports:
one row per unique farmer/child (raw + analytical fields together),
Annex-matched vs Not-on-List kept on separate sheets, with Quarterly/Status/
School summary sheets pivoted from the detail — applied consistently across
Household, Inspection, and Follow Up.
"""
from __future__ import annotations
import pandas as pd
from .data_loader import load_survey_file, load_farmer_list
from .quarters import add_quarter_column
from . import analysis as A


class AnalysisResult:
    def __init__(self):
        self.warnings = []
        self.sheets = {}       # sheet_name -> DataFrame or dict of DataFrames
        self.exec_summary = {}
        self.meta = {}


def run_complete_analysis(programme_cfg, farmer_list_path: str, household_path: str,
                           inspection_path: str, follow_up_path: str,
                           progress_cb=None) -> AnalysisResult:
    def report(msg):
        if progress_cb:
            progress_cb(msg)

    result = AnalysisResult()

    report("Loading Farmer List...")
    fl = load_farmer_list(farmer_list_path, programme_cfg.farmerlist_header_map)
    result.warnings += [f"[Farmer List] {w}" for w in fl.warnings]
    farmer_list_ids = set(fl.df["FARMER_ID"].dropna()) if "FARMER_ID" in fl.df.columns else set()
    farmer_gender_map = {}
    if "FARMER_ID" in fl.df.columns and "GENDER" in fl.df.columns:
        fl_gender = fl.df.dropna(subset=["FARMER_ID"]).drop_duplicates(subset=["FARMER_ID"])
        genders = fl_gender["GENDER"].astype(str).str.strip().str.title()
        farmer_gender_map = dict(zip(fl_gender["FARMER_ID"], genders))

    report("Loading Household data...")
    hh_r = load_survey_file(household_path, "household", programme_cfg.survey_header_map)
    result.warnings += [f"[Household] {w}" for w in hh_r.warnings]
    hh = hh_r.df
    if "SURVEY_DATE" in hh.columns:
        add_quarter_column(hh, "SURVEY_DATE", programme_cfg.quarter_map)

    report("Loading Inspection data...")
    insp_r = load_survey_file(inspection_path, "inspection", programme_cfg.survey_header_map)
    result.warnings += [f"[Inspection] {w}" for w in insp_r.warnings]
    insp = insp_r.df
    if "SURVEY_DATE" in insp.columns:
        add_quarter_column(insp, "SURVEY_DATE", programme_cfg.quarter_map)

    report("Loading Follow Up data...")
    fu_r = load_survey_file(follow_up_path, "follow_up", programme_cfg.survey_header_map)
    result.warnings += [f"[Follow Up] {w}" for w in fu_r.warnings]
    fu = fu_r.df
    if "SURVEY_DATE" in fu.columns:
        add_quarter_column(fu, "SURVEY_DATE", programme_cfg.quarter_map)

    report("Checking data quality...")
    dq = A.data_quality_checks(hh, insp, fu, farmer_list_ids,
                                programme_cfg.gender_codes, programme_cfg.school_codes)

    # ---------------------------------------------------------------- Household
    report("Building Household farmer & child detail...")
    hh_farmer = A.farmer_detail_table(hh, farmer_list_ids, farmer_gender_map)
    hh_child = A.child_detail_table(hh, farmer_list_ids, programme_cfg.gender_codes, programme_cfg.school_codes)
    hh_quarterly = A.quarterly_summary_table(hh_farmer["on_list"], hh_child["on_list"],
                                              hh_farmer["not_on_list"], hh_child["not_on_list"])
    hh_status = A.status_summary_table(hh_child["on_list"], hh_farmer["on_list"])
    hh_school = A.school_education_summary(hh_child["on_list"])
    hh_match_by_group = A.annex_match_by_group(hh_farmer["on_list"], hh_farmer["not_on_list"])
    hh_child_gender = A.child_gender_summary(hh_child["on_list"])
    hh_farmer_gender = A.farmer_gender_summary(hh_farmer["on_list"])
    hh_enum_gender = A.enumerator_gender_summary(hh)

    # ---------------------------------------------------------------- Inspection
    report("Building Inspection farmer & child detail...")
    insp_farmer = A.farmer_detail_table(insp, farmer_list_ids, farmer_gender_map)
    insp_child = A.child_detail_table(insp, farmer_list_ids, programme_cfg.gender_codes, programme_cfg.school_codes)
    insp_quarterly = A.quarterly_summary_table(insp_farmer["on_list"], insp_child["on_list"],
                                                insp_farmer["not_on_list"], insp_child["not_on_list"])
    insp_match_by_group = A.annex_match_by_group(insp_farmer["on_list"], insp_farmer["not_on_list"])
    insp_child_gender = A.child_gender_summary(insp_child["on_list"])
    insp_farmer_gender = A.farmer_gender_summary(insp_farmer["on_list"])
    insp_enum_gender = A.enumerator_gender_summary(insp)

    report("Comparing Household vs Inspection...")
    hh_vs_insp = A.household_vs_inspection(hh, insp)

    # ---------------------------------------------------------------- Follow Up
    report("Analysing Follow Up and case identification...")
    fu_result = A.follow_up_analysis(fu, programme_cfg.case_id_lag_months)
    fu_cases = fu_result["cases"].copy()
    if not fu_cases.empty and "FARMER_ID" in fu_cases.columns:
        fu_cases["Farmer List Match"] = fu_cases["FARMER_ID"].apply(
            lambda fid: "Present" if fid in farmer_list_ids else "Absent"
        )
    fu_timeline = A.follow_up_visit_timeline(fu)
    fu_alerts = A.follow_up_alerts(fu_result["cases"], programme_cfg.case_id_lag_months)

    report("Building Enumerator Analysis...")
    farmers_without_children_ids = set(
        hh_farmer["on_list"].loc[hh_farmer["on_list"]["Children Profiled"] == 0, "Farmer ID"]
    ) if not hh_farmer["on_list"].empty else set()
    enum_analysis = A.enumerator_analysis(hh, insp, fu, farmers_without_children_ids)

    report("Combining farmers with no children across Household and Inspection...")
    no_children_combined = A.combined_farmers_without_children(
        hh_farmer["on_list"], hh_farmer["not_on_list"],
        insp_farmer["on_list"], insp_farmer["not_on_list"],
    )

    result.sheets = {
        "hh_farmer_on_list": hh_farmer["on_list"],
        "hh_farmer_not_on_list": hh_farmer["not_on_list"],
        "hh_child_on_list": hh_child["on_list"],
        "hh_child_not_on_list": hh_child["not_on_list"],
        "hh_quarterly": hh_quarterly,
        "hh_status": hh_status,
        "hh_school": hh_school,
        "hh_match_by_group": hh_match_by_group,
        "hh_child_gender": hh_child_gender,
        "hh_farmer_gender": hh_farmer_gender,
        "hh_enum_gender": hh_enum_gender,
        "insp_farmer_on_list": insp_farmer["on_list"],
        "insp_farmer_not_on_list": insp_farmer["not_on_list"],
        "insp_child_on_list": insp_child["on_list"],
        "insp_child_not_on_list": insp_child["not_on_list"],
        "insp_quarterly": insp_quarterly,
        "insp_match_by_group": insp_match_by_group,
        "insp_child_gender": insp_child_gender,
        "insp_farmer_gender": insp_farmer_gender,
        "insp_enum_gender": insp_enum_gender,
        "hh_vs_insp": hh_vs_insp,
        "fu_summary": fu_result["summary"],
        "fu_cases": fu_cases,
        "fu_timeline_events": fu_timeline["events"],
        "fu_timeline_summary": fu_timeline["summary"],
        "fu_alerts": fu_alerts,
        "enumerator": enum_analysis,
        "no_children_combined_detail": no_children_combined["detail"],
        "no_children_combined_summary": no_children_combined["unique_summary"],
        "data_quality": dq,
    }

    farmer_groups = sorted(set(hh.get("FARMER_GROUP", pd.Series(dtype=str)).dropna().unique()))
    date_min = hh["SURVEY_DATE"].min() if "SURVEY_DATE" in hh.columns and hh["SURVEY_DATE"].notna().any() else None
    date_max = hh["SURVEY_DATE"].max() if "SURVEY_DATE" in hh.columns and hh["SURVEY_DATE"].notna().any() else None
    result.meta = {
        "farmer_groups": farmer_groups,
        "date_min": date_min,
        "date_max": date_max,
        "farmer_list_file": fl.df["SOURCE_FILE"].iloc[0] if not fl.df.empty else "",
        "farmer_list_count": len(farmer_list_ids),
    }

    result.exec_summary = {
        "Unique Farmers Profiled (On List)": len(hh_farmer["on_list"]),
        "Unique Children Profiled (On List)": len(hh_child["on_list"]),
        "Farmers with No Child Under 18 (On List)": int((hh_farmer["on_list"]["Children Profiled"] == 0).sum()) if not hh_farmer["on_list"].empty else 0,
        "Farmers NOT on Annex List (Household)": len(hh_farmer["not_on_list"]),
        "Children NOT on Annex List (Household)": len(hh_child["not_on_list"]),
        "Unique Farmers Inspected (On List)": len(insp_farmer["on_list"]),
        "Unique Children Inspected (On List)": len(insp_child["on_list"]),
        "Farmers NOT on Annex List (Inspection)": len(insp_farmer["not_on_list"]),
        "Unique Follow Up Cases": fu_result["summary"]["Unique Child Labour Cases"].iloc[0] if not fu_result["summary"].empty else 0,
        "Follow Up Pass": fu_result["summary"]["Follow Up Pass"].iloc[0] if not fu_result["summary"].empty else 0,
        "Follow Up Fail": fu_result["summary"]["Follow Up Fail"].iloc[0] if not fu_result["summary"].empty else 0,
        "Fully Remediated": fu_result["summary"]["Fully Remediated"].iloc[0] if not fu_result["summary"].empty else 0,
        "Follow Up Alerts (Next Visit Due/Overdue)": len(fu_alerts),
        "Data Quality Exceptions": len(dq),
    }

    report("Analysis complete.")
    return result
