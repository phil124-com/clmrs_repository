"""
Programme configuration system for CLMRS Data Analyzer.

Each programme defines:
  - column mapping (raw header -> internal field name) per dataset type
  - status/code value meanings
  - quarter convention
  - farmer groups (optional, informational)

Programmes are stored as JSON under %APPDATA%/CLMRS_Analyzer/programmes/ (or
./programmes/ during development) so the user's own mappings persist and can
be edited without touching code. This module defines the BUILT-IN default
for the Butter (LS Country) programme, verified against real Butter export
files (Household/Inspection/Follow-up share one 50+ column ODK/Kobo schema,
distinguished by the "Survey Type" field; Farmer List / Annex has its own
10-column schema).
"""
from __future__ import annotations
import json
import os
from dataclasses import dataclass, field, asdict
from typing import Optional


def user_data_dir() -> str:
    """Where programme configs and logs live, next to the .exe, no admin rights needed."""
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    path = os.path.join(base, "CLMRS_Analyzer")
    os.makedirs(path, exist_ok=True)
    return path


def programmes_dir() -> str:
    path = os.path.join(user_data_dir(), "programmes")
    os.makedirs(path, exist_ok=True)
    return path


# Internal canonical field names used throughout the engine, independent of
# whatever header text a given programme's export happens to use.
CORE_SURVEY_FIELDS = [
    "FARMER_ID", "FARMER_NAME", "CHILD_ID", "CHILD_NAME", "STATUS", "ACTIVE",
    "AGE", "GENDER", "GENDER_CODE", "REGISTERED_BY_ID", "REGISTERED_BY_NAME",
    "DATE_REGISTERED", "COUNTRY", "REGION", "DISTRICT", "PLACE",
    "FARMER_GROUP", "SECTION", "NO_CHILDREN_FLAG", "NO_CHILDREN_DATE",
    "FLAG_REASON", "REMEDIATION_PROPOSED", "FOLLOWUP_REASON",
    "SURVEYED_BY_ID", "SURVEYED_BY_NAME", "ENUM_GENDER", "ENUM_GENDER_CODE",
    "SURVEY_DATE", "VISITED_BY_ID", "VISITED_BY_NAME", "GPS_LAT", "GPS_LON",
    "SURVEY_TYPE", "RELATIONSHIP", "SCHOOL_ATTENDANCE", "SCHOOL_ATTENDANCE_CODE",
]

CORE_FARMERLIST_FIELDS = [
    "FARMER_ID", "FARMER_NAME", "SOCIETY", "DISTRICT", "TOTAL_AREA",
    "NUM_PLOTS", "GENDER", "FARMER_STATUS", "GPS_LAT", "GPS_LON",
]

# Values that a raw header might take across programmes, mapped to the
# canonical field. Matching is case-insensitive and whitespace-trimmed.
DEFAULT_SURVEY_HEADER_MAP = {
    "farmer id": "FARMER_ID", "farmer_id": "FARMER_ID", "farmerid": "FARMER_ID",
    "farmer name": "FARMER_NAME",
    "child id": "CHILD_ID", "child_id": "CHILD_ID", "childid": "CHILD_ID",
    "child name": "CHILD_NAME",
    "status": "STATUS",
    "active": "ACTIVE",
    "age": "AGE",
    "gender": "GENDER",
    "child gender code": "GENDER_CODE",
    "enumerator id who registered the child": "REGISTERED_BY_ID",
    "enumerator name who registered the child": "REGISTERED_BY_NAME",
    "date when child registered": "DATE_REGISTERED",
    "country": "COUNTRY",
    "region": "REGION",
    "district": "DISTRICT",
    "place": "PLACE",
    "farmer group": "FARMER_GROUP",
    "section": "SECTION",
    "farmer visited with no children": "NO_CHILDREN_FLAG",
    "date when farmer visited with no children": "NO_CHILDREN_DATE",
    "reasons for flagging": "FLAG_REASON",
    "remediation action proposed": "REMEDIATION_PROPOSED",
    "reasons for follow-up pass/fail": "FOLLOWUP_REASON",
    "enumerator id who surveyed the child": "SURVEYED_BY_ID",
    "enumerator name who surveyed the child": "SURVEYED_BY_NAME",
    "enumerator gender": "ENUM_GENDER",
    "enumerator gender code": "ENUM_GENDER_CODE",
    "date when the survey was taken": "SURVEY_DATE",
    "enumerator id who visited": "VISITED_BY_ID",
    "enumerator name who visited": "VISITED_BY_NAME",
    "gps coordinates of survey(latitude)": "GPS_LAT",
    "gps coordinates of survey(longitude)": "GPS_LON",
    "survey type": "SURVEY_TYPE",
    "relationship with farmer/worker": "RELATIONSHIP",
    "school attendance": "SCHOOL_ATTENDANCE",
    "school attendance code": "SCHOOL_ATTENDANCE_CODE",
    "if attending school or dropped out of school, what is the child's current level of education?": "EDUCATION_LEVEL",
}

DEFAULT_FARMERLIST_HEADER_MAP = {
    "farmer id": "FARMER_ID",
    "farmer name": "FARMER_NAME",
    "society": "SOCIETY",
    "district": "DISTRICT",
    "total area": "TOTAL_AREA",
    "number of plots": "NUM_PLOTS",
    "gender": "GENDER",
    "farmer status": "FARMER_STATUS",
    "household gps latitude": "GPS_LAT",
    "household gps longitude": "GPS_LON",
}


@dataclass
class ProgrammeConfig:
    name: str
    survey_header_map: dict = field(default_factory=lambda: dict(DEFAULT_SURVEY_HEADER_MAP))
    farmerlist_header_map: dict = field(default_factory=lambda: dict(DEFAULT_FARMERLIST_HEADER_MAP))
    gender_codes: dict = field(default_factory=lambda: {"1": "Male", "2": "Female"})
    school_codes: dict = field(default_factory=lambda: {"0": "Dropped Out", "1": "Attending School"})
    # Quarter convention: month number (1-12) -> quarter label. Butter's fiscal
    # convention is Oct-Dec=Q1 ... Jul-Sep=Q4, per standing training-analysis rules.
    quarter_map: dict = field(default_factory=lambda: {
        10: "Q1", 11: "Q1", 12: "Q1",
        1: "Q2", 2: "Q2", 3: "Q2",
        4: "Q3", 5: "Q3", 6: "Q3",
        7: "Q4", 8: "Q4", 9: "Q4",
    })
    farmer_groups: list = field(default_factory=list)
    # Case-identification lag rules, keyed by follow-up STATUS value (lowercased).
    case_id_lag_months: dict = field(default_factory=lambda: {
        "follow up pass": 3, "follow up fail": 3, "fully remediated": 6,
    })
    notes: str = ""

    def to_json(self) -> str:
        return json.dumps(asdict(self), indent=2)

    @staticmethod
    def from_json(text: str) -> "ProgrammeConfig":
        d = json.loads(text)
        return ProgrammeConfig(**d)

    def save(self):
        path = os.path.join(programmes_dir(), f"{self.name}.json")
        with open(path, "w", encoding="utf-8") as f:
            f.write(self.to_json())

    @staticmethod
    def load(name: str) -> Optional["ProgrammeConfig"]:
        path = os.path.join(programmes_dir(), f"{name}.json")
        if not os.path.exists(path):
            return None
        with open(path, "r", encoding="utf-8") as f:
            return ProgrammeConfig.from_json(f.read())

    @staticmethod
    def list_saved() -> list:
        d = programmes_dir()
        return sorted(f[:-5] for f in os.listdir(d) if f.endswith(".json"))


def butter_default() -> ProgrammeConfig:
    """Built-in default for the Butter (LS Country) programme, verified against
    real 2025-2026 Household/Inspection/Follow-up exports and the combined
    Annex farmer list covering all four farmer groups."""
    cfg = ProgrammeConfig(
        name="Butter",
        farmer_groups=["CCA ASSIN FOSU", "SWEDRU", "CAPE COAST", "B/ASIKUMA", "B/ASIKUMA B"],
        notes=(
            "Household/Inspection/Follow-up share one Kobo export schema, split by "
            "Survey Type. Annex farmer list is combined across all four farmer groups. "
            "A child can appear across multiple datasets; Status reflects the child's "
            "current cumulative CLMRS case status, not a per-row-only value."
        ),
    )
    return cfg


def ensure_builtin_programmes():
    """Seed the built-in Butter programme on first run if not already saved
    (a user's own edits to it are never overwritten)."""
    if ProgrammeConfig.load("Butter") is None:
        butter_default().save()
