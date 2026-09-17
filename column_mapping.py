"""
Maps a workbook's actual column headers to internal canonical field names,
using a programme's saved mapping first and falling back to fuzzy matching
for anything unrecognised (so a new programme with slightly different header
text still mostly auto-maps, and only genuinely new headers need the user's
input in the mapping screen).
"""
from __future__ import annotations
import difflib
from typing import Optional


def _norm(s: str) -> str:
    s = str(s).replace("\u2019", "'").replace("\u2018", "'")  # curly -> straight apostrophe
    return " ".join(s.strip().lower().split())


def map_headers(raw_headers: list, header_map: dict, fuzzy_threshold: float = 0.86):
    """
    Returns:
      mapping: dict {raw_header: internal_field or None}
      unmapped: list of raw headers that could not be confidently mapped
    Matching order: exact normalised match -> fuzzy match against known
    normalised keys -> unmapped (left for the user to map manually).
    """
    norm_map = {_norm(k): v for k, v in header_map.items()}
    known_norms = list(norm_map.keys())

    mapping = {}
    unmapped = []
    used_fields = set()

    for raw in raw_headers:
        n = _norm(raw)
        if n in norm_map:
            field = norm_map[n]
            mapping[raw] = field
            used_fields.add(field)
            continue
        # fuzzy fallback
        close = difflib.get_close_matches(n, known_norms, n=1, cutoff=fuzzy_threshold)
        if close:
            field = norm_map[close[0]]
            if field not in used_fields:  # don't double-assign a field fuzzily
                mapping[raw] = field
                used_fields.add(field)
                continue
        mapping[raw] = None
        unmapped.append(raw)

    return mapping, unmapped


def apply_mapping(df, mapping: dict):
    """Rename a dataframe's columns to internal field names; columns with no
    mapping are prefixed EXTRA__ and kept (never dropped, for traceability
    per the auditability requirement)."""
    rename = {}
    seen_fields = set()
    for raw, field in mapping.items():
        if field is None:
            rename[raw] = f"EXTRA__{raw}"
        elif field in seen_fields:
            # a field already claimed by an earlier column (shouldn't normally
            # happen) -> keep this one as an extra so nothing is silently lost
            rename[raw] = f"EXTRA__{raw}"
        else:
            rename[raw] = field
            seen_fields.add(field)
    return df.rename(columns=rename)
