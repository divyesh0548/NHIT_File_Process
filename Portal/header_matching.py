"""
Shared normalization for header keyword / column matching.

Rules applied to both DB keywords and file header cells before compare:
strip, casefold (lowercase), remove whitespace, remove symbols.
Only letters and digits remain.
"""


def normalize_header_match(value) -> str:
    if value is None:
        return ""
    try:
        import pandas as pd

        if pd.isna(value):
            return ""
    except Exception:
        pass
    s = str(value).strip().casefold()
    return "".join(ch for ch in s if ch.isalnum())
