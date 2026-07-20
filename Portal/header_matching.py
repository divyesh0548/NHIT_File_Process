"""
Shared normalization for header keyword / column matching.

Rules: strip, casefold, remove all whitespace (including between words).
DB keywords and file header cells are compared using this form.
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
    return "".join(ch for ch in s if not ch.isspace())
