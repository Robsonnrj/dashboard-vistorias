# -*- coding: utf-8 -*-
import unicodedata
import pandas as pd
from typing import List, Optional

def norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii")
    return s.strip().lower()

def pick_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    if df is None or df.empty:
        return None
    for cand in candidates:
        for c in df.columns:
            if norm(c) == norm(cand):
                return c
    for cand in candidates:
        tgt = norm(cand)
        for c in df.columns:
            if tgt in norm(c):
                return c
    return None
