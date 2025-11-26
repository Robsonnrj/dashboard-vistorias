import unicodedata
import pandas as pd
from typing import List, Optional

def norm(s: str) -> str:
    """Normaliza string removendo acento, caixa e espaços laterais."""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii")
    return s.strip().lower()

def pick_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    """
    Retorna a primeira coluna do DataFrame que bate com um dos nomes candidatos.
    Normaliza nomes para tolerar variações.
    """
    if df is None or df.empty:
        return None
    col_norm = {norm(c): c for c in df.columns}
    for cand in candidates:
        n_cand = norm(cand)
        if n_cand in col_norm:
            return col_norm[n_cand]
    for cand in candidates:
        tgt = norm(cand)
        for nc, orig in col_norm.items():
            if tgt in nc:
                return orig
    return None

def clean(x) -> str:
    return "" if pd.isna(x) else str(x).strip()

def date_or(x, default: pd.Timestamp) -> pd.Timestamp:
    d = pd.to_datetime(x, errors="coerce")
    return d if pd.notna(d) else default
