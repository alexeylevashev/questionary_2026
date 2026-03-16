"""I/O utilities: reading files, normalising text, parsing reference dictionaries."""

from __future__ import annotations

import re
from collections import defaultdict
from pathlib import Path
from typing import Optional

import geopandas as gpd
import pandas as pd


# ---------------------------------------------------------------------------
# Text normalisation
# ---------------------------------------------------------------------------

def normalize_excel_text(x) -> Optional[str]:
    """Normalise a cell value from Excel: collapse whitespace, strip, lower-case.

    Handles various Unicode space characters and line-break artefacts produced
    by Excel.  Returns None for blank/NaN values.
    """
    if pd.isna(x):
        return None
    s = str(x)
    # Various Unicode spaces
    s = s.replace("\u00A0", " ")   # NBSP
    s = s.replace("\u2007", " ")   # figure space
    s = s.replace("\u202F", " ")   # narrow NBSP
    # Line breaks
    s = s.replace("\r", " ").replace("\n", " ")
    # Collapse tabs / multiple spaces
    s = re.sub(r"[ \t]+", " ", s).strip()
    return s.lower() if s else None


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Strip whitespace and remove line-break artefacts from column names."""
    df = df.copy()
    df.columns = (
        df.columns.astype(str)
        .str.replace("\r", "", regex=False)
        .str.replace("\n", " ", regex=False)
        .str.strip()
    )
    return df


# ---------------------------------------------------------------------------
# Column lookup helpers
# ---------------------------------------------------------------------------

def find_col(df: pd.DataFrame, target: str) -> str:
    """Return the exact column name or the first column that starts with *target*.

    Raises KeyError if nothing found.
    """
    if target in df.columns:
        return target
    candidates = [c for c in df.columns if str(c).strip().startswith(target)]
    if not candidates:
        raise KeyError(
            f"Не найден столбец '{target}'. Доступные: {list(df.columns)}"
        )
    return candidates[0]


def find_optional_col(df: pd.DataFrame, target: str) -> Optional[str]:
    """Same as find_col but returns None instead of raising."""
    if target in df.columns:
        return target
    candidates = [c for c in df.columns if str(c).strip().startswith(target)]
    return candidates[0] if candidates else None


def find_date_column(df: pd.DataFrame, target: str = "Дата перемещений") -> str:
    """Locate the date column even if its name has minor artefacts."""
    if target in df.columns:
        return target
    candidates = [c for c in df.columns if "Дата перемещ" in str(c)]
    if candidates:
        return candidates[0]
    raise KeyError(
        f"Не найден столбец с датой перемещений. Ожидалось '{target}'. "
        f"Доступные: {list(df.columns)}"
    )


# ---------------------------------------------------------------------------
# Transport helpers
# ---------------------------------------------------------------------------

def first_transport_value(x) -> Optional[str]:
    """Return the first transport type from a comma/semicolon-separated string."""
    if pd.isna(x):
        return None
    s = str(x).strip()
    if not s:
        return None
    parts = re.split(r"[;,/\n|]+", s)
    first = parts[0].strip() if parts else s.strip()
    return first if first else None


def is_blank(x) -> bool:
    if pd.isna(x):
        return True
    return str(x).strip() == ""


# ---------------------------------------------------------------------------
# Reference-file parser (Пары.xlsx, Соцгруппы.xlsx, Транспорт.xlsx)
# ---------------------------------------------------------------------------

def parse_columns_as_groups(path_xlsx: str | Path) -> tuple[dict, dict]:
    """Parse a reference workbook where columns are group names and cells are variants.

    Returns:
        mapping   – {normalized_value: group_name}
        conflicts – {normalized_value: [groupA, groupB, ...]} for values appearing
                    in more than one group column
    """
    df = pd.read_excel(path_xlsx, header=0, dtype=str)
    # Drop Unnamed columns that pandas adds for empty header cells
    df = df.loc[:, ~df.columns.astype(str).str.lower().str.startswith("unnamed")]

    mapping: dict[str, str] = {}
    conflicts: dict[str, list] = defaultdict(list)

    for col in df.columns:
        group_name = str(col).strip()
        if not group_name:
            continue
        for raw_val in df[col].dropna().tolist():
            key = normalize_excel_text(raw_val)
            if not key:
                continue
            if key in mapping and mapping[key] != group_name:
                if mapping[key] not in conflicts[key]:
                    conflicts[key].append(mapping[key])
                if group_name not in conflicts[key]:
                    conflicts[key].append(group_name)
                continue
            mapping[key] = group_name

    return mapping, dict(conflicts)


# ---------------------------------------------------------------------------
# File loaders
# ---------------------------------------------------------------------------

def load_surveys(path: str | Path) -> pd.DataFrame:
    """Load the survey workbook and normalise column names.

    Falls back to CSV reading if the file is named .xlsx but contains CSV data.
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Файл анкет не найден: {path}")

    try:
        df = pd.read_excel(path)
    except Exception:
        try:
            df = pd.read_csv(path)
        except Exception as e:
            raise RuntimeError(f"Не удалось прочитать файл анкет '{path}': {e}") from e

    return normalize_columns(df)


def load_geojson(path: str | Path, crs: str = "EPSG:4326") -> gpd.GeoDataFrame:
    """Load a GeoJSON file and ensure it is in the requested CRS."""
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Файл границ не найден: {path}")

    gdf = gpd.read_file(path)
    if gdf.crs is None:
        gdf = gdf.set_crs(crs)
    else:
        gdf = gdf.to_crs(crs)
    return gdf


# ---------------------------------------------------------------------------
# Output path helper
# ---------------------------------------------------------------------------

def safe_output_path(path: str | Path) -> Path:
    """If *path* is locked (open in Excel etc.), return an available variant.

    Tries <stem>_1.xlsx, <stem>_2.xlsx, ... up to 200 attempts.
    """
    p = Path(path)
    parent = p.parent if str(p.parent) else Path(".")
    stem = p.stem
    suffix = p.suffix or ".xlsx"

    def _can_open(candidate: Path) -> bool:
        try:
            with open(candidate, "a+b"):
                return True
        except (PermissionError, OSError):
            return False

    if _can_open(p):
        return p

    for i in range(1, 200):
        cand = parent / f"{stem}_{i}{suffix}"
        if _can_open(cand):
            return cand

    raise PermissionError(f"Не удалось подобрать доступное имя для файла: {path}")
