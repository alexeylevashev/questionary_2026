"""Stage 3 – OD pairs, transport groups, and origin-destination matrices."""

from __future__ import annotations

import pandas as pd

from .config import Config
from .io_utils import (
    parse_columns_as_groups,
    normalize_excel_text,
    first_transport_value,
    is_blank,
)


def run_od(cfg: Config, df: pd.DataFrame) -> pd.DataFrame:
    """Assign OD groups and transport groups; return enriched DataFrame.

    Adds columns to *df*:
        _o_grp          – origin point group (from Пары.xlsx)
        _d_grp          – destination point group
        _od_pair        – "<origin_group> - <dest_group>"
        _transport_group – transport group (from Транспорт.xlsx);
                           blank Transport field → "Пешком"

    Args:
        cfg – project config
        df  – DataFrame after run_status (must have Пункт отправления /
              Пункт прибытия / Транспорт columns)

    Returns:
        Enriched DataFrame.
    """
    cols = cfg.columns

    # ------------------------------------------------------------------
    # 1. Load reference dictionaries
    # ------------------------------------------------------------------
    pairs_map, pairs_conflicts = parse_columns_as_groups(cfg.pairs_reference_path)
    transport_map, transport_conflicts = parse_columns_as_groups(
        cfg.transport_reference_path
    )

    if pairs_conflicts:
        print(
            f"  Внимание: конфликты в Пары.xlsx "
            f"({len(pairs_conflicts)} значений в нескольких группах)"
        )
    if transport_conflicts:
        print(
            f"  Внимание: конфликты в Транспорт.xlsx "
            f"({len(transport_conflicts)} значений в нескольких группах)"
        )

    # ------------------------------------------------------------------
    # 2. Assign OD groups
    # ------------------------------------------------------------------
    df_out = df.copy()

    origin_col = cols.origin_point
    dest_col = cols.dest_point

    df_out["_o_key"] = df_out[origin_col].apply(normalize_excel_text)
    df_out["_d_key"] = df_out[dest_col].apply(normalize_excel_text)

    df_out["_o_grp"] = df_out["_o_key"].map(pairs_map).fillna("Не определено")
    df_out["_d_grp"] = df_out["_d_key"].map(pairs_map).fillna("Не определено")
    df_out["_od_pair"] = df_out["_o_grp"].astype(str) + " - " + df_out["_d_grp"].astype(str)

    unmatched_o = df_out.loc[
        (df_out["_o_grp"] == "Не определено") & df_out[origin_col].notna(),
        origin_col,
    ].unique()
    unmatched_d = df_out.loc[
        (df_out["_d_grp"] == "Не определено") & df_out[dest_col].notna(),
        dest_col,
    ].unique()
    if len(unmatched_o) or len(unmatched_d):
        print(
            f"  Не сопоставлено пунктам: "
            f"отправления – {len(unmatched_o)}, "
            f"прибытия – {len(unmatched_d)} уникальных значений"
        )

    # ------------------------------------------------------------------
    # 3. Assign transport group
    # ------------------------------------------------------------------
    transport_col = cols.transport

    def _map_transport(x):
        if is_blank(x):
            return "Пешком"
        first = first_transport_value(x)
        if first is None:
            return "Пешком"
        key = normalize_excel_text(first)
        return transport_map.get(key, "Не определено")

    df_out["_transport_group"] = df_out[transport_col].apply(_map_transport)

    # ------------------------------------------------------------------
    # 4. Summary
    # ------------------------------------------------------------------
    pair_counts = df_out["_od_pair"].value_counts()
    transport_counts = df_out["_transport_group"].value_counts()
    print(f"  OD-пар уникальных: {len(pair_counts)}")
    print(f"  Транспортных групп: {len(transport_counts)}")

    return df_out


def build_od_matrices(
    df: pd.DataFrame,
    id_col: str,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Build absolute and per-respondent OD matrices.

    Returns:
        mat_abs      – pivot table: rows=origin group, cols=destination group, values=trip count
        mat_per_resp – same divided by unique respondent count
    """
    total_resp = int(df[id_col].nunique()) or 1
    mat_abs = pd.pivot_table(
        df,
        index="_o_grp",
        columns="_d_grp",
        values=id_col,
        aggfunc="size",
        fill_value=0,
    )
    mat_per_resp = mat_abs / total_resp
    return mat_abs, mat_per_resp
