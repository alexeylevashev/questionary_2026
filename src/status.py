"""Stage 2 – assign social groups and compute mobility statistics."""

from __future__ import annotations

import pandas as pd

from .config import Config
from .io_utils import parse_columns_as_groups


def run_status(
    cfg: Config,
    df: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Assign a social group to every row and compute group-level statistics.

    Social groups are defined in Соцгруппы.xlsx: each column header is a group
    name and the cells below are variants of the social-status field value that
    belong to that group.

    Args:
        cfg – project config
        df  – filtered survey DataFrame (output of run_filter)

    Returns:
        df_out   – input DataFrame with added column 'Соцгруппа'
        stats_df – summary table: group, trips, respondents, avg trips per person
    """
    cols = cfg.columns

    # ------------------------------------------------------------------
    # 1. Load reference and build mapping
    # ------------------------------------------------------------------
    mapping, conflicts = parse_columns_as_groups(cfg.social_groups_reference_path)

    if conflicts:
        print(
            f"  Внимание: конфликты в Соцгруппы.xlsx "
            f"({len(conflicts)} значений в нескольких группах)"
        )

    # ------------------------------------------------------------------
    # 2. Assign social group
    # ------------------------------------------------------------------
    df_out = df.copy()
    status_col = cols.social_status

    if status_col not in df_out.columns:
        raise KeyError(f"Не найден столбец '{status_col}' в данных.")

    # Normalise for lookup: lower + strip (mirrors parse_columns_as_groups)
    normalised = (
        df_out[status_col]
        .astype(str)
        .str.strip()
        .str.lower()
        .str.replace(r"\s+", " ", regex=True)
    )
    df_out["Соцгруппа"] = normalised.map(mapping).fillna("Не определено")

    unmatched = df_out.loc[
        (df_out["Соцгруппа"] == "Не определено")
        & df_out[status_col].notna()
        & (df_out[status_col].astype(str).str.strip() != ""),
        status_col,
    ].unique()
    if len(unmatched):
        print(
            f"  Не сопоставлено соцгруппе {len(unmatched)} уникальных статусов. "
            f"Примеры: {list(unmatched[:5])}"
        )

    # ------------------------------------------------------------------
    # 3. Compute statistics
    # ------------------------------------------------------------------
    id_col = cols.id

    stats_df = (
        df_out.groupby("Соцгруппа")
        .agg(
            Передвижения=(id_col, "count"),
            Респонденты=(id_col, pd.Series.nunique),
        )
        .reset_index()
    )
    stats_df["Среднее передвижений"] = (
        stats_df["Передвижения"] / stats_df["Респонденты"]
    ).round(2)

    total_resp = int(df_out[id_col].nunique())
    total_trips = len(df_out)
    avg_overall = round(total_trips / total_resp, 2) if total_resp else 0

    print(
        f"  Респондентов: {total_resp}, "
        f"передвижений: {total_trips}, "
        f"средняя подвижность: {avg_overall}"
    )

    return df_out, stats_df
