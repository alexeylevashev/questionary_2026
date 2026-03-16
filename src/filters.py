"""Stage 1 – spatial filtering of survey trips by origin/destination zones."""

from __future__ import annotations

from typing import List

import geopandas as gpd
import pandas as pd

from .coords import parse_and_fix_coords
from .config import Config
from .io_utils import find_date_column


DAY_WORKDAYS = "будни"
DAY_WEEKENDS = "выходные"
DAY_ALL = "все"


def run_filter(
    cfg: Config,
    gdf_zones: gpd.GeoDataFrame,
    df_surveys: pd.DataFrame,
    origin_zones: List[str],
    dest_zones: List[str],
    day_type: str,
) -> pd.DataFrame:
    """Filter survey trips by zone membership and day type.

    Args:
        cfg          – project config
        gdf_zones    – GeoDataFrame with zone polygons (EPSG:4326)
        df_surveys   – raw survey DataFrame (normalised column names)
        origin_zones – list of zone names selected as origins
        dest_zones   – list of zone names selected as destinations
        day_type     – one of DAY_WORKDAYS / DAY_WEEKENDS / DAY_ALL

    Returns:
        Filtered DataFrame with added columns:
            zone_start, zone_end,
            _geom_start (Point), _geom_end (Point)
    """
    cols = cfg.columns
    name_field = cfg.boundary_name_field
    df = df_surveys.copy()

    # ------------------------------------------------------------------
    # 1. Date parsing and day-type filter
    # ------------------------------------------------------------------
    date_col = find_date_column(df, cols.date)

    df[date_col] = pd.to_datetime(
        df[date_col].astype(str).str.strip(),
        format="%Y-%m-%d",
        errors="coerce",
    )
    # Fallback for rows that didn't parse with strict format
    bad_mask = df[date_col].isna()
    if bad_mask.any():
        df.loc[bad_mask, date_col] = pd.to_datetime(
            df.loc[bad_mask, date_col].astype(str).str.strip(),
            errors="coerce",
            dayfirst=False,
        )

    missing = int(df[date_col].isna().sum())
    if missing:
        print(f"  Исключено строк без даты: {missing}")
    df = df.dropna(subset=[date_col]).copy()

    is_weekend = df[date_col].dt.dayofweek >= 5  # 5=Sat, 6=Sun

    if day_type == DAY_WORKDAYS:
        df = df.loc[~is_weekend].copy()
    elif day_type == DAY_WEEKENDS:
        df = df.loc[is_weekend].copy()
    # DAY_ALL – keep everything

    print(f"  Строк после фильтра по дням ({day_type}): {len(df)}")

    if df.empty:
        return df

    # ------------------------------------------------------------------
    # 2. Coordinate parsing and auto-correction
    # ------------------------------------------------------------------
    zones_union = gdf_zones.geometry.union_all() if hasattr(gdf_zones.geometry, "union_all") \
        else gdf_zones.geometry.unary_union

    def _parse_col(coord_series):
        results = coord_series.apply(
            lambda x: pd.Series(parse_and_fix_coords(x, zones_union))
        )
        results.columns = ["_geom", "_swapped", "_fixed_str"]
        return results

    origin_parsed = _parse_col(df[cols.origin_coords])
    dest_parsed = _parse_col(df[cols.dest_coords])

    df["_geom_start"] = origin_parsed["_geom"].values
    df["_geom_end"] = dest_parsed["_geom"].values

    # Fix swapped coordinate strings in the original columns
    swap_start = origin_parsed["_swapped"].values
    swap_end = dest_parsed["_swapped"].values

    df.loc[swap_start, cols.origin_coords] = origin_parsed.loc[
        origin_parsed["_swapped"], "_fixed_str"
    ].values
    df.loc[swap_end, cols.dest_coords] = dest_parsed.loc[
        dest_parsed["_swapped"], "_fixed_str"
    ].values

    swapped_count = int(swap_start.sum()) + int(swap_end.sum())
    if swapped_count:
        print(f"  Исправлено перепутанных координат: {swapped_count}")

    # Drop rows where coordinates couldn't be parsed
    before = len(df)
    df = df.dropna(subset=["_geom_start", "_geom_end"]).copy()
    dropped = before - len(df)
    if dropped:
        print(f"  Исключено строк с нераспознанными координатами: {dropped}")

    # ------------------------------------------------------------------
    # 3. Spatial join: assign zone to each trip endpoint
    # ------------------------------------------------------------------
    zones_slim = gdf_zones[[name_field, "geometry"]].copy()

    gdf_start = gpd.GeoDataFrame(df, geometry="_geom_start", crs=cfg.crs)
    joined_start = gpd.sjoin(gdf_start, zones_slim, how="left", predicate="within")
    # sjoin may produce duplicates if a point touches multiple zones; keep first
    joined_start = joined_start[~joined_start.index.duplicated(keep="first")]
    df["zone_start"] = joined_start[name_field].values

    gdf_end = gpd.GeoDataFrame(df, geometry="_geom_end", crs=cfg.crs)
    joined_end = gpd.sjoin(gdf_end, zones_slim, how="left", predicate="within")
    joined_end = joined_end[~joined_end.index.duplicated(keep="first")]
    df["zone_end"] = joined_end[name_field].values

    # ------------------------------------------------------------------
    # 4. Filter by origin/destination zones (bidirectional)
    # ------------------------------------------------------------------
    mask_fwd = (
        df["zone_start"].isin(origin_zones) & df["zone_end"].isin(dest_zones)
    )
    mask_bwd = (
        df["zone_start"].isin(dest_zones) & df["zone_end"].isin(origin_zones)
    )

    result = df[mask_fwd | mask_bwd].copy()
    print(f"  Подходящих поездок найдено: {len(result)}")

    return result
