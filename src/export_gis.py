"""GIS export: GeoPackage, GeoJSON, and OD desire lines."""

from __future__ import annotations

from pathlib import Path
from typing import Optional

import geopandas as gpd
import pandas as pd
from shapely.geometry import LineString, Point

from .config import Config


# Columns to include as attributes in GIS layers (drop internal helpers)
_DROP_PREFIXES = ("_geom", "_swapped", "_fixed")


def _attribute_df(df: pd.DataFrame) -> pd.DataFrame:
    """Drop geometry and internal working columns."""
    drop_cols = [
        c for c in df.columns
        if any(c.startswith(p) for p in _DROP_PREFIXES)
        or c in ("index_right",)
    ]
    return df.drop(columns=drop_cols, errors="ignore")


def _make_origins_gdf(df: pd.DataFrame, crs: str) -> gpd.GeoDataFrame:
    """Build a GeoDataFrame of trip origin points."""
    mask = df["_geom_start"].notna()
    sub = df.loc[mask].copy()
    attrs = _attribute_df(sub)
    return gpd.GeoDataFrame(attrs, geometry=sub.loc[mask, "_geom_start"].values, crs=crs)


def _make_destinations_gdf(df: pd.DataFrame, crs: str) -> gpd.GeoDataFrame:
    """Build a GeoDataFrame of trip destination points."""
    mask = df["_geom_end"].notna()
    sub = df.loc[mask].copy()
    attrs = _attribute_df(sub)
    return gpd.GeoDataFrame(attrs, geometry=sub.loc[mask, "_geom_end"].values, crs=crs)


def _make_od_lines_gdf(df: pd.DataFrame, crs: str) -> gpd.GeoDataFrame:
    """Build a GeoDataFrame of OD desire lines (LineString from origin to destination)."""
    mask = df["_geom_start"].notna() & df["_geom_end"].notna()
    sub = df.loc[mask].copy()

    lines = [
        LineString([row["_geom_start"], row["_geom_end"]])
        for _, row in sub.iterrows()
    ]
    attrs = _attribute_df(sub)
    return gpd.GeoDataFrame(attrs, geometry=lines, crs=crs)


def export_gis(
    cfg: Config,
    df: pd.DataFrame,
    gdf_zones: gpd.GeoDataFrame,
    run_tag: str,
) -> dict[str, Path]:
    """Export GIS layers to GeoPackage and/or GeoJSON.

    Layers:
        origins      – trip origin points
        destinations – trip destination points
        od_lines     – desire lines (origin → destination)
        boundaries   – study-area zone polygons

    Returns a dict mapping layer names and format keys to file paths.
    """
    if "_geom_start" not in df.columns or "_geom_end" not in df.columns:
        print("  GIS-экспорт пропущен: геометрия не найдена в данных.")
        return {}

    crs = cfg.crs
    layer_names = cfg.gis_export.layers
    formats = cfg.gis_export.formats

    origins_gdf = _make_origins_gdf(df, crs)
    destinations_gdf = _make_destinations_gdf(df, crs)
    od_lines_gdf = _make_od_lines_gdf(df, crs)
    boundaries_gdf = gdf_zones.to_crs(crs)

    layers = {
        layer_names.get("origins", "origins"): origins_gdf,
        layer_names.get("destinations", "destinations"): destinations_gdf,
        layer_names.get("od_lines", "od_lines"): od_lines_gdf,
        layer_names.get("boundaries", "boundaries"): boundaries_gdf,
    }

    paths: dict[str, Path] = {}

    # GeoPackage — all layers in one file
    if "gpkg" in formats:
        gpkg_path = cfg.output_dir / f"{run_tag}_gis.gpkg"
        # Remove existing file to avoid layer append issues
        if gpkg_path.exists():
            gpkg_path.unlink()

        for layer_name, gdf in layers.items():
            if gdf.empty:
                continue
            _safe_write_gpkg(gdf, gpkg_path, layer_name)

        paths["gpkg"] = gpkg_path
        print(f"  GeoPackage сохранён: {gpkg_path.name}")

    # GeoJSON — one file per layer
    if "geojson" in formats:
        for layer_name, gdf in layers.items():
            if gdf.empty:
                continue
            geojson_path = cfg.output_dir / f"{run_tag}_{layer_name}.geojson"
            _safe_write_geojson(gdf, geojson_path)
            paths[f"geojson_{layer_name}"] = geojson_path
        print(f"  GeoJSON файлы сохранены в {cfg.output_dir}")

    # Shapefile
    if "shp" in formats:
        for layer_name, gdf in layers.items():
            if gdf.empty:
                continue
            shp_dir = cfg.output_dir / f"{run_tag}_{layer_name}_shp"
            shp_dir.mkdir(exist_ok=True)
            shp_path = shp_dir / f"{layer_name}.shp"
            _safe_write_shp(gdf, shp_path)
            paths[f"shp_{layer_name}"] = shp_path
        print(f"  Shapefile файлы сохранены в {cfg.output_dir}")

    return paths


# ---------------------------------------------------------------------------
# Safe write helpers (stringify non-serialisable columns before export)
# ---------------------------------------------------------------------------

def _prepare_gdf(gdf: gpd.GeoDataFrame) -> gpd.GeoDataFrame:
    """Convert problematic column types to strings for file export."""
    gdf = gdf.copy()
    for col in gdf.columns:
        if col == gdf.geometry.name:
            continue
        if gdf[col].dtype == object:
            # Check if any value is not a basic type
            try:
                sample = gdf[col].dropna().head(5).tolist()
                if any(not isinstance(v, (str, int, float, bool)) for v in sample):
                    gdf[col] = gdf[col].astype(str)
            except Exception:
                gdf[col] = gdf[col].astype(str)
    return gdf


def _safe_write_gpkg(gdf: gpd.GeoDataFrame, path: Path, layer: str):
    try:
        _prepare_gdf(gdf).to_file(path, layer=layer, driver="GPKG")
    except Exception as e:
        print(f"  Предупреждение: ошибка записи слоя '{layer}' в GPKG: {e}")


def _safe_write_geojson(gdf: gpd.GeoDataFrame, path: Path):
    try:
        _prepare_gdf(gdf).to_file(path, driver="GeoJSON")
    except Exception as e:
        print(f"  Предупреждение: ошибка записи GeoJSON '{path.name}': {e}")


def _safe_write_shp(gdf: gpd.GeoDataFrame, path: Path):
    try:
        _prepare_gdf(gdf).to_file(path, driver="ESRI Shapefile")
    except Exception as e:
        print(f"  Предупреждение: ошибка записи Shapefile '{path.name}': {e}")
