"""Load and expose project configuration from config.yaml."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import List

import yaml


@dataclass
class BwModelConfig:
    bin_minutes: int
    max_minutes: int
    e_bounds: tuple[float, float]
    f_bounds: tuple[float, float]
    g_bounds: tuple[float, float]


@dataclass
class ColumnsConfig:
    id: str
    social_status: str
    date: str
    origin_address: str
    origin_coords: str
    dest_address: str
    dest_coords: str
    origin_point: str
    dest_point: str
    departure_time: str
    arrival_time: str
    transport: str
    comment: str


@dataclass
class GisExportConfig:
    formats: List[str]
    layers: dict


@dataclass
class QgisProjectConfig:
    filename: str
    title: str


@dataclass
class Config:
    # Resolved absolute paths
    data_dir: Path
    output_dir: Path
    surveys_path: Path
    boundaries_path: Path
    pairs_reference_path: Path
    social_groups_reference_path: Path
    transport_reference_path: Path

    # Geo settings
    crs: str
    boundary_name_field: str

    # Analysis settings
    columns: ColumnsConfig
    bw_model: BwModelConfig
    stat_fields_by_transport: List[str]
    stat_fields_simple: List[str]

    # GIS export
    gis_export: GisExportConfig

    # QGIS project
    qgis_project: QgisProjectConfig


def load_config(config_path: str | Path | None = None) -> Config:
    """Load config.yaml and return a Config object with resolved paths."""
    if config_path is None:
        # Look for config.yaml next to the project root (one level above src/)
        src_dir = Path(__file__).parent
        config_path = src_dir.parent / "config.yaml"

    config_path = Path(config_path)
    if not config_path.exists():
        raise FileNotFoundError(f"Файл конфигурации не найден: {config_path}")

    with open(config_path, encoding="utf-8") as f:
        raw = yaml.safe_load(f)

    project_root = config_path.parent

    paths = raw["paths"]
    data_dir = project_root / paths["data_dir"]
    output_dir = project_root / paths["output_dir"]

    # Create output dir if it doesn't exist
    output_dir.mkdir(parents=True, exist_ok=True)

    geo = raw.get("geo", {})
    analysis = raw.get("analysis", {})
    cols_raw = analysis.get("columns", {})
    bw_raw = analysis.get("bw_model", {})
    gis_raw = raw.get("gis_export", {})
    qgis_raw = raw.get("qgis_project", {})

    return Config(
        data_dir=data_dir,
        output_dir=output_dir,
        surveys_path=data_dir / paths["surveys"],
        boundaries_path=data_dir / paths["boundaries"],
        pairs_reference_path=data_dir / paths["pairs_reference"],
        social_groups_reference_path=data_dir / paths["social_groups_reference"],
        transport_reference_path=data_dir / paths["transport_reference"],
        crs=geo.get("crs", "EPSG:4326"),
        boundary_name_field=geo.get("boundary_name_field", "NAME"),
        columns=ColumnsConfig(
            id=cols_raw.get("id", "ID"),
            social_status=cols_raw.get("social_status", "Социальный статус"),
            date=cols_raw.get("date", "Дата перемещений"),
            origin_address=cols_raw.get("origin_address", "Адрес отправления"),
            origin_coords=cols_raw.get("origin_coords", "Координаты отправления"),
            dest_address=cols_raw.get("dest_address", "Адрес прибытия"),
            dest_coords=cols_raw.get("dest_coords", "Координаты прибытия"),
            origin_point=cols_raw.get("origin_point", "Пункт отправления"),
            dest_point=cols_raw.get("dest_point", "Пункт прибытия"),
            departure_time=cols_raw.get("departure_time", "Время отправления"),
            arrival_time=cols_raw.get("arrival_time", "Время прибытия"),
            transport=cols_raw.get("transport", "Транспорт"),
            comment=cols_raw.get("comment", "Комментарий"),
        ),
        bw_model=BwModelConfig(
            bin_minutes=bw_raw.get("bin_minutes", 5),
            max_minutes=bw_raw.get("max_minutes", 120),
            e_bounds=tuple(bw_raw.get("e_bounds", [0.01, 20.0])),
            f_bounds=tuple(bw_raw.get("f_bounds", [-20.0, 20.0])),
            g_bounds=tuple(bw_raw.get("g_bounds", [-2.0, 2.0])),
        ),
        stat_fields_by_transport=analysis.get("stat_fields_by_transport", []),
        stat_fields_simple=analysis.get("stat_fields_simple", []),
        gis_export=GisExportConfig(
            formats=gis_raw.get("formats", ["gpkg", "geojson"]),
            layers=gis_raw.get("layers", {}),
        ),
        qgis_project=QgisProjectConfig(
            filename=qgis_raw.get("filename", "transport_survey.qgs"),
            title=qgis_raw.get("title", "Транспортное обследование"),
        ),
    )
