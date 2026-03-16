"""CLI entry point for the transport survey processing tool.

Run:
    python -m src.cli
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import List


def _ask_single(options: List[str], prompt: str) -> str:
    """Show a numbered menu and return the chosen option (single choice)."""
    print(f"\n{prompt}")
    for i, opt in enumerate(options, start=1):
        print(f"  {i}. {opt}")
    while True:
        try:
            raw = input("Введите номер: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nОтменено.")
            sys.exit(0)
        try:
            idx = int(raw) - 1
            if 0 <= idx < len(options):
                return options[idx]
        except ValueError:
            pass
        print("Некорректный ввод. Введите число из списка.")


def _ask_multi(options: List[str], prompt: str) -> List[str]:
    """Show a numbered menu and return one or more chosen options."""
    print(f"\n{prompt}")
    for i, opt in enumerate(options, start=1):
        print(f"  {i}. {opt}")
    while True:
        try:
            raw = input("Введите номер(а) через запятую: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nОтменено.")
            sys.exit(0)
        try:
            indices = [int(x.strip()) - 1 for x in raw.split(",")]
            selected = [options[i] for i in indices if 0 <= i < len(options)]
            if selected:
                return selected
        except (ValueError, IndexError):
            pass
        print("Некорректный ввод. Попробуйте снова.")


def _build_run_tag(
    origin_zones: List[str],
    dest_zones: List[str],
    day_type: str,
) -> str:
    """Build a short filesystem-safe tag for output file names."""
    def _slug(names: List[str]) -> str:
        combined = "_".join(names)
        # Keep Cyrillic, Latin, digits, underscore; replace the rest with -
        import re
        return re.sub(r"[^\w]", "-", combined, flags=re.UNICODE)[:40]

    return f"{day_type}_{_slug(origin_zones)}_to_{_slug(dest_zones)}"


def main():
    print("=" * 60)
    print("  ТРАНСПОРТНОЕ ОБСЛЕДОВАНИЕ — обработка анкет")
    print("=" * 60)

    # ------------------------------------------------------------------
    # 1. Load configuration and input data
    # ------------------------------------------------------------------
    try:
        from .config import load_config
        from .io_utils import load_surveys, load_geojson
    except ImportError:
        from src.config import load_config
        from src.io_utils import load_surveys, load_geojson

    cfg = load_config()

    print("\nЗагрузка входных данных...")
    try:
        gdf_zones = load_geojson(cfg.boundaries_path, crs=cfg.crs)
    except FileNotFoundError as e:
        print(f"Ошибка: {e}")
        sys.exit(1)

    try:
        df_surveys = load_surveys(cfg.surveys_path)
    except FileNotFoundError as e:
        print(f"Ошибка: {e}")
        sys.exit(1)

    print(f"  Загружено строк анкет: {len(df_surveys)}")
    print(f"  Загружено зон: {len(gdf_zones)}")

    # ------------------------------------------------------------------
    # 2. User menu — 3 questions only
    # ------------------------------------------------------------------
    name_field = cfg.boundary_name_field
    if name_field not in gdf_zones.columns:
        # Try to find an alternative name field
        candidates = [c for c in gdf_zones.columns if "name" in c.lower()]
        if candidates:
            name_field = candidates[0]
            print(f"  Поле названия территории: '{name_field}'")
        else:
            print(
                f"Ошибка: поле '{cfg.boundary_name_field}' не найдено в GeoJSON. "
                f"Доступные поля: {list(gdf_zones.columns)}"
            )
            sys.exit(1)

    zone_names = sorted(gdf_zones[name_field].astype(str).unique().tolist())

    origin_zones = _ask_multi(zone_names, "Выберите территории ОТПРАВЛЕНИЯ:")
    dest_zones = _ask_multi(zone_names, "Выберите территории ПРИБЫТИЯ:")
    day_type = _ask_single(
        ["будни", "выходные", "все"],
        "Выберите тип дней:",
    )

    run_tag = _build_run_tag(origin_zones, dest_zones, day_type)
    print(f"\nМетка запуска: {run_tag}")
    print(f"Результаты будут сохранены в: {cfg.output_dir}\n")

    # ------------------------------------------------------------------
    # 3. Automatic pipeline
    # ------------------------------------------------------------------
    try:
        from .filters import run_filter, DAY_WORKDAYS, DAY_WEEKENDS, DAY_ALL
        from .status import run_status
        from .od_matrix import run_od
        from .eva import run_eva
        from .excel_report import (
            write_filter_report,
            write_status_report,
            write_od_report,
            write_eva_report,
        )
        from .export_gis import export_gis
        from .qgis_project import write_qgis_project
    except ImportError:
        from src.filters import run_filter, DAY_WORKDAYS, DAY_WEEKENDS, DAY_ALL
        from src.status import run_status
        from src.od_matrix import run_od
        from src.eva import run_eva
        from src.excel_report import (
            write_filter_report,
            write_status_report,
            write_od_report,
            write_eva_report,
        )
        from src.export_gis import export_gis
        from src.qgis_project import write_qgis_project

    # Map user-facing day labels to internal constants
    _day_map = {"будни": DAY_WORKDAYS, "выходные": DAY_WEEKENDS, "все": DAY_ALL}
    day_type_internal = _day_map[day_type]

    # --- Stage 1: Filter ---
    print("─" * 50)
    print("Этап 1 / 4 — Фильтрация по территориям...")
    df = run_filter(cfg, gdf_zones, df_surveys, origin_zones, dest_zones, day_type_internal)

    if df.empty:
        print("Нет данных после фильтрации. Измените условия выборки.")
        sys.exit(0)

    write_filter_report(cfg, df, run_tag)

    # --- Stage 2: Social groups ---
    print("─" * 50)
    print("Этап 2 / 4 — Назначение соцгрупп...")
    df, stats_df = run_status(cfg, df)
    write_status_report(cfg, df, stats_df, run_tag)

    # --- Stage 3: OD matrices ---
    print("─" * 50)
    print("Этап 3 / 4 — OD-матрицы и транспорт...")
    df = run_od(cfg, df)
    write_od_report(cfg, df, run_tag)

    # --- Stage 4: EVA ---
    print("─" * 50)
    print("Этап 4 / 4 — EVA-анализ (BW-модель)...")
    eva_results = run_eva(cfg, df)
    write_eva_report(cfg, eva_results, run_tag)

    # --- GIS export ---
    print("─" * 50)
    print("Экспорт GIS-слоёв...")
    gis_paths = export_gis(cfg, df, gdf_zones, run_tag)

    gpkg_path = gis_paths.get("gpkg")
    write_qgis_project(cfg, gpkg_path, run_tag)

    # ------------------------------------------------------------------
    # 4. Done
    # ------------------------------------------------------------------
    print("─" * 50)
    print(f"\nГотово! Все файлы сохранены в: {cfg.output_dir}")
    print("Созданные файлы:")
    for f in sorted(cfg.output_dir.iterdir()):
        if run_tag in f.name:
            print(f"  {f.name}")


if __name__ == "__main__":
    main()
