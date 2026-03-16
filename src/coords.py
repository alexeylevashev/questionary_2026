"""Coordinate parsing and auto-correction for survey data.

Survey data often contains swapped lat/lon values.  This module detects and
fixes such errors by checking whether a point falls within the study-area
polygons.
"""

from __future__ import annotations

from typing import Optional

import pandas as pd
from shapely.geometry import Point


def parse_and_fix_coords(
    coord_str,
    zones_union=None,
) -> tuple[Optional[Point], bool, Optional[str]]:
    """Parse a coordinate string and fix swapped lat/lon if needed.

    The function tries both interpretations of the two numbers:
    - candidate A: (v1, v2) treated as (lon, lat)
    - candidate B: (v2, v1) treated as (lon, lat)  [swapped]

    If *zones_union* (a Shapely geometry covering the study area) is provided,
    the candidate that falls within it is preferred.  Falls back to a simple
    heuristic (|v| > 90 → longitude) when both or neither candidate is inside.

    Returns:
        point      – shapely.geometry.Point(lon, lat) or None
        swapped    – True if the original string had lat/lon in wrong order
        fixed_str  – corrected "lon, lat" string, or None if parsing failed
    """
    if pd.isna(coord_str) or str(coord_str).strip() == "":
        return None, False, None

    def _valid(lon: float, lat: float) -> bool:
        return (-180 <= lon <= 180) and (-90 <= lat <= 90)

    try:
        clean = str(coord_str).replace('"', "").replace("'", "").strip()
        parts = clean.split(",") if "," in clean else clean.split()
        if len(parts) != 2:
            return None, False, None

        p0, p1 = str(parts[0]).strip(), str(parts[1]).strip()
        v1, v2 = float(p0), float(p1)

        cand_a = (v1, v2)  # as-is: (lon, lat)
        cand_b = (v2, v1)  # swapped: (lon, lat)

        if zones_union is not None:
            in_a = _valid(*cand_a) and Point(*cand_a).within(zones_union)
            in_b = _valid(*cand_b) and Point(*cand_b).within(zones_union)

            if in_a and not in_b:
                return Point(*cand_a), False, f"{p0}, {p1}"
            if in_b and not in_a:
                return Point(*cand_b), True, f"{p1}, {p0}"
            # Both or neither inside — fall through to heuristic

        # Heuristic: the value with |v| > 90 is the longitude
        if abs(v1) > 90 and _valid(*cand_a):
            return Point(*cand_a), False, f"{p0}, {p1}"
        if abs(v2) > 90 and _valid(*cand_b):
            return Point(*cand_b), True, f"{p1}, {p0}"

        # Both values in valid range — assume (lon, lat) as written
        if _valid(*cand_a):
            return Point(*cand_a), False, f"{p0}, {p1}"

        if _valid(*cand_b):
            return Point(*cand_b), True, f"{p1}, {p0}"

        return None, False, None

    except Exception:
        return None, False, None
