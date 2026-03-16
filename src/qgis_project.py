"""Generate a QGIS project file (.qgs) with pre-configured layers."""

from __future__ import annotations

import uuid
from pathlib import Path
from typing import Optional

from .config import Config


# ---------------------------------------------------------------------------
# Colour presets per layer
# ---------------------------------------------------------------------------

_LAYER_STYLES = {
    "origins": {
        "color": "0,114,189,255",       # blue
        "size": "2.6",
        "shape": "circle",
    },
    "destinations": {
        "color": "217,83,25,255",       # orange-red
        "size": "2.6",
        "shape": "circle",
    },
    "od_lines": {
        "color": "128,128,128,180",     # semi-transparent grey
        "width": "0.26",
    },
    "boundaries": {
        "fill_color": "0,0,0,0",        # transparent fill
        "outline_color": "50,50,50,255",
        "outline_width": "0.46",
    },
}


def _uid() -> str:
    return "{" + str(uuid.uuid4()).upper() + "}"


def _point_layer_xml(
    layer_id: str,
    name: str,
    gpkg_path: str,
    layer_name: str,
    color: str,
    size: str,
) -> str:
    return f"""
    <maplayer type="vector" geometry="Point" autoRefreshEnabled="0">
      <id>{layer_id}</id>
      <datasource>{gpkg_path}|layername={layer_name}</datasource>
      <keywordList><value/></keywordList>
      <layername>{name}</layername>
      <srs><spatialrefsys><authid>EPSG:4326</authid></spatialrefsys></srs>
      <renderer-v2 type="singleSymbol">
        <symbols>
          <symbol type="marker" name="0" alpha="1" clip_to_extent="1" force_rhr="0">
            <data_defined_properties><Option type="Map"><Option type="QString" value="" name="name"/><Option name="properties"/><Option type="QString" value="collection" name="type"/></Option></data_defined_properties>
            <layer pass="0" class="SimpleMarker" enabled="1" locked="0">
              <Option type="Map">
                <Option type="QString" value="{color}" name="color"/>
                <Option type="QString" value="{size}" name="size"/>
                <Option type="QString" value="circle" name="name"/>
              </Option>
            </layer>
          </symbol>
        </symbols>
      </renderer-v2>
    </maplayer>"""


def _line_layer_xml(
    layer_id: str,
    name: str,
    gpkg_path: str,
    layer_name: str,
    color: str,
    width: str,
) -> str:
    return f"""
    <maplayer type="vector" geometry="Line" autoRefreshEnabled="0">
      <id>{layer_id}</id>
      <datasource>{gpkg_path}|layername={layer_name}</datasource>
      <keywordList><value/></keywordList>
      <layername>{name}</layername>
      <srs><spatialrefsys><authid>EPSG:4326</authid></spatialrefsys></srs>
      <renderer-v2 type="singleSymbol">
        <symbols>
          <symbol type="line" name="0" alpha="1" clip_to_extent="1" force_rhr="0">
            <data_defined_properties><Option type="Map"><Option type="QString" value="" name="name"/><Option name="properties"/><Option type="QString" value="collection" name="type"/></Option></data_defined_properties>
            <layer pass="0" class="SimpleLine" enabled="1" locked="0">
              <Option type="Map">
                <Option type="QString" value="{color}" name="line_color"/>
                <Option type="QString" value="{width}" name="line_width"/>
              </Option>
            </layer>
          </symbol>
        </symbols>
      </renderer-v2>
    </maplayer>"""


def _polygon_layer_xml(
    layer_id: str,
    name: str,
    gpkg_path: str,
    layer_name: str,
    fill_color: str,
    outline_color: str,
    outline_width: str,
) -> str:
    return f"""
    <maplayer type="vector" geometry="Polygon" autoRefreshEnabled="0">
      <id>{layer_id}</id>
      <datasource>{gpkg_path}|layername={layer_name}</datasource>
      <keywordList><value/></keywordList>
      <layername>{name}</layername>
      <srs><spatialrefsys><authid>EPSG:4326</authid></spatialrefsys></srs>
      <renderer-v2 type="singleSymbol">
        <symbols>
          <symbol type="fill" name="0" alpha="1" clip_to_extent="1" force_rhr="0">
            <data_defined_properties><Option type="Map"><Option type="QString" value="" name="name"/><Option name="properties"/><Option type="QString" value="collection" name="type"/></Option></data_defined_properties>
            <layer pass="0" class="SimpleFill" enabled="1" locked="0">
              <Option type="Map">
                <Option type="QString" value="{fill_color}" name="color"/>
                <Option type="QString" value="{outline_color}" name="outline_color"/>
                <Option type="QString" value="{outline_width}" name="outline_width"/>
              </Option>
            </layer>
          </symbol>
        </symbols>
      </renderer-v2>
    </maplayer>"""


def write_qgis_project(
    cfg: Config,
    gpkg_path: Optional[Path],
    run_tag: str,
) -> Path:
    """Generate a .qgs XML project file that references layers in *gpkg_path*.

    If *gpkg_path* is None (GIS export was skipped), the project is still
    created but with placeholder datasources.
    """
    out_path = cfg.output_dir / f"{run_tag}_project.qgs"

    # Use relative path from project file to gpkg (same directory)
    if gpkg_path is not None:
        gpkg_rel = Path("./") / gpkg_path.name
        gpkg_str = str(gpkg_rel).replace("\\", "/")
    else:
        gpkg_str = "./transport.gpkg"

    layer_names = cfg.gis_export.layers
    lyr_origins = layer_names.get("origins", "origins")
    lyr_dests = layer_names.get("destinations", "destinations")
    lyr_od = layer_names.get("od_lines", "od_lines")
    lyr_bounds = layer_names.get("boundaries", "boundaries")

    id_origins = _uid()
    id_dests = _uid()
    id_od = _uid()
    id_bounds = _uid()

    s_o = _LAYER_STYLES["origins"]
    s_d = _LAYER_STYLES["destinations"]
    s_l = _LAYER_STYLES["od_lines"]
    s_b = _LAYER_STYLES["boundaries"]

    origins_xml = _point_layer_xml(id_origins, "Отправления", gpkg_str, lyr_origins,
                                   s_o["color"], s_o["size"])
    dests_xml = _point_layer_xml(id_dests, "Прибытия", gpkg_str, lyr_dests,
                                 s_d["color"], s_d["size"])
    od_xml = _line_layer_xml(id_od, "OD-линии", gpkg_str, lyr_od,
                             s_l["color"], s_l["width"])
    bounds_xml = _polygon_layer_xml(id_bounds, "Границы территорий", gpkg_str, lyr_bounds,
                                    s_b["fill_color"], s_b["outline_color"], s_b["outline_width"])

    # Layer order in legend (bottom to top): boundaries → od_lines → origins → destinations
    legend_order = (
        f'<legendlayer name="Границы территорий" showFeatureCount="0" open="1" checked="Qt::Checked" drawingOrder="-1">'
        f'<filegroup open="1" hidden="0"><legendlayerfile isInOverview="0" layerid="{id_bounds}" visible="1"/></filegroup></legendlayer>'
        f'<legendlayer name="OD-линии" showFeatureCount="0" open="1" checked="Qt::Checked" drawingOrder="-1">'
        f'<filegroup open="1" hidden="0"><legendlayerfile isInOverview="0" layerid="{id_od}" visible="1"/></filegroup></legendlayer>'
        f'<legendlayer name="Отправления" showFeatureCount="0" open="1" checked="Qt::Checked" drawingOrder="-1">'
        f'<filegroup open="1" hidden="0"><legendlayerfile isInOverview="0" layerid="{id_origins}" visible="1"/></filegroup></legendlayer>'
        f'<legendlayer name="Прибытия" showFeatureCount="0" open="1" checked="Qt::Checked" drawingOrder="-1">'
        f'<filegroup open="1" hidden="0"><legendlayerfile isInOverview="0" layerid="{id_dests}" visible="1"/></filegroup></legendlayer>'
    )

    qgs_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE qgis PUBLIC 'http://mrcc.com/qgis.dtd' 'SYSTEM'>
<qgis version="3.28" projectname="{cfg.qgis_project.title}">
  <title>{cfg.qgis_project.title}</title>
  <autotransaction active="0"/>
  <evaluateDefaultValues active="0"/>
  <trust active="0"/>
  <projectlayers>
    {bounds_xml}
    {od_xml}
    {origins_xml}
    {dests_xml}
  </projectlayers>
  <layerorder>
    <layer id="{id_bounds}"/>
    <layer id="{id_od}"/>
    <layer id="{id_origins}"/>
    <layer id="{id_dests}"/>
  </layerorder>
  <properties/>
  <legend updateDrawingOrder="true">
    {legend_order}
  </legend>
  <mapcanvas annotationsVisible="1" name="theMapCanvas">
    <units>degrees</units>
    <extent>
      <xmin>98</xmin><ymin>46</ymin><xmax>120</xmax><ymax>58</ymax>
    </extent>
    <rotation>0</rotation>
    <destinationsrs><spatialrefsys><authid>EPSG:4326</authid></spatialrefsys></destinationsrs>
  </mapcanvas>
  <projectlayerorder/>
</qgis>
"""

    out_path.write_text(qgs_xml, encoding="utf-8")
    print(f"  QGIS-проект сохранён: {out_path.name}")
    return out_path
