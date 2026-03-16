# -*- coding: utf-8 -*-
import arcpy
import os
from datetime import datetime

arcpy.env.overwriteOutput = True

# =========================================================
# PARÁMETROS DE ENTRADA (HERRAMIENTA)
# =========================================================
POLIGONO_ENTRADA = arcpy.GetParameterAsText(0)
EXPEDIENTE_NUEVO = arcpy.GetParameterAsText(1)

# =========================================================
# VALIDACIONES DE EXPEDIENTE
# =========================================================
if not EXPEDIENTE_NUEVO:
    raise Exception("El expediente es obligatorio.")

if " " in EXPEDIENTE_NUEVO:
    raise Exception("El expediente NO puede contener espacios.")

if EXPEDIENTE_NUEVO != EXPEDIENTE_NUEVO.upper():
    raise Exception("El expediente debe estar en MAYÚSCULAS.")

# =========================================================
# RUTAS FIJAS
# =========================================================
GDB = r"C:\Users\echavarria\Documents\ArcGIS\Projects\Catastro Minero AGOL\Catastro Minero AGOL.gdb"
CM_BASE = os.path.join(GDB, "CM_BASE")
CM_ATRIBUTOS = os.path.join(GDB, "CM_ATRIBUTOS")

REPORTES_TRASLAPES = (
    r"C:\Users\echavarria\OneDrive - MINAE Costa Rica"
    r"\2-REPORTES\Reportes_GIS\Traslapes"
)
os.makedirs(REPORTES_TRASLAPES, exist_ok=True)

TS = datetime.now().strftime("%Y%m%d_%H%M%S")
REPORTE_XLSX = os.path.join(
    REPORTES_TRASLAPES, f"TRASLAPES_{EXPEDIENTE_NUEVO}_{TS}.xlsx"
)

# =========================================================
# VALIDACIONES DE EXISTENCIA
# =========================================================
if not arcpy.Exists(POLIGONO_ENTRADA):
    raise Exception("El polígono de entrada no existe.")

if not arcpy.Exists(CM_BASE):
    raise Exception("No existe CM_BASE.")

if not arcpy.Exists(CM_ATRIBUTOS):
    raise Exception("No existe CM_ATRIBUTOS.")

# =========================================================
# CARGAR ESTADOS DESDE CM_ATRIBUTOS (DICCIONARIO)
# =========================================================
estado_por_expediente = {}

# Se intenta usar Estado o ESTADO de forma tolerante
estado_field = None
for f in arcpy.ListFields(CM_ATRIBUTOS):
    if f.name.upper() == "ESTADO":
        estado_field = f.name
        break

if not estado_field:
    raise Exception("CM_ATRIBUTOS no tiene un campo Estado/ESTADO.")

with arcpy.da.SearchCursor(CM_ATRIBUTOS, ["EXPEDIENTE", estado_field]) as cur:
    for exp, est in cur:
        if exp:
            estado_por_expediente[exp] = est

# =========================================================
# INSERTAR NUEVO POLÍGONO EN CM_BASE
# =========================================================
arcpy.AddMessage("Insertando nuevo polígono en CM_BASE...")

geom_nueva = None

with arcpy.da.SearchCursor(POLIGONO_ENTRADA, ["SHAPE@"]) as scur, arcpy.da.InsertCursor(
    CM_BASE, ["SHAPE@", "EXPEDIENTE"]
) as icur:

    for (geom,) in scur:
        geom_nueva = geom
        icur.insertRow((geom, EXPEDIENTE_NUEVO))
        break

if not geom_nueva:
    raise Exception("No se pudo obtener la geometría de entrada.")

# =========================================================
# ZOOM AUTOMÁTICO (ARCGIS PRO)
# =========================================================
aprx = arcpy.mp.ArcGISProject("CURRENT")
view = aprx.activeView

if view and hasattr(view, "camera"):
    ext = geom_nueva.extent

    # expandir manualmente 20% alrededor
    dx = (ext.XMax - ext.XMin) * 0.10
    dy = (ext.YMax - ext.YMin) * 0.10

    ext2 = arcpy.Extent(ext.XMin - dx, ext.YMin - dy, ext.XMax + dx, ext.YMax + dy)

    cam = view.camera
    cam.setExtent(ext2)
    view.camera = cam

# =========================================================
# ANÁLISIS DE TRASLAPES
# =========================================================
arcpy.AddMessage("Analizando traslapes contra CM_BASE...")

area_nueva = geom_nueva.area
traslapes = []

with arcpy.da.SearchCursor(CM_BASE, ["SHAPE@", "EXPEDIENTE"]) as cur:
    for geom, expediente in cur:

        # Saltar el polígono recién insertado
        if expediente == EXPEDIENTE_NUEVO:
            continue

        # Obtener estado desde CM_ATRIBUTOS (diccionario)
        estado = estado_por_expediente.get(expediente, "NO_DEFINIDO")

        # Ignorar expedientes archivados o extintos
        if estado in ("ARCHIVADO", "EXTINTO"):
            continue

        if geom:
            # Intersección real (solo traslape con área)
            inter = geom.intersect(geom_nueva, 4)  # 4 = POLYGON
            if not inter or inter.area == 0:
                continue

            porcentaje = round((inter.area / area_nueva) * 100, 2)

            traslapes.append(
                {
                    "EXPEDIENTE": expediente,
                    "ESTADO": estado,
                    "AREA_TRASLAPE": round(inter.area, 2),
                    "PORCENTAJE_TRASLAPE": porcentaje,
                }
            )

# =========================================================
# REPORTE DE TRASLAPES (EXCEL)
# =========================================================
if traslapes:
    arcpy.AddWarning("⚠️ Se detectaron traslapes con otros expedientes.")

    import pandas as pd

    df = pd.DataFrame(traslapes)
    df.insert(0, "EXPEDIENTE_NUEVO", EXPEDIENTE_NUEVO)

    with pd.ExcelWriter(REPORTE_XLSX, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="TRASLAPES", index=False)

    arcpy.AddWarning(f"Reporte generado en:\n{REPORTE_XLSX}")

else:
    arcpy.AddMessage("No se detectaron traslapes.")

arcpy.AddMessage("Proceso finalizado correctamente.")
