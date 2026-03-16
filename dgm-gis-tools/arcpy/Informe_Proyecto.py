# -*- coding: utf-8 -*-

import arcpy
import os
import pandas as pd
from datetime import datetime

arcpy.env.overwriteOutput = True

# =====================================================
# CONFIGURACION GENERAL
# =====================================================
project = arcpy.mp.ArcGISProject("CURRENT")
gdb = project.defaultGeodatabase

if not gdb or not arcpy.Exists(gdb):
    raise Exception("No se detecto una Geodatabase por defecto valida.")

OUTPUT_DIR = r"C:\Users\echavarria\OneDrive - MINAE Costa Rica\2-REPORTES\Reportes_GIS\Rep_Auto_ContenidoProyecto"
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

output_xlsx = os.path.join(OUTPUT_DIR, "INVENTARIO_PROYECTO_GIS_" + timestamp + ".xlsx")

# =====================================================
# CONTENEDORES
# =====================================================
maps_rows = []
layers_rows = []
fc_rows = []
tbl_rows = []
toolbox_rows = []
script_rows = []
notebook_rows = []
log_rows = []

# =====================================================
# MAPAS Y CAPAS
# =====================================================
for m in project.listMaps():
    maps_rows.append({"MAPA": m.name})

    for lyr in m.listLayers():
        layers_rows.append(
            {
                "MAPA": m.name,
                "CAPA": lyr.name,
                "TIPO": lyr.longName,
                "VISIBLE": lyr.visible,
            }
        )

# =====================================================
# FEATURE CLASSES Y TABLAS
# =====================================================
arcpy.env.workspace = gdb

for fc in arcpy.ListFeatureClasses():
    d = arcpy.Describe(fc)
    fc_rows.append(
        {
            "NOMBRE": fc,
            "TIPO": "FeatureClass",
            "GEOMETRIA": d.shapeType,
            "RUTA": os.path.join(gdb, fc),
        }
    )

for tbl in arcpy.ListTables():
    tbl_rows.append({"NOMBRE": tbl, "TIPO": "Table", "RUTA": os.path.join(gdb, tbl)})

# =====================================================
# TOOLBOXES Y SCRIPTS
# =====================================================
for tbx in project.toolboxes:
    if not isinstance(tbx, str):
        continue

    tbx_name = os.path.basename(tbx)

    toolbox_rows.append({"TOOLBOX": tbx_name, "RUTA": tbx})

    try:
        arcpy.ImportToolbox(tbx)
        tools = arcpy.ListTools("*")
        if tools:
            for t in tools:
                try:
                    d = arcpy.Describe(t)
                    script_file = getattr(d, "scriptFile", None)
                    if script_file and isinstance(script_file, str):
                        script_rows.append(
                            {
                                "HERRAMIENTA": t,
                                "SCRIPT_PY": os.path.basename(script_file),
                                "RUTA_SCRIPT": script_file,
                                "TOOLBOX": tbx_name,
                            }
                        )
                except:
                    pass
    except:
        pass

# =====================================================
# NOTEBOOKS
# =====================================================
proj_dir = os.path.dirname(project.filePath)

for root, dirs, files in os.walk(proj_dir):
    for f in files:
        if f.lower().endswith(".ipynb"):
            notebook_rows.append({"NOTEBOOK": f, "RUTA": os.path.join(root, f)})

# =====================================================
# LOGS AUTOMATICOS (UPSERT / AUDITORIA)
# =====================================================
LOG_DIR = r"C:\CatastroMineroC\BaseDatosC"

if os.path.exists(LOG_DIR):
    for f in os.listdir(LOG_DIR):
        if f.lower().startswith("log_") and f.lower().endswith(".csv"):
            full_path = os.path.join(LOG_DIR, f)
            log_rows.append(
                {
                    "ARCHIVO_LOG": f,
                    "RUTA": full_path,
                    "FECHA_MODIFICACION": datetime.fromtimestamp(
                        os.path.getmtime(full_path)
                    ).strftime("%Y-%m-%d %H:%M:%S"),
                }
            )

# =====================================================
# EXPORTAR A EXCEL
# =====================================================
writer = pd.ExcelWriter(output_xlsx, engine="openpyxl")

if maps_rows:
    pd.DataFrame(maps_rows).to_excel(writer, "MAPAS", index=False)

if layers_rows:
    pd.DataFrame(layers_rows).to_excel(writer, "CAPAS_POR_MAPA", index=False)

if fc_rows:
    pd.DataFrame(fc_rows).to_excel(writer, "FEATURE_CLASSES", index=False)

if tbl_rows:
    pd.DataFrame(tbl_rows).to_excel(writer, "TABLAS", index=False)

if toolbox_rows:
    pd.DataFrame(toolbox_rows).to_excel(writer, "TOOLBOXES", index=False)

if script_rows:
    pd.DataFrame(script_rows).to_excel(writer, "SCRIPTS_PY", index=False)

if notebook_rows:
    pd.DataFrame(notebook_rows).to_excel(writer, "NOTEBOOKS", index=False)

if log_rows:
    pd.DataFrame(log_rows).to_excel(writer, "LOGS_AUTOMATICOS", index=False)

writer.close()

# =====================================================
# MENSAJES
# =====================================================
arcpy.AddMessage("INVENTARIO EXTENDIDO DEL PROYECTO GENERADO")
arcpy.AddMessage("Archivo:")
arcpy.AddMessage(output_xlsx)
arcpy.AddMessage("Mapas: " + str(len(maps_rows)))
arcpy.AddMessage("Capas: " + str(len(layers_rows)))
arcpy.AddMessage("Feature Classes: " + str(len(fc_rows)))
arcpy.AddMessage("Tablas: " + str(len(tbl_rows)))
arcpy.AddMessage("Toolboxes: " + str(len(toolbox_rows)))
arcpy.AddMessage("Scripts Python: " + str(len(script_rows)))
arcpy.AddMessage("Notebooks: " + str(len(notebook_rows)))
arcpy.AddMessage("Logs detectados: " + str(len(log_rows)))
