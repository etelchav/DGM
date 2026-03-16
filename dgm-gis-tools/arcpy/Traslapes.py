import arcpy
import os
from datetime import datetime

arcpy.env.overwriteOutput = True

# ==========================================
# PARAMETROS
# ==========================================

expediente_input = arcpy.GetParameterAsText(0)
cm_base = arcpy.GetParameterAsText(1)

if not expediente_input:
    arcpy.AddError("Debe ingresar un número de expediente.")
    raise Exception("Expediente vacío.")

# ==========================================
# CREAR LAYER COMPLETA
# ==========================================

arcpy.MakeFeatureLayer_management(cm_base, "lyr_base")

geom_base = None
area_base = 0

# ==========================================
# BUSCAR GEOMETRIA BASE
# ==========================================

with arcpy.da.SearchCursor(
    "lyr_base",
    ["SHAPE@", "CM_BASE.Expediente"]
) as cursor:

    for row in cursor:
        if row[1] == expediente_input:
            geom_base = row[0]
            area_base = geom_base.area
            break

if geom_base is None:
    arcpy.AddError("Expediente no encontrado.")
    raise Exception("Expediente no encontrado.")

if area_base == 0:
    arcpy.AddError("El expediente tiene área 0.")
    raise Exception("Área inválida.")

# ==========================================
# ANALISIS
# ==========================================

resultados = []

with arcpy.da.SearchCursor(
    "lyr_base",
    ["SHAPE@", "CM_BASE.Expediente", "CM_ATRIBUTOS.Estado"]
) as cursor:

    for row in cursor:

        geom_otro = row[0]
        expediente_otro = row[1]
        estado_otro = row[2]

        # Saltar el mismo expediente
        if expediente_otro == expediente_input:
            continue

        # CONTROL DE ESTADO
        if estado_otro is None or str(estado_otro).strip() == "":
            estado_otro = "SIN_ESTADO"
        else:
            estado_upper = estado_otro.upper()
            if estado_upper in ["ARCHIVADO", "EXTINTO"]:
                continue

        # Si no se tocan
        if geom_base.disjoint(geom_otro):
            continue

        inter = geom_base.intersect(geom_otro, 4)
        area_inter = inter.area

        if area_inter > 0:
            tipo = "TRASLAPA"
            porcentaje = (area_inter / area_base) * 100
        else:
            tipo = "TOCA"
            porcentaje = 0

        resultados.append((
            expediente_otro,
            estado_otro,
            round(area_inter, 2),
            round(porcentaje, 2),
            tipo
        ))

# Ordenar por porcentaje descendente
resultados.sort(key=lambda x: x[3], reverse=True)

# ==========================================
# REPORTE
# ==========================================

ruta_reporte = r"C:\Users\echavarria\OneDrive - MINAE Costa Rica\2-REPORTES\Reportes_GIS\Traslapes"

if not os.path.exists(ruta_reporte):
    os.makedirs(ruta_reporte)

fecha = datetime.now().strftime("%Y%m%d_%H%M%S")

archivo_txt = os.path.join(
    ruta_reporte,
    f"Reporte_Traslapes_{expediente_input}_{fecha}.txt"
)

with open(archivo_txt, "w", encoding="utf-8") as f:

    f.write("=========================================\n")
    f.write("REPORTE DE TRASLAPES - CM_BASE\n")
    f.write("=========================================\n\n")
    f.write(f"Expediente base: {expediente_input}\n")
    f.write(f"Área base: {round(area_base, 2)} m²\n")
    f.write(f"Fecha: {datetime.now()}\n\n")

    if not resultados:
        f.write("No se detectaron traslapes válidos.\n")
    else:
        for r in resultados:
            f.write(
                f"Expediente: {r[0]} | "
                f"Estado: {r[1]} | "
                f"Tipo: {r[4]} | "
                f"Área traslapada: {r[2]} m² | "
                f"% respecto al base: {r[3]} %\n"
            )

# ==========================================
# MENSAJES EN PRO
# ==========================================

if not resultados:
    arcpy.AddMessage("No se detectaron traslapes válidos.")
else:
    arcpy.AddMessage("Traslapes detectados:")
    for r in resultados:
        arcpy.AddMessage(f"{r[0]} | {r[1]} | {r[3]} %")

arcpy.AddMessage(f"Reporte generado en: {archivo_txt}")
