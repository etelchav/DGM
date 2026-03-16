import arcpy
import os
from datetime import datetime

arcpy.env.overwriteOutput = True

# ==========================================
# PARAMETROS
# ==========================================

expediente_input = arcpy.GetParameterAsText(0).strip()
cm_base = arcpy.GetParameterAsText(1)

if not expediente_input:
    arcpy.AddError("Debe ingresar un número de expediente.")
    raise Exception("Expediente vacío.")

if not cm_base:
    arcpy.AddError("Debe indicar la capa CM_BASE.")
    raise Exception("Capa CM_BASE vacía.")

# ==========================================
# FUNCIONES AUXILIARES
# ==========================================


def normalizar(txt):
    if txt is None:
        return ""
    return str(txt).strip().upper()


def buscar_campo(capa, candidatos):
    """
    Busca un campo real dentro de la capa, tolerando:
    - nombres simples: Expediente, ESTADO
    - nombres calificados: CM_BASE.Expediente
    """
    fields = arcpy.ListFields(capa)
    mapa = {}

    for f in fields:
        mapa[f.name.upper()] = f.name
        base = f.name.split(".")[-1].upper()
        mapa[base] = f.name

    for c in candidatos:
        c_up = c.upper()
        if c_up in mapa:
            return mapa[c_up]

    return None


# ==========================================
# CREAR LAYER
# ==========================================

arcpy.MakeFeatureLayer_management(cm_base, "lyr_base")

campo_expediente = buscar_campo("lyr_base", ["Expediente", "EXPEDIEN_1", "ID"])
campo_estado = buscar_campo("lyr_base", ["ESTADO", "Estado"])

if not campo_expediente:
    arcpy.AddError("No se encontró el campo de expediente en CM_BASE.")
    raise Exception("Falta campo de expediente.")

if not campo_estado:
    arcpy.AddError("No se encontró el campo de estado en CM_BASE.")
    raise Exception("Falta campo de estado.")

arcpy.AddMessage(f"Campo expediente detectado: {campo_expediente}")
arcpy.AddMessage(f"Campo estado detectado: {campo_estado}")

geom_base = None
area_base = 0

# ==========================================
# BUSCAR GEOMETRIA BASE
# ==========================================

with arcpy.da.SearchCursor("lyr_base", ["SHAPE@", campo_expediente]) as cursor:

    for row in cursor:
        exp_actual = str(row[1]).strip() if row[1] is not None else ""
        if exp_actual == expediente_input:
            geom_base = row[0]
            area_base = geom_base.area
            break

if geom_base is None:
    arcpy.AddError(f"Expediente no encontrado: {expediente_input}")
    raise Exception("Expediente no encontrado.")

if area_base == 0:
    arcpy.AddError("El expediente tiene área 0.")
    raise Exception("Área inválida.")

# ==========================================
# ANALISIS DE TRASLAPES
# ==========================================

resultados = []

with arcpy.da.SearchCursor(
    "lyr_base", ["SHAPE@", campo_expediente, campo_estado]
) as cursor:

    for row in cursor:
        geom_otro = row[0]
        expediente_otro = str(row[1]).strip() if row[1] is not None else ""
        estado_otro = str(row[2]).strip() if row[2] is not None else ""

        # Saltar el mismo expediente
        if expediente_otro == expediente_input:
            continue

        # Control de estado desde CM_BASE
        estado_upper = normalizar(estado_otro)

        if estado_upper == "":
            estado_otro = "SIN_ESTADO"
        elif estado_upper in ["ARCHIVADO", "EXTINTO", "NO_UBICADO"]:
            continue

        # Validar geometría
        if geom_otro is None:
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

        resultados.append(
            (
                expediente_otro,
                estado_otro if estado_otro else "SIN_ESTADO",
                round(area_inter, 2),
                round(porcentaje, 2),
                tipo,
            )
        )

# Ordenar por porcentaje descendente
resultados.sort(key=lambda x: x[3], reverse=True)

# ==========================================
# REPORTE
# ==========================================

ruta_reporte = (
    r"C:\Users\echavarria\OneDrive - MINAE Costa Rica\2-REPORTES\Reportes_GIS\Traslapes"
)

if not os.path.exists(ruta_reporte):
    os.makedirs(ruta_reporte)

fecha = datetime.now().strftime("%Y%m%d_%H%M%S")

archivo_txt = os.path.join(
    ruta_reporte, f"Reporte_Traslapes_{expediente_input}_{fecha}.txt"
)

with open(archivo_txt, "w", encoding="utf-8") as f:

    f.write("=========================================\n")
    f.write("REPORTE DE TRASLAPES - CM_BASE\n")
    f.write("=========================================\n\n")
    f.write(f"Expediente base: {expediente_input}\n")
    f.write(f"Campo expediente usado: {campo_expediente}\n")
    f.write(f"Campo estado usado: {campo_estado}\n")
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
