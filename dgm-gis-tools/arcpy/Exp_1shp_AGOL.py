import arcpy
import os
import zipfile
import re
import unicodedata
from datetime import datetime


def normalizar_nombre_campo(nombre, usados):
    """
    Convierte a minúsculas, sin tildes, sin caracteres especiales.
    Limita a 10 caracteres por compatibilidad shapefile.
    Evita duplicados.
    """
    nombre = unicodedata.normalize("NFKD", nombre)
    nombre = "".join(c for c in nombre if not unicodedata.combining(c))
    nombre = nombre.lower()
    nombre = re.sub(r"[^a-z0-9_]", "_", nombre)
    nombre = re.sub(r"_+", "_", nombre).strip("_")

    if not nombre:
        nombre = "campo"

    if nombre[0].isdigit():
        nombre = f"f_{nombre}"

    base = nombre[:10]

    nuevo = base
    i = 1
    while nuevo in usados:
        sufijo = str(i)
        nuevo = base[: 10 - len(sufijo)] + sufijo
        i += 1

    usados.add(nuevo)
    return nuevo


def main():
    arcpy.env.overwriteOutput = True

    # =========================
    # PARÁMETRO DE ENTRADA
    # =========================
    capa_entrada = arcpy.GetParameterAsText(0)

    if not capa_entrada:
        raise ValueError("Debe indicar una capa de entrada.")

    if not arcpy.Exists(capa_entrada):
        raise ValueError(f"No existe la capa de entrada: {capa_entrada}")

    # =========================
    # CARPETA FIJA DE SALIDA
    # =========================
    output_folder = r"C:\Users\echavarria\OneDrive - MINAE Costa Rica\2-REPORTES\Respaldos_CM_2026\1Shape"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # =========================
    # NOMBRE DE SALIDA
    # =========================
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = f"cm_base_{timestamp}"
    output_shapefile = os.path.join(output_folder, f"{base_name}.shp")
    zip_filename = os.path.join(output_folder, f"{base_name}.zip")

    # =========================
    # CAMPOS EN ORDEN DESEADO
    # AJUSTADOS A LA CAPA REAL
    # =========================
    columnas_deseadas = [
        "GlobalID",
        "Expediente",
        "EXPEDIENTE_KEY",
        "ESTE",
        "NORTE",
        "NOMBRE",
        "ID",
        "ESTADO",
        "FECHA_INGRESO",
        "PROYECTO",
        "CLASE",
        "SUBCLASE",
        "MATERIALES",
        "PROVINCIA",
        "CANTON",
        "DISTRITO",
        "REGION",
        "AREA",
        "UNIDAD",
        "CAUCE",
        "LONGITUD_RIO",
        "CATASTRO",
        "FINCA",
        "CONTACTO_COMERCIAL",
        "TELEFONO_COMERCIAL",
        "WHATSAPP_COMERCIAL",
        "CORREO_COMERCIAL",
        "WEB_COMERCIAL",
        "METODO_EXPLOTACION",
        "REGIMEN_JURIDICO",
        "FINALIDAD",
        "GRUPO_CATASTRO",
        "Fecha_Actualizacion",
        "EN_REPORTE",
        "FECHA_REPORTE",
        "Shape_Length",
        "Shape_Area",
    ]

    # =========================
    # LEER CAMPOS REALES
    # =========================
    fields = arcpy.ListFields(capa_entrada)
    campos_reales = [f.name for f in fields]
    campos_upper = {f.name.upper(): f.name for f in fields}

    arcpy.AddMessage("Campos reales detectados en la capa:")
    arcpy.AddMessage(", ".join(campos_reales))

    # =========================
    # VALIDAR CAMPO ESTADO
    # =========================
    campo_estado = campos_upper.get("ESTADO")
    if not campo_estado:
        raise ValueError(
            "La capa no contiene el campo ESTADO. No se puede aplicar el filtro."
        )

    # =========================
    # FILTRO VÁLIDO PARA GDB
    # EXCLUIR ARCHIVADO Y NO_UBICADO
    # =========================
    campo_estado_delim = arcpy.AddFieldDelimiters(capa_entrada, campo_estado)
    where_clause = (
        f"{campo_estado_delim} IS NULL OR "
        f"UPPER({campo_estado_delim}) <> 'ARCHIVADO' AND "
        f"UPPER({campo_estado_delim}) <> 'NO_UBICADO'"
    )

    arcpy.AddMessage(f"Filtro aplicado: {where_clause}")

    capa_filtrada = "cm_base_filtrada_tmp"
    arcpy.management.MakeFeatureLayer(capa_entrada, capa_filtrada, where_clause)

    conteo = int(arcpy.management.GetCount(capa_filtrada)[0])
    if conteo == 0:
        raise ValueError(
            "No hay registros para exportar luego de excluir ARCHIVADO y NO_UBICADO."
        )

    arcpy.AddMessage(f"✅ Registros a exportar: {conteo}")

    # =========================
    # CAMPOS EXISTENTES
    # =========================
    columnas_existentes = []
    columnas_faltantes = []

    for col in columnas_deseadas:
        col_real = campos_upper.get(col.upper())
        if col_real:
            columnas_existentes.append(col_real)
        else:
            columnas_faltantes.append(col)

    if not columnas_existentes:
        raise ValueError(
            "No se encontró ninguno de los campos esperados en la capa de entrada."
        )

    arcpy.AddMessage(f"✅ Campos a exportar: {', '.join(columnas_existentes)}")

    if columnas_faltantes:
        arcpy.AddWarning(
            f"⚠ Campos no encontrados y omitidos: {', '.join(columnas_faltantes)}"
        )

    # =========================
    # MAPEO DE CAMPOS + RENOMBRE
    # =========================
    field_mappings = arcpy.FieldMappings()
    usados = set()
    log_mapeo = []

    for field_name in columnas_existentes:
        fm = arcpy.FieldMap()
        fm.addInputField(capa_filtrada, field_name)

        out_field = fm.outputField
        nombre_nuevo = normalizar_nombre_campo(field_name, usados)
        out_field.name = nombre_nuevo
        out_field.aliasName = nombre_nuevo
        fm.outputField = out_field

        field_mappings.addFieldMap(fm)
        log_mapeo.append((field_name, nombre_nuevo))

    # =========================
    # EXPORTAR SHAPEFILE
    # =========================
    arcpy.conversion.FeatureClassToFeatureClass(
        in_features=capa_filtrada,
        out_path=output_folder,
        out_name=base_name,
        field_mapping=field_mappings,
    )

    if not os.path.exists(output_shapefile):
        raise RuntimeError("El shapefile no se generó correctamente.")

    arcpy.AddMessage(f"📂 Shapefile exportado: {output_shapefile}")

    arcpy.AddMessage("Mapeo final de campos:")
    for origen, destino in log_mapeo:
        arcpy.AddMessage(f" - {origen} -> {destino}")

    # =========================
    # ZIP SIN XML
    # =========================
    with zipfile.ZipFile(zip_filename, "w", zipfile.ZIP_DEFLATED) as zipf:
        for ext in [".shp", ".shx", ".dbf", ".prj", ".cpg"]:
            ruta = os.path.join(output_folder, f"{base_name}{ext}")
            if os.path.exists(ruta):
                zipf.write(ruta, os.path.basename(ruta))

    arcpy.AddMessage(f"📦 ZIP generado: {zip_filename}")

    # =========================
    # LIMPIEZA
    # =========================
    arcpy.management.Delete(capa_filtrada)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        arcpy.AddError(f"❌ Error durante la exportación: {str(e)}")
        raise
