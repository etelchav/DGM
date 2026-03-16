# -*- coding: utf-8 -*-
import arcpy
import os
import zipfile
from datetime import datetime
import traceback

arcpy.env.overwriteOutput = True


def mensaje(txt):
    arcpy.AddMessage(txt)
    print(txt)


def advertencia(txt):
    arcpy.AddWarning(txt)
    print(txt)


def error(txt):
    arcpy.AddError(txt)
    print(txt)


def crear_zip_shapefile(carpeta_origen, nombre_base, ruta_zip):
    extensiones_validas = [
        ".shp",
        ".shx",
        ".dbf",
        ".prj",
        ".cpg",
        ".sbn",
        ".sbx",
        ".fbn",
        ".fbx",
        ".ain",
        ".aih",
        ".atx",
        ".ixs",
        ".mxs",
    ]

    with zipfile.ZipFile(ruta_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for archivo in os.listdir(carpeta_origen):
            base, ext = os.path.splitext(archivo)
            if (
                base.lower() == nombre_base.lower()
                and ext.lower() in extensiones_validas
            ):
                ruta_archivo = os.path.join(carpeta_origen, archivo)
                zf.write(ruta_archivo, arcname=archivo)


def eliminar_componentes_shapefile(carpeta, nombre_base):
    for archivo in os.listdir(carpeta):
        base, ext = os.path.splitext(archivo)
        if base.lower() == nombre_base.lower() and ext.lower() != ".zip":
            try:
                os.remove(os.path.join(carpeta, archivo))
            except Exception as e:
                advertencia(f"No se pudo eliminar {archivo}: {e}")


def exportar_grupo(capa_entrada, campo_grupo, valor_grupo, carpeta_trabajo, timestamp):
    nombre_base = f"CM_2026_{valor_grupo}_{timestamp}"
    nombre_capa_temp = f"lyr_{valor_grupo}"

    ruta_shp = os.path.join(carpeta_trabajo, f"{nombre_base}.shp")
    ruta_zip = os.path.join(carpeta_trabajo, f"{nombre_base}.zip")

    mensaje(f"Procesando grupo: {valor_grupo}")

    if arcpy.Exists(nombre_capa_temp):
        arcpy.management.Delete(nombre_capa_temp)

    arcpy.management.MakeFeatureLayer(capa_entrada, nombre_capa_temp)

    campo_delim = arcpy.AddFieldDelimiters(capa_entrada, campo_grupo)
    campo_estado = arcpy.AddFieldDelimiters(capa_entrada, "ESTADO")

    where = (
        f"{campo_delim} = '{valor_grupo}' "
        f"AND {campo_estado} NOT IN ('NO_UBICADO', 'ARCHIVADO')"
    )

    arcpy.management.SelectLayerByAttribute(nombre_capa_temp, "NEW_SELECTION", where)

    conteo = int(arcpy.management.GetCount(nombre_capa_temp)[0])
    mensaje(f"Registros seleccionados para {valor_grupo}: {conteo}")

    if conteo == 0:
        advertencia(
            f"No hay registros para el grupo {valor_grupo}. No se generará shapefile."
        )
        arcpy.management.Delete(nombre_capa_temp)
        return None

    arcpy.management.CopyFeatures(nombre_capa_temp, ruta_shp)
    mensaje(f"Shapefile exportado: {ruta_shp}")

    crear_zip_shapefile(carpeta_trabajo, nombre_base, ruta_zip)
    mensaje(f"ZIP creado: {ruta_zip}")

    eliminar_componentes_shapefile(carpeta_trabajo, nombre_base)

    arcpy.management.Delete(nombre_capa_temp)
    return ruta_zip


def main():
    try:
        capa_entrada = arcpy.GetParameterAsText(0)

        if not capa_entrada:
            raise ValueError("Debe indicar la capa de entrada.")

        carpeta_base_fija = r"C:\Users\echavarria\OneDrive - MINAE Costa Rica\2-REPORTES\Respaldos_CM_2026"
        campo_grupo = "GRUPO_CATASTRO"
        grupos = ["METALICOS", "NO_METALICOS", "PERMISO_ESPECIAL"]

        if not os.path.exists(carpeta_base_fija):
            os.makedirs(carpeta_base_fija)

        timestamp = datetime.now().strftime("%d%m%y_%H%M%S")
        carpeta_final = os.path.join(carpeta_base_fija, f"EXP_CM_{timestamp}")
        os.makedirs(carpeta_final, exist_ok=True)

        mensaje("========================================")
        mensaje("EXPORTACIÓN SEGMENTADA CM_2026")
        mensaje("========================================")
        mensaje(f"Capa entrada: {capa_entrada}")
        mensaje(f"Carpeta final automática: {carpeta_final}")

        campos = [f.name.upper() for f in arcpy.ListFields(capa_entrada)]
        if campo_grupo.upper() not in campos:
            raise ValueError(f"No existe el campo requerido: {campo_grupo}")
        if "ESTADO" not in campos:
            raise ValueError("No existe el campo requerido: ESTADO")

        zips_generados = []

        for grupo in grupos:
            ruta_zip = exportar_grupo(
                capa_entrada=capa_entrada,
                campo_grupo=campo_grupo,
                valor_grupo=grupo,
                carpeta_trabajo=carpeta_final,
                timestamp=timestamp,
            )
            if ruta_zip:
                zips_generados.append(ruta_zip)

        mensaje("========================================")
        mensaje("PROCESO FINALIZADO")
        mensaje("========================================")
        mensaje(f"Total ZIP generados: {len(zips_generados)}")

        for z in zips_generados:
            mensaje(f" - {z}")

    except Exception as e:
        error("Ocurrió un error en la herramienta.")
        error(str(e))
        error(traceback.format_exc())
        raise


if __name__ == "__main__":
    main()
