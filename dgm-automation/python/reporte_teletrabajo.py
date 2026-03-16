import pandas as pd
import os
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
import sys
import unicodedata

# ============================
# CONFIGURACIÓN
# ============================

CARPETA_BASE = (
    r"C:\Users\echavarria\OneDrive - MINAE Costa Rica\2-REPORTES\Reportes_Teletrabajo"
)
USUARIO_DEFAULT = "ETELBERTO CHAVARRIA CAMACHO"

# ============================
# UTILIDADES
# ============================


def quitar_tildes(texto):
    return "".join(
        c
        for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )


def normalizar_columnas(df):
    nuevas = []
    for col in df.columns:
        col = col.strip()
        col = quitar_tildes(col)
        col = col.upper()
        col = col.replace(" ", "_")
        col = col.replace("(", "")
        col = col.replace(")", "")
        nuevas.append(col)
    df.columns = nuevas
    return df


# ============================
# PARAMETROS
# ============================


def pedir_parametros():
    print("=== GENERADOR REPORTE TELETRABAJO ===\n")

    mes = int(input("Ingrese el MES (1-12): "))
    anio = int(input("Ingrese el AÑO (ejemplo 2026): "))

    usuario = input(f"Nombre del Usuario [{USUARIO_DEFAULT}]: ").strip()
    if usuario == "":
        usuario = USUARIO_DEFAULT

    print("\nSeleccione el archivo Excel del Informe de Correos...")
    root = tk.Tk()
    root.withdraw()
    archivo_correos = filedialog.askopenfilename(
        title="Seleccione Informe de Correos",
        filetypes=[("Excel files", "*.xlsx *.xls")],
    )

    if not archivo_correos:
        print("No se seleccionó archivo de correos.")
        sys.exit()

    return mes, anio, usuario.upper(), archivo_correos


# ============================
# LECTURA LABORES
# ============================


def filtrar_por_fecha(df, mes, anio, columna_fecha):
    df[columna_fecha] = pd.to_datetime(
        df[columna_fecha], errors="coerce", dayfirst=True
    )
    return df[(df[columna_fecha].dt.month == mes) & (df[columna_fecha].dt.year == anio)]


def leer_archivos_carpeta(mes, anio, usuario):

    data_total = []

    for archivo in os.listdir(CARPETA_BASE):

        if archivo.endswith(".xlsx") and "REPORTE TELETRABAJO" not in archivo.upper():

            ruta = os.path.join(CARPETA_BASE, archivo)

            try:
                df = pd.read_excel(ruta)
                df = normalizar_columnas(df)

                posibles_fechas = [c for c in df.columns if "FECHA" in c]
                if not posibles_fechas:
                    continue

                columna_fecha = posibles_fechas[0]

                df = filtrar_por_fecha(df, mes, anio, columna_fecha)

                if "USUARIO" in df.columns:
                    df = df[df["USUARIO"].str.upper() == usuario]

                df = df.rename(columns={columna_fecha: "FECHA"})

                data_total.append(df)

            except Exception as e:
                print(f"Error leyendo {archivo}: {e}")

    if not data_total:
        return pd.DataFrame()

    return pd.concat(data_total, ignore_index=True)


# ============================
# LECTURA CORREOS
# ============================


def procesar_correos(ruta_correos, mes, anio):

    df = pd.read_excel(ruta_correos)
    df = normalizar_columnas(df)

    if "FECHA_DE_ENVIO" not in df.columns:
        print("No se encontró columna FECHA_DE_ENVIO.")
        print("Columnas detectadas:", df.columns.tolist())
        sys.exit()

    df = filtrar_por_fecha(df, mes, anio, "FECHA_DE_ENVIO")

    df = df.rename(
        columns={
            "FECHA_DE_ENVIO": "FECHA",
            "DESTINATARIOS": "DESTINATARIOS",
            "DESTINATARIOS": "DESTINATARIOS",
        }
    )

    if "DESTINATARIOS" not in df.columns:
        df["DESTINATARIOS"] = ""

    if "ASUNTO" not in df.columns:
        df["ASUNTO"] = ""

    return df[["FECHA", "ASUNTO", "DESTINATARIOS"]]


# ============================
# CONSOLIDACIÓN
# ============================


def consolidar(df_labores, df_correos):

    columnas_base = ["FECHA", "ASUNTO", "EXPEDIENTE", "CONSECUTIVO"]

    for col in columnas_base:
        if col not in df_labores.columns:
            df_labores[col] = ""

    df_labores = df_labores[columnas_base].copy()

    if df_correos.empty:
        df_labores["DESTINATARIOS"] = ""
        return df_labores.sort_values("FECHA")

    df_final = pd.merge(df_labores, df_correos, on=["FECHA", "ASUNTO"], how="left")

    correos_no_rel = df_correos.merge(
        df_labores, on=["FECHA", "ASUNTO"], how="left", indicator=True
    )

    correos_no_rel = correos_no_rel[correos_no_rel["_merge"] == "left_only"]

    if not correos_no_rel.empty:
        extra = correos_no_rel.copy()
        extra["EXPEDIENTE"] = ""
        extra["CONSECUTIVO"] = ""
        df_final = pd.concat([df_final, extra], ignore_index=True)

    return df_final.sort_values("FECHA")


# ============================
# GUARDADO
# ============================


def guardar_reporte(df, mes, anio, usuario):

    nombre_mes = datetime(anio, mes, 1).strftime("%B").upper()
    nombre_archivo = f"Reporte Teletrabajo_{nombre_mes}-{anio}_{usuario}.xlsx"
    ruta_salida = os.path.join(CARPETA_BASE, nombre_archivo)

    df.to_excel(ruta_salida, index=False)

    print("\nREPORTE GENERADO:")
    print(ruta_salida)


# ============================
# MAIN
# ============================

if __name__ == "__main__":

    mes, anio, usuario, archivo_correos = pedir_parametros()

    print("\nProcesando archivos de labores...")
    df_labores = leer_archivos_carpeta(mes, anio, usuario)

    print("Procesando archivo de correos...")
    df_correos = procesar_correos(archivo_correos, mes, anio)

    print("Consolidando información...")
    df_final = consolidar(df_labores, df_correos)

    guardar_reporte(df_final, mes, anio, usuario)
