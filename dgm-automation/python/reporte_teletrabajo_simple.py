import pandas as pd
import os
from datetime import datetime
import unicodedata
import re

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
        for c in unicodedata.normalize("NFD", str(texto))
        if unicodedata.category(c) != "Mn"
    )


def normalizar_columnas(df):
    nuevas = []
    for col in df.columns:
        col = col.strip()
        col = quitar_tildes(col).upper()
        col = col.replace(" ", "_").replace("(", "").replace(")", "")
        nuevas.append(col)
    df.columns = nuevas
    return df


def extraer_expediente(texto):
    if pd.isna(texto):
        return ""
    patron = r"\d{4}-[A-Z]+-[A-Z]+-\d+"
    match = re.search(patron, str(texto).upper())
    return match.group(0) if match else ""


# ============================
# MAIN
# ============================


def main():

    print("=== REPORTE TELETRABAJO CON EXPEDIENTE ===\n")

    mes = int(input("Ingrese el MES (1-12): "))
    anio = int(input("Ingrese el AÑO (ejemplo 2026): "))

    usuario = (
        input(f"Nombre del Usuario [{USUARIO_DEFAULT}]: ").strip() or USUARIO_DEFAULT
    )
    usuario = usuario.upper()

    filas = []

    for archivo in os.listdir(CARPETA_BASE):

        if not archivo.lower().endswith(".xlsx"):
            continue

        if "reporte teletrabajo" in archivo.lower():
            continue

        ruta = os.path.join(CARPETA_BASE, archivo)

        try:
            df = pd.read_excel(ruta)
            df = normalizar_columnas(df)

            columnas_fecha = [c for c in df.columns if "FECHA" in c]
            if not columnas_fecha:
                continue

            col_fecha = columnas_fecha[0]
            df[col_fecha] = pd.to_datetime(
                df[col_fecha], errors="coerce", dayfirst=True
            )

            df = df[(df[col_fecha].dt.month == mes) & (df[col_fecha].dt.year == anio)]
            if df.empty:
                continue

            if "USUARIO" in df.columns:
                df["USUARIO"] = df["USUARIO"].astype(str).str.upper()
                df = df[df["USUARIO"] == usuario]
                if df.empty:
                    continue

            df = df.rename(columns={col_fecha: "FECHA"})

            if "ASUNTO" not in df.columns:
                df["ASUNTO"] = ""

            # 🔥 EXTRAER EXPEDIENTE DESDE ASUNTO
            df["EXPEDIENTE"] = df["ASUNTO"].apply(extraer_expediente)

            if "CONSECUTIVO" not in df.columns:
                if "CONSECTIVO" in df.columns:
                    df = df.rename(columns={"CONSECTIVO": "CONSECUTIVO"})
                else:
                    df["CONSECUTIVO"] = ""

            df_out = df[["FECHA", "ASUNTO", "EXPEDIENTE", "CONSECUTIVO"]].copy()

            filas.append(df_out)

        except Exception:
            continue

    if filas:
        df_final = pd.concat(filas, ignore_index=True).sort_values("FECHA")
    else:
        df_final = pd.DataFrame(
            columns=["FECHA", "ASUNTO", "EXPEDIENTE", "CONSECUTIVO"]
        )

    nombre_mes = datetime(anio, mes, 1).strftime("%B").upper()
    nombre_archivo = f"Reporte Teletrabajo_{nombre_mes}-{anio}_{usuario}.xlsx"
    ruta_salida = os.path.join(CARPETA_BASE, nombre_archivo)

    df_final.to_excel(ruta_salida, index=False)

    print("\nREPORTE GENERADO:")
    print(ruta_salida)


if __name__ == "__main__":
    main()
