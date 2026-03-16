import pandas as pd
import os
from datetime import datetime

print("===================================")
print("MERGE DE EXCEL POR EXPEDIENTE")
print("VERSION AUDITABLE")
print("===================================")

# Pedir rutas
archivo_A = input("Ruta Excel A (ORIGEN): ").strip('"')
archivo_B = input("Ruta Excel B (A ACTUALIZAR): ").strip('"')

# Carpeta de salida
carpeta_salida = r"C:\Users\echavarria\OneDrive - MINAE Costa Rica\2-REPORTES\Python"

# Leer archivos
df_A = pd.read_excel(archivo_A)
df_B = pd.read_excel(archivo_B)

# Normalizar encabezados
df_A.columns = df_A.columns.str.upper().str.strip()
df_B.columns = df_B.columns.str.upper().str.strip()

# Verificar columna clave
if "EXPEDIENTE" not in df_A.columns or "EXPEDIENTE" not in df_B.columns:
    print("ERROR: ambos archivos deben tener columna EXPEDIENTE")
    exit()

# Normalizar campo expediente
df_A["EXPEDIENTE"] = df_A["EXPEDIENTE"].astype(str).str.strip()
df_B["EXPEDIENTE"] = df_B["EXPEDIENTE"].astype(str).str.strip()


# Función para comparar valores correctamente
def normalizar(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip().upper()


# Detectar columnas comunes
columnas_comunes = list(set(df_A.columns) & set(df_B.columns))

if "EXPEDIENTE" in columnas_comunes:
    columnas_comunes.remove("EXPEDIENTE")

print("\nColumnas evaluadas:")
for c in columnas_comunes:
    print("-", c)

# Crear índice rápido
df_A_index = df_A.set_index("EXPEDIENTE")

# Crear columna de validación
df_B["VALIDACION_CAMBIO"] = ""

# Lista para reporte de cambios
reporte_cambios = []

actualizados = 0
sin_cambio = 0
no_en_origen = 0

# Recorrer registros
for i, row in df_B.iterrows():

    expediente = row["EXPEDIENTE"]

    if expediente in df_A_index.index:

        cambio = False

        for col in columnas_comunes:

            valor_A = df_A_index.loc[expediente, col]
            valor_B = row[col]

            if normalizar(valor_A) != normalizar(valor_B):

                df_B.at[i, col] = valor_A
                cambio = True

                reporte_cambios.append(
                    {
                        "EXPEDIENTE": expediente,
                        "CAMPO": col,
                        "VALOR_ANTERIOR": valor_B,
                        "VALOR_NUEVO": valor_A,
                    }
                )

        if cambio:
            df_B.at[i, "VALIDACION_CAMBIO"] = "ACTUALIZADO"
            actualizados += 1
        else:
            df_B.at[i, "VALIDACION_CAMBIO"] = "SIN_CAMBIO"
            sin_cambio += 1

    else:

        df_B.at[i, "VALIDACION_CAMBIO"] = "NO_EN_ORIGEN"
        no_en_origen += 1

# Convertir reporte a DataFrame
df_cambios = pd.DataFrame(reporte_cambios)

# Crear nombre archivo con fecha
nombre_A = os.path.splitext(os.path.basename(archivo_A))[0]
nombre_B = os.path.splitext(os.path.basename(archivo_B))[0]

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

archivo_salida = f"MERGE_{nombre_A}_CON_{nombre_B}_{timestamp}.xlsx"

ruta_salida = os.path.join(carpeta_salida, archivo_salida)

# Guardar Excel
with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:

    df_B.to_excel(writer, sheet_name="DATOS_ACTUALIZADOS", index=False)

    if not df_cambios.empty:
        df_cambios.to_excel(writer, sheet_name="REPORTE_CAMBIOS", index=False)

print("\n===================================")
print("RESULTADO DEL PROCESO")
print("===================================")

print("EXPEDIENTES ACTUALIZADOS:", actualizados)
print("EXPEDIENTES SIN CAMBIO:", sin_cambio)
print("EXPEDIENTES NO EN ORIGEN:", no_en_origen)

print("\nArchivo generado en:")
print(ruta_salida)

print("\nProceso terminado correctamente.")
