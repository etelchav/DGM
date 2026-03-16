import pandas as pd
import os
from datetime import datetime

# =====================================================
# CONFIGURACIÓN
# =====================================================

carpeta_origen = (
    r"C:\Users\echavarria\OneDrive - MINAE Costa Rica\2-REPORTES\Reportes_Teletrabajo"
)
carpeta_destino = os.path.join(carpeta_origen, "Consolidado")
os.makedirs(carpeta_destino, exist_ok=True)

usuario_objetivo = "ETELBERTO CHAVARRIA CAMACHO"
anio = "2026"

# =====================================================
# BUSCAR ARCHIVOS
# =====================================================

archivos = [
    f
    for f in os.listdir(carpeta_origen)
    if f.endswith(".xlsx") and not f.startswith("Informe_Teletrabajo_Etelberto")
]

if not archivos:
    print("No se encontraron archivos para consolidar.")
    exit()

# =====================================================
# CONSOLIDACIÓN
# =====================================================

lista_dfs = []

for archivo in archivos:
    ruta = os.path.join(carpeta_origen, archivo)
    df = pd.read_excel(ruta)
    lista_dfs.append(df)

df_consolidado = pd.concat(lista_dfs, ignore_index=True)
df_consolidado = df_consolidado.dropna(how="all")

# =====================================================
# NORMALIZAR COLUMNAS
# =====================================================

df_consolidado.columns = df_consolidado.columns.str.strip().str.upper()

# =====================================================
# FILTRAR POR USUARIO
# =====================================================

if "USUARIO" not in df_consolidado.columns:
    print("La columna USUARIO no existe en los archivos.")
    exit()

df_consolidado = df_consolidado[
    df_consolidado["USUARIO"].astype(str).str.upper().str.strip() == usuario_objetivo
]

# =====================================================
# ELIMINAR COLUMNAS NO DESEADAS
# =====================================================

if "CONSECUTIVO" in df_consolidado.columns:
    df_consolidado = df_consolidado.drop(columns=["CONSECUTIVO"])

# =====================================================
# MANEJO DE FECHA
# =====================================================

mes_nombre = "Sin datos"  # ← SIEMPRE definida

if "FECHA" in df_consolidado.columns:

    df_consolidado["FECHA"] = df_consolidado["FECHA"].astype(str).str.strip()

    df_consolidado["FECHA"] = pd.to_datetime(
        df_consolidado["FECHA"], errors="coerce", dayfirst=True
    )

    fechas_validas = df_consolidado["FECHA"].dropna()

    if not fechas_validas.empty:
        mes_num = int(fechas_validas.dt.month.iloc[0])
        mes_nombre = datetime(1900, mes_num, 1).strftime("%B").capitalize()

    df_consolidado = df_consolidado.sort_values(by="FECHA")

# =====================================================
# CREAR TITULO
# =====================================================

titulo = f"Informe de Teletrabajo de Ing. Etelberto Chavarría Camacho del mes {mes_nombre} de {anio}"

# =====================================================
# EXPORTAR ARCHIVO
# =====================================================

timestamp = datetime.now().strftime("%Y%d%m%H%M%S")
nombre_final = f"Informe_Teletrabajo_Etelberto_{timestamp}.xlsx"
ruta_final = os.path.join(carpeta_destino, nombre_final)

with pd.ExcelWriter(ruta_final, engine="openpyxl") as writer:
    df_consolidado.to_excel(writer, index=False, startrow=2)
    hoja = writer.sheets["Sheet1"]
    hoja["A1"] = titulo

print("====================================")
print("Informe generado correctamente")
print("Total registros:", len(df_consolidado))
print("Archivo generado en:")
print(ruta_final)
