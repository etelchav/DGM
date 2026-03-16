# -*- coding: utf-8 -*-
import pandas as pd
import re
import unicodedata

# ==========================
# CONFIGURACIÓN
# ==========================
ARCHIVO_ENTRADA = r"C:\CatastroMineroC\BaseDatosC\CAMPOS Y CARGA MASIVA 28-2-2026B.xlsx"
ARCHIVO_SALIDA = r"C:\CatastroMineroC\BaseDatosC\CAMPOS_Y_CARGA_MASIVA_CLASIFICADO.xlsx"


# ==========================
# FUNCION LIMPIAR TEXTO
# ==========================
def limpiar(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).strip()
    texto = unicodedata.normalize("NFKD", texto)
    texto = texto.encode("ascii", "ignore").decode("utf-8")
    return texto.upper()


# ==========================
# CARGA
# ==========================
df = pd.read_excel(ARCHIVO_ENTRADA)

# 🔹 LIMPIAR ENCABEZADOS
df.columns = [limpiar(c) for c in df.columns]

# 🔹 LIMPIAR CONTENIDO
for col in df.columns:
    if df[col].dtype == "object":
        df[col] = df[col].apply(limpiar)

# ==========================
# VALIDAR COLUMNAS
# ==========================
columnas_requeridas = [
    "ESTADO",
    "MATERIALES",
    "REGIMEN_JURIDICO",
    "SUBCLASE",
    "FINALIDAD",
    "EXPEDIENTE",
]

for col in columnas_requeridas:
    if col not in df.columns:
        print(f"❌ Falta columna requerida: {col}")
        print("Columnas encontradas:", df.columns.tolist())
        raise SystemExit

# ==========================
# CLASIFICACIÓN
# ==========================
df["GRUPO_CATASTRO"] = "NO_METALICOS"

# 1) ARCHIVADOS
cond_arch = df["ESTADO"] == "ARCHIVADO"
df.loc[cond_arch, "GRUPO_CATASTRO"] = "ARCHIVADOS"

# 2) METALICOS
cond_metal = df["MATERIALES"].str.contains(r"(ORO|PLATA|COBRE|HIERRO)", na=False)
df.loc[~cond_arch & cond_metal, "GRUPO_CATASTRO"] = "METALICOS"

# 3) PERMISO_ESPECIAL
cond_permiso = (
    (df["REGIMEN_JURIDICO"] == "PERMISO_ESPECIAL")
    | ((df["SUBCLASE"] != "PRIVADO") & (df["SUBCLASE"] != ""))
    | (df["FINALIDAD"] == "EMERGENCIA")
    | (df["EXPEDIENTE"].str.contains(r"(MUN|CNE)", na=False))
)

df.loc[~cond_arch & ~cond_metal & cond_permiso, "GRUPO_CATASTRO"] = "PERMISO_ESPECIAL"

# ==========================
# EXPORTAR
# ==========================
resumen = df["GRUPO_CATASTRO"].value_counts()

with pd.ExcelWriter(ARCHIVO_SALIDA, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="DATOS")
    resumen.to_excel(writer, sheet_name="RESUMEN")

print("\n✔ Clasificación completada correctamente.")
print(resumen)
from datetime import datetime
import os

# ==========================
# GENERAR NOMBRE DINÁMICO
# ==========================
fecha_hora = datetime.now().strftime("%Y%m%d_%H%M%S")
nombre_archivo = f"AGRUPADO_{fecha_hora}.xlsx"

carpeta_salida = r"C:\CatastroMineroC\BaseDatosC"
ruta_final = os.path.join(carpeta_salida, nombre_archivo)

# ==========================
# EXPORTAR
# ==========================
resumen = df["GRUPO_CATASTRO"].value_counts()

with pd.ExcelWriter(ruta_final, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="DATOS")
    resumen.to_excel(writer, sheet_name="RESUMEN")

print("\n✔ Archivo generado correctamente:")
print(ruta_final)
