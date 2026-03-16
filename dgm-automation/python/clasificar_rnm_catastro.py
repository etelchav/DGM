import pandas as pd
import datetime
import os

# ==========================================
# RUTAS
# ==========================================
ruta_entrada = r"C:\CatastroMineroC\BaseDatosC\BD_RNM_AGOL.xlsx"
ruta_salida = r"C:\CatastroMineroC\BaseDatosC\BD_RNM_CLASIFICADO_20260227_1919.xlsx"

col_grupo = "GRUPO_CATASTRO"

# ==========================================
# CARGAR ARCHIVO
# ==========================================
df = pd.read_excel(ruta_entrada)

# Normalizar columnas críticas
for col in ["ESTADO", "CLASE", "MATERIALES", "EXPEDIENTE"]:
    if col in df.columns:
        df[col] = df[col].astype(str).str.upper().str.strip()

# Crear columna grupo
df[col_grupo] = ""

# ==========================================
# PASO 1 — ARCHIVADOS
# ==========================================
mask_arch = df["ESTADO"] == "ARCHIVADO"
df.loc[mask_arch, col_grupo] = "ARCHIVADOS"

# ==========================================
# PASO 2 — METALICOS
# ==========================================
mask_libres = df[col_grupo] == ""

mask_exp = df["EXPEDIENTE"].str.contains("EXP", na=False)
mask_sub = df["EXPEDIENTE"].str.contains("SUB", na=False)
mask_clase = df["CLASE"].isin(["EXPLORACIÓN", "SUBTERRÁNEO"])
mask_mat = df["MATERIALES"].str.contains("ORO|PLATA", na=False)

mask_metal = mask_libres & (mask_exp | mask_sub | mask_clase | mask_mat)
df.loc[mask_metal, col_grupo] = "METALICOS"

# ==========================================
# PASO 3 — PERMISOS_ESPECIALES (ACTUALIZADO)
# ==========================================
mask_libres = df[col_grupo] == ""

anio_actual = datetime.datetime.now().year

# Extraer año del expediente
df["ANIO_EXP"] = df["EXPEDIENTE"].str[:4]
df["ANIO_EXP"] = pd.to_numeric(df["ANIO_EXP"], errors="coerce")

df["ANTIGUEDAD"] = anio_actual - df["ANIO_EXP"]

# Regla 1 — CNE directo
mask_cne = df["EXPEDIENTE"].str.contains("CNE", na=False)

# Regla 2 — CAN-MUN directo
mask_can_mun = df["EXPEDIENTE"].str.contains("CAN-MUN", na=False)

# Regla 3 — MUN con antigüedad > 4
mask_mun = df["EXPEDIENTE"].str.contains("MUN", na=False)
mask_mun_antiguo = mask_mun & (df["ANTIGUEDAD"] > 4)

mask_perm = mask_libres & (mask_cne | mask_can_mun | mask_mun_antiguo)

df.loc[mask_perm, col_grupo] = "PERMISOS_ESPECIALES"

# Regla 1 — CNE directo
mask_cne = df["EXPEDIENTE"].str.contains("CNE", na=False)

# Regla 2 — MUN con antigüedad > 4
mask_mun = df["EXPEDIENTE"].str.contains("MUN", na=False)
mask_mun_antiguo = mask_mun & (df["ANTIGUEDAD"] > 4)

mask_perm = mask_libres & (mask_cne | mask_mun_antiguo)

df.loc[mask_perm, col_grupo] = "PERMISOS_ESPECIALES"

# ==========================================
# PASO 4 — NO_METALICOS
# ==========================================
mask_libres = df[col_grupo] == ""
df.loc[mask_libres, col_grupo] = "NO_METALICOS"

# ==========================================
# VALIDACIÓN
# ==========================================
print("\nResumen por grupo:")
print(df[col_grupo].value_counts())

print("\nRegistros sin grupo:", len(df[df[col_grupo] == ""]))

# ==========================================
# EXPORTAR
# ==========================================
if os.path.exists(ruta_salida):
    os.remove(ruta_salida)

df.to_excel(ruta_salida, index=False)

print("\nArchivo generado correctamente en:")
print(ruta_salida)
