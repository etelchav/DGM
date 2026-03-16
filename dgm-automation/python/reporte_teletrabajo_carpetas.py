import os
import re
from pathlib import Path
from datetime import datetime
import pandas as pd
from openpyxl.styles import Font

# ============================================================
# CONFIGURACION
# ============================================================
CARPETA_BASE = Path(r"C:\Users\echavarria\OneDrive - MINAE Costa Rica\DOCUMENTOS 2026")
CARPETA_SALIDA = Path(
    r"C:\Users\echavarria\OneDrive - MINAE Costa Rica\2-REPORTES\Reportes_Teletrabajo"
)

EXTENSIONES_VALIDAS = {
    ".pdf",
    ".doc",
    ".docx",
    ".xls",
    ".xlsx",
    ".ppt",
    ".pptx",
    ".txt",
    ".csv",
    ".msg",
}


# ============================================================
# FUNCIONES
# ============================================================
def limpiar_texto(texto: str) -> str:
    if texto is None:
        return ""
    texto = str(texto).strip()
    texto = re.sub(r"\s+", " ", texto)
    texto = re.sub(r"_+", "_", texto)
    return texto.strip(" _-")


def separar_nombre_archivo(nombre_archivo: str):
    """
    Formato esperado:
    1_2_3
    donde:
    1 = consecutivo
    2 = asunto
    3 = expediente

    Ejemplo:
    DGM-TOP-O-058-2026_Reserva de área_2026-CAN-PRI-008.pdf
    """
    stem = Path(nombre_archivo).stem
    partes = [limpiar_texto(p) for p in stem.split("_") if limpiar_texto(p)]

    consecutivo = ""
    asunto = ""
    expediente = ""

    if len(partes) >= 3:
        consecutivo = partes[0]
        expediente = partes[-1]
        asunto = " - ".join(partes[1:-1]) if len(partes) > 3 else partes[1]
    elif len(partes) == 2:
        consecutivo = partes[0]
        asunto = partes[1]
    elif len(partes) == 1:
        consecutivo = partes[0]

    return consecutivo, asunto, expediente


def obtener_fecha_modificacion(path_archivo: Path):
    try:
        ts = path_archivo.stat().st_mtime
        return datetime.fromtimestamp(ts)
    except Exception:
        return None


def nombre_mes(mes: int) -> str:
    meses = {
        1: "ENERO",
        2: "FEBRERO",
        3: "MARZO",
        4: "ABRIL",
        5: "MAYO",
        6: "JUNIO",
        7: "JULIO",
        8: "AGOSTO",
        9: "SETIEMBRE",
        10: "OCTUBRE",
        11: "NOVIEMBRE",
        12: "DICIEMBRE",
    }
    return meses.get(mes, f"MES_{mes}")


def recorrer_documentos_filtrados(carpeta_base: Path, anio: int, mes: int):
    registros = []

    if not carpeta_base.exists():
        raise FileNotFoundError(f"No existe la carpeta base: {carpeta_base}")

    for root, dirs, files in os.walk(carpeta_base):
        root_path = Path(root)

        for archivo in files:
            path_archivo = root_path / archivo

            if path_archivo.suffix.lower() not in EXTENSIONES_VALIDAS:
                continue

            fecha_mod = obtener_fecha_modificacion(path_archivo)
            if fecha_mod is None:
                continue

            if fecha_mod.year != anio or fecha_mod.month != mes:
                continue

            consecutivo, asunto, expediente = separar_nombre_archivo(path_archivo.name)

            try:
                carpeta_relativa = str(path_archivo.parent.relative_to(carpeta_base))
            except Exception:
                carpeta_relativa = str(path_archivo.parent)

            registros.append(
                {
                    "FECHA_MODIFICACION": fecha_mod,
                    "CARPETA": carpeta_relativa,
                    "ARCHIVO": path_archivo.name,
                    "CONSECUTIVO": consecutivo,
                    "ASUNTO": asunto,
                    "EXPEDIENTE": expediente,
                    "EXTENSION": path_archivo.suffix.lower(),
                    "ANIO": fecha_mod.year,
                    "MES": fecha_mod.month,
                    "RUTA_COMPLETA": str(path_archivo),
                }
            )

    return registros


def crear_excel(df: pd.DataFrame, ruta_salida: Path, anio: int, mes: int):
    mes_texto = nombre_mes(mes)
    titulo = f"Informe de Teletrabajo Etelberto Chavarría Camacho - {mes_texto} {anio}"

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        # Hoja principal iniciando en fila 3
        df.to_excel(writer, sheet_name="REPORTE", index=False, startrow=2)

        resumen_carpeta = (
            df.groupby("CARPETA", dropna=False)
            .size()
            .reset_index(name="CANTIDAD_DOCUMENTOS")
            .sort_values(["CANTIDAD_DOCUMENTOS", "CARPETA"], ascending=[False, True])
        )
        resumen_carpeta.to_excel(
            writer, sheet_name="RESUMEN_CARPETA", index=False, startrow=2
        )

        inconsistencias = df[
            (df["CONSECUTIVO"].astype(str).str.strip() == "")
            | (df["ASUNTO"].astype(str).str.strip() == "")
            | (df["EXPEDIENTE"].astype(str).str.strip() == "")
        ].copy()
        inconsistencias.to_excel(
            writer, sheet_name="REVISAR_NOMBRES", index=False, startrow=2
        )

        libro = writer.book

        # Aplicar título fila 1 y dejar fila 2 libre en todas las hojas
        for nombre_hoja in ["REPORTE", "RESUMEN_CARPETA", "REVISAR_NOMBRES"]:
            ws = libro[nombre_hoja]
            ws["A1"] = titulo
            ws["A1"].font = Font(size=15, bold=True)

        # Ajuste de ancho de columnas
        for hoja in libro.worksheets:
            for col in hoja.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    valor = "" if cell.value is None else str(cell.value)
                    if len(valor) > max_len:
                        max_len = len(valor)
                hoja.column_dimensions[col_letter].width = min(max_len + 2, 70)

        # Congelar encabezados de la tabla
        ws_reporte = libro["REPORTE"]
        ws_reporte.freeze_panes = "A4"

        ws_resumen = libro["RESUMEN_CARPETA"]
        ws_resumen.freeze_panes = "A4"

        ws_revisar = libro["REVISAR_NOMBRES"]
        ws_revisar.freeze_panes = "A4"


def main():
    print("========================================")
    print("REPORTE DE TELETRABAJO POR MES")
    print("========================================")
    print(f"Carpeta base: {CARPETA_BASE}")
    print(f"Carpeta salida: {CARPETA_SALIDA}")
    print()

    anio = int(input("Ingrese el año (ejemplo 2026): ").strip())
    mes = int(input("Ingrese el mes (1-12): ").strip())

    if mes < 1 or mes > 12:
        print("Mes inválido.")
        return

    CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)

    mes_texto = nombre_mes(mes)
    nombre_archivo = f"Informe_Teletrabajo_EtelbertoCHC_{mes_texto}_{anio}.xlsx"
    ruta_salida = CARPETA_SALIDA / nombre_archivo

    print()
    print("Leyendo documentos...")

    registros = recorrer_documentos_filtrados(CARPETA_BASE, anio, mes)

    if not registros:
        print(f"No se encontraron documentos para {mes_texto} de {anio}.")
        return

    df = pd.DataFrame(registros)

    columnas = [
        "FECHA_MODIFICACION",
        "CARPETA",
        "ARCHIVO",
        "CONSECUTIVO",
        "ASUNTO",
        "EXPEDIENTE",
        "EXTENSION",
        "ANIO",
        "MES",
        "RUTA_COMPLETA",
    ]
    df = df[columnas].sort_values(
        by=["FECHA_MODIFICACION", "CARPETA", "ARCHIVO"], ascending=[True, True, True]
    )

    crear_excel(df, ruta_salida, anio, mes)

    print()
    print("Reporte generado correctamente:")
    print(ruta_salida)
    print(f"Total de documentos incluidos: {len(df)}")


if __name__ == "__main__":
    main()
