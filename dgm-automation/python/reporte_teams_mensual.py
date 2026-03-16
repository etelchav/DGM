import win32com.client
import pandas as pd
from datetime import datetime
import os

print("================================")
print("REPORTE MENSUAL DE REUNIONES")
print("Cuenta: echavarria@minae.go.cr")
print("================================")

# ----------------------------
# PEDIR PERIODO
# ----------------------------

anio = int(input("Ingrese el año (ej: 2026): "))
mes = int(input("Ingrese el mes (1-12): "))

inicio = datetime(anio, mes, 1)

if mes == 12:
    fin = datetime(anio + 1, 1, 1)
else:
    fin = datetime(anio, mes + 1, 1)

# ----------------------------
# CONECTAR A OUTLOOK
# ----------------------------

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

calendar = namespace.Folders["echavarria@minae.go.cr"].Folders["Calendario"]

# ----------------------------
# FILTRAR POR RANGO (forma correcta)
# ----------------------------

items = calendar.Items
items.Sort("[Start]")

inicio_str = inicio.strftime("%m/%d/%Y %H:%M")
fin_str = fin.strftime("%m/%d/%Y %H:%M")

filtro = f"[Start] >= '{inicio_str}' AND [Start] < '{fin_str}'"

items_filtrados = items.Restrict(filtro)
items_filtrados.Sort("[Start]")

print("\nRevisando reuniones en el rango...\n")

data = []

data = []

for meeting in items_filtrados:
    try:
        if meeting.Class != 26:
            continue

        # Solo reuniones (no citas personales)
        if meeting.MeetingStatus == 0:
            continue

        # Debe tener participantes
        if meeting.Recipients.Count == 0:
            continue

        # Opcional: excluir canceladas
        if meeting.Subject.lower().startswith("cancelled"):
            continue

        hora_inicio = meeting.Start.time()
        hora_fin = meeting.End.time()
        duracion = meeting.Duration / 60

        participantes = [att.Name for att in meeting.Recipients]

        data.append(
            [
                meeting.Start.date(),
                hora_inicio,
                hora_fin,
                round(duracion, 2),
                meeting.Subject,
                meeting.Organizer,
                ", ".join(participantes),
            ]
        )

    except:
        continue

# ----------------------------
# VALIDAR RESULTADO
# ----------------------------

if not data:
    print("⚠ No se encontraron reuniones en ese mes.")
    exit()

df = pd.DataFrame(
    data,
    columns=[
        "FECHA",
        "HORA_INICIO",
        "HORA_FIN",
        "DURACION_HORAS",
        "ASUNTO",
        "ORGANIZADOR",
        "PARTICIPANTES",
    ],
)

total_horas = round(df["DURACION_HORAS"].sum(), 2)
total_reuniones = len(df)

# ----------------------------
# EXPORTAR
# ----------------------------

ruta_destino = r"C:\Users\echavarria\OneDrive - MINAE Costa Rica\2-REPORTES\Reportes_Teletrabajo\Teams"
os.makedirs(ruta_destino, exist_ok=True)

nombre_archivo = f"Reporte_Reuniones_{anio}_{mes:02d}.xlsx"
ruta_excel = os.path.join(ruta_destino, nombre_archivo)

with pd.ExcelWriter(ruta_excel, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Detalle", index=False)

    resumen = pd.DataFrame(
        {"TOTAL_REUNIONES": [total_reuniones], "TOTAL_HORAS": [total_horas]}
    )

    resumen.to_excel(writer, sheet_name="Resumen", index=False)

print("\n================================")
print("Reporte generado correctamente")
print("Total reuniones:", total_reuniones)
print("Total horas:", total_horas)
print("Archivo:", ruta_excel)
print("================================")
