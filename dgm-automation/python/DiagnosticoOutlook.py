import win32com.client
from datetime import datetime

print("======================================")
print("INICIANDO DIAGNÓSTICO OUTLOOK")
print("======================================")

# =====================================================
# CONEXIÓN
# =====================================================

try:
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    print("✔ Conectado a Outlook correctamente")
except Exception as e:
    print("✖ Error conectando a Outlook:", e)
    exit()

print("\n======================================")
print("BUZONES DISPONIBLES")
print("======================================")

for i in range(1, namespace.Folders.Count + 1):
    print("-", namespace.Folders.Item(i).Name)

# =====================================================
# RECORRER TODOS LOS BUZONES Y CALENDARIOS
# =====================================================

print("\n======================================")
print("ANALIZANDO CALENDARIOS")
print("======================================")

for i in range(1, namespace.Folders.Count + 1):

    buzon = namespace.Folders.Item(i)
    print(f"\nBUZÓN: {buzon.Name}")

    for j in range(1, buzon.Folders.Count + 1):
        carpeta = buzon.Folders.Item(j)

        # Tipo 9 = Calendario
        if carpeta.DefaultItemType == 1:  # 1 = olAppointmentItem

            print("  → Calendario detectado:", carpeta.Name)

            try:
                items = carpeta.Items
                items.Sort("[Start]")
                items.IncludeRecurrences = True

                print("     Total elementos:", items.Count)

                # Buscar eventos en febrero 2026
                inicio = datetime(2026, 2, 1)
                fin = datetime(2026, 3, 1)

                encontrados = 0

                for meeting in items:
                    try:
                        if meeting.Class != 26:
                            continue

                        fecha = meeting.Start

                        if fecha < inicio:
                            continue

                        if fecha >= fin:
                            break

                        print("       Fecha:", fecha)
                        print("       Asunto:", meeting.Subject)
                        print("       Organizer:", meeting.Organizer)
                        print(
                            "       IsOnlineMeeting:",
                            getattr(meeting, "IsOnlineMeeting", "N/A"),
                        )
                        print(
                            "       OnlineMeetingProvider:",
                            getattr(meeting, "OnlineMeetingProvider", "N/A"),
                        )
                        print("       Location:", meeting.Location)
                        print("       --------------------------------")

                        encontrados += 1

                        if encontrados >= 20:
                            break

                    except:
                        continue

                if encontrados == 0:
                    print("       ⚠ No se encontraron eventos en febrero 2026")

            except Exception as e:
                print("     Error leyendo calendario:", e)

print("\n======================================")
print("FIN DEL DIAGNÓSTICO")
print("======================================")
