README – MACROS EXCEL DGM (ADDAX / HERMES)
========================================

Proyecto: Normalización, Auditoría y Preparación ETL
Institución: Dirección de Geología y Minas – MINAE
Responsable: Etelberto Chavarría
Año: 2026


1. OBJETIVO GENERAL
------------------
Este conjunto de macros en Excel permite:

- Normalizar datos provenientes de ADDAX y HERMES
- Preparar tablas para su consumo por procesos ETL (ArcGIS / AGOL)
- Auditar la calidad de los datos (duplicados, estados vacíos)
- Generar reportes automáticos y trazables para control institucional

Las macros están diseñadas para ejecutarse sobre la hoja activa
y posteriormente ser empaquetadas como complemento XLAM institucional.


2. ESTRUCTURA DE ARCHIVOS
------------------------

Carpeta: Dev\Excel

- modADDAX_Auditoria.bas
  Auditoría post-normalización de datos ADDAX.

- modADDAX_PreETL.bas
  Preparación de tabla ADDAX para procesos ETL ArcGIS.

- Normalizar_HERMES_Estado_Concesiona.bas
  Normalización de datos provenientes del sistema HERMES.

- Normalizador_Work.xlsm
  Libro de pruebas y validación de macros.

- README_Macros_Excel_DGM.txt
  Documento descriptivo del conjunto de macros.


3. DESCRIPCIÓN DE MACROS
-----------------------

3.1 ADDAX – Normalización de datos
Macro: ADDAX_Normalizar_Concesiona_Proyecto_Estado
Módulo: modNormalizar_Textos_Addax

Acciones:
- Normaliza a MAYÚSCULAS SIN TILDES los campos:
  CONCESIONA, PROYECTO, PROVINCIA, CANTON, DISTRITO
- Normaliza los valores del campo ESTADO
- Fuerza columnas clave como TEXTO
- Reordena la columna ESTADO junto a EXPEDIENTE

Resultado:
Tabla homogénea y consistente, lista para auditoría.


3.2 ADDAX – Auditoría post-normalización
Macro: ADDAX_Auditar_Tabla_Post_Normalizacion
Módulo: modADDAX_Auditoria

Acciones:
- Detecta duplicados por EXPEDIENTE
- Conserva el registro con mayor cantidad de información
- Elimina registros duplicados
- Corrige ESTADO vacío a “PENDIENTE”
- Genera hoja AUDITORIA_ADDAX
- Exporta reporte automático a:

  C:\Users\echavarria\OneDrive - MINAE Costa Rica\
  2-REPORTES\Reportes_Macros_Excel\

- Muestra un resumen ejecutivo en pantalla (MsgBox)

Resultado:
Control de calidad completo y trazable del proceso.


3.3 ADDAX – Preparación para ETL
Macro: ADDAX_Preparar_Tabla_ETL
Módulo: modADDAX_PreETL

Acciones:
- Ajusta tipos de datos
- Asegura campos clave como TEXTO
- Deja la estructura compatible con ArcGIS Pro y ArcGIS Online

Resultado:
Tabla lista para ser consumida por procesos ETL GIS.


3.4 HERMES – Normalización
Macro: Normalizar_BD_HERMES
Módulo: modNormalizar_BD_HERMES

Acciones:
- Normaliza estados administrativos
- Limpia textos
- Aplica reglas específicas del sistema HERMES

Nota:
Este flujo es independiente del flujo ADDAX.


4. ORDEN RECOMENDADO DE EJECUCIÓN (ADDAX)
----------------------------------------

1) ADDAX_Normalizar_Concesiona_Proyecto_Estado
2) ADDAX_Auditar_Tabla_Post_Normalizacion
3) ADDAX_Preparar_Tabla_ETL


5. CONSIDERACIONES
------------------
- Las macros trabajan sobre la hoja activa.
- No requieren selección manual de rangos.
- No modificar nombres de columnas sin ajustar las macros.
- Los reportes generados deben conservarse como respaldo institucional.


6. ESTADO DEL PROYECTO
---------------------
- Probado con datos reales
- Estructura limpia y ordenada
- Listo para empaquetar como complemento XLAM institucional
