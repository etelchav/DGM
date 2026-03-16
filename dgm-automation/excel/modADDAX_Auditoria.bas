Attribute VB_Name = "modADDAX_Auditoria"
Option Explicit

' =====================================================
' ADDAX | AUDITORIA POST NORMALIZACION
' =====================================================

Private Const RUTA_REPORTES As String = _
"C:\Users\echavarria\OneDrive - MINAE Costa Rica\2-REPORTES\Reportes_Macros_Excel\"

' =====================================================
' MACRO PRINCIPAL
' =====================================================
Public Sub ADDAX_Auditar_Tabla_Post_Normalizacion()

    Dim ws As Worksheet
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Activa una hoja de datos válida.", vbCritical
        Exit Sub
    End If
    Set ws = ActiveSheet

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' =========================
    ' CONTADORES RESUMEN
    ' =========================
    Dim cntDuplicados As Long
    Dim cntEliminados As Long
    Dim cntEstadoCorregido As Long

    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    ' =====================================================
    ' BUSCAR COLUMNAS CLAVE
    ' =====================================================
    Dim colExp As Long, colEstado As Long
    colExp = BuscarColumna(ws, "EXPEDIENTE")
    colEstado = BuscarColumna(ws, "ESTADO")

    If colExp = 0 Or colEstado = 0 Then
        MsgBox "Faltan columnas requeridas: EXPEDIENTE o ESTADO.", vbCritical
        GoTo Salir
    End If

    ' =====================================================
    ' CREAR / LIMPIAR HOJA AUDITORIA
    ' =====================================================
    Dim auditWs As Worksheet
    On Error Resume Next
    Set auditWs = ThisWorkbook.Worksheets("AUDITORIA_ADDAX")
    On Error GoTo 0

    If auditWs Is Nothing Then
        Set auditWs = ThisWorkbook.Worksheets.Add(After:=ws)
        auditWs.Name = "AUDITORIA_ADDAX"
    Else
        auditWs.Cells.Clear
    End If

    auditWs.Range("A1:C1").Value = Array("EXPEDIENTE", "TIPO", "DETALLE")
    Dim auditRow As Long: auditRow = 2

    ' =====================================================
    ' DICCIONARIO DE EXPEDIENTES
    ' =====================================================
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = 2 To lastRow
        Dim exp As String
        exp = Trim(ws.Cells(r, colExp).Value)

        If exp <> "" Then
            If Not dict.Exists(exp) Then
                dict.Add exp, r
            Else
                cntDuplicados = cntDuplicados + 1

                Dim filaBase As Long
                filaBase = dict(exp)

                Dim scoreBase As Long, scoreNueva As Long
                scoreBase = ConteoDatos(ws, filaBase, lastCol)
                scoreNueva = ConteoDatos(ws, r, lastCol)

                If scoreNueva > scoreBase Then
                    ws.Rows(filaBase).Delete
                    dict(exp) = r - 1
                Else
                    ws.Rows(r).Delete
                End If

                cntEliminados = cntEliminados + 1

                auditWs.Cells(auditRow, 1).Value = exp
                auditWs.Cells(auditRow, 2).Value = "DUPLICADO"
                auditWs.Cells(auditRow, 3).Value = "Se conservó el registro más completo"
                auditRow = auditRow + 1

                lastRow = lastRow - 1
                r = r - 1
            End If
        End If
    Next r

    ' =====================================================
    ' ESTADO VACIO ? PENDIENTE
    ' =====================================================
    For r = 2 To lastRow
        If Trim(ws.Cells(r, colEstado).Value) = "" Then
            ws.Cells(r, colEstado).Value = "PENDIENTE"
            cntEstadoCorregido = cntEstadoCorregido + 1

            auditWs.Cells(auditRow, 1).Value = ws.Cells(r, colExp).Value
            auditWs.Cells(auditRow, 2).Value = "ESTADO_VACIO"
            auditWs.Cells(auditRow, 3).Value = "Asignado automáticamente a PENDIENTE"
            auditRow = auditRow + 1
        End If
    Next r

    auditWs.Columns.AutoFit

    ' =====================================================
    ' EXPORTAR REPORTE A DISCO
    ' =====================================================
    Dim wbRep As Workbook
    Dim nombreArchivo As String

    nombreArchivo = "AUDITORIA_ADDAX_" & Format(Now, "yyyy-mm-dd_hhmmss") & ".xlsx"

    auditWs.Copy
    Set wbRep = ActiveWorkbook

    Application.DisplayAlerts = False
    wbRep.SaveAs Filename:=RUTA_REPORTES & nombreArchivo, _
                 FileFormat:=xlOpenXMLWorkbook
    wbRep.Close False
    Application.DisplayAlerts = True

    ' =====================================================
    ' RESUMEN VISUAL EN PANTALLA
    ' =====================================================
    MsgBox _
        "AUDITORÍA ADDAX – RESUMEN" & vbCrLf & vbCrLf & _
        "Registros finales: " & lastRow - 1 & vbCrLf & _
        "Duplicados detectados: " & cntDuplicados & vbCrLf & _
        "Registros eliminados: " & cntEliminados & vbCrLf & _
        "Estados corregidos: " & cntEstadoCorregido & vbCrLf & vbCrLf & _
        "Reporte generado en:" & vbCrLf & _
        RUTA_REPORTES, _
        vbInformation, "Auditoría ADDAX"

Salir:
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

' =====================================================
' FUNCIONES DE APOYO
' =====================================================
Private Function ConteoDatos(ws As Worksheet, fila As Long, lastCol As Long) As Long
    Dim c As Long, count As Long
    For c = 1 To lastCol
        If Trim(ws.Cells(fila, c).Value) <> "" Then count = count + 1
    Next c
    ConteoDatos = count
End Function

Private Function BuscarColumna(ws As Worksheet, nombre As String) As Long
    Dim c As Long, lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        If Normaliza(ws.Cells(1, c).Value) = Normaliza(nombre) Then
            BuscarColumna = c
            Exit Function
        End If
    Next c
    BuscarColumna = 0
End Function

Private Function Normaliza(txt As String) As String
    txt = UCase(txt)
    txt = Application.WorksheetFunction.Clean(txt)
    txt = Replace(txt, "Á", "A")
    txt = Replace(txt, "É", "E")
    txt = Replace(txt, "Í", "I")
    txt = Replace(txt, "Ó", "O")
    txt = Replace(txt, "Ú", "U")
    txt = Replace(txt, "Ü", "U")
    Normaliza = Trim(txt)
End Function


