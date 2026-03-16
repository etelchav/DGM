Attribute VB_Name = "modADDAX_PreETL"
Option Explicit

' ==================================================
' === ADDAX | PREPARAR TABLA PARA ETL ARCGIS =======
' ==================================================

Private Const RUTA_REPORTES As String = _
"C:\Users\echavarria\OneDrive - MINAE Costa Rica\2-REPORTES\Reportes_Macros_Excel\"

' -------------------------
' UTILIDADES
' -------------------------
Private Function RemoveAccents(ByVal txt As String) As String
    Dim a As String, b As String, i As Long
    a = "ÁÀÄÂÉÈËÊÍÌÏÎÓÒÖÔÚÙÜÛÑáàäâéèëêíìïîóòöôúùüûñ"
    b = "AAAAEEEEIIIIOOOOUUUUNaaaaeeeeiiiioooouuuun"
    For i = 1 To Len(a)
        txt = Replace(txt, Mid(a, i, 1), Mid(b, i, 1))
    Next i
    RemoveAccents = txt
End Function

Private Function CleanKey(ByVal txt As String) As String
    txt = Trim(txt)
    txt = RemoveAccents(txt)
    txt = LCase(txt)
    Do While InStr(txt, "  ") > 0
        txt = Replace(txt, "  ", " ")
    Loop
    CleanKey = txt
End Function

Private Function BuscarColumna(ByVal ws As Worksheet, ByVal nombre As String) As Long
    Dim c As Long, lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        If CleanKey(ws.Cells(1, c).Value) = CleanKey(nombre) Then
            BuscarColumna = c
            Exit Function
        End If
    Next c
    BuscarColumna = 0
End Function

' -------------------------
' MACRO PRINCIPAL
' -------------------------
Public Sub ADDAX_Preparar_Tabla_ETL()

    Dim ws As Worksheet
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Activá una hoja de datos válida.", vbCritical
        Exit Sub
    End If
    Set ws = ActiveSheet

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    ' ==================================================
    ' 0) FORZAR TEXTO EN COLUMNAS A, B, F
    ' ==================================================
    Dim colTxt As Variant, r As Long
    For Each colTxt In Array(1, 2, 6) ' A, B, F
        ws.Columns(colTxt).NumberFormat = "@"
        For r = 2 To lastRow
            If Not IsEmpty(ws.Cells(r, colTxt).Value) Then
                ws.Cells(r, colTxt).Value = CStr(ws.Cells(r, colTxt).Value)
            End If
        Next r
    Next colTxt

    ' ==================================================
    ' 1) NORMALIZAR ENCABEZADOS (ETL-safe)
    ' ==================================================
    Dim c As Long, h As String
    For c = 1 To lastCol
        h = CStr(ws.Cells(1, c).Value)
        h = UCase(RemoveAccents(Trim(h)))
        ws.Cells(1, c).Value = h
    Next c

    ' ==================================================
    ' 2) UBICAR COLUMNAS CLAVE
    ' ==================================================
    Dim colEstado As Long, colNombre As Long
    Dim colProyecto As Long, colProv As Long, colCanton As Long, colDistrito As Long

    colEstado = BuscarColumna(ws, "ESTADO")
    colNombre = BuscarColumna(ws, "NOMBRE")
    colProyecto = BuscarColumna(ws, "PROYECTO")
    colProv = BuscarColumna(ws, "PROVINCIA")
    colCanton = BuscarColumna(ws, "CANTON")
    colDistrito = BuscarColumna(ws, "DISTRITO")

    If colEstado = 0 Or colNombre = 0 Then
        MsgBox "Faltan columnas requeridas: ESTADO o NOMBRE.", vbCritical
        GoTo Salir
    End If

    ' ==================================================
    ' 3) DICCIONARIO DE ESTADOS (CATÁLOGO)
    ' ==================================================
    Dim estados As Object
    Set estados = CreateObject("Scripting.Dictionary")

    estados.Add "archivado", "ARCHIVADO"
    estados.Add "reservado", "RESERVADO"
    estados.Add "vigente", "VIGENTE"
    estados.Add "permiso especial", "PERMISO_ESPECIAL"
    estados.Add "permiso_especial", "PERMISO_ESPECIAL"
    estados.Add "formalizado", "FORMALIZADO"
    estados.Add "no ubicado", "NO_UBICADO"
    estados.Add "en revision legal", "EN_REVISION_LEGAL"
    estados.Add "suspendido", "SUSPENDIDO"
    estados.Add "extinto", "EXTINTO"
    estados.Add "temporal", "PERMISO_ESPECIAL"
    estados.Add "pendiente de ubicar", "PENDIENTE_UBICAR"

    ' ==================================================
    ' 4) CREAR LOG EXTERNO
    ' ==================================================
    Dim logWb As Workbook, logWs As Worksheet
    Set logWb = Workbooks.Add
    Set logWs = logWb.Sheets(1)
    logWs.Name = "LOG_ADDAX_PRE_ETL"
    logWs.Range("A1:D1").Value = Array("FILA", "COLUMNA", "ORIGINAL", "NORMALIZADO")

    Dim logRow As Long: logRow = 2
    Dim vOrig As String, vKey As String, vNew As String

    ' ==================================================
    ' 5) PROCESAR FILAS
    ' ==================================================
    For r = 2 To lastRow

        ' ---- ESTADO
        vOrig = CStr(ws.Cells(r, colEstado).Value)
        vKey = CleanKey(vOrig)
        If estados.Exists(vKey) Then
            vNew = estados(vKey)
            If vOrig <> vNew Then
                ws.Cells(r, colEstado).Value = vNew
                logWs.Cells(logRow, 1).Value = r
                logWs.Cells(logRow, 2).Value = "ESTADO"
                logWs.Cells(logRow, 3).Value = vOrig
                logWs.Cells(logRow, 4).Value = vNew
                logRow = logRow + 1
            End If
        End If

        ' ---- NOMBRE
        vOrig = CStr(ws.Cells(r, colNombre).Value)
        vNew = UCase(RemoveAccents(Trim(vOrig)))
        If vOrig <> vNew Then
            ws.Cells(r, colNombre).Value = vNew
            logWs.Cells(logRow, 1).Value = r
            logWs.Cells(logRow, 2).Value = "NOMBRE"
            logWs.Cells(logRow, 3).Value = vOrig
            logWs.Cells(logRow, 4).Value = vNew
            logRow = logRow + 1
        End If

        ' ---- CAMPOS TERRITORIALES Y PROYECTO
        Call NormalizarCampo(ws, r, colProyecto, "PROYECTO", logWs, logRow)
        Call NormalizarCampo(ws, r, colProv, "PROVINCIA", logWs, logRow)
        Call NormalizarCampo(ws, r, colCanton, "CANTON", logWs, logRow)
        Call NormalizarCampo(ws, r, colDistrito, "DISTRITO", logWs, logRow)

    Next r

    logWs.Columns.AutoFit

    ' ==================================================
    ' 6) GUARDAR LOG
    ' ==================================================
    Dim rutaLog As String
    rutaLog = RUTA_REPORTES & _
              "LOG_ADDAX_PRE_ETL_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"

    Application.DisplayAlerts = False
    logWb.SaveAs rutaLog, xlOpenXMLWorkbook
    logWb.Close False
    Application.DisplayAlerts = True

    MsgBox "ADDAX PRE-ETL completado. LOG guardado en:" & vbCrLf & rutaLog, vbInformation

Salir:
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

' -------------------------
' NORMALIZAR CAMPO GENERICO
' -------------------------
Private Sub NormalizarCampo( _
    ByVal ws As Worksheet, _
    ByVal r As Long, _
    ByVal colIdx As Long, _
    ByVal nombreCampo As String, _
    ByRef logWs As Worksheet, _
    ByRef logRow As Long)

    If colIdx = 0 Then Exit Sub

    Dim vOrig As String, vNew As String
    vOrig = CStr(ws.Cells(r, colIdx).Value)
    vNew = UCase(RemoveAccents(Trim(vOrig)))

    If vOrig <> vNew Then
        ws.Cells(r, colIdx).Value = vNew
        logWs.Cells(logRow, 1).Value = r
        logWs.Cells(logRow, 2).Value = nombreCampo
        logWs.Cells(logRow, 3).Value = vOrig
        logWs.Cells(logRow, 4).Value = vNew
        logRow = logRow + 1
    End If
End Sub


