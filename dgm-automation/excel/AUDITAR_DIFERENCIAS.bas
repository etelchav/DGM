Attribute VB_Name = "AUDITAR_DIFERENCIAS"
Option Explicit

Public Sub Auditar_Migracion_HERMES_vs_ADDAX()

    Dim rutaHermes As String, rutaAddax As String
    Dim wbH As Workbook, wbA As Workbook
    Dim wsH As Worksheet, wsA As Worksheet
    Dim wbR As Workbook
    Dim wsDif As Worksheet, wsOk As Worksheet
    Dim dictH As Object, dictA As Object
    Dim colExpH As Long, colEstH As Long
    Dim colExpA As Long, colEstA As Long
    Dim lastRowH As Long, lastRowA As Long
    Dim r As Long, filaDif As Long, filaOk As Long
    Dim k As Variant
    
    Dim salidaPath As String, salidaFile As String
    salidaPath = "C:\Users\echavarria\OneDrive - MINAE Costa Rica\2-REPORTES\Reportes_Macros_Excel\Rep_migracion\"
    salidaFile = salidaPath & "Rep_Migracion_" & Format(Now, "yyyymmdd_hhnn") & ".xlsx"
    
    '========================
    ' Selección de archivos
    '========================
    rutaHermes = SeleccionarArchivo("Seleccione el Excel HERMES (base)")
    If rutaHermes = "" Then Exit Sub
    
    rutaAddax = SeleccionarArchivo("Seleccione el Excel ADDAX (migrado)")
    If rutaAddax = "" Then Exit Sub
    
    Set wbH = Workbooks.Open(rutaHermes, ReadOnly:=True)
    Set wbA = Workbooks.Open(rutaAddax, ReadOnly:=True)
    
    Set wsH = wbH.Sheets(1)
    Set wsA = wbA.Sheets(1)
    
    Set dictH = CreateObject("Scripting.Dictionary")
    Set dictA = CreateObject("Scripting.Dictionary")
    
    '========================
    ' Buscar columnas
    '========================
    colExpH = BuscarColumna(wsH, "EXPEDIENTE")
    colEstH = BuscarColumna(wsH, "ESTADO")
    colExpA = BuscarColumna(wsA, "EXPEDIENTE")
    colEstA = BuscarColumna(wsA, "ESTADO")
    
    If colExpH = 0 Or colEstH = 0 Or colExpA = 0 Or colEstA = 0 Then
        MsgBox "No se encontraron columnas EXPEDIENTE o ESTADO.", vbCritical
        GoTo CerrarTodo
    End If
    
    '========================
    ' Cargar HERMES (TEXTO)
    '========================
    lastRowH = wsH.Cells(wsH.Rows.count, colExpH).End(xlUp).Row
    For r = 2 To lastRowH
        Dim expH As String
        expH = Trim(CStr(wsH.Cells(r, colExpH).Value))
        If expH <> "" Then
            dictH(expH) = Trim(CStr(wsH.Cells(r, colEstH).Value))
        End If
    Next r
    
    '========================
    ' Cargar ADDAX (TEXTO)
    '========================
    lastRowA = wsA.Cells(wsA.Rows.count, colExpA).End(xlUp).Row
    For r = 2 To lastRowA
        Dim expA As String
        expA = Trim(CStr(wsA.Cells(r, colExpA).Value))
        If expA <> "" Then
            dictA(expA) = Trim(CStr(wsA.Cells(r, colEstA).Value))
        End If
    Next r
    
    '========================
    ' Crear libro de reporte
    '========================
    Set wbR = Workbooks.Add
    Set wsDif = wbR.Sheets(1)
    wsDif.Name = "REPORTE_MIGRACION"
    
    Set wsOk = wbR.Sheets.Add(After:=wsDif)
    wsOk.Name = "COINCIDENCIAS_OK"
    
    ' Formato TEXTO para EXPEDIENTE
    wsDif.Columns(1).NumberFormat = "@"
    wsOk.Columns(1).NumberFormat = "@"
    
    ' Encabezados
    wsDif.Range("A1:D1").Value = Array("EXPEDIENTE", "ESTADO_HERMES", "ESTADO_ADDAX", "OBSERVACION")
    wsOk.Range("A1:B1").Value = Array("EXPEDIENTE", "ESTADO")
    
    filaDif = 2
    filaOk = 2
    
    Dim faltanAddax As Long, faltanHermes As Long, estadosDif As Long, okCount As Long
    
    '========================
    ' Comparar desde HERMES
    '========================
    For Each k In dictH.Keys
        If Not dictA.exists(k) Then
            wsDif.Cells(filaDif, 1).Value = k
            wsDif.Cells(filaDif, 2).Value = dictH(k)
            wsDif.Cells(filaDif, 4).Value = "NO EXISTE EN ADDAX"
            filaDif = filaDif + 1
            faltanAddax = faltanAddax + 1
            
        ElseIf dictH(k) <> dictA(k) Then
            wsDif.Cells(filaDif, 1).Value = k
            wsDif.Cells(filaDif, 2).Value = dictH(k)
            wsDif.Cells(filaDif, 3).Value = dictA(k)
            wsDif.Cells(filaDif, 4).Value = "ESTADO DIFERENTE"
            filaDif = filaDif + 1
            estadosDif = estadosDif + 1
            
        Else
            wsOk.Cells(filaOk, 1).Value = k
            wsOk.Cells(filaOk, 2).Value = dictH(k)
            filaOk = filaOk + 1
            okCount = okCount + 1
        End If
    Next k
    
    '========================
    ' Extra en ADDAX
    '========================
    For Each k In dictA.Keys
        If Not dictH.exists(k) Then
            wsDif.Cells(filaDif, 1).Value = k
            wsDif.Cells(filaDif, 3).Value = dictA(k)
            wsDif.Cells(filaDif, 4).Value = "NO EXISTE EN HERMES"
            filaDif = filaDif + 1
            faltanHermes = faltanHermes + 1
        End If
    Next k
    
    wsDif.Columns.AutoFit
    wsOk.Columns.AutoFit
    
    wbR.SaveAs salidaFile
    wbR.Close False
    
    MsgBox _
        "AUDITORÍA DE MIGRACIÓN FINALIZADA" & vbCrLf & vbCrLf & _
        "Archivo generado:" & vbCrLf & salidaFile & vbCrLf & vbCrLf & _
        "HERMES: " & dictH.count & vbCrLf & _
        "ADDAX: " & dictA.count & vbCrLf & _
        "? Coincidencias correctas: " & okCount & vbCrLf & _
        "? Faltan en ADDAX: " & faltanAddax & vbCrLf & _
        "? Faltan en HERMES: " & faltanHermes & vbCrLf & _
        "? Estados distintos: " & estadosDif, _
        vbInformation

CerrarTodo:
    wbH.Close False
    wbA.Close False

End Sub

Function SeleccionarArchivo(titulo As String) As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = titulo
        .Filters.Clear
        .Filters.Add "Excel", "*.xlsx; *.xlsm"
        .AllowMultiSelect = False
        If .Show = -1 Then
            SeleccionarArchivo = .SelectedItems(1)
        Else
            SeleccionarArchivo = ""
        End If
    End With
End Function
Function BuscarColumna(ws As Worksheet, nombre As String) As Long
    Dim c As Long
    For c = 1 To ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
        If UCase(Trim(ws.Cells(1, c).Value)) = UCase(nombre) Then
            BuscarColumna = c
            Exit Function
        End If
    Next c
    BuscarColumna = 0
End Function



