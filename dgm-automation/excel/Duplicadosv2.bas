Attribute VB_Name = "Duplicadosv2"
Option Explicit

' =========================================================
' AUDITORIA v2 - DUPLICADOS (EXPEDIENTE+INGRESO) + ESTADO (D)
'   Hoja 2: listado de filas involucradas en duplicados
'   Hoja 3: grupos duplicados con ESTADO distinto (o mensaje si todos iguales)
' =========================================================
Public Sub AUDIT_Duplicadosv2()

    Dim wsData As Worksheet
    Dim wbData As Workbook
    Dim wsDup As Worksheet, wsEst As Worksheet

    Dim lastRow As Long, lastCol As Long
    Dim colExp As Long, colIng As Long, colEst As Long
    Dim r As Long

    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Activa una hoja de datos válida.", vbCritical
        Exit Sub
    End If
    Set wsData = ActiveSheet
    Set wbData = wsData.Parent

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo EH

    lastRow = wsData.Cells(wsData.Rows.count, 1).End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.count).End(xlToLeft).Column

    If lastRow < 2 Then
        MsgBox "No hay filas de datos para analizar.", vbExclamation
        GoTo Salir
    End If

    ' --- Buscar columnas por encabezado; fallback por posición
    colExp = FindCol(wsData, "EXPEDIENTE"): If colExp = 0 Then colExp = 1   'A
    colIng = FindCol(wsData, "INGRESO"):   If colIng = 0 Then colIng = 5   'E
    colEst = FindCol(wsData, "ESTADO"):    If colEst = 0 Then colEst = 4   'D

    If colExp > lastCol Or colIng > lastCol Or colEst > lastCol Then
        MsgBox "No se pudieron determinar columnas válidas para EXPEDIENTE/INGRESO/ESTADO.", vbCritical
        GoTo Salir
    End If

    ' --- Hoja 2: duplicados
    Set wsDup = GetOrCreateSheetN(wbData, 2)
    wsDup.Visible = xlSheetVisible
    wsDup.Cells.Clear

    ' --- Hoja 3: análisis de ESTADO
    Set wsEst = GetOrCreateSheetN(wbData, 3)
    wsEst.Visible = xlSheetVisible
    wsEst.Cells.Clear

    ' --- Diccionarios
    Dim dictCount As Object, dictRows As Object
    Set dictCount = CreateObject("Scripting.Dictionary") ' key -> count
    Set dictRows = CreateObject("Scripting.Dictionary")  ' key -> "2,7,15"

    ' --- 1er pase: contar duplicados por clave (EXP + ING)
    For r = 2 To lastRow
        Dim exp As String, ingKey As String, key As String

        exp = Trim(CStr(wsData.Cells(r, colExp).Value))
        ingKey = NormalizeIngreso(wsData.Cells(r, colIng).Value)

        If exp <> "" And ingKey <> "" Then
            key = exp & "||" & ingKey

            If Not dictCount.Exists(key) Then
                dictCount.Add key, 1
                dictRows.Add key, CStr(r)
            Else
                dictCount(key) = CLng(dictCount(key)) + 1
                dictRows(key) = CStr(dictRows(key)) & "," & CStr(r)
            End If
        End If
    Next r

    ' ======================================================
    ' HOJA 2: REPORTE DE DUPLICADOS (todas las filas)
    ' ======================================================
    wsDup.Range("A1:H1").Value = Array( _
        "CLAVE", "EXPEDIENTE", "INGRESO_NORMALIZADO", "CANTIDAD", _
        "FILA", "INGRESO_ORIGINAL", "ESTADO", "HOJA_ORIGEN" _
    )
    wsDup.Rows(1).Font.Bold = True

    Dim outRowDup As Long: outRowDup = 2
    Dim k As Variant
    Dim totalClavesDup As Long, totalFilasDup As Long

    For Each k In dictCount.Keys
        If CLng(dictCount(k)) > 1 Then
            totalClavesDup = totalClavesDup + 1

            Dim parts() As String: parts = Split(CStr(k), "||")
            Dim expOut As String: expOut = parts(0)
            Dim ingOut As String: ingOut = parts(1)

            Dim filas() As String: filas = Split(CStr(dictRows(k)), ",")
            Dim i As Long
            For i = LBound(filas) To UBound(filas)
                Dim filaNum As Long: filaNum = CLng(filas(i))

                wsDup.Cells(outRowDup, 1).Value = CStr(k)
                wsDup.Cells(outRowDup, 2).Value = expOut
                wsDup.Cells(outRowDup, 3).Value = ingOut
                wsDup.Cells(outRowDup, 4).Value = CLng(dictCount(k))
                wsDup.Cells(outRowDup, 5).Value = filaNum
                wsDup.Cells(outRowDup, 6).Value = wsData.Cells(filaNum, colIng).Text
                wsDup.Cells(outRowDup, 7).Value = wsData.Cells(filaNum, colEst).Text
                wsDup.Cells(outRowDup, 8).Value = wsData.Name

                outRowDup = outRowDup + 1
                totalFilasDup = totalFilasDup + 1
            Next i
        End If
    Next k

    wsDup.Columns.AutoFit

    ' ======================================================
    ' HOJA 3: ANALISIS DE ESTADO DENTRO DE DUPLICADOS
    '   - Solo lista grupos donde ESTADO difiere
    '   - Si todos iguales, MsgBox
    ' ======================================================
    wsEst.Range("A1:G1").Value = Array( _
        "CLAVE", "EXPEDIENTE", "INGRESO_NORMALIZADO", "CANTIDAD", _
        "ESTADOS_DISTINTOS", "FILAS_INVOLUCRADAS", "HOJA_ORIGEN" _
    )
    wsEst.Rows(1).Font.Bold = True

    Dim outRowEst As Long: outRowEst = 2
    Dim gruposConEstadoDiferente As Long
    Dim gruposConEstadoIgual As Long

    For Each k In dictCount.Keys
        If CLng(dictCount(k)) > 1 Then

            Dim filas2() As String: filas2 = Split(CStr(dictRows(k)), ",")
            Dim dictEstados As Object: Set dictEstados = CreateObject("Scripting.Dictionary")

            Dim j As Long
            For j = LBound(filas2) To UBound(filas2)
                Dim f As Long: f = CLng(filas2(j))

                Dim estRaw As String
                estRaw = Trim(CStr(wsData.Cells(f, colEst).Value))

                Dim estKey As String
                estKey = UCase$(CollapseSpaces(CleanText(estRaw)))

                If estKey = "" Then estKey = "(VACIO)"

                If Not dictEstados.Exists(estKey) Then dictEstados.Add estKey, 1
            Next j

            Dim parts2() As String: parts2 = Split(CStr(k), "||")

            If dictEstados.count > 1 Then
                gruposConEstadoDiferente = gruposConEstadoDiferente + 1

                wsEst.Cells(outRowEst, 1).Value = CStr(k)
                wsEst.Cells(outRowEst, 2).Value = parts2(0)
                wsEst.Cells(outRowEst, 3).Value = parts2(1)
                wsEst.Cells(outRowEst, 4).Value = CLng(dictCount(k))
                wsEst.Cells(outRowEst, 5).Value = JoinDictKeys(dictEstados, " | ")
                wsEst.Cells(outRowEst, 6).Value = CStr(dictRows(k))
                wsEst.Cells(outRowEst, 7).Value = wsData.Name

                outRowEst = outRowEst + 1
            Else
                gruposConEstadoIgual = gruposConEstadoIgual + 1
            End If

        End If
    Next k

    wsEst.Columns.AutoFit

    If totalClavesDup = 0 Then
        wsEst.Range("A3").Value = "No se detectaron duplicados por (EXPEDIENTE + INGRESO)."
        wsDup.Activate
        MsgBox "No se detectaron duplicados por (EXPEDIENTE + INGRESO).", vbInformation
        GoTo Salir
    End If

    If gruposConEstadoDiferente = 0 Then
        wsEst.Range("A3").Value = "Duplicados detectados, pero TODOS tienen el mismo ESTADO dentro de cada grupo."
        wsEst.Range("A4").Value = "Grupos duplicados analizados:"
        wsEst.Range("B4").Value = totalClavesDup
        wsEst.Activate

        MsgBox "Duplicados detectados: " & totalClavesDup & vbCrLf & _
               "Resultado: en TODOS los grupos, el ESTADO es IGUAL entre registros duplicados.", vbInformation
    Else
        wsEst.Activate
        MsgBox "Duplicados detectados: " & totalClavesDup & vbCrLf & _
               "Grupos con ESTADO distinto: " & gruposConEstadoDiferente & vbCrLf & _
               "Revisá la Hoja 3.", vbExclamation
    End If

Salir:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

EH:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
    Resume Salir

End Sub

' ========================= Helpers =========================

Private Function FindCol(ws As Worksheet, headerName As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        If NormalizeHeader(ws.Cells(1, c).Value) = NormalizeHeader(headerName) Then
            FindCol = c
            Exit Function
        End If
    Next c
    FindCol = 0
End Function

Private Function NormalizeHeader(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v)
    s = CleanText(s)
    s = UCase$(s)
    s = CollapseSpaces(s)
    NormalizeHeader = s
End Function

Private Function NormalizeIngreso(ByVal v As Variant) As String
    If IsDate(v) Then
        NormalizeIngreso = Format$(CDate(v), "yyyy-mm-dd")
    Else
        Dim s As String
        s = CStr(v)
        s = UCase$(CollapseSpaces(CleanText(s)))
        NormalizeIngreso = s
    End If
End Function

Private Function CleanText(ByVal s As String) As String
    s = Application.WorksheetFunction.Clean(s)
    s = Replace(s, ChrW(160), " ") ' NBSP
    CleanText = Trim$(s)
End Function

Private Function CollapseSpaces(ByVal s As String) As String
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    CollapseSpaces = s
End Function

Private Function GetOrCreateSheetN(wb As Workbook, ByVal n As Long) As Worksheet
    If wb.Worksheets.count >= n Then
        Set GetOrCreateSheetN = wb.Worksheets(n)
    Else
        Dim i As Long
        For i = wb.Worksheets.count + 1 To n
            wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.count)
        Next i
        Set GetOrCreateSheetN = wb.Worksheets(n)
    End If
End Function

Private Function JoinDictKeys(dict As Object, ByVal sep As String) As String
    Dim k As Variant, s As String
    For Each k In dict.Keys
        If s <> "" Then s = s & sep
        s = s & CStr(k)
    Next k
    JoinDictKeys = s
End Function
