Attribute VB_Name = "modNormalizar_BD_HERMES"
Option Explicit

' =========================
' UTILIDADES
' =========================
Private Function TrimCompact(ByVal s As String) As String
    Dim t As String
    t = Trim$(CStr(s))
    Do While InStr(t, "  ") > 0
        t = Replace$(t, "  ", " ")
    Loop
    TrimCompact = t
End Function

Private Function RemoveDiacritics(ByVal s As String) As String
    Dim t As String: t = CStr(s)
    t = Replace$(t, "á", "a"): t = Replace$(t, "é", "e")
    t = Replace$(t, "í", "i"): t = Replace$(t, "ó", "o")
    t = Replace$(t, "ú", "u"): t = Replace$(t, "ñ", "n")
    t = Replace$(t, "Á", "A"): t = Replace$(t, "É", "E")
    t = Replace$(t, "Í", "I"): t = Replace$(t, "Ó", "O")
    t = Replace$(t, "Ú", "U"): t = Replace$(t, "Ñ", "N")
    RemoveDiacritics = t
End Function

Private Function KeyNorm(ByVal s As String) As String
    KeyNorm = LCase$(RemoveDiacritics(TrimCompact(CStr(s))))
End Function

' =========================
' BUSCAR COLUMNAS
' =========================
Private Function FindCol(ws As Worksheet, ByVal header As String) As Long
    Dim c As Long, lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        If KeyNorm(ws.Cells(1, c).Value) = KeyNorm(header) Then
            FindCol = c
            Exit Function
        End If
    Next c
    FindCol = 0
End Function

' =========================
' MACRO PRINCIPAL
' =========================
Public Sub Normalizar_HERMES_Concesiona_Proyecto_Estado()

    Dim ws As Worksheet
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Activá una hoja de datos válida.", vbCritical
        Exit Sub
    End If
    Set ws = ActiveSheet

    Dim prevCalc As XlCalculation
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        prevCalc = .Calculation
        .Calculation = xlCalculationManual
    End With

    On Error GoTo CleanExit

    ' --- localizar columnas
    Dim cExp As Long, cEstado As Long, cCon As Long, cProy As Long
    cExp = FindCol(ws, "Expediente")
    cEstado = FindCol(ws, "Estado")
    cCon = FindCol(ws, "Concesiona")
    cProy = FindCol(ws, "Proyecto")

    If cExp = 0 Or cEstado = 0 Then
        MsgBox "No se encontró Expediente o Estado.", vbCritical
        GoTo CleanExit
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, cExp).End(xlUp).Row

    ' --- diccionario de ESTADOS
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "archivado", "ARCHIVADO"
    dict.Add "reservado", "RESERVADO"
    dict.Add "vigente", "VIGENTE"
    dict.Add "permiso especial", "PERMISO_ESPECIAL"
    dict.Add "permiso_especial", "PERMISO_ESPECIAL"
    dict.Add "formalizado", "FORMALIZADO"
    dict.Add "no ubicado", "NO_UBICADO"
    dict.Add "no_ubicado", "NO_UBICADO"
    dict.Add "en revision legal", "EN_REVISION_LEGAL"
    dict.Add "en revisión legal", "EN_REVISION_LEGAL"
    dict.Add "revisiã³n legal", "EN_REVISION_LEGAL"
    dict.Add "suspendido", "SUSPENDIDO"
    dict.Add "extinto", "EXTINTO"
    dict.Add "temporal", "PERMISO_ESPECIAL"
    dict.Add "pendiente de ubicar", "PENDIENTE_UBICAR"
    dict.Add "pendiente_ubicar", "PENDIENTE_UBICAR"

    Dim r As Long, k As String

    For r = 2 To lastRow

        ' CONCESIONA ? MAYÚSCULAS
        If cCon <> 0 Then
            ws.Cells(r, cCon).Value = UCase$(TrimCompact(ws.Cells(r, cCon).Value))
        End If

        ' PROYECTO ? MAYÚSCULAS
        If cProy <> 0 Then
            ws.Cells(r, cProy).Value = UCase$(TrimCompact(ws.Cells(r, cProy).Value))
        End If

        ' ESTADO ? NORMALIZADO
        k = KeyNorm(ws.Cells(r, cEstado).Value)
        If dict.Exists(k) Then
            ws.Cells(r, cEstado).Value = dict(k)
        End If

    Next r

    ' --- mover ESTADO a la derecha de EXPEDIENTE
    If cEstado <> cExp + 1 Then
        ws.Columns(cEstado).Cut
        ws.Columns(cExp + 1).Insert Shift:=xlToRight
        Application.CutCopyMode = False
    End If

CleanExit:
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = prevCalc
    End With

    MsgBox "Normalización aplicada: Concesiona, Proyecto, Estado y reordenamiento de columnas.", vbInformation

End Sub

