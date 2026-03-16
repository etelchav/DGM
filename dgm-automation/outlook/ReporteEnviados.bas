Attribute VB_Name = "ReporteEnviados"
Sub ExportarCorreosPorMes()

    ' --- Declaraciones ---
    Dim olApp As Outlook.Application
    Dim olNs As Outlook.NameSpace
    Dim carpetaEnviados As Outlook.MAPIFolder
    Dim item As Object
    Dim correo As Outlook.MailItem
    Dim xlApp As Object, xlWB As Object, xlSheet As Object
    Dim fila As Long
    Dim rutaArchivo As String
    Dim carpetaDestino As String
    Dim mesFiltro As Integer, annoFiltro As Integer
    Dim nombreArchivo As String
    Dim inputMes As Variant
    Dim inputAnno As Variant

    ' --- Solicitar Mes ---
    inputMes = InputBox("Ingrese el MES (1-12) del reporte:", "Seleccionar Mes")
    If Not IsNumeric(inputMes) Then Exit Sub
    mesFiltro = CInt(inputMes)

    If mesFiltro < 1 Or mesFiltro > 12 Then
        MsgBox "Mes inválido. Debe estar entre 1 y 12.", vbExclamation
        Exit Sub
    End If

    ' --- Solicitar Año ---
    inputAnno = InputBox("Ingrese el AÑO del reporte (ej: 2026):", "Seleccionar Año")
    If Not IsNumeric(inputAnno) Then Exit Sub
    annoFiltro = CInt(inputAnno)

    ' --- Ruta fija institucional ---
    carpetaDestino = "C:\Users\echavarria\OneDrive - MINAE Costa Rica\2-REPORTES\Reportes_Outlook\"

    ' --- Crear carpeta si no existe ---
    If Dir(carpetaDestino, vbDirectory) = "" Then
        MkDir carpetaDestino
    End If

    ' --- Inicializar Outlook ---
    Set olApp = Outlook.Application
    Set olNs = olApp.GetNamespace("MAPI")
    Set carpetaEnviados = olNs.GetDefaultFolder(olFolderSentMail)

    ' --- Inicializar Excel ---
    Set xlApp = CreateObject("Excel.Application")
    Set xlWB = xlApp.Workbooks.Add
    Set xlSheet = xlWB.Sheets(1)

    ' --- Encabezados ---
    With xlSheet
        .Cells(1, 1).Value = "Fecha de envío"
        .Cells(1, 2).Value = "Asunto"
        .Cells(1, 3).Value = "Destinatario(s)"
    End With

    fila = 2

    ' --- Recorrer correos enviados ---
    For Each item In carpetaEnviados.items
        If item.Class = olMail Then
            Set correo = item
            If Month(correo.SentOn) = mesFiltro And Year(correo.SentOn) = annoFiltro Then
                With xlSheet
                    .Cells(fila, 1).Value = correo.SentOn
                    .Cells(fila, 2).Value = correo.Subject
                    .Cells(fila, 3).Value = correo.To
                End With
                fila = fila + 1
            End If
        End If
    Next item

    ' --- Formato ---
    If fila > 2 Then
        xlSheet.Columns("A:C").AutoFit
        xlSheet.Range("A1:C" & fila - 1).Sort Key1:=xlSheet.Range("A2"), Order1:=1, header:=xlYes
    End If

    ' --- Nombre del archivo ---
 nombreArchivo = "Correos_Enviados_" & Format(mesFiltro, "00") & "_" & annoFiltro & ".xlsx"
rutaArchivo = carpetaDestino & nombreArchivo

    ' --- Guardar ---
    xlWB.SaveAs rutaArchivo
    xlWB.Close SaveChanges:=True
    xlApp.Quit

    MsgBox "Archivo generado correctamente en:" & vbCrLf & rutaArchivo, vbInformation, "Exportación completada"

    ' --- Limpieza ---
    Set xlSheet = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    Set carpetaEnviados = Nothing
    Set olNs = Nothing
    Set olApp = Nothing

End Sub

