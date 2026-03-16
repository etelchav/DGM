Attribute VB_Name = "EnviarMasivo"
Option Explicit

' ================== CONFIG ==================
Private Const STORE_NAME As String = "notificacionesrnm@minae.go.cr"
Private Const CONTACTS_NAME_ES As String = "Contactos"
Private Const CONTACTS_NAME_EN As String = "Contacts"
Private Const USERS_FOLDER_NAME As String = "Usuarios"

' Pausas anti-bloqueo (ajustable)
Private Const PAUSA_CADA As Long = 50
Private Const PAUSA_SEGUNDOS As Long = 15

' Archivos de control (se crean en Escritorio)
Private Function GetBaseFolder() As String
    GetBaseFolder = Environ$("USERPROFILE") & "\Desktop\DGM_Masivo"
End Function

Private Function GetLogCsvPath() As String
    GetLogCsvPath = GetBaseFolder() & "\Log_Envios.csv"
End Function

Private Function GetCheckpointPath() As String
    GetCheckpointPath = GetBaseFolder() & "\Checkpoint_ultimo_email.txt"
End Function
' ============================================


' ==========================================================
' MODO REAL: envía N correos (límite configurable) y continúa donde quedó
' ==========================================================
Public Sub EnviarMasivo_Usuarios_Enviar_ConLimite_Y_Log()

    ' 1) Correo base abierto
    If Application.ActiveInspector Is Nothing Then
        MsgBox "Abrí el correo base (asunto + cuerpo + adjuntos) y dejalo abierto.", vbExclamation
        Exit Sub
    End If

    Dim mailBase As Outlook.MailItem
    Set mailBase = Application.ActiveInspector.CurrentItem
    If mailBase Is Nothing Or mailBase.Class <> olMail Then
        MsgBox "El elemento activo no es un correo.", vbExclamation
        Exit Sub
    End If

    ' Forzar que Outlook “fije” el contenido del base
    On Error Resume Next
    mailBase.Save
    On Error GoTo 0

    Dim subj As String: subj = Trim$(mailBase.Subject)
    Dim htmlBase As String, textBase As String
    htmlBase = "": textBase = ""
    On Error Resume Next
    htmlBase = mailBase.HTMLBody
    textBase = mailBase.Body
    On Error GoTo 0

    If Len(subj) = 0 Then
        MsgBox "El correo base no tiene ASUNTO. Poné el asunto y volvé a correr.", vbExclamation
        Exit Sub
    End If
    If Len(Trim$(htmlBase)) = 0 And Len(Trim$(textBase)) = 0 Then
        MsgBox "El correo base no tiene CUERPO. Escribí el contenido y volvé a correr.", vbExclamation
        Exit Sub
    End If

    ' 2) Pedir límite
    Dim limiteStr As String, limite As Long
    limiteStr = InputBox("¿Cuántos correos querés enviar en esta ejecución?" & vbCrLf & _
                         "Ejemplos: 50, 100, 200", "Límite de envío", "50")
    limiteStr = Trim$(limiteStr)
    If Len(limiteStr) = 0 Then Exit Sub
    If Not IsNumeric(limiteStr) Then
        MsgBox "El límite debe ser un número.", vbExclamation
        Exit Sub
    End If
    limite = CLng(limiteStr)
    If limite <= 0 Then
        MsgBox "El límite debe ser mayor que 0.", vbExclamation
        Exit Sub
    End If

    ' 3) Carpeta Usuarios
    Dim usuarios As Outlook.Folder
    Set usuarios = GetUsuariosFolder()
    If usuarios Is Nothing Then Exit Sub

    ' 4) Construir lista de emails (ordenada) desde ContactItem
    Dim emails() As String
    emails = GetSortedEmailsFromFolder(usuarios)

    If (Not Not emails) = 0 Then
        MsgBox "No encontré contactos con email en: " & vbCrLf & usuarios.folderPath, vbExclamation
        Exit Sub
    End If

    Dim total As Long
    total = UBound(emails) - LBound(emails) + 1

    ' 5) Leer checkpoint (último email procesado)
    Dim lastEmail As String
    lastEmail = ReadCheckpoint()

    ' 6) Confirmación
    Dim msg As String
    msg = "Store: " & STORE_NAME & vbCrLf & _
          "Carpeta: " & usuarios.folderPath & vbCrLf & _
          "Total emails detectados: " & total & vbCrLf & _
          "Límite esta ejecución: " & limite & vbCrLf & vbCrLf & _
          "Continuar desde (checkpoint): " & IIf(Len(lastEmail) = 0, "(vacío = inicio)", lastEmail) & vbCrLf & vbCrLf & _
          "Se enviará 1 correo por destinatario (Para). ¿Continuar?"

    If MsgBox(msg, vbQuestion + vbYesNo, "Confirmar envío") <> vbYes Then Exit Sub

    ' 7) Preparar log y carpeta
    EnsureFolderExists GetBaseFolder()
    EnsureLogHasHeader GetLogCsvPath()

    ' 8) Enviar (continuando donde quedó)
    Dim enviados As Long, intentados As Long
    enviados = 0
    intentados = 0

    Dim i As Long
    For i = LBound(emails) To UBound(emails)

        Dim emailTo As String
        emailTo = emails(i)

        ' Saltar hasta pasar el checkpoint
        If Len(lastEmail) > 0 Then
            If StrComp(emailTo, lastEmail, vbTextCompare) <= 0 Then
                GoTo Siguiente
            End If
        End If

        If enviados >= limite Then Exit For

        intentados = intentados + 1

        On Error GoTo EH

        Dim m As Outlook.MailItem
        Set m = mailBase.Copy
        m.Display  ' fuerza inicialización del cuerpo

        With m
            .To = emailTo
            .CC = ""
            .BCC = ""
            .Subject = subj
            If Len(Trim$(htmlBase)) > 0 Then
                .HTMLBody = htmlBase
            Else
                .Body = textBase
            End If
            .Send
        End With

        enviados = enviados + 1

        ' Log OK
        AppendLog GetLogCsvPath(), Now, emailTo, subj, "ENVIADO", ""

        ' Guardar checkpoint (último enviado)
        WriteCheckpoint emailTo

        ' Pausas anti-bloqueo
        If enviados Mod PAUSA_CADA = 0 Then
            EsperarSegundos PAUSA_SEGUNDOS
        End If

        GoTo Siguiente

EH:
        ' Log ERROR y continuar
        AppendLog GetLogCsvPath(), Now, emailTo, subj, "ERROR", Replace(Err.Description, vbCrLf, " ")
        Err.Clear
        ' Importante: NO avanzamos checkpoint en error
        On Error GoTo 0

Siguiente:
        DoEvents
    Next i

    MsgBox "Listo." & vbCrLf & _
           "Enviados en esta ejecución: " & enviados & vbCrLf & _
           "Intentados (pasando checkpoint): " & intentados & vbCrLf & vbCrLf & _
           "Log: " & GetLogCsvPath() & vbCrLf & _
           "Checkpoint: " & GetCheckpointPath(), vbInformation

End Sub


' ==========================================================
' Diagnóstico rápido: contar emails en Usuarios (ordenados)
' ==========================================================
Public Sub Diagnostico_ListarPrimeros20Emails()
    Dim usuarios As Outlook.Folder
    Set usuarios = GetUsuariosFolder()
    If usuarios Is Nothing Then Exit Sub

    Dim emails() As String
    emails = GetSortedEmailsFromFolder(usuarios)

    Dim s As String, i As Long, n As Long
    s = "Primeros emails (ordenados):" & vbCrLf
    If (Not Not emails) = 0 Then
        MsgBox "No hay emails.", vbExclamation
        Exit Sub
    End If

    n = WorksheetFunctionMin(20, UBound(emails) - LBound(emails) + 1)
    For i = LBound(emails) To LBound(emails) + n - 1
        s = s & " - " & emails(i) & vbCrLf
    Next i

    MsgBox s, vbInformation
End Sub


' ================== HELPERS OUTLOOK ==================

Private Function GetUsuariosFolder() As Outlook.Folder

    Dim ns As Outlook.NameSpace
    Set ns = Application.Session

    Dim storeRoot As Outlook.Folder
    On Error Resume Next
    Set storeRoot = ns.Folders.item(STORE_NAME)
    On Error GoTo 0

    If storeRoot Is Nothing Then
        MsgBox "No se encontró el store: " & STORE_NAME, vbCritical
        Set GetUsuariosFolder = Nothing
        Exit Function
    End If

    Dim contactos As Outlook.Folder
    Set contactos = Nothing

    On Error Resume Next
    Set contactos = storeRoot.Folders.item(CONTACTS_NAME_ES)
    On Error GoTo 0

    If contactos Is Nothing Then
        On Error Resume Next
        Set contactos = storeRoot.Folders.item(CONTACTS_NAME_EN)
        On Error GoTo 0
    End If

    If contactos Is Nothing Then
        MsgBox "No se pudo abrir 'Contactos' (o 'Contacts') dentro del store " & STORE_NAME & ".", vbCritical
        Set GetUsuariosFolder = Nothing
        Exit Function
    End If

    Dim usuarios As Outlook.Folder
    Set usuarios = Nothing

    On Error Resume Next
    Set usuarios = contactos.Folders.item(USERS_FOLDER_NAME)
    On Error GoTo 0

    If usuarios Is Nothing Then
        MsgBox "No se encontró la subcarpeta '" & USERS_FOLDER_NAME & "' dentro de:" & vbCrLf & contactos.folderPath, vbCritical
        Set GetUsuariosFolder = Nothing
        Exit Function
    End If

    Set GetUsuariosFolder = usuarios
End Function

Private Function GetSortedEmailsFromFolder(ByVal f As Outlook.Folder) As String()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1 ' TextCompare

    Dim it As Object
    For Each it In f.items
        If TypeName(it) = "ContactItem" Then
            Dim c As Outlook.ContactItem
            Set c = it

            Dim e As String
            e = Trim$(c.Email1Address)
            If Len(e) = 0 Then e = Trim$(c.Email2Address)
            If Len(e) = 0 Then e = Trim$(c.Email3Address)

            If Len(e) > 0 Then
                If Not dict.Exists(LCase$(e)) Then dict.Add LCase$(e), e
            End If
        End If
    Next it

    If dict.Count = 0 Then
        Dim emptyArr() As String
        GetSortedEmailsFromFolder = emptyArr
        Exit Function
    End If

    Dim arr() As String
    ReDim arr(0 To dict.Count - 1)

    Dim k As Variant, idx As Long
    idx = 0
    For Each k In dict.keys
        arr(idx) = dict.item(k)
        idx = idx + 1
    Next k

    QuickSortStrings arr, LBound(arr), UBound(arr)

    GetSortedEmailsFromFolder = arr
End Function


' ================== SORT ==================
Private Sub QuickSortStrings(ByRef arr() As String, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long
    Dim pivot As String, temp As String
    i = first
    j = last
    pivot = arr((first + last) \ 2)

    Do While i <= j
        Do While StrComp(arr(i), pivot, vbTextCompare) < 0
            i = i + 1
        Loop
        Do While StrComp(arr(j), pivot, vbTextCompare) > 0
            j = j - 1
        Loop
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop

    If first < j Then QuickSortStrings arr, first, j
    If i < last Then QuickSortStrings arr, i, last
End Sub


' ================== LOG + CHECKPOINT ==================
Private Sub EnsureLogHasHeader(ByVal path As String)
    If FileExists(path) Then Exit Sub
    Dim header As String
    header = "FechaHora,Email,Asunto,Estado,Detalle" & vbCrLf
    WriteTextFile path, header
End Sub

Private Sub AppendLog(ByVal path As String, ByVal dt As Date, ByVal email As String, ByVal subj As String, ByVal estado As String, ByVal detalle As String)
    Dim line As String
    line = CsvEscape(Format$(dt, "yyyy-mm-dd hh:nn:ss")) & "," & _
           CsvEscape(email) & "," & _
           CsvEscape(subj) & "," & _
           CsvEscape(estado) & "," & _
           CsvEscape(detalle) & vbCrLf
    AppendTextFile path, line
End Sub

Private Function ReadCheckpoint() As String
    If Not FileExists(GetCheckpointPath()) Then
        ReadCheckpoint = ""
        Exit Function
    End If
    ReadCheckpoint = Trim$(ReadTextFile(GetCheckpointPath()))
End Function

Private Sub WriteCheckpoint(ByVal lastEmail As String)
    EnsureFolderExists GetBaseFolder()
    WriteTextFile GetCheckpointPath(), Trim$(lastEmail)
End Sub


' ================== FILE UTILS ==================
Private Sub EnsureFolderExists(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
End Sub

Private Function FileExists(ByVal path As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(path)
End Function

Private Sub WriteTextFile(ByVal path As String, ByVal content As String)
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(path, True, True) ' overwrite, Unicode
    ts.Write content
    ts.Close
End Sub

Private Sub AppendTextFile(ByVal path As String, ByVal content As String)
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(path, 8, True, True) ' ForAppending=8, create, Unicode
    ts.Write content
    ts.Close
End Sub

Private Function ReadTextFile(ByVal path As String) As String
    On Error GoTo EH

    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Si no existe, devolver vacío
    If Not fso.FileExists(path) Then
        ReadTextFile = ""
        Exit Function
    End If

    Set ts = fso.OpenTextFile(path, 1, False, True) ' ForReading=1, Unicode

    ' Si está vacío, no leer
    If ts.AtEndOfStream Then
        ReadTextFile = ""
        ts.Close
        Exit Function
    End If

    ReadTextFile = ts.ReadAll
    ts.Close
    Exit Function

EH:
    ' Si da cualquier error, devolver vacío y seguir
    ReadTextFile = ""
End Function


Private Function CsvEscape(ByVal s As String) As String
    s = Replace(s, """", """""")
    If InStr(s, ",") > 0 Or InStr(s, vbCr) > 0 Or InStr(s, vbLf) > 0 Then
        CsvEscape = """" & s & """"
    Else
        CsvEscape = s
    End If
End Function


' ================== MISC ==================
Private Sub EsperarSegundos(ByVal segundos As Long)
    Dim t As Double
    t = Timer + segundos
    Do While Timer < t
        DoEvents
    Loop
End Sub

Private Function WorksheetFunctionMin(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then WorksheetFunctionMin = a Else WorksheetFunctionMin = b
End Function
Public Sub ResetearCheckpoint_EnvioMasivo()
    Dim baseFolder As String
    baseFolder = Environ$("USERPROFILE") & "\Desktop\DGM_Masivo"

    Dim checkpointPath As String
    checkpointPath = baseFolder & "\Checkpoint_ultimo_email.txt"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Crear carpeta si no existe
    If Not fso.FolderExists(baseFolder) Then
        fso.CreateFolder baseFolder
    End If

    ' Borrar checkpoint si existe
    If fso.FileExists(checkpointPath) Then
        fso.DeleteFile checkpointPath, True
    End If

    MsgBox "Checkpoint reiniciado." & vbCrLf & _
           "La próxima ejecución comenzará desde el inicio.", vbInformation
End Sub

Public Sub EnviarMasivo_Usuarios_Confirmar_Reset_o_Continuar()

    Dim lastEmail As String
    lastEmail = ReadCheckpoint()

    Dim msg As String
    msg = "ENVÍO MASIVO DGM" & vbCrLf & vbCrLf & _
          "YES = CONTINUAR desde el último enviado (checkpoint)" & vbCrLf & _
          "NO  = RESETEAR (campaña nueva desde cero)" & vbCrLf & _
          "CANCEL = Cancelar" & vbCrLf & vbCrLf & _
          "Checkpoint actual:" & vbCrLf & GetCheckpointPath() & vbCrLf & _
          "Último email guardado:" & vbCrLf & _
          IIf(Len(lastEmail) = 0, "(vacío = inicio)", lastEmail)

    Dim r As VbMsgBoxResult
    r = MsgBox(msg, vbQuestion + vbYesNoCancel, "Confirmar envío masivo")

    If r = vbCancel Then Exit Sub

    If r = vbNo Then
        ' Resetear campaña
        ResetearCheckpoint_EnvioMasivo
    End If

    ' Ejecutar envío normal (el tuyo, sin cambios)
    EnviarMasivo_Usuarios_Enviar_ConLimite_Y_Log

End Sub

