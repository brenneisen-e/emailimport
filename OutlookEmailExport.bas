Attribute VB_Name = "OutlookEmailExport"
' ============================================================================
' OUTLOOK EMAIL JSON EXPORT
' ============================================================================
' Dieses Makro exportiert Emails aus Outlook als JSON-Datei
' für den Import ins HTA Email-Review Tool
'
' VERWENDUNG:
' 1. In Excel: ALT+F11 öffnen
' 2. Dieses Modul importieren (Datei → Datei importieren)
' 3. Button/Makro "ExportOutlookToJSON" ausführen
' ============================================================================

Option Explicit

' Hauptfunktion: Exportiert Outlook Emails als JSON
Public Sub ExportOutlookToJSON()
    On Error GoTo ErrorHandler

    Dim outlookApp As Object
    Dim namespace As Object
    Dim folder As Object
    Dim items As Object
    Dim item As Object
    Dim fso As Object
    Dim jsonFile As Object
    Dim jsonPath As String
    Dim jsonContent As String
    Dim emailCount As Long
    Dim i As Long
    Dim cutoffDate As Date
    Dim selectedFolder As Object

    ' Progress anzeigen
    Application.StatusBar = "Initialisiere Outlook..."

    ' Outlook verbinden
    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo ErrorHandler

    If outlookApp Is Nothing Then
        MsgBox "Fehler: Outlook konnte nicht geöffnet werden!", vbCritical
        Exit Sub
    End If

    Set namespace = outlookApp.GetNamespace("MAPI")

    ' Ordner-Auswahl
    Set selectedFolder = namespace.PickFolder
    If selectedFolder Is Nothing Then
        MsgBox "Kein Ordner ausgewählt. Export abgebrochen.", vbInformation
        Application.StatusBar = False
        Exit Sub
    End If

    Set folder = selectedFolder
    Set items = folder.items

    ' Sortiere nach Empfangsdatum
    items.Sort "[ReceivedTime]", True

    ' Datum: Letzte 7 Tage
    cutoffDate = DateAdd("d", -7, Date)

    Application.StatusBar = "Lade Emails aus Outlook..."

    ' JSON Header
    jsonContent = "[" & vbCrLf
    emailCount = 0

    ' Durchlaufe Emails (max 100)
    For i = 1 To items.Count
        If emailCount >= 100 Then Exit For

        Set item = items.item(i)

        ' Nur Emails
        If item.Class = 43 Then ' olMail
            ' Nur letzte 7 Tage
            If item.ReceivedTime >= cutoffDate Then
                ' Progress
                If emailCount Mod 10 = 0 Then
                    Application.StatusBar = "Verarbeite Email " & (emailCount + 1) & "..."
                End If

                ' Email zu JSON hinzufügen
                If emailCount > 0 Then jsonContent = jsonContent & "," & vbCrLf
                jsonContent = jsonContent & EmailToJSON(item)
                emailCount = emailCount + 1
            End If
        End If
    Next i

    ' JSON Footer
    jsonContent = jsonContent & vbCrLf & "]"

    If emailCount = 0 Then
        MsgBox "Keine Emails der letzten 7 Tage gefunden!", vbInformation
        Application.StatusBar = False
        Exit Sub
    End If

    ' Speichere JSON im Downloads-Ordner
    Application.StatusBar = "Speichere JSON..."

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Downloads-Pfad (fest)
    Dim downloadsPath As String
    downloadsPath = Environ("USERPROFILE") & "\Downloads"

    ' Falls Downloads nicht existiert, verwende Desktop
    If Not fso.FolderExists(downloadsPath) Then
        downloadsPath = Environ("USERPROFILE") & "\Desktop"
    End If

    jsonPath = downloadsPath & "\outlook_emails_" & _
               Format(Now, "yyyymmdd_hhnnss") & ".json"

    ' Schreibe Datei
    Set jsonFile = fso.CreateTextFile(jsonPath, True, True) ' Unicode
    jsonFile.Write jsonContent
    jsonFile.Close

    Application.StatusBar = False

    ' Öffne Downloads-Ordner
    On Error Resume Next
    CreateObject("Shell.Application").Explore downloadsPath
    On Error GoTo ErrorHandler

    ' Erfolgsmeldung
    MsgBox "✓ Export erfolgreich!" & vbCrLf & vbCrLf & _
           "Emails exportiert: " & emailCount & vbCrLf & _
           "Ordner: " & downloadsPath & vbCrLf & _
           "Datei: " & fso.GetFileName(jsonPath) & vbCrLf & vbCrLf & _
           "Der Downloads-Ordner wurde geöffnet." & vbCrLf & vbCrLf & _
           "Jetzt kannst du im HTA Tool auf 'JSON aus Downloads laden' klicken.", _
           vbInformation, "Outlook JSON Export"

    ' Cleanup
    Set item = Nothing
    Set items = Nothing
    Set folder = Nothing
    Set namespace = Nothing
    Set outlookApp = Nothing
    Set fso = Nothing
    Set jsonFile = Nothing

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "Fehler beim Export:" & vbCrLf & vbCrLf & _
           "Nummer: " & Err.Number & vbCrLf & _
           "Beschreibung: " & Err.Description, _
           vbCritical, "Export Fehler"
End Sub

' Konvertiert eine Email zu JSON
Private Function EmailToJSON(item As Object) As String
    Dim json As String
    Dim senderEmail As String
    Dim senderName As String
    Dim subject As String
    Dim body As String
    Dim receivedTime As String
    Dim attachments As String
    Dim i As Long

    On Error Resume Next

    ' Sender Email
    senderEmail = GetSenderEmail(item)

    ' Sender Name
    senderName = item.SenderName
    If senderName = "" Then senderName = "Unbekannt"

    ' Subject
    subject = item.subject
    If subject = "" Then subject = "(Kein Betreff)"

    ' Body (Plain Text bevorzugt, sonst HTML)
    body = item.body
    If body = "" Then
        body = item.HTMLBody
    End If

    ' Empfangsdatum
    receivedTime = Format(item.ReceivedTime, "yyyy-mm-dd hh:nn:ss")

    ' Anhänge
    attachments = ""
    If item.attachments.Count > 0 Then
        For i = 1 To item.attachments.Count
            If attachments <> "" Then attachments = attachments & ", "
            attachments = attachments & item.attachments.item(i).FileName
        Next i
    End If

    ' JSON erstellen
    json = "  {" & vbCrLf
    json = json & "    ""von_email"": """ & EscapeJSON(senderEmail) & """," & vbCrLf
    json = json & "    ""von_name"": """ & EscapeJSON(senderName) & """," & vbCrLf
    json = json & "    ""betreff"": """ & EscapeJSON(subject) & """," & vbCrLf
    json = json & "    ""text"": """ & EscapeJSON(body) & """," & vbCrLf
    json = json & "    ""datum"": """ & receivedTime & """," & vbCrLf
    json = json & "    ""anhänge"": """ & EscapeJSON(attachments) & """," & vbCrLf
    json = json & "    ""kategorie"": ""Incident""," & vbCrLf
    json = json & "    ""status"": ""Offen""," & vbCrLf
    json = json & "    ""cluster"": ""Allgemein""," & vbCrLf
    json = json & "    ""agentur"": ""Unbekannt""" & vbCrLf
    json = json & "  }"

    EmailToJSON = json
End Function

' Holt Sender Email-Adresse
Private Function GetSenderEmail(item As Object) As String
    Dim email As String

    On Error Resume Next

    ' Versuche verschiedene Methoden
    email = item.SenderEmailAddress

    ' Wenn Exchange-Format, versuche SMTP-Adresse
    If InStr(email, "/") > 0 Or InStr(email, "=") > 0 Then
        If Not item.Sender Is Nothing Then
            email = item.Sender.GetExchangeUser().PrimarySmtpAddress
        End If
    End If

    ' Fallback
    If email = "" Then
        email = "unknown@unknown.com"
    End If

    GetSenderEmail = email
End Function

' Escape JSON Strings (entferne problematische Zeichen)
Private Function EscapeJSON(text As String) As String
    Dim result As String
    result = text

    ' Ersetze Backslash
    result = Replace(result, "\", "\\")

    ' Ersetze Anführungszeichen
    result = Replace(result, """", "\""")

    ' Ersetze Steuerzeichen
    result = Replace(result, vbCr, "\r")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")

    ' Kürze auf max 10000 Zeichen
    If Len(result) > 10000 Then
        result = Left(result, 10000) & "... (gekürzt)"
    End If

    EscapeJSON = result
End Function
