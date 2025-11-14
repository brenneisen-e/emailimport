' ============================================================================
' OutlookExportMacro.vba
' VBA-Makro für Outlook zum Exportieren ausgewählter Emails als JSON
' ============================================================================
'
' INSTALLATION:
' 1. In Outlook: ALT + F11 drücken (öffnet VBA-Editor)
' 2. Im VBA-Editor: Einfügen → Modul
' 3. Diesen Code in das neue Modul kopieren
' 4. Speichern und VBA-Editor schließen
'
' VERWENDUNG:
' 1. In Outlook eine oder mehrere Emails auswählen
' 2. ALT + F8 drücken
' 3. Makro "ExportSelectedEmailsToJSON" auswählen und ausführen
'
' Optional: Button in Outlook-Ribbon hinzufügen für schnellen Zugriff
' ============================================================================

Option Explicit

' Hauptfunktion zum Exportieren ausgewählter Emails
Public Sub ExportSelectedEmailsToJSON()
    Dim objExplorer As Outlook.Explorer
    Dim objSelection As Outlook.Selection
    Dim objMail As Outlook.MailItem
    Dim objFSO As Object
    Dim objFile As Object
    Dim strJSON As String
    Dim strFilePath As String
    Dim i As Integer
    Dim emailCount As Integer

    On Error GoTo ErrorHandler

    ' Outlook Explorer und Selection abrufen
    Set objExplorer = Application.ActiveExplorer
    Set objSelection = objExplorer.Selection

    ' Prüfen ob Emails ausgewählt sind
    If objSelection.Count = 0 Then
        MsgBox "Bitte wähle mindestens eine Email aus!", vbExclamation, "Keine Emails ausgewählt"
        Exit Sub
    End If

    ' Nur Mail-Items zählen
    emailCount = 0
    For i = 1 To objSelection.Count
        If TypeOf objSelection.Item(i) Is Outlook.MailItem Then
            emailCount = emailCount + 1
        End If
    Next i

    If emailCount = 0 Then
        MsgBox "Die Auswahl enthält keine Emails!", vbExclamation, "Keine Emails gefunden"
        Exit Sub
    End If

    ' Dateiname vorschlagen
    Dim defaultFileName As String
    defaultFileName = "hypercare_emails_" & Format(Now, "yyyymmdd_hhnnss") & ".json"

    ' Speichern-Dialog anzeigen
    strFilePath = GetSaveFilePath(defaultFileName)
    If strFilePath = "" Then
        Exit Sub ' Benutzer hat abgebrochen
    End If

    ' JSON aufbauen
    strJSON = "[" & vbCrLf

    Dim firstItem As Boolean
    firstItem = True

    ' Fortschrittsanzeige
    Application.StatusBar = "Exportiere Emails..."

    For i = 1 To objSelection.Count
        If TypeOf objSelection.Item(i) Is Outlook.MailItem Then
            Set objMail = objSelection.Item(i)

            If Not firstItem Then
                strJSON = strJSON & "," & vbCrLf
            End If
            firstItem = False

            strJSON = strJSON & "  {" & vbCrLf
            strJSON = strJSON & "    ""datum"": """ & FormatDateTime(objMail.ReceivedTime, vbGeneralDate) & """," & vbCrLf
            strJSON = strJSON & "    ""von_email"": """ & EscapeJSON(GetSenderEmail(objMail)) & """," & vbCrLf
            strJSON = strJSON & "    ""von_name"": """ & EscapeJSON(objMail.SenderName) & """," & vbCrLf
            strJSON = strJSON & "    ""betreff"": """ & EscapeJSON(objMail.Subject) & """," & vbCrLf
            strJSON = strJSON & "    ""text"": """ & EscapeJSON(GetEmailBody(objMail)) & """," & vbCrLf
            strJSON = strJSON & "    ""anhänge"": " & GetAttachmentsJSON(objMail) & vbCrLf
            strJSON = strJSON & "  }"

            ' Fortschritt aktualisieren
            Application.StatusBar = "Exportiere Email " & i & " von " & emailCount & "..."
        End If
    Next i

    strJSON = strJSON & vbCrLf & "]"

    ' JSON in Datei schreiben
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.CreateTextFile(strFilePath, True, True) ' True = überschreiben, True = Unicode
    objFile.Write strJSON
    objFile.Close

    ' Erfolgsmeldung
    Application.StatusBar = False
    MsgBox "Erfolgreich " & emailCount & " Email(s) exportiert!" & vbCrLf & vbCrLf & _
           "Datei: " & strFilePath & vbCrLf & vbCrLf & _
           "Du kannst die JSON-Datei jetzt in die BGAV Hypercare Email Review App laden.", _
           vbInformation, "Export erfolgreich"

    ' Ordner öffnen
    If MsgBox("Ordner im Explorer öffnen?", vbYesNo + vbQuestion, "Ordner öffnen") = vbYes Then
        Shell "explorer.exe /select," & Chr(34) & strFilePath & Chr(34), vbNormalFocus
    End If

    ' Aufräumen
    Set objFile = Nothing
    Set objFSO = Nothing
    Set objMail = Nothing
    Set objSelection = Nothing
    Set objExplorer = Nothing

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "Fehler beim Export:" & vbCrLf & vbCrLf & _
           "Fehler-Nr: " & Err.Number & vbCrLf & _
           "Beschreibung: " & Err.Description, _
           vbCritical, "Export-Fehler"
End Sub

' Hilfsfunktion: Speichern-Dialog
Private Function GetSaveFilePath(defaultFileName As String) As String
    Dim objDialog As Object
    Dim strPath As String

    On Error Resume Next

    ' FileDialog verwenden (Outlook 2010+)
    Set objDialog = Application.CreateObject("UserForm1")

    ' Fallback: InputBox verwenden
    strPath = Environ("USERPROFILE") & "\Downloads\" & defaultFileName
    strPath = InputBox("Speicherort für JSON-Export:" & vbCrLf & _
                      "(Vollständiger Pfad mit Dateiname)", _
                      "Emails exportieren", _
                      strPath)

    GetSaveFilePath = strPath
    Set objDialog = Nothing
End Function

' Hilfsfunktion: SMTP-Email-Adresse des Absenders ermitteln
Private Function GetSenderEmail(objMail As Outlook.MailItem) As String
    Dim strEmail As String

    On Error Resume Next

    ' Versuche SMTP-Adresse zu ermitteln
    If objMail.SenderEmailType = "EX" Then
        ' Exchange-Adresse: Versuche SMTP-Adresse zu bekommen
        Dim objSender As Outlook.AddressEntry
        Set objSender = objMail.Sender

        If Not objSender Is Nothing Then
            If objSender.AddressEntryUserType = olExchangeUserAddressEntry Or _
               objSender.AddressEntryUserType = olExchangeRemoteUserAddressEntry Then
                Dim objExUser As Outlook.ExchangeUser
                Set objExUser = objSender.GetExchangeUser
                If Not objExUser Is Nothing Then
                    strEmail = objExUser.PrimarySmtpAddress
                End If
            End If
        End If
    End If

    ' Fallback: Standard-Email-Adresse
    If strEmail = "" Then
        strEmail = objMail.SenderEmailAddress
    End If

    ' Wenn immer noch leer, Platzhalter verwenden
    If strEmail = "" Then
        strEmail = "unknown@unknown.com"
    End If

    GetSenderEmail = strEmail
End Function

' Hilfsfunktion: Email-Body extrahieren (Plain Text bevorzugt)
Private Function GetEmailBody(objMail As Outlook.MailItem) As String
    Dim strBody As String

    On Error Resume Next

    ' Plain Text bevorzugen
    If Len(objMail.Body) > 0 Then
        strBody = objMail.Body
    Else
        ' Fallback: HTML-Body (ohne Tags)
        strBody = StripHTML(objMail.HTMLBody)
    End If

    ' Auf maximale Länge begrenzen (optional)
    If Len(strBody) > 50000 Then
        strBody = Left(strBody, 50000) & "... (gekürzt)"
    End If

    GetEmailBody = strBody
End Function

' Hilfsfunktion: HTML-Tags entfernen
Private Function StripHTML(strHTML As String) As String
    Dim objRegEx As Object
    Dim strText As String

    On Error Resume Next

    strText = strHTML

    ' Einfache Regex-basierte Bereinigung
    Set objRegEx = CreateObject("VBScript.RegExp")
    objRegEx.Global = True
    objRegEx.IgnoreCase = True

    ' BR-Tags durch Zeilenumbrüche ersetzen
    objRegEx.Pattern = "<br\s*/?>"
    strText = objRegEx.Replace(strText, vbCrLf)

    ' Alle HTML-Tags entfernen
    objRegEx.Pattern = "<[^>]+>"
    strText = objRegEx.Replace(strText, "")

    ' HTML-Entities dekodieren
    strText = Replace(strText, "&nbsp;", " ")
    strText = Replace(strText, "&amp;", "&")
    strText = Replace(strText, "&lt;", "<")
    strText = Replace(strText, "&gt;", ">")
    strText = Replace(strText, "&quot;", """")

    StripHTML = Trim(strText)
    Set objRegEx = Nothing
End Function

' Hilfsfunktion: Anhänge als JSON-Array
Private Function GetAttachmentsJSON(objMail As Outlook.MailItem) As String
    Dim strJSON As String
    Dim objAttachment As Outlook.Attachment
    Dim firstAttachment As Boolean

    strJSON = "["
    firstAttachment = True

    For Each objAttachment In objMail.Attachments
        If Not firstAttachment Then
            strJSON = strJSON & ", "
        End If
        firstAttachment = False

        strJSON = strJSON & """" & EscapeJSON(objAttachment.FileName) & """"
    Next objAttachment

    strJSON = strJSON & "]"
    GetAttachmentsJSON = strJSON
End Function

' Hilfsfunktion: JSON-Escape für Strings
Private Function EscapeJSON(strText As String) As String
    Dim strResult As String

    strResult = strText
    strResult = Replace(strResult, "\", "\\")
    strResult = Replace(strResult, """", "\""")
    strResult = Replace(strResult, vbCr, "\r")
    strResult = Replace(strResult, vbLf, "\n")
    strResult = Replace(strResult, vbTab, "\t")

    EscapeJSON = strResult
End Function

' Alternative Funktion: Exportiere ganzen Ordner
Public Sub ExportFolderToJSON()
    Dim objFolder As Outlook.Folder
    Dim objItem As Object
    Dim strJSON As String
    Dim strFilePath As String
    Dim emailCount As Integer
    Dim firstItem As Boolean

    On Error GoTo ErrorHandler

    ' Aktuellen Ordner abrufen
    Set objFolder = Application.ActiveExplorer.CurrentFolder

    If objFolder.Items.Count = 0 Then
        MsgBox "Der Ordner enthält keine Emails!", vbExclamation, "Leerer Ordner"
        Exit Sub
    End If

    ' Bestätigung
    If MsgBox("Möchtest du alle " & objFolder.Items.Count & " Elemente aus dem Ordner '" & _
              objFolder.Name & "' exportieren?", vbYesNo + vbQuestion, "Ordner exportieren") = vbNo Then
        Exit Sub
    End If

    ' Dateiname vorschlagen
    Dim defaultFileName As String
    defaultFileName = "hypercare_" & objFolder.Name & "_" & Format(Now, "yyyymmdd_hhnnss") & ".json"

    strFilePath = GetSaveFilePath(defaultFileName)
    If strFilePath = "" Then Exit Sub

    ' JSON aufbauen
    strJSON = "[" & vbCrLf
    firstItem = True
    emailCount = 0

    For Each objItem In objFolder.Items
        If TypeOf objItem Is Outlook.MailItem Then
            Dim objMail As Outlook.MailItem
            Set objMail = objItem

            emailCount = emailCount + 1
            Application.StatusBar = "Exportiere Email " & emailCount & "..."

            If Not firstItem Then
                strJSON = strJSON & "," & vbCrLf
            End If
            firstItem = False

            strJSON = strJSON & "  {" & vbCrLf
            strJSON = strJSON & "    ""datum"": """ & FormatDateTime(objMail.ReceivedTime, vbGeneralDate) & """," & vbCrLf
            strJSON = strJSON & "    ""von_email"": """ & EscapeJSON(GetSenderEmail(objMail)) & """," & vbCrLf
            strJSON = strJSON & "    ""von_name"": """ & EscapeJSON(objMail.SenderName) & """," & vbCrLf
            strJSON = strJSON & "    ""betreff"": """ & EscapeJSON(objMail.Subject) & """," & vbCrLf
            strJSON = strJSON & "    ""text"": """ & EscapeJSON(GetEmailBody(objMail)) & """," & vbCrLf
            strJSON = strJSON & "    ""anhänge"": " & GetAttachmentsJSON(objMail) & vbCrLf
            strJSON = strJSON & "  }"
        End If
    Next objItem

    strJSON = strJSON & vbCrLf & "]"

    ' Datei schreiben
    Dim objFSO As Object
    Dim objFile As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.CreateTextFile(strFilePath, True, True)
    objFile.Write strJSON
    objFile.Close

    Application.StatusBar = False
    MsgBox "Erfolgreich " & emailCount & " Email(s) exportiert!" & vbCrLf & vbCrLf & _
           "Datei: " & strFilePath, vbInformation, "Export erfolgreich"

    Set objFile = Nothing
    Set objFSO = Nothing
    Set objFolder = Nothing
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "Fehler: " & Err.Description, vbCritical, "Export-Fehler"
End Sub
