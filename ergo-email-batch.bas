Attribute VB_Name = "ErgoEmailBatch"
' ============================================================================
' ERGO EMAIL BATCH - Excel VBA Tool
' ============================================================================
' Liest Emails ein und schreibt sie in ein Excel-Tabellenblatt.
'
' MODUS 1 (Outlook live):
'   - Outlook-Ordner ueber PickFolder-Dialog waehlen
'   - "Letzte X Tage" angeben
'   - Alle Mails im Zeitraum werden gelesen
'
' MODUS 2 (Ordner-Scan):
'   - Beliebigen Ordner waehlen (Laufwerk / Netzlaufwerk)
'   - Alle .msg und .eml Dateien werden gelesen
'   - Keine Outlook-MAPI noetig (nur fuer .msg)
'
' OUTPUT:
'   - Tabellenblatt "Emails" mit Spalten:
'     Nr | Datum | Absender Name | Absender Email | Betreff | Text | Anhaenge
'   - "JSON exportieren" erzeugt Beladungsdatei fuer Web-Tool (index.html)
'
' INSTALLATION:
'   1. Excel oeffnen, leere Mappe -> "Speichern unter" -> Typ ".xlsm"
'   2. Alt+F11 (VBA-Editor) -> Datei -> Datei importieren -> ergo-email-batch.bas
'   3. Alt+F8 -> Hauptmenu auswaehlen -> Ausfuehren
' ============================================================================

Option Explicit

Public Const SHEET_DATA As String = "Emails"
Public Const COL_COUNT As Long = 7

' Email-Klasse-IDs in Outlook
Private Const OL_MAIL_CLASS As Long = 43

' === HAUPTMENUE =============================================================
Public Sub Hauptmenu()
    Dim antwort As VbMsgBoxResult
    antwort = MsgBox( _
        "ERGO Email Batch - Modus waehlen:" & vbCrLf & vbCrLf & _
        "JA      = Outlook live (Postfach-Ordner + Tage)" & vbCrLf & _
        "NEIN    = Ordner mit .msg / .eml Dateien scannen" & vbCrLf & _
        "ABBRECH = Beenden", _
        vbYesNoCancel + vbQuestion, "ERGO Email Batch")

    If antwort = vbYes Then
        ModusOutlook
    ElseIf antwort = vbNo Then
        ModusOrdner
    End If
End Sub

' === SHORTCUT-MAKROS (Alt+F8 sichtbar) ======================================
Public Sub Outlook_Live_Abrufen()
    ModusOutlook
End Sub

Public Sub Ordner_Scannen()
    ModusOrdner
End Sub

Public Sub JSON_Exportieren()
    ExportiereJson
End Sub

' === MODUS 1: OUTLOOK LIVE ==================================================
Public Sub ModusOutlook()
    Dim olApp As Object, ns As Object, folder As Object
    On Error GoTo Fehler

    ' Outlook holen oder starten
    Set olApp = GetOutlook()
    If olApp Is Nothing Then
        MsgBox "Outlook konnte nicht gestartet werden.", vbCritical
        Exit Sub
    End If
    Set ns = olApp.GetNamespace("MAPI")

    ' Ordner via PickFolder-Dialog (zeigt alle Postfaecher mit Baum-Auswahl)
    Set folder = ns.PickFolder()
    If folder Is Nothing Then Exit Sub

    ' Tage zurueck abfragen
    Dim eingabe As String
    eingabe = InputBox( _
        "Wie viele Tage zurueck?" & vbCrLf & vbCrLf & _
        "Beispiel: 7 = letzte Woche, 30 = letzter Monat", _
        "Zeitraum", "7")
    If Trim(eingabe) = "" Then Exit Sub

    Dim daysBack As Long
    daysBack = CLng(Val(eingabe))
    If daysBack <= 0 Then
        MsgBox "Bitte eine positive Zahl eingeben.", vbExclamation
        Exit Sub
    End If

    Dim cutoff As Date
    cutoff = DateAdd("d", -daysBack, Now)

    ' Items holen + sortieren
    Application.ScreenUpdating = False
    Application.StatusBar = "Lade Emails aus '" & folder.Name & "'..."

    Dim items As Object
    Set items = folder.Items
    items.Sort "[ReceivedTime]", True ' descending

    BlattVorbereiten

    Dim row As Long: row = 2
    Dim verarbeitet As Long: verarbeitet = 0
    Dim it As Object
    Dim i As Long
    For i = 1 To items.Count
        On Error Resume Next
        Set it = items.Item(i)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo Fehler
            GoTo NextItem
        End If
        On Error GoTo Fehler

        If it.Class = OL_MAIL_CLASS Then
            ' Sortierung desc -> sobald aelter -> abbrechen
            If it.ReceivedTime < cutoff Then Exit For

            ZeileSchreiben row, _
                FormatDatumDE(it.ReceivedTime), _
                IsoDatum(it.ReceivedTime), _
                CStr(it.SenderName), _
                AbsenderEmail(it), _
                CStr(it.Subject), _
                BodyBereinigen(CStr(it.Body)), _
                AnhangNamen(it)
            row = row + 1
            verarbeitet = verarbeitet + 1
            If verarbeitet Mod 10 = 0 Then
                Application.StatusBar = "Verarbeitet: " & verarbeitet & " Mails..."
                DoEvents
            End If
        End If
NextItem:
    Next i

    BlattAbschliessen row - 1
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox verarbeitet & " Emails geladen aus '" & folder.Name & "'.", vbInformation, "Fertig"
    Exit Sub

Fehler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Fehler: " & Err.Description, vbCritical
End Sub

' === MODUS 2: ORDNER-SCAN ===================================================
Public Sub ModusOrdner()
    Dim folderPath As String
    On Error GoTo Fehler

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Ordner mit .msg / .eml Dateien waehlen"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        MsgBox "Ordner nicht gefunden: " & folderPath, vbCritical
        Exit Sub
    End If

    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)

    ' Outlook nur lazy laden falls .msg-Dateien gefunden werden
    Dim olApp As Object
    Set olApp = Nothing

    BlattVorbereiten

    Application.ScreenUpdating = False
    Application.StatusBar = "Scanne Ordner..."

    Dim row As Long: row = 2
    Dim verarbeitet As Long: verarbeitet = 0
    Dim file As Object
    Dim ext As String
    For Each file In folder.Files
        ext = LCase(fso.GetExtensionName(file.Name))
        If ext = "msg" Then
            If olApp Is Nothing Then Set olApp = GetOutlook()
            If Not olApp Is Nothing Then
                If MsgDateiLesen(olApp, file.Path, row) Then
                    row = row + 1
                    verarbeitet = verarbeitet + 1
                End If
            End If
        ElseIf ext = "eml" Then
            If EmlDateiLesen(fso, file.Path, row) Then
                row = row + 1
                verarbeitet = verarbeitet + 1
            End If
        End If

        If verarbeitet Mod 10 = 0 And verarbeitet > 0 Then
            Application.StatusBar = "Verarbeitet: " & verarbeitet & " Dateien..."
            DoEvents
        End If
    Next file

    BlattAbschliessen row - 1
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox verarbeitet & " Email-Dateien verarbeitet aus '" & folderPath & "'.", _
           vbInformation, "Fertig"
    Exit Sub

Fehler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Fehler: " & Err.Description, vbCritical
End Sub

' === MSG-Datei via Outlook ==================================================
Private Function MsgDateiLesen(olApp As Object, pfad As String, row As Long) As Boolean
    On Error GoTo Fehler
    Dim it As Object
    Set it = olApp.Session.OpenSharedItem(pfad)

    ZeileSchreiben row, _
        FormatDatumDE(it.ReceivedTime), _
        IsoDatum(it.ReceivedTime), _
        CStr(it.SenderName), _
        AbsenderEmail(it), _
        CStr(it.Subject), _
        BodyBereinigen(CStr(it.Body)), _
        AnhangNamen(it)

    On Error Resume Next
    it.Close 1 ' olDiscard
    On Error GoTo 0

    MsgDateiLesen = True
    Exit Function
Fehler:
    MsgDateiLesen = False
End Function

' === EML-Datei (RFC822 Text) ================================================
Private Function EmlDateiLesen(fso As Object, pfad As String, row As Long) As Boolean
    On Error GoTo Fehler

    Dim ts As Object
    Set ts = fso.OpenTextFile(pfad, 1, False, 0) ' ForReading, ASCII

    ' Header sammeln (bis zur ersten Leerzeile)
    Dim headers As Object
    Set headers = CreateObject("Scripting.Dictionary")
    Dim aktuellKey As String: aktuellKey = ""
    Dim aktuellVal As String: aktuellVal = ""
    Dim line As String
    Dim headerEnde As Boolean: headerEnde = False

    Do While Not ts.AtEndOfStream And Not headerEnde
        line = ts.ReadLine()
        If Trim(line) = "" Then
            If aktuellKey <> "" Then headers(aktuellKey) = aktuellVal
            headerEnde = True
        ElseIf Left(line, 1) = " " Or Left(line, 1) = vbTab Then
            ' Folding-Zeile (Fortsetzung des vorherigen Headers)
            aktuellVal = aktuellVal & " " & LTrim(line)
        Else
            ' Neuer Header
            If aktuellKey <> "" Then headers(aktuellKey) = aktuellVal
            Dim posDp As Long: posDp = InStr(line, ":")
            If posDp > 0 Then
                aktuellKey = LCase(Trim(Left(line, posDp - 1)))
                aktuellVal = Trim(Mid(line, posDp + 1))
            End If
        End If
    Loop

    ' Body lesen (Rest der Datei)
    Dim body As String: body = ""
    Do While Not ts.AtEndOfStream
        body = body & ts.ReadLine() & vbCrLf
    Loop
    ts.Close

    ' Anhang-Namen aus Body extrahieren (Content-Disposition: filename="...")
    Dim anhaenge As String: anhaenge = ExtrahiereAnhangFilenames(body)

    ' Body bereinigen (nur erster text/plain Block, oder alles wenn nicht multipart)
    Dim bodyClean As String: bodyClean = EmlBodyExtrahieren(body)
    bodyClean = BodyBereinigen(bodyClean)

    ' Datum parsen
    Dim datStr As String: datStr = ""
    If headers.Exists("date") Then datStr = headers("date")
    Dim dat As Date: dat = ParseEmlDatum(datStr)

    ' Absender-Name + Email aus "From" extrahieren
    Dim fromStr As String: fromStr = ""
    If headers.Exists("from") Then fromStr = headers("from")
    Dim absName As String, absMail As String
    AbsenderParsen fromStr, absName, absMail

    Dim subj As String: subj = ""
    If headers.Exists("subject") Then subj = MimeDecodeSimple(headers("subject"))

    ZeileSchreiben row, _
        FormatDatumDE(dat), _
        IsoDatum(dat), _
        absName, _
        absMail, _
        subj, _
        bodyClean, _
        anhaenge

    EmlDateiLesen = True
    Exit Function
Fehler:
    EmlDateiLesen = False
End Function

Private Function EmlBodyExtrahieren(rawBody As String) As String
    ' Falls multipart: erste text/plain Section nehmen
    Dim lc As String: lc = LCase(rawBody)
    Dim posPlain As Long: posPlain = InStr(lc, "content-type: text/plain")
    If posPlain = 0 Then
        EmlBodyExtrahieren = rawBody
        Exit Function
    End If

    ' Nach text/plain bis zur naechsten Boundary oder text/html
    Dim startBody As Long
    startBody = InStr(posPlain, rawBody, vbCrLf & vbCrLf)
    If startBody = 0 Then
        EmlBodyExtrahieren = rawBody
        Exit Function
    End If
    startBody = startBody + 4

    Dim endeBody As Long
    endeBody = InStr(startBody, rawBody, vbCrLf & "--")
    If endeBody = 0 Then endeBody = Len(rawBody) + 1

    EmlBodyExtrahieren = Mid(rawBody, startBody, endeBody - startBody)
End Function

Private Function ExtrahiereAnhangFilenames(rawBody As String) As String
    Dim result As String: result = ""
    Dim suchStr As String: suchStr = LCase(rawBody)
    Dim pos As Long: pos = 1
    Do
        pos = InStr(pos, suchStr, "filename=")
        If pos = 0 Then Exit Do
        Dim startName As Long: startName = pos + 9
        Dim endeName As Long
        Dim erstesZeichen As String: erstesZeichen = Mid(rawBody, startName, 1)
        If erstesZeichen = """" Then
            startName = startName + 1
            endeName = InStr(startName, rawBody, """")
        Else
            endeName = startName
            Do While endeName <= Len(rawBody)
                Dim c As String: c = Mid(rawBody, endeName, 1)
                If c = ";" Or c = vbCr Or c = vbLf Or c = " " Then Exit Do
                endeName = endeName + 1
            Loop
        End If
        If endeName > startName Then
            Dim fname As String: fname = Mid(rawBody, startName, endeName - startName)
            If result = "" Then result = fname Else result = result & "; " & fname
        End If
        pos = endeName + 1
    Loop
    ExtrahiereAnhangFilenames = result
End Function

' === SHEET-HELPER ===========================================================
Private Sub BlattVorbereiten()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_DATA)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SHEET_DATA
    Else
        ws.Cells.Clear
    End If

    ' Header
    ws.Cells(1, 1).Value = "Nr"
    ws.Cells(1, 2).Value = "Datum"
    ws.Cells(1, 3).Value = "Datum_ISO"
    ws.Cells(1, 4).Value = "Absender Name"
    ws.Cells(1, 5).Value = "Absender Email"
    ws.Cells(1, 6).Value = "Betreff"
    ws.Cells(1, 7).Value = "Text"
    ws.Cells(1, 8).Value = "Anhaenge"

    With ws.Range("A1:H1")
        .Font.Bold = True
        .Interior.Color = RGB(185, 28, 28) ' ERGO rot
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 22

    ws.Activate
End Sub

Private Sub ZeileSchreiben(row As Long, datumDE As String, datumIso As String, _
                            absName As String, absMail As String, _
                            betreff As String, text As String, anhaenge As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_DATA)

    ws.Cells(row, 1).Value = row - 1
    ws.Cells(row, 2).Value = datumDE
    ws.Cells(row, 3).Value = datumIso
    ws.Cells(row, 4).Value = absName
    ws.Cells(row, 5).Value = absMail
    ws.Cells(row, 6).Value = betreff
    ' Excel-Limit: max 32767 Zeichen pro Zelle
    If Len(text) > 32000 Then text = Left(text, 32000) & " [...gekuerzt]"
    ws.Cells(row, 7).Value = text
    ws.Cells(row, 8).Value = anhaenge
End Sub

Private Sub BlattAbschliessen(letzteZeile As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_DATA)

    ' Spaltenbreiten
    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 25
    ws.Columns("E").ColumnWidth = 30
    ws.Columns("F").ColumnWidth = 50
    ws.Columns("G").ColumnWidth = 80
    ws.Columns("H").ColumnWidth = 30
    ws.Columns("G").WrapText = True

    ' Datenbereich nur wenn Daten vorhanden
    If letzteZeile >= 2 Then
        ws.Range("A1:H" & letzteZeile).Borders.LineStyle = xlContinuous
        ws.Range("A1:H" & letzteZeile).Borders.Color = RGB(220, 220, 220)
        ' Autofilter
        ws.Range("A1:H1").AutoFilter
    End If

    ' Freeze Pane
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
End Sub

' === JSON-EXPORT ============================================================
Public Sub ExportiereJson()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_DATA)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Tabellenblatt '" & SHEET_DATA & "' nicht gefunden. Erst Emails abrufen.", vbExclamation
        Exit Sub
    End If

    Dim letzteZeile As Long
    letzteZeile = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If letzteZeile < 2 Then
        MsgBox "Keine Daten zum Exportieren.", vbExclamation
        Exit Sub
    End If

    ' Speichern-Dialog
    Dim defaultName As String
    defaultName = "ergo_emails_" & Format(Now, "yyyy-mm-dd_hh-nn") & ".json"

    Dim pfad As Variant
    pfad = Application.GetSaveAsFilename( _
        InitialFileName:=defaultName, _
        FileFilter:="JSON-Datei (*.json), *.json", _
        Title:="JSON-Beladungsdatei speichern")
    If pfad = False Then Exit Sub

    ' JSON bauen
    Dim json As String
    json = "[" & vbCrLf
    Dim i As Long
    For i = 2 To letzteZeile
        json = json & "  {" & vbCrLf
        json = json & "    ""datum"": " & JsonString(CStr(ws.Cells(i, 3).Value)) & "," & vbCrLf
        json = json & "    ""von_name"": " & JsonString(CStr(ws.Cells(i, 4).Value)) & "," & vbCrLf
        json = json & "    ""von_email"": " & JsonString(CStr(ws.Cells(i, 5).Value)) & "," & vbCrLf
        json = json & "    ""betreff"": " & JsonString(CStr(ws.Cells(i, 6).Value)) & "," & vbCrLf
        json = json & "    ""text"": " & JsonString(CStr(ws.Cells(i, 7).Value)) & "," & vbCrLf
        json = json & "    ""anhaenge"": " & JsonAnhangArray(CStr(ws.Cells(i, 8).Value)) & vbCrLf
        json = json & "  }"
        If i < letzteZeile Then json = json & ","
        json = json & vbCrLf
    Next i
    json = json & "]" & vbCrLf

    ' UTF-8 schreiben via ADODB.Stream
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' Text
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText json
    stream.SaveToFile CStr(pfad), 2 ' overwrite
    stream.Close

    MsgBox (letzteZeile - 1) & " Emails als JSON gespeichert:" & vbCrLf & pfad, _
           vbInformation, "Beladungsdatei erstellt"
End Sub

Private Function JsonString(s As String) As String
    Dim r As String: r = s
    r = Replace(r, "\", "\\")
    r = Replace(r, """", "\""")
    r = Replace(r, vbCrLf, "\n")
    r = Replace(r, vbLf, "\n")
    r = Replace(r, vbCr, "\n")
    r = Replace(r, vbTab, "\t")
    JsonString = """" & r & """"
End Function

Private Function JsonAnhangArray(s As String) As String
    If Trim(s) = "" Then
        JsonAnhangArray = "[]"
        Exit Function
    End If
    Dim teile() As String
    teile = Split(s, ";")
    Dim out As String: out = "["
    Dim i As Long
    For i = 0 To UBound(teile)
        If i > 0 Then out = out & ", "
        out = out & JsonString(Trim(teile(i)))
    Next i
    out = out & "]"
    JsonAnhangArray = out
End Function

' === HELPER =================================================================
Private Function GetOutlook() As Object
    On Error Resume Next
    Dim app As Object
    Set app = GetObject(, "Outlook.Application")
    If app Is Nothing Then Set app = CreateObject("Outlook.Application")
    Set GetOutlook = app
End Function

Private Function FormatDatumDE(d As Variant) As String
    On Error Resume Next
    If IsDate(d) Then
        FormatDatumDE = Format(d, "dd.mm.yyyy hh:nn")
    Else
        FormatDatumDE = ""
    End If
End Function

Private Function IsoDatum(d As Variant) As String
    On Error Resume Next
    If IsDate(d) Then
        IsoDatum = Format(d, "yyyy-mm-dd") & "T" & Format(d, "hh:nn:ss")
    Else
        IsoDatum = ""
    End If
End Function

Private Function AbsenderEmail(it As Object) As String
    On Error Resume Next
    Dim typ As String: typ = it.SenderEmailType
    If typ = "EX" Then
        Dim sender As Object
        Set sender = it.Sender
        If Not sender Is Nothing Then
            Dim eu As Object
            Set eu = sender.GetExchangeUser()
            If Not eu Is Nothing Then
                AbsenderEmail = eu.PrimarySmtpAddress
                If AbsenderEmail <> "" Then Exit Function
            End If
        End If
    End If
    AbsenderEmail = it.SenderEmailAddress
End Function

Private Function AnhangNamen(it As Object) As String
    On Error Resume Next
    Dim result As String: result = ""
    Dim n As Long: n = it.Attachments.Count
    Dim i As Long
    For i = 1 To n
        If result = "" Then
            result = it.Attachments.Item(i).FileName
        Else
            result = result & "; " & it.Attachments.Item(i).FileName
        End If
    Next i
    AnhangNamen = result
End Function

Private Function BodyBereinigen(text As String) As String
    Dim r As String: r = text
    r = Replace(r, vbCrLf, vbLf)
    r = Replace(r, vbCr, vbLf)
    ' Drei oder mehr Leerzeilen -> zwei
    Do While InStr(r, vbLf & vbLf & vbLf) > 0
        r = Replace(r, vbLf & vbLf & vbLf, vbLf & vbLf)
    Loop
    BodyBereinigen = Trim(Replace(r, vbLf, vbCrLf))
End Function

Private Sub AbsenderParsen(fromHeader As String, ByRef name As String, ByRef mail As String)
    name = "": mail = ""
    Dim s As String: s = Trim(fromHeader)
    If s = "" Then Exit Sub

    Dim posLt As Long: posLt = InStr(s, "<")
    Dim posGt As Long: posGt = InStr(s, ">")
    If posLt > 0 And posGt > posLt Then
        mail = Trim(Mid(s, posLt + 1, posGt - posLt - 1))
        name = Trim(Left(s, posLt - 1))
        ' Anfuehrungszeichen entfernen
        If Len(name) >= 2 Then
            If Left(name, 1) = """" And Right(name, 1) = """" Then
                name = Mid(name, 2, Len(name) - 2)
            End If
        End If
    Else
        ' Nur Email oder nur Name
        If InStr(s, "@") > 0 Then
            mail = s
        Else
            name = s
        End If
    End If
End Sub

Private Function ParseEmlDatum(s As String) As Date
    On Error Resume Next
    ' RFC822 z.B. "Thu, 01 Jan 2026 10:00:00 +0100"
    Dim clean As String: clean = Trim(s)
    ' Wochentag-Praefix abschneiden
    Dim posKomma As Long: posKomma = InStr(clean, ",")
    If posKomma > 0 And posKomma < 5 Then clean = Trim(Mid(clean, posKomma + 1))
    ' Zeitzone abschneiden (alles ab " +" oder " -" am Ende)
    Dim posTz As Long: posTz = InStr(clean, " +")
    If posTz = 0 Then posTz = InStr(clean, " -")
    If posTz > 10 Then clean = Left(clean, posTz - 1)
    Dim d As Date
    d = CDate(clean)
    If Err.Number = 0 Then
        ParseEmlDatum = d
    Else
        ParseEmlDatum = 0
    End If
End Function

Private Function MimeDecodeSimple(s As String) As String
    ' Sehr simpler MIME-Decoder fuer "=?utf-8?Q?...?=" und "=?utf-8?B?...?=" -
    ' falls nicht erkannt, original zurueckgeben.
    ' (Volle MIME-Dekodierung ist out-of-scope)
    Dim r As String: r = s
    If InStr(r, "=?") = 0 Then
        MimeDecodeSimple = r
        Exit Function
    End If
    ' Quoted-Printable einfache Faelle: "=?utf-8?Q?Hello_World?="
    r = Replace(r, "=?utf-8?Q?", "")
    r = Replace(r, "=?UTF-8?Q?", "")
    r = Replace(r, "=?iso-8859-1?Q?", "")
    r = Replace(r, "?=", "")
    r = Replace(r, "_", " ")
    MimeDecodeSimple = r
End Function
