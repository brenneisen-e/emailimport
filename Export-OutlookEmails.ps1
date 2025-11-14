# ============================================================================
# Export-OutlookEmails.ps1
# Exportiert ausgewählte Outlook-Emails als JSON für die BGAV Hypercare App
# ============================================================================

<#
.SYNOPSIS
    Exportiert ausgewählte Outlook-Emails als JSON-Datei.

.DESCRIPTION
    Dieses Script greift auf aktuell in Outlook ausgewählte Emails zu und
    exportiert sie im JSON-Format, das von der BGAV Hypercare Email Review App
    gelesen werden kann.

.PARAMETER OutputPath
    Pfad für die JSON-Ausgabedatei. Wenn nicht angegeben, wird ein Dialog geöffnet.

.PARAMETER FolderPath
    Optional: Outlook-Ordnerpfad zum Export aller Emails aus diesem Ordner.

.EXAMPLE
    .\Export-OutlookEmails.ps1
    Exportiert ausgewählte Emails und fragt nach dem Speicherort.

.EXAMPLE
    .\Export-OutlookEmails.ps1 -OutputPath "C:\Temp\emails.json"
    Exportiert ausgewählte Emails zur angegebenen Datei.

.EXAMPLE
    .\Export-OutlookEmails.ps1 -FolderPath "Posteingang\Hypercare"
    Exportiert alle Emails aus dem angegebenen Ordner.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$OutputPath,

    [Parameter(Mandatory=$false)]
    [string]$FolderPath
)

# Funktion zum Konvertieren von HTML zu Plain Text
function Convert-HtmlToText {
    param([string]$Html)

    if ([string]::IsNullOrEmpty($Html)) {
        return ""
    }

    # Einfache HTML-Bereinigung
    $text = $Html -replace '<br\s*/?>', "`n"
    $text = $text -replace '<[^>]+>', ''
    $text = $text -replace '&nbsp;', ' '
    $text = $text -replace '&amp;', '&'
    $text = $text -replace '&lt;', '<'
    $text = $text -replace '&gt;', '>'
    $text = $text -replace '&quot;', '"'

    return $text.Trim()
}

# Funktion zum Extrahieren von Email-Daten
function Export-EmailData {
    param($MailItem)

    try {
        # Absender-Informationen extrahieren
        $senderName = if ($MailItem.SenderName) { $MailItem.SenderName } else { "Unbekannt" }
        $senderEmail = if ($MailItem.SenderEmailAddress) {
            # SMTP-Adresse extrahieren (bei Exchange-Adressen)
            if ($MailItem.SenderEmailAddress -like "/O=*") {
                try {
                    $sender = $MailItem.Sender
                    if ($sender.GetExchangeUser()) {
                        $sender.GetExchangeUser().PrimarySmtpAddress
                    } else {
                        $MailItem.SenderEmailAddress
                    }
                } catch {
                    $MailItem.SenderEmailAddress
                }
            } else {
                $MailItem.SenderEmailAddress
            }
        } else {
            "unknown@unknown.com"
        }

        # Email-Body extrahieren (Plain Text bevorzugen)
        $bodyText = if ($MailItem.Body) {
            $MailItem.Body
        } elseif ($MailItem.HTMLBody) {
            Convert-HtmlToText -Html $MailItem.HTMLBody
        } else {
            ""
        }

        # Anhänge sammeln
        $attachments = @()
        if ($MailItem.Attachments.Count -gt 0) {
            foreach ($attachment in $MailItem.Attachments) {
                $attachments += $attachment.FileName
            }
        }

        # Email-Objekt erstellen
        return [PSCustomObject]@{
            datum = $MailItem.ReceivedTime.ToString("yyyy-MM-ddTHH:mm:ss")
            von_email = $senderEmail
            von_name = $senderName
            betreff = if ($MailItem.Subject) { $MailItem.Subject } else { "(Kein Betreff)" }
            text = $bodyText
            anhänge = $attachments
        }
    } catch {
        Write-Warning "Fehler beim Verarbeiten einer Email: $_"
        return $null
    }
}

# Hauptlogik
try {
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host "  BGAV Hypercare - Outlook Email Export" -ForegroundColor Cyan
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host ""

    # Outlook-Objekt erstellen
    Write-Host "Verbinde mit Outlook..." -ForegroundColor Yellow
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")

    $emails = @()

    if ($FolderPath) {
        # Ordner-Export
        Write-Host "Exportiere Emails aus Ordner: $FolderPath" -ForegroundColor Yellow

        # Ordner finden
        $folder = $namespace.Folders.Item(1) # Standardpostfach
        foreach ($folderName in $FolderPath.Split('\')) {
            $folder = $folder.Folders.Item($folderName)
        }

        $totalCount = $folder.Items.Count
        Write-Host "Gefunden: $totalCount Email(s)" -ForegroundColor Green

        $counter = 0
        foreach ($mailItem in $folder.Items) {
            $counter++
            Write-Progress -Activity "Exportiere Emails" -Status "Email $counter von $totalCount" -PercentComplete (($counter / $totalCount) * 100)

            if ($mailItem.Class -eq 43) { # olMail
                $emailData = Export-EmailData -MailItem $mailItem
                if ($emailData) {
                    $emails += $emailData
                }
            }
        }
        Write-Progress -Activity "Exportiere Emails" -Completed

    } else {
        # Ausgewählte Emails exportieren
        Write-Host "Lese ausgewählte Emails..." -ForegroundColor Yellow

        $explorer = $outlook.ActiveExplorer()
        $selection = $explorer.Selection

        if ($selection.Count -eq 0) {
            Write-Host "FEHLER: Keine Emails ausgewählt!" -ForegroundColor Red
            Write-Host ""
            Write-Host "Bitte wähle in Outlook eine oder mehrere Emails aus und führe das Script erneut aus." -ForegroundColor Yellow
            exit 1
        }

        Write-Host "Gefunden: $($selection.Count) ausgewählte Email(s)" -ForegroundColor Green
        Write-Host ""

        $counter = 0
        foreach ($mailItem in $selection) {
            $counter++
            Write-Progress -Activity "Exportiere Emails" -Status "Email $counter von $($selection.Count)" -PercentComplete (($counter / $selection.Count) * 100)

            if ($mailItem.Class -eq 43) { # olMail
                $emailData = Export-EmailData -MailItem $mailItem
                if ($emailData) {
                    $emails += $emailData
                    Write-Host "✓ $($mailItem.Subject)" -ForegroundColor Green
                }
            }
        }
        Write-Progress -Activity "Exportiere Emails" -Completed
    }

    Write-Host ""
    Write-Host "Erfolgreich verarbeitet: $($emails.Count) Email(s)" -ForegroundColor Green

    if ($emails.Count -eq 0) {
        Write-Host "WARNUNG: Keine Emails zum Exportieren!" -ForegroundColor Yellow
        exit 0
    }

    # Ausgabepfad bestimmen
    if (-not $OutputPath) {
        Add-Type -AssemblyName System.Windows.Forms
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "JSON Dateien (*.json)|*.json|Alle Dateien (*.*)|*.*"
        $saveDialog.Title = "Emails exportieren nach..."
        $saveDialog.FileName = "hypercare_emails_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"

        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $OutputPath = $saveDialog.FileName
        } else {
            Write-Host "Export abgebrochen." -ForegroundColor Yellow
            exit 0
        }
    }

    # JSON exportieren
    Write-Host ""
    Write-Host "Exportiere nach: $OutputPath" -ForegroundColor Yellow

    $jsonContent = $emails | ConvertTo-Json -Depth 10
    $jsonContent | Out-File -FilePath $OutputPath -Encoding UTF8

    Write-Host ""
    Write-Host "================================================" -ForegroundColor Green
    Write-Host "  ✓ Export erfolgreich abgeschlossen!" -ForegroundColor Green
    Write-Host "================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Exportierte Emails: $($emails.Count)" -ForegroundColor Cyan
    Write-Host "Datei: $OutputPath" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Du kannst die JSON-Datei jetzt in die BGAV Hypercare Email Review App laden." -ForegroundColor Yellow
    Write-Host ""

    # Optional: Datei im Explorer öffnen
    $openFile = Read-Host "Ordner im Explorer öffnen? (J/N)"
    if ($openFile -eq 'J' -or $openFile -eq 'j') {
        $folder = Split-Path -Parent $OutputPath
        Start-Process explorer.exe -ArgumentList $folder
    }

} catch {
    Write-Host ""
    Write-Host "FEHLER beim Export:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "Stack Trace:" -ForegroundColor Yellow
    Write-Host $_.Exception.StackTrace -ForegroundColor Yellow
    exit 1
} finally {
    # COM-Objekte freigeben
    if ($outlook) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
