/**
 * Import Core Module
 * Main import logic for JSON to Excel
 */

/**
 * Select JSON file via file dialog
 */
function selectJsonFile() {
    try {
        var shell = new ActiveXObject("Shell.Application");
        var folder = shell.BrowseForFolder(0, "JSON-Datei auswählen", 0x4000, "shell:Downloads");
        if (folder && folder.Self) {
            // This returns a folder - need to use different approach
            // Use InputBox as fallback
            var wshell = new ActiveXObject("WScript.Shell");
            var path = wshell.Exec('powershell -command "[System.Reflection.Assembly]::LoadWithPartialName(\'System.windows.forms\') | Out-Null; $dialog = New-Object System.Windows.Forms.OpenFileDialog; $dialog.Filter = \'JSON files (*.json)|*.json\'; $dialog.InitialDirectory = [Environment]::GetFolderPath(\'UserProfile\') + \'\\Downloads\'; if($dialog.ShowDialog() -eq \'OK\') { $dialog.FileName }"').StdOut.ReadAll().trim();
            if (path) {
                loadJsonFile(path);
            }
        }
    } catch (e) {
        // Fallback: prompt for path
        var path = prompt("JSON-Datei Pfad eingeben:", "");
        if (path) {
            loadJsonFile(path);
        }
    }
}

/**
 * Load JSON file
 * @param {string} path - File path
 */
function loadJsonFile(path) {
    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        if (!fso.FileExists(path)) {
            document.getElementById('jsonStatus').style.display = 'block';
            document.getElementById('jsonStatus').className = 'status error';
            document.getElementById('jsonStatus').innerHTML = 'Datei nicht gefunden: ' + path;
            return;
        }

        // Read file with UTF-16LE encoding
        var stream = new ActiveXObject("ADODB.Stream");
        stream.Type = 2; // adTypeText
        stream.Charset = "UTF-16LE";
        stream.Open();
        stream.LoadFromFile(path);
        var content = stream.ReadText();
        stream.Close();

        // Remove BOM if present
        if (content.charCodeAt(0) === 0xFEFF) {
            content = content.substring(1);
        }

        // Parse JSON
        jsonData = JSON.parse(content);
        document.getElementById('jsonPath').value = path;
        document.getElementById('jsonStatus').style.display = 'block';
        document.getElementById('jsonStatus').className = 'status success';
        document.getElementById('jsonStatus').innerHTML = jsonData.length + ' Emails geladen';

        checkImportReady();

    } catch (e) {
        document.getElementById('jsonStatus').style.display = 'block';
        document.getElementById('jsonStatus').className = 'status error';
        document.getElementById('jsonStatus').innerHTML = 'Fehler: ' + e.message;
        jsonData = null;
        checkImportReady();
    }
}

/**
 * Start the import process
 */
function startImport() {
    if (!targetWorkbook || !jsonData) {
        alert('Excel und JSON müssen verbunden sein!');
        return;
    }

    document.getElementById('importProgress').style.display = 'block';
    document.getElementById('btnImport').disabled = true;
    document.getElementById('importSuccess').style.display = 'none';
    document.getElementById('importError').style.display = 'none';

    setTimeout(doImport, 100);
}

/**
 * Execute the import
 */
function doImport() {
    var currentStep = 'Start';
    var currentEmailIndex = -1;
    var currentEmailBetreff = '';

    try {
        currentStep = 'Worksheet öffnen';
        var sheetName = document.getElementById('sheetName').value || 'Übersicht';
        var worksheet;
        try {
            worksheet = targetWorkbook.Worksheets(sheetName);
        } catch (e) {
            worksheet = targetWorkbook.Worksheets(1);
        }

        // Find table if exists
        currentStep = 'Tabelle suchen';
        var listObject = null;
        var tableLastRow = 0;
        try {
            listObject = worksheet.ListObjects(1);
            if (listObject) {
                var dataBodyRange = listObject.DataBodyRange;
                if (dataBodyRange) {
                    tableLastRow = dataBodyRange.Row + dataBodyRange.Rows.Count - 1;
                } else {
                    tableLastRow = listObject.HeaderRowRange.Row;
                }
            }
        } catch (e) {
            listObject = null;
        }

        // Read existing data
        currentStep = 'Bestehende Daten lesen';
        document.getElementById('importProgressText').innerText = 'Lese bestehende Daten...';
        var existingData = readExistingData(worksheet);

        currentStep = 'Letzte Zeile finden';
        var lastRow;
        if (listObject && tableLastRow > 0) {
            lastRow = tableLastRow;
        } else {
            lastRow = worksheet.Cells(worksheet.Rows.Count, 1).End(-4162).Row;
            if (lastRow === 1 && !worksheet.Cells(1, 1).Value) lastRow = 0;
        }
        var nextRow = lastRow + 1;

        var total = jsonData.length;
        var newCount = 0;
        var skipCount = 0;
        var updateCount = 0;

        for (var i = 0; i < total; i++) {
            currentEmailIndex = i;
            currentStep = 'Email ' + (i+1) + ' laden';
            var email = jsonData[i];
            currentEmailBetreff = email.betreff || '(kein Betreff)';

            if (i % 3 === 0) {
                updateImportProgress(i, total);
                document.getElementById('importProgressText').innerText =
                    'Verarbeite ' + (i+1) + '/' + total + ' (Neu: ' + newCount + ', Aktualisiert: ' + updateCount + ', Übersprungen: ' + skipCount + ')';
            }

            currentStep = 'Email ' + (i+1) + ' Key erstellen';
            var emailKey = createEmailKey(email.datum, email.betreff);
            var existingRow = existingData[emailKey];

            if (existingRow) {
                // Email exists - check for new replies
                currentStep = 'Email ' + (i+1) + ' Antworten prüfen';
                if (email.antworten && email.antworten.length > 0) {
                    var currentReplies = worksheet.Cells(existingRow, 14).Value || '';

                    currentStep = 'Email ' + (i+1) + ' Antworten formatieren';
                    var newRepliesText = formatReplies(email.antworten);

                    var normalizedCurrent = fixEncoding(currentReplies).toLowerCase();
                    var normalizedNew = fixEncoding(newRepliesText).toLowerCase();

                    if (newRepliesText && normalizedCurrent.indexOf(normalizedNew.substring(0, 50)) === -1) {
                        currentStep = 'Email ' + (i+1) + ' Antworten schreiben';
                        var combinedReplies = currentReplies ? newRepliesText + '\n\n' + currentReplies : newRepliesText;
                        combinedReplies = sanitizeText(combinedReplies);
                        worksheet.Cells(existingRow, 14).Value = combinedReplies;

                        currentStep = 'Email ' + (i+1) + ' Timestamps fetten';
                        boldTimestamps(worksheet.Cells(existingRow, 14));

                        // Update ReplyIDs
                        currentStep = 'Email ' + (i+1) + ' ReplyIDs aktualisieren';
                        try {
                            var currentReplyIds = worksheet.Cells(existingRow, 24).Value || '';
                            var newReplyIds = [];
                            for (var nr = 0; nr < email.antworten.length; nr++) {
                                if (email.antworten[nr].replyId) {
                                    newReplyIds.push(email.antworten[nr].replyId);
                                }
                            }
                            if (newReplyIds.length > 0) {
                                var combinedIds = currentReplyIds ? currentReplyIds + ',' + newReplyIds.join(',') : newReplyIds.join(',');
                                worksheet.Cells(existingRow, 24).Value = combinedIds;
                            }
                        } catch (replyIdErr) {}

                        updateCount++;
                    } else {
                        skipCount++;
                    }
                } else {
                    skipCount++;
                }
            } else {
                // New email
                currentStep = 'Email ' + (i+1) + ' schreiben (Zeile ' + nextRow + ')';
                try {
                    var actualRow = nextRow;
                    if (listObject) {
                        try {
                            var newListRow = listObject.ListRows.Add();
                            actualRow = newListRow.Range.Row;
                        } catch (addErr) {}
                    }
                    writeEmailRow(worksheet, actualRow, email);
                    existingData[emailKey] = actualRow;
                    nextRow = actualRow + 1;
                    newCount++;
                } catch (writeErr) {
                    currentStep = 'Email ' + (i+1) + ' FEHLER - übersprungen';
                    try {
                        var errorRow = nextRow;
                        if (listObject) {
                            try {
                                var errListRow = listObject.ListRows.Add();
                                errorRow = errListRow.Range.Row;
                            } catch (e3) {}
                        }
                        worksheet.Cells(errorRow, 1).Value = 'HC-' + (errorRow - 1);
                        worksheet.Cells(errorRow, 2).Value = email.datum || '';
                        worksheet.Cells(errorRow, 12).Value = email.betreff || '';
                        worksheet.Cells(errorRow, 13).Value = '[FEHLER beim Import: ' + writeErr.message + ']';
                        existingData[emailKey] = errorRow;
                        nextRow = errorRow + 1;
                        skipCount++;
                    } catch (e2) {
                        skipCount++;
                    }
                }
            }
        }

        // Auto-Erledigt
        currentStep = 'Auto-Erledigt prüfen';
        var erledigtCount = 0;
        try {
            for (var convId in existingExcelRows) {
                if (existingExcelRows.hasOwnProperty(convId) && !foundInInbox[convId]) {
                    var erledigtRow = existingExcelRows[convId];
                    try {
                        worksheet.Cells(erledigtRow, 9).Value = 'Erledigt';
                        erledigtCount++;
                    } catch (erlErr) {}
                }
            }
        } catch (autoErr) {}

        currentStep = 'Speichern';
        updateImportProgress(total, total);
        targetWorkbook.Save();

        document.getElementById('importProgress').style.display = 'none';
        document.getElementById('importSuccess').style.display = 'block';
        var successMsg = '<strong>' + newCount + '</strong> neue Vorgange importiert<br>' +
            '<strong>' + updateCount + '</strong> Antworten aktualisiert<br>' +
            '<strong>' + skipCount + '</strong> Duplikate ubersprungen';
        if (erledigtCount > 0) {
            successMsg += '<br><strong>' + erledigtCount + '</strong> Vorgange auf Erledigt gesetzt';
        }
        document.getElementById('importSuccessMsg').innerHTML = successMsg;
        document.getElementById('btnImport').disabled = false;

        // Step 5 completed
        completeStep(5);

    } catch (e) {
        document.getElementById('importProgress').style.display = 'none';
        document.getElementById('importError').style.display = 'block';

        var errorDetails = 'Fehler bei: ' + currentStep + '\n';
        if (_writeRowCurrentField) {
            errorDetails += 'Letztes Feld: ' + _writeRowCurrentField + '\n';
        }
        errorDetails += 'Email-Index: ' + (currentEmailIndex + 1) + ' von ' + (jsonData ? jsonData.length : 0) + '\n';
        errorDetails += 'Betreff: ' + currentEmailBetreff + '\n\n';
        errorDetails += 'Fehlermeldung: ' + e.message + '\n';
        if (e.number) errorDetails += 'Fehlernummer: ' + e.number + '\n';

        if (currentEmailIndex >= 0 && jsonData && jsonData[currentEmailIndex]) {
            var debugEmail = jsonData[currentEmailIndex];
            errorDetails += '\n=== Email-Daten ===\n';
            errorDetails += 'Datum: ' + (debugEmail.datum || '') + '\n';
            errorDetails += 'Von: ' + (debugEmail.von_name || debugEmail.von_email || '') + '\n';
            errorDetails += 'Betreff: ' + (debugEmail.betreff || '') + '\n';
        }

        document.getElementById('importErrorMsg').innerHTML =
            '<textarea readonly style="width:100%;height:250px;font-size:11px;background:#1a1a2e;color:#ef4444;border:1px solid #ef4444;padding:8px;font-family:monospace;" onclick="this.select()">' +
            errorDetails.replace(/</g, '&lt;').replace(/>/g, '&gt;') +
            '</textarea><br><small>Klicken zum Kopieren</small>';
        document.getElementById('btnImport').disabled = false;
    }
}

