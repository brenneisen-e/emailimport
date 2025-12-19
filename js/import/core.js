/**
 * Import Core Module
 * Main import logic for JSON to Excel
 */

/**
 * Debug logging helper - writes to visible debug textarea
 */
function debugLog(msg) {
    try {
        var debugSection = document.getElementById('debugSection');
        var debugOutput = document.getElementById('debugOutput');
        if (debugSection && debugOutput) {
            debugSection.style.display = 'block';
            debugOutput.value += msg + '\n';
            debugOutput.scrollTop = debugOutput.scrollHeight;
        }
    } catch (e) {}
}

function clearDebugLog() {
    try {
        var debugOutput = document.getElementById('debugOutput');
        if (debugOutput) {
            debugOutput.value = '';
        }
    } catch (e) {}
}

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

        // Read file using ADODB.Stream for proper UTF-8 support
        var content = '';
        try {
            var stream = new ActiveXObject("ADODB.Stream");
            stream.Type = 2; // adTypeText
            stream.Charset = "UTF-8";
            stream.Open();
            stream.LoadFromFile(path);
            content = stream.ReadText(-1); // -1 = adReadAll
            stream.Close();
        } catch (streamErr) {
            // Fallback to FSO if ADODB fails
            var file = fso.OpenTextFile(path, 1, false, 0); // 0 = ASCII/Default
            content = file.ReadAll();
            file.Close();
        }

        // Remove BOM if present
        if (content.charCodeAt(0) === 0xFEFF || content.charCodeAt(0) === 0xEF) {
            content = content.substring(1);
        }
        // Also check for UTF-8 BOM sequence that might appear as characters
        if (content.substring(0, 3) === '\xEF\xBB\xBF') {
            content = content.substring(3);
        }

        // Parse JSON
        var parsed = JSON.parse(content);

        // Handle different JSON formats
        if (parsed && parsed.conversations) {
            // New conversation format - convert to array
            jsonData = convertConversationsToEmailArray(parsed.conversations);
            document.getElementById('jsonStatus').innerHTML = jsonData.length + ' Vorgänge aus Konversations-Export geladen';
        } else if (parsed && parsed._exportInfo && parsed.emails) {
            // Debug format with export info
            jsonData = parsed.emails;
            // Add sentOnly flag for bearbeitet format if missing
            jsonData = addSentOnlyFlagIfMissing(jsonData);
            document.getElementById('jsonStatus').innerHTML = jsonData.length + ' Emails geladen';
        } else if (Array.isArray(parsed)) {
            // Direct array (bearbeitet_*.json format)
            jsonData = parsed;
            // Add sentOnly flag for bearbeitet format if missing
            jsonData = addSentOnlyFlagIfMissing(jsonData);
            // DEBUG: Count sentOnly
            var sentOnlyCount = 0;
            for (var si = 0; si < jsonData.length; si++) {
                if (jsonData[si].sentOnly) sentOnlyCount++;
            }
            debugLog('=== ' + sentOnlyCount + ' von ' + jsonData.length + ' als sentOnly markiert ===');
            document.getElementById('jsonStatus').innerHTML = jsonData.length + ' Emails geladen' + (sentOnlyCount > 0 ? ' (' + sentOnlyCount + ' sentOnly)' : '');
        } else {
            // Single email
            jsonData = [parsed];
            document.getElementById('jsonStatus').innerHTML = '1 Email geladen';
        }

        document.getElementById('jsonPath').value = path;
        document.getElementById('jsonStatus').style.display = 'block';
        document.getElementById('jsonStatus').className = 'status success';

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
 * Add sentOnly flag to emails if missing (for bearbeitet format)
 * Detects sent-only conversations (replies without original incoming request).
 *
 * A conversation is "sentOnly" when:
 * - The root message was SENT by us (not received in inbox)
 * - There are no incoming messages in the conversation
 *
 * This typically happens when we reply to an old request that's outside
 * the export timeframe - we have our reply but not the original request.
 *
 * @param {Array} emails - Array of email objects
 * @returns {Array} Emails with sentOnly flag set
 */
function addSentOnlyFlagIfMissing(emails) {
    if (!emails || !Array.isArray(emails)) return emails;

    // Known mailbox addresses that indicate "sent from us"
    var mailboxPatterns = [
        'provisionsservice@barmenia.de',
        'provisionsservice@'
    ];

    for (var i = 0; i < emails.length; i++) {
        var email = emails[i];

        // Skip if sentOnly is already explicitly set
        if (email.sentOnly === true || email.sentOnly === false) continue;

        var vonEmail = (email.von_email || '').toLowerCase();

        // Check if the root message was sent FROM our mailbox (not received)
        var isSentFromMailbox = false;
        for (var m = 0; m < mailboxPatterns.length; m++) {
            if (vonEmail.indexOf(mailboxPatterns[m]) !== -1) {
                isSentFromMailbox = true;
                break;
            }
        }

        // Check if there are any INCOMING messages (received from external)
        var hasIncomingMessage = false;
        var antworten = email.antworten || email.alleAntworten || [];
        for (var j = 0; j < antworten.length; j++) {
            var reply = antworten[j];
            // isIncoming=true means the reply was received (not sent by us)
            if (reply.isIncoming === true) {
                hasIncomingMessage = true;
                break;
            }
        }

        // Mark as sentOnly if:
        // - Root message was sent FROM our mailbox (we sent it, not received)
        // - AND there are no incoming replies in the conversation
        // This indicates we only have outgoing messages - likely replies to old requests
        if (isSentFromMailbox && !hasIncomingMessage) {
            email.sentOnly = true;
        } else {
            email.sentOnly = false;
        }
    }

    return emails;
}

/**
 * Convert conversation format to flat email array for Excel import
 * @param {Object} conversations - Conversations object
 * @returns {Array} Array of emails with antworten
 */
function convertConversationsToEmailArray(conversations) {
    var result = [];

    for (var convId in conversations) {
        if (!conversations.hasOwnProperty(convId)) continue;

        var conv = conversations[convId];
        var messages = conv.messages || [];
        if (messages.length === 0) continue;

        // Skip "Online-Abschluss Vermittlerzuordnung" conversations (not needed)
        var convSubject = (conv.subject || '').toLowerCase();
        if (convSubject.indexOf('online-abschluss vermittlerzuordnung') !== -1) continue;

        // Sort by timestamp
        messages.sort(function(a, b) {
            var timeA = a.receivedTime || a.sentOn || '';
            var timeB = b.receivedTime || b.sentOn || '';
            return timeA < timeB ? -1 : (timeA > timeB ? 1 : 0);
        });

        // Find root (first inbox message, or first message)
        var rootMsg = null;
        var hasInboxRoot = false;
        for (var i = 0; i < messages.length; i++) {
            if (messages[i].folder === 'inbox') {
                rootMsg = messages[i];
                hasInboxRoot = true;
                break;
            }
        }
        if (!rootMsg) rootMsg = messages[0];

        // Build email object
        var email = {
            emailId: rootMsg.entryID || '',
            conversationId: convId,
            datum: rootMsg.receivedTime || rootMsg.sentOn || '',
            von_email: rootMsg.senderEmail || '',
            von_name: rootMsg.senderName || '',
            betreff: rootMsg.subject || conv.subject || '',
            text: rootMsg.body || rootMsg.bodyPreview || '',
            htmlBody: rootMsg.htmlBody || '',
            kategorie: '',
            status: '',
            antwort: '',
            kommentar: '',
            sentOnly: !hasInboxRoot,  // Flag: only sent messages, no inbox root
            antworten: []
        };

        // Add non-root messages as antworten (for sentOnly, add ALL messages including root)
        for (var j = 0; j < messages.length; j++) {
            var msg = messages[j];
            // For normal conversations: skip root. For sentOnly: include all as potential replies
            if (!email.sentOnly && msg.entryID === rootMsg.entryID) continue;

            var isSent = msg.folder === 'sent';
            email.antworten.push({
                datum: msg.sentOn || msg.receivedTime || '',
                von: isSent ? 'Ich' : (msg.senderName || msg.senderEmail || ''),
                text: msg.body || msg.bodyPreview || '',
                isIncoming: !isSent,
                replyId: msg.entryID || ''
            });
        }

        result.push(email);
    }

    return result;
}

/**
 * Start the import process
 */
function startImport() {
    if (!targetWorkbook || !jsonData) {
        alert('Excel und JSON muessen verbunden sein!');
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
        currentStep = 'Worksheet oeffnen';
        var sheetName = document.getElementById('sheetName').value || 'Uebersicht';
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

        // Count existing entries for debugging
        var existingCount = Object.keys(existingData.byConvId).length;
        var existingKeyCount = Object.keys(existingData.byKey).length;
        clearDebugLog();
        debugLog('=== Excel hat ' + existingCount + ' Eintraege (ConvID) ===');

        // Show existing convIds
        if (existingCount > 0) {
            for (var ecid in existingData.byConvId) {
                debugLog('  Excel: ' + ecid.substring(0, 8) + '... -> Zeile ' + existingData.byConvId[ecid]);
            }
        }

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

            // Find existing row: by ConversationID if available, otherwise by date+subject
            // IMPORTANT: If email has conversationId, ONLY check by convId (most reliable)
            // The emailKey check is only for legacy emails without conversationId
            var existingRow = null;
            if (email.conversationId) {
                // Has conversationId - only check by conversationId
                if (existingData.byConvId[email.conversationId]) {
                    existingRow = existingData.byConvId[email.conversationId];
                }
            } else {
                // No conversationId - fallback to date+subject key
                if (existingData.byKey[emailKey]) {
                    existingRow = existingData.byKey[emailKey];
                }
            }

            // DEBUG: Log why email is being processed
            var debugReason = '';
            if (existingRow) {
                debugReason = 'EXISTING (row ' + existingRow + ')';
            } else if (email.sentOnly) {
                debugReason = 'SENT-ONLY';
            } else {
                debugReason = 'NEW';
            }
            debugLog('Email ' + (i+1) + ': ' + debugReason + ' - ' + (email.betreff || '').substring(0, 30));

            if (existingRow) {
                // Email exists - check for new replies using ReplyIDs
                currentStep = 'Email ' + (i+1) + ' Antworten prüfen';
                // For sentOnly conversations, mark that we found a match
                var matchedSentOnly = email.sentOnly ? true : false;
                if (email.antworten && email.antworten.length > 0) {
                    // Get existing ReplyIDs from Excel
                    var currentReplyIdsStr = '';
                    try {
                        currentReplyIdsStr = worksheet.Cells(existingRow, 24).Value || '';
                    } catch (e) {}

                    // Build set of existing IDs for fast lookup
                    var existingIds = {};
                    if (currentReplyIdsStr) {
                        var idList = currentReplyIdsStr.split(',');
                        for (var ei = 0; ei < idList.length; ei++) {
                            var trimmedId = idList[ei].trim();
                            if (trimmedId) existingIds[trimmedId] = true;
                        }
                    }

                    // Filter to only truly new replies
                    var newReplies = [];
                    for (var nr = 0; nr < email.antworten.length; nr++) {
                        var reply = email.antworten[nr];
                        if (reply.replyId && !existingIds[reply.replyId]) {
                            newReplies.push(reply);
                        } else if (!reply.replyId) {
                            // No replyId - check by timestamp
                            var replyTimestamp = reply.datum || '';
                            var currentRepliesText = '';
                            try {
                                currentRepliesText = worksheet.Cells(existingRow, 14).Value || '';
                            } catch (e) {}
                            if (replyTimestamp && currentRepliesText.indexOf(replyTimestamp) === -1) {
                                newReplies.push(reply);
                            }
                        }
                    }

                    if (newReplies.length > 0) {
                        currentStep = 'Email ' + (i+1) + ' ' + newReplies.length + ' neue Antworten formatieren';
                        var newRepliesText = formatReplies(newReplies);

                        currentStep = 'Email ' + (i+1) + ' Antworten schreiben';
                        var currentReplies = '';
                        try {
                            currentReplies = worksheet.Cells(existingRow, 14).Value || '';
                        } catch (e) {}
                        var combinedReplies = currentReplies ? currentReplies + '\n\n' + newRepliesText : newRepliesText;
                        // Use sanitizeForExcel to remove \r chars that cause _x000D_
                        combinedReplies = sanitizeForExcel(combinedReplies);
                        worksheet.Cells(existingRow, 14).Value = combinedReplies;

                        currentStep = 'Email ' + (i+1) + ' Timestamps fetten';
                        boldTimestamps(worksheet.Cells(existingRow, 14));

                        // Update ReplyIDs
                        currentStep = 'Email ' + (i+1) + ' ReplyIDs aktualisieren';
                        try {
                            var newReplyIds = [];
                            for (var nri = 0; nri < newReplies.length; nri++) {
                                if (newReplies[nri].replyId) {
                                    newReplyIds.push(newReplies[nri].replyId);
                                }
                            }
                            if (newReplyIds.length > 0) {
                                var combinedIds = currentReplyIdsStr ? currentReplyIdsStr + ',' + newReplyIds.join(',') : newReplyIds.join(',');
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
                // New email - but skip if sentOnly (no matching case in Excel)
                if (email.sentOnly) {
                    // sentOnly conversation with no matching case - skip entirely
                    // These are replies to old cases not in this Excel file
                    skipCount++;
                    continue;
                }

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
                    existingData.byKey[emailKey] = actualRow;
                    if (email.conversationId) {
                        existingData.byConvId[email.conversationId] = actualRow;
                    }
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
                        existingData.byKey[emailKey] = errorRow;
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
            '<strong>' + skipCount + '</strong> Duplikate/Sent-Only ubersprungen';
        if (existingCount > 0) {
            successMsg += '<br><em style="color:#666">(Excel hatte bereits ' + existingCount + ' Vorgange)</em>';
        }
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

