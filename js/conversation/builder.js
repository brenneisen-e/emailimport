/**
 * Conversation Builder Module
 * Collects emails from Inbox + Sent, deduplicates, and builds conversation trees
 */

/**
 * Clean X500 Exchange address to readable name
 * @param {string} addr - Email address or X500 path
 * @returns {string} Clean email or name
 */
function cleanX500Address(addr) {
    if (!addr) return '';
    // If already clean email, return as-is
    if (addr.indexOf('@') > -1 && addr.indexOf('/cn=') === -1) {
        return addr;
    }
    // Extract name from X500 format: /o=.../cn=hash-Name
    if (addr.indexOf('/cn=') > -1) {
        // Remove trailing comma/whitespace first
        var cleaned = addr.replace(/[,\s]+$/, '');
        var match = cleaned.match(/cn=([^\/]+)$/i);
        if (match) {
            // Remove leading GUID (like "f4fadad6c63f4b8e93e39ddb6abbf8d7-")
            return match[1].replace(/^[a-f0-9]{32}-/i, '');
        }
    }
    return addr;
}

// State for async conversation extraction
var convExportState = {
    phase: 0, // 0=init, 1=inbox, 2=sent, 3=building, 4=done
    allEmails: [],
    seenEntryIds: {}, // For deduplication
    inboxFolder: null,
    sentFolder: null,
    currentFolder: null,
    currentItems: null,
    currentIdx: 0,
    maxItems: 0,
    cutoffDate: null,
    conversations: null,
    mailboxName: ''
};

/**
 * Start conversation-based export
 * @param {string} folderId - Selected folder ID or 'inbox'
 * @param {number} days - Number of days to export
 */
function startConversationExport(folderId, days) {
    // Reset state
    convExportState.phase = 0;
    convExportState.allEmails = [];
    convExportState.seenEntryIds = {};
    convExportState.conversations = null;

    // Calculate cutoff date
    convExportState.cutoffDate = new Date();
    convExportState.cutoffDate.setDate(convExportState.cutoffDate.getDate() - days);

    // Get mailbox name
    try {
        convExportState.mailboxName = selectedStore.DisplayName || '';
    } catch (e) {
        convExportState.mailboxName = '';
    }

    // Get folders
    try {
        // Inbox folder
        if (folderId === 'inbox') {
            convExportState.inboxFolder = selectedStore.GetDefaultFolder(6);
        } else {
            convExportState.inboxFolder = outlookNS.GetFolderFromID(folderId);
        }

        // Sent folder - always use default sent folder
        convExportState.sentFolder = selectedStore.GetDefaultFolder(5);

    } catch (e) {
        showExportError("Fehler beim Laden der Ordner: " + e.message);
        return;
    }

    // Show progress
    document.getElementById('exportProgress').style.display = 'block';
    document.getElementById('exportProgressText').innerText = 'Phase 1/4: Lade Inbox...';
    document.getElementById('exportProgressFill').style.width = '5%';

    // Start with inbox
    convExportState.phase = 1;
    startFolderExtraction(convExportState.inboxFolder, 'inbox');
}

/**
 * Start extracting from a folder
 * @param {Object} folder - Outlook folder
 * @param {string} folderType - 'inbox' or 'sent'
 */
function startFolderExtraction(folder, folderType) {
    try {
        convExportState.currentFolder = folder;

        // Filter by date
        var dateField = folderType === 'sent' ? '[SentOn]' : '[ReceivedTime]';
        var dateFilter = dateField + " >= '" + formatDateForRestrict(convExportState.cutoffDate) + "'";

        try {
            convExportState.currentItems = folder.Items.Restrict(dateFilter);
        } catch (e) {
            convExportState.currentItems = folder.Items;
        }

        try {
            convExportState.currentItems.Sort(dateField, true);
        } catch (e) {}

        convExportState.maxItems = Math.min(convExportState.currentItems.Count, 1000);
        convExportState.currentIdx = 1;

        setTimeout(function() { processConvExportBatch(folderType); }, 10);

    } catch (e) {
        // Skip to next phase if folder fails
        advanceConvExportPhase();
    }
}

/**
 * Process a batch of emails from current folder
 * @param {string} folderType - 'inbox' or 'sent'
 */
function processConvExportBatch(folderType) {
    var batchSize = 30;
    var processed = 0;

    while (processed < batchSize && convExportState.currentIdx <= convExportState.maxItems) {
        var i = convExportState.currentIdx;
        convExportState.currentIdx++;
        processed++;

        try {
            var item = convExportState.currentItems.Item(i);
            if (item.Class !== 43) continue; // Only mail items

            // Check date
            var itemDate = folderType === 'sent' ?
                new Date(item.SentOn) : new Date(item.ReceivedTime);
            if (itemDate < convExportState.cutoffDate) continue;

            // Check for duplicates (by EntryID) - within this export
            var entryId = item.EntryID || '';
            if (convExportState.seenEntryIds[entryId]) continue;

            // Extract email (we'll filter at conversation level later)
            var email = extractEmailForConversation(item, folderType);

            // Mark if this email is already in Excel
            email.alreadyInExcel = (typeof existingEmailIds !== 'undefined' && existingEmailIds[entryId]) ? true : false;

            convExportState.allEmails.push(email);
            convExportState.seenEntryIds[entryId] = true;

        } catch (e) {}
    }

    // Update progress
    var phasePct = convExportState.phase === 1 ? 10 : 50;
    var phaseRange = convExportState.phase === 1 ? 40 : 40;
    var folderPct = Math.round((convExportState.currentIdx / convExportState.maxItems) * phaseRange);
    document.getElementById('exportProgressFill').style.width = (phasePct + folderPct) + '%';

    var phaseText = convExportState.phase === 1 ? 'Phase 1/4: Inbox' : 'Phase 2/4: Sent';
    document.getElementById('exportProgressText').innerText =
        phaseText + ' - ' + convExportState.currentIdx + '/' + convExportState.maxItems +
        ' (' + convExportState.allEmails.length + ' Emails)';

    // Continue or next phase
    if (convExportState.currentIdx <= convExportState.maxItems) {
        setTimeout(function() { processConvExportBatch(folderType); }, 10);
    } else {
        advanceConvExportPhase();
    }
}

/**
 * Advance to next export phase
 */
function advanceConvExportPhase() {
    convExportState.phase++;

    if (convExportState.phase === 2) {
        // Phase 2: Extract sent items
        document.getElementById('exportProgressText').innerText = 'Phase 2/4: Lade Sent Items...';
        document.getElementById('exportProgressFill').style.width = '50%';
        startFolderExtraction(convExportState.sentFolder, 'sent');

    } else if (convExportState.phase === 3) {
        // Phase 3: Build conversation hierarchy
        document.getElementById('exportProgressText').innerText = 'Phase 3/4: Baue Konversationen...';
        document.getElementById('exportProgressFill').style.width = '90%';

        setTimeout(function() {
            try {
                convExportState.conversations = buildConversationHierarchy(convExportState.allEmails);
                advanceConvExportPhase();
            } catch (e) {
                showExportError("Fehler beim Konversations-Aufbau: " + e.message);
            }
        }, 50);

    } else if (convExportState.phase === 4) {
        // Phase 4: Save export
        document.getElementById('exportProgressText').innerText = 'Phase 4/4: Speichere...';
        document.getElementById('exportProgressFill').style.width = '95%';

        setTimeout(saveConversationExport, 50);
    }
}

/**
 * Save the conversation export to JSON file
 */
function saveConversationExport() {
    try {
        // Filter conversations: keep only those with at least one NEW email
        var filteredConversations = {};
        var skippedConvCount = 0;
        var newEmailCount = 0;
        var totalEmailCount = 0;

        for (var convId in convExportState.conversations) {
            var conv = convExportState.conversations[convId];
            var hasNewEmail = false;
            var newInConv = 0;

            for (var mi = 0; mi < conv.messages.length; mi++) {
                totalEmailCount++;
                if (!conv.messages[mi].alreadyInExcel) {
                    hasNewEmail = true;
                    newInConv++;
                }
            }

            if (hasNewEmail) {
                filteredConversations[convId] = conv;
                newEmailCount += newInConv;
            } else {
                skippedConvCount++;
            }
        }

        // Check if there are any new conversations
        var filteredCount = Object.keys(filteredConversations).length;
        if (filteredCount === 0) {
            var noNewMsg = 'Keine neuen Emails gefunden.';
            if (skippedConvCount > 0) {
                noNewMsg += '<br><strong>' + totalEmailCount + '</strong> Emails in <strong>' + skippedConvCount + '</strong> Konversationen bereits in Excel.';
            }
            showExportSuccess(noNewMsg);
            document.getElementById('exportProgressFill').style.width = '100%';
            document.getElementById('exportProgressText').innerText = 'Keine neuen Emails';
            return;
        }

        var exportData = createExportStructure(
            filteredConversations,
            convExportState.mailboxName
        );

        var fso = new ActiveXObject("Scripting.FileSystemObject");
        var shell = new ActiveXObject("WScript.Shell");
        var shellApp = new ActiveXObject("Shell.Application");

        var downloads;
        try {
            downloads = shellApp.NameSpace("shell:Downloads").Self.Path;
        } catch (e) {
            downloads = shell.ExpandEnvironmentStrings("%USERPROFILE%") + "\\Downloads";
        }

        var today = new Date();
        var dateStr = today.getFullYear() + '-' +
            (today.getMonth() + 1 < 10 ? '0' : '') + (today.getMonth() + 1) + '-' +
            (today.getDate() < 10 ? '0' : '') + today.getDate();

        // Save conversation format
        var filename = downloads + "\\hypercare_conversations_" + dateStr + ".json";
        var file = fso.CreateTextFile(filename, true, true);
        file.Write(JSON.stringify(exportData, null, 2));
        file.Close();

        // Show success
        var successMsg = '<strong>' + newEmailCount + '</strong> neue Emails in <strong>' + filteredCount + '</strong> Konversationen';
        if (skippedConvCount > 0) {
            successMsg += '<br><strong>' + skippedConvCount + '</strong> Konversationen bereits vollstaendig in Excel';
        }
        successMsg += '<br>Gespeichert: ' + filename;

        showExportSuccess(successMsg);
        completeStep(3);

        document.getElementById('exportProgressFill').style.width = '100%';
        document.getElementById('exportProgressText').innerText = 'Export abgeschlossen!';

    } catch (e) {
        showExportError("Fehler beim Speichern: " + e.message);
    }
}

/**
 * Extract original inquiry from quoted text in sent item
 * @param {string} body - Email body with quotes
 * @returns {Object|null} Extracted inquiry info or null
 */
function extractQuotedInquiry(body) {
    if (!body) return null;

    // Pattern to find quoted email header (German format)
    // Von: Name <email>
    // Gesendet: Datum
    // An: ...
    // Betreff: ...
    var vonMatch = body.match(/Von:\s*([^\r\n<]+)(?:<([^>]+)>)?/i);
    var gesendetMatch = body.match(/Gesendet:\s*([^\r\n]+)/i);
    var betreffMatch = body.match(/Betreff:\s*([^\r\n]+)/i);

    if (!vonMatch) return null;

    var senderName = (vonMatch[1] || '').trim();
    var senderEmail = (vonMatch[2] || '').trim();

    // Find the position where the quoted content starts (after the header block)
    var headerEnd = body.indexOf('\r\n\r\n', body.indexOf('Betreff:'));
    if (headerEnd === -1) headerEnd = body.indexOf('\n\n', body.indexOf('Betreff:'));

    var quotedText = '';
    if (headerEnd > 0) {
        quotedText = body.substring(headerEnd).trim();
    }

    // Parse date from German format
    var originalDate = '';
    if (gesendetMatch && gesendetMatch[1]) {
        originalDate = parseGermanDateToISO(gesendetMatch[1].trim());
    }

    return {
        von_name: senderName,
        von_email: senderEmail || senderName,
        datum: originalDate,
        betreff: betreffMatch ? betreffMatch[1].trim() : '',
        text: quotedText
    };
}

/**
 * Parse German date format to ISO
 * @param {string} germanDate - e.g. "Dienstag, 2. Dezember 2025 18:00"
 * @returns {string} ISO date or original string
 */
function parseGermanDateToISO(germanDate) {
    var months = {
        'januar': '01', 'februar': '02', 'märz': '03', 'april': '04',
        'mai': '05', 'juni': '06', 'juli': '07', 'august': '08',
        'september': '09', 'oktober': '10', 'november': '11', 'dezember': '12'
    };

    // Match: [weekday,] day. month year [time]
    var match = germanDate.match(/(\d{1,2})\.\s*(\w+)\s+(\d{4})\s*(\d{1,2}:\d{2})?/i);
    if (match) {
        var day = match[1].padStart(2, '0');
        var monthName = match[2].toLowerCase();
        var year = match[3];
        var time = match[4] || '00:00';
        var month = months[monthName] || '01';
        return year + '-' + month + '-' + day + 'T' + time + ':00.000Z';
    }
    return germanDate;
}

/**
 * Convert conversation structure to legacy format for web app
 * @param {Object} conversations - Conversation structure
 * @returns {Array} Legacy email array with antworten
 */
function convertToLegacyFormat(conversations) {
    var result = [];

    for (var convId in conversations) {
        var conv = conversations[convId];
        var messages = conv.messages;

        if (messages.length === 0) continue;

        // Find root message (depth 0 or earliest)
        var rootMsg = null;
        var rootTime = null;
        var hasInboxRoot = false;

        for (var i = 0; i < messages.length; i++) {
            var msg = messages[i];
            var isInbox = msg.folder === 'inbox';
            var msgTime = msg.receivedTime || msg.sentOn;

            // Prefer inbox messages as root
            if (isInbox && msg.depth === 0) {
                if (!rootMsg || (msgTime && (!rootTime || msgTime < rootTime))) {
                    rootMsg = msg;
                    rootTime = msgTime;
                    hasInboxRoot = true;
                }
            }
        }

        // If no inbox root found, use earliest message
        if (!rootMsg) {
            for (var j = 0; j < messages.length; j++) {
                var msg2 = messages[j];
                var msgTime2 = msg2.receivedTime || msg2.sentOn;
                if (!rootMsg || (msgTime2 && (!rootTime || msgTime2 < rootTime))) {
                    rootMsg = msg2;
                    rootTime = msgTime2;
                }
            }
        }

        if (!rootMsg) continue;

        // Check if root is a sent item without real inbox root
        var rootIsSentReply = !hasInboxRoot && rootMsg.folder === 'sent';

        // If root is a sent reply, try to extract original inquiry from quoted text
        if (rootIsSentReply) {
            var quotedInquiry = extractQuotedInquiry(rootMsg.body || '');
            if (quotedInquiry && quotedInquiry.von_email) {
                // Build legacy email from extracted inquiry
                var legacyEmail = {
                    emailId: rootMsg.entryID,
                    conversationId: convId,
                    datum: quotedInquiry.datum || rootMsg.receivedTime || rootMsg.sentOn || '',
                    von_email: quotedInquiry.von_email,
                    von_name: quotedInquiry.von_name,
                    betreff: quotedInquiry.betreff || rootMsg.subject.replace(/^(AW:|RE:|WG:|FW:)\s*/gi, '').trim(),
                    text: quotedInquiry.text || '',
                    anhaenge: [],
                    kategorie: rootMsg.kategorie || '',
                    antworten: []
                };

                // Add the sent message as first reply
                legacyEmail.antworten.push({
                    datum: rootMsg.sentOn || rootMsg.receivedTime || '',
                    an: cleanX500Address(rootMsg.recipients[0] ? rootMsg.recipients[0].email : ''),
                    von: '',
                    text: removeEmailQuotes(rootMsg.body || ''),
                    replyId: rootMsg.entryID,
                    isIncoming: false,
                    depth: 1
                });

                // Add other messages as additional replies
                for (var k = 0; k < messages.length; k++) {
                    var reply = messages[k];
                    if (reply.entryID === rootMsg.entryID) continue;

                    var replyTime = reply.sentOn || reply.receivedTime || '';
                    var isSent = reply.folder === 'sent';

                    legacyEmail.antworten.push({
                        datum: replyTime,
                        an: isSent ? cleanX500Address(reply.recipients[0] ? reply.recipients[0].email : '') : '',
                        von: isSent ? '' : reply.senderName,
                        text: removeEmailQuotes(reply.body || reply.bodyPreview || ''),
                        replyId: reply.entryID,
                        isIncoming: !isSent,
                        depth: reply.depth
                    });
                }

                // Sort antworten by date
                legacyEmail.antworten.sort(function(a, b) {
                    return (a.datum || '') < (b.datum || '') ? -1 : 1;
                });

                result.push(legacyEmail);
                continue; // Skip normal processing
            }
        }

        // Normal processing: Build legacy email object
        var legacyEmail = {
            emailId: rootMsg.entryID,
            conversationId: convId,
            datum: rootMsg.receivedTime || rootMsg.sentOn || '',
            von_email: rootMsg.senderEmail,
            von_name: rootMsg.senderName,
            betreff: rootMsg.subject,
            text: removeEmailQuotes(rootMsg.body || ''),
            anhaenge: rootMsg.anhaenge || [],
            kategorie: rootMsg.kategorie || '',
            antworten: []
        };

        // Add other messages as antworten
        for (var k = 0; k < messages.length; k++) {
            var reply = messages[k];
            if (reply.entryID === rootMsg.entryID) continue;

            var replyTime = reply.sentOn || reply.receivedTime || '';
            var isSent = reply.folder === 'sent';

            legacyEmail.antworten.push({
                datum: replyTime,
                an: isSent ? cleanX500Address(reply.recipients[0] ? reply.recipients[0].email : '') : '',
                von: isSent ? '' : reply.senderName,
                text: removeEmailQuotes(reply.body || reply.bodyPreview || ''),
                replyId: reply.entryID,
                isIncoming: !isSent,
                depth: reply.depth
            });
        }

        // Sort antworten by date
        legacyEmail.antworten.sort(function(a, b) {
            return (a.datum || '') < (b.datum || '') ? -1 : 1;
        });

        result.push(legacyEmail);
    }

    // Sort by most recent activity
    result.sort(function(a, b) {
        var dateA = a.antworten.length > 0 ? a.antworten[a.antworten.length - 1].datum : a.datum;
        var dateB = b.antworten.length > 0 ? b.antworten[b.antworten.length - 1].datum : b.datum;
        return (dateB || '') > (dateA || '') ? 1 : -1;
    });

    return result;
}
