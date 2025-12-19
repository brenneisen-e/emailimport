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
    document.getElementById('exportProgressText').innerText = 'Synchronisiere Outlook...';
    document.getElementById('exportProgressFill').style.width = '2%';

    // Force Outlook sync before starting export
    try {
        // Try to trigger send/receive for all accounts
        var syncObjects = outlookNS.SyncObjects;
        if (syncObjects && syncObjects.Count > 0) {
            for (var i = 1; i <= syncObjects.Count; i++) {
                try {
                    syncObjects.Item(i).Start();
                } catch (e) {}
            }
        }
    } catch (e) {}

    // Wait for sync to complete, then start export
    setTimeout(function() {
        document.getElementById('exportProgressText').innerText = 'Phase 1/4: Lade Inbox...';
        document.getElementById('exportProgressFill').style.width = '5%';
        convExportState.phase = 1;
        startFolderExtraction(convExportState.inboxFolder, 'inbox');
    }, 1500);  // 1.5 seconds to let sync start
}

/**
 * Start extracting from a folder
 * @param {Object} folder - Outlook folder
 * @param {string} folderType - 'inbox' or 'sent'
 */
function startFolderExtraction(folder, folderType) {
    try {
        convExportState.currentFolder = folder;

        // Force refresh of folder items (helps with Outlook caching issues)
        try {
            // Access GetTable to trigger internal refresh - this forces Outlook to update
            var table = folder.GetTable();
            // Read first row to force actual data access
            if (table && !table.EndOfTable) {
                try { table.GetNextRow(); } catch (e) {}
            }
            table = null;
        } catch (e) {}

        // Delay to let Outlook fully sync folder data
        setTimeout(function() {
            continueExtraction(folder, folderType);
        }, 500);

    } catch (e) {
        // Skip to next phase if folder fails
        advanceConvExportPhase();
    }
}

/**
 * Continue extraction after sync delay
 */
function continueExtraction(folder, folderType) {
    try {
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

        convExportState.maxItems = Math.min(convExportState.currentItems.Count, 2000);
        convExportState.currentIdx = 1;

        setTimeout(function() { processConvExportBatch(folderType); }, 10);

    } catch (e) {
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
        var totalEmailCount = 0;
        var brandNewConvCount = 0;  // Completely new conversations
        var updatedConvCount = 0;   // Existing conversations with new replies
        var newRepliesInUpdated = 0; // Actual count of new replies in existing conversations

        for (var convId in convExportState.conversations) {
            var conv = convExportState.conversations[convId];
            var hasNewEmail = false;
            var hasExistingEmail = false;  // Does any email already exist in Excel?
            var newInConv = 0;

            for (var mi = 0; mi < conv.messages.length; mi++) {
                totalEmailCount++;
                if (conv.messages[mi].alreadyInExcel) {
                    hasExistingEmail = true;  // Conversation exists in Excel
                } else {
                    hasNewEmail = true;
                    newInConv++;
                }
            }

            if (hasNewEmail) {
                filteredConversations[convId] = conv;
                // Track: is this completely new or existing with updates?
                if (hasExistingEmail) {
                    // At least one message existed in Excel = existing conversation with updates
                    updatedConvCount++;
                    newRepliesInUpdated += newInConv;
                } else {
                    // No messages existed in Excel = brand new conversation
                    brandNewConvCount++;
                }
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

        // Get final counts after sent-only/online-abschluss filtering
        var finalConvCount = exportData.conversationCount;
        var filterInfo = exportData._filterInfo || {};

        // Check if all conversations were filtered out
        if (finalConvCount === 0) {
            var noNewMsg = 'Keine neuen Vorg&auml;nge gefunden.';
            showExportSuccess(noNewMsg);
            document.getElementById('exportProgressFill').style.width = '100%';
            document.getElementById('exportProgressText').innerText = 'Keine relevanten Vorgaenge';
            return;
        }

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

        // Show success message with detailed counts
        var msgParts = [];

        // New conversations (completely new)
        if (brandNewConvCount > 0) {
            msgParts.push('<strong>' + brandNewConvCount + '</strong> neue' + (brandNewConvCount > 1 ? ' Vorg&auml;nge' : 'r Vorgang'));
        }

        // Existing conversations with new replies
        if (updatedConvCount > 0 && newRepliesInUpdated > 0) {
            msgParts.push('<strong>' + newRepliesInUpdated + '</strong> neue Antwort' + (newRepliesInUpdated > 1 ? 'en' : '') + ' in <strong>' + updatedConvCount + '</strong> bestehende' + (updatedConvCount > 1 ? 'n Vorg&auml;ngen' : 'm Vorgang'));
        }

        var successMsg = msgParts.join('<br>');

        if (skippedConvCount > 0) {
            successMsg += '<br>' + skippedConvCount + ' Vorg' + (skippedConvCount > 1 ? '&auml;nge' : 'ang') + ' bereits vollst&auml;ndig in Excel';
        }

        showExportSuccess(successMsg);
        completeStep(3);

        document.getElementById('exportProgressFill').style.width = '100%';
        document.getElementById('exportProgressText').innerText = 'Export abgeschlossen!';

    } catch (e) {
        showExportError("Fehler beim Speichern: " + e.message);
    }
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

        // Check if this is a sent-only conversation (no inbox root)
        // These should only be used to add replies to EXISTING cases, not create new ones
        var sentOnlyConversation = !hasInboxRoot;

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
            sentOnly: sentOnlyConversation,  // Flag for import to handle
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
