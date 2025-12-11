/**
 * Export Core Module
 * Main export logic and batch processing
 */

// Note: exportState is declared in config.js

// Last exported emails for direct import
var lastExportedEmails = null;

/**
 * Check if export can be started (folder selected AND Excel connected)
 */
function checkExportReady() {
    var folderSelected = document.getElementById('folderSelect').value !== '';
    var excelConnected = exportWorkbook !== null;
    var btn = document.getElementById('exportBtn');

    if (folderSelected && excelConnected) {
        btn.disabled = false;
        btn.innerHTML = '&#128229; Emails exportieren';
    } else if (folderSelected && !excelConnected) {
        btn.disabled = true;
        btn.innerHTML = '&#128229; Erst Excel verbinden!';
    } else {
        btn.disabled = true;
        btn.innerHTML = '&#128229; Emails exportieren';
    }
}

/**
 * Start the export process
 */
function startExport() {
    var folderId = document.getElementById('folderSelect').value;
    var days = parseInt(document.getElementById('daysInput').value) || 7;

    if (!folderId) { alert('Bitte Ordner auswahlen!'); return; }

    document.getElementById('exportProgress').style.display = 'block';
    document.getElementById('exportBtn').disabled = true;
    document.getElementById('exportSuccess').style.display = 'none';
    document.getElementById('exportError').style.display = 'none';
    document.getElementById('exportActions').style.display = 'none';

    setTimeout(function() { doExport(folderId, days); }, 100);
}

/**
 * Execute the export
 * @param {string} folderId - Folder ID or 'inbox'
 * @param {number} days - Number of days to export
 */
function doExport(folderId, days) {
    try {
        // Phase 1: Get folder
        document.getElementById('exportProgressText').innerText = 'Phase 1/3: Ordner wird geladen...';
        document.getElementById('exportProgressFill').style.width = '5%';

        var folder = (folderId === 'inbox') ? selectedStore.GetDefaultFolder(6) : outlookNS.GetFolderFromID(folderId);

        // Calculate cutoff date
        exportState.cutoffDate = new Date();
        exportState.cutoffDate.setDate(exportState.cutoffDate.getDate() - days);

        // Use Restrict() to filter emails
        var dateFilter = "[ReceivedTime] >= '" + formatDateForRestrict(exportState.cutoffDate) + "'";
        try {
            exportState.items = folder.Items.Restrict(dateFilter);
        } catch (restrictErr) {
            exportState.items = folder.Items;
        }
        try { exportState.items.Sort("[ReceivedTime]", true); } catch (e) {}

        exportState.emails = [];
        exportState.total = Math.min(exportState.items.Count, 500);
        exportState.current = 1;
        exportState.oldCount = 0;
        exportState.repliesFound = 0;
        exportState.duplicatesSkipped = 0;
        exportState.batchDuplicates = 0;
        exportState.duplicatesWithReplies = 0;
        exportState.processedConvIds = {};

        // Phase 2: Cache sent items
        document.getElementById('exportProgressText').innerText = 'Phase 2/3: Lade gesendete Emails...';
        document.getElementById('exportProgressFill').style.width = '10%';

        setTimeout(function() {
            cacheSentItems(days);
        }, 50);

    } catch (e) {
        showExportError("Fehler: " + e.message);
    }
}

/**
 * Process a batch of emails
 */
function processEmailBatch() {
    var batchSize = 20;
    var processed = 0;

    while (processed < batchSize && exportState.current <= exportState.total) {
        var i = exportState.current;
        exportState.current++;
        processed++;

        try {
            var item = exportState.items.Item(i);
            if (item.Class !== 43) continue;
            var receivedTime = new Date(item.ReceivedTime);
            if (receivedTime < exportState.cutoffDate) {
                exportState.oldCount++;
                if (exportState.oldCount >= 20) {
                    exportState.current = exportState.total + 1;
                    break;
                }
                continue;
            }

            var emailData = extractEmail(item);

            // Filter out specific subjects
            var subjectLower = (emailData.betreff || '').toLowerCase();
            if (subjectLower.indexOf('online-abschluss vermittlerzuordnung ändern') !== -1) {
                continue;
            }

            // Mark as found in Inbox
            if (emailData.conversationId) {
                foundInInbox[emailData.conversationId] = true;
            }

            // Process email with duplicate and reply checking
            processExportEmail(emailData);

        } catch (e) {}
    }

    // Update UI
    var pct = 30 + Math.round((exportState.current / exportState.total) * 70);
    document.getElementById('exportProgressFill').style.width = pct + '%';
    var newCount = exportState.emails.length - exportState.duplicatesWithReplies;
    var statusText = 'Phase 3/3: Email ' + exportState.current + '/' + exportState.total +
        ' (' + newCount + ' neu';
    if (exportState.duplicatesWithReplies > 0) {
        statusText += ', ' + exportState.duplicatesWithReplies + ' Updates';
    }
    if (exportState.duplicatesSkipped > 0) {
        statusText += ', ' + exportState.duplicatesSkipped + ' in Excel';
    }
    if (exportState.batchDuplicates > 0) {
        statusText += ', ' + exportState.batchDuplicates + ' doppelt';
    }
    statusText += ', ' + exportState.repliesFound + ' Antworten)';
    document.getElementById('exportProgressText').innerText = statusText;

    // Continue or finish
    if (exportState.current <= exportState.total) {
        setTimeout(processEmailBatch, 10);
    } else {
        finishExport();
    }
}

/**
 * Process a single email for export
 * @param {Object} emailData - Email data object
 */
function processExportEmail(emailData) {
    // Check for duplicates
    var hasExcelData = false;
    for (var k in existingEmailKeys) { hasExcelData = true; break; }
    var isDuplicate = false;
    var isConversationContinuation = false;
    var existingReplyText = '';
    var existingReplyIdList = '';
    var convKey = 'conv|' + (emailData.conversationId || '');
    var emailKey = createExportEmailKey(emailData.datum, emailData.betreff);

    if (hasExcelData) {
        if (emailData.conversationId && existingEmailKeys[convKey]) {
            if (emailData.emailId && existingEmailIds[emailData.emailId]) {
                isDuplicate = true;
            } else {
                isConversationContinuation = true;
            }
            existingReplyText = existingReplies[convKey] || '';
            existingReplyIdList = existingReplyIds[convKey] || '';
        } else if (existingEmailKeys[emailKey]) {
            isDuplicate = true;
            existingReplyText = existingReplies[emailKey] || '';
            existingReplyIdList = existingReplyIds[emailKey] || '';
        }
    }

    // Find replies from sent items
    var allReplies = findReplies(emailData);
    exportState.repliesFound += allReplies.length;

    if (isDuplicate) {
        handleDuplicateEmail(emailData, allReplies, existingReplyText, existingReplyIdList);
    } else if (isConversationContinuation) {
        handleConversationContinuation(emailData, allReplies, existingReplyText, existingReplyIdList, convKey, emailKey);
    } else {
        handleNewEmail(emailData, allReplies);
    }
}

/**
 * Handle duplicate email (check for new replies only)
 */
function handleDuplicateEmail(emailData, allReplies, existingReplyText, existingReplyIdList) {
    var newReplies = filterNewReplies(allReplies, existingReplyText, existingReplyIdList);

    if (newReplies.length > 0) {
        emailData.antworten = newReplies;
        emailData.alleAntworten = allReplies;
        emailData.existingReplies = existingReplyText;
        emailData.existingReplyIds = existingReplyIdList;
        emailData.isDuplicateWithNewReplies = true;
        exportState.duplicatesWithReplies++;

        // Clean email text
        if (emailData.htmlBody) {
            var contentResult = extractNewContentFromHtml(emailData.htmlBody);
            emailData.text = contentResult.newContent || removeEmailQuotes(emailData.text);
        } else {
            emailData.text = removeEmailQuotes(emailData.text);
        }
        exportState.emails.push(emailData);
    } else {
        exportState.duplicatesSkipped++;
    }
}

/**
 * Handle conversation continuation (new email in existing conversation)
 */
function handleConversationContinuation(emailData, allReplies, existingReplyText, existingReplyIdList, convKey, emailKey) {
    var originalRawText = emailData.text;

    // Extract new content
    var incomingReplyText = '';
    if (emailData.htmlBody) {
        var contentResult = extractNewContentFromHtml(emailData.htmlBody);
        incomingReplyText = contentResult.newContent || removeEmailQuotes(emailData.text);
    } else {
        incomingReplyText = removeEmailQuotes(emailData.text);
    }

    // Create incoming reply object
    var incomingReply = {
        datum: emailData.datum,
        an: '',
        text: incomingReplyText,
        fullText: originalRawText,
        replyId: emailData.emailId,
        internetMessageId: emailData.internetMessageId || '',
        inReplyTo: emailData.inReplyTo || '',
        isNew: true,
        isIncoming: true,
        threadDepth: 1,
        threadPosition: 0
    };

    // Combine with sent replies
    var allNewReplies = [incomingReply];
    var filteredSentReplies = filterNewReplies(allReplies, existingReplyText, existingReplyIdList);
    for (var sr = 0; sr < filteredSentReplies.length; sr++) {
        allNewReplies.push(filteredSentReplies[sr]);
    }

    // Use original request text from Excel
    var originalAnfrage = existingAnfragen[convKey] || existingAnfragen[emailKey] || '';
    emailData.text = originalAnfrage || removeEmailQuotes(originalRawText);

    // Only override datum for root emails (without inReplyTo)
    var isLikelyRootCC = !emailData.inReplyTo;
    if (isLikelyRootCC) {
        var originalDate = existingDates[convKey] || existingDates[emailKey] || '';
        if (originalDate) {
            emailData.datum = originalDate;
        } else if (emailData.conversationStartDate) {
            emailData.datum = emailData.conversationStartDate;
        } else {
            var quotedDate = extractOldestDateFromQuotes(originalRawText);
            if (quotedDate) {
                emailData.datum = formatDate(quotedDate);
            }
        }
    }

    // Use original subject
    var originalSubject = existingSubjects[convKey] || existingSubjects[emailKey] || '';
    if (originalSubject) {
        emailData.betreff = originalSubject;
    } else if (emailData.originalSubject) {
        emailData.betreff = emailData.originalSubject;
    } else {
        emailData.betreff = stripSubjectPrefixes(emailData.betreff);
    }

    emailData.antworten = allNewReplies;
    emailData.alleAntworten = allReplies;
    emailData.existingReplies = existingReplyText;
    emailData.existingReplyIds = existingReplyIdList;
    emailData.isDuplicateWithNewReplies = true;
    emailData.isConversationContinuation = true;
    exportState.duplicatesWithReplies++;
    exportState.emails.push(emailData);
}

/**
 * Handle completely new email
 */
function handleNewEmail(emailData, allReplies) {
    var inboxDate = new Date(emailData.datum);
    var oldestSentItem = null;
    var oldestSentDate = null;

    // Find oldest sent item
    for (var si = 0; si < allReplies.length; si++) {
        var replyDate = new Date(allReplies[si].datum);
        if (!oldestSentDate || replyDate < oldestSentDate) {
            oldestSentDate = replyDate;
            oldestSentItem = allReplies[si];
        }
    }

    // Check if sent item started the conversation
    if (oldestSentItem && oldestSentDate < inboxDate) {
        // Extract content from inbox email
        var inboxReplyText = '';
        if (emailData.htmlBody) {
            var contentResult = extractNewContentFromHtml(emailData.htmlBody);
            inboxReplyText = contentResult.newContent || removeEmailQuotes(emailData.text);
        } else {
            inboxReplyText = removeEmailQuotes(emailData.text);
        }

        // Inbox email becomes a reply
        var inboxAsReply = {
            datum: emailData.datum,
            an: '',
            text: inboxReplyText,
            fullText: emailData.text,
            replyId: emailData.emailId,
            internetMessageId: emailData.internetMessageId || '',
            inReplyTo: emailData.inReplyTo || '',
            isNew: true,
            isIncoming: true,
            threadDepth: 1,
            threadPosition: 0
        };

        // Build replies without oldest sent item
        var repliesWithoutOldest = [];
        for (var ri = 0; ri < allReplies.length; ri++) {
            if (allReplies[ri] !== oldestSentItem) {
                allReplies[ri].isNew = true;
                repliesWithoutOldest.push(allReplies[ri]);
            }
        }
        repliesWithoutOldest.push(inboxAsReply);

        // Use sent item as Anfrage
        emailData.text = oldestSentItem.text;
        emailData.datum = oldestSentItem.datum;
        emailData.antworten = repliesWithoutOldest;

        if (oldestSentItem.to) {
            emailData.von_name = 'An: ' + oldestSentItem.to;
            emailData.von_email = '';
        }
    } else {
        // Inbox email is the oldest - normal case
        emailData.antworten = allReplies;
        for (var r = 0; r < allReplies.length; r++) {
            allReplies[r].isNew = true;
        }
    }

    // Only override datum for root emails
    var isLikelyRoot = !emailData.inReplyTo;
    if (isLikelyRoot) {
        if (emailData.conversationStartDate) {
            var convStartDate = new Date(emailData.conversationStartDate);
            var currentDate = new Date(emailData.datum);
            if (convStartDate < currentDate) {
                emailData.datum = emailData.conversationStartDate;
            }
        } else if (!oldestSentItem || oldestSentDate >= inboxDate) {
            var quotedDate = extractOldestDateFromQuotes(emailData.text);
            if (quotedDate) {
                emailData.datum = formatDate(quotedDate);
            }
        }
    }

    // Use original subject
    if (emailData.originalSubject) {
        emailData.betreff = emailData.originalSubject;
    } else {
        emailData.betreff = stripSubjectPrefixes(emailData.betreff);
    }

    // Clean email text
    if (emailData.htmlBody) {
        var textResult = extractNewContentFromHtml(emailData.htmlBody);
        emailData.text = textResult.newContent || removeEmailQuotes(emailData.text);
    } else {
        emailData.text = removeEmailQuotes(emailData.text);
    }

    exportState.emails.push(emailData);
}

/**
 * Finish the export process
 */
function finishExport() {
    if (exportState.emails.length === 0) {
        showExportError("Keine neuen Emails gefunden.");
        return;
    }

    // Use unified conversation processing
    var groupedEmails = groupEmailsByConversationUnifiedMode(exportState.emails);

    saveExportFile(groupedEmails, exportState.repliesFound, exportState.duplicatesSkipped, exportState.duplicatesWithReplies, exportState.batchDuplicates);
}

/**
 * Group emails by conversation using unified mode
 * @param {Array} inboxEmails - Array of inbox emails
 * @returns {Array} Grouped and processed conversations
 */
function groupEmailsByConversationUnifiedMode(inboxEmails) {
    var result = [];

    // Convert inbox emails to unified format
    var allEmails = [];
    var addedReplyIds = {};

    for (var i = 0; i < inboxEmails.length; i++) {
        var inboxEmail = inboxEmails[i];
        var unifiedEmail = convertToUnifiedFormat(inboxEmail, 'inbox');

        // IMPORTANT: Preserve original antworten for fallback
        unifiedEmail.originalAntworten = inboxEmail.antworten || [];
        unifiedEmail.originalVonName = inboxEmail.von_name || '';

        allEmails.push(unifiedEmail);

        // Add ALL pre-existing replies (both incoming AND outgoing/sent)
        if (inboxEmail.antworten && inboxEmail.antworten.length > 0) {
            for (var ai = 0; ai < inboxEmail.antworten.length; ai++) {
                var reply = inboxEmail.antworten[ai];
                var replyId = reply.replyId || reply.internetMessageId || ('reply_' + ai);
                if (addedReplyIds[replyId]) continue;
                addedReplyIds[replyId] = true;

                var unifiedReply = {
                    entryId: reply.replyId || '',
                    conversationId: inboxEmail.conversationId || '',
                    internetMessageId: reply.internetMessageId || '',
                    inReplyTo: reply.inReplyTo || '',
                    references: [],
                    receivedTime: reply.datum ? new Date(reply.datum) : new Date(),
                    sentOn: reply.datum ? new Date(reply.datum) : new Date(),
                    timestamp: reply.datum ? new Date(reply.datum) : new Date(),
                    von_email: reply.vonEmail || '',
                    von_name: reply.von || reply.an || '',  // Use 'an' for outgoing
                    recipients: reply.an ? [{name: reply.an, email: '', type: 1}] : [],
                    betreff: inboxEmail.betreff || '',
                    text: reply.text || '',
                    htmlBody: reply.htmlBody || '',
                    anhaenge: [],
                    outlookKategorie: '',
                    folderType: reply.isIncoming ? 'inbox' : 'sent',
                    folderName: '',
                    isIncoming: reply.isIncoming !== false,  // Default to true if not specified
                    threadPosition: 0,
                    threadDepth: reply.threadDepth || 1,
                    parentMessageId: reply.parentMessageId || reply.inReplyTo || null,
                    parentFoundInThread: false,
                    isThreadRoot: false,
                    newContent: reply.text || '',
                    normalizedSubject: inboxEmail.normalizedSubject || ''
                };
                allEmails.push(unifiedReply);
            }
        }
    }

    // Add sent items
    var sentByConvId = {};
    for (var j = 0; j < sentItemsCache.length; j++) {
        var sentItem = sentItemsCache[j];
        var convId = sentItem.conversationId || '';
        if (convId) {
            if (!sentByConvId[convId]) {
                sentByConvId[convId] = [];
            }
            sentByConvId[convId].push(sentItem);
        }
    }

    var addedSentIds = {};
    for (var k = 0; k < inboxEmails.length; k++) {
        var email = inboxEmails[k];
        var convId = email.conversationId || '';
        if (convId && sentByConvId[convId]) {
            var matchingSent = sentByConvId[convId];
            for (var m = 0; m < matchingSent.length; m++) {
                var sent = matchingSent[m];
                if (addedSentIds[sent.entryId]) continue;
                addedSentIds[sent.entryId] = true;
                var unifiedSent = convertSentItemToUnifiedFormat(sent);
                allEmails.push(unifiedSent);
            }
        }
    }

    // Group by ConversationID
    var conversations = groupByConversation(allEmails);

    // Process each conversation
    for (var convId in conversations) {
        var conversationEmails = conversations[convId];
        if (conversationEmails.length === 0) continue;

        var conversationJson = processConversationUnified(conversationEmails);

        if (!conversationJson.normalizedSubject) {
            conversationJson.normalizedSubject = normalizeSubject(conversationJson.betreff || '');
        }

        result.push(conversationJson);
    }

    // Sort by last activity date
    result.sort(function(a, b) {
        var dateA = new Date(a.lastActivityDate || a.datum);
        var dateB = new Date(b.lastActivityDate || b.datum);
        return dateB - dateA;
    });

    return result;
}

