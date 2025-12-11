/**
 * Sent Items Cache Module
 * Cache and index sent items for efficient reply matching
 */

// Note: sentItemsCache, sentItemsByConvId, sentItemsByInReplyTo,
// sentItemsByReference, sentItemsByNormSubject are declared in config.js

// Sent cache state for async processing
// Note: sentCacheState is declared in config.js

/**
 * Start caching sent items
 * @param {number} days - Number of days to look back
 */
function cacheSentItems(days) {
    sentItemsCache = [];
    sentItemsByConvId = {};
    sentItemsByInReplyTo = {};
    sentItemsByReference = {};
    sentItemsByNormSubject = {};
    sentCacheState.days = days;
    sentCacheState.cutoffDate = exportState.cutoffDate;
    sentCacheState.folders = [];
    sentCacheState.folderIdx = 0;

    // Only search sent folder of selected mailbox
    try { sentCacheState.folders.push(selectedStore.GetDefaultFolder(5)); } catch (e) {}

    // Start first folder
    startNextSentFolder();
}

/**
 * Start processing the next sent folder
 */
function startNextSentFolder() {
    if (sentCacheState.folderIdx >= sentCacheState.folders.length) {
        // All folders done - start Phase 3
        document.getElementById('exportProgressText').innerText =
            'Phase 2/3 fertig: ' + sentItemsCache.length + ' gesendete Emails aus ' + sentCacheState.folders.length + ' Ordnern geladen';
        document.getElementById('exportProgressFill').style.width = '30%';
        setTimeout(processEmailBatch, 10);
        return;
    }

    try {
        var folder = sentCacheState.folders[sentCacheState.folderIdx];

        // Use Restrict() to filter sent items by date
        var dateFilter = "[SentOn] >= '" + formatDateForRestrict(sentCacheState.cutoffDate) + "'";
        try {
            sentCacheState.items = folder.Items.Restrict(dateFilter);
        } catch (restrictErr) {
            sentCacheState.items = folder.Items;
        }
        try { sentCacheState.items.Sort("[SentOn]", true); } catch (e) {}
        sentCacheState.maxItems = sentCacheState.items.Count;
        sentCacheState.itemIdx = 1;
        setTimeout(processSentBatch, 10);
    } catch (e) {
        sentCacheState.folderIdx++;
        setTimeout(startNextSentFolder, 10);
    }
}

/**
 * Process a batch of sent items
 */
function processSentBatch() {
    var batchSize = 50;
    var processed = 0;

    while (processed < batchSize && sentCacheState.itemIdx <= sentCacheState.maxItems) {
        var i = sentCacheState.itemIdx;
        sentCacheState.itemIdx++;
        processed++;

        try {
            var item = sentCacheState.items.Item(i);
            if (item.Class !== 43) continue;

            var sentDate = new Date(item.SentOn);
            if (sentDate < sentCacheState.cutoffDate) {
                sentCacheState.itemIdx = sentCacheState.maxItems + 1;
                break;
            }

            var body = '';
            try { body = item.Body || ''; } catch (e) {}

            var htmlBody = '';
            try { htmlBody = item.HTMLBody || ''; if (htmlBody.length > 500000) htmlBody = ''; } catch (e) {}

            var toRecipient = '';
            try { toRecipient = item.To || ''; } catch (e) {}

            var convId = '';
            try { convId = item.ConversationID || ''; } catch (e) {}

            var entryId = '';
            try { entryId = item.EntryID || ''; } catch (e) {}

            // Extract extended headers
            var headers = getEmailHeaders(item);

            var subject = '';
            try { subject = item.Subject || ''; } catch (e) {}

            var sentItem = {
                sentDate: sentDate,
                conversationId: convId,
                entryId: entryId,
                body: removeEmailQuotes(body),
                htmlBody: htmlBody,
                to: toRecipient,
                subject: subject,
                internetMessageId: headers.internetMessageId,
                inReplyTo: headers.inReplyTo,
                references: headers.references,
                normalizedSubject: normalizeSubject(subject)
            };
            sentItemsCache.push(sentItem);

            // Index by ConversationID
            if (convId) {
                if (!sentItemsByConvId[convId]) {
                    sentItemsByConvId[convId] = [];
                }
                sentItemsByConvId[convId].push(sentItem);
            }

            // Index by In-Reply-To
            if (headers.inReplyTo) {
                if (!sentItemsByInReplyTo[headers.inReplyTo]) {
                    sentItemsByInReplyTo[headers.inReplyTo] = [];
                }
                sentItemsByInReplyTo[headers.inReplyTo].push(sentItem);
            }

            // Index by References
            if (headers.references && headers.references.length > 0) {
                for (var ri = 0; ri < headers.references.length; ri++) {
                    var refId = headers.references[ri];
                    if (!sentItemsByReference[refId]) {
                        sentItemsByReference[refId] = [];
                    }
                    sentItemsByReference[refId].push(sentItem);
                }
            }
        } catch (e) {}
    }

    // Update progress
    var totalFolders = sentCacheState.folders.length;
    var pct = 10 + Math.round(((sentCacheState.folderIdx * 1000 + sentCacheState.itemIdx) / (totalFolders * 1000)) * 20);
    document.getElementById('exportProgressFill').style.width = pct + '%';
    var folderName = '';
    try { folderName = sentCacheState.folders[sentCacheState.folderIdx].Name; } catch (e) { folderName = 'Ordner ' + (sentCacheState.folderIdx + 1); }
    document.getElementById('exportProgressText').innerText =
        'Phase 2/3: ' + folderName + ' (' + (sentCacheState.folderIdx + 1) + '/' + totalFolders + ')' +
        ' - Email ' + sentCacheState.itemIdx + '/' + sentCacheState.maxItems +
        ' (' + sentItemsCache.length + ' geladen)';

    // Continue or next folder
    if (sentCacheState.itemIdx <= sentCacheState.maxItems) {
        setTimeout(processSentBatch, 10);
    } else {
        sentCacheState.folderIdx++;
        setTimeout(startNextSentFolder, 10);
    }
}

