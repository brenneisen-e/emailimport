/**
 * Sent Items Cache Module
 * Cache and index sent items for efficient reply matching
 */

// Note: sentItemsCache, sentItemsByConvId, sentItemsByInReplyTo,
// sentItemsByReference, sentItemsByNormSubject are declared in config.js

// Sent cache state for async processing
// Note: sentCacheState is declared in config.js

// Sent cache statistics
var sentCacheStats = {
    totalCached: 0,
    withConvId: 0,
    withInReplyTo: 0,
    withSubject: 0,
    withoutAnyIndex: 0
};

/**
 * Normalize subject for indexing (same logic as matching.js)
 * @param {string} subject - Original subject
 * @returns {string} Normalized subject
 */
function normalizeSubjectForIndex(subject) {
    if (!subject) return '';
    var normalized = subject.replace(/^(re:|aw:|fwd:|wg:|fw:)\s*/gi, '');
    while (/^(re:|aw:|fwd:|wg:|fw:)\s*/i.test(normalized)) {
        normalized = normalized.replace(/^(re:|aw:|fwd:|wg:|fw:)\s*/gi, '');
    }
    normalized = normalized.replace(/\s+/g, ' ').trim().toLowerCase();
    return normalized;
}

/**
 * Start caching sent items with multiple indexes
 * @param {number} days - Number of days to look back
 */
function cacheSentItems(days) {
    // Initialize all indexes
    sentItemsCache = [];
    sentItemsByConvId = {};
    sentItemsByInReplyTo = {};
    sentItemsByNormSubject = {};

    // Reset stats
    sentCacheStats.totalCached = 0;
    sentCacheStats.withConvId = 0;
    sentCacheStats.withInReplyTo = 0;
    sentCacheStats.withSubject = 0;
    sentCacheStats.withoutAnyIndex = 0;

    sentCacheState.days = days;
    // Search sent items further back than inbox (replies might be older)
    sentCacheState.cutoffDate = new Date();
    sentCacheState.cutoffDate.setDate(sentCacheState.cutoffDate.getDate() - (days + 30));
    sentCacheState.folders = [];
    sentCacheState.folderIdx = 0;

    // Collect ALL sent folders (like the working old version)
    // 1. Selected store's sent folder
    try { sentCacheState.folders.push(selectedStore.GetDefaultFolder(5)); } catch (e) {}

    // 2. Default namespace sent folder
    try {
        var defaultSent = outlookNS.GetDefaultFolder(5);
        var isDup = false;
        for (var f = 0; f < sentCacheState.folders.length; f++) {
            try { if (sentCacheState.folders[f].EntryID === defaultSent.EntryID) isDup = true; } catch (e) {}
        }
        if (!isDup) sentCacheState.folders.push(defaultSent);
    } catch (e) {}

    // 3. All other stores' sent folders
    try {
        for (var s = 1; s <= outlookNS.Stores.Count; s++) {
            try {
                var storeSent = outlookNS.Stores.Item(s).GetDefaultFolder(5);
                var isDup2 = false;
                for (var f2 = 0; f2 < sentCacheState.folders.length; f2++) {
                    try { if (sentCacheState.folders[f2].EntryID === storeSent.EntryID) isDup2 = true; } catch (e) {}
                }
                if (!isDup2) sentCacheState.folders.push(storeSent);
            } catch (e) {}
        }
    } catch (e) {}

    // Start first folder
    startNextSentFolder();
}

/**
 * Start processing the next sent folder
 */
function startNextSentFolder() {
    if (sentCacheState.folderIdx >= sentCacheState.folders.length) {
        // All folders done - show stats and start Phase 3
        var statsMsg = 'Phase 2/3 fertig: ' + sentItemsCache.length + ' gesendete Emails geladen';
        statsMsg += ' (ConvId: ' + sentCacheStats.withConvId;
        statsMsg += ', InReplyTo: ' + sentCacheStats.withInReplyTo;
        statsMsg += ', Subject: ' + sentCacheStats.withSubject + ')';
        document.getElementById('exportProgressText').innerText = statsMsg;
        document.getElementById('exportProgressFill').style.width = '30%';

        // Reset matching stats before Phase 3
        if (typeof resetMatchingStats === 'function') {
            resetMatchingStats();
        }

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

            // Extract all relevant fields
            var body = '';
            try { body = item.Body || ''; } catch (e) {}

            var toRecipient = '';
            try { toRecipient = item.To || ''; } catch (e) {}

            var convId = '';
            try { convId = item.ConversationID || ''; } catch (e) {}

            var entryId = '';
            try { entryId = item.EntryID || ''; } catch (e) {}

            var subject = '';
            try { subject = item.Subject || ''; } catch (e) {}

            // Get MAPI headers for InReplyTo matching
            var inReplyTo = '';
            var internetMsgId = '';
            try {
                var headers = getEmailHeaders(item);
                inReplyTo = headers.inReplyTo || '';
                internetMsgId = headers.internetMessageId || '';
            } catch (e) {}

            // Create sent item with all matching fields
            var sentItem = {
                sentDate: sentDate,
                conversationId: convId,
                entryId: entryId,
                body: removeEmailQuotes(body),
                to: toRecipient,
                subject: subject,
                normalizedSubject: normalizeSubjectForIndex(subject),
                inReplyTo: inReplyTo,
                internetMessageId: internetMsgId
            };
            sentItemsCache.push(sentItem);
            sentCacheStats.totalCached++;

            // Track indexing for stats
            var hasAnyIndex = false;

            // Index by ConversationID (primary)
            if (convId) {
                if (!sentItemsByConvId[convId]) {
                    sentItemsByConvId[convId] = [];
                }
                sentItemsByConvId[convId].push(sentItem);
                sentCacheStats.withConvId++;
                hasAnyIndex = true;
            }

            // Index by InReplyTo (secondary - for cross-client threads)
            if (inReplyTo) {
                if (!sentItemsByInReplyTo[inReplyTo]) {
                    sentItemsByInReplyTo[inReplyTo] = [];
                }
                sentItemsByInReplyTo[inReplyTo].push(sentItem);
                sentCacheStats.withInReplyTo++;
                hasAnyIndex = true;
            }

            // Index by normalized subject (fallback)
            if (sentItem.normalizedSubject && sentItem.normalizedSubject.length > 5) {
                if (!sentItemsByNormSubject[sentItem.normalizedSubject]) {
                    sentItemsByNormSubject[sentItem.normalizedSubject] = [];
                }
                sentItemsByNormSubject[sentItem.normalizedSubject].push(sentItem);
                sentCacheStats.withSubject++;
                hasAnyIndex = true;
            }

            if (!hasAnyIndex) {
                sentCacheStats.withoutAnyIndex++;
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

/**
 * Get sent cache statistics for debugging
 * @returns {Object} Statistics object
 */
function getSentCacheStats() {
    return {
        totalCached: sentCacheStats.totalCached,
        withConvId: sentCacheStats.withConvId,
        withInReplyTo: sentCacheStats.withInReplyTo,
        withSubject: sentCacheStats.withSubject,
        withoutAnyIndex: sentCacheStats.withoutAnyIndex,
        uniqueConvIds: Object.keys(sentItemsByConvId).length,
        uniqueInReplyTo: Object.keys(sentItemsByInReplyTo).length,
        uniqueSubjects: Object.keys(sentItemsByNormSubject).length
    };
}

