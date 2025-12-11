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
    // SIMPLIFIED to match old working version
    sentItemsCache = [];
    sentItemsByConvId = {};  // Only index we need
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

            // SIMPLIFIED to match old working version
            var body = '';
            try { body = item.Body || ''; } catch (e) {}

            var toRecipient = '';
            try { toRecipient = item.To || ''; } catch (e) {}

            var convId = '';
            try { convId = item.ConversationID || ''; } catch (e) {}

            var entryId = '';
            try { entryId = item.EntryID || ''; } catch (e) {}

            // Match old working version's sent item structure
            var sentItem = {
                sentDate: sentDate,
                conversationId: convId,
                entryId: entryId,
                body: removeEmailQuotes(body),
                to: toRecipient
            };
            sentItemsCache.push(sentItem);

            // Index by ConversationID (only - like old version)
            if (convId) {
                if (!sentItemsByConvId[convId]) {
                    sentItemsByConvId[convId] = [];
                }
                sentItemsByConvId[convId].push(sentItem);
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

