/**
 * Thread Matching Module
 * Match sent items to incoming emails using multiple methods
 */

/**
 * Find replies for an email using multi-layer matching
 * @param {Object} originalEmail - The original email to find replies for
 * @returns {Array} Array of reply objects
 */
function findRepliesMultiLayer(originalEmail) {
    var replies = [];
    var seenIds = {};  // Prevent duplicates

    try {
        var originalConvId = originalEmail.conversationId || '';
        var originalMsgId = originalEmail.internetMessageId || '';
        var originalNormSubject = normalizeSubject(originalEmail.betreff || '');

        // Layer 1: ConversationID (Primary - Outlook native)
        if (originalConvId && sentItemsByConvId[originalConvId]) {
            var convMatches = sentItemsByConvId[originalConvId];
            for (var i = 0; i < convMatches.length; i++) {
                var cached = convMatches[i];
                if (seenIds[cached.entryId]) continue;
                seenIds[cached.entryId] = true;

                if (cached.body && cached.body.trim() !== '') {
                    replies.push(createReplyObject(cached));
                }
            }
        }

        // Layer 2: In-Reply-To header (direct parent-child relationship)
        if (originalMsgId && sentItemsByInReplyTo[originalMsgId]) {
            var replyToMatches = sentItemsByInReplyTo[originalMsgId];
            for (var j = 0; j < replyToMatches.length; j++) {
                var cached = replyToMatches[j];
                if (seenIds[cached.entryId]) continue;
                seenIds[cached.entryId] = true;

                if (cached.body && cached.body.trim() !== '') {
                    replies.push(createReplyObject(cached));
                }
            }
        }

        // Layer 3: References header (thread chain membership)
        if (originalMsgId && sentItemsByReference[originalMsgId]) {
            var refMatches = sentItemsByReference[originalMsgId];
            for (var k = 0; k < refMatches.length; k++) {
                var cached = refMatches[k];
                if (seenIds[cached.entryId]) continue;
                seenIds[cached.entryId] = true;

                if (cached.body && cached.body.trim() !== '') {
                    replies.push(createReplyObject(cached));
                }
            }
        }

    } catch (e) {}

    return replies;
}

/**
 * Create a reply object from cached sent item
 * Enhanced with date extraction and content extraction
 * @param {Object} cached - Cached sent item
 * @returns {Object} Reply object
 */
function createReplyObject(cached) {
    var replyDatum = formatDate(cached.sentDate);
    var datumExtracted = false;
    var replyText = cached.body ? cached.body.trim() : '';
    var fullText = replyText;

    // Try to extract actual date from HTML headers
    if (cached.htmlBody) {
        var dateResult = extractDateFromHtmlHeader(cached.htmlBody);
        if (dateResult.success) {
            replyDatum = formatDate(dateResult.date);
            datumExtracted = true;
        }

        // Extract only new content from HTML
        var contentResult = extractNewContentFromHtml(cached.htmlBody);
        if (contentResult.newContent) {
            replyText = contentResult.newContent;
        }
    }

    return {
        datum: replyDatum,
        datumExtracted: datumExtracted,
        an: cached.to || '',
        text: replyText,
        fullText: fullText,
        htmlBody: cached.htmlBody || '',
        replyId: cached.entryId || '',
        internetMessageId: cached.internetMessageId || '',
        inReplyTo: cached.inReplyTo || '',
        isIncoming: false,  // Sent items are outgoing
        threadDepth: 1,     // Default depth
        threadPosition: 0,  // Will be set after sorting
        parentMessageId: cached.inReplyTo || ''
    };
}

/**
 * Simple findReplies for backward compatibility
 * Uses only ConversationID matching
 * @param {Object} originalEmail - The original email
 * @returns {Array} Array of reply objects
 */
function findReplies(originalEmail) {
    return findRepliesMultiLayer(originalEmail);
}

/**
 * Filter replies to only include ones not already in Excel
 * Uses ReplyID (EntryID) for precise matching, falls back to text matching
 * @param {Array} allReplies - All found replies
 * @param {string} existingReplyText - Existing reply text from Excel
 * @param {string} existingReplyIdList - Comma-separated existing reply IDs
 * @returns {Array} New replies only
 */
function filterNewReplies(allReplies, existingReplyText, existingReplyIdList) {
    if (!allReplies || allReplies.length === 0) return [];

    // Parse existing reply IDs into a lookup set
    var existingIdSet = {};
    if (existingReplyIdList) {
        var idParts = String(existingReplyIdList).split(',');
        for (var p = 0; p < idParts.length; p++) {
            var id = idParts[p].trim();
            if (id) existingIdSet[id] = true;
        }
    }

    // If no existing data at all - all are new
    if (!existingReplyText && !existingReplyIdList) {
        for (var i = 0; i < allReplies.length; i++) {
            allReplies[i].isNew = true;
        }
        return allReplies;
    }

    var newReplies = [];
    // Normalize text for fallback matching
    var existingNormalized = existingReplyText ? fixEncoding(existingReplyText).toLowerCase() : '';

    for (var j = 0; j < allReplies.length; j++) {
        var reply = allReplies[j];
        var replyText = (reply.text || '').trim();

        // Skip empty replies
        if (!replyText) continue;

        // Primary check: Use ReplyID if available (precise matching)
        if (reply.replyId && existingIdSet[reply.replyId]) {
            reply.isNew = false;
            continue;
        }

        // Fallback check: Text matching (for older data without IDs)
        if (existingNormalized) {
            var normalizedReply = fixEncoding(replyText);
            var checkText = normalizedReply.substring(0, 80).toLowerCase().trim();
            if (checkText.length < 10) checkText = normalizedReply.toLowerCase().trim();

            if (existingNormalized.indexOf(checkText) !== -1) {
                reply.isNew = false;
                continue;
            }
        }

        // Not found - this is new
        reply.isNew = true;
        newReplies.push(reply);
    }

    return newReplies;
}

