/**
 * Thread Matching Module
 * Match sent items to incoming emails using multiple methods
 */

/**
 * Find replies for an email - SIMPLIFIED to match old working version
 * Uses only ConversationID matching (like the old version that worked)
 * @param {Object} originalEmail - The original email to find replies for
 * @returns {Array} Array of reply objects
 */
function findRepliesMultiLayer(originalEmail) {
    var replies = [];

    try {
        var originalConvId = originalEmail.conversationId || '';

        // No ConversationID = can't match reliably (same as old version)
        if (!originalConvId) return replies;

        // Use lookup map for O(1) access (same as old version)
        var matchingItems = sentItemsByConvId[originalConvId];
        if (!matchingItems || matchingItems.length === 0) return replies;

        for (var i = 0; i < matchingItems.length; i++) {
            var cached = matchingItems[i];
            if (cached.body && cached.body.trim() !== '') {
                replies.push(createReplyObject(cached));
            }
        }
    } catch (e) {}

    return replies;
}

/**
 * Create a reply object from cached sent item
 * SIMPLIFIED to match old working version exactly
 * @param {Object} cached - Cached sent item
 * @returns {Object} Reply object
 */
function createReplyObject(cached) {
    // Use simple structure like old working version
    // Old version: { datum, an, text, replyId }
    return {
        datum: formatDate(cached.sentDate),
        an: cached.to || '',
        text: cached.body ? cached.body.trim() : '',
        replyId: cached.entryId || ''
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

