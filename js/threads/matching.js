/**
 * Thread Matching Module
 * Match sent items to incoming emails using multiple methods
 */

// Debug stats for tracking match effectiveness
var matchingStats = {
    byConversationId: 0,
    byInternetMessageId: 0,
    bySubject: 0,
    noMatch: 0,
    totalProcessed: 0
};

/**
 * Normalize subject for matching
 * Removes Re:, AW:, Fwd:, WG: prefixes and extra whitespace
 * @param {string} subject - Original subject
 * @returns {string} Normalized subject
 */
function normalizeSubjectForMatching(subject) {
    if (!subject) return '';
    // Remove common reply/forward prefixes (German and English)
    var normalized = subject.replace(/^(re:|aw:|fwd:|wg:|fw:)\s*/gi, '');
    // Remove multiple prefixes
    while (/^(re:|aw:|fwd:|wg:|fw:)\s*/i.test(normalized)) {
        normalized = normalized.replace(/^(re:|aw:|fwd:|wg:|fw:)\s*/gi, '');
    }
    // Normalize whitespace
    normalized = normalized.replace(/\s+/g, ' ').trim().toLowerCase();
    return normalized;
}

/**
 * Check if two dates are within a time window (default 30 days)
 * @param {Date} date1 - First date
 * @param {Date} date2 - Second date
 * @param {number} maxDays - Maximum days apart (default 30)
 * @returns {boolean} True if within window
 */
function isWithinTimeWindow(date1, date2, maxDays) {
    if (!date1 || !date2) return false;
    maxDays = maxDays || 30;
    var diffMs = Math.abs(date1.getTime() - date2.getTime());
    var diffDays = diffMs / (1000 * 60 * 60 * 24);
    return diffDays <= maxDays;
}

/**
 * Find replies for an email using multiple matching strategies
 * Strategy priority:
 * 1. ConversationID (most reliable for Outlook)
 * 2. InternetMessageId/InReplyTo headers
 * 3. Subject + Time window (fallback)
 * @param {Object} originalEmail - The original email to find replies for
 * @returns {Array} Array of reply objects
 */
function findRepliesMultiLayer(originalEmail) {
    var replies = [];
    var addedEntryIds = {};  // Prevent duplicates
    matchingStats.totalProcessed++;

    try {
        var originalConvId = originalEmail.conversationId || '';
        var originalMsgId = originalEmail.internetMessageId || '';
        var originalSubject = normalizeSubjectForMatching(originalEmail.betreff);
        var originalDate = originalEmail.datum ? new Date(originalEmail.datum) : null;

        // Strategy 1: ConversationID matching (primary - most reliable for Outlook)
        if (originalConvId && sentItemsByConvId[originalConvId]) {
            var convMatches = sentItemsByConvId[originalConvId];
            for (var i = 0; i < convMatches.length; i++) {
                var cached = convMatches[i];
                if (cached.body && cached.body.trim() !== '' && !addedEntryIds[cached.entryId]) {
                    replies.push(createReplyObject(cached, 'conversationId'));
                    addedEntryIds[cached.entryId] = true;
                    matchingStats.byConversationId++;
                }
            }
        }

        // Strategy 2: InternetMessageId / InReplyTo matching (for cross-client threads)
        if (originalMsgId && sentItemsByInReplyTo && sentItemsByInReplyTo[originalMsgId]) {
            var replyToMatches = sentItemsByInReplyTo[originalMsgId];
            for (var j = 0; j < replyToMatches.length; j++) {
                var cached2 = replyToMatches[j];
                if (cached2.body && cached2.body.trim() !== '' && !addedEntryIds[cached2.entryId]) {
                    replies.push(createReplyObject(cached2, 'inReplyTo'));
                    addedEntryIds[cached2.entryId] = true;
                    matchingStats.byInternetMessageId++;
                }
            }
        }

        // Strategy 3: Subject + Time window matching (fallback for emails without ConversationID)
        if (originalSubject && originalSubject.length > 5 && sentItemsByNormSubject) {
            var subjectMatches = sentItemsByNormSubject[originalSubject];
            if (subjectMatches && subjectMatches.length > 0) {
                for (var k = 0; k < subjectMatches.length; k++) {
                    var cached3 = subjectMatches[k];
                    // Skip if already added or empty body
                    if (!cached3.body || cached3.body.trim() === '' || addedEntryIds[cached3.entryId]) {
                        continue;
                    }
                    // Check time window: sent item must be after original (reply) and within 30 days
                    var sentDate = cached3.sentDate;
                    if (originalDate && sentDate) {
                        var isAfterOriginal = sentDate >= originalDate;
                        var withinWindow = isWithinTimeWindow(originalDate, sentDate, 30);
                        if (isAfterOriginal && withinWindow) {
                            replies.push(createReplyObject(cached3, 'subject'));
                            addedEntryIds[cached3.entryId] = true;
                            matchingStats.bySubject++;
                        }
                    }
                }
            }
        }

        // Track no-match cases
        if (replies.length === 0) {
            matchingStats.noMatch++;
        }

    } catch (e) {
        // Log error but don't fail
        try { console.log('findRepliesMultiLayer error: ' + e.message); } catch(e2) {}
    }

    // Sort replies by date (oldest first)
    replies.sort(function(a, b) {
        var dateA = a.datumRaw ? new Date(a.datumRaw) : new Date();
        var dateB = b.datumRaw ? new Date(b.datumRaw) : new Date();
        return dateA - dateB;
    });

    return replies;
}

/**
 * Create a reply object from cached sent item
 * @param {Object} cached - Cached sent item
 * @param {string} matchMethod - How this reply was matched
 * @returns {Object} Reply object
 */
function createReplyObject(cached, matchMethod) {
    return {
        datum: formatDate(cached.sentDate),
        datumRaw: cached.sentDate,
        an: cached.to || '',
        text: cached.body ? cached.body.trim() : '',
        replyId: cached.entryId || '',
        matchMethod: matchMethod || 'unknown'
    };
}

/**
 * Simple findReplies for backward compatibility
 * Uses multi-layer matching
 * @param {Object} originalEmail - The original email
 * @returns {Array} Array of reply objects
 */
function findReplies(originalEmail) {
    return findRepliesMultiLayer(originalEmail);
}

/**
 * Get matching statistics for debugging
 * @returns {Object} Statistics object
 */
function getMatchingStats() {
    var total = matchingStats.byConversationId + matchingStats.byInternetMessageId + matchingStats.bySubject;
    return {
        byConversationId: matchingStats.byConversationId,
        byInternetMessageId: matchingStats.byInternetMessageId,
        bySubject: matchingStats.bySubject,
        noMatch: matchingStats.noMatch,
        totalProcessed: matchingStats.totalProcessed,
        totalMatches: total,
        matchRate: matchingStats.totalProcessed > 0 ?
            Math.round((matchingStats.totalProcessed - matchingStats.noMatch) / matchingStats.totalProcessed * 100) : 0
    };
}

/**
 * Reset matching statistics
 */
function resetMatchingStats() {
    matchingStats.byConversationId = 0;
    matchingStats.byInternetMessageId = 0;
    matchingStats.bySubject = 0;
    matchingStats.noMatch = 0;
    matchingStats.totalProcessed = 0;
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

