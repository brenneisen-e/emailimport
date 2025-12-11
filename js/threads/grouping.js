/**
 * Thread Grouping Module
 * Group emails by conversation and manage thread relationships
 */

/**
 * Group emails by ConversationID
 * @param {Array} emails - Array of email objects
 * @returns {Object} Map of conversationId -> email array
 */
function groupByConversation(emails) {
    var conversations = {};

    for (var i = 0; i < emails.length; i++) {
        var email = emails[i];
        var convId = email.conversationId || 'orphan:' + email.internetMessageId;

        if (!conversations[convId]) {
            conversations[convId] = [];
        }
        conversations[convId].push(email);
    }

    return conversations;
}

/**
 * Remove duplicates from email list using InternetMessageId
 * @param {Array} emails - Array of email objects
 * @returns {Array} Deduplicated array
 */
function removeDuplicateEmails(emails) {
    var seen = {};
    var unique = [];

    for (var i = 0; i < emails.length; i++) {
        var email = emails[i];
        var msgId = email.internetMessageId;

        // Without Message-ID: use hash fallback
        if (!msgId || msgId === '') {
            msgId = 'hash:' + (email.timestamp ? email.timestamp.getTime() : 0) + '|' +
                    email.von_email + '|' + email.betreff;
        }

        if (!seen[msgId]) {
            seen[msgId] = email;
            unique.push(email);
        } else {
            // Duplicate found - keep the one with more info
            var existing = seen[msgId];
            if (email.htmlBody && email.htmlBody.length > (existing.htmlBody || '').length) {
                // Replace with more complete version
                var index = unique.indexOf(existing);
                if (index >= 0) {
                    unique[index] = email;
                    seen[msgId] = email;
                }
            }
        }
    }

    return unique;
}

/**
 * Assign parent relationships using InReplyTo and References
 * @param {Array} emails - Array of email objects
 * @param {Object} messageIdMap - Map of internetMessageId -> email
 * @returns {Array} Emails with parent relationships assigned
 */
function assignParentRelationships(emails, messageIdMap) {
    for (var i = 0; i < emails.length; i++) {
        var email = emails[i];
        email.parentMessageId = null;
        email.parentFoundInThread = false;

        // Method 1: In-Reply-To header (PRIMARY - highest priority)
        if (email.inReplyTo && email.inReplyTo !== '') {
            if (messageIdMap[email.inReplyTo]) {
                email.parentMessageId = email.inReplyTo;
                email.parentFoundInThread = true;
                continue;
            }
        }

        // Method 2: References header (last entry that's in the thread)
        if (email.references && email.references.length > 0) {
            // Go from end to start
            for (var j = email.references.length - 1; j >= 0; j--) {
                var refId = email.references[j];
                if (messageIdMap[refId]) {
                    email.parentMessageId = refId;
                    email.parentFoundInThread = true;
                    break;
                }
            }
        }
    }

    return emails;
}

/**
 * Identify the true thread root
 * @param {Array} emails - Array of email objects with parent relationships
 * @returns {Object} The root email object
 */
function identifyThreadRoot(emails) {
    // Candidates: All emails without found parent in thread
    var rootCandidates = [];
    for (var i = 0; i < emails.length; i++) {
        if (!emails[i].parentFoundInThread) {
            rootCandidates.push(emails[i]);
        }
    }

    var root = null;

    // Case 1: Exactly one candidate
    if (rootCandidates.length === 1) {
        root = rootCandidates[0];
        root.isThreadRoot = true;
        return root;
    }

    // Case 2: Multiple candidates (e.g., incomplete conversation)
    if (rootCandidates.length > 1) {
        // Choose the oldest email
        rootCandidates.sort(function(a, b) {
            return a.timestamp - b.timestamp;
        });
        root = rootCandidates[0];
        root.isThreadRoot = true;

        // The other candidates: set parent to root if they come after
        for (var j = 1; j < rootCandidates.length; j++) {
            var orphan = rootCandidates[j];
            if (orphan.timestamp > root.timestamp) {
                orphan.parentMessageId = root.internetMessageId;
                orphan.parentFoundInThread = true;
                orphan.parentAssignedByFallback = true;
            }
        }

        return root;
    }

    // Case 3: No candidate (all have parent, but parent missing)
    // Choose oldest email as "virtual" root
    if (rootCandidates.length === 0) {
        emails.sort(function(a, b) {
            return a.timestamp - b.timestamp;
        });
        root = emails[0];
        root.isThreadRoot = true;
        root.isVirtualRoot = true;
        root.parentMessageId = null;
        root.parentFoundInThread = false;
        return root;
    }

    return null;
}

/**
 * Convert inbox email (old format) to unified format
 * @param {Object} email - Old format email
 * @param {string} folderType - 'inbox' or 'sent'
 * @returns {Object} Unified format email
 */
function convertToUnifiedFormat(email, folderType) {
    return {
        entryId: email.emailId || '',
        conversationId: email.conversationId || '',
        internetMessageId: email.internetMessageId || '',
        inReplyTo: email.inReplyTo || '',
        references: email.references || [],

        receivedTime: email.datum ? new Date(email.datum) : new Date(),
        sentOn: email.datum ? new Date(email.datum) : new Date(),
        timestamp: email.datum ? new Date(email.datum) : new Date(),

        von_email: email.von_email || '',
        von_name: email.von_name || '',
        recipients: [],

        betreff: email.betreff || '',
        text: email.text || '',
        htmlBody: email.htmlBody || '',
        anhaenge: email.anhaenge || [],
        outlookKategorie: email.outlookKategorie || '',
        folderType: folderType || 'inbox',
        folderName: email.folderName || '',
        isIncoming: (folderType !== 'sent'),

        threadPosition: 0,
        threadDepth: 0,
        parentMessageId: null,
        parentFoundInThread: false,
        isThreadRoot: false,
        newContent: '',
        normalizedSubject: normalizeSubject(email.betreff || '')
    };
}

/**
 * Convert sent item to unified format
 * @param {Object} sent - Sent item from cache
 * @returns {Object} Unified format email
 */
function convertSentItemToUnifiedFormat(sent) {
    return {
        entryId: sent.entryId || '',
        conversationId: sent.conversationId || '',
        internetMessageId: sent.internetMessageId || '',
        inReplyTo: sent.inReplyTo || '',
        references: sent.references || [],

        receivedTime: sent.sentDate ? new Date(sent.sentDate) : new Date(),
        sentOn: sent.sentDate ? new Date(sent.sentDate) : new Date(),
        timestamp: sent.sentDate ? new Date(sent.sentDate) : new Date(),

        von_email: '',
        von_name: '',
        recipients: [{name: sent.to || '', email: '', type: 1}],

        betreff: sent.subject || '',
        text: sent.body || '',
        htmlBody: sent.htmlBody || '',
        anhaenge: [],
        outlookKategorie: '',
        folderType: 'sent',
        folderName: 'Sent',
        isIncoming: false,

        threadPosition: 0,
        threadDepth: 0,
        parentMessageId: sent.inReplyTo || null,
        parentFoundInThread: false,
        isThreadRoot: false,
        newContent: sent.body || '',
        normalizedSubject: sent.normalizedSubject || normalizeSubject(sent.subject || '')
    };
}

