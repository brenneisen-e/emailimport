/**
 * Thread Depth Module
 * Calculate thread depths and positions for conversation structure
 */

/**
 * Calculate thread depth for all emails
 * @param {Array} emails - Array of email objects with parent relationships
 * @param {Object} messageIdMap - Map of internetMessageId -> email
 * @returns {Array} Emails with threadDepth assigned
 */
function calculateThreadDepthsUnified(emails, messageIdMap) {
    var depthCache = {};

    function getDepth(email) {
        // Cache hit
        if (depthCache[email.internetMessageId] !== undefined) {
            return depthCache[email.internetMessageId];
        }

        // Base: Root has depth 0
        if (email.isThreadRoot || !email.parentMessageId) {
            depthCache[email.internetMessageId] = 0;
            return 0;
        }

        // Find parent
        var parent = messageIdMap[email.parentMessageId];

        // Parent not in thread (orphan)
        if (!parent) {
            depthCache[email.internetMessageId] = 0;
            return 0;
        }

        // Recursion: depth = parent depth + 1
        var depth = getDepth(parent) + 1;
        depthCache[email.internetMessageId] = depth;
        return depth;
    }

    // Calculate for all emails
    for (var i = 0; i < emails.length; i++) {
        emails[i].threadDepth = getDepth(emails[i]);
    }

    return emails;
}

/**
 * Assign thread positions (chronological order)
 * @param {Array} emails - Array of email objects
 * @returns {Array} Emails with threadPosition assigned
 */
function assignThreadPositionsUnified(emails) {
    // Sort chronologically (oldest first)
    emails.sort(function(a, b) {
        return a.timestamp - b.timestamp;
    });

    // Assign positions
    for (var i = 0; i < emails.length; i++) {
        emails[i].threadPosition = i + 1;
    }

    return emails;
}

/**
 * Process a single conversation with unified logic
 * @param {Array} emailsInConversation - Array of emails in one conversation
 * @returns {Object} Processed conversation JSON
 */
function processConversationUnified(emailsInConversation) {
    // Step 1: Remove duplicates
    var emails = removeDuplicateEmails(emailsInConversation);

    // Step 2: Create message ID index
    var messageIdMap = {};
    for (var i = 0; i < emails.length; i++) {
        if (emails[i].internetMessageId) {
            messageIdMap[emails[i].internetMessageId] = emails[i];
        }
    }

    // Step 3: Assign parent relationships
    emails = assignParentRelationships(emails, messageIdMap);

    // Step 4: Identify thread root
    var threadRoot = identifyThreadRoot(emails);

    // Step 5: Calculate thread depths
    emails = calculateThreadDepthsUnified(emails, messageIdMap);

    // Step 6: Assign thread positions (chronological)
    emails = assignThreadPositionsUnified(emails);

    // Step 7: Extract new content for all emails
    for (var j = 0; j < emails.length; j++) {
        var email = emails[j];
        if (email.htmlBody) {
            var contentResult = extractNewContentFromHtml(email.htmlBody);
            if (contentResult.newContent) {
                email.newContent = contentResult.newContent;
            } else {
                email.newContent = removeEmailQuotes(email.text);
            }
        } else {
            email.newContent = removeEmailQuotes(email.text);
        }
    }

    // Step 8: Build conversation JSON
    return buildConversationJsonUnified(threadRoot, emails);
}

/**
 * Build conversation JSON from processed emails
 * @param {Object} threadRoot - The root email of the thread
 * @param {Array} allEmails - All emails in the conversation
 * @returns {Object} Conversation JSON object
 */
function buildConversationJsonUnified(threadRoot, allEmails) {
    // Sort by threadPosition
    allEmails.sort(function(a, b) {
        return a.threadPosition - b.threadPosition;
    });

    // Replies = all except root
    var replies = [];
    for (var i = 0; i < allEmails.length; i++) {
        if (!allEmails[i].isThreadRoot) {
            replies.push(allEmails[i]);
        }
    }

    // Conversation start date = root timestamp
    var conversationStartDate = threadRoot.timestamp;

    // Last activity date
    var lastActivityDate = threadRoot.timestamp;
    for (var j = 0; j < allEmails.length; j++) {
        if (allEmails[j].timestamp > lastActivityDate) {
            lastActivityDate = allEmails[j].timestamp;
        }
    }

    // For outgoing emails (sent items as root), show recipient instead of sender
    var rootVonName = threadRoot.von_name;
    var rootVonEmail = threadRoot.von_email;
    if (!threadRoot.isIncoming && threadRoot.recipients && threadRoot.recipients.length > 0) {
        // This is an outgoing email - show "An: [recipient]"
        rootVonName = 'An: ' + (threadRoot.recipients[0].name || threadRoot.recipients[0].email || '');
        rootVonEmail = threadRoot.recipients[0].email || '';
    }

    var json = {
        // === IDENTIFICATION ===
        emailId: threadRoot.entryId,
        conversationId: threadRoot.conversationId,
        internetMessageId: threadRoot.internetMessageId,
        threadId: 'conv:' + threadRoot.conversationId,

        // === TIMESTAMPS ===
        datum: formatDate(threadRoot.timestamp),
        conversationStartDate: formatDate(conversationStartDate),
        lastActivityDate: formatDate(lastActivityDate),

        // === SENDER (Root) ===
        von_email: rootVonEmail,
        von_name: rootVonName,

        // === SUBJECT ===
        betreff: threadRoot.betreff,
        originalSubject: threadRoot.betreff,
        normalizedSubject: threadRoot.normalizedSubject,

        // === CONTENT (only new content!) ===
        text: threadRoot.newContent || threadRoot.text,
        htmlBody: threadRoot.htmlBody,

        // === THREAD METADATA ===
        threadPosition: threadRoot.threadPosition,
        threadDepth: 0,
        isThreadRoot: true,
        messageCount: allEmails.length,

        // === HEADER REFERENCES ===
        inReplyTo: threadRoot.inReplyTo || '',
        references: threadRoot.references || [],

        // === ATTACHMENTS ===
        anhaenge: threadRoot.anhaenge,

        // === OUTLOOK CATEGORIES ===
        outlookKategorie: threadRoot.outlookKategorie,

        // === REPLIES ===
        antworten: []
    };

    // Build reply objects
    for (var k = 0; k < replies.length; k++) {
        var email = replies[k];

        var reply = {
            // Timestamp
            datum: formatDate(email.timestamp),

            // People
            von: email.von_name,
            vonEmail: email.von_email,

            // For outgoing: recipients
            an: email.isIncoming ? '' : formatRecipientsUnified(email.recipients),

            // Content (only new content!)
            text: email.newContent || email.text,
            htmlBody: email.htmlBody,

            // Thread structure
            threadPosition: email.threadPosition,
            threadDepth: email.threadDepth,
            parentMessageId: email.parentMessageId || '',

            // Identification
            replyId: email.entryId,
            internetMessageId: email.internetMessageId,
            inReplyTo: email.inReplyTo || '',

            // Direction
            isIncoming: email.isIncoming,
            isNew: true,

            // Additional flags
            isDirectReplyToRoot: email.parentMessageId === threadRoot.internetMessageId
        };

        json.antworten.push(reply);
    }

    // FALLBACK: If no replies found through unified processing, use original antworten
    if (json.antworten.length === 0 && threadRoot.originalAntworten && threadRoot.originalAntworten.length > 0) {
        for (var f = 0; f < threadRoot.originalAntworten.length; f++) {
            var origReply = threadRoot.originalAntworten[f];
            json.antworten.push({
                datum: origReply.datum || '',
                von: origReply.von || '',
                vonEmail: origReply.vonEmail || '',
                an: origReply.an || '',
                text: origReply.text || '',
                htmlBody: origReply.htmlBody || '',
                threadPosition: origReply.threadPosition || (f + 2),
                threadDepth: origReply.threadDepth || 1,
                parentMessageId: origReply.parentMessageId || origReply.inReplyTo || '',
                replyId: origReply.replyId || '',
                internetMessageId: origReply.internetMessageId || '',
                inReplyTo: origReply.inReplyTo || '',
                isIncoming: origReply.isIncoming !== false,
                isNew: true,
                isDirectReplyToRoot: true
            });
        }
        // Update message count
        json.messageCount = 1 + json.antworten.length;
    }

    // Use original von_name if available (preserves "An: ..." from handleNewEmail)
    if (threadRoot.originalVonName && threadRoot.originalVonName !== '') {
        json.von_name = threadRoot.originalVonName;
    }

    return json;
}

/**
 * Format recipients for JSON
 * @param {Array} recipients - Array of recipient objects
 * @returns {string} Formatted recipients string
 */
function formatRecipientsUnified(recipients) {
    if (!recipients || recipients.length === 0) return '';
    var names = [];
    for (var i = 0; i < recipients.length; i++) {
        names.push(recipients[i].name || recipients[i].email || '');
    }
    return names.join('; ');
}

