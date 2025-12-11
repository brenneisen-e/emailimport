/**
 * Conversation Extractor Module
 * Extract emails with full ConversationIndex support for hierarchy building
 */

// MAPI Property for ConversationIndex (binary)
var PR_CONVERSATION_INDEX_BIN = "http://schemas.microsoft.com/mapi/proptag/0x00710102";

// MAPI Property for SMTP Address (resolves Exchange X500 to SMTP)
var PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F";

/**
 * Get SMTP email address from sender (resolves Exchange X500 addresses)
 * @param {Object} mailItem - Outlook mail item
 * @returns {string} SMTP email address
 */
function getSenderSMTPAddress(mailItem) {
    try {
        // Check if sender is Exchange type
        var emailType = mailItem.SenderEmailType;

        if (emailType === 'EX') {
            // Try to get SMTP address via Sender object
            try {
                var sender = mailItem.Sender;
                if (sender) {
                    // Try GetExchangeUser first
                    try {
                        var exchUser = sender.GetExchangeUser();
                        if (exchUser && exchUser.PrimarySmtpAddress) {
                            return exchUser.PrimarySmtpAddress;
                        }
                    } catch (e1) {}

                    // Try PropertyAccessor
                    try {
                        var propAccessor = sender.PropertyAccessor;
                        var smtpAddr = propAccessor.GetProperty(PR_SMTP_ADDRESS);
                        if (smtpAddr) return smtpAddr;
                    } catch (e2) {}
                }
            } catch (e3) {}

            // Fallback: Try PropertyAccessor on mail item itself
            try {
                var mailProp = mailItem.PropertyAccessor;
                // PR_SENDER_SMTP_ADDRESS
                var senderSmtp = mailProp.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x5D01001F");
                if (senderSmtp) return senderSmtp;
            } catch (e4) {}
        }

        // Not Exchange or couldn't resolve - return original
        return mailItem.SenderEmailAddress || '';
    } catch (e) {
        return mailItem.SenderEmailAddress || '';
    }
}

/**
 * Get SMTP email address from recipient (resolves Exchange X500 addresses)
 * @param {Object} recipient - Outlook recipient object
 * @returns {string} SMTP email address
 */
function getRecipientSMTPAddress(recipient) {
    try {
        var addrType = recipient.AddressEntry ? recipient.AddressEntry.Type : '';

        if (addrType === 'EX') {
            try {
                var exchUser = recipient.AddressEntry.GetExchangeUser();
                if (exchUser && exchUser.PrimarySmtpAddress) {
                    return exchUser.PrimarySmtpAddress;
                }
            } catch (e1) {}

            try {
                var propAccessor = recipient.AddressEntry.PropertyAccessor;
                var smtpAddr = propAccessor.GetProperty(PR_SMTP_ADDRESS);
                if (smtpAddr) return smtpAddr;
            } catch (e2) {}
        }

        return recipient.Address || '';
    } catch (e) {
        return recipient.Address || '';
    }
}

/**
 * Extract ConversationIndex as Hex string from mail item
 * @param {Object} mailItem - Outlook mail item
 * @returns {string} Hex string of ConversationIndex
 */
function getConversationIndexHex(mailItem) {
    try {
        var propAccessor = mailItem.PropertyAccessor;
        var indexBytes = propAccessor.GetProperty(PR_CONVERSATION_INDEX_BIN);

        if (!indexBytes) return '';

        var hexStr = '';
        // Convert binary array to hex string
        // indexBytes is a VBArray in HTA/COM
        try {
            var arr = new VBArray(indexBytes).toArray();
            for (var i = 0; i < arr.length; i++) {
                var hex = arr[i].toString(16).toUpperCase();
                if (hex.length === 1) hex = '0' + hex;
                hexStr += hex;
            }
        } catch (vbErr) {
            // Alternative: try direct iteration if not VBArray
            if (typeof indexBytes.length !== 'undefined') {
                for (var j = 0; j < indexBytes.length; j++) {
                    var hex2 = indexBytes[j].toString(16).toUpperCase();
                    if (hex2.length === 1) hex2 = '0' + hex2;
                    hexStr += hex2;
                }
            }
        }
        return hexStr;
    } catch (e) {
        // Fallback: try using ConversationIndex property directly (base64)
        try {
            var convIndex = mailItem.ConversationIndex;
            if (convIndex) {
                // Try to convert base64 to hex
                return base64ToHex(convIndex);
            }
        } catch (e2) {}
        return '';
    }
}

/**
 * Convert base64 string to hex string
 * @param {string} base64 - Base64 encoded string
 * @returns {string} Hex string
 */
function base64ToHex(base64) {
    if (!base64) return '';
    try {
        // Base64 decode table
        var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
        var lookup = {};
        for (var i = 0; i < chars.length; i++) {
            lookup[chars[i]] = i;
        }

        // Remove padding
        base64 = base64.replace(/=/g, '');

        var bytes = [];
        for (var j = 0; j < base64.length; j += 4) {
            var n = (lookup[base64[j]] << 18) |
                    (lookup[base64[j + 1]] << 12) |
                    ((lookup[base64[j + 2]] || 0) << 6) |
                    (lookup[base64[j + 3]] || 0);

            bytes.push((n >> 16) & 0xFF);
            if (base64[j + 2]) bytes.push((n >> 8) & 0xFF);
            if (base64[j + 3]) bytes.push(n & 0xFF);
        }

        var hex = '';
        for (var k = 0; k < bytes.length; k++) {
            var h = bytes[k].toString(16).toUpperCase();
            if (h.length === 1) h = '0' + h;
            hex += h;
        }
        return hex;
    } catch (e) {
        return '';
    }
}

/**
 * Calculate depth from ConversationIndex hex
 * Base header = 22 bytes (44 hex chars)
 * Each response adds 5 bytes (10 hex chars)
 * @param {string} hexIndex - Hex string of ConversationIndex
 * @returns {number} Depth level (0 = root)
 */
function calculateDepthFromIndex(hexIndex) {
    if (!hexIndex || hexIndex.length < 44) return 0;
    return Math.floor((hexIndex.length - 44) / 10);
}

/**
 * Get parent's ConversationIndex by removing last 5-byte block
 * @param {string} hexIndex - Hex string of ConversationIndex
 * @returns {string} Parent's hex index, or empty if root
 */
function getParentIndexHex(hexIndex) {
    if (!hexIndex || hexIndex.length <= 44) return '';
    return hexIndex.substring(0, hexIndex.length - 10);
}

/**
 * Extract email in conversation format with all required properties
 * @param {Object} mailItem - Outlook mail item
 * @param {string} folder - 'inbox' or 'sent'
 * @returns {Object} Email object for conversation building
 */
function extractEmailForConversation(mailItem, folder) {
    var email = {
        entryID: '',
        conversationID: '',
        conversationIndex: '',
        conversationIndexHex: '',
        subject: '',
        senderEmail: '',
        senderName: '',
        recipients: [],
        ccRecipients: [],
        receivedTime: null,
        sentOn: null,
        folder: folder || 'inbox',
        bodyPreview: '',
        body: '',
        hasAttachments: false,
        importance: 1,
        messageClass: 'IPM.Note',
        // Calculated fields
        depth: 0,
        parentIndexHex: '',
        parentEntryID: null,
        childEntryIDs: [],
        // Additional fields for compatibility
        kategorie: '',
        anhaenge: []
    };

    try {
        // Entry ID (unique identifier)
        try { email.entryID = mailItem.EntryID || ''; } catch (e) {}

        // Conversation ID
        try { email.conversationID = mailItem.ConversationID || ''; } catch (e) {}

        // Conversation Index (critical for hierarchy!)
        try {
            email.conversationIndex = mailItem.ConversationIndex || '';
            email.conversationIndexHex = getConversationIndexHex(mailItem);

            // Fallback: if conversationIndexHex is empty, use conversationIndex directly
            // (Outlook sometimes returns index in usable format already)
            if (!email.conversationIndexHex && email.conversationIndex) {
                email.conversationIndexHex = email.conversationIndex;
            }

            // Calculate depth and parent from index
            if (email.conversationIndexHex) {
                email.depth = calculateDepthFromIndex(email.conversationIndexHex);
                email.parentIndexHex = getParentIndexHex(email.conversationIndexHex);
            }
        } catch (e) {}

        // Subject
        try { email.subject = mailItem.Subject || ''; } catch (e) {}

        // Sender (resolve Exchange X500 to SMTP)
        try {
            email.senderEmail = getSenderSMTPAddress(mailItem);
            email.senderName = mailItem.SenderName || '';
        } catch (e) {}

        // Recipients (To) - resolve Exchange X500 to SMTP
        try {
            var recipients = mailItem.Recipients;
            for (var r = 1; r <= recipients.Count; r++) {
                try {
                    var recip = recipients.Item(r);
                    var recipEmail = getRecipientSMTPAddress(recip);
                    if (recip.Type === 1) { // To
                        email.recipients.push({
                            name: recip.Name || '',
                            email: recipEmail
                        });
                    } else if (recip.Type === 2) { // CC
                        email.ccRecipients.push({
                            name: recip.Name || '',
                            email: recipEmail
                        });
                    }
                } catch (e) {}
            }
        } catch (e) {}

        // Timestamps
        try {
            var received = new Date(mailItem.ReceivedTime);
            if (!isNaN(received.getTime())) {
                email.receivedTime = received.toISOString();
            }
        } catch (e) {}

        try {
            var sent = new Date(mailItem.SentOn);
            if (!isNaN(sent.getTime())) {
                email.sentOn = sent.toISOString();
            }
        } catch (e) {}

        // Body preview (first 500 chars)
        try {
            var fullBody = mailItem.Body || '';
            email.body = fullBody.length > 50000 ? fullBody.substring(0, 50000) : fullBody;
            email.bodyPreview = fullBody.substring(0, 500);
        } catch (e) {}

        // Attachments
        try {
            var attCount = mailItem.Attachments.Count;
            email.hasAttachments = attCount > 0;
            for (var a = 1; a <= attCount; a++) {
                try {
                    email.anhaenge.push(mailItem.Attachments.Item(a).FileName);
                } catch (e) {}
            }
        } catch (e) {}

        // Importance (0=low, 1=normal, 2=high)
        try { email.importance = mailItem.Importance || 1; } catch (e) {}

        // Message class
        try { email.messageClass = mailItem.MessageClass || 'IPM.Note'; } catch (e) {}

        // Category
        try {
            var cats = mailItem.Categories || '';
            if (cats) {
                email.kategorie = cats.split(',')[0].trim();
            }
        } catch (e) {}

    } catch (e) {}

    return email;
}

/**
 * Build conversation hierarchy from flat list of emails
 * @param {Array} emails - Array of extracted emails
 * @returns {Object} Conversations grouped by conversationID with hierarchy
 */
function buildConversationHierarchy(emails) {
    var conversations = {};
    var indexToEntryId = {}; // Map conversationIndexHex -> entryID
    var entryIdToEmail = {}; // Map entryID -> email object

    // First pass: Group by conversationID and build index maps
    for (var i = 0; i < emails.length; i++) {
        var email = emails[i];
        var convId = email.conversationID || 'no-conversation';

        // Initialize conversation if needed
        if (!conversations[convId]) {
            conversations[convId] = {
                conversationID: convId,
                subject: '',
                participantCount: 0,
                messageCount: 0,
                firstMessageDate: null,
                lastMessageDate: null,
                messages: [],
                participants: {}
            };
        }

        conversations[convId].messages.push(email);
        conversations[convId].messageCount++;

        // Build index maps
        if (email.conversationIndexHex) {
            indexToEntryId[email.conversationIndexHex] = email.entryID;
        }
        entryIdToEmail[email.entryID] = email;

        // Track participants
        if (email.senderEmail) {
            conversations[convId].participants[email.senderEmail] = email.senderName || email.senderEmail;
        }
    }

    // Second pass: Build parent-child relationships
    for (var convId in conversations) {
        var conv = conversations[convId];
        var messages = conv.messages;

        // Sort by conversationIndexHex length (shortest = oldest/root)
        messages.sort(function(a, b) {
            var lenA = (a.conversationIndexHex || '').length;
            var lenB = (b.conversationIndexHex || '').length;
            if (lenA !== lenB) return lenA - lenB;
            // If same length, sort by timestamp
            var timeA = a.sentOn || a.receivedTime || '';
            var timeB = b.sentOn || b.receivedTime || '';
            return timeA < timeB ? -1 : (timeA > timeB ? 1 : 0);
        });

        // Build hierarchy
        for (var j = 0; j < messages.length; j++) {
            var msg = messages[j];

            // Find parent by prefix match
            if (msg.parentIndexHex && indexToEntryId[msg.parentIndexHex]) {
                msg.parentEntryID = indexToEntryId[msg.parentIndexHex];
                // Add this message as child of parent
                var parent = entryIdToEmail[msg.parentEntryID];
                if (parent && parent.childEntryIDs) {
                    parent.childEntryIDs.push(msg.entryID);
                }
            }

            // Update conversation metadata
            var msgTime = msg.sentOn || msg.receivedTime;
            if (msgTime) {
                if (!conv.firstMessageDate || msgTime < conv.firstMessageDate) {
                    conv.firstMessageDate = msgTime;
                    conv.subject = msg.subject; // Use earliest message's subject
                }
                if (!conv.lastMessageDate || msgTime > conv.lastMessageDate) {
                    conv.lastMessageDate = msgTime;
                }
            }
        }

        // Set participant count
        conv.participantCount = Object.keys(conv.participants).length;
        delete conv.participants; // Remove helper object
    }

    return conversations;
}

/**
 * Create export JSON structure
 * @param {Object} conversations - Grouped conversations
 * @param {string} mailboxName - Name of the mailbox
 * @returns {Object} Export-ready JSON structure
 */
function createExportStructure(conversations, mailboxName) {
    var totalEmails = 0;
    var byEntryID = {};
    var byConversationIndex = {};

    for (var convId in conversations) {
        var conv = conversations[convId];
        totalEmails += conv.messageCount;

        for (var i = 0; i < conv.messages.length; i++) {
            var msg = conv.messages[i];
            byEntryID[msg.entryID] = convId;
            if (msg.conversationIndexHex) {
                byConversationIndex[msg.conversationIndexHex] = msg.entryID;
            }
        }
    }

    return {
        exportDate: new Date().toISOString(),
        mailboxName: mailboxName || '',
        totalEmails: totalEmails,
        conversationCount: Object.keys(conversations).length,
        conversations: conversations,
        index: {
            byEntryID: byEntryID,
            byConversationIndex: byConversationIndex
        }
    };
}
