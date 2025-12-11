/**
 * MAPI Properties Module
 * Extract email headers and properties using MAPI PropertyAccessor
 */

// MAPI Property Tags for email headers
var PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";
var PR_IN_REPLY_TO_ID = "http://schemas.microsoft.com/mapi/proptag/0x1042001F";
var PR_INTERNET_REFERENCES = "http://schemas.microsoft.com/mapi/proptag/0x1039001F";
var PR_CONVERSATION_INDEX = "http://schemas.microsoft.com/mapi/proptag/0x00710102";

/**
 * Extract email headers using PropertyAccessor
 * @param {Object} item - Outlook mail item
 * @returns {Object} Headers object with internetMessageId, inReplyTo, references, conversationIndex
 */
function getEmailHeaders(item) {
    var headers = {
        internetMessageId: '',
        inReplyTo: '',
        references: [],
        conversationIndex: ''
    };

    try {
        var pa = item.PropertyAccessor;
        if (!pa) return headers;

        // Internet Message-ID (unique identifier across email systems)
        try {
            headers.internetMessageId = pa.GetProperty(PR_INTERNET_MESSAGE_ID) || '';
        } catch (e) {}

        // In-Reply-To header (Message-ID of the parent email)
        try {
            headers.inReplyTo = pa.GetProperty(PR_IN_REPLY_TO_ID) || '';
        } catch (e) {}

        // References header (chain of Message-IDs in thread)
        try {
            var refs = pa.GetProperty(PR_INTERNET_REFERENCES) || '';
            if (refs) {
                // References is a space or newline separated list of Message-IDs
                headers.references = refs.split(/[\s\r\n]+/).filter(function(r) {
                    return r && r.indexOf('@') > -1;
                });
            }
        } catch (e) {}

        // Conversation Index (Outlook's internal thread tracking)
        try {
            var convIdx = pa.GetProperty(PR_CONVERSATION_INDEX);
            if (convIdx) {
                headers.conversationIndex = convIdx;
            }
        } catch (e) {}

    } catch (e) {
        // PropertyAccessor not available or failed
    }

    return headers;
}

/**
 * Get the conversation start date from Outlook using GetConversation API
 * @param {Object} item - Outlook mail item
 * @returns {Date|null} Oldest date in conversation or null
 */
function getConversationStartDate(item) {
    try {
        var conv = item.GetConversation();
        if (conv) {
            var rootItems = conv.GetRootItems();
            if (rootItems && rootItems.Count > 0) {
                var oldestDate = null;
                for (var i = 1; i <= rootItems.Count; i++) {
                    var rootItem = rootItems.Item(i);
                    try {
                        var itemDate = rootItem.ReceivedTime || rootItem.SentOn;
                        if (itemDate) {
                            var d = new Date(itemDate);
                            if (!isNaN(d.getTime())) {
                                if (!oldestDate || d < oldestDate) {
                                    oldestDate = d;
                                }
                            }
                        }
                    } catch (e) {}
                }
                return oldestDate;
            }
        }
    } catch (e) {}
    return null;
}

/**
 * Get the original subject from conversation root item
 * @param {Object} item - Outlook mail item
 * @returns {string|null} Original subject or null
 */
function getConversationOriginalSubject(item) {
    try {
        var conv = item.GetConversation();
        if (conv) {
            var rootItems = conv.GetRootItems();
            if (rootItems && rootItems.Count > 0) {
                var oldestDate = null;
                var originalSubject = null;
                for (var i = 1; i <= rootItems.Count; i++) {
                    var rootItem = rootItems.Item(i);
                    try {
                        var itemDate = rootItem.ReceivedTime || rootItem.SentOn;
                        var itemSubject = rootItem.Subject;
                        if (itemDate) {
                            var d = new Date(itemDate);
                            if (!isNaN(d.getTime())) {
                                if (!oldestDate || d < oldestDate) {
                                    oldestDate = d;
                                    originalSubject = itemSubject;
                                }
                            }
                        }
                    } catch (e) {}
                }
                return originalSubject;
            }
        }
    } catch (e) {}
    return null;
}

