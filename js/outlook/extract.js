/**
 * Email Extraction Module
 * Extract email data from Outlook mail items
 */

// MAPI Property for SMTP Address (resolves Exchange X500 to SMTP)
var PR_SMTP_ADDRESS_EXTRACT = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F";

/**
 * Get SMTP email address from sender (resolves Exchange X500 addresses)
 * @param {Object} mailItem - Outlook mail item
 * @returns {string} SMTP email address
 */
function getSenderSMTPAddressLegacy(mailItem) {
    try {
        var emailType = mailItem.SenderEmailType;

        if (emailType === 'EX') {
            try {
                var sender = mailItem.Sender;
                if (sender) {
                    try {
                        var exchUser = sender.GetExchangeUser();
                        if (exchUser && exchUser.PrimarySmtpAddress) {
                            return exchUser.PrimarySmtpAddress;
                        }
                    } catch (e1) {}

                    try {
                        var propAccessor = sender.PropertyAccessor;
                        var smtpAddr = propAccessor.GetProperty(PR_SMTP_ADDRESS_EXTRACT);
                        if (smtpAddr) return smtpAddr;
                    } catch (e2) {}
                }
            } catch (e3) {}

            try {
                var mailProp = mailItem.PropertyAccessor;
                var senderSmtp = mailProp.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x5D01001F");
                if (senderSmtp) return senderSmtp;
            } catch (e4) {}
        }

        return mailItem.SenderEmailAddress || '';
    } catch (e) {
        return mailItem.SenderEmailAddress || '';
    }
}

/**
 * Extract all data from an Outlook mail item
 * @param {Object} mailItem - Outlook mail item
 * @returns {Object} Extracted email data
 */
function extractEmail(mailItem) {
    var email = {
        emailId: '',
        conversationId: '',
        datum: '',
        von_email: '',
        von_name: '',
        betreff: '',
        text: '',
        htmlBody: '',
        anhaenge: [],
        antworten: [],  // Initialize replies array (old version had this)
        kategorie: '',  // Outlook category (old field name for web app)
        outlookKategorie: '',
        folderName: '',
        conversationStartDate: '',
        originalSubject: '',
        internetMessageId: '',
        inReplyTo: '',
        references: []
    };

    try {
        // Basic identifiers
        try { email.emailId = mailItem.EntryID || ''; } catch (e) {}
        try { email.conversationId = mailItem.ConversationID || ''; } catch (e) {}

        // Timestamps
        try {
            var receivedTime = new Date(mailItem.ReceivedTime);
            if (!isNaN(receivedTime.getTime())) {
                email.datum = formatDate(receivedTime);
            }
        } catch (e) {}

        // Sender (resolve Exchange X500 to SMTP)
        try {
            email.von_email = getSenderSMTPAddressLegacy(mailItem);
            email.von_name = mailItem.SenderName || '';
        } catch (e) {}

        // Subject
        try { email.betreff = mailItem.Subject || '(Kein Betreff)'; } catch (e) {}

        // Body
        try {
            email.text = mailItem.Body || '';
            if (email.text.length > 50000) email.text = email.text.substring(0, 50000) + '...';
        } catch (e) {}

        // HTML Body
        try {
            email.htmlBody = mailItem.HTMLBody || '';
            if (email.htmlBody.length > 500000) email.htmlBody = '';
        } catch (e) {}

        // Attachments
        try {
            for (var a = 1; a <= mailItem.Attachments.Count; a++) {
                email.anhaenge.push(mailItem.Attachments.Item(a).FileName);
            }
        } catch (e) {}

        // Categories
        try {
            var cats = mailItem.Categories || '';
            if (cats) {
                var firstCat = cats.split(',')[0].trim();
                email.outlookKategorie = firstCat;
                email.kategorie = firstCat;  // Old field name for web app compatibility
            }
        } catch (e) {}

        // Folder name
        try { email.folderName = mailItem.Parent.Name || ''; } catch (e) {}

        // Conversation start date
        try {
            var convStart = getConversationStartDate(mailItem);
            if (convStart) {
                email.conversationStartDate = formatDate(convStart);
            }
        } catch (e) {}

        // Original subject from conversation root
        try {
            var origSubject = getConversationOriginalSubject(mailItem);
            if (origSubject) {
                email.originalSubject = origSubject;
            }
        } catch (e) {}

        // MAPI headers for thread linking
        var headers = getEmailHeaders(mailItem);
        email.internetMessageId = headers.internetMessageId;
        email.inReplyTo = headers.inReplyTo;
        email.references = headers.references;

    } catch (e) {}

    return email;
}

/**
 * Extract email in unified format for batch processing
 * @param {Object} mailItem - Outlook mail item
 * @param {string} folderType - 'inbox' or 'sent'
 * @returns {Object} Unified email object
 */
function extractEmailUnified(mailItem, folderType) {
    var email = {
        entryId: '',
        conversationId: '',
        internetMessageId: '',
        inReplyTo: '',
        references: [],
        receivedTime: null,
        sentOn: null,
        timestamp: null,
        von_email: '',
        von_name: '',
        recipients: [],
        betreff: '',
        text: '',
        htmlBody: '',
        anhaenge: [],
        outlookKategorie: '',
        folderType: folderType || 'inbox',
        folderName: '',
        isIncoming: true,
        threadPosition: 0,
        threadDepth: 0,
        parentMessageId: null,
        parentFoundInThread: false,
        isThreadRoot: false,
        newContent: '',
        normalizedSubject: ''
    };

    try {
        // Identifiers
        try { email.entryId = mailItem.EntryID || ''; } catch (e) {}
        try { email.conversationId = mailItem.ConversationID || ''; } catch (e) {}

        // Timestamps
        try {
            email.receivedTime = new Date(mailItem.ReceivedTime);
            if (isNaN(email.receivedTime.getTime())) email.receivedTime = null;
        } catch (e) {}

        try {
            email.sentOn = new Date(mailItem.SentOn);
            if (isNaN(email.sentOn.getTime())) email.sentOn = null;
        } catch (e) {}

        // Sender (resolve Exchange X500 to SMTP)
        try {
            email.von_email = getSenderSMTPAddressLegacy(mailItem);
            email.von_name = mailItem.SenderName || '';
        } catch (e) {}

        // Recipients
        try {
            var recipients = mailItem.Recipients;
            for (var r = 1; r <= recipients.Count; r++) {
                try {
                    var recip = recipients.Item(r);
                    email.recipients.push({
                        name: recip.Name || '',
                        email: recip.Address || '',
                        type: recip.Type // 1=To, 2=CC, 3=BCC
                    });
                } catch (e) {}
            }
        } catch (e) {}

        // Subject
        try { email.betreff = mailItem.Subject || '(Kein Betreff)'; } catch (e) {}

        // Body
        try {
            email.text = mailItem.Body || '';
            if (email.text.length > 50000) email.text = email.text.substring(0, 50000) + '...';
        } catch (e) {}

        // HTML Body
        try {
            email.htmlBody = mailItem.HTMLBody || '';
            if (email.htmlBody.length > 500000) email.htmlBody = '';
        } catch (e) {}

        // Attachments
        try {
            for (var a = 1; a <= mailItem.Attachments.Count; a++) {
                email.anhaenge.push(mailItem.Attachments.Item(a).FileName);
            }
        } catch (e) {}

        // Categories
        try {
            var cats = mailItem.Categories || '';
            if (cats) {
                email.outlookKategorie = cats.split(',')[0].trim();
            }
        } catch (e) {}

        // Folder name
        try { email.folderName = mailItem.Parent.Name || ''; } catch (e) {}

        // MAPI headers for thread linking
        var headers = getEmailHeaders(mailItem);
        email.internetMessageId = headers.internetMessageId;
        email.inReplyTo = headers.inReplyTo;
        email.references = headers.references;

        // Calculate timestamp (use sentOn for sent items, receivedTime for inbox)
        if (folderType === 'sent' && email.sentOn && !isNaN(email.sentOn.getTime())) {
            email.timestamp = email.sentOn;
        } else if (email.receivedTime && !isNaN(email.receivedTime.getTime())) {
            email.timestamp = email.receivedTime;
        } else {
            email.timestamp = new Date();
        }

        // Calculate direction
        email.isIncoming = (folderType !== 'sent');

        // Normalized subject for matching
        email.normalizedSubject = normalizeSubject(email.betreff);

    } catch (e) {}

    return email;
}

