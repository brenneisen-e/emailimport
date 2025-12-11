/**
 * Email Extraction Module
 * Extract email data from Outlook mail items
 */

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

        // Sender
        try {
            email.von_email = mailItem.SenderEmailAddress || '';
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
                email.outlookKategorie = cats.split(',')[0].trim();
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

        // Sender
        try {
            email.von_email = mailItem.SenderEmailAddress || '';
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

