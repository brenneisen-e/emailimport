/**
 * File Save Module
 * Save export results to JSON file
 */

/**
 * Save exported emails to JSON file
 * @param {Array} emails - Array of email objects
 * @param {number} repliesFound - Count of replies found
 * @param {number} duplicatesSkipped - Count of duplicates skipped
 * @param {number} duplicatesWithReplies - Count of duplicates with new replies
 * @param {number} batchDuplicates - Count of duplicates within batch
 * @param {Object} debugInfo - Optional debug information
 */
function saveExportFile(emails, repliesFound, duplicatesSkipped, duplicatesWithReplies, batchDuplicates, debugInfo) {
    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        var shell = new ActiveXObject("WScript.Shell");
        var shellApp = new ActiveXObject("Shell.Application");

        // Try to get Downloads folder path
        var downloads;
        try {
            downloads = shellApp.NameSpace("shell:Downloads").Self.Path;
        } catch (e) {
            downloads = shell.ExpandEnvironmentStrings("%USERPROFILE%") + "\\Downloads";
        }

        // Create filename with date
        var today = new Date();
        var dateStr = today.getFullYear() + '-' +
            (today.getMonth() + 1 < 10 ? '0' : '') + (today.getMonth() + 1) + '-' +
            (today.getDate() < 10 ? '0' : '') + today.getDate();
        var filename = downloads + "\\hypercare_emails_" + dateStr + ".json";

        // Build JSON string with proper formatting
        var jsonContent = buildJsonString(emails, debugInfo);

        // Write using FileSystemObject (works with network paths)
        // CreateTextFile(path, overwrite, unicode)
        // unicode=true creates UTF-16LE file with BOM
        var file = fso.CreateTextFile(filename, true, true);
        file.Write(jsonContent);
        file.Close();

        // Store for direct import
        lastExportedEmails = emails;

        // Build success message with reply statistics
        var newCount = emails.length - duplicatesWithReplies;
        var successMsg = '<strong>' + newCount + '</strong> neue Emails';
        if (duplicatesWithReplies > 0) {
            successMsg += ', <strong>' + duplicatesWithReplies + '</strong> mit neuen Antworten';
        }
        if (duplicatesSkipped > 0) {
            successMsg += ', <strong>' + duplicatesSkipped + '</strong> bereits in Excel';
        }
        if (batchDuplicates > 0) {
            successMsg += ', <strong>' + batchDuplicates + '</strong> Duplikate';
        }

        // Add reply match statistics
        if (debugInfo) {
            successMsg += '<br><strong>Antworten:</strong> ' + debugInfo.emailsWithReplies + '/' + debugInfo.totalEmails + ' Emails (' + debugInfo.replyRate + ')';
            if (debugInfo.matchingStats) {
                var ms = debugInfo.matchingStats;
                successMsg += '<br><small>Matches: ConvId=' + ms.byConversationId + ', InReplyTo=' + ms.byInternetMessageId + ', Subject=' + ms.bySubject + '</small>';
            }
        }

        successMsg += '<br>Gespeichert: ' + filename;

        showExportSuccess(successMsg);

        // Step 3 completed - Export done
        completeStep(3);

    } catch (e) {
        showExportError("Fehler beim Speichern: " + e.message);
    }
}

/**
 * Build JSON string from email array
 * @param {Array} emails - Array of email objects
 * @param {Object} debugInfo - Optional debug information
 * @returns {string} JSON string
 */
function buildJsonString(emails, debugInfo) {
    // Sanitize all text fields before JSON encoding
    for (var i = 0; i < emails.length; i++) {
        var email = emails[i];
        email.text = sanitizeText(email.text || '');
        email.betreff = sanitizeText(email.betreff || '');
        email.von_name = sanitizeText(email.von_name || '');

        // Also sanitize replies
        if (email.antworten) {
            for (var r = 0; r < email.antworten.length; r++) {
                var reply = email.antworten[r];
                reply.text = sanitizeText(reply.text || '');
                reply.an = sanitizeText(reply.an || '');
            }
        }
    }

    // Create export object with optional debug info
    if (debugInfo) {
        var exportObj = {
            _exportInfo: debugInfo,
            emails: emails
        };
        return JSON.stringify(exportObj, null, 2);
    }

    return JSON.stringify(emails, null, 2);
}

/**
 * Open the web app for email editing
 */
function openWebApp() {
    try {
        var shell = new ActiveXObject("WScript.Shell");
        shell.Run("https://emailimport.pages.dev/");

        // Step 4 completed - Web-App opened
        completeStep(4);
    } catch (e) {
        showExportError("Fehler beim \u00D6ffnen der Web-App: " + e.message);
    }
}

/**
 * Open the Downloads folder
 */
function openDownloads() {
    try {
        var shell = new ActiveXObject("WScript.Shell");
        var shellApp = new ActiveXObject("Shell.Application");
        var downloads;
        try {
            downloads = shellApp.NameSpace("shell:Downloads").Self.Path;
        } catch (e) {
            downloads = shell.ExpandEnvironmentStrings("%USERPROFILE%") + "\\Downloads";
        }
        shell.Run('explorer "' + downloads + '"');
    } catch (e) {
        showExportError("Fehler beim \u00D6ffnen des Ordners: " + e.message);
    }
}

