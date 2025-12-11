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
 */
function saveExportFile(emails, repliesFound, duplicatesSkipped, duplicatesWithReplies, batchDuplicates) {
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
        var jsonContent = buildJsonString(emails);

        // Write as UTF-16LE (compatible with HTA/JScript)
        var stream = new ActiveXObject("ADODB.Stream");
        stream.Type = 2; // adTypeText
        stream.Charset = "UTF-16LE";
        stream.Open();
        stream.WriteText(jsonContent);
        stream.SaveToFile(filename, 2); // adSaveCreateOverWrite
        stream.Close();

        // Store for direct import
        lastExportedEmails = emails;

        // Show success
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
 * @returns {string} JSON string
 */
function buildJsonString(emails) {
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

    return JSON.stringify(emails, null, 2);
}

/**
 * Open the web app for email editing
 */
function openWebApp() {
    try {
        var shell = new ActiveXObject("WScript.Shell");
        shell.Run("https://brenneisen-e.github.io/emailimport/");

        // Step 4 completed - Web-App opened
        completeStep(4);
    } catch (e) {
        showExportError("Fehler beim Öffnen der Web-App: " + e.message);
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
        showExportError("Fehler beim Öffnen des Ordners: " + e.message);
    }
}

