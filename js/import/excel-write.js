/**
 * Excel Write Module
 * Write email data to Excel worksheet
 */

// Global for tracking current field being written (for error debugging)
var _writeRowCurrentField = '';

/**
 * Write a single email row to Excel
 * @param {Object} worksheet - Excel worksheet
 * @param {number} row - Row number
 * @param {Object} email - Email data
 */
function writeEmailRow(worksheet, row, email) {
    _writeRowCurrentField = 'Start';

    // Copy formatting from row above
    if (row > 2) {
        try {
            _writeRowCurrentField = 'Format kopieren';
            var sourceRange = worksheet.Range(worksheet.Cells(row - 1, 1), worksheet.Cells(row - 1, 25));
            var targetRange = worksheet.Range(worksheet.Cells(row, 1), worksheet.Cells(row, 25));
            sourceRange.Copy();
            targetRange.PasteSpecial(-4122); // xlPasteFormats
            try { worksheet.Application.CutCopyMode = false; } catch(e) {}
        } catch (e) {}
    }

    _writeRowCurrentField = 'Antworten formatieren';
    var repliesText = '';
    try {
        repliesText = formatReplies(email.antworten);
        // Excel cell limit is 32767 chars, keep it safe at 30000
        if (repliesText.length > 30000) {
            repliesText = repliesText.substring(0, 30000) + '\n\n[... weitere Antworten gekürzt ...]';
        }
    } catch (repliesErr) {
        repliesText = '[Fehler: ' + repliesErr.message + ']';
    }

    _writeRowCurrentField = 'provisionData lesen';
    var pd = email.provisionData || {};

    // Absender
    _writeRowCurrentField = 'Absender erstellen';
    var absender = email.von_name || email.von_email || '';
    if (pd.vorname_vermittler || pd.nachname_vermittler) {
        absender = ((pd.vorname_vermittler || '') + ' ' + (pd.nachname_vermittler || '')).trim();
    }
    absender = fixEncoding(absender);

    // Agentur
    _writeRowCurrentField = 'Agentur erstellen';
    var agentur = fixEncoding(email.agentur || '');

    // Anfrage (mit Quote-Entfernung fuer Memory-Schutz)
    _writeRowCurrentField = 'Anfrage erstellen';
    var anfrage = '';
    if (email.anfrage && email.anfrage.trim()) {
        anfrage = removeQuotedContent(email.anfrage);
        anfrage = fixEncoding(anfrage);
    } else if (pd.nachricht) {
        anfrage = fixEncoding(pd.nachricht);
    } else {
        anfrage = removeQuotedContent(email.text || '');
        anfrage = fixEncoding(anfrage);
    }

    // BD-Nummer
    _writeRowCurrentField = 'BD-Nummer erstellen';
    var bdNummer = pd.vermittlernummer_vermittler || email.vermittlernr || '';

    // Betreff
    _writeRowCurrentField = 'Betreff erstellen';
    var betreff = fixEncoding(email.betreff || '');

    // Kategorie
    _writeRowCurrentField = 'Kategorie erstellen';
    var kategorie = fixEncoding(email.kategorie || '');

    // Status
    _writeRowCurrentField = 'Status erstellen';
    var status = fixEncoding(email.status || 'Neu');

    // Clusters
    _writeRowCurrentField = 'Clusters erstellen';
    var clusters = fixEncoding((email.clusters || []).join(', '));

    // Kommentar
    _writeRowCurrentField = 'Kommentar erstellen';
    var kommentar = fixEncoding(email.kommentar || '');

    // Bearbeiter
    _writeRowCurrentField = 'Bearbeiter erstellen';
    var bearbeiter = '';
    if (email.bearbeiter) {
        bearbeiter = '' + email.bearbeiter;
    } else if (email.outlookKategorie) {
        bearbeiter = '' + email.outlookKategorie;
    }

    // Antwort-Text
    _writeRowCurrentField = 'Antwort-Text erstellen';
    var antwortText = '';
    var userAntwort = '';
    try {
        if (email.antwort) {
            userAntwort = String(email.antwort);
            userAntwort = fixEncoding(userAntwort);
        }
    } catch (e) {
        userAntwort = '[Fehler beim Lesen der Benutzerantwort]';
    }

    if (userAntwort && repliesText) {
        antwortText = userAntwort + '\n\n--- Gefundene Antworten ---\n\n' + repliesText;
    } else {
        antwortText = userAntwort || repliesText || '';
    }

    // Calculate "Datum neuste Antwort" (latest reply date)
    _writeRowCurrentField = 'Neuste Antwort berechnen';
    var neusteDatum = '';
    try {
        if (email.antworten && email.antworten.length > 0) {
            var latestDate = null;
            for (var ai = 0; ai < email.antworten.length; ai++) {
                var replyDate = email.antworten[ai].datum;
                if (replyDate) {
                    var rd = new Date(replyDate);
                    if (!isNaN(rd.getTime()) && (!latestDate || rd > latestDate)) {
                        latestDate = rd;
                    }
                }
            }
            if (latestDate) {
                neusteDatum = latestDate.getFullYear() + '-' +
                    pad(latestDate.getMonth() + 1) + '-' +
                    pad(latestDate.getDate()) + ' ' +
                    pad(latestDate.getHours()) + ':' +
                    pad(latestDate.getMinutes()) + ':' +
                    pad(latestDate.getSeconds());
            }
        }
    } catch (e) {}

    // Write cells
    _writeRowCurrentField = 'Spalte 1 (ID)';
    worksheet.Cells(row, 1).Value = 'HC-' + (row - 1);

    _writeRowCurrentField = 'Spalte 2 (Datum Anfrage)';
    try {
        var datumValue = email.datum || '';
        if (datumValue) {
            var d = new Date(datumValue);
            if (!isNaN(d.getTime())) {
                var excelDateStr = d.getFullYear() + '-' +
                    pad(d.getMonth() + 1) + '-' +
                    pad(d.getDate()) + ' ' +
                    pad(d.getHours()) + ':' +
                    pad(d.getMinutes()) + ':' +
                    pad(d.getSeconds());
                worksheet.Cells(row, 2).Value = excelDateStr;
                worksheet.Cells(row, 2).NumberFormat = 'DD.MM.YYYY HH:MM';
            } else {
                worksheet.Cells(row, 2).Value = datumValue;
            }
        }
    } catch (e) {
        worksheet.Cells(row, 2).Value = email.datum || '';
    }

    _writeRowCurrentField = 'Spalte 3 (Datum neuste Antwort)';
    if (neusteDatum) {
        worksheet.Cells(row, 3).Value = neusteDatum;
        worksheet.Cells(row, 3).NumberFormat = 'DD.MM.YYYY HH:MM';
    }

    _writeRowCurrentField = 'Spalte 4 (Kanal)';
    worksheet.Cells(row, 4).Value = 'E-Mail';

    _writeRowCurrentField = 'Spalte 5 (BD-Nummer)';
    worksheet.Cells(row, 5).Value = bdNummer;

    _writeRowCurrentField = 'Spalte 6 (Agentur)';
    worksheet.Cells(row, 6).Value = agentur;

    _writeRowCurrentField = 'Spalte 7 (Absender)';
    worksheet.Cells(row, 7).Value = absender;

    _writeRowCurrentField = 'Spalte 8 (Kategorie)';
    worksheet.Cells(row, 8).Value = kategorie;

    _writeRowCurrentField = 'Spalte 9 (Status)';
    worksheet.Cells(row, 9).Value = status;

    _writeRowCurrentField = 'Spalte 10 (In Bearbeitung von)';
    worksheet.Cells(row, 10).Value = bearbeiter;

    _writeRowCurrentField = 'Spalte 11 (Cluster)';
    worksheet.Cells(row, 11).Value = clusters;

    _writeRowCurrentField = 'Spalte 12 (leer)';
    // Column 12 reserved

    _writeRowCurrentField = 'Spalte 13 (Betreff)';
    worksheet.Cells(row, 13).Value = betreff;

    _writeRowCurrentField = 'Spalte 14 (Anfrage)';
    // Truncate and sanitize for Excel (max 30000 chars for safety)
    var safeAnfrage = truncateForExcel(sanitizeForExcel(fixEncoding(anfrage)), 30000);
    worksheet.Cells(row, 14).Value = safeAnfrage;

    _writeRowCurrentField = 'Spalte 15 (Antwort)';
    // Truncate and sanitize for Excel (max 30000 chars for safety)
    var safeAntwort = truncateForExcel(sanitizeForExcel(fixEncoding(antwortText)), 30000);
    worksheet.Cells(row, 15).Value = safeAntwort;

    // Bold timestamps in reply cell
    _writeRowCurrentField = 'Timestamps fetten';
    boldTimestamps(worksheet.Cells(row, 15));

    _writeRowCurrentField = 'Spalte 16-21 (leer)';
    // Columns 16-21 reserved

    _writeRowCurrentField = 'Spalte 22 (Kommentar)';
    worksheet.Cells(row, 22).Value = kommentar;

    _writeRowCurrentField = 'Spalte 23 (ConversationID)';
    worksheet.Cells(row, 23).Value = email.conversationId || '';

    _writeRowCurrentField = 'Spalte 24 (EmailID)';
    worksheet.Cells(row, 24).Value = email.emailId || '';

    _writeRowCurrentField = 'Spalte 25 (ReplyIDs)';
    var replyIdList = [];
    if (email.antworten) {
        for (var ri = 0; ri < email.antworten.length; ri++) {
            if (email.antworten[ri].replyId) {
                replyIdList.push(email.antworten[ri].replyId);
            }
        }
    }
    worksheet.Cells(row, 25).Value = replyIdList.join(',');

    _writeRowCurrentField = 'Fertig';
}

/**
 * Format replies array to text
 * @param {Array} antworten - Array of reply objects
 * @returns {string} Formatted text
 */
function formatReplies(antworten) {
    if (!antworten) return '';

    try {
        var len = 0;
        try { len = antworten.length || 0; } catch (e) { return ''; }
        if (len === 0) return '';

        var text = '';
        for (var r = 0; r < len && r < 20; r++) {
            try {
                var reply = antworten[r];
                if (!reply) continue;

                var replyDatum = '';
                try {
                    if (reply.datum !== undefined && reply.datum !== null) {
                        replyDatum = sanitizeForExcel('' + reply.datum);
                    }
                } catch (e) { replyDatum = ''; }

                var replyText = '';
                try {
                    if (reply.text !== undefined && reply.text !== null) {
                        replyText = '' + reply.text;
                        // WICHTIG: Zuerst zitierte Inhalte entfernen (sonst Memory-Probleme)
                        replyText = removeQuotedContent(replyText);
                        replyText = fixEncoding(replyText);
                        replyText = sanitizeForExcel(replyText);
                    }
                } catch (e) { replyText = '[Fehler]'; }

                // Kuerzen auf max 2000 Zeichen pro Antwort
                if (replyText.length > 2000) {
                    replyText = replyText.substring(0, 2000) + '...';
                }

                var direction = '';
                if (reply.isIncoming) {
                    direction = '[Eingehend]';
                } else if (reply.an) {
                    direction = '[An: ' + sanitizeText('' + reply.an) + ']';
                }

                text += '=== ' + replyDatum + ' ' + direction + ' ===\n' + replyText + '\n\n';
            } catch (replyErr) {
                text += '=== [Fehler beim Lesen dieser Antwort] ===\n\n';
            }
        }
        return sanitizeForExcel(text.trim());
    } catch (e) {
        return '[Fehler: ' + (e.message || e) + ']';
    }
}

/**
 * Bold timestamp patterns in a cell
 * @param {Object} cell - Excel cell
 */
function boldTimestamps(cell) {
    try {
        var text = cell.Value || '';
        if (!text || typeof text !== 'string') return;

        // First, ensure entire cell is NOT bold (reset from format copy)
        try {
            cell.Font.Bold = false;
        } catch (e) {}

        // Now apply bold only to timestamp patterns
        var regex = new RegExp('===[^=]+===', 'g');
        var match;

        while ((match = regex.exec(text)) !== null) {
            var start = match.index + 1; // Excel Characters is 1-indexed
            var length = match[0].length;
            try {
                cell.Characters(start, length).Font.Bold = true;
            } catch (e) {}
        }
    } catch (e) {}
}

/**
 * Read existing data from Excel for duplicate checking
 * @param {Object} worksheet - Excel worksheet
 * @returns {Object} Map of key -> row number
 */
function readExistingData(worksheet) {
    var data = {};
    try {
        var lastRow = worksheet.Cells(worksheet.Rows.Count, 1).End(-4162).Row;
        if (lastRow <= 1) return data;

        for (var row = 2; row <= lastRow; row++) {
            var datum = worksheet.Cells(row, 2).Value || '';
            var betreff = worksheet.Cells(row, 13).Value || '';  // Column 13 = Betreff

            if (datum || betreff) {
                var key = createEmailKey(datum, betreff);
                data[key] = row;
            }
        }
    } catch (e) {}
    return data;
}

/**
 * Create unique key from date and subject
 * @param {*} datum - Date value
 * @param {string} betreff - Subject
 * @returns {string} Key string
 */
function createEmailKey(datum, betreff) {
    var dateStr = '';
    if (datum) {
        try {
            var d = new Date(datum);
            dateStr = d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate();
        } catch (e) {
            dateStr = String(datum).substring(0, 10);
        }
    }
    var subjectStr = String(betreff || '').toLowerCase().trim().substring(0, 100);
    return dateStr + '|' + subjectStr;
}

