/**
 * Date Utilities
 * Formatting, parsing, and date extraction functions
 */

/**
 * Format a date as ISO 8601 string (YYYY-MM-DDTHH:MM:SS)
 */
function formatDate(date) {
    var d = new Date(date);
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) + 'T' + pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds());
}

/**
 * Pad a number with leading zero
 */
function pad(n) {
    return n < 10 ? '0' + n : n;
}

/**
 * Format date for Outlook Restrict() filter
 */
function formatDateForRestrict(date) {
    var d = new Date(date);
    return (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear() + ' ' + d.getHours() + ':' + d.getMinutes();
}

/**
 * Create a date key for duplicate detection (YYYY-MM-DD format)
 */
function formatDateKey(datum) {
    if (!datum) return '';
    try {
        var d = new Date(datum);
        if (isNaN(d.getTime())) return '';
        return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
    } catch (e) {
        return '';
    }
}

/**
 * German months mapping for date parsing
 */
var germanMonths = {
    'januar': 0, 'februar': 1, 'märz': 2, 'maerz': 2, 'april': 3,
    'mai': 4, 'juni': 5, 'juli': 6, 'august': 7, 'september': 8,
    'oktober': 9, 'november': 10, 'dezember': 11
};

/**
 * English months mapping for date parsing
 */
var englishMonths = {
    'january': 0, 'february': 1, 'march': 2, 'april': 3,
    'may': 4, 'june': 5, 'july': 6, 'august': 7, 'september': 8,
    'october': 9, 'november': 10, 'december': 11
};

/**
 * Extract the oldest date from quoted email headers in text
 * Used as fallback when conversation metadata is not available
 */
function extractOldestDateFromQuotes(text) {
    if (!text) return null;

    var dates = [];

    // German pattern: "Gesendet: Montag, 1. Dezember 2025 12:50"
    var germanPattern = /Gesendet:\s*\w+,\s*(\d{1,2})\.\s*(\w+)\s+(\d{4})\s+(\d{1,2}):(\d{2})/gi;
    var match;

    while ((match = germanPattern.exec(text)) !== null) {
        var day = parseInt(match[1], 10);
        var monthName = match[2].toLowerCase();
        var year = parseInt(match[3], 10);
        var hour = parseInt(match[4], 10);
        var minute = parseInt(match[5], 10);

        var month = germanMonths[monthName];
        if (month !== undefined) {
            var date = new Date(year, month, day, hour, minute, 0);
            if (!isNaN(date.getTime())) {
                dates.push(date);
            }
        }
    }

    // English pattern: "Sent: Monday, December 1, 2025 12:50 PM"
    var englishPattern = /Sent:\s*\w+,\s*(\w+)\s+(\d{1,2}),\s*(\d{4})\s+(\d{1,2}):(\d{2})(?:\s*(AM|PM))?/gi;

    while ((match = englishPattern.exec(text)) !== null) {
        var monthName = match[1].toLowerCase();
        var day = parseInt(match[2], 10);
        var year = parseInt(match[3], 10);
        var hour = parseInt(match[4], 10);
        var minute = parseInt(match[5], 10);
        var ampm = match[6] ? match[6].toUpperCase() : '';

        var month = englishMonths[monthName];
        if (month !== undefined) {
            if (ampm === 'PM' && hour < 12) hour += 12;
            if (ampm === 'AM' && hour === 12) hour = 0;

            var date = new Date(year, month, day, hour, minute, 0);
            if (!isNaN(date.getTime())) {
                dates.push(date);
            }
        }
    }

    // Return the oldest date found
    if (dates.length === 0) return null;

    var oldest = dates[0];
    for (var i = 1; i < dates.length; i++) {
        if (dates[i] < oldest) {
            oldest = dates[i];
        }
    }

    return oldest;
}

/**
 * Extract date from HTML email headers
 * Supports German and English Outlook formats
 * Returns: { date: Date object, success: boolean }
 */
function extractDateFromHtmlHeader(htmlBody) {
    if (!htmlBody) return { date: null, success: false };

    var result = { date: null, success: false };
    var match;

    // Pattern 1: German HTML format - "Gesendet:</b> Dienstag, 9. Dezember 2025 16:05"
    var germanHtmlPattern = /Gesendet:<\/b>\s*\w+,\s*(\d{1,2})\.\s*(\w+)\s+(\d{4})\s+(\d{1,2}):(\d{2})/i;
    match = germanHtmlPattern.exec(htmlBody);
    if (match) {
        var day = parseInt(match[1], 10);
        var monthName = match[2].toLowerCase();
        var year = parseInt(match[3], 10);
        var hour = parseInt(match[4], 10);
        var minute = parseInt(match[5], 10);
        var month = germanMonths[monthName];

        if (month !== undefined) {
            var date = new Date(year, month, day, hour, minute, 0);
            if (!isNaN(date.getTime())) {
                result.date = date;
                result.success = true;
                return result;
            }
        }
    }

    // Pattern 2: English HTML format - "Sent:</b> Tuesday, December 9, 2025 3:28 PM"
    var englishHtmlPattern = /Sent:<\/b>\s*\w+,\s*(\w+)\s+(\d{1,2}),\s*(\d{4})\s+(\d{1,2}):(\d{2})(?:\s*(AM|PM))?/i;
    match = englishHtmlPattern.exec(htmlBody);
    if (match) {
        var monthName = match[1].toLowerCase();
        var day = parseInt(match[2], 10);
        var year = parseInt(match[3], 10);
        var hour = parseInt(match[4], 10);
        var minute = parseInt(match[5], 10);
        var ampm = match[6] ? match[6].toUpperCase() : '';
        var month = englishMonths[monthName];

        if (month !== undefined) {
            if (ampm === 'PM' && hour < 12) hour += 12;
            if (ampm === 'AM' && hour === 12) hour = 0;

            var date = new Date(year, month, day, hour, minute, 0);
            if (!isNaN(date.getTime())) {
                result.date = date;
                result.success = true;
                return result;
            }
        }
    }

    // Pattern 3: German HTML format without bold
    var germanPlainPattern = /Gesendet:\s*\w+,\s*(\d{1,2})\.\s*(\w+)\s+(\d{4})\s+(\d{1,2}):(\d{2})/i;
    match = germanPlainPattern.exec(htmlBody);
    if (match) {
        var day = parseInt(match[1], 10);
        var monthName = match[2].toLowerCase();
        var year = parseInt(match[3], 10);
        var hour = parseInt(match[4], 10);
        var minute = parseInt(match[5], 10);
        var month = germanMonths[monthName];

        if (month !== undefined) {
            var date = new Date(year, month, day, hour, minute, 0);
            if (!isNaN(date.getTime())) {
                result.date = date;
                result.success = true;
                return result;
            }
        }
    }

    // Pattern 4: English HTML format without bold
    var englishPlainPattern = /Sent:\s*\w+,\s*(\w+)\s+(\d{1,2}),\s*(\d{4})\s+(\d{1,2}):(\d{2})(?:\s*(AM|PM))?/i;
    match = englishPlainPattern.exec(htmlBody);
    if (match) {
        var monthName = match[1].toLowerCase();
        var day = parseInt(match[2], 10);
        var year = parseInt(match[3], 10);
        var hour = parseInt(match[4], 10);
        var minute = parseInt(match[5], 10);
        var ampm = match[6] ? match[6].toUpperCase() : '';
        var month = englishMonths[monthName];

        if (month !== undefined) {
            if (ampm === 'PM' && hour < 12) hour += 12;
            if (ampm === 'AM' && hour === 12) hour = 0;

            var date = new Date(year, month, day, hour, minute, 0);
            if (!isNaN(date.getTime())) {
                result.date = date;
                result.success = true;
                return result;
            }
        }
    }

    return result;
}
