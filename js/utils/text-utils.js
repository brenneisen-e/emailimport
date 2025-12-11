/**
 * Text Utilities
 * Text sanitization, encoding fixes, and email quote removal
 */

/**
 * Sanitize text for safe display (remove null bytes, normalize whitespace)
 */
function sanitizeText(text) {
    if (!text) return '';
    return text
        .replace(/\x00/g, '')  // Remove null bytes
        .replace(/\r\n/g, '\n')  // Normalize line endings
        .replace(/\r/g, '\n')
        .trim();
}

/**
 * Fix encoding issues (common in German emails)
 * Handles UTF-8 double-encoding and other common encoding problems
 */
function fixEncoding(text) {
    if (!text) return '';
    return text
        // UTF-8 double-encoded German umlauts
        .replace(/Ã¤/g, 'ä')
        .replace(/Ã¶/g, 'ö')
        .replace(/Ã¼/g, 'ü')
        .replace(/Ã„/g, 'Ä')
        .replace(/Ã–/g, 'Ö')
        .replace(/Ãœ/g, 'Ü')
        .replace(/ÃŸ/g, 'ß')
        // Additional common double-encoding patterns
        .replace(/Ã¼/g, 'ü')
        .replace(/Ã¶/g, 'ö')
        .replace(/Ã¤/g, 'ä')
        .replace(/Ã©/g, 'é')
        .replace(/Ã¨/g, 'è')
        .replace(/Ã /g, 'à')
        .replace(/Ã¢/g, 'â')
        .replace(/Ã®/g, 'î')
        .replace(/Ã´/g, 'ô')
        .replace(/Ã»/g, 'û')
        .replace(/Ã§/g, 'ç')
        // Special characters
        .replace(/â‚¬/g, '€')
        .replace(/â€"/g, '–')
        .replace(/â€"/g, '—')
        .replace(/â€˜/g, "'")
        .replace(/â€™/g, "'")
        .replace(/â€œ/g, '"')
        .replace(/â€/g, '"')
        .replace(/â€¦/g, '...')
        .replace(/Â /g, ' ')
        .replace(/Â´/g, "'")
        .replace(/Â°/g, '°')
        .replace(/Â§/g, '§')
        .replace(/Â©/g, '©')
        .replace(/Â®/g, '®')
        // Remove invisible/problematic Unicode characters
        .replace(/[\u200B-\u200D\uFEFF]/g, '')  // Zero-width chars
        .replace(/\u00A0/g, ' ');  // Non-breaking space to regular space
}

/**
 * Truncate text for Excel cell (max 32767 chars, we use 30000 for safety)
 */
function truncateForExcel(text, maxLength) {
    if (!text) return '';
    maxLength = maxLength || 30000;
    if (text.length <= maxLength) return text;
    return text.substring(0, maxLength) + '\n[... Text gekürzt ...]';
}

/**
 * Sanitize text for Excel (handle special characters)
 */
function sanitizeForExcel(text) {
    if (!text) return '';

    // Remove control characters except newlines and tabs
    var result = text.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');

    // Replace common problematic characters
    result = result.replace(/\u2018|\u2019/g, "'");  // Smart quotes
    result = result.replace(/\u201C|\u201D/g, '"');  // Smart double quotes
    result = result.replace(/\u2013/g, '-');         // En dash
    result = result.replace(/\u2014/g, '--');        // Em dash
    result = result.replace(/\u2026/g, '...');       // Ellipsis

    return result;
}

/**
 * Normalize subject line for thread matching
 * Removes RE:, AW:, FW:, WG: prefixes and normalizes to lowercase
 */
function normalizeSubject(subject) {
    if (!subject) return '';
    return subject
        .replace(/^(RE|AW|FW|WG|Fwd|Antwort):\s*/gi, '')
        .replace(/^(RE|AW|FW|WG|Fwd|Antwort):\s*/gi, '')  // Remove nested prefixes
        .replace(/^(RE|AW|FW|WG|Fwd|Antwort):\s*/gi, '')
        .toLowerCase()
        .trim();
}

/**
 * Strip RE:/AW:/FW:/WG: prefixes from subject
 */
function stripSubjectPrefixes(subject) {
    if (!subject) return '';
    return subject
        .replace(/^(RE|AW|FW|WG|Fwd|Antwort):\s*/gi, '')
        .replace(/^(RE|AW|FW|WG|Fwd|Antwort):\s*/gi, '')
        .replace(/^(RE|AW|FW|WG|Fwd|Antwort):\s*/gi, '')
        .trim();
}

/**
 * Remove quoted content aggressively (for HTA Excel import)
 * Cuts off everything after common quote markers to prevent memory issues
 */
function removeQuotedContent(text) {
    if (!text) return '';

    // Limit input to prevent memory issues with huge strings
    var safeText = text.length > 50000 ? text.substring(0, 50000) : text;

    // Find the earliest quote marker and cut there
    var cutPatterns = [
        /_{5,}\s*\r?\n\s*Von:/i,                    // _____ Von:
        /_{5,}\s*\r?\n\s*From:/i,                   // _____ From:
        /-{5,}\s*Urspr[uü]ngliche\s+Nachricht/i,    // ----- Urspruengliche Nachricht
        /-{5,}\s*Original\s*Message/i,              // ----- Original Message
        /-{5,}\s*Weitergeleitete\s+Nachricht/i,     // ----- Weitergeleitete Nachricht
        /\r?\n\s*Von:\s+[^\r\n]+\r?\n\s*Gesendet:/i,  // Von: ... Gesendet:
        /\r?\n\s*From:\s+[^\r\n]+\r?\n\s*Sent:/i,     // From: ... Sent:
        /\r?\n\s*Am\s+\d[^\r\n]+schrieb[^\r\n]*:/i,   // Am ... schrieb ...:
        /\r?\n\s*On\s+[^\r\n]+wrote:/i                // On ... wrote:
    ];

    var result = safeText;
    var earliestCut = result.length;

    for (var i = 0; i < cutPatterns.length; i++) {
        var match = result.search(cutPatterns[i]);
        if (match > 0 && match < earliestCut) {
            earliestCut = match;
        }
    }

    if (earliestCut < result.length) {
        result = result.substring(0, earliestCut);
    }

    return result.trim();
}

/**
 * Remove email signature from text
 */
function removeSignature(text) {
    if (!text) return '';

    // Common signature delimiters
    var patterns = [
        /\n--\s*\n[\s\S]*$/,                    // Standard "-- " delimiter
        /\nMit freundlichen Grüßen[\s\S]*$/i,   // German
        /\nBeste Grüße[\s\S]*$/i,               // German
        /\nViele Grüße[\s\S]*$/i,               // German
        /\nKind regards[\s\S]*$/i,              // English
        /\nBest regards[\s\S]*$/i,              // English
        /\nRegards[\s\S]*$/i,                   // English
        /\n_{3,}[\s\S]*$/,                      // Underscores separator
        /\n-{3,}[\s\S]*$/                       // Dashes separator
    ];

    var result = text;
    for (var i = 0; i < patterns.length; i++) {
        result = result.replace(patterns[i], '');
    }

    return result.trim();
}

/**
 * Remove quoted email text (replies/forwards)
 * Keeps only the new content
 */
function removeEmailQuotes(text) {
    if (!text) return '';

    var lines = text.split('\n');
    var result = [];
    var inQuote = false;
    var foundQuoteStart = false;

    for (var i = 0; i < lines.length; i++) {
        var line = lines[i];
        var trimmedLine = line.trim();

        // Check for quote start markers
        if (!foundQuoteStart) {
            // German: "Von: ... Gesendet: ..."
            if (/^Von:\s*.+/i.test(trimmedLine) ||
                /^From:\s*.+/i.test(trimmedLine) ||
                /^-----\s*Urspr.+Nachricht/i.test(trimmedLine) ||
                /^-----\s*Original\s*Message/i.test(trimmedLine) ||
                /^_{5,}/.test(trimmedLine) ||
                /^-{5,}/.test(trimmedLine)) {

                // Look ahead to confirm this is a quote header
                if (i + 1 < lines.length) {
                    var nextLine = lines[i + 1].trim();
                    if (/^(Gesendet|Sent|An|To|Betreff|Subject|Cc|Von|From):/i.test(nextLine) ||
                        nextLine === '') {
                        foundQuoteStart = true;
                        inQuote = true;
                        continue;
                    }
                }
            }

            // "Am ... schrieb ..." pattern
            if (/^Am\s+.+\s+schrieb\s+.+:/i.test(trimmedLine) ||
                /^On\s+.+\s+wrote:/i.test(trimmedLine)) {
                foundQuoteStart = true;
                inQuote = true;
                continue;
            }
        }

        // Skip quoted lines
        if (inQuote) {
            continue;
        }

        // Skip lines starting with ">"
        if (/^>/.test(trimmedLine)) {
            continue;
        }

        result.push(line);
    }

    return result.join('\n').trim();
}

/**
 * Convert HTML to plain text
 */
function htmlToPlainText(html) {
    if (!html) return '';

    var text = html;

    // Remove script and style tags with content
    text = text.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '');
    text = text.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '');

    // Convert common block elements to newlines
    text = text.replace(/<\/(p|div|tr|li|h[1-6])>/gi, '\n');
    text = text.replace(/<br\s*\/?>/gi, '\n');
    text = text.replace(/<\/td>/gi, '\t');

    // Remove all remaining tags
    text = text.replace(/<[^>]+>/g, '');

    // Decode HTML entities
    text = text.replace(/&nbsp;/gi, ' ');
    text = text.replace(/&amp;/gi, '&');
    text = text.replace(/&lt;/gi, '<');
    text = text.replace(/&gt;/gi, '>');
    text = text.replace(/&quot;/gi, '"');
    text = text.replace(/&#39;/gi, "'");
    text = text.replace(/&auml;/gi, 'ä');
    text = text.replace(/&ouml;/gi, 'ö');
    text = text.replace(/&uuml;/gi, 'ü');
    text = text.replace(/&Auml;/gi, 'Ä');
    text = text.replace(/&Ouml;/gi, 'Ö');
    text = text.replace(/&Uuml;/gi, 'Ü');
    text = text.replace(/&szlig;/gi, 'ß');

    // Clean up whitespace
    text = text.replace(/[ \t]+/g, ' ');
    text = text.replace(/\n\s*\n\s*\n/g, '\n\n');

    return text.trim();
}
