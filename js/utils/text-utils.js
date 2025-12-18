/**
 * Text Utilities
 * Text sanitization, encoding fixes, and email quote removal
 */

/**
 * Sanitize text for safe display (remove null bytes, normalize whitespace)
 * IMPORTANT: Always removes \r to prevent _x000D_ in Excel
 */
function sanitizeText(text) {
    if (!text) return '';
    return text
        .replace(/\x00/g, '')  // Remove null bytes
        .replace(/\r\n/g, '\n')  // Normalize line endings (CRLF -> LF)
        .replace(/\r/g, '\n')   // Remove stray CR
        .replace(/_x000D_/g, '') // Remove Excel-escaped CR if present
        .trim();
}

/**
 * Fix encoding issues (common in German emails)
 * Handles UTF-8 double-encoding using Unicode escapes for HTA compatibility
 */
function fixEncoding(text) {
    if (!text) return '';

    var result = text;

    // UTF-8 double-encoded German umlauts (Latin-1 interpretation)
    // Pattern: UTF-8 bytes interpreted as Latin-1, then re-encoded
    result = result.replace(/\u00C3\u00A4/g, '\u00E4');  // Ã¤ -> ä
    result = result.replace(/\u00C3\u00B6/g, '\u00F6');  // Ã¶ -> ö
    result = result.replace(/\u00C3\u00BC/g, '\u00FC');  // Ã¼ -> ü
    result = result.replace(/\u00C3\u0084/g, '\u00C4');  // Ã„ -> Ä
    result = result.replace(/\u00C3\u0096/g, '\u00D6');  // Ã– -> Ö
    result = result.replace(/\u00C3\u009C/g, '\u00DC');  // Ãœ -> Ü (Latin-1)
    result = result.replace(/\u00C3\u009F/g, '\u00DF');  // ÃŸ -> ß (Latin-1)

    // UTF-8 double-encoded German umlauts (Windows-1252 interpretation)
    // When bytes 0x80-0x9F are read as Windows-1252 instead of Latin-1
    result = result.replace(/\u00C3\u201E/g, '\u00C4');  // Ã„ -> Ä (Win-1252: 84 = „)
    result = result.replace(/\u00C3\u2013/g, '\u00D6');  // Ã– -> Ö (Win-1252: 96 = –)
    result = result.replace(/\u00C3\u0153/g, '\u00DC');  // Ãœ -> Ü (Win-1252: 9C = œ)
    result = result.replace(/\u00C3\u0178/g, '\u00DF');  // ÃŸ -> ß (Win-1252: 9F = Ÿ)

    // Double-encoded special characters
    result = result.replace(/\u00C3\u00A9/g, '\u00E9');  // Ã© -> é
    result = result.replace(/\u00C3\u00A8/g, '\u00E8');  // Ã¨ -> è
    result = result.replace(/\u00C3\u00A0/g, '\u00E0');  // Ã  -> à
    result = result.replace(/\u00C3\u00A2/g, '\u00E2');  // Ã¢ -> â
    result = result.replace(/\u00C3\u00AE/g, '\u00EE');  // Ã® -> î
    result = result.replace(/\u00C3\u00B4/g, '\u00F4');  // Ã´ -> ô
    result = result.replace(/\u00C3\u00BB/g, '\u00FB');  // Ã» -> û
    result = result.replace(/\u00C3\u00A7/g, '\u00E7');  // Ã§ -> ç

    // Triple-byte UTF-8 sequences double-encoded (arrows, quotes, etc.)
    // → (U+2192) = E2 86 92 in UTF-8 -> â†' when double-encoded
    result = result.replace(/\u00E2\u0086\u0092/g, '\u2192');  // â†' -> →
    result = result.replace(/\u00E2\u0086\u0090/g, '\u2190');  // â†? -> ←

    // German quotes double-encoded (Latin-1 interpretation)
    result = result.replace(/\u00E2\u0080\u009E/g, '\u201E');  // â€ž -> „
    result = result.replace(/\u00E2\u0080\u009C/g, '\u201C');  // â€œ -> "
    result = result.replace(/\u00E2\u0080\u009D/g, '\u201D');  // â€? -> "
    result = result.replace(/\u00E2\u0080\u0098/g, '\u2018');  // â€˜ -> '
    result = result.replace(/\u00E2\u0080\u0099/g, '\u2019');  // â€™ -> '

    // German quotes double-encoded (Windows-1252 interpretation)
    // In Win-1252: 0x80=€, 0x98=˜, 0x99=™, 0x9C=œ, 0x9E=ž
    result = result.replace(/\u00E2\u20AC\u02DC/g, '\u2018');  // â€˜ -> ' (Win-1252)
    result = result.replace(/\u00E2\u20AC\u2122/g, '\u2019');  // â€™ -> ' (Win-1252)
    result = result.replace(/\u00E2\u20AC\u0153/g, '\u201C');  // â€œ -> " (Win-1252)
    result = result.replace(/\u00E2\u20AC\u017E/g, '\u201E');  // â€ž -> „ (Win-1252)

    // Euro and other symbols
    result = result.replace(/\u00E2\u0082\u00AC/g, '\u20AC');  // â‚¬ -> €
    result = result.replace(/\u00E2\u0080\u0093/g, '\u2013');  // â€" -> –
    result = result.replace(/\u00E2\u0080\u0094/g, '\u2014');  // â€" -> —
    result = result.replace(/\u00E2\u0080\u00A6/g, '\u2026');  // â€¦ -> …

    // Stray  characters from double-encoding
    result = result.replace(/\u00C2\u00A0/g, ' ');  // Â  -> space (NBSP)
    result = result.replace(/\u00C2\u00B4/g, "'");  // Â´ -> '
    result = result.replace(/\u00C2\u00B0/g, '\u00B0');  // Â° -> °
    result = result.replace(/\u00C2\u00A7/g, '\u00A7');  // Â§ -> §
    result = result.replace(/\u00C2\u00A9/g, '\u00A9');  // Â© -> ©
    result = result.replace(/\u00C2\u00AE/g, '\u00AE');  // Â® -> ®

    // Remove invisible/problematic Unicode characters
    result = result.replace(/[\u200B-\u200D\uFEFF]/g, '');  // Zero-width chars
    result = result.replace(/\u00A0/g, ' ');  // Non-breaking space to regular space

    return result;
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
 * CRITICAL: This function MUST remove \r to prevent _x000D_ appearing in Excel cells
 */
function sanitizeForExcel(text) {
    if (!text) return '';

    // FIRST: Remove carriage returns to prevent _x000D_ in Excel
    var result = text.replace(/\r\n/g, '\n').replace(/\r/g, '');

    // Remove any existing _x000D_ sequences (already escaped)
    result = result.replace(/_x000D_/g, '');

    // Remove control characters except newlines and tabs
    result = result.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');

    // Remove emojis and other 4-byte UTF-8 characters (surrogate pairs)
    // This prevents issues in HTA/JScript and Excel
    result = result.replace(/[\uD800-\uDBFF][\uDC00-\uDFFF]/g, '');

    // Remove other problematic Unicode symbols
    result = result.replace(/[\u2600-\u26FF]/g, '');   // Misc symbols (sun, stars, etc.)
    result = result.replace(/[\u2700-\u27BF]/g, '');   // Dingbats
    result = result.replace(/[\u2300-\u23FF]/g, '');   // Misc technical symbols
    result = result.replace(/[\u2B50-\u2B55]/g, '');   // Stars, circles
    result = result.replace(/[\u203C\u2049]/g, '');    // Exclamation marks
    result = result.replace(/[\u00AE\u00A9\u2122]/g, ''); // ®, ©, ™

    // Replace common problematic characters
    result = result.replace(/\u2018|\u2019/g, "'");  // Smart quotes
    result = result.replace(/\u201C|\u201D/g, '"');  // Smart double quotes
    result = result.replace(/\u2013/g, '-');         // En dash
    result = result.replace(/\u2014/g, '--');        // Em dash
    result = result.replace(/\u2026/g, '...');       // Ellipsis
    result = result.replace(/\u00B7/g, '-');         // Middle dot
    result = result.replace(/\u2022/g, '-');         // Bullet point
    result = result.replace(/\u25CF/g, '-');         // Black circle
    result = result.replace(/\u25CB/g, 'o');         // White circle
    result = result.replace(/\u2713|\u2714/g, '[x]'); // Checkmarks
    result = result.replace(/\u2717|\u2718/g, '[ ]'); // X marks
    result = result.replace(/\u00AB/g, '<<');        // «
    result = result.replace(/\u00BB/g, '>>');        // »

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
