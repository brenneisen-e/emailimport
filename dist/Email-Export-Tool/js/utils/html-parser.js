/**
 * HTML Parser
 * Extract new content from HTML emails (remove quoted text)
 */

/**
 * Extract only new content from HTML email body
 * Removes quoted replies and forwards
 * Returns: { newContent: string, hasQuotes: boolean }
 */
function extractNewContentFromHtml(htmlBody) {
    if (!htmlBody) return { newContent: '', hasQuotes: false };

    var result = { newContent: '', hasQuotes: false };

    // Find the position of quote separators in HTML
    // Comprehensive patterns for all major email clients
    var quotePatterns = [
        // === OUTLOOK-STYLE PATTERNS ===
        // Border-top separator (most common Outlook pattern)
        /<div[^>]*style\s*=\s*["'][^"']*border-top\s*:\s*solid\s*#E1E1E1[^"']*["'][^>]*>/i,
        /<div[^>]*style\s*=\s*["'][^"']*border-top\s*:\s*solid\s+#[A-Fa-f0-9]{6}[^"']*["'][^>]*>/i,
        /<div[^>]*style\s*=\s*["'][^"']*border:\s*none\s*;\s*border-top:\s*solid[^"']*["'][^>]*>/i,

        // Outlook header block with span
        /<p[^>]*class\s*=\s*["']?MsoNormal["']?[^>]*>\s*<b[^>]*>\s*<span[^>]*>\s*(Von|From)\s*:\s*<\/span>\s*<\/b>/i,

        // Simple header block start
        /<p[^>]*><b>Von:<\/b>/i,
        /<p[^>]*><b>From:<\/b>/i,
        /<b>Von:<\/b>/i,
        /<b>From:<\/b>/i,

        // === GMAIL-STYLE PATTERNS ===
        /<blockquote[^>]*class\s*=\s*["'][^"']*gmail_quote[^"']*["'][^>]*>/i,
        /<div[^>]*class\s*=\s*["'][^"']*gmail_quote[^"']*["'][^>]*>/i,
        /<div[^>]*class\s*=\s*["']gmail_extra["'][^>]*>/i,

        // === APPLE MAIL / THUNDERBIRD ===
        /<blockquote[^>]*type\s*=\s*["']cite["'][^>]*>/i,

        // === GENERIC PATTERNS ===
        // "Am ... schrieb" / "On ... wrote"
        /<div[^>]*>Am\s+[^<]+schrieb[^<]*:/i,
        /<div[^>]*>On\s+[^<]+wrote:/i,

        // Original message separator
        /-----\s*Urspr[^-]*Nachricht\s*-----/i,
        /-----\s*Original\s*Message\s*-----/i,

        // Horizontal rule as separator
        /<hr[^>]*>/i
    ];

    var earliestPosition = htmlBody.length;
    var foundQuote = false;

    for (var i = 0; i < quotePatterns.length; i++) {
        var match = quotePatterns[i].exec(htmlBody);
        if (match && match.index < earliestPosition) {
            earliestPosition = match.index;
            foundQuote = true;
        }
    }

    if (foundQuote) {
        result.hasQuotes = true;
        var newHtml = htmlBody.substring(0, earliestPosition);

        // Convert to plain text
        result.newContent = htmlToPlainText(newHtml);
    } else {
        // No quotes found - convert entire body
        result.newContent = htmlToPlainText(htmlBody);
    }

    // Clean up the result
    result.newContent = result.newContent
        .replace(/^\s+/, '')  // Trim start
        .replace(/\s+$/, '')  // Trim end
        .replace(/\n{3,}/g, '\n\n');  // Max 2 consecutive newlines

    return result;
}
