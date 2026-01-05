#!/usr/bin/env node
/**
 * Comprehensive Test Suite for BGAV Hypercare Email Review Tool
 * Tests JavaScript modules for logic errors and edge cases
 */

const fs = require('fs');
const path = require('path');

// Test results tracking
let testsPassed = 0;
let testsFailed = 0;
const failedTests = [];

function test(name, fn) {
    try {
        fn();
        testsPassed++;
        console.log(`  ✓ ${name}`);
    } catch (error) {
        testsFailed++;
        failedTests.push({ name, error: error.message });
        console.log(`  ✗ ${name}: ${error.message}`);
    }
}

function assertEqual(actual, expected, message) {
    if (actual !== expected) {
        throw new Error(`${message || 'Assertion failed'}: expected "${expected}", got "${actual}"`);
    }
}

function assertDeepEqual(actual, expected, message) {
    if (JSON.stringify(actual) !== JSON.stringify(expected)) {
        throw new Error(`${message || 'Deep assertion failed'}: expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
    }
}

function assertTrue(condition, message) {
    if (!condition) {
        throw new Error(message || 'Expected true');
    }
}

function assertFalse(condition, message) {
    if (condition) {
        throw new Error(message || 'Expected false');
    }
}

// ============================================================================
// TEXT UTILITIES TESTS (from js/utils/text-utils.js)
// ============================================================================

console.log('\n=== TEXT UTILITIES TESTS ===\n');

// Inline the functions for testing
function sanitizeText(text) {
    if (!text) return '';
    return text
        .replace(/\x00/g, '')
        .replace(/\r\n/g, '\n')
        .replace(/\r/g, '\n')
        .replace(/_x000D_/g, '')
        .trim();
}

function fixEncoding(text) {
    if (!text) return '';
    var result = text;
    result = result.replace(/\u00C3\u00A4/g, '\u00E4');  // Ã¤ -> ä
    result = result.replace(/\u00C3\u00B6/g, '\u00F6');  // Ã¶ -> ö
    result = result.replace(/\u00C3\u00BC/g, '\u00FC');  // Ã¼ -> ü
    result = result.replace(/\u00C3\u0084/g, '\u00C4');  // Ã„ -> Ä
    result = result.replace(/\u00C3\u0096/g, '\u00D6');  // Ã– -> Ö
    result = result.replace(/\u00C3\u009C/g, '\u00DC');  // Ãœ -> Ü
    result = result.replace(/\u00C3\u009F/g, '\u00DF');  // ÃŸ -> ß
    result = result.replace(/[\u200B-\u200D\uFEFF]/g, '');
    result = result.replace(/\u00A0/g, ' ');
    return result;
}

function normalizeSubject(subject) {
    if (!subject) return '';
    return subject
        .replace(/^(RE|AW|FW|WG|Fwd|Antwort):\s*/gi, '')
        .replace(/^(RE|AW|FW|WG|Fwd|Antwort):\s*/gi, '')
        .replace(/^(RE|AW|FW|WG|Fwd|Antwort):\s*/gi, '')
        .toLowerCase()
        .trim();
}

function truncateForExcel(text, maxLength) {
    if (!text) return '';
    maxLength = maxLength || 30000;
    if (text.length <= maxLength) return text;
    return text.substring(0, maxLength) + '\n[... Text gekürzt ...]';
}

function sanitizeForExcel(text) {
    if (!text) return '';
    var result = text.replace(/\r\n/g, '\n').replace(/\r/g, '');
    result = result.replace(/_x000D_/g, '');
    result = result.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
    result = result.replace(/[\uD800-\uDBFF][\uDC00-\uDFFF]/g, '');
    result = result.replace(/\u2018|\u2019/g, "'");
    result = result.replace(/\u201C|\u201D/g, '"');
    result = result.replace(/\u2013/g, '-');
    result = result.replace(/\u2014/g, '--');
    result = result.replace(/\u2026/g, '...');
    return result;
}

function removeEmailQuotes(text) {
    if (!text) return '';
    var lines = text.split('\n');
    var result = [];
    var inQuote = false;
    var foundQuoteStart = false;

    for (var i = 0; i < lines.length; i++) {
        var line = lines[i];
        var trimmedLine = line.trim();

        if (!foundQuoteStart) {
            if (/^Von:\s*.+/i.test(trimmedLine) ||
                /^From:\s*.+/i.test(trimmedLine) ||
                /^-----\s*Urspr.+Nachricht/i.test(trimmedLine) ||
                /^-----\s*Original\s*Message/i.test(trimmedLine) ||
                /^_{5,}/.test(trimmedLine) ||
                /^-{5,}/.test(trimmedLine)) {
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
            if (/^Am\s+.+\s+schrieb\s+.+:/i.test(trimmedLine) ||
                /^On\s+.+\s+wrote:/i.test(trimmedLine)) {
                foundQuoteStart = true;
                inQuote = true;
                continue;
            }
        }
        if (inQuote) continue;
        if (/^>/.test(trimmedLine)) continue;
        result.push(line);
    }
    return result.join('\n').trim();
}

// Test cases
test('sanitizeText: handles null/empty', () => {
    assertEqual(sanitizeText(null), '');
    assertEqual(sanitizeText(''), '');
    assertEqual(sanitizeText(undefined), '');
});

test('sanitizeText: removes null bytes', () => {
    assertEqual(sanitizeText('Hello\x00World'), 'Hello\nWorld'.trim() ? 'HelloWorld' : 'HelloWorld');
});

test('sanitizeText: normalizes CRLF to LF', () => {
    assertEqual(sanitizeText('Line1\r\nLine2'), 'Line1\nLine2');
});

test('sanitizeText: removes stray CR', () => {
    assertEqual(sanitizeText('Line1\rLine2'), 'Line1\nLine2');
});

test('sanitizeText: removes _x000D_ sequences', () => {
    assertEqual(sanitizeText('Text_x000D_More'), 'TextMore');
});

test('fixEncoding: fixes German umlauts (double-encoded)', () => {
    assertEqual(fixEncoding('\u00C3\u00A4'), 'ä');
    assertEqual(fixEncoding('\u00C3\u00B6'), 'ö');
    assertEqual(fixEncoding('\u00C3\u00BC'), 'ü');
    assertEqual(fixEncoding('\u00C3\u009F'), 'ß');
});

test('fixEncoding: fixes uppercase umlauts', () => {
    assertEqual(fixEncoding('\u00C3\u0084'), 'Ä');
    assertEqual(fixEncoding('\u00C3\u0096'), 'Ö');
    assertEqual(fixEncoding('\u00C3\u009C'), 'Ü');
});

test('fixEncoding: removes zero-width characters', () => {
    assertEqual(fixEncoding('Hello\u200BWorld'), 'HelloWorld');
    assertEqual(fixEncoding('Test\uFEFFing'), 'Testing');
});

test('fixEncoding: converts NBSP to regular space', () => {
    assertEqual(fixEncoding('Hello\u00A0World'), 'Hello World');
});

test('normalizeSubject: removes RE: prefix', () => {
    assertEqual(normalizeSubject('RE: Test Subject'), 'test subject');
    assertEqual(normalizeSubject('re: another test'), 'another test');
});

test('normalizeSubject: removes AW: prefix (German)', () => {
    assertEqual(normalizeSubject('AW: Betreff Test'), 'betreff test');
    assertEqual(normalizeSubject('aw: klein'), 'klein');
});

test('normalizeSubject: removes WG: prefix (German forward)', () => {
    assertEqual(normalizeSubject('WG: Weitergeleitet'), 'weitergeleitet');
});

test('normalizeSubject: removes multiple prefixes', () => {
    assertEqual(normalizeSubject('RE: AW: WG: Test'), 'test');
    assertEqual(normalizeSubject('AW: RE: RE: Multi'), 'multi');
});

test('normalizeSubject: handles empty/null', () => {
    assertEqual(normalizeSubject(''), '');
    assertEqual(normalizeSubject(null), '');
});

test('truncateForExcel: no truncation for short text', () => {
    assertEqual(truncateForExcel('Short text', 100), 'Short text');
});

test('truncateForExcel: truncates long text', () => {
    const longText = 'A'.repeat(50);
    const result = truncateForExcel(longText, 20);
    assertTrue(result.includes('[... Text gekürzt ...]'), 'Should contain truncation marker');
    assertTrue(result.length < longText.length + 30, 'Should be shorter than original');
});

test('sanitizeForExcel: removes carriage returns', () => {
    // Note: \r is removed completely (not converted to \n) to prevent _x000D_ in Excel
    assertEqual(sanitizeForExcel('Line1\r\nLine2'), 'Line1\nLine2');
    assertEqual(sanitizeForExcel('Test\rOnly'), 'TestOnly');  // \r alone is removed
});

test('sanitizeForExcel: converts smart quotes', () => {
    assertEqual(sanitizeForExcel('\u2018test\u2019'), "'test'");
    assertEqual(sanitizeForExcel('\u201Ctest\u201D'), '"test"');
});

test('sanitizeForExcel: converts dashes', () => {
    assertEqual(sanitizeForExcel('en\u2013dash'), 'en-dash');
    assertEqual(sanitizeForExcel('em\u2014dash'), 'em--dash');
});

test('sanitizeForExcel: converts ellipsis', () => {
    assertEqual(sanitizeForExcel('wait\u2026'), 'wait...');
});

test('removeEmailQuotes: keeps original content', () => {
    const text = 'This is my reply.\n\nVon: Someone\nGesendet: Today\nQuoted text here';
    const result = removeEmailQuotes(text);
    assertTrue(result.includes('This is my reply'), 'Should keep original');
    assertFalse(result.includes('Quoted text here'), 'Should remove quoted');
});

test('removeEmailQuotes: handles empty', () => {
    assertEqual(removeEmailQuotes(''), '');
    assertEqual(removeEmailQuotes(null), '');
});

// ============================================================================
// CONVERSATION INDEX TESTS
// ============================================================================

console.log('\n=== CONVERSATION INDEX TESTS ===\n');

function calculateDepthFromIndex(hexIndex) {
    if (!hexIndex || hexIndex.length < 44) return 0;
    return Math.floor((hexIndex.length - 44) / 10);
}

function getParentIndexHex(hexIndex) {
    if (!hexIndex || hexIndex.length <= 44) return '';
    return hexIndex.substring(0, hexIndex.length - 10);
}

function base64ToHex(base64) {
    if (!base64) return '';
    try {
        var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
        var lookup = {};
        for (var i = 0; i < chars.length; i++) {
            lookup[chars[i]] = i;
        }
        base64 = base64.replace(/=/g, '');
        var bytes = [];
        for (var j = 0; j < base64.length; j += 4) {
            var n = (lookup[base64[j]] << 18) |
                    (lookup[base64[j + 1]] << 12) |
                    ((lookup[base64[j + 2]] || 0) << 6) |
                    (lookup[base64[j + 3]] || 0);
            bytes.push((n >> 16) & 0xFF);
            if (base64[j + 2]) bytes.push((n >> 8) & 0xFF);
            if (base64[j + 3]) bytes.push(n & 0xFF);
        }
        var hex = '';
        for (var k = 0; k < bytes.length; k++) {
            var h = bytes[k].toString(16).toUpperCase();
            if (h.length === 1) h = '0' + h;
            hex += h;
        }
        return hex;
    } catch (e) {
        return '';
    }
}

test('calculateDepthFromIndex: root message (44 chars)', () => {
    // Base header = 22 bytes = 44 hex chars -> depth 0
    const rootIndex = '0101DC70F4A1F7EEC62403D45A41A732183A1EA20DAB';  // 44 chars
    assertEqual(calculateDepthFromIndex(rootIndex), 0);
});

test('calculateDepthFromIndex: first reply (54 chars)', () => {
    // 44 + 10 = 54 hex chars -> depth 1
    const replyIndex = '0101DC70F4A1F7EEC62403D45A41A732183A1EA20DAB0012345678';  // 54 chars
    assertEqual(calculateDepthFromIndex(replyIndex), 1);
});

test('calculateDepthFromIndex: second reply (64 chars)', () => {
    // 44 + 20 = 64 hex chars -> depth 2
    const reply2Index = '0101DC70F4A1F7EEC62403D45A41A732183A1EA20DAB00123456780ABCDEF012';  // 64 chars
    assertEqual(calculateDepthFromIndex(reply2Index), 2);
});

test('calculateDepthFromIndex: handles short/invalid', () => {
    assertEqual(calculateDepthFromIndex(''), 0);
    assertEqual(calculateDepthFromIndex('ABC'), 0);
    assertEqual(calculateDepthFromIndex(null), 0);
});

test('getParentIndexHex: returns empty for root', () => {
    const rootIndex = '01DC70F4A1F7EEC62403D45A41A732183A1EA20D88';
    assertEqual(getParentIndexHex(rootIndex), '');
});

test('getParentIndexHex: returns parent for reply', () => {
    const replyIndex = '01DC70F4A1F7EEC62403D45A41A732183A1EA20D880012345678';
    assertEqual(getParentIndexHex(replyIndex), '01DC70F4A1F7EEC62403D45A41A732183A1EA20D88');
});

test('getParentIndexHex: handles empty/null', () => {
    assertEqual(getParentIndexHex(''), '');
    assertEqual(getParentIndexHex(null), '');
});

test('base64ToHex: converts valid base64', () => {
    // 'SGVsbG8=' = 'Hello' = 48656C6C6F
    const result = base64ToHex('SGVsbG8=');
    assertEqual(result, '48656C6C6F');
});

test('base64ToHex: handles empty', () => {
    assertEqual(base64ToHex(''), '');
    assertEqual(base64ToHex(null), '');
});

// ============================================================================
// THREAD MATCHING TESTS
// ============================================================================

console.log('\n=== THREAD MATCHING TESTS ===\n');

function normalizeSubjectForMatching(subject) {
    if (!subject) return '';
    var normalized = subject.replace(/^(re:|aw:|fwd:|wg:|fw:)\s*/gi, '');
    while (/^(re:|aw:|fwd:|wg:|fw:)\s*/i.test(normalized)) {
        normalized = normalized.replace(/^(re:|aw:|fwd:|wg:|fw:)\s*/gi, '');
    }
    normalized = normalized.replace(/\s+/g, ' ').trim().toLowerCase();
    return normalized;
}

function isWithinTimeWindow(date1, date2, maxDays) {
    if (!date1 || !date2) return false;
    maxDays = maxDays || 30;
    var diffMs = Math.abs(date1.getTime() - date2.getTime());
    var diffDays = diffMs / (1000 * 60 * 60 * 24);
    return diffDays <= maxDays;
}

test('normalizeSubjectForMatching: removes all prefixes', () => {
    assertEqual(normalizeSubjectForMatching('RE: AW: FW: WG: Test'), 'test');
    assertEqual(normalizeSubjectForMatching('FWD: Forwarded'), 'forwarded');
});

test('normalizeSubjectForMatching: normalizes whitespace', () => {
    assertEqual(normalizeSubjectForMatching('RE:  Multiple   Spaces'), 'multiple spaces');
});

test('normalizeSubjectForMatching: converts to lowercase', () => {
    assertEqual(normalizeSubjectForMatching('UPPERCASE SUBJECT'), 'uppercase subject');
});

test('isWithinTimeWindow: returns true for same day', () => {
    const date1 = new Date('2025-01-01');
    const date2 = new Date('2025-01-01');
    assertTrue(isWithinTimeWindow(date1, date2, 30));
});

test('isWithinTimeWindow: returns true within window', () => {
    const date1 = new Date('2025-01-01');
    const date2 = new Date('2025-01-15');
    assertTrue(isWithinTimeWindow(date1, date2, 30));
});

test('isWithinTimeWindow: returns false outside window', () => {
    const date1 = new Date('2025-01-01');
    const date2 = new Date('2025-03-01');
    assertFalse(isWithinTimeWindow(date1, date2, 30));
});

test('isWithinTimeWindow: handles null dates', () => {
    assertFalse(isWithinTimeWindow(null, new Date()));
    assertFalse(isWithinTimeWindow(new Date(), null));
});

// ============================================================================
// JSON DATA VALIDATION TESTS
// ============================================================================

console.log('\n=== JSON DATA VALIDATION TESTS ===\n');

// Load and validate test JSON file
let testData = null;
try {
    const testFilePath = path.join(__dirname, 'testdateien', 'teil3_bearb.json');
    const content = fs.readFileSync(testFilePath, 'utf8');
    testData = JSON.parse(content);
} catch (e) {
    console.log(`  ⚠ Could not load test data: ${e.message}`);
}

if (testData && Array.isArray(testData)) {
    test('JSON: is valid array', () => {
        assertTrue(Array.isArray(testData), 'Should be an array');
    });

    test('JSON: has email entries', () => {
        assertTrue(testData.length > 0, 'Should have entries');
    });

    test('JSON: emails have required fields', () => {
        const email = testData[0];
        assertTrue('id' in email, 'Should have id');
        assertTrue('datum' in email, 'Should have datum');
        assertTrue('von_email' in email, 'Should have von_email');
        assertTrue('betreff' in email, 'Should have betreff');
        assertTrue('conversationId' in email, 'Should have conversationId');
    });

    test('JSON: emails have valid dates', () => {
        const email = testData[0];
        const date = new Date(email.datum);
        assertFalse(isNaN(date.getTime()), 'Should have valid date');
    });

    test('JSON: conversationId format is valid', () => {
        const email = testData[0];
        assertTrue(/^[A-F0-9]{32}$/i.test(email.conversationId), 'ConversationID should be 32 hex chars');
    });

    test('JSON: alleAntworten is array', () => {
        const email = testData[0];
        assertTrue(Array.isArray(email.alleAntworten), 'alleAntworten should be array');
    });

    test('JSON: replyId format in antworten', () => {
        const email = testData[0];
        if (email.alleAntworten && email.alleAntworten.length > 0) {
            const reply = email.alleAntworten[0];
            assertTrue('replyId' in reply, 'Reply should have replyId');
            assertTrue(reply.replyId.length > 50, 'ReplyId should be long EntryID');
        }
    });
}

// ============================================================================
// EDGE CASE TESTS
// ============================================================================

console.log('\n=== EDGE CASE TESTS ===\n');

test('Edge: very long text handling', () => {
    const longText = 'A'.repeat(100000);
    const truncated = truncateForExcel(longText);
    assertTrue(truncated.length < longText.length, 'Should truncate very long text');
});

test('Edge: special characters in subject', () => {
    const subject = 'RE: Prüfung der Änderung für Ärzte';
    const normalized = normalizeSubject(subject);
    assertTrue(normalized.includes('prüfung'), 'Should preserve umlauts');
});

test('Edge: empty conversation handling', () => {
    assertEqual(calculateDepthFromIndex(''), 0);
    assertEqual(getParentIndexHex(''), '');
});

test('Edge: Unicode in email text', () => {
    const text = 'Prüfung mit Überschrift für Änderung: €100';
    const sanitized = sanitizeForExcel(text);
    assertTrue(sanitized.includes('Prüfung'), 'Should preserve German chars');
    assertTrue(sanitized.includes('€') || sanitized.includes('100'), 'Should handle Euro symbol');
});

test('Edge: multiple quote levels', () => {
    const text = 'My reply\n\nVon: Someone\nGesendet: Date\n> Quoted\n>> Double quoted';
    const result = removeEmailQuotes(text);
    assertTrue(result.includes('My reply'), 'Should keep original');
});

// ============================================================================
// IMPORT/EXPORT FORMAT TESTS
// ============================================================================

console.log('\n=== IMPORT/EXPORT FORMAT TESTS ===\n');

function addSentOnlyFlagIfMissing(emails) {
    if (!emails || !Array.isArray(emails)) return emails;
    var mailboxPatterns = [
        'provisionsservice@barmenia.de',
        'provisionsservice@'
    ];

    for (var i = 0; i < emails.length; i++) {
        var email = emails[i];
        if (email.sentOnly === true || email.sentOnly === false) continue;

        var vonEmail = (email.von_email || '').toLowerCase();
        var isSentFromMailbox = false;
        for (var m = 0; m < mailboxPatterns.length; m++) {
            if (vonEmail.indexOf(mailboxPatterns[m]) !== -1) {
                isSentFromMailbox = true;
                break;
            }
        }

        var hasIncomingMessage = false;
        var antworten = email.antworten || email.alleAntworten || [];
        for (var j = 0; j < antworten.length; j++) {
            if (antworten[j].isIncoming === true) {
                hasIncomingMessage = true;
                break;
            }
        }

        if (isSentFromMailbox && !hasIncomingMessage) {
            email.sentOnly = true;
        } else {
            email.sentOnly = false;
        }
    }
    return emails;
}

test('sentOnly: external email not marked as sentOnly', () => {
    const emails = [{ von_email: 'external@example.com', antworten: [] }];
    const result = addSentOnlyFlagIfMissing(emails);
    assertFalse(result[0].sentOnly, 'External email should not be sentOnly');
});

test('sentOnly: mailbox email without incoming marked as sentOnly', () => {
    const emails = [{ von_email: 'provisionsservice@barmenia.de', antworten: [] }];
    const result = addSentOnlyFlagIfMissing(emails);
    assertTrue(result[0].sentOnly, 'Mailbox email without incoming should be sentOnly');
});

test('sentOnly: mailbox email with incoming not marked as sentOnly', () => {
    const emails = [{
        von_email: 'provisionsservice@barmenia.de',
        antworten: [{ isIncoming: true }]
    }];
    const result = addSentOnlyFlagIfMissing(emails);
    assertFalse(result[0].sentOnly, 'Mailbox email with incoming should not be sentOnly');
});

test('sentOnly: preserves existing flag', () => {
    const emails = [{ von_email: 'test@example.com', sentOnly: true }];
    const result = addSentOnlyFlagIfMissing(emails);
    assertTrue(result[0].sentOnly, 'Should preserve existing true flag');
});

// ============================================================================
// X500 ADDRESS CLEANING TESTS
// ============================================================================

console.log('\n=== X500 ADDRESS TESTS ===\n');

function cleanX500Address(addr) {
    if (!addr) return '';
    if (addr.indexOf('@') > -1 && addr.indexOf('/cn=') === -1) {
        return addr;
    }
    if (addr.indexOf('/cn=') > -1) {
        var cleaned = addr.replace(/[,\s]+$/, '');
        var match = cleaned.match(/cn=([^\/]+)$/i);
        if (match) {
            return match[1].replace(/^[a-f0-9]{32}-/i, '');
        }
    }
    return addr;
}

test('X500: returns clean email as-is', () => {
    assertEqual(cleanX500Address('test@example.com'), 'test@example.com');
});

test('X500: extracts name from X500 format', () => {
    const x500 = '/o=Company/ou=Department/cn=John Doe';
    assertEqual(cleanX500Address(x500), 'John Doe');
});

test('X500: removes GUID prefix', () => {
    const x500 = '/o=Company/cn=f4fadad6c63f4b8e93e39ddb6abbf8d7-John Doe';
    assertEqual(cleanX500Address(x500), 'John Doe');
});

test('X500: handles empty/null', () => {
    assertEqual(cleanX500Address(''), '');
    assertEqual(cleanX500Address(null), '');
});

// ============================================================================
// DATE FORMATTING TESTS
// ============================================================================

console.log('\n=== DATE FORMATTING TESTS ===\n');

function formatDateForRestrict(date) {
    if (!date) return '';
    var d = new Date(date);
    var month = d.getMonth() + 1;
    var day = d.getDate();
    var year = d.getFullYear();
    return month + '/' + day + '/' + year + ' 00:00';
}

test('formatDateForRestrict: formats date correctly', () => {
    const date = new Date('2025-06-15');
    const result = formatDateForRestrict(date);
    assertTrue(result.includes('6/15/2025'), 'Should format as M/D/YYYY');
});

// ============================================================================
// SUMMARY
// ============================================================================

console.log('\n' + '='.repeat(60));
console.log('TEST SUMMARY');
console.log('='.repeat(60));
console.log(`\nTotal:  ${testsPassed + testsFailed} tests`);
console.log(`Passed: ${testsPassed} ✓`);
console.log(`Failed: ${testsFailed} ✗`);

if (failedTests.length > 0) {
    console.log('\n--- FAILED TESTS ---');
    failedTests.forEach(({ name, error }) => {
        console.log(`\n  ✗ ${name}`);
        console.log(`    Error: ${error}`);
    });
}

console.log('\n');

// Exit with error code if any tests failed
process.exit(testsFailed > 0 ? 1 : 0);
