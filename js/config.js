/**
 * BGAV Hypercare Email Export - Configuration
 * Global constants and state objects
 */

// MAPI Property Tags for email headers
var PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";
var PR_IN_REPLY_TO_ID = "http://schemas.microsoft.com/mapi/proptag/0x1042001F";
var PR_INTERNET_REFERENCES = "http://schemas.microsoft.com/mapi/proptag/0x1039001F";
var PR_CONVERSATION_INDEX = "http://schemas.microsoft.com/mapi/proptag/0x00710102";

// Global Outlook objects
var outlookApp = null;
var outlookNS = null;
var selectedStore = null;

// Sent items cache and lookup maps
var sentItemsCache = [];
var sentItemsByConvId = {};
var sentItemsByInReplyTo = {};
var sentItemsByReference = {};
var sentItemsByNormSubject = {};

// Track conversations found in inbox (for auto-erledigt feature)
var foundInInbox = {};

// Export state object
var exportState = {
    folder: null,
    items: null,
    emails: [],
    total: 0,
    current: 1,
    oldCount: 0,
    cutoffDate: null,
    repliesFound: 0,
    duplicatesSkipped: 0,
    batchDuplicates: 0,
    duplicatesWithReplies: 0,
    processedConvIds: {}
};

// Sent cache state for async processing
var sentCacheState = {
    folders: [],
    folderIdx: 0,
    itemIdx: 1,
    maxItems: 0,
    items: null,
    cutoffDate: null,
    days: 7
};

// Import state
var importState = {
    jsonData: null,
    jsonPath: '',
    excel: null,
    workbook: null,
    worksheet: null,
    current: 0,
    total: 0,
    existingKeys: {},
    skipped: 0,
    updated: 0
};

// Excel data caches for export duplicate detection
var existingEmailKeys = {};
var existingEmailIds = {};
var existingReplies = {};
var existingReplyIds = {};
var existingAnfragen = {};
var existingDates = {};
var existingSubjects = {};
