/**
 * Excel Connection Module
 * Connect to Excel for duplicate checking and import
 */

// Excel connection for export duplicate checking
var exportExcelApp = null;
var exportWorkbook = null;

// Note: existingEmailKeys, existingEmailIds, existingReplies, existingReplyIds,
// existingAnfragen, existingDates, existingSubjects, foundInInbox
// are declared in config.js

// Excel rows lookup for auto-erledigt feature
var existingExcelRows = {};

// Import state
var excelApp = null;
var targetWorkbook = null;
var jsonData = null;

/**
 * Detect and connect to Excel for export duplicate checking
 */
function detectExcelForExport() {
    document.getElementById('exportExcelStatus').value = "Suche Excel...";
    existingEmailKeys = {};
    existingReplies = {};
    existingEmailIds = {};
    existingReplyIds = {};
    existingAnfragen = {};
    existingDates = {};
    existingSubjects = {};
    existingExcelRows = {};
    foundInInbox = {};

    try {
        // Try to connect to running Excel
        try {
            exportExcelApp = GetObject("", "Excel.Application");
        } catch (e1) {
            exportExcelApp = new ActiveXObject("Excel.Application");
        }
        exportExcelApp.Visible = true;

        // Find Hypercare workbook
        exportWorkbook = null;
        for (var i = 1; i <= exportExcelApp.Workbooks.Count; i++) {
            try {
                var wb = exportExcelApp.Workbooks(i);
                var wbName = wb.Name.toLowerCase().replace(/[_\-\s]/g, '');
                if (wbName.indexOf('hypercare') !== -1) {
                    exportWorkbook = wb;
                    break;
                }
            } catch (e) {}
        }

        if (!exportWorkbook) {
            // Try SharePoint
            document.getElementById('exportExcelStatus').value = "Offne SharePoint-Datei...";
            var sharePointUrl = "https://barmeniagroup.sharepoint.com/sites/SAPUnity-WorkstreamICM/Freigegebene%20Dokumente/40%20Team%20Simulation,%20Kalibrierung,%20Verhandlung,%20Vertrag%20&%20EFM/Tracking_Hypercare.xlsm";
            try {
                exportWorkbook = exportExcelApp.Workbooks.Open(sharePointUrl);
            } catch (e) {
                document.getElementById('exportExcelStatus').value = "Excel nicht gefunden - Duplikate werden nicht gefiltert";
                return;
            }
        }

        // Read existing data from Excel
        document.getElementById('exportExcelStatus').value = "Lese bestehende Eintraege...";
        var worksheet;
        try {
            worksheet = exportWorkbook.Worksheets("Uebersicht");
        } catch (e) {
            worksheet = exportWorkbook.Worksheets(1);
        }

        var lastRow = worksheet.Cells(worksheet.Rows.Count, 1).End(-4162).Row;
        var count = 0;

        // Read data for duplicate detection
        for (var row = 2; row <= lastRow; row++) {
            var conversationId = worksheet.Cells(row, 22).Value || '';
            var datum = worksheet.Cells(row, 2).Value || '';
            var betreff = worksheet.Cells(row, 12).Value || '';
            var antworten = worksheet.Cells(row, 14).Value || '';
            var emailId = worksheet.Cells(row, 23).Value || '';
            var replyIds = worksheet.Cells(row, 24).Value || '';
            var anfrage = worksheet.Cells(row, 13).Value || '';
            var status = worksheet.Cells(row, 9).Value || '';

            // Primary key: ConversationID
            if (conversationId) {
                var convKey = 'conv|' + conversationId;
                existingEmailKeys[convKey] = true;
                if (antworten) existingReplies[convKey] = antworten;
                if (replyIds) existingReplyIds[convKey] = replyIds;
                if (anfrage) existingAnfragen[convKey] = anfrage;
                if (datum) existingDates[convKey] = datum;
                if (betreff) existingSubjects[convKey] = betreff;
                if (String(status).toLowerCase() !== 'erledigt') {
                    existingExcelRows[conversationId] = row;
                }
            }
            // Fallback key: date + subject
            if (datum || betreff) {
                var key = createExportEmailKey(datum, betreff);
                existingEmailKeys[key] = true;
                if (antworten) existingReplies[key] = antworten;
                if (replyIds) existingReplyIds[key] = replyIds;
                if (anfrage) existingAnfragen[key] = anfrage;
                if (datum) existingDates[key] = datum;
                if (betreff) existingSubjects[key] = betreff;
            }
            // Track emailId
            if (emailId) {
                existingEmailIds[emailId] = true;
            }
            // Track all replyIds individually for duplicate detection
            if (replyIds) {
                var replyIdList = String(replyIds).split(',');
                for (var ri = 0; ri < replyIdList.length; ri++) {
                    var rid = replyIdList[ri].trim();
                    if (rid) {
                        existingEmailIds[rid] = true;
                    }
                }
            }
            if (conversationId || datum || betreff) count++;
        }

        document.getElementById('exportExcelStatus').value = "OK: " + exportWorkbook.Name + " (" + count + " Eintraege)";
        checkExportReady();

        // Step 1 completed
        completeStep(1);

    } catch (e) {
        document.getElementById('exportExcelStatus').value = "Fehler: " + e.message;
        existingEmailKeys = {};
        existingEmailIds = {};
        existingReplyIds = {};
        existingAnfragen = {};
        existingDates = {};
        existingSubjects = {};
        existingExcelRows = {};
        foundInInbox = {};
        exportWorkbook = null;
        checkExportReady();
    }
}

/**
 * Create email key for export duplicate detection
 * @param {*} datum - Date value
 * @param {string} betreff - Subject
 * @returns {string} Key string
 */
function createExportEmailKey(datum, betreff) {
    var dateStr = '';
    if (datum) {
        try {
            var d;
            if (typeof datum === 'number') {
                d = new Date((datum - 25569) * 86400 * 1000);
            } else {
                d = new Date(datum);
            }
            if (!isNaN(d.getTime())) {
                dateStr = d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate();
            } else {
                dateStr = String(datum).substring(0, 10);
            }
        } catch (e) {
            dateStr = String(datum).substring(0, 10);
        }
    }
    var subjectStr = String(betreff || '').toLowerCase().trim();
    subjectStr = subjectStr.replace(/^(re|aw|fw|wg|fwd|antwort|antw):\s*/gi, '');
    subjectStr = subjectStr.replace(/^(re|aw|fw|wg|fwd|antwort|antw):\s*/gi, '');
    subjectStr = subjectStr.substring(0, 100);
    return dateStr + '|' + subjectStr;
}

/**
 * Detect Excel for import
 */
function detectExcel() {
    document.getElementById('excelStatus').value = "Suche Excel...";

    try {
        try {
            excelApp = GetObject("", "Excel.Application");
        } catch (e1) {
            excelApp = new ActiveXObject("Excel.Application");
        }
        excelApp.Visible = true;

        targetWorkbook = null;
        for (var i = 1; i <= excelApp.Workbooks.Count; i++) {
            try {
                var wb = excelApp.Workbooks(i);
                var wbName = wb.Name.toLowerCase().replace(/[_\-\s]/g, '');
                if (wbName.indexOf('hypercare') !== -1) {
                    targetWorkbook = wb;
                    break;
                }
            } catch (e) {}
        }

        if (targetWorkbook) {
            document.getElementById('excelStatus').value = "OK: " + targetWorkbook.Name;
            checkImportReady();
        } else {
            document.getElementById('excelStatus').value = "Hypercare-Datei nicht gefunden";
        }
    } catch (e) {
        document.getElementById('excelStatus').value = "Fehler: " + e.message;
    }
}

/**
 * Check if import can be started
 */
function checkImportReady() {
    var btn = document.getElementById('btnImport');
    if (targetWorkbook && jsonData) {
        btn.disabled = false;
    } else {
        btn.disabled = true;
    }
}

