/**
 * Tabs Module
 * Tab switching functionality for Export/Import views
 */

/**
 * Switch between Export and Import tabs
 * @param {string} tab - 'export' or 'import'
 */
function switchTab(tab) {
    var tabs = document.getElementsByClassName('tab');
    for (var i = 0; i < tabs.length; i++) {
        tabs[i].className = 'tab' + (i === (tab === 'export' ? 0 : 1) ? ' active' : '');
    }
    document.getElementById('tabExport').className = 'tab-content' + (tab === 'export' ? ' active' : '');
    document.getElementById('tabImport').className = 'tab-content' + (tab === 'import' ? ' active' : '');

    // When switching to import tab, advance to step 5 if step 4 is completed
    if (tab === 'import' && completedSteps[4]) {
        currentStep = 5;
        updateStepIndicator(0, false);
    }

    // If switching to import and Excel was already connected during export, reuse it
    if (tab === 'import' && exportWorkbook && !targetWorkbook) {
        try {
            // Test if Excel connection is still valid
            var testName = exportWorkbook.Name;
            excelApp = exportExcelApp;
            targetWorkbook = exportWorkbook;
            document.getElementById('excelStatus').value = "OK: " + testName + " (vom Export übernommen)";
            checkImportReady();
        } catch (e) {
            // Excel connection lost - reset
            exportWorkbook = null;
            exportExcelApp = null;
            document.getElementById('excelStatus').value = "Excel-Verbindung verloren - bitte neu verbinden";
        }
    }

    // Auto-load today's JSON file when switching to import tab
    if (tab === 'import' && !jsonData) {
        autoLoadTodayJson();
    }
}

/**
 * Auto-load today's bearbeitet JSON file
 */
function autoLoadTodayJson() {
    try {
        var shell = new ActiveXObject("WScript.Shell");
        var shellApp = new ActiveXObject("Shell.Application");
        var fso = new ActiveXObject("Scripting.FileSystemObject");

        var downloads;
        try { downloads = shellApp.NameSpace("shell:Downloads").Self.Path; }
        catch (e) { downloads = shell.ExpandEnvironmentStrings("%USERPROFILE%") + "\\Downloads"; }

        // Build today's filename
        var today = new Date();
        var dateStr = today.getFullYear() + '-' +
            (today.getMonth() + 1 < 10 ? '0' : '') + (today.getMonth() + 1) + '-' +
            (today.getDate() < 10 ? '0' : '') + today.getDate();
        var expectedFile = downloads + '\\hypercare_bearbeitet_' + dateStr + '.json';

        if (fso.FileExists(expectedFile)) {
            loadJsonFile(expectedFile);
        }
    } catch (e) {
        // Silently fail - user can manually browse
    }
}

