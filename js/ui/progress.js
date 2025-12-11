/**
 * Progress Module
 * Progress bar and status display functions
 */

/**
 * Update export progress bar
 * @param {number} current - Current item being processed
 * @param {number} total - Total items to process
 * @param {string} phase - Current phase description
 */
function updateExportProgress(current, total, phase) {
    var pct = Math.round((current / total) * 100);
    document.getElementById('exportProgressFill').style.width = pct + '%';
    document.getElementById('exportProgressText').innerText = phase || ('Exportiere ' + current + '/' + total + '...');
}

/**
 * Update import progress bar
 * @param {number} current - Current item being processed
 * @param {number} total - Total items to process
 */
function updateImportProgress(current, total) {
    document.getElementById('importProgressFill').style.width = Math.round((current / total) * 100) + '%';
    document.getElementById('importProgressText').innerText = 'Importiere ' + current + '/' + total + '...';
}

/**
 * Show export progress container
 */
function showExportProgress() {
    document.getElementById('exportProgress').style.display = 'block';
    document.getElementById('exportSuccess').style.display = 'none';
    document.getElementById('exportError').style.display = 'none';
    document.getElementById('exportActions').style.display = 'none';
}

/**
 * Hide export progress container
 */
function hideExportProgress() {
    document.getElementById('exportProgress').style.display = 'none';
}

/**
 * Show export success message
 * @param {string} message - Success message to display
 */
function showExportSuccess(message) {
    hideExportProgress();
    document.getElementById('exportSuccess').style.display = 'block';
    document.getElementById('exportSuccessMsg').innerHTML = message;
    document.getElementById('exportActions').style.display = 'block';
    document.getElementById('exportBtn').disabled = false;
}

/**
 * Show export error message
 * @param {string} message - Error message to display
 */
function showExportError(message) {
    hideExportProgress();
    document.getElementById('exportError').style.display = 'block';
    document.getElementById('exportErrorMsg').innerText = message;
    document.getElementById('exportBtn').disabled = false;
}

/**
 * Show import progress container
 */
function showImportProgress() {
    document.getElementById('importProgress').style.display = 'block';
    document.getElementById('importSuccess').style.display = 'none';
    document.getElementById('importError').style.display = 'none';
}

/**
 * Hide import progress container
 */
function hideImportProgress() {
    document.getElementById('importProgress').style.display = 'none';
}

/**
 * Show import success message
 * @param {string} message - Success message to display
 */
function showImportSuccess(message) {
    hideImportProgress();
    document.getElementById('importSuccess').style.display = 'block';
    document.getElementById('importSuccessMsg').innerHTML = message;
    document.getElementById('btnImport').disabled = false;
}

/**
 * Show import error message
 * @param {string} message - Error message to display
 */
function showImportError(message) {
    hideImportProgress();
    document.getElementById('importError').style.display = 'block';
    document.getElementById('importErrorMsg').innerHTML = message;
    document.getElementById('btnImport').disabled = false;
}

