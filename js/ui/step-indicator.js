/**
 * Step Indicator Module
 * Visual workflow step indicator and button state management
 */

// Step tracking state
var currentStep = 1;
var completedSteps = {};

/**
 * Update step indicator UI and button states
 * @param {number} stepNumber - Step to mark as completed (0 to just refresh)
 * @param {boolean} markCompleted - Whether to mark the step as completed
 */
function updateStepIndicator(stepNumber, markCompleted) {
    // Mark the specified step as completed if requested
    if (markCompleted && stepNumber > 0 && stepNumber <= 5) {
        completedSteps[stepNumber] = true;
    }

    // Update all step circles
    for (var i = 1; i <= 5; i++) {
        var stepEl = document.getElementById('step' + i);
        if (stepEl) {
            // Remove all state classes
            stepEl.className = 'step-circle step-' + i;

            if (completedSteps[i]) {
                stepEl.className += ' completed';
            }
            if (i === currentStep && !completedSteps[i]) {
                stepEl.className += ' active';
            }
        }
    }

    // Update button/select states based on current step
    updateButtonStates();
}

/**
 * Update button states based on current step
 * Manages visual highlighting and disabled states
 */
function updateButtonStates() {
    // Step 1: Excel verbinden
    var btnExcel = document.getElementById('btnExcelConnect');
    // Step 2: Postfach auswählen
    var selectMailbox = document.getElementById('mailboxSelect');
    var selectFolder = document.getElementById('folderSelect');
    // Step 3: Emails exportieren
    var btnExport = document.getElementById('exportBtn');
    // Step 4: Web-App öffnen
    var btnWebApp = document.getElementById('btnWebApp');
    // Step 5: In Excel importieren
    var btnImport = document.getElementById('btnImport');

    // Remove all step classes first
    var allElements = [btnExcel, selectMailbox, selectFolder, btnExport, btnWebApp, btnImport];
    for (var i = 0; i < allElements.length; i++) {
        if (allElements[i]) {
            allElements[i].className = allElements[i].className
                .replace(/\s*btn-active-step/g, '')
                .replace(/\s*btn-inactive-step/g, '')
                .replace(/\s*select-active-step/g, '')
                .replace(/\s*select-inactive-step/g, '');
        }
    }

    // Step 1: Excel verbinden
    if (currentStep === 1 && !completedSteps[1]) {
        if (btnExcel) btnExcel.className += ' btn-active-step';
        if (selectMailbox) selectMailbox.className += ' select-inactive-step';
        if (btnExport) btnExport.className += ' btn-inactive-step';
    } else if (completedSteps[1]) {
        if (btnExcel) btnExcel.className += ' btn-inactive-step';
    } else {
        if (btnExcel) btnExcel.className += ' btn-inactive-step';
    }

    // Step 2: Postfach/Ordner auswählen
    if (currentStep === 2 && !completedSteps[2]) {
        if (selectMailbox) selectMailbox.className += ' select-active-step';
        if (selectFolder) selectFolder.className += ' select-active-step';
        if (btnExport) btnExport.className += ' btn-inactive-step';
    } else if (completedSteps[2]) {
        if (selectMailbox) selectMailbox.className += ' select-inactive-step';
        if (selectFolder) selectFolder.className += ' select-inactive-step';
    } else {
        if (selectMailbox) selectMailbox.className += ' select-inactive-step';
        if (selectFolder) selectFolder.className += ' select-inactive-step';
    }

    // Step 3: Emails exportieren
    if (currentStep === 3 && !completedSteps[3]) {
        if (btnExport) btnExport.className += ' btn-active-step';
    } else if (completedSteps[3]) {
        if (btnExport) btnExport.className += ' btn-inactive-step';
    } else {
        if (btnExport) btnExport.className += ' btn-inactive-step';
    }

    // Step 4: Web-App öffnen
    if (currentStep === 4 && !completedSteps[4]) {
        if (btnWebApp) btnWebApp.className += ' btn-active-step';
    } else if (completedSteps[4]) {
        if (btnWebApp) btnWebApp.className += ' btn-inactive-step';
    }

    // Step 5: In Excel importieren
    if (currentStep === 5 && !completedSteps[5]) {
        if (btnImport) btnImport.className += ' btn-active-step';
    } else if (completedSteps[5]) {
        if (btnImport) btnImport.className += ' btn-inactive-step';
    } else {
        if (btnImport) btnImport.className += ' btn-inactive-step';
    }
}

/**
 * Advance to next step without marking current as completed
 * @param {number} stepNumber - Step number to advance to
 */
function advanceToStep(stepNumber) {
    if (stepNumber > currentStep) {
        currentStep = stepNumber;
        updateStepIndicator(0, false);
    }
}

/**
 * Complete a step and advance to next
 * @param {number} stepNumber - Step number to mark as completed
 */
function completeStep(stepNumber) {
    completedSteps[stepNumber] = true;
    if (stepNumber >= currentStep && stepNumber < 5) {
        currentStep = stepNumber + 1;
    }
    updateStepIndicator(stepNumber, true);
}

/**
 * Reset all steps to initial state
 */
function resetSteps() {
    currentStep = 1;
    completedSteps = {};
    updateStepIndicator(0, false);
}

