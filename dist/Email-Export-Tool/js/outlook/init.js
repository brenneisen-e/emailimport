/**
 * Outlook Initialization Module
 * Initialize Outlook COM objects and load mailboxes
 */

/**
 * Initialize Outlook application and namespace
 */
function initOutlook() {
    try {
        outlookApp = new ActiveXObject("Outlook.Application");
        outlookNS = outlookApp.GetNamespace("MAPI");
        loadMailboxes();
    } catch (e) {
        showExportError("Outlook nicht gefunden: " + e.message);
    }
}

/**
 * Load available mailboxes into dropdown
 */
function loadMailboxes() {
    var select = document.getElementById('mailboxSelect');
    select.innerHTML = '<option value="">-- Postfach wahlen --</option>';
    try {
        var stores = outlookNS.Stores;
        for (var i = 1; i <= stores.Count; i++) {
            try {
                var store = stores.Item(i);
                var option = document.createElement('option');
                option.value = i;
                option.text = store.DisplayName;
                select.appendChild(option);
            } catch (e) {}
        }
    } catch (e) {}
}

/**
 * Load folders for selected mailbox
 */
function loadFolders() {
    var mailboxSelect = document.getElementById('mailboxSelect');
    var folderSelect = document.getElementById('folderSelect');
    var exportBtn = document.getElementById('exportBtn');

    if (!mailboxSelect.value) {
        folderSelect.innerHTML = '<option value="">-- Erst Postfach wahlen --</option>';
        folderSelect.disabled = true;
        exportBtn.disabled = true;
        return;
    }

    try {
        selectedStore = outlookNS.Stores.Item(parseInt(mailboxSelect.value));
        var rootFolder = selectedStore.GetRootFolder();
        folderSelect.innerHTML = '';

        try {
            var inbox = selectedStore.GetDefaultFolder(6);
            var option = document.createElement('option');
            option.value = 'inbox';
            option.text = 'Posteingang (' + inbox.Items.Count + ')';
            folderSelect.appendChild(option);
        } catch (e) {}

        addSubfolders(rootFolder.Folders, folderSelect, '');
        folderSelect.disabled = false;
        checkExportReady();

        // Step 2 completed - Mailbox selected
        completeStep(2);
    } catch (e) {
        showExportError("Fehler: " + e.message);
    }
}

/**
 * Recursively add subfolders to dropdown
 * @param {Object} folders - Outlook Folders collection
 * @param {HTMLSelectElement} select - Dropdown element
 * @param {string} prefix - Indentation prefix
 */
function addSubfolders(folders, select, prefix) {
    for (var i = 1; i <= folders.Count; i++) {
        try {
            var folder = folders.Item(i);
            if (folder.DefaultItemType === 0) {
                var option = document.createElement('option');
                option.value = folder.EntryID;
                option.text = prefix + folder.Name + ' (' + folder.Items.Count + ')';
                select.appendChild(option);
                if (folder.Folders.Count > 0) {
                    addSubfolders(folder.Folders, select, prefix + '  ');
                }
            }
        } catch (e) {}
    }
}

