Office.onReady(function (info) {
    if (info.host === Office.HostType.Excel) {
        // The add-in is ready
        console.log('Add-in is ready');
    }
});

function openUrls() {
    Excel.run(function (context) {
        // Get the currently selected range
        var range = context.workbook.getSelectedRange();
        range.load("values");

        return context.sync().then(function () {
            const urls = range.values.flat(); // Flatten the array (Excel stores 2D arrays)
            const validUrls = urls.filter(url => isValidUrl(url));

            if (validUrls.length === 0) {
                Office.context.ui.displayDialogAsync("No valid URLs found.");
                return;
            }

            if (validUrls.length > 20) {
                // Confirm with the user before opening many URLs
                if (!confirm(`You are about to open ${validUrls.length} URLs. Do you want to continue?`)) {
                    return;
                }
            }

            validUrls.forEach(url => window.open(url, '_blank'));
        });
    }).catch(function (error) {
        console.error(error);
    });
}

function isValidUrl(string) {
    try {
        new URL(string);
        return true;
    } catch (_) {
        return false;
    }
}
