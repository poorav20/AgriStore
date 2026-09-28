// Google Apps Script — Paste this in Extensions → Apps Script
// Sheet ID: 1RP4-iGQzsAlfv1Dyq3z1hnol2gMrdnMuFatrVdz98bM

const SHEET_ID = '1RP4-iGQzsAlfv1Dyq3z1hnol2gMrdnMuFatrVdz98bM';

function doPost(e) {
    try {
        const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();

        // Add headers if sheet is empty
        if (sheet.getLastRow() === 0) {
            sheet.appendRow(['Timestamp', 'Name', 'Phone', 'Message']);
            sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
        }

        // Read form data sent as URL-encoded parameters
        const name = e.parameter.name || '';
        const phone = e.parameter.phone || '';
        const message = e.parameter.message || '';

        sheet.appendRow([
            new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }),
            name,
            phone,
            message
        ]);

        return ContentService
            .createTextOutput(JSON.stringify({ status: 'success' }))
            .setMimeType(ContentService.MimeType.JSON);
    } catch (error) {
        return ContentService
            .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
            .setMimeType(ContentService.MimeType.JSON);
    }
}

function doGet() {
    return ContentService
        .createTextOutput('Enquiry API is running!')
        .setMimeType(ContentService.MimeType.TEXT);
}
