/**
 * Student Participation Certificate Automation
 * Google Form -> Google Sheet -> Google Doc template -> PDF -> Email
 *
 * Template placeholders (inside the Google Doc): {{Name}}, {{Event}}, {{Date}}
 * Sheet columns (row 1 headers): Timestamp | Name | Email | Event | Status
 */
const CONFIG = {
  TEMPLATE_DOC_ID: 'PASTE_TEMPLATE_DOC_ID',
  OUTPUT_FOLDER_ID: 'PASTE_DRIVE_FOLDER_ID',
  SHEET_NAME: 'Form Responses 1',
  COL: { NAME: 2, EMAIL: 3, EVENT: 4, STATUS: 5 }   // 1-based column numbers
};

/** Run once: Triggers > Add trigger > generateAllCertificates (optional). */
function generateAllCertificates() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][CONFIG.COL.STATUS - 1] === 'Sent') continue;
    processRow_(sheet, i + 1, rows[i]);
  }
}

/** Installable trigger: From spreadsheet > On form submit. */
function onFormSubmit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  processRow_(sheet, row, sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0]);
}

function processRow_(sheet, rowIndex, row) {
  const name  = row[CONFIG.COL.NAME - 1];
  const email = row[CONFIG.COL.EMAIL - 1];
  const event = row[CONFIG.COL.EVENT - 1];
  try {
    const pdf = createCertificate_(name, event);
    MailApp.sendEmail({
      to: email,
      subject: 'Your Participation Certificate - ' + event,
      body: 'Hi ' + name + ',\n\nThank you for participating in ' + event +
            '. Your certificate is attached.\n\nRegards,\nOrganizing Team',
      attachments: [pdf]
    });
    sheet.getRange(rowIndex, CONFIG.COL.STATUS).setValue('Sent');
  } catch (err) {
    sheet.getRange(rowIndex, CONFIG.COL.STATUS).setValue('Error: ' + err.message);
  }
}

function createCertificate_(name, event) {
  const folder = DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID);
  const copy = DriveApp.getFileById(CONFIG.TEMPLATE_DOC_ID)
                       .makeCopy('Certificate - ' + name, folder);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();
  body.replaceText('{{Name}}', name);
  body.replaceText('{{Event}}', event);
  body.replaceText('{{Date}}', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy'));
  doc.saveAndClose();

  const pdf = copy.getAs(MimeType.PDF).setName('Certificate - ' + name + '.pdf');
  folder.createFile(pdf);
  copy.setTrashed(true);      // keep only the PDF
  return pdf;
}
