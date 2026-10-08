// Telliusco Website Lead Capture
// Deploy as Web App: Execute as "Me", Access "Anyone"
var SPREADSHEET_ID = '1NLlEQSPZPdSzXQ9sBg2-U_Gz4m2ud1ZaLSlATthdP6I';
var NOTIFY_EMAIL = 'info@telliusco.com';
var DRIVE_FOLDER_NAME = 'Telliusco Resumes';
function doPost(e) {
  try {
    var p = e.parameter || {};
    var type = p.inquiry_type || '';
    var isEmployer = type.indexOf('Employer') === 0;
    var now = new Date();
    var sheetName = isEmployer ? 'Hire Leads' : 'Apply Leads';
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    var resumeUrl = 'No file';
    if (e.files && e.files.resume) {
      var blob = e.files.resume;
      var folder = getOrCreateFolder(DRIVE_FOLDER_NAME);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      resumeUrl = file.getUrl();
    }
    var row;
    if (isEmployer) {
      row = [now, p.company || '', p.name || '', p.email || '', p.phone || '', '', '', '', '', '', p.message || '', 'Website', 'en'];
    } else {
      row = [now, p.name || '', p.email || '', p.phone || '', '', '', '', p.message || '', resumeUrl, 'Website', 'en'];
    }
    sheet.appendRow(row);
    var subject = 'New ' + (isEmployer ? 'Hire' : 'Apply') + ' Lead: ' + (p.name || p.company || '');
    var body = 'New lead from telliusco.com\n\n' + 'Type: ' + type + '\n' + 'Name: ' + (p.name || '') + '\n' + 'Company: ' + (p.company || '') + '\n' + 'Email: ' + (p.email || '') + '\n' + 'Phone: ' + (p.phone || '') + '\n' + 'Message: ' + (p.message || '') + '\n' + (resumeUrl !== 'No file' ? 'Resume: ' + resumeUrl + '\n' : '') + '\nLogged to: ' + sheetName;
    MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
    return ContentService.createTextOutput(JSON.stringify({ok: true})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: String(err)})).setMimeType(ContentService.MimeType.JSON);
  }
}
function getOrCreateFolder(name) {
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}
function doGet() {
  return ContentService.createTextOutput('Telliusco lead endpoint is live.');
}
