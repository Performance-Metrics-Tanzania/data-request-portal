// Configuration
const SHEET_NAME = "Data Request Portal"; // Change to match your sheet name
const SPREADSHEET_ID = "1rASlDUz9iPfMduLaqsbLxYe_7s0XN555rBF7n0rJNBI"; // Replace with your Google Sheet ID
const NOTIFICATION_EMAILS = "mercy.idindili@experienceeducate.org,nemes.umela@experienceeducate.org";

/**
 * doPost function receives data from the HTML form
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);

    if (!sheet) {
      createNewSheet(SHEET_NAME, spreadsheet);
      sheet = spreadsheet.getSheetByName(SHEET_NAME);
    }

    const dateSubmitted = new Date(data.timestamp);
    const rowData = [
      dateSubmitted.toLocaleDateString('en-US'),
      data.name || "",
      data.requesterEmail || "",
      data.department || "",
      data.category || "",
      data.problem || "",
      data.priority || "Medium",
      data.deadline || "",
      data.status || "New",
      data.additionalNotes || ""
    ];

    sheet.appendRow(rowData);
    sendNotificationEmail(data);

    return ContentService.createTextOutput(
      JSON.stringify({ success: true, message: "Request submitted successfully" })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log('Error: ' + error);
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, message: error.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function sendNotificationEmail(data) {
  const subject = `New Data Request: ${data.category || 'Request'} (${data.priority || 'Medium'})`;
  const body = `
    <p>A new request was submitted through the portal.</p>
    <p><strong>Name:</strong> ${data.name || 'Unknown'}</p>
    <p><strong>Email:</strong> ${data.requesterEmail || 'Not provided'}</p>
    <p><strong>Department:</strong> ${data.department || 'Not provided'}</p>
    <p><strong>Category:</strong> ${data.category || 'Not provided'}</p>
    <p><strong>Priority:</strong> ${data.priority || 'Medium'}</p>
    <p><strong>Deadline:</strong> ${data.deadline || 'Not provided'}</p>
    <p><strong>Status:</strong> ${data.status || 'New'}</p>
    <p><strong>Problem/Request:</strong><br>${data.problem || 'Not provided'}</p>
    <p><strong>Notes:</strong><br>${data.additionalNotes || 'None'}</p>
  `;

  MailApp.sendEmail({
    to: NOTIFICATION_EMAILS,
    subject: subject,
    htmlBody: body
  });
}

/**
 * Creates a new sheet with proper headers
 */
function createNewSheet(sheetName, spreadsheet) {
  const newSheet = spreadsheet.insertSheet(sheetName);
  const headers = [
    "Date Submitted",
    "Requester Name",
    "Requester Email",
    "Team/Department",
    "Category",
    "Problem/Request",
    "Priority",
    "Requested Timeline",
    "Status",
    "Notes"
  ];

  newSheet.appendRow(headers);

  const headerRange = newSheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#667eea");
  headerRange.setFontColor("white");
  headerRange.setFontWeight("bold");

  newSheet.setColumnWidth(1, 120);
  newSheet.setColumnWidth(2, 150);
  newSheet.setColumnWidth(3, 180);
  newSheet.setColumnWidth(4, 150);
  newSheet.setColumnWidth(5, 140);
  newSheet.setColumnWidth(6, 260);
  newSheet.setColumnWidth(7, 100);
  newSheet.setColumnWidth(8, 120);
  newSheet.setColumnWidth(9, 100);
  newSheet.setColumnWidth(10, 200);
}

/**
 * Helper function to get the spreadsheet ID
 * Run this in the Apps Script editor to find your ID
 */
function getSpreadsheetId() {
  return SpreadsheetApp.getActiveSpreadsheet().getId();
}

/**
 * Helper function to get the Apps Script URL for deployment
 * After deploying, run this to get the URL
 */
function getDeploymentUrl() {
  return "Use the URL from the 'Deploy as web app' dialog";
}
