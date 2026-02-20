/**
 * !!! IMPORTANT !!!
 * If you see permission errors, select 'authorizeScript' in the dropdown and click "Run" once.
 */
function authorizeScript() {
  console.log("Authorizing...");
  const ss = SpreadsheetApp.create("Temp_Auth_Check");
  DriveApp.getFileById(ss.getId()).setTrashed(true);
  const quota = MailApp.getRemainingDailyQuota();
  console.log("Permissions granted. You can now use the Web App.");
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('dialog')
      .evaluate()
      .setTitle('Invoice Generator')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * v1.1 PRODUCTION LOGIC
 * - Sequencing tracked by parsing full ID in cell H8.
 * - Updated PDF Filename format.
 */
function generateAndSendInvoice(requestData) {
  try {
    // --- CONFIGURATION ---
    const SPREADSHEET_ID = '1FS14tsaxxxpKT3yN1_zpfGxoMBbBSAZl5g273i7D2_8';
    const SHEET_ID = '790763898';
    
    // --- 1. Open Sheet & Handle Sequencing (Cell H8) ---
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheets()[0]; 
    
    // Get the FULL string from H8 (e.g., "IMI1120250050")
    let lastInvoiceID = sheet.getRange("H8").getValue().toString();
    let currentSequence = 50; // Default start if H8 is empty or invalid
    
    // Extract the last 4 digits using Regex
    const match = lastInvoiceID.match(/(\d{4})$/);
    if (match) {
      currentSequence = parseInt(match[1], 10);
    }
    
    // Increment sequence
    const nextSequence = currentSequence + 1;
    const sequenceStr = String(nextSequence).padStart(4, '0');

    // --- 2. Prepare Invoice Data ---
    const getMonthNum = (mon) => {
      const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
      const idx = months.indexOf(mon);
      return idx > -1 ? String(idx + 1).padStart(2, '0') : '00';
    };
    
    const monthNum = getMonthNum(requestData.month);
    
    // Build NEW Full Invoice Number
    const invoiceNumber = `IMI${monthNum}${requestData.year}${sequenceStr}`;
    
    const description = `Payment for cleaning of IMI ${requestData.month} ${requestData.year}`;
    const amount = parseFloat(requestData.amount);
    
    // --- 3. Update Google Sheet Invoice Fields ---
    
    // Update Tracking Cell H8 with the NEW FULL ID
    sheet.getRange("H8").setValue(invoiceNumber);
    
    // Update Invoice # field H10 (Visible on PDF)
    sheet.getRange("H10").setValue(invoiceNumber);
    
    // Description: B21
    sheet.getRange("B21").setValue(description);
    
    // Amount: H21
    sheet.getRange("H21").setValue(amount);
    
    // Force updates before PDF generation
    SpreadsheetApp.flush();

    // --- 4. Export Sheet as PDF ---
    // NEW: Construct custom filename
    const pdfFilename = `Cory's Cleaning Invoice - ${requestData.month} ${requestData.year} ${invoiceNumber}`;
    const pdfBlob = getPdfFromSheet(SPREADSHEET_ID, SHEET_ID, pdfFilename);
    
    // --- 5. Send Email ---
    const recipientList = [];
    recipientList.push('christianne.herrera13@gmail.com'); 
    recipientList.push('ms_enriquez2002@yahoo.ca'); // ENABLED for Production
    
    const recipientString = recipientList.join(',');
    const subject = `Invoice ${invoiceNumber} from Cory's Cleaning`;
    const body = `Dear Team,\n\nThe Google Sheet has been updated and the invoice generated.\n\n` +
                 `Description: ${description}\n` +
                 `Total Due: $${amount.toFixed(2)}\n\n` +
                 `Please find the invoice attached.\n\n` +
                 `Regards,\nAutomated Invoice System`;

    if (recipientString) {
      GmailApp.sendEmail(recipientString, subject, body, {
        attachments: [pdfBlob],
        name: "Cory's Cleaning"
      });
      return { 
        success: true, 
        message: `Invoice ${invoiceNumber} generated & sent to: ${recipientString}` 
      };
    } else {
      return { success: true, message: `Invoice ${invoiceNumber} saved. (No recipients enabled).` };
    }

  } catch (e) {
    console.error(e);
    return { success: false, message: "Error: " + e.toString() };
  }
}

function getPdfFromSheet(ssId, sheetId, filename) {
  const token = ScriptApp.getOAuthToken();
  const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=pdf` +
              `&gid=${sheetId}&size=7&portrait=true&fitw=true&gridlines=false` +
              `&printtitle=false&sheetnames=false&pagenumbers=false&fzr=false`;

  const options = {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    throw new Error("Failed to export PDF: " + response.getContentText());
  }
  return response.getBlob().setName(`${filename}.pdf`);
}
