/**
 * GOOGLE APPS SCRIPT FOR HIRARC APPROVAL
 * 
 * Instructions:
 * 1. Open your new Google Sheet designed for HIRARC Approvals.
 * 2. Click Extensions > Apps Script.
 * 3. Delete all code and paste this entire file.
 * 4. Click Deploy > New Deployment.
 * 5. Choose "Web app", set "Who has access" to "Anyone".
 * 6. Copy the Web App URL and paste it into line 250 of hirarc.html
 */

function doPost(e) {
  var headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST",
    "Access-Control-Allow-Headers": "Content-Type"
  };

  try {
    var data = JSON.parse(e.postData.contents);
    
    // This tells the script to put the data into Tab 2 ("Signed Approvals")
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Signed Approvals");
    
    // Failsafe: if the tab name is different, just use whatever is active
    if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    }

    var timestamp = new Date();
    
    // Now expecting 8 columns exactly!
    sheet.appendRow([
      timestamp,               // Col A: Timestamp
      data.department,         // Col B: Department
      data.owner_name,         // Col C: Area Owner Name
      data.employee_id,        // Col D: Employee ID
      data.date,               // Col E: Review Date
      data.document_version,   // Col F: Document Version (The ultra-smart CSV data!)
      data.declaration,        // Col G: ISO Acknowledgment
      data.signature           // Col H: Digital Signature (Base64)
    ]);
    
    return ContentService.createTextOutput(JSON.stringify({"status": "success"}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      "status": "error", 
      "message": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
