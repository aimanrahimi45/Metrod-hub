// --- RUN THIS FUNCTION FIRST TO AUTHORIZE ---
// (Handover Step: Select this function from the dropdown and click 'Run' to trigger Google's permission popup)
function setupAuthorization() {
  DriveApp.getFiles();
  SpreadsheetApp.getActiveSpreadsheet();
}

function doPost(e) {
  var headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST",
    "Access-Control-Allow-Headers": "Content-Type"
  };

  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var timestamp = new Date();
    
    // --- GOOGLE DRIVE PHOTO SAVING LOGIC ---
    // PASTE YOUR SUBFOLDER ID HERE!
    var FOLDER_ID = "1nf9uaVeJjxR5yOU3-JHpz30fX9yI2NCw";
    var photoUrl = "";
    
    if (data.photo) {
      // Decode the heavily compressed JPEG
      var base64Data = data.photo.split(",")[1];
      var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), "image/jpeg", "Induction_" + data.name + "_" + timestamp.getTime() + ".jpg");
      
      // Save to your specific Drive folder
      var folder = DriveApp.getFolderById(FOLDER_ID);
      var file = folder.createFile(blob);
      
      // Generate a shareable link to put in the Excel sheet
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      // Use the raw image URL so the Audit Hub can render it as a picture!
      photoUrl = "https://lh3.googleusercontent.com/d/" + file.getId();
    }
    // ---------------------------------------

    // Make sure your Google Sheet has 9 columns now!
    sheet.appendRow([
      timestamp,         // Col A
      data.name,         // Col B
      data.ic,           // Col C
      data.company,      // Col D
      data.date,         // Col E
      data.inducted_by,  // Col F
      data.declaration,  // Col G
      data.signature,    // Col H
      photoUrl           // Col I (The Live Photo Link!)
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

// ========================================================
// SECURE WEB APP GET LISTENER (DASHBOARD API)
// ========================================================
const DASHBOARD_PIN = PropertiesService.getScriptProperties().getProperty("DASHBOARD_PIN") || "9911";

function doGet(e) {
  try {
    const action = e.parameter.action;
    const pin = e.parameter.pin;
    
    // 1. SECURE SERVER-SIDE LOOKUP (No PIN required, only returns matching single status)
    if (action === "lookupWorker") {
      const searchIC = e.parameter.ic;
      if (!searchIC) {
        return ContentService.createTextOutput(JSON.stringify({ 
          status: "ERROR", 
          message: "IC search query is required" 
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getActiveSheet();
      const rows = sheet.getDataRange().getValues();
      
      // Col A: Timestamp, Col B: Name, Col C: IC Number
      const worker = rows.find(r => r[2] && String(r[2]).includes(searchIC));
      
      if (worker) {
        let inductionDateStr = "";
        if (worker[4] instanceof Date) {
          inductionDateStr = Utilities.formatDate(worker[4], "GMT+8", "yyyy-MM-dd");
        } else {
          inductionDateStr = String(worker[4]);
        }
        
        return ContentService.createTextOutput(JSON.stringify({
          status: "SUCCESS",
          found: true,
          name: worker[1], // Worker Name (Col B)
          date: inductionDateStr // Induction Date (Col E)
        })).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({
          status: "SUCCESS",
          found: false
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    // 2. DASHBOARD RETRIEVAL (Requires DASHBOARD_PIN)
    if (pin !== DASHBOARD_PIN) {
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "ERROR", 
        message: "Unauthorized: Invalid PIN" 
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];
    const data = [];
    
    for (let i = 1; i < rows.length; i++) {
      const obj = {};
      for (let j = 0; j < headers.length; j++) {
        let val = rows[i][j];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, "GMT+8", "yyyy-MM-dd HH:mm:ss");
        }
        obj[headers[j]] = val;
      }
      data.push(obj);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "SUCCESS",
      data: data
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "ERROR",
      message: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
