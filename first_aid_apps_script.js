/**
 * METROD AUTOMATION SYSTEM - FIRST AID BOX INSPECTION (RELATIONAL DATABASE SYSTEM)
 * 
 * Instructions:
 * 1. Open your Google Sheet.
 * 2. Click "Extensions" > "Apps Script".
 * 3. Replace all existing code with this new script.
 * 4. Select the "setupSheet" function in the dropdown at the top and click "Run".
 *    This will automatically create your two clean relational tabs:
 *    - "First Aid Checklist Logs"
 *    - "First Aid Checklist Details"
 *    and delete the old cluttered tab automatically!
 * 5. Click "Deploy" > "New Deployment" > Choose "Web App" > Set Execute as "Me" and Who has access to "Anyone" > Deploy.
 * 6. Copy the new Web App URL (if it changes) and paste it into first_aid.html!
 */

// Fetches the secure PIN from Google Apps Script private Project Properties.
// Default fallback is "9911" if not configured in your Settings panel.
const DASHBOARD_PIN = PropertiesService.getScriptProperties().getProperty("DASHBOARD_PIN") || "9911";

// ========================================================
// 1. SETUP RELATIONAL SHEET STRUCTURE
// ========================================================
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Setup logs sheet
  let logsSheet = ss.getSheetByName("First Aid Checklist Logs");
  if (!logsSheet) {
    logsSheet = ss.insertSheet("First Aid Checklist Logs");
  }
  
  // 2. Setup details sheet
  let detailsSheet = ss.getSheetByName("First Aid Checklist Details");
  if (!detailsSheet) {
    detailsSheet = ss.insertSheet("First Aid Checklist Details");
  }

  // Auto-delete the old cluttered 83-column sheet if it exists
  const oldSheet = ss.getSheetByName("First Aid Checklist");
  if (oldSheet) {
    try {
      ss.deleteSheet(oldSheet);
    } catch(err) {
      // Ignore if deletion is blocked or already deleted
    }
  }
  
  // Format Logs Headers
  const logHeaders = [
    "Audit ID", 
    "Timestamp", 
    "Date of Inspection", 
    "Company", 
    "Department", 
    "Section", 
    "Box ID", 
    "Location",
    "Cleanliness Condition",
    "Cleanliness Remarks",
    "Inspection Findings",
    "Inspected By Name",
    "Inspected By Position",
    "Signature URL"
  ];
  
  logsSheet.clear();
  logsSheet.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]);
  const logHeaderRange = logsSheet.getRange(1, 1, 1, logHeaders.length);
  logHeaderRange.setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
  logsSheet.setFrozenRows(1);
  logsSheet.autoResizeColumns(1, logHeaders.length);

  // Format Details Headers
  const detailHeaders = [
    "Audit ID",
    "Item ID",
    "Item Name",
    "Required Standard",
    "Quantity Available",
    "Expiry Date",
    "Remarks"
  ];

  detailsSheet.clear();
  detailsSheet.getRange(1, 1, 1, detailHeaders.length).setValues([detailHeaders]);
  const detailHeaderRange = detailsSheet.getRange(1, 1, 1, detailHeaders.length);
  detailHeaderRange.setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
  detailsSheet.setFrozenRows(1);
  detailsSheet.autoResizeColumns(1, detailHeaders.length);

  // Success alert dialog
  const ui = SpreadsheetApp.getUi();
  ui.alert("🎉 Success!", "Your Relational Two-Tab Database has been configured successfully!\n\n1. 'First Aid Checklist Logs' (14 columns)\n2. 'First Aid Checklist Details' (7 columns)\n\nThe old cluttered 83-column sheet has been deleted automatically. Ready to deploy your Web App!", ui.ButtonSet.OK);
}

// ========================================================
// 2. WEB APP POST LISTENER
// ========================================================
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logsSheet = ss.getSheetByName("First Aid Checklist Logs");
    const detailsSheet = ss.getSheetByName("First Aid Checklist Details");
    const data = JSON.parse(e.postData.contents);
    
    // Save signature base64 image as file in Google Drive
    let signatureUrl = "";
    if (data.signature) {
      const base64Data = data.signature.split(",")[1];
      const decodedBytes = Utilities.base64Decode(base64Data);
      const blob = Utilities.newBlob(decodedBytes, "image/png", `Sig_${data.boxId.replace(/\//g, '_')}_${data.inspectDate}.png`);
      
      const folderId = "1NEfd1I5zYDRXkhvizjokhX_K4JxBF293";
      let folder;
      try {
        folder = DriveApp.getFolderById(folderId);
      } catch (err) {
        // Fallback folder creation
        const folders = DriveApp.getFoldersByName("Metrod Signatures");
        if (folders.hasNext()) {
          folder = folders.next();
        } else {
          folder = DriveApp.createFolder("Metrod Signatures");
        }
      }
      
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      signatureUrl = file.getUrl();
    }
    
    // Generate Unique transaction Audit ID: FA-BOXCODE-YYYYMMDD-HHMMSS
    const boxCode = data.boxId.replace(/[^a-zA-Z0-9]/g, "");
    const dateCode = data.inspectDate.replace(/-/g, "");
    const timeCode = Utilities.formatDate(new Date(), "GMT+8", "HHmmss");
    const auditId = `FA-${boxCode}-${dateCode}-${timeCode}`;

    // Cleanliness Condition (Item 24) mappings
    const cleanCond = data[`item_24_avail`] || "Good";
    const cleanRem = data[`item_24_remarks`] || "-";

    // 1. Write metadata record to Logs sheet
    const logRow = [
      auditId,
      new Date(), // Timestamp
      data.inspectDate,
      data.company,
      data.department,
      data.section,
      data.boxId,
      data.location,
      cleanCond,
      cleanRem,
      data.findings || "-",
      data.officerName || "-",
      data.officerPos || "-",
      signatureUrl
    ];
    logsSheet.appendRow(logRow);
    
    // 2. Write individual items 1 to 23 vertically to Details sheet
    const items = [
      { id: 1, name: "Triangular Bandage 100cm", req: "5pcs" },
      { id: 2, name: "Eye Dressing No 16", req: "3pkt" },
      { id: 3, name: "Sterile Gamgee Pad 25cm", req: "3pkt" },
      { id: 4, name: "Sterile Gauze Pad 7.5cm", req: "6pkt" },
      { id: 5, name: "Sterile Gauze Pad 10cm", req: "6pkt" },
      { id: 6, name: "Elastic Bandage", req: "3pkt" },
      { id: 7, name: "W.O.W Bandage 2.5cm", req: "8pcs" },
      { id: 8, name: "W.O.W Bandage 5.0cm", req: "8pcs" },
      { id: 9, name: "W.O.W Bandage 7.5cm", req: "8pcs" },
      { id: 10, name: "Instant Ice Pack", req: "6pkt" },
      { id: 11, name: "Sterile Non-Adherent Pad", req: "6pkt" },
      { id: 12, name: "Pair of Glove", req: "6pkt" },
      { id: 13, name: "Scissors", req: "1pcs" },
      { id: 14, name: "Adhesive Tape", req: "1pcs" },
      { id: 15, name: "Bactigras", req: "2pcs" },
      { id: 16, name: "Yellow Antiseptic Liquid", req: "1pcs" },
      { id: 17, name: "Cotton Bud 100pcs", req: "1pkt" },
      { id: 18, name: "CPR Face Shield", req: "3pcs" },
      { id: 19, name: "Adhesive Plaster", req: "60pcs" },
      { id: 20, name: "Safety Pin", req: "36pcs" },
      { id: 21, name: "Thermometer", req: "1pcs" },
      { id: 22, name: "Waste Bag", req: "3pcs" },
      { id: 23, name: "First Aid Manual", req: "1pcs" }
    ];

    items.forEach(item => {
      const avail = data[`item_${item.id}_avail`] || "-";
      const exp = data[`item_${item.id}_exp`] || "-";
      const rem = data[`item_${item.id}_remarks`] || "-";

      const detailRow = [
        auditId,
        item.id,
        item.name,
        item.req,
        avail,
        exp,
        rem
      ];
      detailsSheet.appendRow(detailRow);
    });
    
    return ContentService.createTextOutput("SUCCESS").setMimeType(ContentService.MimeType.TEXT);
    
  } catch (err) {
    return ContentService.createTextOutput("ERROR: " + err.message).setMimeType(ContentService.MimeType.TEXT);
  }
}

// ========================================================
// 3. SECURE WEB APP GET LISTENER (DASHBOARD API)
// ========================================================
function doGet(e) {
  try {
    const pin = e.parameter.pin;
    
    // Check PIN authorization
    if (pin !== DASHBOARD_PIN) {
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "ERROR", 
        message: "Unauthorized: Invalid PIN" 
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    const action = e.parameter.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (action === "getLogs") {
      const sheet = ss.getSheetByName("First Aid Checklist Logs");
      const rows = sheet.getDataRange().getValues();
      const headers = rows[0];
      const data = [];
      
      for (let i = 1; i < rows.length; i++) {
        const obj = {};
        for (let j = 0; j < headers.length; j++) {
          let val = rows[i][j];
          if (val instanceof Date) {
            val = Utilities.formatDate(val, "GMT+8", "yyyy-MM-dd");
          }
          obj[headers[j]] = val;
        }
        data.push(obj);
      }
      
      return ContentService.createTextOutput(JSON.stringify({
        status: "SUCCESS",
        data: data
      })).setMimeType(ContentService.MimeType.JSON);
      
    } else if (action === "getDetails") {
      const sheet = ss.getSheetByName("First Aid Checklist Details");
      const rows = sheet.getDataRange().getValues();
      const headers = rows[0];
      const data = [];
      
      for (let i = 1; i < rows.length; i++) {
        const obj = {};
        for (let j = 0; j < headers.length; j++) {
          let val = rows[i][j];
          if (val instanceof Date) {
            val = Utilities.formatDate(val, "GMT+8", "yyyy-MM-dd");
          }
          obj[headers[j]] = val;
        }
        data.push(obj);
      }
      
      return ContentService.createTextOutput(JSON.stringify({
        status: "SUCCESS",
        data: data
      })).setMimeType(ContentService.MimeType.JSON);
      
    } else {
      return ContentService.createTextOutput(JSON.stringify({
        status: "ERROR",
        message: "Invalid Action"
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "ERROR",
      message: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
