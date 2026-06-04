// ========================================================
// METROD AUTOMATION SYSTEM - PPE MANAGEMENT BACKEND SCRIPT
// ========================================================

const DASHBOARD_PIN = PropertiesService.getScriptProperties().getProperty("DASHBOARD_PIN") || "9911";

// Safe sheet target helper
function getPpeSheet(ss) {
  var name = "PPE Requests";
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

// 1. SETUP SHEET HEADERS & STYLING
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getPpeSheet(ss);
  
  const headers = [
    "Request ID", "Timestamp", "Staff ID", "Staff Name", "Department", 
    "Supervisor Name", "PPE Type", "Size", "Color/Specs", "Replacement Reason", 
    "Condition Remarks", "Status", "Authorized By", "Action Date"
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#0f766e").setFontColor("#ffffff"); // Teal theme
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  
  try {
    SpreadsheetApp.getUi().alert("🎉 PPE Requests Setup Complete!", "Headers have been formatted safely.", SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (err) {
    Logger.log("🎉 PPE Requests Setup Complete! Headers formatted successfully.");
  }
}

// Helper to generate sequential Request ID (e.g., REQ-00001)
function getNextRequestId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return "REQ-00001";
  }
  const lastReqId = sheet.getRange(lastRow, 1).getValue().toString();
  const match = lastReqId.match(/^REQ-(\d+)$/);
  if (match) {
    const nextNum = parseInt(match[1], 10) + 1;
    return "REQ-" + nextNum.toString().padStart(5, '0');
  }
  return "REQ-" + (lastRow).toString().padStart(5, '0');
}

// 2. WEB APP POST HANDLER (CREATION & APPROVALS)
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getPpeSheet(ss);
    const data = JSON.parse(e.postData.contents);
    const timestamp = new Date();
    
    // ACTION A: Approve/Reject or Dispatch Pending Request
    if (data.action === "updateRequestStatus") {
      if (data.pin !== DASHBOARD_PIN) {
        return ContentService.createTextOutput(JSON.stringify({ status: "ERROR", message: "Unauthorized PIN" })).setMimeType(ContentService.MimeType.JSON);
      }
      
      const rows = sheet.getDataRange().getValues();
      let foundRow = -1;
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === String(data.requestId).trim()) {
          foundRow = i + 1;
          break;
        }
      }
      
      if (foundRow !== -1) {
        sheet.getRange(foundRow, 12).setValue(data.status); // Update Status (Col L)
        sheet.getRange(foundRow, 13).setValue(data.authorizedBy); // Update Authorized By (Col M)
        sheet.getRange(foundRow, 14).setValue(new Date()); // Update Action Date (Col N)
        SpreadsheetApp.flush();
        return ContentService.createTextOutput(JSON.stringify({ status: "SUCCESS", message: "Request status updated" })).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({ status: "ERROR", message: "Request ID not found" })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    // ACTION B: Log New Request
    const requestId = getNextRequestId(sheet);
    const status = data.status || "Approved / Dispatched";
    const actionDate = (status !== "Pending Approval") ? timestamp : "";
    const authorizedBy = (status !== "Pending Approval") ? (data.authorizedBy || "Safety Officer") : "";
    
    sheet.appendRow([
      requestId,
      timestamp,
      data.staffId,
      data.staffName,
      data.department,
      data.supervisorName || "SHO",
      data.ppeType,
      data.size || "-",
      data.colorSpecs || "-",
      data.replacementReason || "Damaged",
      data.conditionRemarks || "-",
      status,
      authorizedBy,
      actionDate
    ]);
    
    SpreadsheetApp.flush();
    return ContentService.createTextOutput(JSON.stringify({ status: "SUCCESS", requestId: requestId })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "ERROR", message: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 3. WEB APP GET HANDLER (LOOKUPS & ANALYTICS)
function doGet(e) {
  try {
    const action = e.parameter.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getPpeSheet(ss);
    
    // ACTION A: Check Last Issue (Smart 6-Month Warning Lookup)
    if (action === "checkLastIssue") {
      const staffId = e.parameter.staffId;
      const ppeType = e.parameter.ppeType;
      
      if (!staffId || !ppeType) {
        return ContentService.createTextOutput(JSON.stringify({ status: "ERROR", message: "Missing staffId or ppeType" })).setMimeType(ContentService.MimeType.JSON);
      }
      
      const rows = sheet.getDataRange().getValues();
      let lastIssueDateObj = null;
      
      // Loop backwards from bottom to find the latest approved/dispatched record matching staffId and ppeType
      for (let i = rows.length - 1; i >= 1; i--) {
        const rowStaffId = String(rows[i][2]).trim().toLowerCase();
        const rowPpeType = String(rows[i][6]).trim().toLowerCase();
        const rowStatus = String(rows[i][11]).trim().toLowerCase();
        
        if (rowStaffId === staffId.trim().toLowerCase() && 
            rowPpeType === ppeType.trim().toLowerCase() && 
            rowStatus.indexOf("approved") !== -1) {
          
          // Use Action Date (index 13) or fall back to Timestamp (index 1)
          const dateVal = rows[i][13] || rows[i][1];
          if (dateVal instanceof Date) {
            lastIssueDateObj = dateVal;
            break;
          }
        }
      }
      
      if (lastIssueDateObj) {
        const today = new Date();
        const diffMonths = (today.getFullYear() - lastIssueDateObj.getFullYear()) * 12 + (today.getMonth() - lastIssueDateObj.getMonth());
        const formattedDate = Utilities.formatDate(lastIssueDateObj, "GMT+8", "yyyy-MM-dd");
        
        return ContentService.createTextOutput(JSON.stringify({
          status: "SUCCESS",
          found: true,
          lastDate: formattedDate,
          diffMonths: diffMonths
        })).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({
          status: "SUCCESS",
          found: false
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    // ACTION B: Get All Requests (Requires PIN)
    const pin = e.parameter.pin;
    if (pin !== DASHBOARD_PIN) {
      return ContentService.createTextOutput(JSON.stringify({ status: "ERROR", message: "Unauthorized PIN" })).setMimeType(ContentService.MimeType.JSON);
    }
    
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
    
    return ContentService.createTextOutput(JSON.stringify({ status: "SUCCESS", data: data })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "ERROR", message: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}
