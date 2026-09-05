// Dynamic PIN stored in Script Properties with fallback
function getAdminPin() {
  var prop = PropertiesService.getScriptProperties().getProperty("ADMIN_PIN");
  return prop ? prop : "admin123";
}

function setAdminPin(newPin) {
  PropertiesService.getScriptProperties().setProperty("ADMIN_PIN", newPin);
  return "Admin passcode updated successfully!";
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('ATHER ENERGY // ATTENDANCE PORTAL')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. Role-based Login & PIN Verification
function loginUser(role, pin, empId) {
  if (role === "admin") {
    return pin === getAdminPin() ? { success: true, role: "admin" } : { success: false, message: "Invalid Admin Passcode!" };
  } else if (role === "employee") {
    var details = getEmployeeDetails(empId);
    return details.found ? { success: true, role: "employee", empId: empId, name: details.name, email: details.email } : { success: false, message: "Emp ID not found in Master Sheet!" };
  }
}

// Helper: Check if day is Sunday
function isSunday(dayNum, monthSheetName) {
  try {
    var parts = monthSheetName.split("-");
    var monthStr = parts[0];
    var yearStr = parts[1];

    var monthNames = ["Jan", "Feb", "Mar", "April", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    var monthIdx = -1;
    for (var i = 0; i < monthNames.length; i++) {
      if (monthNames[i].toLowerCase() === monthStr.toLowerCase()) {
        monthIdx = i;
        break;
      }
    }

    if (monthIdx === -1) return false;
    var fullYear = 2000 + parseInt(yearStr, 10);
    var dateObj = new Date(fullYear, monthIdx, parseInt(dayNum, 10));
    return dateObj.getDay() === 0;
  } catch (e) {
    return false;
  }
}

// 2. Lookup Employee Details (ID in Col A, Name in Col B)
function getEmployeeDetails(empId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Sheet");
  if (!sheet) return { found: false, name: '', email: '' };
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { found: false, name: '', email: '' };

  var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toUpperCase() === empId.toString().trim().toUpperCase()) {
      return { found: true, name: data[i][1], email: "" };
    }
  }
  return { found: false, name: '', email: '' };
}

// 3. Fetch Full Employee Records List for Admin Panel
function getFullEmployeeRecords() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Sheet");
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  var list = [];

  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) {
      list.push({
        empId: data[i][0].toString().trim(),
        name: data[i][1] || "",
        contractName: data[i][2] || "",
        staying: data[i][3] || "",
        nativePlace: data[i][4] || "",
        doj: data[i][5] ? String(data[i][5]).split("T")[0] : "",
        ph: data[i][6] || "",
        reportingManager: data[i][7] || ""
      });
    }
  }
  return list;
}

// 4. Add New Employee to Master Sheet
function addEmployee(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Master Sheet");
  if (!sheet) return { success: false, message: "Master Sheet tab not found!" };

  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < existing.length; i++) {
      if (existing[i][0] && existing[i][0].toString().trim().toUpperCase() === data.empId.toString().trim().toUpperCase()) {
        return { success: false, message: "Employee ID (" + data.empId + ") already exists!" };
      }
    }
  }

  sheet.appendRow([
    data.empId,
    data.name,
    data.contractName,
    data.staying,
    data.nativePlace,
    data.doj,
    data.ph,
    data.reportingManager
  ]);

  return { success: true, message: "Employee " + data.name + " (" + data.empId + ") added successfully!" };
}

// 5. Remove Employee from Master Sheet
function removeEmployee(empId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Master Sheet");
  if (!sheet) return { success: false, message: "Master Sheet tab not found!" };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: "No data in Master Sheet." };

  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toUpperCase() === empId.toString().trim().toUpperCase()) {
      sheet.deleteRow(i + 2);
      return { success: true, message: "Employee " + empId + " removed successfully!" };
    }
  }

  return { success: false, message: "Employee ID " + empId + " not found." };
}

// 6. Master Employee List Dropdown Helper
function getMasterEmployeeList() {
  var records = getFullEmployeeRecords();
  return records.map(r => ({ name: r.name, empId: r.empId }));
}

// 7. Update Daily Attendance directly in Month Sheet
function updateDailyAttendance(monthSheetName, empId, targetDay, newStatus) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(monthSheetName);
  if (!sheet) return "Error: Tab '" + monthSheetName + "' not found.";

  var data = sheet.getDataRange().getValues();
  var headers = data[0]; 

  var colIndex = -1;
  for (var c = 3; c < headers.length; c++) {
    if (headers[c].toString().trim() === targetDay.toString().trim()) {
      colIndex = c;
      break;
    }
  }

  if (colIndex === -1) return "Error: Date " + targetDay + " not found.";

  var rowIndex = -1;
  for (var r = 1; r < data.length; r++) {
    if (data[r][0] && data[r][0].toString().trim().toUpperCase() === empId.toString().trim().toUpperCase()) {
      rowIndex = r + 1;
      break;
    }
  }

  if (rowIndex === -1) return "Error: Emp ID " + empId + " not found.";

  sheet.getRange(rowIndex, colIndex + 1).setValue(newStatus);
  return "Attendance for " + empId + " on Day " + targetDay + " updated to '" + newStatus + "'!";
}

// 8. Fetch Single Employee Calendar
function getCalendarData(empId, monthSheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(monthSheetName);
  
  if (!sheet) return { success: false, message: "Tab '" + monthSheetName + "' not found." };

  var data = sheet.getDataRange().getValues();
  var dayHeaders = data[0]; 
  var empRow = null;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toUpperCase() === empId.toString().trim().toUpperCase()) {
      empRow = data[i];
      break;
    }
  }

  if (!empRow) return { success: false, message: "Record for " + empId + " not found." };

  var daysMap = {};
  var presentCount = 0;
  var totalWorkingDays = 0;

  for (var col = 3; col < dayHeaders.length; col++) {
    var dayNum = dayHeaders[col];
    if (dayNum !== "" && dayNum !== undefined) {
      var cellVal = empRow[col] !== undefined ? empRow[col].toString().trim() : "";
      
      if ((cellVal === "" || cellVal === "-") && isSunday(dayNum, monthSheetName)) {
        cellVal = "W/o";
      }

      var cleanVal = cellVal.toUpperCase();
      daysMap[dayNum] = cellVal;

      if (cleanVal !== "W/O" && cleanVal !== "HOL" && cleanVal !== "-" && cellVal !== "") {
        totalWorkingDays++;
        if (cleanVal === "P" || cleanVal === "DHL") {
          presentCount++;
        }
      }
    }
  }

  var attendancePercentage = totalWorkingDays > 0 ? ((presentCount / totalWorkingDays) * 100).toFixed(1) : "0.0";

  return { 
    success: true, 
    name: empRow[1], 
    attendance: daysMap,
    percentage: attendancePercentage,
    presentCount: presentCount,
    workingDays: totalWorkingDays
  };
}

// 9. Fetch Overall Team Calendar
function getOverallTeamCalendar(monthSheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(monthSheetName);
  if (!sheet) return { success: false, message: "Tab '" + monthSheetName + "' not found." };

  var data = sheet.getDataRange().getValues();
  var headers = data[0]; 
  var days = [];

  for (var col = 3; col < headers.length; col++) {
    if (headers[col] !== "" && headers[col] !== undefined) {
      days.push(headers[col]);
    }
  }

  var teamData = [];
  for (var i = 1; i < data.length; i++) {
    var empId = data[i][0];
    var empName = data[i][1];

    if (empName && empId) {
      var dailyMap = {};
      for (var c = 3; c < headers.length; c++) {
        var dayNum = headers[c];
        if (dayNum !== "" && dayNum !== undefined) {
          var val = data[i][c] !== undefined ? data[i][c].toString().trim() : "";
          if ((val === "" || val === "-") && isSunday(dayNum, monthSheetName)) {
            val = "W/o";
          }
          dailyMap[dayNum] = val;
        }
      }
      teamData.push({ empName: empName, empId: empId, dailyStatus: dailyMap });
    }
  }

  return { success: true, days: days, teamData: teamData };
}

// 10. Fetch Team Attendance %
function getAllEmployeesAttendanceStats(monthSheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(monthSheetName);
  if (!sheet) return { success: false, message: "Tab '" + monthSheetName + "' not found." };

  var data = sheet.getDataRange().getValues();
  var dayHeaders = data[0];
  var statsList = [];

  for (var i = 1; i < data.length; i++) {
    var empId = data[i][0];
    var empName = data[i][1];
    
    if (empName && empId) {
      var presentCount = 0;
      var totalWorkingDays = 0;

      for (var col = 3; col < dayHeaders.length; col++) {
        var dayNum = dayHeaders[col];
        var cellVal = data[i][col] !== undefined ? data[i][col].toString().trim() : "";
        if ((cellVal === "" || cellVal === "-") && isSunday(dayNum, monthSheetName)) {
          cellVal = "W/o";
        }

        var cleanVal = cellVal.toUpperCase();
        if (cleanVal !== "W/O" && cleanVal !== "HOL" && cleanVal !== "-" && cellVal !== "") {
          totalWorkingDays++;
          if (cleanVal === "P" || cleanVal === "DHL") {
            presentCount++;
          }
        }
      }

      var pct = totalWorkingDays > 0 ? parseFloat(((presentCount / totalWorkingDays) * 100).toFixed(1)) : 0.0;
      statsList.push({ empName: empName, empId: empId, present: presentCount, workingDays: totalWorkingDays, percentage: pct });
    }
  }

  return { success: true, list: statsList };
}

// 11. Submit Request
function submitRequest(formData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheetName = formData.requestType === "Leave" ? "Leave Requests" : "Permission Requests";
  var sheet = ss.getSheetByName(targetSheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(targetSheetName);
    sheet.appendRow(["Timestamp", "Emp ID", "Name", "Type", "Start Date", "End Date", "Reason", "Status"]);
  }
  
  sheet.appendRow([new Date(), formData.empId, formData.empName, formData.requestType, formData.startDate, formData.endDate, formData.reason, "Pending"]);
  return "Request submitted successfully!";
}

// 12. Fetch Pending Requests
function getPendingRequests() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var requests = [];

  ["Leave Requests", "Permission Requests"].forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var status = data[i][7] ? data[i][7].toString().trim() : "";
        if (status === "Pending") {
          requests.push({
            sheetName: sheetName,
            rowIndex: i + 1,
            empId: data[i][1] ? data[i][1].toString() : "",
            empName: data[i][2] ? data[i][2].toString() : "",
            type: data[i][3] ? data[i][3].toString() : "",
            startDate: data[i][4] ? String(data[i][4]) : "",
            endDate: data[i][5] ? String(data[i][5]) : "",
            reason: data[i][6] ? data[i][6].toString() : ""
          });
        }
      }
    }
  });
  return requests;
}

// 13. Process Approval
function processApproval(sheetName, rowIndex, status) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Error: Sheet Tab Disconnected";
  
  sheet.getRange(rowIndex, 8).setValue(status);
  return "Status updated to " + status.toUpperCase();
}

// 14. Admin Analytics
function getAdminDashboardData(selectedMonth, targetDay) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("Master Sheet");
  var monthSheet = ss.getSheetByName(selectedMonth);

  var totalEmployees = masterSheet ? (masterSheet.getLastRow() - 1) : 0;
  var leaveList = [];
  var sameDateList = [];
  var totalAbsentCount = 0;

  var requestReasons = {};
  ["Leave Requests", "Permission Requests"].forEach(function(sheetName) {
    var reqSheet = ss.getSheetByName(sheetName);
    if (reqSheet) {
      var reqData = reqSheet.getDataRange().getValues();
      for (var r = 1; r < reqData.length; r++) {
        var empId = reqData[r][1] ? reqData[r][1].toString().trim().toUpperCase() : "";
        var reason = reqData[r][6] ? reqData[r][6].toString().trim() : "";
        var status = reqData[r][7] ? reqData[r][7].toString().trim() : "";
        if (empId && status === "Approved") {
          requestReasons[empId] = reason;
        }
      }
    }
  });

  if (monthSheet) {
    var data = monthSheet.getDataRange().getValues();
    var headers = data[0]; 
    
    var targetColIndex = -1;
    if (targetDay) {
      for (var c = 3; c < headers.length; c++) {
        if (headers[c].toString().trim() === targetDay.toString().trim()) {
          targetColIndex = c;
          break;
        }
      }
    }

    for (var i = 1; i < data.length; i++) {
      var empId = data[i][0] ? data[i][0].toString().trim() : "";
      var empName = data[i][1];
      var empIdKey = empId.toUpperCase();
      
      var offDays = [];
      var lastStatus = "";

      for (var col = 3; col < headers.length; col++) {
        var dateNum = headers[col]; 
        var status = data[i][col] ? data[i][col].toString().trim() : "";
        
        if ((status === "" || status === "-") && isSunday(dateNum, selectedMonth)) {
          status = "W/o";
        }

        var cleanStatus = status.toUpperCase();
        
        if (status !== "" && status !== "-" && cleanStatus !== "P" && cleanStatus !== "DHL" && cleanStatus !== "W/O" && cleanStatus !== "HOL") {
          totalAbsentCount++;
          lastStatus = cleanStatus;
          offDays.push("Date " + dateNum + " (" + status + ")");
        }

        if (col === targetColIndex && status !== "" && status !== "-" && cleanStatus !== "P" && cleanStatus !== "DHL" && cleanStatus !== "W/O" && cleanStatus !== "HOL") {
          sameDateList.push({
            empName: empName,
            empId: empId,
            status: status,
            reason: requestReasons[empIdKey] || "No Reason Specified"
          });
        }
      }

      if (offDays.length > 0) {
        var reasonText = requestReasons[empIdKey] || "Reason not updated";
        var leaveCategory = "Leave";

        if (lastStatus === "PER") {
          leaveCategory = "Without Leave (Permission)";
        }

        leaveList.push({
          empName: empName,
          empId: empId,
          days: offDays.join(", "),
          category: leaveCategory,
          reason: reasonText
        });
      }
    }
  }

  return { totalEmployees: totalEmployees, totalAbsences: totalAbsentCount, leaveList: leaveList, sameDateList: sameDateList };
}