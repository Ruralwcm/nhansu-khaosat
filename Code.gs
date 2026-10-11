// ========================================
// GAS Backend - Khảo sát Nhân sự Rural
// Copy to script.google.com → Deploy as Web App
// ========================================

var SHEET_ID = "1G5FVzbnLEPBVbLlYM664UEMQ5a6v49G7grQHzSR9_Tw";

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action || "submit";
    
    if (action === "import_master") return importMaster(data);
    return submitSurvey(data);
  } catch(err) {
    return json({success: false, error: err.toString()});
  }
}

function doGet(e) {
  var action = e.parameter.action || "";
  if (action === "history") return getHistory();
  if (action === "stores") return getStores();
  return json({status: "OK", action: action || "none"});
}

// ========== SURVEY SUBMISSION ==========
function submitSurvey(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("NhanSu_KhaoSat");
  if (!sheet) {
    sheet = ss.insertSheet("NhanSu_KhaoSat");
    sheet.appendRow(["Thời gian","Mã CH","Tên CH","GĐV","QLKV","Tỉnh","MNV","Họ tên","Chức danh","TT Hệ thống","Xác nhận","TT Khảo sát","Ngày bắt đầu","Ngày kết thúc","Đã gửi HR","Ghi chú"]);
  }
  
  var ts = data.ts || new Date().toISOString();
  var maCH = data.maCost || "", tenCH = data.tenCH || "", gdv = data.gdv || "", qlkv = data.qlkv || "", tinh = data.tinh || "";
  var employees = data.employees || {}, empMeta = data.empMeta || {};
  
  var rows = [];
  for (var mnv in employees) {
    var emp = employees[mnv], meta = empMeta[mnv] || {};
    rows.push([ts, maCH, tenCH, gdv, qlkv, tinh, mnv, meta.n||"", meta.c||"", meta.t||"", emp.correct ? "Đúng" : "Sai", emp.s||"", emp.startDate||"", emp.endDate||"", emp.sentHR ? "Có" : "Không", emp.note||""]);
  }
  if (rows.length > 0) sheet.getRange(sheet.getLastRow()+1, 1, rows.length, rows[0].length).setValues(rows);
  return json({success: true, rows: rows.length});
}

// ========== IMPORT MASTER DATA ==========
function importMaster(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("NhanSu_Master");
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet("NhanSu_Master");
  
  if (data.headers) sheet.appendRow(data.headers);
  var rows = data.rows || [];
  if (rows.length > 0) {
    // Batch write in chunks of 5000
    for (var i = 0; i < rows.length; i += 5000) {
      var chunk = rows.slice(i, Math.min(i+5000, rows.length));
      sheet.getRange(sheet.getLastRow()+1, 1, chunk.length, chunk[0].length).setValues(chunk);
    }
  }
  return json({success: true, total: rows.length});
}

// ========== GET SURVEY HISTORY ==========
function getHistory() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("NhanSu_KhaoSat");
  if (!sheet) return json({history: []});
  
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return json({history: []});
  
  var headers = data[0];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) row[headers[j]] = data[i][j];
    result.push(row);
  }
  return json({history: result});
}

// ========== GET STORE LIST ==========
function getStores() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("NhanSu_Master");
  if (!sheet) return json({stores: []});
  
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return json({stores: []});
  
  var headers = data[0];
  var stores = {};
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var cost = String(r[5] || ""); // Mã Cost column
    if (!stores[cost]) stores[cost] = {maCost: cost, tenCH: r[6]||"", gdv: r[2]||"", qlkv: r[3]||"", tinh: r[4]||"", employees: []};
    stores[cost].employees.push({mnv: r[0], hoTen: r[8], chucDanh: r[9], gioiTinh: r[10], trinhTrang: r[11], startDate: r[12], endDate: r[13]});
  }
  return json({stores: Object.values(stores)});
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}