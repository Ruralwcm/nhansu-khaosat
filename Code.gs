// GAS Backend - Khao Sat Nhan Su
// Deploy as Web App, copy URL to index.html

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.openById("1MFb9UE76jL_fBRya0KKFOZm2JRej0troawqvoGMJ2OE");
    var sheet = ss.getSheetByName("NhanSu_KhaoSat");
    
    if (!sheet) {
      sheet = ss.insertSheet("NhanSu_KhaoSat");
      sheet.appendRow(["Thoi gian", "Ma CH", "Ten CH", "GDV", "QLKV", "Tinh", "MNV", "Ho ten", "Chuc danh", "TT He thong", "Xac nhan", "TT Khao sat", "Ngay bat dau", "Ngay ket thuc", "Da gui HR", "Ghi chu"]);
    }
    
    var timestamp = data.ts || new Date().toISOString();
    var maCH = data.maCost || "";
    var tenCH = data.tenCH || "";
    var gdv = data.gdv || "";
    var qlkv = data.qlkv || "";
    var tinh = data.tinh || "";
    
    var rows = [];
    var employees = data.employees || {};
    var empMeta = data.empMeta || {};
    
    for (var mnv in employees) {
      var emp = employees[mnv];
      var meta = empMeta[mnv] || {};
      rows.push([
        timestamp,
        maCH, tenCH, gdv, qlkv, tinh,
        mnv,
        meta.n || "",
        meta.c || "",
        meta.t || "",
        emp.correct ? "Dung" : "Sai",
        emp.s || "",
        emp.startDate || "",
        emp.endDate || "",
        emp.sentHR ? "Co" : "Khong",
        emp.note || ""
      ]);
    }
    
    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
    
    return ContentService.createTextOutput(JSON.stringify({success: true, rows: rows.length}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({success: false, error: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({status: "OK", action: e.parameter.action || "none"}))
    .setMimeType(ContentService.MimeType.JSON);
}