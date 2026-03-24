/**
 * @include https://www.googleapis.com/auth/drive
 */

// ============================================================
//  OnSite Check-in — Code.gs
//  ใช้ google.script.run จาก Index.html (ไม่ใช้ fetch/URL)
// ============================================================

// ─── Config ──────────────────────────────────────────────────
var SHEET_NAME = 'CheckIn';
var FOLDER_ID = '1R2ClGBWv8Hgp3tYpX8EMDjgcJw-nZYQr';  // Drive folder ID
var ADMIN_EMAIL = '';   // อีเมลแจ้งเตือน (ปล่อยว่างถ้าไม่ใช้)
// ─────────────────────────────────────────────────────────────


// ─── doGet: เสิร์ฟหน้า HTML ──────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('OnSite Field Check-in')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function testFullCheckin() {
  var file = DriveApp.getFolderById(FOLDER_ID).getFiles().next();
  var b64 = Utilities.base64Encode(file.getBlob().getBytes());
  
  var fakeData = {
    timestamp:   new Date().toISOString(),
    tsDisplay:   '24/03/2026, 16:00:00',
    staff:       'ธีรวัจน์ ครองโม่ง',
    client:      'คุณเบียร์',
    location:    'เวลโกรบางนา',
    job:         'เทส',
    lat:         '13.771248',
    lng:         '100.607687',
    accuracy:    '64',
    address:     'แขวงพลับพลา เขตวังทองหลาง กรุงเทพมหานคร 10310',
    mapLink:     'https://maps.google.com/?q=13.771248,100.607687',
    photoBase64: 'data:image/jpeg;base64,' + b64,
    purpose:     'นำเสนอ',
    custgroup:   'กรุงเทพ',
    contact:     'คุณเบียร์',
    prodgroup:   'A',
    progress:    'นำเสนอสินค้าเรียบร้อย รอติดตามผล',
    projvalue:   '500000'
  };

  var result = submitCheckin(fakeData);
  Logger.log(JSON.stringify(result));
}

// ─── submitCheckin ────────────────────────────────────────────
function submitCheckin(data) {
  try {
    var sheet = getOrCreateSheet();
    var photoLink = '';

    // ── Step 1: อัพรูปลง Drive (ถ้ามีข้อมูลรูปภาพ) ──────────────────────────────
    if (data && data.photoBase64 && data.photoBase64.length > 100) {
      try {
        var folder = DriveApp.getFolderById(FOLDER_ID);
        var b64 = data.photoBase64.includes(',') ? data.photoBase64.split(',')[1] : data.photoBase64;

        var safeName = (data.staff || 'unknown').replace(/[^a-zA-Zก-๙0-9]/g, '_');
        var fileName = safeName + '_' + new Date().getTime() + '.jpg';

        var blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', fileName);
        var file = folder.createFile(blob);
        
        // ดึง link ก่อน แล้วค่อย setSharing แยก
        photoLink = 'https://drive.google.com/file/d/' + file.getId() + '/view?usp=sharing';
        
        try {
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        } catch (shareErr) {
          Logger.log('Share warning (ไม่กระทบ link): ' + shareErr.toString());
        }

        Logger.log('✅ Photo uploaded: ' + photoLink);

      } catch (driveErr) {
        Logger.log('❌ Drive Error: ' + driveErr.toString());
        photoLink = '';
      }
    }

    // ── Step 2: เตรียม Formula สำหรับ Hyperlink ──────────────────────────────
    var mapFormula = data.mapLink ? '=HYPERLINK("' + data.mapLink + '","แผนที่")' : '';
    var photoFormula = (photoLink && photoLink.startsWith('http'))
      ? '=HYPERLINK("' + photoLink + '","ดูรูป")'
      : photoLink;

    // ── Step 3: บันทึกข้อมูลลง Sheet ในครั้งเดียว (Column A - R) ──────────────────
    sheet.appendRow([
      data.timestamp || '',   // A - Timestamp ISO
      data.tsDisplay || '',   // B - TimestampTH
      data.staff || '',   // C - Staff
      data.client || '',   // D - Client
      data.location || '',   // E - Location
      data.job || '',   // F - Job
      data.lat || '',   // G - Lat
      data.lng || '',   // H - Lng
      data.accuracy || '',   // I - Accuracy
      data.address || '',   // J - Address
      mapFormula,             // K - MapLink (เป็น Formula เลย)
      photoFormula,           // L - PhotoLink (เป็น Formula เลย)
      data.purpose || '',   // M - Purpose
      data.custgroup || '',   // N - CustomerGroup
      data.contact || '',   // O - Contact
      data.prodgroup || '',   // P - ProductGroup
      data.progress || '',   // Q - Progress
      data.projvalue || ''    // R - ProjectValue
    ]);

    // ── Step 4: ตกแต่ง Row ล่าสุด ───────────────────────────────────────────
    var lastRow = sheet.getLastRow();

    // จัด Style (ถ้ามีฟังก์ชันนี้อยู่แล้ว)
    if (typeof applyRowStyle === "function") {
      applyRowStyle(sheet, lastRow);
    }

    // ตั้งสีตัวอักษรสำหรับ Link ให้เป็นสีเขียว
    if (mapFormula) sheet.getRange(lastRow, 11).setFontColor('#1D9E75');
    if (photoFormula && photoLink.startsWith('http')) {
      sheet.getRange(lastRow, 12).setFontColor('#1D9E75');
    }

    // ── Step 5: แจ้งเตือนอีเมล (ใส่ใน try-catch เพื่อไม่ให้กระทบการบันทึกหลัก) ──────
    if (typeof ADMIN_EMAIL !== 'undefined' && ADMIN_EMAIL) {
      try {
        notifyEmail(data, photoLink);
      } catch (emailErr) {
        Logger.log('Email Error: ' + emailErr.toString());
      }
    }

    return {
      status: 'ok',
      photoLink: photoLink,
      row: lastRow
    };

  } catch (err) {
    Logger.log('Final Error: ' + err.toString());
    return { status: 'error', message: 'บันทึกไม่สำเร็จ: ' + err.toString() };
  }
}

// ─── getCheckinData ───────────────────────────────────────────
//     เรียกจาก: google.script.run.getCheckinData()
function getCheckinData() {
  try {
    var sheet = getOrCreateSheet();
    var rows = sheet.getDataRange().getValues();
    var formulas = sheet.getDataRange().getFormulas();

    var clean = rows.map(function (row, i) {
      return row.map(function (cell, j) {
        // ถ้า cell มี formula =HYPERLINK ให้ดึง URL ออกมา
        var formula = formulas[i][j] || '';
        if (formula.toUpperCase().indexOf('HYPERLINK') !== -1) {
          var match = formula.match(/HYPERLINK\("([^"]+)"/i);
          return match ? match[1] : '';
        }
        if (cell instanceof Error) return '';
        if (cell === null || cell === undefined) return '';
        return String(cell);
      });
    });

    return { status: 'ok', data: clean };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

// ─── getSummary ───────────────────────────────────────────────
function getSummary() {
  var rows = getOrCreateSheet().getDataRange().getValues();
  if (rows.length <= 1) return { total: 0, today: 0, clients: 0, staff: 0 };
  var data = rows.slice(1), today = new Date().toDateString();
  var clients = new Set(), staffSet = new Set(), todayCnt = 0;
  data.forEach(function (r) {
    if (r[0] && new Date(r[0]).toDateString() === today) todayCnt++;
    if (r[3]) clients.add(r[3]);
    if (r[2]) staffSet.add(r[2]);
  });
  return { total: data.length, today: todayCnt, clients: clients.size, staff: staffSet.size };
}


// ─── Helper: อัพรูป → Drive ──────────────────────────────────
function uploadPhoto(base64Str, staffName, timestamp) {
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var b64 = base64Str.includes(',') ? base64Str.split(',')[1] : base64Str;
  var safeName = (staffName || 'unknown').replace(/[^a-zA-Zก-๙0-9]/g, '_');
  var fileName = safeName + '_' + new Date(timestamp).getTime() + '.jpg';
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', fileName);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/file/d/' + file.getId() + '/view?usp=sharing';
}


// ─── Helper: Hyperlink ───────────────────────────────────────
function setHyperlink(sheet, row, col, url, label) {
  if (!url) return;
  var cell = sheet.getRange(row, col);
  cell.setFormula('=HYPERLINK("' + url + '","' + label + '")');
  cell.setFontColor('#1D9E75');
}


// ─── Helper: Style แถวข้อมูล ─────────────────────────────────
function applyRowStyle(sheet, rowNum) {
  var bg = (rowNum % 2 === 0) ? '#F9FAFB' : '#FFFFFF';
  var range = sheet.getRange(rowNum, 1, 1, 18);
  range.setBackground(bg).setFontSize(10).setFontColor('#111827').setVerticalAlignment('middle');
  sheet.setRowHeight(rowNum, 28);
  sheet.getRange(rowNum, 7, 1, 2).setFontFamily('Courier New').setFontColor('#00C896').setHorizontalAlignment('right');
  sheet.getRange(rowNum, 1, 1, 2).setHorizontalAlignment('center');
}


// ─── Helper: ดึง/สร้าง Sheet ──────────────────────────────────
function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { sheet = ss.insertSheet(SHEET_NAME); setupHeaders(sheet); }
  return sheet;
}


// ─── Helper: ตั้ง Headers ────────────────────────────────────
function setupHeaders(sheet) {
  sheet.appendRow(['Timestamp', 'TimestampTH', 'Staff', 'Client', 'Location', 'Job', 'Lat', 'Lng', 'Accuracy', 'Address', 'MapLink', 'PhotoLink', 'Purpose', 'CustomerGroup', 'Contact', 'ProductGroup', 'Progress', 'ProjectValue']);
  var h = sheet.getRange(1, 1, 1, 18);
  h.setBackground('#0a0b0d').setFontColor('#00C896').setFontWeight('bold')
    .setFontFamily('Courier New').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 34);
  sheet.setFrozenRows(1);
  [175, 145, 130, 155, 155, 200, 90, 90, 80, 230, 100, 100, 120, 120, 130, 130, 200, 100].forEach(function (w, i) { sheet.setColumnWidth(i + 1, w); });
}


// ─── Helper: Email ───────────────────────────────────────────
function notifyEmail(data, photoLink) {
  try {
    GmailApp.sendEmail(
      ADMIN_EMAIL,
      '[OnSite] ' + (data.staff || '') + ' → ' + (data.client || ''),
      'พนักงาน : ' + (data.staff || '-') + '\nลูกค้า  : ' + (data.client || '-') +
      '\nสถานที่ : ' + (data.location || '-') + '\nงาน     : ' + (data.job || '-') +
      '\nเวลา    : ' + (data.tsDisplay || '-') + '\nGPS     : ' + (data.lat || '-') + ', ' + (data.lng || '-') +
      '\nแผนที่  : ' + (data.mapLink || '-') + '\nรูป     : ' + (photoLink || '-')
    );
  } catch (e) { }
}


// ─── Menu ─────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('OnSite Admin')
    .addItem('Reformat sheet', 'reformatAll')
    .addSeparator()
    .addItem('ดู Summary', 'showSummary')
    .addToUi();
}

function reformatAll() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert('ไม่พบ sheet "' + SHEET_NAME + '"'); return; }
  var last = sheet.getLastRow();
  setupHeaders(sheet);
  for (var r = 2; r <= last; r++) applyRowStyle(sheet, r);
  sheet.getRange(1, 1, last, 18).setBorder(true, true, true, true, true, true, '#222629', SpreadsheetApp.BorderStyle.SOLID);
  SpreadsheetApp.getUi().alert('Reformat เรียบร้อย (' + (last - 1) + ' แถว)');
}

function showSummary() {
  var s = getSummary();
  SpreadsheetApp.getUi().alert('OnSite Summary',
    'เช็คอินทั้งหมด : ' + s.total + ' ครั้ง\nวันนี้          : ' + s.today + ' ครั้ง\nลูกค้า          : ' + s.clients + ' บริษัท\nทีมงาน         : ' + s.staff + ' คน',
    SpreadsheetApp.getUi().ButtonSet.OK);
}
