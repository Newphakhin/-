/**
 * ============================================================================
 * CKP TRANSPORT — Google Apps Script สำหรับเชื่อมรับข้อมูลอัตโนมัติลง Google Sheets
 * รองรับทั้ง 2 รูปแบบ:
 *   1) สำเนาโซน (แยกแท็บตามเส้นทาง + แท็บงานเหมา + สูตร COUNTIF/SUM)
 *   2) ใบเบิกโครงการ (แท็บใบเบิกสรุปยอด + ใบบันทึกตัดหน้าแผ่นละ 30 แถว + แท็บเหมา)
 *
 * วิธีติดตั้ง (ทำครั้งเดียว 2 นาที):
 *   1. เปิดไฟล์ Google Sheets ของบริษัท
 *   2. ไปที่เมนู ส่วนขยาย (Extensions) -> Apps Script
 *   3. ลบโค้ดเดิมออกทั้งหมด แล้ววางโค้ดนี้ลงไป แล้วกดบันทึก (Ctrl+S)
 *   4. กดปุ่ม "ทำให้ใช้งานได้" (Deploy) -> "การทำให้ใช้งานได้รายการใหม่" (New deployment)
 *      - เลือกประเภท: เว็บแอป (Web app)
 *      - ดำเนินการในฐานะ (Execute as): ฉัน (Me)
 *      - ผู้ที่มีสิทธิ์เข้าถึง (Who has access): ทุกคน (Anyone)
 *   5. กด "ทำให้ใช้งานได้" (อนุญาตสิทธิ์ครั้งแรก) แล้วคัดลอก Web app URL มาวางในหน้าเว็บ CKP Transport
 * ============================================================================
 */

var COMPANY_ZONE = 'บริษัท จักรพงษ์ ขนส่ง (2003) จำกัด';
var COMPANY_BILLING = 'บริษัท จักรพงษ์ ขนส่ง ( 2003 ) จำกัด';
var DEFAULT_LOGO_URL = 'https://worker-dispatch-xi.vercel.app/logo.jpg';
var CKP_LOGO_JPG_BASE64 = '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Old02oltxbExVoU4nhgrCIJQnIZw0wbLG7ofulTmYfnMbNQ/QXphS9ijkQSUg5qCYrLRQ4wVwkPvPBYuge1Xb8FMmdWZ6rkfdReqWOloPIYe98R7/A7bsUiGZL1M+tyl1AidvCnM8mnCgQEViZjoN9Ql4xYzrlTum2C1jJkI9D65NPs7aZRgSh5Ek3qxTFlECU1bpVh2ZK9hiVMEsxs5HOj2PT3fSr4lwn1nncr2+jsG+gaAMaWgYPf+IS/AnIiHZxwJZCWCzcqgAaVeQAWGctqYHs0JwthMVUrFm5U8pEUpBEpV2WbUexwOFLQkEq6I1puHs56iJoNEMn2lYwbveuOa2K1J0E1fmfKz80TGcggo3ZGBMe5QcLiBZzu51bIGKJKGTgl+BOREOzj/AOH/AP/aAAwDAQACAAMCAAAQMM8888YsMw08888o888gE88888sY888o888c8888888sk88o88U48w400084sc8o84s08s8sogccwM8o8osMsMc4kEc8Ac8o88w88888oc88c48o88sY8gU8wg08E48o8oMMcI88EsY8UE8o884MIMc8cIM8I88o888g08888884c88o8888k08884sc888o8w88wQ8A8E8w4w0os0cQsUgA4Mks848o8UwEAYo8o8oQAo8owwwwwwwwwwwwwwwg/9oACAEDAAE/EP4Af//aAAgBAgABPxD+AH//2gAIAQECAT8QioKjygCrAXVsBUybWdfVwunJetS9npdn3k64exQkUIkiXP5+QFAef566R66lR/kj05Unws+jP9RShyF6q2nJy2ndZ/xNXuOg+p/+U6l2/fiPM6dKOjlbf6l6+etTJM4+5yrNQQ55B0HnHGtcKKgHk7DLzu9om77U6kBnMCrmv8CZniEIZIlC+3FOqh8vPQzNdTqzjiwKlf4ihJVhKRKsJKToPlOUnOoMp4BUjLlHdf8Ak/ZgQrJ2hEco5nk2nZdzsvphIP8AJG5UjpnBRC8kH4qmUjsxh70BkvME41hMCnBwDwGHGAz63TKy4x50binrq+SJC7j0391twwlI4Inj4mcc/BGlCHpwi6DVuM3Qt8L27BtX+MUIowqIvG763pruL6rTMHCPlaYUxkO2beAN99AcELUOd6PHBLcvb7alwrun4x9QPqhghNg/NWZPMTJ1Uc5vBn0Ww0g2HK6YVaa8q1PfNsUGXtg9Cq1HCIHAKa+D3rA75tglZB1zfmoQg5IyeT7cpX57cowAQFgpXLYLrVzsPf57tKkuFGVXVwmpHm2sRpXIYCLmXHolnXCZv7jLXfNsFMnBqqoG+a7+aYZLM4a9vV0pw75tjIPW+Zp4GXV/ZeSfkGacJUgEzy6Phmd1x62RmtcXfNqa1Nz9Rwl0CEajy65T2Y2HBi75tjhKJJkrym91dK/dT5PfNsQrHUIALqtRjMMtxvPwd82wgGEl6jbJ/A4UZH75u3HeSNfUw3yrye+bY8pRwFLpMn/lAROiEIR8HfNvCiLoHeBDaDyc8ATb96PhKcO+beDZihpez+1OPfNsIMyiqKPe0muMz98dV0Ma17xjitpQ9Wxhb+WTJ39UKsMroVRxKsPl9d0ocexlndlS8TBqDMRya75thyO/ZYVAMAkGRGkmvaTTQfFTh3zbwYSpazC5Y9zxo9MwzEkO4KdR4IIWk9is608EZzccUhZyvOjo5EfAfeiVCyzty9q3+7UYTBKzisds2xkdqK2EXZduo1NdOr4a75t4MJhSAGaulCLF5msZXrJ8bXqvJf7dRIM3Lj4oi7g4LNVYFTB5AtyUNd82xkAnKhYRLiJR5KXtT94QDnxO/wDtTih+8bqg75t4MKfkgp5LNLXaD5Ip2zgGE2OiV3zhf7Gjbn3z9VGSVpJCamcTCfVKq8r/AAFqcCB3nKuyPqiNh1dFtw5vbgmoa1QL4OCaMFhVJA7XlXZH1RGw6ui20Vpscuf43z5WsecFA4z5U+LcsOHlqV57VcE0uFuVmYTSUP8ACK0GPVFVzxvC9w8wRsZh5WpPxPrDR8/Z423iWpkiwp51AsXyy0mmpn+EC/BZT5hWtysuiojyF87VfwARAAiiRHRmvdCT7MulRiq/pw/JK7YUcA1wak96gQQQAQAaAfxmG2BibKs1zWf/AInxXKgv2PdhjzNPHO29cG3Zbu1PZRKy78aP0lInfNdnrQZNgQmwWeYih6hQf9h+U+/IGXgU5SoXIkl1eodG6wKNGMKJEdRMAZWTxKDDQzRBWwEryqPLx+IFCFbBKrbKo0ComCYYog8QYBqrR2lfOoaIqwDqJVm9Ux9ka1bWc2lNdftgiESUYANVa9pP2HyouIiAodRKPXckG42FTz6iqTd2uNzZKTQpNwoP+w/KffkDLwPAwhwWL6tMGWQugSRHTRvjBJbhnbPVVd7TmUAPXhkMQ4JmKJIZP3vau1bqX1jvKvRkI1MikEU6mrHsEOVcMxv8fas4kFi5MMauBC5qyka9SzQM8tuqOpVzee+8sAHC5w529zoHR0wb5yeN0wUipkOJMj5K0wZZC6BJEdNE4z8hBFaHvrvVIcnuRL+LnXfBc+VgiVliZrBit923VY2rn3Zzs/RwnCGp/sR5YGDVqMgU2ZuqB2wV9ItK7w3QixBV7Zvc7KXi0879mOMYdBT2L17XXP7BVzorLzg0HwK8tABMMjIrk+ihK69vT/oUzh76J91QdzrURRbF3yftB13bdhHaAtsZOXW6zpn4JNOkFLch/USZdKMtBsnsH4I8tJdzDWuCignjnRCRq/emz/k28gBa5HB2aywDQQIJEOeiyEAORWY3FW6eM6voKIuUvPBEFAkjXnKBSlUl7wkVABWyUaBQn26qo8LjIjUdCuJ5R3ND1UKK7gh722BElyKMp2NMc8qMRz4PHAomRDMApQsUFMmdzOnvoaLF5l2dOHXJK0y1wtcjg7NZf+G//9k=';
var THAI_MONTHS = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
var MIN_ZONE_ROWS = 45;
var OP_ROWS_PER_SHEET = 30;
var CHARTER_FILLS = ['#92D050', '#CCC0DA', '#B8CCE4', '#E6B8B7', '#FFFF00'];

function placeCkpLogo(ws, rangeA1, logoUrl) {
  var cell = ws.getRange(rangeA1);
  cell.merge().setHorizontalAlignment('center').setVerticalAlignment('middle');

  var url = (logoUrl && /^https?:\/\//i.test(String(logoUrl))) ? String(logoUrl) : DEFAULT_LOGO_URL;

  // 1) รูปฝังในเซลล์จากลิงก์โลโก้บริษัท (พิมพ์/ส่งออก PDF ติดภาพชัดเจน ไม่ลอยทับข้อความ)
  if (typeof SpreadsheetApp.newCellImage === 'function') {
    try {
      var img = SpreadsheetApp.newCellImage()
        .setSourceUrl(url)
        .setAltTextTitle('CKP TRANSPORT')
        .setAltTextDescription('โลโก้บริษัท จักรพงษ์ ขนส่ง (2003) จำกัด')
        .build();
      cell.setValue(img);
      return true;
    } catch (_) {}

    // 2) สำรอง: ใช้ Data URI แบบฝังในโค้ดโดยตรง (ไม่ต้องพึ่งอินเทอร์เน็ตภายนอก)
    try {
      var imgData = SpreadsheetApp.newCellImage()
        .setSourceUrl('data:image/jpeg;base64,' + CKP_LOGO_JPG_BASE64)
        .setAltTextTitle('CKP TRANSPORT')
        .build();
      cell.setValue(imgData);
      return true;
    } catch (_) {}
  }

  // 3) สำรอง: สูตร =IMAGE(...) ของ Google Sheets (ปรับขนาดพอดีช่องอัตโนมัติ)
  try {
    cell.setFormula('=IMAGE("' + url.replace(/"/g, '') + '", 1)');
    return true;
  } catch (_) {}

  // 4) สำรองสุดท้าย: insertImage แบบลอยจาก Base64 Blob
  try {
    var blob = Utilities.newBlob(Utilities.base64Decode(CKP_LOGO_JPG_BASE64), 'image/jpeg', 'ckp-logo.jpg');
    var overImg = ws.insertImage(blob, 1, 1, 6, 4);
    if (overImg && typeof overImg.setWidth === 'function') {
      overImg.setWidth(80).setHeight(78);
    }
    return true;
  } catch (_) {}

  return false;
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return jsonOut({
    ok: true,
    status: 'ready',
    spreadsheetName: ss.getName(),
    spreadsheetUrl: ss.getUrl()
  });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    var body = JSON.parse(raw);

    if (body.action === 'ping') {
      return jsonOut({
        ok: true,
        spreadsheetName: ss.getName(),
        spreadsheetUrl: ss.getUrl()
      });
    }

    var kind = body.kind === 'billing' ? 'billing' : 'zone';
    var logs = Array.isArray(body.logs) ? body.logs : [];
    var sheets = (kind === 'billing')
      ? rebuildBillingSpreadsheet(ss, logs, body)
      : rebuildZoneSpreadsheet(ss, logs, body);

    SpreadsheetApp.flush();
    return jsonOut({
      ok: true,
      kind: kind,
      sheets: sheets,
      sheetCount: sheets.length,
      logCount: logs.length,
      spreadsheetName: ss.getName(),
      spreadsheetUrl: ss.getUrl(),
      updatedAt: new Date().toISOString()
    });
  } catch (err) {
    return jsonOut({
      ok: false,
      error: err && err.message ? String(err.message) : String(err)
    });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

// ----------------------------------------------------------------------------
// Helpers
// ----------------------------------------------------------------------------
function parseIso(iso) {
  var m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? { y: Number(m[1]), m: Number(m[2]), d: Number(m[3]) } : null;
}

function toBE(iso) {
  var p = parseIso(iso);
  return p ? (p.d + '/' + p.m + '/' + (p.y + 543)) : String(iso || '');
}

function showTime(t) {
  var m = String(t || '').match(/^(\d{1,2}):(\d{2})/);
  return m ? (Number(m[1]) + '.' + m[2]) : String(t || '');
}

function cleanSheetName(name, fallback) {
  var s = String(name || '').replace(/[\[\]\*\?\/\\:]/g, ' ').replace(/\s+/g, ' ').trim();
  return (s || fallback || 'Sheet').slice(0, 31).trim();
}

function createNamer() {
  var used = {};
  return function (raw, fallback) {
    var base = cleanSheetName(raw, fallback);
    var cand = base;
    var n = 2;
    while (used[cand.toLowerCase()]) {
      var sfx = ' (' + n + ')';
      cand = base.slice(0, 31 - sfx.length).trim() + sfx;
      n++;
    }
    used[cand.toLowerCase()] = true;
    return cand;
  };
}

function sortLogsChrono(logs) {
  return logs.slice().sort(function (a, b) {
    if (a.trip_date !== b.trip_date) return String(a.trip_date || '').localeCompare(String(b.trip_date || ''));
    var ta = String((a && (a.dep_time || a.arr_time)) || '99:99');
    var tb = String((b && (b.dep_time || b.arr_time)) || '99:99');
    var tc = ta.localeCompare(tb);
    if (tc !== 0) return tc;
    return String(a.id || '').localeCompare(String(b.id || ''));
  });
}

function groupByKey(items, keyFn) {
  var order = [];
  var map = {};
  for (var i = 0; i < items.length; i++) {
    var k = keyFn(items[i]);
    if (!map[k]) {
      map[k] = [];
      order.push(k);
    }
    map[k].push(items[i]);
  }
  return order.map(function (k) { return [k, map[k]]; });
}

function periodTextOf(start, end) {
  var a = parseIso(start);
  var b = parseIso(end);
  if (!a || !b) return '-';
  if (a.y === b.y && a.m === b.m) {
    if (a.d === b.d) return a.d + ' ' + THAI_MONTHS[a.m - 1] + ' ' + (a.y + 543);
    return a.d + ' - ' + b.d + ' ' + THAI_MONTHS[a.m - 1] + ' ' + (a.y + 543);
  }
  return toBE(start) + ' - ' + toBE(end);
}

function resetSpreadsheetWithTemp(ss) {
  var tmpName = '__ckp_sync_' + Date.now();
  var tmp = ss.insertSheet(tmpName);
  var all = ss.getSheets();
  for (var i = 0; i < all.length; i++) {
    if (all[i].getSheetId() !== tmp.getSheetId()) {
      ss.deleteSheet(all[i]);
    }
  }
  return tmp;
}

// ----------------------------------------------------------------------------
// 1) รูปแบบสำเนาโซน (Zone Copy)
// ----------------------------------------------------------------------------
function rebuildZoneSpreadsheet(ss, logs, opts) {
  var tmp = resetSpreadsheetWithTemp(ss);
  var namer = createNamer();
  var sorted = sortLogsChrono(logs);
  var normal = sorted.filter(function (l) { return l.trip_type !== 'charter'; });
  var charter = sorted.filter(function (l) { return l.trip_type === 'charter'; });
  var createdNames = [];

  var groups = groupByKey(normal, function (l) {
    return l.route_sheet_name || (l.origin + ' - ' + l.destination);
  });
  for (var g = 0; g < groups.length; g++) {
    var sName = namer(groups[g][0], 'สายวิ่ง');
    buildZoneSheet(ss.insertSheet(sName), groups[g][1], 'normal');
    createdNames.push(sName);
  }
  if (charter.length > 0) {
    var cName = namer('งานเหมา', 'งานเหมา');
    buildZoneSheet(ss.insertSheet(cName), charter, 'charter');
    createdNames.push(cName);
  }
  if (createdNames.length === 0) {
    var eName = namer('สำเนาโซน', 'สำเนาโซน');
    buildZoneSheet(ss.insertSheet(eName), [], 'normal');
    createdNames.push(eName);
  }

  ss.deleteSheet(tmp);
  return createdNames;
}

function buildZoneSheet(ws, rows, kind) {
  var isCharter = kind === 'charter';
  var colWidths = isCharter
    ? [72, 100, 105, 250, 72, 250, 95, 280]
    : [65, 92, 105, 190, 68, 95, 90, 120];
  for (var c = 0; c < colWidths.length; c++) ws.setColumnWidth(c + 1, colWidths[c]);

  // หัวกระดาษ 3 แถว
  ws.getRange('A1:H1').merge().setValue(COMPANY_ZONE)
    .setFontFamily('Arial').setFontSize(20).setFontWeight('bold').setHorizontalAlignment('center');
  ws.getRange('A2:H2').merge().setValue('บันทึกรายการรถ 6 ล้อ รับ-ส่งคนงาน')
    .setFontFamily('Arial').setFontSize(15).setFontWeight('bold').setHorizontalAlignment('center');
  ws.getRange('A3:H3').merge().setValue('โครงการ  ')
    .setFontFamily('Arial').setFontSize(13).setFontWeight('bold').setFontColor('#FF0000').setHorizontalAlignment('center');
  ws.setRowHeight(1, 34);
  ws.setRowHeight(2, 26);
  ws.setRowHeight(3, 22);
  ws.setRowHeight(4, 18);

  // แถว 5 หัวตาราง
  var headers = isCharter
    ? [['ลำดับ', 'วัน เดือน ปี', 'ทะเบียนรถ', 'จาก', 'เวลา', 'ถึง', 'ราคา', 'หมายเหตุ']]
    : [['ลำดับ', 'วัน เดือน ปี', 'ทะเบียนรถ', 'จาก', 'เวลา', 'จำนวนคน', 'ถึง', 'หมายเหตุ']];
  ws.getRange(5, 1, 1, 8)
    .setValues(headers)
    .setFontFamily('Arial').setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);

  // เตรียมตารางสรุปด้านล่าง (เฉพาะสายปกติ)
  var distinctColC = [];
  var seenC = {};
  var primaryOrigin = (!isCharter && rows.length > 0) ? String(rows[0].origin || '').trim() : '';
  if (!isCharter) {
    for (var i = 0; i < rows.length; i++) {
      var lbl = (rows[i].display_mode_zone === 'license_plate')
        ? String(rows[i].license_plate || rows[i].worker_name || '').trim()
        : String(rows[i].worker_name || rows[i].license_plate || '').trim();
      if (lbl && !seenC[lbl]) {
        seenC[lbl] = true;
        distinctColC.push(lbl);
      }
    }
  }

  var lastDataRow = 5 + rows.length;
  var sumStartRow = (!isCharter && distinctColC.length > 0) ? lastDataRow + 2 : 0;
  var sumEndRow = sumStartRow > 0 ? sumStartRow + distinctColC.length - 1 : 0;
  var sumTotalRow = sumEndRow > 0 ? sumEndRow + 1 : 0;
  var totalRows = Math.max(MIN_ZONE_ROWS, rows.length, sumTotalRow > 0 ? sumTotalRow - 5 : 0);

  var vals = [];
  var bgs = [];
  var aligns = [];
  var prevDate = '';
  var prevGroup = '';
  var groupIdx = -1;

  for (var rIdx = 0; rIdx < totalRows; rIdx++) {
    var rowNum = 6 + rIdx;
    var log = rows[rIdx];
    var rowVals = [rIdx + 1, '', '', '', '', '', '', ''];
    var rowBg = ['#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF'];
    var rowAl = ['center', 'center', 'center', 'center', 'center', 'center', 'center', 'left'];

    if (log) {
      var showDate = log.trip_date !== prevDate;
      prevDate = log.trip_date;
      if (isCharter) {
        var gKey = log.trip_date + '||' + log.origin + '||' + log.destination;
        if (gKey !== prevGroup) {
          prevGroup = gKey;
          groupIdx = (groupIdx + 1) % CHARTER_FILLS.length;
        }
        var fill = CHARTER_FILLS[groupIdx];
        rowVals[1] = showDate ? toBE(log.trip_date) : '';
        rowVals[2] = String(log.license_plate || log.worker_name || '');
        rowVals[3] = String(log.origin || '');
        rowVals[4] = showTime(log.dep_time);
        rowVals[5] = String(log.destination || '');
        rowVals[6] = log.unit_price ? Number(log.unit_price) : '';
        rowVals[7] = String(log.remark || '');
        for (var ci = 1; ci <= 6; ci++) rowBg[ci] = fill;
        if (rowVals[7]) rowBg[7] = fill;
        rowAl[3] = 'left';
        rowAl[5] = 'left';
        rowAl[6] = 'right';
      } else {
        var colC = (log.display_mode_zone === 'license_plate')
          ? String(log.license_plate || log.worker_name || '')
          : String(log.worker_name || log.license_plate || '');
        var isReturn = primaryOrigin && String(log.origin || '').trim() !== primaryOrigin;
        var hasPass = log.passengers !== null && log.passengers !== undefined && log.passengers !== '';
        var hasRem = Boolean(String(log.remark || '').trim());

        rowVals[1] = showDate ? toBE(log.trip_date) : '';
        rowVals[2] = colC;
        rowVals[3] = String(log.origin || '');
        rowVals[4] = showTime(log.dep_time);
        rowVals[5] = hasPass ? Number(log.passengers) : '';
        rowVals[6] = '';
        rowVals[7] = String(log.remark || '');

        rowBg[1] = '#FFFF00';
        rowBg[2] = '#FFFF00';
        rowBg[3] = '#FFFF00';
        rowBg[4] = '#FFFF00';
        if (hasPass) rowBg[5] = '#FFFF00';
        if (hasRem) rowBg[7] = '#FFFF00';
        rowAl[3] = isReturn ? 'right' : 'left';
        rowAl[7] = 'center';
      }
    } else if (sumStartRow > 0 && rowNum >= sumStartRow && rowNum <= sumEndRow) {
      var sLabel = distinctColC[rowNum - sumStartRow];
      rowVals[5] = sLabel;
      rowVals[6] = '=COUNTIF(C$6:C$' + lastDataRow + ',F' + rowNum + ')';
      rowBg[5] = '#92D050';
      rowBg[6] = '#FFFF00';
    } else if (sumTotalRow > 0 && rowNum === sumTotalRow) {
      rowVals[6] = '=SUM(G' + sumStartRow + ':G' + sumEndRow + ')';
      rowBg[6] = '#D99694';
    }

    vals.push(rowVals);
    bgs.push(rowBg);
    aligns.push(rowAl);
  }

  var dataRange = ws.getRange(6, 1, totalRows, 8);
  dataRange.setNumberFormat('@');
  dataRange.setValues(vals);
  dataRange.setBackgrounds(bgs);
  dataRange.setHorizontalAlignments(aligns);
  dataRange.setVerticalAlignment('middle');
  dataRange.setFontSize(10);
  dataRange.setBorder(true, true, true, true, true, true);

  if (isCharter) {
    ws.getRange(6, 7, totalRows, 1).setNumberFormat('#,##0.00');
  } else {
    ws.getRange(6, 6, totalRows, 1).setNumberFormat('#,##0');
    if (sumStartRow > 0) {
      ws.getRange(sumStartRow, 7, distinctColC.length + 1, 1).setNumberFormat('#,##0');
    }
  }
}

// ----------------------------------------------------------------------------
// 2) รูปแบบใบเบิกโครงการ (Billing Workbook)
// ----------------------------------------------------------------------------
function rebuildBillingSpreadsheet(ss, logs, opts) {
  var tmp = resetSpreadsheetWithTemp(ss);
  var namer = createNamer();
  var sorted = sortLogsChrono(logs);
  var normal = sorted.filter(function (l) { return l.trip_type !== 'charter'; });
  var charter = sorted.filter(function (l) { return l.trip_type === 'charter'; });

  var start = opts.dateFrom || (sorted[0] ? sorted[0].trip_date : '');
  var end = opts.dateTo || (sorted.length ? sorted[sorted.length - 1].trip_date : '');
  var ctx = {
    siteName: opts.siteName || '',
    periodText: periodTextOf(start, end),
    issueDateBE: opts.issueDate ? toBE(opts.issueDate) : toBE(Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd')),
    logoUrl: opts.logoUrl || DEFAULT_LOGO_URL
  };

  var createdNames = [];
  var sumName = namer('ใบเบิก', 'ใบเบิก');
  buildBillingSummarySheet(ss.insertSheet(sumName), normal, charter, ctx);
  createdNames.push(sumName);

  var pageNo = 1;
  var groups = groupByKey(normal, function (l) {
    return l.route_sheet_name || (l.origin + ' - ' + l.destination);
  });
  for (var g = 0; g < groups.length; g++) {
    pageNo = buildBillingOpSheets(ss, namer, groups[g][0], groups[g][1], pageNo, ctx, createdNames);
  }
  if (charter.length > 0) {
    pageNo = buildBillingOpSheets(ss, namer, 'เหมา', charter, pageNo, ctx, createdNames);
  }
  if (createdNames.length === 1) {
    buildBillingOpSheets(ss, namer, 'แผ่นที่ 1', [], pageNo, ctx, createdNames);
  }

  ss.deleteSheet(tmp);
  return createdNames;
}

function buildBillingSummarySheet(ws, normal, charter, ctx) {
  var widths = [95, 250, 40, 40, 85, 85, 110, 125, 35, 140];
  for (var c = 0; c < widths.length; c++) ws.setColumnWidth(c + 1, widths[c]);
  ws.setRowHeight(1, 34);
  ws.setRowHeight(2, 30);
  ws.setRowHeight(3, 26);

  placeCkpLogo(ws, 'A1:A3', ctx.logoUrl);

  ws.getRange('B1:J1').merge().setValue(COMPANY_BILLING)
    .setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.getRange('B2:H2').merge().setValue('ใบเบิกเงินค่าบริการรถรับ-ส่งพนักงาน')
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.getRange('I2').setValue('วันที่').setFontWeight('bold').setHorizontalAlignment('right').setVerticalAlignment('middle');
  ws.getRange('J2').setValue(ctx.issueDateBE).setNumberFormat('@').setHorizontalAlignment('center').setVerticalAlignment('middle');

  ws.getRange('A4:J4').merge().setValue('โครงการ: ' + (ctx.siteName || 'ทุกโครงการ'))
    .setFontSize(12).setFontWeight('bold').setHorizontalAlignment('left');
  ws.getRange('A5:J5').merge().setValue('งวดวันที่: ' + ctx.periodText)
    .setFontSize(11).setHorizontalAlignment('left');

  // หัวตาราง แถว 7
  ws.getRange('A7').setValue('ลำดับ');
  ws.getRange('B7:D7').merge().setValue('รายละเอียด');
  ws.getRange('E7').setValue('จำนวน');
  ws.getRange('F7').setValue('หน่วย');
  ws.getRange('G7').setValue('ราคา/หน่วย');
  ws.getRange('H7:I7').merge().setValue('จำนวนเงิน');
  ws.getRange('J7').setValue('หมายเหตุ');
  ws.getRange('A7:J7').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#F1F5F9').setBorder(true, true, true, true, true, true);

  // รวมรายการ
  var items = [];
  var nGroups = groupByKey(normal, function (l) {
    return (l.billing_title || l.route_sheet_name || (l.origin + ' - ' + l.destination)) + '||' + (Number(l.unit_price) || 0);
  });
  for (var i = 0; i < nGroups.length; i++) {
    var arr = nGroups[i][1];
    items.push({
      title: arr[0].billing_title || arr[0].route_sheet_name || (arr[0].origin + ' - ' + arr[0].destination),
      count: arr.length,
      unit: 'เที่ยว',
      price: Number(arr[0].unit_price) || 0,
      charter: false
    });
  }
  var cGroups = groupByKey(charter, function (l) {
    return (l.billing_title || 'เหมาขนย้าย') + '||' + (Number(l.unit_price) || 0);
  });
  for (var j = 0; j < cGroups.length; j++) {
    var cArr = cGroups[j][1];
    items.push({
      title: cArr[0].billing_title || 'เหมาขนย้าย',
      count: cArr.length,
      unit: 'เที่ยว',
      price: Number(cArr[0].unit_price) || 0,
      charter: true
    });
  }

  var rowSlots = Math.max(10, items.length);
  var tableEnd = 7 + rowSlots * 2;
  var totalRow = tableEnd + 1;

  for (var r = 8; r <= totalRow; r++) {
    ws.getRange(r, 2, 1, 3).merge();
    ws.getRange(r, 8, 1, 2).merge();
  }
  ws.getRange(8, 1, totalRow - 7, 10).setBorder(true, true, true, true, true, true).setVerticalAlignment('middle');

  for (var k = 0; k < items.length; k++) {
    var row = 8 + k * 2;
    var it = items[k];
    ws.getRange(row, 1).setValue(k + 1).setHorizontalAlignment('center');
    ws.getRange(row, 2).setValue(it.title).setHorizontalAlignment('left');
    ws.getRange(row, 5).setValue(it.count).setHorizontalAlignment('center').setNumberFormat('#,##0');
    ws.getRange(row, 6).setValue(it.unit).setHorizontalAlignment('center');
    ws.getRange(row, 7).setValue(it.price).setHorizontalAlignment('right').setNumberFormat('#,##0.00');
    ws.getRange(row, 8).setFormula('=SUM(E' + row + '*G' + row + ')').setHorizontalAlignment('right').setNumberFormat('#,##0.00');
    if (it.charter) {
      ws.getRange(row + 1, 2).setValue('( ขนย้าย )').setFontWeight('bold').setHorizontalAlignment('center');
    }
  }

  ws.getRange(totalRow, 2).setValue('รวม').setFontWeight('bold').setHorizontalAlignment('center');
  ws.getRange(totalRow, 8).setFormula('=SUM(H8:I' + tableEnd + ')').setFontWeight('bold').setHorizontalAlignment('right').setNumberFormat('#,##0.00');

  var sig1 = totalRow + 3;
  var sig2 = totalRow + 4;
  ws.getRange(sig1, 5).setValue('ผู้ตรวจรับ / Checked by');
  ws.getRange(sig1, 7, 1, 3).merge().setBorder(false, false, true, false, false, false);
  ws.getRange(sig2, 5).setValue('วันที่ / Date');
  ws.getRange(sig2, 7, 1, 3).merge().setBorder(false, false, true, false, false, false);
}

function buildBillingOpSheets(ss, namer, baseName, rows, startPageNo, ctx, createdNames) {
  var chunks = [];
  for (var i = 0; i < rows.length; i += OP_ROWS_PER_SHEET) {
    chunks.push(rows.slice(i, i + OP_ROWS_PER_SHEET));
  }
  if (chunks.length === 0) chunks.push([]);

  var cleanBase = cleanSheetName(baseName, 'ใบบันทึก');
  var pageNo = startPageNo;
  var prevDate = '';

  for (var k = 0; k < chunks.length; k++) {
    var suffix = k === 0 ? '' : ' (' + (k + 1) + ')';
    var rawName = k === 0 ? cleanBase : (cleanBase.slice(0, 31 - suffix.length).trim() + suffix);
    var sName = namer(rawName, 'ใบบันทึก');
    var ws = ss.insertSheet(sName);
    createdNames.push(sName);

    var widths = [55, 88, 95, 165, 65, 165, 30, 30, 65, 70, 70, 125];
    for (var c = 0; c < widths.length; c++) ws.setColumnWidth(c + 1, widths[c]);
    ws.setRowHeight(1, 32);
    ws.setRowHeight(2, 28);
    ws.setRowHeight(3, 26);

    placeCkpLogo(ws, 'A1:B3', ctx.logoUrl);

    ws.getRange('C1:J1').merge().setValue(COMPANY_BILLING)
      .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
    ws.getRange('K1:L1').merge().setValue('แผ่นที่').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
    ws.getRange('C2:J2').merge().setValue('ใบบันทึกการปฏิบัติงานรถรับ-ส่งพนักงาน')
      .setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
    ws.getRange('K2:L2').merge().setValue(pageNo).setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
    ws.getRange('A4:L4').merge().setValue('โครงการ: ' + (ctx.siteName || 'ทุกโครงการ') + '   |   เส้นทาง: ' + cleanBase + '   |   งวดวันที่: ' + ctx.periodText)
      .setFontSize(11).setFontWeight('bold');

    // หัวตารางแถว 6
    ws.getRange('A6').setValue('ลำดับ');
    ws.getRange('B6').setValue('วันที่');
    ws.getRange('C6').setValue('ทะเบียนรถ');
    ws.getRange('D6').setValue('ต้นทาง (จาก)');
    ws.getRange('E6').setValue('เวลาออก');
    ws.getRange('F6:H6').merge().setValue('ปลายทาง (ถึง)');
    ws.getRange('I6').setValue('เวลาถึง');
    ws.getRange('J6').setValue('จำนวนคน');
    ws.getRange('K6').setValue('เซ็นรับ');
    ws.getRange('L6').setValue('หมายเหตุ');
    ws.getRange('A6:L6').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBackground('#F1F5F9').setBorder(true, true, true, true, true, true);

    var chunk = chunks[k];
    var grid = [];
    for (var rIdx = 0; rIdx < OP_ROWS_PER_SHEET; rIdx++) {
      var rNum = 7 + rIdx;
      ws.getRange(rNum, 6, 1, 3).merge();
      var log = chunk[rIdx];
      var line = [k * OP_ROWS_PER_SHEET + rIdx + 1, '', '', '', '', '', '', '', '', '', '', ''];
      if (log) {
        var showDate = log.trip_date !== prevDate;
        prevDate = log.trip_date;
        line[1] = showDate ? toBE(log.trip_date) : '';
        line[2] = String(log.license_plate || '');
        line[3] = String(log.origin || '');
        line[4] = showTime(log.dep_time);
        line[5] = String(log.destination || '');
        line[8] = showTime(log.arr_time);
        line[9] = (log.passengers !== null && log.passengers !== undefined && log.passengers !== '') ? String(Number(log.passengers)) : '';
        line[11] = String(log.remark || '');
      }
      grid.push(line);
    }

    var bodyRange = ws.getRange(7, 1, OP_ROWS_PER_SHEET, 12);
    bodyRange.setNumberFormat('@');
    bodyRange.setValues(grid);
    bodyRange.setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(10)
      .setBorder(true, true, true, true, true, true);

    ws.getRange('H37').setValue('ผู้ตรวจ / Checked by').setFontWeight('bold');
    ws.getRange('K37:L37').merge().setValue('..............................................').setHorizontalAlignment('center');
    ws.getRange('H38').setValue('วันที่ / Date').setFontWeight('bold');
    ws.getRange('K38:L38').merge().setValue('......... / ............... / ............').setHorizontalAlignment('center');

    pageNo++;
  }
  return pageNo;
}
