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
var THAI_MONTHS = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
var MIN_ZONE_ROWS = 45;
var OP_ROWS_PER_SHEET = 30;
var CHARTER_FILLS = ['#92D050', '#CCC0DA', '#B8CCE4', '#E6B8B7', '#FFFF00'];

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
    issueDateBE: opts.issueDate ? toBE(opts.issueDate) : toBE(Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd'))
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
  var widths = [65, 260, 40, 40, 85, 85, 110, 125, 35, 140];
  for (var c = 0; c < widths.length; c++) ws.setColumnWidth(c + 1, widths[c]);

  ws.getRange('A1:J1').merge().setValue(COMPANY_BILLING)
    .setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center');
  ws.getRange('A2:H2').merge().setValue('ใบเบิกเงินค่าบริการรถรับ-ส่งพนักงาน')
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  ws.getRange('I2').setValue('วันที่').setFontWeight('bold').setHorizontalAlignment('right');
  ws.getRange('J2').setValue(ctx.issueDateBE).setNumberFormat('@').setHorizontalAlignment('center');

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

    ws.getRange('A1:J1').merge().setValue(COMPANY_BILLING)
      .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
    ws.getRange('K1').setValue('แผ่นที่').setFontWeight('bold').setHorizontalAlignment('center');
    ws.getRange('A2:J2').merge().setValue('ใบบันทึกการปฏิบัติงานรถรับ-ส่งพนักงาน')
      .setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center');
    ws.getRange('K2').setValue(pageNo).setFontWeight('bold').setHorizontalAlignment('center');
    ws.getRange('A4:L4').merge().setValue('โครงการ: ' + (ctx.siteName || 'ทุกโครงการ') + '   |   เส้นทาง: ' + cleanBase)
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
