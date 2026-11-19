/**
 * ตัวสร้างไฟล์ Excel (.xlsx) ด้วย ExcelJS — 2 รูปแบบตามเอกสารของบริษัท
 *   1) สำเนาโซน  : buildZoneCopyWorkbook(ExcelJS, logs, opts)
 *   2) ใบเบิกโครงการ : buildBillingWorkbook(ExcelJS, logs, opts)
 * ไม่มีข้อมูลตัวอย่างในโค้ดนี้ — ทุกแถวมาจากบันทึกการเดินทาง (travel_logs) ที่ส่งเข้ามาเท่านั้น
 */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.CkpExcel = api;
})(typeof window !== 'undefined' ? window : globalThis, function () {
  'use strict';

  const COMPANY = 'บริษัท จักรพงษ์ ขนส่ง (2003) จำกัด';
  const THAI_MONTHS_SHORT = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  const MIN_ZONE_ROWS = 45;
  const OP_ROWS_PER_SHEET = 30;

  // ---------------------------------------------------------------- helpers
  function parseIso(iso) {
    const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
    return m ? { y: Number(m[1]), m: Number(m[2]), d: Number(m[3]) } : null;
  }

  /** YYYY-MM-DD → ว/ด/ป (พ.ศ.) เช่น 1/8/2569 */
  function toBE(iso) {
    const p = parseIso(iso);
    return p ? `${p.d}/${p.m}/${p.y + 543}` : '';
  }

  function todayIso() {
    const n = new Date();
    return `${n.getFullYear()}-${String(n.getMonth() + 1).padStart(2, '0')}-${String(n.getDate()).padStart(2, '0')}`;
  }

  /** "06:30" → "6:30" (รูปแบบเดียวกับเอกสารต้นฉบับ) */
  function showTime(t) {
    const m = String(t || '').match(/^0?(\d{1,2}):(\d{2})/);
    return m ? `${m[1]}:${m[2]}` : '';
  }

  const timeKey = (l) => String(l.dep_time || l.arr_time || '');

  function sortChrono(logs) {
    return [...(logs || [])].sort((a, b) => {
      if (a.trip_date !== b.trip_date) return String(a.trip_date).localeCompare(String(b.trip_date));
      return timeKey(a).localeCompare(timeKey(b));
    });
  }

  function groupBy(items, keyFn) {
    const map = new Map();
    for (const it of items) {
      const k = keyFn(it);
      if (!map.has(k)) map.set(k, []);
      map.get(k).push(it);
    }
    return map;
  }

  function dateRangeOf(logs, from, to) {
    const dates = (logs || []).map((l) => l.trip_date).filter(Boolean).sort();
    const start = from || dates[0] || todayIso();
    const end = to || dates[dates.length - 1] || start;
    return { start, end };
  }

  function distinctSites(logs) {
    return [...new Set((logs || []).map((l) => String(l.site_name || '').trim()).filter(Boolean))];
  }

  function cleanSheetTitle(raw, fallback) {
    return (
      String(raw || fallback || 'Sheet')
        .replace(/[\\/?*[\]:]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim() ||
      fallback ||
      'Sheet'
    );
  }

  function createSheetNamer() {
    const used = new Set();
    return function name(raw, fallback) {
      const base = cleanSheetTitle(raw, fallback);
      let candidate = base.slice(0, 31).trim() || fallback || 'Sheet';
      let n = 2;
      while (used.has(candidate.toLowerCase())) {
        const suffix = ` (${n})`;
        candidate = base.slice(0, 31 - suffix.length).trim() + suffix;
        n += 1;
      }
      used.add(candidate.toLowerCase());
      return candidate;
    };
  }

  const THIN = { style: 'thin', color: { argb: 'FF000000' } };
  const BORDER_ALL = { top: THIN, left: THIN, bottom: THIN, right: THIN };

  function pageSetup(ws) {
    ws.pageSetup = {
      paperSize: 9, // A4
      orientation: 'portrait',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 1,
      horizontalCentered: true,
      margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 }
    };
  }

  function style(cell, { size = 12, bold = false, h = 'center', v = 'middle', border = false, wrap = false, numFmt } = {}) {
    cell.font = { name: 'Arial', size, bold };
    cell.alignment = { horizontal: h, vertical: v, wrapText: wrap };
    if (border) cell.border = BORDER_ALL;
    if (numFmt) cell.numFmt = numFmt;
  }

  function borderBox(ws, r1, c1, r2, c2) {
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) ws.getCell(r, c).border = BORDER_ALL;
  }

  function newWorkbook(ExcelJS) {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'CKP Transport';
    wb.created = new Date();
    return wb;
  }

  // ============================================================ 1) สำเนาโซน
  function addZoneSheet(wb, sheetName, rows, kind, projectLabel) {
    const ws = wb.addWorksheet(sheetName);
    pageSetup(ws);
    [9, 12, 13, 26, 9, 14, 11, 18].forEach((w, i) => {
      ws.getColumn(i + 1).width = w;
    });

    const titles = [
      ['A1:H1', COMPANY, 25, 38],
      ['A2:H2', 'บันทึกรายการรถ 6 ล้อ รับ-ส่งคนงาน', 18, 28],
      ['A3:H3', `โครงการ  ${projectLabel || ''}`, 15, 24]
    ];
    titles.forEach(([range, text, size, height], i) => {
      ws.mergeCells(range);
      const cell = ws.getCell(range.split(':')[0]);
      cell.value = text;
      style(cell, { size, bold: true });
      ws.getRow(i + 1).height = height;
    });

    const headers =
      kind === 'charter'
        ? ['ลำดับ', 'วัน เดือน ปี', 'ทะเบียนรถ', 'จาก', 'เวลา', 'ถึง', 'ราคา', 'หมายเหตุ']
        : ['ลำดับ', 'วัน เดือน ปี', 'ทะเบียนรถ', 'จาก', 'เวลา', 'จำนวนคน', 'ถึง', 'หมายเหตุ'];
    headers.forEach((text, i) => {
      const cell = ws.getCell(5, i + 1);
      cell.value = text;
      style(cell, { size: 12, border: true });
    });
    ws.getRow(5).height = 22;

    const total = Math.max(MIN_ZONE_ROWS, rows.length);
    let prevDate = '';
    for (let i = 0; i < total; i++) {
      const r = 6 + i;
      const log = rows[i];
      ws.getRow(r).height = 20;
      const set = (col, value, opts) => {
        const cell = ws.getCell(r, col);
        if (value !== undefined && value !== null && value !== '') cell.value = value;
        style(cell, { border: true, ...opts });
      };
      set(1, i + 1);
      if (!log) {
        for (let c = 2; c <= 8; c++) set(c, '');
        continue;
      }
      const showDate = log.trip_date !== prevDate;
      prevDate = log.trip_date;
      set(2, showDate ? toBE(log.trip_date) : '');
      if (kind === 'charter') {
        set(3, log.license_plate);
        set(4, log.origin, { h: 'left' });
        set(5, showTime(log.dep_time));
        set(6, log.destination);
        set(7, Number(log.unit_price) || 0, { numFmt: '#,##0.00' });
        set(8, log.remark, { h: 'left' });
      } else {
        const col3 = log.display_mode_zone === 'license_plate' ? log.license_plate : log.worker_name;
        set(3, col3);
        set(4, log.origin, { h: 'left' });
        set(5, showTime(log.dep_time));
        set(6, log.passengers === null || log.passengers === undefined || log.passengers === '' ? '' : Number(log.passengers));
        set(7, '');
        set(8, log.remark, { h: 'left' });
      }
    }
    return ws;
  }

  function buildZoneCopyWorkbook(ExcelJS, logs, opts = {}) {
    const wb = newWorkbook(ExcelJS);
    const namer = createSheetNamer();
    const sorted = sortChrono(logs);
    const normal = sorted.filter((l) => l.trip_type !== 'charter');
    const charter = sorted.filter((l) => l.trip_type === 'charter');
    const projectOf = (rows) => opts.siteName || distinctSites(rows).join(', ');

    for (const [key, rows] of groupBy(normal, (l) => l.route_sheet_name || `${l.origin} - ${l.destination}`)) {
      addZoneSheet(wb, namer(key, 'สายวิ่ง'), rows, 'normal', projectOf(rows));
    }
    if (charter.length > 0) {
      addZoneSheet(wb, namer('งานเหมา'), charter, 'charter', projectOf(charter));
    }
    if (wb.worksheets.length === 0) {
      addZoneSheet(wb, namer('สำเนา'), [], 'normal', opts.siteName || '');
    }
    return wb;
  }

  // ============================================================ 2) ใบเบิกโครงการ
  function periodLabel(start, end) {
    const a = parseIso(start);
    const b = parseIso(end);
    if (!a || !b) return '';
    if (start === end) return toBE(start);
    if (a.y === b.y && a.m === b.m) return `${a.d}-${b.d}/${a.m}/${a.y + 543}`;
    return `${toBE(start)}-${toBE(end)}`;
  }

  function billingGroups(rows) {
    const groups = groupBy(rows, (l) => `${l.billing_title || l.route_sheet_name}||${Number(l.unit_price) || 0}`);
    return [...groups.values()].map((items) => ({
      title: items[0].billing_title || items[0].route_sheet_name || `${items[0].origin} - ${items[0].destination}`,
      price: Number(items[0].unit_price) || 0,
      count: items.length,
      remark: [...new Set(items.map((l) => String(l.remark || '').trim()).filter(Boolean))].join(', ')
    }));
  }

  function addBillingSummarySheet(wb, normal, charter, ctx) {
    const ws = wb.addWorksheet('ใบเบิก');
    pageSetup(ws);
    [8, 14, 14, 14, 10, 10, 12, 13, 13, 20].forEach((w, i) => {
      ws.getColumn(i + 1).width = w;
    });

    ws.mergeCells('A1:J1');
    ws.getCell('A1').value = 'ใบเบิกผลงานผู้รับเหมา';
    style(ws.getCell('A1'), { size: 20, bold: true });
    ws.getRow(1).height = 32;

    ws.getCell('A2').value = 'ผู้รับเหมาชุด    ';
    ws.mergeCells('C2:F2');
    ws.getCell('C2').value = 'บริษัท จักรพงษ์ ขนส่ง ( 2003 ) จำกัด';
    ws.mergeCells('H2:I2');
    ws.getCell('H2').value = 'วันที่';
    ws.getCell('J2').value = toBE(ctx.issueDate);

    ws.getCell('A3').value = 'โครงการ';
    ws.mergeCells('C3:F3');
    ws.getCell('C3').value = ctx.siteName;
    ws.mergeCells('H3:I3');
    ws.getCell('H3').value = 'ประจำงวด';
    ws.getCell('J3').value = periodLabel(ctx.start, ctx.end);
    for (const addr of ['A2', 'C2', 'H2', 'J2', 'A3', 'C3', 'H3', 'J3']) {
      style(ws.getCell(addr), { size: 13, bold: true, h: addr.startsWith('A') ? 'left' : 'center' });
    }

    // หัวตาราง
    ws.mergeCells('A5:A6');
    ws.mergeCells('B5:D6');
    ws.mergeCells('E5:G5');
    ws.mergeCells('H5:I6');
    ws.mergeCells('J5:J6');
    const heads = { A5: 'ลำดับ', B5: 'รายละเอียดงาน', E5: 'ปริมาณงาน', H5: 'จำนวนเงิน', J5: 'หมายเหตุ', E6: 'จำนวน', F6: 'หน่วย', G6: 'ราคา/หน่วย' };
    for (const [addr, text] of Object.entries(heads)) ws.getCell(addr).value = text;
    borderBox(ws, 5, 1, 6, 10);
    for (let r = 5; r <= 6; r++) for (let c = 1; c <= 10; c++) style(ws.getCell(r, c), { size: 12, bold: true, border: true });

    // รายการ: แถวคู่ 8,10,12,... เว้น 1 แถวคั่น
    const items = [
      ...billingGroups(normal).map((g) => ({ ...g, charter: false })),
      ...billingGroups(charter).map((g) => ({ ...g, charter: true }))
    ];
    let row = 8;
    let seq = 1;
    let grand = 0;
    const placed = [];
    for (const it of items) {
      placed.push({ ...it, row, seq });
      if (it.charter) ws.getCell(row + 1, 2).value = '( ขนย้าย )';
      row += 2;
      seq += 1;
      grand += it.count * it.price;
    }
    const tableEnd = Math.max(28, row - 1);
    const totalRow = tableEnd + 1;

    for (let r = 7; r <= totalRow; r++) {
      ws.mergeCells(`B${r}:D${r}`);
      ws.mergeCells(`H${r}:I${r}`);
    }
    borderBox(ws, 7, 1, totalRow, 10);
    for (let r = 7; r <= totalRow; r++) {
      ws.getRow(r).height = 20;
      for (let c = 1; c <= 10; c++) {
        style(ws.getCell(r, c), { size: 12, border: true, h: c >= 2 && c <= 4 ? 'left' : 'center' });
      }
    }
    ws.getCell('B7').value = 'ค่ารถรับ -  ส่ง   ';

    for (const it of placed) {
      const r = it.row;
      ws.getCell(r, 1).value = it.seq;
      ws.getCell(r, 2).value = it.title;
      ws.getCell(r, 5).value = it.count;
      ws.getCell(r, 6).value = 'เที่ยว';
      ws.getCell(r, 7).value = it.price;
      ws.getCell(r, 7).numFmt = '#,##0.00';
      ws.getCell(r, 8).value = { formula: `SUM(E${r}*G${r})`, result: it.count * it.price };
      ws.getCell(r, 8).numFmt = '#,##0.00';
      if (it.remark) ws.getCell(r, 10).value = it.remark;
    }

    // รวม
    ws.getCell(totalRow, 2).value = 'รวม';
    style(ws.getCell(totalRow, 2), { size: 12, bold: true, border: true });
    ws.getCell(totalRow, 8).value = { formula: `SUM(H8:I${tableEnd})`, result: grand };
    style(ws.getCell(totalRow, 8), { size: 12, bold: true, border: true, numFmt: '#,##0.00' });

    // ลายเซ็น
    const sig1 = totalRow + 3;
    const sig2 = totalRow + 4;
    ws.getCell(sig1, 5).value = 'ผู้ตรวจรับ / Checked by         ';
    ws.getCell(sig2, 5).value = 'วันที่ / Date                            ';
    for (const r of [sig1, sig2]) {
      ws.mergeCells(`G${r}:I${r}`);
      style(ws.getCell(r, 5), { size: 12, h: 'left' });
      for (let c = 7; c <= 9; c++) ws.getCell(r, c).border = { bottom: THIN };
    }
    ws.getRow(sig2 + 1).height = 15;
    return ws;
  }

  function addOperationSheets(wb, namer, baseName, rows, startPageNo, ctx) {
    const chunks = [];
    for (let i = 0; i < rows.length; i += OP_ROWS_PER_SHEET) chunks.push(rows.slice(i, i + OP_ROWS_PER_SHEET));
    if (chunks.length === 0) chunks.push([]);

    const cleanBase = cleanSheetTitle(baseName, 'ใบบันทึก');
    let pageNo = startPageNo;
    chunks.forEach((chunk, k) => {
      const suffix = k === 0 ? '' : ` (${k + 1})`;
      const raw = k === 0 ? cleanBase : `${cleanBase.slice(0, 31 - suffix.length).trim()}${suffix}`;
      const ws = wb.addWorksheet(namer(raw, 'ใบบันทึก'));
      pageSetup(ws);
      [7, 12, 13, 22, 9, 10, 8, 8, 9, 9, 18, 20].forEach((w, i) => {
        ws.getColumn(i + 1).width = w;
      });

      ws.mergeCells('A1:K1');
      ws.getCell('A1').value = 'ใบบันทึกการปฏิบัติงานรถ 6 ล้อ รับ - ส่งพนักงาน';
      style(ws.getCell('A1'), { size: 16, bold: true });
      ws.getRow(1).height = 28;

      ws.getCell('J2').value = 'แผ่นที่';
      ws.mergeCells('K2:L2');
      ws.getCell('K2').value = pageNo;

      ws.getCell('A3').value = 'สถานที่ปฏิบัติงานโครงการ :-';
      ws.mergeCells('D3:E3');
      ws.getCell('D3').value = ctx.siteName;
      ws.mergeCells('F3:G3');
      ws.getCell('F3').value = 'ระหว่างวันที่';
      ws.mergeCells('H3:I3');
      ws.getCell('H3').value = toBE(ctx.start);
      ws.getCell('J3').value = 'ถึงวันที่ ';
      ws.mergeCells('K3:L3');
      ws.getCell('K3').value = toBE(ctx.end);
      for (const addr of ['J2', 'K2', 'A3', 'D3', 'F3', 'H3', 'J3', 'K3']) {
        style(ws.getCell(addr), { size: 11, bold: true, h: addr === 'A3' ? 'left' : 'center' });
      }

      // หัวตาราง 2 ชั้น
      ws.mergeCells('A5:A6');
      ws.mergeCells('B5:B6');
      ws.mergeCells('C5:C6');
      ws.mergeCells('D5:D6');
      ws.mergeCells('F5:H6');
      ws.mergeCells('L5:L6');
      const heads = {
        A5: 'ลำดับ', B5: 'ว/ด/ป', C5: 'ทะเบียนรถ', D5: 'จาก', E5: 'เวลา', E6: '( น. )', F5: 'ถึง',
        I5: 'เวลา', I6: '( น. )', J5: 'จำนวน', J6: '(คน)', K5: 'ลงชื่อผู้ตรวจ', K6: 'แคมป์ Boss /รปภ.', L5: 'หมายเหตุ'
      };
      for (const [addr, text] of Object.entries(heads)) ws.getCell(addr).value = text;
      for (let r = 5; r <= 6; r++) for (let c = 1; c <= 12; c++) style(ws.getCell(r, c), { size: 11, bold: true, border: true });

      // ตารางข้อมูล 30 แถว (แถว 7-36) — ใส่เลขลำดับ/ตีเส้นครบ 30 แถวเสมอ
      let prevDate = '';
      for (let i = 0; i < OP_ROWS_PER_SHEET; i++) {
        const r = 7 + i;
        ws.mergeCells(`F${r}:H${r}`);
        ws.getRow(r).height = 20;
        for (let c = 1; c <= 12; c++) style(ws.getCell(r, c), { size: 11, border: true, h: c === 4 || c === 12 ? 'left' : 'center' });
        ws.getCell(r, 1).value = k * OP_ROWS_PER_SHEET + i + 1;
        const log = chunk[i];
        if (!log) continue;
        const showDate = log.trip_date !== prevDate;
        prevDate = log.trip_date;
        if (showDate) ws.getCell(r, 2).value = toBE(log.trip_date);
        if (log.license_plate) ws.getCell(r, 3).value = log.license_plate;
        if (log.origin) ws.getCell(r, 4).value = log.origin;
        if (log.dep_time) ws.getCell(r, 5).value = showTime(log.dep_time);
        if (log.destination) ws.getCell(r, 6).value = log.destination;
        if (log.arr_time) ws.getCell(r, 9).value = showTime(log.arr_time);
        if (log.passengers !== null && log.passengers !== undefined && log.passengers !== '') {
          ws.getCell(r, 10).value = Number(log.passengers);
        }
        if (log.remark) ws.getCell(r, 12).value = log.remark;
      }

      // ส่วนท้ายสำหรับเซ็นรับรอง
      ws.getCell('H37').value = 'ผู้ตรวจ / Checked by';
      ws.mergeCells('K37:L37');
      ws.getCell('K37').value = '…........................................................................';
      ws.getCell('H38').value = 'วันที่ / Date';
      ws.mergeCells('K38:L38');
      ws.getCell('K38').value = '…............../…..................../…............';
      for (const addr of ['H37', 'K37', 'H38', 'K38']) style(ws.getCell(addr), { size: 11, h: addr.startsWith('H') ? 'right' : 'left' });

      pageNo += 1;
    });
    return pageNo;
  }

  function buildBillingWorkbook(ExcelJS, logs, opts = {}) {
    const wb = newWorkbook(ExcelJS);
    const sorted = sortChrono(logs);
    const normal = sorted.filter((l) => l.trip_type !== 'charter');
    const charter = sorted.filter((l) => l.trip_type === 'charter');
    const { start, end } = dateRangeOf(sorted, opts.dateFrom, opts.dateTo);
    const ctx = {
      siteName: opts.siteName || distinctSites(sorted).join(', '),
      start,
      end,
      issueDate: opts.issueDate || todayIso()
    };

    addBillingSummarySheet(wb, normal, charter, ctx);

    const namer = createSheetNamer();
    namer('ใบเบิก'); // จองชื่อชีตแรกไว้
    let pageNo = 1;
    for (const [key, rows] of groupBy(normal, (l) => l.route_sheet_name || `${l.origin} - ${l.destination}`)) {
      pageNo = addOperationSheets(wb, namer, key, rows, pageNo, ctx);
    }
    if (charter.length > 0) {
      pageNo = addOperationSheets(wb, namer, 'เหมา', charter, pageNo, ctx);
    }
    if (wb.worksheets.length === 1) {
      addOperationSheets(wb, namer, 'แผ่นที่ 1', [], pageNo, ctx);
    }
    return wb;
  }

  // ============================================================ ชื่อไฟล์
  function thaiRange(start, end) {
    const a = parseIso(start);
    const b = parseIso(end);
    if (!a || !b) return '';
    const yy = (p) => String((p.y + 543) % 100).padStart(2, '0');
    if (start === end) return `${a.d} ${THAI_MONTHS_SHORT[a.m - 1]} ${yy(a)}`;
    if (a.y === b.y && a.m === b.m) return `${a.d} - ${b.d} ${THAI_MONTHS_SHORT[a.m - 1]} ${yy(a)}`;
    return `${a.d} ${THAI_MONTHS_SHORT[a.m - 1]} ${yy(a)} - ${b.d} ${THAI_MONTHS_SHORT[b.m - 1]} ${yy(b)}`;
  }

  const safeFile = (s) => String(s).replace(/[\\/:*?"<>|]/g, ' ').replace(/\s+/g, ' ').trim();

  function zoneCopyFilename({ zoneName, siteName, start, end }) {
    const label = zoneName || siteName || '';
    return safeFile(`สำเนา ${label} ( ${thaiRange(start, end)} )`.replace(/\s+/g, ' ')) + '.xlsx';
  }

  function billingFilename({ siteName, start, end }) {
    const a = parseIso(start);
    const b = parseIso(end);
    let range = thaiRange(start, end);
    if (a && b && a.y === b.y && a.m === b.m) range = start === end ? `${a.d}-${a.m}` : `${a.d}-${b.d}-${a.m}`;
    return safeFile(`ใบเบิก${siteName || ''} ( ${range} )`) + '.xlsx';
  }

  return {
    buildZoneCopyWorkbook,
    buildBillingWorkbook,
    zoneCopyFilename,
    billingFilename,
    toBE,
    thaiRange,
    periodLabel,
    dateRangeOf
  };
});
