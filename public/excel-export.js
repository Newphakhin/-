/**
 * ตัวสร้างไฟล์ Excel (.xlsx) ด้วย ExcelJS — จัดรูปแบบเหมือนไฟล์ต้นฉบับบริษัท 100%
 *   1) สำเนาโซน      : buildZoneCopyWorkbook(ExcelJS, logs, opts)
 *      อ้างอิงโครงสร้าง ฟอนต์ สีไฮไลท์ คอลัมน์ ความสูงแถว สูตรสรุปเที่ยว (COUNTIF/SUM) จากต้นฉบับสำเนาโซน
 *   2) ใบเบิกโครงการ : buildBillingWorkbook(ExcelJS, logs, opts)
 *      อ้างอิงโครงสร้าง ฟอนต์ Angsana New / Tahoma / Arial ขนาด สัดส่วนคอลัมน์ ความสูงแถว เส้นประ สูตรคำนวณ และการตัดหน้า 30 แถว จากต้นฉบับใบเบิก
 * ไม่มีข้อมูลตัวอย่างในโค้ดนี้ — ทุกแถวมาจากบันทึกการเดินทาง (travel_logs) ที่ส่งเข้ามาเท่านั้น
 */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.CkpExcel = api;
})(typeof window !== 'undefined' ? window : globalThis, function () {
  'use strict';

  const COMPANY_ZONE = 'บริษัท จักรพงษ์ ขนส่ง (2003) จำกัด';
  const COMPANY_BILLING = 'บริษัท จักรพงษ์ ขนส่ง ( 2003 ) จำกัด';
  const THAI_MONTHS_SHORT = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  const MIN_ZONE_ROWS = 45;
  const OP_ROWS_PER_SHEET = 30;
  const ACCOUNTING_FMT = '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@';

  // สีตามไฟล์ต้นฉบับสำเนาโซน
  const COLOR_BLACK = { argb: 'FF000000' };
  const COLOR_RED = { argb: 'FFFF0000' };
  const FILL_YELLOW = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
  const FILL_GREEN = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF92D050' } };
  const FILL_ROSE = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD99694' } };
  const CHARTER_GROUP_FILLS = [
    { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF92D050' } }, // เขียว
    { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCC0DA' } }, // ม่วงอ่อน
    { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB8CCE4' } }, // ฟ้าอ่อน
    { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6B8B7' } }, // ชมพูอ่อน
    { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }  // เหลือง
  ];

  const THIN = { style: 'thin', color: COLOR_BLACK };
  const DOTTED = { style: 'dotted', color: COLOR_BLACK };
  const BORDER_ALL = { top: THIN, left: THIN, bottom: THIN, right: THIN };

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

  function applyStyle(
    cell,
    {
      fontName = 'Arial',
      size = 12,
      bold = false,
      color = COLOR_BLACK,
      h = 'center',
      v = 'middle',
      border = null,
      fill = null,
      wrap = false,
      numFmt = null
    } = {}
  ) {
    cell.font = { name: fontName, size, bold, color };
    cell.alignment = { horizontal: h, vertical: v, wrapText: wrap };
    if (border) cell.border = border;
    if (fill) cell.fill = fill;
    if (numFmt) cell.numFmt = numFmt;
  }

  function borderBox(ws, r1, c1, r2, c2, border = BORDER_ALL) {
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        ws.getCell(r, c).border = border;
      }
    }
  }

  function newWorkbook(ExcelJS) {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'CKP Transport';
    wb.created = new Date();
    return wb;
  }

  // ============================================================================
  // 1) ไฟล์สำเนาโซน — เหมือนต้นฉบับ "สำเนา27 โซนวิภา" 100%
  // ============================================================================
  function addZoneSheet(wb, sheetName, rows, kind) {
    const ws = wb.addWorksheet(sheetName);
    ws.properties.defaultRowHeight = 14.5;
    ws.views = [{ state: 'normal', zoomScale: 100, zoomScaleNormal: 100, showGridLines: true }];
    ws.pageSetup = {
      paperSize: 9, // A4
      orientation: 'portrait',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 1,
      margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 }
    };

    const colWidths =
      kind === 'charter'
        ? [10, 13.45, 13.45, 37.63, 8.73, 37.54, 13.27, 44]
        : [8.73, 11.18, 11.91, 25.36, 8.27, 13.45, 10.73, 14.63];
    colWidths.forEach((w, i) => {
      ws.getColumn(i + 1).width = w;
    });

    // หัวกระดาษ 3 แถว (แถว 3 "โครงการ  " ตัวแดงเหมือนต้นฉบับ)
    ws.mergeCells('A1:H1');
    ws.getCell('A1').value = COMPANY_ZONE;
    applyStyle(ws.getCell('A1'), { fontName: 'Arial', size: 25, bold: true, color: COLOR_BLACK, h: 'center', v: 'bottom' });
    ws.getRow(1).height = 31.5;

    ws.mergeCells('A2:H2');
    ws.getCell('A2').value = 'บันทึกรายการรถ 6 ล้อ รับ-ส่งคนงาน';
    applyStyle(ws.getCell('A2'), { fontName: 'Arial', size: 18, bold: true, color: COLOR_BLACK, h: 'center', v: 'bottom' });
    ws.getRow(2).height = 23;

    ws.mergeCells('A3:H3');
    ws.getCell('A3').value = 'โครงการ  ';
    applyStyle(ws.getCell('A3'), { fontName: 'Arial', size: 15, bold: true, color: COLOR_RED, h: 'center', v: 'bottom' });
    ws.getRow(3).height = 19;

    ws.getRow(4).height = 19;

    // แถวที่ 5 หัวตาราง (Arial 12pt ไม่หนา กึ่งกลาง ตีกรอบ 4 ด้าน)
    const headers =
      kind === 'charter'
        ? ['ลำดับ', 'วัน เดือน ปี', 'ทะเบียนรถ', 'จาก', 'เวลา', 'ถึง', 'ราคา', 'หมายเหตุ']
        : ['ลำดับ', 'วัน เดือน ปี', 'ทะเบียนรถ', 'จาก', 'เวลา', 'จำนวนคน', 'ถึง', 'หมายเหตุ'];
    headers.forEach((text, i) => {
      const cell = ws.getCell(5, i + 1);
      cell.value = text;
      applyStyle(cell, {
        fontName: 'Arial',
        size: 12,
        bold: false,
        color: COLOR_BLACK,
        h: 'center',
        v: i === 5 || (kind === 'charter' && i === 7) ? 'middle' : 'bottom',
        border: BORDER_ALL
      });
    });
    ws.getRow(5).height = 15.5;

    // ข้อมูลและตารางสรุปจำนวนเที่ยวรายคน/รายทะเบียน (เฉพาะสายวิ่งปกติ)
    if (kind === 'charter') {
      const totalRows = Math.max(MIN_ZONE_ROWS, rows.length);
      let prevDate = '';
      let prevGroupKey = '';
      let groupColorIdx = -1;

      for (let i = 0; i < totalRows; i++) {
        const r = 6 + i;
        const log = rows[i];

        const aCell = ws.getCell(r, 1);
        aCell.value = i + 1;
        applyStyle(aCell, { fontName: 'Calibri', size: 11, h: 'center', v: 'bottom', border: BORDER_ALL });

        if (!log) {
          for (let c = 2; c <= 8; c++) {
            applyStyle(ws.getCell(r, c), {
              fontName: 'Calibri',
              size: 11,
              h: c === 8 ? 'left' : 'center',
              v: c === 8 ? 'middle' : 'bottom',
              border: BORDER_ALL
            });
          }
          continue;
        }

        const gKey = `${log.trip_date}||${log.origin}||${log.destination}`;
        if (gKey !== prevGroupKey) {
          prevGroupKey = gKey;
          groupColorIdx = (groupColorIdx + 1) % CHARTER_GROUP_FILLS.length;
        }
        const rowFill = CHARTER_GROUP_FILLS[groupColorIdx];

        const showDate = log.trip_date !== prevDate;
        prevDate = log.trip_date;

        const setCharterCell = (c, val, opts = {}) => {
          const cell = ws.getCell(r, c);
          if (val !== undefined && val !== null && val !== '') cell.value = val;
          applyStyle(cell, {
            fontName: 'Calibri',
            size: 11,
            h: 'center',
            v: 'bottom',
            border: BORDER_ALL,
            fill: rowFill,
            ...opts
          });
        };

        setCharterCell(2, showDate ? toBE(log.trip_date) : '', { numFmt: '@' });
        setCharterCell(3, log.license_plate, { numFmt: '@' });
        setCharterCell(4, log.origin);
        setCharterCell(5, showTime(log.dep_time), { numFmt: '@' });
        setCharterCell(6, log.destination);
        setCharterCell(7, Number(log.unit_price) || 0, { v: 'middle', numFmt: '#,##0' });
        setCharterCell(8, log.remark, { h: 'left', v: 'middle' });
      }
      return ws;
    }

    // สายวิ่งปกติ: เตรียมข้อมูลสรุปรายคนขับ/ทะเบียน (คอลัมน์ F = ชื่อ/ทะเบียนพื้นเขียว, คอลัมน์ G = COUNTIF พื้นเหลือง, รวม = SUM พื้นชมพู)
    const primaryOrigin = rows.length > 0 ? String(rows[0].origin || '').trim() : '';
    const lastDataRow = 5 + rows.length;
    const distinctColC = [];
    const colCCounts = new Map();
    for (const log of rows) {
      const col3Val = String(log.display_mode_zone === 'license_plate' ? log.license_plate : log.worker_name || '').trim();
      if (!col3Val) continue;
      if (!colCCounts.has(col3Val)) {
        distinctColC.push(col3Val);
        colCCounts.set(col3Val, 0);
      }
      colCCounts.set(col3Val, colCCounts.get(col3Val) + 1);
    }

    const sumStartRow = rows.length > 0 && distinctColC.length > 0 ? lastDataRow + 3 : -1;
    const sumEndRow = sumStartRow > 0 ? sumStartRow + distinctColC.length - 1 : -1;
    const sumTotalRow = sumEndRow > 0 ? sumEndRow + 1 : -1;
    const totalRows = Math.max(MIN_ZONE_ROWS, sumTotalRow > 0 ? sumTotalRow - 5 + 5 : rows.length);

    let prevDate = '';
    for (let i = 0; i < totalRows; i++) {
      const r = 6 + i;
      const log = rows[i];

      // คอลัมน์ A: ลำดับ (ไม่มีพื้นสี)
      const aCell = ws.getCell(r, 1);
      aCell.value = i + 1;
      applyStyle(aCell, { fontName: 'Calibri', size: 11, h: 'center', v: 'bottom', border: BORDER_ALL });

      if (log) {
        const showDate = log.trip_date !== prevDate;
        prevDate = log.trip_date;
        const col3Val = log.display_mode_zone === 'license_plate' ? log.license_plate : log.worker_name;
        const isReturn = primaryOrigin && String(log.origin || '').trim() !== primaryOrigin;
        const hasPass = log.passengers !== null && log.passengers !== undefined && log.passengers !== '';
        const hasRemark = Boolean(String(log.remark || '').trim());

        const setNormal = (c, val, opts = {}) => {
          const cell = ws.getCell(r, c);
          if (val !== undefined && val !== null && val !== '') cell.value = val;
          applyStyle(cell, {
            fontName: 'Calibri',
            size: 11,
            h: 'center',
            v: 'bottom',
            border: BORDER_ALL,
            ...opts
          });
        };

        setNormal(2, showDate ? toBE(log.trip_date) : '', { fill: FILL_YELLOW, numFmt: '@' });
        setNormal(3, col3Val, { fill: FILL_YELLOW, numFmt: '@' });
        setNormal(4, log.origin, { h: isReturn ? 'right' : 'left', fill: FILL_YELLOW });
        setNormal(5, showTime(log.dep_time), { fill: FILL_YELLOW, numFmt: '@' });
        setNormal(6, hasPass ? Number(log.passengers) : '', { v: 'middle', fill: hasPass ? FILL_YELLOW : null });
        setNormal(7, '');
        setNormal(8, log.remark, { h: 'center', fill: hasRemark ? FILL_YELLOW : null });
      } else {
        // แถวว่างใต้ข้อมูล หรือแถวสรุปจำนวนเที่ยวรายคน/ทะเบียน
        for (let c = 2; c <= 8; c++) {
          applyStyle(ws.getCell(r, c), {
            fontName: 'Calibri',
            size: 11,
            h: c === 4 ? 'right' : 'center',
            v: c === 6 ? 'middle' : 'bottom',
            border: BORDER_ALL
          });
        }

        if (sumStartRow > 0 && r >= sumStartRow && r <= sumEndRow) {
          const label = distinctColC[r - sumStartRow];
          const count = colCCounts.get(label) || 0;
          const fCell = ws.getCell(r, 6);
          fCell.value = label;
          applyStyle(fCell, {
            fontName: 'Calibri',
            size: 11,
            h: 'center',
            v: 'bottom',
            border: BORDER_ALL,
            fill: FILL_GREEN
          });
          const gCell = ws.getCell(r, 7);
          gCell.value = { formula: `COUNTIF(C$6:C$${lastDataRow},F${r})`, result: count };
          applyStyle(gCell, {
            fontName: 'Calibri',
            size: 11,
            h: 'center',
            v: 'bottom',
            border: BORDER_ALL,
            fill: FILL_YELLOW
          });
        } else if (sumTotalRow > 0 && r === sumTotalRow) {
          const gCell = ws.getCell(r, 7);
          gCell.value = { formula: `SUM(G${sumStartRow}:G${sumEndRow})`, result: rows.length };
          applyStyle(gCell, {
            fontName: 'Calibri',
            size: 11,
            h: 'center',
            v: 'bottom',
            border: BORDER_ALL,
            fill: FILL_ROSE
          });
        }
      }
    }

    return ws;
  }

  function buildZoneCopyWorkbook(ExcelJS, logs) {
    const wb = newWorkbook(ExcelJS);
    const namer = createSheetNamer();
    const sorted = sortChrono(logs);
    const normal = sorted.filter((l) => l.trip_type !== 'charter');
    const charter = sorted.filter((l) => l.trip_type === 'charter');

    for (const [key, rows] of groupBy(normal, (l) => l.route_sheet_name || `${l.origin} - ${l.destination}`)) {
      addZoneSheet(wb, namer(key, 'สายวิ่ง'), rows, 'normal');
    }
    if (charter.length > 0) {
      addZoneSheet(wb, namer('งานเหมา'), charter, 'charter');
    }
    if (wb.worksheets.length === 0) {
      addZoneSheet(wb, namer('สำเนา'), [], 'normal');
    }
    return wb;
  }

  // ============================================================================
  // 2) ไฟล์ใบเบิกโครงการ — เหมือนต้นฉบับ "ใบเบิกไลฟ์รัชดา-พระราม9" 100%
  // ============================================================================
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
    ws.properties.defaultRowHeight = 14.5;
    ws.views = [{ state: 'normal', zoomScale: 20, zoomScaleNormal: 20, showGridLines: true }];
    ws.pageSetup = {
      paperSize: 9, // A4
      orientation: 'portrait',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 1,
      margins: { left: 0.39, right: 0.47, top: 0.42, bottom: 0.43, header: 0.3, footer: 0.3 }
    };

    // ความกว้างคอลัมน์ตรงต้นฉบับ 100%
    [30.91, 72.36, 60.27, 72.09, 43.36, 45.63, 90, 60, 72.36, 126.91].forEach((w, i) => {
      ws.getColumn(i + 1).width = w;
    });

    // แถวที่ 1: หัวเอกสาร
    ws.mergeCells('A1:J1');
    ws.getCell('A1').value = 'ใบเบิกผลงานผู้รับเหมา';
    applyStyle(ws.getCell('A1'), { fontName: 'Tahoma', size: 65, bold: false, h: 'center', v: 'bottom' });
    ws.getRow(1).height = 141;

    // แถวที่ 2: ผู้รับเหมาชุด + วันที่ (J2 ขีดเส้นประใต้)
    ws.getRow(2).height = 203;
    ws.getCell('A2').value = 'ผู้รับเหมาชุด    ';
    applyStyle(ws.getCell('A2'), { fontName: 'Angsana New', size: 48, bold: true, h: 'left', v: 'bottom' });
    ws.mergeCells('C2:F2');
    ws.getCell('C2').value = COMPANY_BILLING;
    applyStyle(ws.getCell('C2'), { fontName: 'Angsana New', size: 72, bold: true, h: 'center', v: 'bottom' });
    ws.mergeCells('H2:I2');
    ws.getCell('H2').value = 'วันที่';
    applyStyle(ws.getCell('H2'), { fontName: 'Angsana New', size: 55, bold: true, h: 'left', v: 'bottom' });
    ws.getCell('J2').value = toBE(ctx.issueDate);
    applyStyle(ws.getCell('J2'), {
      fontName: 'Tahoma',
      size: 65,
      bold: false,
      h: 'center',
      v: 'bottom',
      border: { bottom: DOTTED },
      numFmt: '@'
    });

    // แถวที่ 3: โครงการ + ประจำงวด (C3:F3 และ J3 ขีดเส้นประใต้)
    ws.getRow(3).height = 128.5;
    ws.getCell('A3').value = 'โครงการ';
    applyStyle(ws.getCell('A3'), { fontName: 'Angsana New', size: 48, bold: true, h: 'left', v: 'bottom' });
    ws.mergeCells('C3:F3');
    ws.getCell('C3').value = ctx.siteName;
    applyStyle(ws.getCell('C3'), { fontName: 'Angsana New', size: 65, bold: true, h: 'center', v: 'bottom' });
    for (let c = 3; c <= 6; c++) ws.getCell(3, c).border = { bottom: DOTTED };
    ws.mergeCells('H3:I3');
    ws.getCell('H3').value = 'ประจำงวด';
    applyStyle(ws.getCell('H3'), { fontName: 'Angsana New', size: 55, bold: true, h: 'left', v: 'bottom' });
    ws.getCell('J3').value = periodLabel(ctx.start, ctx.end);
    applyStyle(ws.getCell('J3'), {
      fontName: 'Tahoma',
      size: 59,
      bold: false,
      h: 'center',
      v: 'bottom',
      border: { bottom: DOTTED },
      numFmt: '@'
    });

    // แถวที่ 4: เส้นขอบบนของหัวตาราง
    ws.getRow(4).height = 45.5;
    for (let c = 1; c <= 10; c++) ws.getCell(4, c).border = { bottom: THIN };

    // แถวที่ 5-6: หัวตาราง 2 ชั้น (Angsana New 72pt Bold)
    ws.getRow(5).height = 114.5;
    ws.getRow(6).height = 133.5;
    ws.mergeCells('A5:A6');
    ws.mergeCells('B5:D6');
    ws.mergeCells('E5:G5');
    ws.mergeCells('H5:I6');
    ws.mergeCells('J5:J6');

    const heads = {
      A5: 'ลำดับ',
      B5: 'รายละเอียดงาน',
      E5: 'ปริมาณงาน',
      H5: 'จำนวนเงิน',
      J5: 'หมายเหตุ',
      E6: 'จำนวน',
      F6: 'หน่วย',
      G6: 'ราคา/หน่วย'
    };
    borderBox(ws, 5, 1, 6, 10, BORDER_ALL);
    for (const [addr, text] of Object.entries(heads)) {
      const cell = ws.getCell(addr);
      cell.value = text;
      applyStyle(cell, {
        fontName: 'Angsana New',
        size: 72,
        bold: true,
        h: 'center',
        v: 'middle',
        border: BORDER_ALL
      });
    }

    // เตรียมรายการแถวคู่ 8, 10, 12... เว้น 1 แถวคั่น
    const items = [
      ...billingGroups(normal).map((g) => ({ ...g, charter: false })),
      ...billingGroups(charter).map((g) => ({ ...g, charter: true }))
    ];
    let nextRow = 8;
    let seq = 1;
    let grand = 0;
    const placed = [];
    for (const it of items) {
      placed.push({ ...it, row: nextRow, seq });
      nextRow += 2;
      seq += 1;
      grand += it.count * it.price;
    }

    const tableEnd = Math.max(28, nextRow - 1);
    const totalRow = tableEnd + 1;

    // ผสานเซลล์และตีกรอบทุกแถวตั้งแต่ 7 ถึง totalRow (ปกติ 7..29)
    for (let r = 7; r <= totalRow; r++) {
      ws.mergeCells(`B${r}:D${r}`);
      ws.mergeCells(`H${r}:I${r}`);
      ws.getRow(r).height = r === 7 ? 131 : 183.5;
    }
    borderBox(ws, 7, 1, totalRow, 10, BORDER_ALL);

    // แถวที่ 7: หัวข้อย่อย "ค่ารถรับ -  ส่ง   "
    ws.getCell('B7').value = 'ค่ารถรับ -  ส่ง   ';
    applyStyle(ws.getCell('B7'), {
      fontName: 'Angsana New',
      size: 70,
      bold: true,
      h: 'left',
      v: 'bottom',
      border: BORDER_ALL
    });

    // ใส่รายการแต่ละแถวคู่ + "( ขนย้าย )" แถวถัดไปสำหรับงานเหมา
    for (const it of placed) {
      const r = it.row;
      ws.getCell(r, 1).value = it.seq;
      applyStyle(ws.getCell(r, 1), { fontName: 'Angsana New', size: 70, bold: false, h: 'center', v: 'bottom', border: BORDER_ALL });

      ws.getCell(r, 2).value = it.title;
      applyStyle(ws.getCell(r, 2), {
        fontName: it.charter ? 'Arial' : 'Tahoma',
        size: 70,
        bold: false,
        h: 'left',
        v: 'bottom',
        border: BORDER_ALL,
        numFmt: '@'
      });

      ws.getCell(r, 5).value = it.count;
      applyStyle(ws.getCell(r, 5), { fontName: 'Tahoma', size: 71, bold: false, h: 'center', v: 'bottom', border: BORDER_ALL });

      ws.getCell(r, 6).value = 'เที่ยว';
      applyStyle(ws.getCell(r, 6), { fontName: 'Tahoma', size: 71, bold: false, h: 'center', v: 'bottom', border: BORDER_ALL });

      ws.getCell(r, 7).value = it.price;
      applyStyle(ws.getCell(r, 7), {
        fontName: 'Tahoma',
        size: 71,
        bold: false,
        h: 'center',
        v: 'bottom',
        border: BORDER_ALL,
        numFmt: ACCOUNTING_FMT
      });

      ws.getCell(r, 8).value = { formula: `SUM(E${r}*G${r})`, result: it.count * it.price };
      applyStyle(ws.getCell(r, 8), {
        fontName: 'Tahoma',
        size: 71,
        bold: false,
        h: 'center',
        v: 'bottom',
        border: BORDER_ALL,
        numFmt: ACCOUNTING_FMT
      });

      if (it.remark) {
        ws.getCell(r, 10).value = it.remark;
        applyStyle(ws.getCell(r, 10), { fontName: 'Angsana New', size: 48, bold: false, h: 'center', v: 'bottom', border: BORDER_ALL });
      }

      if (it.charter) {
        ws.getCell(r + 1, 2).value = '( ขนย้าย )';
        applyStyle(ws.getCell(r + 1, 2), {
          fontName: 'Arial',
          size: 72,
          bold: true,
          h: 'center',
          v: 'bottom',
          border: BORDER_ALL,
          numFmt: '@'
        });
      }
    }

    // แถวรวม (แถวที่ 29 เมื่อมีไม่เกิน 10 รายการ)
    ws.getCell(totalRow, 2).value = 'รวม';
    applyStyle(ws.getCell(totalRow, 2), {
      fontName: 'Angsana New',
      size: 73,
      bold: true,
      h: 'center',
      v: 'bottom',
      border: BORDER_ALL
    });
    ws.getCell(totalRow, 8).value = { formula: `SUM(H8:I${tableEnd})`, result: grand };
    applyStyle(ws.getCell(totalRow, 8), {
      fontName: 'Tahoma',
      size: 70,
      bold: false,
      h: 'center',
      v: 'bottom',
      border: BORDER_ALL,
      numFmt: ACCOUNTING_FMT
    });

    // แถวเว้นว่าง 30-31
    ws.getRow(totalRow + 1).height = 183.5;
    ws.getRow(totalRow + 2).height = 183.5;

    // แถวลายเซ็น 32-33
    const sig1 = totalRow + 3;
    const sig2 = totalRow + 4;
    ws.getRow(sig1).height = 144;
    ws.getRow(sig2).height = 137;
    ws.getRow(sig2 + 1).height = 45;

    ws.getCell(sig1, 5).value = 'ผู้ตรวจรับ / Checked by         ';
    applyStyle(ws.getCell(sig1, 5), { fontName: 'Angsana New', size: 48, bold: false, h: 'left', v: 'bottom' });
    ws.mergeCells(`G${sig1}:I${sig1}`);
    for (let c = 7; c <= 9; c++) ws.getCell(sig1, c).border = { bottom: THIN };

    ws.getCell(sig2, 5).value = 'วันที่ / Date                            ';
    applyStyle(ws.getCell(sig2, 5), { fontName: 'Angsana New', size: 48, bold: false, h: 'left', v: 'bottom' });
    ws.mergeCells(`G${sig2}:I${sig2}`);
    for (let c = 7; c <= 9; c++) ws.getCell(sig2, c).border = { bottom: THIN };

    ws.mergeCells(`F${sig2 + 1}:H${sig2 + 1}`);
    return ws;
  }

  function addOperationSheets(wb, namer, baseName, rows, startPageNo, ctx) {
    const chunks = [];
    for (let i = 0; i < rows.length; i += OP_ROWS_PER_SHEET) chunks.push(rows.slice(i, i + OP_ROWS_PER_SHEET));
    if (chunks.length === 0) chunks.push([]);

    const cleanBase = cleanSheetTitle(baseName, 'ใบบันทึก');
    let pageNo = startPageNo;
    // ต้นฉบับรัน prevDate ต่อเนื่องข้ามแผ่นต่อ (2), (3) ของสายเดียวกัน และรีเซ็ตเมื่อเริ่มสายใหม่
    let prevDate = '';

    chunks.forEach((chunk, k) => {
      const suffix = k === 0 ? '' : ` (${k + 1})`;
      const raw = k === 0 ? cleanBase : `${cleanBase.slice(0, 31 - suffix.length).trim()}${suffix}`;
      const ws = wb.addWorksheet(namer(raw, 'ใบบันทึก'));

      ws.properties.defaultRowHeight = 14.5;
      ws.views = [{ state: 'normal', zoomScale: 26, zoomScaleNormal: 26, showGridLines: true }];
      ws.pageSetup = {
        paperSize: 9, // A4
        orientation: 'portrait',
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 1,
        margins: { left: 0.7, right: 0.7, top: 0.32, bottom: 0.38, header: 0.18, footer: 0.3 }
      };

      // ความกว้างคอลัมน์ตรงต้นฉบับ 100% (A..L)
      [37.91, 49.36, 56.63, 109, 37.91, 32.18, 36.54, 36.91, 29.63, 39, 39, 70.45].forEach((w, i) => {
        ws.getColumn(i + 1).width = w;
      });

      // แถวที่ 1: หัวกระดาษ (Angsana New 55pt Bold)
      ws.mergeCells('A1:K1');
      ws.getCell('A1').value = 'ใบบันทึกการปฏิบัติงานรถ 6 ล้อ รับ - ส่งพนักงาน';
      applyStyle(ws.getCell('A1'), { fontName: 'Angsana New', size: 55, bold: true, h: 'center', v: 'bottom' });
      ws.getRow(1).height = 78.5;

      // แถวที่ 2: แผ่นที่ (K2:L2 เส้นประใต้)
      ws.getRow(2).height = 86;
      ws.getCell('J2').value = 'แผ่นที่';
      applyStyle(ws.getCell('J2'), { fontName: 'Angsana New', size: 48, bold: true, h: 'center', v: 'bottom' });
      ws.mergeCells('K2:L2');
      ws.getCell('K2').value = pageNo;
      applyStyle(ws.getCell('K2'), { fontName: 'Angsana New', size: 60, bold: true, h: 'center', v: 'bottom' });
      ws.getCell('K2').border = { bottom: DOTTED };
      ws.getCell('L2').border = { bottom: DOTTED };

      // แถวที่ 3: สถานที่ปฏิบัติงานโครงการ + ระหว่างวันที่ + ถึงวันที่
      ws.getRow(3).height = 81.5;
      ws.getCell('A3').value = 'สถานที่ปฏิบัติงานโครงการ :-';
      applyStyle(ws.getCell('A3'), { fontName: 'Angsana New', size: 48, bold: true, h: 'left', v: 'bottom' });

      ws.mergeCells('D3:E3');
      ws.getCell('D3').value = ctx.siteName;
      applyStyle(ws.getCell('D3'), { fontName: 'Tahoma', size: 60, bold: true, h: 'center', v: 'bottom' });
      ws.getCell('D3').border = { bottom: DOTTED };
      ws.getCell('E3').border = { bottom: DOTTED };

      ws.mergeCells('F3:G3');
      ws.getCell('F3').value = 'ระหว่างวันที่';
      applyStyle(ws.getCell('F3'), { fontName: 'Angsana New', size: 48, bold: true, h: 'center', v: 'bottom' });

      ws.mergeCells('H3:I3');
      ws.getCell('H3').value = toBE(ctx.start);
      applyStyle(ws.getCell('H3'), { fontName: 'Arial', size: 55, bold: true, h: 'center', v: 'bottom', numFmt: '@' });

      ws.getCell('J3').value = 'ถึงวันที่ ';
      applyStyle(ws.getCell('J3'), { fontName: 'Angsana New', size: 50, bold: true, h: 'center', v: 'bottom' });

      ws.mergeCells('K3:L3');
      ws.getCell('K3').value = toBE(ctx.end);
      applyStyle(ws.getCell('K3'), { fontName: 'Arial', size: 55, bold: true, h: 'center', v: 'bottom', numFmt: '@' });

      // แถวที่ 4: เว้นว่าง + เส้นบนตาราง
      ws.getRow(4).height = 27.5;

      // แถวที่ 5-6: หัวตาราง 2 ชั้น ตรงตามฟอนต์และขนาดต้นฉบับ 100%
      ws.getRow(5).height = 116.5;
      ws.getRow(6).height = 117.5;
      ws.mergeCells('A5:A6');
      ws.mergeCells('B5:B6');
      ws.mergeCells('C5:C6');
      ws.mergeCells('D5:D6');
      ws.mergeCells('F5:H6');
      ws.mergeCells('L5:L6');
      borderBox(ws, 5, 1, 6, 12, BORDER_ALL);

      const headerSpecs = [
        ['A5', 'ลำดับ', 'Tahoma', 42],
        ['B5', 'ว/ด/ป', 'Tahoma', 42],
        ['C5', 'ทะเบียนรถ', 'Tahoma', 42],
        ['D5', 'จาก', 'Tahoma', 42],
        ['E5', 'เวลา', 'Tahoma', 42],
        ['E6', '( น. )', 'Angsana New', 42],
        ['F5', 'ถึง', 'Tahoma', 42],
        ['I5', 'เวลา', 'Tahoma', 42],
        ['I6', '( น. )', 'Angsana New', 42],
        ['J5', 'จำนวน', 'Calibri', 40],
        ['J6', '(คน)', 'Calibri', 44],
        ['K5', 'ลงชื่อผู้ตรวจ', 'Tahoma', 30],
        ['K6', 'แคมป์ Boss /รปภ.', 'Angsana New', 30],
        ['L5', 'หมายเหตุ', 'Tahoma', 45]
      ];
      for (const [addr, text, fontName, size] of headerSpecs) {
        const cell = ws.getCell(addr);
        cell.value = text;
        applyStyle(cell, { fontName, size, bold: true, h: 'center', v: 'middle', border: BORDER_ALL });
      }

      // แถวที่ 7-36: ตารางข้อมูล 30 แถวพอดี (Arial 42pt Bold กึ่งกลาง ตีกรอบและ Merge F:H ครบทุกแถว)
      for (let i = 0; i < OP_ROWS_PER_SHEET; i++) {
        const r = 7 + i;
        ws.mergeCells(`F${r}:H${r}`);
        ws.getRow(r).height = 124;
        for (let c = 1; c <= 12; c++) {
          applyStyle(ws.getCell(r, c), {
            fontName: 'Arial',
            size: 42,
            bold: true,
            h: 'center',
            v: 'middle',
            border: BORDER_ALL,
            numFmt: c === 2 || c === 3 || c === 5 || c === 9 || c === 10 ? '@' : null
          });
        }
        ws.getCell(r, 1).value = k * OP_ROWS_PER_SHEET + i + 1;

        const log = chunk[i];
        if (!log) continue;

        const showDate = log.trip_date !== prevDate;
        prevDate = log.trip_date;
        if (showDate) ws.getCell(r, 2).value = toBE(log.trip_date);
        if (log.license_plate) ws.getCell(r, 3).value = String(log.license_plate);
        if (log.origin) ws.getCell(r, 4).value = String(log.origin);
        if (log.dep_time) ws.getCell(r, 5).value = showTime(log.dep_time);
        if (log.destination) ws.getCell(r, 6).value = String(log.destination);
        if (log.arr_time) ws.getCell(r, 9).value = showTime(log.arr_time);
        if (log.passengers !== null && log.passengers !== undefined && log.passengers !== '') {
          ws.getCell(r, 10).value = String(Number(log.passengers));
        }
        if (log.remark) ws.getCell(r, 12).value = String(log.remark);
      }

      // แถวที่ 37-38: ส่วนท้ายสำหรับเซ็นรับรอง (ตรงฟอนต์/ขนาดต้นฉบับ 100%)
      ws.getRow(37).height = 124;
      ws.getRow(38).height = 111;
      ws.getRow(39).height = 27.5;

      ws.getCell('H37').value = 'ผู้ตรวจ / Checked by';
      applyStyle(ws.getCell('H37'), { fontName: 'Angsana New', size: 48, bold: true, h: 'left', v: 'bottom' });
      ws.mergeCells('K37:L37');
      ws.getCell('K37').value = '…........................................................................';
      applyStyle(ws.getCell('K37'), { fontName: 'Arial', size: 22, bold: true, h: 'center', v: 'bottom', border: { top: THIN } });
      ws.getCell('L37').border = { top: THIN };

      ws.getCell('H38').value = 'วันที่ / Date';
      applyStyle(ws.getCell('H38'), { fontName: 'Angsana New', size: 48, bold: true, h: 'left', v: 'bottom' });
      ws.mergeCells('K38:L38');
      ws.getCell('K38').value = '…............../…..................../…............';
      applyStyle(ws.getCell('K38'), { fontName: 'Angsana New', size: 48, bold: true, h: 'center', v: 'bottom' });

      ws.mergeCells('I39:K39');

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
    namer('ใบเบิก');
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

  // ============================================================================
  // ชื่อไฟล์
  // ============================================================================
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
