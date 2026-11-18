// ============================================================================
// NON-DESTRUCTIVE EXCEL TEMPLATE SYNC ENGINE (ExcelJS)
// Preserves 100% of existing sheets, borders, fonts, merges, formulas & data.
// ============================================================================
(function (globalScope) {
  'use strict';

  const SYNC_SHEET_NAME = '_CKP_SYNCED';
  const HIDDEN_LEGACY_SHEETS = new Set(['_CKP_HISTORY', '_CKP_SYNCED']);

  function getCellPlainText(cellValue) {
    if (cellValue === null || cellValue === undefined) return '';
    if (typeof cellValue === 'object') {
      if (cellValue.result !== undefined && cellValue.result !== null) {
        return String(cellValue.result).trim();
      }
      if (Array.isArray(cellValue.richText)) {
        return cellValue.richText.map((t) => t.text || '').join('').trim();
      }
      if (cellValue.text !== undefined) {
        return String(cellValue.text).trim();
      }
      if (cellValue instanceof Date && !Number.isNaN(cellValue.getTime())) {
        let y = cellValue.getUTCFullYear();
        if (y < 2400) y += 543;
        return `${cellValue.getUTCDate()}/${cellValue.getUTCMonth() + 1}/${y}`;
      }
      return '';
    }
    return String(cellValue).trim();
  }

  function formatShortThaiDateBE(isoDate) {
    if (!isoDate) return '';
    const m = String(isoDate).match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (!m) return String(isoDate);
    const yearBE = Number(m[1]) + 543;
    const month = Number(m[2]);
    const day = Number(m[3]);
    return `${day}/${month}/${yearBE}`;
  }

  function formatExcelTimeDisplay(timeStr) {
    const s = String(timeStr || '').trim();
    if (!s) return '';
    const m = s.match(/^0?(\d{1,2}):(\d{2})/);
    if (m) return `${m[1]}:${m[2]}`;
    return s;
  }

  function splitRouteParts(routeStr, fallbackSite = '') {
    const raw = String(routeStr || '').trim();
    const site = String(fallbackSite || '').trim();
    if (!raw) return { from: site, to: site };
    const spacedParts = raw.split(/\s+(?:-|->|→|ถึง)\s*|\s*(?:-|->|→|ถึง)\s+/).filter(Boolean);
    if (spacedParts.length >= 2) {
      return {
        from: spacedParts[0].trim(),
        to: spacedParts.slice(1).join(' - ').trim()
      };
    }
    const parts = raw.split(/\s*(?:-|->|→|ถึง)\s*/).filter(Boolean);
    if (parts.length >= 2) {
      return {
        from: parts[0].trim(),
        to: parts.slice(1).join('-').trim()
      };
    }
    return {
      from: raw,
      to: site || raw
    };
  }

  function getDirectionalRouteParts(routeStr, timeStr, fallbackSite = '') {
    const { from, to } = splitRouteParts(routeStr, fallbackSite);
    if (!from || !to || from === to) return { from, to };
    const m = String(timeStr || '').trim().match(/^(\d{1,2}):(\d{2})/);
    if (m && Number(m[1]) >= 12) {
      return { from: to, to: from };
    }
    return { from, to };
  }

  function normalizeRouteKey(str) {
    return String(str || '')
      .replace(/[–—]/g, '-')
      .replace(/\s+/g, '')
      .toLowerCase()
      .trim();
  }

  function normalizeRouteToken(str) {
    return String(str || '')
      .replace(/[\s\-–—.+]/g, '')
      .toLowerCase()
      .trim();
  }

  function stripSheetPageSuffix(name) {
    return String(name || '').replace(/\s*\(\d+\)\s*$/, '').trim();
  }

  function sanitizeSheetName(rawName, fallbackName = 'Sheet') {
    const cleaned = String(rawName || '')
      .replace(/[\\/?*[\]:]/g, ' ')
      .trim();
    return (cleaned || fallbackName).slice(0, 28);
  }

  function getUniqueSheetName(wb, desiredName) {
    const base = sanitizeSheetName(desiredName, 'Sheet');
    if (!wb.getWorksheet(base)) return base;
    let counter = 2;
    while (counter < 200) {
      const suffix = ` (${counter})`;
      const candidate = `${base.slice(0, 31 - suffix.length)}${suffix}`;
      if (!wb.getWorksheet(candidate)) return candidate;
      counter += 1;
    }
    return `${base.slice(0, 24)}_${Date.now() % 1000}`;
  }

  function cloneStyleSafe(styleObj) {
    if (!styleObj) return {};
    return JSON.parse(JSON.stringify(styleObj));
  }

  function thinBorderAll() {
    return {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FF000000' } }
    };
  }

  function makeSyncKey(row) {
    if (row && row.log_id) {
      return `id:${String(row.log_id).trim()}`;
    }
    return [
      'sig',
      String(row.worker_name || '').trim(),
      String(row.travel_date || '').trim(),
      String(row.departure_time || '').trim(),
      String(row.arrival_time || '').trim(),
      String(row.route || '').trim(),
      String(row.vehicle_plate || '').trim(),
      String(row.passenger_count ?? '').trim(),
      String(row.note || '').trim()
    ].join('||');
  }

  function sortRowsChronologically(rows) {
    return [...(rows || [])].sort((a, b) => {
      const d1 = String(a.travel_date || '');
      const d2 = String(b.travel_date || '');
      if (d1 !== d2) return d1.localeCompare(d2);
      const t1 = String(a.departure_time || a.arrival_time || '');
      const t2 = String(b.departure_time || b.arrival_time || '');
      return t1.localeCompare(t2);
    });
  }

  function readSyncedKeysFromWorkbook(wb) {
    const syncedSet = new Set();
    const ws = wb.getWorksheet(SYNC_SHEET_NAME);
    if (!ws) return syncedSet;
    for (let r = 2; r <= ws.rowCount; r++) {
      const key = getCellPlainText(ws.getRow(r).getCell(1).value);
      if (key) syncedSet.add(key);
    }
    return syncedSet;
  }

  function appendSyncedKeysToWorkbook(wb, entries) {
    if (!entries || entries.length === 0) return;
    let ws = wb.getWorksheet(SYNC_SHEET_NAME);
    if (!ws) {
      ws = wb.addWorksheet(SYNC_SHEET_NAME, { state: 'veryHidden' });
      ws.getRow(1).values = ['sync_key', 'log_id', 'target_sheet', 'target_row', 'route', 'travel_date'];
    } else {
      ws.state = 'veryHidden';
    }
    let nextR = ws.rowCount + 1;
    if (nextR < 2) nextR = 2;
    for (const item of entries) {
      ws.getRow(nextR).values = [
        item.sync_key,
        item.log_id || '',
        item.target_sheet || '',
        item.target_row || '',
        item.route || '',
        item.travel_date || ''
      ];
      nextR += 1;
    }
  }

  function isCopySheet(ws) {
    if (!ws || HIDDEN_LEGACY_SHEETS.has(ws.name)) return false;
    const r5 = ws.getRow(5);
    const c1 = getCellPlainText(r5.getCell(1).value);
    const c2 = getCellPlainText(r5.getCell(2).value);
    const c3 = getCellPlainText(r5.getCell(3).value);
    return c1.includes('ลำดับ') && c2.includes('วัน') && c3.includes('ทะเบียนรถ');
  }

  function isBillingSummarySheet(ws) {
    if (!ws || HIDDEN_LEGACY_SHEETS.has(ws.name)) return false;
    const c1 = getCellPlainText(ws.getRow(1).getCell(1).value);
    return ws.name === 'ใบเบิก' || c1.includes('ใบเบิกผลงาน');
  }

  function isOperationSheet(ws) {
    if (!ws || HIDDEN_LEGACY_SHEETS.has(ws.name)) return false;
    const c1 = getCellPlainText(ws.getRow(1).getCell(1).value);
    const r5c1 = getCellPlainText(ws.getRow(5).getCell(1).value);
    const r5c2 = getCellPlainText(ws.getRow(5).getCell(2).value);
    return c1.includes('ใบบันทึกการปฏิบัติงาน') || (r5c1.includes('ลำดับ') && r5c2.includes('ว/ด/ป'));
  }

  // Safely read sheet merges array across ExcelJS versions
  function getWorksheetMergesList(ws) {
    if (Array.isArray(ws?.model?.merges)) return ws.model.merges;
    if (ws?._merges && typeof ws._merges === 'object') {
      return Object.keys(ws._merges).map((k) => {
        const m = ws._merges[k];
        return m && m.range ? m.range : k;
      });
    }
    return [];
  }

  function copyWorksheetTemplate(wb, donorWs, newName, clearFromRow) {
    const finalName = getUniqueSheetName(wb, newName);
    const ws = wb.addWorksheet(finalName);

    // Copy column widths
    if (donorWs.columns) {
      ws.columns = donorWs.columns.map((c) => ({
        width: c.width || 14,
        style: c.style ? cloneStyleSafe(c.style) : undefined
      }));
    }
    // Copy explicit col properties 1..12
    for (let c = 1; c <= 12; c++) {
      const srcCol = donorWs.getColumn(c);
      if (srcCol && srcCol.width) {
        ws.getColumn(c).width = srcCol.width;
      }
    }

    if (donorWs.pageSetup) {
      ws.pageSetup = cloneStyleSafe(donorWs.pageSetup);
    }

    const maxCopyRow = clearFromRow ? Math.min(donorWs.rowCount, Math.max(clearFromRow + 35, 45)) : donorWs.rowCount;
    for (let r = 1; r <= maxCopyRow; r++) {
      const srcRow = donorWs.getRow(r);
      const dstRow = ws.getRow(r);
      if (srcRow.height) dstRow.height = srcRow.height;

      for (let c = 1; c <= 12; c++) {
        const srcCell = srcRow.getCell(c);
        const dstCell = dstRow.getCell(c);
        if (srcCell.style) {
          dstCell.style = cloneStyleSafe(srcCell.style);
        }
        if (!clearFromRow || r < clearFromRow) {
          if (srcCell.value !== null && srcCell.value !== undefined) {
            if (typeof srcCell.value === 'object' && srcCell.value.formula) {
              dstCell.value = { formula: srcCell.value.formula, result: srcCell.value.result ?? 0 };
            } else {
              dstCell.value = srcCell.value;
            }
          }
        }
      }
    }

    // Apply merges (for rows < clearFromRow OR explicit template merges)
    const merges = getWorksheetMergesList(donorWs);
    for (const ref of merges) {
      try {
        ws.mergeCells(ref);
      } catch {
        // ignore overlapping/invalid merge
      }
    }

    return ws;
  }

  // Create a clean styled Copy ("สำเนา") worksheet when workbook has no donor sheets yet
  function createStyledCopyWorksheet(wb, sheetName, projectTitle) {
    const finalName = getUniqueSheetName(wb, sheetName);
    const ws = wb.addWorksheet(finalName);
    const widths = [8, 15, 16, 22, 11, 12, 12, 24];
    widths.forEach((w, idx) => {
      ws.getColumn(idx + 1).width = w;
    });

    ws.getRow(1).values = ['บริษัท จักรพงษ์ ขนส่ง (2003) จำกัด'];
    ws.getRow(2).values = ['บันทึกรายการรถ 6 ล้อ รับ-ส่งคนงาน'];
    ws.getRow(3).values = [`โครงการ  ${projectTitle || ''}`];
    ws.getRow(5).values = ['ลำดับ', 'วัน เดือน ปี', 'ทะเบียนรถ', 'จาก', 'เวลา', 'จำนวนคน', 'ถึง', 'หมายเหตุ'];

    ws.mergeCells('A1:H1');
    ws.mergeCells('A2:H2');
    ws.mergeCells('A3:H3');

    for (let r = 1; r <= 3; r++) {
      const cell = ws.getRow(r).getCell(1);
      cell.font = { name: 'Tahoma', size: r === 1 ? 14 : 12, bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      ws.getRow(r).height = 22;
    }

    const r5 = ws.getRow(5);
    r5.height = 22;
    for (let c = 1; c <= 8; c++) {
      const cell = r5.getCell(c);
      cell.font = { name: 'Tahoma', size: 11, bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = thinBorderAll();
    }

    for (let r = 6; r <= 75; r++) {
      const row = ws.getRow(r);
      row.getCell(1).value = r - 5;
      for (let c = 1; c <= 8; c++) {
        const cell = row.getCell(c);
        cell.font = { name: 'Tahoma', size: 11 };
        cell.alignment = { horizontal: c === 4 || c === 8 ? 'left' : 'center', vertical: 'middle' };
        cell.border = thinBorderAll();
      }
    }

    return ws;
  }

  function findBestMatchingCopySheet(wb, routeKey) {
    const normTarget = normalizeRouteKey(routeKey);
    const tokTarget = normalizeRouteToken(routeKey);
    const { from: wantFrom, to: wantTo } = splitRouteParts(routeKey);
    const tokFrom = normalizeRouteToken(wantFrom);
    const tokTo = normalizeRouteToken(wantTo);

    const candidateSheets = wb.worksheets.filter((ws) => isCopySheet(ws) && ws.name !== 'รวมสำเนาทุกสาย');
    for (const ws of candidateSheets) {
      if (normalizeRouteKey(ws.name) === normTarget || normalizeRouteToken(ws.name) === tokTarget) return ws;
    }
    if (tokFrom && tokTo) {
      for (const ws of candidateSheets) {
        const tokName = normalizeRouteToken(ws.name);
        if (tokName.includes(tokFrom) && tokName.includes(tokTo)) return ws;
      }
    }
    return null;
  }

  function applyRowsToCopyWorkbookExcelJS(wb, webRows, defaultProjectTitle = '') {
    const syncedSet = readSyncedKeysFromWorkbook(wb);
    const sortedRows = sortRowsChronologically(webRows);
    const newRows = sortedRows.filter((r) => !syncedSet.has(makeSyncKey(r)));

    if (wb.worksheets.length === 0 && newRows.length === 0) {
      createStyledCopyWorksheet(wb, 'สำเนา_1', defaultProjectTitle);
      return { syncedCount: 0, totalFlowed: syncedSet.size };
    }
    if (newRows.length === 0) {
      return { syncedCount: 0, totalFlowed: syncedSet.size };
    }

    const donorCopyWs = wb.worksheets.find((ws) => isCopySheet(ws));
    let donorDataStyles = null;
    let donorSummaryDriverStyle = null;
    let donorSummaryCountStyle = null;
    if (donorCopyWs) {
      donorDataStyles = [];
      for (let c = 1; c <= 8; c++) {
        donorDataStyles.push(cloneStyleSafe(donorCopyWs.getRow(6).getCell(c).style));
      }
      for (let r = 6; r <= donorCopyWs.rowCount; r++) {
        const gVal = donorCopyWs.getRow(r).getCell(7).value;
        if (gVal && typeof gVal === 'object' && String(gVal.formula || '').toUpperCase().includes('COUNTIF')) {
          donorSummaryDriverStyle = cloneStyleSafe(donorCopyWs.getRow(r).getCell(6).style);
          donorSummaryCountStyle = cloneStyleSafe(donorCopyWs.getRow(r).getCell(7).style);
          break;
        }
      }
    }

    const buckets = new Map();
    for (const r of newRows) {
      const key = (r.route || r.site_name || '').trim() || 'ทั่วไป';
      if (!buckets.has(key)) buckets.set(key, []);
      buckets.get(key).push(r);
    }

    const legacyHistoryWs = wb.getWorksheet('_CKP_HISTORY');
    if (legacyHistoryWs) legacyHistoryWs.state = 'veryHidden';

    const newlySyncedRecords = [];

    for (const [bucketKey, bucketRows] of buckets.entries()) {
      const bucketProject =
        (bucketRows.find((x) => x.site_name)?.site_name || '').trim() || defaultProjectTitle;
      let ws = findBestMatchingCopySheet(wb, bucketKey, bucketRows);

      if (!ws) {
        if (donorCopyWs) {
          ws = copyWorksheetTemplate(wb, donorCopyWs, sanitizeSheetName(bucketKey, 'สำเนา'), 6);
          ws.getRow(3).getCell(1).value = `โครงการ  ${bucketProject}`;
          for (let r = 6; r <= ws.rowCount; r++) {
            const row = ws.getRow(r);
            row.getCell(1).value = r - 5;
            for (let c = 2; c <= 8; c++) {
              row.getCell(c).value = null;
              if (donorDataStyles && donorDataStyles[c - 1]) {
                row.getCell(c).style = cloneStyleSafe(donorDataStyles[c - 1]);
              }
            }
          }
        } else {
          ws = createStyledCopyWorksheet(wb, sanitizeSheetName(bucketKey, 'สำเนา'), bucketProject);
        }
      }

      // Scan existing trip rows (rows >= 6 where col C or D is populated and col G isn't COUNTIF/SUM)
      let lastDataRow = 5;
      let prevDateStr = '';
      let sheetDataStyle = [];
      let sheetSummaryDriverStyle = donorSummaryDriverStyle;
      let sheetSummaryCountStyle = donorSummaryCountStyle;
      let sheetTotalCountStyle = null;
      let previousBottomRow = 50;

      for (let r = 6; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        const cVal = getCellPlainText(row.getCell(3).value);
        const dVal = getCellPlainText(row.getCell(4).value);
        const fVal = getCellPlainText(row.getCell(6).value);
        const gRaw = row.getCell(7).value;
        const gFormula = gRaw && typeof gRaw === 'object' ? String(gRaw.formula || '').toUpperCase() : '';

        if (gFormula.includes('COUNTIF')) {
          if (!sheetSummaryDriverStyle && fVal) {
            sheetSummaryDriverStyle = cloneStyleSafe(row.getCell(6).style);
            sheetSummaryCountStyle = cloneStyleSafe(row.getCell(7).style);
          }
          previousBottomRow = Math.max(previousBottomRow, r + 5);
          continue;
        }
        if (gFormula.includes('SUM(')) {
          sheetTotalCountStyle = cloneStyleSafe(row.getCell(7).style);
          previousBottomRow = Math.max(previousBottomRow, r + 5);
          continue;
        }
        if (cVal || dVal) {
          lastDataRow = r;
          const bText = getCellPlainText(row.getCell(2).value);
          if (bText) prevDateStr = bText;
        }
        if (getCellPlainText(row.getCell(1).value)) {
          previousBottomRow = Math.max(previousBottomRow, r);
        }
      }

      // Capture row border/font styles from lastDataRow or Row 6
      const styleSourceRow = lastDataRow >= 6 ? ws.getRow(lastDataRow) : ws.getRow(6);
      for (let c = 1; c <= 8; c++) {
        const st = styleSourceRow.getCell(c).style;
        sheetDataStyle.push(
          st && Object.keys(st).length > 0
            ? cloneStyleSafe(st)
            : donorDataStyles && donorDataStyles[c - 1]
            ? cloneStyleSafe(donorDataStyles[c - 1])
            : { font: { name: 'Tahoma', size: 11 }, border: thinBorderAll(), alignment: { horizontal: 'center' } }
        );
      }

      // Clear summary & tail below lastDataRow before appending (existing trip rows 6..lastDataRow are 100% UNTOUCHED!)
      const targetTailEnd = Math.max(previousBottomRow, lastDataRow + bucketRows.length + 25, 55);
      for (let r = lastDataRow + 1; r <= targetTailEnd; r++) {
        const row = ws.getRow(r);
        for (let c = 1; c <= 8; c++) {
          const cell = row.getCell(c);
          cell.value = null;
          cell.style = cloneStyleSafe(sheetDataStyle[c - 1]);
        }
      }

      // Append new rows right after lastDataRow
      for (let i = 0; i < bucketRows.length; i++) {
        const item = bucketRows[i];
        const rIdx = lastDataRow + 1 + i;
        const row = ws.getRow(rIdx);
        const seqNo = rIdx - 5;
        const dateBE = formatShortThaiDateBE(item.travel_date);
        const displayDate = dateBE && dateBE !== prevDateStr ? dateBE : null;
        if (dateBE) prevDateStr = dateBE;

        const tripTime = item.departure_time || item.arrival_time;
        const { from } = getDirectionalRouteParts(item.route, tripTime, item.site_name || bucketProject);
        const driverOrPlate = String(item.vehicle_plate || item.worker_name || '').trim();
        const pax =
          item.passenger_count !== null && item.passenger_count !== undefined && item.passenger_count !== ''
            ? Number(item.passenger_count)
            : null;
        const noteText = String(item.note || '').trim() || null;

        row.getCell(1).value = seqNo;
        row.getCell(2).value = displayDate;
        row.getCell(3).value = driverOrPlate;
        row.getCell(4).value = from || null;
        row.getCell(5).value = formatExcelTimeDisplay(tripTime) || null;
        row.getCell(6).value = pax;
        row.getCell(7).value = null;
        row.getCell(8).value = noteText;

        newlySyncedRecords.push({
          sync_key: makeSyncKey(item),
          log_id: item.log_id || '',
          target_sheet: ws.name,
          target_row: rIdx,
          route: item.route || '',
          travel_date: item.travel_date || ''
        });
      }

      const newLastDataRow = lastDataRow + bucketRows.length;

      // Recompute tally across ALL trip rows 6..newLastDataRow (preserving existing + new rows!)
      const tallyMap = new Map();
      for (let r = 6; r <= newLastDataRow; r++) {
        const key = getCellPlainText(ws.getRow(r).getCell(3).value);
        if (!key) continue;
        if (!tallyMap.has(key)) {
          tallyMap.set(key, { firstRow: r, count: 0 });
        }
        tallyMap.get(key).count += 1;
      }

      // 2 spacer rows with seq numbers
      ws.getRow(newLastDataRow + 1).getCell(1).value = newLastDataRow + 1 - 5;
      ws.getRow(newLastDataRow + 2).getCell(1).value = newLastDataRow + 2 - 5;

      let cursor = newLastDataRow + 3;
      const summaryStartRow = cursor;
      let totalTally = 0;

      const defaultSummaryFill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF92D050' }
      };

      for (const [key, info] of tallyMap.entries()) {
        const sRow = ws.getRow(cursor);
        sRow.getCell(1).value = cursor - 5;
        const fCell = sRow.getCell(6);
        const gCell = sRow.getCell(7);
        fCell.value = key;
        gCell.value = {
          formula: `COUNTIF(C:C,C${info.firstRow})`,
          result: info.count
        };
        if (sheetSummaryDriverStyle) {
          fCell.style = cloneStyleSafe(sheetSummaryDriverStyle);
        } else {
          fCell.fill = defaultSummaryFill;
          fCell.border = thinBorderAll();
          fCell.alignment = { horizontal: 'center' };
        }
        if (sheetSummaryCountStyle) {
          gCell.style = cloneStyleSafe(sheetSummaryCountStyle);
        } else {
          gCell.fill = defaultSummaryFill;
          gCell.border = thinBorderAll();
          gCell.alignment = { horizontal: 'center' };
        }
        totalTally += info.count;
        cursor += 1;
      }

      if (tallyMap.size > 0) {
        const summaryEndRow = cursor - 1;
        const tRow = ws.getRow(cursor);
        tRow.getCell(1).value = cursor - 5;
        const tgCell = tRow.getCell(7);
        tgCell.value = {
          formula: `SUM(G${summaryStartRow}:G${summaryEndRow})`,
          result: totalTally
        };
        if (sheetTotalCountStyle) {
          tgCell.style = cloneStyleSafe(sheetTotalCountStyle);
        } else {
          tgCell.border = thinBorderAll();
          tgCell.font = { name: 'Tahoma', size: 11, bold: true };
          tgCell.alignment = { horizontal: 'center' };
        }
        cursor += 1;
      }

      while (cursor <= targetTailEnd) {
        ws.getRow(cursor).getCell(1).value = cursor - 5;
        cursor += 1;
      }
    }

    appendSyncedKeysToWorkbook(wb, newlySyncedRecords);
    return {
      syncedCount: newlySyncedRecords.length,
      totalFlowed: syncedSet.size + newlySyncedRecords.length,
      summaryText: formatSyncSummaryText(newlySyncedRecords)
    };
  }

  function formatSyncSummaryText(records) {
    if (!records || records.length === 0) return '';
    const bySheet = new Map();
    for (const rec of records) {
      const sName = rec.target_sheet || 'Sheet';
      if (!bySheet.has(sName)) bySheet.set(sName, []);
      bySheet.get(sName).push(rec.target_row);
    }
    const parts = [];
    for (const [sName, rows] of bySheet.entries()) {
      const minR = Math.min(...rows);
      const maxR = Math.max(...rows);
      parts.push(
        minR === maxR
          ? `"${sName}" (แถว ${minR})`
          : `"${sName}" (แถว ${minR}-${maxR})`
      );
    }
    return parts.join(', ');
  }

  // Create clean styled Sheet 1 ("ใบเบิก") if starting from a brand-new workbook
  function createStyledBillingSummarySheet(wb, projectTitle, issueDateBE, periodLabel) {
    const ws = wb.addWorksheet('ใบเบิก');
    const widths = [8, 16, 16, 16, 10, 10, 12, 10, 10, 18];
    widths.forEach((w, idx) => {
      ws.getColumn(idx + 1).width = w;
    });

    ws.getRow(1).values = ['ใบเบิกผลงานผู้รับเหมา'];
    ws.mergeCells('A1:J1');
    ws.getRow(1).getCell(1).font = { name: 'Tahoma', size: 16, bold: true };
    ws.getRow(1).getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

    ws.getRow(2).values = ['ผู้รับเหมาชุด    ', '', 'บริษัท จักรพงษ์ ขนส่ง ( 2003 ) จำกัด', '', '', '', '', 'วันที่', '', issueDateBE || ''];
    ws.mergeCells('C2:F2');
    ws.mergeCells('H2:I2');

    ws.getRow(3).values = ['โครงการ', '', projectTitle || '', '', '', '', '', 'ประจำงวด', '', periodLabel || ''];
    ws.mergeCells('C3:F3');
    ws.mergeCells('H3:I3');

    ws.getRow(5).values = ['ลำดับ', 'รายละเอียดงาน', '', '', 'ปริมาณงาน', '', '', 'จำนวนเงิน', '', 'หมายเหตุ'];
    ws.getRow(6).values = ['', '', '', '', 'จำนวน', 'หน่วย', 'ราคา/หน่วย', '', '', ''];
    ws.getRow(7).values = ['', 'ค่ารถรับ -  ส่ง   '];

    ws.mergeCells('A5:A6');
    ws.mergeCells('B5:D6');
    ws.mergeCells('E5:G5');
    ws.mergeCells('H5:I6');
    ws.mergeCells('J5:J6');

    for (let r = 5; r <= 6; r++) {
      for (let c = 1; c <= 10; c++) {
        const cell = ws.getRow(r).getCell(c);
        cell.font = { name: 'Tahoma', size: 11, bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = thinBorderAll();
      }
    }

    for (let r = 7; r <= 28; r++) {
      ws.mergeCells(`B${r}:D${r}`);
      ws.mergeCells(`H${r}:I${r}`);
      for (let c = 1; c <= 10; c++) {
        const cell = ws.getRow(r).getCell(c);
        cell.font = { name: 'Tahoma', size: 11, bold: r === 7 };
        cell.alignment = {
          horizontal: c >= 2 && c <= 4 ? 'left' : 'center',
          vertical: 'middle'
        };
        cell.border = thinBorderAll();
      }
    }

    ws.getRow(29).values = ['', 'รวม', '', '', '', '', '', { formula: 'SUM(H8:I28)', result: 0 }];
    ws.mergeCells('B29:D29');
    ws.mergeCells('H29:I29');
    for (let c = 1; c <= 10; c++) {
      const cell = ws.getRow(29).getCell(c);
      cell.font = { name: 'Tahoma', size: 11, bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = thinBorderAll();
    }

    ws.getRow(32).getCell(5).value = 'ผู้ตรวจรับ / Checked by         ';
    ws.getRow(33).getCell(5).value = 'วันที่ / Date                            ';
    return ws;
  }

  function createStyledOperationSheet(wb, sheetName, pageNo, siteTitle, startDateBE, endDateBE) {
    const finalName = getUniqueSheetName(wb, sheetName);
    const ws = wb.addWorksheet(finalName);
    const widths = [7, 13, 14, 22, 9, 10, 8, 8, 9, 9, 18, 20];
    widths.forEach((w, idx) => {
      ws.getColumn(idx + 1).width = w;
    });

    ws.getRow(1).values = ['ใบบันทึกการปฏิบัติงานรถ 6 ล้อ รับ - ส่งพนักงาน'];
    ws.mergeCells('A1:K1');
    ws.getRow(1).getCell(1).font = { name: 'Tahoma', size: 14, bold: true };
    ws.getRow(1).getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

    ws.getRow(2).values = ['', '', '', '', '', '', '', '', '', 'แผ่นที่', pageNo];
    ws.mergeCells('K2:L2');

    ws.getRow(3).values = [
      'สถานที่ปฏิบัติงานโครงการ :-',
      '',
      '',
      siteTitle || '',
      '',
      'ระหว่างวันที่',
      '',
      startDateBE || '',
      '',
      'ถึงวันที่ ',
      endDateBE || ''
    ];
    ws.mergeCells('D3:E3');
    ws.mergeCells('F3:G3');
    ws.mergeCells('H3:I3');
    ws.mergeCells('K3:L3');

    ws.getRow(5).values = ['ลำดับ', 'ว/ด/ป', 'ทะเบียนรถ', 'จาก', 'เวลา', 'ถึง', '', '', 'เวลา', 'จำนวน', 'ลงชื่อผู้ตรวจ', 'หมายเหตุ'];
    ws.getRow(6).values = ['', '', '', '', '( น. )', '', '', '', '( น. )', '(คน)', 'แคมป์ Boss /รปภ.', ''];

    ws.mergeCells('A5:A6');
    ws.mergeCells('B5:B6');
    ws.mergeCells('C5:C6');
    ws.mergeCells('D5:D6');
    ws.mergeCells('F5:H6');
    ws.mergeCells('L5:L6');

    for (let r = 5; r <= 6; r++) {
      for (let c = 1; c <= 12; c++) {
        const cell = ws.getRow(r).getCell(c);
        cell.font = { name: 'Tahoma', size: 10, bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = thinBorderAll();
      }
    }

    for (let i = 0; i < 30; i++) {
      const r = 7 + i;
      const row = ws.getRow(r);
      row.getCell(1).value = i + 1;
      ws.mergeCells(`F${r}:H${r}`);
      for (let c = 1; c <= 12; c++) {
        const cell = row.getCell(c);
        cell.font = { name: 'Tahoma', size: 10 };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = thinBorderAll();
      }
    }

    ws.getRow(37).values = ['', '', '', '', '', '', '', 'ผู้ตรวจ / Checked by', '', '', '…........................................................................'];
    ws.getRow(38).values = ['', '', '', '', '', '', '', 'วันที่ / Date', '', '', '…............../…..................../…............'];
    ws.mergeCells('K37:L37');
    ws.mergeCells('K38:L38');

    return ws;
  }

  function findMatchingOperationSheets(wb, routeKey) {
    const normTarget = normalizeRouteKey(routeKey);
    const tokTarget = normalizeRouteToken(routeKey);
    const { from: wantFrom, to: wantTo } = splitRouteParts(routeKey);
    const tokFrom = normalizeRouteToken(wantFrom);
    const tokTo = normalizeRouteToken(wantTo);

    const opSheets = wb.worksheets.filter((ws) => isOperationSheet(ws));
    // 1. Exact base name match
    const exactMatches = opSheets.filter((ws) => {
      const base = stripSheetPageSuffix(ws.name);
      return normalizeRouteKey(base) === normTarget || normalizeRouteToken(base) === tokTarget;
    });
    if (exactMatches.length > 0) return exactMatches;

    // 2. Token match against sheet name OR Row 7's จาก + ถึง
    if (tokFrom && tokTo) {
      const tokenMatches = opSheets.filter((ws) => {
        const r7From = normalizeRouteToken(ws.getRow(7).getCell(4).value);
        const r7To = normalizeRouteToken(ws.getRow(7).getCell(6).value);
        if (r7From && r7To) {
          const fMatch = tokFrom.includes(r7From) || r7From.includes(tokFrom);
          const tMatch = tokTo.includes(r7To) || r7To.includes(tokTo);
          const revF = tokFrom.includes(r7To) || r7To.includes(tokFrom);
          const revT = tokTo.includes(r7From) || r7From.includes(tokTo);
          if ((fMatch && tMatch) || (revF && revT)) return true;
        }
        const tokName = normalizeRouteToken(stripSheetPageSuffix(ws.name));
        return tokName.includes(tokFrom) && tokName.includes(tokTo);
      });
      if (tokenMatches.length > 0) return tokenMatches;
    }

    return [];
  }

  function updateBillingSummaryRow(wsBill, routeName, addTripsCount, todayBE) {
    if (!wsBill) return;
    const tokTarget = normalizeRouteToken(routeName);
    const { from: wantFrom, to: wantTo } = splitRouteParts(routeName);
    const tokFrom = normalizeRouteToken(wantFrom);
    const tokTo = normalizeRouteToken(wantTo);

    // Ensure J2 date is populated if empty
    if (!getCellPlainText(wsBill.getRow(2).getCell(10).value) && todayBE) {
      wsBill.getRow(2).getCell(10).value = todayBE;
    }

    // Find "รวม" row index
    let totalRowIdx = 29;
    for (let r = 8; r <= Math.min(wsBill.rowCount, 60); r++) {
      if (getCellPlainText(wsBill.getRow(r).getCell(2).value) === 'รวม') {
        totalRowIdx = r;
        break;
      }
    }

    // 1. Check if this route already exists in rows 8..(totalRowIdx - 1)
    let highestSeq = 0;
    let lastOccupiedRow = 7;
    for (let r = 8; r < totalRowIdx; r++) {
      const row = wsBill.getRow(r);
      const seqNum = Number.parseInt(getCellPlainText(row.getCell(1).value), 10);
      if (!Number.isNaN(seqNum) && seqNum > highestSeq) highestSeq = seqNum;
      const bText = getCellPlainText(row.getCell(2).value);
      if (bText) {
        lastOccupiedRow = r;
        const tokB = normalizeRouteToken(bText);
        const isDirect = tokB === tokTarget;
        const isToken =
          tokFrom &&
          tokTo &&
          tokB.includes(tokFrom) &&
          tokB.includes(tokTo);
        if (isDirect || isToken) {
          const curQty = Number(getCellPlainText(row.getCell(5).value) || 0);
          const nextQty = (Number.isNaN(curQty) ? 0 : curQty) + addTripsCount;
          row.getCell(5).value = nextQty;
          const unitPrice = Number(getCellPlainText(row.getCell(7).value) || 0);
          row.getCell(8).value = {
            formula: `SUM(E${r}*G${r})`,
            result: !Number.isNaN(unitPrice) ? nextQty * unitPrice : 0
          };
          // Ensure total formula sums full item range
          wsBill.getRow(totalRowIdx).getCell(8).value = {
            formula: `SUM(H8:I${totalRowIdx - 1})`,
            result: Number(getCellPlainText(wsBill.getRow(totalRowIdx).getCell(8).value) || 0) +
              (!Number.isNaN(unitPrice) ? addTripsCount * unitPrice : 0)
          };
          return;
        }
      }
    }

    // 2. Pick next empty even slot (leaving 1 spacer row after lastOccupiedRow)
    let insertRowIdx = lastOccupiedRow <= 7 ? 8 : lastOccupiedRow + 2;
    if (insertRowIdx >= totalRowIdx) {
      // No pre-made slots before รวม — overwrite total? No, insert 2 rows before totalRowIdx!
      wsBill.spliceRows(totalRowIdx, 0, ['', '', '', '', '', '', '', '', '', ''], ['', '', '', '', '', '', '', '', '', '']);
      totalRowIdx += 2;
    }

    // Ensure merged cells exist on insertRowIdx (B:D and H:I)
    try { wsBill.mergeCells(`B${insertRowIdx}:D${insertRowIdx}`); } catch {}
    try { wsBill.mergeCells(`H${insertRowIdx}:I${insertRowIdx}`); } catch {}

    const targetRow = wsBill.getRow(insertRowIdx);
    targetRow.getCell(1).value = highestSeq + 1;
    targetRow.getCell(2).value = routeName;
    targetRow.getCell(5).value = addTripsCount;
    targetRow.getCell(6).value = 'เที่ยว';
    targetRow.getCell(8).value = { formula: `SUM(E${insertRowIdx}*G${insertRowIdx})`, result: 0 };

    // Copy border/font style from Row 8 if present
    const styleSrc = wsBill.getRow(8);
    for (let c = 1; c <= 10; c++) {
      const s = styleSrc.getCell(c).style;
      if (s && Object.keys(s).length > 0) {
        targetRow.getCell(c).style = cloneStyleSafe(s);
      } else {
        targetRow.getCell(c).border = thinBorderAll();
      }
    }

    // Update SUM range in รวม row so newly added rows are automatically included
    const totalCell = wsBill.getRow(totalRowIdx).getCell(8);
    let recomputedTotal = 0;
    for (let r = 8; r < totalRowIdx; r++) {
      const qty = Number(getCellPlainText(wsBill.getRow(r).getCell(5).value) || 0);
      const price = Number(getCellPlainText(wsBill.getRow(r).getCell(7).value) || 0);
      if (!Number.isNaN(qty) && !Number.isNaN(price)) {
        recomputedTotal += qty * price;
      }
    }
    totalCell.value = {
      formula: `SUM(H8:I${totalRowIdx - 1})`,
      result: recomputedTotal
    };
  }

  function applyRowsToBillingWorkbookExcelJS(wb, webRows, defaultProjectTitle = '') {
    const syncedSet = readSyncedKeysFromWorkbook(wb);
    const sortedRows = sortRowsChronologically(webRows);
    const newRows = sortedRows.filter((r) => !syncedSet.has(makeSyncKey(r)));

    const dates = webRows.map((r) => r.travel_date).filter(Boolean).sort();
    const todayISO =
      typeof globalScope.getTodayISO === 'function'
        ? globalScope.getTodayISO()
        : new Date().toISOString().slice(0, 10);
    const todayBE = formatShortThaiDateBE(todayISO);
    const startDateBE = dates.length > 0 ? formatShortThaiDateBE(dates[0]) : todayBE;
    const endDateBE = dates.length > 0 ? formatShortThaiDateBE(dates[dates.length - 1]) : todayBE;
    const periodLabel = startDateBE === endDateBE ? startDateBE : `${startDateBE} - ${endDateBE}`;

    let wsBill = wb.worksheets.find((ws) => isBillingSummarySheet(ws));
    let donorOpWs = wb.worksheets.find((ws) => isOperationSheet(ws));

    if (!wsBill) {
      wsBill = createStyledBillingSummarySheet(wb, defaultProjectTitle, todayBE, periodLabel);
    }
    if (wb.worksheets.length === 1 && newRows.length === 0) {
      createStyledOperationSheet(wb, 'แผ่นที่ 1', 1, defaultProjectTitle, startDateBE, endDateBE);
      return { syncedCount: 0, totalFlowed: syncedSet.size };
    }
    if (newRows.length === 0) {
      return { syncedCount: 0, totalFlowed: syncedSet.size };
    }

    // Track max page number ("แผ่นที่") across existing sheets
    let maxPageNo = 0;
    for (const ws of wb.worksheets) {
      if (isOperationSheet(ws)) {
        const pVal = Number.parseInt(getCellPlainText(ws.getRow(2).getCell(11).value), 10);
        if (!Number.isNaN(pVal) && pVal > maxPageNo) maxPageNo = pVal;
        else maxPageNo += 1;
      }
    }

    const buckets = new Map();
    for (const r of newRows) {
      const key = (r.route || r.site_name || '').trim() || defaultProjectTitle || 'ทั่วไป';
      if (!buckets.has(key)) buckets.set(key, []);
      buckets.get(key).push(r);
    }

    const legacyHistoryWs = wb.getWorksheet('_CKP_HISTORY');
    if (legacyHistoryWs) legacyHistoryWs.state = 'veryHidden';

    const newlySyncedRecords = [];

    for (const [bucketKey, bucketRows] of buckets.entries()) {
      const bucketProject =
        (bucketRows.find((x) => x.site_name)?.site_name || '').trim() || defaultProjectTitle;

      // 1. Update Sheet 1 ("ใบเบิก") summary count & SUM
      updateBillingSummaryRow(wsBill, bucketKey, bucketRows.length, todayBE);

      // 2. Flow individual trip rows into matching ใบบันทึก sheet (or append new 30-row sheet(s))
      const matchedSheets = findMatchingOperationSheets(wb, bucketKey);
      let remaining = [...bucketRows];
      let seqOffset = 0;
      let prevDateStr = '';

      if (matchedSheets.length > 0) {
        const lastWs = matchedSheets[matchedSheets.length - 1];
        // Find last occupied row in Rows 7..36
        let lastUsedRow = 6;
        let lastSeq = 0;
        for (let r = 7; r <= 36; r++) {
          const row = lastWs.getRow(r);
          const sNum = Number.parseInt(getCellPlainText(row.getCell(1).value), 10);
          if (!Number.isNaN(sNum)) lastSeq = sNum;
          const bText = getCellPlainText(row.getCell(2).value);
          if (bText) prevDateStr = bText;
          if (
            getCellPlainText(row.getCell(3).value) ||
            getCellPlainText(row.getCell(4).value) ||
            getCellPlainText(row.getCell(6).value)
          ) {
            lastUsedRow = r;
          }
        }
        seqOffset = Math.max(lastSeq - (36 - lastUsedRow), lastUsedRow - 6);
        let fillRow = lastUsedRow + 1;
        while (fillRow <= 36 && remaining.length > 0) {
          const item = remaining.shift();
          const row = lastWs.getRow(fillRow);
          const dateBE = formatShortThaiDateBE(item.travel_date);
          const displayDate = dateBE && dateBE !== prevDateStr ? dateBE : null;
          if (dateBE) prevDateStr = dateBE;

          const { from, to } = getDirectionalRouteParts(
            item.route,
            item.departure_time || item.arrival_time,
            item.site_name || bucketProject
          );
          const pax =
            item.passenger_count !== null && item.passenger_count !== undefined && item.passenger_count !== ''
              ? Number(item.passenger_count)
              : null;

          if (!getCellPlainText(row.getCell(1).value)) {
            row.getCell(1).value = seqOffset + (fillRow - 6);
          }
          row.getCell(2).value = displayDate;
          row.getCell(3).value = String(item.vehicle_plate || item.worker_name || '').trim() || null;
          row.getCell(4).value = from || null;
          row.getCell(5).value = formatExcelTimeDisplay(item.departure_time) || null;
          row.getCell(6).value = to || null;
          row.getCell(9).value = formatExcelTimeDisplay(item.arrival_time) || null;
          row.getCell(10).value = pax;
          row.getCell(12).value = String(item.note || '').trim() || null;

          newlySyncedRecords.push({
            sync_key: makeSyncKey(item),
            log_id: item.log_id || '',
            target_sheet: lastWs.name,
            target_row: fillRow,
            route: item.route || '',
            travel_date: item.travel_date || ''
          });
          seqOffset = Number.parseInt(getCellPlainText(row.getCell(1).value), 10) || seqOffset + 1;
          fillRow += 1;
        }
      }

      // Create additional 30-row sheet(s) if remaining rows still need space
      let pageIndexForRoute = matchedSheets.length;
      while (remaining.length > 0) {
        const chunk = remaining.splice(0, 30);
        maxPageNo += 1;
        pageIndexForRoute += 1;

        const baseTabName = sanitizeSheetName(bucketKey, `แผ่นที่ ${maxPageNo}`);
        const desiredTabName =
          pageIndexForRoute === 1 ? baseTabName : `${baseTabName.slice(0, 24)} (${pageIndexForRoute})`;

        let wsPage;
        const pageDonor = matchedSheets[matchedSheets.length - 1] || donorOpWs;
        if (pageDonor) {
          wsPage = copyWorksheetTemplate(wb, pageDonor, desiredTabName, 7);
          wsPage.getRow(2).getCell(11).value = maxPageNo;
          if (bucketProject) wsPage.getRow(3).getCell(4).value = bucketProject;
          wsPage.getRow(3).getCell(8).value = startDateBE;
          wsPage.getRow(3).getCell(11).value = endDateBE;

          // Clear values in rows 7..36 (keep merges F:H, borders, and copy Row 7 cell styles onto all 30 rows)
          const r7Styles = [];
          for (let c = 1; c <= 12; c++) {
            r7Styles.push(cloneStyleSafe(pageDonor.getRow(7).getCell(c).style));
          }
          for (let i = 0; i < 30; i++) {
            const rIdx = 7 + i;
            const row = wsPage.getRow(rIdx);
            if (pageDonor.getRow(7).height) row.height = pageDonor.getRow(7).height;
            row.getCell(1).value = matchedSheets.length > 0 ? seqOffset + i + 1 : i + 1;
            for (let c = 2; c <= 12; c++) {
              row.getCell(c).value = null;
            }
            for (let c = 1; c <= 12; c++) {
              if (r7Styles[c - 1] && Object.keys(r7Styles[c - 1]).length > 0) {
                row.getCell(c).style = cloneStyleSafe(r7Styles[c - 1]);
              }
            }
            try { wsPage.mergeCells(`F${rIdx}:H${rIdx}`); } catch {}
          }
          // Copy footer rows 37..39
          for (let fr = 37; fr <= Math.min(pageDonor.rowCount, 39); fr++) {
            const srcFr = pageDonor.getRow(fr);
            const dstFr = wsPage.getRow(fr);
            for (let c = 1; c <= 12; c++) {
              dstFr.getCell(c).value = srcFr.getCell(c).value;
              dstFr.getCell(c).style = cloneStyleSafe(srcFr.getCell(c).style);
            }
          }
          if (!donorOpWs) donorOpWs = wsPage;
        } else {
          wsPage = createStyledOperationSheet(
            wb,
            desiredTabName,
            maxPageNo,
            bucketProject,
            startDateBE,
            endDateBE
          );
          donorOpWs = wsPage;
        }

        for (let i = 0; i < chunk.length; i++) {
          const item = chunk[i];
          const rIdx = 7 + i;
          const row = wsPage.getRow(rIdx);
          const dateBE = formatShortThaiDateBE(item.travel_date);
          const displayDate = dateBE && dateBE !== prevDateStr ? dateBE : null;
          if (dateBE) prevDateStr = dateBE;

          const { from, to } = getDirectionalRouteParts(
            item.route,
            item.departure_time || item.arrival_time,
            item.site_name || bucketProject
          );
          const pax =
            item.passenger_count !== null && item.passenger_count !== undefined && item.passenger_count !== ''
              ? Number(item.passenger_count)
              : null;

          row.getCell(2).value = displayDate;
          row.getCell(3).value = String(item.vehicle_plate || item.worker_name || '').trim() || null;
          row.getCell(4).value = from || null;
          row.getCell(5).value = formatExcelTimeDisplay(item.departure_time) || null;
          row.getCell(6).value = to || null;
          row.getCell(9).value = formatExcelTimeDisplay(item.arrival_time) || null;
          row.getCell(10).value = pax;
          row.getCell(12).value = String(item.note || '').trim() || null;

          newlySyncedRecords.push({
            sync_key: makeSyncKey(item),
            log_id: item.log_id || '',
            target_sheet: wsPage.name,
            target_row: rIdx,
            route: item.route || '',
            travel_date: item.travel_date || ''
          });
        }
        seqOffset += 30;
      }
    }

    appendSyncedKeysToWorkbook(wb, newlySyncedRecords);
    return {
      syncedCount: newlySyncedRecords.length,
      totalFlowed: syncedSet.size + newlySyncedRecords.length,
      summaryText: formatSyncSummaryText(newlySyncedRecords)
    };
  }

  globalScope.CkpExcelSync = {
    applyRowsToCopyWorkbookExcelJS,
    applyRowsToBillingWorkbookExcelJS,
    formatShortThaiDateBE,
    formatExcelTimeDisplay,
    splitRouteParts,
    getDirectionalRouteParts
  };
})(typeof window !== 'undefined' ? window : globalThis);
