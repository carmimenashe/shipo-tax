/*
 * ShipoExcel — builds the quarterly claim workbook in the exact layout of the
 * submitted template (Sheet "רבעון N YYYY", 62 columns A..BJ, 5 classification
 * slots per row) and reads such a workbook back.
 *
 * Needs ExcelJS (browser global or require('exceljs')) and ShipoCore.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./core.js'));
  else root.ShipoExcel = factory(root.ShipoCore);
})(typeof self !== 'undefined' ? self : globalThis, function (C) {
  'use strict';

  const HEADER_ROW = 3;
  const FIRST_DATA_ROW = 4;
  const LAST_COL = 62; // BJ

  // Row 1 of the template: which document each column is taken from.
  const SOURCE_ROW = {
    2: 'ר. ייצוא', 3: 'פרופורמה', 4: 'פרופורמה ', 5: 'פרופורמה', 6: 'ר. ייצוא', 7: 'פרופורמה',
    9: 'פרופורמה', 10: 'ש. מטען', 11: 'פרופורמה', 13: 'שטר מטען', 14: 'אין לנו', 15: 'שטר מטען',
    16: 'IM4', 17: 'IM4', 18: 'שטר מטען', 19: 'פרופורמה', 20: 'IM4', 23: 'IM4', 24: 'IM4', 25: 'IM4',
    26: 'IM4', 27: 'IM4',
  };
  // Row 2 of the template: running field numbers (orange) and two notes.
  const NUMBER_ROW = {
    2: 1, 3: 2, 4: 3, 5: 4, 6: 5, 7: 6, 9: 7, 10: 8, 11: 9, 12: 'אין לנו', 13: 10, 14: 'אווירי',
    15: 11, 16: 12, 17: 13, 18: 14, 19: 15, 20: 16, 23: 17, 24: 18, 25: 19, 26: 20, 27: 21, 61: 22,
  };
  const BASE_HEADERS = [
    'סידורי ', "מס' הצהרת יצוא", "מס' חשבון עסקה", 'תאריך חשבון', 'ערך משלוח כולל למכס', 'מטבע',
    'ערך יצוא זכאי לשיפוי ', 'מטבע', 'מונח המכר', 'משקל כולל ברוטו', 'משקל כולל נטו', 'מספר 1-EUR',
    "מס' שטר מטען", "מס' מכולה", 'תאריך טעינה', "מס' הצהרת יבוא IM4 ", 'תאריך שחרור הסחורה ',
    'מדינת יעד', 'נמל',
  ];
  const SLOT_HEADERS = (k) => [
    `סיווג ${k}\nזכאי לשיפוי`, `אתר גידול סיווג ${k}`, `מיקוד גידול סיווג ${k}`, `תיאור גידול סיווג ${k}`,
    `משקל נטו סיווג ${k}`, `ערך למכס סיווג ${k}`, `אחוז המכס סיווג ${k}`, `סכום המכס סיווג ${k}`,
  ];
  const TAIL_HEADERS = ['מטבע סכום שיפוי מבוקש', 'סה"כ סכום שיפוי מבוקש', 'מטבע סכום שיפוי מבוקש'];

  function headers() {
    const h = BASE_HEADERS.slice();
    for (let k = 1; k <= C.SLOTS_PER_ROW; k++) h.push(...SLOT_HEADERS(k));
    h.push(...TAIL_HEADERS);
    return h;
  }

  const slotCol = (slot, field) => 20 + slot * 8 + field; // slot 0..4, field 0..7
  const COL = {
    serial: 1, exportDecl: 2, invoiceNo: 3, invoiceDate: 4, invoiceTotal: 5, invoiceCurrency: 6,
    eligibleValue: 7, eligibleCurrency: 8, incoterm: 9, grossKg: 10, netKg: 11, eur1: 12, awb: 13,
    container: 14, loadingDate: 15, im4: 16, releaseDate: 17, destCountry: 18, port: 19,
    claimCurrency: 60, claimTotal: 61, claimCurrency2: 62,
  };

  const COL_WIDTHS = { 2: 16.6, 3: 15, 4: 9.5, 5: 12.4, 6: 6.1, 7: 11.1, 8: 6.1, 9: 9, 10: 9.9, 11: 9.9, 12: 9, 13: 14, 14: 9, 15: 9.5, 16: 11, 17: 9.5, 18: 9, 19: 9, 60: 9, 61: 11, 62: 9 };

  const FILL_GREY = 'FFDBDBDB'; // template: theme accent3 at 60% tint
  const FILL_GREEN = 'FFBAE6C6';
  const FILL_ORANGE = 'FFFFC000';
  const ACCOUNTING = '_ * #,##0.00_ ;_ * \\-#,##0.00_ ;_ * "-"??_ ;_ @_ ';
  const FONT = { name: 'Arial', size: 10 };
  const THIN = { style: 'thin' };
  const BORDER = { top: THIN, left: THIN, bottom: THIN, right: THIN };

  function colLetter(n) {
    let s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  const fill = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const asNumberIfDigits = (v, maxLen = 15) => (/^\d+$/.test(String(v)) && String(v).length <= maxLen && !/^0/.test(String(v)) ? Number(v) : v);
  const fmtNum = (n) => String(C.round2(n));

  function sheetName(year, q) {
    return year && q ? `רבעון ${q} ${year}` : 'שיפוי';
  }

  /** Split shipments into template rows (a shipment with >5 classifications continues on extra rows). */
  function toRows(shipments) {
    const rows = [];
    shipments.forEach((s, idx) => {
      const chunks = [];
      for (let i = 0; i < Math.max(1, s.classifications.length); i += C.SLOTS_PER_ROW) {
        chunks.push(s.classifications.slice(i, i + C.SLOTS_PER_ROW));
      }
      chunks.forEach((cls, part) => rows.push({ s, serial: idx + 1, part, parts: chunks.length, cls }));
    });
    return rows;
  }

  /**
   * meta: { year, q, signerName, signDate (ISO), company, issuesById: {id: [issues]} }
   */
  function buildWorkbook(ExcelJS, shipments, meta = {}) {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'shipo-tax';
    wb.created = new Date();
    const list = C.sortShipments(shipments.filter((s) => s.included !== false));
    buildMainSheet(wb, list, meta);
    buildSummarySheet(wb, list, meta);
    buildChecksSheet(wb, list, meta);
    return wb;
  }

  function buildMainSheet(wb, list, meta) {
    const ws = wb.addWorksheet(sheetName(meta.year, meta.q), {
      views: [{ rightToLeft: true }],
      pageSetup: { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
      properties: { defaultColWidth: 9 },
    });
    for (let c = 1; c <= LAST_COL; c++) {
      ws.getColumn(c).width = COL_WIDTHS[c] || (((c - 20) % 8 === 3 && c >= 20 && c < 60) ? 14 : 9);
    }

    // rows 1-2: source labels and field numbers, as in the template
    for (const [c, label] of Object.entries(SOURCE_ROW)) {
      const cell = ws.getCell(1, +c);
      cell.value = label;
      cell.font = { ...FONT, bold: true };
      cell.alignment = { horizontal: 'center' };
    }
    for (const [c, label] of Object.entries(NUMBER_ROW)) {
      const cell = ws.getCell(2, +c);
      cell.value = label;
      cell.font = { ...FONT, bold: true };
      cell.alignment = { horizontal: 'center' };
      if (typeof label === 'number') cell.fill = fill(FILL_ORANGE);
    }

    // row 3: headers
    const hdr = headers();
    ws.getRow(HEADER_ROW).height = 38.25;
    hdr.forEach((text, i) => {
      const c = i + 1;
      const cell = ws.getCell(HEADER_ROW, c);
      cell.value = text;
      cell.font = { ...FONT, bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      const slot = c >= 20 && c < 60 ? Math.floor((c - 20) / 8) : -1;
      cell.fill = fill(slot >= 0 && slot % 2 === 0 ? FILL_GREEN : FILL_GREY);
      cell.border = BORDER;
    });

    // data
    const rows = toRows(list);
    rows.forEach((r, i) => {
      const rowNo = FIRST_DATA_ROW + i;
      const s = r.s;
      const set = (col, value, numFmt) => {
        const cell = ws.getCell(rowNo, col);
        if (value !== null && value !== undefined && value !== '') cell.value = value;
        if (numFmt) cell.numFmt = numFmt;
        return cell;
      };
      const date = (iso) => (iso ? C.isoToUTCDate(iso) : null);

      if (r.part === 0) {
        set(COL.serial, r.serial);
        set(COL.exportDecl, s.exportDecl ? asNumberIfDigits(s.exportDecl) : '');
        set(COL.invoiceNo, s.invoiceNo);
        set(COL.invoiceDate, date(s.invoiceDate), 'dd.mm.yy');
        set(COL.invoiceTotal, s.invoiceTotal, ACCOUNTING);
        set(COL.invoiceCurrency, s.invoiceCurrency);
        if (s.eligibleParts.length > 1) {
          set(COL.eligibleValue, { formula: s.eligibleParts.map(fmtNum).join('+'), result: C.round2(C.sum(s.eligibleParts)) }, ACCOUNTING);
        } else {
          set(COL.eligibleValue, s.eligibleValue, ACCOUNTING);
        }
        set(COL.eligibleCurrency, s.eligibleCurrency);
        set(COL.incoterm, s.incoterm);
        set(COL.grossKg, s.grossKg, ACCOUNTING);
        set(COL.netKg, s.netKg, ACCOUNTING);
        set(COL.eur1, s.eur1);
        set(COL.awb, s.awb);
        set(COL.container, s.container);
        set(COL.loadingDate, date(s.loadingDate), 'dd.mm.yy');
        const short = C.im4Short(s.im4);
        set(COL.im4, short ? asNumberIfDigits(short) : '');
        set(COL.releaseDate, date(s.releaseDate), 'dd.mm.yy');
        set(COL.destCountry, s.destCountry);
        set(COL.port, s.port);
      } else {
        set(COL.serial, `${r.serial} (המשך)`);
        set(COL.exportDecl, s.exportDecl ? asNumberIfDigits(s.exportDecl) : '');
        set(COL.awb, s.awb);
      }

      const dutyRefs = [];
      r.cls.forEach((c, slot) => {
        set(slotCol(slot, 0), c.hs);
        set(slotCol(slot, 1), c.site);
        set(slotCol(slot, 2), c.zip);
        set(slotCol(slot, 3), c.desc);
        set(slotCol(slot, 4), c.netKg);
        set(slotCol(slot, 5), c.value);
        set(slotCol(slot, 6), c.ratePct === null ? null : C.round2(c.ratePct * 100) / 10000, '0.00%');
        set(slotCol(slot, 7), c.duty);
        if (c.duty !== null) dutyRefs.push(`${colLetter(slotCol(slot, 7))}${rowNo}`);
      });
      set(COL.claimCurrency, s.claimCurrency);
      const subtotal = C.round2(C.sum(r.cls.map((c) => c.duty)));
      if (dutyRefs.length > 1) set(COL.claimTotal, { formula: dutyRefs.join('+'), result: subtotal });
      else if (dutyRefs.length === 1) set(COL.claimTotal, { formula: dutyRefs[0], result: subtotal });
      set(COL.claimCurrency2, s.claimCurrency);

      for (let c = 1; c <= LAST_COL; c++) {
        const cell = ws.getCell(rowNo, c);
        cell.border = BORDER;
        cell.font = FONT;
        if (c === COL.serial) cell.alignment = { horizontal: 'center' };
      }
    });

    const last = FIRST_DATA_ROW + rows.length - 1;

    // signature block, 6 rows below the data like the template
    const sigRow = Math.max(last, HEADER_ROW) + 6;
    const label = (r, c, v) => { const cell = ws.getCell(r, c); cell.value = v; cell.font = FONT; };
    label(sigRow, 2, 'שם מורשה חתימה ');
    label(sigRow, 7, 'תאריך');
    label(sigRow, 9, 'חתימה (בצירוף חותמת)');
    if (meta.signerName) label(sigRow + 1, 2, meta.signerName);
    if (meta.signDate) label(sigRow + 1, 7, C.fmtDateLong(meta.signDate).replace(/\//g, '.'));
    return ws;
  }

  function buildSummarySheet(wb, list, meta) {
    const sum = C.summarize(list, meta);
    const ws = wb.addWorksheet('ריכוז רבעוני', { views: [{ rightToLeft: true }], pageSetup: { paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
    ws.columns = [{ width: 30 }, { width: 22 }, { width: 18 }, { width: 16 }, { width: 16 }, { width: 14 }, { width: 14 }];
    let r = 1;
    const title = ws.getCell(r, 1);
    title.value = `ריכוז בקשת שיפוי — ${C.quarterLabel(meta.year, meta.q)}`;
    title.font = { ...FONT, size: 14, bold: true };
    r += 2;
    const company = meta.company || {};
    const kv = (k, v, fmt) => {
      ws.getCell(r, 1).value = k;
      ws.getCell(r, 1).font = { ...FONT, bold: true };
      const cell = ws.getCell(r, 2);
      cell.value = v;
      cell.font = FONT;
      if (fmt) cell.numFmt = fmt;
      r += 1;
    };
    kv('שם החברה', company.name || '');
    kv('מספר ח.פ', company.id || '');
    kv('תקופה', sum.range ? `${C.fmtDateLong(sum.range.from)} – ${C.fmtDateLong(sum.range.to)}` : '');
    kv('מספר משלוחים', sum.count);
    kv('תאריכי משלוח', sum.firstDate ? `${C.fmtDateLong(sum.firstDate)} – ${C.fmtDateLong(sum.lastDate)}` : '');
    kv('יעד סופי', sum.countries.join(', '));
    kv('נמל/משרד מכס', sum.ports.join(', '));
    kv('ערך יצוא זכאי לשיפוי', C.fmtByCurrency(sum.eligible));
    kv('ערך למכס של הפריטים הזכאים', C.fmtByCurrency(sum.customsValue));
    kv('אחוזי המכס ששולמו', sum.rates.map(C.fmtPct).join(', '));
    kv('סה"כ סכום שיפוי מבוקש', C.fmtByCurrency(sum.duty));
    r += 1;

    const table = (headersRow, rowsData, fmts) => {
      headersRow.forEach((h, i) => {
        const cell = ws.getCell(r, i + 1);
        cell.value = h;
        cell.font = { ...FONT, bold: true };
        cell.fill = fill(FILL_GREY);
        cell.border = BORDER;
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      });
      r += 1;
      rowsData.forEach((row) => {
        row.forEach((v, i) => {
          const cell = ws.getCell(r, i + 1);
          cell.value = v;
          cell.font = FONT;
          cell.border = BORDER;
          if (fmts && fmts[i]) cell.numFmt = fmts[i];
        });
        r += 1;
      });
    };

    ws.getCell(r, 1).value = 'ריכוז לפי פרט מכס';
    ws.getCell(r, 1).font = { ...FONT, size: 12, bold: true };
    r += 1;
    table(['פרט מכס', 'תיאור', 'משלוחים', 'משקל נטו (ק"ג)', 'ערך למכס', 'אחוז מכס', 'סכום המכס'],
      sum.hsRows.map((h) => [h.hs, h.desc, h.shipments, h.netKg, h.value, h.rates.map(C.fmtPct).join(', '), h.duty]),
      [null, null, null, '#,##0.00', '#,##0.00', null, '#,##0.00']);
    const tot = r;
    ws.getCell(tot, 1).value = 'סה"כ';
    ws.getCell(tot, 1).font = { ...FONT, bold: true };
    ws.getCell(tot, 5).value = C.round2(C.sum(sum.hsRows.map((h) => h.value)));
    ws.getCell(tot, 7).value = C.round2(C.sum(sum.hsRows.map((h) => h.duty)));
    [5, 7].forEach((c) => { ws.getCell(tot, c).numFmt = '#,##0.00'; ws.getCell(tot, c).font = { ...FONT, bold: true }; });
    r += 2;

    ws.getCell(r, 1).value = 'ריכוז לפי חודש';
    ws.getCell(r, 1).font = { ...FONT, size: 12, bold: true };
    r += 1;
    table(['חודש', 'משלוחים', 'ערך למכס', 'סכום המכס'],
      sum.months.map((m) => [m.month === '—' ? '—' : `${m.month.slice(5)}/${m.month.slice(0, 4)}`, m.shipments, m.value, m.duty]),
      [null, null, '#,##0.00', '#,##0.00']);
    r += 1;

    ws.getCell(r, 1).value = 'פירוט משלוחים';
    ws.getCell(r, 1).font = { ...FONT, size: 12, bold: true };
    r += 1;
    table(['שטר מטען', "מס' הצהרת יצוא", 'תאריך טעינה', 'הצהרת יבוא IM4', 'פרטי מכס', 'ערך למכס', 'סכום המכס'],
      sum.shipments.map((s) => [s.awb, s.exportDecl, C.fmtDateLong(s.loadingDate), s.im4, s.classifications.map((c) => c.hs).join(', '), C.customsValueTotal(s), C.claimTotal(s)]),
      [null, null, null, null, null, '#,##0.00', '#,##0.00']);
    return ws;
  }

  function buildChecksSheet(wb, list, meta) {
    const ws = wb.addWorksheet('בקרות', { views: [{ rightToLeft: true }] });
    ws.columns = [{ width: 8 }, { width: 16 }, { width: 10 }, { width: 90 }];
    ['#', 'שטר מטען', 'רמה', 'הערה'].forEach((h, i) => {
      const cell = ws.getCell(1, i + 1);
      cell.value = h;
      cell.font = { ...FONT, bold: true };
      cell.fill = fill(FILL_GREY);
    });
    let r = 2;
    const levelName = { error: 'שגיאה', warn: 'אזהרה', info: 'מידע' };
    list.forEach((s, i) => {
      const issues = (meta.issuesById && meta.issuesById[s.id]) || C.validateShipment(s, { year: meta.year, q: meta.q, shipments: list });
      const notes = (s.notes || []).map((n) => ({ level: 'info', msg: n }));
      for (const it of [...issues, ...notes]) {
        ws.getCell(r, 1).value = i + 1;
        ws.getCell(r, 2).value = s.awb;
        ws.getCell(r, 3).value = levelName[it.level] || it.level;
        ws.getCell(r, 4).value = it.msg;
        ws.getCell(r, 4).alignment = { wrapText: true };
        r += 1;
      }
    });
    if (r === 2) ws.getCell(2, 4).value = 'לא נמצאו הערות';
    return ws;
  }

  /* ----------------------------------------------------------------- import */

  function cellVal(cell) {
    let v = cell && cell.value;
    if (v === null || v === undefined) return null;
    if (typeof v === 'object' && !(v instanceof Date)) {
      if ('result' in v) return v.result === undefined ? null : v.result;
      if ('richText' in v) return v.richText.map((t) => t.text).join('');
      if ('text' in v) return v.text;
      if ('error' in v) return null;
    }
    return v;
  }

  function cellFormula(cell) {
    const v = cell && cell.value;
    if (v && typeof v === 'object' && (v.formula || v.sharedFormula)) return v.formula || null;
    return null;
  }

  const clean = (s) => String(s || '').replace(/\s+/g, ' ').trim();

  /** Find the sheet and row that hold the template's header. */
  function findHeader(wb) {
    for (const ws of wb.worksheets) {
      for (let r = 1; r <= Math.min(15, ws.rowCount || 15); r++) {
        const b = clean(cellVal(ws.getCell(r, 2)));
        if (b.includes('הצהרת יצוא')) return { ws, row: r };
      }
    }
    return null;
  }

  /**
   * Reads a claim workbook (one we produced or the hand-filled template).
   * Returns { shipments, signerName, signDate, year, q, sheetName }
   */
  async function importWorkbook(ExcelJS, data) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(data);
    const found = findHeader(wb);
    if (!found) throw new Error('לא נמצאה בקובץ טבלת שיפוי (כותרת "מס\' הצהרת יצוא")');
    const { ws, row: headerRow } = found;

    // map header text -> column, so slightly different layouts still load
    const map = {};
    const hdrs = headers().map(clean);
    for (let c = 1; c <= LAST_COL + 5; c++) {
      const t = clean(cellVal(ws.getCell(headerRow, c)));
      if (!t) continue;
      const idx = hdrs.findIndex((h, i) => h === t && !(i in map));
      if (idx >= 0) map[idx] = c;
    }
    const at = (i) => map[i] || i + 1;

    const shipments = [];
    let signerName = '';
    let signDate = null;
    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const v = (i) => cellVal(ws.getCell(r, at(i)));
      const first = clean(v(0));
      const second = clean(v(1));
      if (second.startsWith('שם מורשה חתימה')) {
        signerName = clean(cellVal(ws.getCell(r + 1, at(1))));
        signDate = C.parseDate(cellVal(ws.getCell(r + 1, at(6))));
        break;
      }
      const awb = clean(v(12));
      const hasData = first || second || awb;
      if (!hasData) continue;

      const classifications = [];
      for (let slot = 0; slot < C.SLOTS_PER_ROW; slot++) {
        const g = (f) => v(19 + slot * 8 + f);
        const cls = {
          hs: C.normHs(g(0)), site: clean(g(1)), zip: clean(g(2)), desc: clean(g(3)),
          netKg: C.num(g(4)), value: C.num(g(5)), ratePct: null, duty: C.num(g(7)),
        };
        const rate = C.num(g(6));
        if (rate !== null) cls.ratePct = C.round2(rate <= 1 ? rate * 100 : rate);
        if (cls.hs || cls.desc || cls.duty !== null) classifications.push(cls);
      }

      if (/המשך/.test(first) && shipments.length) {
        const prev = shipments[shipments.length - 1];
        prev.classifications.push(...classifications);
        const more = C.num(v(60));
        if (prev.fileClaimTotal !== null && more !== null) prev.fileClaimTotal = C.round2(prev.fileClaimTotal + more);
        continue;
      }

      const eligCell = ws.getCell(r, at(6));
      const f = cellFormula(eligCell);
      let eligibleParts = [];
      if (f && /^[\d.+\s]+$/.test(f)) eligibleParts = f.split('+').map(C.num).filter((n) => n !== null);
      const im4 = v(15);

      shipments.push(C.normalizeShipment({
        exportDecl: v(1) === null ? '' : String(v(1)),
        invoiceNo: v(2), invoiceDate: v(3), invoiceTotal: v(4), invoiceCurrency: v(5),
        eligibleValue: v(6), eligibleParts, eligibleCurrency: v(7), incoterm: v(8),
        grossKg: v(9), netKg: v(10), eur1: v(11), awb, container: v(13), loadingDate: v(14),
        im4: im4 === null ? '' : String(im4), releaseDate: v(16), destCountry: v(17), port: v(18),
        classifications, claimCurrency: v(59) || v(61), source: 'excel',
        fileClaimTotal: v(60),
      }));
    }
    const m = ws.name.match(/(\d)\D+(\d{4})/) || ws.name.match(/(\d{4})\D+(\d)/);
    let year = null;
    let q = null;
    if (m) {
      if (m[1].length === 4) { year = +m[1]; q = +m[2]; } else { q = +m[1]; year = +m[2]; }
    }
    if (!(q >= 1 && q <= 4)) {
      const guess = C.suggestQuarter(shipments);
      year = guess ? guess.year : null;
      q = guess ? guess.q : null;
    }
    return { shipments, signerName, signDate, year, q, sheetName: ws.name };
  }

  return { buildWorkbook, importWorkbook, headers, sheetName, toRows, colLetter, COL, SOURCE_ROW, NUMBER_ROW };
});
