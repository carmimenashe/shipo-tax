import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { existsSync, readFileSync } from 'node:fs';

const require = createRequire(import.meta.url);
const C = require('../app/js/core.js');
const X = require('../app/js/excel.js');
const ExcelJS = require('exceljs');

test('number parsing handles both decimal styles', () => {
  assert.equal(C.num('24,991.15'), 24991.15);
  assert.equal(C.num('257,13'), 257.13);
  assert.equal(C.num('1.234,56'), 1234.56);
  assert.equal(C.num('US$ 7,367.78'), 7367.78);
  assert.equal(C.num('5,283'), 5283);
  assert.equal(C.num(''), null);
  assert.equal(C.num({ formula: 'A1', result: 12 }), 12);
});

test('dates and quarters', () => {
  assert.equal(C.parseDate('18.12.24'), '2024-12-18');
  assert.equal(C.parseDate('19/12/2024'), '2024-12-19');
  assert.equal(C.parseDate(45644), '2024-12-18');
  assert.deepEqual(C.quarterOf('2024-12-18'), { year: 2024, q: 4 });
  assert.deepEqual(C.quarterRange(2024, 4), { from: '2024-10-01', to: '2024-12-31' });
  assert.deepEqual(C.quarterRange(2025, 1), { from: '2025-01-01', to: '2025-03-31' });
});

test('identifiers', () => {
  assert.equal(C.normHs(910993300), '0910993300');
  assert.equal(C.normHs('09109933 00'), '0910993300');
  assert.equal(C.normAwb('700-12345678'), '700 12345678');
  assert.equal(C.im4Short('24BEI0000012345678'), '12345678');
  assert.equal(C.im4Short('12345678'), '12345678');
  assert.equal(C.joinInvoiceNumbers(['WT24001234/1', 'WT24001234/2']), 'WT24001234/1/2');
  assert.equal(C.normCurrency('ERU'), 'EUR');
});

test('validation flags duty arithmetic, missing fields and quarter', () => {
  const s = C.normalizeShipment({
    awb: '700 12345678', exportDecl: '300012345', invoiceNo: 'WT1', invoiceDate: '2024-12-18', im4: 'X', releaseDate: '2024-12-19', destCountry: 'בלגיה',
    loadingDate: '2025-01-02', dutyDocTotal: 100,
    classifications: [{ hs: '0709999090', value: 654.54, ratePct: 12.8, duty: 122.18 }],
  });
  const issues = C.validateShipment(s, { year: 2024, q: 4, shipments: [s] });
  const msgs = issues.map((i) => i.msg).join('\n');
  assert.match(msgs, /83\.78/); // 654.54 × 12.8% = 83.78
  assert.match(msgs, /שונה מסכום המכס במסמכים/);
  assert.match(msgs, /אינו ברבעון 4 2024/);
});

test('page grouping links pages by AWB / invoice / declaration and shares payment allocations', () => {
  const pages = [
    { fileIndex: 0, fileName: 'a.pdf', page: 1, type: 'application_form', keys: {} },
    { fileIndex: 0, fileName: 'a.pdf', page: 2, type: 'air_waybill', keys: { awb_numbers: ['700-12345678'] } },
    { fileIndex: 0, fileName: 'a.pdf', page: 3, type: 'duty_invoice', keys: { awb_numbers: ['70012345678'], invoice_numbers: ['WT24001234/1'] } },
    { fileIndex: 1, fileName: 'b.pdf', page: 1, type: 'packing_list', keys: { invoice_numbers: ['WT24001234/2'] } },
    { fileIndex: 2, fileName: 'c.pdf', page: 1, type: 'air_waybill', keys: { awb_numbers: ['700 12345690'] } },
    { fileIndex: 3, fileName: 'pay.pdf', page: 1, type: 'duty_payment', keys: { awb_numbers: ['700-1234-5678', '700-1234-5690'] } },
  ];
  const groups = C.groupPages(pages);
  assert.equal(groups.length, 2);
  const g1 = groups.find((g) => g.awbs.includes('70012345678'));
  assert.equal(g1.pages.length, 5); // 3 from a.pdf, packing list via invoice base, shared payment
  const g2 = groups.find((g) => g.awbs.includes('70012345690'));
  assert.equal(g2.pages.length, 2);
});

test('Excel round trip keeps every field', async () => {
  const data = C.exampleData();
  const wb = X.buildWorkbook(ExcelJS, data.shipments, { year: 2024, q: 4, signerName: 'ישראל ישראלי', signDate: '2025-01-20', company: data.settings.company });
  const buf = await wb.xlsx.writeBuffer();
  const back = await X.importWorkbook(ExcelJS, buf);
  assert.equal(back.year, 2024);
  assert.equal(back.q, 4);
  assert.equal(back.signerName, 'ישראל ישראלי');
  assert.equal(back.signDate, '2025-01-20');
  assert.equal(back.shipments.length, data.shipments.length);
  const a = C.sortShipments(data.shipments);
  back.shipments.forEach((s, i) => {
    const o = a[i];
    for (const f of ['exportDecl', 'invoiceNo', 'invoiceDate', 'invoiceTotal', 'invoiceCurrency', 'eligibleValue', 'incoterm', 'grossKg', 'netKg', 'awb', 'loadingDate', 'releaseDate', 'destCountry', 'port', 'claimCurrency']) {
      assert.deepEqual(s[f], o[f], `${f} of row ${i + 1}`);
    }
    assert.equal(s.im4, C.im4Short(o.im4));
    assert.deepEqual(s.eligibleParts, o.eligibleParts.length > 1 ? o.eligibleParts : []);
    assert.deepEqual(s.classifications, o.classifications);
  });
});

test('more than five classifications continue on an extra row', async () => {
  const cls = Array.from({ length: 7 }, (_, i) => ({ hs: `07099990${i}0`, desc: `P${i}`, value: 100, ratePct: 10, duty: 10 }));
  const s = C.normalizeShipment({ awb: '700 11111111', exportDecl: '123456789', loadingDate: '2024-10-01', classifications: cls });
  const wb = X.buildWorkbook(ExcelJS, [s], { year: 2024, q: 4 });
  const ws = wb.worksheets[0];
  assert.equal(ws.getCell('A5').value, '1 (המשך)');
  const back = await X.importWorkbook(ExcelJS, await wb.xlsx.writeBuffer());
  assert.equal(back.shipments.length, 1);
  assert.equal(back.shipments[0].classifications.length, 7);
});

// Optional: the real hand-filled template, kept outside the repo (private data).
const REAL = process.env.SHIPO_TEMPLATE;
test('imports the hand-filled template', { skip: !REAL || !existsSync(REAL) }, async () => {
  const res = await X.importWorkbook(ExcelJS, readFileSync(REAL));
  assert.ok(res.shipments.length > 0);
  assert.ok(res.year >= 2020 && res.q >= 1 && res.q <= 4);
  for (const s of res.shipments) {
    assert.ok(s.awb, 'awb');
    assert.ok(s.classifications.length > 0, 'classifications');
    assert.equal(s.claimCurrency, 'EUR');
  }
  console.log(JSON.stringify({ n: res.shipments.length, year: res.year, q: res.q, signer: !!res.signerName, first: res.shipments[0] }, null, 1).slice(0, 1500));
});

test('invoice basis: main invoice by default, sum on request, typed values kept', () => {
  const x = {
    awb_number: '700 12345678', invoice_numbers: ['WT24001234/2', 'WT24001234/1'], invoice_total: 99, net_weight_kg: 1,
    invoices: [
      { number: 'WT24001234/2', total: 1259.48, currency: 'USD', net_weight_kg: 32 },
      { number: 'WT24001234/1', total: 24991.15, currency: 'USD', net_weight_kg: 3291 },
    ],
  };
  const s = C.fromExtraction(x, []);
  C.applyInvoiceBasis(s, 'first');
  assert.equal(s.invoiceTotal, 24991.15);
  assert.equal(s.netKg, 3291);
  assert.equal(s.invoiceNo, 'WT24001234/1/2');
  C.applyInvoiceBasis(s, 'all');
  assert.equal(s.invoiceTotal, 26250.63);
  assert.equal(s.netKg, 3323);
  s.totalsEdited = true;
  s.invoiceTotal = 1;
  C.applyInvoiceBasis(s, 'first');
  assert.equal(s.invoiceTotal, 1);
  // no invoice list (e.g. loaded from Excel): values untouched
  const t = C.normalizeShipment({ invoiceTotal: 500, netKg: 10 });
  C.applyInvoiceBasis(t, 'all');
  assert.equal(t.invoiceTotal, 500);
  // survives a save/load round trip
  const back = C.normalizeShipment(JSON.parse(JSON.stringify(s)));
  assert.equal(back.invoices.length, 2);
  assert.equal(back.totalsEdited, true);
});
