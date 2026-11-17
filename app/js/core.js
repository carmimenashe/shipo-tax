/*
 * ShipoCore — data model, normalisation, validation and quarterly consolidation
 * for exporter-indemnity claims ("הקרן לשיפוי יצואנים בגין פגיעה בהסכמי סחר").
 *
 * Pure functions only (no DOM), so it runs in the browser and under Node tests.
 */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.ShipoCore = api;
})(typeof self !== 'undefined' ? self : globalThis, function () {
  'use strict';

  const SLOTS_PER_ROW = 5; // the template has room for 5 duty classifications per row

  /* ------------------------------------------------------------------ numbers */

  /** Parse "24,991.15", "257,13", "US$ 7,367.78", 0.07 … into a Number, or null. */
  function num(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return Number.isFinite(v) ? v : null;
    if (typeof v === 'object' && v && 'result' in v) return num(v.result);
    let s = String(v).trim();
    if (!s) return null;
    const neg = /^\(.*\)$/.test(s) || /^-/.test(s.replace(/[^\d.,-]/g, ''));
    s = s.replace(/[^\d.,]/g, '');
    if (!s) return null;
    const lastDot = s.lastIndexOf('.');
    const lastComma = s.lastIndexOf(',');
    if (lastDot >= 0 && lastComma >= 0) {
      // both present: the later one is the decimal separator
      if (lastComma > lastDot) s = s.replace(/\./g, '').replace(',', '.');
      else s = s.replace(/,/g, '');
    } else if (lastComma >= 0) {
      // only commas: thousands if groups of 3, else decimal comma
      if (/^\d{1,3}(,\d{3})+$/.test(s)) s = s.replace(/,/g, '');
      else s = s.replace(/,/g, '.');
    } else if ((s.match(/\./g) || []).length > 1) {
      // 1.234.567 -> thousands dots
      s = s.replace(/\./g, '');
    }
    const n = parseFloat(s);
    if (!Number.isFinite(n)) return null;
    return neg ? -n : n;
  }

  const round2 = (n) => (n === null || n === undefined ? null : Math.round((n + Number.EPSILON) * 100) / 100);
  const sum = (arr) => arr.reduce((a, b) => a + (Number.isFinite(b) ? b : 0), 0);

  function fmtMoney(n, digits = 2) {
    if (n === null || n === undefined || !Number.isFinite(n)) return '';
    return n.toLocaleString('en-US', { minimumFractionDigits: digits, maximumFractionDigits: digits });
  }

  function fmtPct(p) {
    if (p === null || p === undefined || !Number.isFinite(p)) return '';
    return `${+p.toFixed(2)}%`;
  }

  /* -------------------------------------------------------------------- dates */

  const pad2 = (n) => String(n).padStart(2, '0');

  /** Accepts ISO, dd.mm.yy, dd/mm/yyyy, dd-mm-yyyy, Date, Excel serial → "YYYY-MM-DD" or null. */
  function parseDate(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) {
      if (isNaN(v)) return null;
      return `${v.getUTCFullYear()}-${pad2(v.getUTCMonth() + 1)}-${pad2(v.getUTCDate())}`;
    }
    if (typeof v === 'object' && v && 'result' in v) return parseDate(v.result);
    if (typeof v === 'number') {
      if (v > 20000 && v < 80000) { // Excel serial date
        const d = new Date(Math.round((v - 25569) * 86400000));
        return parseDate(d);
      }
      return null;
    }
    const s = String(v).trim();
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return valid(+m[1], +m[2], +m[3]);
    m = s.match(/^(\d{1,2})[./\-\s](\d{1,2})[./\-\s](\d{2,4})/);
    if (m) {
      let y = +m[3];
      if (y < 100) y += 2000;
      return valid(y, +m[2], +m[1]);
    }
    return null;
    function valid(y, mo, d) {
      if (mo < 1 || mo > 12 || d < 1 || d > 31 || y < 1990 || y > 2100) return null;
      return `${y}-${pad2(mo)}-${pad2(d)}`;
    }
  }

  /** "2024-12-18" → "18.12.24" (the template's style) */
  function fmtDateShort(iso) {
    if (!iso) return '';
    const [y, m, d] = iso.split('-');
    return `${d}.${m}.${y.slice(2)}`;
  }

  /** "2024-12-18" → "18/12/2024" */
  function fmtDateLong(iso) {
    if (!iso) return '';
    const [y, m, d] = iso.split('-');
    return `${d}/${m}/${y}`;
  }

  function isoToUTCDate(iso) {
    if (!iso) return null;
    const [y, m, d] = iso.split('-').map(Number);
    return new Date(Date.UTC(y, m - 1, d));
  }

  /* ----------------------------------------------------------------- quarters */

  const ROMAN = ['I', 'II', 'III', 'IV'];

  function quarterOf(iso) {
    if (!iso) return null;
    const [y, m] = iso.split('-').map(Number);
    return { year: y, q: Math.floor((m - 1) / 3) + 1 };
  }

  function quarterRange(year, q) {
    const startMonth = (q - 1) * 3 + 1;
    const endMonth = startMonth + 2;
    const lastDay = new Date(Date.UTC(year, endMonth, 0)).getUTCDate();
    return { from: `${year}-${pad2(startMonth)}-01`, to: `${year}-${pad2(endMonth)}-${pad2(lastDay)}` };
  }

  function quarterLabel(year, q) {
    return `רבעון ${q} ${year}`;
  }

  /** The date that decides which quarter a shipment belongs to. */
  function shipmentDate(s) {
    return s.loadingDate || s.invoiceDate || s.releaseDate || null;
  }

  /** Most common quarter among the shipments (ties → latest). */
  function suggestQuarter(shipments) {
    const counts = new Map();
    for (const s of shipments) {
      const qq = quarterOf(shipmentDate(s));
      if (!qq) continue;
      const k = qq.year * 10 + qq.q;
      counts.set(k, (counts.get(k) || 0) + 1);
    }
    if (!counts.size) return null;
    const best = [...counts.entries()].sort((a, b) => b[1] - a[1] || b[0] - a[0])[0][0];
    return { year: Math.floor(best / 10), q: best % 10 };
  }

  /* ------------------------------------------------------------ identifiers */

  /** HS code → digits only; a 9-digit code lost its leading zero in Excel, put it back. */
  function normHs(v) {
    if (v === null || v === undefined) return '';
    let s = String(typeof v === 'object' && 'result' in v ? v.result : v).replace(/\D/g, '');
    if (s.length === 9 || s.length === 7 || s.length === 5) s = '0' + s;
    return s;
  }

  /** "70012345678", "700-1234 5678" → "700 12345678" */
  function normAwb(v) {
    if (!v) return '';
    const d = String(v).replace(/\D/g, '');
    if (d.length === 11) return `${d.slice(0, 3)} ${d.slice(3)}`;
    return String(v).trim();
  }

  const awbKey = (v) => String(v || '').replace(/\D/g, '');

  /** MRN "24BEI0000012345678" → "12345678" (significant digits, as used in the submitted Excel). */
  function im4Short(v) {
    if (!v) return '';
    const s = String(v).replace(/\s/g, '');
    const m = s.match(/(\d+)$/);
    if (!m) return s;
    if (s.length <= 10) return s; // already short
    return m[1].replace(/^0+/, '') || s;
  }

  function normCurrency(v) {
    if (!v) return '';
    let s = String(v).trim().toUpperCase();
    if (s === 'ERU' || s === '€' || s === 'EURO' || s === 'יורו') s = 'EUR';
    if (s === 'US$' || s === '$' || s === 'דולר') s = 'USD';
    return s.replace(/[^A-Z]/g, '').slice(0, 3);
  }

  /** ["WT24001234/1","WT24001234/2"] → "WT24001234/1/2" */
  function joinInvoiceNumbers(list) {
    const items = (list || []).map((x) => String(x).trim()).filter(Boolean);
    if (!items.length) return '';
    const groups = new Map();
    const order = [];
    for (const it of items) {
      const m = it.match(/^(.*?)\/(\d+)$/);
      const base = m ? m[1] : it;
      const part = m ? m[2] : null;
      if (!groups.has(base)) { groups.set(base, []); order.push(base); }
      if (part && !groups.get(base).includes(part)) groups.get(base).push(part);
    }
    return order.map((b) => {
      const parts = groups.get(b).sort((x, y) => x - y);
      return parts.length ? `${b}/${parts.join('/')}` : b;
    }).join(', ');
  }

  /* ------------------------------------------------------------ shipment model */

  let idCounter = 0;
  const newId = () => `s${Date.now().toString(36)}${(idCounter++).toString(36)}`;

  function emptyClassification() {
    return { hs: '', site: '', zip: '', desc: '', netKg: null, value: null, ratePct: null, duty: null };
  }

  function emptyShipment() {
    return {
      id: newId(),
      exportDecl: '', invoiceNo: '', invoiceDate: null, invoiceTotal: null, invoiceCurrency: '',
      invoices: [], totalsEdited: false,
      eligibleValue: null, eligibleParts: [], eligibleCurrency: '',
      incoterm: '', grossKg: null, netKg: null, eur1: '', awb: '', container: '',
      loadingDate: null, im4: '', releaseDate: null, destCountry: '', port: '',
      classifications: [], claimCurrency: 'EUR',
      dutyDocTotal: null, dutyEvidence: '', fileClaimTotal: null,
      notes: [], files: [], source: 'manual', included: true,
    };
  }

  function normClassification(c) {
    const out = emptyClassification();
    if (!c) return out;
    out.hs = normHs(c.hs ?? c.hs_code);
    out.site = String(c.site ?? c.growing_site ?? '').trim();
    out.zip = String(c.zip ?? c.growing_zip ?? '').trim();
    out.desc = String(c.desc ?? c.description ?? '').trim();
    out.netKg = num(c.netKg ?? c.net_weight_kg);
    out.value = num(c.value ?? c.customs_value);
    let rate = c.ratePct ?? c.duty_rate_percent;
    rate = num(rate);
    out.ratePct = rate;
    out.duty = num(c.duty ?? c.duty_amount);
    return out;
  }

  /** Bring any partial/raw object (AI output, Excel row, saved JSON) to the canonical shape. */
  function normalizeShipment(raw) {
    const s = Object.assign(emptyShipment(), raw && raw.id ? { id: raw.id } : {});
    if (!raw) return s;
    const str = (v) => (v === null || v === undefined ? '' : String(v).trim());
    s.exportDecl = str(raw.exportDecl).replace(/\s/g, '');
    s.invoiceNo = str(raw.invoiceNo);
    s.invoiceDate = parseDate(raw.invoiceDate);
    s.invoiceTotal = num(raw.invoiceTotal);
    s.invoiceCurrency = normCurrency(raw.invoiceCurrency);
    s.eligibleParts = (raw.eligibleParts || []).map(num).filter((n) => n !== null);
    s.eligibleValue = num(raw.eligibleValue);
    if (s.eligibleValue === null && s.eligibleParts.length) s.eligibleValue = round2(sum(s.eligibleParts));
    s.eligibleCurrency = normCurrency(raw.eligibleCurrency) || s.invoiceCurrency;
    s.incoterm = str(raw.incoterm).toUpperCase().replace(/^CFR$/, 'C&F').replace(/^CNF$/, 'C&F');
    s.grossKg = num(raw.grossKg);
    s.netKg = num(raw.netKg);
    s.eur1 = str(raw.eur1);
    s.awb = normAwb(raw.awb);
    s.container = str(raw.container);
    s.loadingDate = parseDate(raw.loadingDate);
    s.im4 = str(raw.im4).replace(/\s/g, '');
    s.releaseDate = parseDate(raw.releaseDate);
    s.destCountry = str(raw.destCountry);
    s.port = str(raw.port).toUpperCase();
    s.classifications = (raw.classifications || []).map(normClassification)
      .filter((c) => c.hs || c.desc || c.duty !== null || c.value !== null);
    s.claimCurrency = normCurrency(raw.claimCurrency) || 'EUR';
    s.dutyDocTotal = num(raw.dutyDocTotal);
    s.dutyEvidence = str(raw.dutyEvidence);
    s.notes = Array.isArray(raw.notes) ? raw.notes.map(String).filter(Boolean) : [];
    s.files = Array.isArray(raw.files) ? raw.files.slice() : [];
    s.source = raw.source || 'manual';
    s.included = raw.included !== false;
    s.fileClaimTotal = num(raw.fileClaimTotal);
    s.invoices = (raw.invoices || []).map((inv) => ({
      no: String(inv.no ?? inv.number ?? '').trim(),
      total: num(inv.total),
      currency: normCurrency(inv.currency),
      netKg: num(inv.netKg ?? inv.net_weight_kg),
    })).filter((inv) => inv.no || inv.total !== null);
    s.totalsEdited = !!raw.totalsEdited;
    s.edited = !!raw.edited;
    return s;
  }

  /**
   * Convert the extraction model's JSON (snake_case schema in prompts) to a shipment.
   */
  function fromExtraction(x, files) {
    if (!x) return normalizeShipment(null);
    const raw = {
      exportDecl: x.export_declaration_number,
      invoiceNo: Array.isArray(x.invoice_numbers) ? joinInvoiceNumbers(x.invoice_numbers) : x.invoice_numbers,
      invoiceDate: x.invoice_date,
      invoiceTotal: x.invoice_total,
      invoices: x.invoices,
      invoiceCurrency: x.invoice_currency,
      eligibleParts: x.eligible_export_value_parts,
      eligibleValue: x.eligible_export_value,
      eligibleCurrency: x.eligible_export_value_currency || x.invoice_currency,
      incoterm: x.incoterm,
      grossKg: x.gross_weight_kg,
      netKg: x.net_weight_kg,
      eur1: x.eur1_number,
      awb: x.awb_number,
      container: x.container_number,
      loadingDate: x.loading_date,
      im4: x.im4_number,
      releaseDate: x.release_date,
      destCountry: x.destination_country,
      port: x.port,
      classifications: x.classifications,
      claimCurrency: x.duty_currency,
      dutyDocTotal: x.duty_total_on_documents,
      dutyEvidence: x.duty_payment_evidence,
      notes: x.issues,
      files,
      source: 'ai',
    };
    return normalizeShipment(raw);
  }

  /** Main invoice first: WT24001234/1 before WT24001234/2. */
  function sortedInvoices(s) {
    const part = (no) => { const m = String(no).match(/\/(\d+)$/); return m ? +m[1] : 0; };
    return s.invoices.slice().sort((a, b) => String(a.no).replace(/\/\d+$/, '').localeCompare(String(b.no).replace(/\/\d+$/, '')) || part(a.no) - part(b.no));
  }

  /**
   * "ערך משלוח כולל למכס" and the net weight, from the shipment's duty invoices:
   *   basis "first" — the main invoice only (as in the Excel that was filed)
   *   basis "all"   — the sum of all the shipment's duty invoices
   * Leaves the values alone when there is no invoice list or the user typed them.
   */
  function applyInvoiceBasis(s, basis = 'first') {
    if (!s.invoices || !s.invoices.length || s.totalsEdited) return s;
    const list = sortedInvoices(s);
    if (basis === 'all') {
      const cur = list[0].currency || s.invoiceCurrency;
      const same = list.filter((i) => !i.currency || i.currency === cur);
      s.invoiceTotal = same.some((i) => i.total !== null) ? round2(sum(same.map((i) => i.total))) : s.invoiceTotal;
      s.netKg = list.some((i) => i.netKg !== null) ? round2(sum(list.map((i) => i.netKg))) : s.netKg;
      if (cur) s.invoiceCurrency = cur;
    } else {
      const first = list[0];
      if (first.total !== null) s.invoiceTotal = first.total;
      if (first.netKg !== null) s.netKg = first.netKg;
      if (first.currency) s.invoiceCurrency = first.currency;
    }
    return s;
  }

  const claimTotal = (s) => round2(sum(s.classifications.map((c) => c.duty)));
  const customsValueTotal = (s) => round2(sum(s.classifications.map((c) => c.value)));

  /* -------------------------------------------------------------- validation */

  /**
   * Returns [{level:'error'|'warn'|'info', field, msg}] in Hebrew.
   * ctx: { year, q, shipments } — the quarter being prepared and all rows (for duplicates).
   */
  function validateShipment(s, ctx = {}) {
    const issues = [];
    const add = (level, field, msg) => issues.push({ level, field, msg });

    const required = [
      ['exportDecl', "מס' הצהרת יצוא"], ['invoiceNo', "מס' חשבון עסקה"], ['invoiceDate', 'תאריך חשבון'],
      ['awb', "מס' שטר מטען"], ['im4', "מס' הצהרת יבוא IM4"], ['releaseDate', 'תאריך שחרור הסחורה'],
      ['destCountry', 'מדינת יעד'],
    ];
    for (const [f, label] of required) if (!s[f]) add('error', f, `חסר: ${label}`);
    if (!s.classifications.length) add('error', 'classifications', 'אין סיווג זכאי לשיפוי (לא נמצא מכס ששולם בהצהרת היבוא)');

    if (s.exportDecl && !/^\d{6,15}$/.test(s.exportDecl)) add('warn', 'exportDecl', "מס' הצהרת היצוא אינו מספרי או באורך חריג");

    s.classifications.forEach((c, i) => {
      const n = i + 1;
      if (!c.hs) add('error', `c${i}.hs`, `סיווג ${n}: חסר פרט מכס`);
      else if (!/^\d{8,10}$/.test(c.hs)) add('warn', `c${i}.hs`, `סיווג ${n}: פרט המכס ${c.hs} אינו בן 8–10 ספרות`);
      if (c.duty === null) add('error', `c${i}.duty`, `סיווג ${n}: חסר סכום מכס`);
      if (c.value === null) add('warn', `c${i}.value`, `סיווג ${n}: חסר ערך למכס`);
      if (c.ratePct === null) add('warn', `c${i}.ratePct`, `סיווג ${n}: חסר אחוז מכס`);
      if (c.value !== null && c.ratePct !== null && c.duty !== null) {
        const expected = c.value * c.ratePct / 100;
        const tol = Math.max(0.05, expected * 0.01);
        if (Math.abs(expected - c.duty) > tol) {
          add('warn', `c${i}.duty`, `סיווג ${n}: ${fmtMoney(c.value)} × ${fmtPct(c.ratePct)} = ${fmtMoney(expected)}, אך סכום המכס הוא ${fmtMoney(c.duty)}`);
        }
      }
      if (!c.zip && !c.site) add('info', `c${i}.site`, `סיווג ${n}: לא צוין אתר/מיקוד גידול`);
    });

    const total = claimTotal(s);
    if (s.dutyDocTotal !== null && s.classifications.length && Math.abs(s.dutyDocTotal - total) > 0.05) {
      add('warn', 'dutyDocTotal', `סכום הסיווגים (${fmtMoney(total)}) שונה מסכום המכס במסמכים (${fmtMoney(s.dutyDocTotal)})`);
    }
    if (s.fileClaimTotal !== null && s.classifications.length && Math.abs(s.fileClaimTotal - total) > 0.05) {
      add('warn', 'claimTotal', `בקובץ שנטען הסה"כ היה ${fmtMoney(s.fileClaimTotal)}, אבל סכום הסיווגים הוא ${fmtMoney(total)} (ייתכן שסיווג לא נכלל בנוסחה)`);
    }
    if (s.eligibleParts.length && s.eligibleValue !== null && Math.abs(sum(s.eligibleParts) - s.eligibleValue) > 0.05) {
      add('warn', 'eligibleValue', 'ערך היצוא הזכאי אינו שווה לסכום השורות שלו');
    }
    const shipmentTotal = s.invoices.length ? round2(sum(s.invoices.filter((i) => !i.currency || i.currency === s.invoiceCurrency).map((i) => i.total))) : s.invoiceTotal;
    if (s.eligibleValue !== null && shipmentTotal !== null && s.eligibleCurrency === s.invoiceCurrency && s.eligibleValue > shipmentTotal + 0.01) {
      add('warn', 'eligibleValue', 'ערך היצוא הזכאי גבוה מערך המשלוח הכולל');
    }
    if (s.grossKg !== null && s.netKg !== null && s.netKg > s.grossKg) add('warn', 'netKg', 'המשקל הנטו גבוה מהברוטו');
    if (s.loadingDate && s.releaseDate && s.releaseDate < s.loadingDate) add('warn', 'releaseDate', 'תאריך השחרור מוקדם מתאריך הטעינה');

    if (ctx.year && ctx.q) {
      const d = shipmentDate(s);
      const qq = quarterOf(d);
      if (qq && (qq.year !== ctx.year || qq.q !== ctx.q)) {
        add('warn', 'loadingDate', `תאריך המשלוח (${fmtDateLong(d)}) אינו ב${quarterLabel(ctx.year, ctx.q)}`);
      }
    }
    if (ctx.shipments) {
      // the same goods description under a different customs item elsewhere in the quarter is usually a typo
      const codesByDesc = new Map();
      ctx.shipments.filter((o) => o.included !== false).forEach((o) => o.classifications.forEach((c) => {
        const d = (c.desc || '').toUpperCase().replace(/\s+/g, ' ').trim();
        if (!d || !c.hs) return;
        if (!codesByDesc.has(d)) codesByDesc.set(d, new Map());
        const m = codesByDesc.get(d);
        m.set(c.hs, (m.get(c.hs) || 0) + 1);
      }));
      s.classifications.forEach((c, i) => {
        const d = (c.desc || '').toUpperCase().replace(/\s+/g, ' ').trim();
        const m = codesByDesc.get(d);
        if (!m || m.size < 2 || !c.hs) return;
        const mine = m.get(c.hs) || 0;
        const [topCode, topCount] = [...m.entries()].sort((a, b) => b[1] - a[1])[0];
        if (topCode !== c.hs && topCount > mine) {
          add('warn', `c${i}.hs`, `סיווג ${i + 1}: "${c.desc}" מופיע ברבעון ${topCount} פעמים עם פרט מכס ${topCode}, וכאן ${c.hs}`);
        }
      });
      const k = awbKey(s.awb);
      const dupAwb = k && ctx.shipments.some((o) => o !== s && o.id !== s.id && o.included !== false && awbKey(o.awb) === k);
      if (dupAwb) add('error', 'awb', `שטר המטען ${s.awb} מופיע יותר מפעם אחת`);
      const dupDecl = s.exportDecl && ctx.shipments.some((o) => o !== s && o.id !== s.id && o.included !== false && o.exportDecl === s.exportDecl);
      if (dupDecl) add('warn', 'exportDecl', `הצהרת היצוא ${s.exportDecl} מופיעה יותר מפעם אחת`);
    }
    return issues;
  }

  function issueCounts(issues) {
    return {
      error: issues.filter((i) => i.level === 'error').length,
      warn: issues.filter((i) => i.level === 'warn').length,
      info: issues.filter((i) => i.level === 'info').length,
    };
  }

  /* ------------------------------------------------------------ consolidation */

  /** Sort by shipment date, then AWB. */
  function sortShipments(list) {
    return list.slice().sort((a, b) => {
      const da = shipmentDate(a) || '9999';
      const db = shipmentDate(b) || '9999';
      return da < db ? -1 : da > db ? 1 : awbKey(a.awb).localeCompare(awbKey(b.awb));
    });
  }

  function addTo(map, key, amount) {
    if (!Number.isFinite(amount)) return;
    map[key || '—'] = round2((map[key || '—'] || 0) + amount);
  }

  /**
   * Quarter-level summary used by the Excel summary sheet and the PDF form.
   */
  function summarize(shipments, { year, q } = {}) {
    const list = sortShipments(shipments.filter((s) => s.included !== false));
    const dates = list.map(shipmentDate).filter(Boolean).sort();
    const byHs = new Map();
    const rates = new Set();
    const countries = new Set();
    const ports = new Set();
    const duty = {};
    const customsValue = {};
    const eligible = {};
    const invoiceTotals = {};
    const byMonth = new Map();
    let classificationCount = 0;
    let netKg = 0;

    for (const s of list) {
      if (s.destCountry) countries.add(s.destCountry);
      if (s.port) ports.add(s.port);
      addTo(eligible, s.eligibleCurrency, s.eligibleValue);
      addTo(invoiceTotals, s.invoiceCurrency, s.invoiceTotal);
      const month = (shipmentDate(s) || '').slice(0, 7) || '—';
      if (!byMonth.has(month)) byMonth.set(month, { month, shipments: 0, duty: 0, value: 0 });
      const mrow = byMonth.get(month);
      mrow.shipments += 1;
      for (const c of s.classifications) {
        classificationCount += 1;
        addTo(duty, s.claimCurrency, c.duty);
        addTo(customsValue, s.claimCurrency, c.value);
        if (Number.isFinite(c.ratePct)) rates.add(+c.ratePct.toFixed(2));
        if (Number.isFinite(c.netKg)) netKg += c.netKg;
        mrow.duty = round2(mrow.duty + (c.duty || 0));
        mrow.value = round2(mrow.value + (c.value || 0));
        const key = c.hs || '—';
        if (!byHs.has(key)) byHs.set(key, { hs: key, descs: new Set(), netKg: 0, value: 0, duty: 0, rates: new Set(), shipments: new Set() });
        const h = byHs.get(key);
        if (c.desc) h.descs.add(c.desc);
        h.netKg = round2(h.netKg + (c.netKg || 0));
        h.value = round2(h.value + (c.value || 0));
        h.duty = round2(h.duty + (c.duty || 0));
        if (Number.isFinite(c.ratePct)) h.rates.add(+c.ratePct.toFixed(2));
        h.shipments.add(s.id);
      }
    }
    const hsRows = [...byHs.values()].map((h) => ({
      hs: h.hs, desc: [...h.descs].join(', '), netKg: h.netKg, value: h.value, duty: h.duty,
      rates: [...h.rates].sort((a, b) => a - b), shipments: h.shipments.size,
    })).sort((a, b) => b.duty - a.duty);

    return {
      year, q,
      range: year && q ? quarterRange(year, q) : null,
      shipments: list,
      count: list.length,
      firstDate: dates[0] || null,
      lastDate: dates[dates.length - 1] || null,
      classificationCount,
      netKg: round2(netKg),
      hsRows,
      hsCodes: hsRows.map((h) => h.hs).filter((h) => h !== '—'),
      rates: [...rates].sort((a, b) => a - b),
      countries: [...countries],
      ports: [...ports],
      duty, customsValue, eligible, invoiceTotals,
      months: [...byMonth.values()].sort((a, b) => a.month.localeCompare(b.month)),
    };
  }

  /** {EUR: 10, USD: 5} → "10.00 EUR + 5.00 USD" */
  function fmtByCurrency(map) {
    const entries = Object.entries(map || {}).filter(([, v]) => Number.isFinite(v));
    if (!entries.length) return '';
    return entries.map(([cur, v]) => `${fmtMoney(v)} ${cur === '—' ? '' : cur}`.trim()).join(' + ');
  }

  /* ------------------------------------------------------- page-fact grouping */

  /**
   * Groups extracted pages into shipments. Pages are linked when they share an
   * AWB number, an invoice base number (e.g. WT24001234) or an export
   * declaration number. Pages without keys stay with their file's main group.
   * pages: [{fileIndex, fileName, page, type, keys:{awb_numbers, invoice_numbers, export_declaration_number}, ...}]
   * Returns [{pages:[...], files:[names], awb}]
   */
  function groupPages(pages) {
    const parent = pages.map((_, i) => i);
    const find = (i) => (parent[i] === i ? i : (parent[i] = find(parent[i])));
    const union = (a, b) => { a = find(a); b = find(b); if (a !== b) parent[b] = a; };
    const keyOwner = new Map();
    const multiAwb = new Set(); // pages that list several AWBs (payment allocations) are shared evidence

    const keysOf = (p) => {
      const k = p.keys || {};
      const out = [];
      const awbs = (k.awb_numbers || []).map(awbKey).filter((x) => x.length >= 8);
      if (awbs.length > 1 && new Set(awbs).size > 1) multiAwb.add(p);
      else awbs.forEach((a) => out.push(`awb:${a}`));
      (k.invoice_numbers || []).forEach((inv) => {
        const base = String(inv).trim().toUpperCase().replace(/\s/g, '').replace(/\/\d+$/, '').replace(/^W(\d)/, 'WT$1');
        if (base.length >= 5) out.push(`inv:${base}`);
      });
      const decl = String(k.export_declaration_number || '').replace(/\D/g, '');
      if (decl.length >= 6) out.push(`decl:${decl}`);
      return out;
    };

    const pageKeys = pages.map(keysOf);
    pages.forEach((p, i) => {
      if (multiAwb.has(p)) return;
      for (const key of pageKeys[i]) {
        if (keyOwner.has(key)) union(keyOwner.get(key), i);
        else keyOwner.set(key, i);
      }
    });

    // keyless pages: attach to the dominant keyed group of the same file
    const byFile = new Map();
    pages.forEach((p, i) => {
      if (!byFile.has(p.fileIndex)) byFile.set(p.fileIndex, []);
      byFile.get(p.fileIndex).push(i);
    });
    for (const idxs of byFile.values()) {
      const keyed = idxs.filter((i) => pageKeys[i].length && !multiAwb.has(pages[i]));
      const roots = new Map();
      keyed.forEach((i) => roots.set(find(i), (roots.get(find(i)) || 0) + 1));
      if (!roots.size) { idxs.forEach((i) => !multiAwb.has(pages[i]) && union(idxs[0], i)); continue; }
      const main = [...roots.entries()].sort((a, b) => b[1] - a[1])[0][0];
      idxs.filter((i) => !pageKeys[i].length && !multiAwb.has(pages[i])).forEach((i) => union(main, i));
    }

    const groups = new Map();
    pages.forEach((p, i) => {
      if (multiAwb.has(p)) return;
      const r = find(i);
      if (!groups.has(r)) groups.set(r, []);
      groups.get(r).push(p);
    });

    const result = [...groups.values()].map((gp) => {
      const awbs = new Set();
      gp.forEach((p) => ((p.keys || {}).awb_numbers || []).forEach((a) => awbKey(a).length >= 8 && awbs.add(awbKey(a))));
      return { pages: gp, awbs: [...awbs], files: [...new Set(gp.map((p) => p.fileName))] };
    });

    // shared evidence pages go to every group whose AWB they mention
    for (const p of multiAwb) {
      const mentioned = new Set((p.keys.awb_numbers || []).map(awbKey));
      for (const g of result) if (g.awbs.some((a) => mentioned.has(a))) g.pages.push(p);
    }
    // groups that hold nothing but application forms / unrelated pages are dropped by the caller
    return result;
  }

  /* -------------------------------------------------------------- example data */

  /** Fictional data so the tool opens in a working state. Not real shipments. */
  function exampleData() {
    const mk = (o) => normalizeShipment(Object.assign({ source: 'example' }, o));
    return {
      settings: {
        company: { name: 'חברת דוגמה ליצוא בע"מ', id: '510000000', street: 'הדוגמה', houseNo: '1', city: 'עיר לדוגמה', zone: '', zip: '1234567', phone: '03-0000000' },
        applicant: { firstName: 'ישראל', lastName: 'ישראלי', idNo: '000000000', role: 'מנכ"ל', phone: '050-0000000' },
        goodsDescription: 'תבלינים',
      },
      year: 2024, q: 4,
      shipments: [
        mk({ exportDecl: '300000101', invoiceNo: 'WT24000101/1/2', invoiceDate: '2024-10-08', invoiceTotal: 21450.6, invoiceCurrency: 'USD', eligibleParts: [310.2, 2650.4], eligibleCurrency: 'USD', incoterm: 'C&F', grossKg: 4620, netKg: 3105, awb: '700 50000101', loadingDate: '2024-10-08', im4: '24BEI0000060000101', releaseDate: '2024-10-09', destCountry: 'בלגיה', port: 'LIEGE', claimCurrency: 'EUR', dutyDocTotal: 262.4,
          classifications: [
            { hs: '0910993300', site: 'מושב לדוגמה א', zip: '1000001', desc: 'THYM', netKg: 48, value: 240.5, ratePct: 7, duty: 16.84 },
            { hs: '0703900000', site: 'מושב לדוגמה ב', zip: '1000002', desc: 'CIBOULETTE', netKg: 300, value: 2361.1, ratePct: 10.4, duty: 245.55 },
          ] }),
        mk({ exportDecl: '300000155', invoiceNo: 'WT24000155', invoiceDate: '2024-11-03', invoiceTotal: 18320, invoiceCurrency: 'USD', eligibleParts: [2212.75], eligibleCurrency: 'USD', incoterm: 'C&F', grossKg: 3980, netKg: 2510, awb: '700 50000155', loadingDate: '2024-11-03', im4: '24BEI0000060000155', releaseDate: '2024-11-04', destCountry: 'בלגיה', port: 'LIEGE', claimCurrency: 'EUR', dutyDocTotal: 203.88,
          classifications: [
            { hs: '0703900000', site: 'מושב לדוגמה ב', zip: '1000002', desc: 'CIBOULETTE', netKg: 250, value: 1960.4, ratePct: 10.4, duty: 203.88 },
          ] }),
        mk({ exportDecl: '300000198', invoiceNo: 'WT24000198/1/2', invoiceDate: '2024-12-15', invoiceTotal: 25010.9, invoiceCurrency: 'USD', eligibleParts: [2801.3, 760.1, 1199.9], eligibleCurrency: 'USD', incoterm: 'C&F', grossKg: 5300, netKg: 3330, awb: '700 50000198', loadingDate: '2024-12-15', im4: '24BEI0000060000198', releaseDate: '2024-12-16', destCountry: 'בלגיה', port: 'LIEGE', claimCurrency: 'EUR', dutyDocTotal: 316.91,
          classifications: [
            { hs: '0703900000', site: 'מושב לדוגמה ב', zip: '1000002', desc: 'CIBOULETTE', netKg: 300, value: 2290.6, ratePct: 10.4, duty: 238.22 },
            { hs: '0709999090', site: 'מושב לדוגמה א', zip: '1000001', desc: 'LEGUMES FRAIS', netKg: 80, value: 605.3, ratePct: 12.8, duty: 77.48 },
            { hs: '0805501090', site: 'מושב לדוגמה א', zip: '1000001', desc: 'CITRONS AUTRES', netKg: 30, value: 18.9, ratePct: 6.4, duty: 1.21 },
          ] }),
      ],
    };
  }

  return {
    SLOTS_PER_ROW, ROMAN,
    num, round2, sum, fmtMoney, fmtPct,
    parseDate, fmtDateShort, fmtDateLong, isoToUTCDate,
    quarterOf, quarterRange, quarterLabel, shipmentDate, suggestQuarter,
    normHs, normAwb, awbKey, im4Short, normCurrency, joinInvoiceNumbers,
    emptyShipment, emptyClassification, normalizeShipment, normClassification, fromExtraction,
    claimTotal, customsValueTotal, applyInvoiceBasis, sortedInvoices,
    validateShipment, issueCounts,
    sortShipments, summarize, fmtByCurrency,
    groupPages, exampleData,
  };
});
