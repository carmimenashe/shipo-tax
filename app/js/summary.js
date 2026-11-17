/*
 * ShipoSummary — the summary for the fund ("דף סיכום לרשות המסים"):
 *   1-2. the official request form filled at quarter level (consolidated)
 *   3+.  appendix A: every shipment of the quarter; appendix B: totals per customs item
 * and, optionally, the same form per shipment (as was attached to each shipment file).
 *
 * Pages are built as HTML strings (pure, testable) and turned into a PDF in the
 * browser: HTML -> SVG <foreignObject> -> canvas -> jsPDF. The browser lays out
 * the Hebrew text itself, so right-to-left text and mixed numbers render correctly.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./core.js'));
  else root.ShipoSummary = factory(root.ShipoCore);
})(typeof self !== 'undefined' ? self : globalThis, function (C) {
  'use strict';

  const PAGE_W = 794; // A4 at 96 dpi
  const PAGE_H = 1123;

  const esc = (v) => String(v === null || v === undefined ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  const DEFAULT_ATTACHMENTS = [
    { key: 'form', label: 'טופס בקשה של היצואן', checked: true },
    { key: 'exportDecl', label: 'רשימון ייצוא ושטר מטען', checked: true },
    { key: 'invoice', label: 'חשבון עסקה', checked: true },
    { key: 'origin', label: 'הצהרת חשבונית (אם יצואן מאושר) או תעודת מקור (EUR.1) בציון מקום היצור', checked: false },
    { key: 'importDecl', label: 'רשימון יבוא', checked: true },
    { key: 'payment', label: 'חשבונית/קבלת תשלום או מסמך רשמי אחר (לרבות רשימון יבוא) המעידים על התשלום בפועל למכס האירופאי (וכן לגבי מדינות אפט"א) תוך ציון גובה התשלום', checked: true },
    { key: 'deficit', label: 'במקרה של דרישת "גרעון" מהמכס האירופאי – קבלה על תשלום המכס בפועל או אסמכתא על ערבות שהופקדה', checked: false },
    { key: 'cpa', label: 'אישור עדכני מטעם רואה החשבון של החברה בדבר זהותם של מורשי החתימה בחברה', checked: false },
    { key: 'excel', label: 'גליון אקסל מפורט', checked: true },
  ];

  const DECLARATIONS = [
    'המכס שהוטל על ידי הרשויות בחו"ל על המשלוח המצוין לעיל הוטל רק בגין היותו מיוצר באחד, או יותר, מהשטחים המפורטים להלן: רצועת עזה, הגדה המערבית, ירושלים המזרחית, רמת הגולן.',
    'המכס הנדרש שולם בפועל.',
    'הסכום המבוקש לשיפוי אינו כולל עלויות הובלה, מע"מ, ריביות או קנסות.',
    'אני מורשה לחתימה בשם החברה. רצ"ב אישור עדכני מטעם רואה החשבון של החברה בדבר זהותם של מורשי החתימה בחברה.',
    'הטובין אשר בגינם אני דורש החזר מכס עומדים בכללי המקור שנקבעו בהסכם אירופי – ים תיכוני, וידוע לי כי במידה ויימצא כי הטובין אינם עומדים בכללי המקור איאלץ להחזיר את הכספים שהתקבלו בידיי בעקבות בקשה זו בצירוף הפרשי הצמדה וריבית כחוק עד לביצוע התשלום המלא בפועל.',
    'ידוע לי כי עלי להצהיר בבקשה זו ובכל המסמכים הנלווים אליה עובדות נכונות ומדויקות וכי אם לא אעשה כן, אהיה צפוי לעונשים הקבועים בחוק ותתבטל למפרע זכאותי לשיפוי מהקרן לשיפוי יצואנים (להלן: "הקרן"). בנסיבות בהן תתבטל זכאותי לשיפוי כאמור, תהיה הקרן זכאית לדרוש ממני את החזרת הכספים שהתקבלו בידיי בעקבות הגשת בקשה זו, בצירוף הפרשי הצמדה וריבית כחוק עד לביצוע התשלום המלא בפועל.',
    'ידוע לי כי ככל שתאושר בקשתי לשיפוי, ביצוע התשלום ייעשה ע"י הקרן לשיפוי יצואנים.',
  ];

  const STYLE = `
    *{box-sizing:border-box;margin:0;padding:0}
    .pg{width:${PAGE_W}px;height:${PAGE_H}px;background:#fff;color:#111;direction:rtl;font-family:PdfSans,Arial,'Noto Sans Hebrew',sans-serif;font-size:13.5px;line-height:1.45;padding:58px 64px 48px;position:relative;overflow:hidden}
    .ttl{font-family:PdfSerif,PdfSans,Arial,serif;font-size:21px;font-weight:700;text-align:center;text-decoration:underline;text-underline-offset:5px;margin-bottom:14px}
    .sub{text-align:center;font-size:13px;color:#333;margin:-6px 0 12px}
    h3{font-size:14.5px;font-weight:700;text-align:center;text-decoration:underline;font-style:italic;margin:16px 0 9px}
    .row{display:flex;gap:18px;align-items:baseline;margin:6px 0;flex-wrap:wrap}
    .f{display:flex;gap:6px;align-items:baseline}
    .lbl{white-space:nowrap}
    .val{font-weight:700;border-bottom:1px solid #222;min-width:90px;padding:0 6px 1px;text-align:center;unicode-bidi:plaintext}
    .val.w{min-width:200px}.val.xw{min-width:360px;text-align:right}
    .ltr{direction:ltr;unicode-bidi:isolate}
    .qrow{display:flex;justify-content:space-between;margin:8px 10px 2px}
    .q{display:flex;align-items:center;gap:8px;font-size:14px}
    .box{width:30px;height:24px;border:1.5px solid #222;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:15px}
    .small{font-size:11px;color:#333}
    .hs{direction:ltr;text-align:right;font-weight:700;line-height:1.55}
    .hsl{display:flex;gap:14px;align-items:flex-start;margin:6px 0}
    ol.decl{padding-right:22px;margin:4px 0}
    ol.decl li{margin:0 0 7px;padding-right:4px;text-align:justify}
    .chk{display:flex;gap:8px;align-items:flex-start;margin:0 0 6px}
    .cb{width:15px;height:15px;border:1.2px solid #222;flex:none;margin-top:3px;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700}
    .sig{display:flex;gap:28px;margin-top:22px;flex-wrap:wrap}
    .sig .f{min-width:200px}
    .foot{position:absolute;bottom:22px;left:64px;right:64px;display:flex;justify-content:space-between;font-size:10.5px;color:#555;border-top:1px solid #bbb;padding-top:5px}
    table{width:100%;border-collapse:collapse;font-size:10.5px;line-height:1.3}
    th,td{border:1px solid #444;padding:3px 4px;text-align:center;vertical-align:middle}
    th{background:#e6e6e6;font-weight:700}
    td.n{direction:ltr;font-variant-numeric:tabular-nums;white-space:nowrap}
    tr.tot td{font-weight:700;background:#f2f2f2}
    .note{font-size:11.5px;color:#333;margin-top:10px}
    .kv{display:grid;grid-template-columns:auto 1fr;gap:4px 14px;margin:6px 0 10px;font-size:13px}
    .kv b{font-weight:700}
  `;

  /* -------------------------------------------------------------- helpers */

  const money = (n) => C.fmtMoney(n);
  const v = (x, cls = '') => `<span class="val ${cls}">${esc(x) || '&#160;'}</span>`;
  const field = (label, value, cls) => `<span class="f"><span class="lbl">${esc(label)}</span>${v(value, cls)}</span>`;

  function quarterBoxes(q) {
    return `<div class="qrow">${[1, 2, 3, 4].map((n) => `<span class="q"><span>רבעון ${C.ROMAN[n - 1]}</span><span class="box">${n === q ? 'V' : ''}</span></span>`).join('')}</div>`;
  }

  function footer(meta, pageNo, pageCount) {
    const company = (meta.settings && meta.settings.company) || {};
    return `<div class="foot"><span>${esc(company.name || '')} · ${esc(C.quarterLabel(meta.year, meta.q))}</span><span>עמוד ${pageNo} מתוך ${pageCount}</span></div>`;
  }

  /**
   * One request form (2 pages). `exp` holds the export-details block:
   * { description, hsCodes[], dateText, destination, notes, fob, rates, duty, refText }
   */
  function formPages(meta, exp) {
    const st = meta.settings || {};
    const co = st.company || {};
    const ap = st.applicant || {};
    const hs = exp.hsCodes.length ? exp.hsCodes : [''];
    const hsLines = [];
    for (let i = 0; i < Math.max(5, hs.length); i++) hsLines.push(`<div>${hs[i] ? `<span class="ltr">${esc(hs[i])}</span>` : '__________'} .${i + 1}</div>`);
    const half = Math.ceil(hsLines.length / 2);
    const hsCols = hsLines.length > 6
      ? `<div class="hs">${hsLines.slice(0, half).join('')}</div><div class="hs">${hsLines.slice(half).join('')}</div>`
      : `<div class="hs">${hsLines.join('')}</div>`;

    const p1 = `
      <div class="ttl">טופס בקשה לשיפוי מהקרן לשיפוי יצואנים בגין פגיעה בהסכמי סחר</div>
      <div class="row" style="justify-content:center">${field('שנה:', meta.year)}</div>
      ${quarterBoxes(meta.q)}
      <h3>פרטי החברה הפונה בבקשה</h3>
      <div class="row">${field('שם החברה:', co.name, 'w')}${field('מספר ח.פ (או ח.צ):', co.id)}</div>
      <div class="row">${field('כתובת: (נא לא לציין תא-דואר) רחוב', co.street, 'w')}${field("מס'", co.houseNo)}</div>
      <div class="row">${field('ישוב/עיר:', co.city)}${field('אזור תעשייה:', co.zone)}</div>
      <div class="row">${field('מיקוד:', co.zip)}${field('טלפון:', co.phone)}</div>
      <h3>מגיש הבקשה בשם החברה הפונה</h3>
      <div class="sub small">(חייב להיות מורשה לחתימה מטעם החברה)</div>
      <div class="row">${field('שם: (פרטי)', ap.firstName)}${field('(משפחה)', ap.lastName)}</div>
      <div class="row">${field('ת.ז:', ap.idNo)}${field('תפקיד בחברה:', ap.role)}</div>
      <div class="row">${field('טלפון:', ap.phone)}</div>
      <h3>פרטי היצוא בגינם מוגשת הבקשה</h3>
      <div class="row">${field('תיאור הטובין (טקסט חופשי):', exp.description, 'w')}${exp.refText ? `<span class="val ltr">${esc(exp.refText)}</span>` : ''}</div>
      <div class="hsl"><span class="lbl">פריט מכס (לפחות 6 ספרות):</span>${hsCols}</div>
      <div class="row">${field('תאריך המשלוח:', exp.dateText, 'w')}${field('יעד סופי:', exp.destination)}</div>
      <div class="row">${field('הערות (טקסט חופשי):', exp.notes, 'xw')}</div>
      <div class="row">${field('ערך המשלוח (FOB):', exp.fob, 'w')}</div>
      <div class="row">${field('אחוז המכס ששולם בפועל:', exp.rates)}${field('סכום המכס ששולם בפועל:', exp.duty, 'w')}</div>`;

    const attachments = (st.attachments && st.attachments.length ? st.attachments : DEFAULT_ATTACHMENTS);
    const signDate = meta.signDate ? C.fmtDateLong(meta.signDate) : '';
    const p2 = `
      <h3 style="text-align:right;margin-top:0">מצורפים המסמכים הבאים:</h3>
      ${attachments.map((a) => `<div class="chk"><span class="cb">${a.checked ? '✓' : ''}</span><span>${esc(a.label)}</span></div>`).join('')}
      <h3 style="text-align:right">הצהרת מגיש הבקשה:</h3>
      <div>אני הח"מ מצהיר כי:</div>
      <ol class="decl">${DECLARATIONS.map((d) => `<li>${esc(d)}</li>`).join('')}</ol>
      <div class="sig">${field('תאריך:', signDate)}${field('שם החברה:', co.name, 'w')}</div>
      <div class="sig">${field('חתימה (בצירוף חותמת):', '', 'w')}</div>
      <div class="sig">${field('שם/שמות החותמים מטעם החברה:', [ap.firstName, ap.lastName].filter(Boolean).join(' '), 'w')}${field('ציון תפקיד החותמים:', ap.role)}</div>`;
    return [p1, p2];
  }

  /* ----------------------------------------------------- quarterly summary */

  function consolidatedExport(meta, sum) {
    const st = meta.settings || {};
    const descs = [...new Set(sum.hsRows.map((h) => h.desc).filter(Boolean))].join(', ');
    const dateText = sum.firstDate
      ? `${C.fmtDateLong(sum.firstDate)} – ${C.fmtDateLong(sum.lastDate)} (${sum.count} משלוחים)`
      : '';
    return {
      description: st.goodsDescription || descs,
      hsCodes: sum.hsCodes,
      dateText,
      destination: sum.countries.join(', '),
      notes: `בקשה מרוכזת ל${C.quarterLabel(meta.year, meta.q)}: ${sum.count} משלוחים. פירוט בנספחים א'–ב' ובגליון האקסל.`,
      fob: C.fmtByCurrency(sum.customsValue),
      rates: sum.rates.map(C.fmtPct).join(', '),
      duty: C.fmtByCurrency(sum.duty),
      refText: '',
    };
  }

  const APPX_ROWS_FIRST = 26;
  const APPX_ROWS_NEXT = 32;

  function appendixA(meta, sum) {
    const rows = sum.shipments.map((s, i) => `<tr>
      <td>${i + 1}</td><td class="n">${esc(s.awb)}</td><td class="n">${esc(s.exportDecl)}</td>
      <td class="n">${esc(s.invoiceNo)}</td><td class="n">${esc(C.fmtDateShort(s.loadingDate))}</td>
      <td class="n">${esc(C.im4Short(s.im4))}</td><td class="n">${esc(C.fmtDateShort(s.releaseDate))}</td>
      <td class="n" style="white-space:normal">${s.classifications.map((c) => esc(c.hs)).join('<br>')}</td>
      <td class="n">${money(C.customsValueTotal(s))}</td><td class="n">${money(C.claimTotal(s))}</td></tr>`);
    const head = `<tr><th>#</th><th>שטר מטען</th><th>הצהרת יצוא</th><th>חשבון עסקה</th><th>תאריך טעינה</th><th>IM4</th><th>תאריך שחרור</th><th>פרטי מכס</th><th>ערך למכס</th><th>סכום המכס</th></tr>`;
    const total = `<tr class="tot"><td colspan="8">סה"כ (${sum.count} משלוחים)</td><td class="n">${esc(C.fmtByCurrency(sum.customsValue))}</td><td class="n">${esc(C.fmtByCurrency(sum.duty))}</td></tr>`;
    const pages = [];
    let i = 0;
    let first = true;
    while (i < rows.length || first) {
      const take = first ? APPX_ROWS_FIRST : APPX_ROWS_NEXT;
      const chunk = rows.slice(i, i + take);
      i += take;
      const last = i >= rows.length;
      pages.push(`
        <div class="ttl">נספח א' – ריכוז המשלוחים ב${esc(C.quarterLabel(meta.year, meta.q))}${first ? '' : ' (המשך)'}</div>
        ${first ? `<div class="kv"><b>תקופה:</b><span>${esc(sum.range ? `${C.fmtDateLong(sum.range.from)} – ${C.fmtDateLong(sum.range.to)}` : '')}</span>
          <b>מספר משלוחים:</b><span>${sum.count}</span>
          <b>ערך יצוא זכאי לשיפוי:</b><span>${esc(C.fmtByCurrency(sum.eligible))}</span>
          <b>מטבע השיפוי:</b><span>${esc(Object.keys(sum.duty).join(', ') || '')}</span></div>` : ''}
        <table>${head}${chunk.join('')}${last ? total : ''}</table>
        ${last ? '<div class="note">ערך למכס וסכום המכס – לפי הצהרות היבוא (IM4) בגין הפריטים הזכאים לשיפוי בלבד. סכום השיפוי המבוקש אינו כולל מע"מ, הובלה, ריביות או קנסות.</div>' : ''}`);
      first = false;
    }
    return pages;
  }

  function appendixB(meta, sum) {
    const rows = sum.hsRows.map((h) => `<tr><td class="n">${esc(h.hs)}</td><td>${esc(h.desc)}</td><td>${h.shipments}</td>
      <td class="n">${money(h.netKg)}</td><td class="n">${money(h.value)}</td><td class="n">${esc(h.rates.map(C.fmtPct).join(', '))}</td><td class="n">${money(h.duty)}</td></tr>`).join('');
    const months = sum.months.map((m) => `<tr><td class="n">${esc(m.month === '—' ? '—' : `${m.month.slice(5)}/${m.month.slice(0, 4)}`)}</td><td>${m.shipments}</td><td class="n">${money(m.value)}</td><td class="n">${money(m.duty)}</td></tr>`).join('');
    return [`
      <div class="ttl">נספח ב' – ריכוז לפי פרט מכס</div>
      <table><tr><th>פרט מכס</th><th>תיאור</th><th>משלוחים</th><th>משקל נטו (ק"ג)</th><th>ערך למכס</th><th>אחוז מכס</th><th>סכום המכס</th></tr>${rows}
      <tr class="tot"><td colspan="3">סה"כ</td><td class="n">${money(sum.netKg)}</td><td class="n">${esc(C.fmtByCurrency(sum.customsValue))}</td><td></td><td class="n">${esc(C.fmtByCurrency(sum.duty))}</td></tr></table>
      <h3 style="margin-top:26px">ריכוז לפי חודש</h3>
      <table><tr><th>חודש</th><th>משלוחים</th><th>ערך למכס</th><th>סכום המכס</th></tr>${months}</table>`];
  }

  /** All pages of the quarterly summary, as HTML strings (inner content of each page). */
  function quarterlyPages(meta, shipments) {
    const sum = C.summarize(shipments, meta);
    const pages = [...formPages(meta, consolidatedExport(meta, sum)), ...appendixA(meta, sum), ...appendixB(meta, sum)];
    return wrapPages(meta, pages);
  }

  /** The form for one shipment, as attached to each shipment's file. */
  function shipmentFormExport(meta, s) {
    const st = meta.settings || {};
    const rates = [...new Set(s.classifications.map((c) => c.ratePct).filter(Number.isFinite))].sort((a, b) => a - b);
    return {
      description: st.goodsDescription || [...new Set(s.classifications.map((c) => c.desc).filter(Boolean))].join(', '),
      hsCodes: s.classifications.map((c) => c.hs),
      dateText: C.fmtDateLong(C.shipmentDate(s)),
      destination: s.destCountry,
      notes: [s.awb ? `שטר מטען ${s.awb}` : '', s.im4 ? `IM4 ${s.im4}` : ''].filter(Boolean).join(' · '),
      fob: s.classifications.length ? `${money(C.customsValueTotal(s))} ${s.claimCurrency}` : '',
      rates: rates.map(C.fmtPct).join(', '),
      duty: s.classifications.length ? `${money(C.claimTotal(s))} ${s.claimCurrency}` : '',
      refText: s.awb,
    };
  }

  function shipmentFormPages(meta, shipments) {
    const list = C.sortShipments(shipments.filter((s) => s.included !== false));
    const pages = [];
    list.forEach((s) => pages.push(...formPages(meta, shipmentFormExport(meta, s))));
    return wrapPages(meta, pages);
  }

  function wrapPages(meta, inner) {
    return inner.map((html, i) => `<div class="pg">${html}${footer(meta, i + 1, inner.length)}</div>`);
  }

  /* ------------------------------------------------------- PDF (browser) */

  function pageSvg(pageHtml, fontCss) {
    const xhtml = `<div xmlns="http://www.w3.org/1999/xhtml"><style>${fontCss || ''}${STYLE}</style>${pageHtml}</div>`;
    // XHTML needs self-closing void elements
    const fixed = xhtml.replace(/<br>/g, '<br/>');
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${PAGE_W}" height="${PAGE_H}"><foreignObject x="0" y="0" width="100%" height="100%">${fixed}</foreignObject></svg>`;
  }

  function loadImage(src) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = () => reject(new Error('page render failed'));
      img.src = src;
    });
  }

  /** Draws one page into a canvas at `scale` and returns a JPEG data URL. */
  async function renderPage(pageHtml, { fontCss, scale = 2 } = {}) {
    const svg = pageSvg(pageHtml, fontCss);
    const url = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
    const img = await loadImage(url);
    // give embedded fonts a moment on the first decode
    if (img.decode) { try { await img.decode(); } catch (e) { /* drawn anyway */ } }
    const canvas = document.createElement('canvas');
    canvas.width = PAGE_W * scale;
    canvas.height = PAGE_H * scale;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
    const data = canvas.toDataURL('image/jpeg', 0.9);
    canvas.width = canvas.height = 0;
    return data;
  }

  /** pages: HTML strings from quarterlyPages/shipmentFormPages. Returns a Blob (application/pdf). */
  async function renderPdf(pages, { jsPDF, fontCss, onProgress } = {}) {
    if (!jsPDF) throw new Error('PDF library did not load');
    // warm up the embedded fonts once so the first page is not drawn with a fallback face
    if (fontCss) await renderPage('<div class="pg"><div class="ttl">א</div><div>a1 א</div></div>', { fontCss, scale: 0.2 });
    const pdf = new jsPDF({ unit: 'pt', format: 'a4', orientation: 'portrait', compress: true });
    const w = pdf.internal.pageSize.getWidth();
    const h = pdf.internal.pageSize.getHeight();
    for (let i = 0; i < pages.length; i++) {
      const img = await renderPage(pages[i], { fontCss });
      if (i > 0) pdf.addPage();
      pdf.addImage(img, 'JPEG', 0, 0, w, h, undefined, 'FAST');
      onProgress && onProgress(i + 1, pages.length);
    }
    return pdf.output('blob');
  }

  /** For on-screen preview: the page HTML plus the CSS it needs. */
  function previewCss(fontCss) {
    return `${fontCss || ''}${STYLE}`;
  }

  return {
    PAGE_W, PAGE_H, DEFAULT_ATTACHMENTS, DECLARATIONS,
    quarterlyPages, shipmentFormPages, consolidatedExport, shipmentFormExport,
    renderPdf, renderPage, pageSvg, previewCss,
  };
});
