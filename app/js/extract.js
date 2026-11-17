/*
 * ShipoExtract — reads scanned shipment documents with Claude.
 *
 * Stage A: every page image is classified and its fields are read ("page facts").
 * Grouping: pages are tied into shipments by AWB / invoice / export-declaration numbers
 *           (ShipoCore.groupPages), so one combined PDF per shipment and loose
 *           per-document files both work.
 * Stage B: one text-only call per shipment turns its page facts into a claim row.
 *
 * Engines:
 *   - "sample": inside claude.ai, the artifact asks Claude on the viewer's own account.
 *   - "api":    standalone page with the user's Anthropic API key (official SDK).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./core.js'));
  else root.ShipoExtract = factory(root.ShipoCore);
})(typeof self !== 'undefined' ? self : globalThis, function (C) {
  'use strict';

  const SDK_URL = 'https://cdn.jsdelivr.net/npm/@anthropic-ai/sdk@1.3.0/+esm';
  const DEFAULT_MODEL = 'claude-opus-5-5';
  const TARGET_LONG_SIDE = 1600;
  const MAX_PAGES_PER_CALL = 8;

  /* ----------------------------------------------------------------- prompts */

  const PAGE_PROMPT = `You are reading scanned export and customs documents of an Israeli exporter.
The attached images are pages {PAGES} of the file "{FILE}", in that order (one image per page).
Return one object per page.

Page types:
- "application_form": the Hebrew form "טופס בקשה לשיפוי מהקרן לשיפוי יצואנים" or its second page (attached documents, declarations, signatures).
- "il_export_declaration": Israeli export declaration "הצהרת יצוא" (מדינת ישראל – אגף המכס והמע"מ / רשות המסים).
- "air_waybill": air waybill or bill of lading.
- "duty_invoice": the exporter's commercial / duty invoice ("Duty Invoice No. …", proforma, חשבון עסקה).
- "packing_list": packing list of an invoice, listing grower / vendor, city and ZIP code per product.
- "eu_import_declaration": EU import declaration (Single Administrative Document "COMMUNAUTÉ EUROPÉENNE", box 1 "IM", often called IM4); usually one article per page.
- "duty_payment": proof that the EU import duty was paid: a forwarder or airline invoice for "Import duties", a per-AWB allocation list, a receipt.
- "eur1_or_origin": EUR.1 certificate or invoice origin declaration.
- "phytosanitary", "other".

"keys" — identifiers printed on the page:
- awb_numbers: every air waybill / bill of lading number on the page (e.g. "700-12345678"). An allocation list can name several.
- invoice_numbers: the exporter's own invoice numbers (e.g. "WT24001234/1"). On an IM4, box 44 references such as "N935 WT24001234/1+2" mean WT24001234/1 and WT24001234/2. Do not put here the short sequence numbers of the export declaration's "מספר חשבון" field (e.g. 432, 7283).
- export_declaration_number: on an Israeli export declaration only, the 9-digit "מספר הצהרה לדיווח מע\\"מ" (not the 14-digit "מספר הצהרה").

"fields" by page type. Leave out anything not on the page. Numbers are JSON numbers with "." as the decimal point (the European "257,13" is 257.13). Dates are YYYY-MM-DD.
- il_export_declaration: {vat_report_declaration_number, declaration_number, release_date, loading_date, destination_country (as printed), awb_number, total_fob, total_fob_currency (the foreign currency, not ₪), gross_weight_kg, invoices:[{number, value, currency, hs_code, quantity, unit}]}
- air_waybill: {awb_number, gross_weight_kg, pieces, executed_date, flight_date, departure, destination}
- duty_invoice: {invoice_number, invoice_date (the flight date printed on it), terms_of_trade, total_amount, total_currency (currency of the Total column), total_weight_kg, lines:[{description, net_kg, total, total_currency}]}
- packing_list: {invoice_number, lines:[{description, grower_no, vendor_name, city, zip, units}]}
- eu_import_declaration: {mrn (box A reference, e.g. "24BEI0000012345678", no spaces), declaration_date (box A date), customs_office (box A, e.g. "BUR BIERSET (GRACE-HOLLOGNE)"), place (city in box 20 / C, e.g. "LIEGE"), country (country of that customs office, in English), article_no (box 32), goods_description (box 31, e.g. "THYM"), hs_code (box 33, all digits including the leading zero, e.g. "0910993300"), gross_kg (box 35), net_kg (box 38), statistical_value (box 46), currency (box 22), duty_lines:[{type ("A00" customs duty, "B00" VAT, …), base, rate_percent, amount, note (e.g. "Dispense")}], payment_total (the "Report de paiement" amount, if shown), references (box 44 lines)}
- duty_payment: {document (issuer and number), date, currency, total, per_awb:[{awb, amount, currency}]}
- eur1_or_origin: {number, date}
- phytosanitary / other: {summary}
- application_form: {}

Read numbers digit by digit; never guess. If a value is unreadable, leave it out and name it in "unclear".

Reply with JSON only, in this shape:
{"pages":[{"page":<page number>,"type":"...","keys":{"awb_numbers":[],"invoice_numbers":[],"export_declaration_number":null},"fields":{},"unclear":[]}]}`;

  const SHIPMENT_PROMPT = `You are preparing one row of an Israeli exporter's quarterly claim to the fund "הקרן לשיפוי יצואנים בגין פגיעה בהסכמי סחר". The fund refunds EU import duty that was charged because goods were grown in areas that do not get the EU–Israel preferential origin. Below is the data read from the scanned documents of ONE shipment (file and page per item).

Build the row by these rules:
- awb_number: the shipment's air waybill number as "700 12345678" (3 digits, space, 8 digits).
- export_declaration_number: the 9-digit "מספר הצהרה לדיווח מע\\"מ" from the Israeli export declaration.
- invoice_numbers: the exporter duty invoices of this shipment that are in the documents (e.g. ["WT24001234/1","WT24001234/2"]); not invoices that are only referenced.
- invoice_date: the duty invoice date (its flight date). invoice_total / invoice_currency: the sum of the duty invoices' totals, in the currency of their Total column.
- incoterm: terms of trade as printed (e.g. "C&F").
- gross_weight_kg: from the air waybill. net_weight_kg: total weight of the duty invoices.
- loading_date: the air waybill execution or flight date; if missing, the export declaration loading date.
- im4_number: the MRN of the EU import declaration (box A). release_date: its date. If there are several MRNs, use the one that carries the duty and mention the others in issues.
- destination_country: the country of the EU customs office that charged the duty, in Hebrew (בלגיה, הולנד, גרמניה, צרפת, …). port: that office's city or airport in Latin capitals (e.g. LIEGE).
- classifications: one entry per IM4 article on which customs duty was charged (an A00 line with an amount above 0). Ignore VAT (B00) lines.
  - hs_code: box 33, 10 digits with the leading zero. description: box 31 as printed (e.g. "CIBOULETTE").
  - net_weight_kg: box 38. customs_value: the A00 base. duty_rate_percent: the A00 rate as a percent number (7 for 7%). duty_amount: the A00 amount.
  - growing_site / growing_zip: the grower's city or settlement and ZIP code from the packing list, for the products this article covers. Match names across languages (CIBOULETTE = CHIVES, THYM = THYME, CITRONS = citrus such as BUDDHA FINGERS, LEGUMES FRAIS = other fresh herbs or vegetables such as TARRAGON). Join several with ", ". Leave empty if unknown.
- eligible_export_value_parts: the duty-invoice line totals (in the invoice total currency) of the products covered by the classifications. eligible_export_value: their sum.
- duty_currency: the currency of the IM4 duty (usually EUR). duty_total_on_documents: the duty total for THIS shipment shown on the documents ("Report de paiement" on the IM4, or this AWB's line in a payment allocation), else null.
- duty_payment_evidence: one short line naming the proof of payment (e.g. "forwarder invoice 7654321 – 339.46 EUR for this AWB"), or "".
- eur1_number, container_number: only if printed, else "".
- issues: up to 5 short notes in Hebrew about what a reviewer must check before filing: missing or unreadable values, or numbers in this row that disagree between documents (for example the duty total). Leave out differences that do not affect the row. An empty list if none.

Use only values found in the documents and never invent numbers. Use null for anything you cannot find.

Reply with JSON only, with exactly these keys:
{"awb_number":"","export_declaration_number":"","invoice_numbers":[],"invoice_date":"","invoice_total":0,"invoice_currency":"","incoterm":"","gross_weight_kg":0,"net_weight_kg":0,"eur1_number":"","container_number":"","loading_date":"","im4_number":"","release_date":"","destination_country":"","port":"","classifications":[{"hs_code":"","description":"","growing_site":"","growing_zip":"","net_weight_kg":0,"customs_value":0,"duty_rate_percent":0,"duty_amount":0}],"eligible_export_value_parts":[],"eligible_export_value":0,"eligible_export_value_currency":"","duty_currency":"","duty_total_on_documents":null,"duty_payment_evidence":"","issues":[]}

Documents of this shipment:
`;

  function pagePrompt(fileName, pageNumbers) {
    const label = pageNumbers.length === 1 ? `${pageNumbers[0]}` : `${pageNumbers.join(', ')}`;
    return PAGE_PROMPT.replace('{PAGES}', label).replace('{FILE}', fileName);
  }

  /** Compact page facts for stage B (application forms are our output, not a source). */
  function shipmentPrompt(pages) {
    const docs = pages
      .filter((p) => p.type !== 'application_form')
      .map((p) => ({ file: p.fileName, page: p.page, type: p.type, keys: p.keys, fields: p.fields, unclear: p.unclear && p.unclear.length ? p.unclear : undefined }));
    return SHIPMENT_PROMPT + JSON.stringify(docs);
  }

  /* ------------------------------------------------------------- JSON helper */

  function parseJsonLoose(text) {
    if (text && typeof text === 'object') return text;
    let t = String(text || '').trim();
    const fence = t.match(/```(?:json)?\s*([\s\S]*?)```/);
    if (fence) t = fence[1].trim();
    const start = t.search(/[[{]/);
    if (start > 0) t = t.slice(start);
    const endObj = t.lastIndexOf('}');
    const endArr = t.lastIndexOf(']');
    const end = Math.max(endObj, endArr);
    if (end >= 0) t = t.slice(0, end + 1);
    return JSON.parse(t);
  }

  /* ---------------------------------------------------------------- engines */

  class EngineError extends Error {
    constructor(code, message) { super(message || code); this.code = code; }
  }

  /** Claude inside claude.ai (artifact "sample" capability). Resolves null when unavailable. */
  async function createSampleEngine({ tier = 'default' } = {}) {
    const claude = typeof window !== 'undefined' ? window.claude : null;
    if (!claude || typeof claude.use !== 'function') return null;
    let sample;
    try { sample = await claude.use('sample'); } catch (e) { return null; }
    if (!sample) return null;
    let limits = null;
    try { limits = await sample.limits(); } catch (e) { limits = null; }
    if (!limits || !limits.images) return { kind: 'sample', usable: false, reason: 'images' };
    const maxImages = Math.max(1, Math.min(limits.images.maxCount || 1, MAX_PAGES_PER_CALL));
    return {
      kind: 'sample',
      usable: true,
      maxImages,
      maxImageBytes: limits.images.maxInputBytes || 5 * 1024 * 1024,
      maxPromptBytes: limits.maxPromptBytes || 200000,
      async ask(prompt, images, { signal } = {}) {
        try {
          const opts = { modelTier: tier, signal, cache: false };
          if (images && images.length) opts.images = images;
          return await sample.json(prompt, opts);
        } catch (e) {
          if (e && e.code === 'invalid_json' && e.text) {
            try { return parseJsonLoose(e.text); } catch (_) { /* fall through */ }
          }
          throw new EngineError((e && e.code) || 'upstream_error', (e && e.message) || String(e));
        }
      },
    };
  }

  /** Standalone use with the user's own API key, through the official Anthropic SDK. */
  async function createApiEngine({ apiKey, model = DEFAULT_MODEL } = {}) {
    if (!apiKey) return null;
    const mod = await import(/* webpackIgnore: true */ SDK_URL);
    const Anthropic = mod.default || mod.Anthropic;
    const client = new Anthropic({ apiKey, dangerouslyAllowBrowser: true, maxRetries: 3 });
    const toB64 = (blob) => new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => resolve(String(r.result).split(',')[1]);
      r.onerror = () => reject(r.error);
      r.readAsDataURL(blob);
    });
    return {
      kind: 'api',
      usable: true,
      maxImages: MAX_PAGES_PER_CALL,
      maxImageBytes: 5 * 1024 * 1024,
      maxPromptBytes: 2000000,
      async ask(prompt, images, { signal } = {}) {
        const content = [];
        for (const img of images || []) {
          content.push({ type: 'image', source: { type: 'base64', media_type: img.type || 'image/jpeg', data: await toB64(img) } });
        }
        content.push({ type: 'text', text: prompt });
        let msg;
        try {
          msg = await client.beta.messages.create({
            model,
            max_tokens: 16000,
            betas: ['server-side-fallback-2026-07-01'],
            fallbacks: 'default',
            output_config: { effort: 'high' },
            messages: [{ role: 'user', content }],
          }, { signal });
        } catch (e) {
          const status = e && e.status;
          const code = status === 401 ? 'bad_key' : status === 429 ? 'rate_limited' : (e && e.name === 'AbortError') ? 'cancelled' : 'upstream_error';
          throw new EngineError(code, (e && e.message) || String(e));
        }
        if (msg.stop_reason === 'refusal') throw new EngineError('refused', 'The request was declined');
        const text = msg.content.filter((b) => b.type === 'text').map((b) => b.text).join('');
        try { return parseJsonLoose(text); } catch (e) { throw new EngineError('invalid_json', 'Reply was not valid JSON'); }
      },
    };
  }

  /* --------------------------------------------------------- page rendering */

  async function canvasToBlob(canvas, quality) {
    return new Promise((resolve) => canvas.toBlob(resolve, 'image/jpeg', quality));
  }

  async function encodeCanvas(canvas, maxBytes) {
    let q = 0.82;
    let blob = await canvasToBlob(canvas, q);
    while (blob && blob.size > maxBytes && q > 0.4) {
      q -= 0.12;
      blob = await canvasToBlob(canvas, q);
    }
    return blob;
  }

  function isPdf(file) {
    return /pdf$/i.test(file.type) || /\.pdf$/i.test(file.name);
  }
  function isImage(file) {
    return /^image\//.test(file.type) || /\.(jpe?g|png|webp|gif|bmp|tiff?)$/i.test(file.name);
  }

  /**
   * Opens a file and returns { pageCount, render(pageNo) -> Promise<Blob> }.
   */
  async function openDocument(file, { maxImageBytes = 5 * 1024 * 1024 } = {}) {
    if (isPdf(file)) {
      const pdfjs = (typeof window !== 'undefined' && window.pdfjsLib) || null;
      if (!pdfjs) throw new EngineError('no_pdfjs', 'PDF reader did not load');
      const data = new Uint8Array(await file.arrayBuffer());
      const doc = await pdfjs.getDocument({ data, isEvalSupported: false, disableFontFace: true }).promise;
      return {
        pageCount: doc.numPages,
        async render(n) {
          const page = await doc.getPage(n);
          const vp1 = page.getViewport({ scale: 1 });
          const scale = TARGET_LONG_SIDE / Math.max(vp1.width, vp1.height);
          const vp = page.getViewport({ scale });
          const canvas = document.createElement('canvas');
          canvas.width = Math.round(vp.width);
          canvas.height = Math.round(vp.height);
          const ctx = canvas.getContext('2d');
          ctx.fillStyle = '#fff';
          ctx.fillRect(0, 0, canvas.width, canvas.height);
          await page.render({ canvasContext: ctx, viewport: vp }).promise;
          page.cleanup();
          const blob = await encodeCanvas(canvas, maxImageBytes);
          canvas.width = canvas.height = 0;
          return blob;
        },
        close() { try { doc.destroy(); } catch (e) { /* ignore */ } },
      };
    }
    if (isImage(file)) {
      return {
        pageCount: 1,
        async render() {
          const bmp = await createImageBitmap(file);
          const scale = Math.min(1, TARGET_LONG_SIDE / Math.max(bmp.width, bmp.height));
          const canvas = document.createElement('canvas');
          canvas.width = Math.round(bmp.width * scale);
          canvas.height = Math.round(bmp.height * scale);
          const ctx = canvas.getContext('2d');
          ctx.fillStyle = '#fff';
          ctx.fillRect(0, 0, canvas.width, canvas.height);
          ctx.drawImage(bmp, 0, 0, canvas.width, canvas.height);
          return encodeCanvas(canvas, maxImageBytes);
        },
        close() {},
      };
    }
    throw new EngineError('unsupported', `סוג קובץ לא נתמך: ${file.name}`);
  }

  /* ------------------------------------------------------------ job runner */

  /** Runs async tasks with a concurrency limit; stops early when signal aborts. */
  async function runQueue(tasks, concurrency, signal) {
    let next = 0;
    const results = new Array(tasks.length);
    async function worker() {
      while (next < tasks.length) {
        if (signal && signal.aborted) return;
        const i = next++;
        try { results[i] = { ok: true, value: await tasks[i]() }; } catch (e) { results[i] = { ok: false, error: e }; }
      }
    }
    await Promise.all(Array.from({ length: Math.min(concurrency, tasks.length) }, worker));
    return results;
  }

  const sleep = (ms, signal) => new Promise((resolve) => {
    const t = setTimeout(resolve, ms);
    if (signal) signal.addEventListener('abort', () => { clearTimeout(t); resolve(); }, { once: true });
  });

  /** One call with a single delayed retry on rate limiting or a transient upstream error. */
  async function askWithRetry(engine, prompt, images, signal, onWait) {
    try {
      return await engine.ask(prompt, images, { signal });
    } catch (e) {
      if (!['rate_limited', 'upstream_error', 'empty_completion', 'invalid_json', 'queue_overflow'].includes(e.code)) throw e;
      if (onWait) onWait(e.code);
      await sleep(e.code === 'rate_limited' ? 45000 : 5000, signal);
      if (signal && signal.aborted) throw new EngineError('cancelled', 'cancelled');
      return engine.ask(prompt, images, { signal });
    }
  }

  /**
   * Stage A for one file. Returns page facts [{fileIndex, fileName, page, type, keys, fields, unclear}].
   * progress(done, total, note)
   */
  async function readFilePages(engine, file, fileIndex, { signal, progress } = {}) {
    const doc = await openDocument(file, { maxImageBytes: engine.maxImageBytes });
    const total = doc.pageCount;
    const batches = [];
    for (let p = 1; p <= total; p += engine.maxImages) {
      batches.push(Array.from({ length: Math.min(engine.maxImages, total - p + 1) }, (_, i) => p + i));
    }
    let done = 0;
    const facts = [];
    progress && progress(0, total, 'reading');
    try {
      for (const pagesNo of batches) {
        if (signal && signal.aborted) throw new EngineError('cancelled', 'cancelled');
        const images = [];
        for (const n of pagesNo) images.push(await doc.render(n));
        const reply = await askWithRetry(engine, pagePrompt(file.name, pagesNo), images, signal, (code) => progress && progress(done, total, code));
        const list = Array.isArray(reply) ? reply : (reply && reply.pages) || [];
        pagesNo.forEach((n, i) => {
          const got = list.find((x) => Number(x.page) === n) || list[i] || {};
          facts.push({
            fileIndex, fileName: file.name, page: n,
            type: got.type || 'other',
            keys: normalizeKeys(got.keys),
            fields: got.fields || {},
            unclear: Array.isArray(got.unclear) ? got.unclear : [],
          });
        });
        done += pagesNo.length;
        progress && progress(done, total, 'reading');
      }
    } finally {
      doc.close();
    }
    return facts;
  }

  function normalizeKeys(k) {
    k = k || {};
    const arr = (v) => (Array.isArray(v) ? v : v ? [v] : []).map(String).filter(Boolean);
    return {
      awb_numbers: arr(k.awb_numbers),
      invoice_numbers: arr(k.invoice_numbers),
      export_declaration_number: k.export_declaration_number ? String(k.export_declaration_number) : null,
    };
  }

  /** Stage B for one group of pages → shipment (ShipoCore model). */
  async function buildShipment(engine, group, { signal } = {}) {
    const prompt = shipmentPrompt(group.pages);
    const reply = await askWithRetry(engine, prompt, null, signal);
    return C.fromExtraction(reply, group.files);
  }

  /** Does a group contain anything worth turning into a claim row? */
  function groupIsRelevant(group) {
    return group.pages.some((p) => ['eu_import_declaration', 'il_export_declaration', 'air_waybill', 'duty_invoice'].includes(p.type));
  }

  return {
    PAGE_PROMPT, SHIPMENT_PROMPT, DEFAULT_MODEL,
    pagePrompt, shipmentPrompt, parseJsonLoose,
    createSampleEngine, createApiEngine, EngineError,
    openDocument, isPdf, isImage,
    runQueue, readFilePages, buildShipment, groupIsRelevant, normalizeKeys,
  };
});
