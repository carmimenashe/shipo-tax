/* UI controller for the quarterly indemnity claim builder. */
(function () {
  'use strict';
  const C = window.ShipoCore;
  const X = window.ShipoExcel;
  const S = window.ShipoSummary;
  const E = window.ShipoExtract;
  const STORE_KEY = 'shipo.v1';
  const $ = (id) => document.getElementById(id);
  const esc = (v) => String(v === null || v === undefined ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  /* ------------------------------------------------------------------ state */

  function defaultSettings() {
    return {
      company: { name: '', id: '', street: '', houseNo: '', city: '', zone: '', zip: '', phone: '' },
      applicant: { firstName: '', lastName: '', idNo: '', role: '', phone: '' },
      goodsDescription: '',
      attachments: S.DEFAULT_ATTACHMENTS.map((a) => ({ ...a })),
      tier: 'default',
      invoiceBasis: 'first',
      signDate: null,
    };
  }

  function lastCompletedQuarter() {
    const now = new Date();
    let q = Math.floor(now.getMonth() / 3); // previous quarter (0 = Q4 of last year)
    let y = now.getFullYear();
    if (q === 0) { q = 4; y -= 1; }
    return { year: y, q };
  }

  let state = null;
  const files = []; // {key, file, kind, status, done, total, note, msg}
  let running = null;
  let openId = null;
  let downloadsCap;

  function load() {
    let saved = null;
    try { saved = JSON.parse(localStorage.getItem(STORE_KEY) || 'null'); } catch (e) { saved = null; }
    if (saved && saved.version === 1) {
      state = {
        version: 1,
        settings: Object.assign(defaultSettings(), saved.settings || {}),
        year: saved.year, q: saved.q, quarterAuto: !!saved.quarterAuto,
        shipments: (saved.shipments || []).map(C.normalizeShipment),
        pages: saved.pages || [],
        isExample: !!saved.isExample,
      };
      state.settings.company = Object.assign(defaultSettings().company, state.settings.company || {});
      state.settings.applicant = Object.assign(defaultSettings().applicant, state.settings.applicant || {});
      return;
    }
    const ex = C.exampleData();
    state = {
      version: 1,
      settings: Object.assign(defaultSettings(), { company: ex.settings.company, applicant: ex.settings.applicant, goodsDescription: ex.settings.goodsDescription }),
      year: ex.year, q: ex.q, quarterAuto: true,
      shipments: ex.shipments, pages: [], isExample: true,
    };
  }

  let saveTimer = null;
  function save() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => {
      const data = { ...state };
      try {
        localStorage.setItem(STORE_KEY, JSON.stringify(data));
      } catch (e) {
        try { localStorage.setItem(STORE_KEY, JSON.stringify({ ...data, pages: [] })); } catch (_) { /* storage unavailable */ }
      }
    }, 250);
  }

  function clearExample() {
    if (!state.isExample) return;
    const blank = defaultSettings();
    state.settings.company = blank.company;
    state.settings.applicant = blank.applicant;
    state.settings.goodsDescription = '';
    state.shipments = [];
    state.pages = [];
    state.isExample = false;
    state.quarterAuto = true;
    openId = null;
    fillSettings();
  }

  /* ---------------------------------------------------------------- helpers */

  let toastTimer = null;
  function toast(msg, ms = 3800) {
    const t = $('toast');
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { t.hidden = true; }, ms);
  }

  const ERROR_COPY = {
    not_granted: 'לא ניתן אישור לשימוש ב-Claude. אפשר לאשר בפעם הבאה שתלחצו "קרא את המסמכים".',
    rate_limited: 'הגעתם למגבלת השימוש ב-Claude כרגע. המתינו כמה דקות ולחצו שוב "קרא את המסמכים" — מה שכבר נקרא נשמר.',
    images_unavailable: 'Claude לא יכול לקבל תמונות בתצוגה הזו. פתחו את הכלי בדפדפן במחשב.',
    image_rejected: 'אחד העמודים נדחה. נסו לשמור את הקובץ מחדש כ-PDF.',
    prompt_too_large: 'המשלוח גדול מדי לבקשה אחת. פצלו את הקובץ.',
    refused: 'Claude סירב לקרוא את המסמך.',
    bad_key: 'מפתח ה-API שגוי. בדקו אותו בהגדרות.',
    cancelled: 'העיבוד נעצר.',
    no_pdfjs: 'קורא ה-PDF לא נטען. רעננו את הדף.',
    unsupported: 'סוג קובץ שאינו נתמך.',
    sampling_disabled: 'השימוש ב-Claude מתוך הכלי כבוי בחשבון זה.',
    session_expired: 'פג תוקף החיבור. רעננו את הדף ונסו שוב.',
  };
  const errCopy = (e) => ERROR_COPY[e && e.code] || `שגיאה בקריאה (${(e && (e.code || e.message)) || 'לא ידוע'}). נסו שוב.`;

  function meta() {
    const ap = state.settings.applicant;
    const issuesById = {};
    included().forEach((s) => { issuesById[s.id] = issuesOf(s); });
    return {
      year: state.year, q: state.q, settings: state.settings, company: state.settings.company,
      signerName: [ap.firstName, ap.lastName].filter(Boolean).join(' '),
      signDate: state.settings.signDate || new Date().toISOString().slice(0, 10),
      issuesById,
    };
  }

  const included = () => state.shipments.filter((s) => s.included !== false);
  const issuesOf = (s) => C.validateShipment(s, { year: state.year, q: state.q, shipments: state.shipments });

  async function saveFile(name, blob) {
    if (downloadsCap === undefined) {
      downloadsCap = null;
      if (window.claude && typeof window.claude.use === 'function') {
        try { downloadsCap = await window.claude.use('downloads'); } catch (e) { downloadsCap = null; }
      }
    }
    if (downloadsCap) {
      try {
        await downloadsCap.save({ filename: name, data: blob });
        toast(`הקובץ ${name} נשמר`);
      } catch (e) {
        if (e && e.code === 'declined') return;
        toast(e && e.code === 'rate_limited' ? 'חלון שמירה כבר פתוח. נסו שוב בעוד רגע.' : 'לא ניתן לשמור קבצים בתצוגה הזו.');
      }
      return;
    }
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }

  /* --------------------------------------------------------------- quarter */

  function renderQuarter() {
    const nowY = new Date().getFullYear();
    const years = new Set();
    for (let y = nowY - 3; y <= nowY + 1; y++) years.add(y);
    if (state.year) years.add(state.year);
    $('qYear').innerHTML = [...years].sort().map((y) => `<option value="${y}">${y}</option>`).join('');
    $('qYear').value = String(state.year);
    $('qQuarter').value = String(state.q);
    $('stampQ').textContent = `Q${state.q}`;
  }

  function autoQuarter() {
    if (!state.quarterAuto) return;
    const g = C.suggestQuarter(state.shipments);
    if (g) { state.year = g.year; state.q = g.q; renderQuarter(); }
  }

  /* ------------------------------------------------------------------ KPIs */

  function renderKpis() {
    const sum = C.summarize(state.shipments, { year: state.year, q: state.q });
    $('kDuty').textContent = C.fmtByCurrency(sum.duty) || '0.00';
    $('kCount').textContent = String(sum.count);
    $('kValue').textContent = C.fmtByCurrency(sum.customsValue) || '0.00';
    let errs = 0;
    let warns = 0;
    included().forEach((s) => { const c = C.issueCounts(issuesOf(s)); errs += c.error; warns += c.warn; });
    $('kIssues').textContent = String(errs + warns);
    $('kIssuesBox').classList.toggle('alert', errs > 0);
    $('exampleBanner').hidden = !state.isExample;
    $('outNote').textContent = errs
      ? `יש ${errs} שגיאות ו-${warns} אזהרות בטבלה. מומלץ לתקן לפני ההגשה.`
      : warns ? `${warns} אזהרות לבדיקה בטבלה.` : sum.count ? 'הנתונים עברו את כל הבדיקות.' : '';
  }

  /* ----------------------------------------------------------------- files */

  const fileKey = (f) => `${f.name}|${f.size}`;
  function kindOf(f) {
    if (/\.xlsx$/i.test(f.name)) return 'excel';
    if (/\.json$/i.test(f.name)) return 'json';
    if (E.isPdf(f) || E.isImage(f)) return 'doc';
    return 'other';
  }

  function addFiles(list) {
    for (const f of list) {
      const key = fileKey(f);
      if (files.some((x) => x.key === key)) continue;
      const kind = kindOf(f);
      files.push({ key, file: f, kind, status: kind === 'other' ? 'error' : 'pending', done: 0, total: 0, note: '', msg: kind === 'other' ? ERROR_COPY.unsupported : '' });
    }
    renderFiles();
    // spreadsheets and backups load right away; documents wait for the button
    const quick = files.filter((x) => x.status === 'pending' && x.kind !== 'doc');
    if (quick.length) importQuick(quick);
  }

  function renderFiles() {
    $('fileList').innerHTML = files.map((f) => {
      let st = '';
      let cls = '';
      if (f.status === 'pending') st = f.kind === 'doc' ? 'ממתין לקריאה' : 'נטען…';
      if (f.status === 'queued') st = 'בתור';
      if (f.status === 'running') st = f.total ? `קורא עמודים ${f.done}/${f.total}${f.note === 'rate_limited' ? ' · ממתין למגבלת שימוש' : ''}` : 'פותח את הקובץ…';
      if (f.status === 'done') { st = f.msg || (f.total ? `נקרא · ${f.total} עמודים` : 'נטען'); cls = 'done'; }
      if (f.status === 'error') { st = f.msg || 'שגיאה'; cls = 'err'; }
      const pct = f.total ? Math.round((f.done / f.total) * 100) : 0;
      return `<li class="file"><span class="name" title="${esc(f.file.name)}">${esc(f.file.name)}</span><span class="st ${cls}">${esc(st)}</span>${f.status === 'running' ? `<span class="bar"><i style="width:${pct}%"></i></span>` : ''}</li>`;
    }).join('');
    const hasPendingDocs = files.some((f) => f.kind === 'doc' && (f.status === 'pending' || f.status === 'error'));
    $('processBtn').disabled = !!running || !hasPendingDocs;
    $('stopBtn').hidden = !running;
    $('clearFiles').hidden = !!running || !files.length;
  }

  async function importQuick(list) {
    for (const f of list) {
      try {
        if (f.kind === 'excel') {
          const res = await X.importWorkbook(window.ExcelJS, await f.file.arrayBuffer());
          if (state.isExample) clearExample();
          const wasEmpty = !state.shipments.length;
          res.shipments.forEach((s) => upsert(s));
          if (wasEmpty && res.year && res.q) { state.year = res.year; state.q = res.q; state.quarterAuto = false; }
          const ap = state.settings.applicant;
          if (res.signerName && !ap.firstName && !ap.lastName) {
            const parts = res.signerName.split(/\s+/);
            ap.firstName = parts.shift() || '';
            ap.lastName = parts.join(' ');
            fillSettings();
          }
          f.status = 'done';
          f.msg = `נטענו ${res.shipments.length} משלוחים מהאקסל`;
        } else if (f.kind === 'json') {
          const data = JSON.parse(await f.file.text());
          if (!data || !Array.isArray(data.shipments)) throw new Error('not a backup');
          state = {
            version: 1,
            settings: Object.assign(defaultSettings(), data.settings || {}),
            year: data.year || state.year, q: data.q || state.q, quarterAuto: false,
            shipments: data.shipments.map(C.normalizeShipment), pages: data.pages || [], isExample: false,
          };
          fillSettings();
          f.status = 'done';
          f.msg = `גיבוי נטען · ${state.shipments.length} משלוחים`;
        }
      } catch (e) {
        f.status = 'error';
        f.msg = f.kind === 'excel' ? `לא ניתן לקרוא את האקסל: ${e.message || e}` : 'קובץ הגיבוי אינו תקין';
      }
    }
    autoQuarter();
    renderAll();
    save();
  }

  function upsert(s) {
    C.applyInvoiceBasis(s, state.settings.invoiceBasis || 'first');
    const k = C.awbKey(s.awb);
    let idx = k ? state.shipments.findIndex((o) => C.awbKey(o.awb) === k) : -1;
    if (idx < 0 && s.exportDecl) idx = state.shipments.findIndex((o) => o.exportDecl && o.exportDecl === s.exportDecl);
    if (idx >= 0) {
      const old = state.shipments[idx];
      s.id = old.id;
      s.included = old.included;
      if (old.edited && s.source === 'ai') s.notes = [...(s.notes || []), 'עודכן מחדש מהמסמכים; עריכות ידניות קודמות הוחלפו.'];
      state.shipments[idx] = s;
    } else {
      state.shipments.push(s);
    }
  }

  /* ------------------------------------------------------------- processing */

  let engineInfo = null;

  async function detectEngine() {
    const chip = $('engine');
    const inClaude = !!(window.claude && typeof window.claude.use === 'function');
    if (inClaude) {
      const eng = await E.createSampleEngine({ tier: state.settings.tier });
      if (eng && eng.usable) {
        engineInfo = 'sample';
        chip.textContent = 'קריאה באמצעות Claude';
        chip.className = 'engine on';
        return;
      }
    }
    engineInfo = 'api';
    $('apiKeyFld').hidden = false;
    chip.textContent = inClaude ? 'Claude לא זמין כאן · אפשר מפתח API' : 'נדרש מפתח API לקריאה';
    chip.className = 'engine off';
  }

  async function getEngine() {
    if (engineInfo === 'sample') {
      const eng = await E.createSampleEngine({ tier: state.settings.tier });
      if (eng && eng.usable) return eng;
    }
    const key = $('apiKey').value.trim();
    if (!key) return null;
    try {
      return await E.createApiEngine({ apiKey: key });
    } catch (e) {
      toast('לא ניתן לטעון את ספריית Anthropic. בדקו את החיבור לאינטרנט.');
      return null;
    }
  }

  async function processAll() {
    const docs = files.filter((f) => f.kind === 'doc' && (f.status === 'pending' || f.status === 'error'));
    if (!docs.length || running) return;
    const engine = await getEngine();
    if (!engine) {
      $('settingsBox').open = true;
      $('apiKey').focus();
      toast('כדי לקרוא מסמכים, פתחו את הכלי מתוך Claude או הזינו מפתח API בהגדרות.', 6000);
      return;
    }
    if (state.isExample) clearExample();
    running = new AbortController();
    const signal = running.signal;
    docs.forEach((f) => { f.status = 'queued'; f.msg = ''; f.done = 0; f.total = 0; });
    renderFiles();
    setStatus(`קורא ${docs.length} קבצים. כל קובץ נקרא עמוד אחר עמוד, וזה יכול לקחת כמה דקות לקובץ.`);

    const newPages = [];
    const tasks = docs.map((f) => async () => {
      if (signal.aborted) return;
      f.status = 'running';
      renderFiles();
      try {
        const facts = await E.readFilePages(engine, f.file, f.key, {
          signal,
          progress: (done, total, note) => { f.done = done; f.total = total; f.note = note; renderFiles(); },
        });
        state.pages = state.pages.filter((p) => p.fileIndex !== f.key).concat(facts);
        newPages.push(...facts);
        f.status = 'done';
        const kinds = new Set(facts.map((p) => p.type));
        f.msg = `נקרא · ${facts.length} עמודים${kinds.has('eu_import_declaration') ? '' : ' · לא נמצא IM4'}`;
        save();
      } catch (e) {
        f.status = 'error';
        f.msg = errCopy(e);
      }
      renderFiles();
    });
    await E.runQueue(tasks, 2, signal);

    if (newPages.length && !signal.aborted) {
      const fresh = new Set(newPages);
      const groups = C.groupPages(state.pages).filter(E.groupIsRelevant).filter((g) => g.pages.some((p) => fresh.has(p)));
      let done = 0;
      setStatus(`מרכיב ${groups.length} משלוחים מהעמודים שנקראו… (0/${groups.length})`);
      const build = groups.map((g) => async () => {
        try {
          const s = await E.buildShipment(engine, g, { signal });
          upsert(s);
          autoQuarter();
          renderAll();
          save();
        } catch (e) {
          if (e.code !== 'cancelled') toast(`משלוח ${g.awbs[0] || g.files[0]}: ${errCopy(e)}`, 6000);
        }
        done += 1;
        setStatus(`מרכיב ${groups.length} משלוחים מהעמודים שנקראו… (${done}/${groups.length})`);
      });
      await E.runQueue(build, 2, signal);
      setStatus(signal.aborted ? 'העיבוד נעצר. מה שכבר נקרא נשמר.' : `הסתיים: ${groups.length} משלוחים עודכנו. בדקו את השורות המסומנות בטבלה.`);
    } else {
      setStatus(signal.aborted ? 'העיבוד נעצר. מה שכבר נקרא נשמר.' : 'לא נקראו עמודים חדשים.');
    }
    running = null;
    renderAll();
    save();
  }

  const setStatus = (t) => { $('statusLine').textContent = t; };

  /* ------------------------------------------------------------------ table */

  function statusPill(s, issues) {
    if (s.included === false) return '<span class="pill">לא נכלל</span>';
    const c = C.issueCounts(issues);
    if (c.error) return `<span class="pill err">${c.error} שגיאות${c.warn ? ` · ${c.warn} אזהרות` : ''}</span>`;
    if (c.warn) return `<span class="pill warn">${c.warn} לבדיקה</span>`;
    return '<span class="pill ok">תקין</span>';
  }

  function rowCells(s, i, issues) {
    const hs = s.classifications.map((c) => `<span class="chip">${esc(c.hs || '?')}</span>`).join('');
    return `<td class="num">${i + 1}</td>
      <td class="mono">${esc(s.awb) || '—'}</td>
      <td class="mono">${esc(C.fmtDateShort(s.loadingDate)) || '—'}</td>
      <td class="mono">${esc(s.exportDecl) || '—'}</td>
      <td class="mono">${esc(C.im4Short(s.im4)) || '—'}</td>
      <td><div class="chips">${hs || '—'}</div></td>
      <td class="num money">${esc(C.fmtMoney(C.customsValueTotal(s)))}</td>
      <td class="num money"><b>${esc(C.fmtMoney(C.claimTotal(s)))}</b> ${esc(s.claimCurrency)}</td>
      <td>${statusPill(s, issues)}</td>`;
  }

  function renderTable() {
    const list = C.sortShipments(state.shipments);
    if (!list.length) {
      $('shipRows').innerHTML = '<tr><td colspan="9" class="empty">עדיין אין משלוחים. העלו מסמכים או טענו אקסל של רבעון.</td></tr>';
      $('shipFoot').innerHTML = '';
      return;
    }
    $('shipRows').innerHTML = list.map((s, i) => {
      const issues = issuesOf(s);
      const open = s.id === openId;
      return `<tr class="ship${open ? ' open' : ''}${s.included === false ? ' excluded' : ''}" data-id="${esc(s.id)}" tabindex="0" aria-expanded="${open}">${rowCells(s, i, issues)}</tr>${open ? `<tr class="editor"><td colspan="9">${editorHtml(s, issues)}</td></tr>` : ''}`;
    }).join('');
    renderFoot();
  }

  function renderFoot() {
    const sum = C.summarize(state.shipments, { year: state.year, q: state.q });
    $('shipFoot').innerHTML = `<tr><td colspan="6">סה"כ ${sum.count} משלוחים · ${sum.classificationCount} סיווגים</td><td class="num money">${esc(C.fmtByCurrency(sum.customsValue))}</td><td class="num money">${esc(C.fmtByCurrency(sum.duty))}</td><td></td></tr>`;
  }

  function refreshRow(s) {
    const list = C.sortShipments(state.shipments);
    const i = list.findIndex((o) => o.id === s.id);
    const tr = document.querySelector(`tr.ship[data-id="${CSS.escape(s.id)}"]`);
    const issues = issuesOf(s);
    if (tr) {
      tr.innerHTML = rowCells(s, i, issues);
      tr.classList.toggle('excluded', s.included === false);
    }
    const ul = document.querySelector('.editor .issues');
    if (ul) ul.outerHTML = issuesHtml(s, issues);
    renderFoot();
    renderKpis();
    schedulePreview();
    save();
  }

  /* ----------------------------------------------------------------- editor */

  const FIELDS = [
    ['חשבון עסקה ורשימון יצוא', [
      ['exportDecl', "מס' הצהרת יצוא", 'text'], ['invoiceNo', "מס' חשבון עסקה", 'text'], ['invoiceDate', 'תאריך חשבון', 'date'],
      ['invoiceTotal', 'ערך משלוח כולל למכס', 'num'], ['invoiceCurrency', 'מטבע', 'text'],
      ['eligibleParts', 'ערך יצוא זכאי לשיפוי (אפשר a+b+c)', 'parts'], ['eligibleCurrency', 'מטבע', 'text'],
      ['incoterm', 'מונח המכר', 'text'], ['netKg', 'משקל כולל נטו', 'num'],
    ]],
    ['שטר מטען', [
      ['awb', "מס' שטר מטען", 'text'], ['loadingDate', 'תאריך טעינה', 'date'], ['grossKg', 'משקל כולל ברוטו', 'num'],
      ['eur1', 'מספר EUR-1', 'text'], ['container', "מס' מכולה", 'text'],
    ]],
    ['הצהרת יבוא (IM4) ותשלום', [
      ['im4', "מס' הצהרת יבוא IM4", 'text'], ['releaseDate', 'תאריך שחרור הסחורה', 'date'], ['destCountry', 'מדינת יעד', 'text'],
      ['port', 'נמל', 'text'], ['claimCurrency', 'מטבע השיפוי', 'text'], ['dutyDocTotal', 'סכום המכס לפי המסמכים', 'num'],
      ['dutyEvidence', 'הוכחת תשלום', 'text'],
    ]],
  ];
  const CLS_FIELDS = [['hs', 'פרט מכס'], ['desc', 'תיאור'], ['site', 'אתר גידול'], ['zip', 'מיקוד גידול'], ['netKg', 'משקל נטו'], ['value', 'ערך למכס'], ['ratePct', 'אחוז מכס'], ['duty', 'סכום המכס']];
  const NUM_CLS = new Set(['netKg', 'value', 'ratePct', 'duty']);

  function fieldValue(s, f, type) {
    if (type === 'parts') return s.eligibleParts.length > 1 ? s.eligibleParts.map((n) => C.round2(n)).join('+') : (s.eligibleValue ?? '');
    const v = s[f];
    return v === null || v === undefined ? '' : v;
  }

  function issuesHtml(s, issues) {
    const notes = (s.notes || []).map((n) => ({ level: 'info', msg: n }));
    const all = [...issues, ...notes];
    if (!all.length) return '<ul class="issues"><li class="info">אין הערות למשלוח הזה.</li></ul>';
    return `<ul class="issues">${all.map((it) => `<li class="${it.level}">${esc(it.msg)}</li>`).join('')}</ul>`;
  }

  function editorHtml(s, issues) {
    const bad = new Set(issues.filter((i) => i.level === 'error').map((i) => i.field));
    const groups = FIELDS.map(([title, fields]) => `<div class="ed-h">${esc(title)}</div><div class="ed-grid">${fields.map(([f, label, type]) =>
      `<label class="fld${bad.has(f) ? ' bad' : ''}"><span>${esc(label)}</span><input data-f="${f}" data-t="${type}" type="${type === 'date' ? 'date' : 'text'}" ${type === 'num' || type === 'parts' ? 'inputmode="decimal" dir="ltr"' : 'dir="auto"'} value="${esc(fieldValue(s, f, type))}"></label>`).join('')}</div>`).join('');
    const cls = `<div class="ed-h">סיווגים זכאים לשיפוי (שורה לכל פריט ב-IM4 שחויב במכס)</div>
      <div class="cls-wrap"><table class="cls"><thead><tr><th>#</th>${CLS_FIELDS.map(([, l]) => `<th>${esc(l)}</th>`).join('')}<th></th></tr></thead><tbody>
      ${s.classifications.map((c, i) => `<tr><td class="num">${i + 1}</td>${CLS_FIELDS.map(([f]) => `<td><input data-c="${i}" data-cf="${f}" value="${esc(c[f] ?? '')}" ${NUM_CLS.has(f) || f === 'hs' || f === 'zip' ? 'dir="ltr" inputmode="decimal"' : 'dir="auto"'}></td>`).join('')}<td><button class="btn small danger" type="button" data-act="delcls" data-i="${i}" aria-label="מחק סיווג ${i + 1}">מחק</button></td></tr>`).join('')}
      </tbody></table></div>
      <div class="row" style="margin-top:8px"><button class="btn small" type="button" data-act="addcls">הוסף סיווג</button></div>`;
    const src = s.files && s.files.length ? `<p class="src">מקור: ${s.files.map((f) => `<bdi>${esc(f)}</bdi>`).join(', ')}${s.source === 'excel' ? ' (אקסל)' : ''}</p>` : s.source === 'excel' ? '<p class="src">מקור: אקסל שנטען</p>' : '';
    return `${issuesHtml(s, issues)}${groups}${cls}
      <div class="row" style="margin-top:16px;justify-content:space-between">
        <label class="row" style="gap:6px"><input type="checkbox" data-act="include" ${s.included !== false ? 'checked' : ''}> לכלול בבקשה לרבעון</label>
        <div class="row"><button class="btn small" type="button" data-act="close">סגור</button><button class="btn small danger" type="button" data-act="delete">מחק משלוח</button></div>
      </div>${src}`;
  }

  function currentOpen() { return state.shipments.find((s) => s.id === openId); }

  function onEditorChange(ev) {
    const s = currentOpen();
    if (!s) return;
    const t = ev.target;
    if (t.dataset.f) {
      const f = t.dataset.f;
      const type = t.dataset.t;
      const raw = t.value;
      if (type === 'num') {
        s[f] = C.num(raw);
        if (f === 'invoiceTotal' || f === 'netKg') s.totalsEdited = true;
      }
      else if (type === 'date') s[f] = C.parseDate(raw);
      else if (type === 'parts') {
        const parts = raw.split('+').map(C.num).filter((n) => n !== null);
        s.eligibleParts = parts.length > 1 ? parts : [];
        s.eligibleValue = parts.length ? C.round2(C.sum(parts)) : null;
      } else if (f === 'awb') s.awb = C.normAwb(raw);
      else if (f === 'claimCurrency' || f === 'invoiceCurrency' || f === 'eligibleCurrency') s[f] = C.normCurrency(raw);
      else s[f] = raw.trim();
    } else if (t.dataset.cf) {
      const c = s.classifications[+t.dataset.c];
      const f = t.dataset.cf;
      if (!c) return;
      if (NUM_CLS.has(f)) c[f] = C.num(t.value);
      else if (f === 'hs') c[f] = C.normHs(t.value);
      else c[f] = t.value.trim();
    } else if (t.dataset.act === 'include') {
      s.included = t.checked;
    } else return;
    s.edited = true;
    refreshRow(s);
  }

  function onTableClick(ev) {
    const act = ev.target.closest('[data-act]');
    if (act && act.tagName === 'BUTTON') {
      const s = currentOpen();
      if (!s) return;
      const a = act.dataset.act;
      if (a === 'addcls') s.classifications.push(C.emptyClassification());
      if (a === 'delcls') s.classifications.splice(+act.dataset.i, 1);
      if (a === 'close') openId = null;
      if (a === 'delete') {
        if (act.dataset.confirm !== '1') { act.dataset.confirm = '1'; act.textContent = 'ללחוץ שוב למחיקה'; return; }
        state.shipments = state.shipments.filter((o) => o.id !== s.id);
        openId = null;
        toast('המשלוח נמחק');
      }
      s.edited = true;
      renderAll();
      save();
      return;
    }
    const tr = ev.target.closest('tr.ship');
    if (tr) {
      openId = openId === tr.dataset.id ? null : tr.dataset.id;
      renderTable();
    }
  }

  /* ---------------------------------------------------------------- preview */

  let previewTimer = null;
  function schedulePreview() { clearTimeout(previewTimer); previewTimer = setTimeout(renderPreview, 350); }

  function renderPreview() {
    const box = $('preview');
    const pages = S.quarterlyPages(meta(), state.shipments).slice(0, 3);
    box.innerHTML = '';
    pages.forEach((html, i) => {
      const sheet = document.createElement('div');
      sheet.className = 'sheet';
      sheet.setAttribute('aria-label', `עמוד ${i + 1}`);
      const root = sheet.attachShadow ? sheet.attachShadow({ mode: 'open' }) : sheet;
      root.innerHTML = `<style>${S.previewCss('')}.scaler{transform:scale(.35);transform-origin:top right;position:absolute;top:0;right:0}</style><div class="scaler">${html}</div>`;
      box.appendChild(sheet);
    });
  }

  /* ---------------------------------------------------------------- outputs */

  const outName = (kind, ext) => `${kind} רבעון ${state.q}-${state.year}.${ext}`;

  async function withBusy(btn, label, fn) {
    const old = btn.textContent;
    btn.disabled = true;
    btn.textContent = label;
    try { await fn((t) => { btn.textContent = t; }); } catch (e) { toast(`ההפקה נכשלה: ${e.message || e}`, 6000); }
    btn.textContent = old;
    btn.disabled = false;
  }

  async function downloadExcel(ev) {
    if (!included().length) { toast('אין משלוחים להפקה.'); return; }
    await withBusy(ev.currentTarget, 'מכין אקסל…', async () => {
      const wb = X.buildWorkbook(window.ExcelJS, state.shipments, meta());
      const buf = await wb.xlsx.writeBuffer();
      await saveFile(outName('שיפוי', 'xlsx'), new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    });
  }

  async function downloadPdf(ev, which) {
    if (!included().length) { toast('אין משלוחים להפקה.'); return; }
    await withBusy(ev.currentTarget, 'מכין PDF…', async (setLabel) => {
      const m = meta();
      const pages = which === 'forms' ? S.shipmentFormPages(m, state.shipments) : S.quarterlyPages(m, state.shipments);
      const blob = await S.renderPdf(pages, {
        jsPDF: window.jspdf && window.jspdf.jsPDF, fontCss: window.ShipoFonts && window.ShipoFonts.css,
        onProgress: (d, t) => setLabel(`מכין PDF… ${d}/${t}`),
      });
      await saveFile(outName(which === 'forms' ? 'טפסים לפי משלוח' : 'דף סיכום', 'pdf'), blob);
    });
  }

  async function downloadBackup() {
    const data = { ...state, exportedAt: new Date().toISOString() };
    await saveFile(outName('גיבוי שיפוי', 'json'), new Blob([JSON.stringify(data, null, 1)], { type: 'application/json' }));
  }

  /* --------------------------------------------------------------- settings */

  const SET_FIELDS = [
    ['co_', 'company', ['name', 'id', 'street', 'houseNo', 'city', 'zone', 'zip', 'phone']],
    ['ap_', 'applicant', ['firstName', 'lastName', 'idNo', 'role', 'phone']],
  ];

  function fillSettings() {
    const st = state.settings;
    for (const [prefix, group, keys] of SET_FIELDS) for (const k of keys) $(prefix + k).value = st[group][k] || '';
    $('goodsDescription').value = st.goodsDescription || '';
    $('signDate').value = st.signDate || '';
    $('tier').value = st.tier || 'default';
    $('invoiceBasis').value = st.invoiceBasis || 'first';
    const atts = st.attachments && st.attachments.length ? st.attachments : S.DEFAULT_ATTACHMENTS.map((a) => ({ ...a }));
    st.attachments = atts;
    $('attachments').innerHTML = '<div class="set-h" style="margin:0">מסמכים מצורפים (מסומנים בטופס)</div>' + atts.map((a, i) =>
      `<label><input type="checkbox" data-att="${i}" ${a.checked ? 'checked' : ''}> <span>${esc(a.label)}</span></label>`).join('');
  }

  function onSettingsInput(ev) {
    const t = ev.target;
    const st = state.settings;
    if (t.dataset.att !== undefined) { st.attachments[+t.dataset.att].checked = t.checked; }
    else if (t.id === 'goodsDescription') st.goodsDescription = t.value;
    else if (t.id === 'signDate') st.signDate = t.value || null;
    else if (t.id === 'tier') st.tier = t.value;
    else if (t.id === 'invoiceBasis') {
      if (ev.type !== 'change') return;
      st.invoiceBasis = t.value;
      state.shipments.forEach((s) => C.applyInvoiceBasis(s, st.invoiceBasis));
      renderAll();
    }
    else if (t.id === 'apiKey') return;
    else {
      for (const [prefix, group] of SET_FIELDS) {
        if (t.id.startsWith(prefix)) st[group][t.id.slice(prefix.length)] = t.value;
      }
    }
    schedulePreview();
    save();
  }

  /* ------------------------------------------------------------------- init */

  function renderAll() {
    renderQuarter();
    renderKpis();
    renderTable();
    renderFiles();
    schedulePreview();
  }

  function init() {
    load();
    // fonts used by the summary pages, so the preview matches the PDF
    if (window.ShipoFonts) {
      const st = document.createElement('style');
      st.textContent = window.ShipoFonts.css;
      document.head.appendChild(st);
    }
    if (window.pdfjsLib && window.pdfjsWorker) {
      // the worker runs on the page itself (the frame cannot load cross-origin workers)
      window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
    }
    fillSettings();
    renderAll();

    $('qYear').addEventListener('change', (e) => { state.year = +e.target.value; state.quarterAuto = false; renderAll(); save(); });
    $('qQuarter').addEventListener('change', (e) => { state.q = +e.target.value; state.quarterAuto = false; renderAll(); save(); });
    $('clearExample').addEventListener('click', () => { clearExample(); renderAll(); save(); toast('הדוגמה נמחקה'); });

    const drop = $('drop');
    drop.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('over'); });
    drop.addEventListener('dragleave', () => drop.classList.remove('over'));
    drop.addEventListener('drop', (e) => { e.preventDefault(); drop.classList.remove('over'); addFiles([...e.dataTransfer.files]); });
    $('fileInput').addEventListener('change', (e) => { addFiles([...e.target.files]); e.target.value = ''; });
    $('processBtn').addEventListener('click', processAll);
    $('stopBtn').addEventListener('click', () => { if (running) running.abort(); setStatus('עוצר אחרי הבקשה הנוכחית…'); });
    $('clearFiles').addEventListener('click', () => { files.splice(0, files.length); renderFiles(); setStatus(''); });

    $('addManual').addEventListener('click', () => {
      if (state.isExample) clearExample();
      const s = C.emptyShipment();
      s.classifications.push(C.emptyClassification());
      state.shipments.push(s);
      openId = s.id;
      renderAll();
      save();
    });
    const rows = $('shipRows');
    rows.addEventListener('click', onTableClick);
    rows.addEventListener('change', onEditorChange);
    rows.addEventListener('keydown', (e) => {
      if ((e.key === 'Enter' || e.key === ' ') && e.target.matches('tr.ship')) { e.preventDefault(); onTableClick(e); }
    });

    $('settingsForm').addEventListener('input', onSettingsInput);
    $('settingsForm').addEventListener('change', onSettingsInput);
    $('dlExcel').addEventListener('click', downloadExcel);
    $('dlSummary').addEventListener('click', (e) => downloadPdf(e, 'summary'));
    $('dlForms').addEventListener('click', (e) => downloadPdf(e, 'forms'));
    $('dlBackup').addEventListener('click', downloadBackup);

    detectEngine();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
