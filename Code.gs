/**
 * Mail Merge with inline images — 2026 edition
 * © David Sung (2019–2026) · https://github.com/david097/GAS-Mail-Merge-with-inline-images
 *
 * Sends personalised emails from a Gmail draft to every row of a Google Sheet.
 *  - Messages are built as raw MIME and sent through the Gmail API, so emoji (4-byte characters)
 *    in the subject, body and sender name survive, and Gmail "Send mail as" aliases work.
 *  - Inline images and attachments in the draft are preserved.
 *  - Optional: a Google Slides page rendered per recipient and embedded with {{SlideImage}}.
 *  - UI text is localised (see i18n.gs and Lang_*.gs).
 *
 * Required: Apps Script editor > Services (+) > Gmail API
 *           (and Google Slides API if you use the slide image feature).
 *           Both are already declared in appsscript.json.
 */

const CONFIG = {
  MENU_NAME: '▼Mail Merge▼',
  STATUS_HEADER: 'Mail Merge Status', // kept in English for compatibility with existing sheets
  STATUS_SENT: 'EMAIL_SENT',
  SENT_AT_HEADER: 'Mail Merge Sent At', // last successful send time (also kept in English)
  // Headers recognised as the email column (compared after lower-casing and removing spaces/symbols)
  EMAIL_HEADERS: ['email', 'emailaddress', 'mail', 'mailaddress', '이메일', '이메일주소', '메일', '메일주소'],
  MAX_DRAFTS: 20,
  MAX_SLIDES: 20,
  TIME_LIMIT_MS: 5 * 60 * 1000, // stop safely before Apps Script's 6-minute execution limit
  SLIDE_CID: 'mm_slide_image',
  PROP_SENDER_NAME: 'mm_sender_name',
};

/* ======================================================================
 * Menu / dialogs
 * ====================================================================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(CONFIG.MENU_NAME)
    .addItem(t('menu_open'), 'showSidebar')
    .addItem(t('menu_clear'), 'clearSheet')
    .addSeparator()
    .addItem(t('menu_manual'), 'showManual')
    .addToUi();
}

function showSidebar() {
  const lang = getLang_();
  const tpl = HtmlService.createTemplateFromFile('MailMerge');
  tpl.lang = lang;
  tpl.s = getStrings_(lang);
  tpl.languages = getLanguages_();
  const html = tpl.evaluate().setTitle(t('sidebar_title'));
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Shows Manual_<lang>.html, falling back to the English manual. */
function showManual() {
  let html;
  try {
    html = HtmlService.createHtmlOutputFromFile('Manual_' + getLang_());
  } catch (e) {
    html = HtmlService.createHtmlOutputFromFile('Manual_' + I18N.DEFAULT_LANG);
  }
  SpreadsheetApp.getUi().showModalDialog(html.setWidth(580).setHeight(540), t('manual_title'));
}

/** Clears everything below the header row of the active sheet. */
function clearSheet() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const answer = ui.alert(t('clear_title'), t('clear_confirm', { sheet: sheet.getName() }), ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  const rows = sheet.getMaxRows() - 1;
  if (rows > 0) sheet.getRange(2, 1, rows, sheet.getMaxColumns()).clearContent();
}

/* ======================================================================
 * Functions called from the sidebar (google.script.run)
 * ====================================================================== */

/** Everything the sidebar needs, in a single round trip. */
function getSidebarData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const senders = listSenders_();
  return {
    drafts: listDrafts_(),
    senders: senders,
    sheets: ss.getSheets().map(s => s.getName()),
    activeSheet: ss.getActiveSheet().getName(),
    senderName: getDefaultSenderName_(senders),
  };
}

function getDrafts() {
  return listDrafts_();
}

/** Saves the chosen language and reopens the sidebar in that language. */
function changeLanguage(code) {
  setLanguage(code);
  showSidebar();
}

/** Google Slides files in the user's Drive. */
function getSlides() {
  const files = DriveApp.searchFiles(`mimeType = '${MimeType.GOOGLE_SLIDES}' and trashed = false`);
  const list = [];
  while (files.hasNext() && list.length < CONFIG.MAX_SLIDES) {
    const f = files.next();
    list.push({ name: f.getName(), url: f.getUrl() });
  }
  return list;
}

function createSlide() {
  const pres = SlidesApp.create(t('slide_new_name', { date: new Date().toLocaleString() }));
  return { name: pres.getName(), url: pres.getUrl() };
}

/** Pre-send check: how many rows will be sent / were already sent / have an invalid address. */
function getSendPreview(sheetName) {
  const table = readTable_(getSheet_(sheetName), false);
  let pending = 0, sent = 0, invalid = 0;
  table.rows.forEach(r => {
    if (r.sent) sent++;
    else if (!isValidEmail_(r.email)) invalid++;
    else pending++;
  });
  return { pending: pending, sent: sent, invalid: invalid };
}

/**
 * Runs the mail merge.
 * @param {Object} opts {draftId, sheetName, senderEmail, senderName, cc, bcc, mode: 'test'|'run', slideUrl}
 */
function startMailMerge(opts) {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(1000)) throw new Error(t('err_busy'));
  try {
    return runMailMerge_(opts);
  } finally {
    lock.releaseLock();
  }
}

/* ======================================================================
 * Mail merge core
 * ====================================================================== */

function runMailMerge_(opts) {
  const startedAt = Date.now();
  const isTest = opts.mode === 'test';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();

  if (typeof Gmail === 'undefined') throw new Error(t('err_gmail_service'));
  if (!opts.draftId) throw new Error(t('err_no_draft_selected'));

  const sheet = getSheet_(opts.sheetName);
  const table = readTable_(sheet, !isTest); // test mode never modifies the sheet
  const template = loadTemplate_(opts.draftId);
  const cc = parseAddressList_(opts.cc, 'CC');
  const bcc = parseAddressList_(opts.bcc, 'BCC');
  const me = Session.getActiveUser().getEmail();
  const from = formatAddress_(opts.senderName, opts.senderEmail || me);

  if (opts.senderName) {
    PropertiesService.getUserProperties().setProperty(CONFIG.PROP_SENDER_NAME, opts.senderName);
  }

  // Check API availability once up front, so a disabled API produces one clear message
  // instead of an error on every row.
  checkApi_('Gmail', () => Gmail.Users.getProfile('me'));

  let slideId = opts.slideUrl ? SlidesApp.openByUrl(opts.slideUrl).getId() : null;
  if (slideId && !template.usesSlideImage) slideId = null; // no {{SlideImage}} in the draft → skip slides
  if (slideId) {
    if (typeof Slides === 'undefined') throw new Error(t('err_slides_service'));
    checkApi_('Google Slides', () => Slides.Presentations.get(slideId, { fields: 'presentationId' }));
  }

  const log = [];
  const summary = { sent: 0, skipped: 0, failed: 0, remaining: 0, stoppedEarly: false, missingFields: [] };
  const missing = {};

  const targets = isTest ? table.rows.slice(0, 1) : table.rows;
  if (targets.length === 0) throw new Error(t('err_no_rows'));

  ss.toast(isTest ? t('toast_test') : t('toast_run'), 'Mail Merge', -1);

  for (let i = 0; i < targets.length; i++) {
    const row = targets[i];

    if (!isTest && Date.now() - startedAt > CONFIG.TIME_LIMIT_MS) {
      summary.stoppedEarly = true;
      summary.remaining = targets.slice(i).filter(r => !r.sent).length;
      break;
    }
    if (!isTest && row.sent) {
      summary.skipped++;
      log.push({ type: 'skip', text: t('log_already_sent', { email: row.email, row: row.rowNumber }) });
      continue;
    }

    const to = isTest ? me : row.email;
    if (!isValidEmail_(to)) {
      summary.failed++;
      setStatus_(sheet, row, table.statusCol, 'ERROR: ' + t('status_invalid_email'), isTest);
      log.push({ type: 'error', text: t('log_invalid_email', { row: row.rowNumber, value: to || t('value_empty') }) });
      continue;
    }

    try {
      const msg = renderMessage_(template, row, slideId, missing);
      const mime = buildMime_({
        from: from,
        replyTo: opts.senderEmail || me,
        to: to,
        cc: isTest ? [] : cc,
        bcc: isTest ? [] : bcc,
        subject: (isTest ? '[TEST] ' : '') + msg.subject,
        html: msg.html,
        plain: msg.plain,
        inline: msg.inline,
        attachments: template.attachments,
      });
      withRetry_(() => sendRawMime_(mime));
      summary.sent++;
      markSent_(sheet, row, table, isTest);
      log.push({ type: 'ok', text: isTest ? t('log_test_sent', { to: to, row: row.rowNumber }) : to });
    } catch (e) {
      summary.failed++;
      setStatus_(sheet, row, table.statusCol, 'ERROR: ' + e.message, isTest);
      log.push({ type: 'error', text: `${to} — ${e.message}` });
    }
  }

  summary.missingFields = Object.keys(missing);
  ss.toast(t('toast_done'), 'Mail Merge', 3);
  showResultDialog_(opts, sheet, template, summary, log, isTest, tz);
  return summary;
}

/** Fills subject, HTML body, plain body and inline images for one row. */
function renderMessage_(template, row, slideId, missing) {
  const inline = template.inline.slice();
  let slideCid = null;
  if (slideId) {
    inline.push({ cid: CONFIG.SLIDE_CID, blob: renderSlideImage_(slideId, row.pairs).setName('slide.png') });
    slideCid = CONFIG.SLIDE_CID;
  }
  return {
    subject: fillTemplate_(template.subject, row.lookup, { html: false, missing: missing }),
    html: fillTemplate_(template.html, row.lookup, { html: true, slideCid: slideCid, missing: missing }),
    plain: fillTemplate_(template.plain, row.lookup, { html: false, missing: missing }),
    inline: inline,
  };
}

/* ======================================================================
 * Reading the sheet
 * ====================================================================== */

function getSheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error(t('err_sheet_not_found', { name: name }));
  return sheet;
}

/**
 * Makes sure the "Mail Merge Status" and "Mail Merge Sent At" columns exist (run mode only).
 * - Status: appended after the last column if missing.
 * - Sent At: placed right after Status. If that column already holds data, a new column is
 *   inserted instead, so existing data is never overwritten. Later runs find it by name.
 */
function ensureStatusColumns_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(h => String(h).trim());

  let statusCol = headers.indexOf(CONFIG.STATUS_HEADER) + 1;
  if (!statusCol) {
    statusCol = lastCol + 1;
    ensureColumnExists_(sheet, statusCol);
    sheet.getRange(1, statusCol).setValue(CONFIG.STATUS_HEADER);
  }

  if (headers.indexOf(CONFIG.SENT_AT_HEADER) < 0) {
    const col = statusCol + 1;
    if (col > sheet.getMaxColumns()) {
      sheet.insertColumnAfter(statusCol);
    } else {
      const rows = Math.max(sheet.getLastRow(), 1);
      const used = sheet.getRange(1, col, rows, 1).getDisplayValues().some(r => String(r[0]).trim() !== '');
      if (used) sheet.insertColumnAfter(statusCol);
    }
    sheet.getRange(1, col).setValue(CONFIG.SENT_AT_HEADER);
  }
}

function ensureColumnExists_(sheet, col) {
  const max = sheet.getMaxColumns();
  if (col > max) sheet.insertColumnsAfter(max, col - max);
}

/**
 * Reads the sheet into row objects.
 * - Keeps the real sheet row number, so statuses land on the right row even with blank rows.
 * - Uses getDisplayValues so dates and numbers appear exactly as formatted in the sheet.
 * - The status / sent-at columns are not available as merge fields.
 */
function readTable_(sheet, createStatusColumns) {
  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) {
    throw new Error(t('err_sheet_empty', { name: sheet.getName() }));
  }
  if (createStatusColumns) ensureStatusColumns_(sheet);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(h => String(h).trim());

  let statusIdx = headers.indexOf(CONFIG.STATUS_HEADER);
  if (statusIdx < 0) { // preview only: the column doesn't exist yet
    statusIdx = headers.length;
    headers.push(CONFIG.STATUS_HEADER);
  }
  const sentAtIdx = headers.indexOf(CONFIG.SENT_AT_HEADER);

  const emailIdx = headers.findIndex(h => CONFIG.EMAIL_HEADERS.indexOf(normalizeKey_(h)) >= 0);
  if (emailIdx < 0) throw new Error(t('err_no_email_column'));

  const rows = [];
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getDisplayValues();
    values.forEach((cells, i) => {
      if (cells.every(c => String(c).trim() === '')) return; // skip completely empty rows
      const pairs = [];
      headers.forEach((h, c) => {
        if (h && c !== statusIdx && c !== sentAtIdx) pairs.push([h, String(cells[c])]);
      });
      rows.push({
        rowNumber: i + 2,
        email: String(cells[emailIdx]).trim(),
        sent: isSentStatus_(cells[statusIdx]),
        pairs: pairs,
        lookup: buildLookup_(pairs),
      });
    });
  }
  return { headers: headers, rows: rows, statusCol: statusIdx + 1, sentAtCol: sentAtIdx + 1 };
}

/**
 * True when the status cell means "already sent".
 * Case, spaces, '_' and '-' are ignored: EMAIL_SENT, Email_Sent, email sent, EMAIL-SENT all match.
 */
function isSentStatus_(value) {
  const norm = s => String(s || '').toUpperCase().replace(/[\s_-]+/g, '');
  return norm(value) === norm(CONFIG.STATUS_SENT);
}

function setStatus_(sheet, row, statusCol, text, isTest) {
  if (isTest) return;
  sheet.getRange(row.rowNumber, statusCol).setValue(String(text).slice(0, 300));
}

/** Marks a row as sent and records the send time (success only; never in test mode). */
function markSent_(sheet, row, table, isTest) {
  if (isTest) return;
  setStatus_(sheet, row, table.statusCol, CONFIG.STATUS_SENT, false);
  if (table.sentAtCol > 0) {
    sheet.getRange(row.rowNumber, table.sentAtCol).setValue(new Date()).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }
}

/* ======================================================================
 * Template (Gmail draft)
 * ====================================================================== */

function listDrafts_() {
  const tz = Session.getScriptTimeZone();
  return GmailApp.getDrafts().slice(0, CONFIG.MAX_DRAFTS).map(d => {
    const m = d.getMessage();
    const subject = m.getSubject() || t('draft_no_subject');
    return {
      id: d.getId(),
      label: truncate_(subject, 40) + '  · ' + Utilities.formatDate(m.getDate(), tz, 'MM/dd HH:mm'),
    };
  });
}

function loadTemplate_(draftId) {
  let draft = null;
  try {
    draft = GmailApp.getDraft(draftId);
  } catch (e) {
    // handled below
  }
  if (!draft) throw new Error(t('err_draft_not_found'));

  const msg = draft.getMessage();
  const inlineBlobs = msg.getAttachments({ includeInlineImages: true, includeAttachments: false });
  const attachments = msg.getAttachments({ includeInlineImages: false, includeAttachments: true });
  const mapped = mapInlineImages_(msg.getBody(), inlineBlobs);

  return {
    subject: msg.getSubject() || '',
    html: mapped.html,
    plain: msg.getPlainBody() || '',
    inline: mapped.inline,
    attachments: attachments,
    usesSlideImage: /\{\{[^{}]*slide\s*image[^{}]*\}\}/i.test(mapped.html),
  };
}

/**
 * Links each <img> in the draft HTML to its inline image blob via Content-ID.
 * - Handles both src="cid:..." and data-surl="cid:..." (older Gmail markup).
 * - Matches by alt (file name) first, then falls back to document order.
 */
function mapInlineImages_(html, blobs) {
  const byName = {};
  blobs.forEach(b => { byName[normalizeFileName_(b.getName())] = b; });

  const inline = [];
  const used = [];
  const unmatched = [];
  const seen = {};

  const newHtml = html.replace(/<img\b[^>]*>/gi, tag => {
    const src = getAttr_(tag, 'src') || '';
    const surl = getAttr_(tag, 'data-surl') || '';
    const cidMatch = /^cid:(.+)$/i.exec(src) || /^cid:(.+)$/i.exec(surl);
    if (!cidMatch) return tag; // external images etc. are left as-is
    const cid = cidMatch[1].trim();

    if (!seen[cid]) {
      seen[cid] = true;
      const blob = byName[normalizeFileName_(getAttr_(tag, 'alt') || '')];
      if (blob && used.indexOf(blob) < 0) {
        inline.push({ cid: cid, blob: blob });
        used.push(blob);
      } else {
        unmatched.push(cid);
      }
    }
    // If src is not a cid: URL (e.g. "?view=att..."), point it at the cid.
    return /^cid:/i.test(src) ? tag : setAttr_(tag, 'src', 'cid:' + cid);
  });

  const rest = blobs.filter(b => used.indexOf(b) < 0);
  unmatched.forEach((cid, i) => {
    if (rest[i]) inline.push({ cid: cid, blob: rest[i] });
  });
  return { html: newHtml, inline: inline };
}

/* ======================================================================
 * Merge fields
 * ====================================================================== */

/** Header → value lookup, keyed by original name, lower-case name and normalised name. */
function buildLookup_(pairs) {
  const map = {};
  pairs.forEach(p => {
    const h = p[0], v = p[1];
    [h, h.toLowerCase(), normalizeKey_(h)].forEach(k => {
      if (k && !(k in map)) map[k] = v;
    });
  });
  return map;
}

/**
 * Replaces {{Field}} markers with row values.
 * - Recognises markers even when Gmail inserts tags inside them, e.g. {{<span>Name</span>}}.
 * - Values are HTML-escaped in the HTML body, with line breaks turned into <br>.
 * - Uses a replacer function so values containing "$" are inserted literally.
 */
function fillTemplate_(template, lookup, options) {
  if (!template) return '';
  return template.replace(/\{\{([^{}]{1,200})\}\}/g, (whole, inner) => {
    const key = cleanPlaceholder_(inner);
    if (normalizeKey_(key) === 'slideimage') {
      if (options.html && options.slideCid) {
        return `<img src="cid:${options.slideCid}" alt="slide" style="max-width:100%">`;
      }
      return '';
    }
    let value = lookup[key];
    if (value === undefined) value = lookup[key.toLowerCase()];
    if (value === undefined) value = lookup[normalizeKey_(key)];
    if (value === undefined) {
      if (options.missing) options.missing[key] = true;
      return '';
    }
    return options.html ? escapeHtml_(value).replace(/\r?\n/g, '<br>') : value;
  });
}

function cleanPlaceholder_(inner) {
  return inner
    .replace(/<[^>]*>/g, '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/["'\u201C\u201D\u2018\u2019]/g, '')
    .trim();
}

/** Lower-cases and keeps only letters/digits (Latin and Hangul): "Email Address" → "emailaddress". */
function normalizeKey_(s) {
  return String(s).toLowerCase().replace(/[^0-9a-z\u3131-\u318E\uAC00-\uD7A3]/g, '');
}

/* ======================================================================
 * Slide image ({{SlideImage}})
 * ====================================================================== */

/** Duplicates the first slide, fills in the row's values, renders it as PNG, then deletes the copy. */
function renderSlideImage_(presentationId, pairs) {
  const pres = SlidesApp.openById(presentationId);
  const copy = pres.getSlides()[0].duplicate(); // the first slide is the template
  const copyId = copy.getObjectId();
  try {
    pairs.forEach(p => copy.replaceAllText('{{' + p[0] + '}}', p[1]));
    pres.saveAndClose();

    const thumb = Slides.Presentations.Pages.getThumbnail(presentationId, copyId, {
      'thumbnailProperties.mimeType': 'PNG',
      'thumbnailProperties.thumbnailSize': 'LARGE',
    });
    return UrlFetchApp.fetch(thumb.contentUrl).getBlob();
  } finally {
    // Always remove the temporary copy so the presentation doesn't grow with every recipient.
    const p = SlidesApp.openById(presentationId);
    const s = p.getSlideById(copyId);
    if (s) s.remove();
    p.saveAndClose();
  }
}

/* ======================================================================
 * MIME building / sending (this is what makes emoji work)
 * ====================================================================== */

/**
 * Builds an RFC 2822 / 2045 / 2047 message.
 * Non-ASCII header text is encoded as =?UTF-8?B?...?= and bodies as UTF-8 base64,
 * so 4-byte characters such as emoji are preserved.
 *
 * Structure: mixed( related( alternative(text, html), inline images... ), attachments... )
 */
function buildMime_(m) {
  const textPart = {
    headers: ['Content-Type: text/plain; charset=UTF-8', 'Content-Transfer-Encoding: base64'],
    body: wrap76_(base64Utf8_(m.plain || htmlToText_(m.html))),
  };
  const htmlPart = {
    headers: ['Content-Type: text/html; charset=UTF-8', 'Content-Transfer-Encoding: base64'],
    body: wrap76_(base64Utf8_(m.html || '')),
  };
  let root = multipart_('alternative', [textPart, htmlPart]);

  const inline = m.inline || [];
  if (inline.length) {
    root = multipart_('related', [root].concat(inline.map(img => blobPart_(img.blob, 'inline', img.cid))));
  }
  const attachments = m.attachments || [];
  if (attachments.length) {
    root = multipart_('mixed', [root].concat(attachments.map(b => blobPart_(b, 'attachment'))));
  }

  const headers = [
    'MIME-Version: 1.0',
    'From: ' + m.from,
    'To: ' + sanitizeHeader_(m.to),
  ];
  if (m.replyTo) headers.push('Reply-To: ' + sanitizeHeader_(m.replyTo));
  if (m.cc && m.cc.length) headers.push('Cc: ' + m.cc.map(sanitizeHeader_).join(', '));
  if (m.bcc && m.bcc.length) headers.push('Bcc: ' + m.bcc.map(sanitizeHeader_).join(', '));
  headers.push('Subject: ' + encodeHeaderWord_(sanitizeHeader_(m.subject)));

  return renderPart_({ headers: headers.concat(root.headers), body: root.body });
}

function multipart_(subtype, parts) {
  const boundary = 'mm_' + subtype + '_' + Utilities.getUuid().replace(/-/g, '');
  const body = parts.map(p => '--' + boundary + '\r\n' + renderPart_(p)).join('\r\n') + '\r\n--' + boundary + '--';
  return { headers: [`Content-Type: multipart/${subtype}; boundary="${boundary}"`], body: body };
}

function blobPart_(blob, disposition, cid) {
  const name = encodeHeaderWord_(sanitizeHeader_(blob.getName() || 'file'), true);
  const type = blob.getContentType() || 'application/octet-stream';
  const headers = [
    `Content-Type: ${type}; name="${name}"`,
    `Content-Disposition: ${disposition}; filename="${name}"`,
    'Content-Transfer-Encoding: base64',
  ];
  if (cid) headers.push(`Content-ID: <${cid}>`, `X-Attachment-Id: ${cid}`);
  return { headers: headers, body: wrap76_(Utilities.base64Encode(blob.getBytes())) };
}

function renderPart_(part) {
  return part.headers.join('\r\n') + '\r\n\r\n' + part.body;
}

function sendRawMime_(mime) {
  return Gmail.Users.Messages.send({ raw: Utilities.base64EncodeWebSafe(mime, Utilities.Charset.UTF_8) }, 'me');
}

/** Retries transient failures (rate limits, backend errors) with back-off. */
function withRetry_(fn) {
  const waits = [1000, 3000, 8000];
  for (let i = 0; ; i++) {
    try {
      return fn();
    } catch (e) {
      const retryable = /rate|limit|429|500|503|backend|timed out|temporar/i.test(e.message);
      if (!retryable || i >= waits.length) throw e;
      Utilities.sleep(waits[i]);
    }
  }
}

/**
 * RFC 2047 encoded-word. ASCII text is returned as-is; anything else is UTF-8/base64 encoded.
 * Splits by character (not UTF-16 unit) so an emoji's surrogate pair is never cut in half.
 * @param {boolean} singleLine join words without folding (used for file names)
 */
function encodeHeaderWord_(text, singleLine) {
  text = String(text || '');
  if (/^[\x20-\x7E]*$/.test(text)) return text;
  const MAX_BYTES = 45; // 45 bytes → 60 base64 chars → 72-char encoded-word
  const words = [];
  let chunk = '', bytes = 0;
  Array.from(text).forEach(ch => {
    const n = utf8Length_(ch);
    if (bytes + n > MAX_BYTES && chunk) {
      words.push(chunk);
      chunk = '';
      bytes = 0;
    }
    chunk += ch;
    bytes += n;
  });
  if (chunk) words.push(chunk);
  return words.map(w => '=?UTF-8?B?' + base64Utf8_(w) + '?=').join(singleLine ? '' : '\r\n ');
}

/** "Name <email>", encoding the display name when it contains non-ASCII characters. */
function formatAddress_(name, email) {
  name = sanitizeHeader_(name || '').trim();
  email = sanitizeHeader_(email || '').trim();
  if (!name) return email;
  if (/^[\x20-\x7E]*$/.test(name)) return `"${name.replace(/(["\\])/g, '\\$1')}" <${email}>`;
  return `${encodeHeaderWord_(name)} <${email}>`;
}

/* ======================================================================
 * Result dialog
 * ====================================================================== */

function showResultDialog_(opts, sheet, template, summary, log, isTest, tz) {
  const color = { ok: '#1a73e8', skip: '#888', error: '#d93025' };
  const label = { ok: t('label_ok'), skip: t('label_skip'), error: t('label_error') };
  const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  const e = escapeHtml_;

  let html = '<div style="font:13px/1.6 sans-serif">';
  html += '<div style="background:#f1f3f4;padding:10px;border-radius:6px;color:#444">';
  html += `${e(t('result_subject'))}: ${e(template.subject)}<br>`;
  html += `${e(t('result_sheet'))}: ${e(sheet.getName())}<br>`;
  html += `${e(t('result_from'))}: ${e(opts.senderName || '')} &lt;${e(opts.senderEmail || '')}&gt;<br>`;
  if (!isTest && opts.cc) html += `CC: ${e(opts.cc)}<br>`;
  if (!isTest && opts.bcc) html += `BCC: ${e(opts.bcc)}<br>`;
  html += `${e(t('result_finished'))}: ${now}</div>`;

  html += `<p>${e(t('result_summary', { sent: summary.sent, skipped: summary.skipped, failed: summary.failed }))}</p>`;
  if (summary.stoppedEarly) {
    html += `<p style="color:#e37400"><b>${e(t('result_stopped', { n: summary.remaining }))}</b></p>`;
  }
  if (summary.missingFields.length) {
    const fields = summary.missingFields.map(f => '{{' + f + '}}').join(', ');
    html += `<p style="color:#e37400">${e(t('result_missing', { fields: fields }))}</p>`;
  }
  html += '<div style="max-height:260px;overflow:auto;border-top:1px solid #eee;padding-top:6px">';
  log.forEach(l => {
    html += `<div><span style="color:${color[l.type]}">${e(label[l.type])}</span> : ${e(l.text)}</div>`;
  });
  html += '</div></div>';

  const out = HtmlService.createHtmlOutput(html).setWidth(540).setHeight(460);
  SpreadsheetApp.getUi().showModalDialog(out, isTest ? t('result_title_test') : t('result_title_run'));
}

/* ======================================================================
 * Sender addresses
 * ====================================================================== */

/** Gmail "Send mail as" addresses (primary first). Falls back to GmailApp aliases. */
function listSenders_() {
  try {
    const res = Gmail.Users.Settings.SendAs.list('me');
    const list = (res.sendAs || [])
      .filter(s => s.isPrimary || s.verificationStatus === 'accepted')
      .map(s => ({ email: s.sendAsEmail, name: s.displayName || '', primary: !!s.isPrimary }));
    list.sort((a, b) => Number(b.primary) - Number(a.primary));
    if (list.length) return list;
  } catch (e) {
    // Gmail API not available → fall back below
  }
  const me = Session.getActiveUser().getEmail();
  return [{ email: me, name: '', primary: true }].concat(
    GmailApp.getAliases().map(a => ({ email: a, name: '', primary: false }))
  );
}

function getDefaultSenderName_(senders) {
  const saved = PropertiesService.getUserProperties().getProperty(CONFIG.PROP_SENDER_NAME);
  if (saved) return saved;
  const primary = (senders || []).filter(s => s.primary)[0];
  if (primary && primary.name) return primary.name;
  try {
    return DriveApp.getRootFolder().getOwner().getName();
  } catch (e) {
    return '';
  }
}

/* ======================================================================
 * Utilities
 * ====================================================================== */

/**
 * Calls an API once and turns "API disabled" errors into a readable message with the enable link.
 * (Scripts linked to a standard Google Cloud project must enable APIs in the Cloud Console.)
 */
function checkApi_(apiName, fn) {
  try {
    fn();
  } catch (e) {
    const msg = String(e.message || e);
    if (/has not been used in project|is disabled|SERVICE_DISABLED|accessNotConfigured/i.test(msg)) {
      const url = (msg.match(/https:\/\/console\.(?:developers|cloud)\.google\.com\/\S+/) || [''])[0].replace(/[.,)]+$/, '');
      throw new Error(t('err_api_disabled', { api: apiName, url: url }));
    }
    throw e;
  }
}

function parseAddressList_(text, label) {
  const list = String(text || '').split(/[,;\s]+/).map(s => s.trim()).filter(Boolean);
  const bad = list.filter(a => !isValidEmail_(a));
  if (bad.length) throw new Error(t('err_invalid_address', { label: label, list: bad.join(', ') }));
  return list;
}

function isValidEmail_(s) {
  return /^[^\s@<>,;"]+@[^\s@<>,;"]+\.[^\s@<>,;"]+$/.test(String(s || '').trim());
}

/** Prevents header injection via line breaks. */
function sanitizeHeader_(s) {
  return String(s || '').replace(/[\r\n]+/g, ' ');
}

function escapeHtml_(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function htmlToText_(html) {
  return String(html || '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|li|tr|h\d)>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/&amp;/gi, '&');
}

function getAttr_(tag, name) {
  const m = new RegExp('\\s' + name + '\\s*=\\s*("([^"]*)"|\'([^\']*)\')', 'i').exec(tag);
  if (!m) return null;
  return (m[2] !== undefined ? m[2] : m[3]).replace(/&amp;/g, '&');
}

function setAttr_(tag, name, value) {
  const re = new RegExp('(\\s' + name + '\\s*=\\s*)("[^"]*"|\'[^\']*\')', 'i');
  if (re.test(tag)) return tag.replace(re, (all, prefix) => prefix + '"' + value + '"');
  return tag.replace(/^<img\b/i, '<img ' + name + '="' + value + '"');
}

function normalizeFileName_(s) {
  return String(s || '').replace(/\s+/g, '').toLowerCase();
}

function base64Utf8_(s) {
  return Utilities.base64Encode(String(s), Utilities.Charset.UTF_8);
}

function wrap76_(b64) {
  return b64.replace(/.{1,76}/g, '$&\r\n').replace(/\r\n$/, '');
}

function utf8Length_(ch) {
  const cp = ch.codePointAt(0);
  return cp < 0x80 ? 1 : cp < 0x800 ? 2 : cp < 0x10000 ? 3 : 4;
}

/** Truncates by character so emoji are never cut in half. */
function truncate_(s, max) {
  const chars = Array.from(String(s));
  return chars.length > max ? chars.slice(0, max).join('') + '…' : chars.join('');
}
