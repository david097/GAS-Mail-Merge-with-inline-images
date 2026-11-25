/**
 * Mail Merge with inline images — 2026 정리판
 * Original © David Sung (2019~2023), 2023-11-15 버전 기준으로 정리
 *
 * 주요 변경
 *  - 발송을 GmailApp.sendEmail → Gmail API(raw MIME)로 변경
 *    · GmailApp 은 이모지(4바이트 문자)를 깨뜨리고, MailApp 은 보내는 주소(from) 지정이 안 됨
 *    · Gmail API 는 둘 다 해결 (보내는 주소 별칭 + 이모지 모두 정상)
 *  - 빈 행이 있으면 상태(EMAIL_SENT)가 엉뚱한 행에 기록되던 버그 수정
 *  - 한글 헤더({{이름}}) 지원, 헤더가 비어 있거나 한글이면 필드가 한 칸씩 밀리던 버그 수정
 *  - 초안 선택을 "목록 순번"이 아닌 "초안 ID"로 변경 (다른 초안이 발송되는 문제 방지)
 *  - 한 행 실패해도 전체가 멈추지 않고, 실패 사유를 상태 열에 기록
 *  - 6분 실행 제한 전에 안전하게 멈추고, 다시 실행하면 이어서 발송
 *  - 슬라이드 이미지 사용 시 원본 슬라이드에 복사본이 계속 쌓이던 문제 수정
 *  - 쓰지 않던 함수/파일 정리, 'Mail Merge.txt' 생성 꼼수 제거
 *
 * 필수 설정: Apps Script 편집기 > 서비스(+) > "Gmail API" 추가
 *           (슬라이드 이미지 기능을 쓰면 "Google Slides API"도 추가)
 */

const CONFIG = {
  MENU_NAME: '▼Mail Merge▼',
  STATUS_HEADER: 'Mail Merge Status',
  STATUS_SENT: 'EMAIL_SENT',
  // 이메일 열로 인식할 헤더 (소문자, 공백/기호 제거 후 비교)
  EMAIL_HEADERS: ['email', 'emailaddress', 'mail', 'mailaddress', '이메일', '이메일주소', '메일', '메일주소'],
  MAX_DRAFTS: 20,
  MAX_SLIDES: 20,
  TIME_LIMIT_MS: 5 * 60 * 1000, // Apps Script 6분 제한 전에 멈춤
  SLIDE_CID: 'mm_slide_image',
  PROP_SENDER_NAME: 'mm_sender_name',
};

/* ======================================================================
 * 메뉴 / 화면
 * ====================================================================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(CONFIG.MENU_NAME)
    .addItem('메일머지 열기', 'showSidebar')
    .addItem('현재 시트 내용 지우기', 'clearSheet')
    .addSeparator()
    .addItem('사용법', 'showManual')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('MailMerge').setTitle('Mail Merge');
  SpreadsheetApp.getUi().showSidebar(html);
}

function showManual() {
  const html = HtmlService.createHtmlOutputFromFile('UserManual').setWidth(560).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Mail Merge 사용법');
}

/** 현재 시트에서 머리행(1행)을 제외한 내용을 지움 */
function clearSheet() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const answer = ui.alert('주의', `'${sheet.getName()}' 시트의 내용(1행 제외)을 모두 지울까요?`, ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  const rows = sheet.getMaxRows() - 1;
  if (rows > 0) sheet.getRange(2, 1, rows, sheet.getMaxColumns()).clearContent();
}

/* ======================================================================
 * 사이드바에서 호출하는 함수 (google.script.run)
 * ====================================================================== */

/** 사이드바를 여는 데 필요한 정보를 한 번에 반환 */
function getSidebarData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const senders = listSenders_();
  return {
    drafts: listDrafts_(),
    senders: senders,
    sheets: ss.getSheets().map(s => s.getName()),
    activeSheet: ss.getActiveSheet().getName(),
    senderName: getDefaultSenderName_(senders),
    me: Session.getActiveUser().getEmail(),
  };
}

function getDrafts() {
  return listDrafts_();
}

/** 내 드라이브의 Google Slides 목록 */
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
  const pres = SlidesApp.create('Mail Merge 슬라이드 - ' + new Date().toLocaleString());
  return { name: pres.getName(), url: pres.getUrl() };
}

/** 발송 전 확인용: 대상/이미 발송/오류 건수 */
function getSendPreview(sheetName) {
  const table = readTable_(getSheet_(sheetName), false);
  let pending = 0, sent = 0, invalid = 0;
  table.rows.forEach(r => {
    if (r.status === CONFIG.STATUS_SENT) sent++;
    else if (!isValidEmail_(r.email)) invalid++;
    else pending++;
  });
  return { pending: pending, sent: sent, invalid: invalid };
}

/**
 * 메일머지 실행
 * @param {Object} opts {draftId, sheetName, senderEmail, senderName, cc, bcc, mode:'test'|'run', slideUrl}
 */
function startMailMerge(opts) {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(1000)) throw new Error('이미 메일머지가 실행 중입니다. 끝난 뒤 다시 시도해 주세요.');
  try {
    return runMailMerge_(opts);
  } finally {
    lock.releaseLock();
  }
}

/* ======================================================================
 * 메일머지 본체
 * ====================================================================== */

function runMailMerge_(opts) {
  const startedAt = Date.now();
  const isTest = opts.mode === 'test';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();

  if (typeof Gmail === 'undefined') {
    throw new Error('Gmail API 서비스가 추가되지 않았습니다. Apps Script 편집기 > 서비스(+) > Gmail API 를 추가해 주세요.');
  }
  if (!opts.draftId) throw new Error('메일 템플릿(임시보관함 초안)을 선택해 주세요.');

  const sheet = getSheet_(opts.sheetName);
  const table = readTable_(sheet, true);
  const template = loadTemplate_(opts.draftId);
  const cc = parseAddressList_(opts.cc, 'CC');
  const bcc = parseAddressList_(opts.bcc, 'BCC');
  const me = Session.getActiveUser().getEmail();
  const from = formatAddress_(opts.senderName, opts.senderEmail || me);

  if (opts.senderName) {
    PropertiesService.getUserProperties().setProperty(CONFIG.PROP_SENDER_NAME, opts.senderName);
  }

  const slideId = opts.slideUrl ? SlidesApp.openByUrl(opts.slideUrl).getId() : null;
  const log = [];
  const summary = { sent: 0, skipped: 0, failed: 0, remaining: 0, stoppedEarly: false, missingFields: [] };
  const missing = {};

  const targets = isTest ? table.rows.slice(0, 1) : table.rows;
  if (targets.length === 0) throw new Error('보낼 데이터 행이 없습니다.');

  ss.toast(isTest ? '테스트 메일 발송 중...' : '메일머지 발송 중...', 'Mail Merge', -1);

  for (let i = 0; i < targets.length; i++) {
    const row = targets[i];

    if (!isTest && Date.now() - startedAt > CONFIG.TIME_LIMIT_MS) {
      summary.stoppedEarly = true;
      summary.remaining = targets.slice(i).filter(r => r.status !== CONFIG.STATUS_SENT).length;
      break;
    }
    if (!isTest && row.status === CONFIG.STATUS_SENT) {
      summary.skipped++;
      log.push({ type: 'skip', text: `${row.email} (${row.rowNumber}행) — 이미 발송됨` });
      continue;
    }

    const to = isTest ? me : row.email;
    if (!isValidEmail_(to)) {
      summary.failed++;
      setStatus_(sheet, row, table.statusCol, 'ERROR: 이메일 주소 없음/형식 오류', isTest);
      log.push({ type: 'error', text: `${row.rowNumber}행 — 이메일 주소가 비었거나 형식이 잘못됨 (${to || '빈 칸'})` });
      continue;
    }

    try {
      const msg = renderMessage_(template, row, slideId, missing);
      const mime = buildMime_({
        from: from,
        replyTo: opts.senderEmail || me, // 2023 버전의 replyTo 동작 유지
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
      setStatus_(sheet, row, table.statusCol, CONFIG.STATUS_SENT, isTest);
      log.push({ type: 'ok', text: isTest ? `테스트 메일 → ${to} (${row.rowNumber}행 데이터 사용)` : `${to}` });
    } catch (e) {
      summary.failed++;
      setStatus_(sheet, row, table.statusCol, 'ERROR: ' + e.message, isTest);
      log.push({ type: 'error', text: `${to} — ${e.message}` });
    }
  }

  summary.missingFields = Object.keys(missing);
  ss.toast('완료되었습니다.', 'Mail Merge', 3);
  showResultDialog_(opts, sheet, template, summary, log, isTest, tz);
  return summary;
}

/** 한 행의 데이터로 제목/본문/인라인 이미지를 완성 */
function renderMessage_(template, row, slideId, missing) {
  const inline = template.inline.slice();
  let slideCid = null;
  if (slideId && template.usesSlideImage) {
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
 * 시트 읽기
 * ====================================================================== */

function getSheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error(`'${name}' 시트를 찾을 수 없습니다.`);
  return sheet;
}

/**
 * 시트를 읽어 행 목록으로 반환.
 * - 실제 시트 행 번호(rowNumber)를 함께 보관 → 빈 행이 있어도 상태가 올바른 행에 기록됨
 * - getDisplayValues 사용 → 날짜/숫자가 시트에 보이는 그대로 들어감
 */
function readTable_(sheet, createStatusColumn) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) throw new Error(`'${sheet.getName()}' 시트가 비어 있습니다.`);

  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(h => String(h).trim());

  let statusIdx = headers.indexOf(CONFIG.STATUS_HEADER);
  if (statusIdx < 0) {
    statusIdx = headers.length;
    headers.push(CONFIG.STATUS_HEADER);
    if (createStatusColumn) sheet.getRange(1, statusIdx + 1).setValue(CONFIG.STATUS_HEADER);
  }

  const emailIdx = headers.findIndex(h => CONFIG.EMAIL_HEADERS.indexOf(normalizeKey_(h)) >= 0);
  if (emailIdx < 0) {
    throw new Error("이메일 열을 찾을 수 없습니다. 이메일이 들어있는 열의 머리행을 'Email' 또는 '이메일'로 적어주세요.");
  }

  const rows = [];
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getDisplayValues();
    values.forEach((cells, i) => {
      if (cells.every(c => String(c).trim() === '')) return; // 완전히 빈 행은 건너뜀
      const pairs = [];
      headers.forEach((h, c) => {
        if (h && c !== statusIdx) pairs.push([h, String(cells[c])]);
      });
      rows.push({
        rowNumber: i + 2,
        email: String(cells[emailIdx]).trim(),
        status: String(cells[statusIdx]).trim(),
        pairs: pairs,
        lookup: buildLookup_(pairs),
      });
    });
  }
  return { headers: headers, rows: rows, statusCol: statusIdx + 1 };
}

function setStatus_(sheet, row, statusCol, text, isTest) {
  if (isTest) return;
  sheet.getRange(row.rowNumber, statusCol).setValue(String(text).slice(0, 300));
}

/* ======================================================================
 * 템플릿 (Gmail 초안)
 * ====================================================================== */

function listDrafts_() {
  const tz = Session.getScriptTimeZone();
  return GmailApp.getDrafts().slice(0, CONFIG.MAX_DRAFTS).map(d => {
    const m = d.getMessage();
    const subject = m.getSubject() || '(제목 없음)';
    return {
      id: d.getId(),
      label: truncate_(subject, 40) + '  · ' + Utilities.formatDate(m.getDate(), tz, 'MM/dd HH:mm'),
    };
  });
}

function loadTemplate_(draftId) {
  let draft;
  try {
    draft = GmailApp.getDraft(draftId);
  } catch (e) {
    draft = null;
  }
  if (!draft) throw new Error('선택한 초안을 찾을 수 없습니다. 사이드바에서 초안 목록을 새로고침해 주세요.');

  const msg = draft.getMessage();
  const inlineBlobs = msg.getAttachments({ includeInlineImages: true, includeAttachments: false });
  const attachments = msg.getAttachments({ includeInlineImages: false, includeAttachments: true });
  const mapped = mapInlineImages_(msg.getBody(), inlineBlobs);
  const html = mapped.html;

  return {
    subject: msg.getSubject() || '',
    html: html,
    plain: msg.getPlainBody() || '',
    inline: mapped.inline,
    attachments: attachments,
    usesSlideImage: /\{\{[^{}]*slide\s*image[^{}]*\}\}/i.test(html),
  };
}

/**
 * 초안 HTML의 <img> 와 인라인 이미지 Blob 을 Content-ID 로 연결.
 * - src="cid:..." 또는 data-surl="cid:..." 둘 다 처리 (예전 코드가 쓰던 방식 포함)
 * - alt(파일명)로 먼저 찾고, 못 찾으면 순서대로 연결
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
    if (!cidMatch) return tag; // 외부 URL 이미지 등은 그대로
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
    // src 가 cid 가 아니면(예: ?view=att...) cid 로 교체
    return /^cid:/i.test(src) ? tag : setAttr_(tag, 'src', 'cid:' + cid);
  });

  const rest = blobs.filter(b => used.indexOf(b) < 0);
  unmatched.forEach((cid, i) => {
    if (rest[i]) inline.push({ cid: cid, blob: rest[i] });
  });
  return { html: newHtml, inline: inline };
}

/* ======================================================================
 * 머지 필드 치환
 * ====================================================================== */

/** 헤더 → 값 조회표. 원래 이름 / 소문자 / 정규화(공백·기호 제거) 세 가지 키로 등록 */
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
 * {{필드}} 를 값으로 치환.
 * - Gmail 편집기가 {{<span>Name</span>}} 처럼 태그를 끼워 넣어도 인식
 * - HTML 본문에는 값을 이스케이프하고 줄바꿈은 <br> 로
 * - 값에 $ 가 있어도 깨지지 않도록 함수형 replace 사용
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

/** 소문자 + 영문/숫자/한글만 남김 ("Email Address" → "emailaddress") */
function normalizeKey_(s) {
  return String(s).toLowerCase().replace(/[^0-9a-z\u3131-\u318E\uAC00-\uD7A3]/g, '');
}

/* ======================================================================
 * 슬라이드 이미지 ({{SlideImage}})
 * ====================================================================== */

function renderSlideImage_(presentationId, pairs) {
  const pres = SlidesApp.openById(presentationId);
  const copy = pres.getSlides()[0].duplicate(); // 첫 장을 템플릿으로 사용
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
    // 복사한 장은 반드시 삭제 (예전 코드는 발송할 때마다 슬라이드가 쌓였음)
    const p = SlidesApp.openById(presentationId);
    const s = p.getSlideById(copyId);
    if (s) s.remove();
    p.saveAndClose();
  }
}

/* ======================================================================
 * MIME 메시지 생성 / 발송 (이모지 대응의 핵심)
 * ====================================================================== */

/**
 * RFC 2822/2045/2047 형식의 메일 원문을 만든다.
 * 모든 헤더의 비 ASCII 문자는 =?UTF-8?B?...?= 로, 본문은 UTF-8 base64 로 인코딩하므로
 * 이모지 같은 4바이트 문자도 깨지지 않는다.
 *
 * 구조: mixed( related( alternative(text, html), 인라인이미지... ), 첨부... )
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

/** 일시적 오류(요청 과다 등)는 잠시 기다렸다 재시도 */
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
 * RFC 2047 encoded-word. ASCII 면 그대로, 아니면 UTF-8 base64 로 인코딩.
 * 이모지(서로게이트 쌍)가 중간에 잘리지 않도록 '문자 단위'로 나눠 encoded-word 를 만든다.
 * @param {boolean} singleLine 파일명처럼 한 줄이어야 하는 경우 접기(folding) 없이 붙임
 */
function encodeHeaderWord_(text, singleLine) {
  text = String(text || '');
  if (/^[\x20-\x7E]*$/.test(text)) return text;
  const MAX_BYTES = 45; // base64 후 60자 → encoded-word 75자 이내
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

/** "이름 <email>" 형식. 이름에 한글/이모지가 있으면 인코딩 */
function formatAddress_(name, email) {
  name = sanitizeHeader_(name || '').trim();
  email = sanitizeHeader_(email || '').trim();
  if (!name) return email;
  if (/^[\x20-\x7E]*$/.test(name)) return `"${name.replace(/(["\\])/g, '\\$1')}" <${email}>`;
  return `${encodeHeaderWord_(name)} <${email}>`;
}

/* ======================================================================
 * 결과 창
 * ====================================================================== */

function showResultDialog_(opts, sheet, template, summary, log, isTest, tz) {
  const color = { ok: '#1a73e8', skip: '#888', error: '#d93025' };
  const label = { ok: '성공', skip: '건너뜀', error: '실패' };
  const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');

  let html = '<div style="font:13px/1.6 sans-serif">';
  html += '<div style="background:#f1f3f4;padding:10px;border-radius:6px;color:#444">';
  html += `제목: ${escapeHtml_(template.subject)}<br>`;
  html += `시트: ${escapeHtml_(sheet.getName())}<br>`;
  html += `보낸 사람: ${escapeHtml_(opts.senderName || '')} &lt;${escapeHtml_(opts.senderEmail || '')}&gt;<br>`;
  if (!isTest && opts.cc) html += `CC: ${escapeHtml_(opts.cc)}<br>`;
  if (!isTest && opts.bcc) html += `BCC: ${escapeHtml_(opts.bcc)}<br>`;
  html += `완료 시각: ${now}</div>`;

  html += `<p><b>성공 ${summary.sent}</b> · 건너뜀 ${summary.skipped} · <span style="color:${color.error}">실패 ${summary.failed}</span></p>`;
  if (summary.stoppedEarly) {
    html += `<p style="color:#e37400"><b>실행 시간 제한으로 멈췄습니다.</b> 남은 ${summary.remaining}건은 다시 RUN 하면 이어서 발송됩니다.</p>`;
  }
  if (summary.missingFields.length) {
    html += `<p style="color:#e37400">시트에 없는 필드라 빈 값으로 들어간 항목: ${summary.missingFields.map(f => '{{' + escapeHtml_(f) + '}}').join(', ')}</p>`;
  }
  html += '<div style="max-height:260px;overflow:auto;border-top:1px solid #eee;padding-top:6px">';
  log.forEach(l => {
    html += `<div><span style="color:${color[l.type]}">${label[l.type]}</span> : ${escapeHtml_(l.text)}</div>`;
  });
  html += '</div></div>';

  const out = HtmlService.createHtmlOutput(html).setWidth(520).setHeight(460);
  SpreadsheetApp.getUi().showModalDialog(out, isTest ? '테스트 메일 결과' : '메일머지 결과');
}

/* ======================================================================
 * 보내는 사람
 * ====================================================================== */

/** Gmail '다른 주소에서 메일 보내기'에 등록된 주소 목록 (기본 주소가 맨 앞) */
function listSenders_() {
  try {
    const res = Gmail.Users.Settings.SendAs.list('me');
    const list = (res.sendAs || [])
      .filter(s => s.isPrimary || s.verificationStatus === 'accepted')
      .map(s => ({ email: s.sendAsEmail, name: s.displayName || '', primary: !!s.isPrimary }));
    list.sort((a, b) => Number(b.primary) - Number(a.primary));
    if (list.length) return list;
  } catch (e) {
    // Gmail API 미설정 등 → 아래 방식으로 대체
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
 * 유틸
 * ====================================================================== */

function parseAddressList_(text, label) {
  const list = String(text || '').split(/[,;\s]+/).map(s => s.trim()).filter(Boolean);
  const bad = list.filter(a => !isValidEmail_(a));
  if (bad.length) throw new Error(`${label} 주소 형식이 올바르지 않습니다: ${bad.join(', ')}`);
  return list;
}

function isValidEmail_(s) {
  return /^[^\s@<>,;"]+@[^\s@<>,;"]+\.[^\s@<>,;"]+$/.test(String(s || '').trim());
}

/** 헤더 값에 줄바꿈이 들어가 헤더가 조작되는 것을 방지 */
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

/** 이모지가 반으로 잘리지 않도록 '문자' 단위로 자르기 */
function truncate_(s, max) {
  const chars = Array.from(String(s));
  return chars.length > max ? chars.slice(0, max).join('') + '…' : chars.join('');
}
