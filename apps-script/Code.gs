const RECORDS_TABS = { essay: '논술', practice: '실무' };
const HEADERS = ['기록 ID', '이름', '주제', '작성일', '모드', '제한 시간(초)', '실제 소요 시간(초)', '구간별 기록(JSON)', '저장 시각'];

function doPost(e) {
  try {
    if (!e || !e.postData || e.postData.contents.length > 10000) return json_({ok: false});
    const body = JSON.parse(e.postData.contents);
    const props = PropertiesService.getScriptProperties();
    const secret = props.getProperty('TIMER_SHARED_SECRET');
    if (!secret || body.secret !== secret) return json_({ok: false});
    if (typeof body.user !== 'string' || !/^[a-z0-9][a-z0-9._-]{2,31}$/.test(body.user)) return json_({ok: false});
    if (body.action === 'signup' || body.action === 'account') return accountRequest_(body, props);
    if (body.action === 'list') {
      const records = Object.keys(RECORDS_TABS)
        .flatMap(mode => listRecords_(recordsSheet_(props, mode), body.user))
        .sort((a, b) => new Date(b.savedAt) - new Date(a.savedAt));
      return json_({ok: true, records: records.slice(0, 30)});
    }
    if (body.action !== 'save' || !validRecord_(body.record)) {
      return json_({ok: false, code: 'INVALID_RECORD'});
    }

    const r = body.record;
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sheet = recordsSheet_(props, r.mode);
      const ownerHeader = sheet.getRange(1, 10);
      const header = ownerHeader.getValue();
      if (!header) ownerHeader.setValue('계정 ID');
      else if (header !== '계정 ID') throw new Error('J1 must contain the account ID header');
      const last = sheet.getLastRow();
      const duplicate = last > 1 && sheet.getRange(2, 1, last - 1, 1)
        .createTextFinder(r.id).matchEntireCell(true).findNext();
      if (duplicate && sheet.getRange(duplicate.getRow(), 10).getDisplayValue() !== body.user) {
        return json_({ok: false, code: "INVALID_RECORD"});
      }
      if (!duplicate) {
        sheet.appendRow([r.id, "'" + r.name.trim(), "'" + r.topic.trim(), "'" + r.date, r.mode,
          r.total, r.elapsed, JSON.stringify(r.stages), new Date(), "'" + body.user]);
        SpreadsheetApp.flush();
      }
      return json_({ok: true, id: r.id});
    } finally {
      lock.releaseLock();
    }
  } catch {
    return json_({ok: false});
  }
}

function recordsSheet_(props, mode) {
  const sheet = SpreadsheetApp.openById(props.getProperty('SPREADSHEET_ID')).getSheetByName(RECORDS_TABS[mode]);
  if (!sheet || sheet.getRange(1, 1, 1, HEADERS.length).getDisplayValues()[0].join('|') !== HEADERS.join('|')) {
    throw new Error('논술·실무 탭과 첫 행의 열 제목을 확인하세요.');
  }
  return sheet;
}

function listRecords_(sheet, user) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  if (sheet.getRange(1, 10).getValue() !== '계정 ID') return [];
  const rows = sheet.getRange(2, 1, last - 1, 10).getValues();
  return rows.reverse().filter(row => row[9] === user).slice(0, 30).map(row => ({
    id: row[0], name: row[1], topic: row[2], date: row[3], mode: row[4], total: row[5],
    elapsed: row[6], stages: JSON.parse(row[7] || '[]'), savedAt: row[8]
  }));
}

function validRecord_(r) {
  const text = (s, max) => typeof s === 'string' && s.trim().length > 0 && s.length <= max;
  const seconds = (n, max) => Number.isInteger(n) && n >= 0 && n <= max;
  return !!r && /^[a-f0-9]{8}(?:-[a-f0-9]{4}){3}-[a-f0-9]{12}$/.test(r.id) &&
    text(r.name, 60) && text(r.topic, 200) && typeof r.date === 'string' &&
    /^\d{4}-\d{2}-\d{2}$/.test(r.date) && Number.isFinite(Date.parse(r.date)) &&
    new Date(r.date).toISOString().slice(0, 10) === r.date &&
    ['essay', 'practice'].includes(r.mode) && seconds(r.total, 7200) && r.total >= 600 &&
    seconds(r.elapsed, 86400) && r.elapsed > 0 && Array.isArray(r.stages) &&
    r.stages.length === 6 && r.stages.every(s => s && text(s.name, 60) &&
      seconds(s.planned, 7200) && (s.actual === null || seconds(s.actual, 86400)));
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}


function accountRequest_(body, props) {
  if (!/^[a-f0-9]{64}$/.test(body.client)) return json_({ok: false});
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const cache = CacheService.getScriptCache();
    const keys = ['auth-ip:' + body.client, 'auth-user:' + body.user];
    const counts = keys.map(key => Number(cache.get(key) || 0));
    if (counts[0] >= 50 || counts[1] >= 12) return json_({ok: false, code: 'RATE_LIMIT'});
    keys.forEach((key, i) => cache.put(key, String(counts[i] + 1), 300));

    const book = SpreadsheetApp.openById(props.getProperty('SPREADSHEET_ID'));
    let sheet = book.getSheetByName('타이머 계정');
    if (!sheet) {
      if (body.action === 'account') return json_({ok: true, account: null});
      sheet = book.insertSheet('타이머 계정');
      sheet.appendRow(['계정 ID', '비밀번호 솔트', '비밀번호 해시']);
    }
    const count = sheet.getLastRow() - 1;
    const cell = count > 0 ? sheet.getRange(2, 1, count, 1)
      .createTextFinder(body.user).matchEntireCell(true).findNext() : null;
    if (body.action === 'account') {
      const row = cell ? sheet.getRange(cell.getRow(), 2, 1, 2).getDisplayValues()[0] : null;
      return json_({ok: true, account: row ? {salt: row[0], hash: row[1]} : null});
    }
    if (!/^[a-f0-9]{32}$/.test(body.salt) || !/^[a-f0-9]{128}$/.test(body.hash)) return json_({ok: false});
    if (cell) return json_({ok: false, code: 'USER_EXISTS'});
    sheet.appendRow(["'" + body.user, "'" + body.salt, "'" + body.hash]);
    SpreadsheetApp.flush();
    return json_({ok: true});
  } finally {
    lock.releaseLock();
  }
}
