const RECORDS_TABS = { essay: '논술', practice: '실무' };
const HEADERS = ['기록 ID', '이름', '주제', '작성일', '모드', '제한 시간(초)', '실제 소요 시간(초)', '구간별 기록(JSON)', '저장 시각'];

function doPost(e) {
  try {
    if (!e || !e.postData || e.postData.contents.length > 10000) return json_({ok: false});
    const body = JSON.parse(e.postData.contents);
    const props = PropertiesService.getScriptProperties();
    const secret = props.getProperty('TIMER_SHARED_SECRET');
    if (!secret || body.secret !== secret) return json_({ok: false});
    if (body.action === 'list') {
      const records = Object.keys(RECORDS_TABS)
        .flatMap(mode => listRecords_(recordsSheet_(props, mode)))
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
      const last = sheet.getLastRow();
      const duplicate = last > 1 && sheet.getRange(2, 1, last - 1, 1)
        .createTextFinder(r.id).matchEntireCell(true).findNext();
      if (!duplicate) {
        sheet.appendRow([r.id, "'" + r.name.trim(), "'" + r.topic.trim(), "'" + r.date, r.mode,
          r.total, r.elapsed, JSON.stringify(r.stages), new Date()]);
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

function listRecords_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const first = Math.max(2, last - 29);
  return sheet.getRange(first, 1, last - first + 1, HEADERS.length).getValues().reverse().map(row => ({
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
