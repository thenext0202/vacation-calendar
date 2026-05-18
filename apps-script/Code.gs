/**
 * 여름휴가 일정 공유 - Google Apps Script 백엔드
 *
 * 시트 구조 (자동 생성):
 *   A: company | B: name | C: day1 | D: day2 | E: created_at
 *
 * API:
 *   GET  /exec                                  → 전체 레코드 조회
 *   POST /exec {action:'add',    company, name, days:[d1,d2]}
 *   POST /exec {action:'delete', company, name}
 */

const SHEET_NAME = 'records';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['company', 'name', 'day1', 'day2', 'created_at']);
    sheet.getRange('A1:E1').setFontWeight('bold').setBackground('#e0f2fe');
    sheet.setFrozenRows(1);
  }
  // 날짜 컬럼(C,D)은 텍스트 포맷 — Sheets가 "2026-08-11"을 자동으로 Date로 변환하지 않도록 차단
  sheet.getRange('C:D').setNumberFormat('@');
  return sheet;
}

// Date 객체든 문자열이든 'YYYY-MM-DD'로 정규화
function formatDay_(v) {
  if (v instanceof Date) {
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, '0');
    const d = String(v.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return String(v);
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  const records = [];
  for (let i = 1; i < data.length; i++) {
    const [company, name, day1, day2] = data[i];
    if (!company || !name) continue;
    records.push({
      company: String(company),
      name: String(name),
      days: [formatDay_(day1), formatDay_(day2)]
    });
  }
  return jsonOut_(records);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const body = JSON.parse(e.postData.contents);
    const sheet = getSheet_();

    if (body.action === 'add') {
      if (!body.company || !body.name || !body.days || body.days.length !== 2) {
        return jsonOut_({ ok: false, error: 'invalid_payload' });
      }
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === body.company && data[i][1] === body.name) {
          return jsonOut_({ ok: false, error: 'duplicate' });
        }
      }
      sheet.appendRow([
        body.company,
        body.name,
        body.days[0],
        body.days[1],
        new Date()
      ]);
      return jsonOut_({ ok: true });
    }

    if (body.action === 'delete') {
      if (!body.company || !body.name) {
        return jsonOut_({ ok: false, error: 'invalid_payload' });
      }
      const data = sheet.getDataRange().getValues();
      let deleted = 0;
      for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][0] === body.company && data[i][1] === body.name) {
          sheet.deleteRow(i + 1);
          deleted++;
        }
      }
      return jsonOut_({ ok: true, deleted: deleted });
    }

    return jsonOut_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
