/**
 * Add-on endpoints for the Nil Transaction BO + BO Booking dashboard.
 *
 * IMPORTANT: This is an ADDITIONAL .gs file for the existing Apps Script project.
 * Follow APPS-SCRIPT-UPGRADE.md to route the existing doGet/doPost functions to
 * these helpers, then redeploy the web app as a new deployment version.
 *
 * Shared exclusions are stored in a dedicated sheet named:
 *   Dashboard Office Exclusions
 * Columns: office_id, office_name, excluded, updated_at
 */

const EXCLUSIONS_SHEET = 'Dashboard Office Exclusions';
const EXCLUSIONS_HEADERS = ['office_id','office_name','excluded','updated_at'];

/** GET action helper: return Daily Nil Transactions on/after a date (inclusive). */
function getDailySince_(sinceValue) {
  const since = normalizeSheetDate_(sinceValue);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(since)) {
    throw new Error('getDailySince requires since=YYYY-MM-DD.');
  }
  const sheet = getRequiredSheet_(DAILY_SHEET);
  const lastRow = sheet.getLastRow();
  const totalCount = Math.max(0, lastRow - 1);
  if (!totalCount) return {ok:true, action:'getDailySince', since:since, rows:[], count:0, totalCount:0, timestamp:new Date().toISOString()};
  // Read the date column first, then only return rows that match the date window.
  const dateValues = sheet.getRange(2, 3, totalCount, 1).getDisplayValues();
  const rowNumbers = [];
  dateValues.forEach(function(row, i) {
    const d = normalizeSheetDate_(row[0]);
    if (d && d >= since) rowNumbers.push(i + 2);
  });
  const rows = [];
  // Contiguous reads are combined into blocks to avoid one API call per record.
  const blocks = [];
  rowNumbers.forEach(function(rowNum) {
    const last = blocks[blocks.length - 1];
    if (last && rowNum === last.start + last.count) last.count++;
    else blocks.push({start:rowNum,count:1});
  });
  blocks.forEach(function(block) {
    const values = sheet.getRange(block.start, 1, block.count, DAILY_HEADERS.length).getDisplayValues();
    values.forEach(function(row) {
      const obj = {};
      DAILY_HEADERS.forEach(function(header, i) { obj[header] = row[i] === undefined ? '' : row[i]; });
      rows.push(obj);
    });
  });
  return {ok:true, action:'getDailySince', since:since, rows:rows, count:rows.length, totalCount:totalCount, timestamp:new Date().toISOString()};
}

/** GET action helper: return booking rows whose Booking Date is on/after since. */
function getBookingSince_(sinceValue) {
  const since = normalizeSheetDate_(sinceValue);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(since)) {
    throw new Error('getBookingSince requires since=YYYY-MM-DD.');
  }
  const sheet = getRequiredSheet_(BOOKING_SHEET);
  const lastRow = sheet.getLastRow();
  const totalCount = Math.max(0, lastRow - 1);
  if (!totalCount) return {ok:true, action:'getBookingSince', since:since, rows:[], count:0, totalCount:0, timestamp:new Date().toISOString()};
  const dates = sheet.getRange(2, 11, totalCount, 1).getDisplayValues();
  const rowNumbers = [];
  dates.forEach(function(row, i) {
    const d = normalizeSheetDate_(row[0]);
    if (d && d >= since) rowNumbers.push(i + 2);
  });
  const rows = [];
  const blocks = [];
  rowNumbers.forEach(function(rowNum) {
    const last = blocks[blocks.length - 1];
    if (last && rowNum === last.start + last.count) last.count++;
    else blocks.push({start:rowNum,count:1});
  });
  blocks.forEach(function(block) {
    const values = sheet.getRange(block.start, 1, block.count, BOOKING_HEADERS.length).getDisplayValues();
    values.forEach(function(row) {
      const obj = {};
      BOOKING_KEYS.forEach(function(key, i) { obj[key] = row[i] === undefined ? '' : row[i]; });
      rows.push(obj);
    });
  });
  return {ok:true, action:'getBookingSince', since:since, rows:rows, count:rows.length, totalCount:totalCount, timestamp:new Date().toISOString()};
}

function ensureExclusionsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(EXCLUSIONS_SHEET);
  if (!sheet) sheet = ss.insertSheet(EXCLUSIONS_SHEET);
  if (sheet.getLastRow() === 0) sheet.getRange(1,1,1,EXCLUSIONS_HEADERS.length).setValues([EXCLUSIONS_HEADERS]);
  const headers = sheet.getRange(1,1,1,EXCLUSIONS_HEADERS.length).getDisplayValues()[0];
  const valid = EXCLUSIONS_HEADERS.every(function(h,i){return String(headers[i]||'').trim()===h;});
  if (!valid) throw new Error('Header mismatch in "'+EXCLUSIONS_SHEET+'". Expected: '+EXCLUSIONS_HEADERS.join(', '));
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,EXCLUSIONS_HEADERS.length).setFontWeight('bold').setBackground('#dbeafe').setWrap(true);
  return sheet;
}

/** GET action helper: shared exclusion IDs, loaded by every dashboard browser. */
function getExclusions_() {
  const sheet = ensureExclusionsSheet_();
  const lastRow = sheet.getLastRow();
  const excludedOfficeIds = [];
  if (lastRow > 1) {
    sheet.getRange(2,1,lastRow-1,3).getDisplayValues().forEach(function(row) {
      const id = String(row[0]||'').trim();
      const excluded = String(row[2]||'').trim().toLowerCase();
      if (id && (excluded === 'true' || excluded === 'yes' || excluded === '1')) excludedOfficeIds.push(id);
    });
  }
  return {ok:true, action:'getExclusions', excludedOfficeIds:excludedOfficeIds, count:excludedOfficeIds.length, timestamp:new Date().toISOString()};
}

/** POST action helper: atomically replace the shared exclusion list. */
function saveExclusions_(ids) {
  if (!Array.isArray(ids)) throw new Error('saveExclusions requires excludedOfficeIds as an array.');
  const unique = Array.from(new Set(ids.map(function(x){return String(x||'').trim();}).filter(Boolean)));
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = ensureExclusionsSheet_();
    const masterById = new Map();
    readRowsAsObjects_(MASTER_SHEET, MASTER_HEADERS).forEach(function(o) {
      const id = String(o.office_id||'').trim();
      if (id) masterById.set(id, String(o.office_name||''));
    });
    const now = new Date();
    const values = unique.map(function(id) {
      return [id, masterById.get(id)||'', true, now];
    });
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.getRange(2,1,lastRow-1,EXCLUSIONS_HEADERS.length).clearContent();
    if (values.length) {
      sheet.getRange(2,1,values.length,EXCLUSIONS_HEADERS.length).setValues(values);
      sheet.getRange(2,4,values.length,1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    }
    SpreadsheetApp.flush();
    return {sheet:EXCLUSIONS_SHEET, excludedCount:values.length, timestamp:now.toISOString()};
  } finally {
    lock.releaseLock();
  }
}
