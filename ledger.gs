// ledger.gs -- normalized long-table layer (the migration target).
//
// Four flat sheets replace the per-month positional sheets:
//   Txns             id | date | month | description | amount | method | category | discretionary | ben | jenna
//   Fixed Log        id | month | name | source | amount | paid | paid_date
//   Income Log       id | month | date | source | amount
//   Variable History month | category | amount        (display/TTM only -- NEVER summed into expenses)
//
// Headers on row 1, data rows 2+. Rows are located by stable id (col A),
// never by row index -- the stale-rowIndex bug class cannot exist here.
// Appends only; no formulas below the data block; no positional ranges to rot.
//
// THE AGGREGATION RULE (enforced here and only here):
//   Every dollar counts exactly once.
//   - A transaction counts via its row. category (onetime/groceries/gas) is a
//     label, not a separate bucket to re-add.
//   - transfers = rows with category 'transfer' (deposits in, moves between
//     accounts, pass-through money) PLUS any Checking row with amount < 0
//     (an overlooked transfer: money arriving in checking is never a
//     negative expense). Excluded from expenses AND from income. Negative
//     amounts on cards (cash back, refunds) stay in and net against spend.
//   - A fixed expense counts via its paid flag, including Extra loan payments
//     and utilities (Electricity, Water live ONLY in Fixed Log).
//   - totalExpenses = (txn sum - transfers) + fixedPaid. Nothing else.
//   - Variable History is estimation/display data. It never feeds a total.
//   This rule exists because the old sheets double-counted Electricity+Gas
//   (in both the fixed SUMIF and the actual-variable sum) every month since
//   Dec 2024 ($5,541.88) and Dogs/Vet in 5 months ($1,937.80): the dogs
//   "actual" was a formula re-summing cells already counted as transactions.
//   Transfers were previously a hand-picked SUM of cell refs, rebuilt by hand
//   each month -- the pick was forgotten at least once (a deposit counted as
//   negative expense). Here it is a category, assigned once at entry.

var LEDGER_TXNS    = 'Txns';
var LEDGER_FIXED   = 'Fixed Log';
var LEDGER_INCOME  = 'Income Log';
var LEDGER_VARHIST = 'Variable History';

var LEDGER_CACHE_VERSION = 'ledger_v1'; // bump on every getLedgerMonthly shape change
var LEDGER_CACHE_KEY     = LEDGER_CACHE_VERSION + '_monthly';

var LEDGER_CATEGORIES = ['onetime', 'groceries', 'gas', 'transfer'];

var LEDGER_HEADERS = {};
LEDGER_HEADERS[LEDGER_TXNS]    = ['id','date','month','description','amount','method','category','discretionary','ben','jenna'];
LEDGER_HEADERS[LEDGER_FIXED]   = ['id','month','name','source','amount','paid','paid_date'];
LEDGER_HEADERS[LEDGER_INCOME]  = ['id','month','date','source','amount'];
LEDGER_HEADERS[LEDGER_VARHIST] = ['month','category','amount'];

// MANUAL SHEET SETUP (no script function creates sheets -- scripted
// insertSheet on this workbook triggers full recalculation and times out).
// Create four tabs by hand, paste the header row into row 1 of each,
// View > Freeze > 1 row, then import the migration CSVs at A1 (the CSVs
// include the header row; "Replace data at selected cell"):
//   Txns:             id  date  month  description  amount  method  category  discretionary  ben  jenna
//   Fixed Log:        id  month  name  source  amount  paid  paid_date
//   Income Log:       id  month  date  source  amount
//   Variable History: month  category  amount
// Optionally create an empty hidden tab '_TestScratchLedger_' so testLedger
// never has to insertSheet either.

// -- tolerant readers (imported CSV cells may be text where booleans/dates
//    are expected; Sheets may hand back Date objects or strings) ----------
function _lb(v) { // boolean
  if (v === true) return true;
  return String(v).toUpperCase() === 'TRUE';
}
function _ln(v) { // number
  if (typeof v === 'number' && isFinite(v)) return v;
  var n = parseFloat(String(v).replace(/[$,]/g, ''));
  return isFinite(n) ? n : null;
}
function _ldate(v) { // -> 'YYYY-MM-DD' or ''
  if (v instanceof Date && !isNaN(v.getTime())) {
    return v.getFullYear() + '-' +
      ('0' + (v.getMonth() + 1)).slice(-2) + '-' +
      ('0' + v.getDate()).slice(-2);
  }
  var s = String(v || '').trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}
function _lmonth(v) { // -> 'YYYY-MM' or ''
  if (v instanceof Date && !isNaN(v.getTime())) {
    return v.getFullYear() + '-' + ('0' + (v.getMonth() + 1)).slice(-2);
  }
  var s = String(v || '').trim();
  if (/^\d{4}-\d{2}$/.test(s)) return s;
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s.slice(0, 7);
  return '';
}

// -- Pure aggregation over raw values() arrays (header row included).
//    Kept pure so testLedger can drive it with fixtures and so an offline
//    run against the migration CSVs can predict its output exactly. --------
function _ledgerAggregate(txnRows, fixedRows, incomeRows) {
  var per = {};
  function m(mk) {
    if (!per[mk]) per[mk] = { onetime: 0, groceries: 0, gas: 0, transfers: 0,
      discretionary: 0, benDisc: 0, jennaDisc: 0, fixedPaid: 0, income: 0,
      spend: 0, totalExpenses: 0, net: 0, txnCount: 0 };
    return per[mk];
  }
  var i, r;
  for (i = 1; i < txnRows.length; i++) {
    r = txnRows[i];
    var mk = _lmonth(r[2]) || _lmonth(r[1]);
    var amt = _ln(r[4]);
    if (!mk || amt === null) continue;
    var b = m(mk);
    b.txnCount++;
    var cat = String(r[6] || 'onetime');
    var isTransfer = (cat === 'transfer') || (String(r[5]) === 'Checking' && amt < 0);
    if (isTransfer) { b.transfers += amt; continue; } // money movement: not expense, not income
    if (cat === 'groceries')  b.groceries += amt;
    else if (cat === 'gas')   b.gas += amt;
    else                      b.onetime += amt;
    if (_lb(r[7])) { // discretionary
      b.discretionary += amt;
      var ben = _lb(r[8]), jenna = _lb(r[9]);
      var nSplit = (ben ? 1 : 0) + (jenna ? 1 : 0);
      if (nSplit > 0) {
        if (ben)   b.benDisc   += amt / nSplit;
        if (jenna) b.jennaDisc += amt / nSplit;
      }
    }
  }
  for (i = 1; i < fixedRows.length; i++) {
    r = fixedRows[i];
    var fmk = _lmonth(r[1]);
    var famt = _ln(r[4]);
    if (!fmk || famt === null) continue;
    if (_lb(r[5])) m(fmk).fixedPaid += famt;
  }
  for (i = 1; i < incomeRows.length; i++) {
    r = incomeRows[i];
    var imk = _lmonth(r[1]) || _lmonth(r[2]);
    var iamt = _ln(r[4]);
    if (!imk || iamt === null) continue;
    m(imk).income += iamt;
  }
  var months = [], k;
  for (k in per) {
    var b2 = per[k];
    b2.spend = b2.onetime + b2.groceries + b2.gas;
    b2.totalExpenses = b2.spend + b2.fixedPaid;
    b2.net = b2.income - b2.totalExpenses;
    months.push(k);
  }
  months.sort();
  months.reverse(); // newest first, matching getMonthlySheetNames convention
  return { months: months, perMonth: per };
}

// -- READ: ledgerMonthly (the getMonthlyData successor) --------------------
// One pass over three flat tables. No per-month cache tiers needed: the
// whole history is a few thousand rows read in one getDataRange each.
// A short CacheService TTL still smooths repeat loads.
function getLedgerMonthly() {
  var cache = CacheService.getScriptCache();
  try {
    var hit = cache.get(LEDGER_CACHE_KEY);
    if (hit) return JSON.parse(hit);
  } catch (e) {}
  var txnRows    = readSheet(LEDGER_TXNS);
  var fixedRows  = readSheet(LEDGER_FIXED);
  var incomeRows = readSheet(LEDGER_INCOME);
  if (!txnRows) return { error: 'Sheet not found: ' + LEDGER_TXNS + ' -- create the ledger sheets (see manual setup notes at top of ledger.gs) and import the migration CSVs.' };
  var out = _ledgerAggregate(txnRows, fixedRows || [[]], incomeRows || [[]]);
  out.updated = new Date().toISOString();
  try { cache.put(LEDGER_CACHE_KEY, JSON.stringify(out), 300); } catch (e2) {}
  return out;
}

function invalidateLedgerCache() {
  try { CacheService.getScriptCache().remove(LEDGER_CACHE_KEY); } catch (e) {}
}

// -- READ: ledgerTxns -- one month's transactions, by stable id -----------
function getLedgerTxns(month) {
  var mk = _lmonth(month);
  if (!mk) return { error: 'ledgerTxns needs month=YYYY-MM' };
  var rows = readSheet(LEDGER_TXNS);
  if (!rows) return { error: 'Sheet not found: ' + LEDGER_TXNS };
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (_lmonth(r[2]) !== mk) continue;
    out.push({
      id: parseInt(r[0], 10), date: _ldate(r[1]), month: mk,
      description: String(r[3] || ''), amount: _ln(r[4]),
      method: String(r[5] || ''), category: String(r[6] || 'onetime'),
      discretionary: _lb(r[7]), ben: _lb(r[8]), jenna: _lb(r[9])
    });
  }
  out.sort(function(a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : a.id - b.id; });
  return { month: mk, txns: out };
}

// -- shared write plumbing --------------------------------------------------
function _ledgerToday() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function _ledgerWithLock(fn) {
  var lock = LockService.getScriptLock();
  var got = false;
  try { got = lock.tryLock(5000); } catch (e) {}
  if (!got) return { error: 'Could not acquire lock -- please try again.' };
  try { return fn(); }
  finally { try { lock.releaseLock(); } catch (e2) {} }
}
function _ledgerNextId(values) {
  var maxId = 0;
  for (var i = 1; i < values.length; i++) {
    var n = parseInt(values[i][0], 10);
    if (!isNaN(n) && n > maxId) maxId = n;
  }
  return maxId + 1;
}
function _ledgerFindRowById(values, id) { // -> sheet row number, or -1
  var want = parseInt(id, 10);
  for (var i = 1; i < values.length; i++) {
    if (parseInt(values[i][0], 10) === want) return i + 1;
  }
  return -1;
}

// Validate + normalize a txn payload. Returns { error } or { rec } where
// rec is the 9 data columns (everything after id). Single place -- add,
// update, and split all route through here.
function _ledgerValidateTxn(p) {
  var dateStr = _ldate(p.date);
  if (!dateStr) return { error: 'Invalid or missing date (need YYYY-MM-DD).' };
  var amount = _ln(p.amount);
  if (amount === null || amount === 0) return { error: 'Amount must be a nonzero number.' };
  var category = String(p.category || 'onetime');
  if (LEDGER_CATEGORIES.indexOf(category) < 0)
    return { error: 'Unknown category: ' + category + ' (want onetime|groceries|gas|transfer)' };
  var method = String(p.method || '').trim();
  if (!method) return { error: 'Payment method is required.' };
  var description = String(p.description || '').trim();
  if (!description) {
    if (category === 'gas') description = 'Gas';
    else return { error: 'Description is required.' };
  }
  // month is an explicit assignment, not always month(date): a July-1 charge
  // can belong to June's books, exactly as the old sheets allowed.
  var mk = _lmonth(p.month) || _lmonth(dateStr);
  return { rec: [dateStr, mk, description, amount, method, category,
                 !!p.discretionary, !!p.ben, !!p.jenna] };
}

// -- WRITE: ledgerAddTxn ----------------------------------------------------
// Validated append. id = max existing id + 1 under LockService (two
// simultaneous adds must never share an id; edit/delete locate rows by id).
// Wrapped in _withIdem by both routers; invalidates caches internally.
function ledgerAddTxn(p) {
  var v = _ledgerValidateTxn(p);
  if (v.error) return v;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LEDGER_TXNS);
  if (!sheet) return { error: 'Sheet not found: ' + LEDGER_TXNS };
  return _ledgerWithLock(function () {
    invalidateSheet(LEDGER_TXNS);
    var values = readSheet(sheet);
    var newId = _ledgerNextId(values);
    var row = Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(row, 1, 1, 10).setValues([[newId].concat(v.rec)]);
    SpreadsheetApp.flush();
    invalidateLedgerCache();
    invalidateSheet(LEDGER_TXNS);
    return { ok: true, id: newId, month: v.rec[1] };
  });
}

// -- WRITE: ledgerUpdateTxn -------------------------------------------------
// Full-record replace by id: client sends the complete corrected txn (same
// fields as add, plus id). Partial updates are a foot-gun on money rows --
// the record you validated is the record you store.
function ledgerUpdateTxn(p) {
  if (!p.id) return { error: 'ledgerUpdateTxn needs id.' };
  var v = _ledgerValidateTxn(p);
  if (v.error) return v;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LEDGER_TXNS);
  if (!sheet) return { error: 'Sheet not found: ' + LEDGER_TXNS };
  return _ledgerWithLock(function () {
    invalidateSheet(LEDGER_TXNS);
    var values = readSheet(sheet);
    var row = _ledgerFindRowById(values, p.id);
    if (row < 0) return { error: 'Txn id ' + p.id + ' not found.' };
    sheet.getRange(row, 2, 1, 9).setValues([v.rec]);
    SpreadsheetApp.flush();
    invalidateLedgerCache();
    invalidateSheet(LEDGER_TXNS);
    return { ok: true, id: parseInt(p.id, 10), month: v.rec[1] };
  });
}

// -- WRITE: ledgerDeleteTxn -------------------------------------------------
// deleteRow is safe on a flat table (no formulas below the data block), and
// row shifts are irrelevant because everything locates by id. A repeat
// delete of the same id fails loudly ('not found') -- genuine network
// retries are already absorbed by _withIdem before reaching here.
function ledgerDeleteTxn(p) {
  if (!p.id) return { error: 'ledgerDeleteTxn needs id.' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LEDGER_TXNS);
  if (!sheet) return { error: 'Sheet not found: ' + LEDGER_TXNS };
  return _ledgerWithLock(function () {
    invalidateSheet(LEDGER_TXNS);
    var values = readSheet(sheet);
    var row = _ledgerFindRowById(values, p.id);
    if (row < 0) return { error: 'Txn id ' + p.id + ' not found.' };
    sheet.deleteRow(row);
    SpreadsheetApp.flush();
    invalidateLedgerCache();
    invalidateSheet(LEDGER_TXNS);
    return { ok: true, id: parseInt(p.id, 10) };
  });
}

// -- WRITE: ledgerSplitTxn --------------------------------------------------
// Atomic split under one lock: the original row is rewritten as parts[0]
// (keeping its id), the remaining parts append with new ids. INVARIANT:
// the parts must sum to the original amount (a split conserves money);
// anything else is an update or an add, and is rejected loudly.
function _ledgerPrepSplit(orig, parts) { // pure; orig = {date,month,method,...}
  if (!parts || !parts.length || parts.length < 2)
    return { error: 'Split needs at least 2 parts.' };
  var recs = [], sum = 0;
  for (var i = 0; i < parts.length; i++) {
    var part = parts[i];
    var merged = {
      date: part.date !== undefined ? part.date : orig.date,
      month: part.month !== undefined ? part.month : orig.month,
      description: part.description,
      amount: part.amount,
      method: part.method !== undefined ? part.method : orig.method,
      category: part.category !== undefined ? part.category : orig.category,
      discretionary: part.discretionary !== undefined ? part.discretionary : orig.discretionary,
      ben: part.ben !== undefined ? part.ben : orig.ben,
      jenna: part.jenna !== undefined ? part.jenna : orig.jenna
    };
    var v = _ledgerValidateTxn(merged);
    if (v.error) return { error: 'Part ' + (i + 1) + ': ' + v.error };
    recs.push(v.rec);
    sum += v.rec[3];
  }
  if (Math.abs(sum - orig.amount) > 0.005)
    return { error: 'Split parts sum to ' + sum.toFixed(2) + ' but original is ' +
                    orig.amount.toFixed(2) + ' -- a split must conserve the amount.' };
  return { recs: recs };
}

function ledgerSplitTxn(p) {
  if (!p.id) return { error: 'ledgerSplitTxn needs id.' };
  var parts = p.parts;
  if (typeof parts === 'string') { try { parts = JSON.parse(parts); } catch (e) { return { error: 'parts is not valid JSON.' }; } }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LEDGER_TXNS);
  if (!sheet) return { error: 'Sheet not found: ' + LEDGER_TXNS };
  return _ledgerWithLock(function () {
    invalidateSheet(LEDGER_TXNS);
    var values = readSheet(sheet);
    var row = _ledgerFindRowById(values, p.id);
    if (row < 0) return { error: 'Txn id ' + p.id + ' not found.' };
    var r = values[row - 1];
    var orig = { date: _ldate(r[1]), month: _lmonth(r[2]), description: String(r[3] || ''),
                 amount: _ln(r[4]), method: String(r[5] || ''), category: String(r[6] || 'onetime'),
                 discretionary: _lb(r[7]), ben: _lb(r[8]), jenna: _lb(r[9]) };
    var prep = _ledgerPrepSplit(orig, parts);
    if (prep.error) return prep;
    var ids = [parseInt(p.id, 10)];
    sheet.getRange(row, 2, 1, 9).setValues([prep.recs[0]]);
    var nextId = _ledgerNextId(values);
    var appendAt = Math.max(sheet.getLastRow() + 1, 2);
    var newRows = [];
    for (var i = 1; i < prep.recs.length; i++) {
      newRows.push([nextId].concat(prep.recs[i]));
      ids.push(nextId);
      nextId++;
    }
    sheet.getRange(appendAt, 1, newRows.length, 10).setValues(newRows);
    SpreadsheetApp.flush();
    invalidateLedgerCache();
    invalidateSheet(LEDGER_TXNS);
    return { ok: true, ids: ids };
  });
}

// -- READ: ledgerFixed -- one month's fixed rows, by stable id --------------
function getLedgerFixed(month) {
  var mk = _lmonth(month);
  if (!mk) return { error: 'ledgerFixed needs month=YYYY-MM' };
  var rows = readSheet(LEDGER_FIXED);
  if (!rows) return { error: 'Sheet not found: ' + LEDGER_FIXED };
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (_lmonth(r[1]) !== mk) continue;
    out.push({ id: parseInt(r[0], 10), month: mk, name: String(r[2] || ''),
               source: String(r[3] || ''), amount: _ln(r[4]), paid: _lb(r[5]),
               paid_date: _ldate(r[6]) });
  }
  out.sort(function (a, b) { return a.id - b.id; });
  return { month: mk, fixed: out };
}

// -- WRITE: ledgerAddFixed --------------------------------------------------
function ledgerAddFixed(p) {
  var mk = _lmonth(p.month);
  if (!mk) return { error: 'ledgerAddFixed needs month=YYYY-MM.' };
  var name = String(p.name || '').trim();
  if (!name) return { error: 'Fixed expense name is required.' };
  var amount = _ln(p.amount);
  if (amount === null) return { error: 'Amount must be a number.' };
  var paid = !!p.paid && String(p.paid).toUpperCase() !== 'FALSE';
  var paidDate = paid ? (_ldate(p.paid_date) || _ledgerToday()) : '';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LEDGER_FIXED);
  if (!sheet) return { error: 'Sheet not found: ' + LEDGER_FIXED };
  return _ledgerWithLock(function () {
    invalidateSheet(LEDGER_FIXED);
    var values = readSheet(sheet);
    var newId = _ledgerNextId(values);
    var row = Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(row, 1, 1, 7).setValues([[newId, mk, name,
      String(p.source || '').trim(), amount, paid, paidDate]]);
    SpreadsheetApp.flush();
    invalidateLedgerCache();
    invalidateSheet(LEDGER_FIXED);
    return { ok: true, id: newId, month: mk };
  });
}

// -- WRITE: ledgerUpdateFixed -----------------------------------------------
// Partial update by id: only supplied keys change (dashboard ops are
// single-field: toggle paid, edit cost, rename). paid=true without a date
// stamps today; paid=false clears the date -- a paid_date on an unpaid row
// is a contradiction we refuse to store.
function ledgerUpdateFixed(p) {
  if (!p.id) return { error: 'ledgerUpdateFixed needs id.' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LEDGER_FIXED);
  if (!sheet) return { error: 'Sheet not found: ' + LEDGER_FIXED };
  return _ledgerWithLock(function () {
    invalidateSheet(LEDGER_FIXED);
    var values = readSheet(sheet);
    var row = _ledgerFindRowById(values, p.id);
    if (row < 0) return { error: 'Fixed id ' + p.id + ' not found.' };
    var r = values[row - 1].slice();
    if (p.name !== undefined) {
      var nm = String(p.name).trim();
      if (!nm) return { error: 'Name cannot be blank.' };
      r[2] = nm;
    }
    if (p.source !== undefined) r[3] = String(p.source).trim();
    if (p.amount !== undefined) {
      var amt = _ln(p.amount);
      if (amt === null) return { error: 'Amount must be a number.' };
      r[4] = amt;
    }
    if (p.paid !== undefined) {
      var paid = (p.paid === true) || String(p.paid).toUpperCase() === 'TRUE';
      r[5] = paid;
      r[6] = paid ? (_ldate(p.paid_date) || _ldate(r[6]) || _ledgerToday()) : '';
    } else if (p.paid_date !== undefined) {
      if (!_lb(r[5])) return { error: 'Cannot set paid_date on an unpaid row.' };
      r[6] = _ldate(p.paid_date) || _ledgerToday();
    }
    sheet.getRange(row, 2, 1, 6).setValues([[r[1], r[2], r[3], r[4], r[5], r[6]]]);
    SpreadsheetApp.flush();
    invalidateLedgerCache();
    invalidateSheet(LEDGER_FIXED);
    return { ok: true, id: parseInt(p.id, 10) };
  });
}

// -- WRITE: ledgerSeedFixedMonth --------------------------------------------
// The surviving fragment of addMonth: copy the most recent prior month's
// fixed rows into the target month, all unpaid. Refuses if the target
// already has rows (re-seeding would duplicate; delete the rows in the
// sheet first if a re-seed is genuinely wanted).
function ledgerSeedFixedMonth(p) {
  var mk = _lmonth(p.month);
  if (!mk) return { error: 'ledgerSeedFixedMonth needs month=YYYY-MM.' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LEDGER_FIXED);
  if (!sheet) return { error: 'Sheet not found: ' + LEDGER_FIXED };
  return _ledgerWithLock(function () {
    invalidateSheet(LEDGER_FIXED);
    var values = readSheet(sheet);
    var srcMonth = '', i;
    for (i = 1; i < values.length; i++) {
      var m = _lmonth(values[i][1]);
      if (!m) continue;
      if (m === mk) return { error: mk + ' already has fixed rows -- not seeding twice.' };
      if (m < mk && m > srcMonth) srcMonth = m;
    }
    if (!srcMonth) return { error: 'No prior month found to seed from.' };
    var newRows = [], nextId = _ledgerNextId(values);
    for (i = 1; i < values.length; i++) {
      if (_lmonth(values[i][1]) !== srcMonth) continue;
      newRows.push([nextId, mk, String(values[i][2] || ''), String(values[i][3] || ''),
                    _ln(values[i][4]) || 0, false, '']);
      nextId++;
    }
    var row = Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(row, 1, newRows.length, 7).setValues(newRows);
    SpreadsheetApp.flush();
    invalidateLedgerCache();
    invalidateSheet(LEDGER_FIXED);
    return { ok: true, month: mk, seededFrom: srcMonth, count: newRows.length };
  });
}

// -- Tests (compact, failure-first; pure fns on fixtures, write path on a
//    hidden scratch sheet -- never the live Txns sheet) ---------------------
function testLedger() {
  var fails = [], passes = 0;
  function ok(cond, msg) { if (cond) passes++; else fails.push('[ledger] ' + msg); }
  function near(a, b) { return Math.abs(a - b) < 0.005; }

  // helpers
  ok(_lb(true) && _lb('TRUE') && _lb('True') && !_lb(false) && !_lb('FALSE') && !_lb(''), 'boolean coercion');
  ok(_ln('1,234.50') === 1234.5 && _ln('$10') === 10 && _ln('x') === null, 'number coercion');
  ok(_lmonth('2026-06-15') === '2026-06' && _lmonth('2026-06') === '2026-06' && _lmonth('junk') === '', 'month coercion');
  ok(_ldate(new Date(2026, 5, 15)) === '2026-06-15', 'date from Date object');

  // aggregation rule on a fixture: every dollar exactly once
  var T = [LEDGER_HEADERS[LEDGER_TXNS],
    [1, '2026-06-02', '2026-06', 'Transfer in',  -1562.5, 'Checking', 'onetime', 'FALSE', 'FALSE', 'FALSE'], // checking negative: auto-transfer
    [2, '2026-06-03', '2026-06', 'Beach chairs',  107.99, 'Costco',   'onetime', 'TRUE',  'TRUE',  'TRUE'],
    [3, '2026-06-22', '2026-06', 'Cash Back',     -43.67, 'Prime',    'onetime', 'FALSE', 'TRUE',  'TRUE'],  // card negative: nets in
    [4, '2026-06-22', '2026-06', 'Wegmans',        129.02, 'AMEX',    'groceries','FALSE','TRUE',  'TRUE'],
    [5, '2026-06-07', '2026-06', 'Gas',             32.63, 'Costco',  'gas',     'FALSE', 'TRUE',  'TRUE'],
    [6, '2026-06-10', '2026-06', 'Claude sub',      21.60, 'Wells Fargo','onetime','TRUE','TRUE',  'FALSE'], // ben-only disc
    [7, '2026-07-01', '2026-06', 'July-dated, June books', 10.00, 'Prime', 'onetime', 'FALSE', 'TRUE', 'TRUE'],
    [8, '2026-06-15', '2026-06', 'Move to 529',   5000.00, 'Checking', 'transfer', 'FALSE', 'FALSE', 'FALSE']]; // explicit transfer: positive, excluded
  var F = [LEDGER_HEADERS[LEDGER_FIXED],
    [1, '2026-06', 'Electricity + Gas', 'Checking', 298.24, 'TRUE',  '2026-06-12'],
    [2, '2026-06', 'Peloton+',          'Wells Fargo', 49.99, 'FALSE', ''],           // unpaid: excluded
    [3, '2026-06', 'Extra loan payments','Checking', 100.00, 'TRUE',  '2026-06-15']]; // counts (by rule)
  var I = [LEDGER_HEADERS[LEDGER_INCOME],
    [1, '2026-06', '2026-06-05', 'Schrodinger', 4237.93]];
  var agg = _ledgerAggregate(T, F, I);
  var b = agg.perMonth['2026-06'];
  ok(!!b, 'month bucket exists');
  if (b) {
    ok(near(b.transfers, -1562.5 + 5000.00), 'transfers: checking-negative + explicit, both signs: ' + b.transfers);
    ok(near(b.onetime, 107.99 - 43.67 + 21.60 + 10.00), 'onetime nets cash back, includes month-assigned row: ' + b.onetime);
    ok(near(b.groceries, 129.02) && near(b.gas, 32.63), 'category sums');
    ok(near(b.fixedPaid, 398.24), 'fixedPaid = paid rows incl extra loan, excl unpaid: ' + b.fixedPaid);
    ok(near(b.totalExpenses, (107.99 - 43.67 + 21.60 + 10.00) + 129.02 + 32.63 + 398.24), 'every dollar exactly once: ' + b.totalExpenses);
    ok(near(b.discretionary, 107.99 + 21.60), 'discretionary attribute, not addend');
    ok(near(b.benDisc, 107.99 / 2 + 21.60) && near(b.jennaDisc, 107.99 / 2), 'ben/jenna split shares');
    ok(near(b.net, 4237.93 - b.totalExpenses), 'net = income - totalExpenses');
  }
  ok(agg.months.length === 1 && agg.months[0] === '2026-06', 'months list');

  // validation rejections
  ok(!!_ledgerValidateTxn({ date: 'nope', amount: 5, method: 'AMEX', description: 'x' }).error, 'rejects bad date');
  ok(!!_ledgerValidateTxn({ date: '2026-06-01', amount: 0, method: 'AMEX', description: 'x' }).error, 'rejects zero amount');
  ok(!!_ledgerValidateTxn({ date: '2026-06-01', amount: 5, method: '', description: 'x' }).error, 'rejects empty method');
  ok(!!_ledgerValidateTxn({ date: '2026-06-01', amount: 5, method: 'AMEX', description: '', category: 'onetime' }).error, 'rejects empty onetime description');
  ok(!_ledgerValidateTxn({ date: '2026-06-01', amount: 5, method: 'Costco', description: '', category: 'gas' }).error, 'gas defaults description');

  // split preparation: conservation invariant, inheritance, rejections (pure)
  var so = { date: '2026-06-10', month: '2026-06', description: 'Costco run', amount: 150.00,
             method: 'Costco', category: 'onetime', discretionary: false, ben: true, jenna: true };
  var sp = _ledgerPrepSplit(so, [
    { description: 'Groceries part', amount: 100.00, category: 'groceries' },
    { description: 'Beach towels', amount: 50.00, discretionary: true }]);
  ok(!sp.error && sp.recs.length === 2, 'split accepts conserving parts: ' + (sp.error || 'ok'));
  ok(!sp.error && sp.recs[0][5] === 'groceries' && sp.recs[1][4] === 'Costco', 'split parts inherit method, override category');
  ok(!!_ledgerPrepSplit(so, [{ description: 'a', amount: 100 }, { description: 'b', amount: 49 }]).error, 'split rejects non-conserving parts');
  ok(!!_ledgerPrepSplit(so, [{ description: 'a', amount: 150 }]).error, 'split rejects single part');

  // write paths on a scratch sheet, two phases: txn lifecycle, then fixed
  // lifecycle (same tab, re-headed between phases). Never touches live sheets.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scratch = ss.getSheetByName('_TestScratchLedger_');
  if (!scratch) {
    scratch = ss.insertSheet('_TestScratchLedger_'); // create once, reused forever
    scratch.hideSheet();
  }
  function rehead(sheetKey) {
    scratch.clearContents();
    var hdr = LEDGER_HEADERS[sheetKey];
    scratch.getRange(1, 1, 1, hdr.length).setValues([hdr]);
    invalidateSheet('_TestScratchLedger_');
  }

  // -- phase 1: txn lifecycle --
  rehead(LEDGER_TXNS);
  var real = LEDGER_TXNS;
  try {
    LEDGER_TXNS = '_TestScratchLedger_';
    var r1 = ledgerAddTxn({ date: '2026-06-01', amount: 12.5, method: 'AMEX', description: 'test a' });
    var r2 = ledgerAddTxn({ date: '2026-06-02', amount: 7.25, method: 'Prime', description: 'test b', discretionary: true, ben: true });
    ok(r1.ok && r1.id === 1, 'first id = 1: ' + JSON.stringify(r1));
    ok(r2.ok && r2.id === 2, 'second id = 2: ' + JSON.stringify(r2));
    var got = getLedgerTxns('2026-06');
    ok(got.txns && got.txns.length === 2 && got.txns[1].discretionary === true, 'read back by month with flags');

    var ru = ledgerUpdateTxn({ id: 2, date: '2026-06-03', month: '2026-05', description: 'test b fixed',
                               amount: 9.99, method: 'Costco', category: 'gas' });
    ok(ru.ok && ru.month === '2026-05', 'update rewrites record incl month reassignment: ' + JSON.stringify(ru));
    got = getLedgerTxns('2026-05');
    ok(got.txns.length === 1 && near(got.txns[0].amount, 9.99) && got.txns[0].method === 'Costco', 'updated row reads back in new month');
    ok(!!ledgerUpdateTxn({ id: 99, date: '2026-06-01', amount: 1, method: 'AMEX', description: 'x' }).error, 'update of missing id fails loudly');

    var rs = ledgerSplitTxn({ id: 1, parts: [
      { description: 'part 1', amount: 8.5, category: 'groceries' },
      { description: 'part 2', amount: 4.0, discretionary: true }] });
    ok(rs.ok && rs.ids.length === 2 && rs.ids[0] === 1 && rs.ids[1] === 3, 'split keeps original id, appends next id: ' + JSON.stringify(rs));
    got = getLedgerTxns('2026-06');
    var splitSum = 0;
    for (var si = 0; si < got.txns.length; si++) splitSum += got.txns[si].amount;
    ok(got.txns.length === 2 && near(splitSum, 12.5), 'split conserved the amount on-sheet: ' + splitSum);
    ok(!!ledgerSplitTxn({ id: 3, parts: [{ description: 'a', amount: 1 }, { description: 'b', amount: 1 }] }).error, 'non-conserving split rejected at the sheet');

    var rd = ledgerDeleteTxn({ id: 3 });
    ok(rd.ok, 'delete by id: ' + JSON.stringify(rd));
    ok(!!ledgerDeleteTxn({ id: 3 }).error, 'repeat delete fails loudly');
    ok(getLedgerTxns('2026-06').txns.length === 1, 'deleted row is gone');
  } finally {
    LEDGER_TXNS = real;
    invalidateSheet('_TestScratchLedger_');
  }

  // -- phase 2: fixed lifecycle --
  rehead(LEDGER_FIXED);
  var realF = LEDGER_FIXED;
  try {
    LEDGER_FIXED = '_TestScratchLedger_';
    var f1 = ledgerAddFixed({ month: '2026-06', name: 'Mortgage', source: 'Checking', amount: 2000 });
    var f2 = ledgerAddFixed({ month: '2026-06', name: 'Spotify', source: 'AMEX', amount: 19.99, paid: true, paid_date: '2026-06-09' });
    ok(f1.ok && f1.id === 1 && f2.ok && f2.id === 2, 'fixed adds with sequential ids');
    var gf = getLedgerFixed('2026-06');
    ok(gf.fixed.length === 2 && gf.fixed[1].paid === true && gf.fixed[1].paid_date === '2026-06-09', 'fixed read back with paid state');

    ok(ledgerUpdateFixed({ id: 1, paid: true }).ok, 'toggle paid on');
    gf = getLedgerFixed('2026-06');
    ok(gf.fixed[0].paid === true && !!gf.fixed[0].paid_date, 'paid=true stamps a date');
    ok(ledgerUpdateFixed({ id: 1, paid: false }).ok && getLedgerFixed('2026-06').fixed[0].paid_date === '', 'paid=false clears the date');
    ok(ledgerUpdateFixed({ id: 1, amount: 2100.5 }).ok && near(getLedgerFixed('2026-06').fixed[0].amount, 2100.5), 'partial amount update');
    ok(!!ledgerUpdateFixed({ id: 1, paid_date: '2026-06-15' }).error, 'paid_date on unpaid row refused');

    var sd = ledgerSeedFixedMonth({ month: '2026-07' });
    ok(sd.ok && sd.count === 2 && sd.seededFrom === '2026-06', 'seed copies prior month: ' + JSON.stringify(sd));
    gf = getLedgerFixed('2026-07');
    ok(gf.fixed.length === 2 && gf.fixed[0].paid === false && gf.fixed[0].paid_date === '', 'seeded rows are unpaid');
    ok(!!ledgerSeedFixedMonth({ month: '2026-07' }).error, 'double-seed refused');
  } finally {
    LEDGER_FIXED = realF;
    invalidateSheet('_TestScratchLedger_');
  }

  if (fails.length) { Logger.log('testLedger FAILURES:\n' + fails.join('\n')); }
  Logger.log('testLedger: ' + passes + ' passed, ' + fails.length + ' failed');
  return { passed: passes, failed: fails };
}
