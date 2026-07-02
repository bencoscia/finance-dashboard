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

// One-time setup: create the four sheets with headers. Never overwrites --
// an existing sheet is left untouched and reported. Run from the editor,
// then import the migration CSVs (paste starting at row 2, matching order).
function setupLedgerSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var made = [], skipped = [], names = [LEDGER_TXNS, LEDGER_FIXED, LEDGER_INCOME, LEDGER_VARHIST];
  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    if (ss.getSheetByName(name)) { skipped.push(name); continue; }
    var sh = ss.insertSheet(name);
    var hdr = LEDGER_HEADERS[name];
    sh.getRange(1, 1, 1, hdr.length).setValues([hdr]).setFontWeight('bold');
    sh.setFrozenRows(1);
    made.push(name);
  }
  return { ok: true, created: made, skippedExisting: skipped };
}

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
  if (!txnRows) return { error: 'Sheet not found: ' + LEDGER_TXNS + ' -- run setupLedgerSheets() and import the migration CSVs.' };
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

// -- WRITE: ledgerAddTxn ----------------------------------------------------
// Validated append. id = max existing id + 1 under LockService (same pattern
// as addHsaReceipt -- two simultaneous adds must never share an id, because
// future edit/delete paths locate rows by id). Wrapped in _withIdem by both
// routers; invalidates the ledger cache internally on success.
function ledgerAddTxn(p) {
  var dateStr = _ldate(p.date);
  if (!dateStr) return { error: 'Invalid or missing date (need YYYY-MM-DD).' };
  var dchk = new Date(dateStr + 'T12:00:00');
  if (isNaN(dchk.getTime())) return { error: 'Not a real date: ' + dateStr };

  var amount = _ln(p.amount);
  if (amount === null || amount === 0) return { error: 'Amount must be a nonzero number.' };

  var category = String(p.category || 'onetime');
  if (LEDGER_CATEGORIES.indexOf(category) < 0)
    return { error: 'Unknown category: ' + category + ' (want onetime|groceries|gas)' };

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

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LEDGER_TXNS);
  if (!sheet) return { error: 'Sheet not found: ' + LEDGER_TXNS + ' -- run setupLedgerSheets().' };

  var lock = LockService.getScriptLock();
  var gotLock = false;
  try { gotLock = lock.tryLock(5000); } catch (e) {}
  if (!gotLock) return { error: 'Could not acquire lock -- please try again.' };

  try {
    invalidateSheet(LEDGER_TXNS);
    var values = readSheet(sheet);
    var maxId = 0;
    for (var i = 1; i < values.length; i++) {
      var n = parseInt(values[i][0], 10);
      if (!isNaN(n) && n > maxId) maxId = n;
    }
    var newId = maxId + 1;
    var row = Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(row, 1, 1, 10).setValues([[
      newId, dateStr, mk, description, amount, method, category,
      !!p.discretionary, !!p.ben, !!p.jenna
    ]]);
    SpreadsheetApp.flush();
    invalidateLedgerCache();
    invalidateSheet(LEDGER_TXNS);
    return { ok: true, id: newId, month: mk };
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
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
  ok(!!ledgerAddTxnValidateOnly({ date: 'nope', amount: 5, method: 'AMEX', description: 'x' }), 'rejects bad date');
  ok(!!ledgerAddTxnValidateOnly({ date: '2026-06-01', amount: 0, method: 'AMEX', description: 'x' }), 'rejects zero amount');
  ok(!!ledgerAddTxnValidateOnly({ date: '2026-06-01', amount: 5, method: '', description: 'x' }), 'rejects empty method');
  ok(!!ledgerAddTxnValidateOnly({ date: '2026-06-01', amount: 5, method: 'AMEX', description: '', category: 'onetime' }), 'rejects empty onetime description');
  ok(!ledgerAddTxnValidateOnly({ date: '2026-06-01', amount: 5, method: 'Costco', description: '', category: 'gas' }), 'gas defaults description');

  // write path on scratch sheet: sequential ids
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scratch = ss.getSheetByName('_TestScratchLedger_');
  if (!scratch) {
    scratch = ss.insertSheet('_TestScratchLedger_');
    scratch.hideSheet();
  }
  scratch.clearContents();
  var hdr = LEDGER_HEADERS[LEDGER_TXNS];
  scratch.getRange(1, 1, 1, hdr.length).setValues([hdr]);
  var real = LEDGER_TXNS;
  try {
    LEDGER_TXNS = '_TestScratchLedger_';
    var r1 = ledgerAddTxn({ date: '2026-06-01', amount: 12.5, method: 'AMEX', description: 'test a' });
    var r2 = ledgerAddTxn({ date: '2026-06-02', amount: 7.25, method: 'Prime', description: 'test b', discretionary: true, ben: true });
    ok(r1.ok && r1.id === 1, 'first id = 1: ' + JSON.stringify(r1));
    ok(r2.ok && r2.id === 2, 'second id = 2: ' + JSON.stringify(r2));
    var got = getLedgerTxns('2026-06');
    ok(got.txns && got.txns.length === 2 && got.txns[1].discretionary === true, 'read back by month with flags');
  } finally {
    LEDGER_TXNS = real;
    invalidateSheet('_TestScratchLedger_');
  }

  if (fails.length) { Logger.log('testLedger FAILURES:\n' + fails.join('\n')); }
  Logger.log('testLedger: ' + passes + ' passed, ' + fails.length + ' failed');
  return { passed: passes, failed: fails };
}

// Validation-only twin used by tests (returns the error string or null,
// never touches the sheet). Keep its checks in lockstep with ledgerAddTxn.
function ledgerAddTxnValidateOnly(p) {
  var dateStr = _ldate(p.date);
  if (!dateStr) return 'bad date';
  var amount = _ln(p.amount);
  if (amount === null || amount === 0) return 'bad amount';
  var category = String(p.category || 'onetime');
  if (LEDGER_CATEGORIES.indexOf(category) < 0) return 'bad category';
  if (!String(p.method || '').trim()) return 'bad method';
  if (!String(p.description || '').trim() && category !== 'gas') return 'bad description';
  return null;
}
