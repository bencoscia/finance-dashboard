// ============================================================
// perf.gs -- UX_SPEC Phase 1B (instrumentation) + Phase 2A
//
// New file in the Apps Script project (File > + > Script, name it "perf").
// Apps Script flattens every file into one global scope, so nothing here may
// reuse a top-level name from another file (and there must be NO doGet/doPost
// here -- a second doGet silently replaces the router; see INDEX section 4).
//
// Contents:
//   _perfTimeIt / _perfStampMs        -- server timing (_t per section of all, _ms per response)
//   _ledgerCleanTags          -- the ONE tag normaliser (ledgerSetTags + ledgerAddTxn)
//   _ledgerCardTotals         -- pure: per-method TTM charges (UX_SPEC D4)
//   ledgerCardTotals          -- reader wrapper; getCardTotals delegates here post-cutover
//   testLedgerCardTotals      -- fixtures only, never touches a sheet (fast suite)
// ============================================================

// Run fn, record its wall time in t[name] (ms), return its result. Records the
// time even when fn throws, so a failing section still shows up in _t.
function _perfTimeIt(t, name, fn) {
  var t0 = Date.now();
  try { return fn(); }
  finally { t[name] = Date.now() - t0; }
}

// Server execution ms on a response object. Arrays / non-objects are left alone.
// Safe after _withIdem: the idem cache entry is serialized before this runs, so
// _ms never lands in a cached result.
function _perfStampMs(data, t0) {
  if (data && typeof data === 'object' && !(data instanceof Array)) data._ms = Date.now() - t0;
  return data;
}

// Tags: lowercase, trimmed, deduped, no commas (col K is comma-joined), no blanks.
// Accepts an array, a JSON array string, or a comma string.
function _ledgerCleanTags(tags) {
  if (typeof tags === 'string') {
    try { tags = JSON.parse(tags); } catch (e) { tags = String(tags).split(','); }
  }
  if (!tags || !(tags instanceof Array)) tags = [];
  var clean = [];
  for (var i = 0; i < tags.length; i++) {
    if (tags[i] === null || tags[i] === undefined) continue;
    var t = String(tags[i]).trim().toLowerCase();
    if (t && t.indexOf(',') < 0 && clean.indexOf(t) < 0) clean.push(t);
  }
  return clean;
}

// Lap timers for breaking one function's time into named steps.
// t = _perfLaps(); ...work...; _perfLap(t, 'step'); ...; _perfLapsDone(t) -> {step: ms, ...}
function _perfLaps() { return { _at: Date.now() }; }
function _perfLap(t, name) {
  if (!t) return;
  var now = Date.now();
  t[name] = (t[name] || 0) + (now - t._at);
  t._at = now;
}
function _perfLapsDone(t) { if (t) delete t._at; return t; }
// base + prefixed copy of sub (sub may be null). Returns base.
function _perfMergeT(base, sub, prefix) {
  if (sub) for (var k in sub) if (typeof sub[k] === 'number') base[prefix + k] = sub[k];
  return base;
}
// Set by _computeCardBalances / ledgerComputeMonths on a real (uncached) compute.
// Apps Script re-evaluates globals per execution, so these never leak across requests.
var _perfCardT = null, _perfMonthsT = null;

// 'YYYY-MM' + n months -> 'YYYY-MM'
function _perfMkAdd(mk, n) {
  var y = parseInt(mk.slice(0, 4), 10), m = parseInt(mk.slice(5, 7), 10) - 1 + n;
  y += Math.floor(m / 12); m = ((m % 12) + 12) % 12;
  return y + '-' + ('0' + (m + 1)).slice(-2);
}

// -- Phase 2A: ledger-backed card totals (UX_SPEC decision D4) -----------------
// Pure over raw values() arrays (header row included), so fixtures can drive it.
//   Numerator, per payment method: Txns.amount (all categories; refunds net as
//     negatives) + PAID Fixed Log rows whose source is that method. Card payments
//     live in 'Card Payments', not Txns, so they are excluded by construction.
//     Checking only: transfers are money movement, not spend -- excluded with the
//     same predicate _ledgerAggregate uses (category 'transfer', or a negative
//     Checking amount).
//   Window: the 12 most recent COMPLETE months by month key; curMonthKey is
//     excluded. If the ledger's history is shorter, the window starts at the
//     earliest month present in either table.
//   Divisor (returned as months): months in the window -- never the number of
//     months that happen to have rows for a method.
// Returns { totals, counts, months, label:'TTM', from, to } -- the legacy shape
// plus from/to month keys (null when the window is empty).
function _ledgerCardTotals(txnRows, fixedRows, curMonthKey) {
  var to = _perfMkAdd(curMonthKey, -1), from = _perfMkAdd(curMonthKey, -12);
  var earliest = '', i, r, mk;
  function seen(m) { if (m && (!earliest || m < earliest)) earliest = m; }
  for (i = 1; i < txnRows.length; i++)   seen(_lmonth(txnRows[i][2]) || _lmonth(txnRows[i][1]));
  for (i = 1; i < fixedRows.length; i++) seen(_lmonth(fixedRows[i][1]));
  if (earliest && earliest > from) from = earliest;
  if (!earliest || from > to) {
    return { totals: {}, counts: {}, months: 0, label: 'TTM', from: null, to: null };
  }
  var months = 0, k = from;
  while (k <= to && months < 13) { months++; k = _perfMkAdd(k, 1); }

  var totals = {}, counts = {};
  function add(method, amt) {
    totals[method] = (totals[method] || 0) + amt;
    counts[method] = (counts[method] || 0) + 1;
  }
  for (i = 1; i < txnRows.length; i++) {
    r = txnRows[i];
    mk = _lmonth(r[2]) || _lmonth(r[1]);
    var amt = _ln(r[4]), method = String(r[5] || '').trim();
    if (!mk || mk < from || mk > to || amt === null || !method) continue;
    if (method === 'Checking' && (String(r[6] || '') === 'transfer' || amt < 0)) continue;
    add(method, amt);
  }
  for (i = 1; i < fixedRows.length; i++) {
    r = fixedRows[i];
    mk = _lmonth(r[1]);
    var famt = _ln(r[4]), src = String(r[3] || '').trim();
    if (!mk || mk < from || mk > to || famt === null || !src || !_lb(r[5])) continue;
    add(src, famt);
  }
  for (var m2 in totals) totals[m2] = Math.round(totals[m2] * 100) / 100;
  return { totals: totals, counts: counts, months: months, label: 'TTM', from: from, to: to };
}

// Cached like the monthly aggregate (same 5-min TTL) and cleared by
// invalidateLedgerCache, i.e. by every ledger write and by onEdit on Txns /
// Fixed Log. The key carries the current month, so a month rollover can never
// serve last month's window.
function _perfCardTotalsKey() {
  return LEDGER_CACHE_VERSION + '_cardtotals_' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
}
function ledgerCardTotals() {
  var cache = CacheService.getScriptCache(), key = _perfCardTotalsKey();
  try { var hit = cache.get(key); if (hit) return JSON.parse(hit); } catch (e) {}
  var txnRows = readSheet(LEDGER_TXNS);
  if (!txnRows) return { error: 'Sheet not found: ' + LEDGER_TXNS };
  var fixedRows = readSheet(LEDGER_FIXED) || [[]];
  var cur = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  var out = _ledgerCardTotals(txnRows, fixedRows, cur);
  try { cache.put(key, JSON.stringify(out), 300); } catch (e2) {}
  return out;
}

// -- Tests (pure; fixtures only) ----------------------------------------------
function testLedgerCardTotals() {
  section('Ledger card totals (D4) + tag normaliser');
  var TH = ['id','date','month','description','amount','method','category','discretionary','ben','jenna','tags','pending'];
  var FH = ['id','month','name','source','amount','paid','paid_date'];
  var T = [TH,
    [1, '2025-10-31', '2025-10', 'before window',   500.00, 'Wells Fargo', 'onetime', 'FALSE','TRUE','TRUE'],
    [2, '2025-11-02', '2025-11', 'first month',      10.00, 'Wells Fargo', 'onetime', 'FALSE','TRUE','TRUE'],
    [3, '2026-03-05', '2026-03', 'Wegmans',          40.00, 'Wells Fargo', 'groceries','FALSE','TRUE','TRUE'],
    [4, '2026-03-09', '2026-03', 'refund',          -15.00, 'Wells Fargo', 'onetime', 'FALSE','TRUE','TRUE'],
    [5, '2026-11-01', '2025-10', 'Nov-dated, Oct books', 7.00, 'Wells Fargo', 'onetime', 'FALSE','TRUE','TRUE'], // month key wins: out
    [6, '2026-10-30', '2026-10', 'last month',        5.00, 'AMEX',        'gas',     'FALSE','TRUE','TRUE'],
    [7, '2026-11-03', '2026-11', 'current month',   999.00, 'AMEX',        'onetime', 'FALSE','TRUE','TRUE'], // partial month: out
    [8, '2026-06-01', '2026-06', 'Move to 529',    1000.00, 'Checking',    'transfer','FALSE','FALSE','FALSE'],
    [9, '2026-06-02', '2026-06', 'Transfer in',    -200.00, 'Checking',    'onetime', 'FALSE','FALSE','FALSE'],
    [10,'2026-06-03', '2026-06', 'Hardware store',   30.00, 'Checking',    'onetime', 'FALSE','TRUE','TRUE']];
  var F = [FH,
    [1, '2026-02', 'Peloton+',  'Wells Fargo', 49.99, 'TRUE',  '2026-02-03'],
    [2, '2026-02', 'Unpaid',    'Wells Fargo', 80.00, 'FALSE', ''],
    [3, '2026-11', 'Current',   'Wells Fargo', 12.00, 'TRUE',  '2026-11-02'],
    [4, '2026-04', 'Electric',  'Checking',   100.00, 'TRUE',  '2026-04-12']];
  var r = _ledgerCardTotals(T, F, '2026-11');
  assertEqual([r.from, r.to, r.months], ['2025-11', '2026-10', 12], 'window = 12 complete months, current month excluded');
  assertApprox(r.totals['Wells Fargo'], 10 + 40 - 15 + 49.99, 'WF: all categories, refund nets, paid fixed in, unpaid/current/pre-window out');
  assertEqual(r.counts['Wells Fargo'], 4, 'WF count = 3 txns + 1 paid fixed');
  assertApprox(r.totals['AMEX'], 5, 'AMEX: partial current month excluded');
  assertApprox(r.totals['Checking'], 30 + 100, 'Checking: explicit + negative transfers excluded, paid fixed in');
  assert(r.label === 'TTM', 'legacy label kept');

  // short history: earliest month bounds the window; divisor = months in window
  var T2 = [TH, [1, '2026-01-10', '2026-01', 'x', 12.00, 'Citi', 'onetime', 'FALSE','TRUE','TRUE']];
  var r2 = _ledgerCardTotals(T2, [FH], '2026-03');
  assertEqual([r2.from, r2.to, r2.months], ['2026-01', '2026-02', 2], 'short history: window starts at earliest ledger month');
  // sparse rows: divisor is NOT the count of months with rows
  var T3 = [TH, [1, '2026-02-10', '2026-02', 'x', 120.00, 'Citi', 'onetime', 'FALSE','TRUE','TRUE'],
                [2, '2025-06-10', '2025-06', 'old', 1.00, 'Prime Visa', 'onetime', 'FALSE','TRUE','TRUE']];
  assertEqual(_ledgerCardTotals(T3, [FH], '2026-11').months, 12, 'divisor stays 12 when only 1 month has rows');
  // empty / brand-new ledger
  assertEqual(_ledgerCardTotals([TH], [FH], '2026-11').months, 0, 'empty ledger -> months 0, no divide-by-zero downstream');
  assertEqual(_ledgerCardTotals([TH, [1,'2026-11-01','2026-11','x',1,'Citi','onetime','','','']], [FH], '2026-11').from, null, 'only current-month rows -> empty window');
  assertEqual(_perfMkAdd('2026-01', -1), '2025-12', 'month math across year');

  // tag normaliser shared by ledgerSetTags and ledgerAddTxn
  assertEqual(_ledgerCleanTags([' Trip ', 'trip', 'a,b', '', null, 'Kids']), ['trip', 'kids'], 'tags: trim, lowercase, dedupe, drop commas/blanks/null');
  assertEqual(_ledgerCleanTags('["x","X"]'), ['x'], 'tags: JSON string');
  assertEqual(_ledgerCleanTags('x, y'), ['x', 'y'], 'tags: comma string');
  assertEqual(_ledgerCleanTags(undefined), [], 'tags: missing -> []');

  // lap timers
  var lt = _perfLaps(); _perfLap(lt, 'a'); _perfLap(lt, 'a'); _perfLapsDone(lt);
  assert(typeof lt.a === 'number' && lt._at === undefined, 'laps accumulate a repeated step and drop _at');
  assertEqual(_perfMergeT({ x: 1 }, { y: 2, z: 'no' }, 'm.'), { x: 1, 'm.y': 2 }, 'merge prefixes numeric fields only');
  assertEqual(_perfMergeT({ x: 1 }, null, 'm.'), { x: 1 }, 'merge tolerates null');

  // _perfStampMs never touches arrays; _perfTimeIt records even on throw
  var a = [1]; _perfStampMs(a, Date.now()); assert(a._ms === undefined, '_perfStampMs skips arrays');
  var t = {}; try { _perfTimeIt(t, 'boom', function () { throw new Error('x'); }); } catch (e) {}
  assert(typeof t.boom === 'number', '_perfTimeIt records a throwing section');
}
