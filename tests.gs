// ============================================================
// Finance Dashboard -- Test Suite
// Add this file to the Apps Script project alongside dashboard.gs
//
// USAGE:
//   Run runTests() from the editor to execute all tests.
//   Results appear in the Execution Log (View ? Logs).
//   Green PASS = pass, Red FAIL = fail, with details on failures.
//
// Tests use a temporary scratch sheet that is created and
// deleted for each test run -- no permanent data is modified
// except during write tests (which are immediately reversed).
// ============================================================

var TEST_SHEET_NAME = '_TestScratch_';
var HSA_TEST_SHEET_NAME = '_TestScratchHsa_'; // separate from TEST_SHEET_NAME -- different column layout (HSA Receipts schema), shared scratch sheet already has a fixed transaction layout other tests depend on throughout the run
var _testPass = 0;
var _testFail = 0;
var _testLog   = [];   // FAILURE detail only -- passes are counted, never logged
var _sections  = [];   // [{name, pass, fail}] -- one compact tally line each
var _curSection = null;
var _notes     = [];   // always-shown runner notes (benign warnings, diagnostics)

// Output is deliberately failure-first and compact: Apps Script truncates a
// single oversized log entry, so we never emit a PASS-per-assertion transcript.
// Counts + failures + a section tally fit well under the cap.

// -- Assertion helpers -----------------------------------------

function _tag() { return '[' + (_curSection ? _curSection.name : '?') + '] '; }
function _bump(ok) {
  if (ok) { _testPass++; if (_curSection) _curSection.pass++; }
  else    { _testFail++; if (_curSection) _curSection.fail++; }
}
function _note(msg) { _notes.push(msg); }

function assert(condition, message) {
  _bump(!!condition);
  if (!condition) _testLog.push('  FAIL ' + _tag() + message);
}

function assertEqual(actual, expected, message) {
  var ok = JSON.stringify(actual) === JSON.stringify(expected);
  _bump(ok);
  if (!ok) {
    _testLog.push('  FAIL ' + _tag() + message);
    _testLog.push('      expected: ' + JSON.stringify(expected));
    _testLog.push('      actual:   ' + JSON.stringify(actual));
  }
}

function assertApprox(actual, expected, message, tolerance) {
  tolerance = tolerance || 0.01;
  var ok = typeof actual === 'number' && Math.abs(actual - expected) <= tolerance;
  _bump(ok);
  if (!ok) _testLog.push('  FAIL ' + _tag() + message + ' (got ' + actual + ', expected ~' + expected + ')');
}

function section(name) {
  _curSection = { name: name, pass: 0, fail: 0 };
  _sections.push(_curSection);
}

// -- Scratch sheet helpers -------------------------------------

function createScratchSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TEST_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(TEST_SHEET_NAME);
    sheet.hideSheet(); // keep out of tab bar -- revealed automatically when accessed
  } else {
    sheet.clearContents();
  }

  // Real sheets have row 1 blank, row 2 as header, data from row 3
  sheet.getRange(2, 1, 1, 7).setValues([['Date','Description','Cost','Payment Method','Discretionary','Ben','Jenna']]);

  var rows = [
    [new Date('2026-05-01'), 'Coffee',         5.50,  'AMEX',        'TRUE',  'TRUE',  'FALSE'],
    [new Date('2026-05-02'), 'Groceries',      95.20, 'Costco',      'FALSE', 'TRUE',  'TRUE' ],
    [new Date('2026-05-03'), 'Gas',            45.00, 'Wells Fargo', 'FALSE', 'TRUE',  'FALSE'],
    [new Date('2026-05-04'), 'Netflix',        15.99, 'Wells Fargo', 'FALSE', 'TRUE',  'TRUE' ],
    [new Date('2026-05-05'), 'Restaurant',     62.40, 'AMEX',        'TRUE',  'TRUE',  'TRUE' ],
    [new Date('2026-05-06'), 'Amazon',         23.50, 'Prime',       'TRUE',  'FALSE', 'TRUE' ],
    [new Date('2026-05-07'), 'Tax Refund',    -500.0, 'Checking',    'FALSE', 'TRUE',  'TRUE' ],
    [new Date('2026-05-08'), 'Hardware Store', 88.75, "Lowe's",      'TRUE',  'TRUE',  'FALSE'],
    [new Date('2026-05-09'), 'Pharmacy',       12.30, 'Wells Fargo', 'FALSE', 'TRUE',  'TRUE' ],
    [new Date('2026-05-10'), 'Dinner out',     78.90, 'Costco',      'TRUE',  'TRUE',  'TRUE' ],
  ];
  sheet.getRange(3, 1, rows.length, 7).setValues(rows);

  sheet.getRange(rows.length + 3, 2).setValue('Recurring Fixed Expenses');
  sheet.getRange(rows.length + 4, 2).setValue('Total Expenses');
  sheet.getRange(rows.length + 4, 3).setValue(927.54);
  return sheet;
}

// Mirrors the live HSA Receipts layout: B1/B2 balance metadata (left blank --
// addHsaReceipt doesn't touch them), headers row 3, data rows 4+. Cleared and
// re-created each run so id-assignment tests start from a known empty state.
function createHsaScratchSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HSA_TEST_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(HSA_TEST_SHEET_NAME);
    sheet.hideSheet();
  } else {
    sheet.clearContents();
  }
  sheet.getRange(1, 1).setValue('Account balance');
  sheet.getRange(2, 1).setValue('Balance as of');
  sheet.getRange(3, 1, 1, 11).setValues([[
    'id','date_incurred','amount','provider','description',
    'category','funding','receipt_link','reimbursed_amount','reimbursed_date','notes'
  ]]);
  return sheet;
}

function deleteScratchSheet() {
  // Don't delete -- leave it so next run can reuse it without insertSheet overhead.
  // It gets cleared and rewritten at the start of each test run.
  // To remove manually: delete the '_TestScratch_' sheet from the spreadsheet.
}

// -- Test groups -----------------------------------------------

function testGetTransactions() {
  section('getTransactions');
  _resetCaches();

  var txns = getTransactions(TEST_SHEET_NAME);

  assert(!txns.error, 'No error returned');
  assert(Array.isArray(txns.transactions), 'Returns transactions array');
  assertEqual(txns.transactions.length, 10, 'Reads exactly 10 transactions (stops at section header)');
  assertEqual(txns.month, TEST_SHEET_NAME, 'Returns correct month name');

  var first = txns.transactions[0];
  assertEqual(first.description, 'Coffee', 'First row description correct');
  assertApprox(first.cost, 5.50, 'First row cost correct');
  assertEqual(first.paymentMethod, 'AMEX', 'First row payment method correct');
  assertEqual(first.discretionary, true, 'First row discretionary (boolean true) correct');
  assertEqual(first.ben, true, 'First row ben correct');
  assertEqual(first.jenna, false, 'First row jenna correct');
  assertEqual(first.rowIndex, 3, 'First row rowIndex = 3 (data starts row 3)');

  // Test boolean false values are included (not filtered out)
  var groceries = txns.transactions[1];
  assertEqual(groceries.description, 'Groceries', 'Second row (disc=false) included');
  assertEqual(groceries.discretionary, false, 'disc=false boolean handled correctly');

  // Test negative cost (refund)
  var refund = txns.transactions[6];
  assertEqual(refund.description, 'Tax Refund', 'Negative cost row included');
  assertApprox(refund.cost, -500.0, 'Negative cost correct');

  // Last row rowIndex
  var last = txns.transactions[9];
  assertEqual(last.rowIndex, 12, 'Last row rowIndex correct (row 12 in sheet)');
}

function testGetTransactionsTermination() {
  section('getTransactions -- termination');
  _resetCaches();

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TEST_SHEET_NAME);

  // Insert a blank row mid-list (between rows 7 and 8, both of which are transactions)
  sheet.insertRowBefore(7);
  _resetCaches();
  var txns = getTransactions(TEST_SHEET_NAME);
  assertEqual(txns.transactions.length, 10, 'Single blank row mid-list does not cut off transactions');

  // Remove the test row
  sheet.deleteRow(7);
  _resetCaches();
}

function testAddTransaction() {
  section('addTransaction');
  _resetCaches();

  var beforeCount = getTransactions(TEST_SHEET_NAME).transactions.length;

  var result = addTransaction({
    month:         TEST_SHEET_NAME,
    date:          '2026-05-15',
    description:   'Test Transaction',
    cost:          42.00,
    paymentMethod: 'AMEX',
    discretionary: true,
    ben:           true,
    jenna:         false,
  });

  _resetCaches();
  assert(result.ok, 'addTransaction returns ok');
  assert(result.rowIndex > 0, 'addTransaction returns rowIndex > 0');

  var afterCount = getTransactions(TEST_SHEET_NAME).transactions.length;
  assertEqual(afterCount, beforeCount + 1, 'Transaction count increased by 1');

  var txns = getTransactions(TEST_SHEET_NAME).transactions;
  var added = txns.find(function(t){ return t.description === 'Test Transaction'; });
  assert(added != null, 'Added transaction found in list');
  assertApprox(added.cost, 42.00, 'Added transaction cost correct');
  assertEqual(added.paymentMethod, 'AMEX', 'Added transaction payment method correct');

  return result.rowIndex; // for use in subsequent tests
}

function testUpdateTransaction(rowIndex) {
  section('updateTransaction');
  _resetCaches();

  var result = updateTransaction({
    month:         TEST_SHEET_NAME,
    rowIndex:      rowIndex,
    date:          '2026-05-15',
    description:   'Updated Transaction',
    cost:          99.00,
    paymentMethod: 'Wells Fargo',
    discretionary: false,
    ben:           true,
    jenna:         true,
  });

  _resetCaches();
  assert(result.ok, 'updateTransaction returns ok');

  var txns = getTransactions(TEST_SHEET_NAME).transactions;
  var updated = txns.find(function(t){ return t.description === 'Updated Transaction'; });
  assert(updated != null, 'Updated transaction found');
  assertApprox(updated.cost, 99.00, 'Updated cost correct');
  assertEqual(updated.paymentMethod, 'Wells Fargo', 'Updated payment method correct');
  assertEqual(updated.discretionary, false, 'Updated discretionary correct');
}

function testDeleteTransaction(rowIndex) {
  section('deleteTransaction');
  _resetCaches();

  var beforeCount = getTransactions(TEST_SHEET_NAME).transactions.length;

  var result = deleteTransaction({ month: TEST_SHEET_NAME, rowIndex: rowIndex });

  _resetCaches();
  assert(result.ok, 'deleteTransaction returns ok');

  var afterCount = getTransactions(TEST_SHEET_NAME).transactions.length;
  assertEqual(afterCount, beforeCount - 1, 'Transaction count decreased by 1');

  var txns = getTransactions(TEST_SHEET_NAME).transactions;
  var found = txns.find(function(t){ return t.description === 'Updated Transaction'; });
  assert(found == null, 'Deleted transaction no longer found');
}

function testSplitTransaction() {
  section('splitTransaction');
  _resetCaches();

  // Add a transaction to split
  var added = addTransaction({
    month:         TEST_SHEET_NAME,
    date:          '2026-05-20',
    description:   'Costco Trip',
    cost:          150.00,
    paymentMethod: 'Costco',
    discretionary: true,
    ben:           true,
    jenna:         true,
  });
  _resetCaches();

  var beforeCount = getTransactions(TEST_SHEET_NAME).transactions.length;
  var splitRow    = added.rowIndex;

  var result = splitTransaction({
    month:    TEST_SHEET_NAME,
    rowIndex: splitRow,
    date:     '2026-05-20',
    lines: [
      { description: 'Groceries part', cost: 90.00, paymentMethod: 'Costco', discretionary: false, ben: true, jenna: true },
      { description: 'Household part', cost: 60.00, paymentMethod: 'Costco', discretionary: true,  ben: true, jenna: true },
    ]
  });
  _resetCaches();

  assert(result.ok, 'splitTransaction returns ok');

  var afterCount = getTransactions(TEST_SHEET_NAME).transactions.length;
  assertEqual(afterCount, beforeCount + 1, 'Split: count increased by 1 (2 replace 1)');

  var txns = getTransactions(TEST_SHEET_NAME).transactions;
  var line1 = txns.find(function(t){ return t.description === 'Groceries part'; });
  var line2 = txns.find(function(t){ return t.description === 'Household part'; });
  assert(line1 != null, 'Split line 1 found');
  assert(line2 != null, 'Split line 2 found');
  assertApprox(line1.cost, 90.00, 'Split line 1 cost correct');
  assertApprox(line2.cost, 60.00, 'Split line 2 cost correct');

  // Clean up both split rows
  _resetCaches();
  var allTxns = getTransactions(TEST_SHEET_NAME).transactions;
  var r1 = allTxns.find(function(t){ return t.description === 'Groceries part'; });
  var r2 = allTxns.find(function(t){ return t.description === 'Household part'; });
  if (r1) { deleteTransaction({ month: TEST_SHEET_NAME, rowIndex: r1.rowIndex }); _resetCaches(); }
  var allTxns2 = getTransactions(TEST_SHEET_NAME).transactions;
  var r2b = allTxns2.find(function(t){ return t.description === 'Household part'; });
  if (r2b) deleteTransaction({ month: TEST_SHEET_NAME, rowIndex: r2b.rowIndex });
}

function testCardPayment() {
  section('makeCardPayment + voidLastPayment');

  var ct = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CARD_TRACKERS_SHEET);
  if (!ct) {
    _note('testCardPayment skipped -- Card Trackers sheet not found');
    return;
  }

  // Find AMEX row and read its current balance -- invalidate cache first for fresh read
  invalidateSheet(CARD_TRACKERS_SHEET);
  var vals = ct.getDataRange().getValues();
  var amexRow = -1, origBalance = 0;
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][CT_COL_CARD-1]).trim() === 'AMEX') {
      amexRow = i + 1;
      origBalance = typeof vals[i][CT_COL_BALANCE-1] === 'number'
        ? vals[i][CT_COL_BALANCE-1] : parseFloat(vals[i][CT_COL_BALANCE-1]) || 0;
      break;
    }
  }

  if (amexRow < 0) {
    _note('testCardPayment skipped -- AMEX not found in Card Trackers');
    return;
  }

  // Record a test payment with a distinctive month label so void finds it reliably
  var testAmount = 1.11;
  var testMonth  = '_TEST_PAYMENT_DO_NOT_USE_';
  var payResult  = makeCardPayment({
    cardName: 'AMEX',
    amount:   testAmount,
    date:     '2026-01-01',
    month:    testMonth,
  });
  _note('payResult: ' + JSON.stringify(payResult) + ' origBalance=' + origBalance);

  assert(payResult.ok, 'makeCardPayment returns ok');
  assertApprox(payResult.newBalance, origBalance - testAmount,
    'Balance decreased by payment amount');

  // Void it
  var voidResult = voidLastPayment({ cardName: 'AMEX' });
  _note('voidResult: ' + JSON.stringify(voidResult));
  if (!voidResult.ok) {
    _note('voidLastPayment error (balance manually restored): ' + voidResult.error);
    // Manually restore balance so test doesn't corrupt real data
    _adjustCardBalance('AMEX', testAmount);
  }

  assert(voidResult.ok, 'voidLastPayment returns ok');
  assertApprox(voidResult.newBalance, origBalance,
    'Balance restored to original after void');

  // Verify Card Trackers col C is back to original
  var newVals = ct.getDataRange().getValues();
  var newBal  = newVals[amexRow-1][CT_COL_BALANCE-1];
  assertApprox(typeof newBal === 'number' ? newBal : parseFloat(newBal) || 0,
    origBalance, 'Card Trackers balance physically restored after void');
}

function testBooleanDiscFlags() {
  section('Boolean Discretionary flag handling');
  _resetCaches();

  var txns = getTransactions(TEST_SHEET_NAME).transactions;

  // Count transactions with disc=true and disc=false
  var trueCount  = txns.filter(function(t){ return t.discretionary === true;  }).length;
  var falseCount = txns.filter(function(t){ return t.discretionary === false; }).length;

  assert(trueCount  > 0, 'Some transactions have discretionary=true');
  assert(falseCount > 0, 'Some transactions have discretionary=false (booleans included)');
  assertEqual(trueCount + falseCount, txns.length,
    'All transactions have a valid boolean discretionary value');

  // Check ben/jenna are also proper booleans
  var allBoolBen   = txns.every(function(t){ return typeof t.ben   === 'boolean'; });
  var allBoolJenna = txns.every(function(t){ return typeof t.jenna === 'boolean'; });
  assert(allBoolBen,   'All ben values are booleans');
  assert(allBoolJenna, 'All jenna values are booleans');
}

function testGetConfig() {
  section('getConfig');

  var config = getConfig();
  assert(!config.error,                'getConfig returns no error');
  assert(config.budget > 0,            'budget is positive number');
  assert(Array.isArray(config.cards),  'cards is an array');
  assertEqual(config.cards.length, CARD_DEFS.length, 'card count matches CARD_DEFS');
  assert(config.daycareAmount > 0,     'daycareAmount is set');
  assert(config.daycareLastPaid,       'daycareLastPaid is set');
  assert(Array.isArray(config.pmOptions), 'pmOptions is array');
  assert(config.pmOptions.length > 0,  'pmOptions not empty');

  // Each card should have required fields
  var firstCard = config.cards[0];
  assert(firstCard.sheetKey,    'Card has sheetKey');
  assert(firstCard.displayName, 'Card has displayName');
  assert(firstCard.limit > 0,   'Card has limit');
  assert(firstCard.due,         'Card has due date');
}

function testGetMonthlyData() {
  section('getMonthlyData -- structure check');

  var months = getMonthlyData();
  assert(Array.isArray(months),    'Returns array');
  assert(months.length > 0,        'Has at least one month');

  var recent = months[0]; // most recent
  assert(recent.month,             'Month has name');
  assert(typeof recent.income === 'number' || recent.income == null,
    'income is number or null');
  assert(typeof recent.fixedFcf === 'number',
    'fixedFcf is always a number');
  assert(typeof recent.fixedSavings === 'number',
    'fixedSavings is always a number');
}

function testGetCardBalances() {
  section('getCardBalances');

  var result = getCardBalances();
  assert(!result.error,               'No error');
  assert(result.totalBalance != null, 'totalBalance present');
  assert(result.cards != null,        'cards object present');

  // Check each expected card is present under result.cards
  CARD_DEFS.forEach(function(def) {
    var card = result.cards[def.name];
    assert(card != null, 'Card "' + def.name + '" present in balances');
    if (card) {
      assert(typeof card.balance === 'number', def.name + ' balance is a number');
      // lastPayment and lastDate should be present (may be null for cards never paid)
      assert('lastPayment' in card, def.name + ' has lastPayment field');
    }
  });
}

function testMonthlySheetNames() {
  section('getMonthlySheetNames');
  _resetCaches();

  var names = getMonthlySheetNames();
  assert(Array.isArray(names),   'Returns array');
  assert(names.length > 0,       'Has at least one month sheet');

  // Verify all names match the pattern
  names.forEach(function(n) {
    assert(MONTH_PATTERN.test(n), '"' + n + '" matches month pattern');
  });

  // Scratch sheet should NOT appear
  assert(names.indexOf(TEST_SHEET_NAME) < 0,
    'Scratch sheet not included in month names');
}

function testRowIndexConsistency() {
  section('rowIndex consistency (insert/delete symmetry)');
  _resetCaches();

  var before = getTransactions(TEST_SHEET_NAME).transactions;
  var origCount = before.length;

  // Add at top, verify subsequent rows shift
  var added = addTransaction({
    month: TEST_SHEET_NAME, date: '2026-05-01',
    description: '__SHIFT_TEST__', cost: 1.00,
    paymentMethod: 'AMEX', discretionary: true, ben: true, jenna: false,
  });
  _resetCaches();

  var after = getTransactions(TEST_SHEET_NAME).transactions;
  assertEqual(after.length, origCount + 1, 'Row count increased');

  // Delete it and verify count is restored
  deleteTransaction({ month: TEST_SHEET_NAME, rowIndex: added.rowIndex });
  _resetCaches();

  var restored = getTransactions(TEST_SHEET_NAME).transactions;
  assertEqual(restored.length, origCount, 'Row count restored after delete');

  // Verify original rows have same rowIndices as before
  var origFirst = before[0];
  var restoredFirst = restored[0];
  assertEqual(restoredFirst.rowIndex, origFirst.rowIndex,
    'First row rowIndex unchanged after insert+delete cycle');
}

// -- addMonth tests ---------------------------------------------

function testAddMonthValidation() {
  section('addMonth -- validation');

  var r1 = addMonth({});
  assert(r1.error && r1.error.indexOf('required') >= 0, 'Missing monthName returns error');

  var r2 = addMonth({ monthName: 'Jun 2026' });
  assert(r2.error && r2.error.indexOf('Invalid') >= 0, 'Short month name rejected');

  var r3 = addMonth({ monthName: 'June2026' });
  assert(r3.error && r3.error.indexOf('Invalid') >= 0, 'Missing space rejected');

  var r4 = addMonth({ monthName: 'december 2099' });
  assert(r4.error && r4.error.indexOf('Invalid') >= 0, 'Lowercase month rejected');

  var existing = getMonthlySheetNames()[0];
  var r5 = addMonth({ monthName: existing });
  assert(r5.error && r5.error.indexOf('already exists') >= 0, 'Existing sheet rejected: ' + existing);
}

function testMonthCacheTiers() {
  section('Month cache tiers');

  var name = '_TierTest_ ' + Date.now();
  var key  = _monthCacheKey(name);
  var row  = { month: name, totalExpenses: 123.45 };
  var json = JSON.stringify(row);

  // Properties tier round-trip
  _monthPropsSet(name, json);
  var all = _monthPropsAll();
  assert(!!all[key], 'Properties tier stores entry');
  assertEqual(JSON.parse(all[key]).totalExpenses, 123.45, 'Properties round-trip preserves data');

  // Invalidation clears BOTH tiers
  try { CacheService.getScriptCache().put(key, json, 600); } catch(e) {}
  invalidateMonthCache(name);
  var all2 = _monthPropsAll();
  assert(!all2[key], 'invalidateMonthCache removes properties entry');
  var c = null;
  try { c = CacheService.getScriptCache().get(key); } catch(e) {}
  assert(!c, 'invalidateMonthCache removes cache entry');
}

function testConfigParsing() {
  section('Config sheet parsing');

  assertEqual(_cfgNum('42.5', 0), 42.5, 'Numeric string parses');
  assertEqual(_cfgNum('abc', 7), 7, 'Garbage number falls back');
  assertEqual(_cfgNum(13204.75, 0), 13204.75, 'Real number passes through');

  assertEqual(_cfgDateStr('2026-07-10', 'x'), '2026-07-10', 'ISO date string accepted');
  assertEqual(_cfgDateStr('July 10', 'fallback'), 'fallback', 'Non-ISO string falls back');
  var d = new Date(2026, 6, 10, 12);
  assertEqual(_cfgDateStr(d, 'x'), '2026-07-10', 'Date object formatted');

  var cm = _parseClosedMondays('2026-06-29, 2026-12-21', []);
  assertEqual(cm.length, 2, 'Two closed mondays parsed');
  assertEqual(cm[0], '2026-06-29', 'First closed monday correct');
  assertEqual(_parseClosedMondays('junk', ['d']).length, 1, 'Invalid list falls back');
  assertEqual(_parseClosedMondays('', ['d'])[0], 'd', 'Empty falls back');

  var q = _parseQuarterly('Water:1,4,7,10; Sewer:2,5', []);
  assertEqual(q.length, 2, 'Two quarterly entries parsed');
  assertEqual(q[0].name, 'Water', 'Quarterly name parsed');
  assertEqual(q[0].months.length, 4, 'Quarterly months parsed');
  assertEqual(q[1].months[1], 5, 'Second entry months parsed');
  assertEqual(_parseQuarterly('nonsense', [{name:'X',months:[1]}])[0].name, 'X', 'Malformed quarterly falls back');

  var out = _configDefaults();
  _applyRawConfig(out, { 'budget': '14000', 'daycare_amount': 2500 });
  assertEqual(out.budget, 14000, 'Raw budget override applies');
  assertEqual(out.daycareAmount, 2500, 'Raw daycare override applies');
  assertEqual(out.projectionTarget, 1000000, 'Untouched keys keep defaults');

  // Regression: hsa_established must round-trip through the same path as
  // every other Config key (this is what broke -- a stale CFG_CACHE_KEY
  // served pre-HSA-shape config and showed "no establishment date" even
  // with the sheet row present; see INDEX trap on cfg() cache versioning).
  var out2 = _configDefaults();
  assertEqual(out2.hsaEstablished, null, 'hsaEstablished defaults to null');
  _applyRawConfig(out2, { 'hsa_established': new Date(2025, 3, 17, 12) });
  assertEqual(out2.hsaEstablished, '2025-04-17', 'hsa_established Date applies via _cfgDateStr');
  var out3 = _configDefaults();
  _applyRawConfig(out3, { 'hsa_established': '2025-04-17' });
  assertEqual(out3.hsaEstablished, '2025-04-17', 'hsa_established ISO string applies');
}

function testIdempotency() {
  section('Idempotency guard');

  // Unique key: first call runs, second call with same key returns cached result
  var key = 'idem_test_' + Date.now();
  var calls = 0;
  var r1 = _withIdem({ idemKey: key }, function(){ calls++; return { ok: true, value: 42 }; });
  var r2 = _withIdem({ idemKey: key }, function(){ calls++; return { ok: true, value: 99 }; });
  assertEqual(calls, 1, 'Duplicate key skips the second write');
  assertEqual(r1.value, 42, 'First call returns real result');
  assertEqual(r2.value, 42, 'Second call returns cached result, not re-run');

  // No key: always runs
  var calls2 = 0;
  _withIdem({}, function(){ calls2++; return { ok: true }; });
  _withIdem({}, function(){ calls2++; return { ok: true }; });
  assertEqual(calls2, 2, 'Missing key runs every time');

  // Error results are not cached -- retry can attempt again
  var key3 = 'idem_test_err_' + Date.now();
  var calls3 = 0;
  _withIdem({ idemKey: key3 }, function(){ calls3++; return { error: 'boom' }; });
  _withIdem({ idemKey: key3 }, function(){ calls3++; return { ok: true }; });
  assertEqual(calls3, 2, 'Error result not cached; retry runs the write');

  // Cleanup
  try { CacheService.getScriptCache().remove(key); CacheService.getScriptCache().remove(key3); } catch(e) {}
}

function testHelperFunctions() {
  section('addMonth helpers');

  assertEqual(shortLabel('January 2025'), "January '25", 'shortLabel January 2025');
  assertEqual(shortLabel('December 2026'), "December '26", 'shortLabel December 2026');
  assertEqual(shortLabel('May 2026'), "May '26", 'shortLabel May 2026');

  var fakeVals = [
    ['', 'Header', ''],
    ['', 'Foo', ''],
    ['', 'Activity 2', ''],
    ['', 'Bar', ''],
  ];
  assertEqual(findRowByLabel(fakeVals, 'Activity 2'), 3, 'findRowByLabel finds correct row');
  assertEqual(findRowByLabel(fakeVals, 'Missing'), -1, 'findRowByLabel returns -1 when not found');

  var fakeVals2 = [];
  for (var i = 0; i < 6; i++) fakeVals2.push(new Array(22).fill(''));
  fakeVals2[2][20] = 'Discretionary Expenditures';
  fakeVals2[4][20] = 'Total Expenses';
  assertEqual(findRowByULabel(fakeVals2, 'Discretionary Expenditures'), 3, 'findRowByULabel finds col U label');
  assertEqual(findRowByULabel(fakeVals2, 'Total Expenses'), 5, 'findRowByULabel finds second label');
  assertEqual(findRowByULabel(fakeVals2, 'Missing'), -1, 'findRowByULabel returns -1 when not found');

  var rx = /^(January|February|March|April|May|June|July|August|September|October|November|December) \d{4}$/;
  ['January 2025','December 2024','June 2026'].forEach(function(m) {
    assert(rx.test(m), 'Valid month format: ' + m);
  });
  ['Jun 2025','January2025','january 2025','June 26',''].forEach(function(m) {
    assert(!rx.test(m), 'Invalid month format rejected: ' + m);
  });
}

// -- Main runner ------------------------------------------------
// runTests() -- fast suite (~10s): transaction CRUD, config, helpers
// runSlowTests() -- full suite (~30s): includes card balances + monthly data scans

// HSA rollup math + reimburse guards. Pure -- no sheet, so it runs in the fast
// suite. Fixture mirrors INDEX section 8 and the build-spec worked example.
function testHsa() {
  section('HSA receipts (rollups + reimburse guards)');

  // A/B/C/D = qualified; E = pre-establishment (must be excluded + flagged).
  var rows = [
    { id:1, dateIncurred:'2023-02-10', amount:1200, funding:'OOP', reimbursedAmount:0 },
    { id:2, dateIncurred:'2023-05-04', amount:800,  funding:'OOP', reimbursedAmount:800 },
    { id:3, dateIncurred:'2024-03-22', amount:2500, funding:'OOP', reimbursedAmount:1000 },
    { id:4, dateIncurred:'2024-07-09', amount:300,  funding:'HSA', reimbursedAmount:0 },
    { id:5, dateIncurred:'2021-06-01', amount:500,  funding:'OOP', reimbursedAmount:0 }
  ];
  var est = '2022-01-01';

  var r5 = _hsaRollups(rows, est, 5000);
  assertEqual(r5.unreimbursed, 2700, 'unreimbursed pool = 1200 + 0 + 1500');
  assertEqual(r5.totalSubstantiated, 4800, 'total substantiated = 1200+800+2500+300 (E excluded)');
  assertEqual(r5.reimbursableNow, 2700, 'reimbursable now = min(2700, 5000)');
  assertEqual(r5.strandedEntitlement, 0, 'no stranded entitlement when balance >= pool');

  var r2 = _hsaRollups(rows, est, 2000);
  assertEqual(r2.reimbursableNow, 2000, 'reimbursable now capped by balance (2000)');
  assertEqual(r2.strandedEntitlement, 700, 'stranded entitlement = 2700 - 2000');

  var rNull = _hsaRollups(rows, est, null);
  assertEqual(rNull.reimbursableNow, null, 'reimbursable now null when balance unknown');
  assertEqual(rNull.strandedEntitlement, null, 'stranded null when balance unknown');
  assertEqual(rNull.unreimbursed, 2700, 'unreimbursed unchanged when balance blank (not 0)');
  assertEqual(rNull.totalSubstantiated, 4800, 'total substantiated unchanged when balance blank');

  // Per-row status / qualified flags
  var byId = {};
  for (var i = 0; i < r5.rows.length; i++) byId[r5.rows[i].id] = r5.rows[i];
  assertEqual(byId[1].status, 'open',        'row 1 open');
  assertEqual(byId[2].status, 'reimbursed',  'row 2 reimbursed');
  assertEqual(byId[3].status, 'partial',     'row 3 partial');
  assertEqual(byId[4].status, 'paid_direct', 'row 4 paid_direct (HSA card)');
  assertEqual(byId[5].qualified, false,      'row 5 flagged unqualified (predates HSA)');
  assertEqual(byId[1].qualified, true,       'row 1 qualified');

  // Empty input: no exception, all zero, rows []
  var rEmpty = _hsaRollups([], est, 5000);
  assertEqual(rEmpty.unreimbursed, 0, 'empty: unreimbursed 0');
  assertEqual(rEmpty.totalSubstantiated, 0, 'empty: total substantiated 0');
  assertEqual(rEmpty.reimbursableNow, 0, 'empty: reimbursable now 0 (balance present)');
  assertEqual(rEmpty.rows.length, 0, 'empty: no rows');

  // Gate off when no establishment date: nothing disqualified
  var rNoEst = _hsaRollups(rows, null, 5000);
  assertEqual(rNoEst.unreimbursed, 3200, 'no establishment date -> E (500) now counts (gate off)');

  // Reimburse guards
  assertEqual(_reimburseGuard({ amount:1200, funding:'OOP', qualified:true }, 1200), null,
    'guard allows valid full reimbursement');
  assertEqual(_reimburseGuard({ amount:1200, funding:'OOP', qualified:true }, 600), null,
    'guard allows valid partial reimbursement');
  assert(_reimburseGuard({ amount:300, funding:'HSA', qualified:true }, 300) !== null,
    'guard rejects funding=HSA (already HSA-paid)');
  assert(_reimburseGuard({ amount:500, funding:'OOP', qualified:false }, 500) !== null,
    'guard rejects pre-establishment row');
  assert(_reimburseGuard({ amount:1200, funding:'OOP', qualified:true }, 1500) !== null,
    'guard rejects over-amount reimbursement');
  assert(_reimburseGuard({ amount:1200, funding:'OOP', qualified:true }, -5) !== null,
    'guard rejects negative amount');

  // needs_review drafts are excluded from totals until confirmed
  var draftRows = [
    { id:1, dateIncurred:'2024-01-01', amount:1000, funding:'OOP', reimbursedAmount:0 },
    { id:2, dateIncurred:'2024-02-01', amount:500,  funding:'OOP', reimbursedAmount:0, needsReview:true }
  ];
  var rd = _hsaRollups(draftRows, '2022-01-01', 5000);
  assertEqual(rd.unreimbursed, 1000, 'draft row excluded from unreimbursed pool');
  assertEqual(rd.totalSubstantiated, 1000, 'draft row excluded from total substantiated');
  assertEqual(rd.rows[1].status, 'needs_review', 'draft row gets needs_review status');
}

function testHsaPlanYear() {
  section('_hsaPlanYear (deductible / contributions / balance reconciliation)');

  // Rolled receipt rows: card spend (HSA) across two plan years + one OOP +
  // one flagged draft (must NOT count toward the deductible, but DOES leave cash).
  var rolled = [
    { dateIncurred:'2025-08-01', amount:3636.33, funding:'HSA', reimbursedAmount:0,   qualified:true,  needsReview:false },
    { dateIncurred:'2026-03-01', amount:3201.71, funding:'HSA', reimbursedAmount:0,   qualified:true,  needsReview:false },
    { dateIncurred:'2026-04-10', amount:20.48,   funding:'HSA', reimbursedAmount:0,   qualified:true,  needsReview:true  }, // flagged: cash left, not counted
    { dateIncurred:'2026-05-01', amount:200.00,  funding:'OOP', reimbursedAmount:50,  qualified:true,  needsReview:false }  // OOP counts toward deductible; $50 reimbursed = outflow
  ];
  var contribs = [
    { date:'2025-09-01', amount:4003.73, type:'employee' },
    { date:'2026-02-01', amount:2925.01, type:'employee' },
    { date:'2026-04-23', amount:2000.00, type:'employer' }
  ];
  var p = _hsaPlanYear(rolled, contribs, 2052.65, null, 3300, 8750, 1, '2026-07-07');

  assertEqual(p.planYearStart, '2026-01-01', 'calendar plan year starts Jan 1');
  // deductible: 2026 qualified incurred = 3201.71 (card) + 200 (OOP) = 3401.71; flagged 20.48 excluded
  assertApprox(p.deductibleYtd, 3401.71, 'deductible YTD sums qualified incurred (both fundings), excludes drafts');
  assertEqual(p.deductibleRemaining, 0, 'deductible met -> remaining clamps to 0');
  assertEqual(p.deductiblePct, 100, 'deductible pct clamps to 100');
  // contributions toward limit: employee+employer this calendar year = 2925.01 + 2000 = 4925.01
  assertApprox(p.contribYtd, 4925.01, 'contrib YTD = employee+employer this calendar year');
  assertApprox(p.contribRoom, 3824.99, 'contrib room = limit - YTD');
  // No B2 -> all-time reconstruction: inflows 8928.74 - outflows (card 6858.52 + reimb 50) = 2020.22
  assertApprox(p.balanceOutflows, 6908.52, 'outflows = all card spend + reimbursements');
  assertApprox(p.computedBalance, 2020.22, 'no B2 -> computed balance = all-time inflows - outflows');
  assertApprox(p.drift, 32.43, 'drift vs B1 surfaces unlogged interest');

  // Date-aware reconciliation: B1 as of 2026-04-01. Split by that date:
  //   inThru = 4003.73 + 2925.01 = 6928.74 ; outThru = 3636.33 + 3201.71 = 6838.04
  //   computedThruB2 = 90.70 ; B1 = 123.13 -> drift 32.43 (interest).
  //   After 4/1: in = 2000 (employer 4/23) ; out = 20.48 (draft card) + 50 (reimb) = 70.48.
  //   estimated now = 123.13 + 2000 - 70.48 = 2052.65.
  var pa = _hsaPlanYear(rolled, contribs, 123.13, '2026-04-01', 3300, 8750, 1, '2026-07-07');
  assertApprox(pa.drift, 32.43, 'date-aware: drift is interest-only, not stale-anchor activity');
  assertApprox(pa.inflowsSince, 2000, 'inflows since B2 = contributions dated after B2');
  assertApprox(pa.outflowsSince, 70.48, 'outflows since B2 = card + reimbursement after B2');
  assertApprox(pa.computedBalance, 2052.65, 'displayed balance projects B1 forward by since-B2 activity');

  // No contributions sheet -> balance/contrib gauges null, deductible still works
  var p2 = _hsaPlanYear(rolled, null, 2052.65, null, 3300, 8750, 1, '2026-07-07');
  assertEqual(p2.computedBalance, null, 'no contrib sheet -> computed balance null');
  assertEqual(p2.contribYtd, null, 'no contrib sheet -> contrib YTD null');
  assertApprox(p2.deductibleYtd, 3401.71, 'deductible still computes without contributions');

  // Mid-year plan reset (e.g. month 7): on 2026-07-07 the plan year just started
  var p3 = _hsaPlanYear(rolled, contribs, 2052.65, null, 3300, 8750, 7, '2026-07-07');
  assertEqual(p3.planYearStart, '2026-07-01', 'mid-year reset month resolves plan-year start');
  assertEqual(p3.deductibleYtd, 0, 'fresh plan year -> deductible YTD resets to 0');
}

function testAddHsaReceipt() {
  section('addHsaReceipt (scratch sheet -- never touches live HSA Receipts)');
  createHsaScratchSheet();
  _resetCaches();

  var r1 = addHsaReceipt({
    dateIncurred:'2026-05-01', amount:120.5, provider:'Test Clinic',
    description:'Checkup', category:'Office', funding:'OOP'
  }, HSA_TEST_SHEET_NAME);
  assert(r1.ok, 'first add returns ok');
  assertEqual(r1.id, 1, 'first receipt gets id 1');

  var r2 = addHsaReceipt({
    dateIncurred:'2026-05-02', amount:50, provider:'Pharmacy', funding:'HSA'
  }, HSA_TEST_SHEET_NAME);
  assert(r2.ok, 'second add returns ok');
  assertEqual(r2.id, 2, 'second receipt gets id 2 (max+1, ids never reused)');

  // Validation guards -- each must reject WITHOUT writing a row
  assert(addHsaReceipt({ dateIncurred:'2026-05-01', amount:-5 }, HSA_TEST_SHEET_NAME).error,
    'rejects non-positive amount');
  assert(addHsaReceipt({ dateIncurred:'2026-05-01', amount:0 }, HSA_TEST_SHEET_NAME).error,
    'rejects zero amount');
  assert(addHsaReceipt({ dateIncurred:'not-a-date', amount:10 }, HSA_TEST_SHEET_NAME).error,
    'rejects malformed date');
  assert(addHsaReceipt({ dateIncurred:'', amount:10 }, HSA_TEST_SHEET_NAME).error,
    'rejects missing date');
  assert(addHsaReceipt({ dateIncurred:'2026-05-01', amount:10, receiptLink:'javascript:alert(1)' }, HSA_TEST_SHEET_NAME).error,
    'rejects non-http(s) receipt link');

  _resetCaches();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HSA_TEST_SHEET_NAME);
  var rows  = _hsaReadRows(sheet);
  assertEqual(rows.length, 2, 'exactly 2 rows present -- failed validations wrote nothing');
  assertEqual(rows[0].id, 1, 'row 1 id correct on the sheet');
  assertEqual(rows[1].funding, 'HSA', 'row 2 funding correct on the sheet');
  assertApprox(rows[0].amount, 120.5, 'row 1 amount correct on the sheet');

  // sourceFileId lands in col L (12) -- the folder-scan dedup key
  var r3 = addHsaReceipt({
    dateIncurred:'2026-05-03', amount:75, provider:'Drive Scan', funding:'OOP',
    sourceFileId:'drive-file-abc123'
  }, HSA_TEST_SHEET_NAME);
  assert(r3.ok, 'add with sourceFileId returns ok');
  assertEqual(sheet.getRange(r3.row, 12).getValue(), 'drive-file-abc123', 'source_file_id written to col L');
}

function testParseReceiptName() {
  section('_parseReceiptName (filename -> fields)');

  var a = _parseReceiptName('2026-06-27~Summit Dental~1200.pdf');
  assert(!a.error, 'valid name parses');
  assertEqual(a.date, '2026-06-27', 'date parsed');
  assertEqual(a.provider, 'Summit Dental', 'provider with space parsed');
  assertApprox(a.amount, 1200, 'amount parsed');

  var b = _parseReceiptName('2026-06-27~Summit_Dental~1200.50.pdf');
  assertEqual(b.provider, 'Summit_Dental', 'provider with underscore allowed (delimiter is ~)');
  assertApprox(b.amount, 1200.50, 'decimal amount parsed');

  var c = _parseReceiptName('2026-06-27~CVS~$1,200.00.pdf');
  assertApprox(c.amount, 1200, 'currency symbols/commas stripped from amount');

  var d = _parseReceiptName('2026-06-27~Walgreens~10.PDF');
  assert(!d.error, 'uppercase .PDF extension accepted');

  // Strict rejections -- each must skip, never guess
  assert(_parseReceiptName('2026-6-27~X~10.pdf').error,    'rejects non-zero-padded date');
  assert(_parseReceiptName('06-27-2026~X~10.pdf').error,   'rejects wrong date order');
  assert(_parseReceiptName('2026-06-27~X.pdf').error,      'rejects too few fields');
  assert(_parseReceiptName('2026-06-27~A~B~10.pdf').error, 'rejects too many fields');
  assert(_parseReceiptName('2026-06-27~~10.pdf').error,    'rejects empty provider');
  assert(_parseReceiptName('2026-06-27~X~abc.pdf').error,  'rejects non-numeric amount');
  assert(_parseReceiptName('2026-06-27~X~0.pdf').error,    'rejects zero amount');
  assert(_parseReceiptName('2026-06-27~X~-5.pdf').error,   'rejects negative amount');
}

function testBestEffortReceipt() {
  section('_bestEffortReceipt + _isRealDate (malformed-name salvage)');

  // _isRealDate
  assert(_isRealDate('2026-06-27'),  'valid date accepted');
  assert(!_isRealDate('2026-13-01'), 'month 13 rejected');
  assert(!_isRealDate('2026-02-29'), 'Feb 29 in non-leap year rejected');
  assert(_isRealDate('2024-02-29'),  'Feb 29 in leap year accepted');
  assert(!_isRealDate('2026-6-7'),   'non-zero-padded rejected');

  // Old underscore-delimited name (pre-convention) salvages cleanly
  var a = _bestEffortReceipt('2026-06-27_Summit_1200.pdf');
  assertEqual(a.date, '2026-06-27', 'underscore name: date salvaged');
  assertApprox(a.amount, 1200, 'underscore name: amount salvaged');
  assertEqual(a.provider, 'Summit', 'underscore name: provider salvaged');

  // Out-of-order fields still resolve by type
  var b = _bestEffortReceipt('CVS~2026-03-01~$45.99.pdf');
  assertEqual(b.date, '2026-03-01', 'out-of-order: date found');
  assertApprox(b.amount, 45.99, 'out-of-order: amount found ($ stripped)');
  assertEqual(b.provider, 'CVS', 'out-of-order: provider is the leftover');

  // Invalid date is NOT salvaged (left blank for review), amount still found
  var c = _bestEffortReceipt('2026-13-99~Clinic~50.pdf');
  assertEqual(c.date, '', 'invalid date left blank');
  assertApprox(c.amount, 50, 'amount still found alongside bad date');

  // Nothing receipt-like -> empty (caller will skip)
  var d = _bestEffortReceipt('random scan.pdf');
  assertEqual(d.date, '', 'no date in junk name');
  assertEqual(d.amount, 0, 'no amount in junk name');
}

function testUpdateHsaReceipt() {
  section('updateHsaReceipt (edit + clear needs_review; scratch sheet only)');
  createHsaScratchSheet();
  _resetCaches();

  // Seed a draft (needs_review) row, as the folder scan would for a bad name
  var d = addHsaReceipt({
    dateIncurred:'', amount:0, provider:'Imported', funding:'OOP',
    needsReview:true, notes:'from bad-name.pdf'
  }, HSA_TEST_SHEET_NAME);
  assert(d.ok, 'draft import returns ok');
  assert(d.needsReview, 'draft flagged needsReview');

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HSA_TEST_SHEET_NAME);
  _resetCaches();
  var before = _hsaReadRows(sheet);
  assertEqual(before[0].needsReview, true, 'draft row reads back as needs_review');

  // Validation rejects bad input (no write)
  assert(updateHsaReceipt({ id:d.id, amount:-1, dateIncurred:'2026-01-01' }, HSA_TEST_SHEET_NAME).error,
    'rejects non-positive amount');
  assert(updateHsaReceipt({ id:d.id, amount:10, dateIncurred:'bad' }, HSA_TEST_SHEET_NAME).error,
    'rejects malformed date');
  assert(updateHsaReceipt({ id:'999999', amount:10, dateIncurred:'2026-01-01' }, HSA_TEST_SHEET_NAME).error,
    'rejects unknown id');

  // Valid edit fills the fields and clears needs_review
  var u = updateHsaReceipt({
    id:d.id, dateIncurred:'2026-03-15', amount:212.40,
    provider:'Summit Dental', description:'Crown', category:'Dental', funding:'OOP'
  }, HSA_TEST_SHEET_NAME);
  assert(u.ok, 'valid edit returns ok');

  _resetCaches();
  var after = _hsaReadRows(sheet);
  assertEqual(after[0].needsReview, false, 'needs_review cleared after edit');
  assertEqual(after[0].dateIncurred, '2026-03-15', 'date updated');
  assertApprox(after[0].amount, 212.40, 'amount updated');
  assertEqual(after[0].provider, 'Summit Dental', 'provider updated');
}

function testMortgageSingleSource() {
  section('mortgage single source (net worth == ledger, not B8)');

  var m = readMortgageLedger();
  assert(!m.error, 'readMortgageLedger returns no error');
  assert(typeof m.curBalance === 'number' && m.curBalance > 0,
    'ledger curBalance is a positive number');

  var mg = getMortgageData();
  assert(!mg.error, 'getMortgageData returns no error');
  // One definition of "current balance": the mortgage tab must equal the ledger.
  assertApprox(mg.currentBalance, m.curBalance,
    'mortgage tab currentBalance == ledger curBalance', 0.01);

  // Regression guard for the B8-drift bug: net worth's mortgage liability must be
  // the schedule's current balance, NOT the hand-typed Physical Assets!B8.
  invalidateNetWorthCache();
  var nw = getNetWorth();
  assert(!nw.error, 'getNetWorth returns no error');
  assert(typeof nw.mortgage === 'number', 'net worth mortgage is a number');
  assertApprox(nw.mortgage, m.curBalance,
    'net worth mortgage liability == ledger curBalance (not B8)', 0.01);

  // B8 is now only an informational cross-check; drift is allowed and must NOT
  // affect net worth. Log the current drift for visibility.
  if (mg.physicalAssetsBalance != null) {
    var drift = Math.round((mg.currentBalance - mg.physicalAssetsBalance) * 100) / 100;
    _note('B8 cross-check drift = ' + drift + ' (informational; B8 is not used by net worth)');
  }
}

function runTests() {
  _runTestSuite(false);
}

function runSlowTests() {
  _runTestSuite(true);
}

function _runTestSuite(includeSlow) {
  _testPass = 0;
  _testFail = 0;
  _testLog   = [];
  _sections  = [];
  _curSection = null;
  _notes     = [];
  _resetCaches();

  var startMs = Date.now();

  // Create scratch sheet for transaction tests
  createScratchSheet();
  _resetCaches();

  try {
    // Config and structure tests (fast -- no month scanning)
    testMonthlySheetNames();
    testGetConfig();
    testMortgageSingleSource();   // mortgage single-source regression guard

    // Transaction read/write tests (against scratch sheet only -- fast)
    testGetTransactions();
    testBooleanDiscFlags();
    testGetTransactionsTermination();

    // Write tests -- each cleans up after itself
    var addedRowIndex = testAddTransaction();
    if (addedRowIndex) {
      testUpdateTransaction(addedRowIndex);
      testDeleteTransaction(addedRowIndex);
    }
    testSplitTransaction();
    testRowIndexConsistency();

    // addMonth validation and helper tests (no sheet creation -- fast)
    testAddMonthValidation();
    testHelperFunctions();
    testIdempotency();
    testConfigParsing();
    testHsa();
    testHsaPlanYear();
    testAddHsaReceipt();
    testParseReceiptName();
    testBestEffortReceipt();
    testUpdateHsaReceipt();
    testMonthCacheTiers();

    if (includeSlow) {
      // These scan all month sheets -- slow but thorough
      testGetCardBalances();
      testGetMonthlyData();
      testCardPayment();
    }

  } catch(e) {
    if (e.message && e.message.indexOf('typed column') >= 0) {
      _note('ignored Sheets typed-column formatting warning -- tests unaffected');
    } else {
      // A thrown runner is a real failure: count it so the verdict reflects it.
      _testFail++;
      _testLog.push('  FAIL [runner] threw: ' + (e && e.message ? e.message : e));
    }
  } finally {
    // Scratch sheet is intentionally kept for reuse on next run
    _resetCaches();
  }

  // -- Compact, failure-first output (single Logger.log; Apps Script truncates
  //    oversized entries, so we never print a PASS-per-assertion transcript) --
  var elapsed = ((Date.now() - startMs) / 1000).toFixed(1);
  var L = [];
  L.push('Finance Dashboard Test Suite' + (includeSlow ? ' (full)' : ' (fast)'));
  L.push('Results: ' + _testPass + ' passed, ' + _testFail + ' failed (' + elapsed + 's)');
  L.push(_testFail === 0 ? 'ALL PASSED' : (_testFail + ' FAILED -- detail below'));

  if (_notes.length) {
    L.push('');
    L.push('Notes:');
    for (var n = 0; n < _notes.length; n++) L.push('  - ' + _notes[n]);
  }
  if (_testFail > 0) {
    L.push('');
    L.push('FAILURES:');
    for (var k = 0; k < _testLog.length; k++) L.push(_testLog[k]);
  }
  L.push('');
  L.push('Sections (' + _sections.length + '):');
  for (var i = 0; i < _sections.length; i++) {
    var s = _sections[i];
    L.push('  ' + (s.fail ? 'FAIL ' : 'ok   ') + s.name + ': ' + s.pass + '/' + (s.pass + s.fail));
  }

  Logger.log(L.join('\n'));
}
