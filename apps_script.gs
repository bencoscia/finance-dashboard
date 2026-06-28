// ============================================================
// See INDEX.md for the file map, invariants, traps, and change protocol. Keep it updated.
// Finance Dashboard -- Google Apps Script v3
// Paste into Tools > Apps Script, deploy as Web App:
//   Execute as: Me  |  Who has access: Anyone
// After any edit, create a NEW deployment for changes to apply.
// ============================================================

// -- Per-request caches (reset at start of every doGet/doPost) --
var _sheetCache  = {};   // name -> values[][]
var _monthsCache = null; // cached sheet name list

function _resetCaches() {
  _sheetCache  = {};
  _monthsCache = null;
}

// Single bulk read per sheet per request -- the single biggest perf win.
// Every getDataRange().getValues() call is replaced by readSheet(name).
function readSheet(nameOrSheet) {
  var name = (typeof nameOrSheet === 'string') ? nameOrSheet : nameOrSheet.getName();
  if (_sheetCache[name]) return _sheetCache[name];
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = (typeof nameOrSheet === 'string') ? ss.getSheetByName(name) : nameOrSheet;
  if (!sheet) return null;
  return _sheetCache[name] = sheet.getDataRange().getValues();
}

// Invalidate a single sheet after a write so subsequent reads in the same
// request see fresh data (important for multi-step operations like addMonth).
function invalidateSheet(name) {
  delete _sheetCache[name];
}

// -- Month sheet discovery ------------------------------------
var MONTH_PATTERN = /^(January|February|March|April|May|June|July|August|September|October|November|December) \d{4}$/;

function getMonthlySheetNames() {
  if (_monthsCache) return _monthsCache;
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var names = [];
  ss.getSheets().forEach(function(s) {
    if (MONTH_PATTERN.test(s.getName())) names.push(s.getName());
  });
  return _monthsCache = names;
}

var PORTFOLIO_SHEET = 'Portfolio Management';
var BUDGET = 13204.75208;
// The month whose ending balance is used as the seed for card balance calculations.
// When you archive and reset, update this to the last archived month.
var SEED_MONTH_CONFIG = 'December 2024';


// -- DISPATCH -------------------------------------------------
// Single entry point for google.script.run calls from HtmlService.
// Handles both read (GET-style) and write (POST-style) actions.
// -- Idempotency guard -----------------------------------------
// Client sends a unique idemKey with create operations. If the same key
// arrives again (an automatic retry after a dropped response), return the
// original result instead of performing the write a second time.
// Keys live in CacheService for 6 hours -- far longer than any retry window.
function _withIdem(p, fn) {
  var key = p && p.idemKey ? String(p.idemKey).slice(0, 240) : '';
  if (!key) return fn();
  var cache = CacheService.getScriptCache();
  try {
    var hit = cache.get(key);
    if (hit) return JSON.parse(hit); // duplicate request -- skip the write
  } catch (e) {}
  var result = fn();
  // Only record the key on success; a failed write should be retryable
  if (result && !result.error) {
    try { cache.put(key, JSON.stringify(result), 21600); } catch (e) {}
  }
  return result;
}

function dispatch(payload) {
  _resetCaches();
  var action = payload.action || '';
  try {
    // Read actions
    if      (action === 'all')             return getAllData();
    else if (action === 'quick')           return getQuickData();
    else if (action === 'monthly')         return getMonthlyData();
    else if (action === 'networth')        return getNetWorth();
    else if (action === 'transactions')    return getTransactions(payload.month);
    else if (action === 'fixed')           return getFixedExpenses(payload.month);
    else if (action === 'variableEntries') return getVariableEntries(payload.month);
    else if (action === 'accounts')        return getAccountsData();
    else if (action === 'cardTotals')      return getCardTotals();
    else if (action === 'cardBalances')    return getCardBalances();
    else if (action === 'loans')           return getLoanData();
    else if (action === 'mortgage')        return getMortgageData();
    else if (action === 'cardTxns')        return getTransactionsByCard(payload.pm, payload.limit);
    else if (action === 'config')          return getConfig();
    else if (action === 'hsa')             return getHsa();
    // Write actions
    else if (action === 'updateTransaction')  { var r = updateTransaction(payload);  if (payload.month) invalidateMonthCache(payload.month); invalidateCardBalanceCache(); return r; }
    else if (action === 'addTransaction')     { var r = _withIdem(payload, function(){ return addTransaction(payload); });     if (payload.month) invalidateMonthCache(payload.month); invalidateCardBalanceCache(); return r; }
    else if (action === 'deleteTransaction')  { var r = deleteTransaction(payload);  if (payload.month) invalidateMonthCache(payload.month); invalidateCardBalanceCache(); return r; }
    else if (action === 'splitTransaction')   { var r = splitTransaction(payload);   if (payload.month) invalidateMonthCache(payload.month); invalidateCardBalanceCache(); return r; }
    else if (action === 'setFixedPaid')       { var r = setFixedPaid(payload);       if (payload.month) invalidateMonthCache(payload.month); return r; }
    else if (action === 'updateFixedCost')    { var r = updateFixedCost(payload);    if (payload.month) invalidateMonthCache(payload.month); return r; }
    else if (action === 'updateFixedName')    { var r = updateFixedName(payload);    if (payload.month) invalidateMonthCache(payload.month); return r; }
    else if (action === 'updateLoanBalance')  return _withIdem(payload, function(){ return updateLoanBalance(payload); }); // invalidates Loans + net worth internally
    else if (action === 'logCreditSnapshot')  return logCreditSnapshot(payload);
    else if (action === 'addMonth')           return addMonth(payload);
    else if (action === 'makeCardPayment')    { var r = _withIdem(payload, function(){ return makeCardPayment(payload); });    invalidateCardBalanceCache(); return r; }
    else if (action === 'voidLastPayment')    { var r = voidLastPayment(payload);    invalidateCardBalanceCache(); return r; }
    else if (action === 'setSeedBalance')     { var r = setSeedBalance(payload);     invalidateCardBalanceCache(); return r; }
    else if (action === 'saveNetWorthSnapshot') { invalidateNetWorthCache(); return saveNetWorthSnapshot(payload); }
    else if (action === 'makeCheckingEntry')  { var r = _withIdem(payload, function(){ return makeCheckingEntry(payload); });  if (payload.month) invalidateMonthCache(payload.month); return r; }
    else if (action === 'addVariableEntry')   { var r = _withIdem(payload, function(){ return addVariableEntry(payload); });   if (payload.month) invalidateMonthCache(payload.month); invalidateCardBalanceCache(); return r; }
    else if (action === 'updateVariableEntry') { invalidateCardBalanceCache(); return updateVariableEntry(payload); }
    else if (action === 'deleteVariableEntry') { invalidateCardBalanceCache(); return deleteVariableEntry(payload); }
    else if (action === 'reimburseReceipt')   return _withIdem(payload, function(){ return reimburseReceipt(payload); }); // invalidates HSA cache internally
    else if (action === 'addHsaReceipt')      return _withIdem(payload, function(){ return addHsaReceipt(payload); });   // invalidates HSA cache internally
    else if (action === 'scanHsaFolder')      return scanHsaFolder(); // idempotent via fileId dedup, not _withIdem
    else return { error: 'Unknown action: ' + action };
  } catch(err) {
    return { error: err.message };
  }
}


function doGet(e) {
  _resetCaches();
  var action = (e.parameter && e.parameter.action) || 'all';
  var data;
  try {
    if      (action === 'all')             data = getAllData();
    else if (action === 'quick')           data = getQuickData();
    else if (action === 'monthly')         data = getMonthlyData();
    else if (action === 'networth')        data = getNetWorth();
    else if (action === 'transactions')    data = getTransactions(e.parameter.month);
    else if (action === 'fixed')           data = getFixedExpenses(e.parameter.month);
    else if (action === 'variableEntries') data = getVariableEntries(e.parameter.month);
    else if (action === 'accounts')        data = getAccountsData();
    else if (action === 'cardTotals')      data = getCardTotals();
    else if (action === 'cardBalances')    data = getCardBalances();
    else if (action === 'loans')           data = getLoanData();
    else if (action === 'mortgage')        data = getMortgageData();
    else if (action === 'cardTxns')        data = getTransactionsByCard(e.parameter.pm, e.parameter.limit);
    else if (action === 'config')          data = getConfig();
    else if (action === 'hsa')             data = getHsa();
    else data = { error: 'Unknown action: ' + action };
  } catch(err) {
    data = { error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  _resetCaches();
  var data;
  try {
    var p = JSON.parse(e.postData.contents);
    _lastPostParams = p; // captured for getQuickData to know which month changed
    if      (p.action === 'updateTransaction')  data = updateTransaction(p);
    else if (p.action === 'addTransaction')     data = _withIdem(p, function(){ return addTransaction(p); });
    else if (p.action === 'deleteTransaction')  data = deleteTransaction(p);
    else if (p.action === 'setFixedPaid')       data = setFixedPaid(p);
    else if (p.action === 'updateFixedCost')    data = updateFixedCost(p);
    else if (p.action === 'updateFixedName')    data = updateFixedName(p);
    else if (p.action === 'updateLoanBalance')  data = _withIdem(p, function(){ return updateLoanBalance(p); });
    else if (p.action === 'logCreditSnapshot')  data = logCreditSnapshot(p);
    else if (p.action === 'addMonth')           data = addMonth(p);
    else if (p.action === 'makeCardPayment')    data = _withIdem(p, function(){ return makeCardPayment(p); });
    else if (p.action === 'voidLastPayment')    data = voidLastPayment(p);
    else if (p.action === 'setSeedBalance')     data = setSeedBalance(p);
    else if (p.action === 'saveNetWorthSnapshot') { data = saveNetWorthSnapshot(p); invalidateNetWorthCache(); }
    else if (p.action === 'splitTransaction')   data = splitTransaction(p);
    else if (p.action === 'addVariableEntry')   data = _withIdem(p, function(){ return addVariableEntry(p); });
    else if (p.action === 'makeCheckingEntry')  data = _withIdem(p, function(){ return makeCheckingEntry(p); });
    else if (p.action === 'updateVariableEntry') data = updateVariableEntry(p);
    else if (p.action === 'deleteVariableEntry') data = deleteVariableEntry(p);
    else if (p.action === 'reimburseReceipt')   data = _withIdem(p, function(){ return reimburseReceipt(p); }); // invalidates HSA cache internally
    else if (p.action === 'addHsaReceipt')      data = _withIdem(p, function(){ return addHsaReceipt(p); });   // invalidates HSA cache internally
    else if (p.action === 'scanHsaFolder')      data = scanHsaFolder(); // idempotent via fileId dedup, not _withIdem
    else data = { error: 'Unknown POST action: ' + p.action };
    // Invalidate caches for the affected month and card balances
    if (p.month) invalidateMonthCache(p.month);
    // Invalidate card balance cache for any action that could change balances
    var BALANCE_ACTIONS = ['updateTransaction','addTransaction','deleteTransaction',
      'splitTransaction','addVariableEntry','updateVariableEntry','deleteVariableEntry',
      'makeCardPayment','voidLastPayment','setSeedBalance','makeCheckingEntry',
      'setFixedPaid','updateFixedCost'];
    if (BALANCE_ACTIONS.indexOf(p.action) >= 0) invalidateCardBalanceCache();
  } catch(err) {
    data = { error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// -- GET: all ------------------------------------------------
function getAllData() {
  var loanData = getLoanData();
  return {
    config:        getConfig(),
    monthly:       getMonthlyData(),
    accounts:      getAccountsData(),
    cardTotals:    getCardTotals(),
    cardBalances:  getCardBalances(),
    budget:        cfg().budget,
    months:        getMonthlySheetNames(),
    loanInterest:  loanData.sheetTotalInterest || loanData.totalInterest || 0,
    loanPrincipal: loanData.sheetTotalPayment  ? (loanData.sheetTotalPayment - (loanData.sheetTotalInterest||0)) : 0,
    updated:       new Date().toISOString()
  };
}

// Fast refresh -- called after every save. Only recomputes the current month;
// all other months served from CacheService. Typically reads 1 sheet, not 17.
function getQuickData() {
  var p = _lastPostParams || {};
  var currentMonth = p.month || getMonthlySheetNames()[0] || '';
  return {
    monthly:       getMonthlyDataQuick(currentMonth),
    months:        getMonthlySheetNames(),
    updated:       new Date().toISOString()
  };
}

// Holds the params from the most recent doPost call so getQuickData
// knows which month to invalidate without an extra parameter.
var _lastPostParams = null;

// -- CACHE UTILITIES -------------------------------------------
// Run warmMonthlyCache() manually after adding a new month sheet,
// or whenever you want to prime the cache cold.
function warmMonthlyCache() {
  var months  = getMonthlySheetNames();
  var cache   = CacheService.getScriptCache();
  var count   = 0;
  var startMs = Date.now();
  months.forEach(function(name) {
    if (Date.now() - startMs > 4.5 * 60 * 1000) {
      Logger.log('Timeout guard: cached ' + count + ' of ' + months.length + ' months');
      return;
    }
    var row = _computeOneMonth(name);
    if (row) {
      var json = JSON.stringify(row);
      try { cache.put(_monthCacheKey(name), json, 21600); count++; } catch(e) {}
      _monthPropsSet(name, json);
    }
  });
  Logger.log('warmMonthlyCache: cached ' + count + ' months in ' + Math.round((Date.now()-startMs)/1000) + 's');
}

// Wipe all cached monthly summaries (useful after changing FCF exclusion list or labels).
function clearMonthlyCache() {
  var months = getMonthlySheetNames();
  var keys   = months.map(function(n){ return _monthCacheKey(n); });
  try {
    CacheService.getScriptCache().removeAll(keys);
  } catch(e) {}
  months.forEach(function(n){ _monthPropsRemove(n); });
  Logger.log('Cleared ' + keys.length + ' monthly cache entries (both tiers)');
}

// -- GET: monthly summaries --------------------------
// Fixed expense rows to exclude from FCF (savings & wealth-building only)
var FCF_EXCLUDE_FIXED = [
  'Wesley 529', 'Max 529', '529',
  'IRA', 'Roth IRA',
  'Extra loan payments',
  'Emergency Fund',
  'Investment', 'Brokerage',
];

// Savings items to track separately (for pre-savings FCF add-back)
var FCF_SAVINGS_ITEMS = [
  'Wesley 529', 'Max 529', '529',
  'IRA', 'Roth IRA',
  'Emergency Fund',
];

// Daycare FSA annual contribution: Ben is paid bimonthly (24 pay periods/yr ? $312.50)
var DAYCARE_FSA_ANNUAL = 24 * 312.50; // $7,500/yr -> $625/mo net benefit

// -- MONTHLY DATA CACHE ---------------------------------------
// CacheService stores pre-computed monthly summaries for up to 6 hours.
// Key: 'monthly_v2_<monthName>'  Value: JSON of one month's summary row
// This avoids re-scanning every month sheet on every quick refresh.

var CACHE_VERSION      = 'monthly_v3'; // bump when getMonthlyData fields change
var CARD_CACHE_VERSION = 'cards_v1';   // bump when getCardBalances fields change
var CARD_CACHE_KEY     = CARD_CACHE_VERSION + '_balances';
var CARD_CACHE_TTL     = 300;          // 5 minutes -- short since balances change often

function _monthCacheKey(name) {
  return CACHE_VERSION + '_' + name.replace(/\s+/g, '_');
}

// -- Durable month-summary tier (PropertiesService) ------------
// CacheService expires after 6h max, so the first load each day was
// recomputing every month sheet. ScriptProperties never expire: closed
// months are effectively immutable, and any edit invalidates both tiers.
// Keys share _monthCacheKey, so CACHE_VERSION bumps apply here too
// (old-version property entries become orphans; clearMonthlyCache sweeps
// the current version, and total size is ~1KB/month -- negligible).
function _monthPropsAll() {
  try { return PropertiesService.getScriptProperties().getProperties(); }
  catch (e) { return {}; }
}
function _monthPropsSet(name, json) {
  try { PropertiesService.getScriptProperties().setProperty(_monthCacheKey(name), json); }
  catch (e) {}
}
function _monthPropsRemove(name) {
  try { PropertiesService.getScriptProperties().deleteProperty(_monthCacheKey(name)); }
  catch (e) {}
}

function _computeOneMonth(name) {
  var values = readSheet(name);
  if (!values) return null;
  var row = { month: name };
  var inIncome = false, income = 0, benIncome = 0, jennaIncome = 0;
  var inFixed  = false;
  var fixedFcf = 0, fixedSavings = 0, daycareGross = 0, daycareReimb = 0;

  for (var i = 0; i < values.length; i++) {
    var r     = values[i];
    var label = String(r[1] || '').trim();
    var val   = parseFloat(r[2]);

    if (!isNaN(val)) {
      if      (label === 'Total Expenses')                               row.totalExpenses  = val;
      else if (label === 'Budget')                                       row.budget         = val;
      else if (label === 'Deficit/Surplus')                              row.surplusDeficit = val;
      else if (label === 'Total Discretionary expenses')                 row.discretionary  = val;
      else if (label === 'Total Recurring Fixed Expenses')               row.fixed          = val;
      else if (label.indexOf('Total Actual Recurring Variable') === 0)   row.variable       = val;
      else if (label === 'Total One-time Expenses')                      row.oneTime        = val;
      else if (label === 'Total Non-Descretionary One-time' ||
               label === 'Total Non-Discretionary One-time')             row.nonDiscOneTime = val;
      else if (label === 'Net for month')                                row.netForMonth    = val;
      else if (inIncome && val > 0) {
        income += val;
        if      (label.indexOf('Schrodinger') >= 0 || label.indexOf('Schr') >= 0) benIncome   += val;
        else if (label.indexOf('St. Joe') >= 0 || label.indexOf('St Joe') >= 0)   jennaIncome += val;
      }
    }

    if (label === 'Income') { inIncome = true; continue; }
    if (inIncome && !label && isNaN(val)) inIncome = false;

    if (label === 'Recurring Fixed Expenses') { inFixed = true; continue; }
    if (inFixed && (label === 'Actual Recurring Variable Expenses' || label === 'Totals')) inFixed = false;

    if (inFixed && label && !isNaN(val)) {
      var isSavings  = FCF_SAVINGS_ITEMS.some(function(kw){ return label.indexOf(kw) >= 0; });
      var isExcluded = FCF_EXCLUDE_FIXED.some(function(kw){ return label.indexOf(kw) >= 0; });
      if (label === 'Daycare') {
        daycareGross += val;
      } else if (label.indexOf('Daycare Reimbursement') >= 0) {
        daycareReimb += val;
      } else {
        if (isSavings)   fixedSavings += val;
        if (!isExcluded) fixedFcf     += val;
      }
    }
  }

  var daycareNet = daycareGross - (cfg().daycareFsaAnnual / 12);
  if (daycareGross > 0) fixedFcf += Math.max(0, daycareNet);

  if (income > 0)      row.income       = Math.round(income      * 100) / 100;
  if (benIncome > 0)   row.benIncome    = Math.round(benIncome   * 100) / 100;
  if (jennaIncome > 0) row.jennaIncome  = Math.round(jennaIncome * 100) / 100;
  row.fixedFcf      = Math.round(fixedFcf      * 100) / 100;
  row.fixedSavings  = Math.round(fixedSavings  * 100) / 100;
  row.daycareGross  = Math.round(daycareGross  * 100) / 100;
  row.daycareNet    = Math.round(Math.max(0, daycareNet) * 100) / 100;
  return row;
}

// Full recalculate -- reads all months, writes results to cache.
// Called by getAllData and the monthly endpoint.
function getMonthlyData() {
  var cache   = CacheService.getScriptCache();
  var props   = _monthPropsAll(); // one read for all months
  var months  = getMonthlySheetNames();
  var results = [];

  months.forEach(function(name) {
    var key    = _monthCacheKey(name);
    var cached = null;
    try { cached = cache.get(key); } catch(e) {}

    var row;
    if (cached) {
      try { row = JSON.parse(cached); } catch(e) {}
    }
    // Tier 2: durable properties (survives CacheService 6h expiry)
    if (!row && props[key]) {
      try { row = JSON.parse(props[key]); } catch(e) {}
      if (row) { try { cache.put(key, props[key], 21600); } catch(e) {} }
    }
    if (!row) {
      row = _computeOneMonth(name);
      if (row) {
        var json = JSON.stringify(row);
        try { cache.put(key, json, 21600); } catch(e) {}
        _monthPropsSet(name, json);
      }
    }
    if (row) results.push(row);
  });
  return results;
}

// Partial recalculate -- only recomputes the current month, serves rest from cache.
// Called by getQuickData after every save. Typically reads just 1 sheet instead of 17.
function getMonthlyDataQuick(currentMonth) {
  var cache   = CacheService.getScriptCache();
  var months  = getMonthlySheetNames();
  var results = [];

  var props = _monthPropsAll();
  months.forEach(function(name) {
    var key = _monthCacheKey(name);
    var row;
    if (name === currentMonth) {
      // Always recompute the current month -- it just changed
      invalidateSheet(name);
      row = _computeOneMonth(name);
      if (row) {
        var json = JSON.stringify(row);
        try { cache.put(key, json, 21600); } catch(e) {}
        _monthPropsSet(name, json);
      }
    } else {
      // Serve older months: cache -> durable properties -> compute
      var cached = null;
      try { cached = cache.get(key); } catch(e) {}
      if (cached) {
        try { row = JSON.parse(cached); } catch(e) {}
      }
      if (!row && props[key]) {
        try { row = JSON.parse(props[key]); } catch(e) {}
        if (row) { try { cache.put(key, props[key], 21600); } catch(e) {} }
      }
      if (!row) {
        row = _computeOneMonth(name);
        if (row) {
          var json2 = JSON.stringify(row);
          try { cache.put(key, json2, 21600); } catch(e) {}
          _monthPropsSet(name, json2);
        }
      }
    }
    if (row) results.push(row);
  });
  return results;
}

// Invalidate one month's cache entry (call when a month's summary labels change).
function invalidateMonthCache(name) {
  try {
    CacheService.getScriptCache().remove(_monthCacheKey(name));
  } catch(e) {}
  _monthPropsRemove(name);
}

// -- GET: transactions ----------------------------------------
function getTransactions(monthName) {
  if (!monthName) return { error: 'month parameter required' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(monthName);
  if (!sheet) return { error: 'Sheet not found: ' + monthName };
  var values = readSheet(sheet);
  var SECTION_HEADERS = [
    'Recurring Fixed Expenses', 'Actual Recurring Variable Expenses',
    'Total Expenses', 'Total One-time Expenses', 'Total Discretionary expenses',
    'Total Non-Descretionary One-time', 'Total Non-Discretionary One-time',
    'Budget', 'Deficit/Surplus', 'Income', 'Net for month',
    'Discretionary', 'Ben Discretionary', 'Jenna Discretionary',
    'Estimated Recurring Variable Expenses', 'Deposits/Withdrawals',
    'Totals', 'Total Estimated Expenses', 'Total Recurring Fixed Expenses',
  ];

  // Find the hard end of the transaction section
  var endRow = values.length;
  for (var i = 2; i < values.length; i++) {
    var label = String(values[i][1] || '').trim();
    if (SECTION_HEADERS.indexOf(label) >= 0) { endRow = i; break; }
    if (label === 'Activity 2' || label === 'Mileage Plus') { endRow = i; break; }
  }

  // Helper: handles both boolean checkboxes and legacy 'True'/'False' strings
  function boolCol(v) {
    if (v === true)  return true;
    if (v === false) return false;
    return String(v || '').toLowerCase() === 'true';
  }

  var txns = [];
  for (var i = 2; i < endRow; i++) {
    var r       = values[i];
    var dateVal = r[0];
    var desc    = String(r[1] || '').trim();
    var cost    = r[2];
    var costNum = parseFloat(cost);
    var disc = r[4] === true  ? 'true'
             : r[4] === false ? 'false'
             : String(r[4] || '').trim().toLowerCase();

    // Only include rows with a valid Discretionary flag (True/False) and a date.
    // Handles both boolean checkboxes and legacy string values.
    if (disc !== 'true' && disc !== 'false') continue;
    if (!dateVal) continue;

    var dateStr = '';
    if (dateVal instanceof Date) {
      dateStr = dateVal.toISOString().split('T')[0];
    } else if (typeof dateVal === 'number' && dateVal > 40000) {
      dateStr = new Date((dateVal - 25569) * 86400 * 1000).toISOString().split('T')[0];
    } else if (dateVal) {
      dateStr = String(dateVal).split('T')[0];
    }

    txns.push({
      rowIndex: i + 1, date: dateStr, description: desc,
      cost: isNaN(costNum) ? null : costNum,
      paymentMethod: String(r[3] || '').trim(),
      discretionary: boolCol(r[4]),
      ben:           boolCol(r[5]),
      jenna:         boolCol(r[6]),
    });
  }
  return { month: monthName, transactions: txns };
}

// -- GET: fixed + variable expenses ---------------------------
function getFixedExpenses(monthName) {
  if (!monthName) return { error: 'month parameter required' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(monthName);
  if (!sheet) return { error: 'Sheet not found: ' + monthName };
  var values = readSheet(sheet);

  var fixed = [], variable = [];

  // Find "Recurring Fixed Expenses" header row
  var fixedStart = -1, variableStart = -1;
  for (var i = 0; i < values.length; i++) {
    var label = String(values[i][1] || '').trim();
    if (label === 'Recurring Fixed Expenses') fixedStart = i + 1;
    if (label === 'Actual Recurring Variable Expenses') variableStart = i + 1;
  }

  // Parse fixed expenses -- carry forward merged col A (source/payment method)
  if (fixedStart > 0) {
    var lastSource = '';
    for (var i = fixedStart; i < values.length; i++) {
      var r = values[i];
      var name = String(r[1] || '').trim();
      var cost = parseFloat(r[2]);
      if (name === 'Actual Recurring Variable Expenses' || name.indexOf('Total') === 0 || name === 'Totals') break;
      if (!name || isNaN(cost)) continue;

      // Carry forward merged col A value
      var rawSource = String(r[0] || '').trim();
      if (rawSource) lastSource = rawSource;

      var paidVal = String(r[3] || '').trim();
      var paid = paidVal.toLowerCase() === 'true';
      var dateStr = '';
      var d = r[4];
      if (d instanceof Date) dateStr = d.toISOString().split('T')[0];
      else if (typeof d === 'number' && d > 40000) dateStr = new Date((d - 25569) * 86400 * 1000).toISOString().split('T')[0];

      fixed.push({
        rowIndex: i + 1,
        source:   lastSource,
        name:     name,
        cost:     cost,
        paid:     paid,
        paidDate: dateStr
      });
    }
  }

  // Parse variable expenses
  if (variableStart > 0) {
    for (var i = variableStart; i < values.length; i++) {
      var r = values[i];
      var name = String(r[1] || '').trim();
      var cost = parseFloat(r[2]);
      if (name.indexOf('Total') === 0 || name === 'Totals' || name === '') {
        if (name.indexOf('Total') === 0) break;
        continue;
      }
      if (!name || isNaN(cost)) continue;
      variable.push({ rowIndex: i + 1, name: name, cost: cost });
    }
  }

  return { month: monthName, fixed: fixed, variable: variable };
}

// -- GET: variable tracker entries (groceries, gas) ------------
// Returns the individual line items from each tracker section.
// Groceries: cols L=date, M=store, N=cost, O=payment  (rows 4-9, total at N17)
// Gas:       cols L=date, M=cost, N=payment            (rows 21-26, total at M34)
function getVariableEntries(monthName) {
  if (!monthName) return { error: 'month required' };
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(monthName);
  if (!sheet) return { error: 'Sheet not found: ' + monthName };
  var values = readSheet(sheet);

  function parseDate(v) {
    if (v instanceof Date) return v.toISOString().split('T')[0];
    if (typeof v === 'number' && v > 40000) return new Date((v-25569)*86400000).toISOString().split('T')[0];
    return v ? String(v).split('T')[0] : '';
  }

  // -- Groceries: find header row (L='Date', M='Store', N='Cost', O='Payment Method')
  var groceries = [], groceryTotalRow = -1, groceryEntryStart = -1, groceryEntryEnd = -1;
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][11]||'').trim()==='Date' && String(values[i][12]||'').trim()==='Store') {
      groceryEntryStart = i + 1; // 0-based index of first entry row
      // Scan forward for 'Total' in col M (idx 12)
      for (var j = i + 1; j < Math.min(i + 20, values.length); j++) {
        if (String(values[j][12]||'').trim() === 'Total') { groceryTotalRow = j + 1; break; }
        if (values[j][11] instanceof Date || (typeof values[j][11]==='number' && values[j][11]>40000) ||
            (!values[j][11] && !values[j][12] && values[j][13]!=null && parseFloat(values[j][13])>0) ||
            (values[j][11] && String(values[j][11]).trim() && String(values[j][11]).trim()!=='Date')) {
          groceryEntryEnd = j; // last valid entry candidate
        }
      }
      break;
    }
  }
  // Collect grocery entries: rows from groceryEntryStart up to but not including groceryTotalRow
  if (groceryEntryStart > 0 && groceryTotalRow > 0) {
    for (var i = groceryEntryStart; i < groceryTotalRow - 1; i++) {
      var r = values[i];
      var cost = parseFloat(r[13]); // col N (idx 13)
      if (isNaN(cost) || (!r[11] && !r[12] && isNaN(cost))) continue; // skip blank + subtotals
      if (String(r[12]||'').trim().match(/Total|AMEX|Citi|Wells|Prime|Discover|United/i)) continue;
      groceries.push({
        rowIndex: i + 1,
        date:    parseDate(r[11]),
        store:   String(r[12]||'').trim(),
        cost:    isNaN(cost) ? null : Math.round(cost*100)/100,
        payment: String(r[14]||'').trim(),
      });
    }
  }

  // -- Gas: find header row after 'Gas' label (L='Date', M='Cost', N='Payment Method')
  var gas = [], gasTotalRow = -1, gasEntryStart = -1;
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][11]||'').trim() === 'Gas') {
      // Find header row immediately after label
      for (var j = i + 1; j < Math.min(i + 5, values.length); j++) {
        if (String(values[j][11]||'').trim() === 'Date') {
          gasEntryStart = j + 1; // first entry row after header
          break;
        }
      }
      if (gasEntryStart < 0) gasEntryStart = i + 2; // fallback: skip label row
      // Scan forward to find Total row
      for (var j = gasEntryStart; j < Math.min(i + 30, values.length); j++) {
        if (String(values[j][11]||'').trim() === 'Total') { gasTotalRow = j + 1; break; }
      }
      break;
    }
  }
  if (gasEntryStart > 0 && gasTotalRow > 0) {
    for (var i = gasEntryStart; i < gasTotalRow - 1; i++) {
      var r = values[i];
      var cost = parseFloat(r[12]); // col M (idx 12)
      // Skip blank rows and SUMIF label rows (United Total, AMEX, Citi, etc.)
      if (!r[11] && isNaN(cost)) continue;
      if (String(r[11]||'').trim().match(/United|AMEX|Citi|Wells|Prime|Discover|Lowe/i)) break; // hit SUMIF block
      if (isNaN(cost)) continue;
      gas.push({
        rowIndex: i + 1,
        date:    parseDate(r[11]),
        cost:    Math.round(cost*100)/100,
        payment: String(r[13]||'').trim(),
      });
    }
  }

  return {
    month: monthName,
    groceries: { entries: groceries, totalRow: groceryTotalRow, entryStart: groceryEntryStart },
    gas:       { entries: gas,       totalRow: gasTotalRow,     entryStart: gasEntryStart },
  };
}

// -- POST: add entry to a variable expense tracker -------------
// category: 'groceries' | 'gas'
// Groceries: insert in cols L-O before Total row (N17 = SUM(N4:N9))
// Gas:       insert in cols L-N before Total row (M34 = SUM(M21:M26))
function addVariableEntry(p) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(p.month);
  if (!sheet) return { error: 'Sheet not found: ' + p.month };
  invalidateSheet(p.month);

  var amount  = parseFloat(p.amount);
  if (isNaN(amount) || amount <= 0) return { error: 'Invalid amount' };
  var dateObj = p.date ? new Date(p.date + 'T12:00:00') : new Date();
  var values  = readSheet(sheet);

  function parseDate(v) {
    if (v instanceof Date) return true;
    if (typeof v === 'number' && v > 40000) return true;
    return false;
  }

  if (p.category === 'groceries') {
    var headerRow = -1, firstSumifRow = -1, totalRow = -1;
    for (var i = 0; i < values.length; i++) {
      var l = String(values[i][11]||'').trim();
      var m = String(values[i][12]||'').trim();
      var n = String(values[i][13]||'').trim();
      // Header: L='Date', M='Store'
      if (l === 'Date' && (m === 'Store' || m === 'Grocery' || m === 'Groceries')) {
        headerRow = i + 1;
        // Scan forward: find first non-entry row (blank L or SUMIF/Total row in M/N)
        // Entry rows have a date in col L; anything else is SUMIF or Total
        for (var j = i + 1; j < Math.min(i + 60, values.length); j++) {
          var jl = String(values[j][11]||'').trim();
          var jm = String(values[j][12]||'').trim();
          if (jm === 'Total') { totalRow = j + 1; break; }
          // SUMIF label rows in col M mark the end of entries
          if (jm.match(/United|AMEX|Citi|Wells|Prime|Discover|Lowe/i)) { totalRow = j + 1; break; }
          var isDate = values[j][11] instanceof Date ||
            (typeof values[j][11] === 'number' && values[j][11] > 40000);
          if (isDate) {
            // Real entry row -- track the last one seen
            // insertRow = the next row after this one (may be blank or SUMIF)
            firstSumifRow = j + 2; // 1-based row after last entry
          }
        }
        break;
      }
    }
    if (headerRow < 0) return { error: 'Groceries header not found in ' + p.month };
    if (totalRow < 0)  return { error: 'Groceries Total row not found in ' + p.month };

    // Use the row right after the last entry as insertRow
    var insertRow = (firstSumifRow > 0 && firstSumifRow < totalRow)
      ? firstSumifRow
      : headerRow + 1;

    // If the target row is already blank (L and M empty, no date), write directly without inserting
    var targetVals = sheet.getRange(insertRow, 12, 1, 5).getValues()[0];
    var targetIsBlank = !targetVals[0] && !targetVals[1] && !targetVals[2];
    if (!targetIsBlank) {
      // Target row has content (shouldn't happen but be safe) -- insert first
      sheet.getRange(insertRow, 12, 1, 5).insertCells(SpreadsheetApp.Dimension.ROWS);
    }
    sheet.getRange(insertRow, 12).setValue(dateObj);          // L = date
    sheet.getRange(insertRow, 13).setValue(p.store || '');    // M = store
    sheet.getRange(insertRow, 14).setValue(amount);           // N = cost
    sheet.getRange(insertRow, 15).setValue(p.payment || '');  // O = payment
    // col P left blank (no data in that column for groceries)

    // Fix SUM range: Total row shifted down by 1, extend its SUM to cover new row
    var newTotalRow = totalRow + 1;
    var totalFormula = sheet.getRange(newTotalRow, 14).getFormula();
    var sumMatch = totalFormula.match(/SUM\(N(\d+):N(\d+)\)/i);
    if (sumMatch) {
      var sumStart = parseInt(sumMatch[1]);
      var sumEnd   = parseInt(sumMatch[2]);
      if (sumEnd < insertRow) {
        sheet.getRange(newTotalRow, 14).setFormula(
          totalFormula.replace(/SUM\(N\d+:N\d+\)/i, 'SUM(N'+sumStart+':N'+insertRow+')')
        );
      }
    }

    // Fix SUMIF ranges if we inserted a new row (not when writing to existing blank)
    if (!targetIsBlank) {
      var newTotalRow = totalRow + 1;
      for (var si = insertRow; si < newTotalRow; si++) {
        var cellFormula = sheet.getRange(si, 14).getFormula();
        if (!cellFormula) continue;
        var newFormula = cellFormula.replace(/([NO]\d+:[NO])(\d+)/gi, function(m, prefix, end) {
          return parseInt(end) < insertRow ? prefix + insertRow : m;
        });
        if (newFormula !== cellFormula) sheet.getRange(si, 14).setFormula(newFormula);
      }
    }

    // Adjust card balance for the charge

    return { ok: true, category: 'groceries', insertRow: insertRow, totalRow: newTotalRow };
  }

  if (p.category === 'gas') {
    var headerRow = -1, totalRow = -1, lastEntryRow = -1;
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][11]||'').trim() === 'Gas') {
        // Find header row (Date | Cost | Payment Method) right after Gas label
        for (var j = i + 1; j < Math.min(i + 5, values.length); j++) {
          if (String(values[j][11]||'').trim() === 'Date') {
            headerRow = j + 1; // 1-based row of header
            break;
          }
        }
        if (headerRow < 0) headerRow = i + 2; // fallback if no header found
        // Scan forward: track last real entry row, stop at United Total or Total
        for (var j = headerRow; j < Math.min(i + 50, values.length); j++) {
          var jl = String(values[j][11]||'').trim();
          if (jl === 'Total') { totalRow = j + 1; break; }
          // United Total (or any card name) marks the end of entries
          if (jl.match(/United|AMEX|Citi|Wells|Prime|Discover|Lowe/i)) { totalRow = j + 1; break; }
          // Real entry: has a date in col L
          var jIsDate = values[j][11] instanceof Date ||
            (typeof values[j][11] === 'number' && values[j][11] > 40000);
          if (jIsDate) lastEntryRow = j + 2; // insert AFTER (1-based)
        }
        break;
      }
    }
    if (headerRow < 0) return { error: 'Gas tracker header not found in ' + p.month };
    if (totalRow < 0)  return { error: 'Gas tracker Total row not found in ' + p.month };

    // Insert right after last entry; if no entries, insert right after header
    var insertRow = (lastEntryRow > 0 && lastEntryRow < totalRow)
      ? lastEntryRow
      : headerRow + 1;

    // If the target row is already blank, write directly without inserting a new row
    var gasTargetVals = sheet.getRange(insertRow, 12, 1, 3).getValues()[0];
    var gasTargetIsBlank = !gasTargetVals[0] && !gasTargetVals[1];
    if (!gasTargetIsBlank) {
      sheet.getRange(insertRow, 12, 1, 3).insertCells(SpreadsheetApp.Dimension.ROWS);
    }
    sheet.getRange(insertRow, 12).setValue(dateObj);
    sheet.getRange(insertRow, 13).setValue(amount);
    sheet.getRange(insertRow, 14).setValue(p.payment || '');

    // Fix SUMIF/SUM ranges only if we inserted a new row
    if (!gasTargetIsBlank) {
      var newTotalRow = totalRow + 1;
      for (var si = insertRow; si <= newTotalRow; si++) {
        var cellFormula = sheet.getRange(si, 13).getFormula();
        if (!cellFormula) continue;
        var extended = cellFormula.replace(/([LMN]\d+:[LMN])(\d+)/gi, function(m, prefix, end) {
          return parseInt(end) < insertRow ? prefix + insertRow : m;
        });
        if (extended !== cellFormula) sheet.getRange(si, 13).setFormula(extended);
      }
    }

    return { ok: true, category: 'gas', headerRow: headerRow, row: insertRow };
  }

  // Diagnostic: show what's in cols L-M for first 10 rows to help debug
  var diagnostic = [];
  for (var i = 0; i < Math.min(25, values.length); i++) {
    var l = String(values[i][11]||'').trim();
    var m = String(values[i][12]||'').trim();
    if (l || m) diagnostic.push('r'+(i+1)+': L='+l+' M='+m);
  }

  return { error: 'Unknown category: ' + p.category, diagnostic: diagnostic };
}

// -- POST: update a variable entry row -------------------------
function updateVariableEntry(p) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(p.month);
  if (!sheet) return { error: 'Sheet not found: ' + p.month };
  invalidateSheet(p.month);
  var row = parseInt(p.rowIndex);
  if (isNaN(row) || row < 1) return { error: 'Invalid rowIndex' };
  var dateObj   = p.date ? new Date(p.date + 'T12:00:00') : null;
  var newAmount = parseFloat(p.amount) || 0;
  var newPm     = (p.payment || '').trim();

  // Read old values to reverse old charge
  var oldVals = sheet.getRange(row, 12, 1, 5).getValues()[0];
  var oldPm, oldCost;
  if (p.category === 'groceries') {
    oldPm   = String(oldVals[3]||'').trim(); // col O (idx 3)
    oldCost = parseFloat(oldVals[2]);         // col N (idx 2)
    if (dateObj) sheet.getRange(row, 12).setValue(dateObj);
    sheet.getRange(row, 13).setValue(p.store || '');
    sheet.getRange(row, 14).setValue(newAmount);
    sheet.getRange(row, 15).setValue(newPm);
  } else if (p.category === 'gas') {
    oldPm   = String(oldVals[2]||'').trim(); // col N (idx 2)
    oldCost = parseFloat(oldVals[1]);         // col M (idx 1)
    if (dateObj) sheet.getRange(row, 12).setValue(dateObj);
    sheet.getRange(row, 13).setValue(newAmount);
    sheet.getRange(row, 14).setValue(newPm);
  } else {
    return { error: 'Unknown category: ' + p.category };
  }
  // Reverse old, apply new
  SpreadsheetApp.flush();
  return { ok: true };
}

// -- POST: delete a variable entry row -------------------------
function deleteVariableEntry(p) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(p.month);
  if (!sheet) return { error: 'Sheet not found: ' + p.month };
  invalidateSheet(p.month);
  var row = parseInt(p.rowIndex);
  if (isNaN(row) || row < 1) return { error: 'Invalid rowIndex' };
  // Read before deleting to reverse charge
  var oldVals = sheet.getRange(row, 12, 1, 5).getValues()[0];
  var oldPm, oldCost;
  if (p.category === 'groceries') {
    oldPm   = String(oldVals[3]||'').trim();
    oldCost = parseFloat(oldVals[2]);
  } else {
    oldPm   = String(oldVals[2]||'').trim();
    oldCost = parseFloat(oldVals[1]);
  }
  // Reverse the card charge before deleting
  // Delete cols L-P only (tracker columns), leaving A-K intact
  sheet.getRange(row, 12, 1, 5).deleteCells(SpreadsheetApp.Dimension.ROWS);
  SpreadsheetApp.flush();
  return { ok: true };
}


function updateFixedCost(p) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(p.month);
  if (!sheet) return { error: 'Sheet not found: ' + p.month };
  invalidateSheet(p.month);
  sheet.getRange(p.rowIndex, 3).setValue(p.cost);
  return { ok: true };
}

// -- POST: update fixed expense name --------------------------
function updateFixedName(p) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(p.month);
  if (!sheet) return { error: 'Sheet not found: ' + p.month };
  invalidateSheet(p.month);
  sheet.getRange(p.rowIndex, 2).setValue(p.name);
  return { ok: true };
}

// -- POST: toggle fixed expense paid status --------------------
function setFixedPaid(p) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(p.month);
  if (!sheet) return { error: 'Sheet not found: ' + p.month };
  invalidateSheet(p.month);
  sheet.getRange(p.rowIndex, 4).setValue(p.paid ? 'True' : '');
  if (p.paid && p.paidDate) {
    sheet.getRange(p.rowIndex, 5).setValue(new Date(p.paidDate + 'T12:00:00'));
  } else if (!p.paid) {
    sheet.getRange(p.rowIndex, 5).setValue('');
  }
  return { ok: true };
}

// -- POST: update transaction ---------------------------------
function updateTransaction(p) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(p.month);
  if (!sheet) return { error: 'Sheet not found: ' + p.month };
  invalidateSheet(p.month);
  var dateObj = (p.date && p.date.length >= 8) ? new Date(p.date + 'T12:00:00') : '';
  if (dateObj instanceof Date && isNaN(dateObj.getTime())) dateObj = '';
  sheet.getRange(p.rowIndex, 1, 1, 7).setValues([[
    dateObj, p.description || '', p.cost != null ? p.cost : '',
    p.paymentMethod || '',
    p.discretionary ? 'True' : 'False',
    p.ben   ? 'True' : 'False',
    p.jenna ? 'True' : 'False'
  ]]);
  return { ok: true };
}

// -- POST: add transaction (chronological insertion) ----------
function addTransaction(p) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(p.month);
  if (!sheet) return { error: 'Sheet not found: ' + p.month };
  invalidateSheet(p.month);
  var values = readSheet(sheet);
  var newDate = p.date ? new Date(p.date + 'T12:00:00') : null;

  // values[] is 0-based; sheet rows are 1-based (values[i] = sheet row i+1)
  var insertRow = 3; // default: first data row (after blank row 1 and headers row 2)
  for (var i = 2; i < values.length; i++) {
    var r    = values[i];
    var desc = String(r[1] || '').trim();
    if (desc === 'Recurring Fixed Expenses') break;
    if (!desc && !r[0] && (r[2] === null || r[2] === '')) break;

    var rowDate = null;
    if (r[0] instanceof Date) rowDate = r[0];
    else if (typeof r[0] === 'number' && r[0] > 40000) rowDate = new Date((r[0]-25569)*86400000);

    if (newDate && rowDate && rowDate > newDate) {
      // Insert BEFORE this row: sheet row = i+1
      insertRow = i + 1;
      break;
    }
    // This row's date <= new date: insert after it, so tentatively place after = i+2
    if (rowDate) insertRow = i + 2;
  }

  // Insert cells in cols A-I only (transaction area).
  // This shifts transaction data down without affecting the right-side
  // tracker columns (L+) which are now in separate sheets anyway.
  // Col J and K are empty buffers; cols A-I cover Date, Desc, Cost,
  // Payment, Disc, Ben, Jenna, and H/I auto-computed formula columns.
  sheet.getRange(insertRow, 1, 1, 9).insertCells(SpreadsheetApp.Dimension.ROWS);
  var dateObj = (p.date && p.date.length >= 8) ? new Date(p.date + 'T12:00:00') : '';
  if (dateObj instanceof Date && isNaN(dateObj.getTime())) dateObj = '';
  sheet.getRange(insertRow, 1, 1, 7).setValues([[
    dateObj, p.description || '', p.cost != null ? p.cost : '',
    p.paymentMethod || '',
    p.discretionary ? 'True' : 'False',
    p.ben   ? 'True' : 'False',
    p.jenna ? 'True' : 'False'
  ]]);
  // Clear H-I in the inserted row so they recalculate from F-G
  sheet.getRange(insertRow, 8, 1, 2).clearContent();
  // Inserted cells inherit no alignment; Ben column (F) is centered by convention
  sheet.getRange(insertRow, 6).setHorizontalAlignment('center');
  SpreadsheetApp.flush();
  return { ok: true, rowIndex: insertRow };
}

// -- Config sheet ----------------------------------------------
// Editable values live in the 'Config' sheet (key | value | notes) so
// they can change without redeploying. The hardcoded vars elsewhere in
// this file remain as DEFAULTS when the sheet or a key is missing.
// Parsed result is cached 5 minutes (edits take effect within that).
// Run setupConfigSheet() in migrate.gs once to create the sheet.
var CONFIG_SHEET  = 'Config';
var CFG_CACHE_KEY = 'cfg_v3'; // v3: added hsaReceiptFolder (v2 added hsaEstablished) -- bump on every shape change (lesson: stale-shape cache, see INDEX trap)
var _cfgMemo = null;

function cfg() {
  if (_cfgMemo) return _cfgMemo;
  var cache = CacheService.getScriptCache();
  try {
    var hit = cache.get(CFG_CACHE_KEY);
    if (hit) { _cfgMemo = JSON.parse(hit); return _cfgMemo; }
  } catch (e) {}
  var out = _configDefaults();
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
    if (sheet) {
      var rows = sheet.getDataRange().getValues();
      var raw = {};
      for (var i = 0; i < rows.length; i++) {
        var k = String(rows[i][0] || '').trim().toLowerCase();
        if (!k || k === 'key') continue;
        raw[k] = rows[i][1];
      }
      _applyRawConfig(out, raw);
    }
  } catch (e) {}
  _cfgMemo = out;
  try { cache.put(CFG_CACHE_KEY, JSON.stringify(out), 300); } catch (e) {}
  return out;
}

function _configDefaults() {
  return {
    budget:               BUDGET,
    daycareAmount:        DC_AMOUNT,
    daycareLastPaid:      DC_LAST_PAID,
    daycareClosedMondays: DC_CLOSED_MONDAYS,
    daycareFsaAnnual:     DAYCARE_FSA_ANNUAL,
    projectionTarget:     1000000,
    seedMonth:            SEED_MONTH_CONFIG,
    quarterlyExpenses:    QUARTERLY_EXPENSES,
    hsaEstablished:       null,
    hsaReceiptFolder:     null,
  };
}

function _cfgNum(v, d) {
  var n = typeof v === 'number' ? v : parseFloat(v);
  return isNaN(n) ? d : n;
}

function _cfgDateStr(v, d) {
  if (v instanceof Date && !isNaN(v.getTime()))
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var s = String(v || '').trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : d;
}

function _parseClosedMondays(v, d) {
  if (v == null || v === '') return d;
  if (v instanceof Date) { var one = _cfgDateStr(v, ''); return one ? [one] : d; }
  var parts = String(v).split(',').map(function(s){ return s.trim(); }).filter(String);
  var ok = parts.filter(function(s){ return /^\d{4}-\d{2}-\d{2}$/.test(s); });
  return ok.length ? ok : d;
}

// Format: 'Water:1,4,7,10; Sewer:2,5,8,11' (1 = January)
function _parseQuarterly(v, d) {
  if (v == null || String(v).trim() === '') return d;
  var entries = String(v).split(';').map(function(s){ return s.trim(); }).filter(String);
  var out = [];
  for (var i = 0; i < entries.length; i++) {
    var ix = entries[i].indexOf(':');
    if (ix < 1) continue;
    var name = entries[i].slice(0, ix).trim();
    var months = entries[i].slice(ix + 1).split(',')
      .map(function(s){ return parseInt(s.trim(), 10); })
      .filter(function(n){ return n >= 1 && n <= 12; });
    if (name && months.length) out.push({ name: name, months: months });
  }
  return out.length ? out : d;
}

function _applyRawConfig(out, raw) {
  if ('budget' in raw)                 out.budget               = _cfgNum(raw['budget'], out.budget);
  if ('daycare_amount' in raw)         out.daycareAmount        = _cfgNum(raw['daycare_amount'], out.daycareAmount);
  if ('daycare_last_paid' in raw)      out.daycareLastPaid      = _cfgDateStr(raw['daycare_last_paid'], out.daycareLastPaid);
  if ('daycare_closed_mondays' in raw) out.daycareClosedMondays = _parseClosedMondays(raw['daycare_closed_mondays'], out.daycareClosedMondays);
  if ('daycare_fsa_annual' in raw)     out.daycareFsaAnnual     = _cfgNum(raw['daycare_fsa_annual'], out.daycareFsaAnnual);
  if ('projection_target' in raw)      out.projectionTarget     = _cfgNum(raw['projection_target'], out.projectionTarget);
  if ('quarterly_expenses' in raw)     out.quarterlyExpenses    = _parseQuarterly(raw['quarterly_expenses'], out.quarterlyExpenses);
  if ('hsa_established' in raw)         out.hsaEstablished       = _cfgDateStr(raw['hsa_established'], out.hsaEstablished);
  if ('hsa_receipt_folder' in raw)      out.hsaReceiptFolder     = String(raw['hsa_receipt_folder'] || '').trim() || null;
  if ('seed_month' in raw) {
    var sm = raw['seed_month'];
    if (sm instanceof Date && !isNaN(sm.getTime()))
      out.seedMonth = Utilities.formatDate(sm, Session.getScriptTimeZone(), 'MMMM yyyy');
    else if (String(sm).trim())
      out.seedMonth = String(sm).trim();
  }
}

function invalidateConfigCache() {
  _cfgMemo = null;
  try { CacheService.getScriptCache().remove(CFG_CACHE_KEY); } catch (e) {}
}

// -- GET: dashboard configuration ------------------------------
// Single source of truth for all values that would otherwise be
// hardcoded in both the dashboard and Apps Script.
// Add this to getAllData so the dashboard always has current config.
function getConfig() {
  var c = cfg();
  return {
    budget:           c.budget,
    daycareAmount:    c.daycareAmount,
    daycareLastPaid:  c.daycareLastPaid,
    daycareClosedMondays: c.daycareClosedMondays,
    daycareFsaAnnual: c.daycareFsaAnnual,
    projectionTarget: c.projectionTarget,
    cards: CARD_DEFS.map(function(d) {
      return {
        // sheetKey = name used in sheet (e.g. 'Citi', 'Wells Fargo')
        // displayName and pm are the same as sheetKey unless overridden here
        sheetKey:    d.name,
        pm:          CARD_PM_MAP[d.name] || d.name,
        displayName: CARD_DISPLAY_MAP[d.name] || d.name,
        limit:       d.limit,
        due:         d.due,
      };
    }),
    pmOptions: ['', 'AMEX', 'Costco', 'Wells Fargo', 'Prime', 'United', "Lowe's", 'Discover', 'Checking', 'Gift Card'],
    quarterlyExpenses: c.quarterlyExpenses,
  };
}

// Maps sheetKey -> payment method label (as it appears in transaction col D)
// Only needed when the PM label differs from the sheetKey
var CARD_PM_MAP = {
  'Citi':       'Costco',
  'Prime Visa': 'Prime',
};

// Maps sheetKey -> full display name shown in the dashboard UI
var CARD_DISPLAY_MAP = {
  'Wells Fargo': 'Wells Fargo Active Cash',
  'Citi':        'Citi Costco Visa',
  'AMEX':        'AMEX Blue Cash',
  'Prime Visa':  'Amazon Prime Visa',
  "Lowe's":      "Lowe's Advantage",
  'United':      'United Explorer',
  'Discover':    'Discover It',
};

// Daycare config -- update here when schedule changes, redeploy once
var DC_AMOUNT          = 2380;
var DC_LAST_PAID       = '2026-04-10';
var DC_CLOSED_MONDAYS  = ['2026-06-29', '2026-12-21'];


// Flat table layout -- one row per card, no scanning for section starts.
// Columns: A=Card, B=Limit, C=Balance, D=LastPaymentDate, E=LastPaymentAmt, F=Utilization%, G=Due
//
// Row 1: header
// Rows 2-8: one row per card (same order as CARD_DEFS)
// Row 10: total balance / total limit

// -- CARD TRACKERS CONSTANTS -----------------------------------
// Defined here (before the helper functions that use them)
var CT_HEADER_ROW  = 1;  // header row number
var CT_FIRST_ROW   = 2;  // first card row
var CT_COL_CARD    = 1;  // A: card name
var CT_COL_LIMIT   = 2;  // B: credit limit
var CT_COL_BALANCE = 3;  // C: computed balance (written by getCardBalances)
var CT_COL_DATE    = 4;  // D: last payment date
var CT_COL_PAYMENT = 5;  // E: last payment amount
var CT_COL_UTIL    = 6;  // F: utilization %
var CT_COL_DUE     = 7;  // G: payment due date
var CT_COL_SEED    = 8;  // H: seed balance at start of tracking period

var CP_COL_DATE    = 1;  // A: payment date
var CP_COL_CARD    = 2;  // B: card name
var CP_COL_AMOUNT  = 3;  // C: payment amount (negative)
var CP_COL_MONTH   = 4;  // D: month sheet name

// -- CARD BALANCE HELPERS -------------------------------------

// Build reverse map: PM label -> sheetKey
// e.g. 'Costco' -> 'Citi', 'Prime' -> 'Prime Visa', 'AMEX' -> 'AMEX'
function _pmToSheetKey(pm) {
  if (!pm) return null;
  var pmTrimmed = pm.trim();
  var pmLower   = pmTrimmed.toLowerCase();
  // Check reverse of CARD_PM_MAP (case-insensitive)
  var keys = Object.keys(CARD_PM_MAP);
  for (var i = 0; i < keys.length; i++) {
    if (CARD_PM_MAP[keys[i]].toLowerCase() === pmLower) return keys[i];
  }
  // Direct match against CARD_DEFS names (case-insensitive)
  for (var i = 0; i < CARD_DEFS.length; i++) {
    if (CARD_DEFS[i].name.toLowerCase() === pmLower) return CARD_DEFS[i].name;
  }
  return null; // not a card (Checking, Gift Card, etc.)
}

// Expenses that only appear in certain months -- show no checkbox when not due
// Format: { name: 'Water', months: [1, 4, 7, 10] } (month numbers, 1=Jan)
var QUARTERLY_EXPENSES = [
  { name: 'Water', months: [1, 4, 7, 10] },
];
var CARD_TRACKERS_SHEET  = 'Card Trackers';
var CARD_PAYMENTS_SHEET  = 'Card Payments';

var CARD_DEFS = [
  { name: 'United',      limit: 14800,  due: '8th'  },
  { name: 'Prime Visa',  limit: 29800,  due: '12th' },
  { name: 'AMEX',        limit: 40700,  due: '19th' },
  { name: 'Citi',        limit: 14000,  due: '28th' },
  { name: 'Wells Fargo', limit: 18800,  due: '22nd' },
  { name: "Lowe's",      limit: 35000,  due: '5th'  },
  { name: 'Discover',    limit: 6600,   due: '28th' },
];

var CT_TOTAL_ROW = CT_FIRST_ROW + CARD_DEFS.length + 1; // blank row then totals

// Append a single payment to the Card Payments history sheet
function _appendCardPayment(cardName, dateObj, amount, month) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cp = ss.getSheetByName(CARD_PAYMENTS_SHEET);
  if (!cp) return; // sheet not created yet -- silently skip
  invalidateSheet(CARD_PAYMENTS_SHEET);
  var lastRow = Math.max(cp.getLastRow(), 1);
  var newRow  = lastRow + 1;
  cp.getRange(newRow, CP_COL_DATE  ).setValue(dateObj);
  cp.getRange(newRow, CP_COL_CARD  ).setValue(cardName);
  cp.getRange(newRow, CP_COL_AMOUNT).setValue(amount);
  cp.getRange(newRow, CP_COL_MONTH ).setValue(month + ''); // force string
  cp.getRange(newRow, CP_COL_DATE,   1, 1).setNumberFormat('yyyy-mm-dd');
  cp.getRange(newRow, CP_COL_AMOUNT, 1, 1).setNumberFormat('$#,##0.00');
  cp.getRange(newRow, CP_COL_MONTH,  1, 1).setNumberFormat('@'); // plain text
}

// -- GET: card balances -- reads flat table from Card Trackers --
function getCardBalances() {
  // Try cache first (5-minute TTL)
  var cache = CacheService.getScriptCache();
  try {
    var cached = cache.get(CARD_CACHE_KEY);
    if (cached) return JSON.parse(cached);
  } catch(e) {}

  var result = _computeCardBalances();

  // Write to cache
  try { cache.put(CARD_CACHE_KEY, JSON.stringify(result), CARD_CACHE_TTL); } catch(e) {}
  return result;
}

function invalidateCardBalanceCache() {
  try { CacheService.getScriptCache().remove(CARD_CACHE_KEY); } catch(e) {}
}

function _computeCardBalances() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // -- Step 1: Read Card Trackers for metadata and seed balances
  var ctSheet = ss.getSheetByName(CARD_TRACKERS_SHEET);
  var ctMeta  = {};
  if (ctSheet) {
    var ctVals = ctSheet.getDataRange().getValues();
    for (var i = 1; i < ctVals.length; i++) {
      var name = String(ctVals[i][CT_COL_CARD-1]||'').trim();
      if (!name || name === 'TOTAL') continue;
      var lastDate = ctVals[i][CT_COL_DATE-1];
      var seed     = ctVals[i][CT_COL_SEED-1];
      ctMeta[name] = {
        limit:             typeof ctVals[i][CT_COL_LIMIT-1]   === 'number' ? ctVals[i][CT_COL_LIMIT-1]   : null,
        due:               ctVals[i][CT_COL_DUE-1]            || null,
        lastPaymentDate:   lastDate instanceof Date ? lastDate : (lastDate ? new Date(lastDate) : null),
        lastPaymentAmount: typeof ctVals[i][CT_COL_PAYMENT-1] === 'number' ? Math.round(ctVals[i][CT_COL_PAYMENT-1]*100)/100 : null,
        seed:              typeof seed === 'number' ? seed : (parseFloat(seed) || 0),
      };
    }
  }

  // -- Step 2: Build PM -> sheetKey map
  var pmToKey = {};
  CARD_DEFS.forEach(function(def) {
    var pm = CARD_PM_MAP[def.name] || def.name;
    pmToKey[pm]       = def.name;
    pmToKey[def.name] = def.name;
  });

  // -- Step 3: Initialise balances from seed
  var balances = {};
  CARD_DEFS.forEach(function(def) {
    balances[def.name] = (ctMeta[def.name] ? ctMeta[def.name].seed : 0) || 0;
  });

  // -- Step 4: Add all transaction charges from Jan 2025 onwards
  // (December 2024 is excluded because the seed balance already includes those charges)
  var SEED_MONTH = cfg().seedMonth; // earliest month sheet to scan (Config sheet: seed_month)
  var monthNames = getMonthlySheetNames();
  monthNames.forEach(function(monthName) {
    if (monthName === SEED_MONTH) return; // skip -- already in seed
    var vals = readSheet(monthName);
    if (!vals) return;
    // Find end of transaction section (same logic as getTransactions)
    var txnEnd = vals.length;
    var SECT = ['Recurring Fixed Expenses','Actual Recurring Variable Expenses',
      'Total Expenses','Total One-time Expenses','Total Discretionary expenses',
      'Total Non-Descretionary One-time','Total Non-Discretionary One-time',
      'Budget','Deficit/Surplus','Income','Net for month','Discretionary',
      'Ben Discretionary','Jenna Discretionary','Estimated Recurring Variable Expenses',
      'Deposits/Withdrawals','Totals','Total Estimated Expenses','Total Recurring Fixed Expenses'];
    for (var si = 2; si < vals.length; si++) {
      var lbl = String(vals[si][1]||'').trim();
      if (SECT.indexOf(lbl) >= 0 || lbl === 'Activity 2' || lbl === 'Mileage Plus') {
        txnEnd = si;
        break;
      }
    }
    for (var i = 2; i < txnEnd; i++) {
      var dateVal = vals[i][0];
      var pm      = String(vals[i][3]||'').trim();
      var cost    = parseFloat(vals[i][2]);
      if (!dateVal || !pm || isNaN(cost)) continue;
      // Validate it's a real transaction row (disc col is a boolean-like value)
      var disc    = vals[i][4];
      var discStr = disc === true ? 'true' : disc === false ? 'false' : String(disc||'').toLowerCase();
      if (discStr !== 'true' && discStr !== 'false') continue;
      var key = pmToKey[pm];
      if (!key || balances[key] === undefined) continue;
      balances[key] += cost;
    }
  });
  Logger.log('AMEX after txn scan: ' + balances['AMEX'] + ' seed was: ' + (ctMeta['AMEX'] ? ctMeta['AMEX'].seed : 'n/a'));
  monthNames.forEach(function(monthName) {
    if (monthName === SEED_MONTH) return; // skip -- already in seed
    var vals = readSheet(monthName);
    if (!vals) return;
    var inGroc = false;
    for (var i = 0; i < vals.length; i++) {
      var l = String(vals[i][11]||'').trim();
      var m = String(vals[i][12]||'').trim();
      if (l === 'Date' && (m === 'Store' || m === 'Grocery' || m === 'Groceries')) { inGroc = true; continue; }
      if (!inGroc) continue;
      if (m.match(/United|AMEX|Citi|Wells|Prime|Discover|Lowe/i) || m === 'Total') break;
      var isDate = vals[i][11] instanceof Date || (typeof vals[i][11] === 'number' && vals[i][11] > 40000);
      if (!isDate) continue;
      var grocCost = parseFloat(vals[i][13]);
      var grocPm   = String(vals[i][14]||'').trim();
      if (isNaN(grocCost) || !grocPm) continue;
      var grocKey = pmToKey[grocPm];
      if (!grocKey || balances[grocKey] === undefined) continue;
      balances[grocKey] += grocCost;
    }
  });

  // -- Step 6: Add gas tracker charges from Jan 2025 onwards
  monthNames.forEach(function(monthName) {
    if (monthName === SEED_MONTH) return; // skip -- already in seed
    var vals = readSheet(monthName);
    if (!vals) return;
    var inGas = false;
    for (var i = 0; i < vals.length; i++) {
      var l = String(vals[i][11]||'').trim();
      if (l === 'Gas') { inGas = true; continue; }
      if (!inGas) continue;
      if (l.match(/United|AMEX|Citi|Wells|Prime|Discover|Lowe/i) || l === 'Total') break;
      var isDate = vals[i][11] instanceof Date || (typeof vals[i][11] === 'number' && vals[i][11] > 40000);
      if (!isDate) continue;
      var gasCost = parseFloat(vals[i][12]);
      var gasPm   = String(vals[i][13]||'').trim();
      if (isNaN(gasCost) || !gasPm) continue;
      var gasKey = pmToKey[gasPm];
      if (!gasKey || balances[gasKey] === undefined) continue;
      balances[gasKey] += gasCost;
    }
  });

  // -- Step 7: Add fixed expense charges per card across all months
  monthNames.forEach(function(monthName) {
    if (monthName === SEED_MONTH) return;
    var vals = readSheet(monthName);
    if (!vals) return;
    // Find fixed expenses header: col A='Source', col B='Recurring Fixed Expenses'
    var fixedStart = -1;
    for (var i = 0; i < vals.length; i++) {
      if (String(vals[i][0]||'').trim() === 'Source' &&
          String(vals[i][1]||'').trim().indexOf('Recurring Fixed') >= 0) {
        fixedStart = i + 1; break;
      }
    }
    if (fixedStart < 0) return;
    var curSource = '';
    for (var i = fixedStart; i < Math.min(fixedStart + 60, vals.length); i++) {
      var colA = String(vals[i][0]||'').trim();
      var colB = String(vals[i][1]||'').trim();
      var colC = parseFloat(vals[i][2]);
      // New source group
      if (colA) curSource = colA;
      // Stop at section markers
      if (colB === 'Actual Recurring Variable Expenses' || colB === 'Totals' ||
          colB === 'Total Recurring Fixed Expenses' || colB === 'Total Estimated Expenses') break;
      // Skip non-card sources and rows without a cost
      if (!curSource || curSource === 'Checking' || isNaN(colC) || !colB) continue;
      var key = pmToKey[curSource];
      if (!key || balances[key] === undefined) continue;
      // Only count if actually paid (col D = 'True' or a date in col E)
      var paid = vals[i][3];
      var paidDate = vals[i][4];
      var isPaid = paid === true || String(paid||'').toLowerCase() === 'true' || paidDate instanceof Date;
      if (!isPaid) continue;
      balances[key] += colC;
    }
  });

  // -- Step 8: Subtract all payments from Card Payments sheet
  var cp = ss.getSheetByName(CARD_PAYMENTS_SHEET);
  if (cp) {
    var cpVals = cp.getDataRange().getValues();
    var pmCounts = {}, pmTotals = {};
    for (var i = 1; i < cpVals.length; i++) {
      var cardName = String(cpVals[i][CP_COL_CARD-1]||'').trim();
      var amount   = typeof cpVals[i][CP_COL_AMOUNT-1] === 'number'
                     ? cpVals[i][CP_COL_AMOUNT-1] : parseFloat(cpVals[i][CP_COL_AMOUNT-1]);
      if (!cardName || isNaN(amount)) continue;
      var month = String(cpVals[i][CP_COL_MONTH-1]||'');
      if (month.indexOf('_TEST_') >= 0) continue;
      if (month === SEED_MONTH) continue;
      if (balances[cardName] === undefined) continue;
      pmCounts[cardName] = (pmCounts[cardName] || 0) + 1;
      pmTotals[cardName] = Math.round(((pmTotals[cardName] || 0) + amount) * 100) / 100;
      balances[cardName] += amount;
    }
    Logger.log('AMEX payments: count=' + (pmCounts['AMEX']||0) + ' total=' + (pmTotals['AMEX']||0));
    Logger.log('WF payments: count=' + (pmCounts['Wells Fargo']||0) + ' total=' + (pmTotals['Wells Fargo']||0));
  }
  Logger.log('WF after payments: ' + balances['Wells Fargo']);
  Logger.log('AMEX after payments: ' + balances['AMEX']);

  // -- Step 9: Write computed balances back to Card Trackers col C
  if (ctSheet) {
    var ctVals2 = ctSheet.getDataRange().getValues();
    for (var i = 1; i < ctVals2.length; i++) {
      var name = String(ctVals2[i][CT_COL_CARD-1]||'').trim();
      if (!name || name === 'TOTAL' || balances[name] === undefined) continue;
      ctSheet.getRange(i + 1, CT_COL_BALANCE).setValue(balances[name]);
    }
    invalidateSheet(CARD_TRACKERS_SHEET);
  }

  // -- Step 10: Build response
  var cards = {};
  CARD_DEFS.forEach(function(def) {
    var meta = ctMeta[def.name] || {};
    cards[def.name] = {
      balance:     Math.round((balances[def.name] || 0) * 100) / 100,
      limit:       meta.limit        || null,
      seed:        meta.seed         || 0,
      lastDate:    meta.lastPaymentDate
                   ? meta.lastPaymentDate.toISOString().split('T')[0] : null,
      lastPayment: meta.lastPaymentAmount || null,
    };
  });

  var totalBalance = 0, totalLimit = 0;
  Object.keys(cards).forEach(function(k) {
    totalBalance += cards[k].balance || 0;
    totalLimit   += cards[k].limit   || 0;
  });

  // Checking balance from most recent month sheet
  var checkingBalance = null, checkingBalanceAlt = null;
  if (monthNames.length) {
    var mSheet = ss.getSheetByName(monthNames[0]);
    var mVals  = mSheet ? readSheet(mSheet) : [];
    for (var i = 0; i < mVals.length; i++) {
      if (String(mVals[i][1]||'').trim() === 'Current Checking Account Balance') {
        var v1 = mVals[i][2], v2 = i+1 < mVals.length ? mVals[i+1][2] : null;
        checkingBalance    = typeof v1==='number' ? Math.round(v1*100)/100 : null;
        checkingBalanceAlt = typeof v2==='number' ? Math.round(v2*100)/100 : null;
        break;
      }
    }
  }

  return {
    month:             CARD_TRACKERS_SHEET,
    usingCardTrackers: true,
    cards:             cards,
    totalBalance:      Math.round(totalBalance*100)/100,
    totalLimit:        totalLimit,
    checkingBalance:   checkingBalance,
    checkingBalanceAlt:checkingBalanceAlt,
    checkingSource:    monthNames.length ? monthNames[0] : null,
  };
}

// -- POST: update seed balance for a card ---------------------
function setSeedBalance(p) {
  var ct = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CARD_TRACKERS_SHEET);
  if (!ct) return { error: 'Card Trackers not found' };
  var cardName = (p.cardName || '').trim();
  var seed     = parseFloat(p.seed);
  if (!cardName) return { error: 'cardName required' };
  if (isNaN(seed)) return { error: 'Invalid seed value' };
  var vals = ct.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][CT_COL_CARD-1]||'').trim() === cardName) {
      ct.getRange(i + 1, CT_COL_SEED).setValue(seed);
      invalidateSheet(CARD_TRACKERS_SHEET);
      SpreadsheetApp.flush();
      return { ok: true, cardName: cardName, seed: seed };
    }
  }
  return { error: 'Card not found: ' + cardName };
}

// -- POST: record a card payment -- updates flat row in Card Trackers --
function makeCardPayment(p) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var cardName = (p.cardName || '').trim();
  var amount   = Math.abs(parseFloat(p.amount));
  var dateObj  = p.date ? new Date(p.date + 'T12:00:00') : new Date();
  if (isNaN(amount) || amount <= 0) return { error: 'Invalid payment amount' };

  var sheet = ss.getSheetByName(CARD_TRACKERS_SHEET);
  if (!sheet) {
    // Fall back to month sheet if Card Trackers not migrated yet
    return _makeCardPaymentLegacy(p);
  }
  invalidateSheet(CARD_TRACKERS_SHEET);

  // Find the card's row in the flat table
  var values = readSheet(sheet);
  var cardRow = -1;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][CT_COL_CARD-1]||'').trim() === cardName) { cardRow = i+1; break; }
  }
  if (cardRow < 0) return { error: 'Card "' + cardName + '" not found in Card Trackers' };

  // Update balance: subtract payment from current balance
  var curBal = values[cardRow-1][CT_COL_BALANCE-1];
  curBal = typeof curBal === 'number' ? curBal : parseFloat(curBal) || 0;
  var newBal = Math.round((curBal - amount) * 100) / 100;

  sheet.getRange(cardRow, CT_COL_BALANCE).setValue(newBal);
  sheet.getRange(cardRow, CT_COL_DATE   ).setValue(dateObj);
  sheet.getRange(cardRow, CT_COL_PAYMENT).setValue(-amount);

  // Append to Card Payments history log
  _appendCardPayment(cardName, dateObj, -amount, p.month || '');

  return { ok: true, card: cardName, newBalance: newBal, row: cardRow };
}

// -- POST: void the most recent payment for a card -------------
// Finds the last row in Card Payments for this card, deletes it,
// and reverses the balance in Card Trackers.
function voidLastPayment(p) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var cardName = (p.cardName || '').trim();
  if (!cardName) return { error: 'cardName required' };

  // Find the last payment row for this card in Card Payments
  var cp = ss.getSheetByName(CARD_PAYMENTS_SHEET);
  if (!cp) return { error: 'Card Payments sheet not found' };

  var cpVals  = cp.getDataRange().getValues();
  var lastRow = -1;
  var lastAmt = null;
  for (var i = cpVals.length - 1; i >= 1; i--) {
    if (String(cpVals[i][CP_COL_CARD-1]||'').trim() === cardName) {
      lastRow = i + 1; // 1-based sheet row
      lastAmt = typeof cpVals[i][CP_COL_AMOUNT-1] === 'number'
                ? cpVals[i][CP_COL_AMOUNT-1]
                : parseFloat(cpVals[i][CP_COL_AMOUNT-1]);
      break;
    }
  }
  if (lastRow < 0) return { error: 'No payments found for ' + cardName };
  if (isNaN(lastAmt)) return { error: 'Could not read payment amount from row ' + lastRow };

  // Delete that row from Card Payments
  cp.deleteRow(lastRow);
  invalidateSheet(CARD_PAYMENTS_SHEET);

  // Reverse the balance in Card Trackers (add back the absolute amount)
  var ct = ss.getSheetByName(CARD_TRACKERS_SHEET);
  if (!ct) return { error: 'Card Trackers sheet not found' };
  invalidateSheet(CARD_TRACKERS_SHEET);

  var ctVals = readSheet(ct);
  for (var i = 1; i < ctVals.length; i++) {
    if (String(ctVals[i][CT_COL_CARD-1]||'').trim() === cardName) {
      var cardRow = i + 1;
      var curBal  = typeof ctVals[i][CT_COL_BALANCE-1] === 'number'
                    ? ctVals[i][CT_COL_BALANCE-1]
                    : parseFloat(ctVals[i][CT_COL_BALANCE-1]) || 0;
      // lastAmt is negative (payment), so subtracting it adds back to balance
      var newBal = Math.round((curBal - lastAmt) * 100) / 100;
      ct.getRange(cardRow, CT_COL_BALANCE).setValue(newBal);

      // Find the new last payment for this card (to update cols D and E)
      var newLastRow = -1, newLastDate = null, newLastAmt = null;
      var cpVals2 = cp.getDataRange().getValues();
      for (var j = cpVals2.length - 1; j >= 1; j--) {
        if (String(cpVals2[j][CP_COL_CARD-1]||'').trim() === cardName) {
          newLastRow  = j + 1;
          newLastDate = cpVals2[j][CP_COL_DATE-1];
          newLastAmt  = cpVals2[j][CP_COL_AMOUNT-1];
          break;
        }
      }
      ct.getRange(cardRow, CT_COL_DATE   ).setValue(newLastDate || '');
      ct.getRange(cardRow, CT_COL_PAYMENT).setValue(newLastAmt  || '');

      SpreadsheetApp.flush();
      return { ok: true, card: cardName, voidedAmount: lastAmt, newBalance: newBal };
    }
  }
  return { error: 'Card "' + cardName + '" not found in Card Trackers' };
}


function _makeCardPaymentLegacy(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(p.month);
  if (!sheet) return { error: 'Sheet not found: ' + p.month };
  invalidateSheet(p.month);

  var values   = readSheet(sheet);
  var cardName = (p.cardName || '').trim();
  var amount   = Math.abs(parseFloat(p.amount));
  var dateObj  = p.date ? new Date(p.date + 'T12:00:00') : new Date();
  var KNOWN    = CARD_DEFS.map(function(d){ return d.name; });
  var inCard   = false, chargesRow = -1, totalRow = -1;

  for (var i = 0; i < values.length; i++) {
    var l = String(values[i][11]||'').trim();
    var n = String(values[i][13]||'').trim();
    if (!inCard && l===cardName && n==='Amounts') { inCard=true; continue; }
    if (inCard) {
      if (l==='Charges') { chargesRow=i+1; continue; }
      if (l==='Total')   { totalRow=i+1; break; }
      if (KNOWN.indexOf(l)>=0 && n==='Amounts') break;
    }
  }
  if (totalRow<0) return { error: 'Card section "'+cardName+'" not found in '+p.month };

  var insertRow = (chargesRow>0?chargesRow:totalRow-1)+1;
  for (var i=(chargesRow>0?chargesRow:0); i<totalRow-1; i++) {
    if (String(values[i][11]||'').trim()==='Payment') insertRow=i+2;
  }
  sheet.getRange(insertRow,12,1,5).insertCells(SpreadsheetApp.Dimension.ROWS);
  sheet.getRange(insertRow,12).setValue('Payment');
  sheet.getRange(insertRow,13).setValue(dateObj);
  sheet.getRange(insertRow,14).setValue(-amount);

  var newTotalRow = totalRow+1;
  var formula = sheet.getRange(newTotalRow,14).getFormula();
  var m = formula.match(/SUM\(N(\d+):N(\d+)\)/i);
  if (m && parseInt(m[2])<insertRow) {
    sheet.getRange(newTotalRow,14).setFormula(
      formula.replace(/SUM\(N\d+:N\d+\)/i,'SUM(N'+m[1]+':N'+insertRow+')')
    );
  }
  return { ok:true, card:cardName, amount:-amount, row:insertRow };
}


// -- POST: delete a transaction row ----------------------------
// Deletes the full row to keep H/I formulas and right-side tracker cols aligned.
function deleteTransaction(p) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(p.month);
  if (!sheet) return { error: 'Sheet not found: ' + p.month };
  invalidateSheet(p.month);
  var rowIndex = parseInt(p.rowIndex);
  if (isNaN(rowIndex) || rowIndex < 1) return { error: 'Invalid rowIndex: ' + p.rowIndex };
  // Read row before deleting so we can reverse its card charge
  var existing = sheet.getRange(rowIndex, 1, 1, 7).getValues()[0];
  var hasDate  = existing[0] instanceof Date || (typeof existing[0] === 'string' && existing[0].length > 0);
  var hasDesc  = String(existing[1]||'').trim().length > 0;
  if (!hasDate && !hasDesc) {
    return { error: 'Row ' + rowIndex + ' appears empty -- may be stale rowIndex. Reload and try again.' };
  }
  // Delete cols A-I only -- leaves right-side tracker columns (L+) intact
  sheet.getRange(rowIndex, 1, 1, 9).deleteCells(SpreadsheetApp.Dimension.ROWS);
  SpreadsheetApp.flush();
  return { ok: true, deletedRow: rowIndex };
}

// -- POST: split a transaction into multiple rows --------------
// Deletes the original row (cols A-I), then inserts the split
// lines in chronological order at the same position.
// p: { month, rowIndex, date, lines: [{description,cost,paymentMethod,discretionary,ben,jenna}] }
function splitTransaction(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(p.month);
  if (!sheet) return { error: 'Sheet not found: ' + p.month };
  invalidateSheet(p.month);

  var lines    = p.lines || [];
  if (lines.length < 2) return { error: 'Need at least 2 lines to split' };

  var rowIndex = parseInt(p.rowIndex);
  if (isNaN(rowIndex) || rowIndex < 1) return { error: 'Invalid rowIndex' };

  var dateObj = (p.date && p.date.length >= 8) ? new Date(p.date + 'T12:00:00') : new Date();
  if (dateObj instanceof Date && isNaN(dateObj.getTime())) dateObj = new Date();

  // Step 1: Delete original row cols A-I only
  sheet.getRange(rowIndex, 1, 1, 9).deleteCells(SpreadsheetApp.Dimension.ROWS);

  // Step 2: Insert cols A-I for remaining split lines (one slot freed by delete)
  for (var i = 0; i < lines.length - 1; i++) {
    sheet.getRange(rowIndex, 1, 1, 9).insertCells(SpreadsheetApp.Dimension.ROWS);
  }

  // Step 3: Write each split line; clear H-I so formulas recalculate cleanly
  for (var i = 0; i < lines.length; i++) {
    var r = rowIndex + i;
    var line = lines[i];
    sheet.getRange(r, 8, 1, 2).clearContent();
    sheet.getRange(r, 1, 1, 7).setValues([[
      dateObj,
      line.description || '',
      line.cost != null ? line.cost : '',
      line.paymentMethod || '',
      line.discretionary ? 'True' : 'False',
      line.ben   ? 'True' : 'False',
      line.jenna ? 'True' : 'False'
    ]]);
    sheet.getRange(r, 6).setHorizontalAlignment('center');
  }

  SpreadsheetApp.flush(); // ensure all writes are committed before dashboard reloads transactions
  return { ok: true, month: p.month, originalRow: rowIndex, linesInserted: lines.length };
}


function makeCheckingEntry(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(p.month);
  if (!sheet) return { error: 'Sheet not found: ' + p.month };

  var amount = parseFloat(p.amount);
  if (isNaN(amount)) return { error: 'Invalid amount' };
  var desc    = (p.description || '').trim() || (amount >= 0 ? 'Deposit' : 'Withdrawal');
  var dateObj = p.date ? new Date(p.date + 'T12:00:00') : new Date();
  invalidateSheet(p.month);
  var values  = readSheet(sheet);

  // Find Income header: B='Income', C='Cost'
  var incomeHeaderRow = -1;
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][1]||'').trim()==='Income' && String(values[i][2]||'').trim()==='Cost') {
      incomeHeaderRow = i + 1; break;
    }
  }
  if (incomeHeaderRow < 0) return { error: 'Income section not found' };

  // Find 'Net for month' row (end of income block)
  var netRow = -1;
  for (var i = incomeHeaderRow; i < values.length; i++) {
    if (String(values[i][1]||'').trim() === 'Net for month') { netRow = i + 1; break; }
  }
  if (netRow < 0) return { error: '"Net for month" row not found' };

  // Insert before the first blank buffer row inside the income block
  var insertRow = netRow - 1;
  for (var i = incomeHeaderRow; i < netRow - 1; i++) {
    if (!String(values[i][1]||'').trim() && (values[i][2]===null || values[i][2]==='')) {
      insertRow = i + 1; break;
    }
  }

  // Insert cells in cols A-E (shift entire income row including date col A and col E)
  sheet.getRange(insertRow, 1, 1, 5).insertCells(SpreadsheetApp.Dimension.ROWS);
  sheet.getRange(insertRow, 2).setValue(desc).setHorizontalAlignment('right');
  sheet.getRange(insertRow, 3).setValue(amount).setNumberFormat('$#,##0.00').setHorizontalAlignment('right');
  sheet.getRange(insertRow, 4).setValue(dateObj);

  // Net row shifted down -- verify SUM covers insertRow
  var newNetRow = netRow + 1;
  var netFormula = sheet.getRange(newNetRow, 3).getFormula();
  var sumMatch = netFormula.match(/SUM\(C(\d+):C(\d+)\)/i);
  if (sumMatch) {
    var sumStart = parseInt(sumMatch[1]);
    var sumEnd   = parseInt(sumMatch[2]);
    if (sumEnd < insertRow) {
      sheet.getRange(newNetRow, 3).setFormula(
        netFormula.replace(/SUM\(C\d+:C\d+\)/i, 'SUM(C'+sumStart+':C'+insertRow+')')
      );
    }
  }

  return { ok: true, month: p.month, description: desc, amount: amount, row: insertRow };
}

function addMonth(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var monthName = (p.monthName || '').trim();
  if (!monthName) return { error: 'monthName is required' };

  var valid = /^(January|February|March|April|May|June|July|August|September|October|November|December) \d{4}$/.test(monthName);
  if (!valid) return { error: 'Invalid format. Use e.g. "May 2026"' };

  if (ss.getSheetByName(monthName)) return { error: 'Sheet "' + monthName + '" already exists' };

  var template = ss.getSheetByName('Template');
  if (!template) return { error: 'Template sheet not found' };

  // Find insert position: right after Card Trackers (or Card Payments, or Template -- whichever is last)
  var allSheets   = ss.getSheets();
  var templateIdx = -1, insertAfterIdx = -1;
  for (var i = 0; i < allSheets.length; i++) {
    var n = allSheets[i].getName();
    if (n === 'Template')     templateIdx    = i;
    if (n === CARD_TRACKERS_SHEET || n === CARD_PAYMENTS_SHEET || n === 'Template') {
      if (i > insertAfterIdx) insertAfterIdx = i;
    }
  }
  if (insertAfterIdx < 0) insertAfterIdx = templateIdx; // fallback

  // Previous month = most recent existing month sheet (for history carryover)
  var monthNames = getMonthlySheetNames();
  var prevName   = monthNames.length > 0 ? monthNames[0] : null;
  var prevSheet  = prevName ? ss.getSheetByName(prevName) : null;

  // Step 1: Fill template yellow cells with previous month's totals
  if (prevSheet && prevName) {
    fillTemplateYellowCells(template, prevSheet, prevName, monthName);
  }

  // Step 2: Copy the now-filled template to the new month sheet
  var newSheet = template.copyTo(ss);
  newSheet.setName(monthName);
  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(insertAfterIdx + 2); // 1-based: puts it right after Card Trackers

  // Step 3: Remove all yellow fills from the new month sheet
  clearYellowFills(newSheet);

  // Step 3b: Write checking balance formula with this month's name hardcoded.
  // Scans for the 'Current Checking Account Balance' label in col B to find
  // the right row -- resilient to row shifts between months.
  var values = newSheet.getDataRange().getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][1]||'').trim() === 'Current Checking Account Balance') {
      var checkingRow = i + 1; // 1-based
      // Formula: income - fixed paid + card payments this month from Card Payments tab
      // SUM(C105:C109) range is approximate -- use dynamic MATCH to find income rows
      // For now use the same structure as the previous month's formula
      var prevFormula = prevSheet
        ? prevSheet.getRange(checkingRow, 3).getFormula()
        : '';
      // Replace the hardcoded month name in the SUMIF with this month's name
      if (prevFormula && prevFormula.indexOf('Card Payments') >= 0) {
        // Swap out old month name string in the formula
        var newFormula = prevFormula.replace(
          /"[A-Za-z]+ \d{4}"/,
          '"' + monthName + '"'
        );
        newSheet.getRange(checkingRow, 3).setFormula(newFormula);
      } else {
        // First time -- write the full formula fresh
        // Detect income rows: find row with label 'Income' and sum the 4 rows after it
        var incomeStart = -1;
        for (var j = 0; j < values.length; j++) {
          if (String(values[j][1]||'').trim() === 'Income') { incomeStart = j + 2; break; }
        }
        // Detect fixed section: SUMIF(D70:D83,TRUE,C70:C83) -- use same range as template
        // Build formula: income block + card payments - fixed paid
        if (incomeStart > 0) {
          var incomeEnd = incomeStart + 4;
          var formula =
            '=SUM(C' + incomeStart + ':C' + incomeEnd + ')' +
            '-SUMIF(D70:D83,TRUE,C70:C83)' +
            '+SUMIF(\'Card Payments\'!D:D,"' + monthName + '",\'Card Payments\'!C:C)';
          newSheet.getRange(checkingRow, 3).setFormula(formula);
        }
      }
      break;
    }
  }

  // Step 4: Re-initialize the template for next month
  // (re-add new empty yellow rows at top of each history section, remove yellow from row below)
  reinitTemplateForNextMonth(template, monthName);

  return { ok: true, monthName: monthName };
}

// Helper: short label "April 2026" -> "April '26"
function shortLabel(name) {
  var p = name.split(' ');
  return p[0] + " '" + p[1].slice(2);
}

// Helper: find row number by searching col B for exact label text
function findRowByLabel(values, label) {
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][1] || '').trim() === label) return i + 1;
  }
  return -1;
}

// Helper: find row where col M (col 13, idx 12) = 'Total' (grocery gas tracker)
function findColMTotal(values, startSearch) {
  for (var i = (startSearch || 0); i < values.length; i++) {
    if (String(values[i][12] || '').trim() === 'Total') return i + 1;
  }
  return -1;
}
function findColMTotalForGas(values) {
  // Gas tracker has a separate 'Total' in col M -- it's the SECOND one
  var count = 0;
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][12] || '').trim() === 'Total') {
      count++;
      if (count === 2) return i + 1;
    }
  }
  return -1;
}

// Helper: find label in col Q (idx 16)
function findRowByQLabel(values, label) {
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][16] || '').trim() === label) return i + 1;
  }
  return -1;
}

// Helper: find label in col U (idx 20)
function findRowByULabel(values, label) {
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][20] || '').trim() === label) return i + 1;
  }
  return -1;
}

// -- GET: card totals from transaction history -----------------
// Sums all transaction costs by payment method across every month sheet.
// -- GET: loan data from Loans sheet --------------------------
function getLoanData() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Loans');
  if (!sheet) return { error: 'Loans sheet not found' };
  var values = readSheet(sheet);

  var loans = [];
  var totalBalance = 0, totalInterest = 0, totalPayment = 0, totalPrincipal = 0, activeCount = 0;

  // Rows 4-25 (index 3-24): A=origNum, B=renum, C=type, D=balance, E=rate, F=interest, G=principal, H=payment
  // Every TYPED row is returned -- including paid-off ($0) loans -- each tagged
  // with its 1-based sheet row, so the dashboard table is a faithful 1:1 mirror
  // of Loans!C4:H25 and every Balance cell maps to a known edit target (D{row}).
  // Only truly empty rows (no Type in col C) are skipped.
  for (var i = 3; i <= 24 && i < values.length; i++) {
    var r        = values[i];
    var type     = String(r[2] || '').trim();
    if (!type) continue;
    var num      = String(r[1] || r[0] || '').trim();
    var balance  = parseFloat(r[3]); if (isNaN(balance)) balance = 0;
    var rate     = parseFloat(r[4]);
    var interest = parseFloat(r[5]);
    var principal= parseFloat(r[6]);
    var payment  = parseFloat(r[7]);

    var loanObj = {
      row:       i + 1,           // 1-based sheet row; edit target is D{row}
      num:       num,
      type:      type,
      balance:   Math.round(balance  * 100) / 100,
      rate:      isNaN(rate)     ? null : Math.round(rate * 10000) / 100,
      interest:  isNaN(interest) ? null : Math.round(interest  * 100) / 100,
      principal: isNaN(principal)? null : Math.round(principal * 100) / 100,
      payment:   isNaN(payment)  ? 0    : Math.round(payment   * 100) / 100,
    };
    loans.push(loanObj);
    if (loanObj.balance > 0) activeCount++;
    totalBalance   += loanObj.balance;
    totalInterest  += loanObj.interest  || 0;
    totalPrincipal += loanObj.principal || 0;
    totalPayment   += loanObj.payment;
  }

  // Row 26 (index 25) = sheet totals (already evaluated)
  var totalRow   = values[25] || [];
  var sheetTotalBalance  = parseFloat(totalRow[3]);
  var sheetTotalInterest = parseFloat(totalRow[5]);
  var sheetTotalPayment  = parseFloat(totalRow[7]);

  return {
    loans:          loans,
    activeCount:    activeCount,                              // loans with balance > 0
    totalBalance:   Math.round(totalBalance   * 100) / 100,
    totalInterest:  Math.round(totalInterest  * 100) / 100,
    totalPrincipal: Math.round(totalPrincipal * 100) / 100,
    totalPayment:   Math.round(totalPayment   * 100) / 100,
    // Prefer sheet-evaluated totals if available (formulas already computed)
    sheetTotalBalance:  isNaN(sheetTotalBalance)  ? null : Math.round(sheetTotalBalance  * 100) / 100,
    sheetTotalInterest: isNaN(sheetTotalInterest) ? null : Math.round(sheetTotalInterest * 100) / 100,
    sheetTotalPayment:  isNaN(sheetTotalPayment)  ? null : Math.round(sheetTotalPayment  * 100) / 100,
  };
}

// -- POST: edit a single loan's outstanding balance ----------
// Writes Loans!D{row} only. Interest (F), Principal (G), New Balance (I),
// %-of-total (J) and the Total row (D26/F26/H26/I26) are all sheet formulas,
// so they recompute on flush. We read the recomputed values back and return
// them, so the dashboard updates FROM the sheet (single source of truth)
// rather than re-deriving interest/totals locally (which would drift from the
// sheet's exact day-count interest convention).
function updateLoanBalance(p) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Loans');
  if (!sheet) return { error: 'Loans sheet not found' };

  var row = parseInt(p.row, 10);
  var bal = parseFloat(p.balance);
  // Guardrails -- fail loudly, never silently:
  //  - row must be inside the loan table (4-25); never the header (3) or Total (26)
  //  - balance must be a finite number >= 0
  //  - the target row must actually be a loan (Type present in col C)
  // The Type check defends against a stale/forged row index writing into the
  // wrong cell.
  if (isNaN(row) || row < 4 || row > 25)        return { error: 'Loan row out of range (4-25): ' + p.row };
  if (isNaN(bal) || !isFinite(bal) || bal < 0)  return { error: 'Invalid balance: ' + p.balance };

  var typeCell = sheet.getRange(row, 3).getValue(); // col C = Type
  if (!String(typeCell || '').trim())           return { error: 'Row ' + row + ' is not a loan (no Type) -- refusing to write' };

  bal = Math.round(bal * 100) / 100;
  sheet.getRange(row, 4).setValue(bal);  // col D = Balance (the only editable column)
  invalidateSheet('Loans');
  invalidateNetWorthCache();             // Net Worth reads Loans!D26 -> studentLoans
  SpreadsheetApp.flush();                // force formula recompute before read-back

  // Authoritative recomputed values straight from the sheet
  var loanInterest  = parseFloat(sheet.getRange(row, 6).getValue());  // col F = Interest
  var totalBalance  = parseFloat(sheet.getRange(26, 4).getValue());   // D26 = SUM(D4:D25)
  var totalInterest = parseFloat(sheet.getRange(26, 6).getValue());   // F26
  var totalPayment  = parseFloat(sheet.getRange(26, 8).getValue());   // H26

  return {
    ok:            true,
    row:           row,
    balance:       bal,
    loanInterest:  isNaN(loanInterest)  ? null : Math.round(loanInterest  * 100) / 100,
    totalBalance:  isNaN(totalBalance)  ? null : Math.round(totalBalance  * 100) / 100,
    totalInterest: isNaN(totalInterest) ? null : Math.round(totalInterest * 100) / 100,
    totalPayment:  isNaN(totalPayment)  ? null : Math.round(totalPayment  * 100) / 100,
  };
}

// -- GET: mortgage amortization schedule (Loans sheet, columns M-W) ----------
// Layout (1-based): M=date, N=description/paid-flag(TRUE), O=principal balance,
// P=payment, Q=interest, R=mortgage insurance, S=property tax, T=principal,
// U=equity%. Header constants: N3=original rate, N4=valuation, N5=purchase.
// A 'Refinance' marker row (M='Refinance') is NOT a payment -- its O cell holds
// the new rate, and the schedule's interest switches to it on the next row.
// The schedule is a FULL amortization to payoff; rows with N=TRUE are logged
// actuals, the remainder is the sheet's own forward projection. We read it as
// the source of truth rather than re-deriving the schedule. Read-only.
// Single source of truth for the mortgage's CURRENT principal balance.
// Reads the Loans!M:W amortization block and returns the parsed ledger plus the
// "current" row = the last row flagged paid (N===TRUE), else the last row.
// BOTH getMortgageData (mortgage tab) and _computeNetWorth (net-worth liability)
// consume this, so the two views can never disagree. See INDEX: B8 and the ledger
// were previously two hand-maintained sources ~one month's principal apart; this
// collapses them to one definition.
function readMortgageLedger(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Loans');
  if (!sheet) return { error: 'Loans sheet not found' };
  // Bounded read (A1:W{<=500}) -- the schedule ends well before row 500 and
  // this avoids the sheet's stray far-down formatting cells.
  var lastRow = Math.min(sheet.getLastRow() || 500, 500);
  var v  = sheet.getRange(1, 1, lastRow, 23).getValues();   // 0-based; M-W = idx 12-22
  var tz = Session.getScriptTimeZone();
  function num(x){ var f = parseFloat(x); return isNaN(f) ? null : f; }
  function r2(x){ var f = parseFloat(x); return isNaN(f) ? null : Math.round(f * 100) / 100; }
  function pct(x){ return (typeof x === 'number') ? Math.round(x * 10000) / 100 : null; }
  function at(row1, col1){ var row = v[row1 - 1]; return row ? row[col1 - 1] : null; }

  var origRate  = num(at(3, 14));   // N3
  var valuation = num(at(4, 14));   // N4
  var purchase  = num(at(5, 14));   // N5

  var ledger = [], refiRate = null, refiAtIdx = null, lastPaidIdx = -1;
  for (var i = 7; i < v.length; i++) {                 // row 8 (idx 7) onward
    var mc = v[i][12];                                 // M
    if (typeof mc === 'string' && mc.toLowerCase().indexOf('refinance') === 0) {
      refiRate  = num(v[i][14]);                       // O on marker row = new rate
      refiAtIdx = ledger.length;                       // next pushed row is first post-refi month
      continue;
    }
    var bal = v[i][14];                                // O = balance
    if (typeof bal !== 'number') { if (ledger.length) break; else continue; }
    var paid = (v[i][13] === true);                    // N === TRUE => logged actual
    if (paid) lastPaidIdx = ledger.length;
    ledger.push({
      date:      (v[i][12] instanceof Date) ? Utilities.formatDate(v[i][12], tz, 'yyyy-MM-dd') : null,
      paid:      paid,
      balance:   Math.round(bal * 100) / 100,
      payment:   r2(v[i][15]),                          // P (total: P&I + escrow)
      interest:  r2(v[i][16]),                          // Q
      escrow:    Math.round(((num(v[i][17]) || 0) + (num(v[i][18]) || 0)) * 100) / 100, // R(PMI)+S(tax)
      principal: r2(v[i][19]),                          // T
      equityPct: pct(v[i][20])                          // U
    });
  }
  if (!ledger.length) return { error: 'No mortgage rows found in Loans!M:W' };

  // "current" = last logged-actual month; before any payment is logged, the first row.
  var curIdx = (lastPaidIdx >= 0) ? lastPaidIdx : ledger.length - 1;
  return {
    origRate:    origRate,
    valuation:   valuation,
    purchase:    purchase,
    ledger:      ledger,
    refiRate:    refiRate,
    refiAtIdx:   refiAtIdx,
    lastPaidIdx: lastPaidIdx,
    curIdx:      curIdx,
    curBalance:  ledger[curIdx].balance,   // <-- the ONE mortgage-balance definition
    costInitial:   r2(at(9, 23)),    // W9  initial projected total P&I
    costProjected: r2(at(10, 23)),   // W10 current projected total P&I
    costSavings:   r2(at(11, 23))    // W11 saved vs original via extra principal
  };
}

function getMortgageData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var m  = readMortgageLedger(ss);
  if (m.error) return m;
  function r2(x){ var f = parseFloat(x); return isNaN(f) ? null : Math.round(f * 100) / 100; }

  // B8 is now LEGACY: a hand-typed cell that net worth no longer reads. We still
  // surface it so the tab can flag if it has drifted from the schedule (a manual
  // cross-check), but it is NOT a source of truth.
  var pa = ss.getSheetByName('Physical Assets');
  var b8 = pa ? r2(pa.getRange(8, 2).getValue()) : null;

  var payoff  = m.ledger[m.ledger.length - 1];
  var cur     = m.ledger[m.curIdx];
  var effRate = (m.refiRate != null) ? m.refiRate : m.origRate;

  return {
    origRate:      m.origRate != null ? Math.round(m.origRate * 10000) / 100 : null, // %
    refiRate:      m.refiRate != null ? Math.round(m.refiRate * 10000) / 100 : null, // %
    rate:          effRate    != null ? Math.round(effRate    * 10000) / 100 : null, // % in effect
    valuation:     r2(m.valuation),
    purchasePrice: r2(m.purchase),
    currentBalance: cur.balance,
    currentDate:    cur.date,
    equity:        (m.valuation != null) ? Math.round((m.valuation - cur.balance) * 100) / 100 : null,
    equityPct:     cur.equityPct,
    payment:       cur.payment,
    principalAndInterest: (cur.interest != null && cur.principal != null) ? Math.round((cur.interest + cur.principal) * 100) / 100 : null,
    escrow:        cur.escrow,
    payoffDate:    payoff.date,
    refiAtIdx:     m.refiAtIdx,
    lastPaidIdx:   m.lastPaidIdx,
    costInitial:   m.costInitial,    // W9
    costProjected: m.costProjected,  // W10
    costSavings:   m.costSavings,    // W11
    physicalAssetsBalance: b8,       // legacy B8 -- informational cross-check only
    mortgageSource: 'ledger',        // net worth derives from this same ledger now
    ledger: m.ledger
  };
}

// -- GET: all transactions for a given payment method ---------
// Scans every month sheet, returns transactions where col D = pm.
// Also includes grocery/gas tracker entries attributed to that pm.
function getTransactionsByCard(pm, limitStr) {
  if (!pm) return { error: 'pm required' };
  var limit   = parseInt(limitStr) || 500;
  var months  = getMonthlySheetNames();
  var results = [];

  months.forEach(function(name) {
    var values = readSheet(name);
    if (!values) return;

    function parseDate(v) {
      if (v instanceof Date) return v.toISOString().split('T')[0];
      if (typeof v === 'number' && v > 40000) return new Date((v-25569)*86400000).toISOString().split('T')[0];
      return String(v||'').split('T')[0];
    }

    // Transaction rows
    for (var i = 2; i < values.length; i++) {
      var r    = values[i];
      var desc = String(r[1]||'').trim();
      var rowPm = String(r[3]||'').trim();
      if (desc === 'Recurring Fixed Expenses') break;
      if (!desc && !r[0] && (r[2]===null||r[2]==='')) break;
      var cost = parseFloat(r[2]);
      if (isNaN(cost) || rowPm !== pm) continue;
      results.push({
        month:         name,
        date:          parseDate(r[0]),
        description:   desc,
        cost:          Math.round(cost*100)/100,
        paymentMethod: rowPm,
        discretionary: r[4] === true || String(r[4]||'').trim().toLowerCase() === 'true',
        ben:           r[5] === true || String(r[5]||'').trim().toLowerCase() === 'true',
        jenna:         r[6] === true || String(r[6]||'').trim().toLowerCase() === 'true',
        source:        'transaction',
      });
    }

    // Grocery tracker entries attributed to this pm (col O = payment)
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][11]||'').trim()==='Date' && String(values[i][12]||'').trim()==='Store') {
        for (var j = i+1; j < Math.min(i+25, values.length); j++) {
          if (String(values[j][12]||'').trim()==='Total') break;
          var cost = parseFloat(values[j][13]);
          var epPm = String(values[j][14]||'').trim();
          if (!isNaN(cost) && epPm === pm) {
            results.push({
              month: name, date: parseDate(values[j][11]),
              description: String(values[j][12]||'Groceries').trim(),
              cost: Math.round(cost*100)/100,
              paymentMethod: pm, source: 'groceries',
              discretionary: false, ben: false, jenna: false,
            });
          }
        }
        break;
      }
    }

    // Gas tracker entries (col N = payment)
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][11]||'').trim()==='Date' && String(values[i][12]||'').trim()==='Cost' &&
          String(values[i][13]||'').trim()==='Payment Method') {
        for (var j = i+1; j < Math.min(i+25, values.length); j++) {
          if (String(values[j][11]||'').trim()==='Total') break;
          var cost = parseFloat(values[j][12]);
          var epPm = String(values[j][13]||'').trim();
          if (!isNaN(cost) && epPm === pm) {
            results.push({
              month: name, date: parseDate(values[j][11]),
              description: 'Gas', cost: Math.round(cost*100)/100,
              paymentMethod: pm, source: 'gas',
              discretionary: false, ben: false, jenna: false,
            });
          }
        }
        break;
      }
    }
  });

  // Sort newest first
  results.sort(function(a,b){ return (b.date||'').localeCompare(a.date||''); });
  return { pm: pm, transactions: results.slice(0, limit), total: results.length };
}


// Single pass per sheet: collects transactions + grocery + gas in one loop.
function getCardTotals() {
  var monthNames = getMonthlySheetNames().slice(0, 12); // TTM
  var totals = {}, counts = {};

  monthNames.forEach(function(name) {
    var values = readSheet(name);
    if (!values) return;

    var txnDone = false;
    var grocHeaderIdx = -1, gasHeaderIdx = -1;

    for (var i = 0; i < values.length; i++) {
      var r   = values[i];
      var l11 = String(r[11] || '').trim(); // col L
      var l12 = String(r[12] || '').trim(); // col M
      var l13 = String(r[13] || '').trim(); // col N

      // -- Transactions (rows 2+, stop at fixed section) --
      if (!txnDone && i >= 2) {
        var desc = String(r[1] || '').trim();
        var pm   = String(r[3] || '').trim();
        if (desc === 'Recurring Fixed Expenses' || (!desc && !r[0] && (r[2] === null || r[2] === ''))) {
          txnDone = true;
        } else {
          var cost = parseFloat(r[2]);
          if (!isNaN(cost) && pm) { totals[pm]=(totals[pm]||0)+cost; counts[pm]=(counts[pm]||0)+1; }
        }
      }

      // -- Detect tracker headers --
      if (l11 === 'Date') {
        if (l12 === 'Store' && grocHeaderIdx < 0) grocHeaderIdx = i;
        else if (l12 === 'Cost' && l13 === 'Payment Method' && gasHeaderIdx < 0) gasHeaderIdx = i;
      }

      // -- Grocery entries --
      if (grocHeaderIdx >= 0 && i > grocHeaderIdx && i <= grocHeaderIdx + 20) {
        if (l12 === 'Total') { grocHeaderIdx = -2; }
        else {
          var cost = parseFloat(r[13]);            // col N
          var pm   = String(r[14] || '').trim();   // col O
          if (!isNaN(cost) && pm) { totals[pm]=(totals[pm]||0)+cost; counts[pm]=(counts[pm]||0)+1; }
        }
      }

      // -- Gas entries --
      if (gasHeaderIdx >= 0 && i > gasHeaderIdx && i <= gasHeaderIdx + 20) {
        if (l11 === 'Total') { gasHeaderIdx = -2; }
        else {
          var cost = parseFloat(r[12]);            // col M
          var pm   = String(r[13] || '').trim();   // col N
          if (!isNaN(cost) && pm) { totals[pm]=(totals[pm]||0)+cost; counts[pm]=(counts[pm]||0)+1; }
        }
      }
    }
  });

  Object.keys(totals).forEach(function(k) { totals[k] = Math.round(totals[k]*100)/100; });
  return { totals: totals, counts: counts, months: monthNames.length, label: 'TTM' };
}

// -- GET: accounts data (Credit history + Loans + Discover) ---
function getAccountsData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = {};

  // Credit sheet
  var creditSheet = ss.getSheetByName('Credit');
  if (creditSheet) {
    var cv = readSheet(creditSheet);
    var history = [];
    for (var i = 1; i < cv.length; i++) {
      var r = cv[i];
      var dateVal = r[0], bal = parseFloat(r[1]), lim = parseFloat(r[2]);
      if (!dateVal || isNaN(bal)) continue;
      var dateStr;
      if (dateVal instanceof Date) dateStr = dateVal.toISOString().split('T')[0];
      else if (typeof dateVal === 'number') dateStr = new Date((dateVal-25569)*86400000).toISOString().split('T')[0];
      else dateStr = String(dateVal).split('T')[0];
      history.push({ date: dateStr, balance: Math.round(bal*100)/100, limit: isNaN(lim)?null:lim });
    }
    result.creditHistory = history; // already newest-first
  }

  // Discover Savings -- col D = running balance (most recent row 0)
  var discSheet = ss.getSheetByName('Discover Savings');
  if (discSheet) {
    var dv = readSheet(discSheet);
    if (dv.length > 0) result.discoverBal = parseFloat(dv[0][3]) || null;
  }

  // Loans -- find Total row
  var loanSheet = ss.getSheetByName('Loans');
  if (loanSheet) {
    var lv = readSheet(loanSheet);
    for (var i = 0; i < lv.length; i++) {
      if (String(lv[i][2]).trim() === 'Total') {
        result.loanTotal = parseFloat(lv[i][3]) || null;
        break;
      }
    }
  }

  return result;
}

// -- POST: log new credit snapshot row ------------------------
function logCreditSnapshot(p) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Credit');
  if (!sheet) return { error: 'Credit sheet not found' };
  invalidateSheet('Credit');
  var values = readSheet(sheet);
  // Get current limit from most recent row if not provided
  var prevLimit = p.limit || (values.length > 1 ? parseFloat(values[1][2]) : 0);
  var util = prevLimit > 0 ? p.balance / prevLimit : 0;
  // TTM average: average of last 11 entries + this one
  var recentBals = values.slice(1,12).map(function(r){ return parseFloat(r[1])||0; });
  recentBals.unshift(p.balance);
  var ttm = recentBals.slice(0,12).reduce(function(s,v){return s+v;},0) / Math.min(recentBals.length,12);
  // Insert at row 2 (after header), shifting existing rows down in cols A-E only
  sheet.getRange(2, 1, 1, 5).insertCells(SpreadsheetApp.Dimension.ROWS);
  var dateObj = new Date(p.date + 'T12:00:00');
  sheet.getRange(2, 1, 1, 5).setValues([[dateObj, p.balance, prevLimit, util, Math.round(ttm*100)/100]]);
  return { ok: true };
}

// -- GET: net worth --------------------------------------------
var NET_WORTH_SHEET      = 'Net Worth';
var NW_CACHE_KEY         = 'nw_v3_data';
var NW_CACHE_TTL         = 300; // 5 minutes

function getNetWorth() {
  var cache = CacheService.getScriptCache();
  try {
    var cached = cache.get(NW_CACHE_KEY);
    if (cached) {
      var parsed = JSON.parse(cached);
      // Sanity check: never serve a zeroed result. If a bad compute got
      // cached (flaky read mid-recalc), fall through and recompute fresh.
      if (_nwLooksValid(parsed)) return parsed;
      cache.remove(NW_CACHE_KEY);
    }
  } catch(e) {}
  var result = _computeNetWorth();
  // Trim before caching -- CacheService has 100KB limit per key
  // Keep last 260 history points (~5 years weekly) and thin projection points
  try {
    if (!_nwLooksValid(result)) return result; // never cache a bad result
    var cacheable = JSON.parse(JSON.stringify(result)); // deep copy
    if (cacheable.history && cacheable.history.length > 260) {
      cacheable.history = cacheable.history.slice(-260);
    }
    if (cacheable.projections) {
      cacheable.projections = cacheable.projections.map(function(p) {
        // Keep every other point to halve the size
        return {
          rate: p.rate, label: p.label, milestone: p.milestone,
          weeklyContrib: p.weeklyContrib,
          points: p.points.filter(function(_,i){ return i % 2 === 0; })
        };
      });
    }
    var json = JSON.stringify(cacheable);
    if (json.length < 90000) { // only cache if under 90KB
      cache.put(NW_CACHE_KEY, json, NW_CACHE_TTL);
    }
  } catch(e) {}
  return result;
}

// A net worth result is plausible only if investments came through and
// the weekly history was found. Zeros mean a failed/partial read.
function _nwLooksValid(r) {
  return r && !r.error &&
         typeof r.investments === 'number' && r.investments > 0 &&
         r.history && r.history.length > 0;
}

function invalidateNetWorthCache() {
  try { CacheService.getScriptCache().remove(NW_CACHE_KEY); } catch(e) {}
}

function _computeNetWorth() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pmSheet = ss.getSheetByName(PORTFOLIO_SHEET);
  if (!pmSheet) return { error: 'Portfolio Management sheet not found' };
  var values = readSheet(pmSheet);

  // -- Current investment accounts (rows 2-17, col B=label, C=value) --
  var accounts = {}, investments = 0, liquid = 0;
  var seenLabels = {};
  values.slice(0, 20).forEach(function(r) {
    var label = String(r[1]||'').trim(), val = parseFloat(r[2]);
    if (isNaN(val)) return;
    if (label === 'Net worth' || label === 'Net worth (ex SDGR)') return;
    if (label === 'Liquid') { liquid = val; return; }
    if (label && val > 0) {
      // Deduplicate: if label already seen, append a counter
      var key = label;
      if (seenLabels[label] !== undefined) {
        seenLabels[label]++;
        key = label + ' (' + seenLabels[label] + ')';
      } else {
        seenLabels[label] = 0;
      }
      accounts[key] = Math.round(val * 100) / 100;
    }
  });
  // C18 = total investments
  investments = typeof values[17][2] === 'number' ? values[17][2] : parseFloat(values[17][2]) || 0;

  // -- Weekly investment history (Portfolio Management rows 25+, cols B=date, C=value) --
  var history = [], inTracker = false;
  values.forEach(function(r) {
    if (String(r[0]||'').trim() === 'Net worth tracker' ||
        String(r[1]||'').trim() === 'Date') { inTracker = true; return; }
    if (!inTracker) return;
    var dateVal = r[1], worth = parseFloat(r[2]);
    if (!dateVal || isNaN(worth) || worth <= 0) return;
    var dateStr;
    if (dateVal instanceof Date) dateStr = dateVal.toISOString().split('T')[0];
    else if (typeof dateVal === 'number') dateStr = new Date((dateVal-25569)*86400*1000).toISOString().split('T')[0];
    else dateStr = String(dateVal).split('T')[0];
    if (dateStr) history.push({ date: dateStr, investments: Math.round(worth) });
  });

  // -- Cash: Discover Savings H4 = balance ex-NHSC (NHSC funds earmarked for student loans)
  // H4 is labeled 'ex. NHSC' -- this is the amount that's actually discretionary cash
  var cash = 0;
  var dsSheet = ss.getSheetByName('Discover Savings');
  if (dsSheet) {
    var dsVals = readSheet(dsSheet);
    // H4 = index 3 (row), col H = index 7 -- ex-NHSC balance
    if (dsVals.length > 3 && typeof dsVals[3][7] === 'number') {
      cash = dsVals[3][7];
    } else if (dsVals.length > 1 && typeof dsVals[1][3] === 'number') {
      // Fallback to D2 full balance if H4 not found
      cash = dsVals[1][3];
    }
  }

  // -- Physical assets: Physical Assets!E7 --
  var physicalAssets = 0;
  var paSheet = ss.getSheetByName('Physical Assets');
  if (paSheet) {
    var paVals = readSheet(paSheet);
    // E7 = index 6, col E = index 4
    if (paVals.length > 6 && typeof paVals[6][4] === 'number') physicalAssets = paVals[6][4];
  }

  // -- Liabilities --
  var mortgage = 0, studentLoans = 0;
  // Mortgage: the current principal balance is the SINGLE source of truth, held in
  // the Loans ledger (last row flagged paid) and read via readMortgageLedger(). The
  // mortgage tab reads the same value, so the two views can never disagree. We no
  // longer read the hand-typed Physical Assets!B8 (see INDEX: mortgage reconciliation).
  var mort = readMortgageLedger(ss);
  if (!mort.error && typeof mort.curBalance === 'number') {
    mortgage = mort.curBalance;
  } else {
    // Degrade gracefully, never silently: a failed ledger read must NOT book $0 of
    // mortgage debt (that overstates net worth by the full balance, and _nwLooksValid
    // does not check liabilities so it would cache the bad result). Fall back to the
    // legacy B8 cell and log loudly.
    if (paSheet) {
      var paVals2 = readSheet(paSheet);
      if (paVals2.length > 7 && typeof paVals2[7][1] === 'number') mortgage = paVals2[7][1];
    }
    Logger.log('netWorth: mortgage ledger unavailable (' + (mort.error || 'no curBalance') + '); fell back to Physical Assets!B8 = ' + mortgage);
  }
  // Student loans: Loans!D26 (index 25, col D = index 3)
  var loansSheet = ss.getSheetByName('Loans');
  if (loansSheet) {
    var loansVals = readSheet(loansSheet);
    if (loansVals.length > 25 && typeof loansVals[25][3] === 'number') studentLoans = loansVals[25][3];
  }

  // -- Snapshots from Net Worth sheet --
  var snapshots = [];
  var nwSheet = ss.getSheetByName(NET_WORTH_SHEET);
  if (nwSheet) {
    var nwVals = readSheet(nwSheet);
    for (var i = 1; i < nwVals.length; i++) {
      var r = nwVals[i];
      var d = r[0];
      var dateStr;
      if (d instanceof Date) dateStr = d.toISOString().split('T')[0];
      else if (typeof d === 'number') dateStr = new Date((d-25569)*86400*1000).toISOString().split('T')[0];
      else dateStr = String(d||'').split('T')[0];
      if (!dateStr || dateStr.length < 8) continue;
      snapshots.push({
        date:           dateStr,
        investments:    typeof r[1]==='number' ? Math.round(r[1]) : null,
        cash:           typeof r[2]==='number' ? Math.round(r[2]) : null,
        physicalAssets: typeof r[3]==='number' ? Math.round(r[3]) : null,
        mortgage:       typeof r[4]==='number' ? Math.round(r[4]) : null,
        studentLoans:   typeof r[5]==='number' ? Math.round(r[5]) : null,
        carLoans:       typeof r[6]==='number' ? Math.round(r[6]) : null,
        fivetwonine:    typeof r[7]==='number' ? Math.round(r[7]) : null,
        netWorth:       typeof r[8]==='number' ? Math.round(r[8]) : null,
      });
    }
  }

  // -- Compute current net worth --
  var totalAssets     = investments + cash + physicalAssets;
  var totalLiabilities= mortgage + studentLoans;
  var netWorth        = Math.round((totalAssets - totalLiabilities) * 100) / 100;

  // -- 529 totals (separable accounts) --
  var fivetwonine = Math.round(((accounts['NY State 529 - Wesley']||0) + (accounts['NY State 529 - Maxwell']||0)) * 100) / 100;

  // -- $1M projections at 6%, 8%, 10% nominal annual return --
  var PROJECTION_TARGET = cfg().projectionTarget;
  var projections = [];
  if (history.length > 0) {
    var currentVal  = history[history.length - 1].investments || 0;
    var currentDate = new Date(history[history.length - 1].date);

    // TTM deposits: sum col D (index 3) from weekly tracker over past 52 weeks
    var ttmCutoff  = new Date(currentDate.getTime() - 365 * 24 * 60 * 60 * 1000);
    var ttmDeposits = 0;
    var inDep = false;
    for (var di = 0; di < values.length; di++) {
      var dr = values[di];
      if (!inDep) {
        if (String(dr[0]||'').trim() === 'Net worth tracker' ||
            String(dr[1]||'').trim() === 'Date') { inDep = true; }
        continue;
      }
      var depDate = dr[1], dep = parseFloat(dr[3]);
      if (!depDate || isNaN(dep) || dep <= 0) continue;
      var depD = depDate instanceof Date ? depDate : new Date((depDate - 25569) * 86400 * 1000);
      if (depD >= ttmCutoff) ttmDeposits += dep;
    }
    var weeklyContrib = ttmDeposits / 52;

    [0.06, 0.08, 0.10].forEach(function(rate) {
      var weeklyRate = Math.pow(1 + rate, 1/52) - 1;
      var val = currentVal;
      var weeks = 0;
      var points = [{ date: currentDate.toISOString().split('T')[0], value: Math.round(val) }];
      while (val < PROJECTION_TARGET && weeks < 52 * 15) {
        val = val * (1 + weeklyRate) + weeklyContrib;
        weeks++;
        if (weeks % 4 === 0) {
          var pd = new Date(currentDate.getTime() + weeks * 7 * 24 * 60 * 60 * 1000);
          points.push({ date: pd.toISOString().split('T')[0], value: Math.round(val) });
        }
      }
      var milestoneDate = null;
      if (val >= PROJECTION_TARGET) {
        var md = new Date(currentDate.getTime() + weeks * 7 * 24 * 60 * 60 * 1000);
        milestoneDate = md.toISOString().split('T')[0];
      }
      projections.push({
        rate:          rate,
        label:         Math.round(rate * 100) + '%',
        points:        points,
        milestone:     milestoneDate,
        weeklyContrib: Math.round(weeklyContrib),
      });
    });
  }

  return {
    investments:     Math.round(investments * 100) / 100,
    cash:            Math.round(cash * 100) / 100,
    physicalAssets:  Math.round(physicalAssets * 100) / 100,
    mortgage:        Math.round(mortgage * 100) / 100,
    studentLoans:    Math.round(studentLoans * 100) / 100,
    carLoans:        0, // not tracked in sheet yet
    fivetwonine:     fivetwonine,
    liquid:          Math.round(liquid * 100) / 100,
    totalAssets:     Math.round(totalAssets * 100) / 100,
    totalLiabilities:Math.round(totalLiabilities * 100) / 100,
    netWorth:        netWorth,
    accounts:        accounts,
    history:         history,
    snapshots:       snapshots,
    projections:     projections,
    snapshotExists:  !!nwSheet,
    codeVersion:     'nw-2026-06-16-mortgage-ledger',
  };
}

// Create the Net Worth snapshot sheet if it doesn't exist
function ensureNetWorthSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NET_WORTH_SHEET);
  if (sheet) return sheet;
  sheet = ss.insertSheet(NET_WORTH_SHEET);
  var headers = ['Date','Investments','Cash','Physical Assets','Mortgage','Student Loans','Car Loans','529 Balance','Net Worth'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1a3a5c').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 110);
  for (var i = 2; i <= headers.length; i++) sheet.setColumnWidth(i, 120);
  sheet.getRange(1, 1, 1, headers.length).setNumberFormat('@');
  SpreadsheetApp.flush();
  return sheet;
}

// Save a net worth snapshot row
function saveNetWorthSnapshot(p) {
  var sheet = ensureNetWorthSheet();
  var investments  = parseFloat(p.investments)  || 0;
  var cash         = parseFloat(p.cash)         || 0;
  var physical     = parseFloat(p.physicalAssets)|| 0;
  var mortgage     = parseFloat(p.mortgage)     || 0;
  var studentLoans = parseFloat(p.studentLoans) || 0;
  var carLoans     = parseFloat(p.carLoans)     || 0;
  var fivetwonine  = parseFloat(p.fivetwonine)  || 0;
  var netWorth     = investments + cash + physical - mortgage - studentLoans - carLoans;
  var dateObj      = p.date ? new Date(p.date + 'T12:00:00') : new Date();
  var row = [dateObj, investments, cash, physical, mortgage, studentLoans, carLoans, fivetwonine || '', Math.round(netWorth * 100) / 100];
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);
  sheet.getRange(lastRow + 1, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(lastRow + 1, 2, 1, row.length - 1).setNumberFormat('$#,##0');
  SpreadsheetApp.flush();
  return { ok: true, netWorth: Math.round(netWorth * 100) / 100, row: lastRow + 1 };
}


// ============================================================
// HSA RECEIPTS -- reimbursement-entitlement ledger
// The 'HSA Receipts' sheet holds one row per qualified medical expense. An
// out-of-pocket (funding=OOP) expense that has not been reimbursed is a
// tax-free withdrawal claim on the HSA; the receipt is its substantiation.
// Two jobs: total the outstanding claims (the unreimbursed pool) and make each
// claim reimbursable EXACTLY once (idempotency + the reimburse guards).
// Sheet layout (see INDEX section 8): B1 = current HSA balance (manual, blank
// if unknown), B2 = balance as-of date, header row 3, data rows 4+. Columns
// A..K: id, date_incurred, amount, provider, description, category, funding,
// receipt_link, reimbursed_amount, reimbursed_date, notes.
// ============================================================
var HSA_SHEET         = 'HSA Receipts';
var HSA_DATA_ROW      = 4;          // 1-indexed first data row (headers on row 3)
var HSA_CACHE_VERSION = 'hsa_v1';   // bump on getHsa response-shape change
var HSA_CACHE_KEY     = HSA_CACHE_VERSION + '_data';

function _hsaR2(x) { var f = parseFloat(x); return isNaN(f) ? 0 : Math.round(f * 100) / 100; }

function _hsaDateStr(v) {
  if (v instanceof Date && !isNaN(v.getTime()))
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var s = String(v == null ? '' : v).trim();
  return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(0, 10) : '';
}

// Pure rollup -- no sheet access, so testHsa() can call it with a fixture.
// rows: [{ id, dateIncurred, amount, provider, description, category, funding,
//          receiptLink, reimbursedAmount, reimbursedDate }]
// established: 'YYYY-MM-DD' or null (gate OFF when null -- cannot disqualify what
//   we cannot compare; the response flags established:null so the UI warns).
// hsaBalance: number or null (dependent metrics null when balance unknown).
function _hsaRollups(rows, established, hsaBalance) {
  var unreimbursed = 0, totalSubstantiated = 0;
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r       = rows[i];
    var amount  = _hsaR2(r.amount);
    var reimb   = _hsaR2(r.reimbursedAmount);
    var funding = (String(r.funding || 'OOP').toUpperCase() === 'HSA') ? 'HSA' : 'OOP';
    var qualified = true;
    if (established && r.dateIncurred) qualified = (String(r.dateIncurred) >= String(established));
    var status;
    if (funding === 'HSA')      status = 'paid_direct';
    else if (reimb <= 0)        status = 'open';
    else if (reimb < amount)    status = 'partial';
    else                        status = 'reimbursed';
    if (qualified) {
      totalSubstantiated += amount;
      if (funding === 'OOP') unreimbursed += (amount - reimb);
    }
    out.push({
      id:               r.id,
      dateIncurred:     r.dateIncurred || '',
      amount:           amount,
      provider:         r.provider || '',
      description:      r.description || '',
      category:         r.category || '',
      funding:          funding,
      receiptLink:      r.receiptLink || '',
      reimbursedAmount: reimb,
      reimbursedDate:   r.reimbursedDate || '',
      status:           status,
      qualified:        qualified
    });
  }
  unreimbursed       = _hsaR2(unreimbursed);
  totalSubstantiated = _hsaR2(totalSubstantiated);
  var reimbursableNow = (hsaBalance == null) ? null : _hsaR2(Math.min(unreimbursed, hsaBalance));
  var stranded        = (hsaBalance == null) ? null : _hsaR2(Math.max(unreimbursed - hsaBalance, 0));
  return {
    rows:                out,
    unreimbursed:        unreimbursed,
    totalSubstantiated:  totalSubstantiated,
    reimbursableNow:     reimbursableNow,
    strandedEntitlement: stranded
  };
}

// Reimburse guards -- pure, returns an error string or null. Tested directly.
// row must carry { amount (number), funding, qualified }.
function _reimburseGuard(row, reimbursedAmount) {
  if (!row) return 'Receipt not found.';
  if (String(row.funding).toUpperCase() === 'HSA')
    return 'This expense was paid with the HSA card; it is not reimbursable.';
  if (!row.qualified)
    return 'This expense predates the HSA establishment date and does not qualify.';
  var a = parseFloat(reimbursedAmount);
  if (isNaN(a) || a < 0) return 'Reimbursement amount must be zero or more.';
  if (a > _hsaR2(row.amount) + 0.005) return 'Reimbursement exceeds the expense amount.';
  return null;
}

// Read the HSA Receipts data rows (row 4+) into parsed objects. Empty table
// (lastRow < 4) returns [] without ever calling getRange with 0 rows.
function _hsaReadRows(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < HSA_DATA_ROW) return [];
  var n      = lastRow - (HSA_DATA_ROW - 1);
  var values = sheet.getRange(HSA_DATA_ROW, 1, n, 11).getValues();
  var rows   = [];
  for (var i = 0; i < values.length; i++) {
    var v  = values[i];
    var id = v[0];
    if ((id === '' || id == null) && (v[2] === '' || v[2] == null)) continue; // blank spacer
    rows.push({
      id:               (typeof id === 'number') ? id : String(id),
      dateIncurred:     _hsaDateStr(v[1]),
      amount:           v[2],
      provider:         v[3],
      description:      v[4],
      category:         v[5],
      funding:          v[6],
      receiptLink:      (v[7] == null) ? '' : String(v[7]),
      reimbursedAmount: v[8],
      reimbursedDate:   _hsaDateStr(v[9])
    });
  }
  return rows;
}

// -- GET: HSA receipts + rollups -------------------------------
function getHsa() {
  var cache = CacheService.getScriptCache();
  try { var hit = cache.get(HSA_CACHE_KEY); if (hit) return JSON.parse(hit); } catch (e) {}

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HSA_SHEET);
  if (!sheet) return { error: 'Sheet not found: ' + HSA_SHEET };

  // B1 = balance (blank => unknown => null, NEVER 0); B2 = as-of date.
  var b1          = sheet.getRange(1, 2).getValue();
  var hsaBalance  = (b1 === '' || b1 == null || isNaN(parseFloat(b1))) ? null : _hsaR2(b1);
  var balanceAsOf = _hsaDateStr(sheet.getRange(2, 2).getValue()) || null;

  var established = cfg().hsaEstablished || null;
  var roll = _hsaRollups(_hsaReadRows(sheet), established, hsaBalance);

  var result = {
    established:         established,
    hsaBalance:         hsaBalance,
    balanceAsOf:        balanceAsOf,
    unreimbursed:       roll.unreimbursed,
    reimbursableNow:    roll.reimbursableNow,
    totalSubstantiated: roll.totalSubstantiated,
    strandedEntitlement: roll.strandedEntitlement,
    rows:               roll.rows
  };
  try { cache.put(HSA_CACHE_KEY, JSON.stringify(result), 300); } catch (e) {}
  return result;
}

function invalidateHsaCache() {
  try { CacheService.getScriptCache().remove(HSA_CACHE_KEY); } catch (e) {}
}

// -- POST: reimburse a receipt (absolute reimbursed_amount) ----
// Sets reimbursed_amount to an ABSOLUTE value (idempotent-friendly: a retry
// writes the same number). Locates the row BY id (linear scan), never by
// position, and only updates cells in place -- so it cannot hit the stale-row
// -index bug class. Wrapped in _withIdem in both routers.
function reimburseReceipt(p) {
  var id = (p.id == null) ? '' : String(p.id);
  if (!id) return { error: 'Missing receipt id.' };

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HSA_SHEET);
  if (!sheet) return { error: 'Sheet not found: ' + HSA_SHEET };

  invalidateSheet(HSA_SHEET);
  var values = readSheet(sheet);
  var rowNum = -1, amount = 0, funding = 'OOP', dateIncurred = '';
  for (var i = HSA_DATA_ROW - 1; i < values.length; i++) {
    if (String(values[i][0]) === id) {
      rowNum       = i + 1;
      amount       = _hsaR2(values[i][2]);
      funding      = String(values[i][6] || 'OOP');
      dateIncurred = _hsaDateStr(values[i][1]);
      break;
    }
  }
  if (rowNum < 0) return { error: 'Receipt not found: ' + id };

  var established = cfg().hsaEstablished || null;
  var qualified   = !(established && dateIncurred) || (String(dateIncurred) >= String(established));
  var guard = _reimburseGuard({ amount: amount, funding: funding, qualified: qualified }, p.reimbursedAmount);
  if (guard) return { error: guard };

  var amt  = _hsaR2(p.reimbursedAmount);
  var dStr = _hsaDateStr(p.reimbursedDate) || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  sheet.getRange(rowNum, 9).setValue(amt).setNumberFormat('$#,##0.00');
  if (amt > 0) sheet.getRange(rowNum, 10).setValue(new Date(dStr + 'T12:00:00')).setNumberFormat('yyyy-mm-dd');
  else         sheet.getRange(rowNum, 10).clearContent(); // un-reimburse -> clear the date
  SpreadsheetApp.flush();
  invalidateHsaCache();
  return { ok: true, id: id, reimbursedAmount: amt, reimbursedDate: (amt > 0 ? dStr : '') };
}

// -- POST: add a new receipt row -------------------------------
// Plain append -- HSA Receipts has no formulas or sections below the data
// block (verified against the live sheet), so this is a single setValues
// call, not the insertCells/SUM-repair dance the monthly sheets need.
// LockService guards id assignment: two near-simultaneous adds (e.g. both
// household members logging a receipt at once) must never compute the same
// "next id" -- a duplicate id would corrupt the by-id lookup reimburseReceipt
// depends on. Wrapped in _withIdem in both routers for retry-safety.
function addHsaReceipt(p, sheetNameOverride) {
  var amount = parseFloat(p.amount);
  if (isNaN(amount) || amount <= 0) return { error: 'Amount must be a positive number.' };
  var dateStr = String(p.dateIncurred || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return { error: 'Invalid or missing date incurred.' };
  var funding = (String(p.funding || 'OOP').toUpperCase() === 'HSA') ? 'HSA' : 'OOP';
  var receiptLink = String(p.receiptLink || '').trim();
  if (receiptLink && !/^https?:\/\//i.test(receiptLink))
    return { error: 'Receipt link must start with http:// or https://' };

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetNameOverride || HSA_SHEET);
  if (!sheet) return { error: 'Sheet not found: ' + (sheetNameOverride || HSA_SHEET) };

  var lock = LockService.getScriptLock();
  var gotLock = false;
  try { gotLock = lock.tryLock(5000); } catch (e) {}
  if (!gotLock) return { error: 'Could not acquire lock -- please try again.' };

  try {
    invalidateSheet(sheetNameOverride || HSA_SHEET);
    var values = readSheet(sheet);
    var maxId = 0;
    for (var i = HSA_DATA_ROW - 1; i < values.length; i++) {
      var n = parseInt(values[i][0], 10);
      if (!isNaN(n) && n > maxId) maxId = n;
    }
    var newId = maxId + 1;
    var insertRow = Math.max(sheet.getLastRow() + 1, HSA_DATA_ROW);
    // Col L (12) = source_file_id: the Drive fileId when created by the folder
    // scan, blank for manual adds. This is the dedup key scanHsaFolder reads to
    // avoid re-importing a file. Written in the same setValues as the id so it
    // lands atomically under the lock.
    var row = [
      newId, new Date(dateStr + 'T12:00:00'), amount,
      String(p.provider || '').trim(), String(p.description || '').trim(),
      String(p.category || '').trim(), funding, receiptLink, '', '',
      String(p.notes || '').trim(), String(p.sourceFileId || '')
    ];
    sheet.getRange(insertRow, 1, 1, row.length).setValues([row]);
    sheet.getRange(insertRow, 1).setNumberFormat('0');
    sheet.getRange(insertRow, 2).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(insertRow, 3).setNumberFormat('$#,##0.00');
    SpreadsheetApp.flush();
    invalidateHsaCache();
    return { ok: true, id: newId, row: insertRow };
  } finally {
    lock.releaseLock();
  }
}

// Parse a receipt filename of the form  DATE~provider~amount.pdf  into fields.
// Pure -- unit-tested by testParseReceiptName. Delimiter is '~' (NOT '_' or
// '-') so provider may contain spaces, underscores, and hyphens. Strict:
// returns { error } on anything that doesn't match exactly, so the scan SKIPS
// (and reports) a malformed name rather than guessing a wrong date/amount.
function _parseReceiptName(fname) {
  var base = String(fname || '').replace(/\.pdf$/i, '');
  var parts = base.split('~');
  if (parts.length !== 3) return { error: 'name must be DATE~provider~amount.pdf' };
  var date     = parts[0].trim();
  var provider = parts[1].trim();
  var amtStr   = parts[2].trim().replace(/[$,]/g, '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { error: 'bad date (need YYYY-MM-DD)' };
  if (!provider) return { error: 'missing provider' };
  var amount = parseFloat(amtStr);
  if (isNaN(amount) || amount <= 0) return { error: 'bad amount' };
  return { date: date, provider: provider, amount: amount };
}

// -- POST / trigger: import new receipt PDFs from the Drive folder ----------
// Takes no args so a future time-based trigger can call it directly. Idempotent
// by construction: each created row stores its Drive fileId in col L, and a file
// whose id is already present is skipped -- so re-running (button re-click,
// retry, or hourly trigger) never double-imports. A script lock serializes
// scans so two concurrent runs can't both see the same file as "new". Strict
// filename parsing: unmatched names are skipped and returned in `skipped` with a
// reason (fail loud, never guess). Auto-created rows default funding=OOP (the
// safe default -- the receipt shows up as a reimbursable claim, not silently
// dropped); category/description stay blank for you to fill in.
function scanHsaFolder() {
  var folderId = cfg().hsaReceiptFolder;
  if (!folderId) return { error: 'Set hsa_receipt_folder in Config (the Drive folder ID).' };

  var folder;
  try { folder = DriveApp.getFolderById(folderId); }
  catch (e) { return { error: 'Cannot open Drive folder (check the ID and that the script can access it): ' + e.message }; }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HSA_SHEET);
  if (!sheet) return { error: 'Sheet not found: ' + HSA_SHEET };

  var lock = LockService.getScriptLock();
  var gotLock = false;
  try { gotLock = lock.tryLock(20000); } catch (e) {}
  if (!gotLock) return { error: 'A scan is already running -- try again in a moment.' };

  try {
    invalidateSheet(HSA_SHEET);
    var values = readSheet(sheet);
    var seen = {};
    for (var i = HSA_DATA_ROW - 1; i < values.length; i++) {
      var fid = values[i][11]; // col L
      if (fid) seen[String(fid)] = true;
    }

    var created = [], skipped = [];
    var it = folder.getFilesByType('application/pdf');
    while (it.hasNext()) {
      var file  = it.next();
      var theId = file.getId();
      var fname = file.getName();
      if (seen[theId]) continue; // already imported -- silent (not noise)

      var parsed = _parseReceiptName(fname);
      if (parsed.error) { skipped.push({ name: fname, reason: parsed.error }); continue; }

      var res = addHsaReceipt({
        dateIncurred: parsed.date,
        amount:       parsed.amount,
        provider:     parsed.provider,
        funding:      'OOP',
        receiptLink:  file.getUrl(),
        sourceFileId: theId
      });
      if (res.error) { skipped.push({ name: fname, reason: res.error }); continue; }

      seen[theId] = true;
      created.push({ name: fname, id: res.id });
    }

    invalidateHsaCache();
    return { ok: true, created: created, createdCount: created.length, skipped: skipped, skippedCount: skipped.length };
  } finally {
    lock.releaseLock();
  }
}
