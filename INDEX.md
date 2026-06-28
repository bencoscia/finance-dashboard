# INDEX — Household Finance Dashboard

**Read this first.** It maps the files, states the invariants, and lists the exact touchpoints to update when you add things. Keep it accurate: when this drifts from the code, the code wins — fix this file. Where a number here can rot (counts, dates), it's marked *(verify)*.

---

## 1. Architecture in one breath

A **Google Sheet is the single source of truth.** A **single-file HTML dashboard** (`index.html`) is a read/write *view* of it; it holds no authoritative state. A **Google Apps Script web app** (`apps_script.gs`) is the only thing that touches the sheet: `doGet` serves reads, `doPost` serves writes, both returning JSON. The dashboard talks to the deployed Web App URL (`SCRIPT_URL`) via `fetch`. `SCRIPT_URL` is loaded from a **gitignored `config.js`** (`window.CONFIG.SCRIPT_URL`); a saved URL in `localStorage['fin_url']` overrides it, and when neither is present the boot path shows the setup screen.

```
Google Sheet  ◄──►  apps_script.gs (Web App: doGet reads / doPost writes)  ◄──fetch──►  index.html (browser)
   (truth)            tests.gs · debug.gs · migrate.gs live in same Apps Script project
```

**Source of truth for code + docs: the public GitHub repo `bencoscia/finance-dashboard`.** `index.html` is served from **GitHub Pages** off that repo. The backend lives in a separate Apps Script project; the repo's `.gs` files are reference copies, not the live backend. Claude reads the live files via `raw.githubusercontent.com` (CDN-cached ~5 min); pushes are manual — you commit. `config.js` (the `SCRIPT_URL` bearer secret) is gitignored and never in the repo.

---

## 2. File catalog

| File | Size~ | Role | Notes |
|---|---|---|---|
| `index.html` | ~285 KB | Entire frontend: HTML + CSS + JS in one file; Chart.js 4.4.1 via CDN | Served from GitHub Pages. Self-contained, no build step. `SCRIPT_URL` comes from `config.js`, not baked in. |
| `apps_script.gs` | 116 KB | Backend Web App over the Sheet | Deployed as a Web App. **Changes here require a redeploy** (§7). |
| `tests.gs` | 28 KB | Apps Script test suite, run from the editor | Fast suite: idempotency, config parsing, month-cache tiers, etc. |
| `debug.gs` | 9 KB | Diagnostic fns, run manually from the editor | Never called by the dashboard. |
| `migrate.gs` | 18 KB | One-time setup/migration fns | Idempotent unless a fn says otherwise (e.g. `setupConfigSheet`). |
| `audit.md` | 22 KB | Standing audit: bug history, invariants, resolved + carried-forward items | The "why" record. Authoritative on past decisions. |
| `net_worth_plan.md` | 4 KB | Design plan: Net Worth tab → FI tracking | Roadmap, not yet fully built. |
| `transcript-2026-06-14.md` | 34 KB | Prior working-session transcript | History/reference only. |
| `INDEX.md` | — | This map + consistency rules | Update on every structural change. **In the repo.** |
| `config.js` | — | Holds `window.CONFIG.SCRIPT_URL` (the Web App URL) | **Gitignored, local-only.** Bearer secret — never commit/serve. Absent on Pages → setup screen collects the URL into localStorage. |
| `config.example.js` | — | Placeholder template for `config.js` | Committed. Copy → `config.js`, fill in the URL. |
| `.gitignore` | — | Excludes `config.js` | Committed. |

`index.html` is served from **GitHub Pages** off the public repo `bencoscia/finance-dashboard` (the source of truth for code + docs). The `.gs` files are separate files **inside one Apps Script project** (the repo holds reference copies). `config.js` is **never committed** (gitignored).

---

## 3. Invariants — do not violate

1. **The Sheet is truth.** The dashboard never invents state; it reads from / writes to the Sheet. Client caches (`allData`, `nwData`, `loanData`, `mortgageData`) are views and must be invalidated when a write changes them.
2. **Zero inline `on*=` attributes in HTML.** No `onclick="…"` (or any `on…=`) in rendered markup. Interactions go through `data-action`/`data-args` + the document-level `ACTION_MAP` dispatcher, or through `.onclick`/`.oninput` **property closures**. The JS-string-inside-HTML-attribute escaping bug class caused three outages and is now structurally extinct — do not reintroduce it. (`.onclick`/`.oninput` property assignments are *allowed* — real closures, never the hazard; audit.md's count of "seven" is stale, it's ~12 now *(verify)*.)
3. **Dual-router rule.** A read action must be registered in **both** dispatchers in `apps_script.gs`: the `doGet` chain (~L137) **and** the unified action chain (~L93). A write goes in `doPost` (~L146) and, if it mutates money or rows, is wrapped in `_withIdem`.
4. **No silent staleness.** Any value interpolated into already-rendered HTML must be re-rendered by every write that changes it; cross-tab caches affecting net worth must be nulled on relevant writes.
5. **Protan-safe visuals (Ben is protanopic).** Never encode meaning by hue alone. Separate series by **lightness, dash pattern, axis, shape, and a text label** — at least two non-color channels, always.
6. **`.gs` is ES5 + ASCII only.** `var`/`function(){}`, no arrow fns, no template literals, no non-ASCII characters. `index.html` JS follows the same `var`/`function`/string-concatenation style; no template literals in rendered HTML strings.
7. **Chart.js discipline.** Destroy the instance before re-rendering (`if(cX){cX.destroy();cX=null;}`). There is **no time-scale adapter** loaded — use category-string labels + `maxTicksLimit`. Chart.js cannot read CSS custom properties — resolve them with `getComputedStyle` (helper: `cssv(name, fallback)`).
8. **Bump cache versions on shape change.** Change a cached payload's fields → bump the matching `CACHE_VERSION` (`monthly_v3`) or `CARD_CACHE_VERSION` (`cards_v1`).

---

## 4. Traps / failure modes (each has bitten before)

- **Deployment skew.** "New deployment" mints a *new URL that serves the old code forever*. Always **pencil-edit the existing deployment → New version** (same URL). Symptom of a stale deploy: `Unknown action: <x>` from a read/write you just added.
- **Dual-router drift.** Forgetting to add a read to *both* chains (Invariant 3) → works via one entry point, fails via the other.
- **Interpolated-UI staleness.** A write updates the Sheet but not the baked-in HTML/caches → the tab shows old numbers until reload.
- **`--blue` is green.** `--blue` = `--green` = `#1E5C46`. The actual blue is `--c2` `#3E6B8F`. Misnamed token — read §6 before picking colors.
- **Mortgage balance — one source (resolved 2026-06-16).** The `Loans!M–W` ledger is the source of truth; `Physical Assets!B8` is now a **formula** mirroring its current balance — `=XLOOKUP(TRUE, Loans!N8:N400, Loans!O8:O400, "NO PAID ROW", 0, -1)` (balance on the last `TRUE`-flagged row). Net worth reads B8, the mortgage tab reads the ledger, and they can't drift because B8 derives from the ledger. The one manual input is the **col-N paid flag** — tick it the month a payment posts. **Do not overtype B8 with a literal** — that re-creates the old hand-maintained duplicate (the ~$343 / one-month drift). Two Sheets gotchas: the Excel `LOOKUP(2,1/(…))` last-match idiom returns `#N/A` here — use XLOOKUP; and bound the range to the schedule’s last row so a stray `TRUE` in another loan block below can’t be matched (XLOOKUP `-1` searches bottom-up). If B8 ever reads `NO PAID ROW`, net worth silently books $0 mortgage.
- **`config.js` isn't served on Pages.** It's gitignored, so on GitHub Pages `window.CONFIG` is undefined and `SCRIPT_URL` is empty — by design. The setup screen is the Pages onboarding (paste the URL once → saved to that browser's `localStorage['fin_url']`). `bootDashboard()` **must** fall back to `showSetup()` when there's no URL; a missing `else` there once left Pages blank-white. Works locally because `config.js` is present.
- **No client-side password gate.** One was added and removed — it's theater: page source (hence any hash) is world-readable on Pages, and `SCRIPT_URL`, not the gate, protects the data. Don't reintroduce. Real access control = `SCRIPT_URL` secrecy now, a server-validated token later (§9).
- **`SCRIPT_URL` precedence:** `localStorage['fin_url']` (set via setup) wins over `config.js`. A stale saved URL silently overrides a fresh `config.js` — clear it via the setup screen if you repoint.
- **HSA double-reimbursement (cardinal risk).** Paying out a receipt twice is an over-withdrawal the IRS can claw back. Two layers: `reimburseReceipt` is `_withIdem`-wrapped **and** writes `reimbursed_amount` as an **absolute** value (a retry writes the same number, never adds), plus a server `_reimburseGuard` rejecting `> amount`, negative, `funding=HSA`, and pre-establishment rows. The UI "Reimburse to total ($)" field is the absolute total, not an increment — keep it so.
- **HSA establishment gate.** Only expenses on/after `Config!hsa_established` qualify; earlier rows are flagged ⚠ and excluded from every total. **Blank `hsa_established` → gate OFF** (nothing disqualified) and the tab shows a ⚠ banner — set it or totals overstate.
- **HSA balance is `B1`, blank ≠ 0.** `HSA Receipts!B1` (manual) is the only balance source; **blank → `null` → renders "—", never `$0`**, and `reimbursableNow`/`strandedEntitlement` stay `null` (not computed against 0). HSA is **not** in net worth yet (§9). Receipts store a **Drive link**, never an embedded image.
- **`cfg()` cache is shape-versioned (`CFG_CACHE_KEY`, 5-min TTL) — bump it on every Config-shape change, not just `getNetWorth`'s.** Adding `hsaEstablished` without bumping `cfg_v1`→`cfg_v2` meant a config cached moments before the redeploy kept serving the old shape (missing the field) for up to 5 min — surfaced as "No HSA establishment date set" even with the sheet row correct. Same root cause as the documented `nw_v1`→`nw_v3` incident (§ bug class 2), now also true of `cfg()`. (Bumped again to `cfg_v3` for `hsaReceiptFolder`.) `testConfigParsing` asserts the round-trips, but it can't catch a forgotten version bump — that's a manual checklist item.
- **`HSA_CACHE_VERSION` bumped `hsa_v1`→`hsa_v2`** when `getHsa` rows gained `needsReview`. Same rule as the `cfg()`/`nw` trap above: any change to a cached payload's shape needs the version bump or stale reads persist up to the TTL.
- **`needs_review` drafts are deliberately invisible to totals.** A scanned-but-malformed receipt is present (and visible, sorted to top with ✎) but contributes **$0** to `unreimbursed`/`totalSubstantiated`/the pool until you confirm it via Edit. Intentional (don't trust unverified data) — but it means "I scanned it and the pool didn't move" is expected, not a bug.
- **The HSA projection is a two-engine, nominal-dollar model.** Engine 1 (recurring tax avoided = spend×(tc+f)) is certain and market-independent; Engine 2 (the hoard) is the only market-dependent part and only matters on receipts you *keep* (don't reimburse). Don't conflate them, and don't read the old "$8k/yr for 30 yrs, hold forever" framing back in — that was ~80× too high because it used the contribution limit instead of the real claim rate and never reimbursed. `hsaProjCompute` is pure and cross-checked offline; at the demo inputs ($2,800/$500/20/30) it must produce Engine 1 $1,054/yr and Engine 2 HSA $43,145 / IRA $31,078 / taxable $22,469. Spend/hoard auto-fill from the ledger TTM; a user override (saved in `localStorage['hsa_calc_v2']`) wins over the TTM prefill.
- **Folder scan: filename is load-bearing data.** `scanHsaFolder` parses `DATE~provider~amount.pdf` into the row's date/provider/amount — a filename typo becomes a wrong ledger entry. `_parseReceiptName` is strict and **skips + reports** anything off-pattern rather than guessing. Delimiter is `~` (provider may contain spaces/underscores; it may **not** contain `~`).
- **Folder scan dedups on Drive `fileId` (col L), not on content** — so a receipt entered *both* manually and via the scan creates **two** rows (the manual one has no fileId to match). Pick one path per receipt; the scan's created/skipped report makes a surprise duplicate visible immediately. Auto-created rows are `funding=OOP` (safe default — surfaces as a reimbursable claim) with blank category/description to fill in.
- **`DriveApp` needs the Drive OAuth scope.** Adding the folder scan introduces a new scope; the first `scanHsaFolder` run prompts re-authorization. After deploying, run it once from the editor (or via the button) and approve, or the web app returns an auth error.

---

## 5. Server reference (`apps_script.gs`)

**Read actions (15)** — registered in both chains, served by `doGet`:
`all · quick · monthly · networth · transactions · fixed · variableEntries · accounts · cardTotals · cardBalances · cardTxns · config · loans · mortgage · hsa`

**Write actions (22)** — `doPost` (★ = `_withIdem`-wrapped):
`addTransaction★ · updateTransaction · deleteTransaction · splitTransaction · setFixedPaid · updateFixedCost · updateFixedName · updateLoanBalance★ · makeCardPayment★ · voidLastPayment · setSeedBalance · addVariableEntry★ · updateVariableEntry · deleteVariableEntry · makeCheckingEntry★ · saveNetWorthSnapshot · logCreditSnapshot · addMonth · reimburseReceipt★ · addHsaReceipt★ · scanHsaFolder · updateHsaReceipt★`

`scanHsaFolder` is **not** `_withIdem`-wrapped — it's idempotent by construction (dedup on Drive `fileId` stored in `HSA Receipts` col L), takes no args (so a future time-trigger can call it directly), and serializes via `LockService`. Reads the Drive folder named in `Config!hsa_receipt_folder` (requires the Drive OAuth scope — first run prompts re-authorization). A malformed filename is imported as a **draft** (`needs_review`, col M) via best-effort parse rather than skipped; only a name with neither a date nor an amount is skipped. `updateHsaReceipt` is the general edit path (fix any field; clears `needs_review`); reimbursed amount/date stay on `reimburseReceipt` so that remains the single money-moving path.

**Sheets touched:** `Loans`, `Physical Assets`, `Credit`, `Discover Savings`, `Portfolio Management`, `Net Worth`, `Card Payments`, `Card Trackers`, `Config`, `HSA Receipts`, `Template`, and the monthly sheets (`"<Month> YYYY"`).

**Key helpers:** `readSheet(name|sheet)` (cached per-request 2-D values) · `_withIdem(payload, fn)` (idempotency keys in CacheService, 6 h) · `cfg()` (reads `Config`, 5-min cache under shape-versioned `CFG_CACHE_KEY` — bump on any shape change, see §4 trap) · cache invalidation per month/card/net-worth/HSA.

---

## 6. Client reference (`index.html`)

**Tabs (9):** `thisweek` (landing; driven by `all`/`quick` + Quick Add) · `accounts` · `txns` · `fixed` (Recurring expenses) · `budget` (merged Monthly+Breakdown) · `networth` · `loans` (**label: "Student Loans"**, id stays `loans`) · `mortgage` · `hsa`.

**Quick Add (6 types):** `txn · groceries · gas · payment · checking · hsa`. Global trigger (`+ Add`, header) defaults to `txn`; the HSA tab's own `+ Add receipt` button opens the same modal preset to `hsa` (`showQuickAdd(e)` reads `data-qa-type` off the clicked element, falling back to `txn`). The `hsa` type is the only one not gated by `qaCurMonth()` — it writes to `HSA Receipts`, not a monthly sheet — and is the only type with an explicit demo-mode guard (the other 5 currently lack one — pre-existing gap, not yet fixed). The HSA tab also has a **Scan folder** button (`scanHsaReceipts` → POST `scanHsaFolder`); created/skipped counts show as a banner and the skip list persists as a `hsa-note` via `hsaScanReport` until the next scan.

**Loaders:** `loadAccounts · loadFixed · loadLoans · loadMortgage · loadNetWorth · loadTxns · loadHsa · loadSaved` (+ initial `all` fetch). Tab lazy-load is wired in `showTab(id)` (`if(id==='x') loadX();`).

**Boot:** entry point `bootDashboard()` (called once at load). URL precedence: `localStorage['fin_url']` > `window.CONFIG.SCRIPT_URL` (from `config.js`); if empty → `showSetup()`. This replaced a client-side password/unlock gate that was added then removed (§4).

**Demo mode:** `isDemo` gates every loader to a `demoX()` synthesizer: `demoData · demoAccounts · demoTxns · demoFixed · demoVarEntries · demoNetWorth · demoCardTotals · demoCardBalances · demoLoans · demoMortgage · demoHsa`.

**Interaction model:** one delegated `click` listener → `ACTION_MAP[action]` or generic `window[action](...data-args)`; per-row controls carry `data-action`/`data-args` (numbers auto-parsed). No inline `on*=` (Invariant 2).

**Chart instances:** `cMonthly · cBreakdown · cNW · cNWSnap · cCredit · cLoanProj · cMtgHist · cMtgProj · cHsaProj`. Each destroyed before re-render.

**HSA tab specifics.** Ledger render `renderHsa` floats `needsReview` drafts to the top (✎) and puts **Edit + Reimburse** buttons per row. `showHsaEditModal(id)`/`hideHsaEditModal`/`submitHsaEdit` drive the edit modal (`#hsa-edit-backdrop`) → `updateHsaReceipt`; this is the in-dashboard correction path for both malformed-import drafts and any wrong field, and clears `needs_review`. Module vars: `hsaData · hsaReimburseId · hsaEditId · hsaShowOpenOnly · hsaScanReport`. The **projection calculator** is a two-engine HSA value model, independent of receipt data but auto-filling from it: `hsaCalcInit()` (idempotent — attaches input listeners once, restores inputs from `localStorage['hsa_calc_v2']`, else `HSA_CALC_DEFAULTS`), `hsaTtm()` (reads trailing-12-month routed **spend** = HSA-card payments + reimbursements taken, and **hoard** = unreimbursed OOP qualified, straight from `hsaData.rows`; drafts excluded), `hsaCalcApplyTtm()` (prefills spend/hoard from TTM unless the user overrode them — a saved value wins; `renderHsa` calls it after the ledger loads so the async data lands in the fields), `readHsaCalc()`, `onHsaCalcChange()`, pure `hsaProjCompute(g)`, and `renderHsaProj()`. **Engine 1** (recurring, market-independent) = routed spend × (tc+f), the income+payroll tax avoided each year by paying medical with pre-tax dollars. **Engine 2** (the hoard, compounding) grows the kept receipts (rate H/yr for Nc years) tax-free vs the same pre-tax dollars in an IRA (post-FICA, taxed at retirement rate out) or taxable (post-FICA+income, dividend drag + LTCG); the `cHsaProj` chart shows the three after-tax curves. Spend and hoard are **disjoint** slices of the contribution, so the contribution-side break isn't double-counted. Verified offline — at the demo inputs ($2,800 spend / $500 hoard / 20 yrs / 30 horizon): Engine 1 $1,054/yr ($21,084 cum), Engine 2 HSA $43,145 / IRA $31,078 / taxable $22,469. Wired in `showTab`: `if(id==='hsa'){ loadHsa(); hsaCalcInit(); }`. Chart is protan-safe by **dash pattern + line order + label**, not hue (HSA solid, IRA `[7,4]`, taxable `[2,3]`).

**Helpers worth knowing:** `$`=getElementById · `esc()` · `fmt(value, decimals=0)` (null→"—", negatives in accounting parens) · `apiFetch(base, action, params)` (GET, cache-busting) · `genIdemKey` + `apiPost` (writes) · `cssv(name, fallback)` (CSS var → JS) · `mdate(iso)` ("Mon YYYY") · `monthsFromNow(iso)` · `amortize(balance, annualRatePct, monthlyPI)` (single-balance payoff; shared primitive). **Note:** there is **no** `smo()` — use `mdate`.

**CSS tokens (banknote palette).** Series colors: `--c1` green `#1E5C46` · `--c2` blue `#3E6B8F` · `--c3`/`--gold`/`--amber` `#B08D3E` · `--c4`/`--terra`/`--red` `#C44E36`. Surfaces `--bg/--surface/--surface2/--surface3`; lines `--border`/`--border2`; text `--text/--muted/--faint`. Fonts: `--display` Besley (serif), `--sans` Public Sans, `--mono` Spline Sans Mono. Radii `--r`/`--rl`; `--shadow`. Tabular numerals throughout.

**Mobile / responsive.** One primary breakpoint — `@media (max-width:768px)`, a single documented block near the end of `<style>` — holds the phone rules: wide tables (`.tbl/.loan-table/.fx-table/.var-entry-table`) become `display:block;overflow-x:auto` so they scroll sideways instead of squishing; the header condenses to one row (refresh/settings drop to icons via `.btn-lbl{display:none}`, `.hide-mobile` hides `+ month`, timestamp hidden, title shrunk); the tab bar scrolls horizontally (`touch-action:pan-x;overscroll-behavior-x:contain`); gutters tighten; inputs forced to 16px (stops iOS focus-zoom). The component grid-collapses at `600/820/1100` are intentionally separate (different component widths). **Layout-only — never adds color-only meaning (Invariant 5).** CSS can't put a breakpoint in a variable, so "one breakpoint" = the same `768px` used consistently.

---

## 7. Deployment workflow

**Frontend (`index.html`, CSS, docs):** commit + push → **GitHub Pages** serves it. `raw.githubusercontent.com` and Pages are CDN-cached ~5 min — hard-refresh after a push. `config.js` is **never committed** (local-only; on Pages the setup screen collects the URL into localStorage).

**Backend (`apps_script.gs`):** edit in the Apps Script editor, then **Deploy → Manage deployments → pencil-edit the existing deployment → Version: New version → Deploy.** Never "New deployment" (new URL = stale code). Commit the updated `.gs` so the repo's reference copy doesn't drift. Smoke test: open the affected tab; `Unknown action: <x>` means the deploy didn't take.

**Claude's loop:** reads live files via `raw.githubusercontent.com`; **cannot push** — you commit/push every change. Within a session the file Claude last produced is the base; across sessions, fetch from the repo. ~90 stale deployments exist — archive opportunistically (§9).

---

## 8. Spreadsheet schemas that matter

- **`Loans!A4:J25` — student loans.** A origNum, B renum, C type, **D balance (edit target for `updateLoanBalance`)**, E rate, F interest, G principal, H payment due, I new balance, J % of balance. Row 26 totals (D26 balance, F26 interest); H27 avg rate; H28 % grad school.
- **`Loans!M–W` — mortgage amortization (full schedule to payoff Feb 2055).** Header: N3 original rate, N4 valuation (`='Physical Assets'!B6`), N5 purchase. Ledger rows 8+: M date, N description/paid-flag (`TRUE`=logged actual), O balance, P payment (P&I + escrow), Q interest, R PMI, S property tax, T principal, U equity %. **Row 40 = `Refinance` marker** (its O cell = new rate; interest switches N3→O40 after it). Cost box: W9 initial projected P&I, W10 current projected P&I (`=SUM(Q:Q)+SUM(T:T)`), W11 savings (`=W9-W10`). Read-only in the dashboard via `getMortgageData`. `Physical Assets!B8` mirrors this ledger's current balance (last `TRUE`-flagged row's O) via formula, and is what `_computeNetWorth` reads for the mortgage liability.
- **`Physical Assets`.** B3–B5 home estimates (Redfin/Zillow/realtor), B6 avg valuation, **B8 mortgage balance — formula `=XLOOKUP(TRUE, Loans!N8:N400, Loans!O8:O400, "NO PAID ROW", 0, -1)` deriving the ledger's current-month balance; net-worth source (was hand-typed, now self-maintaining — don't overtype)**, B9/B10 equity $ / %.
- **Monthly sheets `"<Month> YYYY"`.** Per-month transactions/fixed/variable; the hot path for `getMonthlyData` (cached: CacheService + durable PropertiesService tier; recompute only changed months).
- **`Config`.** `key | value | notes`; read by `cfg()`. Holds `BUDGET`, `DC_AMOUNT`, `DC_CLOSED_MONDAYS`, `QUARTERLY_EXPENSES`, `PROJECTION_TARGET`, `SEED_MONTH_CONFIG` (each with a hardcoded fallback), and `hsa_established` (HSA qualifying-date gate; **no fallback — blank disables the gate**, see §4), `hsa_receipt_folder` (Drive folder ID the receipt scan reads).
- **`Net Worth`.** Snapshots written by `saveNetWorthSnapshot`.
- **`HSA Receipts` — reimbursement-entitlement ledger.** Metadata: **`B1` = current HSA balance (manual; blank→`null`, never 0), `B2` = balance as-of date.** Headers on **row 3**, data **rows 4+**. Cols A–M: A `id` (stable positive int, never reused — `reimburseReceipt`/`updateHsaReceipt` locate rows by id, not position), B `date_incurred`, C `amount`, D `provider`, E `description`, F `category`, G `funding` (`OOP`|`HSA`), H `receipt_link` (Drive URL), I `reimbursed_amount` (**absolute**; written by `reimburseReceipt`), J `reimbursed_date`, K `notes`, L `source_file_id` (Drive fileId when created by the folder scan; blank for manual adds — the dedup key), **M `needs_review` (`TRUE` for a malformed-import draft; blank once confirmed)**. Read by `getHsa` → `{established, hsaBalance, balanceAsOf, unreimbursed, reimbursableNow, totalSubstantiated, strandedEntitlement, rows[]}` where each row carries `needsReview` (cols L is internal, not exposed). **A `needs_review` draft is excluded from every total until confirmed** (it's unverified data — `_hsaRollups` gives it `status:'needs_review'` and skips it from `unreimbursed`/`totalSubstantiated`); it sorts to the top of the ledger and shows the ✎ glyph. Written by `addHsaReceipt` (plain append — no formulas below the data block; **id = max existing id + 1 inside a `LockService.getScriptLock()`**; `needsReview:true` relaxes validation so a partial draft can be stored), `reimburseReceipt`, `scanHsaFolder`, and **`updateHsaReceipt`** (the general edit/correction path — validates strictly, writes B–H, and clears col M; reimbursed amount/date are deliberately **not** editable here so `reimburseReceipt` stays the single money-moving path). Folder scan imports PDFs from `Config!hsa_receipt_folder` named **`DATE~provider~amount.pdf`** (delimiter `~`; `_parseReceiptName` is strict). **A name that fails strict parse is salvaged by `_bestEffortReceipt`** (tolerates `~` or `_`, pulls a calendar-valid date via `_isRealDate` and a positive amount, rest → provider) and imported as a `needs_review` draft; only a name with **neither** a date nor an amount is skipped + reported. Pure helpers `_hsaRollups` / `_reimburseGuard` / `_parseReceiptName` / `_bestEffortReceipt` / `_isRealDate` are unit-tested (`testHsa`, `testParseReceiptName`, `testBestEffortReceipt`); `addHsaReceipt` + `updateHsaReceipt` have scratch-sheet integration tests (`testAddHsaReceipt`, `testUpdateHsaReceipt`, sheet `_TestScratchHsa_` — never the live sheet). **Correct any row in-dashboard via the edit modal (✎ Edit on every row); the Sheet remains a valid fallback.**

---

## 9. Open items / carried forward

- **This-Week stacked debt chart — PENDING (next build).** Stacked area of all debts over time (student + mortgage [+ car]) on the shared `amortize` primitive; must lazy-load `loans`+`mortgage` in the background so the landing tab's first paint stays fast. Protan-safe via lightness-separated bands + labels.
- **Mortgage balance reconciliation — ✓ RESOLVED (2026-06-16).** `Physical Assets!B8` changed from a hand-typed literal to a formula deriving the `Loans` ledger's current-month balance (`=XLOOKUP(TRUE, Loans!N8:N400, Loans!O8:O400, "NO PAID ROW", 0, -1)`). Net worth still reads B8, but B8 now mirrors the ledger, so the two can't drift. **No code change / no redeploy.**
- **Car-loan tab — planned** (3rd debt tab). `amortize` is already shared and ready; mirror the Mortgage tab's structure.
- **Server-validated access token — DEFERRED.** For multi-user access (Ben's wife). Random token in Script Properties; client sends it every request; server rejects mismatches. Lives in `config.js`/localStorage like the URL. Until then, access control = `SCRIPT_URL` secrecy (entered via the setup screen per browser).
- **Mobile UI polish — carried forward.** Done: 768 breakpoint, table scroll, one-row header (§6). Next: tap targets ≥44px; Chart.js readability on narrow screens (fewer ticks / compact legend via `window.matchMedia` — the one thing CSS can't express).
- **Deployment hygiene.** ~90 stale Apps Script deployments — archive them; use the single-deployment pencil-edit flow going forward.
- **audit.md staleness.** "Seven `.onclick`" → now ~12 by design (not a regression).
- **Resolved (per audit):** inline-onclick removal (65→0), idempotency keys, Config sheet, `getMonthlyData` cold-path caching. **This session:** hosting decided (GitHub Pages + public repo); `SCRIPT_URL` externalized to gitignored `config.js`; client password gate added then removed (theater); `bootDashboard` no-URL→`showSetup` fallback; split-payment branch fix (a single remaining txn line updates in place instead of being sent to `splitTransaction`); mobile responsive block.

---

## 10. Change protocol — update these together

**Add a read endpoint:** write `getX()` → add `action==='x'` to **both** the `doGet` chain and the unified action chain → client `apiFetch(SCRIPT_URL,'x')` → list it in §5 → **redeploy**.

**Add a write endpoint:** write the mutator → add to `doPost` (wrap in `_withIdem` if it touches money/rows; invalidate affected month/card/net-worth caches) → list in §5 → **redeploy**.

**Add a tab:** nav `<button data-tab="x">` → `<div id="tab-x" class="tab-panel">` panel → `if(id==='x') loadX();` in `showTab` → `loadX()` (gated by `isDemo`→`demoX()`) → `renderX()` → `demoX()` → list in §6. Redeploy only if it needs a new read endpoint.

**Add a chart:** declare `var cX=null;` → destroy before re-render → category labels + `cssv()` tokens → at least two non-color channels (Invariant 5) → add to §6 chart list.

**Change a cached payload's fields:** bump the matching `CACHE_VERSION`/`CARD_CACHE_VERSION`.

**Change a sheet schema:** update the reader's column map + dependent formulas + §8 here; verify against a real export before relying on it.

**Frontend change (HTML/CSS/JS):** commit + push → Pages serves it; hard-refresh (CDN ~5 min). No Apps Script redeploy — that's backend-only.

**Add a mobile rule:** put it in the single `@media (max-width:768px)` block near the end of `<style>`; layout-only; hold Invariant 5 (no color-only meaning).

**Any structural change:** update this file in the same pass and **commit it** — Done isn't done until INDEX, the code, and the Sheet agree.
