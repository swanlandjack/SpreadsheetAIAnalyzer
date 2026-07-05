# Excel AI Analyzer — Session Handoff (for a FRESH next chat)
Written: July 5, 2026 · Author context: Jack Lau, AI programmer · Machine: MacBook arm64
Repo: github.com/swanlandjack/SpreadsheetAIAnalyzer (private, origin/main)
Local: ~/Developer/excel-ai-electron  (NOT iCloud — iCloud corrupts .git; GitHub is the backup)

────────────────────────────────────────────────────────────────────────
## READ FIRST — where we are in one paragraph
The parse→clean→route→LLM-handoff pipeline works and is committed+pushed
(HEAD had P1 at commit 4868962 / 74c94fc). Since then we improved the
statement/flat detector twice more (v3, then v4) by STRESS-TESTING with new
spreadsheets. **v4 is the latest and is NOT yet applied to server.py or
committed** — applying + committing v4 is the FIRST task next session (§4).
The next PHASE is a testing loop: keep generating new `testX.xlsx`, run them
through the parser AND the app's LLM, and build confidence in both. Everything
needed to resume is below.

────────────────────────────────────────────────────────────────────────
## 1. WHAT WE ARE TESTING (the whole point)

The product's value hinges on PARSING messy real-world spreadsheets correctly,
then handing clean tables to a local LLM. Two things must be trustworthy:

  (A) PARSING — does eparse find every table, promote the right header, clean
      dtypes (no 1970 dates, numbers-as-text handled), and correctly classify
      each table as:
        - "statement" = hierarchical financials (line items + subtotals + totals
          stacked) -> the app must NOT sum a column; read the labelled row.
        - "flat"      = ordinary records -> safe to aggregate/sum.
      Mis-classification is a WRONG-OUTPUT bug (e.g. summing a balance-sheet
      column double-counts subtotals), so it's the #1 correctness risk.

  (B) ROUTING + LLM HANDOFF — when a question NAMES a sheet/table, send THAT
      table (not the selected tab); "each/all sheets" -> compact catalog, not a
      raw dump that overflows num_ctx=4096. (This is done + verified.)

  LLM ANSWER QUALITY / output formatting = a SEPARATE later phase. For now we
  "assume the LLM works as is." Stage-done = clean tables correctly handed off.

Two automated test assets exist and must stay green:
  - REGRESSION CORPUS: backend/tests/fixtures/  (synthetic, deterministic)
      run_fixtures.py imports the REAL server.py funcs and asserts each fixture
      against expected.json. Current: 30 pass, 0 FAIL, 1 known (banner #08).
  - confirm_samples.py: quick parse-check over any set of real workbooks (flags
      col_N headers, 1970 dates, zero tables, crashes).

────────────────────────────────────────────────────────────────────────
## 2. DETECTOR HISTORY (detect_table_shape) — WHY each change

The statement/flat classifier evolved through real test files. Each version
fixed a distinct real-world failure class. **v4 is current.**

  ORIGINAL (numeric-count based): flagged statement if >=2 year cols OR a "Total"
      row OR NaN-section rows. Too trigger-happy.

  P1 (committed, 4868962): PERIOD-HEADER based. statement iff >=2 period columns
      (years/quarters/months) OR section structure OR single-value+>=2 Totals.
      Dropped the numeric-count trigger.
      FIXED: fund holdings tables with a "Per cent of portfolio in top 10"
      footer were wrongly 'statement' (suppressed valid weight sums).
      (UBSGreaterChina / CapitalGroup / MorganStanley now correctly 'flat'.)

  v3 (NOT committed — superseded by v4): anchored _TOTAL_RE to line START and
      required a real total row to carry a NUMERIC value.
      FIXED: a FLAT payroll table with a prose footnote "…Total company
      headcount: 1,525" was wrongly 'statement' (the word "total" in a footnote
      tripped it). Pinned by fixture 29.

  v4 (LATEST — apply this): extends period detection to real financial year
      formats (FY 2024, FY24, 2024E, 2024A, 2024 (Projected)) AND makes ">=2 real
      subtotal rows" a statement trigger on its own (no NaN-section requirement);
      _TOTAL_RE also matches "Operating Income"/"Income Before".
      FIXED: real P&L / Balance Sheet with `FY 2024` headers + indented section
      subtotals were wrongly 'flat' (a false-NEGATIVE — would let the app sum a
      P&L column). Exposed by test5.xlsx. Pinned by fixtures 30 + 31.

  Progression summary:
      P1  killed a false-POSITIVE (fund holdings)
      v3  killed a false-POSITIVE (payroll footnote)
      v4  killed a false-NEGATIVE (FY-header P&L / section-total balance sheet)

────────────────────────────────────────────────────────────────────────
## 3. TEST FILE INVENTORY (all in Samples/, FLAT — no subdirs)

Convention (Jack): test files are named `testX.xlsx` (X = number). More will be
generated. Sample data lives in Samples/. Code downloads to ~/Downloads first,
then copied into the repo tree. ALWAYS `cd` before any command.

  Samples/
    test.xlsx        4 sheets, simple flat            COMMITTED (clean)
    test2.xlsx       6 sheets: Dashboard + 4 real statements + Portfolio
                     -> the statement-guard baseline  COMMITTED (clean)
    test3.xlsx       14 sheets, REAL fund data, FT/Bloomberg URLs, IMPORTHTML
                     -> GITIGNORED (real-ish; regenerable)
    test4.xlsx       14+ sheets, banner-stripped multi-fund
                     -> GITIGNORED (real-ish)
    test5.xlsx       4 sheets: Profit & Loss (FY headers), Balance Sheet (2
                     tables), Cash Flow Projection, Payroll Register.
                     url_cells=0 -> CLEAN, commit-safe. EXPOSED the v4 bug.
    (future testX.xlsx — check each: URL/real-data => gitignore; clean => commit)

    DEMO workbooks (fake data, user-facing, commit-safe):
      1_Portfolio_Dashboard.xlsx     multi-table-on-one-sheet showcase
      2_Financial_Statements.xlsx    statement guard
      3_Sales_by_Region.xlsx         120-row exact aggregation
      4_Multi_Fund_Comparison.xlsx   routing
      Northwind_10K_Financials.xlsx  ** the CFO-grade 10-K ** balanced
                                     (Assets=L+E each yr; NI rolls to retained
                                     earnings). 6 sheets: Income Stmt / Balance
                                     Sheet / Cash Flow / Payroll / Segments.
      make_cfo.py / make_demos.py    generators (random.seed fixed -> reproducible)
      QUESTIONS.md / README.md       suggested demo questions + expected answers

  RULE for new testX before committing:
    cd ~/Developer/excel-ai-electron/backend
    uv run --with openpyxl python3 -c "…"   # count url_cells (see §6 snippet)
    - url_cells>0 or real names  -> `echo 'Samples/testX.xlsx' >> .gitignore`
    - clean/synthetic            -> git add it

────────────────────────────────────────────────────────────────────────
## 4. FIRST TASKS NEXT SESSION (apply + commit v4)

v4 files are in ~/Downloads (or regenerate): detect_table_shape_v4.py,
expected_v4.json, 30_pl_fy_headers.xlsx, 31_bs_section_totals.xlsx.
Also fixture 29 (29_payroll_footnote_total.xlsx) belongs in the corpus.

  STEP 1 — place files:
    cd ~/Developer/excel-ai-electron/backend
    cp ~/Downloads/detect_table_shape_v4.py /tmp/dts4.py
    cp ~/Downloads/expected_v4.json tests/fixtures/expected.json
    cp ~/Downloads/29_payroll_footnote_total.xlsx tests/fixtures/files/
    cp ~/Downloads/30_pl_fy_headers.xlsx tests/fixtures/files/
    cp ~/Downloads/31_bs_section_totals.xlsx tests/fixtures/files/

  STEP 2 — replace the detector block (line-range, NOT regex — re.sub chokes on
  the \s in the replacement string). This finds the block from _YEAR_RE to
  def clean_dataframe:
    cd ~/Developer/excel-ai-electron/backend
    python3 - << 'PYEOF'
    import pathlib
    p = pathlib.Path("server.py"); lines = p.read_text().splitlines(keepends=True)
    new = pathlib.Path("/tmp/dts4.py").read_text().rstrip("\n") + "\n"
    start = next(i for i,l in enumerate(lines) if l.startswith("_YEAR_RE"))
    end   = next(i for i,l in enumerate(lines) if l.startswith("def clean_dataframe"))
    assert start < end
    p.write_text("".join(lines[:start] + [new, "\n"] + lines[end:]))
    print(f"replaced {start+1}..{end} with v4")
    PYEOF
    python3 -c "import ast; ast.parse(open('server.py').read()); print('syntax ok')"

  STEP 3 — verify (must all pass BEFORE commit):
    cd ~/Developer/excel-ai-electron/backend
    uv run --with eparse --with pandas --with openpyxl --with xlrd --with numpy \
      python tests/run_fixtures.py --server server.py \
      --files tests/fixtures/files --expected tests/fixtures/expected.json
    # expect: 30 pass, 0 FAIL, 1 known
    # then real-file spot check (P&L=statement, Payroll=flat):
    uv run --with eparse --with pandas --with openpyxl --with xlrd --with numpy python3 - \
      ../Samples/test5.xlsx ../Samples/Northwind_10K_Financials.xlsx ../Samples/test2.xlsx << 'PYEOF'
    import sys, importlib.util as u, warnings; warnings.filterwarnings("ignore")
    m=u.module_from_spec(u.spec_from_file_location("m","server.py"))
    u.spec_from_file_location("m","server.py").loader.exec_module(m)
    for fn in sys.argv[1:]:
        print("===", fn.split("/")[-1])
        for r in m.extract_regions(open(fn,"rb").read(), fn):
            print("  ", m.detect_table_shape(m.clean_dataframe(r["df"])), "|", r["label"][:36])
    PYEOF
    # PASS: test5 -> P&L/BalanceSheet x2/CashFlow = statement, Payroll = flat
    #       10-K  -> statements=statement, Payroll=flat, EPS=flat
    #       test2 -> 4 statements = statement, rest flat

  STEP 4 — commit v4 + fixtures + clean samples (NOT test3/test4):
    cd ~/Developer/excel-ai-electron
    git status                       # confirm test3/test4 absent (gitignored)
    git add backend/server.py backend/tests/fixtures/expected.json \
            backend/tests/fixtures/files/29_payroll_footnote_total.xlsx \
            backend/tests/fixtures/files/30_pl_fy_headers.xlsx \
            backend/tests/fixtures/files/31_bs_section_totals.xlsx \
            Samples/test5.xlsx Samples/make_cfo.py Samples/Northwind_10K_Financials.xlsx \
            Samples/1_Portfolio_Dashboard.xlsx Samples/2_Financial_Statements.xlsx \
            Samples/3_Sales_by_Region.xlsx Samples/4_Multi_Fund_Comparison.xlsx \
            Samples/make_demos.py
    git diff --cached --name-only    # final check
    git commit -m "Statement detector v4: FY-year headers + multi-subtotal branch; fixtures 29/30/31; demos incl 10-K; flatten Samples/"
    git push origin main             # keychain has new PAT; user swanlandjack

  NOTE: the Samples/ flatten (dropped demo/ subdir) also needs `git add -A Samples/`
  if the old Samples/demo/* deletions aren't staged yet.

  STEP 5 — rebuild the Mac binary so v4 ships in the .app (source is ahead of binary):
    cd ~/Developer/excel-ai-electron/backend && pwd
    rm -rf dist/server-mac
    uv run --with pyinstaller --with eparse --with flask --with flask-cors \
      --with pandas --with openpyxl --with xlrd --with requests --with numpy \
      pyinstaller --noconfirm --clean server.spec
    file dist/server-mac/server      # Mach-O arm64
    # smoke test frozen binary (foreground; curl in 2nd tab; cold-start ~10-20s):
    #   ./dist/server-mac/server
    #   ID=$(curl -s -F "file=@../Samples/test5.xlsx" localhost:3000/upload | \
    #        python3 -c "import sys,json;print(json.load(sys.stdin)['uploaded'][0]['id'])")
    #   curl -s "localhost:3000/text/$ID?sheet=Profit%20%26%20Loss" | grep -E "TYPE|COLUMN"

────────────────────────────────────────────────────────────────────────
## 5. THE NEXT PHASE — building confidence in PARSING + LLM

Goal Jack stated: keep generating new spreadsheets, test them, until confident
in BOTH the parsing and the LLM answers. Suggested loop:

  THE TEST LOOP (repeat per new file):
    1. Generate/obtain a new testX.xlsx (or a new demo domain: HR, inventory,
       budget, invoice register, multi-currency, quarterly actuals-vs-budget…).
    2. Check it's clean (url_cells) -> decide commit vs gitignore.
    3. PARSE check: run it through server.py (detect_table_shape per table).
       Eyeball: right # tables? right statement/flat? headers clean? no 1970?
    4. If a mis-classification -> it's a NEW bug class. Add a MINIMAL fixture to
       backend/tests/fixtures/ that isolates it + an expected.json entry, THEN
       fix detect_table_shape until corpus is green again (test-driven).
    5. LLM check (in the app): upload, ask the file's key questions, judge the
       ANSWER (routing correct? statement not summed? aggregation exact?).
    6. Re-run the full corpus to confirm no regression. Commit.

  WHAT TO PROBE for parsing confidence (known weak spots / not-yet-covered):
    - Merged multi-row headers (e.g. "2024" spanning Q1–Q4)  [eparse weakness]
    - KPI cards / single merged cells above a table
    - Multiple tables stacked with tiny gaps
    - Wide statements (5+ period columns), quarterly + annual mixed
    - Actual vs Budget vs Variance column layouts (NOT year headers)
    - Non-ASCII / CJK headers (Jack is HK-based; real files may have these)
    - Footnotes with numbers (already fixed once via v3 — probe variants)
    - Very large sheets (perf + sample-cap behavior)
    - .csv edge cases (quoted commas/newlines already covered; try semicolons,
      European decimal commas)

  WHAT TO PROBE for LLM confidence (separate, later, but start noticing):
    - Does it read the right labelled row on statements (never sum a column)?
    - Does routing pick the named table over the selected tab?
    - Are exact aggregations correct over ALL rows (not a 15-row sample)?
    - Cross-statement questions (does 10-K NI match on IS and Cash Flow?)
    - Does "each sheet" use the catalog, not overflow context?
    Consider: a small manual ANSWER-KEY per demo (QUESTIONS.md already has this
    for the 10-K) so LLM answers can be spot-graded, not just vibes.

  KNOWN-OPEN parser items (deferred, evidence logged):
    - #08 banner-above-header (no blank separator): per-sheet strip TESTED and
      REJECTED (destroys real test3 sheets, e.g. FranklinUS->empty). Needs a
      smarter multi-row banner detector. Shelved. test4 was hand-stripped.
    - P2 formula/URL bleed-in (IMPORTHTML cells, URL in Ticker): NOT a
      data_only fix (eparse owns its read; cached value IS the URL). Needs a
      content sanitizer. Cosmetic; model tolerates it. Deferred.
    - Merged multi-row headers (#10): eparse flattens them. Deferred.

────────────────────────────────────────────────────────────────────────
## 6. HANDY SNIPPETS

  # URL / real-data check for a new testX (decide commit vs gitignore)
  cd ~/Developer/excel-ai-electron/backend
  uv run --with openpyxl python3 - ../Samples/testX.xlsx << 'PYEOF'
  import sys, openpyxl, re
  U=re.compile(r'https?://'); wb=openpyxl.load_workbook(sys.argv[1],data_only=True,read_only=True)
  n=sum(1 for ws in wb.worksheets for row in ws.iter_rows(values_only=True) for v in row if v and U.search(str(v)))
  print("sheets",wb.sheetnames,"url_cells",n)
  PYEOF

  # dev run (no rebuild for logic changes)
  cd ~/Developer/excel-ai-electron/backend
  uv run --with flask --with flask-cors --with pandas --with openpyxl --with xlrd \
    --with requests --with numpy --with eparse server.py

  # confirm ALL real samples parse clean
  cd ~/Developer/excel-ai-electron/backend
  uv run --with eparse --with pandas --with openpyxl --with xlrd --with numpy \
    python confirm_samples.py ../Samples/*.xlsx

────────────────────────────────────────────────────────────────────────
## 7. STATE FLAGS / GOTCHAS
  - v4 applied+committed?  NO -> §4 is the first job.
  - Binary ships v4?       NO -> §4 step 5 rebuild.
  - origin/main:           had P1 (74c94fc); v3 was never committed (skip it);
                           commit v4 directly.
  - Samples/ flatten:      done on disk; ensure `git add -A Samples/` captures the
                           demo/ deletions in the v4 commit.
  - STATUS-stage-complete.md describes P1's detector logic -> now STALE (v4).
    Update it (or note "see this handoff for detector v4").
  - Token: new PAT in macOS keychain; old PAT revoked (401). Remote URL tokenless.
  - Monetization (Mac-first, Jack IS an Apple Developer): notarization needs a
    Developer ID Application cert (verify: `security find-identity -v -p codesigning`)
    + app-specific password (self-serve at appleid.apple.com) OR App Store Connect
    API key (BLOCKED — only Account Holder Katherine can request that). License
    model recommendation: activate-once-online -> cache token -> run offline.
    Merchant-of-record (Paddle/Lemon Squeezy) for keys+VAT. All DEFERRED.
  - NEVER put .git in iCloud. NEVER re.sub-replace code with \s in the replacement.
    ALWAYS `cd` before commands. ALWAYS check url_cells before committing a testX.
────────────────────────────────────────────────────────────────────────
```
```
