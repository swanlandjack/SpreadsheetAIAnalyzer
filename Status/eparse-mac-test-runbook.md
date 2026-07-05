# Excel AI Analyzer — eparse Integration: Mac Test → Windows Runbook

Session: July 4, 2026. Covers the eparse multi-table parse layer: how to run and
test it on Mac first, then convert to Windows via GitHub Actions.

---

## 0. Golden rules (learned the hard way)

- **`.git` never lives in iCloud.** iCloud makes `… 2` conflict copies of `.git`
  internals and corrupts the repo (July incident). Repo lives in `~/Developer`.
- **Master (source only) in iCloud, repo in `~/Developer`, source of truth on GitHub.**
  Copy is one-way, iCloud → Developer, *source files only, never `.git`*.
- **`pwd` before any destructive command.** No clever `cd || (...)` before `rm`.
- **One terminal command at a time, verify each**, especially copies.
- Stage folders in iCloud are NOT reliably newest-by-number. `Stage4` here was
  *older* than Developer (pre-July-speed-fix). Verify content, don't trust the name.

---

## 1. Paths

```
REPO (run/build here):   ~/Developer/excel-ai-electron
  backend/server.py        <- the file that changed
  backend/server.spec      <- NOT used by the Windows workflow (see §7)
  backend/.venv            <- uv venv
  .github/workflows/build-server-windows-x64.yml

MASTER (source backup):  ~/Documents/Programming/2026Programming/TestFolderCVS/
                         TestOllamaCVS_Electron/Stage4/excel-ai-electron/backend/
  (has NO .git — good. Note: was stale/permission-locked in July session.)
```

---

## 2. What changed in server.py (why we're testing)

Old: one `pandas.read_excel` per sheet → assumes ONE flat table, header in row 1.
New parse layer for `.xlsx/.xlsm`:

```
workbook
  -> eparse: detect ALL tables per sheet (region detection)
  -> promote eparse header row -> clean_dataframe -> dataframe_to_text
  -> pandas fallback for any sheet eparse finds no table in (nothing dropped)
```

- Each table = one addressable **region**, keyed by a **label that embeds the sheet name**:
  - 1 table on a sheet  -> label = sheet name           (old behaviour preserved)
  - N tables on a sheet -> label = `Sheet · Header`      (deduped with `@anchor`)
- `.csv` and legacy `.xls` keep the original flat pandas path.
- **Statement guard**: financial statements stack line items + subtotals + totals in
  one column, so summing double-counts (267,000 Total Assets column sums to 1,678,000).
  Such tables are auto-detected; their column sums/avgs are SUPPRESSED in the text and
  full labelled rows shown instead.
- Constants kept from July speed fix: `MAX_SAMPLE_ROWS=15`, `num_ctx=4096`.
  New: `STATEMENT_MAX_ROWS=60` (statements are row-lookup, show more rows).

---

## 3. Install the new server.py into the repo

Skip the iCloud master if it's permission-locked (July state). The repo + GitHub are
the real source of truth. Put the new file straight into Developer.

```bash
cd ~/Developer/excel-ai-electron/backend
cp server.py server.py.bak                 # backup current (July) version
cp ~/Downloads/server.py ./server.py       # drop in the new eparse version
```

Verify it's the RIGHT file (not stale, July fix intact):

```bash
grep -c "eparse" server.py                                  # expect ~20+
grep -nE "^MAX_SAMPLE_ROWS|STATEMENT_MAX_ROWS|num_ctx" server.py
# expect: MAX_SAMPLE_ROWS = 15 | STATEMENT_MAX_ROWS = 60 | num_ctx": 4096
```

If a value is wrong (e.g. `MAX_SAMPLE_ROWS = 100`, `num_ctx 16384`), you copied a
stale file — restore and re-copy:

```bash
cp server.py.bak server.py
```

---

## 4. Environment (uv venv)

```bash
cd ~/Developer/excel-ai-electron/backend
uv venv                                     # once; creates .venv
source .venv/bin/activate                   # prompt shows (backend)
```

Confirm you're in the repo's venv:

```bash
which python                                # -> .../Developer/excel-ai-electron/backend/.venv/bin/python
```

Install deps (eparse pulls lxml + peewee + click transitively):

```bash
uv pip install flask flask-cors pandas openpyxl xlrd requests numpy eparse
python -c "import eparse, lxml, peewee; from eparse.core import get_df_from_file; print('ok', eparse.__version__)"
# expect: ok 0.9.2
```

> The venv used HERE must be the same env any local PyInstaller build reads.

---

## 5. Run + smoke test

Terminal A — Ollama:

```bash
ollama serve            # leave running
ollama list             # need qwen2.5 ; else: ollama pull qwen2.5
```

Terminal B — backend:

```bash
cd ~/Developer/excel-ai-electron/backend && source .venv/bin/activate
python server.py        # banner ends: ...(eparse multi-table)  http://localhost:3000
```

Terminal C — checks:

```bash
curl -s localhost:3000/health
# expect: {"status":"ok","ollama":true,...}
```

Notes:
- Browsing to `http://localhost:3000/` returns **"Not Found"** — CORRECT. Flask has no
  `/` route; the UI is served by Electron, not Flask. 404 = server is up and routing.
- `"ollama":false` -> start `ollama serve`; backend itself is still fine.

---

## 6. The actual parse test (no model needed)

Upload a file and print the table classification in one shot:

```bash
ID=$(curl -s -F "file=@$HOME/Downloads/test2.xlsx" localhost:3000/upload \
  | python3 -c "import sys,json;d=json.load(sys.stdin)['uploaded'][0];print(d['id']);[print(' ',t['kind'],'|',t['source'],'|',t['label'],'|',f\"{t['rows']}x{t['cols']}\") for t in d.get('tables',[])]")
echo "ID=$ID"
```

Dump the EXACT text a table sends to the model (quote label; escape `$`):

```bash
curl -s "localhost:3000/text/$ID?sheet=Detailed Balance Sheet" | head -50
# statement PASS: "TYPE: financial statement", "Source sheet:", "sum/avg omitted"
#   on numeric cols, "Total Assets" visible, and NO 1,678,000 anywhere.
```

Whole workbook / other endpoints:

```bash
curl -s "localhost:3000/text/$ID" | less          # all tables concatenated
curl -s "localhost:3000/files" | python3 -m json.tool
curl -s -X DELETE localhost:3000/files            # reset the in-memory store
```

### What to eyeball on every file
- **Table count** matches Excel (side-by-side AND stacked tables split).
- **`kind`** correct: real statements -> `statement` (sums suppressed);
  transaction/holding lists -> `flat` (sums present).
- **No `1970-01-01`** in numeric columns; **no `col_1..col_N`** headers
  (that combo = the old flatten failing).
- **`source=pandas-fallback`** -> eparse found no table on that sheet; open it in
  Excel and confirm it's genuinely KPI-only / empty, not a missed real table.

### Edge cases worth throwing at it (test/test2 don't cover these)
- Merged multi-row headers (`2024` spanning `Q1 Q2 Q3 Q4`) — eparse's weak spot.
- A KPI-only sheet (merged single cells, no 2x2 block).
- Blank spacer rows *inside* one table (does eparse split it?).
- A statement with month or single-period columns (not years) — does the detector
  still catch it, or false-negative to `flat` and expose a bad column sum?
- A genuine `.xls` and a `.csv` — these bypass eparse; confirm fallback still labels/classifies.

### Full loop (with model)
Launch the Electron app, drag a file in, pick the table tab (tabs = the per-table
labels), ask a question. Answer should cite which table/sheet it used and never sum a
statement column.

---

## 7. Convert to Windows (GitHub Actions)

**Critical:** the workflow runs `pyinstaller ... server.py` with CLI `--hidden-import`
flags. Passing a `.py` (not a `.spec`) makes PyInstaller synthesize its own spec and
**ignore `backend/server.spec`**. So the Windows fix lives in the YAML, not the spec.

Edit `.github/workflows/build-server-windows-x64.yml` — two changes:

1. Install step — add `eparse` to the pip line:
   ```
   pip install flask flask-cors pandas openpyxl xlrd requests numpy pyinstaller eparse
   ```
2. Build step — add before `server.py`:
   ```
   --hidden-import=eparse `
   --hidden-import=eparse.core `
   --hidden-import=eparse.interfaces `
   --hidden-import=peewee `
   --hidden-import=lxml `
   --hidden-import=lxml.etree `
   --exclude-module=psycopg2 `
   ```
   - `eparse.core` is imported LAZILY inside a function in server.py — PyInstaller's
     static scan can miss lazy imports, so name it explicitly.
   - `peewee` / `eparse.interfaces`: satisfy the static import chain (never hit at runtime).
   - `lxml`: the one new C-extension transitive dep — watch it in the build log if the
     run ever fails.
   - `--exclude-module=psycopg2`: optional; keeps the log clean (driver not installed).

Then push and build:

```bash
cd ~/Developer/excel-ai-electron
git add backend/server.py .github/workflows/build-server-windows-x64.yml
git commit -m "eparse multi-table parse layer + statement guard; workflow deps"
git push origin main
```

- GitHub -> Actions -> **Build server.exe (Windows x64)** -> Run workflow.
- Download artifact `server-windows-x64.zip` -> unzip -> `server.exe` + `_internal/`.
- **Clean** `backend/dist/server/` completely, then copy in ONLY the fresh
  `server.exe` + `_internal/` (stale binaries get bundled otherwise).
- `cd ~/Developer/excel-ai-electron && npx electron-builder --win`
- Edit the v1.0 release -> delete old Windows `.exe` -> upload new -> update release.

### Windows install gotcha (from July)
In-place update over a **running** app can't overwrite a locked `server.exe` — it
updates `_internal/` but SKIPS the `.exe`, leaving the OLD binary running.
Fix: reboot the PC (clears the lock) -> Task Manager confirm no `server.exe` ->
uninstall -> delete leftover
`C:\Users\<name>\AppData\Local\Programs\Excel AI Analyzer\` -> reinstall fresh ->
verify `server.exe` Properties date is the NEW build.

---

## 8. Known open items (post-test)

- **pandas 3/4 warning**: `select_dtypes(include=["object"])` in `clean_dataframe`
  will change semantics in pandas 4 (str vs object). Cosmetic now; harden by passing
  explicit dtypes before a pandas 4 upgrade.
- **Statement detector false positives**: a 2-column fund table with a "Total" row can
  wrongly flag `statement`. Candidate fix: only treat a Total-row match as `statement`
  when the table also has >=3 numeric columns OR year/period headers.
- **`Overview`-style wide sheets**: a large single detected table (e.g. 157x15) may be
  several stacked tables eparse merged — inspect before trusting.
- **No auto-routing yet**: with no table tab selected, all tables concatenate and can
  overflow 4096 ctx on big workbooks. Per-table tabs = manual routing for now.
  Next increment: catalog builder + keyword/small-LLM router carrying a shape tag.

---

## 9. Rollback

```bash
cd ~/Developer/excel-ai-electron/backend
cp server.py.bak server.py        # restore July version
```
GitHub remains source of truth; the previous release asset is unchanged until you
explicitly replace it.
```
```
