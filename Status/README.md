# Parser Regression Corpus

28 fixtures isolating one parse stressor each, + expected.json + a runner that
imports the REAL functions from server.py (extract_regions, clean_dataframe,
detect_table_shape). Parse-layer only: deterministic, no model.

## Install into the repo
    mkdir -p ~/Developer/excel-ai-electron/backend/tests/fixtures
    cp -r files expected.json ~/Developer/excel-ai-electron/backend/tests/fixtures/
    cp run_fixtures.py ~/Developer/excel-ai-electron/backend/tests/

## Run (from backend/, in the uv env so eparse is present)
    cd ~/Developer/excel-ai-electron/backend
    uv run --with eparse --with pandas --with openpyxl --with xlrd --with numpy \
      python tests/run_fixtures.py --server server.py \
      --files tests/fixtures/files --expected tests/fixtures/expected.json

    # -v to print detected regions per file

## Reading results
    PASS   = matches expected
    FAIL   = regression — must fix (exit code 1)
    KNOWN  = a documented open issue (banner #08, formula #17); not counted as failure

## Current known failures (as of this session)
    08  URL banner, no separator   -> banner-strip fix is shelved
    14  holdings w/ Total footer    -> statement false-positive (the UBSGreaterChina bug):
        3-numeric-col holdings table with a 'Per cent...' footer still flags 'statement'.
        Fix target: require year/period HEADERS (not just numeric-col count) to flag statement.

## Adding your own real files (do this — highest value)
    1. Drop anonymized real workbook into files/  (fake the numbers, keep the LAYOUT)
    2. Add an entry to expected.json:  "myfile.xlsx": {"min_tables": N, "tables": [{"kind":"flat"}]}
    3. Re-run. Real files break in ways synthetic ones don't.

## Notes
    - #28 is an .xlsx placeholder for the .xls path (true .xls needs xlwt to author).
    - Fixtures with "note" in expected.json document a known-weak area, not a hard assert.
