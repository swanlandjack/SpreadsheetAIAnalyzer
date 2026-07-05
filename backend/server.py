"""
Excel AI Analyzer — Python Backend
Uses Flask + Pandas + eparse + Ollama to analyze uploaded Excel/CSV files.

Install:
    pip install flask flask-cors pandas openpyxl xlrd requests numpy eparse

Run:
    python server.py

Ollama must be running:
    ollama serve
    ollama pull qwen2.5

------------------------------------------------------------------------------
WHAT CHANGED (eparse integration)
------------------------------------------------------------------------------
The old pipeline read each sheet with a single `pandas.read_excel`, assuming ONE
flat table with headers in row 1. Real finance workbooks break that (multiple
tables per sheet, KPI cards, headers not in row 1, side-by-side tables).

New parse layer for .xlsx/.xlsm:
    workbook
      -> eparse: detect ALL tables per sheet (region detection)
      -> per table: promote eparse's header row -> clean_dataframe -> text
      -> pandas fallback for any sheet eparse finds no table in (nothing dropped)

Each detected table becomes one addressable "region", keyed by a LABEL that
always carries its sheet name:
    * one table on a sheet   -> label = sheet name          (old behaviour kept)
    * many tables on a sheet -> label = "Sheet · Header"     (deduped with @anchor)

.csv and legacy .xls keep the original flat pandas path.

Statement guard: financial statements (Income Statement / Balance Sheet / etc.)
stack line items together with their subtotals and grand totals in one column,
so summing the column double-counts (e.g. a 267,000 Total Assets column sums to
1,678,000). Such tables are detected and their misleading column sums/averages
are SUPPRESSED in the text sent to the model; instead the full labelled rows are
shown for row-lookup.
"""

import io
import json
import os
import re
import tempfile
import traceback
import warnings

import numpy as np
import pandas as pd
import requests
from flask import Flask, Response, jsonify, request, stream_with_context
from flask_cors import CORS

# Force UTF-8 encoding on Windows to prevent charmap errors
import sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if sys.stderr.encoding != 'utf-8':
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

# eparse prints nothing critical; silence noisy pandas date-inference warnings
warnings.filterwarnings("ignore", message="Could not infer format")

app = Flask(__name__)
CORS(app)

# ── In-memory store ────────────────────────────────────────────────────────────
# { file_id: { id, name, sheets: {label: df}, texts: {label: str}, regions: [meta] } }
# NOTE: "sheets"/"texts" are now keyed by REGION LABEL (which embeds the sheet
# name). For simple one-table-per-sheet files the label IS the sheet name, so all
# existing routes behave identically.
uploaded_files: dict = {}

# ── Config ─────────────────────────────────────────────────────────────────────
OLLAMA_BASE_URL = os.getenv("OLLAMA_BASE_URL", "http://localhost:11434")
DEFAULT_MODEL   = os.getenv("OLLAMA_MODEL", "qwen3:4b-instruct")

MAX_SAMPLE_ROWS      = 15   # rows shown for a FLAT table (small = fast CPU prefill; aggregates come from COLUMN STATISTICS which cover ALL rows)
STATEMENT_MAX_ROWS   = 60   # statements are row-lookup; show more rows, no column sums
MAX_COLS             = 30   # trim very wide sheets


# ══════════════════════════════════════════════════════════════════════════════
#  STEP 1 — READ  (upload bytes >> list of table REGIONS)
# ══════════════════════════════════════════════════════════════════════════════
#
#  A "region" is one detected table:
#     { sheet, anchor, header, df, source, label? }
#  `label` is assigned by label_regions() once all regions for a file are known.

def _promote_header(df: pd.DataFrame) -> pd.DataFrame:
    """
    eparse returns a DataFrame with POSITIONAL integer columns (1,2,3…) and the
    real header living in row 0. Promote that first row to be the column names so
    the rest of the pipeline (clean_dataframe / dataframe_to_text) works unchanged.
    """
    if df.empty:
        return df
    header = [str(v) for v in df.iloc[0].values]
    out = df.iloc[1:].copy()
    out.columns = header
    return out.reset_index(drop=True)


def _pandas_sheet_df(path: str, sheet_name) -> pd.DataFrame:
    return pd.read_excel(
        path, sheet_name=sheet_name, engine="openpyxl",
        na_values=["", "N/A", "NA", "n/a", "#N/A"],
    )


def _regions_from_xlsx(path: str) -> list:
    """
    eparse multi-table crawl of an .xlsx/.xlsm file, grouped by sheet, with a
    pandas fallback for any sheet eparse finds no table in (so nothing is lost).
    """
    import openpyxl

    regions = []
    seen_sheets = set()

    # 1) eparse — every detected table across every sheet
    try:
        from eparse.core import get_df_from_file
        for df, anchor, header, sheet in get_df_from_file(path):
            promoted = _promote_header(df)
            if promoted.empty:
                continue
            seen_sheets.add(sheet)
            regions.append({
                "sheet":  sheet,
                "anchor": str(anchor),
                "header": str(header),
                "df":     promoted,
                "source": "eparse",
            })
    except Exception:
        # eparse failed entirely — fall through; the pandas pass below covers all sheets
        print("  [eparse] extraction failed, falling back to pandas per-sheet")
        traceback.print_exc()

    # 2) pandas fallback for sheets eparse skipped (e.g. KPI-cards-only sheets)
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        sheet_names = list(wb.sheetnames)
        wb.close()
    except Exception:
        sheet_names = []

    for name in sheet_names:
        if name in seen_sheets:
            continue
        try:
            df = _pandas_sheet_df(path, name)
            if not df.dropna(how="all").empty:
                regions.append({
                    "sheet":  name,
                    "anchor": "A1",
                    "header": "",
                    "df":     df,
                    "source": "pandas-fallback",
                })
        except Exception:
            print(f"  [fallback] could not read sheet {name!r}")

    return regions


def extract_regions(raw: bytes, filename: str) -> list:
    """
    Turn an upload into a list of labelled table regions.

      .csv          -> one flat region ("Data")            [pandas]
      .xls          -> one flat region per sheet           [pandas, legacy engine]
      .xlsx / .xlsm -> eparse multi-table + pandas fallback [region detection]

    Returns regions each carrying a unique `label` that embeds the sheet name.
    """
    ext = filename.rsplit(".", 1)[-1].lower()
    regions = []

    # ── CSV ──────────────────────────────────────────────────────────────────
    if ext == "csv":
        buf = io.BytesIO(raw)
        for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
            try:
                buf.seek(0)
                df = pd.read_csv(buf, encoding=enc, on_bad_lines="skip",
                                 low_memory=False)
                regions.append({"sheet": "Data", "anchor": "A1", "header": "",
                                "df": df, "source": "pandas"})
                print(f"  [CSV] encoding={enc}  rows={len(df)}")
                break
            except (UnicodeDecodeError, Exception):
                continue
        if not regions:
            raise ValueError("Could not decode CSV with any common encoding.")

    # ── XLS (legacy binary) — keep the simple pandas path ────────────────────
    elif ext == "xls":
        xl = pd.ExcelFile(io.BytesIO(raw), engine="xlrd")
        for name in xl.sheet_names:
            df = xl.parse(name, na_values=["", "N/A", "NA", "n/a", "#N/A"])
            regions.append({"sheet": name, "anchor": "A1", "header": "",
                            "df": df, "source": "pandas"})
        print(f"  [XLS]  sheets={xl.sheet_names}")

    # ── XLSX / XLSM — eparse multi-table ─────────────────────────────────────
    elif ext in ("xlsx", "xlsm"):
        # eparse reads from a path, so spool the upload to a temp file
        with tempfile.NamedTemporaryFile(suffix=f".{ext}", delete=False) as tmp:
            tmp.write(raw)
            tmp_path = tmp.name
        try:
            regions = _regions_from_xlsx(tmp_path)
        finally:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass
        # last-ditch: if nothing at all was found, read every sheet flat
        if not regions:
            xl = pd.ExcelFile(io.BytesIO(raw), engine="openpyxl")
            for name in xl.sheet_names:
                df = xl.parse(name, na_values=["", "N/A", "NA", "n/a", "#N/A"])
                regions.append({"sheet": name, "anchor": "A1", "header": "",
                                "df": df, "source": "pandas-fallback"})
        srcs = {}
        for r in regions:
            srcs[r["source"]] = srcs.get(r["source"], 0) + 1
        print(f"  [XLSX] tables={len(regions)}  by_source={srcs}")

    else:
        raise ValueError(
            f"Unsupported extension: .{ext}  "
            "(supported: .xlsx  .xls  .xlsm  .csv)"
        )

    return label_regions(regions)


def label_regions(regions: list) -> list:
    """
    Assign each region a unique display label that always embeds its sheet name:
      * exactly one table on a sheet  -> label = sheet name (old behaviour)
      * multiple tables on a sheet    -> label = "Sheet · Header"
                                         (fall back to "@anchor" on collision)
    """
    by_sheet = {}
    for r in regions:
        by_sheet.setdefault(r["sheet"], []).append(r)

    used = set()
    for sheet, rs in by_sheet.items():
        if len(rs) == 1:
            rs[0]["label"] = _uniquify(sheet, used)
        else:
            for r in rs:
                header = r.get("header", "").strip()
                base = f"{sheet} \u00b7 {header}".rstrip(" \u00b7") if header else sheet
                if base in used:
                    base = f"{base} @{r['anchor']}"
                r["label"] = _uniquify(base, used)
    return regions


def _uniquify(label: str, used: set) -> str:
    lbl = label
    i = 2
    while lbl in used:
        lbl = f"{label} ({i})"
        i += 1
    used.add(lbl)
    return lbl


# ══════════════════════════════════════════════════════════════════════════════
#  STEP 1.5 — CLASSIFY TABLE SHAPE  (flat vs. hierarchical statement)
# ══════════════════════════════════════════════════════════════════════════════

_YEAR_RE  = re.compile(r"^(19|20)\d{2}(\.0)?$")
_TOTAL_RE = re.compile(r"(?i)\b(?:total|subtotal|net\s|gross\s)")


# ── Period-header helpers (place near _YEAR_RE / _TOTAL_RE) ────────────────────
_MONTHS = {"jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec",
           "january","february","march","april","june","july","august","september",
           "october","november","december"}
_QTR_RE = re.compile(r"^(q[1-4]|[1-4]q|h[12])(\s*\d{2,4})?$", re.I)

def _is_period(col) -> bool:
    c = str(col).strip()
    return bool(_YEAR_RE.match(c)) or bool(_QTR_RE.match(c)) or c.lower() in _MONTHS


def detect_table_shape(df: pd.DataFrame) -> str:
    """
    statement = hierarchical financial statement (line items + subtotals/totals
                stacked in a column) -> do NOT sum columns.
    flat      = ordinary record list -> safe to aggregate.

    A statement is identified by PERIOD columns (years / quarters / months) as
    headers, OR explicit section structure (a Total row AND all-NaN section-header
    rows), OR a single value column with >=2 Total-like rows. Numeric-column COUNT
    is NOT a trigger: a fund holdings table can have several numeric columns and a
    'Per cent of portfolio...' footer yet is flat.
    """
    if df.shape[0] < 2 or df.shape[1] < 2:
        return "flat"

    # (a) >=2 period columns -> statement (years / quarters / months matrix)
    if sum(1 for c in df.columns if _is_period(c)) >= 2:
        return "statement"

    first_col = df.iloc[:, 0].astype(str)
    has_total = bool(first_col.str.contains(_TOTAL_RE, regex=True).any())
    num = df.select_dtypes(include="number")

    # (b) section-structured statement: a Total row AND all-NaN section-header rows
    if has_total and not num.empty:
        allnan = int(num.isna().all(axis=1).sum())
        labelled = int(first_col.str.len().gt(0).sum())
        if allnan >= 1 and labelled > allnan:
            return "statement"

    # (c) single value column with >=2 Total-like rows (income statement, one period)
    if num.shape[1] == 1:
        if int(first_col.str.contains(_TOTAL_RE, regex=True).sum()) >= 2:
            return "statement"

    return "flat"

def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Standard cleanup so the data Ollama sees is well-formed:

    1. Drop fully-empty rows and columns (common in XLS with merged headers).
    2. Sanitise column names (strip whitespace, replace special chars).
    3. Deduplicate column names.
    4. Try to parse date-like string columns.
    5. Infer better dtypes.
    6. Trim to MAX_COLS columns.
    """
    # 1. Drop fully-empty rows/cols
    df = df.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)

    # 2 & 3. Clean and deduplicate column names
    new_cols = []
    seen = {}
    for col in df.columns:
        c = str(col).strip()
        c = re.sub(r"\s+", " ", c)
        c = re.sub(r"[^\w\s\-\(\)\[\]\/\%\#\$\.]", "_", c)
        c = c.strip("_")
        if not c or c.lower().startswith("unnamed") or c.lower() == "nan":
            c = f"col_{len(new_cols) + 1}"
        if c in seen:
            seen[c] += 1
            c = f"{c}_{seen[c]}"
        else:
            seen[c] = 0
        new_cols.append(c)
    df.columns = new_cols

    # 4. Coerce object columns to their best type: numeric > datetime > text.
    #    Order matters. eparse hands numeric columns back as `object` (the header
    #    row forced object dtype), and pd.to_datetime() would read a bare int like
    #    50 as epoch-nanoseconds -> 1970-01-01. Numeric-first prevents that; only
    #    genuine datetime OBJECTS or date-looking STRINGS are parsed as dates.
    import datetime as _dt

    def _is_dateish(v):
        return isinstance(v, (_dt.datetime, _dt.date, pd.Timestamp))

    for col in df.columns:
        _s = df[col]
        if (pd.api.types.is_numeric_dtype(_s)
                or pd.api.types.is_datetime64_any_dtype(_s)
                or pd.api.types.is_bool_dtype(_s)):
            continue
        s      = df[col]
        sample = s.dropna()
        if sample.empty:
            continue
        head = sample.head(20)

        # (a) already datetime-like objects (common from eparse)
        if head.map(_is_dateish).mean() > 0.7:
            df[col] = pd.to_datetime(s, errors="coerce")
            continue

        # (b) numeric? (do this BEFORE any date parsing)
        if pd.to_numeric(sample, errors="coerce").notna().mean() > 0.7:
            df[col] = pd.to_numeric(s, errors="coerce")
            continue

        # (c) date-like STRINGS only (never bare numbers)
        if head.map(lambda v: isinstance(v, str)).mean() > 0.7:
            if pd.to_datetime(sample, errors="coerce").notna().mean() > 0.7:
                df[col] = pd.to_datetime(s, errors="coerce")

    # 5. Infer better dtypes
    df = df.infer_objects()

    # 6. Trim columns
    if len(df.columns) > MAX_COLS:
        print(f"  [CLEAN] Trimming {len(df.columns)} >> {MAX_COLS} columns")
        df = df.iloc[:, :MAX_COLS]

    return df


# ══════════════════════════════════════════════════════════════════════════════
#  STEP 3 — CONVERT TO TEXT  (DataFrame >> rich plain text for Ollama)
# ══════════════════════════════════════════════════════════════════════════════

def dtype_label(series: pd.Series) -> str:
    dt = series.dtype
    if pd.api.types.is_datetime64_any_dtype(dt):   return "date/time"
    if pd.api.types.is_bool_dtype(dt):             return "boolean"
    if pd.api.types.is_integer_dtype(dt):          return "integer"
    if pd.api.types.is_float_dtype(dt):            return "decimal"
    return "text"


def fmt_num(v) -> str:
    """Format a number cleanly: no sci notation, commas, trailing zeros stripped."""
    if pd.isna(v):
        return "N/A"
    if isinstance(v, (int, np.integer)) or (isinstance(v, float) and v == int(v)):
        return f"{int(v):,}"
    return f"{v:,.4f}".rstrip("0").rstrip(".")


def numeric_stats_line(series: pd.Series) -> str:
    s = series.dropna()
    if s.empty:
        return "    (all values blank)"
    parts = [
        f"sum={fmt_num(s.sum())}",
        f"avg={fmt_num(s.mean())}",
        f"median={fmt_num(s.median())}",
        f"min={fmt_num(s.min())}",
        f"max={fmt_num(s.max())}",
    ]
    nulls = series.isna().sum()
    if nulls:
        parts.append(f"blanks={nulls}")
    return "    " + ",  ".join(parts)


def numeric_range_line(series: pd.Series) -> str:
    """Statement-safe: min/max only, NO sum/avg (they double-count subtotals)."""
    s = series.dropna()
    if s.empty:
        return "    (all values blank)"
    return (f"    min={fmt_num(s.min())},  max={fmt_num(s.max())}"
            f"   [sum/avg omitted — column mixes line items with subtotals]")


def date_stats_line(series: pd.Series) -> str:
    s = series.dropna()
    if s.empty:
        return "    (all values blank)"
    return (f"    earliest={s.min().date()},  "
            f"latest={s.max().date()},  "
            f"blanks={series.isna().sum()}")


def text_stats_line(series: pd.Series) -> str:
    n_unique = series.nunique()
    n_null   = series.isna().sum()
    top5     = series.value_counts().head(5).index.tolist()
    top_str  = ", ".join(f'"{v}"' for v in top5)
    return f"    {n_unique} unique values,  blanks={n_null},  top: [{top_str}]"


def dataframe_to_text(df: pd.DataFrame, label: str,
                      sheet_name: str = None, is_statement: bool = False) -> str:
    """
    Convert a cleaned DataFrame (one table region) to a structured plain-text
    document for Ollama.

    Layout:
        TABLE OVERVIEW       — label, source sheet, row/col counts, columns
        COLUMN STATISTICS    — type + full-dataset stats  (FLAT tables)
          or COLUMN NOTES    — min/max only, no sums      (STATEMENT tables)
        DATA SAMPLE          — first N rows                (FLAT: N=15)
          or FULL STATEMENT  — up to 60 rows for row-lookup (STATEMENT)
    """
    lines = []
    n_rows, n_cols = df.shape
    cap          = STATEMENT_MAX_ROWS if is_statement else MAX_SAMPLE_ROWS
    is_truncated = n_rows > cap

    # ── Overview ──────────────────────────────────────────────────────────────
    lines.append("=" * 68)
    lines.append(f"TABLE: {label}")
    if sheet_name:
        lines.append(f"Source sheet: {sheet_name}")
    if is_statement:
        lines.append("TYPE: financial statement — this column layout stacks line "
                     "items together with their subtotals and totals. Do NOT sum a "
                     "column; read the specific labelled row (e.g. 'Total Assets').")
    lines.append("=" * 68)
    lines.append(f"Total rows    : {n_rows:,}")
    lines.append(f"Total columns : {n_cols}")
    lines.append(f"Columns       : {', '.join(map(str, df.columns.tolist()))}")
    lines.append("")

    # ── Column statistics / notes ─────────────────────────────────────────────
    lines.append("-" * 68)
    if is_statement:
        lines.append("COLUMN NOTES  (statement — do NOT sum columns; read labelled rows)")
    else:
        lines.append("COLUMN STATISTICS  (computed on ALL rows)")
    lines.append("-" * 68)
    for col in df.columns:
        series = df[col]
        lines.append(f"  [{dtype_label(series)}]  {col}")
        if pd.api.types.is_numeric_dtype(series):
            lines.append(numeric_range_line(series) if is_statement
                         else numeric_stats_line(series))
        elif pd.api.types.is_datetime64_any_dtype(series):
            lines.append(date_stats_line(series))
        else:
            lines.append(text_stats_line(series))
    lines.append("")

    # ── Data sample / full statement ──────────────────────────────────────────
    lines.append("-" * 68)
    if is_statement:
        if is_truncated:
            lines.append(f"STATEMENT ROWS  (first {cap} of {n_rows:,} — read specific rows for values)")
        else:
            lines.append(f"FULL STATEMENT  ({n_rows} rows — read the specific labelled row for each value)")
    elif is_truncated:
        lines.append(
            f"DATA SAMPLE  (first {cap} of {n_rows:,} rows shown)\n"
            f"NOTE: Use the statistics above for totals/averages — "
            f"they cover ALL {n_rows:,} rows."
        )
    else:
        lines.append(f"FULL DATA  ({n_rows} rows)")
    lines.append("-" * 68)

    sample = df.head(cap).copy()

    # Format datetime cols as readable strings
    for col in sample.columns:
        if pd.api.types.is_datetime64_any_dtype(sample[col]):
            sample[col] = sample[col].dt.strftime("%Y-%m-%d").fillna("")
        else:
            sample[col] = sample[col].fillna("").astype(str)

    # Fixed-width columns (cap at 28 chars)
    widths = {
        col: min(28, max(len(str(col)),
                         sample[col].str.len().max() if not sample[col].empty else 0))
        for col in sample.columns
    }

    def row_line(values):
        return "  " + "  ".join(
            str(v)[:widths[c]].ljust(widths[c])
            for c, v in zip(sample.columns, values)
        )

    lines.append(row_line(sample.columns))
    lines.append("  " + "  ".join("-" * widths[c] for c in sample.columns))
    for _, row in sample.iterrows():
        lines.append(row_line([row[c] for c in sample.columns]))

    if is_truncated:
        lines.append(f"\n  ... ({n_rows - cap:,} more rows not shown)")

    lines.append("=" * 68)
    return "\n".join(lines)


# ══════════════════════════════════════════════════════════════════════════════
#  STEP 4 — OLLAMA  (text + question >> streaming answer)
# ══════════════════════════════════════════════════════════════════════════════

SYSTEM_PROMPT = """\
You are an expert data analyst assistant.
You are analyzing a spreadsheet called "{filename}".
Tables available (each labelled by its source sheet): {sheet_list}

Each data block you receive contains:
  1. TABLE OVERVIEW  — source sheet, total rows, column names
  2. COLUMN STATISTICS or COLUMN NOTES  — per-column type and figures
  3. DATA SAMPLE / STATEMENT ROWS  — rows as a text table

Rules:
  • Totals / averages / min / max on a FLAT table >> read from COLUMN STATISTICS.
  • For a table marked TYPE: financial statement >> NEVER sum a column. The column
    already contains subtotals and totals mixed with line items. Read the specific
    labelled row instead (e.g. "Total Assets", "Net Income").
  • Row lookups and pattern questions >> use the data rows shown.
  • Always state which table/sheet an answer came from.
  • If data is truncated, say so and use stats for aggregate answers.
  • Use markdown tables when comparing multiple values.
  • Show brief workings for any calculations.
  • Never invent data that is not in the provided text.
  • If a question cannot be answered from the data, say so clearly.{statement_note}\
"""


# ══════════════════════════════════════════════════════════════════════════════
#  ROUTING — pick which table(s) to send based on the question
# ══════════════════════════════════════════════════════════════════════════════

CTX_CHAR_BUDGET = 9000   # keep concatenated tables within reach of num_ctx=4096

_WHOLE_WB_RE = re.compile(
    r"\b(each|all|every|per)\s+(sheet|tab|table|fund)s?\b|\bevery\s+one\b|"
    r"\bwhole\s+(workbook|file|spreadsheet)\b|\bsummar(y|ise|ize)\b|\boverview\s+of\b",
    re.I,
)

def _norm(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())

def route_question(question, regions, selected_label=None):
    """
    Decide what data to send.
      ('catalog', [all labels])  -> compact index of every table (whole-workbook Qs)
      ('tables',  [labels])      -> full text for named/selected tables
    Precedence: whole-workbook phrasing > tables named in the question >
                explicit tab selection > fallback.
    """
    q  = question or ""
    qn = _norm(q)

    if _WHOLE_WB_RE.search(q):
        return "catalog", [r["label"] for r in regions]

    named = []
    for r in regions:
        for cand in (r["label"], r["sheet"]):
            cn = _norm(cand)
            if len(cn) >= 4 and cn in qn:
                named.append(r["label"]); break
    named = list(dict.fromkeys(named))
    if named:
        return "tables", named

    if selected_label:
        return "tables", [selected_label]

    if len(regions) == 1:
        return "tables", [regions[0]["label"]]
    return "catalog", [r["label"] for r in regions]

def build_catalog(regions):
    lines = ["TABLE CATALOG (each table's label, source sheet, size, type — "
             "ask about a specific one to see its rows):"]
    for r in regions:
        lines.append(f"  - {r['label']}  [sheet={r['sheet']}, "
                     f"{r['rows']}x{r['cols']}, {r['kind']}]")
    return "\n".join(lines)

def build_user_msg(texts, regions, sheet, question):
    """
    Route the question to the right table(s). `sheet` is the label of the
    currently-selected tab (may be None). Named tables in the question override it.
    """
    mode, labels = route_question(question, regions, sheet)

    if mode == "catalog":
        block = build_catalog(regions)
        return (f"DATA (catalog only — name a table to see its rows):\n\n{block}"
                f"\n\n---\nQUESTION: {question}")

    chosen, total, dropped = [], 0, []
    for lbl in labels:
        t = texts.get(lbl, "")
        if not t:
            continue
        if total + len(t) > CTX_CHAR_BUDGET and chosen:
            dropped.append(lbl); continue
        chosen.append(t); total += len(t)
    body = "\n\n".join(chosen) if chosen else "(no matching table found)"
    if dropped:
        body += (f"\n\n[NOTE: {len(dropped)} more table(s) omitted for length: "
                 f"{', '.join(dropped)}. Ask about them individually.]")
    return f"DATA:\n\n{body}\n\n---\nQUESTION: {question}"


def build_system(info: dict) -> str:
    labels = list(info["sheets"].keys())
    statements = [m["label"] for m in info.get("regions", []) if m.get("kind") == "statement"]
    note = ""
    if statements:
        note = ("\n  • These tables are financial statements — do not sum their "
                "columns: " + ", ".join(statements))
    return SYSTEM_PROMPT.format(
        filename=info["name"],
        sheet_list=", ".join(labels),
        statement_note=note,
    )


def ollama_stream(model: str, system: str, messages: list):
    payload = {
        "model":    model,
        "stream":   True,
        "messages": [{"role": "system", "content": system}] + messages,
        "options":  {"temperature": 0.15, "num_ctx": 4096},
    }
    with requests.post(
        f"{OLLAMA_BASE_URL}/api/chat",
        json=payload,
        stream=True,
        timeout=180,
    ) as resp:
        resp.raise_for_status()
        for raw_line in resp.iter_lines():
            if not raw_line:
                continue
            chunk = json.loads(raw_line)
            token = chunk.get("message", {}).get("content", "")
            if token:
                yield token
            if chunk.get("done"):
                break


# ══════════════════════════════════════════════════════════════════════════════
#  ROUTES
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/upload", methods=["POST"])
def upload():
    """
    POST multipart/form-data  field: 'file'

    Pipeline: extract_regions (eparse) >> clean_dataframe >> dataframe_to_text
    Each detected table becomes an addressable region keyed by a label that
    embeds its sheet name.
    """
    if "file" not in request.files:
        return jsonify({"error": "No 'file' field in request"}), 400

    results = []
    for f in request.files.getlist("file"):
        if not f.filename:
            continue
        print(f"\n>> Uploading: {f.filename}")
        try:
            raw     = f.read()
            regions = extract_regions(raw, f.filename)

            sheets = {}   # label -> cleaned df
            texts  = {}   # label -> text
            meta   = []   # per-region metadata (for API + system prompt)

            for r in regions:
                clean = clean_dataframe(r["df"])
                if clean.empty:
                    continue
                kind  = detect_table_shape(clean)
                label = r["label"]
                sheets[label] = clean
                texts[label]  = dataframe_to_text(
                    clean, label, sheet_name=r["sheet"],
                    is_statement=(kind == "statement"),
                )
                meta.append({
                    "label":  label,
                    "sheet":  r["sheet"],
                    "kind":   kind,
                    "source": r["source"],
                    "anchor": r["anchor"],
                    "rows":   len(clean),
                    "cols":   len(clean.columns),
                })

            if not sheets:
                results.append({"name": f.filename, "error": "No readable tables found"})
                continue

            # Terminal preview
            for m in meta:
                print(f"  [{m['kind']:9}] {m['label']}  "
                      f"(sheet={m['sheet']}, {m['rows']}x{m['cols']}, {m['source']})")

            file_id = "f{}_{}".format(
                len(uploaded_files) + 1,
                re.sub(r"[^a-zA-Z0-9_]", "_", f.filename)
            )
            uploaded_files[file_id] = {
                "id":      file_id,
                "name":    f.filename,
                "sheets":  sheets,
                "texts":   texts,
                "regions": meta,
            }
            results.append({
                "id":      file_id,
                "name":    f.filename,
                "sheets":  list(sheets.keys()),                       # labels (frontend tabs)
                "tables":  meta,                                      # rich structure
                "rows":    {lbl: len(df) for lbl, df in sheets.items()},
                "columns": {lbl: list(df.columns) for lbl, df in sheets.items()},
            })

        except Exception as e:
            traceback.print_exc()
            results.append({"name": f.filename, "error": str(e)})

    return jsonify({"uploaded": results})


@app.route("/files", methods=["GET"])
def list_files():
    return jsonify({"files": [
        {"id": fid, "name": info["name"],
         "sheets": list(info["sheets"].keys()),
         "rows": {s: len(df) for s, df in info["sheets"].items()}}
        for fid, info in uploaded_files.items()
    ]})


@app.route("/files/<file_id>", methods=["DELETE"])
def delete_file(file_id):
    if file_id in uploaded_files:
        del uploaded_files[file_id]
        return jsonify({"deleted": file_id})
    return jsonify({"error": "Not found"}), 404


@app.route("/files", methods=["DELETE"])
def clear_files():
    uploaded_files.clear()
    return jsonify({"cleared": True})


@app.route("/preview/<file_id>", methods=["GET"])
def preview(file_id):
    """Return JSON row preview (used by the HTML table view)."""
    if file_id not in uploaded_files:
        return jsonify({"error": "File not found"}), 404
    info  = uploaded_files[file_id]
    sheet = request.args.get("sheet", list(info["sheets"].keys())[0])
    n     = int(request.args.get("n", 20))
    df    = info["sheets"].get(sheet)
    if df is None:
        return jsonify({"error": "Sheet not found"}), 404
    return jsonify({
        "file":    info["name"],
        "sheet":   sheet,
        "columns": list(df.columns),
        "rows":    df.head(n).fillna("").astype(str).values.tolist(),
        "total":   len(df),
    })


@app.route("/text/<file_id>", methods=["GET"])
def get_text(file_id):
    """
    Return the exact text that will be sent to Ollama.
    Useful for verifying the conversion is correct.

        GET /text/<file_id>?sheet=<label>
    """
    if file_id not in uploaded_files:
        return jsonify({"error": "File not found"}), 404
    info  = uploaded_files[file_id]
    sheet = request.args.get("sheet")
    texts = info["texts"]
    body  = texts[sheet] if (sheet and sheet in texts) else "\n\n".join(texts.values())
    return Response(body, content_type="text/plain; charset=utf-8")


@app.route("/chat", methods=["POST"])
def chat():
    """
    POST { file_id, question, sheet?, model?, history? }
    `sheet` is a region label (e.g. "Dashboard · Asset" or "Income Statement").
    Streams Ollama response as plain text.
    """
    body     = request.get_json(force=True)
    file_id  = body.get("file_id", "").strip()
    question = body.get("question", "").strip()
    sheet    = body.get("sheet") or None
    model    = body.get("model") or DEFAULT_MODEL
    history  = body.get("history", [])

    if not file_id:
        return jsonify({"error": "file_id required"}), 400
    if not question:
        return jsonify({"error": "question required"}), 400
    if file_id not in uploaded_files:
        return jsonify({"error": f"File not loaded: {file_id}"}), 404

    info     = uploaded_files[file_id]
    system   = build_system(info)
    user_msg = build_user_msg(info["texts"], info["regions"], sheet, question)
    messages = list(history) + [{"role": "user", "content": user_msg}]

    print(f"\n[CHAT] {info['name']} | table={sheet} | model={model}")
    print(f"       Q: {question[:100]}")

    def generate():
        try:
            for token in ollama_stream(model, system, messages):
                yield token
        except requests.ConnectionError:
            yield "\n\n❌ Cannot reach Ollama. Run: `ollama serve`"
        except requests.HTTPError as e:
            yield f"\n\n❌ Ollama HTTP error: {e}"
        except Exception as e:
            traceback.print_exc()
            yield f"\n\n❌ Server error: {e}"

    return Response(
        stream_with_context(generate()),
        content_type="text/plain; charset=utf-8",
        headers={"X-Accel-Buffering": "no", "Cache-Control": "no-cache"},
    )


@app.route("/models", methods=["GET"])
def list_models():
    try:
        r      = requests.get(f"{OLLAMA_BASE_URL}/api/tags", timeout=5)
        models = [m["name"] for m in r.json().get("models", [])]
        return jsonify({"models": models})
    except Exception as e:
        return jsonify({"models": [], "error": str(e)})


@app.route("/pull", methods=["POST"])
def pull_model():
    """
    POST { "model": "qwen3:1.7b" }

    Streams pull progress back as newline-delimited JSON, forwarded
    directly from Ollama's /api/pull endpoint.
    """
    body  = request.get_json(force=True)
    model = (body.get("model") or "").strip()

    if not model:
        return jsonify({"error": "model name required"}), 400

    print(f"\n[PULL] Starting pull: {model}")

    def generate():
        try:
            with requests.post(
                f"{OLLAMA_BASE_URL}/api/pull",
                json={"name": model, "stream": True},
                stream=True,
                timeout=3600,   # large models can take a long time
            ) as resp:
                resp.raise_for_status()
                for raw_line in resp.iter_lines():
                    if raw_line:
                        print(f"  [PULL] {raw_line[:120]}")
                        yield raw_line + b"\n"
        except requests.ConnectionError:
            yield json.dumps({"error": "Cannot reach Ollama"}).encode() + b"\n"
        except Exception as e:
            traceback.print_exc()
            yield json.dumps({"error": str(e)}).encode() + b"\n"

    return Response(
        stream_with_context(generate()),
        content_type="application/x-ndjson",
        headers={"X-Accel-Buffering": "no", "Cache-Control": "no-cache"},
    )


@app.route("/delete-model", methods=["POST"])
def delete_model():
    """POST { "model": "qwen3:1.7b" } — removes a model from Ollama."""
    body  = request.get_json(force=True)
    model = (body.get("model") or "").strip()
    if not model:
        return jsonify({"error": "model name required"}), 400
    try:
        r = requests.delete(
            f"{OLLAMA_BASE_URL}/api/delete",
            json={"name": model},
            timeout=30,
        )
        return jsonify({"deleted": model, "status": r.status_code})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/health", methods=["GET"])
def health():
    ollama_ok = False
    try:
        r         = requests.get(f"{OLLAMA_BASE_URL}/api/tags", timeout=3)
        ollama_ok = r.ok
    except Exception:
        pass
    return jsonify({
        "status":       "ok",
        "ollama":       ollama_ok,
        "files_loaded": len(uploaded_files),
        "model":        DEFAULT_MODEL,
    })


# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("=" * 60)
    print("  Excel AI Analyzer — Backend (eparse multi-table)")
    print(f"  Ollama URL    : {OLLAMA_BASE_URL}")
    print(f"  Default model : {DEFAULT_MODEL}")
    print(f"  Server        : http://localhost:3000")
    print("")
    print("  Routes:")
    print("    POST   /upload          upload Excel/CSV files")
    print("    POST   /chat            ask a question (streaming)")
    print("    GET    /text/<id>       see exact text sent to Ollama")
    print("    GET    /preview/<id>    JSON row preview")
    print("    GET    /files           list loaded files")
    print("    DELETE /files/<id>      remove one file")
    print("    DELETE /files           remove all files")
    print("    GET    /models          list Ollama models")
    print("    GET    /health          server + Ollama status")
    print("=" * 60)
    app.run(host="0.0.0.0", port=3000, debug=False)
