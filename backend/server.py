"""
Excel AI Analyzer — Python Backend
Uses Flask + Pandas + Ollama to analyze uploaded Excel/CSV files.

Install:
    pip install flask flask-cors pandas openpyxl xlrd requests numpy

Run:
    python server.py

Ollama must be running:
    ollama serve
    ollama pull qwen2.5
"""

import io
import json
import os
import re
import traceback

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

app = Flask(__name__)
CORS(app)

# ── In-memory store ────────────────────────────────────────────────────────────
# { file_id: { id, name, sheets: {name: df}, texts: {name: str} } }
uploaded_files: dict = {}

# ── Config ─────────────────────────────────────────────────────────────────────
OLLAMA_BASE_URL = os.getenv("OLLAMA_BASE_URL", "http://localhost:11434")
DEFAULT_MODEL   = os.getenv("OLLAMA_MODEL",   "qwen2.5")

MAX_SAMPLE_ROWS = 15    # rows shown in the data sample sent to Ollama (small = fast CPU prefill; aggregates come from COLUMN STATISTICS which cover ALL rows)
MAX_COLS        = 30    # trim very wide sheets


# ══════════════════════════════════════════════════════════════════════════════
#  STEP 1 — READ  (upload bytes >> dict of DataFrames)
# ══════════════════════════════════════════════════════════════════════════════

def read_excel_bytes(raw: bytes, filename: str) -> dict:
    """
    Read raw bytes from an Excel/CSV upload.
    Returns { sheet_name: pd.DataFrame }

    Handles:
      .xlsx / .xlsm  >> openpyxl
      .xls           >> xlrd  (legacy binary format)
      .csv           >> auto-detects encoding
    """
    ext = filename.rsplit(".", 1)[-1].lower()
    buf = io.BytesIO(raw)
    sheets = {}

    # ── CSV ──────────────────────────────────────────────────────────────────
    if ext == "csv":
        for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
            try:
                buf.seek(0)
                df = pd.read_csv(buf, encoding=enc, on_bad_lines="skip",
                                 low_memory=False)
                sheets["Data"] = df
                print(f"  [CSV] encoding={enc}  rows={len(df)}")
                break
            except (UnicodeDecodeError, Exception):
                continue
        if not sheets:
            raise ValueError("Could not decode CSV with any common encoding.")

    # ── XLS (legacy binary) ──────────────────────────────────────────────────
    elif ext == "xls":
        buf.seek(0)
        xl = pd.ExcelFile(buf, engine="xlrd")
        for name in xl.sheet_names:
            df = xl.parse(name, na_values=["", "N/A", "NA", "n/a", "#N/A"])
            sheets[name] = df
        print(f"  [XLS]  sheets={xl.sheet_names}")

    # ── XLSX / XLSM ──────────────────────────────────────────────────────────
    elif ext in ("xlsx", "xlsm"):
        buf.seek(0)
        xl = pd.ExcelFile(buf, engine="openpyxl")
        for name in xl.sheet_names:
            df = xl.parse(name, na_values=["", "N/A", "NA", "n/a", "#N/A"])
            sheets[name] = df
        print(f"  [XLSX] sheets={xl.sheet_names}")

    else:
        raise ValueError(
            f"Unsupported extension: .{ext}  "
            "(supported: .xlsx  .xls  .xlsm  .csv)"
        )

    return sheets


# ══════════════════════════════════════════════════════════════════════════════
#  STEP 2 — CLEAN  (DataFrame >> tidy DataFrame)
# ══════════════════════════════════════════════════════════════════════════════

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
        if not c or c.lower().startswith("unnamed"):
            c = f"col_{len(new_cols) + 1}"
        if c in seen:
            seen[c] += 1
            c = f"{c}_{seen[c]}"
        else:
            seen[c] = 0
        new_cols.append(c)
    df.columns = new_cols

    # 4. Try to parse date-like string columns
    for col in df.select_dtypes(include=["object"]).columns:
        sample = df[col].dropna().head(20)
        if sample.empty:
            continue
        parsed = pd.to_datetime(sample, errors="coerce")
        if parsed.notna().mean() > 0.7:
            df[col] = pd.to_datetime(df[col], errors="coerce")

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


def dataframe_to_text(df: pd.DataFrame, sheet_name: str) -> str:
    """
    Convert a cleaned DataFrame to a structured plain-text document.

    Layout:
        SHEET OVERVIEW       — row/col counts, column list
        COLUMN STATISTICS    — type + full-dataset stats per column
        DATA SAMPLE          — aligned text table (first MAX_SAMPLE_ROWS rows)
    """
    lines = []
    n_rows, n_cols = df.shape
    is_truncated   = n_rows > MAX_SAMPLE_ROWS

    # ── Overview ──────────────────────────────────────────────────────────────
    lines.append("=" * 68)
    lines.append(f"SHEET: {sheet_name}")
    lines.append("=" * 68)
    lines.append(f"Total rows    : {n_rows:,}")
    lines.append(f"Total columns : {n_cols}")
    lines.append(f"Columns       : {', '.join(df.columns.tolist())}")
    lines.append("")

    # ── Column statistics (full dataset) ─────────────────────────────────────
    lines.append("-" * 68)
    lines.append("COLUMN STATISTICS  (computed on ALL rows)")
    lines.append("-" * 68)
    for col in df.columns:
        series = df[col]
        label  = dtype_label(series)
        lines.append(f"  [{label}]  {col}")
        if pd.api.types.is_numeric_dtype(series):
            lines.append(numeric_stats_line(series))
        elif pd.api.types.is_datetime64_any_dtype(series):
            lines.append(date_stats_line(series))
        else:
            lines.append(text_stats_line(series))
    lines.append("")

    # ── Data sample ───────────────────────────────────────────────────────────
    lines.append("-" * 68)
    if is_truncated:
        lines.append(
            f"DATA SAMPLE  (first {MAX_SAMPLE_ROWS} of {n_rows:,} rows shown)\n"
            f"NOTE: Use the statistics above for totals/averages — "
            f"they cover ALL {n_rows:,} rows."
        )
    else:
        lines.append(f"FULL DATA  ({n_rows} rows)")
    lines.append("-" * 68)

    sample = df.head(MAX_SAMPLE_ROWS).copy()

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
        lines.append(f"\n  ... ({n_rows - MAX_SAMPLE_ROWS:,} more rows not shown)")

    lines.append("=" * 68)
    return "\n".join(lines)


def build_file_texts(sheets: dict) -> dict:
    """Convert all cleaned sheets to text. Returns { sheet_name: text }"""
    return {name: dataframe_to_text(df, name) for name, df in sheets.items()}


# ══════════════════════════════════════════════════════════════════════════════
#  STEP 4 — OLLAMA  (text + question >> streaming answer)
# ══════════════════════════════════════════════════════════════════════════════

SYSTEM_PROMPT = """\
You are an expert data analyst assistant.
You are analyzing a spreadsheet called "{filename}".
Sheets available: {sheet_list}

Each data block you receive contains:
  1. SHEET OVERVIEW  — total rows and column names
  2. COLUMN STATISTICS  — computed from the FULL dataset (use these for aggregates)
  3. DATA SAMPLE  — first {sample_rows} rows as a text table

Rules:
  • Totals / averages / min / max >> always read from COLUMN STATISTICS.
  • Row lookups and pattern questions >> use the DATA SAMPLE.
  • If data is truncated, say so and use stats for aggregate answers.
  • Use markdown tables when comparing multiple values.
  • Show brief workings for any calculations.
  • Never invent data that is not in the provided text.
  • If a question cannot be answered from the data, say so clearly.\
"""


def build_system(info: dict) -> str:
    return SYSTEM_PROMPT.format(
        filename=info["name"],
        sheet_list=", ".join(info["sheets"].keys()),
        sample_rows=MAX_SAMPLE_ROWS,
    )


def build_user_msg(texts: dict, sheet: str, question: str) -> str:
    if sheet and sheet in texts:
        data_block = texts[sheet]
    else:
        data_block = "\n\n".join(texts.values())
    return f"DATA:\n\n{data_block}\n\n---\nQUESTION: {question}"


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

    Pipeline: read_excel_bytes >> clean_dataframe >> dataframe_to_text
    All three representations stored in memory.
    """
    if "file" not in request.files:
        return jsonify({"error": "No 'file' field in request"}), 400

    results = []
    for f in request.files.getlist("file"):
        if not f.filename:
            continue
        print(f"\n>> Uploading: {f.filename}")
        try:
            raw    = f.read()
            sheets = read_excel_bytes(raw, f.filename)
            clean  = {name: clean_dataframe(df) for name, df in sheets.items()}
            texts  = build_file_texts(clean)

            # Terminal preview
            for name, txt in texts.items():
                preview = txt.split("\n")[:10]
                print(f"  [TEXT PREVIEW — {name}]")
                for ln in preview:
                    print(f"    {ln}")
                print(f"    ... total {len(txt):,} chars")

            file_id = "f{}_{}".format(
                len(uploaded_files) + 1,
                re.sub(r"[^a-zA-Z0-9_]", "_", f.filename)
            )
            uploaded_files[file_id] = {
                "id":     file_id,
                "name":   f.filename,
                "sheets": clean,
                "texts":  texts,
            }
            results.append({
                "id":      file_id,
                "name":    f.filename,
                "sheets":  list(clean.keys()),
                "rows":    {s: len(df) for s, df in clean.items()},
                "columns": {s: list(df.columns) for s, df in clean.items()},
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

        GET /text/<file_id>?sheet=Sheet1
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
    user_msg = build_user_msg(info["texts"], sheet, question)
    messages = list(history) + [{"role": "user", "content": user_msg}]

    print(f"\n[CHAT] {info['name']} | sheet={sheet} | model={model}")
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

    Each line is one of:
      {"status":"pulling manifest"}
      {"status":"downloading ...","completed":N,"total":N}
      {"status":"success"}
      {"error":"..."}
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
    print("  Excel AI Analyzer — Backend")
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
