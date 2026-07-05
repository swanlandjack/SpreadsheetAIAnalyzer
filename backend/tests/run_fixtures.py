#!/usr/bin/env python3
"""
Parser regression runner. Imports the REAL parse functions from server.py and
asserts each fixture against expected.json. Parse-layer only — no model, fully
deterministic.

Usage (from backend/, or point --server at server.py):
    python run_fixtures.py
    python run_fixtures.py --server ../server.py --files tests/fixtures/files
"""
import argparse, importlib.util, json, os, sys, traceback

def load_server(path):
    spec = importlib.util.spec_from_file_location("srv", path)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m

def regions_for(srv, path):
    """Run the real pipeline: extract_regions -> clean -> detect_table_shape."""
    raw = open(path, "rb").read()
    regs = srv.extract_regions(raw, os.path.basename(path))
    out = []
    for r in regs:
        clean = srv.clean_dataframe(r["df"])
        if clean.empty:
            continue
        kind = srv.detect_table_shape(clean)
        out.append({
            "label": r["label"], "sheet": r["sheet"], "kind": kind,
            "rows": len(clean), "cols": len(clean.columns),
            "columns": [str(c) for c in clean.columns],
            "has_1970": _has_1970(clean),
        })
    return out

def _has_1970(df):
    import pandas as pd
    for c in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[c]):
            s = df[c].dropna()
            if len(s) and (s.dt.year <= 1970).any():
                return True
    return False

def check(name, spec, regs):
    """Return list of failure strings ([] = pass)."""
    fails = []
    if spec.get("must_not_crash") and regs is None:
        return ["crashed"]
    if regs is None:
        return ["crashed (no regions)"]
    if "min_tables" in spec and len(regs) < spec["min_tables"]:
        fails.append(f"tables {len(regs)} < min {spec['min_tables']}")
    if "sheets_present" in spec:
        got = {r["sheet"] for r in regs}
        for s in spec["sheets_present"]:
            if s not in got:
                fails.append(f"sheet {s!r} missing")
    for i, tspec in enumerate(spec.get("tables", [])):
        # match against the best-fitting region (any that satisfies), else region i
        cands = regs
        if "must_have_header" in tspec:
            h = tspec["must_have_header"]
            cands = [r for r in regs if any(h.lower() in c.lower() for c in r["columns"])]
            if not cands:
                fails.append(f"table[{i}] no region has header ~{h!r}")
                continue
        r = cands[0] if cands else (regs[i] if i < len(regs) else None)
        if r is None:
            fails.append(f"table[{i}] absent"); continue
        if "kind" in tspec and r["kind"] != tspec["kind"]:
            fails.append(f"table[{i}] kind={r['kind']} expected {tspec['kind']}")
        if "min_cols" in tspec and r["cols"] < tspec["min_cols"]:
            fails.append(f"table[{i}] cols {r['cols']} < {tspec['min_cols']}")
        if "min_rows" in tspec and r["rows"] < tspec["min_rows"]:
            fails.append(f"table[{i}] rows {r['rows']} < {tspec['min_rows']}")
        if "max_cols_after_trim" in tspec and r["cols"] > tspec["max_cols_after_trim"]:
            fails.append(f"table[{i}] cols {r['cols']} > trim {tspec['max_cols_after_trim']}")
        if tspec.get("no_1970") and r["has_1970"]:
            fails.append(f"table[{i}] has 1970 dates")
    return fails

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--server", default="server.py")
    ap.add_argument("--files", default="tests/fixtures/files")
    ap.add_argument("--expected", default="tests/fixtures/expected.json")
    ap.add_argument("-v", action="store_true", help="verbose: show regions per file")
    a = ap.parse_args()

    srv = load_server(a.server)
    exp = json.load(open(a.expected))
    files = sorted(os.listdir(a.files))

    npass = nfail = nknown = 0
    for fn in files:
        if fn not in exp:
            continue
        spec = exp[fn]
        known = any("KNOWN" in str(t.get("note","")) for t in spec.get("tables",[])) \
                or "KNOWN" in str(spec.get("note",""))
        try:
            regs = regions_for(srv, os.path.join(a.files, fn))
        except Exception as e:
            regs = None
            if a.v: traceback.print_exc()
        fails = check(fn, spec, regs)
        if a.v and regs is not None:
            for r in regs:
                print(f"      {r['kind']:9} {r['label'][:40]:40} {r['rows']}x{r['cols']}")
        if not fails:
            print(f"PASS  {fn}"); npass += 1
        elif known:
            print(f"KNOWN {fn}: {'; '.join(fails)}"); nknown += 1
        else:
            print(f"FAIL  {fn}: {'; '.join(fails)}"); nfail += 1

    print(f"\n{npass} pass, {nfail} FAIL, {nknown} known-issue  (of {npass+nfail+nknown})")
    sys.exit(1 if nfail else 0)

if __name__ == "__main__":
    main()
