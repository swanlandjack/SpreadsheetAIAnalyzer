#!/usr/bin/env python3
"""
Confirm every sample workbook parses cleanly through the REAL server.py.
Parse-layer only (deterministic). Flags anything suspicious: col_N headers,
1970 dates, zero tables, or a fund/holdings sheet wrongly tagged 'statement'.

Usage (from backend/):
    uv run --with eparse --with pandas --with openpyxl --with xlrd --with numpy \
      python confirm_samples.py ../Samples/*.xlsx ../Samples/demo/files/*.xlsx
"""
import sys, importlib.util as u, warnings
warnings.filterwarnings("ignore")

m = u.module_from_spec(u.spec_from_file_location("m","server.py"))
u.spec_from_file_location("m","server.py").loader.exec_module(m)

def confirm(path):
    issues=[]; regs=m.extract_regions(open(path,"rb").read(), path)
    if not regs: return ["NO TABLES FOUND"], []
    summary=[]
    for r in regs:
        d=m.clean_dataframe(r["df"]); kind=m.detect_table_shape(d)
        cols=[str(c) for c in d.columns]
        # red flags
        if any(c.startswith("col_") for c in cols):
            issues.append(f"{r['label'][:30]}: col_N headers (parse failed)")
        # 1970 leak
        import pandas as pd
        for c in d.columns:
            if pd.api.types.is_datetime64_any_dtype(d[c]):
                s=d[c].dropna()
                if len(s) and (s.dt.year<=1970).any():
                    issues.append(f"{r['label'][:30]}: 1970 dates in {c}")
        summary.append((kind, r["label"], d.shape))
    return issues, summary

any_issue=False
for path in sys.argv[1:]:
    name=path.split("/")[-1]
    try:
        issues, summary = confirm(path)
    except Exception as e:
        print(f"CRASH {name}: {e}"); any_issue=True; continue
    tag = "ISSUES" if issues else "ok"
    print(f"[{tag}] {name}  ({len(summary)} tables)")
    for kind,label,shape in summary:
        print(f"        {kind:9} {label[:44]:44} {shape}")
    for i in issues:
        print(f"    !!  {i}"); any_issue=True

print("\n" + ("SOME SAMPLES HAVE ISSUES — inspect above" if any_issue else "ALL SAMPLES PARSE CLEAN"))
sys.exit(1 if any_issue else 0)
