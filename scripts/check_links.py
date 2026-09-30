from __future__ import annotations
import json,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
from app.core.function_registry import HANDLERS
def load(name):
    return json.loads((ROOT/"config"/name).read_text(encoding="utf-8"))
defs=load("functions.json");deps=load("function_dependencies.json");layout=load("ui_layout.json")
ids=[]
for g in defs["groups"]:
    for fn in g["functions"]:ids.append(fn["id"])
errors=[]
if len(ids)!=len(set(ids)):errors.append("Duplicate function IDs")
for fid in ids:
    if fid not in HANDLERS:errors.append(f"No handler: {fid}")
    if fid not in deps:errors.append(f"No dependency manifest: {fid}")
for fid in deps:
    if fid not in ids:errors.append(f"Dependency manifest has unknown ID: {fid}")
    for upstream in deps[fid].get("depends_on",[]):
        if upstream not in ids:errors.append(f"{fid}: unknown dependency {upstream}")
for fid in (layout.get("functions",{}) if isinstance(layout,dict) else {}):
    if fid not in ids:errors.append(f"Layout references unknown function: {fid}")
print(f"Functions: {len(ids)}")
print(f"Handlers: {len(HANDLERS)}")
print(f"Dependency records: {len(deps)}")
if errors:
    print("ERRORS:")
    for e in errors:print(" -",e)
    raise SystemExit(1)
print("LINK CHECK: PASS")
