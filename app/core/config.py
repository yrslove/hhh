from __future__ import annotations
import json
from .paths import CONFIG_DIR
def load_json(name:str,default=None):
    p=CONFIG_DIR/name
    if not p.exists():return {} if default is None else default
    with p.open("r",encoding="utf-8") as f:return json.load(f)
def save_json(name:str,data):
    p=CONFIG_DIR/name;tmp=p.with_suffix(p.suffix+".tmp")
    tmp.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding="utf-8");tmp.replace(p)