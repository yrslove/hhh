from __future__ import annotations
import json, uuid
from copy import deepcopy
from .config import load_json, save_json

FILE="ui_layout.json"

def _defaults(functions:dict):
    groups=[]
    for g in functions.get("groups",[]):
        groups.append({"id":g["id"],"title":g["title"],"order":g.get("order",0)})
    items={}
    for g in functions.get("groups",[]):
        for i,fn in enumerate(g.get("functions",[])):
            items[fn["id"]]={"title":fn["title"],"note":"","group_id":g["id"],"order":i}
    return {"groups":groups,"functions":items}

def load_layout(functions:dict):
    data=load_json(FILE,{})
    defaults=_defaults(functions)
    if not data or "groups" not in data:
        save_json(FILE,defaults); return defaults
    existing={g["id"] for g in data["groups"]}
    for g in defaults["groups"]:
        if g["id"] not in existing:data["groups"].append(g)
    for fid,item in defaults["functions"].items():
        data.setdefault("functions",{}).setdefault(fid,item)
    save_json(FILE,data)
    return data

def save_layout(data):
    # sanitize
    groups=[]
    seen=set()
    for idx,g in enumerate(data.get("groups",[])):
        gid=str(g.get("id") or f"group_{uuid.uuid4().hex[:10]}")
        if gid in seen: continue
        seen.add(gid)
        groups.append({"id":gid,"title":str(g.get("title","Група"))[:80],"order":idx})
    funcs={}
    for fid,x in data.get("functions",{}).items():
        funcs[str(fid)]={
            "title":str(x.get("title",""))[:120],
            "note":str(x.get("note",""))[:300],
            "group_id":str(x.get("group_id","")),
            "order":int(x.get("order",0)),
        }
    payload={"groups":groups,"functions":funcs}
    save_json(FILE,payload); return payload

def create_group(functions:dict,title:str):
    layout=load_layout(functions)
    gid="group_"+uuid.uuid4().hex[:10]
    layout["groups"].append({"id":gid,"title":title[:80] or "Нова група","order":len(layout["groups"])})
    return save_layout(layout)

def rename_group(functions:dict,gid:str,title:str):
    layout=load_layout(functions)
    for g in layout["groups"]:
        if g["id"]==gid:g["title"]=title[:80] or g["title"]
    return save_layout(layout)
