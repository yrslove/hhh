"""Capture sanitized Doctor Eleks procedure saves during one manual UI run.

Run from the repository root. This script only listens to browser traffic; it
does not click, submit, modify responses, or read authorization headers.
"""
from __future__ import annotations

import json
import sys
import time
from datetime import datetime
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))

from playwright.sync_api import sync_playwright
from app.workflows.procedure_form import (
    _present, _sanitized_response, sanitize_payload, structured_payload_diff,
)

OUTPUT=ROOT/"runtime"/"procedure-manual-capture.jsonl"
SUCCESS=ROOT/"runtime"/"procedure-manual-success.json"
FINAL=ROOT/"runtime"/"procedure-manual-final-success.json"
DIFF=ROOT/"runtime"/"procedure-manual-diff.json"


def _complete(payload):
    return (payload.get("procedurePeriod")=="30" and
            str(payload.get("is_date_or_period")).casefold()=="true" and
            payload.get("selectedReferralRef")=="digitalReferral" and
            all(_present(payload.get(key)) for key in (
                "id","based_on","code","performer","division",
                "procedurePeriodStart","procedurePeriodEnd","performed_date_time")))


def main():
    OUTPUT.parent.mkdir(parents=True,exist_ok=True)
    pending=[]
    complete=None
    final=None
    sequence=0
    with sync_playwright() as playwright:
        browser=playwright.chromium.connect_over_cdp("http://127.0.0.1:9222")
        page=next(page for context in browser.contexts for page in context.pages
                  if "ehealth.eleks.com" in page.url)

        def on_request(request):
            nonlocal sequence
            if request.method!="POST" or "/api/run/portal" not in request.url:
                return
            try:body=request.post_data_json
            except Exception:return
            if not isinstance(body,dict) or body.get("name")!="spiprocedureSave":
                return
            sequence+=1
            pending.append({"request":request,"sequence":sequence,
                            "timestamp":datetime.now().astimezone().isoformat(),
                            "last_action":"MANUAL_UI","payload":sanitize_payload(body)})

        def on_response(response):
            nonlocal complete,final
            record=next((item for item in pending if item["request"] is response.request),None)
            if record is None:return
            pending.remove(record)
            record.pop("request",None)
            record["status"]=response.status
            try:record["response"]=_sanitized_response(response.json())
            except Exception:record["response"]={"type":"unreadable"}
            with OUTPUT.open("a",encoding="utf-8") as stream:
                stream.write(json.dumps(record,ensure_ascii=False)+"\n")
            payload=record["payload"]
            print(json.dumps({"sequence":record["sequence"],"status":record["status"],
                              "procedurePeriod":payload.get("procedurePeriod"),
                              "toValidate":payload.get("toValidate")},ensure_ascii=False),flush=True)
            if 200<=response.status<300 and _complete(payload):
                complete=record
                SUCCESS.write_text(json.dumps(record,ensure_ascii=False,indent=2),encoding="utf-8")
                failed=None
                source=ROOT/"runtime"/"procedure-save-capture.jsonl"
                if source.exists():
                    for line in source.read_text(encoding="utf-8").splitlines():
                        try:item=json.loads(line)
                        except json.JSONDecodeError:continue
                        candidate=item.get("save_payload") or {}
                        if (item.get("event")=="AUTOSAVE_FAILURE" and
                                _present(candidate.get("based_on")) and
                                str(candidate.get("code","")).casefold()==
                                    str(payload.get("code","")).casefold()):
                            failed=item
                if failed:
                    DIFF.write_text(json.dumps({
                        "failed_status":failed.get("status"),
                        "manual_status":record.get("status"),
                        "failed_action":failed.get("last_action"),
                        "differences":structured_payload_diff(
                            {"payload":failed["save_payload"]},record)},
                        ensure_ascii=False,indent=2),encoding="utf-8")
            if (complete and 200<=response.status<300 and
                    payload.get("id")==complete["payload"].get("id") and
                    str(payload.get("toValidate"))=="1"):
                final=record
                FINAL.write_text(json.dumps(record,ensure_ascii=False,indent=2),encoding="utf-8")

        page.on("request",on_request)
        page.on("response",on_response)
        print("MANUAL_CAPTURE_READY",flush=True)
        deadline=time.monotonic()+900
        try:
            while time.monotonic()<deadline and final is None:
                page.wait_for_timeout(200)
        finally:
            page.remove_listener("request",on_request)
            page.remove_listener("response",on_response)
            browser.close()
        print("MANUAL_CAPTURE_FINAL_SUCCESS" if final else "MANUAL_CAPTURE_STOPPED",flush=True)


if __name__=="__main__":
    main()
