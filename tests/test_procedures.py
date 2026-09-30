import unittest
import tempfile
import time
from pathlib import Path
from datetime import date, datetime
from unittest.mock import patch
from openpyxl import Workbook, load_workbook

from app.workflows.procedures import (_completion_log, _codes_by_type, _specialist_type,
                                     _service_code5, _service_codes_in_row,
                                     _take_referral_to_work, _click_create_procedure,
                                     _select_matching_referral)
import app.workflows.procedures as procedure_workflow
from app.excel.workbook import find_sheet, get_graph_patient, load_procedure_reference
from app.automation.browser import (_ehealth_url, _is_ehealth_application_frame,
                                    _row_has_full_patient_name, _row_has_patient_identity,
                                    _row_matches_birth_date, ManualCheckpoint, AutomationError,
                                    WebEleks, wait_visible_recheck)
from app.core.text import normalize_patient_name
from app.workflows.procedure_form import DoctorEleksProcedureForm, structured_payload_diff, sanitize_payload


class ProcedureCompletionLogTests(unittest.TestCase):
    def test_missing_element_is_reacquired_after_two_second_recheck(self):
        events=[]
        class Element:
            def __init__(self,attempt):self.attempt=attempt
            def wait_for(self,state="visible",timeout=1800):
                events.append(("check",self.attempt))
                if self.attempt==1:raise TimeoutError()
        class Page:
            def wait_for_timeout(self,ms):events.append(("wait",ms))
        attempts=[]
        def locate():
            attempts.append(len(attempts)+1)
            return Element(attempts[-1])
        result=wait_visible_recheck(Page(),locate)
        self.assertEqual(result.attempt,2)
        self.assertEqual(events,[("check",1),("wait",2000),("check",2)])

    def test_browser_alias_lookup_rechecks_after_ui_render(self):
        events=[];state={"ready":False}
        class Locator:
            def count(self):return 1 if state["ready"] else 0
            @property
            def first(self):return self
            def is_visible(self):return state["ready"]
        class Page:
            def wait_for_timeout(self,ms):
                events.append(ms)
                if ms==2000:state["ready"]=True
        web=WebEleks("test",{});web.page=Page()
        web.selector_map={"referral":["referral"]}
        web._candidate_locator=lambda selector:Locator()
        web._log=lambda *args,**kwargs:None
        self.assertTrue(web.locator("referral",timeout=20).is_visible())
        self.assertIn(2000,events)

    def test_content_tab_rechecks_after_render_delay(self):
        events=[];state={"visible":False,"active":False}
        class TabClass:
            def count(self):return 1
            def get_attribute(self,name,timeout=300):return "ct-tab -active" if state["active"] else "ct-tab"
        class Tab:
            def is_visible(self):return state["visible"]
            def bounding_box(self):return {"x":240}
            def locator(self,selector):return TabClass()
            def click(self,timeout=2000):events.append("click");state["active"]=True
        class Matches:
            def count(self):return 1
            def nth(self,index):return Tab()
        class Page:
            def get_by_text(self,label,exact=True):return Matches()
            def wait_for_timeout(self,ms):
                events.append(("wait",ms))
                if ms==2000:state["visible"]=True
        web=WebEleks("test",{});web.page=Page();web._log=lambda *args,**kwargs:None
        self.assertTrue(web.click_content_text("Направлення",timeout=20,retry_click=True))
        self.assertIn(("wait",2000),events)
        self.assertEqual(events.count("click"),1)

    def test_content_tab_retries_failed_navigation_click_once(self):
        events=[];state={"active":False}
        class TabClass:
            def count(self):return 1
            def get_attribute(self,name,timeout=300):return "ct-tab -active" if state["active"] else "ct-tab"
        class Tab:
            def is_visible(self):return True
            def bounding_box(self):return {"x":240}
            def locator(self,selector):return TabClass()
            def click(self,timeout=2000):
                events.append("click")
                if events.count("click")==1:raise TimeoutError()
                state["active"]=True
        class Matches:
            def count(self):return 1
            def nth(self,index):return Tab()
        class Page:
            def get_by_text(self,label,exact=True):return Matches()
            def wait_for_timeout(self,ms):events.append(("wait",ms))
        web=WebEleks("test",{});web.page=Page();web._log=lambda *args,**kwargs:None
        self.assertTrue(web.click_content_text("Направлення",timeout=20,retry_click=True))
        self.assertEqual(events.count("click"),2)
        self.assertIn(("wait",2000),events)

    def test_hidden_active_tab_text_does_not_skip_visible_tab_navigation(self):
        events=[];state={"visible_active":False}
        class TabClass:
            def __init__(self,active):self.active=active
            def count(self):return 1
            def get_attribute(self,name,timeout=300):
                return "ct-tab -active" if self.active else "ct-tab"
        class Tab:
            def __init__(self,visible,active):self.visible=visible;self.active=active
            def is_visible(self):return self.visible
            def bounding_box(self):return {"x":240}
            def locator(self,selector):return TabClass(self.active)
            def click(self,timeout=2000):
                events.append("visible-click")
                state["visible_active"]=True
        class Matches:
            def count(self):return 2
            def nth(self,index):
                return Tab(False,True) if index==0 else Tab(True,state["visible_active"])
        class Page:
            def get_by_text(self,label,exact=True):return Matches()
            def wait_for_timeout(self,ms):events.append(("wait",ms))
        web=WebEleks("test",{});web.page=Page();web._log=lambda *args,**kwargs:None
        self.assertTrue(web.click_content_text("Взаємодії",timeout=20,retry_click=True))
        self.assertEqual(events.count("visible-click"),1)

    def test_active_class_duplicate_does_not_skip_requested_tab_click(self):
        events=[]
        class Tab:
            def __init__(self,active):self.active=active
            def is_visible(self):return True
            def bounding_box(self):return {"x":240,"y":500,"width":90,"height":28}
            def locator(self,selector):
                class Parent:
                    def count(self):return 1
                    def get_attribute(self,name,timeout=300):return "ct-tab -active"
                return Parent()
            def click(self,timeout=2000):events.append("clicked-requested-tab")
        class Matches:
            def count(self):return 1
            def nth(self,index):return Tab(True)
        class Page:
            def get_by_text(self,label,exact=True):return Matches()
            def wait_for_timeout(self,ms):events.append(("wait",ms))
        web=WebEleks("test",{});web.page=Page();web._log=lambda *args,**kwargs:None
        self.assertTrue(web.click_content_text("Взаємодії",timeout=20,retry_click=True))
        self.assertIn("clicked-requested-tab",events)

    def test_failed_mutating_click_is_not_repeated(self):
        events=[]
        class Button:
            def is_visible(self):return True
            def bounding_box(self):return {"x":240}
            def click(self,timeout=2000):events.append("click");raise TimeoutError()
        class Matches:
            def count(self):return 1
            def nth(self,index):return Button()
        class Page:
            def get_by_text(self,label,exact=True):return Matches()
            def wait_for_timeout(self,ms):events.append(("wait",ms))
        web=WebEleks("test",{});web.page=Page();web._log=lambda *args,**kwargs:None
        web.snapshot_error=lambda label:None
        with self.assertRaises(AutomationError):
            web.click_content_text("Взяти у роботу",timeout=20)
        self.assertEqual(events,["click"])

    def test_procedure_reference_loads_only_needed_sheets_without_writing(self):
        with tempfile.TemporaryDirectory() as folder:
            path=Path(folder)/"graph.xlsx"
            wb=Workbook(); distribution=wb.active; distribution.title="Розподіл АР2"
            distribution.cell(2,3).value="Іваненко Іван Іванович"
            distribution.cell(2,18).value="96128/ФТ"
            wb.create_sheet("Unused")["A1"]="not needed"
            schedule=wb.create_sheet("ЕТФТ")
            schedule["C1"]=date(2026,9,3)
            schedule["A2"]="Спеціаліст ФТ"
            schedule["B2"]="09:30"
            schedule["C2"]="Іваненко Іван Іванович"
            wb.save(path); wb.close()
            original=path.read_bytes()
            selected=load_procedure_reference(str(path),"ЕТФТ","Розподіл AR2")
            try:
                self.assertEqual(set(selected.sheetnames),{"Розподіл АР2","ЕТФТ"})
                self.assertEqual(selected["ЕТФТ"]["C2"].value,"Іваненко Іван Іванович")
                self.assertEqual(get_graph_patient(str(path),"Розподіл AR2",
                                 "Іваненко Іван Іванович",workbook=selected)["codes_specialists"],"96128/ФТ")
            finally:
                selected.close()
            self.assertEqual(path.read_bytes(),original)

    def test_lists_each_created_procedure_in_human_readable_form(self):
        items=[
            {"date":"2026-09-05","time":"10:30","specialist":"Буньковський ФТ"},
            {"date":"2026-09-12","time":"11:00","specialist":"Буньковський ФТ"},
        ]
        text=_completion_log(items)
        self.assertIn("Внесено процедур: 2",text)
        self.assertIn("05.09.2026 10:30 — Буньковський ФТ",text)
        self.assertIn("12.09.2026 11:00 — Буньковський ФТ",text)

    def test_reports_when_everything_was_already_confirmed(self):
        self.assertIn("Нових процедур не внесено",_completion_log([]))

    def test_distribution_sheet_accepts_latin_ar2_setting(self):
        """The workbook tab is Ukrainian АР2, while old settings used Latin AR2."""
        with tempfile.TemporaryDirectory() as folder:
            path=Path(folder)/"graph.xlsx"
            wb=Workbook(); ws=wb.active; ws.title="Розподіл АР2"
            ws.cell(2,3).value="Іваненко Іван Іванович"
            ws.cell(2,18).value="96.113 / ET"
            wb.save(path); wb.close()
            from openpyxl import load_workbook
            loaded=load_workbook(path)
            self.assertEqual(find_sheet(loaded,"Розподіл AR2"),"Розподіл АР2")
            loaded.close()
            patient=get_graph_patient(str(path),"Розподіл AR2","Іваненко Іван Іванович")
            self.assertEqual(patient["sheet"],"Розподіл АР2")

    def test_service_code_uses_the_specialists_final_type(self):
        codes=_codes_by_type("96128/фт 96112/ет")
        self.assertEqual(codes,{"ft":"96128","et":"96112"})
        # The final suffix overrides a stale configured profile type.
        cfg={"specialists":{"Пальчик ФТ":{"type":"ET"}}}
        self.assertEqual(_specialist_type("Пальчик ФТ",cfg),"ft")
        self.assertEqual(_specialist_type("Асистент аФТ",{}),"ft")
        self.assertEqual(_specialist_type("Терапевт ЕТ",{}),"et")
        self.assertEqual(_specialist_type("Асистент аЕТ",{}),"et")
        self.assertEqual(_service_code5("96.128"),"96128")
        self.assertEqual(_service_codes_in_row("96113-00 Послуга"),["96113"])

    def test_takes_matching_referral_from_first_four_rows(self):
        events=[]
        class Action:
            def __init__(self,label):self.label=label
            def inner_text(self,timeout=300):return "Дії"
            def is_visible(self):return True
            def wait_for(self,state="visible",timeout=2000):events.append(("visible",self.label,timeout))
            def click(self,timeout=2000):events.append(("action",self.label,timeout))
            def bounding_box(self):return {"x":1100}
        class Controls:
            def __init__(self,items):self.items=items
            def count(self):return len(self.items)
            def nth(self,index):return self.items[index]
            @property
            def first(self):return self.items[0]
        class Row:
            def __init__(self,label,text):self.label=label;self.text=text
            def is_visible(self):return True
            def inner_text(self,timeout=300):return self.text
            def get_by_text(self,text,exact=True):return Controls([Action(self.label)]) if text=="Дії" else Controls([])
            def locator(self,selector):
                if selector=='td[aria-colindex="4"]':
                    return Controls([type("Service",(),{"inner_text":lambda self,timeout=1000:
                      "96113-00 Навчання навичок, що стосуються положення/рухомості/рухів тіла"})()])
                if selector=='[role="button"][aria-label="Вибрати"]':return Controls([Action(self.label)])
                return Controls([])
        rows=[
            Row("one","96112-00 Послуга"),
            Row("two","96130-00 Послуга"),
            Row("target","96113-00 Послуга 7999-1018-3565-7885"),
            Row("four","96114-00 Послуга"),
            Row("ignored","96113-00 Інше направлення"),
        ]
        class Rows:
            def count(self):return len(rows)
            def nth(self,index):return rows[index]
        class PopupButton:
            def wait_for(self,state="visible",timeout=2000):events.append(("button_wait",state))
            def inner_text(self,timeout=1000):return "ОК"
            def click(self,timeout=2000):events.append(("popup_ok",timeout))
        class PopupText:
            def inner_text(self,timeout=2000):return "Заявка на взяття направлення в роботу успішно створена та обробляється"
        class Popup:
            def wait_for(self,state="visible",timeout=5000):events.append(("popup_wait",state))
            def locator(self,selector):return PopupText() if selector==".popup_text" else PopupButton()
        class Page:
            def locator(self,selector):
                if selector==".popup.-notification":return Popup()
                if selector=="li.side-list-item":
                    class SidebarItem:
                        def is_visible(self):return True
                        def inner_text(self,timeout=500):return "Медичні записи"
                        def locator(self,subselector):
                            return type("IconSet",(),{"first":type("Icon",(),{
                              "wait_for":lambda self,state="visible",timeout=2000:None,
                              "click":lambda self,timeout=3000:events.append(("sidebar_click",timeout))})()})()
                    return type("Sidebar",(),{"count":lambda self:1,"nth":lambda self,index:SidebarItem()})()
                return Rows()
            def wait_for_timeout(self,timeout):events.append(("wait",timeout))
            def wait_for_function(self,script,timeout=20000):
                events.append(("queue_ready",timeout))
        class Web:
            action_id="test"
            page=Page()
            def click_content_text(self,labels,**kwargs):events.append(("content",tuple(labels)));return True
            def goto_medical_records(self):events.append(("medical_records_ready",))
        with patch("app.workflows.procedures.log_event"):
            referral=_take_referral_to_work(Web(),"96113")
        self.assertEqual(referral,{"number":"7999101835657885",
                                   "service":"Навчання навичок, що стосуються положення/рухомості/рухів тіла",
                                   "code":"96113"})
        self.assertIn(("action","target",2000),events)
        self.assertNotIn(("action","ignored",2000),events)
        self.assertIn(("wait",2000),events)
        self.assertEqual(events[0][0],"content")
        popup_events=[event for event in events if event[0] in ("popup_wait","button_wait","popup_ok")]
        self.assertEqual(popup_events,[
            ("popup_wait","visible"),("button_wait","visible"),
            ("popup_ok",2000),("popup_wait","hidden")])
        self.assertIn(("queue_ready",20000),events)
        self.assertIn(("wait",1500),events)
        self.assertIn(("medical_records_ready",),events)
        self.assertLess(events.index(("queue_ready",20000)),events.index(("medical_records_ready",)))
        self.assertLess(events.index(("content",("Взяти у роботу","Взяти в роботу","Взять в работу"))),
                        events.index(("popup_ok",2000)))

    def test_referral_modal_selects_newest_exact_number_and_service(self):
        events=[]
        referral={"number":"7999-1018-3565-7885","service":"Навчання навичок, що стосуються положення/рухомості/рухів тіла"}
        records=[
            ("2026-09-18",referral["service"],referral["number"]),
            ("2026-09-20",referral["service"],"1111-1018-3565-7885"),
            ("2026-09-19",referral["service"],referral["number"]),
            ("2026-09-21","Арт-терапія",referral["number"]),
        ]
        class Element:
            def __init__(self,label):self.label=label
            def wait_for(self,state="visible",timeout=5000):pass
            def click(self,timeout=3000):events.append(("click",self.label))
            def inner_text(self,timeout=500):return self.label
            def is_visible(self):return True
        class Cells:
            def __init__(self,record):self.record=record
            def count(self):return 7
            def nth(self,index):return Element({0:self.record[0],1:self.record[1],3:self.record[2]}.get(index,""))
        class Row:
            def __init__(self,index,record):self.index=index;self.record=record
            def is_visible(self):return True
            def wait_for(self,state="visible",timeout=5000):pass
            def locator(self,selector):
                if selector=='td[role="gridcell"]':return Cells(self.record)
                if selector=='[role="button"][aria-label="Вибрати"]':return type("Result",(),{
                    "count":lambda self:1,"first":Element(f"select-{self.index}")})()
                return type("Result",(),{})()
        class Rows:
            def __init__(self):self.items=[Row(i,record) for i,record in enumerate(records)]
            def count(self):return len(self.items)
            def nth(self,index):return self.items[index]
            @property
            def first(self):return self.items[0]
        class Page:
            def get_by_text(self,text,exact=True):return type("Result",(),{"first":Element(text)})()
            def locator(self,selector):return Rows()
            def wait_for_timeout(self,timeout):events.append(("wait",timeout))
            def wait_for_function(self,script,timeout=20000):return True
        web=type("Web",(),{"page":Page(),"action_id":"test"})()
        with patch("app.workflows.procedures.log_event"):
            _select_matching_referral(web,referral)
        self.assertIn(("click","select-2"),events)
        self.assertNotIn(("click","select-0"),events)
        self.assertNotIn(("click","select-3"),events)
        self.assertEqual([e for e in events if e[0]=="wait"],[])

    def test_create_procedure_clicks_add_next_to_its_label(self):
        events=[]
        class Add:
            def wait_for(self,state="visible",timeout=3000):pass
            def click(self,timeout=3000):events.append("add")
        class Row:
            def locator(self,selector):
                self.selector=selector
                return type("Result",(),{"first":Add()})()
        class Label:
            def wait_for(self,state="visible",timeout=3000):pass
            def locator(self,selector):return Row()
        class Page:
            def get_by_text(self,text,exact=True):
                self.text=text
                return type("Result",(),{"first":Label()})()
            def wait_for_timeout(self,timeout):events.append(timeout)
        web=type("Web",(),{"page":Page()})()
        _click_create_procedure(web)
        self.assertEqual(web.page.text,"Створити процедуру:")
        self.assertEqual(events,["add",2000])

    def test_web_automation_refuses_to_open_a_blank_browser(self):
        with self.assertRaises(ManualCheckpoint) as error:
            _ehealth_url({"ehealth_base_url":""})
        self.assertEqual(error.exception.code,"ehealth_url")
        self.assertEqual(_ehealth_url({"ehealth_base_url":"https://eleks.example.test/login"}),
                         "https://eleks.example.test/login")

    def test_continue_uses_operator_auth_confirmation(self):
        from app.automation.browser import WebEleks
        web=WebEleks("test",{},confirmed_authenticated=True)
        web.page=type("Page",(),{"url":"https://ehealth.eleks.com/"})()
        with patch("app.automation.browser.log_event") as event:
            web.require_authenticated()
            event.assert_called_once()

    def test_ehealth_uses_the_internal_application_frame(self):
        self.assertTrue(_is_ehealth_application_frame("https://ehealth.eleks.com/web/ehealth/internal/"))
        self.assertFalse(_is_ehealth_application_frame("https://ehealth.eleks.com/"))

    def test_patient_result_requires_full_name_after_surname_search(self):
        wanted="Нарожний Ілля Вікторович"
        self.assertTrue(_row_has_full_patient_name("НАРОЖНИЙ  Ілля Вікторович\n1985",wanted))
        self.assertFalse(_row_has_full_patient_name("Нарожний Іван Вікторович\n1985",wanted))

    def test_patient_result_accepts_matching_surname_and_initials(self):
        wanted="Приліпко Василь Олегович"
        self.assertTrue(_row_has_patient_identity("Приліпко В. О.\n02.03.1985",wanted))
        self.assertTrue(_row_has_patient_identity("ПРИЛІПКО Василь Олегович",wanted))
        self.assertFalse(_row_has_patient_identity("Приліпко В. І.\n02.03.1985",wanted))

    def test_patient_birth_date_supports_date_and_age_in_results(self):
        dob=date(1985,3,2)
        self.assertTrue(_row_matches_birth_date("Дата народження: 02.03.1985",dob))
        self.assertTrue(_row_matches_birth_date("1985-03-02",dob))

    def test_patient_search_waits_two_seconds_before_search_click(self):
        events=[]
        web=WebEleks("test",{})
        web.page=type("Page",(),{"wait_for_timeout":lambda self,ms:events.append(("wait",ms))})()
        web.goto_medical_records=lambda:events.append(("goto",None))
        class SearchInput:
            def fill(self,value,timeout=5000):events.append(("fill","patient_search",value))
            def input_value(self,timeout=1000):return "Приліпко"
        web.locator=lambda alias,timeout=2000:SearchInput()
        web._log=lambda *args,**kwargs:None
        web._click_patient_search_button=lambda search:events.append(("click","search_button"))
        row=object()
        web._scan_patient_results=lambda name,timeout_ms=2000:([row],[row])
        web._open_patient_result=lambda result,name,kind,data=None:events.append(("open",kind))
        web.select_patient("Приліпко Василь Олегович",date(1985,3,2))
        self.assertEqual(events[:4],[
            ("goto",None),
            ("fill","patient_search","Приліпко"),
            ("wait",2000),
            ("click","search_button"),
        ])
        self.assertEqual(events[-1],("open","identity"))

    def test_medical_records_does_not_reclick_sidebar_when_workspace_is_ready(self):
        events=[]
        web=WebEleks("test",{})
        web.require_authenticated=lambda:None
        web._medical_records_ready=lambda:True
        web.page=type("Page",(),{"locator":lambda self,selector:events.append(selector)})()
        web.goto_medical_records()
        self.assertEqual(events,[])

    def test_current_patient_search_workspace_counts_as_medical_records_ready(self):
        class Field:
            def is_visible(self):return True
        class Fields:
            count_value=1
            @property
            def first(self):return Field()
            def count(self):return self.count_value
        class ActiveItem:
            def inner_text(self,timeout=300):return "menu_book Медичні записи"
            def count(self):return 1
            def nth(self,index):return self
        class Page:
            url="https://eleks.test/#/?action=spzIBWebPatientGeneralList"
            def locator(self,selector):
                if selector=="li.side-list-item.active":return ActiveItem()
                if selector=='input[name="patientSearch"]':return Fields()
                return Fields()
        web=WebEleks("test",{});web.page=Page()
        self.assertTrue(web._medical_records_ready())

    def test_queue_search_field_is_not_a_medical_records_page(self):
        class Page:
            url="https://eleks.test/#/?action=spzIBWebServiceRequestScheduleList"
            def locator(self,selector):raise AssertionError("queue route should fail first")
        web=WebEleks("test",{});web.page=Page()
        self.assertFalse(web._medical_records_ready())

    def test_start_from_referral_queue_opens_and_verifies_medical_records(self):
        events=[];state={"ready":False}
        class Icon:
            def wait_for(self,state="visible",timeout=2000):pass
            def click(self,timeout=3000):events.append("sidebar_click");state_ready()
        def state_ready():state["ready"]=True
        class Item:
            def is_visible(self):return True
            def inner_text(self,timeout=300):return "Медичні записи"
            def locator(self,selector):return type("One",(),{"first":Icon()})()
        class Items:
            def count(self):return 1
            def nth(self,index):return Item()
        class Page:
            url="https://eleks.test/#/?action=spzIBWebServiceRequestScheduleList"
            def locator(self,selector):return Items()
            def wait_for_timeout(self,timeout):events.append(("wait",timeout))
        web=WebEleks("test",{});web.page=Page()
        web.require_authenticated=lambda:None
        web._medical_records_ready=lambda:state["ready"]
        web._log=lambda *args,**kwargs:None
        web.goto_medical_records()
        self.assertEqual(events,["sidebar_click",("wait",1000)])

    def test_patient_result_click_must_open_matching_card(self):
        values=[];events=[]
        class Field:
            def __init__(self,value):self.value=value
            def is_visible(self):return True
            def input_value(self,timeout=200):return self.value
        class Inputs:
            def count(self):return len(values)
            def nth(self,index):return Field(values[index])
        class Page:
            url="https://eleks.test/patients"
            def locator(self,selector):return Inputs()
            def wait_for_timeout(self,timeout):events.append(("wait",timeout))
        class Result:
            def wait_for(self,state="visible",timeout=2000):pass
            def scroll_into_view_if_needed(self,timeout=1000):events.append(("scroll",timeout))
            def click(self,timeout=2000,force=False):
                events.append(("click",timeout,force))
                values.extend(["Приліпко","Ярина","Димитріївна"])
        web=WebEleks("test",{});web.page=Page();web._log=lambda *args,**kwargs:None
        web._patient_card_matches=lambda name:bool(values)
        web._open_patient_result(Result(),"Приліпко Ярина Димитріївна","identity")
        self.assertEqual(values,["Приліпко","Ярина","Димитріївна"])
        self.assertIn(("click",2000,False),events)

    def test_patient_result_clicks_exact_name_span_coordinates(self):
        events=[];opened=[]
        class Span:
            def is_visible(self):return True
            def inner_text(self,timeout=300):return "Приліпко Ярина Димитріївна"
            def scroll_into_view_if_needed(self,timeout=1000):events.append(("scroll",timeout))
            def bounding_box(self):return {"x":192,"y":195,"width":180,"height":18}
            def evaluate(self,script):return {"tag":"SPAN","classes":"text-1 ng-binding"}
        class Spans:
            def count(self):return 1
            def nth(self,index):return Span()
        class Result:
            def wait_for(self,state="visible",timeout=2000):pass
            def locator(self,selector):
                self.selector=selector
                return Spans()
            def click(self,*args,**kwargs):events.append(("wrong_container_click",))
        class Mouse:
            def click(self,x,y):events.append(("mouse",x,y));opened.append(True)
        class Page:
            url="https://eleks.test/patients"
            mouse=Mouse()
            def wait_for_timeout(self,timeout):events.append(("wait",timeout))
        web=WebEleks("test",{});web.page=Page();web._log=lambda *args,**kwargs:None
        web._patient_card_matches=lambda name:bool(opened)
        web._open_patient_result(Result(),"Приліпко Ярина Димитріївна","identity")
        self.assertIn(("mouse",282.0,204.0),events)
        self.assertNotIn(("wrong_container_click",),events)

    def test_patient_card_initials_count_as_verified_open_card(self):
        values={"patientLastName":"Крупеня","patientFirstName":"Н.",
                "patientSecondName":"С."}
        class Field:
            def __init__(self,value):self.value=value
            def is_visible(self):return True
            def input_value(self,timeout=200):return self.value
        class Fields:
            def __init__(self,value):self.value=value
            def count(self):return 1
            def nth(self,index):return Field(self.value)
        class Page:
            def locator(self,selector):
                name=selector.split('"')[1]
                return Fields(values[name])
        web=WebEleks("test",{});web.page=Page()
        self.assertTrue(web._patient_card_matches("Крупеня Назарій Сергійович"))
        values["patientSecondName"]="І."
        self.assertFalse(web._patient_card_matches("Крупеня Назарій Сергійович"))

    def test_patient_result_does_not_repeat_search_after_card_opens(self):
        events=[];opened=[]
        class Page:
            url="https://eleks.test/patients"
            def wait_for_timeout(self,timeout):events.append(("wait",timeout))
        class Result:
            def wait_for(self,state="visible",timeout=2000):pass
            def locator(self,selector):raise LookupError(selector)
            def click(self,timeout=2000,force=False):
                events.append("result_click");opened.append(True)
        web=WebEleks("test",{});web.page=Page();web._log=lambda *args,**kwargs:None
        web._patient_card_matches=lambda name:bool(opened)
        web._click_patient_search_button=lambda search:events.append("second_search")
        web._open_patient_result(Result(),"Крупеня Назарій Сергійович","identity")
        self.assertEqual(events,["result_click"])

    def test_patient_search_finds_virtual_patient_button(self):
        class PatientButton:
            def is_visible(self):return True
            def inner_text(self,timeout=300):return "Приліпко Ярина Дмитріївна\n2023-01-01"
        class Candidates:
            def __init__(self,items):self.items=items
            def count(self):return len(self.items)
            def nth(self,index):return self.items[index]
        class Page:
            def __init__(self):self.selectors=[]
            def locator(self,selector):
                self.selectors.append(selector)
                return Candidates([PatientButton()]) if "role='option'" in selector else Candidates([])
            def wait_for_timeout(self,timeout):pass
        web=WebEleks("test",{});web.page=Page()
        matches,_=web._scan_patient_results("Приліпко Ярина Дмитріївна",timeout_ms=2000)
        self.assertEqual(len(matches),1)
        self.assertEqual(web.page.selectors[0],"div.list_text")
        self.assertIn("role='option'",web.page.selectors[1])

    def test_browser_checkpoint_restores_saved_plan_without_excel_read(self):
        name="Іваненко Іван Іванович"; pkey=normalize_patient_name(name)+"|пальчик фт"
        plan=[{"date":"2026-09-05","time":"10:00","specialist":"Пальчик ФТ",
               "type":"ft","service_code":"96128","sheet":"ЕТФТ","cell":"C2"}]
        with tempfile.TemporaryDirectory() as folder:
            path=Path(folder)/"graph.xlsx"  # The saved plan survives removal of its source file.
            snapshot=procedure_workflow._resume_snapshot(
                str(path),{"dob":date(2018,1,2)},plan,[],"Пальчик ФТ",pkey)
            state={"state":"browser_ready","data":{"resume":snapshot,"completed_cells":[]}}
            with patch("app.workflows.procedures.get_workflow",return_value=state):
                restored=procedure_workflow._resume_procedure_plan(
                    {"_resume_from_checkpoint":True,"_resume_key":pkey},name)
        self.assertEqual(restored[0],str(path))
        self.assertEqual(restored[1]["dob"],"2018-01-02")
        self.assertEqual(restored[2],plan)


class ProcedureFormSafetyTests(unittest.TestCase):
    def test_calendar_waits_for_date_to_reach_editor_dxform_and_angular(self):
        expected=datetime(2026,9,7,14,0)
        old="2026-09-25T23:58:43.49"
        new="2026-09-07T14:00:00"
        def state(angular,dx,editor):
            return {"formData":{"dateStart":angular},
                    "dxFormData":{"dateStart":dx},
                    "editors":{"dateStart":{"value":editor}}}
        states=[state(old,old,old),state(new,new,old),state(new,new,new)]
        waits=[]
        adapter=DoctorEleksProcedureForm.__new__(DoctorEleksProcedureForm)
        adapter.guard=lambda:None
        adapter.read_state=lambda:states.pop(0)
        adapter.page=type("Page",(),{
            "wait_for_timeout":lambda self,ms:waits.append(ms)})()
        result=adapter._wait_datetime_state("dateStart",expected,timeout_ms=1000)
        self.assertEqual(result["editors"]["dateStart"]["value"],new)
        self.assertEqual(waits,[100,100])

    def test_calendar_still_stops_when_date_never_synchronizes(self):
        expected=datetime(2026,9,7,14,0)
        adapter=DoctorEleksProcedureForm.__new__(DoctorEleksProcedureForm)
        adapter.guard=lambda:None
        adapter.read_state=lambda:{"formData":{"dateStart":"2026-09-25T23:58:43.49"},
                                   "dxFormData":{"dateStart":"2026-09-25T23:58:43.49"},
                                   "editors":{"dateStart":{"value":"2026-09-25T23:58:43.49"}}}
        adapter.page=type("Page",(),{"wait_for_timeout":lambda self,ms:None})()
        with patch("app.workflows.procedure_form.time.monotonic",side_effect=[0,0,0.02]):
            with self.assertRaises(ManualCheckpoint) as error:
                adapter._wait_datetime_state("dateStart",expected,timeout_ms=10)
        self.assertEqual(error.exception.code,"procedure_calendar")

    def test_stock_form_follows_manual_control_order(self):
        events=[]
        adapter=DoctorEleksProcedureForm.__new__(DoctorEleksProcedureForm)
        adapter.saves=[]
        adapter.guard=lambda:None
        adapter.read_state=lambda:{"formData":{"selectedReferralCode":"digitalReferral"}}
        adapter._log=lambda *args,**kwargs:None
        adapter.choose_division=lambda code:events.append("division")
        adapter.choose_calendar_datetime=lambda field,value:events.append(field)
        adapter._wait_datetime_state=lambda field,value,timeout_ms=10000:(
            events.append(("wait",field)) or {})
        adapter._state_datetime_matches=lambda *args:False
        adapter._payload_datetime_matches=lambda *args:False
        def choose_period(verify_state=True):
            events.append(("period",verify_state))
            adapter.saves.append({"payload":{"procedurePeriod":"30"}})
        adapter.choose_period_30=choose_period
        adapter.set_complete_state({"date":"2026-09-03T09:30:00",
            "dateStart":"2026-09-03T09:30:00","dateEnd":"2026-09-03T10:00:00",
            "division":"division-1","isDateOrPeriod":"true"})
        self.assertEqual(events,["division","dateStart",("period",False),
                                 ("wait","dateStart"),"dateEnd",("period",True)])

    def test_period_state_waits_for_angular_dxform_and_editor_to_converge(self):
        adapter=DoctorEleksProcedureForm.__new__(DoctorEleksProcedureForm)
        states=[
            {"formData":{"procedurePeriod":"12366"},"dxFormData":{"procedurePeriod":"30"},
             "editors":{"procedurePeriod":{"value":"30"}}},
            {"formData":{"procedurePeriod":"30"},"dxFormData":{"procedurePeriod":"30"},
             "editors":{"procedurePeriod":{"value":"30"}}},
        ]
        adapter.read_state=lambda:states.pop(0)
        adapter.guard=lambda:None
        adapter.page=type("Page",(),{"wait_for_timeout":lambda self,ms:None})()
        self.assertEqual(adapter._wait_period_state(timeout_ms=1000)["formData"]["procedurePeriod"],"30")

    def test_period_state_match_requires_all_three_layers(self):
        self.assertFalse(DoctorEleksProcedureForm._period_state_matches({
            "formData":{"procedurePeriod":"30"},"dxFormData":{"procedurePeriod":"30"},
            "editors":{"procedurePeriod":{"value":None}}}))
        self.assertTrue(DoctorEleksProcedureForm._period_state_matches({
            "formData":{"procedurePeriod":"30"},"dxFormData":{"procedurePeriod":30},
            "editors":{"procedurePeriod":{"value":"30"}}}))

    def test_division_state_waits_for_angular_dxform_and_editor_to_converge(self):
        adapter=DoctorEleksProcedureForm.__new__(DoctorEleksProcedureForm)
        states=[
            {"formData":{"division":"old"},"dxFormData":{"division":"new"},
             "editors":{"division":{"value":"new"}}},
            {"formData":{"division":"new"},"dxFormData":{"division":"new"},
             "editors":{"division":{"value":"new"}}},
        ]
        adapter.read_state=lambda:states.pop(0)
        adapter.guard=lambda:None
        adapter.page=type("Page",(),{"wait_for_timeout":lambda self,ms:None})()
        self.assertEqual(adapter._wait_field_state("division","NEW",1000)["formData"]["division"],"new")

    def test_login_url_blocks_even_previously_confirmed_session(self):
        web=WebEleks.__new__(WebEleks)
        web.page=type("Page",(),{"url":"https://example.test/#/login"})()
        web.confirmed_authenticated=True
        with self.assertRaises(ManualCheckpoint) as error:web.require_authenticated()
        self.assertEqual(error.exception.code,"login")

    def test_matching_existing_record_is_available_for_retry_skip(self):
        class Row:
            def is_visible(self):return True
            def inner_text(self,timeout=500):
                return "2026-09-03 09:30 2026-09-03 10:00 (30 хв) 96130-00 Пальчик Ярослава"
        class Rows:
            def all(self):return [Row()]
        class Page:
            def locator(self,selector):return Rows()
        web=type("Web",(),{"page":Page()})()
        matches=procedure_workflow._existing_procedure_candidates(
            web,{"date":"2026-09-03","time":"09:30"},"96130","Пальчик ФТ")
        self.assertEqual(len(matches),1)

    def test_existing_record_on_second_grid_page_is_found_and_page_restored(self):
        class Button:
            def __init__(self,page,number):self.page=page;self.number=number
            def get_attribute(self,name):return str(self.number)
            def click(self,timeout=3000):self.page.current=self.number
        class Buttons:
            def __init__(self,page,numbers):self.page=page;self.numbers=numbers
            def all(self):return [Button(self.page,n) for n in self.numbers]
            def count(self):return len(self.numbers)
            @property
            def first(self):return Button(self.page,self.numbers[0])
            def click(self,timeout=3000):self.first.click(timeout=timeout)
        class Pager:
            def __init__(self,page):self.page=page
            def locator(self,selector):
                if "dx-button-success" in selector:return ActiveButton(self.page)
                if 'aria-label="' in selector:
                    return Buttons(self.page,[int(selector.split('aria-label="')[1].split('"')[0])])
                return Buttons(self.page,[1,2,3])
        class ActiveButton:
            def __init__(self,page):self.page=page
            def count(self):return 1
            @property
            def first(self):return Button(self.page,self.page.current)
        class Label:
            def __init__(self,page):self.page=page
            def count(self):return 1
            @property
            def first(self):return self
            def locator(self,selector):return Pager(self.page)
        class Row:
            def __init__(self,page):self.page=page
            def count(self):return 1
            @property
            def first(self):return self
            def inner_text(self,timeout=1000):return f"row {self.page.current}"
        class Page:
            current=1
            def get_by_text(self,text,exact=True):return Label(self)
            def locator(self,selector):return Row(self)
            def wait_for_function(self,script,arg=None,timeout=12000):return True
        page=Page();web=type("Web",(),{"page":page})()
        with patch.object(procedure_workflow,"_matching_visible_procedure_rows",
                          side_effect=lambda *_:["existing"] if page.current==2 else []):
            matches=procedure_workflow._existing_procedure_candidates(
                web,{"date":"2026-09-03","time":"09:30"},"96130","Пальчик ФТ")
        self.assertEqual(matches,["existing"])
        self.assertEqual(page.current,1)

    def test_fresh_procedure_skips_full_grid_scan_but_interrupted_attempt_checks(self):
        item={"cell":"AT272"}
        self.assertFalse(procedure_workflow._needs_retry_duplicate_check(
            {"state":"new","data":{}},item,True))
        self.assertFalse(procedure_workflow._needs_retry_duplicate_check(
            {"state":"in_progress","data":{"attempt_tracking_version":1}},item,True))
        self.assertTrue(procedure_workflow._needs_retry_duplicate_check(
            {"state":"in_progress","data":{"attempt_tracking_version":1,
                                             "attempted_cell":"AT272"}},item,True))
        self.assertTrue(procedure_workflow._needs_retry_duplicate_check(
            {"state":"referral_modal_rows","data":{}},item,True))

    def test_created_row_on_current_page_does_not_turn_pager(self):
        item={"date":"2026-09-03","time":"09:30"}
        page=type("Page",(),{"wait_for_timeout":lambda self,ms:None})()
        web=type("Web",(),{"page":page})()
        row="2026-09-03 09:30 procedure record-1"
        with patch.object(procedure_workflow,"_matching_visible_procedure_rows",return_value=[row]),\
             patch.object(procedure_workflow,"_existing_procedure_candidates") as scan:
            result=procedure_workflow._wait_created_procedure_row(
                web,item,"96130","Specialist",expected_id="record-1")
        self.assertEqual(result,row)
        scan.assert_not_called()

    def test_batch_verification_scans_pages_once_for_all_items(self):
        plan=[{"date":"2026-09-03","time":"09:30","cell":"AK270"},
              {"date":"2026-09-04","time":"10:30","cell":"AL270"}]
        rows=["2026-09-03 09:30 96130-00 Пальчик record-1",
              "2026-09-04 10:30 96130-00 Пальчик record-2"]
        web=type("Web",(),{"action_id":"test"})()
        with patch.object(procedure_workflow,"_scan_procedure_grid_pages",return_value=rows) as scan,\
             patch("app.workflows.procedures.log_event"):
            self.assertEqual(procedure_workflow._verify_batch_procedure_rows(
                web,plan,"96130","Пальчик ФТ"),2)
        scan.assert_called_once()

    def test_batch_verification_detects_missing_or_duplicate_row(self):
        item={"date":"2026-09-03","time":"09:30","cell":"AK270"}
        web=type("Web",(),{"action_id":"test"})()
        with patch.object(procedure_workflow,"_scan_procedure_grid_pages",return_value=[]):
            with self.assertRaises(ManualCheckpoint) as error:
                procedure_workflow._verify_batch_procedure_rows(web,[item],"96130","Пальчик ФТ")
        self.assertEqual(error.exception.code,"procedure_batch_verification")

    def _adapter(self,period="30",editor="30",payload_period="30",status=200):
        start="2026-09-05T15:00:00.000Z";end="2026-09-05T15:30:00.000Z"
        data={"procedurePeriod":period,"dateStart":start,"dateEnd":end,"date":start,
              "serviceCode":"(96128-00) Procedure","performerRef":"person-1",
              "division":"division-1","selectedReferralCode":"digitalReferral",
              "isDateOrPeriod":"true"}
        payload_start="2026-09-05T18:00:00";payload_end="2026-09-05T18:30:00"
        payload={"based_on":"based-1","code":"code-uuid","performer":"person-1",
                 "division":"division-1","selectedReferralRef":"digitalReferral",
                 "procedurePeriod":payload_period,"procedurePeriodStart":payload_start,
                 "procedurePeriodEnd":payload_end,"is_date_or_period":"true",
                 "performed_date_time":payload_start,"showErrorMessage":"0"}
        return type("Adapter",(),{
            "failure":None,"target":{"date":"2026-09-05","start_time":"18:00","code":"96128"},
            "saves":[{"sequence":1,"status":status,"payload":payload,"response":{"ok":True}}],
            "guard":lambda self:None,"read_state":lambda self:{"ready":True,"formData":data,
              "dxFormData":dict(data),"editors":{"procedurePeriod":{"value":editor},
              "division":{"value":"division-1"},"dateStart":{"value":start,"display":"05.09.2026, 18:00"},
              "dateEnd":{"value":end,"display":"05.09.2026, 18:30"}}},
            "_log":lambda *args,**kwargs:None})()

    def test_null_procedure_period_blocks_final_preflight(self):
        adapter=self._adapter(period=None,editor=None,payload_period=None)
        with self.assertRaises(ManualCheckpoint) as error:
            procedure_workflow._procedure_autosave_preflight(adapter)
        self.assertEqual(error.exception.code,"procedure_preflight")

    def test_start_at_1800_has_end_at_1830(self):
        target=procedure_workflow._build_procedure_target("Patient",{},"96128","Specialist",
          {"date":"2026-09-05","time":"18:00"})
        self.assertEqual(target["end_time"],"18:30")
        self.assertEqual(target["duration_minutes"],30)

    def test_form_snapshot_reacquires_current_dom_on_each_read(self):
        class Page:
            def __init__(self):self.reads=0
            def evaluate(self,script):
                self.reads+=1
                return {"ready":True,"formData":{},"editors":{"generation":self.reads}}
        adapter=DoctorEleksProcedureForm.__new__(DoctorEleksProcedureForm)
        adapter.page=Page()
        first=adapter.read_state();second=adapter.read_state()
        self.assertEqual(first["editors"]["generation"],1)
        self.assertEqual(second["editors"]["generation"],2)

    def test_autosave_500_raises_before_any_following_selector_action(self):
        interactions=[]
        adapter=DoctorEleksProcedureForm.__new__(DoctorEleksProcedureForm)
        adapter.failure=(500,{"payload":{"procedurePeriod":None},"response":{"error":"bad"}})
        adapter.last_action="FIELD_BATCH_BEGIN";adapter.target={};adapter.web=type("Web",(),{
          "action_id":"test","snapshot_error":lambda self,label:""})()
        adapter._log=lambda *args,**kwargs:None
        adapter._try_read_state=lambda:{"formData":{}}
        with self.assertRaises(ManualCheckpoint) as error:
            adapter.guard();interactions.append("next dropdown")
        self.assertEqual(error.exception.code,"procedure_autosave_failed")
        self.assertEqual(interactions,[])

    def test_http_401_is_auth_checkpoint(self):
        adapter=DoctorEleksProcedureForm.__new__(DoctorEleksProcedureForm)
        adapter.failure=(401,{"payload":{},"response":{}});adapter.last_action="OPEN"
        adapter.target={};adapter._log=lambda *args,**kwargs:None
        with self.assertRaises(ManualCheckpoint) as error:adapter.guard()
        self.assertEqual(error.exception.code,"login")
        self.assertIn("HTTP 401",str(error.exception))

    def test_non_save_portal_401_is_auth_failure(self):
        adapter=DoctorEleksProcedureForm.__new__(DoctorEleksProcedureForm)
        adapter.saves=[];adapter.failure=None;adapter.last_action="FORM_READY"
        adapter.target={};adapter._log=lambda *args,**kwargs:None
        response=type("Response",(),{"request":object(),"status":401,
            "url":"https://example.test/api/run/portal"})()
        adapter._capture_response(response)
        self.assertEqual(adapter.failure[0],401)
        with self.assertRaises(ManualCheckpoint) as error:adapter.guard()
        self.assertEqual(error.exception.code,"login")

    def test_final_save_waits_for_validate_request(self):
        adapter=DoctorEleksProcedureForm.__new__(DoctorEleksProcedureForm)
        adapter.failure=None;adapter.guard=lambda:None
        adapter.saves=[{"sequence":2,"status":200,"payload":{"toValidate":"1","id":"record-1"},
                        "requested_at":time.monotonic()-1,"response":{"keys":[]}}]
        adapter.page=type("Page",(),{"wait_for_timeout":lambda self,ms:None})()
        self.assertEqual(adapter.wait_final_save(1,timeout_ms=100,
                                                 expected_id="RECORD-1")["status"],200)

    def test_initialization_e1035_is_logged_but_not_called_corruption(self):
        logs=[]
        adapter=DoctorEleksProcedureForm.__new__(DoctorEleksProcedureForm)
        adapter._mutations_started=False;adapter.failure=None;adapter.last_action="FORM_OPEN"
        adapter._log=lambda event,data:logs.append((event,data))
        adapter._try_read_state=lambda:{"ready":True}
        message=type("ConsoleMessage",(),{"text":"E1035 - Cannot read properties of undefined (reading 'getInstance')"})()
        adapter._capture_console(message)
        self.assertIsNone(adapter.failure)
        self.assertEqual(logs[0][0],"FORM_INITIALIZATION_CONSOLE")

    def test_e1035_after_successful_save_requires_live_form_check(self):
        adapter=DoctorEleksProcedureForm.__new__(DoctorEleksProcedureForm)
        adapter._mutations_started=True;adapter.failure=None;adapter.last_action="SELECT_DROPDOWN:division"
        adapter._active_record_id="record-1";adapter._console_pending=[]
        adapter._log=lambda *args,**kwargs:None
        adapter._try_read_state=lambda:{"ready":True,
            "formData":{"procedureRef":"record-1","isDateOrPeriod":"true"},
            "editors":{key:{"dataSourceReady":True} for key in
                       ("division","isDateOrPeriod","dateStart","dateEnd","procedurePeriod")}}
        adapter.web=type("Web",(),{"require_authenticated":lambda self:None})()
        message=type("ConsoleMessage",(),{"text":"E1035 - editor failed"})()
        adapter._capture_console(message)
        self.assertIsNone(adapter.failure)
        self.assertTrue(adapter._console_pending)
        adapter.guard()
        self.assertIsNone(adapter.failure)
        self.assertEqual(adapter._console_pending,[])

    def test_e1035_with_missing_required_editor_stops_form(self):
        adapter=DoctorEleksProcedureForm.__new__(DoctorEleksProcedureForm)
        adapter._mutations_started=True;adapter.failure=None;adapter.last_action="SELECT_DIVISION"
        adapter._active_record_id="record-1";adapter._console_pending=[]
        adapter._log=lambda *args,**kwargs:None
        adapter._try_read_state=lambda:{"ready":True,
            "formData":{"procedureRef":"record-1","isDateOrPeriod":"true"},
            "editors":{"division":{"dataSourceReady":True}}}
        adapter.web=type("Web",(),{"require_authenticated":lambda self:None})()
        adapter.page=type("Page",(),{"wait_for_timeout":lambda self,ms:None})()
        adapter._capture_console(type("Message",(),{"text":"E1035 - editor failed"})())
        adapter._confirm_console_state(timeout_ms=20)
        self.assertEqual(adapter.failure[0],"CORRUPTED")

    def test_console_warning_after_confirmed_final_save_allows_closed_form(self):
        events=[]
        adapter=DoctorEleksProcedureForm.__new__(DoctorEleksProcedureForm)
        adapter.failure=None;adapter._final_save_confirmed=True
        adapter._console_pending=["E1035 after patient-list navigation"]
        adapter._log=lambda event,data:events.append(event)
        adapter._confirm_console_state=lambda: self.fail("closed form was probed")
        adapter.guard()
        self.assertEqual(adapter._console_pending,[])
        self.assertEqual(events,["FORM_CLOSED_AFTER_FINAL"])

    def test_2xx_with_complete_payload_passes_preflight(self):
        procedure_workflow._procedure_autosave_preflight(self._adapter())

    def test_final_save_500_stops_before_confirmation(self):
        item={"date":"2026-09-03","time":"09:30","sheet":"ETFT","cell":"AK270"}
        adapter=type("Adapter",(),{"guard":lambda self:None,"close":lambda self:None})()
        with patch("app.workflows.procedures._sign_procedure_later",
                   side_effect=ManualCheckpoint("procedure_autosave_failed","HTTP 500")),\
             patch("app.workflows.procedures._wait_created_procedure_row") as row:
            with self.assertRaises(ManualCheckpoint):
                procedure_workflow._confirm_procedure(
                    object(),adapter,item,"96130","Specialist")
        row.assert_not_called()

    def test_successful_final_http_2xx_is_recognized(self):
        item={"date":"2026-09-03","time":"09:30","sheet":"ETFT","cell":"AK270"}
        events=[]
        adapter=type("Adapter",(),{"guard":lambda self:events.append("guard"),
                                    "close":lambda self:events.append("close"),
                                    "_final_save_confirmed":True})()
        page=type("Page",(),{"wait_for_function":lambda self,*args,**kwargs:
                              events.append("patient_list")})()
        web=type("Web",(),{"page":page,"action_id":"test"})()
        with patch("app.workflows.procedures._sign_procedure_later",
                   side_effect=lambda web:events.append("final_2xx") or "record-1"),\
             patch("app.workflows.procedures._matching_visible_procedure_rows",
                   side_effect=lambda *args:events.append("current_page") or []),\
             patch("app.workflows.procedures._existing_procedure_candidates") as scan,\
             patch("app.workflows.procedures.log_event"):
            result=procedure_workflow._confirm_procedure(
                web,adapter,item,"96130","Specialist")
        self.assertEqual(result,"record-1")
        self.assertEqual(events,["final_2xx","guard","patient_list","current_page","close"])
        scan.assert_not_called()

    def test_final_2xx_without_closed_form_stops_confirmation(self):
        item={"date":"2026-09-03","time":"09:30","sheet":"ETFT","cell":"AK270"}
        adapter=type("Adapter",(),{"guard":lambda self:None,
                                    "_final_save_confirmed":True})()
        page=type("Page",(),{"wait_for_function":lambda self,*args,**kwargs:
                              (_ for _ in ()).throw(TimeoutError())})()
        web=type("Web",(),{"page":page})()
        with patch("app.workflows.procedures._sign_procedure_later",return_value="record-1"):
            with self.assertRaises(ManualCheckpoint) as error:
                procedure_workflow._confirm_procedure(
                    web,adapter,item,"96130","Specialist")
        self.assertEqual(error.exception.code,"procedure_final_navigation")

    def test_visual_thirty_minutes_with_null_form_data_fails(self):
        adapter=self._adapter(period=None,editor="30 хвилин",payload_period="30")
        with self.assertRaises(ManualCheckpoint):
            procedure_workflow._procedure_autosave_preflight(adapter)

    def test_form_data_and_editor_do_not_hide_null_outgoing_payload(self):
        adapter=self._adapter(period="30",editor="30",payload_period=None)
        with self.assertRaises(ManualCheckpoint):
            procedure_workflow._procedure_autosave_preflight(adapter)

    def test_structured_manual_diff_ignores_generated_procedure_id(self):
        diff=structured_payload_diff({"payload":{"id":"old","code":"A"}},
                                     {"payload":{"id":"new","code":"B"}})
        self.assertEqual(diff,{"code":{"failed":"A","manual_success":"B"}})

    def test_portal_param_array_preserves_literal_null_for_preflight(self):
        body={"name":"spiprocedureSave","params":[
            {"name":"procedurePeriod","value":"null"},
            {"name":"is_date_or_period","value":"true"},
            {"name":"Authorization","value":"Bearer secret"}]}
        self.assertEqual(sanitize_payload(body),
                         {"procedurePeriod":"null","is_date_or_period":"true"})

    def test_literal_null_payload_blocks_final_preflight(self):
        with self.assertRaises(ManualCheckpoint):
            procedure_workflow._procedure_autosave_preflight(
                self._adapter(payload_period="null"))


if __name__ == "__main__":
    unittest.main()
