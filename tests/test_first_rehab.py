from __future__ import annotations

import unittest
from datetime import date
from unittest.mock import MagicMock, patch

from app.workflows import rehab_first
from app.workflows import first_interaction
from app.automation.browser import WebEleks, AutomationError


class RehabFirstPureTests(unittest.TestCase):
    def test_diagnosis_parser_ignores_annotations_and_preserves_order(self):
        self.assertEqual(rehab_first.parse_icd_codes("R62.8 + R47.8 (СП) + Z50.9"),
                         ["R62.8", "R47.8", "Z50.9"])

    def test_primary_is_first_code_in_user_source_order(self):
        self.assertEqual(rehab_first.parse_icd_codes(
            "G96.8 R62.8 + R26.8 + R41.8 + R47.8 + Z50.9 + Z54.8"),
            ["G96.8","R62.8","R26.8","R41.8","R47.8","Z50.9","Z54.8"])

    def test_exact_code_matching_does_not_match_prefix_or_extension(self):
        wanted=rehab_first.normalize_icd("R62.8")
        self.assertEqual(wanted,"R62.8")
        self.assertNotEqual(wanted,rehab_first.normalize_icd("R62"))
        self.assertNotEqual(wanted,rehab_first.normalize_icd("R62.80"))

    def test_all_diagnoses_present_keeps_rows_and_z(self):
        plan=rehab_first.plan_diagnosis_reconciliation(
            ["R62.8","R47.8"],[{"code":"R62.8"},{"code":"R47.8"},{"code":"Z50.9"}])
        self.assertEqual([r["code"] for r in plan["keep"]],["R62.8","R47.8"])
        self.assertEqual(plan["replace"],[]);self.assertEqual(plan["delete"],[]);self.assertEqual(plan["add"],[])
        self.assertEqual([r["code"] for r in plan["immutable_z"]],["Z50.9"])

    def test_source_smaller_deletes_mutable_extras_and_keeps_z(self):
        plan=rehab_first.plan_diagnosis_reconciliation(
            ["R62.8"],[{"code":"R62.8"},{"code":"R47.8"},{"code":"Z50.9"}])
        self.assertEqual([r["code"] for r in plan["delete"]],["R47.8"])
        self.assertEqual([r["code"] for r in plan["immutable_z"]],["Z50.9"])

    def test_source_larger_adds_missing_codes(self):
        plan=rehab_first.plan_diagnosis_reconciliation(
            ["R62.8","R47.8","G96.8"],[{"code":"R62.8"}])
        self.assertEqual(plan["add"],["R47.8","G96.8"])

    def test_duplicate_codes_are_preserved_and_never_replaced_or_deleted(self):
        plan=rehab_first.plan_diagnosis_reconciliation(
            ["A22.1","R62.8"],
            [{"code":"R52.2"},{"code":"R52.2"},{"code":"Z50.9"}])
        self.assertEqual(plan["replace"],[])
        self.assertEqual(plan["delete"],[])
        self.assertEqual(plan["add"],["A22.1","R62.8"])
        self.assertEqual([r["code"] for r in plan["protected_duplicates"]],["R52.2","R52.2"])
        self.assertEqual(plan["expected_mutable_counts"],{"A22.1":1,"R62.8":1,"R52.2":2})

    def test_unique_wrong_code_is_replaced_and_single_surplus_is_deleted(self):
        plan=rehab_first.plan_diagnosis_reconciliation(
            ["A22.1"],[{"code":"R52.2"},{"code":"R47.8"},{"code":"Z50.9"}])
        self.assertEqual([(old["code"],new) for old,new in plan["replace"]],[("R52.2","A22.1")])
        self.assertEqual([r["code"] for r in plan["delete"]],["R47.8"])
        self.assertEqual(plan["add"],[])

    def test_wrong_diagnosis_time_is_detected(self):
        self.assertFalse(rehab_first.diagnosis_time_matches("26.09.2026 09:00","26.09.2026 10:00"))
        self.assertTrue(rehab_first.diagnosis_time_matches("26.09.2026 10:00","26.09.2026 10:00"))

    def test_96022_wrong_time_is_detected_for_edit(self):
        actual="26.09.2026 09:00";expected="26.09.2026 10:00"
        self.assertFalse(rehab_first.diagnosis_time_matches(actual,expected))

    def test_report_code_is_extracted_from_scrollable_row_without_its_date(self):
        self.assertEqual(first_interaction._fi_report_code(
            "2026-09-15 10:00:00 96022-00 / Оцінка підтримання здоров'я"),"9602200")
        self.assertEqual(first_interaction._fi_report_code("96.0.22.00 / report"),"9602200")
        self.assertEqual(first_interaction._fi_report_code(
            "2026-09-15 10:00:00 96037-00 / Інше обстеження"),"9603700")
        self.assertEqual(first_interaction._fi_report_code("2026-09-15 10:00:00"),"")

    def test_success_notification_uses_scoped_primary_ok_button(self):
        web=MagicMock()
        popup=MagicMock()
        popup.inner_text.return_value="Діагностичний звіт видалено!"
        button=MagicMock()
        button.inner_text.return_value="ОК"
        button.is_visible.return_value=True
        button.is_enabled.return_value=True
        popup.locator.return_value.count.return_value=1
        popup.locator.return_value.nth.return_value=button
        web.page.locator.return_value.count.return_value=0
        first_interaction._fi_confirm_popup_ok(web,popup,"success notification")
        popup.locator.assert_any_call('.popup_buttons button.btn.btn-primary')
        button.click.assert_called_once_with(timeout=10000)
        popup.wait_for.assert_not_called()
        web.page.wait_for_timeout.assert_any_call(2000)

    def test_barthel_lookup_uses_the_data_cell_and_its_own_row(self):
        web=MagicMock()
        container=MagicMock();cells=MagicMock();cell=MagicMock();row_locator=MagicMock()
        row=MagicMock()
        web.page.locator.return_value.count.return_value=1
        web.page.locator.return_value.nth.return_value=container
        container.locator.return_value=cells
        cells.count.return_value=1;cells.nth.return_value=cell
        cell.inner_text.return_value="Barthel_Extended - Підсумкова оцінка Індекс Бартел (розширений)"
        cell.locator.return_value=row_locator
        row_locator.count.return_value=1
        row_locator.locator.return_value.count.return_value=0
        row_locator.inner_text.return_value=cell.inner_text.return_value+" 60 бал 2026-09-15 10:00"
        found=first_interaction._fi_barthel_exam_rows(web)
        web.page.locator.assert_called_with(".dx-scrollable-content:visible")
        container.locator.assert_called_with(
            'td[aria-describedby]:visible:has-text("Barthel_Extended")')
        self.assertEqual(found,[(row_locator,row_locator.inner_text.return_value)])

    def test_diagnostic_report_reader_scopes_to_dx_scrollable_content(self):
        web=MagicMock()
        container=MagicMock();rows=MagicMock();row=MagicMock()
        web.page.locator.return_value=MagicMock(count=MagicMock(return_value=1))
        web.page.locator.return_value.nth.return_value=container
        container.locator.return_value=rows
        rows.count.return_value=1;rows.nth.return_value=row
        row.locator.return_value.count.return_value=0
        row.inner_text.return_value="2026-09-15 10:00:00 96022-00 / report"
        result=first_interaction._fi_diagnostic_report_rows(web)
        web.page.locator.assert_called_with(".dx-scrollable-content:visible")
        self.assertEqual(result,[(row,"2026-09-15 10:00:00 96022-00 / report")])

    def test_child_removes_96037_and_barthel(self):
        self.assertEqual(rehab_first.validate_age_specific_reports(17,["9602200","9603700"]),[1])
        names=["Barthel_Extended","Other"]
        self.assertEqual(rehab_first.barthel_rows_to_remove(17,names),[0])

    def test_adult_requires_96037_and_leaves_barthel(self):
        self.assertEqual(rehab_first.validate_age_specific_reports(18,["9602200","9603700"]),[])
        self.assertEqual(rehab_first.barthel_rows_to_remove(18,["Barthel_Extended"]),[])
        with self.assertRaisesRegex(ValueError,"96037"):
            rehab_first.validate_age_specific_reports(18,["9602200"])

    def test_96022_is_mandatory(self):
        with self.assertRaisesRegex(ValueError,"96022"):
            rehab_first.validate_age_specific_reports(10,["9603700"])

    def test_latest_referral_wins_and_ties_keep_first_ui_order(self):
        self.assertEqual(rehab_first.newest_referral_index(["24.09.2026","26.09.2026","25.09.2026"]),1)
        self.assertEqual(rehab_first.newest_referral_index(["26.09.2026","26.09.2026"]),0)

    def test_template_drift_reports_expected_and_actual(self):
        self.assertEqual(rehab_first.validate_template_drift({"duration":"60","priority":"Плановий"},
                                                            {"duration":"30","priority":"Плановий"}),
                         [("duration","60","30")])


class StagedWaitTests(unittest.TestCase):
    def _web(self, ready_after):
        state={"now":0.0}
        class Locator:
            @property
            def first(self):return self
            def count(self):return 1
            def is_visible(self):return state["now"]>=ready_after
            def is_enabled(self):return True
        class Page:
            def wait_for_timeout(self,ms):state["now"]+=ms/1000
        web=object.__new__(WebEleks);web.page=Page();web.selector_map={"target":["target"]}
        web._candidate_locator=lambda candidate:Locator()
        web.snapshot_error=lambda label:""
        return web,state

    def test_immediate_and_after_first_stage(self):
        for ready_at in (0,1.5):
            web,state=self._web(ready_at)
            with patch("app.automation.browser.time.monotonic",side_effect=lambda:state["now"]):
                web.wait_for_alias("target")
            self.assertLessEqual(state["now"],2.0)

    def test_after_second_and_third_stages(self):
        for ready_at in (4.0,9.0):
            web,state=self._web(ready_at)
            with patch("app.automation.browser.time.monotonic",side_effect=lambda:state["now"]):
                web.wait_for_alias("target")
            self.assertGreaterEqual(state["now"],ready_at)

    def test_final_timeout_aborts_at_twelve_seconds(self):
        web,state=self._web(99)
        with patch("app.automation.browser.time.monotonic",side_effect=lambda:state["now"]):
            with self.assertRaisesRegex(AutomationError,r"2\+5\+5"):
                web.wait_for_alias("target")
        self.assertGreaterEqual(state["now"],11.9)
        self.assertLessEqual(state["now"],12.1)


class FirstInteractionUiGuardTests(unittest.TestCase):
    def test_diagnosis_pager_can_be_sibling_of_custom_grid(self):
        class Button:
            def __init__(self,n):self.n=n
            def get_attribute(self,name,timeout=300):return f"Page {self.n}" if name=="aria-label" else ""
            def inner_text(self,timeout=300):return str(self.n)
        class ButtonList:
            def __init__(self,items):self.items=items
            def count(self):return len(self.items)
            def nth(self,index):return self.items[index]
            @property
            def first(self):return self.items[0]
        all_pages=ButtonList([Button(1),Button(2)])
        selected=ButtonList([Button(1)])
        class Pager:
            def locator(self,selector):return selected if "dx-selection" in selector else all_pages
        class PagerList:
            def count(self):return 1
            @property
            def first(self):return Pager()
        class Empty:
            def count(self):return 0
        class Parent:
            def locator(self,selector):return PagerList() if selector==".dx-pager:visible" else Empty()
        class Grid:
            def locator(self,selector):
                if selector==".dx-pager:visible":return Empty()
                if selector=="xpath=..":return Parent()
                return Empty()
        class GridList:
            def count(self):return 1
            @property
            def first(self):return Grid()
        class Page:
            def locator(self,selector):
                if selector.startswith("extendedgrid[id^"):return GridList()
                raise AssertionError(selector)
        web=MagicMock();web.page=Page()
        _grid,_pager,pages,current=first_interaction._fi_diagnosis_page_state(web)
        self.assertEqual(pages,[1,2])
        self.assertEqual(current,1)

    def test_row_action_waits_for_a_stable_position(self):
        web=MagicMock()
        row=MagicMock()
        controls=row.locator.return_value
        controls.count.return_value=1
        control=controls.nth.return_value
        control.is_visible.return_value=True
        control.is_enabled.return_value=True
        control.get_attribute.return_value="Редагувати"
        box={"x":100.0,"y":200.0,"width":50.0,"height":20.0}
        control.bounding_box.side_effect=[box,box]

        result=first_interaction._fi_row_action(web,row,["Редагувати"])

        self.assertIs(result,control)
        self.assertEqual(control.bounding_box.call_count,2)

    def test_diagnosis_control_is_scoped_to_open_modal(self):
        web=MagicMock()
        modal=web.page.locator.return_value
        modal.count.return_value=1
        field=modal.first.locator.return_value
        field.count.return_value=1
        field.first.is_visible.return_value=True
        field.first.is_enabled.return_value=True

        result=first_interaction._fi_diagnosis_control(web,"diagnosis_role")

        self.assertIs(result,field.first)
        web.page.locator.assert_called_once_with(".popup.-extended:visible")
        modal.first.locator.assert_called_once_with(
            'input.dx-texteditor-input[role="combobox"][id$="_diagnosesrole"]')

    def test_referral_input_read_tolerates_transient_timeout_for_next_poll(self):
        web=MagicMock()
        locator=web.page.locator.return_value
        locator.count.return_value=1
        field=locator.nth.return_value
        field.is_visible.return_value=True
        field.input_value.side_effect=[TimeoutError("not rendered yet"),"9715-5008-3769"]

        self.assertEqual(first_interaction._fi_read_named_input(web,"serviceRequestNumber"),"")
        self.assertEqual(first_interaction._fi_read_named_input(web,"serviceRequestNumber"),"9715-5008-3769")
        web.page.locator.assert_called_with('input[name="serviceRequestNumber"]')
        self.assertEqual(field.input_value.call_args_list[0].kwargs["timeout"],2000)

    def test_selected_referral_checks_input_values_and_reference(self):
        state={"number":"9715-5008-6324-3769","created":"2026-09-26",
               "has_reference":True,"radio_checked":True}
        self.assertTrue(first_interaction._fi_referral_matches(
            state,"9715-5008-6324-3769","26.09.2026"))
        self.assertFalse(first_interaction._fi_referral_matches(
            state,"9715-5008-6324-3768","26.09.2026"))
        self.assertFalse(first_interaction._fi_referral_matches(
            state,"9715-5008-6324-3769","25.09.2026"))
        self.assertFalse(first_interaction._fi_referral_matches(
            {**state,"has_reference":False},"9715-5008-6324-3769","26.09.2026"))

    def test_facility_display_quotes_and_repeated_number_are_equivalent(self):
        expected="ТОВ 2 Міжнародний центр реабілітації та нейрофізіології"
        shown='ТОВ 2 \\"МІЖНАРОДНИЙ ЦЕНТР РЕАБІЛІТАЦІЇ ТА НЕЙРОФІЗІОЛОГІЇ\\" 2'
        self.assertTrue(first_interaction._fi_place_matches(expected,shown))
        self.assertTrue(first_interaction._fi_place_matches(expected,expected))
        self.assertFalse(first_interaction._fi_place_matches(
            expected,'ТОВ 2 "ІНШИЙ ЦЕНТР РЕАБІЛІТАЦІЇ" 2'))
        self.assertFalse(first_interaction._fi_place_matches(
            expected,'ТОВ 2 "МІЖНАРОДНИЙ ЦЕНТР РЕАБІЛІТАЦІЇ ТА НЕЙРОФІЗІОЛОГІЇ" 3'))

    def test_saved_interaction_guard_requires_start_and_full_hour_end(self):
        target="15.09.2026 10:00"
        self.assertFalse(first_interaction._fi_saved_interaction_matches(
            "2026-09-15 10:00 2026-09-15 10:30 G96.8",target))
        self.assertTrue(first_interaction._fi_saved_interaction_matches(
            "2026-09-15 10:00 2026-09-15 11:00 Реабілітація",target))
        self.assertFalse(first_interaction._fi_saved_interaction_matches(
            "2026-09-15 09:00 2026-09-15 10:00 Реабілітація",target))

    def test_patient_card_accepts_initials_verified_by_existing_helper(self):
        web=MagicMock()
        web._patient_card_matches.return_value=True
        field=web.page.locator.return_value.nth.return_value
        web.page.locator.return_value.count.return_value=1
        field.is_visible.return_value=True
        field.input_value.return_value="02.01.2018"
        with patch.object(first_interaction,"_fi_wait_condition",side_effect=lambda _web,predicate,_desc:self.assertTrue(predicate())):
            first_interaction._fi_verify_patient_card(
                web,"Бадінгар Данієль Ерікович",date(2018,1,2))
        web._patient_card_matches.assert_called_once_with("Бадінгар Данієль Ерікович")
        web.page.locator.assert_called_once_with('input[name="patientBirthDate"]')
        self.assertEqual(field.input_value.call_args.kwargs["timeout"],0)
        web.current_text.assert_not_called()

    def test_patient_card_rejects_different_birth_date(self):
        web=MagicMock()
        web._patient_card_matches.return_value=True
        field=web.page.locator.return_value.nth.return_value
        web.page.locator.return_value.count.return_value=1
        field.is_visible.return_value=True
        field.input_value.return_value="03.01.2018"
        with patch.object(first_interaction,"_fi_wait_condition",side_effect=lambda _web,predicate,_desc:predicate()):
            with self.assertRaisesRegex(AutomationError,"не збігається"):
                first_interaction._fi_verify_patient_card(
                    web,"Бадінгар Данієль Ерікович",date(2018,1,2))

    def test_patient_card_allows_missing_birth_date_field(self):
        web=MagicMock()
        web._patient_card_matches.return_value=True
        web.page.locator.return_value.count.return_value=0
        with patch.object(first_interaction,"log_event") as logged:
            first_interaction._fi_verify_patient_card(
                web,"Бадінгар Данієль Ерікович",date(2018,1,2))
        logged.assert_called_once()
        web._patient_card_matches.assert_called_once_with("Бадінгар Данієль Ерікович")

    def test_patient_card_does_not_wait_if_birth_date_is_unreadable(self):
        web=MagicMock()
        web._patient_card_matches.return_value=True
        web.page.locator.return_value.count.return_value=1
        field=web.page.locator.return_value.nth.return_value
        field.is_visible.return_value=True
        field.input_value.side_effect=TimeoutError("field rerendered")
        with patch.object(first_interaction,"log_event") as logged:
            first_interaction._fi_verify_patient_card(
                web,"Бадінгар Данієль Ерікович",date(2018,1,2))
        self.assertEqual(field.input_value.call_args.kwargs["timeout"],0)
        logged.assert_called_once()

    def test_click_uses_the_locator_that_passed_readiness_check(self):
        web=MagicMock()
        ready=web.wait_for_alias.return_value
        first_interaction._fi_click(web,"save")
        web.wait_for_alias.assert_called_once_with("save",enabled=True)
        web.click.assert_called_once_with("save",locator=ready)

    def test_fill_uses_the_locator_that_passed_readiness_check(self):
        web=MagicMock()
        ready=web.wait_for_alias.return_value
        ready.input_value.return_value="R62.8"
        first_interaction._fi_fill(web,"code_search","R62.8")
        web.fill.assert_called_once_with("code_search","R62.8",locator=ready)

    def test_interaction_start_uses_datetime_calendar_handler(self):
        web=MagicMock()
        ready=web.wait_for_alias.return_value
        with patch.object(first_interaction,"_fi_set_datetime") as set_datetime:
            first_interaction._fi_fill(web,"interaction_start","15.09.2026 10:00")
        set_datetime.assert_called_once_with(web,ready,"interaction_start","15.09.2026 10:00")
        web.fill.assert_not_called()

    def test_section_rows_exclude_visible_modal_rows(self):
        web=MagicMock()
        rows=web.page.locator.return_value
        rows.count.return_value=2
        main=rows.nth.side_effect=[MagicMock(),MagicMock()]
        main[0].locator.return_value.count.return_value=0
        main[0].inner_text.return_value="96022"
        main[1].locator.return_value.count.return_value=1
        result=first_interaction._fi_table_rows(web,"diagnostic_reports")
        web.click_content_text.assert_called_once_with(
            ("Діагностичні звіти","Диагностические отчеты"),retry_click=True)
        web.page.locator.assert_called_once_with("tr.dx-data-row:visible")
        self.assertEqual(result,[(main[0],"96022")])


if __name__=="__main__":
    unittest.main()
