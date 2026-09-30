import tempfile, unittest
from pathlib import Path
from datetime import date,time
from unittest.mock import patch
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from app.excel.schedule import (Appointment,ScheduleIndex,StructureWarning,build_index,
                                parse_schedule,check_frm_before_etft,check_etft_overlaps,
                                classify_frm_occurrences,normalize_patient_name)
from app.workflows import schedule_checks

BLUE=PatternFill("solid",fgColor="92CDDC")
GREEN=PatternFill("solid",fgColor="C6EFCE")

def make_book(path):
    wb=Workbook(); frm=wb.active;frm.title="ФРМ"; et=wb.create_sheet("ЕТФТ")
    for ws in (frm,et):
        ws["C1"]=date(2026,9,4);ws["D1"]=date(2026,9,5)
    # FRM one-hour patient 12:00-13:00 on 04.09
    times=["08:00","08:30","09:00","09:30","10:00","10:30","11:00","11:30","12:00","12:30","13:00","13:30","14:00"]
    for i,t in enumerate(times,2):
        frm.cell(i,1).value="ФРМ Лікар";frm.cell(i,2).value=t
    r=2+times.index("12:00");frm.cell(r,3).value="ІВАНЕНКО Іван Іванович";frm.cell(r,3).fill=BLUE
    frm.cell(r+1,3).fill=BLUE
    # ET/FT section 1
    for i,t in enumerate(times,2):
        et.cell(i,1).value="Халай Тимофій ET";et.cell(i,2).value=t
    # forbidden earlier 10:30-11:30 (hour by continuation)
    r1=2+times.index("10:30");et.cell(r1,3).value="Іваненко Іван Іванович";et.cell(r1,3).fill=GREEN;et.cell(r1+1,3).fill=GREEN
    # allowed 13:00
    r2=2+times.index("13:00");et.cell(r2,3).value="Іваненко Іван Іванович"
    # section 2 repeated time grid causing overlap on 05.09
    start=20
    for i,t in enumerate(times,start):
        et.cell(i,1).value="Інший спеціаліст FT";et.cell(i,2).value=t
    a=start+times.index("11:00");et.cell(a,4).value="Петренко Петро";et.cell(a,4).fill=BLUE;et.cell(a+1,4).fill=BLUE
    b=2+times.index("11:30");et.cell(b,4).value="Петренко Петро";et.cell(b,4).fill=GREEN;et.cell(b+1,4).fill=GREEN
    wb.save(path)

class ScheduleTests(unittest.TestCase):
    def test_rules(self):
        with tempfile.TemporaryDirectory() as d:
            p=Path(d)/"x.xlsx";make_book(p)
            frm,fw=parse_schedule(str(p),"ФРМ",2026);et,ew=parse_schedule(str(p),"ЕТФТ",2026)
            order,hourly=check_frm_before_etft(frm,et)
            self.assertEqual(hourly,1)
            self.assertEqual(len(order),1)
            self.assertIn("10:30",order[0]["etft"]["interval"])
            overlaps=check_etft_overlaps(et)
            p_over=[x for x in overlaps if "Петренко" in x["patient"]]
            # Both visual blocks are 60 minutes in Excel, but ET/FT overlap
            # semantics are fixed at 30 minutes: 11:00 and 11:30 are adjacent.
            self.assertTrue(all(a.duration_minutes==60 for a in et if "Петренко" in a.patient_name_original))
            self.assertEqual(p_over,[])

    @staticmethod
    def appointment(patient,day,start="12:00",end="13:00",*,hourly=True,cell="C3",sheet="ФРМ"):
        return Appointment(
            patient_name_original=patient,
            patient_name_normalized=normalize_patient_name(patient),
            sheet_name=sheet,date=f"2026-09-{day}",start_time=start,end_time=end,
            duration_minutes=60 if hourly else 30,specialist="test",cell_address=cell,
            fill_color="92CDDC",continuation_cell_address=f"{cell}_next" if hourly else None,
        )

    def test_frm_pair_boundaries_and_unusual_gaps(self):
        for later in ("19","21"):
            with self.subTest(gap=later):
                frm=[self.appointment("Іваненко Іван Іванович","05",cell="C3"),
                     self.appointment("Іваненко Іван Іванович",later,cell="D3")]
                _,primary,repeat,unpaired,ar3_infos=classify_frm_occurrences(frm)
                self.assertEqual(len(primary),1)
                self.assertEqual(len(repeat),1)
                self.assertEqual(len(unpaired),0)
                self.assertEqual(ar3_infos,[])
                self.assertEqual(repeat[0].frm_occurrence_type,"REPEAT")
        for later in ("18","22"):
            with self.subTest(unusual_gap=later):
                frm=[self.appointment("Іваненко Іван Іванович","05",cell="C3"),
                     self.appointment("Іваненко Іван Іванович",later,cell="D3")]
                _,primary,repeat,unpaired,ar3_infos=classify_frm_occurrences(frm)
                self.assertEqual((len(primary),len(repeat),len(unpaired)),(0,0,2))
                self.assertEqual(len(ar3_infos),1)
                self.assertEqual(ar3_infos[0].type,"54 пакет група АР3-1 (монопослуга)")

    def test_four_frm_visits_form_two_pairs(self):
        frm=[self.appointment("Іваненко Іван Іванович",day,cell=f"C{i}")
             for i,day in enumerate(("05","20","30"),1)]
        # October is deliberately used for the second normal 15-day cycle.
        frm[-1].date="2026-10-05"
        frm.append(self.appointment("Іваненко Іван Іванович","20",cell="C4"))
        frm[-1].date="2026-10-20"
        _,primary,repeat,unpaired,ar3_infos=classify_frm_occurrences(frm)
        self.assertEqual([a.date for a in primary],["2026-09-05","2026-10-05"])
        self.assertEqual([a.date for a in repeat],["2026-09-20","2026-10-20"])
        self.assertEqual(unpaired,[])
        self.assertEqual(ar3_infos,[])

    def test_names_ignore_trailing_service_note_but_keep_original(self):
        self.assertEqual(
            normalize_patient_name("Дронік Ігор Євгенійович (Повторне 3 2026)"),
            normalize_patient_name("ДРОНІК   Ігор Євгенійович"),
        )
        self.assertEqual(
            normalize_patient_name("Тимошенко Людмила Володимирівна (прибиральниця)"),
            normalize_patient_name("Тимошенко Людмила Володимирівна"),
        )
        original="Тимошенко Людмила Володимирівна (прибиральниця)"
        a=self.appointment(original,"05")
        self.assertEqual(a.patient_name_original,original)

    def test_repeat_is_ignored_for_order_but_not_etft_overlap(self):
        patient="Дронік Ігор Євгенійович (Повторне 3 2026)"
        frm=[self.appointment(patient,"05",cell="C3"),self.appointment(patient,"20",cell="D3")]
        # One pre-FRM ET/FT record on the primary date; exact 13:00 boundary is allowed.
        et=[self.appointment("ДРОНІК   Ігор Євгенійович","05","10:00","11:00",hourly=False,cell="E3",sheet="ЕТФТ"),
            self.appointment("ДРОНІК   Ігор Євгенійович","05","13:00","14:00",hourly=False,cell="E4",sheet="ЕТФТ"),
            # These repeat-date ET/FT visits are still included in overlap detection.
            self.appointment("ДРОНІК   Ігор Євгенійович","20","09:00","10:00",hourly=False,cell="F3",sheet="ЕТФТ"),
            self.appointment("ДРОНІК   Ігор Євгенійович","20","10:00","11:00",hourly=False,cell="F4",sheet="ЕТФТ"),
            self.appointment("ДРОНІК   Ігор Євгенійович","20","10:30","11:30",hourly=False,cell="F5",sheet="ЕТФТ"),
            self.appointment("ДРОНІК   Ігор Євгенійович","20","10:00","11:00",hourly=False,cell="F6",sheet="ЕТФТ")]
        classify_frm_occurrences(frm)
        order,hourly=check_frm_before_etft(frm,et)
        self.assertEqual(hourly,2)
        self.assertEqual(len(order),1)
        self.assertEqual(order[0]["date"],"2026-09-05")
        overlaps=check_etft_overlaps(et)
        self.assertEqual(len(overlaps),1)
        self.assertEqual(overlaps[0]["overlap"],"10:00–10:30")

    def test_etft_overlap_uses_fixed_half_hour_slots(self):
        patient="Іваненко Іван Іванович"
        def et(start,cell):
            # The deliberately one-hour visual interval must be ignored here.
            return self.appointment(patient,"12",start,"13:00",hourly=True,cell=cell,sheet="ЕТФТ")
        for first,second in (("12:00","12:30"),("12:30","13:00")):
            with self.subTest(first=first,second=second):
                self.assertEqual(check_etft_overlaps([et(first,"AS95"),et(second,"AS305")]),[])
        first,second=et("12:00","AS95"),et("12:00","AS305")
        first.specialist="Specialist A"; second.specialist="Specialist B"
        errors=check_etft_overlaps([first,second])
        self.assertEqual(len(errors),1)
        self.assertEqual(errors[0]["first"]["interval"],"12:00–12:30")
        self.assertEqual(errors[0]["second"]["interval"],"12:00–12:30")
        self.assertEqual(errors[0]["overlap"],"12:00–12:30")
        self.assertEqual({errors[0]["first"]["cell"],errors[0]["second"]["cell"]},{"AS95","AS305"})
        self.assertEqual({errors[0]["first"]["specialist"],errors[0]["second"]["specialist"]},{"Specialist A","Specialist B"})

    def test_levchenko_adjacent_slots_are_not_overlap(self):
        patient="Левченко Олексій Ігорович"
        et=[self.appointment(patient,"18","12:00","13:00",hourly=True,cell="AZ40",sheet="ЕТФТ"),
            self.appointment(patient,"18","12:30","13:00",hourly=False,cell="AZ198",sheet="ЕТФТ")]
        self.assertEqual(check_etft_overlaps(et),[])

    def test_frm_order_still_uses_etft_start_time(self):
        patient="Іваненко Іван Іванович"
        frm=[self.appointment(patient,"12","12:00","13:00",cell="C3")]
        before=[self.appointment(patient,"12","12:30","13:00",hourly=False,cell="D3",sheet="ЕТФТ")]
        boundary=[self.appointment(patient,"12","13:00","13:30",hourly=False,cell="D4",sheet="ЕТФТ")]
        self.assertEqual(len(check_frm_before_etft(frm,before)[0]),1)
        self.assertEqual(check_frm_before_etft(frm,boundary)[0],[])

    def test_important_output_hides_technical_entries_and_exposes_ar3_info(self):
        patient="Сторожук Дмитро Павлович"
        first=self.appointment(patient,"02",cell="E9"); first.date="2026-08-02"
        next_=self.appointment(patient,"05",cell="AN9"); next_.date="2026-09-05"
        hourly,primary,repeat,unpaired,ar3_infos=classify_frm_occurrences([first,next_])
        schedule=ScheduleIndex(
            frm_appointments=[first,next_],etft_appointments=[],etft_by_patient_date=build_index([]),
            frm_hourly=hourly,frm_primary=primary,frm_repeat=repeat,frm_unpaired=unpaired,
            structure_warnings=[StructureWarning("ФРМ","V44","Знайдено порожнє ПІБ; клітинку пропущено.")],
            ar3_monoservice=ar3_infos,
        )
        with patch("app.workflows.schedule_checks.graph_path",return_value="test.xlsx"), \
             patch("app.workflows.schedule_checks.settings",return_value={}), \
             patch("app.workflows.schedule_checks.build_schedule_index",return_value=schedule), \
             patch("app.workflows.schedule_checks.log_event"), \
             patch("app.workflows.schedule_checks.save_check_run",return_value=1):
            result=schedule_checks.run_all({})
        data=result.data
        self.assertEqual(data["summary"]["errors"],0)
        self.assertEqual(data["summary"]["ar3_monoservice"],1)
        self.assertNotIn("warnings",data)
        self.assertNotIn("frm_unpaired",data)
        self.assertEqual(data["ar3_monoservice"][0]["first_cell"],"E9")
        self.assertEqual(data["ar3_monoservice"][0]["next_cell"],"AN9")

    def test_schedule_human_log_groups_errors_by_month(self):
        errors=[
            {"date":"2026-09-01","message":"Помилка вересня"},
            {"date":"2026-08-15","message":"Помилка серпня 2"},
            {"date":"2026-08-01","message":"Помилка серпня 1"},
        ]
        output=schedule_checks._schedule_human_log("ЕТ_ФТ_ФРМ_Пересечение времени",errors)
        self.assertEqual(output.count("08 МЕСЯЦ"),1)
        self.assertEqual(output.count("09 МЕСЯЦ"),1)
        self.assertLess(output.index("08 МЕСЯЦ"),output.index("09 МЕСЯЦ"))
        self.assertLess(output.index("Помилка серпня 1"),output.index("Помилка серпня 2"))

if __name__=="__main__":unittest.main()
