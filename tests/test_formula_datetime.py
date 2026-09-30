from __future__ import annotations

import re
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from zipfile import ZipFile

from openpyxl import Workbook

from app.excel.workbook import get_graph_patient
from app.workflows.rehab_first import format_interaction_datetime, parse_interaction_datetime


class FormulaDatetimeTests(unittest.TestCase):
    def _workbook(self, root: Path, *, cached: bool, diagnosis_formula=False) -> Path:
        path=root / "graph.xlsx"
        wb=Workbook();ws=wb.active;ws.title="Patients"
        ws["C2"]="Sample Patient Name"
        ws["M2"]="=A1+1"
        ws["N2"]='="G96.8 R62.8 + R26.8"' if diagnosis_formula else "R62.8"
        wb.save(path);wb.close()
        if cached:
            replacement=root / "cached.xlsx"
            with ZipFile(path) as source, ZipFile(replacement,"w") as target:
                for entry in source.infolist():
                    data=source.read(entry.filename)
                    if entry.filename=="xl/worksheets/sheet1.xml":
                        data,count=re.subn(rb'(<c r="M2"[^>]*><f>.*?</f>)<v\s*/>',
                                           rb'\g<1><v>46280.4166666667</v>',data,count=1)
                        self.assertEqual(count,1)
                        if diagnosis_formula:
                            data,count=re.subn(rb'<c r="N2"[^>]*><f>(.*?)</f><v\s*/></c>',
                                rb'<c r="N2" t="str"><f>\1</f><v>G96.8 R62.8 + R26.8</v></c>',data,count=1)
                            self.assertEqual(count,1)
                    target.writestr(entry,data)
            path.write_bytes(replacement.read_bytes())
        return path

    def test_formula_uses_saved_numeric_result_without_writing_excel(self):
        with tempfile.TemporaryDirectory() as folder:
            path=self._workbook(Path(folder),cached=True)
            before=path.read_bytes()
            record=get_graph_patient(str(path),"Patients","Sample Patient Name")
            self.assertIsInstance(record["first_interaction"],datetime)
            self.assertEqual(format_interaction_datetime(record["first_interaction"]),"15.09.2026 10:00")
            self.assertEqual(path.read_bytes(),before)

    def test_numeric_excel_serial_is_accepted(self):
        parsed=parse_interaction_datetime(46280.4166666667)
        self.assertEqual(parsed.strftime("%d.%m.%Y %H:%M"),"15.09.2026 10:00")

    def test_formula_without_saved_result_fails_closed(self):
        with tempfile.TemporaryDirectory() as folder:
            path=self._workbook(Path(folder),cached=False)
            with self.assertRaisesRegex(ValueError,"M2.*результату формули"):
                get_graph_patient(str(path),"Patients","Sample Patient Name")

    def test_diagnosis_formula_uses_saved_ordered_text_result(self):
        with tempfile.TemporaryDirectory() as folder:
            path=self._workbook(Path(folder),cached=True,diagnosis_formula=True)
            record=get_graph_patient(str(path),"Patients","Sample Patient Name")
            self.assertEqual(record["system_diagnoses"],"G96.8 R62.8 + R26.8")


if __name__=="__main__":
    unittest.main()
