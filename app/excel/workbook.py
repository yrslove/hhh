from __future__ import annotations
from pathlib import Path
from datetime import datetime, date
import shutil, re
from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.utils.datetime import from_excel
from ..core.text import normalize_patient_name, display_name
from ..core.paths import BACKUP_DIR
from ..core.audit import log_event
from ..core.action_control import raise_if_cancelled

def backup(path:str)->str:
    src=Path(path)
    stamp=datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    dst=BACKUP_DIR/f"{stamp}_{src.name}"
    shutil.copy2(src,dst)
    return str(dst)

_SHEET_CONFUSABLES=str.maketrans({
    "a":"а", "b":"в", "c":"с", "e":"е", "h":"н", "i":"і",
    "k":"к", "m":"м", "o":"о", "p":"р", "r":"р", "t":"т", "x":"х", "y":"у",
})

def normalize_sheet_name(value)->str:
    """Match tabs typed with visually identical Latin/Cyrillic letters."""
    return re.sub(r"[\s/\\_-]+","",str(value or "").casefold()).translate(_SHEET_CONFUSABLES)

def find_sheet(wb,wanted:str):
    norm=normalize_sheet_name
    for s in wb.sheetnames:
        if norm(s)==norm(wanted):return s
    return None

def load_procedure_reference(path:str,schedule_sheet:str,distribution_sheet:str):
    """Load only the two sheets used to build one patient's procedure plan.

    This workbook is read-only by convention: its in-memory worksheets are
    never saved. The standard openpyxl loader eagerly parses every sheet.
    """
    from types import SimpleNamespace
    from openpyxl.reader.excel import ExcelReader
    from openpyxl.worksheet._reader import WorksheetReader
    from openpyxl.packaging.relationship import get_rels_path, get_dependents, RelationshipList
    from .schedule import _sheet_name_match

    class SelectedReader(ExcelReader):
        def read_worksheets(self):
            sheets=list(self.parser.find_sheets())
            names=SimpleNamespace(sheetnames=[sheet.name for sheet,_ in sheets])
            selected={_sheet_name_match(names,schedule_sheet),find_sheet(names,distribution_sheet)}
            for sheet,rel in sheets:
                if sheet.name not in selected:
                    continue
                ws=self.wb.create_sheet(sheet.name)
                rels_path=get_rels_path(rel.target)
                ws._rels=(get_dependents(self.archive,rels_path) if rels_path in self.valid_files
                          else RelationshipList())
                with self.archive.open(rel.target) as stream:
                    WorksheetReader(ws,stream,self.shared_strings,self.data_only,self.rich_text).bind_all()
                ws.sheet_state=sheet.state
            # Sheet indexes from the source workbook no longer apply to this
            # subset; defined names are not needed for schedule or patient data.
            self.parser.assign_names=lambda:None

    reader=None
    try:
        reader=SelectedReader(path,data_only=False,rich_text=True)
        reader.read()
        return reader.wb
    except Exception as exc:
        if reader is not None:
            reader.archive.close()
        log_event("excel.selective_reader_fallback","Вибіркове читання Excel недоступне; читаємо книгу звичайним способом.",
                  level="WARNING",action_id="procedures_create",data={"reason":str(exc)})
        return load_workbook(path,data_only=False,rich_text=True)

def find_patient_row(path:str,sheet:str,name:str,name_col=3,workbook=None):
    """Find patient rows; reuse ``workbook`` to avoid a duplicate XLSX load."""
    opened_here=workbook is None
    wb=workbook or load_workbook(path,data_only=False,rich_text=True)
    sn=find_sheet(wb,sheet) or sheet
    if sn not in wb.sheetnames:
        if opened_here:wb.close()
        return None
    ws=wb[sn]; key=normalize_patient_name(name)
    matches=[]
    try:
        for r in range(1,ws.max_row+1):
            if r%100==0:raise_if_cancelled()
            v=ws.cell(r,name_col).value
            if normalize_patient_name(v)==key:matches.append(r)
    finally:
        if opened_here:wb.close()
    return matches

def _schedule_month(value, epoch=None):
    if isinstance(value,datetime):return value.strftime("%Y-%m")
    if isinstance(value,date):return value.strftime("%Y-%m")
    if isinstance(value,(int,float)) and not isinstance(value,bool) and epoch is not None:
        try:return from_excel(value,epoch).strftime("%Y-%m")
        except Exception:return ""
    text=str(value or "").strip()
    for fmt in ("%Y-%m-%d %H:%M:%S","%Y-%m-%d %H:%M","%Y-%m-%d",
                "%d.%m.%Y %H:%M","%d.%m.%Y","%d/%m/%Y","%d-%m-%Y"):
        try:return datetime.strptime(text,fmt).strftime("%Y-%m")
        except ValueError:pass
    return ""


def get_graph_patient(path:str,sheet:str,name:str,workbook=None,month=None):
    """Read one patient, optionally from a workbook already open for this task."""
    opened_here=workbook is None
    wb=workbook or load_workbook(path,data_only=False,rich_text=True)
    try:
        sn=find_sheet(wb,sheet) or sheet
        if sn not in wb.sheetnames:raise ValueError(f"Не знайдено вкладку {sheet}")
        ws=wb[sn]
        matches=find_patient_row(path,sn,name,3,workbook=wb)
        if not matches:raise ValueError("Пацієнта не знайдено у Графік 1.")
        wanted_month=str(month or "").strip()
        if wanted_month and not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])",wanted_month):
            raise ValueError("Місяць пошуку має бути у форматі РРРР-ММ.")
        if wanted_month:
            cached_wb=None
            try:
                if any(ws.cell(row,13).data_type=="f" for row in matches):
                    cached_wb=load_workbook(path,data_only=True,read_only=True)
                cached_ws=cached_wb[sn] if cached_wb else None
                matches=[row for row in matches if _schedule_month(
                    cached_ws.cell(row,13).value if ws.cell(row,13).data_type=="f" and cached_ws else ws.cell(row,13).value,
                    cached_wb.epoch if cached_wb else wb.epoch)==wanted_month]
            finally:
                if cached_wb:cached_wb.close()
            if not matches:
                raise ValueError(f"Знайдено рядки з таким ПІБ, але немає рядка за місяць {wanted_month}.")
        if len(matches)>1:
            raise ValueError(f"Знайдено {len(matches)} рядків з однаковим ПІБ"+
                             (f" за місяць {wanted_month}" if wanted_month else "")+
                             "; потрібне ручне уточнення.")
        r=matches[0]
        interaction_cell=ws.cell(r,13)
        diagnosis_cell=ws.cell(r,14)
        first_interaction=interaction_cell.value
        system_diagnoses=diagnosis_cell.value
        formula_columns=[column for column,cell in ((13,interaction_cell),(14,diagnosis_cell))
                         if cell.data_type=="f"]
        if formula_columns:
            # openpyxl does not calculate formulas. Read saved formula results
            # for both the datetime and ordered diagnosis source cells.
            cached_wb=load_workbook(path,data_only=True,read_only=True)
            try:
                cached_ws=cached_wb[sn]
                if 13 in formula_columns:
                    first_interaction=cached_ws.cell(r,13).value
                    if first_interaction is None or str(first_interaction).strip()=="" or (
                        isinstance(first_interaction,str) and first_interaction.startswith("#")):
                        raise ValueError(
                            f"У {sn}!M{r} немає збереженого результату формули. "
                            "Перерахуйте й збережіть Excel-книгу перед запуском."
                        )
                    if isinstance(first_interaction,(int,float)) and not isinstance(first_interaction,bool):
                        first_interaction=from_excel(first_interaction,cached_wb.epoch)
                if 14 in formula_columns:
                    system_diagnoses=cached_ws.cell(r,14).value
                    if system_diagnoses is None or not str(system_diagnoses).strip() or (
                        isinstance(system_diagnoses,str) and system_diagnoses.startswith("#")):
                        raise ValueError(
                            f"У {sn}!N{r} немає збереженого результату формули. "
                            "Перерахуйте й збережіть Excel-книгу перед запуском."
                        )
            finally:
                cached_wb.close()
        return {
          "row":r,"sheet":sn,"name":display_name(ws.cell(r,3).value),
          "referral":ws.cell(r,5).value,"dob":ws.cell(r,6).value,
          "document_diagnoses":ws.cell(r,9).value,
          "first_interaction":first_interaction,
          "system_diagnoses":system_diagnoses,
          "first_interaction_status":ws.cell(r,15).value,
          "plan_status":ws.cell(r,16).value,
          "plan_detail_status":ws.cell(r,17).value,
          "codes_specialists":ws.cell(r,18).value
        }
    finally:
        if opened_here:wb.close()

def parse_codes(value)->list[str]:
    if value is None:return []
    txt=str(value)
    return list(dict.fromkeys(re.findall(r"\b[A-ZА-ЯІЇЄҐ]?\d{1,3}(?:[.\-]\d{1,3})*(?:\.\d+)?\b",txt,re.I)))

def parse_diagnoses_with_primary(path:str,sheet:str,name:str):
    wb=load_workbook(path,data_only=False,rich_text=True)
    try:
        sn=find_sheet(wb,sheet) or sheet; ws=wb[sn]
        rows=find_patient_row(path,sn,name,3,workbook=wb)
        if len(rows)!=1:raise ValueError("Не вдалося однозначно знайти пацієнта.")
        cell=ws.cell(rows[0],14)
        value=cell.value
        codes=[]; primary=[]
        if isinstance(value,CellRichText):
            for block in value:
                text=str(block)
                found=parse_codes(text); codes.extend(found)
                bold=False
                try:
                    if isinstance(block,TextBlock):
                        bold=bool(getattr(block.font,"b",False))
                except: pass
                if bold:primary.extend(found)
        else:
            codes=parse_codes(value)
            if bool(cell.font and cell.font.bold) and len(codes)==1:primary=list(codes)
        codes=list(dict.fromkeys(codes)); primary=list(dict.fromkeys(primary))
        return {"codes":codes,"primary":primary,"ambiguous_primary":len(primary)!=1 and len(codes)>0,"row":rows[0]}
    finally:
        wb.close()

def write_cell(path:str,sheet:str,row:int,col:int,value,bold:bool|None=None):
    b=backup(path)
    wb=load_workbook(path)
    ws=wb[find_sheet(wb,sheet) or sheet]
    cell=ws.cell(row,col); cell.value=value
    if bold is not None:
        from copy import copy
        f=copy(cell.font); f.bold=bold; cell.font=f
    wb.save(path);wb.close()
    log_event("excel.write","Змінено клітинку Excel",data={"path":path,"sheet":sheet,"cell":cell.coordinate,"backup":b})
    return cell.coordinate

class WorkbookChangedError(RuntimeError):
    """The source workbook changed during the initial read."""


def sync_coordinator_to_graph(coordinator_path:str,graph_path:str,patient_name:str,coord_sheet:str,graph_sheet:str):
    cwb=load_workbook(coordinator_path,data_only=False)
    cs=cwb[coord_sheet] if coord_sheet and coord_sheet in cwb.sheetnames else cwb[cwb.sheetnames[0]]
    key=normalize_patient_name(patient_name)
    # Current workflow says coordinator PІБ is in C; B is a conservative legacy fallback only if C has no exact match.
    rows_c=[r for r in range(1,cs.max_row+1) if normalize_patient_name(cs.cell(r,3).value)==key]
    rows_b=[r for r in range(1,cs.max_row+1) if normalize_patient_name(cs.cell(r,2).value)==key] if not rows_c else []
    found=rows_c or rows_b
    if len(found)!=1:raise ValueError(f"У таблиці координаторів знайдено {len(found)} точних рядків за ПІБ.")
    r=found[0]
    headers={}
    header_row=None
    for rr in range(1,min(20,cs.max_row)+1):
        local={}
        for cc in range(1,cs.max_column+1):
            v=str(cs.cell(rr,cc).value or "").casefold()
            if "направ" in v:local.setdefault("referral",cc)
            if "народ" in v:local.setdefault("dob",cc)
            if "діагноз" in v or "диагноз" in v:local.setdefault("diag",cc)
        if len(local)>len(headers):
            headers=local;header_row=rr
    missing=[x for x in ("referral","dob","diag") if x not in headers]
    if missing:
        raise ValueError("Не знайдено заголовки в таблиці координаторів: "+", ".join(missing)+". Програма не буде вгадувати колонки.")
    referral=cs.cell(r,headers["referral"]).value
    dob=cs.cell(r,headers["dob"]).value
    diag=cs.cell(r,headers["diag"]).value
    cwb.close()
    gwb=load_workbook(graph_path)
    gs=gwb[find_sheet(gwb,graph_sheet) or graph_sheet]
    matches=[rr for rr in range(1,gs.max_row+1) if normalize_patient_name(gs.cell(rr,3).value)==key]
    backup_path=backup(graph_path)
    if matches:
        if len(matches)>1:raise ValueError("У Графік 1 є кілька однакових ПІБ; автоматичний запис зупинено.")
        gr=matches[0]
    else:
        gr=gs.max_row+1
        gs.cell(gr,3).value=patient_name
    gs.cell(gr,5).value=referral
    gs.cell(gr,6).value=dob
    gs.cell(gr,9).value=diag
    gwb.save(graph_path);gwb.close()
    return {"row":gr,"backup":backup_path,"referral":str(referral or ""),"dob":str(dob or ""),"diagnoses":str(diag or "")}


def _coordinator_columns(ws):
    best=None
    for r in range(1,min(20,ws.max_row)+1):
        name_col=None;status_col=None
        for c in range(1,ws.max_column+1):
            v=str(ws.cell(r,c).value or "").casefold()
            if name_col is None and ("піб" in v or "пацієн" in v or "пациент" in v):name_col=c
            if status_col is None and ("статус план" in v or "план лікування" in v or "план лечения" in v):status_col=c
        score=int(bool(name_col))+int(bool(status_col))
        if best is None or score>best[0]:best=(score,r,name_col,status_col)
    if not best or not best[2] or not best[3]:
        raise ValueError("Не знайдено заголовки ПІБ та статусу плану лікування у таблиці координаторів.")
    return best[1],best[2],best[3]

def update_coordinator_status(path:str,patient_name:str,status:str,sheet:str=""):
    wb=load_workbook(path)
    ws=wb[sheet] if sheet and sheet in wb.sheetnames else wb[wb.sheetnames[0]]
    key=normalize_patient_name(patient_name)
    header_row,name_col,status_col=_coordinator_columns(ws)
    rows=[r for r in range(header_row+1,ws.max_row+1) if normalize_patient_name(ws.cell(r,name_col).value)==key]
    if len(rows)!=1:raise ValueError(f"Не вдалося однозначно знайти пацієнта в таблиці координаторів ({len(rows)}).")
    b=backup(path);ws.cell(rows[0],status_col).value=status;wb.save(path);wb.close()
    return {"row":rows[0],"column":status_col,"backup":b}

def get_coordinator_status(path:str,patient_name:str,sheet:str=""):
    wb=load_workbook(path,data_only=True)
    ws=wb[sheet] if sheet and sheet in wb.sheetnames else wb[wb.sheetnames[0]]
    key=normalize_patient_name(patient_name)
    header_row,name_col,status_col=_coordinator_columns(ws)
    rows=[r for r in range(header_row+1,ws.max_row+1) if normalize_patient_name(ws.cell(r,name_col).value)==key]
    if len(rows)!=1:
        wb.close();return None
    value=ws.cell(rows[0],status_col).value;wb.close()
    return value
