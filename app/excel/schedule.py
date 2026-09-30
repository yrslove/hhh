from __future__ import annotations
from dataclasses import dataclass, asdict
from datetime import date, datetime, time, timedelta
from collections import defaultdict
from typing import Any
import re, math, unicodedata, xml.etree.ElementTree as ET
from openpyxl import load_workbook
from openpyxl.styles.colors import COLOR_INDEX
from openpyxl.cell.rich_text import CellRichText, TextBlock
from ..core.text import normalize_patient_name, display_name
from ..core.audit import log_event
from ..core.action_control import raise_if_cancelled, set_progress

UA_MONTHS={
 "січня":1,"січень":1,"лютого":2,"лютий":2,"березня":3,"березень":3,"квітня":4,"квітень":4,
 "травня":5,"травень":5,"червня":6,"червень":6,"липня":7,"липень":7,"серпня":8,"серпень":8,
 "вересня":9,"вересень":9,"жовтня":10,"жовтень":10,"листопада":11,"листопад":11,
 "грудня":12,"грудень":12
}
TIME_RX=re.compile(r"^\s*(\d{1,2})[:.](\d{2})(?::\d{2}(?:[.,]\d+)?)?\s*$")
DATE_RX=re.compile(r"^\s*(\d{1,2})[./-](\d{1,2})(?:[./-](\d{2,4}))?\s*$")

@dataclass(slots=True)
class Appointment:
    patient_name_original:str
    patient_name_normalized:str
    sheet_name:str
    date:str
    start_time:str
    end_time:str
    duration_minutes:int
    specialist:str
    cell_address:str
    fill_color:str
    continuation_cell_address:str|None=None
    row:int=0
    column:int=0
    # Assigned only for hourly ФРМ appointments.  ET/FT appointments are never
    # filtered by this classification.
    frm_occurrence_type:str|None=None
    paired_with_cell:str|None=None
    paired_with_date:str|None=None
    pair_gap_days:int|None=None
    def as_dict(self): return asdict(self)

@dataclass(slots=True)
class StructureWarning:
    sheet:str
    cell:str
    message:str
    def as_dict(self): return asdict(self)

@dataclass(slots=True)
class Ar3MonoserviceInfo:
    patient:str
    first_date:str
    next_date:str
    gap_days:int
    first_cell:str
    next_cell:str
    type:str="54 пакет група АР3-1 (монопослуга)"
    def as_dict(self): return asdict(self)

@dataclass(slots=True)
class ScheduleIndex:
    """One workbook import and the independent indexes used by IMPORTANT_000."""
    frm_appointments:list[Appointment]
    etft_appointments:list[Appointment]
    etft_by_patient_date:dict
    frm_hourly:list[Appointment]
    frm_primary:list[Appointment]
    frm_repeat:list[Appointment]
    frm_unpaired:list[Appointment]  # internal pairing state; never user-facing
    structure_warnings:list[StructureWarning]
    ar3_monoservice:list[Ar3MonoserviceInfo]

def parse_time(v)->time|None:
    if v is None:return None
    if isinstance(v,datetime): return v.time().replace(second=0,microsecond=0)
    if isinstance(v,time): return v.replace(second=0,microsecond=0)
    if isinstance(v,(float,int)) and 0<=float(v)<1:
        mins=round(float(v)*24*60)
        return time((mins//60)%24, mins%60)
    m=TIME_RX.match(str(v))
    if not m:return None
    h,mi=int(m.group(1)),int(m.group(2))
    if h>23 or mi>59:return None
    return time(h,mi)

def parse_date(v, year:int)->date|None:
    if v is None:return None
    if isinstance(v,datetime): return v.date()
    if isinstance(v,date): return v
    s=str(v).strip().casefold()
    m=DATE_RX.match(s)
    if m:
        d,mo=int(m.group(1)),int(m.group(2)); y=int(m.group(3)) if m.group(3) else year
        if y<100:y+=2000
        try:return date(y,mo,d)
        except:return None
    # "4 вересня" / "4 вересня 2026"
    m=re.match(r"^\s*(\d{1,2})\s+([а-яіїєґ]+)(?:\s+(\d{4}))?\s*$",s)
    if m and m.group(2) in UA_MONTHS:
        try:return date(int(m.group(3) or year),UA_MONTHS[m.group(2)],int(m.group(1)))
        except:return None
    return None

def _theme_colors(wb):
    out={}
    raw=getattr(wb,"loaded_theme",None)
    if not raw:return out
    try:
        root=ET.fromstring(raw)
        ns={"a":"http://schemas.openxmlformats.org/drawingml/2006/main"}
        scheme=root.find(".//a:clrScheme",ns)
        if scheme is None:return out
        names=["lt1","dk1","lt2","dk2","accent1","accent2","accent3","accent4","accent5","accent6","hlink","folHlink"]
        for idx,name in enumerate(names):
            el=scheme.find(f"a:{name}",ns)
            if el is None:continue
            child=next(iter(el),None)
            if child is None:continue
            val=child.attrib.get("val") or child.attrib.get("lastClr")
            if val: out[idx]=val.upper()
    except Exception:
        pass
    return out

def _apply_tint(rgb:str,tint:float|None)->str:
    # sufficient for comparing and white detection; preserves deterministic actual-color approximation
    if not rgb:return ""
    rgb=rgb[-6:].upper()
    if tint in (None,0,0.0):return rgb
    try:
        vals=[int(rgb[i:i+2],16) for i in (0,2,4)]
        if tint<0: vals=[round(v*(1+tint)) for v in vals]
        else: vals=[round(v*(1-tint)+255*tint) for v in vals]
        return "".join(f"{max(0,min(255,v)):02X}" for v in vals)
    except:return rgb

def fill_color(cell, theme_map)->str:
    fill=cell.fill
    if not fill or fill.patternType in (None,"none"):
        return "FFFFFF"
    c=fill.fgColor
    typ=getattr(c,"type",None)
    if typ=="rgb" and c.rgb:
        return c.rgb[-6:].upper()
    if typ=="indexed" and c.indexed is not None:
        try:return COLOR_INDEX[int(c.indexed)][-6:].upper()
        except:return f"INDEXED:{c.indexed}"
    if typ=="theme" and c.theme is not None:
        base=theme_map.get(int(c.theme))
        if base:return _apply_tint(base,getattr(c,"tint",0))
        return f"THEME:{c.theme}:{getattr(c,'tint',0)}"
    return str(c.value or "FFFFFF").upper()

def is_white(color:str)->bool:
    c=(color or "").upper().replace("#","")
    return c in {"FFFFFF","FFFFFFFF","00000000","NONE",""}

def _sheet_name_match(wb, wanted:str)->str|None:
    wanted_norm=re.sub(r"[\s/\\_-]+","",wanted.casefold())
    exact=[s for s in wb.sheetnames if re.sub(r"[\s/\\_-]+","",s.casefold())==wanted_norm]
    if exact:return exact[0]
    # cautious aliases only, no fuzzy matching
    aliases={
      "етфт":["етфт","ет/фт","ет-фт","ет фт"],
      "фрм":["фрм","frm"],
    }
    candidates=aliases.get(wanted_norm,[])
    for s in wb.sheetnames:
        sn=re.sub(r"[\s/\\_-]+","",s.casefold())
        if any(sn==re.sub(r"[\s/\\_-]+","",a.casefold()) for a in candidates):return s
    return None

def detect_date_header(ws, year:int, max_row:int|None=None, max_column:int|None=None):
    best=(0,None,{})
    sheet_max_row=ws.max_row if max_row is None else max_row
    sheet_max_column=ws.max_column if max_column is None else max_column
    header_scan_max_row=min(sheet_max_row,30)
    for r in range(1,header_scan_max_row+1):
        dates={}
        for c in range(3,sheet_max_column+1):
            d=parse_date(ws.cell(r,c).value,year)
            if d: dates[c]=d
        if len(dates)>best[0]:best=(len(dates),r,dates)
    return best[1],best[2]

def detect_time_rows(ws, header_row:int|None, max_row:int|None=None):
    rows={}
    start=(header_row+1 if header_row else 1)
    sheet_max_row=ws.max_row if max_row is None else max_row
    for r in range(start,sheet_max_row+1):
        t=parse_time(ws.cell(r,2).value)
        if t: rows[r]=t
    return rows

def _merged_nonstandard(ws, row:int,col:int)->str|None:
    coord=ws.cell(row,col).coordinate
    for rng in ws.merged_cells.ranges:
        if coord in rng:
            if rng.min_row!=rng.max_row or rng.min_col!=rng.max_col:
                return str(rng)
    return None

def parse_schedule(workbook_or_path:Any, sheet_name:str, year:int=2026):
    """Parse one sheet.

    ``workbook_or_path`` intentionally accepts an already opened workbook so a
    combined check can parse both sheets without paying the XLSX load cost twice.
    The path form remains compatible with the existing independent workflows.
    """
    opened_here=isinstance(workbook_or_path,(str,bytes))
    wb=(load_workbook(workbook_or_path,data_only=False,rich_text=True)
        if opened_here else workbook_or_path)
    def finish(appointments,warnings):
        if opened_here:wb.close()
        return appointments,warnings
    raise_if_cancelled()
    actual=_sheet_name_match(wb,sheet_name)
    if not actual:
        return finish([],[StructureWarning(sheet_name,"",f"Вкладку «{sheet_name}» не знайдено. Доступні: {', '.join(wb.sheetnames)}")])
    ws=wb[actual]
    # ``Worksheet.max_row`` and ``max_column`` scan the internal cell mapping
    # on every access.  ET/FT contains thousands of appointments, so reading
    # max_row inside that loop used to rescan the sheet thousands of times.
    max_row=ws.max_row
    max_column=ws.max_column
    theme=_theme_colors(wb)
    header_row,dates=detect_date_header(ws,year,max_row,max_column)
    warnings=[]
    if not dates:
        return finish([],[StructureWarning(actual,"","Не знайдено рядок з розпізнаваними датами.")])
    time_rows=detect_time_rows(ws,header_row,max_row)
    if not time_rows:
        return finish([],[StructureWarning(actual,"","Не знайдено 30-хвилинну часову сітку у колонці B.")])
    # Do not silently ignore patient-looking cells in columns whose header date was not recognized.
    for c in range(3,max_column+1):
        if c in dates:continue
        samples=[]
        for r in list(time_rows.keys())[:min(len(time_rows),200)]:
            v=ws.cell(r,c).value
            if v is not None and display_name(v):
                samples.append(ws.cell(r,c).coordinate)
                if len(samples)>=3:break
        if samples:
            header_value=ws.cell(header_row or 1,c).value
            warnings.append(StructureWarning(actual,samples[0],
                f"У колонці є записи, але дата не розпізнана (заголовок: {header_value!r}); колонку не перевірено. Приклади: {', '.join(samples)}."))
    rows_sorted=sorted(time_rows)
    # Resolve non-standard merges once. Re-checking every populated cell
    # against every merged range is needlessly expensive on large ET/FT sheets.
    nonstandard_merged={}
    for rng in ws.merged_cells.ranges:
        if rng.min_row==rng.max_row and rng.min_col==rng.max_col:
            continue
        label=str(rng)
        for merged_row in range(rng.min_row,rng.max_row+1):
            for merged_column in range(rng.min_col,rng.max_col+1):
                nonstandard_merged[(merged_row,merged_column)]=label
    # section/specialist uses column A and is filled down
    current_specialist=""
    specialist_by_row={}
    for r in range((header_row or 1)+1,max_row+1):
        v=ws.cell(r,1).value
        if v is not None and str(v).strip():
            current_specialist=str(v).strip()
        specialist_by_row[r]=current_specialist

    appointments=[]
    occupied=set()
    for r in rows_sorted:
        raise_if_cancelled()
        t=time_rows[r]
        if t.minute not in (0,30):
            warnings.append(StructureWarning(actual,f"B{r}",f"Час {t.strftime('%H:%M')} не кратний 30 хвилинам."))
            continue
        for c,d in dates.items():
            cell=ws.cell(r,c)
            value=cell.value
            if value is None:
                continue
            if not display_name(value):
                warnings.append(StructureWarning(actual,cell.coordinate,"Знайдено порожнє ПІБ; клітинку пропущено."))
                continue
            if (r,c) in occupied:
                warnings.append(StructureWarning(actual,cell.coordinate,"Одна клітинка інтерпретується повторно; запис пропущено."))
                continue
            merged=nonstandard_merged.get((r,c))
            if merged:
                warnings.append(StructureWarning(actual,cell.coordinate,f"Нестандартно об'єднана клітинка {merged}; інтервал не вгадується."))
                continue
            name=display_name(value)
            original_name=str(value).strip()
            # reject obvious headers/non-patient tokens conservatively
            if len(name)<3:
                warnings.append(StructureWarning(actual,cell.coordinate,"Порожнє або надто коротке ПІБ; запис пропущено."))
                continue
            color=fill_color(cell,theme)
            if color.startswith("THEME:"):
                warnings.append(StructureWarning(actual,cell.coordinate,"Не вдалося перетворити theme-fill у фактичний RGB; тривалість за кольором не буде вгадуватися."))
            duration=30
            cont=None
            next_r=r+1
            if next_r<=max_row and next_r in time_rows:
                nt=time_rows[next_r]
                dt1=datetime.combine(d,t); dt2=datetime.combine(d,nt)
                if int((dt2-dt1).total_seconds()/60)==30:
                    low=ws.cell(next_r,c)
                    low_color=fill_color(low,theme)
                    if color.startswith("THEME:") or low_color.startswith("THEME:"):
                        if (low.value is None or not display_name(low.value)) and low_color==color:
                            warnings.append(StructureWarning(actual,low.coordinate,"Колір продовження theme не розпізнано до RGB; 60 хв не встановлено автоматично."))
                    elif (low.value is None or not display_name(low.value)) and (not is_white(color)) and low_color==color and not is_white(low_color):
                        duration=60
                        cont=low.coordinate
                        occupied.add((next_r,c))
            elif r==rows_sorted[-1] and not is_white(color):
                warnings.append(StructureWarning(actual,cell.coordinate,"Запис стоїть в останньому часовому рядку; нижню клітинку для тривалості перевірити неможливо."))
            start_dt=datetime.combine(d,t)
            end_dt=start_dt+timedelta(minutes=duration)
            appointments.append(Appointment(
                patient_name_original=original_name,
                patient_name_normalized=normalize_patient_name(name),
                sheet_name=actual,
                date=d.isoformat(),
                start_time=t.strftime("%H:%M"),
                end_time=end_dt.time().strftime("%H:%M"),
                duration_minutes=duration,
                specialist=specialist_by_row.get(r,""),
                cell_address=cell.coordinate,
                fill_color=color,
                continuation_cell_address=cont,
                row=r,column=c
            ))
            occupied.add((r,c))
    return finish(appointments,warnings)

def build_index(appointments):
    idx=defaultdict(lambda:defaultdict(list))
    for a in appointments:
        idx[a.patient_name_normalized][a.date].append(a)
    for p in idx.values():
        for arr in p.values():
            arr.sort(key=lambda x:(x.start_time,x.end_time,x.cell_address))
    return idx

def _frm_ar3_info(first:Appointment, next_:Appointment, gap:int)->Ar3MonoserviceInfo:
    return Ar3MonoserviceInfo(
        patient=first.patient_name_original,
        first_date=first.date,
        next_date=next_.date,
        gap_days=gap,
        first_cell=first.cell_address,
        next_cell=next_.cell_address,
    )

def classify_frm_occurrences(frm:list[Appointment]):
    """Classify hourly ФРМ visits without guessing damaged occurrence sequences.

    Only adjacent records at a 14--16 calendar-day distance form a pair.  An
    unpaired record remains eligible for the order check: it is explicitly not
    silently discarded merely because it happens to be the second row.
    """
    hourly=[a for a in frm if a.duration_minutes==60 and a.continuation_cell_address]
    by_patient=defaultdict(list)
    for appointment in hourly:
        appointment.frm_occurrence_type=None
        appointment.paired_with_cell=None
        appointment.paired_with_date=None
        appointment.pair_gap_days=None
        by_patient[appointment.patient_name_normalized].append(appointment)

    primary=[]; repeat=[]; unpaired=[]; ar3_infos=[]
    for records in by_patient.values():
        records.sort(key=lambda a:(a.date,a.start_time,a.cell_address))
        date_counts=defaultdict(int)
        for record in records:
            date_counts[record.date]+=1
        dates_with_duplicates={record_date for record_date,count in date_counts.items() if count>1}

        i=0
        while i<len(records):
            current=records[i]
            if i+1>=len(records):
                current.frm_occurrence_type="UNPAIRED"
                unpaired.append(current)
                i+=1
                continue
            next_=records[i+1]
            gap=(date.fromisoformat(next_.date)-date.fromisoformat(current.date)).days
            # A duplicated date makes either side ambiguous, even when a
            # neighbouring gap happens to fall in the nominal range.
            if 14<=gap<=16 and current.date not in dates_with_duplicates and next_.date not in dates_with_duplicates:
                current.frm_occurrence_type="PRIMARY"
                next_.frm_occurrence_type="REPEAT"
                current.paired_with_cell=next_.cell_address
                current.paired_with_date=next_.date
                current.pair_gap_days=gap
                next_.paired_with_cell=current.cell_address
                next_.paired_with_date=current.date
                next_.pair_gap_days=gap
                primary.append(current)
                repeat.append(next_)
                i+=2
            else:
                current.frm_occurrence_type="UNPAIRED"
                unpaired.append(current)
                ar3_infos.append(_frm_ar3_info(current,next_,gap))
                i+=1
    return hourly,primary,repeat,unpaired,ar3_infos

def build_schedule_index(workbook_path:str, frm_sheet:str, etft_sheet:str, year:int=2026)->ScheduleIndex:
    """Open an XLSX/XLSM once, parse both schedule sheets, then index them."""
    set_progress(12,"Зчитуємо Excel один раз для двох перевірок")
    wb=load_workbook(workbook_path,data_only=False,rich_text=True)
    try:
        frm,frm_warnings=parse_schedule(wb,frm_sheet,year)
        set_progress(43,"Розібрано вкладку ФРМ")
        etft,etft_warnings=parse_schedule(wb,etft_sheet,year)
        set_progress(67,"Розібрано вкладку ЕТ/ФТ")
        hourly,primary,repeat,unpaired,ar3_infos=classify_frm_occurrences(frm)
        return ScheduleIndex(
            frm_appointments=frm,
            etft_appointments=etft,
            etft_by_patient_date=build_index(etft),
            frm_hourly=hourly,
            frm_primary=primary,
            frm_repeat=repeat,
            frm_unpaired=unpaired,
            structure_warnings=frm_warnings+etft_warnings,
            ar3_monoservice=ar3_infos,
        )
    finally:
        wb.close()

def check_frm_before_etft(frm,etft,etft_by_patient_date=None):
    et_idx=etft_by_patient_date if etft_by_patient_date is not None else build_index(etft)
    errors=[]
    hourly=[a for a in frm if a.duration_minutes==60 and a.continuation_cell_address]
    # Direct callers retain the old two-list API while receiving the corrected
    # pair-aware rule.  IMPORTANT_000 supplies an already classified index.
    if any(a.frm_occurrence_type is None for a in hourly):
        classify_frm_occurrences(frm)
    checkable=(a for a in hourly if a.frm_occurrence_type!="REPEAT")
    for f in checkable:
        raise_if_cancelled()
        f_end=datetime.strptime(f.end_time,"%H:%M").time()
        for e in et_idx.get(f.patient_name_normalized,{}).get(f.date,[]):
            e_start=datetime.strptime(e.start_time,"%H:%M").time()
            if e_start < f_end:
                errors.append({
                  "type":"ЕТ/ФТ назначено раніше закінчення ФРМ",
                  "patient":f.patient_name_original,"date":f.date,
                  "frm":{"sheet":f.sheet_name,"cell":f.cell_address,"interval":f"{f.start_time}–{f.end_time}","specialist":f.specialist},
                  "etft":{"sheet":e.sheet_name,"cell":e.cell_address,"interval":f"{e.start_time}–{e.end_time}","specialist":e.specialist},
                  "message":f"{f.patient_name_original}, {f.date}. Прийом ФРМ: {f.start_time}–{f.end_time}. На ЕТ/ФТ знайдено {e.start_time}–{e.end_time}; запис може починатися тільки з {f.end_time}."
                })
    return errors, len(hourly)

def check_etft_overlaps(etft):
    """Find double-booked ET/FT slots using the fixed 30-minute procedure rule.

    ``Appointment.duration_minutes`` retains the Excel visual-block metadata for
    other checks.  It must not turn an ET/FT procedure into a one-hour interval.
    """
    idx=build_index(etft)
    errors=[]
    for patient,by_date in idx.items():
        raise_if_cancelled()
        for d,arr in by_date.items():
            # output-sensitive sweep line: no full all-pairs scan when intervals do not overlap
            active=[]
            for cur in arr:
                cs=datetime.strptime(cur.start_time,"%H:%M")
                ce=cs+timedelta(minutes=30)
                # keep only intervals with end > current start
                active=[entry for entry in active if entry[2]>cs]
                for prev,ps,pe in active:
                    if ps < ce and cs < pe:
                        ov_start=max(ps,cs).strftime("%H:%M"); ov_end=min(pe,ce).strftime("%H:%M")
                        errors.append({
                          "type":"Пересечение приёмов на ЕТ/ФТ",
                          "patient":cur.patient_name_original,"date":d,
                          "first":{"interval":f"{ps.strftime('%H:%M')}–{pe.strftime('%H:%M')}","specialist":prev.specialist,"cell":prev.cell_address},
                          "second":{"interval":f"{cs.strftime('%H:%M')}–{ce.strftime('%H:%M')}","specialist":cur.specialist,"cell":cur.cell_address},
                          "overlap":f"{ov_start}–{ov_end}",
                          "message":f"{cur.patient_name_original}, {d}. Перетин: {ps.strftime('%H:%M')}–{pe.strftime('%H:%M')} і {cs.strftime('%H:%M')}–{ce.strftime('%H:%M')}. Фактичний перетин {ov_start}–{ov_end}."
                        })
                active.append((cur,cs,ce))
    return errors

def serialize(items):
    return [x.as_dict() if hasattr(x,"as_dict") else x for x in items]
