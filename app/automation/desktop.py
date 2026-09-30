from __future__ import annotations
import os, subprocess, time, re
from pathlib import Path
from ..core.audit import log_event
from ..core.action_control import raise_if_cancelled, set_progress
from .browser import ManualCheckpoint, AutomationError

class DesktopEleks:
    def __init__(self, action_id:str, settings:dict):
        self.action_id=action_id; self.settings=settings; self.app=None; self.window=None

    def connect(self):
        raise_if_cancelled()
        if os.name!="nt":
            raise AutomationError("Doctor Eleks Desktop automation підтримується тільки у Windows.")
        try:
            from pywinauto import Application, Desktop
        except Exception as e:
            raise AutomationError("pywinauto не встановлено.") from e
        # First connect to existing window
        try:
            self.window=Desktop(backend="uia").window(title_re=r".*(Doctor\s*Eleks|Доктор\s*Елекс).*")
            self.window.wait("visible",timeout=3)
            return self
        except Exception:pass
        exe=self.settings.get("desktop_eleks_path","")
        if not exe or not Path(exe).exists():
            raise ManualCheckpoint("desktop_path","Doctor Eleks Desktop не запущено і шлях до EXE не налаштовано.")
        set_progress(15,"Запускаємо Doctor Eleks Desktop")
        Application(backend="uia").start(exe)
        self.window=Desktop(backend="uia").window(title_re=r".*(Doctor\s*Eleks|Доктор\s*Елекс).*")
        self.window.wait("visible",timeout=10)
        raise ManualCheckpoint(
            "desktop_ready",
            "Doctor Eleks Desktop відкрито. Переконайтеся, що програма завантажилася та авторизуйтеся за потреби, після чого повторіть запуск.",
        )

    def _controls(self,typ=None):
        try:return self.window.descendants(control_type=typ) if typ else self.window.descendants()
        except:return []

    def dump_tree(self):
        rows=[]
        for c in self._controls():
            try:
                rows.append({"type":c.element_info.control_type,"name":c.window_text()[:160],
                             "auto_id":c.element_info.automation_id})
            except:pass
        return rows[:2000]

    def click_text(self,needles):
        raise_if_cancelled()
        needles=[n.casefold() for n in (needles if isinstance(needles,list) else [needles])]
        for c in self._controls():
            try:
                txt=c.window_text().strip()
                if txt and any(n in txt.casefold() for n in needles):
                    c.click_input()
                    log_event("desktop.click",f"Натиснуто: {txt}",action_id=self.action_id)
                    time.sleep(.4);return True
            except:continue
        return False

    def fill_first_edit(self,value):
        raise_if_cancelled()
        edits=self._controls("Edit")
        if not edits:return False
        # prefer visible enabled edit
        for e in edits:
            try:
                if e.is_visible() and e.is_enabled():
                    e.set_edit_text(value)
                    log_event("desktop.fill","Заповнено поле пошуку",action_id=self.action_id)
                    return True
            except:continue
        return False

    def find_neurologist_report(self,patient_name:str):
        self.connect()
        set_progress(30,"Шукаємо висновок невролога в Doctor Eleks Desktop")
        self.window.set_focus()
        # Home/patient screen: fill visible edit then search
        if not self.fill_first_edit(patient_name.split()[0]):
            raise AutomationError("Не знайдено поле вводу пацієнта у Doctor Eleks Desktop.")
        self.click_text(["Пошук","Поиск"])
        time.sleep(1)
        if not self.click_text([patient_name, patient_name.split()[0]]):
            raise ManualCheckpoint("desktop_patient","Не вдалося однозначно відкрити картку пацієнта. Відкрийте її вручну і повторіть.")
        time.sleep(.7)
        # Documents first, then Files
        for section in (["Документи","Документы"],["Файли пацієнта","Файлы пациента"]):
            raise_if_cancelled()
            if not self.click_text(section):continue
            time.sleep(.7)
            if self.click_text(["консультативний висновок невролога","консультативное заключение невролога","невролог"]):
                time.sleep(.8)
                texts=[]
                for c in self._controls():
                    try:
                        txt=c.window_text().strip()
                        if len(txt)>20:texts.append(txt)
                    except:pass
                content="\n".join(dict.fromkeys(texts))
                if len(content)>100:
                    return content
        # diagnostics without OCR: dump tree for later selector tuning
        tree=self.dump_tree()
        log_event("desktop.report.not_found","Висновок невролога не знайдено автоматично",level="WARNING",
                  action_id=self.action_id,data={"controls":tree[:120]})
        raise ManualCheckpoint("desktop_report",
          "Автоматично не вдалося витягти висновок. Відкрийте консультативний висновок невролога вручну і повторіть; журнал містить дерево UI для калібрування.")
