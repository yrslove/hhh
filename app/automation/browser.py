from __future__ import annotations
import json, os, socket, subprocess, time, shutil, re
from pathlib import Path
from contextlib import contextmanager
from threading import Lock
from urllib.parse import urlparse
from urllib.request import urlopen
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from ..core.paths import ROOT, RUNTIME_DIR, ERROR_DIR, BROWSER_LOG
from ..core.config import load_json
from ..core.audit import log_event
from ..core.action_control import ActionCancelled, raise_if_cancelled, set_progress
from ..core.text import normalize_patient_name

_driver_lock=Lock()
_active_drivers={}


def stop_active_playwright_drivers():
    """Stop Playwright driver processes while leaving the Chrome browser open."""
    with _driver_lock:
        drivers=list(_active_drivers.items())
    stopped=0
    for _pid,proc in drivers:
        try:
            if proc.returncode is None:
                proc.kill()
                stopped+=1
        except (ProcessLookupError,OSError):
            pass
    return stopped

class ManualCheckpoint(RuntimeError):
    def __init__(self, code:str, message:str, data=None):
        super().__init__(message); self.code=code; self.data=data or {}

class AutomationError(RuntimeError): pass

def wait_visible_recheck(page, locate, initial_ms=1800, retry_ms=1800):
    """Resolve a fresh Locator after a two-second UI-render grace period."""
    try:
        locator=locate()
        locator.wait_for(state="visible",timeout=initial_ms)
        return locator
    except ActionCancelled:raise
    except Exception:
        page.wait_for_timeout(2000)
        locator=locate()
        locator.wait_for(state="visible",timeout=retry_ms)
        return locator

def _is_ehealth_application_frame(url:str)->bool:
    return "/web/ehealth/internal" in str(url or "").casefold()

def _row_has_full_patient_name(row_text:str, patient_name:str)->bool:
    """Match a result row by full normalized PІБ, never by surname alone."""
    target=normalize_patient_name(patient_name)
    return bool(target and target in normalize_patient_name(row_text))

def _patient_name_identity(patient_name:str):
    parts=normalize_patient_name(patient_name).split()
    surname=parts[0] if parts else ""
    initials=tuple(part[0] for part in parts[1:3] if part)
    return surname,initials

def _row_has_patient_surname(row_text:str,patient_name:str)->bool:
    surname,_=_patient_name_identity(patient_name)
    row=normalize_patient_name(row_text)
    return bool(surname and re.search(rf"(?<!\w){re.escape(surname)}(?!\w)",row))

def _row_has_patient_identity(row_text:str,patient_name:str)->bool:
    """Accept either full PІБ or the same surname followed by matching initials."""
    if _row_has_full_patient_name(row_text,patient_name):return True
    surname,initials=_patient_name_identity(patient_name)
    if not surname or not initials:return False
    row=normalize_patient_name(row_text)
    match=re.search(rf"(?<!\w){re.escape(surname)}(?!\w)",row)
    if not match:return False
    # Handles both «Приліпко Василь Олегович» and «Приліпко В. О.».
    following=re.findall(r"[^\W\d_]+(?:[-'][^\W\d_]+)*",row[match.end():],flags=re.UNICODE)
    return len(following)>=len(initials) and all(
        following[index].startswith(initial) for index,initial in enumerate(initials)
    )

def _coerce_birth_date(value):
    from datetime import date,datetime
    if not value:return None
    if isinstance(value,datetime):return value.date()
    if isinstance(value,date):return value
    raw=str(value).strip()
    for fmt in ("%Y-%m-%d","%d.%m.%Y","%d/%m/%Y","%d-%m-%Y"):
        try:return datetime.strptime(raw,fmt).date()
        except ValueError:pass
    return None

def _row_matches_birth_date(row_text:str,dob)->bool:
    from datetime import date
    birth=_coerce_birth_date(dob)
    if not birth:return False
    text=str(row_text or "")
    variants={birth.isoformat(),birth.strftime("%d.%m.%Y"),birth.strftime("%d/%m/%Y"),birth.strftime("%d-%m-%Y")}
    if any(value in text for value in variants):return True
    today=date.today();age=today.year-birth.year-((today.month,today.day)<(birth.month,birth.day))
    return bool(re.search(rf"(?<!\d){age}(?!\d)",text))

def _port_open(port:int)->bool:
    with socket.socket() as s:
        s.settimeout(.2)
        return s.connect_ex(("127.0.0.1",port))==0

def _profile_debug_port(profile:Path)->int|None:
    """Reuse only the Chrome endpoint recorded by this browser profile."""
    try:
        port_text,browser_path=(profile/"DevToolsActivePort").read_text(encoding="utf-8").splitlines()[:2]
        port=int(port_text)
        if not _port_open(port):return None
        with urlopen(f"http://127.0.0.1:{port}/json/version",timeout=1) as response:
            endpoint=json.load(response).get("webSocketDebuggerUrl","")
        return port if urlparse(endpoint).path==browser_path else None
    except (OSError,ValueError,IndexError,KeyError,TypeError):
        return None

def _seed_browser_profile(profile:Path,source:Path):
    """Copy account state once, without restoring the source's open tabs."""
    marker=profile/".account_profile_seeded"
    if marker.exists():return
    source_state=source/"Local State"
    source_default=source/"Default"
    try:
        state=json.loads(source_state.read_text(encoding="utf-8"))
        account=(state.get("profile",{}).get("info_cache",{})
                 .get("Default",{}).get("user_name") or "")
    except (OSError,ValueError,AttributeError) as exc:
        raise AutomationError("Не вдалося прочитати профіль Chrome Valeriia для процедур.") from exc
    if not account:
        raise AutomationError("У вихідному профілі Chrome немає облікового запису Valeriia.")
    target_default=profile/"Default"
    target_default.mkdir(parents=True,exist_ok=True)
    try:
        shutil.copy2(source_state,profile/"Local State")
        for name in ("Preferences","Secure Preferences","Web Data","Account Web Data"):
            src=source_default/name
            if src.is_file():shutil.copy2(src,target_default/name)
        # Chrome keeps its Cookies database open with an exclusive Windows
        # handle. The copied account identity is usable without copying that
        # live database; Doctor Eleks may ask for a one-time login.
        network=source_default/"Network"
        if network.is_dir():
            shutil.copytree(network,target_default/"Network",dirs_exist_ok=True,
                            ignore=shutil.ignore_patterns("Cookies","Cookies-journal"))
        for name in ("Local Storage","Sync Data"):
            src=source_default/name
            if src.is_dir():shutil.copytree(src,target_default/name,dirs_exist_ok=True)
    except (OSError,shutil.Error) as exc:
        raise AutomationError("Не вдалося скопіювати стан профілю Chrome Valeriia. Закрийте його вікно та повторіть запуск.") from exc
    marker.write_text("ready",encoding="utf-8")
    log_event("browser.profile.seeded","Для процедур підготовлено профіль Chrome Valeriia без інших вкладок.",
              data={"profile":str(profile)})

def _detect_browser(configured:str=""):
    candidates=[]
    if configured:candidates.append(configured)
    env=[os.environ.get("PROGRAMFILES",""),os.environ.get("PROGRAMFILES(X86)",""),os.environ.get("LOCALAPPDATA","")]
    for root in env:
        if not root:continue
        candidates += [
          os.path.join(root,"Microsoft","Edge","Application","msedge.exe"),
          os.path.join(root,"Google","Chrome","Application","chrome.exe")
        ]
    for p in candidates:
        if p and Path(p).exists():return str(Path(p))
    return shutil.which("msedge") or shutil.which("chrome")

def _ehealth_url(settings:dict)->str:
    """Require the clinic's configured Doctor Eleks URL; never launch blank."""
    url=str(settings.get("ehealth_base_url") or "").strip()
    parsed=urlparse(url)
    if not url or parsed.scheme not in {"http","https"} or not parsed.netloc:
        raise ManualCheckpoint(
            "ehealth_url",
            "Не вказано адресу Doctor Eleks Web. У Налаштуваннях → Doctor Eleks Web заповніть «Адреса eHealth/Doctor Eleks» (https://…), збережіть і повторіть запуск.",
        )
    return url

def ensure_browser(settings:dict,action_id:str|None=None):
    url=_ehealth_url(settings)
    port=int(settings.get("remote_debug_port",9222))
    profile=Path(settings.get("browser_profile_dir") or RUNTIME_DIR/"browser_profile")
    profile.mkdir(parents=True,exist_ok=True)
    if port==0:
        existing=_profile_debug_port(profile)
        if existing:
            log_event("browser.launch.reused","Використовуємо вже відкрите вікно Doctor Eleks.",
                      action_id=action_id,data={"port":existing,"profile":str(profile)})
            return existing,False
    elif _port_open(port):
        log_event("browser.launch.reused","Використовуємо вже відкрите вікно Doctor Eleks.",
                  action_id=action_id,data={"port":port,"profile":str(profile)})
        return port,False
    seed=settings.get("browser_profile_seed_dir")
    if seed:_seed_browser_profile(profile,Path(seed))
    exe=_detect_browser(settings.get("browser_executable",""))
    if not exe:raise AutomationError("Не знайдено Microsoft Edge або Google Chrome. Вкажіть browser_executable у Налаштуваннях.")
    if port==0:
        (profile/"DevToolsActivePort").unlink(missing_ok=True)
    args=[exe,f"--remote-debugging-port={port}","--remote-allow-origins=*",
          f"--user-data-dir={profile}",
          "--no-first-run","--no-default-browser-check","--new-window",url]
    log_event("browser.launch.start","Запускаємо окреме вікно Doctor Eleks.",
              action_id=action_id,data={"profile":str(profile),"dynamic_port":port==0})
    process=subprocess.Popen(args,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
    for _ in range(80):
        raise_if_cancelled()
        ready=_profile_debug_port(profile) if port==0 else port if _port_open(port) else None
        if ready:
            log_event("browser.launch.ready","Браузер відкрито, порт керування готовий.",
                      action_id=action_id,data={"port":ready,"profile":str(profile)})
            return ready,True
        if process.poll() is not None:break
        time.sleep(.25)
    raise AutomationError("Не вдалося відкрити керований Chrome: порт браузера не став доступним. Перевірте вікно Chrome та повторіть запуск.")

class WebEleks:
    def __init__(self, action_id:str, settings:dict, safe_mode:bool|None=None,
                 confirmed_authenticated:bool=False):
        self.action_id=action_id; self.settings=settings
        self.safe_mode=settings.get("safe_mode",True) if safe_mode is None else safe_mode
        # Set only by an explicit local «Продовжити» after browser_ready/login.
        # It acknowledges the operator's visual confirmation, not an API login.
        self.confirmed_authenticated=confirmed_authenticated;self._auth_confirmation_logged=False
        self.selector_map=load_json("selectors.json",{})
        # ``browser_page`` is the outer page; ``page`` is the interactive
        # Doctor Eleks frame when the site is served through a legacy frameset.
        self._pw=None; self.browser=None; self.context=None; self.browser_page=None; self.page=None
        self._driver_pids=[]

    def _application_frame(self):
        """Wait for the live Doctor Eleks frame instead of its frameset shell."""
        for _ in range(60):
            raise_if_cancelled()
            for frame in self.browser_page.frames:
                if _is_ehealth_application_frame(frame.url):
                    return frame
            self.browser_page.wait_for_timeout(250)
        raise ManualCheckpoint(
            "ehealth_frame",
            "Відкрито оболонку Doctor Eleks, але робоча сторінка ще не завантажилася. Зачекайте появи форми пошуку пацієнта і натисніть «Продовжити».",
            {"url":self.browser_page.url},
        )

    def __enter__(self):
        set_progress(15,"Відкриваємо Doctor Eleks у браузері")
        url=_ehealth_url(self.settings)
        port,launched=ensure_browser(self.settings,self.action_id)
        if launched:
            raise ManualCheckpoint(
                "browser_ready",
                "Doctor Eleks відкрито. Переконайтеся, що вікно з'явилося, авторизуйтеся за потреби й повторіть запуск для продовження.",
            )
        raise_if_cancelled()
        endpoint=f"http://127.0.0.1:{port}"
        last_error=None
        for attempt in range(2):
            raise_if_cancelled()
            log_event("browser.cdp_connect.start","Підключаємо Playwright до відкритого браузера.",
                      action_id=self.action_id,data={"port":port,"attempt":attempt+1})
            try:
                self._pw=sync_playwright().start()
                proc=self._pw._impl_obj._connection._transport._proc
                self._driver_pids.append((proc.pid,proc))
                with _driver_lock:
                    _active_drivers[proc.pid]=proc
                raise_if_cancelled()
                # A stale/unresponsive remote-debugging endpoint used to leave
                # the whole action blocked here with no cancellation boundary.
                # Bound each attempt so the operator gets a visible result and
                # can safely retry after restarting the browser.
                self.browser=self._pw.chromium.connect_over_cdp(endpoint,timeout=10000)
                log_event("browser.cdp_connect.success","Playwright підключився до браузера.",
                          action_id=self.action_id,data={"port":port,"attempt":attempt+1})
                break
            except Exception as exc:
                last_error=exc
                if self._pw:
                    try:self._pw.stop()
                    except Exception:pass
                    self._pw=None
                log_event("browser.cdp_connect.failed","Не вдалося підключитися до браузера.",
                          level="WARNING",action_id=self.action_id,
                          data={"port":port,"attempt":attempt+1,"error":str(exc)[:500]})
                raise_if_cancelled()
                if attempt==0:
                    time.sleep(.5)
        if self.browser is None:
            self.__exit__()
            raise ManualCheckpoint(
                "browser_connection",
                "Playwright не зміг ініціалізувати відкритий браузер. "
                "Закрийте лише вікно Doctor Eleks, відкрите цією функцією, та запустіть її повторно.",
                {"port":port,"error":str(last_error)[:500]},
            ) from last_error
        try:
            self.context=self.browser.contexts[0] if self.browser.contexts else self.browser.new_context()
            pages=self.context.pages
            base=url.casefold()
            page=None
            if base:
                page=next((p for p in pages if base in p.url.casefold()),None)
            self.browser_page=page or (pages[-1] if pages else self.context.new_page())
            if base and (not self.browser_page.url or self.browser_page.url=="about:blank"):
                self.browser_page.goto(url,wait_until="domcontentloaded",timeout=30000)
            self.page=self._application_frame()
            set_progress(22,"Doctor Eleks відкрито, перевіряємо доступ")
            raise_if_cancelled()
            return self
        except Exception:
            self.__exit__()
            raise

    def __exit__(self,*exc):
        with _driver_lock:
            for pid,proc in self._driver_pids:
                if _active_drivers.get(pid) is proc:
                    _active_drivers.pop(pid,None)
        self._driver_pids=[]
        try:
            if self._pw:self._pw.stop() # detaches; externally launched browser remains
        except: pass

    def _log(self,kind,alias,data=None):
        payload={"kind":kind,"alias":alias,"url":self.page.url if self.page else "",**(data or {})}
        log_event("browser.action",f"{kind}: {alias}",action_id=self.action_id,data=payload)
        try:
            with BROWSER_LOG.open("a",encoding="utf-8") as f:
                import datetime
                f.write(json.dumps({"time":datetime.datetime.now().isoformat(),**payload},ensure_ascii=False)+"\n")
        except:pass

    def snapshot_error(self,label:str):
        stamp=time.strftime("%Y%m%d_%H%M%S")
        d=ERROR_DIR/f"{stamp}_{self.action_id}_{label}"
        d.mkdir(parents=True,exist_ok=True)
        try:self.browser_page.screenshot(path=str(d/"screen.png"),full_page=True)
        except:pass
        try:(d/"page.html").write_text(self.page.content(),encoding="utf-8")
        except:pass
        return str(d)

    def _candidate_locator(self,candidate:str):
        if candidate.startswith("role="):
            rest=candidate[5:]
            role,name=(rest.split("|",1)+[""])[:2]
            return self.page.get_by_role(role,name=name,exact=False)
        if candidate.startswith("text="):
            return self.page.get_by_text(candidate[5:],exact=False)
        if candidate.startswith("label="):
            return self.page.get_by_label(candidate[6:],exact=False)
        if candidate.startswith("placeholder="):
            return self.page.get_by_placeholder(candidate[12:],exact=False)
        if candidate.startswith("css="):
            return self.page.locator(candidate[4:])
        return self.page.get_by_text(candidate,exact=False)

    def locator(self,alias:str,timeout=1500):
        raise_if_cancelled()
        candidates=self.selector_map.get(alias,[alias])
        for attempt in range(2):
            deadline=time.monotonic()+timeout/1000
            while True:
                raise_if_cancelled()
                for c in candidates:
                    try:
                        loc=self._candidate_locator(c)
                        if loc.count()>0 and loc.first.is_visible():return loc.first
                    except ActionCancelled:raise
                    except Exception:pass
                if time.monotonic()>=deadline:break
                self.page.wait_for_timeout(200)
            if attempt==0:
                self._log("locator.recheck",alias,{"reason":"not_visible"})
                self.page.wait_for_timeout(2000)
        raise AutomationError(f"Не знайдено елемент «{alias}». Перевірте сторінку/селектори.")

    def wait_for_alias(self, alias:str, *, enabled=False):
        """Quiet staged readiness check used by longer transactional workflows."""
        deadlines=(2.0,5.0,5.0)
        for stage,seconds in enumerate(deadlines,1):
            end=time.monotonic()+seconds
            while time.monotonic()<end:
                raise_if_cancelled()
                for candidate in self.selector_map.get(alias,[alias]):
                    try:
                        loc=self._candidate_locator(candidate).first
                        if loc.count() and loc.is_visible() and (not enabled or loc.is_enabled()):
                            return loc
                    except ActionCancelled:raise
                    except Exception:pass
                self.page.wait_for_timeout(250)
        self.snapshot_error("wait_"+re.sub(r"[^a-zA-Z0-9_-]+","_",alias))
        raise AutomationError(f"Не знайдено готовий елемент «{alias}» після очікування 2+5+5 секунд.")

    def click(self,alias:str,optional=False,*,locator=None):
        try:
            raise_if_cancelled()
            loc=locator if locator is not None else self.locator(alias)
            self._log("click",alias,{"text":(loc.inner_text(timeout=500) or "")[:160]})
            if self.safe_mode and alias in {"save_sign","sign_selected"}:
                raise ManualCheckpoint("safe_mode",f"Безпечний режим зупинив фінальне натискання «{alias}». Вимкніть safe_mode після перевірки.")
            before_url=self.page.url
            loc.click(timeout=5000)
            self.page.wait_for_timeout(300)
            self._log("click.ok",alias,{"before_url":before_url,"after_url":self.page.url})
            return True
        except (ManualCheckpoint,ActionCancelled): raise
        except Exception:
            if optional:return False
            self.snapshot_error("click_"+alias)
            raise

    def fill(self,alias:str,value,optional=False,*,locator=None):
        try:
            raise_if_cancelled()
            loc=locator if locator is not None else self.locator(alias)
            self._log("fill",alias,{"value":"[REDACTED]" if alias in {"otp"} else str(value)[:120]})
            loc.fill(str(value),timeout=5000)
            self._log("fill.ok",alias,{"length":len(str(value))})
            return True
        except ActionCancelled:raise
        except Exception:
            if optional:return False
            self.snapshot_error("fill_"+alias); raise

    def choose_text(self,text:str,optional=False):
        try:
            loc=wait_visible_recheck(self.page,
                lambda:self.page.get_by_text(text,exact=False).first)
            self._log("click.text",text)
            loc.click(timeout=5000); self.page.wait_for_timeout(250); return True
        except ActionCancelled:raise
        except Exception:
            if optional:return False
            self.snapshot_error("text");raise

    def click_content_text(self,texts,optional=False,timeout=2000,retry_click=False):
        """Find content in two bounded checks; retry clicks only for safe tabs."""
        labels=(texts,) if isinstance(texts,str) else tuple(texts)
        def tab_active():
            for label in labels:
                matches=self.page.get_by_text(label,exact=True)
                for index in range(matches.count()):
                    try:
                        candidate=matches.nth(index)
                        if not candidate.is_visible():continue
                        tab=matches.nth(index).locator(
                            "xpath=ancestor::div[contains(@class,'ct-tab')][1]")
                        if (tab.count() and
                            "-active" in (tab.get_attribute("class",timeout=300) or "").split()):
                            return True
                    except Exception:pass
            return False
        try:
            # The active-class check can match a duplicate/ancestor label in the
            # currently rendered panel. Always click the requested visible tab;
            # tab navigation is idempotent, and callers validate the destination.
            last_click_error=None
            for attempt in range(2):
                deadline=time.monotonic()+timeout/1000
                while True:
                    raise_if_cancelled()
                    for label in labels:
                        matches=self.page.get_by_text(label,exact=True)
                        for index in range(matches.count()):
                            candidate=matches.nth(index)
                            try:
                                if not candidate.is_visible():continue
                                box=candidate.bounding_box()
                                # The Doctor Eleks sidebar occupies the left edge.
                                if box and box.get("x",0)<175:continue
                            except ActionCancelled:raise
                            except Exception:continue
                            try:
                                self._log("click.content",label,{"x":box.get("x") if box else None,
                                                                 "attempt":attempt+1})
                                candidate.click(timeout=timeout)
                                self.page.wait_for_timeout(300)
                                if retry_click:
                                    ready_by=time.monotonic()+2
                                    while time.monotonic()<ready_by:
                                        if tab_active():return True
                                        self.page.wait_for_timeout(200)
                                    last_click_error=AutomationError(
                                        f"Вкладка «{label}» не стала активною після кліку.")
                                    break
                                return True
                            except ActionCancelled:raise
                            except Exception as exc:
                                last_click_error=exc
                                if not retry_click:
                                    raise AutomationError(
                                      f"Елемент «{label}» знайдено, але клік не підтверджено.") from exc
                                break
                        if last_click_error:break
                    if last_click_error or time.monotonic()>=deadline:break
                    self.page.wait_for_timeout(200)
                if attempt==0:
                    if retry_click and tab_active():return True
                    self._log("content.recheck",", ".join(labels),
                              {"reason":"click_failed" if last_click_error else "not_visible"})
                    self.page.wait_for_timeout(2000)
                    last_click_error=None
            raise AutomationError(f"Не знайдено елемент у картці пацієнта після повторної перевірки: {', '.join(labels)}")
        except ActionCancelled:raise
        except Exception:
            if optional:return False
            self.snapshot_error("content_text")
            raise

    def current_text(self)->str:
        try:
            raise_if_cancelled()
            return self.page.locator("body").inner_text(timeout=3000)
        except ActionCancelled:raise
        except:return ""

    def require_authenticated(self):
        if "/login" in self.page.url.casefold():
            raise ManualCheckpoint("login","Сесія Doctor Eleks потребує повторної авторизації",
                                   {"url":self.page.url})
        if self.confirmed_authenticated:
            if not self._auth_confirmation_logged:
                log_event("browser.auth.operator_confirmed",
                          "Оператор підтвердив авторизовану вкладку Doctor Eleks; продовжуємо браузерну автоматизацію.",
                          action_id=self.action_id,data={"url":self.page.url})
                self._auth_confirmation_logged=True
            return
        body=self.current_text().casefold()
        anchors=["медичні записи","введіть дані для пошуку","пацієнт","взаємодії"]
        if any(a in body for a in anchors):return
        raise ManualCheckpoint("login","Авторизуйтеся у відкритому Doctor Eleks/eHealth, після чого повторіть запуск функції.",
                               {"url":self.page.url})

    def _medical_records_ready(self):
        """Identify the patient workspace across current and legacy UI layouts."""
        try:
            if "action=spzibwebpatientgenerallist" not in self.page.url.casefold():
                return False
            active=self.page.locator("li.side-list-item.active")
            if not any("медичні записи" in normalize_patient_name(active.nth(i).inner_text(timeout=300))
                       for i in range(active.count())):
                return False
            # The current Doctor Eleks build uses a single DevExtreme
            # autocomplete named patientSearch. Older builds expose the
            # patient's three name fields separately.
            current_search=self.page.locator('input[name="patientSearch"]')
            if current_search.count() and current_search.first.is_visible():
                return True
            return all(self.page.locator(f'input[name="{field}"]').first.is_visible()
                       for field in ("patientLastName","patientFirstName","patientSecondName"))
        except Exception:
            return False

    def goto_medical_records(self):
        self.require_authenticated()
        if self._medical_records_ready():return
        if "action=spzibwebpatientgenerallist" in self.page.url.casefold():
            deadline=time.monotonic()+5
            while time.monotonic()<deadline:
                raise_if_cancelled()
                if self._medical_records_ready():return
                self.page.wait_for_timeout(250)
        # Reacquire the sidebar on every attempt: Doctor Eleks replaces it
        # during redirects, and the queue also has a patient-search input.
        for attempt in range(2):
            raise_if_cancelled()
            items=self.page.locator("li.side-list-item")
            for index in range(items.count()):
                item=items.nth(index)
                try:
                    if not item.is_visible():continue
                    label=normalize_patient_name(item.inner_text(timeout=300))
                    if "медичні записи" not in label and "медицинские записи" not in label:
                        continue
                    icon=item.locator(".side-icon").first
                    icon.wait_for(state="visible",timeout=2000)
                    self._log("click.medical_records_sidebar","Медичні записи",{"attempt":attempt+1})
                    icon.click(timeout=3000)
                    break
                except ActionCancelled:raise
                except Exception:continue
            deadline=time.monotonic()+12
            while time.monotonic()<deadline:
                raise_if_cancelled()
                if self._medical_records_ready():
                    # A late queue redirect can replace the workspace shortly
                    # after the first render, so verify it remains selected.
                    self.page.wait_for_timeout(1000)
                    if self._medical_records_ready():
                        self._log("medical_records.ready","Медичні записи",{"attempt":attempt+1})
                        return
                self.page.wait_for_timeout(250)
        raise ManualCheckpoint("medical_records_navigation",
              "Не вдалося відкрити сторінку «Медичні записи». Перевірте поточну сторінку Doctor Eleks.",
              {"url":self.page.url})

    def _click_patient_search_button(self,search_input):
        """Click the Search button nearest to the top patient-search input."""
        input_box=search_input.bounding_box()
        choices=[]
        for attempt in range(2):
            choices=[]
            for label in ("Пошук","Поиск"):
                buttons=self.page.get_by_role("button",name=label,exact=True)
                for index in range(buttons.count()):
                    button=buttons.nth(index)
                    try:
                        if not button.is_visible():continue
                        box=button.bounding_box()
                        if not box:continue
                        dy=abs((box["y"]+box["height"]/2)-(input_box["y"]+input_box["height"]/2)) if input_box else box["y"]
                        dx=abs(box["x"]-(input_box["x"]+input_box["width"])) if input_box else box["x"]
                        choices.append((dy,dx,button,box,label))
                    except Exception:pass
            if choices:break
            if attempt==0:self.page.wait_for_timeout(2000)
        if not choices:raise AutomationError("Не знайдено кнопку «Пошук» біля верхнього поля пошуку пацієнта.")
        _,_,button,box,label=min(choices,key=lambda item:(item[0],item[1]))
        self._log("click.patient_search",label,{"x":box["x"],"y":box["y"]})
        button.click(timeout=2000)
        self.page.wait_for_timeout(300)

    def _patient_card_birth_matches(self,dob)->bool:
        expected=_coerce_birth_date(dob)
        if expected is None:return False
        try:
            fields=self.page.locator('input[name="patientBirthDate"]')
            for index in range(fields.count()):
                field=fields.nth(index)
                if not field.is_visible():continue
                actual=_coerce_birth_date(field.input_value(timeout=300))
                if actual is not None:return actual==expected
        except Exception:pass
        return False

    def _patient_card_matches(self,name:str,dob=None)->bool:
        wanted=normalize_patient_name(name).split()[:3]
        if len(wanted)<2:return False
        # Read the actual patient-card fields.  Looking through every input on
        # the page also sees the search box and can confuse the dropdown with
        # an opened card.
        values=[]
        for field_name in ("patientLastName","patientFirstName","patientSecondName"):
            try:
                fields=self.page.locator(f'input[name="{field_name}"]')
                field=next((fields.nth(i) for i in range(fields.count())
                            if fields.nth(i).is_visible()),None)
                if field is None:raise LookupError(field_name)
                value=normalize_patient_name(field.input_value(timeout=200)).strip(" .")
                values.append(value)
            except Exception:
                values=[]
                break
        if len(values)==3:
            # The patient card can show only first-name and patronymic
            # initials even when the search result matched the full name.
            family_and_first=(values[0]==wanted[0] and
                (values[1]==wanted[1] or (len(values[1])==1 and wanted[1].startswith(values[1]))))
            patronymic_matches=(len(wanted)>2 and (values[2]==wanted[2] or
                (len(values[2])==1 and wanted[2].startswith(values[2]))))
            if family_and_first and patronymic_matches:return True
            # Accept a patronymic spelling variant only when surname, first
            # name, and the exact date of birth identify the same card.
            return family_and_first and dob is not None and self._patient_card_birth_matches(dob)
        values=[]
        inputs=self.page.locator("input")
        for index in range(inputs.count()):
            try:
                field=inputs.nth(index)
                if not field.is_visible():continue
                value=normalize_patient_name(field.input_value(timeout=200))
                if value:values.append(value)
            except Exception:pass
        return bool(values and _row_has_patient_identity(" ".join(values),name))

    def _scan_patient_results(self,name:str,timeout_ms=2000):
        """Wait for the delayed dropdown/table and return verified result rows."""
        deadline=time.monotonic()+timeout_ms/1000
        full_matches=[];surname_matches=[]
        while True:
            raise_if_cancelled()
            full_matches=[];surname_matches=[]
            # Doctor Eleks normally renders the patient as a virtual green
            # button after Search.  Prefer clickable dropdown items and use a
            # table row only as a fallback for another UI version.
            selector_groups=(
                # Current Doctor Eleks result markup.  This is intentionally
                # first: .list_text covers only the green patient row, unlike
                # the surrounding popup which also contains its header.
                "div.list_text",
                # Some builds render the patient's name only in this nested
                # span and put the DOB/age beside it in a separate sibling.
                # Scan the name itself so a visible result is not discarded
                # merely because its parent uses a different class.
                "span.text-1",
                "[role='option'], mat-option, .dropdown-menu button, .dropdown-menu a, .dropdown-menu li, .list-group-item",
                "tr",
                # Last-resort compatibility with older Doctor Eleks builds.
                "button, a, [role='button'], [ng-click], [data-ng-click]",
            )
            for selector in selector_groups:
                candidates=self.page.locator(selector)
                group_full=[];group_surnames=[]
                for i in range(candidates.count()):
                    raise_if_cancelled()
                    try:
                        row=candidates.nth(i)
                        if hasattr(row,"is_visible") and not row.is_visible():continue
                        text=row.inner_text(timeout=300)
                        if _row_has_patient_identity(text,name):group_full.append(row)
                        if _row_has_patient_surname(text,name):group_surnames.append(row)
                    except ActionCancelled:raise
                    except Exception:pass
                if group_full:return group_full,group_surnames
                if group_surnames and not surname_matches:surname_matches=group_surnames
            if time.monotonic()>=deadline:return full_matches,surname_matches
            self.page.wait_for_timeout(200)

    def _open_patient_result(self,row,name:str,match_kind:str,data=None,dob=None):
        raise_if_cancelled()
        try:
            row.wait_for(state="visible",timeout=2000)
        except ActionCancelled:raise
        except Exception as exc:
            self.page.wait_for_timeout(2000)
            matches,_=self._scan_patient_results(name,timeout_ms=2000)
            if len(matches)!=1:
                raise AutomationError(
                    "Результат пацієнта зник під час оновлення списку; повторна перевірка не дала одного збігу.") from exc
            row=matches[0]
            row.wait_for(state="visible",timeout=2000)
        self._log(f"select_patient.{match_kind}",name,data or {})
        before_url=self.page.url
        def exact_name_target(container):
            """Return the name span, never the popup/header around it."""
            try:
                spans=container.locator("span.text-1")
                for index in range(spans.count()):
                    span=spans.nth(index)
                    if not span.is_visible():continue
                    text=span.inner_text(timeout=300)
                    if _row_has_patient_identity(text,name):return span
            except Exception:pass
            return container

        def click_actual_target(container,attempt):
            target=exact_name_target(container)
            try:
                target.scroll_into_view_if_needed(timeout=1000)
            except Exception:pass
            try:
                box=target.bounding_box()
                if box:
                    click_x=box["x"]+box["width"]/2
                    click_y=box["y"]+box["height"]/2
                    try:
                        target_info=target.evaluate("node => ({tag:node.tagName,classes:node.className,text:node.innerText})")
                    except Exception:target_info={}
                    self._log("select_patient.click_target",name,{
                        "attempt":attempt,"x":round(click_x,1),"y":round(click_y,1),
                        "width":round(box["width"],1),"height":round(box["height"],1),
                        "target":target_info,
                    })
                    # Click the centre of span.text-1 itself.  The event then
                    # bubbles to the Angular list-item handler without ever
                    # touching the popup title above it.
                    self.page.mouse.click(click_x,click_y)
                    return True
            except Exception:
                pass
            try:
                target.click(timeout=2000)
                return True
            except Exception:
                try:
                    target.click(timeout=2000,force=True)
                    return True
                except Exception:return False
        # One result click can trigger a slow card rerender. A second Search
        # click at this point searches *inside* the patient card and clears the
        # selected result, so wait for the card instead of searching again.
        if click_actual_target(row,1):
            deadline=time.monotonic()+15
            while time.monotonic()<deadline:
                raise_if_cancelled()
                if self._patient_card_matches(name,dob):
                    self._log("select_patient.opened",name,{"verified":True,"attempt":1,
                              "before_url":before_url,"after_url":self.page.url})
                    return
                self.page.wait_for_timeout(200)
        self._log("select_patient.opened",name,{"verified":False,"before_url":before_url,"after_url":self.page.url})
        self.snapshot_error("patient_not_opened")
        raise AutomationError(f"Результат «{name}» знайдено, але його картка не відкрилася після натискання.")

    def select_patient(self,name:str,dob=None):
        set_progress(28,"Шукаємо пацієнта в Doctor Eleks")
        raise_if_cancelled()
        self.goto_medical_records()
        surname=str(name).strip().split()[0]
        search_input=self.locator("patient_search",timeout=2000)
        search_input.fill(surname,timeout=5000)
        actual=search_input.input_value(timeout=1000).strip()
        self._log("fill.patient_search",surname,{"actual":actual})
        if normalize_patient_name(actual)!=normalize_patient_name(surname):
            raise AutomationError(f"У поле пошуку не вдалося ввести прізвище «{surname}».")
        # The Doctor Eleks search widget can lose a query when Search is
        # clicked immediately after fill.  Let its client-side state settle.
        self.page.wait_for_timeout(2000)
        self._click_patient_search_button(search_input)
        # Results may be a table or a delayed dropdown.  Wait up to two
        # seconds instead of treating the first empty millisecond as failure.
        full_matches,surname_matches=self._scan_patient_results(name,timeout_ms=2500)
        if not full_matches and not surname_matches:
            self.page.wait_for_timeout(2000)
            full_matches,surname_matches=self._scan_patient_results(name,timeout_ms=2500)
        if len(full_matches)==1:
            self._open_patient_result(full_matches[0],name,"identity",{"search_term":surname},dob=dob)
            return
        if len(full_matches)>1:
            if dob:
                dated=[]
                for row in full_matches:
                    raise_if_cancelled()
                    try:
                        if _row_matches_birth_date(row.inner_text(timeout=300),dob):dated.append(row)
                    except ActionCancelled:raise
                    except Exception:pass
                if len(dated)==1:
                    self._open_patient_result(dated[0],name,"birth_date",{"search_term":surname},dob=dob)
                    return
            raise ManualCheckpoint("ambiguous_patient",
              f"Знайдено {len(full_matches)} пацієнтів з прізвищем та ініціалами «{name}». Відкрийте правильну картку вручну.",
              {"matches":len(full_matches),"search_term":surname})
        if surname_matches and dob:
            # If the UI formats initials differently (for example Latin and
            # Cyrillic lookalikes), DOB can still identify one exact patient.
            dated=[]
            for row in surname_matches:
                raise_if_cancelled()
                try:
                    if _row_matches_birth_date(row.inner_text(timeout=300),dob):dated.append(row)
                except ActionCancelled:raise
                except Exception:pass
            if len(dated)==1:
                self._open_patient_result(dated[0],name,"birth_date",
                                          {"search_term":surname,"dob_verified":True},dob=dob)
                return
        if not surname_matches:
            raise AutomationError(f"Пацієнта з прізвищем «{surname}» не знайдено у результатах пошуку.")
        raise ManualCheckpoint("ambiguous_patient",
          f"Знайдено результати за прізвищем «{surname}», але прізвище та ініціали не збігаються з «{name}». Відкрийте правильного пацієнта вручну і повторіть функцію.",
          {"matches":len(surname_matches),"search_term":surname})

    def maybe_wait_kep(self,message="Накладіть КЕП у відкритому вікні, дочекайтеся завершення і повторіть функцію."):
        raise ManualCheckpoint("kep",message)

    def page_has(self,text:str)->bool:
        try:return self.page.get_by_text(text,exact=False).count()>0
        except:return False
