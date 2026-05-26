# -*- coding: utf-8 -*-
r"""
Полный скрипт + ДОБАВЛЕНО:
✅ Если браузер (Chrome/WebDriver) "падает" / закрывается / теряет сессию —
   скрипт автоматически перезапускает браузер и продолжает работу.
✅ Кол-во попыток перезапуска браузера: не меньше 20 (по умолчанию 20).
✅ Прогресс сохраняется в checkpoint-файл (JSON) в папке BASE_DIR, чтобы продолжать
   с последней необработанной строки Excel даже после падения.

ВАЖНО:
- Я НЕ менял вашу бизнес-логику по получению данных, ретраям, 2-му проходу.
- Я добавил "наружный контур" устойчивости: restart loop + checkpoint + resume.

Примечание: этот файл — перенесённое без изменений содержимое прежнего
scripts/sud.py (подача иска в office.sud.kz + добавление участника +
автоподстановка ФИО/адреса по ИИН). Используется пунктом интерфейса
"⚖️ Судебный кабинет - ФИО, адрес" (ключ sudFIO). scripts/sud.py теперь
содержит отдельную новую логику — выгрузку статусов дел с СК по талонам.
"""

import os
import time
import json
import logging
import datetime as _dt
import re as _re  # на будущее

from openpyxl import load_workbook, Workbook

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    StaleElementReferenceException,
    ElementNotInteractableException,
    WebDriverException,
    InvalidSessionIdException,
    NoSuchWindowException,
)

# ========= КОНФИГ =========
USER_AUTH     = "230240016634"        # твой ИИН/БИН
USER_PASSWORD = "n3y&pAM5mD&4zKZ"     # твой пароль

BASE      = "https://office.sud.kz"
LOGIN_URL = f"{BASE}/index.xhtml"
HOME_URL  = f"{BASE}/form/proceedings/services.xhtml"
SEND_DOCS_URL = f"{BASE}/form/send/index.xhtml"
SEARCH_FALLBACK_URL = f"{BASE}/lawsuit/document.xhtml"   # пока не используем

# Паузы
WAIT   = 1.0
RETRY  = 8

# ========= ПУТИ ДЛЯ ПАРСИНГА ИИН =========
BASE_DIR    = r"C:\Users\User\АИС ОИП1\Cуд_ФИО_адреса"
INPUT_XLSX  = os.path.join(BASE_DIR, "СК_ИИН.xlsx")      # при необходимости исправь имя
OUTPUT_DIR  = BASE_DIR

# ✅ НОВОЕ: автоперезапуск браузера
MAX_BROWSER_RESTARTS = 20

# ✅ НОВОЕ: checkpoint-файл для продолжения
CHECKPOINT_PATH = os.path.join(BASE_DIR, "sud_checkpoint.json")

def make_output_filename() -> str:
    # дд.мм.гггг-чч.мм
    return _dt.datetime.now().strftime("%d.%m.%Y-%H.%M") + ".xlsx"

# ✅ НОВОЕ: фиксируем имя output 1 раз на весь запуск (и на рестарты)
RUN_TS = _dt.datetime.now().strftime("%d.%m.%Y-%H.%M")
OUT_XLSX = os.path.join(OUTPUT_DIR, RUN_TS + ".xlsx")

os.makedirs(BASE_DIR, exist_ok=True)

# ========= ЛОГИ =========
logging.basicConfig(
    filename="sud_script.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8",
)

def log_step(msg):
    print(msg)
    logging.info(msg)

def log_warn(msg):
    print("WARN:", msg)
    logging.warning(msg)

def log_err(msg):
    print("ERROR:", msg)
    logging.error(msg)

# короткий лог с временем для блока 4
def log4(msg):
    print(time.strftime("[%H:%M:%S]"), msg, flush=True)
    logging.info(msg)

# ========= НОВОЕ: ошибки/утилиты для рестарта =========
class BrowserCrashed(RuntimeError):
    pass

def is_browser_dead(exc: Exception) -> bool:
    """
    Пытаемся распознать ситуации, когда WebDriver/Chrome "умер" и нужен рестарт.
    """
    if isinstance(exc, (InvalidSessionIdException, NoSuchWindowException)):
        return True
    if isinstance(exc, WebDriverException):
        msg = (str(exc) or "").lower()
        # типовые признаки смерти сессии/браузера
        bad_markers = [
            "disconnected",
            "not reachable",
            "chrome not reachable",
            "invalid session id",
            "session deleted",
            "no such window",
            "target window already closed",
            "cannot determine loading status",
            "unknown error: cannot find",
            "failed to check if window was closed",
        ]
        return any(m in msg for m in bad_markers)
    return False

def safe_quit(driver):
    try:
        driver.quit()
    except Exception:
        pass

# ========= DRIVER =========
def init_driver() -> webdriver.Chrome:
    opts = webdriver.ChromeOptions()
    opts.add_argument("--start-maximized")
    opts.page_load_strategy = "eager"

    prefs = {"profile.managed_default_content_settings.images": 2}
    opts.add_experimental_option("prefs", prefs)
    opts.add_argument("--blink-settings=imagesEnabled=false")
    opts.add_argument("--disable-extensions")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--disable-background-networking")
    opts.add_argument("--disable-renderer-backgrounding")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=opts)
    return driver

# ========= ВСПОМОГАТЕЛЬНЫЕ =========
def find_with_fallbacks(driver, variants, desc, tries=8, delay=0.4):
    """Пробует несколько локаторов с коротким ретраем."""
    last = None
    for _ in range(tries):
        for by, sel in variants:
            try:
                el = driver.find_element(by, sel)
                logging.info(f"Нашёл {desc} по {by}={sel}")
                return el
            except Exception as e:
                last = e
        time.sleep(delay)
    raise NoSuchElementException(f"Не удалось найти {desc}. Последняя ошибка: {last}")

# ========= LOGIN =========
def login(driver):
    log_step("Открываю страницу логина…")
    driver.get(LOGIN_URL)
    time.sleep(WAIT)

    # язык — РУС
    try:
        ru = driver.find_elements(By.XPATH, "//a[contains(.,'РУС') and not(contains(@class,'active'))]")
        if ru:
            ru[0].click()
            log_step("Переключил язык на РУС")
            time.sleep(WAIT)
    except Exception:
        log_warn("Не удалось переключить язык на РУС")

    # ЛОГИН
    login_el = find_with_fallbacks(
        driver,
        variants=[
            (By.ID, "j_idt78:auth:xin"),
            (By.XPATH, "//input[contains(@id,':auth:xin')]"),
            (By.NAME, "j_idt78:auth:xin"),
            (By.XPATH, "//input[contains(@placeholder,'ИИН') or contains(@placeholder,'ЖСН') or contains(@placeholder,'ИИН/БИН')]"),
            (By.CSS_SELECTOR, "input[type='email']"),
        ],
        desc="поле ИИН/БИН"
    )
    pass_el = find_with_fallbacks(
        driver,
        variants=[
            (By.ID, "j_idt78:auth:password"),
            (By.XPATH, "//input[contains(@id,':auth:password')]"),
            (By.XPATH, "//input[@type='password']"),
            (By.XPATH, "//input[contains(@placeholder,'Пароль') or contains(@placeholder,'Құпия')]"),
        ],
        desc="поле Пароль"
    )
    submit = find_with_fallbacks(
        driver,
        variants=[
            (By.CSS_SELECTOR, "input.button-primary[type='submit']"),
            (By.XPATH, "//input[@type='submit' and (contains(@value,'Войти') or contains(@class,'button-primary'))]"),
            (By.XPATH, "//button[contains(.,'Войти')]"),
        ],
        desc="кнопка Войти"
    )

    login_el.clear(); login_el.send_keys(USER_AUTH)
    pass_el.clear();  pass_el.send_keys(USER_PASSWORD)
    log_step("Ввёл ИИН/БИН и Пароль")
    submit.click()
    log_step("Нажал 'Войти'")
    time.sleep(WAIT * 2)

# ========= ПЕРЕХОД В «Подача документов» =========
def go_to_send_docs(driver):
    log_step("Открываю главную страницу (плитки)…")
    driver.get(HOME_URL)
    time.sleep(WAIT)

    link = None
    try:
        link = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, '/form/send/index.xhtml')]"))
        )
    except TimeoutException:
        link = None

    if not link:
        try:
            link = driver.find_element(By.XPATH, "//p[normalize-space()='Подача документов']/ancestor::a")
        except Exception:
            try:
                link = driver.find_element(By.XPATH, "//a[.//p[contains(normalize-space(),'Подача документ')]]")
            except Exception:
                link = None

    if link:
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", link)
        except Exception:
            pass
        try:
            link.click()
        except Exception:
            driver.execute_script("arguments[0].click();", link)
        log_step("Кликнул по плитке 'Подача документов'")
    else:
        log_warn("Плитка не найдена, открываю 'Подача документов' по прямому URL")
        driver.get(SEND_DOCS_URL)

    loaded = False
    try:
        WebDriverWait(driver, 10).until(EC.url_contains("/form/send"))
        loaded = True
    except TimeoutException:
        pass

    if not loaded:
        try:
            WebDriverWait(driver, 6).until(
                EC.presence_of_element_located((
                    By.XPATH,
                    "//h1[contains(.,'Подача документов')] | "
                    "//h2[contains(.,'Подача документов')] | "
                    "//form[contains(@action,'/form/send')]"
                ))
            )
            loaded = True
        except TimeoutException:
            pass

    if not loaded:
        log_warn("Не удалось явно подтвердить загрузку страницы 'Подача документов' — продолжаю по текущему состоянию")

# ========= БЛОК 2: выбор CIVIL/FIRSTINSTANCE/Иск =========
def _css_from_locator(locator):
    by, sel = locator
    if by == By.CSS_SELECTOR:
        return sel
    if by == By.ID:
        css_id = sel.replace(":", r"\:")
        return f"#{css_id}"
    return sel

def _select_by_value_robust(driver, locator, value, desc, timeout=20, tries=3):
    w = WebDriverWait(driver, timeout)

    # ждём <select>
    w.until(EC.presence_of_element_located(locator))

    # ждём наличия нужной опции
    def _opt_present(d):
        try:
            el = d.find_element(*locator)
            return any((o.get_attribute("value") or "") == value
                       for o in el.find_elements(By.TAG_NAME, "option"))
        except StaleElementReferenceException:
            return False
    w.until(_opt_present)

    last_err = None
    for _ in range(tries):
        try:
            el = driver.find_element(*locator)
            Select(el).select_by_value(value)
            w.until(lambda d: d.find_element(*locator).get_attribute("value") == value)
            print(f"✔ {desc} → {value}")
            return
        except (StaleElementReferenceException, TimeoutException) as e:
            last_err = e

    # JS-фолбэк
    ok = driver.execute_script("""
        var el = document.querySelector(arguments[0]);
        if(!el) return false;
        var idx = Array.from(el.options).findIndex(o => (o.value||'') === arguments[1]);
        if(idx < 0) return false;
        el.selectedIndex = idx;
        el.dispatchEvent(new Event('change', {bubbles:true}));
        return true;
    """, _css_from_locator(locator), value)

    if not ok:
        raise TimeoutException(f"Не удалось выбрать '{desc}' со значением {value}. Последняя ошибка: {last_err}")
    WebDriverWait(driver, 10).until(lambda d: d.find_element(*locator).get_attribute("value") == value)
    print(f"✔ {desc} → {value} (через JS)")

def send_claim(driver, timeout=20):
    """
    На странице /form/send/index.xhtml:
      1) Тип производства = CIVIL
      2) Инстанция = FIRSTINSTANCE
      3) Тип документа = Иск (3)
      4) Нажать «Отправить»
    """
    w = WebDriverWait(driver, timeout)

    _select_by_value_robust(
        driver, (By.CSS_SELECTOR, "select[id$=':case-type']"),
        "CIVIL", "Тип производства", timeout=timeout
    )

    _select_by_value_robust(
        driver, (By.CSS_SELECTOR, "select[id$=':instance']"),
        "FIRSTINSTANCE", "Инстанция", timeout=timeout
    )

    _select_by_value_robust(
        driver, (By.CSS_SELECTOR, "select[id$=':request']"),
        "3", "Тип документа", timeout=timeout
    )

    submit_locators = [
        (By.XPATH, "//input[@type='submit' and @value='Отправить']"),
        (By.XPATH, "//input[contains(@id,':j_idt') and @type='submit']"),
        (By.XPATH, "//button[normalize-space()='Отправить']"),
    ]
    btn = None
    for loc in submit_locators:
        try:
            btn = w.until(EC.element_to_be_clickable(loc))
            break
        except TimeoutException:
            continue
    if not btn:
        raise TimeoutException("Кнопка «Отправить» не найдена.")

    try:
        btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", btn)
    print("✔ Нажал «Отправить».")

# ========= БЛОК 3: createRequest + форма участника =========
def _wait(drv, cond, t=25):
    return WebDriverWait(drv, t, ignored_exceptions=(StaleElementReferenceException,)).until(cond)

def _select_value(drv, css, value, desc, t=25):
    sel = _wait(drv, EC.presence_of_element_located((By.CSS_SELECTOR, css)), t)
    for _ in range(3):
        try:
            Select(sel).select_by_value(value)
            _wait(drv, lambda d: d.find_element(By.CSS_SELECTOR, css).get_attribute("value") == value, 10)
            print(f"✔ {desc} → value={value}")
            return
        except StaleElementReferenceException:
            sel = drv.find_element(By.CSS_SELECTOR, css)
        except Exception:
            pass
    # JS-фолбэк
    drv.execute_script("""
        var s = document.querySelector(arguments[0]);
        var val = arguments[1];
        if(!s) return false;
        s.value = val;
        s.dispatchEvent(new Event('change', {bubbles:true}));
        return true;
    """, css, value)
    _wait(drv, lambda d: d.find_element(By.CSS_SELECTOR, css).get_attribute("value") == value, 10)
    print(f"✔ {desc} → value={value} (JS)")

def _click(drv, locator, desc, t=25):
    btn = _wait(drv, EC.element_to_be_clickable(locator), t)
    drv.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    try:
        btn.click()
    except Exception:
        drv.execute_script("arguments[0].click();", btn)
    print(f"✔ Нажал {desc}")

def block3_fill_case_and_open_modal(driver, timeout=40):
    w = WebDriverWait(driver, timeout)

    # убеждаемся, что на createRequest.xhtml
    try:
        w.until(
            lambda d: "/form/requestType2/createRequest.xhtml" in d.current_url
            or d.find_element(By.CSS_SELECTOR, "select[id$=':edit-category']")
        )
    except Exception:
        raise TimeoutException("Страница createRequest.xhtml не открыта")

    # 1) Вид производства по делу: value=2
    _select_value(
        driver,
        "select[id$=':edit-categoryGroup']",
        "2",
        "Вид производства по делу",
        t=timeout,
    )

    # 2) Категория дела: value=27
    _wait(
        driver,
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "select[id$=':edit-category'] option[value='27']")
        ),
        30,
    )
    _select_value(
        driver,
        "select[id$=':edit-category']",
        "27",
        "Категория дела",
        t=timeout,
    )

    # 3) Характер заявления: value=1
    _wait(
        driver,
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "select[id$=':edit-character'] option[value='1']")
        ),
        20,
    )
    _select_value(
        driver,
        "select[id$=':edit-character']",
        "1",
        "Характер заявления",
        t=timeout,
    )

    # 4) Галочка «Дело упрощенного производства»
    try:
        cb = w.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "input[id$=':edit-simpleProcess'][type='checkbox']")
            )
        )
        if not cb.is_selected():
            driver.execute_script("arguments[0].click();", cb)
        print("✔ Галочка «Дело упрощенного производства» установлена")
    except Exception:
        print("⚠️ Не удалось установить галочку «Дело упрощенного производства» (продолжаю)")

    # 5) Открыть модал «Добавить участника процесса»
    add_locators = [
        (By.XPATH, "//button[contains(.,'Добавить участника процесса')]"),
        (By.CSS_SELECTOR, "button[id*='addPerson']"),
    ]
    next_xpath = (
        "//div[contains(@class,'modal') or contains(@class,'modal-dialog')]"
        "//input[@type='button' or @type='submit'][@value='Далее']"
    )
    for loc in add_locators:
        try:
            _click(driver, loc, "«Добавить участника процесса»", t=timeout)
            _wait(driver, EC.presence_of_element_located((By.XPATH, next_xpath)), 20)
            break
        except TimeoutException:
            continue
    else:
        raise TimeoutException("Кнопка «Добавить участника процесса» не найдена")

    # 6) В модалке жмём «Далее» и ждём поле ИИН
    person_iin_css = "input[id$=':person-iin']"

    def modal_switched(d):
        try:
            return d.find_element(By.CSS_SELECTOR, person_iin_css)
        except Exception:
            return False

    for attempt in range(3):
        try:
            btn_next = _wait(
                driver,
                EC.element_to_be_clickable((By.XPATH, next_xpath)),
                15,
            )

            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", btn_next
            )

            try:
                btn_next.click()
            except Exception:
                driver.execute_script("arguments[0].click();", btn_next)

            _wait(driver, modal_switched, 25)
            print("✅ Блок 3 завершён: открыта форма участника с полем ИИН.")
            break

        except StaleElementReferenceException:
            print("⚠ Stale element на кнопке «Далее», пробую ещё раз…")
            if attempt == 2:
                print("❌ Не удалось нажать «Далее» — кнопка всё время устаревает")
                raise
            time.sleep(1)

# ========= БЛОК 4: парсинг ИИН из Excel =========
print("=== BLOCK-4 v5 + browser-restart loaded (СК_ИИН → новый файл с результатами) ===")

AUTOSAVE_EVERY = 25  # каждые N обработанных ИИН автосейв

# профиль по времени
PERSON_TIMEOUT         = 8
RETRIES_PER_PERSON     = 2
PAUSE_BETWEEN_TRIES    = 0.4
COOLDOWN_BETWEEN_ROWS  = 1.1
MIN_SECONDS_PER_ROW    = 1.2
POLL                   = 0.10
MIN_STABLE_SEC         = 0.35
HARD_WAIT_AFTER_SEARCH = 0.20

# локаторы формы «Добавить участника процесса»
IIN  = (By.CSS_SELECTOR, 'input[id$=":person-iin"]')
SUR  = (By.CSS_SELECTOR, 'input[id$=":person-surname"]')
NAM  = (By.CSS_SELECTOR, 'input[id$=":person-firstname"]')
PATR = (By.CSS_SELECTOR, 'input[id$=":person-patronymic"]')
LIVE = (By.CSS_SELECTOR, 'textarea[id$=":person-livePlace"], input[id$=":person-livePlace"]')

# кнопки/диалоги для перезапуска формы
BTN_CLOSE_PERSON = (
    By.XPATH,
    "//input[@value='Закрыть' and contains(@onclick,'hideFizModalDialog')]"
    " | //input[@id='j_idt278:j_idt328']"
)
BTN_ADD_PERSON = (
    By.XPATH,
    "//button[contains(@onclick,'renderAddPersonModalDialog') or contains(.,'Добавить участника процесса')]"
)
BTN_NEXT_SIDE = (
    By.XPATH,
    "//div[contains(@class,'modal')]//input[@type='button' and @value='Далее'] "
    "| //input[@id='j_idt185:j_idt205']"
)

def _norm_iin(val) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    try:
        if "e" in s.lower():
            s = str(int(float(s)))
    except Exception:
        pass
    s = "".join(ch for ch in s if ch.isdigit())
    return s[:12]

def _get_val(drv, loc) -> str:
    try:
        el = drv.find_element(*loc)
        return (el.get_attribute("value") or "").strip()
    except Exception:
        return ""

def _rf_queue_size(drv):
    try:
        return drv.execute_script("""
            try {
              if (window.RichFaces && RichFaces.Queue && typeof RichFaces.Queue.getSize==='function'){
                return RichFaces.Queue.getSize();
              }
            } catch(e){}
            return 0;
        """) or 0
    except Exception:
        return 0

def _wait_queue_quiet(drv, min_quiet=MIN_STABLE_SEC, total=6):
    t_end = time.time() + total
    quiet_start = None
    last = -1
    while time.time() < t_end:
        q = _rf_queue_size(drv)
        if q != last:
            last = q
            log4(f"   ajax queue size = {q}")
        if q == 0:
            if quiet_start is None:
                quiet_start = time.time()
            if time.time() - quiet_start >= min_quiet:
                return True
        else:
            quiet_start = None
        time.sleep(0.15)
    return False

def _values_snapshot(drv):
    return {
        "sur":  _get_val(drv, SUR),
        "name": _get_val(drv, NAM),
        "patr": _get_val(drv, PATR),
        "live": _get_val(drv, LIVE),
    }

def _stable_after_change(drv, before: dict, timeout=PERSON_TIMEOUT):
    w = WebDriverWait(
        drv,
        timeout,
        poll_frequency=POLL,
        ignored_exceptions=(StaleElementReferenceException,),
    )

    def _changed(_):
        now = _values_snapshot(drv)
        changed = any(now[k] != before.get(k, "") for k in now) and any(now.values())
        return now if changed else False

    try:
        changed_vals = w.until(_changed)
        t0 = time.time()
        last = changed_vals
        while time.time() - t0 < MIN_STABLE_SEC:
            time.sleep(0.15)
            now = _values_snapshot(drv)
            if now != last:
                t0 = time.time()
                last = now
        return last
    except TimeoutException:
        return _values_snapshot(drv)

def _click_magnifier(drv):
    # сначала пробуем js-функцию
    try:
        drv.execute_script("if (typeof fillPersonData==='function'){fillPersonData('j_idt278:person-iin');}")
        log4("   вызвал fillPersonData()")
        return True
    except Exception:
        pass

    # потом ищем иконку-лупу
    try:
        span = drv.find_element(
            By.XPATH,
            "//input[contains(@id,':person-iin')]/following-sibling::span[contains(@class,'gbdSearch')]",
        )
        drv.execute_script("arguments[0].scrollIntoView({block:'center'});", span)
        span.click()
        log4("   клик по иконке-лупе")
        return True
    except Exception:
        try:
            span = drv.find_element(By.CSS_SELECTOR, "span.gbdSearch")
            drv.execute_script("arguments[0].scrollIntoView({block:'center'});", span)
            span.click()
            log4("   клик по .gbdSearch (fallback)")
            return True
        except Exception:
            log4("   ⚠ не смог нажать лупу")
            return False

def _trigger_and_wait(drv) -> dict:
    before = _values_snapshot(drv)
    for attempt in range(1, RETRIES_PER_PERSON + 1):
        log4(f"   попытка #{attempt} — запуск поиска")
        _click_magnifier(drv)
        time.sleep(HARD_WAIT_AFTER_SEARCH)
        _wait_queue_quiet(drv, min_quiet=MIN_STABLE_SEC, total=PERSON_TIMEOUT)
        after = _stable_after_change(drv, before, timeout=PERSON_TIMEOUT)
        changed = any(after[k] != before.get(k, "") for k in after) and any(after.values())
        log4(f"   снэпшот: {after}  | changed={changed}")
        if changed:
            return after
        log4("   ↺ не изменилось — подожду и попробую ещё раз")
        time.sleep(PAUSE_BETWEEN_TRIES)
    return {"sur": "", "name": "", "patr": "", "live": ""}

def _safe_click(drv, loc, desc=""):
    try:
        el = WebDriverWait(drv, 6).until(EC.element_to_be_clickable(loc))
    except TimeoutException:
        try:
            el = drv.find_element(*loc)
        except Exception:
            return False
    try:
        drv.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        el.click()
        return True
    except Exception:
        try:
            drv.execute_script("arguments[0].click();", el)
            return True
        except Exception:
            return False

def _reopen_person_modal(drv):
    """
    Закрыть форму участника, снова нажать «Добавить участника процесса»
    и (если нужно) «Далее» в выборе стороны.
    """
    log4("   ↻ перезапуск формы участника…")

    _safe_click(drv, BTN_CLOSE_PERSON, "Закрыть")
    time.sleep(0.7)

    if not _safe_click(drv, BTN_ADD_PERSON, "Добавить участника"):
        log4("   ⚠ не нашёл кнопку «Добавить участника процесса»")
        return False

    _safe_click(drv, BTN_NEXT_SIDE, "Далее")

    try:
        WebDriverWait(drv, 12).until(EC.presence_of_element_located(IIN))
        log4("   форма участника открыта заново")
        return True
    except TimeoutException:
        log4("   ⚠ не дождался поля ИИН после перезапуска")
        return False

def ensure_person_form_open(drv, timeout=15):
    """
    Гарантируем, что открыта форма участника с полем ИИН.
    Если поля нет — пробуем заново открыть модалку.
    """
    try:
        return WebDriverWait(drv, timeout).until(
            EC.presence_of_element_located(IIN)
        )
    except TimeoutException:
        if not _reopen_person_modal(drv):
            raise RuntimeError(
                "Не удалось найти форму участника (поле ИИН). "
                "Возможно, заявка закрылась или сессия слетела."
            )
        return WebDriverWait(drv, timeout).until(
            EC.presence_of_element_located(IIN)
        )

def _safe_save(wb, path):
    try:
        wb.save(path)
        log4(f"💾 autosave to {path}")
        return True
    except Exception as e:
        log4(f"⚠ не удалось сохранить файл: {e}")
        return False

def _split_country_region(live: str):
    country = ""
    region = ""
    if live:
        parts = [p.strip() for p in live.split(",") if p.strip()]
        if parts:
            country = parts[0]
        if len(parts) >= 2:
            region = parts[1]
    return country, region

# ========= НОВОЕ: checkpoint state =========
def load_checkpoint():
    if not os.path.exists(CHECKPOINT_PATH):
        return None
    try:
        with open(CHECKPOINT_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        log_warn(f"Не удалось прочитать checkpoint: {e}")
        return None

def save_checkpoint(state: dict):
    try:
        with open(CHECKPOINT_PATH, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log_warn(f"Не удалось сохранить checkpoint: {e}")

def clear_checkpoint():
    try:
        if os.path.exists(CHECKPOINT_PATH):
            os.remove(CHECKPOINT_PATH)
    except Exception:
        pass

def init_result_file_if_needed(out_path: str):
    """
    Если out_path не существует — создаём новый файл с заголовками.
    Если существует — ничего не делаем (будем продолжать).
    """
    if os.path.exists(out_path):
        return

    wb = Workbook()
    ws = wb.active
    headers = [
        "порядковый №",
        "ИНН",
        "Фамилия",
        "Имя",
        "Отчество",
        "ФИО",
        "Место жительства",
        "Страна",
        "Регион",
        "Статус",
    ]
    for idx, name in enumerate(headers, start=1):
        ws.cell(row=1, column=idx, value=name)
    wb.save(out_path)

def append_result_row(out_path: str, row_idx: int, src_iin: str, sur: str, name: str, patr: str, fio: str, live: str, country: str, region: str, status: str):
    """
    Пишем/перезаписываем конкретную строку row_idx (в файле результата).
    """
    wb = load_workbook(out_path)
    ws = wb.active

    ws.cell(row=row_idx, column=1, value=row_idx - 1)  # порядковый №
    ws.cell(row=row_idx, column=2, value=src_iin)
    ws.cell(row=row_idx, column=3, value=sur)
    ws.cell(row=row_idx, column=4, value=name)
    ws.cell(row=row_idx, column=5, value=patr)
    ws.cell(row=row_idx, column=6, value=fio)
    ws.cell(row=row_idx, column=7, value=live)
    ws.cell(row=row_idx, column=8, value=country)
    ws.cell(row=row_idx, column=9, value=region)
    ws.cell(row=row_idx, column=10, value=status)

    wb.save(out_path)

def get_next_output_row(out_path: str) -> int:
    wb = load_workbook(out_path)
    ws = wb.active
    return ws.max_row + 1  # следующая свободная

def parse_people_from_excel_and_save_v2_resumable(drv, out_path: str):
    """
    Ваша функция блока 4, но с продолжением по checkpoint:
    - next_src_row: с какой строки исходного Excel продолжать
    - failed_rows: проблемные строки (для 2-го прохода)
    - row_map: mapping src_row -> result_row (для перезаписи во 2-м проходе)
    """

    MAX_RESTARTS_PER_ROW = 2

    init_result_file_if_needed(out_path)

    state = load_checkpoint()
    if not state:
        state = {
            "input_xlsx": INPUT_XLSX,
            "output_xlsx": out_path,
            "next_src_row": 2,
            "failed_rows": [],
            "row_map": {},     # key=src_row (str), value=result_row (int)
            "processed_count": 0,
            "phase": "FIRST_PASS",
        }
        save_checkpoint(state)
        log4("Checkpoint не найден — стартую с начала (row 2).")
    else:
        # минимальная валидация
        if state.get("input_xlsx") != INPUT_XLSX:
            log_warn("Checkpoint от другого INPUT_XLSX — игнорирую и начинаю заново.")
            clear_checkpoint()
            return parse_people_from_excel_and_save_v2_resumable(drv, out_path)

        log4(f"Resume из checkpoint: phase={state.get('phase')}, next_src_row={state.get('next_src_row')}")

    # Загружаем исходный файл
    src_wb = load_workbook(INPUT_XLSX, data_only=True)
    src_ws = src_wb.active

    # убеждаемся, что форма участника открыта
    WebDriverWait(drv, 20).until(EC.presence_of_element_located(IIN))

    # --- local vars ---
    processed = int(state.get("processed_count") or 0)
    failed_rows = list(state.get("failed_rows") or [])
    row_map = dict(state.get("row_map") or {})
    next_src_row = int(state.get("next_src_row") or 2)
    phase = state.get("phase") or "FIRST_PASS"

    def checkpoint_update(**kwargs):
        state.update(kwargs)
        save_checkpoint(state)

    try:
        # ---------- ПЕРВЫЙ ПРОХОД ----------
        if phase == "FIRST_PASS":
            out_row = get_next_output_row(out_path)

            for r in range(next_src_row, src_ws.max_row + 1):
                try:
                    # ❗ ВАЖНО: вы в описании писали "колонка D", но в вашем коде стоит column=1.
                    # Я ОСТАВИЛ КАК В ВАШЕМ КОДЕ: column=1.
                    src_iin = _norm_iin(src_ws.cell(row=r, column=1).value)
                    if not src_iin or len(src_iin) != 12:
                        checkpoint_update(next_src_row=r + 1)
                        continue

                    processed += 1
                    t_row_start = time.time()
                    log4(f"#{processed} (Excel row {r}) → ИИН: {src_iin}")

                    success = False
                    data_for_row = {"sur": "", "name": "", "patr": "", "live": ""}

                    for restart in range(MAX_RESTARTS_PER_ROW + 1):
                        if restart > 0:
                            log4(f"   ↻ полный перезапуск формы для этого ИИН (рестарт {restart})")
                            if not _reopen_person_modal(drv):
                                log4("   ❌ не удалось перезапустить форму, выхожу из попыток по этому ИИН")
                                break

                        iin_el = ensure_person_form_open(drv, timeout=20)
                        drv.execute_script("arguments[0].scrollIntoView({block:'center'});", iin_el)

                        # очищаем поле
                        try:
                            iin_el.clear()
                        except Exception:
                            try:
                                drv.execute_script("arguments[0].value='';", iin_el)
                            except Exception:
                                pass

                        # вводим ИИН
                        try:
                            iin_el.send_keys(src_iin)
                        except ElementNotInteractableException:
                            log4("   ⚠ элемент ИИН не интерактивен — ставлю значение через JS")
                            drv.execute_script("arguments[0].value = arguments[1];", iin_el, src_iin)

                        # триггеры change/blur
                        try:
                            drv.execute_script(
                                "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
                                iin_el,
                            )
                            drv.execute_script(
                                "arguments[0].dispatchEvent(new Event('blur',{bubbles:true}));",
                                iin_el,
                            )
                        except Exception:
                            pass

                        before = _values_snapshot(drv)
                        data = _trigger_and_wait(drv)
                        changed = any(data.values()) and any(data[k] != before.get(k, "") for k in data)
                        log4(f"   changed={changed}, data={data}")

                        if changed:
                            data_for_row = data
                            success = True
                            break
                        else:
                            log4("   данных нет / не обновились, пробую перезапустить форму…")

                    sur  = (data_for_row.get("sur") or "").strip()
                    name = (data_for_row.get("name") or "").strip()
                    patr = (data_for_row.get("patr") or "").strip()
                    live = (data_for_row.get("live") or "").strip()
                    fio  = " ".join(x for x in [sur, name, patr] if x)

                    country, region = _split_country_region(live)
                    status = "OK" if success and live else "ERROR_NO_DATA"

                    # пишем в результат (append)
                    append_result_row(out_path, out_row, src_iin, sur, name, patr, fio, live, country, region, status)
                    row_map[str(r)] = out_row
                    log4(f"   → RESULT row {out_row}: status={status}, live='{live}'")
                    out_row += 1

                    if not success:
                        failed_rows.append(r)

                    # checkpoint обновляем ПОСЛЕ строки
                    checkpoint_update(
                        next_src_row=r + 1,
                        failed_rows=failed_rows,
                        row_map=row_map,
                        processed_count=processed,
                        phase="FIRST_PASS",
                    )

                    _wait_queue_quiet(drv, min_quiet=MIN_STABLE_SEC, total=4)
                    time.sleep(COOLDOWN_BETWEEN_ROWS)
                    elapsed = time.time() - t_row_start
                    if elapsed < MIN_SECONDS_PER_ROW:
                        time.sleep(MIN_SECONDS_PER_ROW - elapsed)

                    if processed % AUTOSAVE_EVERY == 0:
                        log4("autosave checkpoint OK")

                except Exception as e:
                    # если умер браузер — прокидываем наверх для рестарта
                    if is_browser_dead(e):
                        raise BrowserCrashed(f"Browser dead during FIRST_PASS at src row {r}: {e}") from e
                    # иначе считаем как сбой строки (но продолжаем)
                    log4(f"⚠ Ошибка на строке {r}: {e}")
                    failed_rows.append(r)
                    checkpoint_update(
                        next_src_row=r + 1,
                        failed_rows=failed_rows,
                        row_map=row_map,
                        processed_count=processed,
                        phase="FIRST_PASS",
                    )

            log4(f"✅ Первый проход завершён. out={out_path}")
            checkpoint_update(
                next_src_row=2,
                failed_rows=failed_rows,
                row_map=row_map,
                processed_count=processed,
                phase="SECOND_PASS",
            )
            phase = "SECOND_PASS"

        # ---------- ВТОРОРОЙ ПРОХОД ----------
        if phase == "SECOND_PASS":
            failed_rows = list(state.get("failed_rows") or [])
            row_map = dict(state.get("row_map") or {})

            if failed_rows:
                log4(f"=== SECOND PASS: найдено {len(failed_rows)} строк с ERROR_NO_DATA, пробуем ещё раз ===")
                # заново открываем форму участника
                go_to_send_docs(drv)
                send_claim(drv)
                block3_fill_case_and_open_modal(drv)
                WebDriverWait(drv, 20).until(EC.presence_of_element_located(IIN))

                for r in failed_rows:
                    try:
                        src_iin = _norm_iin(src_ws.cell(row=r, column=1).value)
                        if not src_iin:
                            continue

                        log4(f"[2ND PASS] row {r} → ИИН: {src_iin}")

                        iin_el = ensure_person_form_open(drv, timeout=20)
                        drv.execute_script("arguments[0].scrollIntoView({block:'center'});", iin_el)

                        try:
                            iin_el.clear()
                        except Exception:
                            try:
                                drv.execute_script("arguments[0].value='';", iin_el)
                            except Exception:
                                pass

                        try:
                            iin_el.send_keys(src_iin)
                        except ElementNotInteractableException:
                            log4("   ⚠ [2ND PASS] элемент ИИН не интерактивен — ставлю значение через JS")
                            drv.execute_script("arguments[0].value = arguments[1];", iin_el, src_iin)

                        try:
                            drv.execute_script(
                                "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
                                iin_el,
                            )
                            drv.execute_script(
                                "arguments[0].dispatchEvent(new Event('blur',{bubbles:true}));",
                                iin_el,
                            )
                        except Exception:
                            pass

                        before = _values_snapshot(drv)
                        data = _trigger_and_wait(drv)
                        changed = any(data.values()) and any(data[k] != before.get(k, "") for k in data)
                        log4(f"   [2ND PASS] changed={changed}, data={data}")

                        if not changed:
                            log4("   [2ND PASS] снова нет данных, оставляю статус как есть")
                            _wait_queue_quiet(drv, min_quiet=MIN_STABLE_SEC, total=4)
                            time.sleep(COOLDOWN_BETWEEN_ROWS)
                            continue

                        sur  = (data.get("sur") or "").strip()
                        name = (data.get("name") or "").strip()
                        patr = (data.get("patr") or "").strip()
                        live = (data.get("live") or "").strip()
                        fio  = " ".join(x for x in [sur, name, patr] if x)
                        country, region = _split_country_region(live)
                        status = "OK" if live else "ERROR_NO_DATA"

                        res_row_idx = row_map.get(str(r))
                        if res_row_idx:
                            append_result_row(out_path, int(res_row_idx), src_iin, sur, name, patr, fio, live, country, region, status)
                            log4(f"   [2ND PASS] → OVERWRITE result row {res_row_idx}: status={status}, live='{live}'")
                        else:
                            log4("   [2ND PASS] WARNING: не нашёл строку результата для этого r")

                        _wait_queue_quiet(drv, min_quiet=MIN_STABLE_SEC, total=4)
                        time.sleep(COOLDOWN_BETWEEN_ROWS)

                        checkpoint_update(phase="SECOND_PASS")

                    except Exception as e:
                        if is_browser_dead(e):
                            raise BrowserCrashed(f"Browser dead during SECOND_PASS at src row {r}: {e}") from e
                        log4(f"⚠ [2ND PASS] Ошибка на строке {r}: {e}")

            log4(f"✅ Второй проход завершён. Итог сохранён: {out_path}")
            checkpoint_update(phase="DONE")

        if state.get("phase") == "DONE":
            log4("🎉 DONE. Удаляю checkpoint.")
            clear_checkpoint()

        return True

    except BrowserCrashed:
        raise
    except Exception as e:
        # если тут умер браузер — тоже наверх
        if is_browser_dead(e):
            raise BrowserCrashed(f"Browser dead (outer parse): {e}") from e
        raise

# ========= НОВОЕ: bootstrap (восстановление UI после рестарта браузера) =========
def bootstrap_to_person_form(drv):
    """
    После каждого рестарта браузера нужно заново:
      login -> go_to_send_docs -> send_claim -> block3 -> открыть форму участника
    """
    login(drv)
    go_to_send_docs(drv)
    send_claim(drv)
    block3_fill_case_and_open_modal(drv)
    WebDriverWait(drv, 20).until(EC.presence_of_element_located(IIN))

# ========= ORCHESTRATOR (с 20 рестартами браузера) =========
def main():
    log_step(f"OUTPUT файл текущего запуска: {OUT_XLSX}")
    log_step(f"CHECKPOINT: {CHECKPOINT_PATH}")
    log_step(f"MAX_BROWSER_RESTARTS: {MAX_BROWSER_RESTARTS}")

    last_exc = None

    for attempt in range(1, MAX_BROWSER_RESTARTS + 1):
        drv = None
        try:
            log_step(f"\n=== BROWSER ATTEMPT {attempt}/{MAX_BROWSER_RESTARTS} ===")
            drv = init_driver()

            bootstrap_to_person_form(drv)

            ok = parse_people_from_excel_and_save_v2_resumable(drv, OUT_XLSX)
            if ok:
                log_step("✅ Работа завершена успешно.")
                safe_quit(drv)
                return

        except BrowserCrashed as e:
            last_exc = e
            log_warn(f"Браузер/сессия упали. Перезапускаю попытку... Причина: {e}")
            safe_quit(drv)
            time.sleep(2.0)
            continue

        except Exception as e:
            last_exc = e
            # если это похоже на падение браузера — тоже рестартим
            if is_browser_dead(e):
                log_warn(f"Похоже на падение браузера. Рестарт... {e}")
                safe_quit(drv)
                time.sleep(2.0)
                continue

            # иначе — это "обычная" ошибка, прекращаем
            log_err(f"❌ Фатальная ошибка (не похоже на падение браузера): {e}")
            safe_quit(drv)
            raise

        finally:
            # если дошли сюда и драйвер ещё жив — закрываем
            safe_quit(drv)

    # если дошли сюда — исчерпали рестарты
    raise RuntimeError(f"Исчерпаны попытки перезапуска браузера ({MAX_BROWSER_RESTARTS}). Последняя ошибка: {last_exc}")

if __name__ == "__main__":
    main()
