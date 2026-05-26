# -*- coding: utf-8 -*-
r"""
Скрипт для подбора адреса и ЧСИ по ИИН:

1) Логинится в office.sud.kz (переключает язык на РУС).
2) Переходит в "Подача документов".
3) Выбирает CIVIL / FIRSTINSTANCE / Иск и нажимает «Отправить».
4) На createRequest.xhtml заполняет поля и открывает форму участника.
5) Берёт ИИН из Excel-файла
      C:\Users\ARMADA\Desktop\Подача по АИС ОИП\Список для подачи заявления ЧСИ через АИС ОИП_Omega.xlsx
   (колонка E),
   по каждому ИИН тянет данные:
      - полный адрес ("Место жительства") пишет в колонку I;
      - по адресу подбирает ФИО ЧСИ из файла
        C:\Users\ARMADA\Desktop\Подача по АИС ОИП\ЧСИ в АИС ОИП.xlsx
        (район — колонка B, ФИО — колонка C, регион — колонка D)
        и пишет ФИО ЧСИ в колонку H.

Логика парсинга данных с Судебного кабинета сохранена из твоего рабочего скрипта.
"""

import os
import time
import logging
import re
from difflib import SequenceMatcher

from openpyxl import load_workbook

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
)

# ========= КОНФИГ =========
USER_AUTH     = "230240016634"        # твой ИИН/БИН
USER_PASSWORD = "n3y&pAM5mD&4zKZ"     # твой пароль

BASE      = "https://office.sud.kz"
LOGIN_URL = f"{BASE}/index.xhtml"
HOME_URL  = f"{BASE}/form/proceedings/services.xhtml"
SEND_DOCS_URL = f"{BASE}/form/send/index.xhtml"

# Паузы
WAIT   = 1.0

# ========= ПУТИ =========
BASE_DIR   = r"C:\Users\user\Desktop\Подача по АИС ОИП"
INPUT_XLSX = os.path.join(BASE_DIR, "Список для подачи заявления ЧСИ через АИС ОИП_Omega.xlsx")
OUT_XLSX   = INPUT_XLSX  # пишем туда же

CHSI_XLSX  = os.path.join(BASE_DIR, "ЧСИ в АИС ОИП.xlsx")

# Колонки в основном файле (номер столбца)
COL_IIN   = 5  # E – ИИН
COL_CHSI  = 8  # H – ФИО ЧСИ
COL_ADDR  = 9  # I – полный адрес

# ========= ЛОГИ =========
logging.basicConfig(
    filename="sud_chsi_script.log",
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

def log4(msg):
    print(time.strftime("[%H:%M:%S]"), msg, flush=True)
    logging.info(msg)

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
    log_step("[LOGIN] Открываю страницу логина…")
    driver.get(LOGIN_URL)
    time.sleep(WAIT)

    # язык — РУС
    try:
        ru = driver.find_elements(By.XPATH, "//a[contains(.,'РУС') and not(contains(@class,'active'))]")
        if ru:
            ru[0].click()
            log_step("[LOGIN] Переключил язык на РУС")
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
    log_step("[LOGIN] Ввёл ИИН/БИН и Пароль")
    submit.click()
    log_step("[LOGIN] Нажал 'Войти'")
    time.sleep(WAIT * 2)

# ========= ПЕРЕХОД В «Подача документов» =========
def go_to_send_docs(driver):
    log_step("[NAV] Открываю главную страницу (плитки)…")
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
        log_step("[NAV] Кликнул по плитке 'Подача документов'")
    else:
        log_warn("[NAV] Плитка не найдена, открываю 'Подача документов' по прямому URL")
        driver.get(SEND_DOCS_URL)

    try:
        WebDriverWait(driver, 10).until(EC.url_contains("/form/send"))
    except TimeoutException:
        log_warn("[NAV] Не удалось явно подтвердить загрузку страницы 'Подача документов'")

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
        (By.XPATH, "//button[normalize-space()='Отправить']"),
        # на всякий случай — казахский вариант
        (By.XPATH, "//input[@type='submit' and @value='Жіберу']"),
        (By.XPATH, "//button[normalize-space()='Жіберу']"),
        # и любой submit внутри формы подачи документов
        (By.XPATH, "//form[contains(@action,'send') or contains(@id,'sendForm')]//input[@type='submit']"),
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

# ========= БЛОК 4: парсинг ИИН из Excel (оставляем как было, но под новый файл) =========
print("=== BLOCK-4 for CHSI loaded (IIN from column E, address->I, CHSI->H) ===")

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

    clicked_next = _safe_click(drv, BTN_NEXT_SIDE, "Далее")

    try:
        WebDriverWait(drv, 12).until(EC.presence_of_element_located(IIN))
        if clicked_next:
            log4("   форма участника открыта заново (после выбора стороны)")
        else:
            log4("   форма участника открыта заново (без выбора стороны)")
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

# ===== ЧТЕНИЕ СПРАВОЧНИКА ЧСИ И "УМНОЕ" СОВПАДЕНИЕ =====
def _normalize_txt(s: str) -> str:
    if not s:
        return ""
    s = str(s).upper()
    s = s.replace("Ё", "Е")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def fio_title_case(s: str) -> str:
    """Нормальное отображение ФИО (каждое слово и часть через дефис с заглавной)."""
    if not s:
        return ""
    s = re.sub(r"\s+", " ", str(s).strip())
    words = []
    for part in s.split(" "):
        sub = [sp.capitalize() for sp in part.split("-")]
        words.append("-".join(sub))
    return " ".join(words)

def load_chsi_table(path: str):
    rows = []
    if not os.path.exists(path):
        log_warn(f"[CHSI] Файл не найден: {path}")
        return rows

    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    for row in ws.iter_rows(min_row=2, values_only=True):
        district = row[1] if len(row) > 1 else None   # колонка B
        fio      = row[2] if len(row) > 2 else None   # колонка C
        region   = row[3] if len(row) > 3 else None   # колонка D

        if not fio:
            continue

        reg_u = _normalize_txt(region)
        dist_u = _normalize_txt(district)
        fio_u = _normalize_txt(fio)

        rows.append({
            "district": district or "",
            "region": region or "",
            "fio": fio or "",
            "district_u": dist_u,
            "region_u": reg_u,
            "fio_u": fio_u,
        })
    wb.close()
    log_step(f"[CHSI] Загружено строк: {len(rows)}")
    return rows

CHSI_ROWS = load_chsi_table(CHSI_XLSX)

# Разрешённые ЧСИ для Карагандинской области
KARAGANDA_CHSI_ALLOWED = {
    _normalize_txt("Алибеков Арман Жасуланович"),
    _normalize_txt("Балтабаев Аян Хайдарович"),
    _normalize_txt("Висмурадов Хусайн Салманович"),
    _normalize_txt("Келдибеков Жазибек Мажирович"),
}

def _find_chsi_by_fio(fio_search: str):
    """Поиск ЧСИ по ФИО в справочнике."""
    if not fio_search:
        return None
    target = _normalize_txt(fio_search)
    for r in CHSI_ROWS:
        if _normalize_txt(r.get("fio", "")) == target:
            return r
    return None

def pick_chsi_by_address(address: str):
    """
    Подбор ЧСИ по адресу.

    1) Сначала применяем жёсткие правила по началу адреса (АЛМАТЫ, АСТАНА,
       ШЫМКЕНТ, все области и детальная логика по Карагандинской области).
    2) Если ни одно правило не сработало — используем старую "умную" логику
       по региону/району, но для КАРАГАНДИНСКОЙ ОБЛАСТИ ограничиваемся
       только 4 разрешёнными ЧСИ.
    """
    if not address:
        return None, 0.0

    addr_u = _normalize_txt(address)

    # --------- ЖЁСТКИЕ ПРАВИЛА ПО ГОРОДАМ-РЕСПУБЛИКАНСКОГО ЗНАЧЕНИЯ ---------
    if addr_u.startswith("КАЗАХСТАН, АЛМАТЫ"):
        return _find_chsi_by_fio("Нугманова Ильнара Султанбековна"), 100.0

    if addr_u.startswith("КАЗАХСТАН, АСТАНА") or addr_u.startswith("КАЗАХСТАН, НУР-СУЛТАН"):
        return _find_chsi_by_fio("Буганаев Уалихан Муратбекович"), 100.0

    if addr_u.startswith("КАЗАХСТАН, ШЫМКЕНТ"):
        return _find_chsi_by_fio("Бейсенова Алия Жумабековна"), 100.0

    # --------- ЖЁСТКИЕ ПРАВИЛА ПО ОБЛАСТЯМ ---------
    if addr_u.startswith("КАЗАХСТАН, АКМОЛИНСКАЯ ОБЛАСТЬ"):
        return _find_chsi_by_fio("Гафурова Зарина Радиковна"), 100.0

    if addr_u.startswith("КАЗАХСТАН, АКТЮБИНСКАЯ ОБЛАСТЬ"):
        return _find_chsi_by_fio("Кулшанов Алмас Бакитович"), 100.0

    if addr_u.startswith("КАЗАХСТАН, АЛМАТИНСКАЯ ОБЛАСТЬ"):
        return _find_chsi_by_fio("Досымбекова Эльмира Есбосыновна"), 100.0

    if addr_u.startswith("КАЗАХСТАН, АТЫРАУСКАЯ ОБЛАСТЬ"):
        return _find_chsi_by_fio("Мамбетова Айгуль Махсутовна"), 100.0

    if addr_u.startswith("КАЗАХСТАН, ВОСТОЧНО-КАЗАХСТАНСКАЯ ОБЛАСТЬ"):
        return _find_chsi_by_fio("Жайсанбаев Куаныш Максатович"), 100.0

    if addr_u.startswith("КАЗАХСТАН, ЖАМБЫЛСКАЯ ОБЛАСТЬ"):
        return _find_chsi_by_fio("Болтаев Еркебулан Бекайдарович"), 100.0

    if addr_u.startswith("КАЗАХСТАН, ЗАПАДНО-КАЗАХСТАНСКАЯ ОБЛАСТЬ"):
        return _find_chsi_by_fio("Сюнтиев Адиль Сапарович"), 100.0

    if addr_u.startswith("КАЗАХСТАН, КОСТАНАЙСКАЯ ОБЛАСТЬ"):
        return _find_chsi_by_fio("Исмухамедов Анвар Маликович"), 100.0

    if addr_u.startswith("КАЗАХСТАН, КЫЗЫЛОРДИНСКАЯ ОБЛАСТЬ"):
        return _find_chsi_by_fio("Алшагиров Ербулат Қайырбекұлы"), 100.0

    if addr_u.startswith("КАЗАХСТАН, МАНГИСТАУСКАЯ ОБЛАСТЬ"):
        return _find_chsi_by_fio("Смаилов Рауан Темиргалиевич"), 100.0

    if addr_u.startswith("КАЗАХСТАН, ОБЛАСТЬ АБАЙ"):
        return _find_chsi_by_fio("Бектемиров Рахат Хамитович"), 100.0

    if addr_u.startswith("КАЗАХСТАН, ОБЛАСТЬ ЖЕТІСУ"):
        return _find_chsi_by_fio("Асбекова Асем Сакеновна"), 100.0

    if addr_u.startswith("КАЗАХСТАН, ОБЛАСТЬ ҰЛЫТАУ"):
        return _find_chsi_by_fio("Длимбетов Ардак Катарбаевич"), 100.0

    if addr_u.startswith("КАЗАХСТАН, ПАВЛОДАРСКАЯ ОБЛАСТЬ"):
        return _find_chsi_by_fio("Шахматалинова Айжан Айдосовна"), 100.0

    if addr_u.startswith("КАЗАХСТАН, СЕВЕРО-КАЗАХСТАНСКАЯ ОБЛАСТЬ"):
        return _find_chsi_by_fio("Темирханова Дина Куандыковна"), 100.0

    if addr_u.startswith("КАЗАХСТАН, ТУРКЕСТАНСКАЯ ОБЛАСТЬ"):
        return _find_chsi_by_fio("Сердалиев Ерланбек Навбатиллаевич"), 100.0

    # --------- ДЕТАЛЬНАЯ ЛОГИКА ПО КАРАГАНДИНСКОЙ ОБЛАСТИ ---------
    if addr_u.startswith("КАЗАХСТАН, КАРАГАНДИНСКАЯ ОБЛАСТЬ, КАРАГАНДА"):
        return _find_chsi_by_fio("Алибеков Арман Жасуланович"), 100.0

    if (
        addr_u.startswith("КАЗАХСТАН, КАРАГАНДИНСКАЯ ОБЛАСТЬ, ТЕМИРТАУ")
        or addr_u.startswith("КАЗАХСТАН, КАРАГАНДИНСКАЯ ОБЛАСТЬ, НУРИНСКИЙ РАЙОН")
        or addr_u.startswith("КАЗАХСТАН, КАРАГАНДИНСКАЯ ОБЛАСТЬ, ОСАКАРОВСКИЙ РАЙОН")
        or addr_u.startswith("КАЗАХСТАН, КАРАГАНДИНСКАЯ ОБЛАСТЬ, КАРКАРАЛИНСКИЙ РАЙОН")
        or addr_u.startswith("КАЗАХСТАН, КАРАГАНДИНСКАЯ ОБЛАСТЬ, БУХАР-ЖЫРАУСКИЙ Р-Н")
    ):
        return _find_chsi_by_fio("Балтабаев Аян Хайдарович"), 100.0

    if (
        addr_u.startswith("КАЗАХСТАН, КАРАГАНДИНСКАЯ ОБЛАСТЬ, АБАЙСКИЙ РАЙОН")
        or addr_u.startswith("КАЗАХСТАН, КАРАГАНДИНСКАЯ ОБЛАСТЬ, ШАХТИНСК")
        or addr_u.startswith("КАЗАХСТАН, КАРАГАНДИНСКАЯ ОБЛАСТЬ, САРАНЬ")
    ):
        return _find_chsi_by_fio("Висмурадов Хусайн Салманович"), 100.0

    if (
        addr_u.startswith("КАЗАХСТАН, КАРАГАНДИНСКАЯ ОБЛАСТЬ, БАЛХАШ")
        or addr_u.startswith("КАЗАХСТАН, КАРАГАНДИНСКАЯ ОБЛАСТЬ, ПРИОЗЕРСК")
        or addr_u.startswith("КАЗАХСТАН, КАРАГАНДИНСКАЯ ОБЛАСТЬ, АКТОГАЙСКИЙ РАЙОН")
        or addr_u.startswith("КАЗАХСТАН, КАРАГАНДИНСКАЯ ОБЛАСТЬ, ШЕТСКИЙ РАЙОН")
    ):
        return _find_chsi_by_fio("Келдибеков Жазибек Мажирович"), 100.0

    # --------- ДАЛЬШЕ — СТАРАЯ "УМНАЯ" ЛОГИКА (ФОЛБЭК) ---------

    best = None
    best_score = 0.0

    # если в адресе Карагандинская область — ограничиваемся 4 ЧСИ
    if "КАРАГАНДИНСКАЯ ОБЛАСТЬ" in addr_u:
        candidates = [r for r in CHSI_ROWS if r.get("fio_u") in KARAGANDA_CHSI_ALLOWED]
        if not candidates:
            candidates = CHSI_ROWS  # на всякий случай
    else:
        candidates = CHSI_ROWS

    for r in candidates:
        score = 0.0
        reg_u = r.get("region_u", "")
        dist_u = r.get("district_u", "")

        if reg_u and reg_u in addr_u:
            score += 2.0
        if dist_u and dist_u in addr_u:
            score += 3.0

        if dist_u:
            ratio = SequenceMatcher(None, dist_u, addr_u).ratio()
            score += ratio  # 0..1

        if score > best_score:
            best_score = score
            best = r

    return (best, best_score) if best else (None, 0.0)


# ===== ОСНОВНОЙ ЦИКЛ ПО ИИН (НОВАЯ ВЕРСИЯ) =====
def parse_chsi_from_excel(drv):
    """
    Новый вариант:
      - ИИН берём из колонки E (5);
      - по каждому ИИН тянем данные с сайта;
      - полный адрес пишем в колонку I (9);
      - по адресу подбираем ЧСИ из CHSI_XLSX и пишем ФИО в колонку H (8).
    """
    MAX_RESTARTS_PER_ROW = 2

    log4("=== START parse_chsi_from_excel ===")
    log4(f"Файл: {INPUT_XLSX}")

    wb = load_workbook(INPUT_XLSX, data_only=True)
    ws = wb.active

    # Заголовки H/I
    header_chsi = ws.cell(row=1, column=COL_CHSI)
    if not header_chsi.value or str(header_chsi.value).strip() == "":
        header_chsi.value = "ФИО ЧСИ (подбор)"

    header_addr = ws.cell(row=1, column=COL_ADDR)
    if not header_addr.value or str(header_addr.value).strip() == "":
        header_addr.value = "Адрес с Судебного кабинета"

    WebDriverWait(drv, 20).until(EC.presence_of_element_located(IIN))

    processed = 0
    failed_rows = []

    try:
        for r in range(2, ws.max_row + 1):
            src_iin = _norm_iin(ws.cell(row=r, column=COL_IIN).value)
            if not src_iin or len(src_iin) != 12:
                continue

            processed += 1
            t_row_start = time.time()
            log4(f"#{processed} (Excel row {r}) → ИИН: {src_iin}")

            success = False

            for restart in range(MAX_RESTARTS_PER_ROW + 1):
                if restart > 0:
                    log4(f"   ↻ полный перезапуск формы для этого ИИН (рестарт {restart})")
                    if not _reopen_person_modal(drv):
                        log4("   ❌ не удалось перезапустить форму, выхожу из попыток по этому ИИН")
                        break

                try:
                    iin_el = ensure_person_form_open(drv, timeout=20)
                except RuntimeError as e:
                    log4(f"   ❌ форма участника недоступна: {e}")
                    break

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
                    live = data.get("live", "")
                    ws.cell(row=r, column=COL_ADDR, value=live)

                    chsi_row, score = pick_chsi_by_address(live)
                    if chsi_row:
                        fio = fio_title_case(chsi_row["fio"])
                        ws.cell(row=r, column=COL_CHSI, value=fio)
                        log4(f"   → CHSI: '{fio}' (score={score:.3f})")
                    else:
                        log4("   → CHSI: не найден по адресу")

                    success = True
                    break
                else:
                    log4("   данных нет / не обновились, пробую перезапустить форму…")

            if not success:
                failed_rows.append(r)
                log4(f"   → row {r}: ERROR_NO_DATA")

            # паузы и очередь ajax
            _wait_queue_quiet(drv, min_quiet=MIN_STABLE_SEC, total=4)
            time.sleep(COOLDOWN_BETWEEN_ROWS)
            elapsed = time.time() - t_row_start
            if elapsed < MIN_SECONDS_PER_ROW:
                time.sleep(MIN_SECONDS_PER_ROW - elapsed)

            if processed % AUTOSAVE_EVERY == 0:
                _safe_save(wb, OUT_XLSX)

        _safe_save(wb, OUT_XLSX)
        log4(f"✅ Первый проход завершён. Сохранено в: {OUT_XLSX}")

        # Второй проход по проблемным (если нужны, можно оставить / убрать)
        if failed_rows:
            log4(f"=== SECOND PASS: {len(failed_rows)} строк с ERROR_NO_DATA ===")
            try:
                go_to_send_docs(drv)
                send_claim(drv)
                block3_fill_case_and_open_modal(drv)
                WebDriverWait(drv, 20).until(EC.presence_of_element_located(IIN))
            except Exception as e:
                log4(f"⚠ не удалось заново открыть форму участника: {e}")
            else:
                for r in failed_rows:
                    src_iin = _norm_iin(ws.cell(row=r, column=COL_IIN).value)
                    if not src_iin:
                        continue

                    log4(f"[2ND PASS] row {r} → ИИН: {src_iin}")

                    try:
                        iin_el = ensure_person_form_open(drv, timeout=20)
                    except RuntimeError as e:
                        log4(f"   ❌ форма участника недоступна при втором проходе: {e}")
                        break

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

                    if changed:
                        live = data.get("live", "")
                        ws.cell(row=r, column=COL_ADDR, value=live)

                        chsi_row, score = pick_chsi_by_address(live)
                        if chsi_row:
                            fio = fio_title_case(chsi_row["fio"])
                            ws.cell(row=r, column=COL_CHSI, value=fio)
                            log4(f"   [2ND PASS] → CHSI: '{fio}' (score={score:.3f})")
                        else:
                            log4("   [2ND PASS] → CHSI: не найден")

                    else:
                        log4("   [2ND PASS] снова нет данных, оставляю пустым")

                    _wait_queue_quiet(drv, min_quiet=MIN_STABLE_SEC, total=4)
                    time.sleep(COOLDOWN_BETWEEN_ROWS)

                _safe_save(wb, OUT_XLSX)
                log4(f"✅ Второй проход завершён. Итог: {OUT_XLSX}")

    finally:
        try:
            _safe_save(wb, OUT_XLSX)
        except Exception:
            pass

# ========= ORCHESTRATOR =========
def run():
    drv = init_driver()
    try:
        login(drv)
        go_to_send_docs(drv)
        send_claim(drv)
        block3_fill_case_and_open_modal(drv)
        parse_chsi_from_excel(drv)
    finally:
        # при желании можно закрывать браузер
        # drv.quit()
        log_step("[EXIT] Скрипт завершён.")

