# -*- coding: utf-8 -*-
"""
Полный объединённый скрипт:

1) Логинится в office.sud.kz
2) Переходит в "Подача документов"
3) Выбирает CIVIL / FIRSTINSTANCE / Иск и нажимает «Отправить»
4) На createRequest.xhtml заполняет поля и открывает форму участника
5) Читает ИИН из Excel, по каждому ИИН тянет адрес и записывает в тот же Excel:
      - ИИН берётся из колонки D;
      - обрабатываются ВСЕ строки;
      - в ту же строку, в колонку O пишется полный адрес (Место жительства);
      - заголовок O1 ставится только если пустой;
      - в колонку P пишется Регион, определённый по адресу из O (заголовок также ставится только если пустой).

6) После Selenium-части:
      - по адресам (O) и регионам (P) заполняет колонку Q
        "Судебный орган с Судебного кабинета" по справочнику
        "Суды по гражданским делам.xlsx" (лист "Возврат") для всех строк.

7) По адресу (O), региону (P) и суду (Q) определяет УГД и БИН УГД
   из справочника "БИН(УГД).xlsx" и записывает в колонки L и M
   листа "Отмены", а также синхронизирует в лист "Данные для шаблонов"
   по уникальному номеру.

8) Формирует файл ProcessImport.xlsx в сетевой папке \\192.168.1.251\A-Omega
"""

import os
import sys
import argparse

# ========= АРГУМЕНТЫ КОМАНДНОЙ СТРОКИ =========
# Разбираются здесь, ДО импорта config/sbordoc_files.logi (см. ниже) — иначе
# COMPANY_ID уходит в config.py уже после того, как он определил ROOT по
# умолчанию, и --company_id перестаёт на что-либо влиять.
def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('--workdir',    type=str, default=None)
    parser.add_argument('--excel_file', type=str, default=None)
    parser.add_argument('--company_id', type=str, default=None)
    return parser.parse_known_args()[0]

args = parse_args()
if not args.company_id or not args.company_id.strip():
    print("❌ ОШИБКА: не передан --company_id — компания не определена, запуск остановлен.")
    sys.exit(1)
os.environ['COMPANY_ID'] = args.company_id.strip()

import time
import logging
import datetime as _dt
import re
import unicodedata
import xml.etree.ElementTree as ET
from pathlib import Path
from dataclasses import dataclass
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
import pandas as pd
from difflib import SequenceMatcher

from openpyxl import load_workbook, Workbook

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from sbordoc_files.logi import LOG_SUMMARY

from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    StaleElementReferenceException,
    ElementNotInteractableException,
)

from config import (
    CREDENTIALS,
    COMPANY_CREDENTIALS,
    #MAIN_EXCEL,
    ROOT,
    TARGET_BASE,
    EMAIL_SOURCE,
    LOG_FILE
)

from utils import (
    normalize,
    filename_contains_iin,
    ensure_client_folder,
    safe_log,
    safe_update_summary
)

# ========= КОНФИГ =========
BASE      = "https://office.sud.kz"
LOGIN_URL = f"{BASE}/index.xhtml"
HOME_URL  = f"{BASE}/form/proceedings/services.xhtml"
SEND_DOCS_URL = f"{BASE}/form/send/index.xhtml"
SEARCH_FALLBACK_URL = f"{BASE}/lawsuit/document.xhtml"
CREATE_REQUEST_URL = f"{BASE}/form/requestType2/createRequest.xhtml"
VIEWSTATE_NAMES = ("javax.faces.ViewState", "jakarta.faces.ViewState")

# Паузы
WAIT   = 1.0
RETRY  = 8

# ========= ПУТИ =========
BASE_DIR   = rf"{ROOT}\Документы для подачи Исков"
INPUT_XLSX = args.excel_file
OUT_XLSX   = INPUT_XLSX

FILE_PEOPLE  = INPUT_XLSX
# Справочники судов и УГД — как и FILE_ENBEKSHI/FILE_SHET ниже, общие для ВСЕХ
# компаний (не зависят от того, чьё это дело), поэтому путь зафиксирован на
# одну сетевую папку, а не строится через ROOT компании.
FILE_COURTS  = rf"\\192.168.1.200\workfolder\Документы для подачи Исков\Шаблоны документов\Суды по гражданским делам.xlsx"
FILE_UGD     = rf"\\192.168.1.200\workfolder\Документы для подачи Исков\Шаблоны документов\БИН(УГД) .xlsx"
# Справочник сёл/округов Енбекшиказахского района общий для всех компаний
# (разбивка района между двумя судами не зависит от того, чьё это дело),
# поэтому путь зафиксирован на одну сетевую папку, а не на ROOT компании.
FILE_ENBEKSHI = rf"\\192.168.1.200\workfolder\Документы для подачи Исков\Шаблоны документов\Енбекшиказахский суд.xlsx"
# Аналогичный справочник для Шетского района Карагандинской области (тоже
# поделён между двумя судами по сёлам/округам). Имя файла — как оно реально
# называется в сетевой папке.
FILE_SHET = rf"\\192.168.1.200\workfolder\Документы для подачи Исков\Шаблоны документов\КРГ Шетский  (1).xlsx"

# ========= КОНФИГ =========
_creds        = COMPANY_CREDENTIALS.get(args.company_id.strip(), COMPANY_CREDENTIALS['1'])
USER_AUTH     = _creds['sk_login']
USER_PASSWORD = _creds['sk_password']

# Папка на сервере
NETWORK_DIR = r"\\192.168.1.251\A-Omega"

os.makedirs(NETWORK_DIR, exist_ok=True)
os.makedirs(BASE_DIR, exist_ok=True)

PROCESSIMPORT_PATH = os.path.join(NETWORK_DIR, "ProcessImport.xlsx")

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

    try:
        ru = driver.find_elements(By.XPATH, "//a[contains(.,'РУС') and not(contains(@class,'active'))]")
        if ru:
            ru[0].click()
            log_step("Переключил язык на РУС")
            time.sleep(WAIT)
    except Exception:
        log_warn("Не удалось переключить язык на РУС")

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
        log_warn("Не удалось явно подтвердить загрузку страницы 'Подача документов' — продолжаю")

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
    w.until(EC.presence_of_element_located(locator))

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
            print(f" {desc} → {value}")
            return
        except (StaleElementReferenceException, TimeoutException) as e:
            last_err = e

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
    print(f" {desc} → {value} (через JS)")

def send_claim(driver, timeout=20):
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
    print("Нажал «Отправить».")

# ========= БЛОК 3: createRequest + форма участника =========
def _wait(drv, cond, t=25):
    return WebDriverWait(drv, t, ignored_exceptions=(StaleElementReferenceException,)).until(cond)

def _select_value(drv, css, value, desc, t=25):
    sel = _wait(drv, EC.presence_of_element_located((By.CSS_SELECTOR, css)), t)
    for _ in range(3):
        try:
            Select(sel).select_by_value(value)
            _wait(drv, lambda d: d.find_element(By.CSS_SELECTOR, css).get_attribute("value") == value, 10)
            print(f" {desc} → value={value}")
            return
        except StaleElementReferenceException:
            sel = drv.find_element(By.CSS_SELECTOR, css)
        except Exception:
            pass
    drv.execute_script("""
        var s = document.querySelector(arguments[0]);
        var val = arguments[1];
        if(!s) return false;
        s.value = val;
        s.dispatchEvent(new Event('change', {bubbles:true}));
        return true;
    """, css, value)
    _wait(drv, lambda d: d.find_element(By.CSS_SELECTOR, css).get_attribute("value") == value, 10)
    print(f" {desc} → value={value} (JS)")

def _click(drv, locator, desc, t=25):
    btn = _wait(drv, EC.element_to_be_clickable(locator), t)
    drv.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    try:
        btn.click()
    except Exception:
        drv.execute_script("arguments[0].click();", btn)
    print(f" Нажал {desc}")

def block3_fill_case_and_open_modal(driver, timeout=40):
    w = WebDriverWait(driver, timeout)

    try:
        w.until(
            lambda d: "/form/requestType2/createRequest.xhtml" in d.current_url
            or d.find_element(By.CSS_SELECTOR, "select[id$=':edit-category']")
        )
    except Exception:
        raise TimeoutException("Страница createRequest.xhtml не открыта")

    _select_value(driver, "select[id$=':edit-categoryGroup']", "2", "Вид производства по делу", t=timeout)

    _wait(driver, EC.presence_of_element_located(
        (By.CSS_SELECTOR, "select[id$=':edit-category'] option[value='27']")), 30)
    _select_value(driver, "select[id$=':edit-category']", "27", "Категория дела", t=timeout)

    _wait(driver, EC.presence_of_element_located(
        (By.CSS_SELECTOR, "select[id$=':edit-character'] option[value='1']")), 20)
    _select_value(driver, "select[id$=':edit-character']", "1", "Характер заявления", t=timeout)

    try:
        cb = w.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "input[id$=':edit-simpleProcess'][type='checkbox']")))
        if not cb.is_selected():
            driver.execute_script("arguments[0].click();", cb)
        print("Галочка «Дело упрощенного производства» установлена")
    except Exception:
        print(" Не удалось установить галочку «Дело упрощенного производства» (продолжаю)")

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

    person_iin_css = "input[id$=':person-iin']"

    def modal_switched(d):
        try:
            return d.find_element(By.CSS_SELECTOR, person_iin_css)
        except Exception:
            return False

    for attempt in range(3):
        try:
            btn_next = _wait(driver, EC.element_to_be_clickable((By.XPATH, next_xpath)), 15)
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn_next)
            try:
                btn_next.click()
            except Exception:
                driver.execute_script("arguments[0].click();", btn_next)
            _wait(driver, modal_switched, 25)
            print(" Блок 3 завершён: открыта форма участника с полем ИИН.")
            break
        except StaleElementReferenceException:
            print("⚠ Stale element на кнопке «Далее», пробую ещё раз…")
            if attempt == 2:
                raise
            time.sleep(1)

# ========= БЛОК 4: парсинг ИИН из Excel =========
print("=== BLOCK-4 loaded (все строки, write address to column O, region to P) ===")

AUTOSAVE_EVERY = 25

PERSON_TIMEOUT         = 8
RETRIES_PER_PERSON     = 2
PAUSE_BETWEEN_TRIES    = 0.4
COOLDOWN_BETWEEN_ROWS  = 1.1
MIN_SECONDS_PER_ROW    = 1.2
POLL                   = 0.10
MIN_STABLE_SEC         = 0.35
HARD_WAIT_AFTER_SEARCH = 0.20

IIN  = (By.CSS_SELECTOR, 'input[id$=":person-iin"]')
SUR  = (By.CSS_SELECTOR, 'input[id$=":person-surname"]')
NAM  = (By.CSS_SELECTOR, 'input[id$=":person-firstname"]')
PATR = (By.CSS_SELECTOR, 'input[id$=":person-patronymic"]')
LIVE = (By.CSS_SELECTOR, 'textarea[id$=":person-livePlace"], input[id$=":person-livePlace"]')

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
        drv, timeout, poll_frequency=POLL,
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
    try:
        drv.execute_script("if (typeof fillPersonData==='function'){fillPersonData('j_idt278:person-iin');}")
        log4("   вызвал fillPersonData()")
        return True
    except Exception:
        pass
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
    try:
        return WebDriverWait(drv, timeout).until(EC.presence_of_element_located(IIN))
    except TimeoutException:
        if not _reopen_person_modal(drv):
            raise RuntimeError("Не удалось найти форму участника (поле ИИН).")
        return WebDriverWait(drv, timeout).until(EC.presence_of_element_located(IIN))

def _safe_save(wb, path):
    try:
        wb.save(path)
        log4(f"💾 autosave to {path}")
        return True
    except Exception as e:
        log4(f"⚠ не удалось сохранить файл: {e}")
        return False

# ===== РЕГИОНЫ =====
REGION_NAMES = [
    "Акмолинская область",
    "Актюбинская область",
    "Алматинская область",
    "город Алматы",
    "город Астана",
    "Атырауская область",
    "Восточно-Казахстанская область",
    "Жамбылская область",
    "Западно-Казахстанская область",
    "Карагандинская область",
    "Костанайская область",
    "Кызылординская область",
    "Мангистауская область",
    "Область Абай",
    "Область Жетісу",
    "Область Ұлытау",
    "Павлодарская область",
    "Северо-Казахстанская область",
    "Туркестанская область",
    "город Шымкент",
]

def extract_region_from_address(address: str) -> str:
    if not address:
        return ""
    up = str(address).upper()

    if re.search(r'\bАСТАНА\b', up) or "НУР-СУЛТАН" in up:
        return "город Астана"
    if re.search(r'\bШЫМКЕНТ\b', up):
        return "город Шымкент"
    if re.search(r'\bАЛМАТЫ\b', up) and "АЛМАТИНСКАЯ ОБЛАСТЬ" not in up:
        return "город Алматы"

    addr_low = up.lower()
    for reg in REGION_NAMES:
        if reg.lower() in addr_low:
            return reg
    return ""

def parse_people_from_excel_and_save_v2(drv):
    """
    Основной цикл по ИИН — обрабатываются ВСЕ строки.
    ИИН берём из колонки D.
    В колонку O пишем адрес, в колонку P — регион.

    ВАЖНО (правка): при неудаче на строке (нет данных / StaleElement / любая
    другая ошибка) скрипт НЕ переходит к следующей строке, а повторяет
    ТЕКУЩУЮ строку заново (с восстановлением формы), пока не получит адрес,
    либо пока не будет исчерпан общий лимит попыток на строку
    (MAX_TOTAL_ATTEMPTS_PER_ROW). Только после этого строка помечается как
    неудачная, и скрипт переходит дальше — чтобы не зависнуть навечно на
    одном "мёртвом" ИИН.
    """
    MAX_RESTARTS_PER_ROW = 2            # рестартов формы внутри одной попытки
    MAX_TOTAL_ATTEMPTS_PER_ROW = 6      # сколько раз в целом повторяем саму строку
    RETRY_ROW_PAUSE = 1.5               # пауза перед повтором той же строки

    log4("=== START parse_people_from_excel_and_save_v2 ===")
    log4(f"Файл-источник: {INPUT_XLSX}")

    wb = load_workbook(INPUT_XLSX, data_only=True)
    try:
        ws = wb["Отмены"]
    except KeyError:
        ws = wb.active

    header_o = ws.cell(row=1, column=15)
    if not header_o.value or str(header_o.value).strip() == "":
        header_o.value = "Актуальный адрес с СК"

    header_p = ws.cell(row=1, column=16)
    if not header_p.value or str(header_p.value).strip() == "":
        header_p.value = "Регион"

    WebDriverWait(drv, 20).until(EC.presence_of_element_located(IIN))

    failed_rows = []
    processed = 0
    consecutive_errors = 0  # счётчик подряд идущих ERROR_NO_DATA
    MAX_CONSECUTIVE_ERRORS = 3  # после 3х подряд — перелогин

    def _relogin(drv):
        log4("   🔄 Перелогиниваюсь и восстанавливаю форму...")
        try:
            login(drv)
            go_to_send_docs(drv)
            send_claim(drv)
            block3_fill_case_and_open_modal(drv)
            WebDriverWait(drv, 20).until(EC.presence_of_element_located(IIN))
            log4("   ✅ Сессия восстановлена")
            return True
        except Exception as e:
            log4(f"   ❌ Не удалось восстановить сессию: {e}")
            return False

    def _restore_form(reason: str):
        log4(f"   ↻ восстанавливаю форму после {reason}…")
        try:
            time.sleep(1)
            _reopen_person_modal(drv)
            log4("   ✅ форма восстановлена (мягкий перезапуск)")
            return
        except Exception as e1:
            log4(f"   ⚠ мягкий перезапуск не помог: {e1}")
        try:
            go_to_send_docs(drv)
            send_claim(drv)
            block3_fill_case_and_open_modal(drv)
            WebDriverWait(drv, 20).until(EC.presence_of_element_located(IIN))
            log4("   ✅ форма восстановлена (полный перезаход)")
        except Exception as e2:
            log4(f"   ❌ полный перезаход тоже не помог: {e2}")

    def _try_fill_row_once(r, src_iin):
        """
        Одна попытка получить данные для строки r (с внутренними
        рестартами формы MAX_RESTARTS_PER_ROW раз). Возвращает True,
        если данные получены и записаны, False — если нет.
        Может бросить исключение (StaleElement / прочее) — обрабатывается
        вызывающим кодом.
        """
        for restart in range(MAX_RESTARTS_PER_ROW + 1):
            if restart > 0:
                log4(f"   ↻ полный перезапуск формы для этого ИИН (рестарт {restart})")
                if not _reopen_person_modal(drv):
                    log4("   ❌ не удалось перезапустить форму")
                    return False

            try:
                iin_el = ensure_person_form_open(drv, timeout=20)
            except RuntimeError as e:
                log4(f"   ❌ форма участника недоступна: {e}")
                return False

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
                log4("   ⚠ элемент ИИН не интерактивен — ставлю значение через JS")
                drv.execute_script("arguments[0].value = arguments[1];", iin_el, src_iin)

            try:
                drv.execute_script(
                    "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));", iin_el)
                drv.execute_script(
                    "arguments[0].dispatchEvent(new Event('blur',{bubbles:true}));", iin_el)
            except Exception:
                pass

            before = _values_snapshot(drv)
            data = _trigger_and_wait(drv)
            changed = any(data.values()) and any(data[k] != before.get(k, "") for k in data)
            log4(f"   changed={changed}, data={data}")

            if changed:
                live = data.get("live", "")
                ws.cell(row=r, column=15, value=live)
                region = extract_region_from_address(live)
                ws.cell(row=r, column=16, value=region)
                log4(f"   → SAVE: row {r}, address='{live}', region='{region}'")
                return True
            else:
                log4("   данных нет, пробую перезапустить форму…")

        return False

    try:
        for r in range(2, ws.max_row + 1):
            src_iin = _norm_iin(ws.cell(row=r, column=4).value)
            if not src_iin or len(src_iin) != 12:
                continue

            processed += 1
            t_row_start = time.time()
            log4(f"#{processed} (Excel row {r}) → ИИН: {src_iin}")

            success = False
            total_attempts = 0

            # Повторяем ЭТУ ЖЕ строку, пока не получим адрес,
            # либо пока не исчерпаем MAX_TOTAL_ATTEMPTS_PER_ROW попыток.
            while not success and total_attempts < MAX_TOTAL_ATTEMPTS_PER_ROW:
                total_attempts += 1

                # Если много ошибок подряд — перелогиниться
                if consecutive_errors >= MAX_CONSECUTIVE_ERRORS:
                    log4(f"   ⚠ {consecutive_errors} ошибок подряд — пробую перелогиниться")
                    if _relogin(drv):
                        consecutive_errors = 0
                    else:
                        log4("   ❌ Перелогин не удался, продолжаю попытки...")

                try:
                    success = _try_fill_row_once(r, src_iin)

                except StaleElementReferenceException as e:
                    log4(f"   ❌ StaleElement на строке {r} (попытка {total_attempts}/{MAX_TOTAL_ATTEMPTS_PER_ROW}): {e}")
                    consecutive_errors += 1
                    _restore_form("StaleElement")

                except Exception as e:
                    log4(f"   ❌ НЕОЖИДАННАЯ ОШИБКА на строке {r} (попытка {total_attempts}/{MAX_TOTAL_ATTEMPTS_PER_ROW}): {type(e).__name__}: {e}")
                    consecutive_errors += 1
                    _restore_form(type(e).__name__)

                if success:
                    consecutive_errors = 0
                    break

                if total_attempts < MAX_TOTAL_ATTEMPTS_PER_ROW:
                    log4(f"   ↺ строка {r} не получилась, повторяю строку (попытка {total_attempts + 1}/{MAX_TOTAL_ATTEMPTS_PER_ROW})…")
                    consecutive_errors += 1
                    time.sleep(RETRY_ROW_PAUSE)
                    # Если ошибок накопилось — перелогин прямо перед следующей попыткой
                    if consecutive_errors >= MAX_CONSECUTIVE_ERRORS:
                        log4(f"   ⚠ {consecutive_errors} ошибок подряд — немедленный перелогин")
                        if _relogin(drv):
                            consecutive_errors = 0

            if not success:
                failed_rows.append(r)
                log4(f"   → SAVE (error): row {r}, address='' (ERROR_NO_DATA после {MAX_TOTAL_ATTEMPTS_PER_ROW} попыток) — перехожу к следующей строке")

            _wait_queue_quiet(drv, min_quiet=MIN_STABLE_SEC, total=4)
            time.sleep(COOLDOWN_BETWEEN_ROWS)
            elapsed = time.time() - t_row_start
            if elapsed < MIN_SECONDS_PER_ROW:
                time.sleep(MIN_SECONDS_PER_ROW - elapsed)

            if processed % AUTOSAVE_EVERY == 0:
                _safe_save(wb, OUT_XLSX)

    finally:
        try:
            _safe_save(wb, OUT_XLSX)
        except Exception:
            pass

    return processed, failed_rows


# ========= БЛОК 4 (HTTP-версия): поиск по ИИН через прямые запросы =========
# Вместо клика по лупе через Selenium — прямая реплика AJAX-запроса
# fillPersonData(), захваченного в HAR (office.sud.kz_CURRENT_PARTICIPANTS).
# Selenium используется только для логина/навигации/восстановления сессии;
# сам цикл по сотням ИИН идёт через requests, как в podacha_iska_v2.py.

def _parse_partial_response(xml_text: str):
    """Разбирает XML partial-response JSF/RichFaces. Возвращает (updates, new_viewstate)."""
    updates = {}
    viewstate = None
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError:
        return updates, viewstate
    for node in root.iter():
        tag = node.tag.rsplit("}", 1)[-1]
        if tag == "update":
            update_id = node.attrib.get("id", "")
            value = node.text or ""
            updates[update_id] = value
            if "ViewState" in update_id and value.strip():
                viewstate = value.strip()
    return updates, viewstate


def _extract_field_value(html_fragment: str) -> str:
    if not html_fragment:
        return ""
    m = re.search(r'value="([^"]*)"', html_fragment)
    return (m.group(1) if m else "").strip()


def _find_update_by_suffix(updates: dict, suffix: str) -> str:
    for uid, html in updates.items():
        if uid.endswith(suffix):
            return _extract_field_value(html)
    return ""


def _dump_debug_html(html: str, tag: str) -> str:
    debug_path = None
    try:
        debug_dir = Path("sud_http_debug")
        debug_dir.mkdir(parents=True, exist_ok=True)
        debug_path = debug_dir / f"{tag}_{time.strftime('%Y%m%d_%H%M%S')}.html"
        debug_path.write_text(html, encoding="utf-8")
    except Exception:
        pass
    return str(debug_path) if debug_path else "<не удалось сохранить>"


def _extract_js_triggered_ajax_id(html: str, func_name: str):
    """
    Находит ajax-id для кнопок вида onclick="FUNC(...)", которые сервер
    реализует через инлайн-скрипт:
        FUNC=function(...){RichFaces.ajax("AJAX_ID", ...)}
    Работает для fillPersonData, renderAddPersonModalDialog, sendRequest —
    в «сыром» HTTP-ответе (без исполнения в браузере) этот текст всегда
    присутствует как есть, в отличие от driver.page_source в Selenium,
    где jQuery/RichFaces вычищает содержимое script после выполнения.
    """
    m = re.search(
        re.escape(func_name) + r"=function\([^)]*\)\{RichFaces\.ajax\(\"([^\"]+)\"",
        html,
    )
    return m.group(1) if m else None


def _extract_person_form_context(html: str):
    """
    Достаёт form_id (из onclick лупы: fillPersonData('FORM_ID:person-iin'))
    и полный ajax-id лупы (из инлайн-скрипта fillPersonData=function...).
    """
    m_form = re.search(r"onclick=\"fillPersonData\('([^']+):person-iin'\)\"", html)
    ajax_component = _extract_js_triggered_ajax_id(html, "fillPersonData")

    if not m_form or not ajax_component:
        debug_path = _dump_debug_html(html, "person_form")
        raise RuntimeError(
            "Не найдена разметка лупы (fillPersonData) на странице — форма участника не открыта. "
            f"Сохранил HTML для диагностики: {debug_path}"
        )

    return m_form.group(1), ajax_component


def _extract_viewstate_from_html(html: str):
    m = re.search(r'name="javax\.faces\.ViewState"[^>]*value="([^"]*)"', html)
    if m:
        return m.group(1)
    m = re.search(r'<input[^>]*value="([^"]*)"[^>]*name="javax\.faces\.ViewState"', html)
    return m.group(1) if m else None


def _http_session_from_cookies(cookies, user_agent: str) -> requests.Session:
    """Строит requests.Session из cookies, полученных один раз при логине в Selenium."""
    session = requests.Session()
    session.headers.update({
        "User-Agent": user_agent,
        "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
        "Origin": BASE,
    })
    for cookie in cookies:
        kwargs = {
            "name": cookie["name"],
            "value": cookie["value"],
            "path": cookie.get("path", "/"),
        }
        if cookie.get("domain"):
            kwargs["domain"] = cookie["domain"]
        session.cookies.set(**kwargs)
    return session


def selenium_login_get_session() -> requests.Session:
    """
    Единственное место во всём HTTP-режиме, где используется Selenium:
    логин, чтобы получить cookies авторизованной сессии. Дальше браузер
    закрывается — вся навигация и сам поиск идут через requests.
    """
    drv = init_driver()
    try:
        login(drv)
        cookies = drv.get_cookies()
        user_agent = drv.execute_script("return navigator.userAgent;")
    finally:
        try:
            drv.quit()
        except Exception:
            pass
    return _http_session_from_cookies(cookies, user_agent)


# ---------- Мини-реплика SudHttpClient (как в podacha_iska_v2.py) ----------

def _merge_partial_updates_into_html(current_html: str, updates: dict) -> str:
    """Накатывает JSF partial-response updates на отслеживаемый полный HTML (как в браузере)."""
    if not updates:
        return current_html
    for update_id, value in updates.items():
        if update_id.endswith("javax.faces.ViewRoot") or update_id == "javax.faces.ViewRoot":
            if value and "<" in value:
                return value
    if not current_html:
        # Ничего ещё не загружено (до первого client.get()) — просто
        # склеиваем фрагменты. Раньше здесь дополнительно проверялось
        # '"<html" not in current_html.lower()', но страницы office.sud.kz
        # часто начинаются с '<!DOCTYPE html>'+'<head>' без буквального тега
        # <html> — из-за этого условие срабатывало всегда, и полноценный
        # merge по id никогда не выполнялся (терялись несвязанные с
        # обновлением элементы формы).
        html_parts = [v for k, v in updates.items() if "ViewState" not in k and v and "<" in v]
        return "\n".join(html_parts) if html_parts else current_html

    soup = BeautifulSoup(current_html, "lxml")
    for update_id, fragment_html in updates.items():
        if "ViewState" in update_id or not fragment_html or "<" not in fragment_html:
            continue
        target = soup.find(id=update_id)
        fragment_soup = BeautifulSoup(fragment_html, "lxml")
        replacement = fragment_soup.find(id=update_id)
        if replacement is None:
            body = fragment_soup.body
            candidates = [
                node for node in (body.contents if body else fragment_soup.contents)
                if getattr(node, "name", None) is not None
            ]
            replacement = candidates[0] if candidates else None
        if target is not None and replacement is not None:
            target.replace_with(replacement)
        elif replacement is not None:
            (soup.body or soup).append(replacement)
    return str(soup)


def _parse_partial_response_with_redirect(xml_text: str):
    """Как _parse_partial_response, но дополнительно достаёт <redirect url=...>."""
    updates, viewstate = _parse_partial_response(xml_text)
    redirect_url = None
    try:
        root = ET.fromstring(xml_text)
        for node in root.iter():
            if node.tag.rsplit("}", 1)[-1] == "redirect":
                redirect_url = node.attrib.get("url")
    except ET.ParseError:
        pass
    return updates, viewstate, redirect_url


@dataclass
class _HttpNavState:
    url: str
    html: str = ""
    viewstate: str = None


class SudHttpClient:
    """
    Минимальная реплика JSF/RichFaces HTTP-клиента для office.sud.kz —
    держит текущий URL, накопленный (смёрженный) HTML страницы и актуальный
    ViewState, как это делает браузер между AJAX-запросами.
    """

    def __init__(self, session: requests.Session, base_url: str = BASE):
        self.session = session
        self.base_url = base_url.rstrip("/")
        self.state = _HttpNavState(url=self.base_url)

    def _check_auth(self, response):
        low_url = response.url.lower()
        low_text = response.text[:5000].lower()
        if (
            "login" in low_url
            or 'name="login"' in low_text
            or ("войти" in low_text and "судебный кабинет" in low_text)
        ):
            raise RuntimeError("HTTP_SESSION_DEAD")

    def get(self, url: str, referer: str = None):
        full_url = urljoin(self.base_url + "/", url)
        headers = {"Referer": referer} if referer else {}
        response = self.session.get(full_url, headers=headers, timeout=60, allow_redirects=True)
        response.raise_for_status()
        self._check_auth(response)
        self.state.url = response.url
        self.state.html = response.text
        vs = _extract_viewstate_from_html(response.text)
        if vs:
            self.state.viewstate = vs
        return response

    def post_form(self, url: str, data: dict, referer: str = None, ajax: bool = True):
        full_url = urljoin(self.base_url + "/", url)
        payload = {k: v for k, v in data.items() if k not in VIEWSTATE_NAMES}
        if self.state.viewstate:
            payload["javax.faces.ViewState"] = self.state.viewstate

        headers = {"Referer": referer or self.state.url}
        if ajax:
            headers.update({"Accept": "*/*", "Faces-Request": "partial/ajax"})

        response = self.session.post(
            full_url, data=payload, headers=headers, timeout=60, allow_redirects=True,
        )
        response.raise_for_status()
        self._check_auth(response)

        content_type = response.headers.get("Content-Type", "").lower()
        if "xml" in content_type or response.text.lstrip().startswith("<?xml"):
            updates, new_viewstate, redirect_url = _parse_partial_response_with_redirect(response.text)
            if new_viewstate:
                self.state.viewstate = new_viewstate
            self.state.html = _merge_partial_updates_into_html(self.state.html, updates)
            self.state.url = response.url
            if redirect_url:
                redirect_full = urljoin(response.url, redirect_url)
                return self.get(redirect_full, referer=response.url)
        else:
            self.state.url = response.url
            self.state.html = response.text
            vs = _extract_viewstate_from_html(response.text)
            if vs:
                self.state.viewstate = vs
        return response


def _patch_select_value(html: str, select_id: str, value: str) -> str:
    """
    Проставляет выбранную опцию <select id=select_id> в отслеживаемом HTML.

    Сервер в partial-response обычно НЕ перерисовывает сам изменившийся
    <select> (перерисовываются только зависимые от него панели) — но
    живой браузер и так помнит выбор пользователя в DOM независимо от
    ответа сервера. Без этого патча следующий шаг цепочки заново
    сериализовал бы это поле как пустое/дефолтное значение.
    """
    soup = BeautifulSoup(html, "lxml")
    sel = soup.find(id=select_id)
    if sel is None:
        return html
    for opt in sel.find_all("option"):
        if opt.has_attr("selected"):
            del opt["selected"]
    match = sel.find("option", attrs={"value": value})
    if match is None:
        match = next((o for o in sel.find_all("option") if o.get_text() == value), None)
    if match is not None:
        match["selected"] = "selected"
    return str(soup)


def _find_container_id(html: str, known_field_suffix: str) -> str:
    """
    Например known_field_suffix=':case-type' находит id="XXX:case-type"
    (XXX может содержать несколько сегментов через ':') и возвращает XXX.
    """
    m = re.search(r'id="([^"]+)' + re.escape(known_field_suffix) + r'"', html)
    if not m:
        raise RuntimeError(f"Не найден элемент с суффиксом {known_field_suffix!r} на странице.")
    return m.group(1)


def _serialize_naming_container(html: str, container_id: str) -> dict:
    """
    Собирает текущие значения всех input/select/textarea с id вида
    'container_id:*' из отслеживаемого HTML — как их отправил бы браузер
    при сабмите формы (нужно, чтобы каждый AJAX-POST нёс полное текущее
    состояние формы, а не только изменившееся поле — именно так это
    выглядит в захваченном HAR).
    """
    soup = BeautifulSoup(html, "lxml")
    prefix = container_id + ":"
    result = {}
    for el in soup.find_all(["input", "select", "textarea"]):
        el_id = el.get("id")
        if not el_id or not el_id.startswith(prefix):
            continue
        name = el.get("name") or el_id
        if el.name == "select":
            # Если у <option> нет атрибута value, браузер при сабмите
            # использует текст опции — иначе плейсхолдеры вида
            # "-- Выберите --" уходили бы как пустая строка.
            chosen = el.find("option", selected=True) or el.find("option")
            if chosen is not None:
                result[name] = chosen.get("value")
                if result[name] is None:
                    result[name] = chosen.get_text()
            else:
                result[name] = ""
        elif el.name == "textarea":
            result[name] = el.get_text() or ""
        else:
            el_type = (el.get("type") or "text").lower()
            if el_type in ("submit", "reset", "image"):
                # Кнопки этого типа реальный браузер включает в сериализацию
                # формы, только если именно они были нажаты (см. HAR: видимая
                # submit-кнопка «Отправить» отсутствует в реальном теле
                # запроса — вместо неё RichFaces дёргает отдельный скрытый
                # <span>-обработчик sendRequest).
                continue
            if el_type in ("checkbox", "radio"):
                if el.has_attr("checked"):
                    result[name] = el.get("value", "on")
            else:
                result[name] = el.get("value", "")
    result[container_id] = container_id
    return result


def _ajax_meta(ajax_component: str, event: str = None) -> dict:
    """
    Служебные поля JSF/RichFaces AJAX-запроса.

    По HAR: для <select onchange="...change...> пары "ajax_component:
    ajax_component" в теле нет — реальное значение поля живёт под своим
    именем. А для любого КЛИКА (и по кнопке за скрытым <span>, и по обычной
    <input onclick="RichFaces.ajax(...)">, например «Далее») RichFaces
    ДОПОЛНИТЕЛЬНО дописывает такую пару в конец тела запроса — в
    application/x-www-form-urlencoded это дублирует ключ поля, и как raw-
    браузер, так и сервер берут последнее значение, то есть эта пара
    ПЕРЕТИРАЕТ «естественное» значение поля (например текст кнопки).
    Поэтому здесь self-ref добавляется для всех событий, кроме "change".
    """
    meta = {
        "javax.faces.source": ajax_component,
        "javax.faces.partial.execute": f"{ajax_component} @component",
        "javax.faces.partial.render": "@component",
        "org.richfaces.ajax.component": ajax_component,
        "rfExt": "null",
        "AJAX:EVENTS_COUNT": "1",
        "javax.faces.partial.ajax": "true",
    }
    if event:
        meta["javax.faces.partial.event"] = event
        if event == "change":
            # "behavior.event" в HAR встречается только у <select onchange>;
            # для клика по кнопке (event="click") его в теле запроса нет —
            # только partial.event.
            meta["javax.faces.behavior.event"] = event
    if event != "change":
        meta[ajax_component] = ajax_component
    return meta


def _ajax_select_change(client: SudHttpClient, url: str, container_id: str, field: str, value: str, referer: str):
    """Реплика onchange="RichFaces.ajax(this,...)" для одного <select>."""
    state = _serialize_naming_container(client.state.html, container_id)
    ajax_component = f"{container_id}:{field}"
    state[ajax_component] = value
    state.update(_ajax_meta(ajax_component, event="change"))
    client.post_form(url, state, referer=referer)
    client.state.html = _patch_select_value(client.state.html, ajax_component, value)


def _click_js_triggered_button(client: SudHttpClient, url: str, container_id: str, func_name: str, referer: str, extra_state: dict = None):
    """Реплика клика по кнопке onclick="FUNC()" (fillPersonData/sendRequest/renderAddPersonModalDialog)."""
    ajax_component = _extract_js_triggered_ajax_id(client.state.html, func_name)
    if not ajax_component:
        debug_path = _dump_debug_html(client.state.html, f"missing_{func_name}")
        raise RuntimeError(f"Не найдена кнопка ({func_name}) на странице. HTML: {debug_path}")

    state = _serialize_naming_container(client.state.html, container_id)
    if extra_state:
        state.update(extra_state)
    state.update(_ajax_meta(ajax_component))
    client.post_form(url, state, referer=referer)


def _extract_plain_ajax_button_id(html: str, near_field_suffix: str):
    """
    Находит id кнопки вида onclick="RichFaces.ajax('ID',event,{'incId':'1'});
    return false;" (без jsf.util.chain, как «Далее» в модалке выбора стороны
    участника) — по соседству с полем near_field_suffix (например ':pp-side'),
    чтобы не спутать с другими похожими кнопками на странице.

    Разбираем через BeautifulSoup, а не голым regex по сырому тексту: после
    _merge_partial_updates_into_html HTML пересобирается через bs4, и
    экранирование кавычек в атрибутах меняется (&quot;.. → одинарные кавычки
    снаружи + литеральные " внутри) — regex по сырым &quot; после этого
    перестаёт совпадать, а bs4 отдаёт логическое значение атрибута всегда
    одинаково, независимо от исходного экранирования.
    """
    field_idx = html.find(near_field_suffix)
    if field_idx < 0:
        return None
    window = html[field_idx: field_idx + 4000]
    soup = BeautifulSoup(window, "lxml")
    for el in soup.find_all("input"):
        onclick = el.get("onclick") or ""
        m = re.search(
            r'RichFaces\.ajax\("([^"]+)",event,\{"incId":"1"\}\s*\);return false;',
            onclick,
        )
        if m:
            return m.group(1)
    return None


def open_person_form_via_http(client: SudHttpClient):
    """
    Полная HTTP-навигация от уже залогиненной сессии (cookies) до открытой
    формы участника с полем ИИН — без единого клика в Selenium:
    send/index.xhtml (CIVIL/FIRSTINSTANCE/тип 3, «Отправить») →
    createRequest.xhtml (категория/характер дела, чекбокс, «Добавить
    участника процесса») → выбор стороны («Далее»).
    Возвращает (form_id, ajax_component) для fillPersonData.
    """
    log4("   [HTTP-nav] GET send/index.xhtml")
    client.get(SEND_DOCS_URL)
    container1 = _find_container_id(client.state.html, ":case-type")

    log4("   [HTTP-nav] case-type=CIVIL")
    _ajax_select_change(client, SEND_DOCS_URL, container1, "case-type", "CIVIL", client.state.url)
    log4("   [HTTP-nav] instance=FIRSTINSTANCE")
    _ajax_select_change(client, SEND_DOCS_URL, container1, "instance", "FIRSTINSTANCE", client.state.url)
    log4("   [HTTP-nav] request=3")
    _ajax_select_change(client, SEND_DOCS_URL, container1, "request", "3", client.state.url)
    log4("   [HTTP-nav] click sendRequest («Отправить»)")
    _click_js_triggered_button(client, SEND_DOCS_URL, container1, "sendRequest", client.state.url)
    # sendRequest отвечает <redirect> на createRequest.xhtml — client уже перешёл туда.
    log4(f"   [HTTP-nav] → {client.state.url}")

    container2 = _find_container_id(client.state.html, ":edit-categoryGroup")

    log4("   [HTTP-nav] edit-categoryGroup=2")
    _ajax_select_change(client, CREATE_REQUEST_URL, container2, "edit-categoryGroup", "2", client.state.url)
    log4("   [HTTP-nav] edit-category=27")
    _ajax_select_change(client, CREATE_REQUEST_URL, container2, "edit-category", "27", client.state.url)
    log4("   [HTTP-nav] edit-character=1")
    _ajax_select_change(client, CREATE_REQUEST_URL, container2, "edit-character", "1", client.state.url)

    # Чекбокс "Дело упрощенного производства" — обычный checkbox без своего
    # ajax-запроса, включаем его прямо в состояние перед кликом «Добавить».
    log4("   [HTTP-nav] click renderAddPersonModalDialog («Добавить участника процесса»)")
    _click_js_triggered_button(
        client, CREATE_REQUEST_URL, container2, "renderAddPersonModalDialog", client.state.url,
        extra_state={f"{container2}:edit-simpleProcess": "on"},
    )

    next_button_id = _extract_plain_ajax_button_id(client.state.html, ":pp-side")
    if not next_button_id:
        debug_path = _dump_debug_html(client.state.html, "missing_next_side_button")
        raise RuntimeError(f"Не найдена кнопка «Далее» в форме выбора стороны. HTML: {debug_path}")

    log4("   [HTTP-nav] click «Далее» (выбор стороны)")
    container3 = next_button_id.rsplit(":", 1)[0]
    state = _serialize_naming_container(client.state.html, container3)
    state.update(_ajax_meta(next_button_id, event="click"))
    client.post_form(CREATE_REQUEST_URL, state, referer=client.state.url)

    log4("   [HTTP-nav] извлекаю контекст формы участника (fillPersonData)")
    return _extract_person_form_context(client.state.html)


def _fetch_person_by_iin_http(http_session, referer_url, form_id, ajax_component, viewstate, iin, timeout=30):
    """Точная реплика запроса fillPersonData из HAR. Возвращает (data_dict, new_viewstate)."""
    data = {
        form_id: form_id,
        f"{form_id}:person-iin": iin,
        f"{form_id}:person-surname": "",
        f"{form_id}:person-firstname": "",
        f"{form_id}:person-patronymic": "",
        f"{form_id}:person-livePlace": "",
        f"{form_id}:person-workPlace": "",
        f"{form_id}:person-phone": "",
        f"{form_id}:person-email": "",
        "javax.faces.ViewState": viewstate,
        "javax.faces.source": ajax_component,
        "javax.faces.partial.execute": f"{ajax_component} @component",
        "javax.faces.partial.render": "@component",
        "param1": f"{form_id}:person-iin",
        "org.richfaces.ajax.component": ajax_component,
        ajax_component: ajax_component,
        "rfExt": "null",
        "AJAX:EVENTS_COUNT": "1",
        "javax.faces.partial.ajax": "true",
    }
    headers = {
        "Accept": "*/*",
        "Faces-Request": "partial/ajax",
        "Referer": referer_url,
        "Origin": BASE,
    }
    resp = http_session.post(
        CREATE_REQUEST_URL,
        data=data,
        headers=headers,
        timeout=timeout,
    )
    resp.raise_for_status()

    low = resp.text[:2000].lower()
    if "войти" in low or "/index.xhtml" in resp.url.lower():
        raise RuntimeError("HTTP_SESSION_DEAD")

    updates, new_viewstate = _parse_partial_response(resp.text)
    if not updates:
        raise RuntimeError("HTTP_EMPTY_RESPONSE")

    result = {
        "sur":  _find_update_by_suffix(updates, ":person-surname"),
        "name": _find_update_by_suffix(updates, ":person-firstname"),
        "patr": _find_update_by_suffix(updates, ":person-patronymic"),
        "live": _find_update_by_suffix(updates, ":person-livePlace"),
    }

    if not any(result.values()):
        # Сервер иногда вместо полей person-* возвращает processPersonListCode
        # (скрытое поле + JS $(...).val('КОД')) — это значит "этот ИИН уже
        # добавлен в текущее дело/участники", а не "не найдено в ГБДУ".
        # Повторный soft-restart формы здесь бесполезен: пока дело то же,
        # сервер снова вернёт этот же код.
        m_dup = re.search(
            r"processPersonListCode[^']*'\]\)\.val\('(\d+)'\)",
            resp.text,
        )
        if m_dup:
            raise RuntimeError(f"HTTP_ALREADY_IN_PROCESS:{m_dup.group(1)}")

    return result, (new_viewstate or viewstate)


def parse_people_via_pure_http():
    """
    Полностью HTTP-версия блока 4 — Selenium используется РОВНО ОДИН РАЗ,
    внутри selenium_login_get_session(), только чтобы залогиниться и
    получить cookies. Вся навигация (open_person_form_via_http) и сам цикл
    поиска по сотням ИИН — requests, как в блоке 2 podacha_iska_v2.py.
    При обрыве сессии/деле "восстановление" = новый Selenium-логин +
    заново пройденная HTTP-навигация (никаких кликов/переоткрытий DOM).
    """
    MAX_TOTAL_ATTEMPTS_PER_ROW = 6
    RETRY_ROW_PAUSE = 1.0
    HTTP_TIMEOUT = 30
    PAUSE_BETWEEN_HTTP_CALLS = 0.25
    MAX_CONSECUTIVE_ERRORS = 3

    log4("=== START parse_people_via_pure_http (Selenium только для логина) ===")
    log4(f"Файл-источник: {INPUT_XLSX}")

    wb = load_workbook(INPUT_XLSX, data_only=True)
    try:
        ws = wb["Отмены"]
    except KeyError:
        ws = wb.active

    header_o = ws.cell(row=1, column=15)
    if not header_o.value or str(header_o.value).strip() == "":
        header_o.value = "Актуальный адрес с СК"

    header_p = ws.cell(row=1, column=16)
    if not header_p.value or str(header_p.value).strip() == "":
        header_p.value = "Регион"

    def _new_client() -> SudHttpClient:
        log4("   🔐 Логин через Selenium (единственный раз за сессию)...")
        session = selenium_login_get_session()
        client = SudHttpClient(session)
        log4("   🧭 HTTP-навигация до формы участника...")
        form_id, ajax_component = open_person_form_via_http(client)
        log4(f"   ✅ Форма открыта: form_id={form_id}, ajax={ajax_component}")
        return client, form_id, ajax_component

    client = form_id = ajax_component = None
    last_error = None
    for attempt in range(1, 4):
        try:
            client, form_id, ajax_component = _new_client()
            break
        except Exception as e:
            last_error = e
            log4(f"   ❌ Попытка {attempt}/3 открыть форму не удалась: {type(e).__name__}: {e}")
            logging.exception("Открытие формы участника не удалось")
            if attempt < 3:
                time.sleep(3)
    if client is None:
        raise RuntimeError(f"Не удалось открыть форму участника после 3 попыток: {last_error}")

    failed_rows = []
    processed = 0
    consecutive_errors = 0

    try:
        for r in range(2, ws.max_row + 1):
            src_iin = _norm_iin(ws.cell(row=r, column=4).value)
            if not src_iin or len(src_iin) != 12:
                continue

            processed += 1
            t_row_start = time.time()
            log4(f"#{processed} (Excel row {r}) → ИИН: {src_iin}")

            success = False
            total_attempts = 0

            while not success and total_attempts < MAX_TOTAL_ATTEMPTS_PER_ROW:
                total_attempts += 1

                if consecutive_errors >= MAX_CONSECUTIVE_ERRORS:
                    log4(f"   ⚠ {consecutive_errors} ошибок подряд — новый логин и новое дело")
                    try:
                        client, form_id, ajax_component = _new_client()
                        consecutive_errors = 0
                    except Exception as e:
                        log4(f"   ❌ Не удалось перелогиниться/открыть форму: {e}")

                log4(f"   HTTP-запрос fillPersonData для ИИН {src_iin}")
                try:
                    data, new_viewstate = _fetch_person_by_iin_http(
                        client.session, client.state.url, form_id, ajax_component,
                        client.state.viewstate, src_iin, timeout=HTTP_TIMEOUT,
                    )
                    client.state.viewstate = new_viewstate
                except RuntimeError as e:
                    if str(e).startswith("HTTP_ALREADY_IN_PROCESS"):
                        log4(
                            f"   ⚠ ИИН {src_iin} уже числится участником текущего дела "
                            f"(сервер вернул {e}) — нужен новый логин/дело"
                        )
                        consecutive_errors = MAX_CONSECUTIVE_ERRORS
                    else:
                        log4(f"   ⚠ {e}")
                        consecutive_errors += 1
                except requests.RequestException as e:
                    log4(f"   ⚠ ошибка HTTP-запроса: {e}")
                    consecutive_errors += 1
                else:
                    changed = any(data.values())
                    log4(f"   ответ: {data} | changed={changed}")
                    if changed:
                        live = data.get("live", "")
                        ws.cell(row=r, column=15, value=live)
                        region = extract_region_from_address(live)
                        ws.cell(row=r, column=16, value=region)
                        log4(f"   → SAVE: row {r}, address='{live}', region='{region}'")
                        success = True
                        consecutive_errors = 0
                    else:
                        log4("   данных нет")
                        consecutive_errors += 1

                if success:
                    break

                if total_attempts < MAX_TOTAL_ATTEMPTS_PER_ROW:
                    log4(f"   ↺ строка {r} не получилась, повторяю (попытка {total_attempts + 1}/{MAX_TOTAL_ATTEMPTS_PER_ROW})…")
                    time.sleep(RETRY_ROW_PAUSE)

            if not success:
                failed_rows.append(r)
                log4(f"   → SAVE (error): row {r}, address='' (ERROR_NO_DATA после {MAX_TOTAL_ATTEMPTS_PER_ROW} попыток) — перехожу к следующей строке")

            time.sleep(PAUSE_BETWEEN_HTTP_CALLS)
            elapsed = time.time() - t_row_start
            if elapsed < 0.3:
                time.sleep(0.3 - elapsed)

            if processed % AUTOSAVE_EVERY == 0:
                _safe_save(wb, OUT_XLSX)

    finally:
        try:
            _safe_save(wb, OUT_XLSX)
        except Exception:
            pass

    return processed, failed_rows


# ========= БЛОК 5: логика выбора суда =========
ADDR_COL_NAME   = "Актуальный адрес с Судебного кабинета"
REGION_COL_NAME = "Регион с Судебного кабинета"
COURT_COL_NAME  = "Судебный орган с Судебного кабинета"

def normalize_kz_text(s: str) -> str:
    s = unicodedata.normalize("NFC", str(s))
    s = s.lower()
    replace_map = str.maketrans({
        "қ": "к", "ғ": "г", "ң": "н", "ү": "у", "ұ": "у",
        "ө": "о", "ә": "а", "һ": "х", "і": "и", "ё": "е",
    })
    s = s.translate(replace_map)
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def norm_region(s: str) -> str:
    s = normalize_kz_text(s)
    s = s.replace("город ", "").replace("г. ", "").replace("г.", "")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def norm_locality_raw(s: str) -> str:
    s = normalize_kz_text(s)
    s = s.replace("р-н", "район").replace("р н", "район")
    s = re.sub(r"\s+", " ", s)
    remove_words = ["район", "аудан", "область", "облысы", "город", "қала", "поселок", "посёлок", "село", "ауыл"]
    for w in remove_words:
        s = s.replace(w, " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def locality_stem(loc_key: str) -> str:
    if not loc_key:
        return ""
    s = loc_key
    s = re.sub(r"(ский|ская|ское|скии|скій)$", "", s)
    s = re.sub(r"(ый|ий)$", "", s)
    s = s.strip()
    return s if s else loc_key

def simplify_court_name(name: str) -> str:
    s = normalize_kz_text(name)
    s = re.sub(r"\(.*?\)", " ", s)
    # Обрезаем всё начиная со слова "област..." — у части судов название
    # региона стоит ПОСЛЕ "области" (например "суд области Жетісу"), и без
    # обрезки оно оставалось в ключе, ложно совпадая с одноимённым районом.
    s = re.split(r"\bобласт\w*", s, maxsplit=1)[0]
    remove_words = [
        "районный суд", "районный  суд", "городской суд",
        "район", "района", "районный", "городской", "суд", "соты",
        "облысы", "города",
        "республики казахстан",
    ]
    for w in remove_words:
        s = s.replace(w, " ")
    # "рк" убираем только как отдельное слово (аббревиатура "РК"), а не
    # подстрокой — иначе она вырезалась и из "туРКестанской".
    s = re.sub(r"\bрк\b", " ", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def str_similarity(a: str, b: str) -> float:
    a = str(a)
    b = str(b)
    if not a or not b:
        return 0.0
    try:
        from rapidfuzz import fuzz
        return float(fuzz.token_set_ratio(a, b))
    except Exception:
        return SequenceMatcher(None, a, b).ratio() * 100.0

def extract_locality_from_address(address: str):
    if pd.isna(address):
        return None
    parts = [p.strip() for p in str(address).split(",") if p.strip()]
    if len(parts) >= 2:
        return parts[1]
    return parts[0] if parts else None

_TWO_COURT_SETTLEMENT_CACHE: dict = {}


def _load_two_court_settlement_map(file_path: str):
    """
    Общий парсер справочников вида "Енбекшиказахский суд.xlsx" /
    "КРГ Шетский.xlsx" — район поделён между ДВУМЯ судами по сёлам/округам,
    и обычная логика по названию района/области этого не различает.

    Колонка A — сёла/округа суда, названного в шапке колонки A, колонка B —
    суда из шапки колонки B. Ячейка может содержать название округа плюс
    (в скобках) список сёл этого округа — в т.ч. НЕСКОЛЬКО скобочных групп в
    одной ячейке — разбираем и округ, и каждое село из каждой группы, т.к. в
    адресе клиента может встретиться любое из этих названий.

    Возвращает (court_a_name, court_b_name, {нормализованное_название: суд}).
    Результат кэшируется по пути файла.
    """
    if file_path in _TWO_COURT_SETTLEMENT_CACHE:
        return _TWO_COURT_SETTLEMENT_CACHE[file_path]

    wb = load_workbook(file_path, data_only=True)
    ws = wb.active

    header = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    court_a_name = str(header[0]).strip()
    court_b_name = str(header[1]).strip()

    settlement_map = {}

    def add_keywords(raw_value, court_name):
        if not raw_value:
            return
        text = str(raw_value)
        # Все скобочные группы вырезаются из "базового" названия округа и
        # разбираются отдельно — в справочнике встречаются ячейки с
        # НЕСКОЛЬКИМИ группами в скобках подряд, например
        # "Шетский сельский округ (село ныне Унрек), (село Куттыбай, ...)".
        paren_groups = re.findall(r"\(([^)]*)\)", text)
        base = re.sub(r"\([^)]*\)", "", text)
        names = [base] + [item for group in paren_groups for item in group.split(",")]
        for name in names:
            name = name.strip(" \xa0,")
            if not name:
                continue
            # "село ныне Унрек" = "село, ныне [переименованное в] Унрек" —
            # слово "ныне" не часть названия.
            name = re.sub(r"\bныне\b\s*", "", name, flags=re.IGNORECASE).strip()
            for prefix in ("село ", "город ", "поселок ", "посёлок ", "ферма ",
                           "сельский округ ", "поселковый округ "):
                if name.lower().startswith(prefix):
                    name = name[len(prefix):].strip()
            name = re.sub(r"\s*(сельский|поселковый)\s+округ\s*$", "", name, flags=re.IGNORECASE).strip()
            key = normalize_kz_text(name).upper() if name else ""
            if key:
                settlement_map.setdefault(key, court_name)

    for row in ws.iter_rows(min_row=2, values_only=True):
        add_keywords(row[0] if len(row) > 0 else None, court_a_name)
        add_keywords(row[1] if len(row) > 1 else None, court_b_name)

    result = (court_a_name, court_b_name, settlement_map)
    _TWO_COURT_SETTLEMENT_CACHE[file_path] = result
    return result


def _pick_court_by_settlement(addr_norm: str, file_path: str, default_court: str, district_label: str) -> str:
    """
    Ищет в адресе конкретное село/округ из справочника `file_path`
    (см. _load_two_court_settlement_map); если не нашли — суд по умолчанию.
    """
    try:
        _, _, settlement_map = _load_two_court_settlement_map(file_path)
    except Exception as e:
        log4(f"⚠ Не удалось прочитать справочник района ({district_label}): {e}")
        return default_court

    for keyword, court_name in settlement_map.items():
        # Границы слова обязательны: например ключ "ЕНБЕК" (село) иначе
        # ложно совпадает внутри самого названия района "ЕНБЕКШИКАЗАХСКИЙ",
        # которое встречается в КАЖДОМ адресе этого района.
        if re.search(r"\b" + re.escape(keyword) + r"\b", addr_norm):
            return court_name
    return default_court


def _pick_enbekshi_court(addr_norm: str) -> str:
    return _pick_court_by_settlement(
        addr_norm, FILE_ENBEKSHI,
        default_court="Енбекшиказахский районный суд",
        district_label="Енбекшиказахский район",
    )


def _pick_shet_court(addr_norm: str) -> str:
    return _pick_court_by_settlement(
        addr_norm, FILE_SHET,
        default_court="Шетский районный суд Карагандинской области",
        district_label="Шетский район",
    )


def pick_court(region_value, address_value, courts_by_region):
    addr_norm = ""
    if not pd.isna(address_value):
        addr_norm = normalize_kz_text(address_value).upper()

        if "КОСТАНАЙСКАЯ ОБЛАСТЬ, КОСТАНАЙСКИЙ РАЙОН" in addr_norm and "ТОБЫЛ" in addr_norm:
            return "Костанайский межрайонный суд Костанайской области"
        if "КОСТАНАЙСКАЯ ОБЛАСТЬ, КОСТАНАЙ," in addr_norm:
            return "Костанайский городской суд Костанайской области (Гражданские дела)"
        if "АТЫРАУСКАЯ ОБЛАСТЬ" in addr_norm and "АТЫРАУ," in addr_norm:
            return "Атырауский городской суд Атырауской области (Гражданские дела)"
        if "ТУРКЕСТАНСКАЯ ОБЛАСТЬ, САУРАНСКИЙ РАЙОН" in addr_norm:
            return "Кентауский городской суд Туркестанской области (Общая юрисдикция)"
        if "ПАВЛОДАРСКАЯ ОБЛАСТЬ, ПАВЛОДАР," in addr_norm:
            return "Межрайонный суд по гражданским делам города Павлодара"
        if "ПАВЛОДАРСКАЯ ОБЛАСТЬ, ПАВЛОДАРСКИЙ РАЙОН" in addr_norm:
            return "Межрайонный суд по гражданским делам города Павлодара"
        if "КЫЗЫЛОРДИНСКАЯ ОБЛАСТЬ" in addr_norm and "КЫЗЫЛОРДА" in addr_norm:
            return "Кызылординский городской суд Кызылординской области (Гражданские дела)"
        if "КАРАГАНДИНСКАЯ ОБЛАСТЬ, САРАНЬ" in addr_norm:
            return "Саранский городской суд Карагандинской области (Общая юрисдикция)"
        if "КОСТАНАЙСКАЯ ОБЛАСТЬ, АРКАЛЫК" in addr_norm:
            return "Аркалыкский городской суд Костанайской области (Общая юрисдикция)"
        if "КАРАГАНДИНСКАЯ ОБЛАСТЬ" in addr_norm and "ШЕТСК" in addr_norm:
            return _pick_shet_court(addr_norm)
        if "КАРАГАНДИНСКАЯ ОБЛАСТЬ" in addr_norm and "КАРАГАНДА," in addr_norm:
            return "Необходимо вручную проставить суд! Так как в данном городе несколько судов."
        if "АКТЮБИНСКАЯ ОБЛАСТЬ" in addr_norm and "АКТОБЕ," in addr_norm:
            # Город Актобе разбит на 2 внутригородских района, которые в
            # справочнике судов так и называются — "Алматы" и "Астана" (не
            # путать с одноимёнными городами/областями).
            if "АЛМАТЫ" in addr_norm:
                return 'Суд № 3 города Актобе Актюбинской области, район "Алматы" (Гражданские дела)'
            if "АСТАНА" in addr_norm:
                return 'Суд города Актобе Актюбинской области, район "Астана" (Гражданские дела)'
            return "Необходимо вручную проставить суд! Так как в данном городе несколько судов."
        if "ЖАМБЫЛСКАЯ ОБЛАСТЬ" in addr_norm:
            if "ТАРАЗ," in addr_norm:
                return "Таразский городской суд Жамбылской области (Гражданские дела)"
            if "КОРДАЙСКИЙ РАЙОН" in addr_norm:
                return "Кордайский районный суд Жамбылской области"
            if "МЕРКЕНСКИЙ РАЙОН" in addr_norm:
                return "Меркенский районный суд Жамбылской области"
            if "ЖУАЛИНСКИЙ РАЙОН" in addr_norm:
                return "Жуалинский районный суд Жамбылской области"
            if "БАЙЗАКСКИЙ РАЙОН" in addr_norm:
                return "Байзакский районный суд Жамбылской области"
            if "ШУСКИЙ РАЙОН" in addr_norm:
                return "Шуский районный суд Жамбылской области"
        if "МАНГИСТАУСКАЯ ОБЛАСТЬ" in addr_norm and "АКТАУ," in addr_norm:
            return "Актауский городской суд Мангистауской области (Гражданские дела)"
        if "ВОСТОЧНО-КАЗАХСТАНСКАЯ ОБЛАСТЬ" in addr_norm and "УСТЬ-КАМЕНОГОРСК," in addr_norm:
            return "Межрайонный суд по гражданским делам города Усть-Каменогорска"
        if "ОБЛАСТЬ ЖЕТІСУ" in addr_norm and "ТАЛДЫКОРГАН," in addr_norm:
            return "Талдыкорганский городской суд области Жетісу (Гражданские дела)"
        if "АЛМАТИНСКАЯ ОБЛАСТЬ" in addr_norm and "ЕНБЕКШИКАЗАХ" in addr_norm:
            return _pick_enbekshi_court(addr_norm)

    if pd.isna(region_value):
        return None

    reg_key = norm_region(region_value)
    candidates_df = courts_by_region.get(reg_key)
    if candidates_df is None or candidates_df.empty:
        return None

    raw_loc = extract_locality_from_address(address_value)
    if not raw_loc:
        return candidates_df["Суды"].iloc[0]

    loc_key = norm_locality_raw(raw_loc)
    if not loc_key:
        return candidates_df["Суды"].iloc[0]

    stem_val = locality_stem(loc_key)
    filtered = candidates_df
    if stem_val:
        mask = filtered["court_key"].apply(lambda ck: stem_val in ck if isinstance(ck, str) else False)
        if mask.any():
            filtered = filtered[mask]

    best_score = -1.0
    best_court = None
    for _, row in filtered.iterrows():
        score = str_similarity(loc_key, row["court_key"])
        if score > best_score:
            best_score = score
            best_court = row["Суды"]

    return best_court


def fill_courts_column():
    print("Читаю файл отчёта:", FILE_PEOPLE)
    df_people = pd.read_excel(FILE_PEOPLE, sheet_name="Отмены")
    print("Столбцы в отчёте:", list(df_people.columns))

    for col in (ADDR_COL_NAME, REGION_COL_NAME, COURT_COL_NAME):
        if col not in df_people.columns:
            raise ValueError(f"В файле нет столбца '{col}'")

    print("Читаю файл судов:", FILE_COURTS)
    df_courts = pd.read_excel(FILE_COURTS, sheet_name="Возврат")
    df_courts["Области"] = df_courts["Области"].ffill()
    df_courts = df_courts.dropna(subset=["Суды"]).copy()
    df_courts["region_key"] = df_courts["Области"].apply(norm_region)
    df_courts["court_key"]  = df_courts["Суды"].apply(simplify_court_name)

    courts_by_region = {
        reg_key: grp.reset_index(drop=True)
        for reg_key, grp in df_courts.groupby("region_key")
    }

    print("Определяю суды для каждой строки...")
    courts_series = df_people.apply(
        lambda row: pick_court(row[REGION_COL_NAME], row[ADDR_COL_NAME], courts_by_region),
        axis=1,
    ).reset_index(drop=True)

    wb = load_workbook(FILE_PEOPLE)
    try:
        ws = wb["Отмены"]
    except KeyError:
        ws = wb.active

    try:
        ws_templates = wb["Данные для шаблонов"]
    except KeyError:
        ws_templates = None
        print("Лист 'Отмены' не найден.")

    templates_court_col = None
    if ws_templates is not None:
        for c in range(1, ws_templates.max_column + 1):
            h = ws_templates.cell(row=1, column=c).value
            if h and "судебный орган" in str(h).lower():
                templates_court_col = c
                break

    ws["Q1"].value = COURT_COL_NAME

    # Строим словарь: уникальный номер → суд, из листа "Отмены"
    uniq_col_idx = None
    for c in range(1, ws.max_column + 1):
        h = ws.cell(row=1, column=c).value
        if h and "уникальный номер" in str(h).lower():
            uniq_col_idx = c
            break

    # Записываем суды в лист "Отмены" и строим словарь
    uniq_to_court = {}
    for i, court_name in enumerate(courts_series, start=2):
        ws[f"Q{i}"].value = court_name
        if uniq_col_idx:
            uniq_val = ws.cell(row=i, column=uniq_col_idx).value
            if uniq_val:
                uniq_to_court[str(uniq_val).strip()] = court_name

    # Записываем в "Данные для шаблонов" по совпадению уникального номера
    if ws_templates is not None and templates_court_col is not None:
        templates_uniq_col = None
        for c in range(1, ws_templates.max_column + 1):
            h = ws_templates.cell(row=1, column=c).value
            if h and "уникальный номер" in str(h).lower():
                templates_uniq_col = c
                break

        if templates_uniq_col is None:
            print("⚠ В листе 'Данные для шаблонов' не найден столбец 'Уникальный номер'")
        else:
            matched = 0
            for r in range(2, ws_templates.max_row + 1):
                uniq_val = ws_templates.cell(row=r, column=templates_uniq_col).value
                if uniq_val is None:
                    continue
                key = str(uniq_val).strip()
                if key in uniq_to_court:
                    ws_templates.cell(row=r, column=templates_court_col, value=uniq_to_court[key])
                    matched += 1
            print(f"  Совпадений найдено: {matched}")

    wb.save(FILE_PEOPLE)
    print("Готово! Значения судов записаны в:")
    print("  - лист 'Отмены', колонка Q")
    if ws_templates is not None and templates_court_col is not None:
        print("  - лист 'Данные для шаблонов', колонка с заголовком 'Судебный орган'")


# ========= БЛОК 6: логика определения УГД =========

# Колонки в листе "Отмены"
UGD_COL_IDX  = 12   # L — название УГД
BIN_COL_IDX  = 13   # M — БИН УГД
# Адрес (O=15), регион (P=16), суд (Q=17) уже заполнены к этому моменту

UGD_COL_HEADER  = "УГД"
BIN_COL_HEADER  = "БИН УГД"

REGION_ALIASES_UGD = {
    "АСТАНА": ["астана", "город астана", "нур султан", "нурсултан"],
    "АЛМАТЫ": ["алматы", "город алматы"],
    "ШЫМКЕНТ": ["шымкент", "город шымкент", "чимкент"],
    "ОБЛАСТЬ АБАЙ": ["область абай", "абайская область"],
    "АКМОЛИНСКАЯ ОБЛАСТЬ": ["акмолинская область"],
    "АКТЮБИНСКАЯ ОБЛАСТЬ": ["актюбинская область", "актобе область"],
    "АЛМАТИНСКАЯ ОБЛАСТЬ": ["алматинская область"],
    "АТЫРАУСКАЯ ОБЛАСТЬ": ["атырауская область"],
    "ВОСТОЧНО-КАЗАХСТАНСКАЯ ОБЛАСТЬ": [
        "вко", "восточно казахстанская область",
        "восточно казахстанская", "восточный казахстан",
    ],
    "ЖАМБЫЛСКАЯ ОБЛАСТЬ": ["жамбылская область"],
    "ЗАПАДНО-КАЗАХСТАНСКАЯ ОБЛАСТЬ": [
        "зко", "западно казахстанская область", "западно казахстанская",
    ],
    "КАРАГАНДИНСКАЯ ОБЛАСТЬ": ["карагандинская область"],
    "КОСТАНАЙСКАЯ ОБЛАСТЬ": ["костанайская область"],
    "КЫЗЫЛОРДИНСКАЯ ОБЛАСТЬ": ["кызылординская область"],
    "МАНГИСТАУСКАЯ ОБЛАСТЬ": ["мангистауская область"],
    "ПАВЛОДАРСКАЯ ОБЛАСТЬ": ["павлодарская область"],
    "СЕВЕРО-КАЗАХСТАНСКАЯ ОБЛАСТЬ": [
        "ско", "северо казахстанская область", "северо казахстанская",
    ],
    "ТУРКЕСТАНСКАЯ ОБЛАСТЬ": ["туркестанская область"],
    "УЛЫТАУ": ["улытау", "область улытау"],
}


def _ugd_norm(text) -> str:
    """Нормализация текста для УГД-поиска (аналог norm() из ugd-скрипта)."""
    if text is None:
        return ""
    text = str(text).lower().strip()
    replace_map = {
        "ё": "е", "ә": "а", "і": "и", "ң": "н",
        "ғ": "г", "ү": "у", "ұ": "у", "қ": "к",
        "һ": "х", "ө": "о",
    }
    for k, v in replace_map.items():
        text = text.replace(k, v)
    text = text.replace("г.", "город ")
    text = text.replace("р-н", "район")
    text = text.replace("р/н", "район")
    text = re.sub(r"[^а-яa-z0-9]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _make_bin(value) -> str:
    if value is None:
        return ""
    value = str(value).strip()
    value = re.sub(r"\.0$", "", value)
    value = re.sub(r"\D", "", value)
    return value.zfill(12) if value else ""


def _ugd_stem(word: str) -> str:
    word = _ugd_norm(word)
    endings = [
        "скому", "цкому", "ского", "цкого",
        "ский", "цкий", "ская", "цкая",
        "ской", "цкой", "ском", "цком",
        "ному", "ного", "ный", "ная", "ной", "ном",
        "кому", "кого", "кий", "кая", "кой", "ком",
        "ому", "его", "ого", "ему",
        "инскому", "инского", "инский", "инская", "инской",
        "лыскому", "лыского", "лыский",
        "ин", "ск", "ий", "ый", "ая", "ой", "ом", "ым", "им",
    ]
    for e in endings:
        if word.endswith(e) and len(word) > len(e) + 3:
            word = word[:-len(e)]
            break
    manual = {
        "сарыарка": "сарыарк", "сарыаркин": "сарыарк",
        "келес": "келес", "келесск": "келес",
        "бородулихин": "бородулих", "бородулих": "бородулих",
        "кзылкогин": "кызылког", "кызылкогин": "кызылког", "кызылкугин": "кызылког",
        "енбекшиказах": "енбекшиказах", "енбекшин": "енбекшин",
        "каратау": "каратау", "ауыраль": "ауэзов",
    }
    return manual.get(word, word)


def _ugd_words_from_text(text: str) -> list:
    return [_ugd_stem(x) for x in _ugd_norm(text).split() if len(_ugd_stem(x)) >= 4]


def _ugd_detect_region(region_text, address_text) -> str:
    combined = _ugd_norm(str(region_text or "") + " " + str(address_text or ""))
    best_region = None
    best_len = 0
    for region, aliases in REGION_ALIASES_UGD.items():
        for alias in aliases:
            a = _ugd_norm(alias)
            if a in combined and len(a) > best_len:
                best_region = region
                best_len = len(a)
    return best_region


def _extract_district_from_address_ugd(address: str) -> str:
    text = _ugd_norm(address)
    words = text.split()
    for i, w in enumerate(words):
        if w in ["район", "аудан", "ауданы"]:
            if i > 0:
                return _ugd_stem(words[i - 1])
    return ""


def _extract_district_from_court_ugd(court: str) -> str:
    text = _ugd_norm(court)
    words = text.split()
    for i, w in enumerate(words):
        if w in ["районный", "район", "аудандык", "ауданы"]:
            if i > 0:
                return _ugd_stem(words[i - 1])
    return ""


def _extract_ugd_district(ugd_name: str) -> str:
    text = _ugd_norm(ugd_name)
    words = text.split()
    for i, w in enumerate(words):
        if w in ["району", "район", "ауданы", "аудан"]:
            if i > 0:
                return _ugd_stem(words[i - 1])
    return ""


def _is_city_ugd(ugd_name: str) -> bool:
    text = _ugd_norm(ugd_name)
    city_words = [
        "по город",           # покрывает "по г.Тараз" → "по город тараз"
        "по атырау",
        "по туркестан",
        "астана жана кала",
        "онтустик",
        "парк информационных технологии",
        "морпорт",
        "по г аксу",
        "по г балхаш",
        "по г приозерск",
        "по г шахтинск",
        "по г темиртау",
        "по г сарань",
        "по г риддер",
        "по г семей",
    ]
    return any(_ugd_norm(x) in text for x in city_words)


def _ugd_similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio()


def _load_ugd_mapping() -> list:
    """Читает справочник БИН(УГД).xlsx и возвращает список словарей."""
    print("Читаю справочник УГД:", FILE_UGD)
    dict_wb = load_workbook(FILE_UGD, data_only=True)
    dict_ws = dict_wb.active

    mapping = []
    current_region = None

    for row in dict_ws.iter_rows(min_row=1, values_only=True):
        col_a = row[0]
        col_b = row[1] if len(row) > 1 else None
        col_c = row[2] if len(row) > 2 else None

        # Строка-заголовок региона: только колонка A заполнена
        if col_a and not col_b and not col_c:
            current_region = str(col_a).strip().upper()
            continue

        if col_b and col_c and current_region:
            ugd_name = str(col_c).strip()
            district_key = _extract_ugd_district(ugd_name)
            mapping.append({
                "region":       current_region,
                "bin":          _make_bin(col_b),
                "ugd":          ugd_name,
                "district_key": district_key,
                "ugd_words":    _ugd_words_from_text(ugd_name),
                "is_city":      _is_city_ugd(ugd_name),
            })

    print(f"  Загружено УГД-записей: {len(mapping)}")
    return mapping


def _pick_ugd(address_value, region_value, court_value, mapping: list) -> tuple:
    """
    ИСПРАВЛЕННАЯ ВЕРСИЯ №2.

    Проблема была в том, что в некоторых областях есть район с ТЕМ ЖЕ
    названием, что и сама область (например "Жамбылская область" →
    район "Жамбылский", или "Мангистауская область" → район
    "Мангистауский"). Из-за этого нечёткий поиск по словам (бывший Шаг 4)
    ловил слово "жамбыл"/"мангистау" из названия ОБЛАСТИ в адресе и
    ошибочно матчил его на одноимённый РАЙОН, даже если в адресе явно
    был указан другой город (Тараз, Актау и т.д.).

    Исправлено:
      1) Явный поиск города (бывший Шаг 5) теперь идёт РАНЬШЕ нечёткого
         поиска по словам (бывший Шаг 4) — точное совпадение приоритетнее.
      2) В нечётком поиске по словам слова, происходящие из названия
         самой области, исключаются из сравнения, чтобы не давать
         ложное совпадение "область == одноимённый район".
    """
    region = _ugd_detect_region(region_value, address_value)

    # Если во всём определённом регионе всего один УГД — берём его сразу,
    # без сверки по району. Иначе поиск по району ломается там, где в
    # справочнике единственная запись региона названа по конкретному
    # району (а не по городу) — например "УГД по Есильскому району" для
    # Астаны или "УГД по Аль-Фарабийскому району" для Шымкента: тогда
    # район из адреса другого клиента того же города просто не совпадёт
    # ни с чем, хотя это тот же единственный УГД. Раньше это было
    # захардкожено только для Астаны — обобщаем на любой такой регион.
    if region:
        region_candidates = [x for x in mapping if x["region"] == region]
        if len(region_candidates) == 1:
            found = region_candidates[0]
            return found["ugd"], found["bin"]

    address_district = _extract_district_from_address_ugd(str(address_value or ""))
    court_district   = _extract_district_from_court_ugd(str(court_value or ""))

    search_keys = []
    if address_district:
        search_keys.append(address_district)
    if court_district and court_district not in search_keys:
        search_keys.append(court_district)

    candidates = mapping
    if region:
        candidates = [x for x in mapping if x["region"] == region]

    found = None

    # ---- ШАГ 1: точное совпадение района ----
    for key in search_keys:
        for item in candidates:
            if item["is_city"]:
                continue
            if item["district_key"] == key:
                found = item
                break
        if found:
            break

    # ---- ШАГ 2: частичное совпадение района ----
    if not found:
        for key in search_keys:
            for item in candidates:
                if item["is_city"]:
                    continue
                ugd_key = item["district_key"]
                if ugd_key and key and (key in ugd_key or ugd_key in key):
                    found = item
                    break
            if found:
                break

    # ---- ШАГ 3: fuzzy-поиск района ----
    if not found:
        best_item  = None
        best_score = 0.0
        for key in search_keys:
            for item in candidates:
                if item["is_city"]:
                    continue
                ugd_key = item["district_key"]
                if not ugd_key:
                    continue
                score = _ugd_similarity(key, ugd_key)
                if score > best_score:
                    best_score = score
                    best_item  = item
        if best_score >= 0.72:
            found = best_item

    # ---- ШАГ 4 (было Шаг 5): городской УГД по явному упоминанию города ----
    # Перенесено ВЫШЕ нечёткого поиска по словам, т.к. явное совпадение
    # города надёжнее и не должно перебиваться случайным совпадением
    # слова "область" с одноимённым районом.
    if not found:
        addr_norm = _ugd_norm(str(address_value or ""))
        city_keywords = {
            "усть каменогорск": "усть каменогорск",
            "семей": "семей",
            "риддер": "риддер",
            "павлодар": "павлодар",
            # "Экибастуз"/"Шахтинск" в справочнике БИН(УГД).xlsx записаны в
            # дательном падеже ("...Экибастузу", "...Шахтинску") — прямое
            # совпадение слова "экибастуз"/"шахтинск" из адреса клиента с ЭТИМ
            # словом справочника ниже (ШАГ 6) не срабатывает, и без этих ключей
            # адрес проваливался в ШАГ 6, где "павлодар"/"сарань" ложно
            # совпадали как подстрока названия области/первый попавшийся
            # городской УГД. city_search ("экибастуз"/"шахтинск") ищется КАК
            # ПОДСТРОКА внутри названия УГД, поэтому падеж справочника тут не
            # мешает — совпадёт и с "...экибастузу", и с "...шахтинску".
            "экибастуз": "экибастуз",
            "шахтинск": "шахтинск",
            "аксу": "аксу",
            "актобе": "актобе",
            "атырау": "атырау",
            "актау": "актау",
            "кызылорда": "кызылорда",
            "талдыкорган": "талдыкорган",
            "туркестан": "туркестан",
            "тараз": "тараз",
            "костанай": "костанай",
            "петропавловск": "петропавловск",
            "кокшетау": "кокшетау",
            "уральск": "уральск",
            "шымкент": "шымкент",
            "алматы": "алматы",
            "астана": "астана",
        }
        for city_kw, city_search in city_keywords.items():
            if re.search(rf'\b{re.escape(city_kw)}\b', addr_norm):
                for item in candidates:
                    if city_search in _ugd_norm(item["ugd"]) and item["is_city"]:
                        found = item
                        break
                break

    # ---- ШАГ 5 (было Шаг 4): поиск по словам из адреса + суда ----
    # Исключаем слова, происходящие из названия самой ОБЛАСТИ — иначе
    # "Жамбылская область" всегда матчится на "Жамбылский район", а
    # "Мангистауская область" — на "Мангистауский район".
    if not found:
        region_words = set()
        if region:
            region_words = set(_ugd_words_from_text(region))

        all_words = set(_ugd_words_from_text(
            str(address_value or "") + " " + str(court_value or "")
        ))
        all_words -= region_words  # <-- ключевое исправление

        best_item    = None
        best_matches = 0
        for item in candidates:
            if item["is_city"]:
                continue
            matches = len(all_words.intersection(set(item["ugd_words"])))
            if matches > best_matches:
                best_matches = matches
                best_item    = item
        if best_matches >= 1:
            found = best_item

    # ---- ШАГ 6: последний fallback — любой городской УГД по совпадению слов ----
    if not found:
        addr_norm_ugd = _ugd_norm(str(address_value or ""))
        best_city_item = None
        best_city_score = 0.0
        for item in candidates:
            if not item["is_city"]:
                continue
            ugd_norm = _ugd_norm(item["ugd"])
            ugd_words = [w for w in ugd_norm.split() if len(w) >= 4 and w not in ["город", "района", "район"]]
            score = sum(1 for w in ugd_words if w in addr_norm_ugd)
            if score > best_city_score:
                best_city_score = score
                best_city_item = item

        if best_city_item and best_city_score > 0:
            found = best_city_item
        elif candidates:
            for item in candidates:
                if item["is_city"]:
                    found = item
                    break

    if found:
        return found["ugd"], found["bin"]
    return "", ""


def fill_ugd_column():
    print("\n=== Определяю УГД и БИН УГД ===")

    ugd_mapping = _load_ugd_mapping()

    wb = load_workbook(FILE_PEOPLE)
    try:
        ws = wb["Отмены"]
    except KeyError:
        ws = wb.active

    if not ws.cell(row=1, column=UGD_COL_IDX).value:
        ws.cell(row=1, column=UGD_COL_IDX).value = UGD_COL_HEADER
    if not ws.cell(row=1, column=BIN_COL_IDX).value:
        ws.cell(row=1, column=BIN_COL_IDX).value = BIN_COL_HEADER

    uniq_col_idx = None
    for c in range(1, ws.max_column + 1):
        h = ws.cell(row=1, column=c).value
        if h and "уникальный номер" in str(h).lower():
            uniq_col_idx = c
            break

    uniq_to_ugd  = {}
    uniq_to_addr = {}  # НОВОЕ

    for r in range(2, ws.max_row + 1):
        address_val = ws.cell(row=r, column=15).value   # O — Актуальный адрес с СК
        region_val  = ws.cell(row=r, column=16).value   # P — Регион
        court_val   = ws.cell(row=r, column=17).value   # Q — Судебный орган

        if not address_val and not region_val and not court_val:
            continue

        ugd_name, bin_val = _pick_ugd(address_val, region_val, court_val, ugd_mapping)

        ws.cell(row=r, column=UGD_COL_IDX).value = ugd_name
        ws.cell(row=r, column=BIN_COL_IDX).value = bin_val

        log4(f"  row {r}: УГД='{ugd_name}', БИН='{bin_val}'")

        if uniq_col_idx:
            uniq_val = ws.cell(row=r, column=uniq_col_idx).value
            if uniq_val:
                key = str(uniq_val).strip()
                uniq_to_ugd[key]  = (ugd_name, bin_val)
                uniq_to_addr[key] = address_val  # НОВОЕ

    # --- Синхронизация в лист "Данные для шаблонов" ---
    try:
        ws_templates = wb["Данные для шаблонов"]
    except KeyError:
        ws_templates = None
        print("⚠ Лист 'Данные для шаблонов' не найден — синхронизация УГД пропущена.")

    if ws_templates is not None and uniq_to_ugd:
        t_uniq_col = None
        t_ugd_col  = None
        t_bin_col  = None
        t_addr_col = None  # НОВОЕ

        for c in range(1, ws_templates.max_column + 1):
            h = str(ws_templates.cell(row=1, column=c).value or "").lower().strip()
            if "уникальный номер" in h:
                t_uniq_col = c
            elif h == "угд":
                t_ugd_col = c
            elif "бин угд" in h:
                t_bin_col = c
            elif "адрес регистрации" in h:  # НОВОЕ
                t_addr_col = c

        if t_ugd_col is None:
            t_ugd_col = ws_templates.max_column + 1
            ws_templates.cell(row=1, column=t_ugd_col).value = UGD_COL_HEADER
            print(f"  Создана колонка '{UGD_COL_HEADER}' в 'Данные для шаблонов' (col {t_ugd_col})")

        if t_bin_col is None:
            t_bin_col = ws_templates.max_column + 1
            ws_templates.cell(row=1, column=t_bin_col).value = BIN_COL_HEADER
            print(f"  Создана колонка '{BIN_COL_HEADER}' в 'Данные для шаблонов' (col {t_bin_col})")

        if t_uniq_col is None:
            print("⚠ В листе 'Данные для шаблонов' не найден столбец 'Уникальный номер' — синхронизация пропущена.")
        else:
            if t_addr_col is None:
                print("⚠ Колонка 'Адрес регистрации' не найдена в листе 'Данные для шаблонов'")

            matched = 0
            for r in range(2, ws_templates.max_row + 1):
                uniq_val = ws_templates.cell(row=r, column=t_uniq_col).value
                if uniq_val is None:
                    continue
                key = str(uniq_val).strip()
                if key in uniq_to_addr:
                    # Только адрес — УГД и БИН не дублируем
                    if t_addr_col:
                        ws_templates.cell(row=r, column=t_addr_col).value = uniq_to_addr[key]
                        matched += 1
            print(f"  Синхронизирован Адрес в 'Данные для шаблонов': {matched} строк")

            wb.save(FILE_PEOPLE)
            print("✅ УГД и БИН УГД записаны в:")
            print(f"  - лист 'Отмены', колонки L и M")
            if ws_templates is not None:
                print(f"  - лист 'Данные для шаблонов', колонки УГД, БИН УГД и Адрес регистрации")


# ========= ЗАГОЛОВКИ ProcessImport =========
PROCESS_IMPORT_HEADERS = [
    'Уникальный номер сделки',
    'Дата подачи заявления на выписку ИЛ',
    'Дата получения ИЛ',
    'Дата передачи ИЛ ЧСИ',
    'Комментарии по суду',
    'Дата подачи заявления на выдачу СП',
    'Дата получения СП',
    'Дата передачи СП ЧСИ',
    'Комментарий',
    'Представитель истца',
    'Номер судебного дела',
    'Судебный орган',
    'Категория дела',
    'Сумма иска',
    'Сумма государственной пошлины',
    'Дата отправки искового заявления',
    'Отклонено',
    'Причина отклонения заявления',
    'Зарегистрировано',
    'Судья',
    'Вынесено определение о возврате искового заявления',
    'Вынесено определение об утверждении соглашения об урегулировании спора',
    'Вынесено определение о рассмотрении дела в порядке упрощенного производства',
    'Вынесено решение первой инстанции',
    'Вынесен судебный приказ',
    'Определение об отмене решения в порядке упрощенного производства',
    'Ответственный',
    'Дата создания',
    'Название процесса',
    'Тип процесса',
    'Статус процесса',
]


# ========= ORCHESTRATOR =========
def run(df_main):
    global LOG_SUMMARY
    processed, failed_rows = 0, []

    try:
        result = parse_people_via_pure_http()
        if result is not None:
            processed, failed_rows = result
    except Exception as e:
        log_err(f"Ошибка в HTTP-блоке: {e}")
        logging.exception("Ошибка в HTTP-блоке (полный traceback ниже, см. sud_script.log)")

    # Блок 5: суды
    fill_courts_column()

    # Блок 6: УГД и БИН УГД
    try:
        fill_ugd_column()
    except Exception as e:
        log_err(f"Ошибка при заполнении УГД: {e}")

    LOG_SUMMARY["АДРЕСА И РЕГИОНЫ"] = {
        "found":     processed - len(failed_rows),
        "total":     processed,
        "not_found": [f"строка {r}" for r in failed_rows],
    }

    print("\n=== ГОТОВО ===")
    print("Создан файл:", OUT_XLSX)
    if args.workdir:
        import shutil
        out_dir = os.path.join(args.workdir, "out")
        os.makedirs(out_dir, exist_ok=True)
        shutil.copy2(OUT_XLSX, os.path.join(out_dir, os.path.basename(OUT_XLSX)))

    return processed - len(failed_rows), len(failed_rows)


if __name__ == "__main__":
    run(None)