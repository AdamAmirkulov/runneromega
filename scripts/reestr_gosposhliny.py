# -*- coding: utf-8 -*-
"""
Реестр на возврат госпошлины.

1) Выгружает из CRM (MS SQL, база crm) должников выбранной компании,
   у которых АИС ОИП окончено, кредит не погашен и госпошлина ещё
   не оформлена (условия и запрос — от заказчика).
2) Заполняет по этим данным data/Шаблоны/Шаблон_реестр.xlsx.
3) Для каждого ИИН ищет актуальный адрес в Судебном кабинете
   (HTTP-версия блока поиска участника — перенесена из
   scripts/poiskvsk.py: Selenium используется только для логина,
   сам поиск по сотням ИИН идёт через requests).
4) Отправляет Excel-реестр в WhatsApp-группу через Wamm Chat, перед
   файлом коротким сообщением тегает два номера из
   config.REESTR_GP_WA_TAGS (текстовая сводка не отправляется).
5) Дополнительно формирует AddressesImport.xlsx (импорт адресов в
   Дельту) по образцу AddressesImportExample.xlsx и кладёт его в папку
   автоимпорта Дельты — CREDENTIALS['path_crm'] из config.py
   (копия — в out/).
"""

import os
import sys
import argparse


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('--workdir', type=str, default=None)
    parser.add_argument('--company_id', type=str, default=None)
    return parser.parse_known_args()[0]


args = parse_args()

if not args.company_id or not args.company_id.strip():
    print("❌ ОШИБКА: не передан --company_id — компания не определена, запуск остановлен.")
    sys.exit(1)
os.environ['COMPANY_ID'] = args.company_id.strip()

import re
import time
import base64
import shutil
import logging
import unicodedata
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime
from dataclasses import dataclass
from difflib import SequenceMatcher
from urllib.parse import urljoin

import requests
import pyodbc
import pandas as pd
from bs4 import BeautifulSoup
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException

from config import CREDENTIALS, DB_COMPANY_FILTER

try:
    from config import REESTR_GP_WA_TAGS
except ImportError:
    REESTR_GP_WA_TAGS = []

# ============================================================
# НАСТРОЙКИ
# ============================================================

PROJECT_ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = PROJECT_ROOT / "data" / "Шаблоны" / "Шаблон_реестр.xlsx"

WORKDIR = Path(args.workdir) if args.workdir else Path(".")
OUT_DIR = WORKDIR / "out"
OUT_DIR.mkdir(parents=True, exist_ok=True)
OUT_XLSX = OUT_DIR / f"Реестр_на_возврат_госпошлины_{datetime.now():%Y%m%d_%H%M}.xlsx"

DB_CONFIG = {
    "server":   "DBSRV",
    "database": "crm",
    "username": "user",
    "password": "Log1cF",
}

# Судебный кабинет — логин/пароль по company_id, как и в остальных
# скриптах, работающих с office.sud.kz.
USER_AUTH = CREDENTIALS['sk_login']
USER_PASSWORD = CREDENTIALS['sk_password']

WAMM_TOKEN = 'dZ0ec89YqunW7GOB'
WAMM_GROUP_ID = '120363160875851590'
WAMM_MSG_URL = f"https://wamm.chat/api2/msg_to/{WAMM_TOKEN}/"
WAMM_FILE_URL = f"https://wamm.chat/api2/file_from_base64/{WAMM_TOKEN}/"
WAMM_MAX_MSG_CHARS = 3500

logging.basicConfig(
    filename="sud_script.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8",
)


def log4(msg):
    print(time.strftime("[%H:%M:%S]"), msg, flush=True)
    logging.info(msg)


# ============================================================
# SQL-ЗАПРОС
# ============================================================
# Условия и формула госпошлины (3% от остатка без расходов за надпись и
# без ранее оплаченной госпошлины) — от заказчика.
# Изменения относительно исходного запроса:
#  - "l.F209 LIKE '%KPI%'" -> фильтр по компании из config.py
#    (DB_COMPANY_FILTER), чтобы работало для любой компании (--company_id);
#  - статус исполнительного документа: было
#    "(constF235.ID IS NULL OR constF235.Caption NOT IN ('Погасил'))",
#    стало строго "constF235.Caption = N'Отмена'";
#  - добавлено "ISNULL(l.F307,0) < 1" — только где госпошлина ещё не оплачена.

SQL_QUERY = f"""
WITH LoanCounts AS (
    SELECT
        c.F293 AS ИИН,
        COUNT(*) AS LoanCount
    FROM loans l
    JOIN clients c ON l.CID = c.ID
    WHERE l.F209 = N'{DB_COMPANY_FILTER}'
    GROUP BY c.F293
)
SELECT
    l.EID AS [Уникальный номер],
    dF246.F249 AS [Продукт],
    (
        SELECT STRING_AGG(
            CASE
                WHEN LEN(value) > 1 THEN UPPER(LEFT(value, 1)) + LOWER(SUBSTRING(value, 2, LEN(value) - 1))
                ELSE UPPER(value)
            END,
            ' '
        )
        FROM STRING_SPLIT(c.FIO, ' ')
    ) AS [ФИО],
    c.F293 AS [ИИН],
    CAST(l.F34 AS DECIMAL(18,2)) AS [Остаток задолженности],
    constF235.Caption AS [Статус исполнительного документа],
    constF236.Caption AS [Статус АИС ОИП],
    s.Caption AS [Статус кредита],
    ISNULL(lc.LoanCount, 0) AS [Кол-во займов],
    CAST(
        (ISNULL(l.F34,0) - ISNULL(l.F315,0) - ISNULL(l.F307,0))
        AS DECIMAL(18,2)
    ) AS [Остаток без расходов за надпись и без госпошлины],
    CAST((ISNULL(l.F34,0) - ISNULL(l.F13,0)) * 0.03 AS DECIMAL(18,2)) AS [Госпошлина 3%],
    NULL AS [УГД],
    NULL AS [БИН УГД],
    CAST(ISNULL(l.F307, 0) AS DECIMAL(18,2)) AS [Оплаченная ранее госпошлина]
FROM
    loans l (NOLOCK)
JOIN
    clients c (NOLOCK) ON l.CID = c.ID
JOIN
    states s ON s.ID = l.State
JOIN
    Dictionary dF246 ON dF246.ID = l.F246
LEFT JOIN
    Constants constF235 ON constF235.ID = l.F235
JOIN
    Constants constF236 ON constF236.ID = l.F236
LEFT JOIN
    LoanCounts lc ON c.F293 = lc.ИИН
WHERE
    s.Caption IN (N'В работе')
    AND constF236.Caption LIKE '%Окончено%'
    AND constF235.Caption = N'Отмена'
    AND ISNULL(l.F307, 0) < 1
    AND l.F158 IS NULL
    AND (l.F220 = 0 OR l.F220 IS NULL)
    AND l.F209 = N'{DB_COMPANY_FILTER}'
GROUP BY
    l.EID, dF246.F249, c.FIO, c.F293, l.F34, constF235.Caption,
    constF236.Caption, s.Caption, lc.LoanCount, l.F315, l.F307, l.F13;
"""


def load_data_from_db() -> pd.DataFrame:
    conn_str = (
        "DRIVER={ODBC Driver 18 for SQL Server};"
        f"SERVER={DB_CONFIG['server']};"
        f"DATABASE={DB_CONFIG['database']};"
        f"UID={DB_CONFIG['username']};"
        f"PWD={DB_CONFIG['password']};"
        "TrustServerCertificate=yes;"
        "Encrypt=no;"
    )
    log4("🔌 Подключение к БД...")
    conn = pyodbc.connect(conn_str, timeout=30)
    try:
        df = pd.read_sql(SQL_QUERY, conn)
    finally:
        conn.close()
    log4(f"✅ Загружено {len(df)} записей из БД")
    return df


# ============================================================
# ШАБЛОН EXCEL
# ============================================================

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
    return s.zfill(12) if s else ""


def fill_template(df: pd.DataFrame):
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Не найден шаблон реестра: {TEMPLATE_PATH}")

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb["Реестр"] if "Реестр" in wb.sheetnames else wb.active

    # В шаблоне строка 2 — «пример заполнения» (жёлтая заливка, формат «₸» в
    # колонке D), строка 4 — подпись-пояснение. Раньше очищались только
    # значения, а заливка/формат/шрифт оставались → первая строка данных
    # выходила жёлтой и с валютой, остальные — обычными. Чистим и стили.
    _plain_fill = PatternFill(fill_type=None)
    _plain_font = Font(name="Calibri", size=11)

    last_row = max(ws.max_row, len(df) + 1)
    for r in range(2, last_row + 1):
        for c in range(1, 9):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            cell.fill = _plain_fill
            cell.font = _plain_font
            cell.number_format = "General"

    for i, row in enumerate(df.to_dict("records"), start=2):
        ws.cell(row=i, column=1, value=row.get("Уникальный номер"))
        ws.cell(row=i, column=2, value=_norm_iin(row.get("ИИН")))
        ws.cell(row=i, column=3, value=row.get("ФИО"))
        gp = row.get("Госпошлина 3%")
        cell_gp = ws.cell(row=i, column=4, value=float(gp) if gp is not None else None)
        cell_gp.number_format = "#,##0.00"
        # УГД / БИН УГД / адрес / суд заполняются позже (адрес — из СК)

    return wb, ws


# ============================================================
# СУДЕБНЫЙ КАБИНЕТ — HTTP-поиск адреса по ИИН
# (перенесено из scripts/poiskvsk.py: Selenium только для логина,
#  сам поиск по ИИН идёт через requests — как в podacha_iska_v2.py)
# ============================================================

BASE = "https://office.sud.kz"
LOGIN_URL = f"{BASE}/index.xhtml"
SEND_DOCS_URL = f"{BASE}/form/send/index.xhtml"
CREATE_REQUEST_URL = f"{BASE}/form/requestType2/createRequest.xhtml"
VIEWSTATE_NAMES = ("javax.faces.ViewState", "jakarta.faces.ViewState")


def init_driver() -> webdriver.Chrome:
    opts = webdriver.ChromeOptions()
    opts.add_argument("--start-maximized")
    opts.page_load_strategy = "eager"
    opts.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})
    opts.add_argument("--blink-settings=imagesEnabled=false")
    opts.add_argument("--disable-extensions")
    opts.add_argument("--disable-gpu")
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=opts)


def find_with_fallbacks(driver, variants, desc, tries=8, delay=0.4):
    last = None
    for _ in range(tries):
        for by, sel in variants:
            try:
                return driver.find_element(by, sel)
            except Exception as e:
                last = e
        time.sleep(delay)
    raise NoSuchElementException(f"Не удалось найти {desc}. Последняя ошибка: {last}")


def sk_login(driver):
    log4("Открываю страницу логина СК…")
    driver.get(LOGIN_URL)
    time.sleep(1.0)
    try:
        ru = driver.find_elements(By.XPATH, "//a[contains(.,'РУС') and not(contains(@class,'active'))]")
        if ru:
            ru[0].click()
            time.sleep(1.0)
    except Exception:
        pass

    login_el = find_with_fallbacks(
        driver,
        [
            (By.ID, "j_idt78:auth:xin"),
            (By.XPATH, "//input[contains(@id,':auth:xin')]"),
            (By.CSS_SELECTOR, "input[type='email']"),
        ],
        "поле ИИН/БИН",
    )
    pass_el = find_with_fallbacks(
        driver,
        [
            (By.ID, "j_idt78:auth:password"),
            (By.XPATH, "//input[contains(@id,':auth:password')]"),
            (By.XPATH, "//input[@type='password']"),
        ],
        "поле Пароль",
    )
    submit = find_with_fallbacks(
        driver,
        [
            (By.CSS_SELECTOR, "input.button-primary[type='submit']"),
            (By.XPATH, "//input[@type='submit' and contains(@class,'button-primary')]"),
        ],
        "кнопка Войти",
    )

    login_el.clear(); login_el.send_keys(USER_AUTH)
    pass_el.clear(); pass_el.send_keys(USER_PASSWORD)
    submit.click()
    log4("Вошёл в Судебный кабинет")
    time.sleep(2.0)


def _http_session_from_cookies(cookies, user_agent: str) -> requests.Session:
    session = requests.Session()
    session.headers.update({
        "User-Agent": user_agent,
        "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
        "Origin": BASE,
    })
    for cookie in cookies:
        kwargs = {"name": cookie["name"], "value": cookie["value"], "path": cookie.get("path", "/")}
        if cookie.get("domain"):
            kwargs["domain"] = cookie["domain"]
        session.cookies.set(**kwargs)
    return session


def selenium_login_get_session() -> requests.Session:
    drv = init_driver()
    try:
        sk_login(drv)
        cookies = drv.get_cookies()
        user_agent = drv.execute_script("return navigator.userAgent;")
    finally:
        try:
            drv.quit()
        except Exception:
            pass
    return _http_session_from_cookies(cookies, user_agent)


def _extract_viewstate_from_html(html: str):
    m = re.search(r'name="javax\.faces\.ViewState"[^>]*value="([^"]*)"', html)
    if m:
        return m.group(1)
    m = re.search(r'<input[^>]*value="([^"]*)"[^>]*name="javax\.faces\.ViewState"', html)
    return m.group(1) if m else None


@dataclass
class _HttpNavState:
    url: str
    html: str = ""
    viewstate: str = None


class SudHttpClient:
    """Минимальная реплика JSF/RichFaces HTTP-клиента для office.sud.kz."""

    def __init__(self, session: requests.Session, base_url: str = BASE):
        self.session = session
        self.base_url = base_url.rstrip("/")
        self.state = _HttpNavState(url=self.base_url)

    def _check_auth(self, response):
        low_url = response.url.lower()
        low_text = response.text[:5000].lower()
        if "login" in low_url or 'name="login"' in low_text or ("войти" in low_text and "судебный кабинет" in low_text):
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

        response = self.session.post(full_url, data=payload, headers=headers, timeout=60, allow_redirects=True)
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
                return self.get(urljoin(response.url, redirect_url), referer=response.url)
        else:
            self.state.url = response.url
            self.state.html = response.text
            vs = _extract_viewstate_from_html(response.text)
            if vs:
                self.state.viewstate = vs
        return response


def _parse_partial_response(xml_text: str):
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


def _parse_partial_response_with_redirect(xml_text: str):
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
    try:
        debug_dir = Path("sud_http_debug")
        debug_dir.mkdir(parents=True, exist_ok=True)
        debug_path = debug_dir / f"{tag}_{time.strftime('%Y%m%d_%H%M%S')}.html"
        debug_path.write_text(html, encoding="utf-8")
        return str(debug_path)
    except Exception:
        return "<не удалось сохранить>"


def _extract_js_triggered_ajax_id(html: str, func_name: str):
    m = re.search(re.escape(func_name) + r"=function\([^)]*\)\{RichFaces\.ajax\(\"([^\"]+)\"", html)
    return m.group(1) if m else None


def _extract_person_form_context(html: str):
    m_form = re.search(r"onclick=\"fillPersonData\('([^']+):person-iin'\)\"", html)
    ajax_component = _extract_js_triggered_ajax_id(html, "fillPersonData")
    if not m_form or not ajax_component:
        debug_path = _dump_debug_html(html, "person_form")
        raise RuntimeError(f"Не найдена разметка лупы (fillPersonData). HTML: {debug_path}")
    return m_form.group(1), ajax_component


def _merge_partial_updates_into_html(current_html: str, updates: dict) -> str:
    if not updates:
        return current_html
    for update_id, value in updates.items():
        if update_id.endswith("javax.faces.ViewRoot") or update_id == "javax.faces.ViewRoot":
            if value and "<" in value:
                return value
    if not current_html:
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
            candidates = [n for n in (body.contents if body else fragment_soup.contents) if getattr(n, "name", None) is not None]
            replacement = candidates[0] if candidates else None
        if target is not None and replacement is not None:
            target.replace_with(replacement)
        elif replacement is not None:
            (soup.body or soup).append(replacement)
    return str(soup)


def _patch_select_value(html: str, select_id: str, value: str) -> str:
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
    m = re.search(r'id="([^"]+)' + re.escape(known_field_suffix) + r'"', html)
    if not m:
        raise RuntimeError(f"Не найден элемент с суффиксом {known_field_suffix!r} на странице.")
    return m.group(1)


def _serialize_naming_container(html: str, container_id: str) -> dict:
    soup = BeautifulSoup(html, "lxml")
    prefix = container_id + ":"
    result = {}
    for el in soup.find_all(["input", "select", "textarea"]):
        el_id = el.get("id")
        if not el_id or not el_id.startswith(prefix):
            continue
        name = el.get("name") or el_id
        if el.name == "select":
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
                continue
            if el_type in ("checkbox", "radio"):
                if el.has_attr("checked"):
                    result[name] = el.get("value", "on")
            else:
                result[name] = el.get("value", "")
    result[container_id] = container_id
    return result


def _ajax_meta(ajax_component: str, event: str = None) -> dict:
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
            meta["javax.faces.behavior.event"] = event
    if event != "change":
        meta[ajax_component] = ajax_component
    return meta


def _ajax_select_change(client: SudHttpClient, url: str, container_id: str, field: str, value: str, referer: str):
    state = _serialize_naming_container(client.state.html, container_id)
    ajax_component = f"{container_id}:{field}"
    state[ajax_component] = value
    state.update(_ajax_meta(ajax_component, event="change"))
    client.post_form(url, state, referer=referer)
    client.state.html = _patch_select_value(client.state.html, ajax_component, value)


def _click_js_triggered_button(client: SudHttpClient, url: str, container_id: str, func_name: str, referer: str, extra_state: dict = None):
    ajax_component = _extract_js_triggered_ajax_id(client.state.html, func_name)
    if not ajax_component:
        debug_path = _dump_debug_html(client.state.html, f"missing_{func_name}")
        raise RuntimeError(f"Не найдена кнопка ({func_name}). HTML: {debug_path}")
    state = _serialize_naming_container(client.state.html, container_id)
    if extra_state:
        state.update(extra_state)
    state.update(_ajax_meta(ajax_component))
    client.post_form(url, state, referer=referer)


def _extract_plain_ajax_button_id(html: str, near_field_suffix: str):
    field_idx = html.find(near_field_suffix)
    if field_idx < 0:
        return None
    window = html[field_idx: field_idx + 4000]
    soup = BeautifulSoup(window, "lxml")
    for el in soup.find_all("input"):
        onclick = el.get("onclick") or ""
        m = re.search(r'RichFaces\.ajax\("([^"]+)",event,\{"incId":"1"\}\s*\);return false;', onclick)
        if m:
            return m.group(1)
    return None


def open_person_form_via_http(client: SudHttpClient):
    log4("[HTTP-nav] GET send/index.xhtml")
    client.get(SEND_DOCS_URL)
    container1 = _find_container_id(client.state.html, ":case-type")

    _ajax_select_change(client, SEND_DOCS_URL, container1, "case-type", "CIVIL", client.state.url)
    _ajax_select_change(client, SEND_DOCS_URL, container1, "instance", "FIRSTINSTANCE", client.state.url)
    _ajax_select_change(client, SEND_DOCS_URL, container1, "request", "3", client.state.url)
    _click_js_triggered_button(client, SEND_DOCS_URL, container1, "sendRequest", client.state.url)
    log4(f"[HTTP-nav] -> {client.state.url}")

    container2 = _find_container_id(client.state.html, ":edit-categoryGroup")
    _ajax_select_change(client, CREATE_REQUEST_URL, container2, "edit-categoryGroup", "2", client.state.url)
    _ajax_select_change(client, CREATE_REQUEST_URL, container2, "edit-category", "27", client.state.url)
    _ajax_select_change(client, CREATE_REQUEST_URL, container2, "edit-character", "1", client.state.url)

    _click_js_triggered_button(
        client, CREATE_REQUEST_URL, container2, "renderAddPersonModalDialog", client.state.url,
        extra_state={f"{container2}:edit-simpleProcess": "on"},
    )

    next_button_id = _extract_plain_ajax_button_id(client.state.html, ":pp-side")
    if not next_button_id:
        debug_path = _dump_debug_html(client.state.html, "missing_next_side_button")
        raise RuntimeError(f"Не найдена кнопка «Далее» в форме выбора стороны. HTML: {debug_path}")

    container3 = next_button_id.rsplit(":", 1)[0]
    state = _serialize_naming_container(client.state.html, container3)
    state.update(_ajax_meta(next_button_id, event="click"))
    client.post_form(CREATE_REQUEST_URL, state, referer=client.state.url)

    return _extract_person_form_context(client.state.html)


def _fetch_person_by_iin_http(http_session, referer_url, form_id, ajax_component, viewstate, iin, timeout=30):
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
    headers = {"Accept": "*/*", "Faces-Request": "partial/ajax", "Referer": referer_url, "Origin": BASE}
    resp = http_session.post(CREATE_REQUEST_URL, data=data, headers=headers, timeout=timeout)
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
        m_dup = re.search(r"processPersonListCode[^']*'\]\)\.val\('(\d+)'\)", resp.text)
        if m_dup:
            raise RuntimeError(f"HTTP_ALREADY_IN_PROCESS:{m_dup.group(1)}")

    return result, (new_viewstate or viewstate)


def fill_addresses(ws):
    """Ищет адрес по ИИН (колонка B) и пишет его в колонку G ('Актуальный адрес с СК')."""
    MAX_TOTAL_ATTEMPTS_PER_ROW = 6
    RETRY_ROW_PAUSE = 1.0
    MAX_CONSECUTIVE_ERRORS = 3

    def _new_client():
        log4("🔐 Логин в СК через Selenium (единственный раз за сессию)...")
        session = selenium_login_get_session()
        client = SudHttpClient(session)
        log4("🧭 HTTP-навигация до формы участника...")
        form_id, ajax_component = open_person_form_via_http(client)
        log4(f"✅ Форма открыта: form_id={form_id}")
        return client, form_id, ajax_component

    client = form_id = ajax_component = None
    last_error = None
    for attempt in range(1, 4):
        try:
            client, form_id, ajax_component = _new_client()
            break
        except Exception as e:
            last_error = e
            log4(f"❌ Попытка {attempt}/3 открыть форму не удалась: {e}")
            if attempt < 3:
                time.sleep(3)
    if client is None:
        log4(f"❌ Не удалось открыть форму участника после 3 попыток: {last_error} — адреса не будут найдены")
        return 0, ws.max_row - 1

    found, not_found = 0, 0
    consecutive_errors = 0

    for r in range(2, ws.max_row + 1):
        src_iin = _norm_iin(ws.cell(row=r, column=2).value)
        if not src_iin or len(src_iin) != 12:
            continue

        log4(f"Строка {r} → ИИН: {src_iin}")
        success = False
        total_attempts = 0

        while not success and total_attempts < MAX_TOTAL_ATTEMPTS_PER_ROW:
            total_attempts += 1
            if consecutive_errors >= MAX_CONSECUTIVE_ERRORS:
                log4(f"⚠ {consecutive_errors} ошибок подряд — новый логин и новое дело")
                try:
                    client, form_id, ajax_component = _new_client()
                    consecutive_errors = 0
                except Exception as e:
                    log4(f"❌ Не удалось перелогиниться: {e}")

            try:
                data, new_viewstate = _fetch_person_by_iin_http(
                    client.session, client.state.url, form_id, ajax_component,
                    client.state.viewstate, src_iin, timeout=30,
                )
                client.state.viewstate = new_viewstate
            except RuntimeError as e:
                if str(e).startswith("HTTP_ALREADY_IN_PROCESS"):
                    consecutive_errors = MAX_CONSECUTIVE_ERRORS
                else:
                    consecutive_errors += 1
                log4(f"⚠ {e}")
            except requests.RequestException as e:
                consecutive_errors += 1
                log4(f"⚠ ошибка HTTP-запроса: {e}")
            else:
                if any(data.values()):
                    live = data.get("live", "")
                    ws.cell(row=r, column=7, value=live)
                    log4(f"→ адрес: {live}")
                    success = True
                    consecutive_errors = 0
                else:
                    consecutive_errors += 1
                    log4("данных нет")

            if success:
                break
            if total_attempts < MAX_TOTAL_ATTEMPTS_PER_ROW:
                time.sleep(RETRY_ROW_PAUSE)

        if success:
            found += 1
        else:
            not_found += 1
            log4(f"→ строка {r}: адрес не найден после {MAX_TOTAL_ATTEMPTS_PER_ROW} попыток")

        time.sleep(0.25)

    return found, not_found


# ============================================================
# СУД / УГД — та же логика, что в scripts/poiskvsk.py
# (Блок 5: выбор суда по адресу/региону; Блок 6: УГД и БИН УГД)
# Справочники — общие для всех компаний сетевые файлы, как в poiskvsk.py.
# ============================================================

FILE_COURTS = r"\\192.168.1.200\workfolder\Документы для подачи Исков\Шаблоны документов\Суды по гражданским делам.xlsx"
FILE_UGD    = r"\\192.168.1.200\workfolder\Документы для подачи Исков\Шаблоны документов\БИН(УГД) .xlsx"
FILE_ENBEKSHI = r"\\192.168.1.200\workfolder\Документы для подачи Исков\Шаблоны документов\Енбекшиказахский суд.xlsx"
FILE_SHET     = r"\\192.168.1.200\workfolder\Документы для подачи Исков\Шаблоны документов\КРГ Шетский  (1).xlsx"

REGION_NAMES = [
    "Акмолинская область", "Актюбинская область", "Алматинская область",
    "город Алматы", "город Астана", "Атырауская область",
    "Восточно-Казахстанская область", "Жамбылская область",
    "Западно-Казахстанская область", "Карагандинская область",
    "Костанайская область", "Кызылординская область",
    "Мангистауская область", "Область Абай", "Область Жетісу",
    "Область Ұлытау", "Павлодарская область",
    "Северо-Казахстанская область", "Туркестанская область",
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
        paren_groups = re.findall(r"\(([^)]*)\)", text)
        base = re.sub(r"\([^)]*\)", "", text)
        names = [base] + [item for group in paren_groups for item in group.split(",")]
        for name in names:
            name = name.strip(" \xa0,")
            if not name:
                continue
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
    try:
        _, _, settlement_map = _load_two_court_settlement_map(file_path)
    except Exception as e:
        log4(f"⚠ Не удалось прочитать справочник района ({district_label}): {e}")
        return default_court

    for keyword, court_name in settlement_map.items():
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


def _load_courts_by_region():
    log4(f"Читаю файл судов: {FILE_COURTS}")
    df_courts = pd.read_excel(FILE_COURTS, sheet_name="Возврат")
    df_courts["Области"] = df_courts["Области"].ffill()
    df_courts = df_courts.dropna(subset=["Суды"]).copy()
    df_courts["region_key"] = df_courts["Области"].apply(norm_region)
    df_courts["court_key"] = df_courts["Суды"].apply(simplify_court_name)
    return {
        reg_key: grp.reset_index(drop=True)
        for reg_key, grp in df_courts.groupby("region_key")
    }


# ---- УГД / БИН УГД ----

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
        "по город", "по атырау", "по туркестан", "астана жана кала",
        "онтустик", "парк информационных технологии", "морпорт",
        "по г аксу", "по г балхаш", "по г приозерск", "по г шахтинск",
        "по г темиртау", "по г сарань", "по г риддер", "по г семей",
    ]
    return any(_ugd_norm(x) in text for x in city_words)


def _ugd_similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio()


def _load_ugd_mapping() -> list:
    log4(f"Читаю справочник УГД: {FILE_UGD}")
    dict_wb = load_workbook(FILE_UGD, data_only=True)
    dict_ws = dict_wb.active

    mapping = []
    current_region = None

    for row in dict_ws.iter_rows(min_row=1, values_only=True):
        col_a = row[0]
        col_b = row[1] if len(row) > 1 else None
        col_c = row[2] if len(row) > 2 else None

        if col_a and not col_b and not col_c:
            current_region = str(col_a).strip().upper()
            continue

        if col_b and col_c and current_region:
            ugd_name = str(col_c).strip()
            district_key = _extract_ugd_district(ugd_name)
            mapping.append({
                "region": current_region,
                "bin": _make_bin(col_b),
                "ugd": ugd_name,
                "district_key": district_key,
                "ugd_words": _ugd_words_from_text(ugd_name),
                "is_city": _is_city_ugd(ugd_name),
            })

    log4(f"  Загружено УГД-записей: {len(mapping)}")
    return mapping


def _pick_ugd(address_value, region_value, court_value, mapping: list) -> tuple:
    region = _ugd_detect_region(region_value, address_value)

    if region:
        region_candidates = [x for x in mapping if x["region"] == region]
        if len(region_candidates) == 1:
            found = region_candidates[0]
            return found["ugd"], found["bin"]

    address_district = _extract_district_from_address_ugd(str(address_value or ""))
    court_district = _extract_district_from_court_ugd(str(court_value or ""))

    search_keys = []
    if address_district:
        search_keys.append(address_district)
    if court_district and court_district not in search_keys:
        search_keys.append(court_district)

    candidates = mapping
    if region:
        candidates = [x for x in mapping if x["region"] == region]

    found = None

    for key in search_keys:
        for item in candidates:
            if item["is_city"]:
                continue
            if item["district_key"] == key:
                found = item
                break
        if found:
            break

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

    if not found:
        best_item = None
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
                    best_item = item
        if best_score >= 0.72:
            found = best_item

    if not found:
        addr_norm = _ugd_norm(str(address_value or ""))
        city_keywords = {
            "усть каменогорск": "усть каменогорск", "семей": "семей", "риддер": "риддер",
            "павлодар": "павлодар", "экибастуз": "экибастуз", "шахтинск": "шахтинск",
            "аксу": "аксу", "актобе": "актобе", "атырау": "атырау", "актау": "актау",
            "кызылорда": "кызылорда", "талдыкорган": "талдыкорган", "туркестан": "туркестан",
            "тараз": "тараз", "костанай": "костанай", "петропавловск": "петропавловск",
            "кокшетау": "кокшетау", "уральск": "уральск", "шымкент": "шымкент",
            "алматы": "алматы", "астана": "астана",
        }
        for city_kw, city_search in city_keywords.items():
            if re.search(rf'\b{re.escape(city_kw)}\b', addr_norm):
                for item in candidates:
                    if city_search in _ugd_norm(item["ugd"]) and item["is_city"]:
                        found = item
                        break
                break

    if not found:
        region_words = set()
        if region:
            region_words = set(_ugd_words_from_text(region))
        all_words = set(_ugd_words_from_text(str(address_value or "") + " " + str(court_value or "")))
        all_words -= region_words

        best_item = None
        best_matches = 0
        for item in candidates:
            if item["is_city"]:
                continue
            matches = len(all_words.intersection(set(item["ugd_words"])))
            if matches > best_matches:
                best_matches = matches
                best_item = item
        if best_matches >= 1:
            found = best_item

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


def fill_courts_and_ugd(ws):
    """Заполняет E (УГД), F (БИН УГД), H (Судебный орган) по адресу из G."""
    courts_by_region = _load_courts_by_region()
    ugd_mapping = _load_ugd_mapping()

    filled = 0
    for r in range(2, ws.max_row + 1):
        address_val = ws.cell(row=r, column=7).value  # G
        if not address_val:
            continue

        region_val = extract_region_from_address(address_val)
        court_val = pick_court(region_val, address_val, courts_by_region)

        # Города с внутригородскими районами (Актобе, Караганда): если СК отдал
        # адрес без района, суд уже помечен «вручную» — тогда и УГД не угадываем,
        # иначе city-fallback подставит почти всегда неверный УГД.
        if court_val and "вручную" in str(court_val).lower():
            ugd_name, bin_val = "Необходимо вручную проставить УГД!", ""
        else:
            ugd_name, bin_val = _pick_ugd(address_val, region_val, court_val, ugd_mapping)

        ws.cell(row=r, column=8, value=court_val)   # H — Судебный орган
        ws.cell(row=r, column=5, value=ugd_name)     # E — УГД
        ws.cell(row=r, column=6, value=bin_val)      # F — БИН УГД
        filled += 1

    log4(f"Суд/УГД определены для {filled} строк")
    return filled


# ============================================================
# ADDRESSES IMPORT (импорт адресов в Дельту)
# Дополнительная функция: по готовому реестру формирует
# AddressesImport.xlsx по образцу AddressesImportExample.xlsx
# (сетевая папка //192.168.1.251/kpi/Example) и кладёт его в папку
# автоимпорта Дельты — CREDENTIALS['path_crm'] из config.py.
# Существующую логику (реестр, СК, суд/УГД, WhatsApp) не трогает.
# ============================================================

ADDRESSES_IMPORT_HEADERS = [
    "Уникальный номер сделки",
    "Страна",
    "Регион",
    "Город/Район",
    "Населённый пункт",
    "Улица",
    "Дом",
    "Квартира",
    "Тип адреса",
    "Тип владельца",
    "Статус адреса",
]

# Папка автоимпорта Дельты для выбранной компании (config.py).
PATH_CRM = CREDENTIALS.get("path_crm") or ""
ADDRESSES_IMPORT_XLSX = (
    Path(PATH_CRM) / "AddressesImport.xlsx" if PATH_CRM
    else OUT_DIR / "AddressesImport.xlsx"
)

AI_COUNTRY = "КАЗАХСТАН"
AI_ADDR_TYPE = "Из судебного кабинета"
AI_OWNER_TYPE = "Клиент"
AI_ADDR_STATUS = "Актуальный"


def _ai_norm_region(part: str) -> str:
    """Первая часть адреса СК -> Регион как в образце: без слова 'ОБЛАСТЬ',
    дефис -> пробел ('СЕВЕРО-КАЗАХСТАНСКАЯ ОБЛАСТЬ' -> 'СЕВЕРО КАЗАХСТАНСКАЯ',
    'ОБЛАСТЬ АБАЙ' -> 'АБАЙ')."""
    s = str(part or "").upper().replace("-", " ")
    s = re.sub(r"\bОБЛАСТ\w*\b", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def create_addresses_import_file(ws):
    """Формирует AddressesImport.xlsx по строкам реестра, где найден адрес
    в СК (колонка G), и сохраняет его в папку автоимпорта Дельты
    (CREDENTIALS['path_crm']). Колонки E–H (нас. пункт / улица / дом /
    квартира) остаются пустыми — как в образце. Копия кладётся в out/."""
    wb_out = Workbook()
    ws_out = wb_out.active
    ws_out.title = "Лист1"

    for col_idx, header in enumerate(ADDRESSES_IMPORT_HEADERS, start=1):
        ws_out.cell(row=1, column=col_idx, value=header)

    out_row = 2
    for r in range(2, ws.max_row + 1):
        uniq_val = ws.cell(row=r, column=1).value          # A — Уникальный номер
        address_val = ws.cell(row=r, column=7).value       # G — адрес с СК
        if not address_val or not str(address_val).strip():
            continue

        parts = [p.strip() for p in str(address_val).split(",") if p.strip()]
        if not parts:
            continue

        region = _ai_norm_region(parts[0])
        gorod = parts[1].upper().strip() if len(parts) > 1 else None

        ws_out.cell(row=out_row, column=1,  value=uniq_val)         # A
        ws_out.cell(row=out_row, column=2,  value=AI_COUNTRY)       # B
        ws_out.cell(row=out_row, column=3,  value=region or None)   # C
        ws_out.cell(row=out_row, column=4,  value=gorod)            # D
        # E–H остаются пустыми
        ws_out.cell(row=out_row, column=9,  value=AI_ADDR_TYPE)     # I
        ws_out.cell(row=out_row, column=10, value=AI_OWNER_TYPE)    # J
        ws_out.cell(row=out_row, column=11, value=AI_ADDR_STATUS)   # K
        out_row += 1

    os.makedirs(ADDRESSES_IMPORT_XLSX.parent, exist_ok=True)
    wb_out.save(ADDRESSES_IMPORT_XLSX)
    log4(f"✅ AddressesImport сформирован: {ADDRESSES_IMPORT_XLSX} (строк: {out_row - 2})")

    # копия в out/, чтобы файл попал в результат прогона
    if ADDRESSES_IMPORT_XLSX.parent != OUT_DIR:
        try:
            shutil.copy2(ADDRESSES_IMPORT_XLSX, OUT_DIR / "AddressesImport.xlsx")
        except Exception as e:
            log4(f"⚠ Не удалось сохранить копию AddressesImport.xlsx в out/: {e}")

    return out_row - 2


# ============================================================
# WHATSAPP (Wamm Chat)
# ============================================================

def send_wamm_message(text: str):
    if len(text) > WAMM_MAX_MSG_CHARS:
        text = text[:WAMM_MAX_MSG_CHARS] + "\n…сообщение сокращено"
    try:
        resp = requests.get(WAMM_MSG_URL, params={"phone": WAMM_GROUP_ID, "text": text}, timeout=30)
        log4(f"Wamm Chat ответ: {resp.status_code} {resp.text[:300]}")
    except Exception as e:
        log4(f"⚠ Не удалось отправить сообщение в WhatsApp (Wamm Chat): {e}")


def send_wamm_file(path: Path):
    try:
        file_base64 = base64.b64encode(path.read_bytes()).decode("ascii")
        resp = requests.post(
            WAMM_FILE_URL,
            params={"phone": WAMM_GROUP_ID},
            data={"file_name": path.name, "file_base64": file_base64},
            timeout=120,
        )
        log4(f"Wamm Chat (файл) ответ: {resp.status_code} {resp.text[:300]}")
    except Exception as e:
        log4(f"⚠ Не удалось отправить файл в WhatsApp (Wamm Chat): {e}")


def build_wa_tags_text() -> str:
    """Текст для упоминания (@тег) двух номеров из config.REESTR_GP_WA_TAGS.
    К Excel-реестру подпись не прикладывается (Wamm Chat не поддерживает
    caption у файлов) — тег уходит отдельным коротким сообщением перед файлом."""
    nums = [re.sub(r"\D", "", str(n)) for n in (REESTR_GP_WA_TAGS or [])]
    nums = [n for n in nums if n]
    return " ".join(f"@{n}" for n in nums)


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    df = load_data_from_db()

    if df.empty:
        log4("Нет дел, подходящих под условия отбора — реестр не формируется.")
        send_wamm_message(
            f"📋 Реестр на возврат госпошлины\nКомпания: {DB_COMPANY_FILTER}\n"
            "Новых дел по условиям отбора не найдено."
        )
        sys.exit(0)

    wb, ws = fill_template(df)
    wb.save(OUT_XLSX)
    log4(f"💾 Реестр сохранён: {OUT_XLSX}")

    addr_found, addr_not_found = fill_addresses(ws)
    wb.save(OUT_XLSX)
    log4(f"💾 Реестр обновлён (адреса СК): {OUT_XLSX}")

    try:
        fill_courts_and_ugd(ws)
        wb.save(OUT_XLSX)
        log4(f"💾 Реестр обновлён (суд/УГД): {OUT_XLSX}")
    except Exception as e:
        log4(f"⚠ Не удалось определить суд/УГД: {e}")

    try:
        create_addresses_import_file(ws)
    except Exception as e:
        log4(f"⚠ Не удалось сформировать AddressesImport.xlsx: {e}")

    tags_text = build_wa_tags_text()
    if tags_text:
        send_wamm_message(tags_text)
    else:
        log4("⚠ config.REESTR_GP_WA_TAGS не заданы — упоминание номеров не отправлено")
    send_wamm_file(OUT_XLSX)

    log4("=== ГОТОВО ===")
