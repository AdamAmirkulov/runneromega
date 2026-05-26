# -*- coding: utf-8 -*-
"""
Выгрузка из Судебного кабинета (office.sud.kz) через HTTP/AJAX.

Портировано из ноутбука Vigruzka_SK_po_Talony_HTTP_API_v2_FINAL.ipynb.
Бизнес-логика (SQL-отбор сделок, JSF/RichFaces HTTP-запросы, парсинг
карточки дела) сохранена без изменений — добавлена только интеграция
с --company_id/config.py, как и в остальных скриптах проекта.

Schema:
    SQL Server (свои талоны для выбранной компании) -> один вход Selenium
    в Судебный кабинет -> cookies -> requests.Session -> JSF/RichFaces
    AJAX-запросы по каждому талону -> Excel.

Selenium используется только один раз для входа. Все талоны обрабатываются
прямыми HTTP-запросами.
"""

import os
import sys
import argparse

# =========================
# АРГУМЕНТЫ ЗАПУСКА — ДО ИМПОРТА config!
# =========================

def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('--workdir', default=None)
    parser.add_argument('--company_id', type=str, default=None)
    # Приняты для обратной совместимости со старыми расписаниями — SQL-выборка
    # (см. ниже) идёт по статусу сделки, а не по диапазону дат, поэтому эти
    # параметры сейчас не используются.
    parser.add_argument('--date_from', default='')
    parser.add_argument('--date_to', default='')
    return parser.parse_args()


args = parse_args()
if not args.company_id or not args.company_id.strip():
    print("❌ ОШИБКА: не передан --company_id — компания не определена, запуск остановлен.")
    sys.exit(1)
os.environ['COMPANY_ID'] = args.company_id.strip()  # ← ОБЯЗАТЕЛЬНО до import config

if args.date_from or args.date_to:
    print(
        "Примечание: параметры --date_from/--date_to приняты, но не используются — "
        "талоны отбираются из MS SQL по статусу сделки (см. SQL_QUERY)."
    )

# =========================
# ДОБАВЛЯЕМ КОРЕНЬ ПРОЕКТА В sys.path
# =========================

project_root = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(project_root)
if project_root not in sys.path:
    sys.path.insert(0, project_root)

# =========================
# УЧЁТНЫЕ ДАННЫЕ ПО КОМПАНИИ — ИЗ config.py
# =========================

from config import CREDENTIALS, DB_COMPANY_FILTER, _company_id

print(f"🏢 Запуск выгрузки с Судебного кабинета для компании ID={_company_id}")

# =========================
# ОСТАЛЬНЫЕ ИМПОРТЫ
# =========================

import re
import html
import time
import shutil
import traceback

import numpy as np
import pandas as pd
import pyodbc
import requests

from copy import deepcopy
from pathlib import Path
from datetime import datetime
from urllib.parse import urljoin
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


# ============================================================
# НАСТРОЙКИ
# ============================================================

# Судебный кабинет — логин/пароль берутся из config.py (COMPANY_CREDENTIALS)
# по company_id, как и в остальных скриптах, работающих с office.sud.kz
# (podacha_iska_v2.py, chsi_podacha_runner.py).
USER_AUTH = CREDENTIALS['sk_login']
USER_PASSWORD = CREDENTIALS['sk_password']

# БД — общая для всех компаний, отбор конкретной компании идёт через
# фильтр l.F209 = DB_COMPANY_FILTER (см. config.py:COMPANY_DB_FILTER),
# так же как в scripts/otmeny.py.
DB_SERVER = "DBSRV"
DB_DATABASE = "crm"
DB_USERNAME = "user"
DB_PASSWORD = "Log1cF"

LOADING_TIME = 60
PAGELOAD_TIMEOUT = 120

BASE_URL = "https://office.sud.kz"
LOGIN_URL = BASE_URL + "/index.xhtml"
MYCASES_URL = BASE_URL + "/form/cases/mycases.xhtml"

HTTP_TIMEOUT = 60
HTTP_RETRIES = 3
PAUSE_BETWEEN_TICKETS = 0.15

# Итоговый файл — под именем, ожидаемым автоимпортом CRM, кладётся в папку
# компании (CREDENTIALS['path_crm']), как и результаты scripts/otmeny.py.
# Копия для скачивания через веб-интерфейс сохраняется в --workdir/out.
OUTPUT_FILENAME = "processimport.xlsx"

WORKDIR = Path(args.workdir) if args.workdir else Path(".")
LOCAL_OUT_DIR = WORKDIR / "out"
LOCAL_OUT_DIR.mkdir(parents=True, exist_ok=True)
LOCAL_OUTPUT_XLSX = LOCAL_OUT_DIR / OUTPUT_FILENAME

NETWORK_OUTPUT_XLSX = Path(CREDENTIALS['path_crm']) / OUTPUT_FILENAME


# ============================================================
# SQL-ЗАПРОС
# ============================================================
# Условия отбора сохранены из исходного ноутбука. Отличие: вместо
# захардкоженного "l.F209 LIKE N'%Омега%'" используется фильтр по компании
# из config.py, чтобы скрипт мог обрабатывать любую из компаний (Омега,
# KPI, Orion, InvestWay) в зависимости от выбранного --company_id.

SQL_QUERY = f"""
WITH Base AS (
    SELECT
        l.EID AS [Уникальный номер],
        pPick.F356 AS [Номер талона в СК],
        l.F158 AS SortDate,
        ROW_NUMBER() OVER (
            PARTITION BY l.EID, pPick.F356
            ORDER BY l.F158 ASC, l.ID ASC
        ) AS rn
    FROM dbo.loans l WITH (NOLOCK)
    JOIN dbo.states s WITH (NOLOCK)                 ON s.ID = l.State
    JOIN dbo.clients c WITH (NOLOCK)                ON c.ID = l.CID
    JOIN dbo.Dictionary dF246 WITH (NOLOCK)         ON dF246.ID = l.F246
    LEFT JOIN dbo.Constants constF236 WITH (NOLOCK) ON constF236.ID = l.F236
    LEFT JOIN dbo.Constants constF235 WITH (NOLOCK) ON constF235.ID = l.F235
    LEFT JOIN dbo.Constants constF238 WITH (NOLOCK) ON constF238.ID = c.F238
    OUTER APPLY (
        SELECT TOP (1)
            NULLIF(
                REPLACE(
                    REPLACE(LTRIM(RTRIM(p.F356)), NCHAR(160), N''),
                    N' ', N''
                ),
                N''
            ) AS F356
        FROM dbo.ProcessCreatedInLoans pl WITH (NOLOCK)
        JOIN dbo.Process p WITH (NOLOCK) ON p.ID = pl.ProcessID
        WHERE pl.LoanID = l.ID
          AND NULLIF(
                REPLACE(
                    REPLACE(LTRIM(RTRIM(p.F356)), NCHAR(160), N''),
                    N' ', N''
                ),
                N''
              ) IS NOT NULL
        ORDER BY p.ID DESC
    ) pPick
    WHERE
        l.F209 = N'{DB_COMPANY_FILTER}'
        AND s.Caption = N'В работе'
        AND (constF236.Caption = N'Окончено' OR constF236.ID IS NULL)
        AND (constF235.ID IS NULL OR constF235.Caption NOT IN (N'погасил'))
        AND pPick.F356 IS NOT NULL
)
SELECT
    [Уникальный номер],
    [Номер талона в СК]
FROM Base
WHERE rn = 1
ORDER BY SortDate ASC;
"""


def load_tickets_from_db():
    conn_str = (
        "DRIVER={ODBC Driver 18 for SQL Server};"
        f"SERVER={DB_SERVER};"
        f"DATABASE={DB_DATABASE};"
        f"UID={DB_USERNAME};"
        f"PWD={DB_PASSWORD};"
        "TrustServerCertificate=yes;"
        "Encrypt=no;"
    )

    print("Подключение к БД...")
    conn = pyodbc.connect(conn_str, timeout=30)
    try:
        df_source = pd.read_sql(SQL_QUERY, conn)
    finally:
        conn.close()

    def norm(v):
        if v is None:
            return ""
        return re.sub(r"\D+", "", str(v).replace("\xa0", " ").strip())

    df_source["Номер талона в СК"] = df_source["Номер талона в СК"].apply(norm)
    df_source = df_source.drop_duplicates(
        subset=["Уникальный номер", "Номер талона в СК"]
    ).reset_index(drop=True)

    tickets = []
    seen = set()
    for x in df_source["Номер талона в СК"]:
        t = norm(x)
        if t and t not in seen:
            seen.add(t)
            tickets.append(t)

    print(f"Получено строк из БД: {len(df_source)}")
    print(f"Уникальных талонов: {len(tickets)}")
    return df_source, tickets


# ============================================================
# ОДНОКРАТНАЯ АВТОРИЗАЦИЯ
# ============================================================

def build_login_driver():
    opts = Options()
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-notifications")
    opts.add_argument("--disable-popup-blocking")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=opts)
    driver.set_page_load_timeout(PAGELOAD_TIMEOUT)
    return driver


def switch_to_ru(driver):
    for _ in range(10):
        try:
            if "судебный кабинет" in (driver.title or "").lower():
                return
            els = driver.find_elements(By.LINK_TEXT, "РУС")
            if els:
                els[0].click()
                time.sleep(1)
                return
        except Exception:
            pass
        time.sleep(0.5)


MAX_LOGIN_ATTEMPTS = 4
LOGIN_RETRY_BACKOFF_SECONDS = 20


def login_and_make_session():
    """
    Первая попытка входа падала дважды подряд на самом первом
    driver.get(LOGIN_URL) с net::ERR_CONNECTION_TIMED_OUT — сетевая
    заминка на стороне Chrome/office.sud.kz, а не логическая ошибка.
    Поэтому весь цикл входа (не только сам GET) оборачивается ретраями
    с полным перезапуском браузера — так же, как рестарт-логика в
    scripts/sud_fio_zayavlenie.py.
    """
    last_exc = None

    for attempt in range(1, MAX_LOGIN_ATTEMPTS + 1):
        if attempt > 1:
            print(f"Повторная попытка входа в Судебный кабинет ({attempt}/{MAX_LOGIN_ATTEMPTS})...")
        else:
            print("Авторизация в Судебном кабинете...")

        try:
            return _try_login_and_make_session()
        except Exception as e:
            last_exc = e
            print(f"⚠ Попытка входа {attempt}/{MAX_LOGIN_ATTEMPTS} не удалась: {e}")
            if attempt < MAX_LOGIN_ATTEMPTS:
                time.sleep(LOGIN_RETRY_BACKOFF_SECONDS)

    raise RuntimeError(
        f"Не удалось авторизоваться в Судебном кабинете после {MAX_LOGIN_ATTEMPTS} попыток."
    ) from last_exc


def _try_login_and_make_session():
    driver = build_login_driver()

    try:
        driver.get(LOGIN_URL)
        time.sleep(1)
        switch_to_ru(driver)

        login_input = WebDriverWait(driver, LOADING_TIME).until(
            EC.visibility_of_element_located(
                (By.XPATH, "//input[@placeholder='ИИН/БИН' or contains(@id,'auth:xin')]")
            )
        )
        login_input.clear()
        login_input.send_keys(USER_AUTH)

        pass_input = WebDriverWait(driver, LOADING_TIME).until(
            EC.visibility_of_element_located(
                (By.CSS_SELECTOR, "input[placeholder='Пароль']")
            )
        )
        pass_input.clear()
        pass_input.send_keys(USER_PASSWORD)

        WebDriverWait(driver, LOADING_TIME).until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "input.button-primary[type='submit']")
            )
        ).click()

        WebDriverWait(driver, LOADING_TIME).until(
            lambda d: "index.xhtml" not in d.current_url
        )

        switch_to_ru(driver)
        driver.get(MYCASES_URL)

        WebDriverWait(driver, LOADING_TIME).until(
            EC.presence_of_element_located(
                (By.XPATH, "//input[contains(@id,':filter-form:filter-requestNumber')]")
            )
        )

        session = requests.Session()

        ua = driver.execute_script("return navigator.userAgent")
        session.headers.update({
            "User-Agent": ua,
            "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
        })

        for c in driver.get_cookies():
            session.cookies.set(
                c["name"],
                c["value"],
                domain=c.get("domain"),
                path=c.get("path", "/")
            )

        r = session.get(MYCASES_URL, timeout=HTTP_TIMEOUT)
        r.raise_for_status()

        if "filter-requestNumber" not in r.text:
            raise RuntimeError(
                "HTTP-сессия не получила страницу 'Мои дела'. "
                "Возможно, сайт использует дополнительную привязку сессии."
            )

        print("✓ Авторизация выполнена, cookies переданы в HTTP-сессию.")
        return session

    finally:
        driver.quit()


# ============================================================
# JSF/RICHFACES — ДИНАМИЧЕСКОЕ ОПРЕДЕЛЕНИЕ ПОЛЕЙ
# ============================================================

def soup_html(text):
    return BeautifulSoup(text or "", "html.parser")


def get_viewstate_from_html(text):
    soup = soup_html(text)
    el = soup.find("input", {"name": "javax.faces.ViewState"})
    return el.get("value", "") if el else ""


def get_viewstate_from_partial(text):
    m = re.search(
        r'<update[^>]+id=["\'][^"\']*javax\.faces\.ViewState[^"\']*["\'][^>]*>'
        r'<!\[CDATA\[(.*?)\]\]>',
        text or "",
        flags=re.S | re.I
    )
    if m:
        return html.unescape(m.group(1).strip())

    m = re.search(
        r'<update[^>]+javax\.faces\.ViewState[^>]*>(.*?)</update>',
        text or "",
        flags=re.S | re.I
    )
    return html.unescape(m.group(1).strip()) if m else ""


def normalize_js_identifier(value):
    """
    Преобразует JavaScript-экранирование в реальный JSF component ID.
    Например:
        filter\\u002Dform  ->  filter-form
    """
    if not value:
        return ""
    value = html.unescape(str(value))
    value = re.sub(
        r"\\u([0-9a-fA-F]{4})",
        lambda m: chr(int(m.group(1), 16)),
        value
    )
    value = value.replace("\\/", "/")
    return value


def get_filter_metadata(page_html):
    soup = soup_html(page_html)

    ticket_input = soup.find(
        "input",
        attrs={"name": re.compile(r":filter-form:filter-requestNumber$")}
    )
    if not ticket_input:
        ticket_input = soup.find(
            "input",
            attrs={"id": re.compile(r":filter-form:filter-requestNumber$")}
        )

    if not ticket_input:
        raise RuntimeError("Не найдено поле номера талона на странице.")

    ticket_name = ticket_input.get("name") or ticket_input.get("id")
    form_name = ticket_name.rsplit(":filter-requestNumber", 1)[0]

    search_btn = None
    form = soup.find(attrs={"id": form_name}) or soup.find("form", attrs={"name": form_name})

    if form:
        candidates = form.find_all(["input", "button"])
    else:
        candidates = soup.find_all(["input", "button"])

    for el in candidates:
        value = (el.get("value") or el.get_text(" ", strip=True) or "").strip().lower()
        if value == "найти":
            search_btn = el
            break

    if not search_btn:
        for el in candidates:
            eid = el.get("id") or el.get("name") or ""
            onclick = el.get("onclick") or ""
            if "filter-form" in eid and ("RichFaces.ajax" in onclick or "javax.faces" in onclick):
                search_btn = el
                break

    search_source = ""
    if search_btn:
        search_source = search_btn.get("name") or search_btn.get("id") or ""

    if not search_source:
        m = re.search(
            r'RichFaces\.ajax\(["\']([^"\']*filter-form:[^"\']+)["\']',
            page_html,
            flags=re.I
        )
        if m:
            search_source = m.group(1)

    if not search_source:
        raise RuntimeError("Не удалось определить JSF-компонент кнопки 'Найти'.")

    select_source = ""
    patterns = [
        r'function\s+viewSelectedRequest\s*\([^)]*\)\s*\{.*?RichFaces\.ajax\(["\']([^"\']+)["\']',
        r'viewSelectedRequest.*?RichFaces\.ajax\(["\']([^"\']+)["\']',
    ]
    for pat in patterns:
        m = re.search(pat, page_html, flags=re.S | re.I)
        if m:
            select_source = m.group(1)
            break

    if not select_source:
        m = re.search(
            r'viewSelectedRequest.*?(?:source|request)\s*[:=(]\s*["\']([^"\']*filter-form:[^"\']+)["\']',
            page_html,
            flags=re.S | re.I
        )
        if m:
            select_source = m.group(1)

    return {
        "form_name": normalize_js_identifier(form_name),
        "ticket_name": normalize_js_identifier(ticket_name),
        "search_source": normalize_js_identifier(search_source),
        "select_source": normalize_js_identifier(select_source),
        "viewstate": get_viewstate_from_html(page_html),
        "page_html": page_html,
    }


def base_form_values(page_html, form_name):
    soup = soup_html(page_html)
    form = soup.find(attrs={"id": form_name}) or soup.find("form", attrs={"name": form_name})
    result = {}

    if not form:
        result[form_name] = form_name
        return result

    for el in form.find_all(["input", "select", "textarea"]):
        name = el.get("name")
        if not name:
            continue

        typ = (el.get("type") or "").lower()

        if typ in {"submit", "button", "image", "file"}:
            continue
        if typ in {"checkbox", "radio"} and not el.has_attr("checked"):
            continue

        if el.name == "select":
            opt = el.find("option", selected=True)
            if not opt:
                opt = el.find("option")
            value = opt.get("value", "") if opt else ""
        elif el.name == "textarea":
            value = el.get_text()
        else:
            value = el.get("value", "")

        result[name] = value

    result[form_name] = form_name
    return result


def make_ajax_headers(referer):
    return {
        "Faces-Request": "partial/ajax",
        "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
        "Referer": referer,
        "Accept": "*/*",
        "X-Requested-With": "XMLHttpRequest",
    }


def extract_updated_list(partial_xml):
    m = re.search(
        r'<update[^>]+id=["\'][^"\']*:filter-form:list["\'][^>]*>'
        r'<!\[CDATA\[(.*?)\]\]>',
        partial_xml or "",
        flags=re.S | re.I
    )
    if m:
        return m.group(1)
    return ""


def extract_case_number_from_search(partial_xml):
    frag = extract_updated_list(partial_xml)
    if not frag:
        return ""

    soup = soup_html(frag)
    card = soup.select_one(".case-item-container")
    if not card:
        return ""

    h3 = card.find("h3")
    if not h3:
        return ""

    return h3.get_text(" ", strip=True).replace("№", "").strip()


def detect_select_source(meta, partial_xml):
    if meta.get("select_source"):
        return normalize_js_identifier(meta["select_source"])

    for text in (partial_xml, meta.get("page_html", "")):
        patterns = [
            r'viewSelectedRequest.*?RichFaces\.ajax\(["\']([^"\']+)["\']',
            r'org\.richfaces\.ajax\.component["\']?\s*[:=]\s*["\']([^"\']+)["\']',
        ]
        for pat in patterns:
            m = re.search(pat, text or "", flags=re.S | re.I)
            if m:
                val = m.group(1)
                if ":filter-form:" in val and val != meta.get("search_source"):
                    return normalize_js_identifier(val)

    return ""


def extract_redirect_url(response):
    if response.url and "/viewRequest.xhtml" in response.url:
        return response.url

    text = response.text or ""

    m = re.search(r'<redirect\s+url=["\']([^"\']+)["\']', text, flags=re.I)
    if m:
        return urljoin(BASE_URL, html.unescape(m.group(1)))

    m = re.search(
        r'(/form/requestType\d*/viewRequest\.xhtml\?i=[A-Za-z0-9]+)',
        text,
        flags=re.I
    )
    if m:
        return urljoin(BASE_URL, m.group(1))

    return ""


def http_post_retry(session, url, data, headers):
    last = None
    for attempt in range(1, HTTP_RETRIES + 1):
        try:
            r = session.post(
                url,
                data=data,
                headers=headers,
                timeout=HTTP_TIMEOUT,
                allow_redirects=True
            )
            r.raise_for_status()
            return r
        except Exception as e:
            last = e
            if attempt < HTTP_RETRIES:
                time.sleep(1.0 * attempt)
    raise last


def search_and_open_case(session, ticket):
    page = session.get(MYCASES_URL, timeout=HTTP_TIMEOUT)
    page.raise_for_status()

    if "filter-requestNumber" not in page.text:
        raise RuntimeError("Сессия Судебного кабинета истекла.")

    meta = get_filter_metadata(page.text)
    values = base_form_values(page.text, meta["form_name"])

    for key in list(values):
        if key.endswith(":filter-requestDate"):
            values[key] = ""
        elif key.endswith(":filter-requestType"):
            values[key] = ""
        elif key.endswith(":district-field"):
            values[key] = ""
        elif key.endswith(":court-field"):
            values[key] = ""

    values[meta["ticket_name"]] = ticket
    values["javax.faces.ViewState"] = meta["viewstate"]

    src = meta["search_source"]
    values["javax.faces.source"] = src
    values["javax.faces.partial.event"] = "click"
    values["javax.faces.partial.execute"] = f"{src} @component"
    values["javax.faces.partial.render"] = "@component"
    values["org.richfaces.ajax.component"] = src
    values[src] = src
    values["rfExt"] = "null"
    values["AJAX:EVENTS_COUNT"] = "1"
    values["javax.faces.partial.ajax"] = "true"

    r_search = http_post_retry(
        session,
        MYCASES_URL,
        values,
        make_ajax_headers(MYCASES_URL)
    )

    case_number = extract_case_number_from_search(r_search.text)

    if not case_number:
        return None, "", "NOT_FOUND"

    new_vs = get_viewstate_from_partial(r_search.text) or meta["viewstate"]
    select_source = detect_select_source(meta, r_search.text)

    if not select_source:
        raise RuntimeError(
            "Найден талон, но не определён JSF-компонент открытия карточки. "
            "Нужно повторно запустить discovery после обновления сайта."
        )

    open_values = base_form_values(page.text, meta["form_name"])

    for key in list(open_values):
        if key.endswith(":filter-requestDate"):
            open_values[key] = ""
        elif key.endswith(":filter-requestType"):
            open_values[key] = ""
        elif key.endswith(":district-field"):
            open_values[key] = ""
        elif key.endswith(":court-field"):
            open_values[key] = ""

    open_values[meta["ticket_name"]] = ticket
    open_values["javax.faces.ViewState"] = new_vs
    open_values["javax.faces.source"] = select_source
    open_values["javax.faces.partial.execute"] = f"{select_source} @component"
    open_values["javax.faces.partial.render"] = "@component"
    open_values["param1"] = "0"
    open_values["org.richfaces.ajax.component"] = select_source
    open_values[select_source] = select_source
    open_values["rfExt"] = "null"
    open_values["AJAX:EVENTS_COUNT"] = "1"
    open_values["javax.faces.partial.ajax"] = "true"

    r_open = session.post(
        MYCASES_URL,
        data=open_values,
        headers=make_ajax_headers(MYCASES_URL),
        timeout=HTTP_TIMEOUT,
        allow_redirects=False
    )
    r_open.raise_for_status()

    detail_url = extract_redirect_url(r_open)

    if not detail_url:
        for rr in [r_open] + list(r_open.history):
            u = extract_redirect_url(rr)
            if u:
                detail_url = u
                break

    if not detail_url:
        raise RuntimeError(
            f"Талон {ticket}: поиск выполнен, но URL карточки не получен."
        )

    detail = session.get(
        detail_url,
        headers={"Referer": MYCASES_URL},
        timeout=HTTP_TIMEOUT
    )
    detail.raise_for_status()

    if "viewRequest.xhtml" not in detail.url:
        raise RuntimeError(
            f"Талон {ticket}: вместо карточки получен URL {detail.url}"
        )

    return detail.text, case_number, "OK"


# ============================================================
# ПАРСИНГ HTML КАРТОЧКИ
# ============================================================

LIST_COL = [
    "list_defendants",
    "plaintiff_representative",
    "number_case",
    "location",
    "judicial_authority",
    "case_category",
    "amount_of_claim",
    "amount_of_state_duties",
    "date_of_departure",
    "date_of_rejection",
    "cause_of_rejection",
    "date_of_registration",
    "name_judge",
    "date_of_return",
    "date_of_agreement",
    "date_of_simplification",
    "date_of_first_instance",
    "date_court_order",
    "date_of_leaving_without_consideration",
    "date_refusal_of_summary_proceedings",
]


DICT_SEARCH_NAME_COL = {
    "Исковое заявление отправлено": "date_of_departure",
    "Заявление успешно отправлено": "date_of_departure",
    "Иск отправлено": "date_of_departure",
    "отклонено": "date_of_rejection",
    "зарегистрировано": "date_of_registration",
    "вынесено определение о возврате искового заявления": "date_of_return",
    "вынесено определение о возвращении заявления": "date_of_return",
    "вынесено определение об утверждении соглашения об урегулировании спора": "date_of_agreement",
    "вынесено определение о рассмотрении дела в порядке упрощенного производства": "date_of_simplification",
    "вынесено определение о возбуждении гражданского дела": "date_of_simplification",
    "вынесено решение первой инстанции": "date_of_first_instance",
    "вынесено судебный приказ": "date_court_order",
    "оставлении заявления без рассмотрения": "date_of_leaving_without_consideration",
    "об отмене решения в порядке упрощенного (письменного) производства": "date_refusal_of_summary_proceedings",
}


def create_dict_for_table():
    return dict(zip(LIST_COL, ["" for _ in LIST_COL]))


def create_dict_defendant():
    return {"defendant_inn": "", "defendant_fio": ""}


def clean_text(x):
    return re.sub(r"\s+", " ", (x or "").replace("\xa0", " ")).strip()


def select_col_by_text(text):
    low = (text or "").lower()
    for kw, col in DICT_SEARCH_NAME_COL.items():
        if kw.lower() in low:
            return col
    return None


def text_cause_of_rejection(text):
    idx = (text or "").rfind("Причина: ")
    if idx >= 0:
        return clean_text((text or "")[idx + 9:])
    return ""


def extract_judge(text):
    m = re.search(
        r"Судья\s*[–-]\s*([А-ЯA-ZЁӘІҢҒҮҰҚӨҺ][^,\n\.]+)",
        text or ""
    )
    return clean_text(m.group(1)) if m else ""


def extract_court_case_number(text):
    nums = re.findall(
        r"№\s*([0-9]{4}-[0-9]{2}-[0-9]-[0-9]+/[0-9]+)",
        text or ""
    )
    return nums[-1].strip() if nums else ""


def parse_dt(text):
    m = re.search(r"(\d{2}\.\d{2}\.\d{4})\s+(\d{2}:\d{2})", text or "")
    if not m:
        return None
    try:
        return datetime.strptime(
            m.group(1) + " " + m.group(2),
            "%d.%m.%Y %H:%M"
        )
    except Exception:
        return None


def first_input_value_by_classes(soup, classes):
    selector = "." + ".".join(classes.split())
    el = soup.select_one(selector)
    if not el:
        return ""
    inp = el.find("input")
    return clean_text(inp.get("value", "")) if inp else ""


def parse_dynamics(soup, out):
    block = soup.find(id="collapseDynamicReview")
    mode = "review"

    if not block:
        block = soup.find(id="collapseHistory")
        mode = "history"

    if not block:
        return

    best_dt = {k: None for k in LIST_COL}
    best_judge_dt = None
    best_case_dt = None
    best_judge = ""
    best_case = ""

    for rec in block.select(".well"):
        if mode == "review":
            ps = rec.find_all("p")
            date_text = clean_text(ps[1].get_text(" ", strip=True)) if len(ps) > 1 else ""
            divs = rec.find_all("div")
            text = clean_text(divs[1].get_text(" ", strip=True)) if len(divs) > 1 else clean_text(rec.get_text(" ", strip=True))
        else:
            span = rec.find("span")
            date_text = clean_text(span.get_text(" ", strip=True)) if span else ""
            text = clean_text(rec.get_text(" ", strip=True))

        dt = parse_dt(date_text)
        date_short = date_text[:10] if re.match(r"\d{2}\.\d{2}\.\d{4}", date_text) else ""

        cno = extract_court_case_number(text)
        if cno and (best_case_dt is None or (dt is not None and dt >= best_case_dt)):
            best_case = cno
            best_case_dt = dt

        judge = extract_judge(text)
        if judge and (best_judge_dt is None or (dt is not None and dt >= best_judge_dt)):
            best_judge = judge
            best_judge_dt = dt

        name_col = select_col_by_text(text)
        if not name_col:
            continue

        cur = best_dt.get(name_col)
        if cur is None or (dt is not None and dt >= cur):
            out[name_col] = date_short
            best_dt[name_col] = dt

            if name_col == "date_of_rejection":
                out["cause_of_rejection"] = text_cause_of_rejection(text)

    if best_judge:
        out["name_judge"] = best_judge
    if best_case and not out.get("number_case"):
        out["number_case"] = best_case


def parse_panels(soup, out):
    for panel in soup.select("fieldset.panel.panel-default.form-wrapper"):
        legend = panel.select_one("legend .panel-title")
        if not legend:
            continue

        title = clean_text(legend.get_text(" ", strip=True)).lower()

        if title == "информация об оплате":
            rows = panel.select("div.panel-body table tbody > tr")
            if rows:
                for row in rows:
                    cells = row.find_all("td")
                    if len(cells) >= 4:
                        out["amount_of_claim"] = clean_text(cells[2].get_text(" ", strip=True))
                        out["amount_of_state_duties"] = clean_text(cells[3].get_text(" ", strip=True))
            else:
                claim = panel.find(id="edit-claim")
                duty = panel.find(id="edit-duty")
                if claim:
                    out["amount_of_claim"] = clean_text(claim.get("value", ""))
                if duty:
                    out["amount_of_state_duties"] = clean_text(duty.get("value", ""))

        elif title == "стороны процесса":
            defendants = []
            table = panel.select_one("div.panel-body div.table")
            if not table:
                table = panel.find("table")

            if table:
                rows = table.find_all("tr")
                if not rows:
                    rows = table.select(".row")

                for row in rows[1:]:
                    cells = row.find_all(["td", "div"], recursive=False)
                    if len(cells) < 4:
                        cells = row.find_all(["td"])

                    vals = [clean_text(c.get_text(" ", strip=True)) for c in cells]

                    if len(vals) >= 4:
                        side = vals[0].lower()
                        if side in {"ответчик", "должник"}:
                            defendants.append({
                                "defendant_inn": vals[2],
                                "defendant_fio": vals[3],
                            })
                        elif side == "представитель":
                            out["plaintiff_representative"] = vals[3]

            out["list_defendants"] = defendants


def parse_case_html(detail_html, ticket, search_case_number):
    soup = soup_html(detail_html)
    out = create_dict_for_table()
    out["ticket_number"] = ticket
    out["number_case"] = search_case_number

    parse_dynamics(soup, out)

    out["location"] = first_input_value_by_classes(
        soup,
        "form-type-textfield form-item-district form-disabled form-item form-group"
    )
    out["judicial_authority"] = first_input_value_by_classes(
        soup,
        "form-type-textfield form-item-court form-disabled form-item form-group"
    )
    out["case_category"] = first_input_value_by_classes(
        soup,
        "form-type-textfield form-item-categoryGroup form-disabled form-item form-group"
    )

    parse_panels(soup, out)

    if not out.get("list_defendants"):
        out["list_defendants"] = [create_dict_defendant()]

    return out


# ============================================================
# ФОРМИРОВАНИЕ ИТОГОВОГО ФАЙЛА
# ============================================================

def format_defendants(lst):
    if not lst or not isinstance(lst, list):
        return ""
    parts = []
    for item in lst:
        inn = (item.get("defendant_inn") or "").strip()
        fio = (item.get("defendant_fio") or "").strip()
        if inn and fio:
            parts.append(f"{inn}: {fio}")
        elif fio:
            parts.append(fio)
        elif inn:
            parts.append(inn)
    return "; ".join(parts)


def get_latest_date(val):
    if pd.isna(val) or not isinstance(val, str) or not val.strip():
        return ""
    parts = [p.strip() for p in val.split(",")]
    if len(parts) == 1:
        return parts[0]

    parsed = []
    for p in parts:
        for fmt in ["%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"]:
            try:
                parsed.append((datetime.strptime(p, fmt), p))
                break
            except ValueError:
                pass

    return max(parsed, key=lambda x: x[0])[1] if parsed else parts[-1]


FINAL_COLS = [
    "Уникальный номер",
    "Номер талона в СК",
    "Дата подачи заявления на выписку ИЛ",
    "Дата получения ИЛ",
    "Дата передачи ИЛ ЧСИ",
    "Комментарии по суду",
    "Дата подачи заявления на выдачу СП",
    "Дата получения СП",
    "Дата передачи СП ЧСИ",
    "Комментарий",
    "Представитель истца",
    "Номер судебного дела",
    "Судебный орган",
    "Категория дела",
    "Сумма иска",
    "Сумма государственной пошлины",
    "Дата отправки искового заявления",
    "Отклонено",
    "Причина отклонения заявления",
    "Зарегистрировано",
    "Судья",
    "Вынесено определение о возврате искового заявления",
    "Вынесено определение об оставлении заявления без рассмотрения",
    "Вынесено определение об утверждении соглашения об урегулировании спора",
    "Вынесено определение о рассмотрении дела в порядке упрощенного производства",
    "Вынесено решение первой инстанции",
    "Вынесен судебный приказ",
    "Определение об отмене решения в порядке упрощенного производства",
    "Ответственный",
    "Дата создания",
    "Название процесса",
    "Тип процесса",
    "Статус процесса",
]


def build_final_dataframe(df_source, parsed_by_ticket):
    list_dicts_cases = []

    for _, src_row in df_source.iterrows():
        unique_number = src_row["Уникальный номер"]
        ticket = re.sub(
            r"\D+",
            "",
            str(src_row["Номер талона в СК"]).replace("\xa0", " ").strip()
        )

        base = parsed_by_ticket.get(ticket)

        if base is None:
            d = create_dict_for_table()
            d["ticket_number"] = ticket
            d["number_case"] = ""
            d["list_defendants"] = [create_dict_defendant()]
        else:
            d = deepcopy(base)
            d["ticket_number"] = ticket

        d["Уникальный номер"] = unique_number
        list_dicts_cases.append(d)

    print(f"\n✓ Обработано итоговых строк: {len(list_dicts_cases)}")

    df = pd.DataFrame(list_dicts_cases)

    df["formatted_defendants"] = (
        df["list_defendants"].apply(format_defendants)
        if "list_defendants" in df.columns else ""
    )

    if "list_defendants" in df.columns:
        df = df.drop(columns=["list_defendants"])

    for col in [
        "date_of_departure",
        "date_of_rejection",
        "date_of_registration",
        "date_of_return",
        "date_of_agreement",
        "date_of_simplification",
        "date_of_first_instance",
        "date_court_order",
        "date_of_leaving_without_consideration",
        "date_refusal_of_summary_proceedings",
    ]:
        if col in df.columns:
            df[col] = df[col].apply(get_latest_date)

    df = df.rename(columns={
        "ticket_number": "Номер талона в СК",
        "number_case": "Номер судебного дела",
        "date_court_order": "Вынесен судебный приказ",
        "date_refusal_of_summary_proceedings": "Определение об отмене решения в порядке упрощенного производства",
        "amount_of_claim": "Сумма иска",
        "amount_of_state_duties": "Сумма государственной пошлины",
        "date_of_departure": "Дата отправки искового заявления",
        "date_of_rejection": "Отклонено",
        "cause_of_rejection": "Причина отклонения заявления",
        "date_of_registration": "Зарегистрировано",
        "name_judge": "Судья",
        "date_of_return": "Вынесено определение о возврате искового заявления",
        "date_of_agreement": "Вынесено определение об утверждении соглашения об урегулировании спора",
        "date_of_simplification": "Вынесено определение о рассмотрении дела в порядке упрощенного производства",
        "date_of_first_instance": "Вынесено решение первой инстанции",
        "plaintiff_representative": "Представитель истца",
        "judicial_authority": "Судебный орган",
        "date_of_leaving_without_consideration": "Вынесено определение об оставлении заявления без рассмотрения",
        "case_category": "Категория дела",
    })

    for col in [
        "Дата подачи заявления на выписку ИЛ",
        "Дата получения ИЛ",
        "Дата передачи ИЛ ЧСИ",
        "Комментарии по суду",
        "Дата подачи заявления на выдачу СП",
        "Дата получения СП",
        "Дата передачи СП ЧСИ",
        "Комментарий",
        "Ответственный",
        "Дата создания",
        "Название процесса",
    ]:
        if col not in df.columns:
            df[col] = np.nan

    df["Тип процесса"] = "1. Упрощённое производство"
    df["Статус процесса"] = "Рассмотрение дела"

    existing = [c for c in FINAL_COLS if c in df.columns]
    rest = [c for c in df.columns if c not in FINAL_COLS]
    df = df[existing + rest]

    cols_to_drop = [
        "location",
        "formatted_defendants",
        "Область (столица, город республиканского значения)",
        "Ответчики",
    ]
    df = df.drop(columns=[c for c in cols_to_drop if c in df.columns])

    return df


# ============================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ============================================================

def main():
    df_source, tickets_uniq = load_tickets_from_db()

    if not tickets_uniq:
        print("Нет талонов для обработки — завершаю без запуска браузера.")
        df_empty = pd.DataFrame(columns=FINAL_COLS)
        df_empty.to_excel(LOCAL_OUTPUT_XLSX, index=False, engine="openpyxl")
        print(f"✓ Пустой итоговый файл сохранён: {LOCAL_OUTPUT_XLSX}")
        return

    session = login_and_make_session()

    parsed_by_ticket = {}

    print("\nНачинаем HTTP/AJAX выгрузку...\n")

    for i, ticket in enumerate(tickets_uniq, start=1):
        print(f"[{i}/{len(tickets_uniq)}] {ticket}", end="", flush=True)

        try:
            detail_html, case_number, status = search_and_open_case(session, ticket)

            if status == "NOT_FOUND":
                print(" -> не найден")
                parsed_by_ticket[ticket] = None
            else:
                parsed_by_ticket[ticket] = parse_case_html(
                    detail_html,
                    ticket,
                    case_number
                )
                print(f" -> OK, дело {case_number}")

        except Exception as e:
            parsed_by_ticket[ticket] = None
            print(f" -> ОШИБКА: {e}")

        if PAUSE_BETWEEN_TICKETS:
            time.sleep(PAUSE_BETWEEN_TICKETS)

    df = build_final_dataframe(df_source, parsed_by_ticket)

    df.to_excel(LOCAL_OUTPUT_XLSX, index=False, engine="openpyxl")
    print(f"✓ Итоговый Excel сохранён: {LOCAL_OUTPUT_XLSX}")
    print(f"Строк: {len(df)}, колонок: {len(df.columns)}")

    try:
        NETWORK_OUTPUT_XLSX.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(LOCAL_OUTPUT_XLSX, NETWORK_OUTPUT_XLSX)
        print(f"✓ Скопировано в папку автоимпорта CRM: {NETWORK_OUTPUT_XLSX}")
    except Exception as e:
        print(f"✗ Не удалось скопировать в папку автоимпорта {NETWORK_OUTPUT_XLSX}: {e}")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        print("КРИТИЧЕСКАЯ ОШИБКА В СКРИПТЕ ВЫГРУЗКИ С СК")
        traceback.print_exc()
        raise
