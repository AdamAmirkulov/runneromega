# -*- coding: utf-8 -*-
"""
АИС ОИП — Отмены / законченные производства (через API).

Извлечено из AISOIP_Status_plus_Otmeny_API_combined.ipynb (Этап 2).

В отличие от исходного ноутбука, где Этап 2 продолжал работу в той же
браузерной сессии, что и Этап 1 (Статусы), здесь скрипт запускается
самостоятельно и открывает собственную авторизованную сессию АИС ОИП.
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
    parser.add_argument('--date_from', default='')
    parser.add_argument('--date_to', default='')
    parser.add_argument('--login', default='')
    parser.add_argument('--password', default='')
    return parser.parse_args()


args = parse_args()
if not args.company_id or not args.company_id.strip():
    print("❌ ОШИБКА: не передан --company_id — компания не определена, запуск остановлен.")
    sys.exit(1)
os.environ['COMPANY_ID'] = args.company_id.strip()  # ← ОБЯЗАТЕЛЬНО до import config

# =========================
# ДОБАВЛЯЕМ КОРЕНЬ ПРОЕКТА В sys.path
# =========================

project_root = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(project_root)
if project_root not in sys.path:
    sys.path.insert(0, project_root)

# =========================
# УЧЕТНЫЕ ДАННЫЕ ПО КОМПАНИИ — ИЗ config.py
# =========================

from config import CREDENTIALS, _company_id, DB_COMPANY_FILTER

print(f"🏢 Запуск АИС ОИП отмен для компании ID={_company_id}")

# =========================
# ГЛОБАЛЬНЫЕ ПАРАМЕТРЫ
# =========================

# Оставлены пути из Omega_status.py / комбинированного ноутбука.
loading_path_part = r'C:\Users\User\АИС ОИП1\load'
run_prod = True

# Логин/пароль АИС ОИП берутся из config.py (COMPANY_CREDENTIALS) по company_id,
# так же как в status.py и остальных скриптах.
# Если --login/--password явно переданы через форму запуска, они переопределяют
# значения по умолчанию для этой компании.
name_login = f"company{_company_id}"
dict_log_pas = {
    name_login: {
        "log": CREDENTIALS['aisoip_login'],
        "pas": CREDENTIALS['aisoip_password'],
    }
}

if args.login and args.password:
    dict_log_pas[name_login] = {"log": args.login, "pas": args.password}

if args.date_from or args.date_to:
    print(
        "Примечание: параметры --date_from/--date_to приняты, но должники "
        "для проверки отмен берутся из MS SQL без фильтра по датам "
        "(логика сохранена как в исходном ноутбуке)."
    )

AISOIP_URL = "https://aisoip.adilet.gov.kz/cabinet/exec-productions"
API_TIMEOUT_SECONDS = 240


# =========================
# ИМПОРТЫ
# =========================

from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from tqdm import tqdm
from pathlib import Path
from datetime import datetime

import base64
import json
import re
import shutil
import time
import traceback

import pandas as pd


# =========================
# АВТОМАТИЧЕСКОЕ ЛОГИРОВАНИЕ
# =========================

class TeeStream:
    """Дублирует stdout/stderr одновременно в консоль и log-файл."""
    def __init__(self, console, log_handle):
        self.console = console
        self.log_handle = log_handle

    def write(self, data):
        self.console.write(data)
        self.log_handle.write(data)
        self.log_handle.flush()

    def flush(self):
        self.console.flush()
        self.log_handle.flush()

    def isatty(self):
        return getattr(self.console, "isatty", lambda: False)()


_original_stdout = sys.stdout
_original_stderr = sys.stderr
_log_handle = None


def start_run_logging(run_dir):
    global _log_handle
    log_path = Path(run_dir) / "run.log"
    _log_handle = open(log_path, "a", encoding="utf-8", buffering=1)

    sys.stdout = TeeStream(_original_stdout, _log_handle)
    sys.stderr = TeeStream(_original_stderr, _log_handle)

    print("=" * 90)
    print("СТАРТ:", datetime.now().strftime("%d.%m.%Y %H:%M:%S"))
    print("Лог:", log_path)
    print("=" * 90)
    return log_path


def finish_run_logging():
    global _log_handle
    try:
        print("=" * 90)
        print("ЗАВЕРШЕНИЕ:", datetime.now().strftime("%d.%m.%Y %H:%M:%S"))
        print("=" * 90)
    except Exception:
        pass

    sys.stdout = _original_stdout
    sys.stderr = _original_stderr

    if _log_handle is not None:
        try:
            _log_handle.flush()
            _log_handle.close()
        except Exception:
            pass
        _log_handle = None


# =========================
# ПАПКА ЗАПУСКА
# =========================

def make_load_dir(name_login, loading_path_part):
    """Создает папку запуска для сессии Selenium (скачивание вспомогательных файлов)."""
    dt_string = datetime.now().strftime("%Y-%m-%d___%H-%M")
    loading_path = Path(loading_path_part) / f"otmeny_{name_login} {dt_string}"
    loading_path_all = loading_path / "all_file"

    loading_path_all.mkdir(parents=True, exist_ok=False)

    return str(loading_path), str(loading_path_all)


# =========================
# АВТОМАТИЧЕСКИЙ ВХОД В АИС ОИП
# =========================

def to_site(loading_path_all):
    """Открывает Chrome, настраивает сессию и страницу АИС ОИП."""
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")

    prefs = {
        "profile.default_content_settings.popups": 0,
        "download.default_directory": str(Path(loading_path_all).resolve()),
        "directory_upgrade": True,
        "profile.managed_default_content_settings.images": 2,
    }
    options.add_experimental_option("prefs", prefs)

    if run_prod:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
    else:
        driver = webdriver.Chrome(options=options)

    driver.implicitly_wait(5)
    driver.set_script_timeout(API_TIMEOUT_SECONDS)
    driver.get(AISOIP_URL)
    time.sleep(2)
    return driver


def is_login_page(driver):
    url = (driver.current_url or "").lower()
    return (
        "authenticationendpoint" in url
        or "login.do" in url
        or len(driver.find_elements(By.CSS_SELECTOR, 'input[name="usernameUserInput"]')) > 0
    )


def write_login(driver, name_login):
    """Автоматически вводит логин/пароль, если открыта страница авторизации."""
    if not is_login_page(driver):
        print("АИС ОИП: активная авторизованная сессия уже есть.")
        return

    login = dict_log_pas[name_login]["log"]
    password = dict_log_pas[name_login]["pas"]

    login_field = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[name="usernameUserInput"]'))
    )
    login_field.clear()
    login_field.send_keys(login)

    password_field = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, "password"))
    )
    password_field.clear()
    password_field.send_keys(password)

    button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, ".buttons>.ui.primary.button.fluid"))
    )
    button.click()

    WebDriverWait(driver, 60).until(lambda d: not is_login_page(d))
    driver.get(AISOIP_URL)
    WebDriverWait(driver, 60).until(
        lambda d: "aisoip.adilet.gov.kz" in (d.current_url or "").lower()
    )
    time.sleep(2)
    print("АИС ОИП: автоматическая авторизация выполнена.")


# =========================
# API АИС ОИП (общие функции fetch/retry)
# =========================

def browser_fetch_json(driver, url, method="GET", body=None):
    """Авторизованный API fetch внутри Chrome; cookies текущей сессии используются автоматически."""
    script = r"""
    const done = arguments[arguments.length - 1];
    const url = arguments[0];
    const method = arguments[1];
    const body = arguments[2];

    fetch(url, {
        method: method,
        credentials: 'include',
        headers: {
            'Accept': 'application/json, text/plain, */*',
            'Content-Type': 'application/json'
        },
        body: body === null ? undefined : JSON.stringify(body)
    })
    .then(async response => {
        const text = await response.text();
        done({status: response.status, text: text});
    })
    .catch(error => done({status: 0, text: String(error)}));
    """

    result = driver.execute_async_script(script, url, method, body)

    if result.get("status") != 200:
        raise RuntimeError(
            f"API JSON error {result.get('status')} for {url}: "
            f"{result.get('text', '')[:1000]}"
        )

    try:
        return json.loads(result["text"])
    except json.JSONDecodeError as exc:
        raise RuntimeError(
            f"Ответ API не является JSON: {result['text'][:1000]}"
        ) from exc


def reauthenticate_aisoip(driver):
    """Переавторизует сессию АИС ОИП, если она истекла (401 от API)."""
    print("API: сессия АИС ОИП истекла (401) — повторная авторизация...")
    try:
        driver.get(AISOIP_URL)
        write_login(driver, name_login)
        driver.get(AISOIP_URL)
        WebDriverWait(driver, 60).until(
            lambda d: "aisoip.adilet.gov.kz" in (d.current_url or "").lower()
        )
        time.sleep(2)
        print("API: переавторизация выполнена.")
        return True
    except Exception as exc:
        print(f"API: ошибка переавторизации: {exc}")
        traceback.print_exc()
        return False


def api_call_with_retry(callable_obj, description, retries=3, driver=None):
    """Автоматические повторные попытки API. При 401 — переавторизация сессии."""
    last_error = None

    for attempt in range(1, retries + 1):
        try:
            print(f"API: {description}. Попытка {attempt}/{retries}")
            result = callable_obj()
            print(f"API: {description} — успешно")
            return result

        except Exception as exc:
            last_error = exc
            print(
                f"API ERROR: {description}. "
                f"Попытка {attempt}/{retries}: {exc}"
            )
            traceback.print_exc()

            if "401" in str(exc) and driver is not None:
                reauthenticate_aisoip(driver)
            elif attempt < retries:
                time.sleep(2 * attempt)

    raise RuntimeError(
        f"Не удалось выполнить: {description}"
    ) from last_error


# ============================================================
# ЭТАП 2. ОТМЕНЫ / ЗАКОНЧЕННЫЕ ПРОИЗВОДСТВА (через API)
# ============================================================

import pyodbc
from collections import defaultdict
from typing import List, Dict
from pdfminer.high_level import extract_text as pdf_extract_text

OTMENY_SCRIPT_VERSION = "2026-08-13_OTMENY_API_COMBINED"
print(f"Версия скрипта отмен: {OTMENY_SCRIPT_VERSION}")

OTMENY_DB_CONFIG = {
    "server":   "DBSRV",
    "database": "crm",
    "username": "user",
    "password": "Log1cF",
    "trusted":  False,
}

OTMENY_LOADING_ROOT = r"C:\Users\User\АИС ОИП1\load_finish\new_load"
# Название компании для фильтра l.F209 берётся из config.py (COMPANY_DB_FILTER)
# по company_id, переданному при запуске (--company_id), так же как CREDENTIALS.
OTMENY_COMPANY_FILTER = DB_COMPANY_FILTER

OTMENY_API_BASE = "https://aisoip.adilet.gov.kz/extperson/api/rest"
OTMENY_SEARCH_PAGE_SIZE = 100


# ============================================================
# 2.1. ПОЛУЧЕНИЕ ДОЛЖНИКОВ ИЗ MS SQL
# Условия сохранены из otmeny_new_2026_api_v3_auth_fixed(1).ipynb
# ============================================================

def otmeny_get_connection():
    if OTMENY_DB_CONFIG["trusted"]:
        conn_str = (
            f"DRIVER={{ODBC Driver 18 for SQL Server}};"
            f"SERVER={OTMENY_DB_CONFIG['server']};"
            f"DATABASE={OTMENY_DB_CONFIG['database']};"
            f"Trusted_Connection=yes;"
            f"TrustServerCertificate=yes;"
        )
    else:
        conn_str = (
            f"DRIVER={{ODBC Driver 18 for SQL Server}};"
            f"SERVER={OTMENY_DB_CONFIG['server']};"
            f"DATABASE={OTMENY_DB_CONFIG['database']};"
            f"UID={OTMENY_DB_CONFIG['username']};"
            f"PWD={OTMENY_DB_CONFIG['password']};"
            f"TrustServerCertificate=yes;"
        )
    return pyodbc.connect(conn_str)


def otmeny_normalize_iin(value) -> str:
    if pd.isna(value):
        return ""
    value = re.sub(r"\D", "", str(value).replace(".0", ""))
    return value.zfill(12) if value else ""


def otmeny_normalize_ip_number(value) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", "", str(value)).strip().casefold()


def otmeny_normalize_text(value) -> str:
    value = (value or "").lower().replace("ё", "е").replace("\xa0", " ")
    value = re.sub(r"[\u200b\u200c\u200d\ufeff]", "", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def otmeny_get_debtors_from_db():
    query = f"""
    SELECT
        l.EID,
        s.Caption         AS State,
        constF236.Caption AS F236,
        l.F146            AS f258,
        l.F62             AS ЧСИ,
        c.F293            AS ИИН,
        c.FIO
    FROM loans l
    JOIN States s ON s.ID = l.State
    JOIN Constants constF236
        ON constF236.ID = l.F236
        AND constF236.Caption = N'Окончено'
    JOIN clients c ON c.ID = l.CID
    WHERE
        l.F235 IS NULL
        AND l.State NOT IN (7, 8)
        AND l.F209 = N'{OTMENY_COMPANY_FILTER}'
    """

    conn = otmeny_get_connection()
    try:
        df = pd.read_sql(query, conn)
    finally:
        conn.close()

    df["ИИН"] = df["ИИН"].apply(otmeny_normalize_iin)
    df["EID"] = df["EID"].astype(str).str.replace(".0", "", regex=False)
    df["f258"] = df["f258"].fillna("").astype(str).str.strip()

    print(f"Этап 2: получено должников из БД: {len(df)}")
    return df


# ============================================================
# 2.2. ПРОВЕРКА ТЕКУЩЕЙ API-СЕССИИ
# ============================================================

def otmeny_check_api_authorization(driver):
    """Проверяет, что собственная сессия скрипта отмен авторизована."""
    test_iin = "930719401037"
    url = f"{OTMENY_API_BASE}/execproc/search?page=0&size=5"
    payload = {
        "iin": test_iin,
        "searchType": False,
    }

    try:
        browser_fetch_json(
            driver,
            url,
            method="POST",
            body=payload,
        )
        print("Этап 2: текущая API-сессия АИС ОИП активна.")
        return True

    except RuntimeError as exc:
        message = str(exc)
        if "401" in message:
            raise RuntimeError(
                "Этап 2: API АИС ОИП вернул 401 Unauthorized "
                "даже после автоматического входа."
            ) from exc
        raise


# ============================================================
# 2.3. API ПОИСК ИП И ДОКУМЕНТОВ
# ============================================================

def otmeny_browser_download_pdf(driver, url: str, out_path: Path):
    script = r"""
    const callback = arguments[arguments.length - 1];
    const url = arguments[0];

    fetch(url, {
        method: 'GET',
        credentials: 'include',
        headers: {
            'Accept': 'application/pdf, application/octet-stream, */*'
        }
    })
    .then(async response => {
        const blob = await response.blob();
        const reader = new FileReader();

        reader.onloadend = () => callback({
            status: response.status,
            contentType: response.headers.get('content-type') || '',
            dataUrl: reader.result
        });

        reader.readAsDataURL(blob);
    })
    .catch(error => callback({
        status: 0,
        contentType: '',
        dataUrl: String(error)
    }));
    """

    result = driver.execute_async_script(script, url)

    if result["status"] != 200:
        raise RuntimeError(
            f"API PDF GET {url}: HTTP {result['status']}: "
            f"{str(result)[:500]}"
        )

    data_url = result["dataUrl"]
    if not isinstance(data_url, str) or "," not in data_url:
        raise RuntimeError("API вернул некорректный PDF data URL")

    pdf_bytes = base64.b64decode(data_url.split(",", 1)[1])

    if not pdf_bytes.startswith(b"%PDF"):
        raise RuntimeError(
            f"Скачанный файл не является PDF. "
            f"Content-Type={result.get('contentType')}"
        )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_bytes(pdf_bytes)


def otmeny_search_all_execprocs_by_iin(driver, iin: str) -> List[Dict]:
    payload = {
        "iin": iin,
        "searchType": False,
    }

    all_rows = []
    page = 0

    while True:
        url = (
            f"{OTMENY_API_BASE}/execproc/search"
            f"?page={page}&size={OTMENY_SEARCH_PAGE_SIZE}"
        )

        data = api_call_with_retry(
            lambda url=url, payload=payload: browser_fetch_json(
                driver,
                url,
                method="POST",
                body=payload,
            ),
            f"этап 2: поиск ИП ИИН={iin}, page={page}",
            driver=driver,
        )

        if isinstance(data, dict):
            rows = data.get("content", []) or []
            pagination = data.get("pagination", {}) or {}
        elif isinstance(data, list):
            rows = data
            pagination = {}
        else:
            rows = []
            pagination = {}

        all_rows.extend(rows)

        is_last = pagination.get("last")
        total_pages = pagination.get("totalPages")

        if is_last is True:
            break
        if isinstance(total_pages, int) and page + 1 >= total_pages:
            break
        if len(rows) < OTMENY_SEARCH_PAGE_SIZE:
            break

        page += 1

    return all_rows


def otmeny_get_documents_by_execproc_id(driver, exec_proc_id) -> List[Dict]:
    url = (
        f"{OTMENY_API_BASE}/execproc/doc/"
        f"{exec_proc_id}?searchType=false"
    )

    data = api_call_with_retry(
        lambda: browser_fetch_json(
            driver,
            url,
            method="GET",
        ),
        f"этап 2: список документов execProcId={exec_proc_id}",
        driver=driver,
    )

    return data if isinstance(data, list) else []


def otmeny_is_target_document(title: str) -> bool:
    title_norm = otmeny_normalize_text(title)

    target_keywords = (
        "прекращении",
        "возвращении",
        "приостановлении",
        "қаулы атқарушылық іс жүргізуді тоқтату туралы",
        "қаулы атқарушылық құжатты қайтару туралы",
    )

    return any(keyword in title_norm for keyword in target_keywords)


def otmeny_safe_filename(value: str) -> str:
    return re.sub(
        r"[^0-9A-Za-zА-Яа-я_\-]+",
        "_",
        str(value),
    )[:100]


def otmeny_download_target_documents(
    driver,
    docs: List[Dict],
    exec_proc_id,
    download_path: Path,
    eid: str,
):
    target_docs = [
        doc for doc in docs
        if otmeny_is_target_document(doc.get("ddocTitle", ""))
    ]

    downloaded = 0

    for index, doc in enumerate(target_docs, start=1):
        did = str(doc.get("did", "")).strip()
        if not did:
            continue

        url = (
            f"{OTMENY_API_BASE}/execproc/doc/"
            f"{exec_proc_id}/{did}?lang=ru"
        )

        out_path = (
            Path(download_path)
            / f"eid_{eid}__v{index}__{otmeny_safe_filename(did)}.pdf"
        )

        try:
            api_call_with_retry(
                lambda url=url, out_path=out_path:
                    otmeny_browser_download_pdf(
                        driver,
                        url,
                        out_path,
                    ),
                f"этап 2: PDF EID={eid}, did={did}",
                driver=driver,
            )
            downloaded += 1

        except Exception:
            fallback_url = (
                f"{OTMENY_API_BASE}/execproc/doc/"
                f"{did}/{did}?lang=ru"
            )

            api_call_with_retry(
                lambda fallback_url=fallback_url, out_path=out_path:
                    otmeny_browser_download_pdf(
                        driver,
                        fallback_url,
                        out_path,
                    ),
                f"этап 2: PDF fallback EID={eid}, did={did}",
                driver=driver,
            )
            downloaded += 1

    return downloaded, len(docs)


def otmeny_choose_exact_execproc(execprocs: List[Dict], f258: str):
    target = otmeny_normalize_ip_number(f258)
    if not target:
        return None

    exact = [
        item for item in execprocs
        if otmeny_normalize_ip_number(
            item.get("execProcNum", "")
        ) == target
    ]

    if not exact:
        return None

    def date_key(item):
        return pd.to_datetime(
            item.get("startDate"),
            errors="coerce",
        )

    return sorted(
        exact,
        key=lambda item: (
            pd.notna(date_key(item)),
            date_key(item)
            if pd.notna(date_key(item))
            else pd.Timestamp.min,
        ),
        reverse=True,
    )[0]


def otmeny_parse_site_api(df_main, loading_path_docs, driver):
    """Использует переданный driver (собственная сессия этого скрипта)."""
    result_list = []

    otmeny_check_api_authorization(driver)

    for _, row in tqdm(
        df_main.iterrows(),
        total=len(df_main),
        desc="Этап 2: проверка ИП через API",
        mininterval=2,
    ):
        iin_iter = otmeny_normalize_iin(row["ИИН"])
        f258_iter = str(row["f258"]).strip()
        eid_iter = str(row["EID"]).replace(".0", "")
        chi_iter = row["ЧСИ"]

        result = {
            "EID": eid_iter,
            "ЧСИ": chi_iter,
            "ИИН": iin_iter,
            "f258": f258_iter,
            "найдено строк": 0,
            "найдено дело": 0,
            "Кол-во файлов": 0,
            "Дата возбуждения судебного производства": None,
            "Сумма иска": None,
            "Номер исполнительного документа": None,
            "Дата выписки исполнительного документа": None,
            "Орган выдавший исполнительный документ": None,
            "Статус ИП API": None,
        }

        try:
            execprocs = otmeny_search_all_execprocs_by_iin(
                driver,
                iin_iter,
            )
            result["найдено строк"] = len(execprocs)

            selected = otmeny_choose_exact_execproc(
                execprocs,
                f258_iter,
            )

            if selected is None:
                result_list.append(result)
                continue

            result["найдено дело"] = 1
            result[
                "Дата возбуждения судебного производства"
            ] = selected.get("startDate")
            result[
                "Сумма иска"
            ] = selected.get("recoveryAmount")
            result[
                "Номер исполнительного документа"
            ] = selected.get("execDocNum")
            result[
                "Дата выписки исполнительного документа"
            ] = selected.get("ilDate")
            result[
                "Орган выдавший исполнительный документ"
            ] = selected.get("ilOrgan_ru")
            result[
                "Статус ИП API"
            ] = selected.get("status_ru")

            exec_proc_id = selected.get("execProcId")

            if exec_proc_id:
                docs = otmeny_get_documents_by_execproc_id(
                    driver,
                    exec_proc_id,
                )

                downloaded, _ = otmeny_download_target_documents(
                    driver=driver,
                    docs=docs,
                    exec_proc_id=exec_proc_id,
                    download_path=Path(loading_path_docs),
                    eid=eid_iter,
                )

                result["Кол-во файлов"] = downloaded

        except Exception as exc:
            result["Ошибка"] = (
                f"{type(exc).__name__}: {exc}"
            )
            print(
                f"\nОшибка этапа 2: "
                f"EID={eid_iter}, "
                f"ИИН={iin_iter}, "
                f"ИП={f258_iter}: "
                f"{result['Ошибка']}"
            )

        result_list.append(result)

    return pd.DataFrame(result_list)


# ============================================================
# 2.4. ОБРАБОТКА PDF
# Логика классификации сохранена из второго скрипта.
# ============================================================

def otmeny_extract_pdf_text(file_path):
    try:
        return pdf_extract_text(str(file_path))
    except Exception as exc:
        print(f"Ошибка чтения PDF {file_path}: {exc}")
        return ""


def otmeny_find_text_between(text):
    match = re.search(
        r"УСТАНОВИЛ\(А\):(.*?)ПОСТАНОВИЛ\(А\):",
        text,
        re.DOTALL,
    )
    if match:
        return match.group(1).strip()

    match = re.search(
        r"УСТАНОВИЛ:(.*?)ПОСТАНОВИЛ:",
        text,
        re.DOTALL,
    )
    if match:
        return match.group(1).strip()

    # Казахский вариант
    match = re.search(
        r"БЕЛГІЛЕДІ:(.*?)ҚАУЛЫ ЕТТІ:",
        text,
        re.DOTALL,
    )
    if match:
        return match.group(1).strip()

    return ""


def otmeny_find_header_and_date(text):
    header = ""

    header_match = re.search(
        r"ПОСТАНОВЛЕНИЕ\s*\n\s*\n(.*?)\n",
        text,
    )
    if header_match:
        header = header_match.group(1).strip().lower()

    # Казахский заголовок
    if not header:
        header_match = re.search(
            r"ҚАУЛЫ\s*\n\s*\n(.*?)\n",
            text,
        )
        if header_match:
            header = header_match.group(1).strip().lower()

    date_value = ""
    date_match = re.search(
        r"(\d{2}\.\d{2}\.\d{4})",
        text,
    )
    if date_match:
        date_value = date_match.group(1)

    return header, date_value


def otmeny_classify(text):
    text_norm = text.lower()

    # Русский язык — подпункты
    if re.search(r"подпункт\w*\s*5[-\s]*1", text_norm):
        return "судебное банкротство"
    if re.search(r"подпункт\w*\s*5[-\s]*2", text_norm):
        return "внесудебное банкротство"
    if re.search(r"подпункт\w*\s*2[-\s]*1", text_norm):
        return "медиация"
    if re.search(r"подпункт\w*\s*7", text_norm):
        return "погасил"
    if re.search(r"подпункт\w*\s*5", text_norm):
        return "отмена"
    if re.search(r"подпункт\w*\s*3", text_norm):
        return "умерший"
    if re.search(r"подпункт\w*\s*1\b", text_norm):
        return "возврат"

    # Казахский язык — тармақша
    if re.search(r"5[-\s]*1\s*тармақша", text_norm):
        return "судебное банкротство"
    if re.search(r"5[-\s]*2\s*тармақша", text_norm):
        return "внесудебное банкротство"
    if re.search(r"2[-\s]*1\s*тармақша", text_norm):
        return "медиация"
    if re.search(r"7\s*тармақша", text_norm):
        return "погасил"
    if re.search(r"5\s*тармақша", text_norm):
        return "отмена"
    if re.search(r"3\s*тармақша", text_norm):
        return "умерший"
    if re.search(r"1\)\s*тармақша|1\s*тармақша\b", text_norm):
        return "возврат"

    return "не определено"


def otmeny_process_pdfs(docs_path):
    eid_texts = defaultdict(list)

    valid_headers = {
        "о прекращении исполнительного производства",
        "о возвращении исполнительного документа",
        "о приостановлении исполнительного производства",
        "қаулы атқарушылық іс жүргізуді тоқтату туралы",
    }

    files = sorted(Path(docs_path).glob("*.pdf"))

    print()
    print("=== ЭТАП 2: ОБРАБОТКА PDF ===")
    print(f"Всего PDF-файлов: {len(files)}")

    for item in files:
        eid = (
            item.stem
            .split("__", maxsplit=1)[0]
            .replace("eid_", "")
        )

        if not eid.isdigit():
            continue

        text = otmeny_extract_pdf_text(item)
        text_target = otmeny_find_text_between(text)

        if not text_target:
            continue

        header, date_value = otmeny_find_header_and_date(text)

        if header not in valid_headers:
            continue

        entry = [
            text_target,
            otmeny_classify(text_target),
            header,
            date_value,
        ]

        if entry not in eid_texts[eid]:
            eid_texts[eid].append(entry)

    print(
        "Найдено EID с подходящими документами: "
        f"{len(eid_texts)}"
    )

    if not eid_texts:
        return pd.DataFrame(columns=["EID"])

    rows = []

    for eid, entries in eid_texts.items():
        flat = [
            value
            for entry in entries
            for value in entry
        ]
        rows.append([eid, flat])

    df = pd.DataFrame(
        rows,
        columns=["EID", "_list"],
    )

    df_expanded = pd.DataFrame(
        df["_list"].tolist()
    )

    columns = []

    for index in range(
        1,
        len(df_expanded.columns) // 4 + 1,
    ):
        columns += [
            f"Установил {index}",
            f"Класс {index}",
            f"Заголовок {index}",
            f"Дата {index}",
        ]

    df_expanded.columns = columns[
        :len(df_expanded.columns)
    ]

    df = pd.concat(
        [df[["EID"]], df_expanded],
        axis=1,
    )

    class_cols = [
        col for col in df.columns
        if col.startswith("Класс")
    ]

    date_cols = [
        col for col in df.columns
        if col.startswith("Дата")
    ]

    if class_cols:
        mask = df[class_cols].apply(
            lambda row: "погасил" in row.values,
            axis=1,
        )
        df.loc[mask, "Класс 1"] = "погасил"

    if date_cols:
        df["Дата 1"] = df[date_cols].apply(
            lambda row: pd.to_datetime(
                row,
                format="%d.%m.%Y",
                errors="coerce",
            ).max(),
            axis=1,
        )

    return df


# ============================================================
# 2.5. ГЛАВНАЯ ФУНКЦИЯ
# ============================================================

def run_otmeny_stage(driver):
    if driver is None:
        raise RuntimeError(
            "Скрипт отмен не может стартовать: driver отсутствует."
        )

    now = datetime.now()
    dt_string = now.strftime("%Y-%m-%d___%H-%M")

    otmeny_loading_path = (
        Path(OTMENY_LOADING_ROOT)
        / dt_string
    )

    loading_path_docs = (
        otmeny_loading_path
        / "docs"
    )

    otmeny_loading_path.mkdir(
        parents=True,
        exist_ok=True,
    )
    loading_path_docs.mkdir(
        parents=True,
        exist_ok=True,
    )

    # 1. Сделки из MS SQL
    df_main_otmeny = otmeny_get_debtors_from_db()

    # 2. Проверка через API
    df_results_otmeny = otmeny_parse_site_api(
        df_main_otmeny,
        loading_path_docs,
        driver,
    )

    # Технический лог API-проверки
    api_log_path = (
        otmeny_loading_path
        / "api_check_log.xlsx"
    )

    df_results_otmeny.to_excel(
        api_log_path,
        index=False,
    )

    print(
        "Этап 2: технический API-лог сохранен:",
        api_log_path,
    )

    # 3. Разбор PDF
    df_texts_otmeny = otmeny_process_pdfs(
        loading_path_docs
    )

    # 4. Объединение результатов API и PDF
    df_results_otmeny["EID"] = (
        df_results_otmeny["EID"]
        .astype(str)
        .str.replace(".0", "", regex=False)
    )

    if (
        not df_texts_otmeny.empty
        and "EID" in df_texts_otmeny.columns
    ):
        df_texts_otmeny["EID"] = (
            df_texts_otmeny["EID"]
            .astype(str)
            .str.replace(".0", "", regex=False)
        )

        df_final_otmeny = pd.merge(
            df_results_otmeny,
            df_texts_otmeny,
            on="EID",
            how="left",
        )
    else:
        df_final_otmeny = (
            df_results_otmeny.copy()
        )

    # 5. Добавляем данные исходных сделок
    df_main_otmeny["EID"] = (
        df_main_otmeny["EID"]
        .astype(str)
        .str.replace(".0", "", regex=False)
    )

    df_final_otmeny = pd.merge(
        df_final_otmeny,
        df_main_otmeny[
            ["EID", "FIO", "State", "F236"]
        ],
        on="EID",
        how="left",
    )

    # 6. Названия колонок итогового файла
    df_final_otmeny = df_final_otmeny.rename(
        columns={
            "EID": "Уникальный номер сделки",
            "F236": "Статус АИС ОИП",
            "Класс 1": "Статус исполнительного документа",
            "Дата 1": "Дата Постановления о прекращении ИП",
            "f258": "Номер исполнительного производства",
        }
    )

    target_columns_otmeny = [
        "Уникальный номер сделки",
        "Статус АИС ОИП",
        "ЧСИ",
        "Дата возбуждения судебного производства",
        "Сумма иска",
        "Номер исполнительного производства",
        "Номер исполнительного документа",
        "Статус исполнительного документа",
        "Дата выписки исполнительного документа",
        "Орган выдавший исполнительный документ",
        "Дата Постановления о прекращении ИП",
    ]

    for column in target_columns_otmeny:
        if column not in df_final_otmeny.columns:
            df_final_otmeny[column] = None

    df_final_otmeny = df_final_otmeny[
        target_columns_otmeny
    ]

    # Формат дат для импорта
    for date_column in [
        "Дата возбуждения судебного производства",
        "Дата выписки исполнительного документа",
        "Дата Постановления о прекращении ИП",
    ]:
        df_final_otmeny[date_column] = pd.to_datetime(
            df_final_otmeny[date_column],
            errors="coerce",
        ).dt.strftime("%d.%m.%Y")

    # 7. Итоговый Excel
    out_path_otmeny = (
        otmeny_loading_path
        / "AIS_OIP_Import.xlsx"
    )

    df_final_otmeny.to_excel(
        out_path_otmeny,
        engine="xlsxwriter",
        index=False,
    )

    print(
        "Этап 2: итоговый файл сохранен:",
        out_path_otmeny,
    )

    # 8. Копирование в папку автоимпорта CRM — своя у каждой компании
    # (CREDENTIALS['path_crm'], см. config.py:COMPANY_CREDENTIALS; та же папка,
    # что и в родных per-company скриптах otmeny_omega.py/kpi_otmeny.py/
    # investway_otmeny.py — раньше здесь был захардкожен путь Омеги
    # \\192.168.1.251\A-Omega для ВСЕХ компаний, из-за чего итоговый файл
    # KPI/Orion/InvestWay никогда не попадал в их собственную папку CRM).
    network_path_otmeny = Path(
        CREDENTIALS["path_crm"]
    )

    try:
        network_path_otmeny.mkdir(
            parents=True,
            exist_ok=True,
        )

        destination_otmeny = (
            network_path_otmeny
            / "AIS_OIP_Import.xlsx"
        )

        shutil.copy2(
            out_path_otmeny,
            destination_otmeny,
        )

        print(
            "Этап 2: скопировано в:",
            destination_otmeny,
        )

    except Exception as exc:
        print(
            "Этап 2: ошибка копирования "
            f"в сетевую папку: {exc}"
        )

    return out_path_otmeny


# ============================================================
# ЗАПУСК
# ============================================================

driver = None
loading_path, loading_path_all = make_load_dir(name_login, loading_path_part)

start_run_logging(loading_path)

try:
    print("***", name_login, "***")
    print("Папка запуска:", loading_path)

    driver = to_site(loading_path_all)
    write_login(driver, name_login)

    driver.get(AISOIP_URL)
    WebDriverWait(driver, 60).until(
        lambda d: "aisoip.adilet.gov.kz" in (d.current_url or "").lower()
    )

    otmeny_result_path = run_otmeny_stage(driver)

    print()
    print("=" * 90)
    print("СКРИПТ ОТМЕН УСПЕШНО ЗАВЕРШЕН")
    print("Итоговый файл:", otmeny_result_path)
    print("=" * 90)

except Exception:
    print("КРИТИЧЕСКАЯ ОШИБКА В СКРИПТЕ ОТМЕН")
    traceback.print_exc()
    raise

finally:
    if driver is not None:
        try:
            driver.quit()
        except Exception:
            traceback.print_exc()
        driver = None

    finish_run_logging()
