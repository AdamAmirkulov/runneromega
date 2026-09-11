# -*- coding: utf-8 -*-
"""
АИС ОИП — Ограничение на выезд (API-версия, по образцу status.py/otmeny.py).

Заменяет старую версию этого скрипта (проверка через клики по UI АИС ОИП,
захардкоженную под одну компанию "kpi"/Омега). Логика проверки перенесена
из отдельного прототипа new_mery_omega.py:
1) список должников без постановления берётся из MS SQL (CRM);
2) по каждому ИИН через API АИС ОИП ищутся исполнительные производства;
3) по каждому производству скачивается тот же xlsx-экспорт документов,
   что и по клику на значок Excel в карточке производства (через fetch
   внутри авторизованной сессии Chrome, без реальных кликов по UI);
4) из названий документов достаются даты извещения / постановления /
   снятия ограничения на выезд.

В отличие от прототипа, здесь (как и в status.py/otmeny.py):
- вход в АИС ОИП автоматический, отдельным Selenium-драйвером (не нужен
  заранее запущенный Chrome с remote-debugging и ручной логин);
- логин/пароль АИС ОИП и фильтр БД по компании берутся из config.py
  (CREDENTIALS / DB_COMPANY_FILTER) по --company_id — скрипт работает для
  любой из компаний, а не только для Омеги;
- параметры запуска приходят через argparse (--workdir/--company_id/...);
- итоговый файл кладётся в <workdir>/out (виден в интерфейсе как результат
  задачи), а урезанный SanctionsImport.xlsx дополнительно копируется в
  папку автоимпорта CRM компании (CREDENTIALS['path_crm']) — аналогично
  тому, как otmeny.py копирует AIS_OIP_Import.xlsx.
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
    parser.add_argument('--iin', default='', help='Проверить только один ИИН (без него — полный прогон по БД)')
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
# УЧЁТНЫЕ ДАННЫЕ ПО КОМПАНИИ — ИЗ config.py
# =========================

from config import CREDENTIALS, _company_id, DB_COMPANY_FILTER

print(f"🏢 Запуск проверки ограничения на выезд для компании ID={_company_id}")

# Логин/пароль АИС ОИП берутся из config.py по company_id, как и в остальных
# скриптах. Если --login/--password переданы явно через форму — они
# переопределяют значения по умолчанию для этой компании.
name_login = f"company{_company_id}"
dict_log_pas = {
    name_login: {
        "log": CREDENTIALS['aisoip_login'],
        "pas": CREDENTIALS['aisoip_password'],
    }
}

if args.login and args.password:
    dict_log_pas[name_login] = {"log": args.login, "pas": args.password}

AISOIP_URL = "https://aisoip.adilet.gov.kz/cabinet/exec-productions"
API_BASE = "https://aisoip.adilet.gov.kz/extperson/api/rest"
API_TIMEOUT_SECONDS = 240

# Обрабатывать ВСЕ исполнительные производства данного ИИН (объединяя
# найденные даты по всем), а не только одно.
PROCESS_ALL_EXECPROCS = True

# Как часто сохранять итоговый Excel по ходу полного прогона
SAVE_EVERY = 25


# =========================
# ИМПОРТЫ
# =========================

from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from pathlib import Path
from datetime import datetime
from typing import List, Dict, Optional, Tuple

import base64
import json
import re
import shutil
import time
import traceback

import pandas as pd
import pyodbc


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
    Path(run_dir).mkdir(parents=True, exist_ok=True)
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

if args.workdir:
    JOB_DIR = Path(args.workdir)
else:
    dt_string = datetime.now().strftime("%Y-%m-%d___%H-%M")
    JOB_DIR = Path(r"C:\Users\User\Desktop") / f"ogranichenie_{_company_id}_{dt_string}"

OUT_DIR = JOB_DIR / "out"
DOWNLOADS_DIR = JOB_DIR / "downloads"
OUT_DIR.mkdir(parents=True, exist_ok=True)
DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)

RESULT_EXCEL_PATH = OUT_DIR / "Список по ограничению на выезд.xlsx"

run_prod = True


# =========================
# БД — ПОЛУЧЕНИЕ СПИСКА ИЗ SQL SERVER
# =========================

DB_CONFIG = {
    "server":   "DBSRV",
    "database": "crm",
    "username": "user",
    "password": "Log1cF",
}

# Фильтр по компании (l.F209) берётся из config.py:COMPANY_DB_FILTER по
# --company_id — так же, как в otmeny.py. Раньше здесь был захардкожен
# `l.F209 LIKE N'%Омега%'`, из-за чего скрипт всегда работал только с
# Омегой, независимо от выбранной в интерфейсе компании.
SQL_QUERY = f"""
SELECT
    l.EID AS [Уникальный номер],
    CAST(LEFT(dF246.F249, 50) AS VARCHAR(50)) AS [Продукт],
    CAST(l.F287 AS VARCHAR(50)) AS [Номер договора],
    c.FIO  AS [ФИО],
    c.F293 AS [ИИН],
    s.Caption AS [Статус кредита],
    l.F34 AS [Остаток задолженности],
    l.F62 AS [ЧСИ],
    l.F146 AS [Номер исполнительного производства],
    NULLIF(constF236.Caption, '') AS [Статус АИС ОИП],
    l.F352 AS [Извещение о временном ограничении на выезд],
    l.F351 AS [Постановление об ограничении на выезд],
    l.F353 AS [Снятие ограничения на выезд]
FROM loans l
LEFT JOIN clients   c         ON c.ID = l.CID
LEFT JOIN states    s         ON s.ID = l.State
LEFT JOIN Dictionary dF246    ON dF246.ID = l.F246
LEFT JOIN Constants constF236 ON constF236.ID = l.F236
WHERE
    s.Caption NOT IN (N'Погашен', N'Обратный выкуп', N'В графике', N'Армия', N'Умерший', N'Банкрот', N'ВосПС', N'Мошенничество')
    AND l.F209 = N'{DB_COMPANY_FILTER}'
    AND constF236.Caption = N'На исполнении'
    AND l.F34 > 173000
    AND l.F147 < DATEADD(DAY, -30, GETDATE())
ORDER BY
    l.EID ASC;
"""


def get_db_connection():
    conn_str = (
        f"DRIVER={{ODBC Driver 18 for SQL Server}};"
        f"SERVER={DB_CONFIG['server']};"
        f"DATABASE={DB_CONFIG['database']};"
        f"UID={DB_CONFIG['username']};"
        f"PWD={DB_CONFIG['password']};"
        f"TrustServerCertificate=yes;"
    )
    return pyodbc.connect(conn_str, timeout=30)


def load_data_from_db() -> pd.DataFrame:
    print("🔌 Подключение к БД...")
    try:
        conn = get_db_connection()
        df = pd.read_sql(SQL_QUERY, conn)
        conn.close()
        print(f"✅ Загружено {len(df)} записей из БД")
        return df
    except Exception as e:
        print(f"❌ Ошибка подключения к БД: {e}")
        raise


# =========================
# АВТОМАТИЧЕСКИЙ ВХОД В АИС ОИП
# =========================

def to_site():
    """Открывает Chrome и переходит на страницу АИС ОИП."""
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")

    prefs = {
        "profile.default_content_settings.popups": 0,
        "download.default_directory": str(DOWNLOADS_DIR.resolve()),
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


def reauthenticate_aisoip(driver):
    """Переавторизует сессию АИС ОИП, если она истекла."""
    print("API: сессия АИС ОИП истекла — повторная авторизация...")
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


# =========================
# API АИС ОИП
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


def api_call_with_retry(callable_obj, description, retries=3, driver=None):
    """Автоматические повторные попытки API. При 401 — переавторизация сессии."""
    last_error = None

    for attempt in range(1, retries + 1):
        try:
            return callable_obj()
        except Exception as exc:
            last_error = exc
            print(f"API ERROR: {description}. Попытка {attempt}/{retries}: {exc}")

            if "401" in str(exc) and driver is not None:
                reauthenticate_aisoip(driver)
            elif attempt < retries:
                time.sleep(2 * attempt)

    raise RuntimeError(f"Не удалось выполнить: {description}") from last_error


def search_execprocs_by_iin(driver, iin: str) -> List[Dict]:
    url = f"{API_BASE}/execproc/search?page=0&size=20"
    payload = {"iin": iin, "searchType": False}

    data = api_call_with_retry(
        lambda: browser_fetch_json(driver, url, method="POST", body=payload),
        f"поиск ИП по ИИН={iin}",
        driver=driver,
    )

    if isinstance(data, dict):
        return data.get("content", []) or []
    if isinstance(data, list):
        return data
    return []


def download_execproc_excel(driver, exec_proc_id, out_path: Path):
    """Качает тот же xlsx, что скачивается по клику на зелёный значок
    Excel рядом с "Документы" в карточке производства."""
    url = f"{API_BASE}/execproc/excel/{exec_proc_id}?docType=doc&lang=ru"

    script = r"""
    const callback = arguments[arguments.length - 1];
    const url = arguments[0];

    fetch(url, {
        method: 'GET',
        credentials: 'include',
        headers: {'Accept': 'application/json, text/plain, */*'}
    })
    .then(async r => {
        const blob = await r.blob();
        const reader = new FileReader();
        reader.onloadend = () => callback({
            status: r.status,
            contentType: r.headers.get('content-type'),
            dataUrl: reader.result
        });
        reader.readAsDataURL(blob);
    })
    .catch(e => callback({status: 0, contentType: '', dataUrl: String(e)}));
    """

    res = driver.execute_async_script(script, url)

    if res["status"] != 200:
        raise RuntimeError(f"download_execproc_excel failed {res['status']}: {str(res)[:500]}")

    data_url = res["dataUrl"]
    if "," not in data_url:
        raise RuntimeError(f"Некорректный dataUrl: {data_url[:200]}")

    file_bytes = base64.b64decode(data_url.split(",", 1)[1])

    # xlsx-файлы — это zip-архивы, начинаются с сигнатуры PK. Если сессия
    # успела протухнуть, вместо файла придёт HTML страницы логина — сюда
    # оно не пройдёт, и вызывающий код переавторизует сессию и повторит.
    if not file_bytes.startswith(b"PK"):
        raise RuntimeError("Скачанный файл не похож на xlsx (нет сигнатуры PK)")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_bytes(file_bytes)


def download_execproc_excel_with_retry(driver, exec_proc_id, out_path: Path, retries=3):
    last_error = None
    for attempt in range(1, retries + 1):
        try:
            download_execproc_excel(driver, exec_proc_id, out_path)
            return
        except Exception as exc:
            last_error = exc
            print(f"API ERROR: скачивание Excel execProcId={exec_proc_id}. Попытка {attempt}/{retries}: {exc}")
            msg = str(exc)
            if "401" in msg or "не похож на xlsx" in msg:
                reauthenticate_aisoip(driver)
            elif attempt < retries:
                time.sleep(2 * attempt)

    raise RuntimeError(f"Не удалось скачать Excel execProcId={exec_proc_id}") from last_error


# =========================
# ПАРСИНГ СКАЧАННОГО EXCEL
# =========================

def normalize_text(s: str) -> str:
    s = (s or "").lower().replace("ё", "е")
    s = s.replace("\xa0", " ")
    s = re.sub(r"[\u200b\u200c\u200d\ufeff]", "", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def extract_dates_from_excel_file(excel_file_path: Path) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """Читает скачанный xlsx (тот же формат, что скачивается по клику на
    Excel-значок в карточке производства) и достаёт три типа дат."""
    try:
        df = pd.read_excel(excel_file_path)
        df.columns = df.columns.str.strip()

        if 'Наименование документа' not in df.columns or 'Дата подписания' not in df.columns:
            print(f"  ❌ Не найдены нужные колонки в {excel_file_path.name}: {list(df.columns)}")
            return None, None, None

        postanovlenie_dates, izveshenie_dates, removal_dates = [], [], []

        for _, row in df.iterrows():
            title = normalize_text(str(row['Наименование документа']))
            date_str = str(row['Дата подписания'])

            date_obj = None
            for fmt in ('%d.%m.%Y %H:%M:%S', '%d.%m.%Y', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d'):
                try:
                    date_obj = datetime.strptime(date_str, fmt)
                    break
                except Exception:
                    continue

            if not date_obj:
                continue

            date_formatted = date_obj.strftime('%d.%m.%Y')

            # 1) Снятие ограничения
            if (
                ('снят' in title or 'сняти' in title or 'отмен' in title)
                and 'огранич' in title
                and ('выезд' in title or 'vyezd' in title)
            ):
                removal_dates.append(date_formatted)
                continue

            # 2) Извещение
            if 'извещ' in title and (
                ('огранич' in title or 'ogranichenii' in title)
                and ('выезд' in title or 'vyezd' in title)
            ):
                izveshenie_dates.append(date_formatted)
                continue

            # 3) Постановление об ограничении
            if (
                ('постановлен' in title or 'postanovlenie' in title or 'kauly boryshkerdin' in title)
                and ('огранич' in title or 'ogranichenii' in title or 'shekteu' in title)
                and ('выезд' in title or 'vyezd' in title or 'shyguyn' in title)
            ):
                postanovlenie_dates.append(date_formatted)

        date_postanovlenie = '; '.join(sorted(set(postanovlenie_dates))) if postanovlenie_dates else None
        date_izveshenie = '; '.join(sorted(set(izveshenie_dates))) if izveshenie_dates else None
        date_removal = '; '.join(sorted(set(removal_dates))) if removal_dates else None

        return date_postanovlenie, date_izveshenie, date_removal

    except Exception as e:
        print(f"  ❌ Ошибка чтения {excel_file_path.name}: {e}")
        return None, None, None


def process_one_iin(driver, iin: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """Возвращает (дата_постановления, дата_извещения, дата_снятия) для одного ИИН.
    Скачивает Excel-экспорт документов по каждому найденному исполнительному
    производству и объединяет найденные даты."""
    execprocs = search_execprocs_by_iin(driver, iin)

    if not execprocs:
        print("  ⚠️ Производства не найдены")
        return None, None, None

    targets = execprocs if PROCESS_ALL_EXECPROCS else execprocs[:1]

    all_postanovlenie, all_izveshenie, all_removal = [], [], []

    for ep in targets:
        exec_proc_id = ep.get("execProcId")
        if not exec_proc_id:
            continue

        tmp_path = DOWNLOADS_DIR / f"iin_{iin}__ep_{exec_proc_id}.xlsx"

        try:
            download_execproc_excel_with_retry(driver, exec_proc_id, tmp_path)
            p, i, r = extract_dates_from_excel_file(tmp_path)

            if p:
                all_postanovlenie.extend(p.split('; '))
            if i:
                all_izveshenie.extend(i.split('; '))
            if r:
                all_removal.extend(r.split('; '))

        except Exception as e:
            print(f"  ⚠️ Не удалось скачать/разобрать экспорт для execProcId={exec_proc_id}: {e}")

        finally:
            try:
                tmp_path.unlink(missing_ok=True)
            except Exception:
                pass

    date_postanovlenie = '; '.join(sorted(set(all_postanovlenie))) if all_postanovlenie else None
    date_izveshenie = '; '.join(sorted(set(all_izveshenie))) if all_izveshenie else None
    date_removal = '; '.join(sorted(set(all_removal))) if all_removal else None

    return date_postanovlenie, date_izveshenie, date_removal


# =========================
# EXCEL — ИТОГОВЫЙ ОТЧЁТ И SanctionsImport
# =========================

def update_main_df_with_dates(df_main: pd.DataFrame, iin, date_postanovlenie, date_izveshenie, date_removal) -> bool:
    iin_formatted = str(iin).strip().zfill(12)
    mask = df_main['ИИН'] == iin_formatted

    if not mask.any():
        print(f"  ❌ ИИН {iin_formatted} не найден в таблице")
        return False

    if date_postanovlenie:
        df_main.loc[mask, 'Постановление об ограничении на выезд'] = date_postanovlenie
    if date_izveshenie:
        df_main.loc[mask, 'Извещение о временном ограничении на выезд'] = date_izveshenie
    if date_removal:
        df_main.loc[mask, 'Снятие ограничения на выезд'] = date_removal

    return True


def save_result_excel(output_path: Path, df_main: pd.DataFrame):
    output_columns = [
        'Уникальный номер', 'Продукт', 'Номер договора', 'ФИО', 'ИИН',
        'Статус кредита', 'Остаток задолженности', 'ЧСИ',
        'Номер исполнительного производства', 'Статус АИС ОИП',
        'Извещение о временном ограничении на выезд',
        'Постановление об ограничении на выезд',
        'Снятие ограничения на выезд',
    ]
    existing_cols = [c for c in output_columns if c in df_main.columns]
    df_out = df_main[existing_cols].copy()

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False, sheet_name="Результат")
        ws = writer.sheets["Результат"]

        if 'ИИН' in existing_cols:
            iin_col_idx = existing_cols.index('ИИН') + 1  # openpyxl — 1-based
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=iin_col_idx)
                if cell.value:
                    cell.value = str(cell.value).strip().zfill(12)
                    cell.number_format = '@'

        for col_cells in ws.columns:
            max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 50)


DATE_COLUMNS_FOR_IMPORT = (
    'Извещение о временном ограничении на выезд',
    'Постановление об ограничении на выезд',
    'Снятие ограничения на выезд',
)


def _latest_single_date(value) -> Optional[str]:
    """Из ячейки с одной или несколькими датами (dd.mm.yyyy, через '; ' —
    так их пишет update_main_df_with_dates, когда по ИИН несколько
    исполнительных производств) возвращает только САМУЮ СВЕЖУЮ.

    CRM-автоимпорт кладёт значение в обычное поле даты и не умеет
    разбирать несколько дат в одной ячейке — раньше это лечили вручную
    отдельным скриптом (Omega Scripts/послдата.py) уже после экспорта,
    из-за чего часть строк с несколькими производствами (и, соответственно,
    несколькими датами) в CRM не подтягивалась, хотя файл успешно
    записывался и копировался — ровно то, на что жаловались изначально."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None

    text = str(value).strip()
    if not text:
        return None

    parts = [p.strip() for p in text.split(';') if p.strip()]
    dates = []
    for part in parts:
        try:
            dates.append(datetime.strptime(part, '%d.%m.%Y'))
        except ValueError:
            continue

    if not dates:
        return text  # не разобрали — оставляем как есть, чтобы не терять данные

    return max(dates).strftime('%d.%m.%Y')


def export_sanctions_import(df_main: pd.DataFrame):
    """Урезанный файл для автоимпорта в CRM (только колонки для импорта) —
    сохраняется в out/ задачи и копируется в папку автоимпорта компании
    (CREDENTIALS['path_crm']), аналогично AIS_OIP_Import.xlsx в otmeny.py."""
    cols_needed = [
        'Уникальный номер',
        'Номер исполнительного производства',
        'Извещение о временном ограничении на выезд',
        'Постановление об ограничении на выезд',
        'Снятие ограничения на выезд',
    ]
    cols_existing = [c for c in cols_needed if c in df_main.columns]
    missing = [c for c in cols_needed if c not in df_main.columns]
    if missing:
        print(f"⚠️ Не найдены столбцы для SanctionsImport: {missing}")

    df_export = df_main[cols_existing].copy()

    for col in DATE_COLUMNS_FOR_IMPORT:
        if col in df_export.columns:
            df_export[col] = df_export[col].apply(_latest_single_date)

    # "Уникальный номер" (EID) должен быть текстом, а не числом — так он
    # записан в эталонном SanctionsImport.xlsx (py scripts/SanctionsImport.xlsx),
    # который реально подхватывался CRM. pd.read_sql возвращает его как
    # int64, и то_excel() без явного приведения пишет числовую ячейку —
    # CRM сопоставляет сделку по этому полю как по тексту, и числовая
    # ячейка с ней не матчится, из-за чего строка молча не импортируется.
    if 'Уникальный номер' in df_export.columns:
        df_export['Уникальный номер'] = (
            df_export['Уникальный номер']
            .astype(str)
            .str.strip()
            .str.replace(r'\.0$', '', regex=True)
        )

    local_path = OUT_DIR / "SanctionsImport.xlsx"
    # Имя листа — как в эталонном файле ("Лист1"), а не дефолтное "Sheet1".
    df_export.to_excel(local_path, index=False, sheet_name="Лист1")
    print(f"✅ SanctionsImport сохранён: {local_path}")

    try:
        network_dir = Path(CREDENTIALS["path_crm"])
        network_dir.mkdir(parents=True, exist_ok=True)
        network_path = network_dir / "SanctionsImport.xlsx"
        shutil.copy2(local_path, network_path)
        print(f"✅ SanctionsImport скопирован в папку автоимпорта CRM: {network_path}")
    except Exception as e:
        print(f"⚠️ Не удалось скопировать SanctionsImport в сетевую папку CRM: {e}")


# =========================
# РЕЖИМ 1: ПРОВЕРКА ОДНОГО ИИН (--iin)
# =========================

def run_debug_dump(driver, iin: str):
    iin = str(iin).strip().zfill(12)
    print(f"\n=== ПРОВЕРКА ОДНОГО ИИН: {iin} ===")
    execprocs = search_execprocs_by_iin(driver, iin)
    print(f"Найдено производств: {len(execprocs)}")
    print(json.dumps(execprocs, ensure_ascii=False, indent=2)[:4000])

    date_postanovlenie, date_izveshenie, date_removal = process_one_iin(driver, iin)
    print(f"\nПостановление: {date_postanovlenie}")
    print(f"Извещение: {date_izveshenie}")
    print(f"Снятие: {date_removal}")


# =========================
# РЕЖИМ 2: ПОЛНЫЙ ПРОГОН ПО БД
# =========================

def run_full_check(driver):
    df_main = load_data_from_db()

    df_main['ИИН'] = df_main['ИИН'].apply(
        lambda x: str(x).strip().zfill(12) if pd.notna(x) and str(x).strip() else None
    )

    df_to_process = df_main[
        df_main['Постановление об ограничении на выезд'].isna()
        | (df_main['Постановление об ограничении на выезд'].astype(str).str.strip() == '')
    ]

    iin_list = df_to_process['ИИН'].dropna().astype(str).str.strip().str.zfill(12).tolist()

    print("📊 СТАТИСТИКА:")
    print(f"   ВСЕГО ИЗ БД: {len(df_main)}")
    print(f"   БЕЗ ПОСТАНОВЛЕНИЯ: {len(df_to_process)}")
    print(f"   К ОБРАБОТКЕ: {len(iin_list)}")
    print("=" * 60)

    processed = 0
    successful = 0
    errors = []
    start_time = time.time()

    for idx, iin in enumerate(iin_list, start=1):
        iteration_start = time.time()
        iin = str(iin).strip().zfill(12)
        print(f"\n[{idx}/{len(iin_list)}] 🔍 ИИН: {iin}")

        try:
            date_postanovlenie, date_izveshenie, date_removal = process_one_iin(driver, iin)

            if date_postanovlenie or date_izveshenie or date_removal:
                if update_main_df_with_dates(df_main, iin, date_postanovlenie, date_izveshenie, date_removal):
                    successful += 1
                    print(f"  ✅ Постановление: {date_postanovlenie} | Извещение: {date_izveshenie} | Снятие: {date_removal}")
            else:
                print("  ⚠️ Нет дат")

            processed += 1

        except Exception as e:
            print(f"  ❌ Ошибка: {e}")
            traceback.print_exc()
            errors.append(iin)

        iteration_time = time.time() - iteration_start
        print(f"  ⏱️ {iteration_time:.1f}с")

        if processed > 0 and processed % SAVE_EVERY == 0:
            save_result_excel(RESULT_EXCEL_PATH, df_main)
            print(f"💾 Промежуточное сохранение после {processed} строк")

    save_result_excel(RESULT_EXCEL_PATH, df_main)
    export_sanctions_import(df_main)

    total_time = time.time() - start_time
    avg_time = total_time / processed if processed else 0

    print(f"\n{'=' * 60}")
    print("📊 ИТОГИ")
    print(f"   Обработано: {processed}/{len(iin_list)}")
    print(f"   Успешно: {successful}")
    print(f"   Ошибки: {len(errors)}")
    print(f"   ⏱️ Общее время: {total_time / 60:.1f} мин")
    print(f"   ⚡ Среднее время на ИИН: {avg_time:.1f}с")
    if errors:
        print(f"   ❌ Список ошибок: {errors[:10]}")
    print("=" * 60)


# =========================
# ЗАПУСК
# =========================

driver = None
start_run_logging(JOB_DIR)

try:
    print("***", name_login, "***")
    print("Папка запуска:", JOB_DIR)

    driver = to_site()
    write_login(driver, name_login)

    driver.get(AISOIP_URL)
    WebDriverWait(driver, 60).until(
        lambda d: "aisoip.adilet.gov.kz" in (d.current_url or "").lower()
    )

    if args.iin and args.iin.strip():
        run_debug_dump(driver, args.iin.strip())
    else:
        run_full_check(driver)

    print()
    print("=" * 90)
    print("СКРИПТ ОГРАНИЧЕНИЯ НА ВЫЕЗД УСПЕШНО ЗАВЕРШЁН")
    print("=" * 90)

except Exception:
    print("КРИТИЧЕСКАЯ ОШИБКА В СКРИПТЕ ОГРАНИЧЕНИЯ НА ВЫЕЗД")
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
