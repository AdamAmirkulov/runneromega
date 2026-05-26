# -*- coding: utf-8 -*-
"""
АИС ОИП — Статусы (Omega, API-выгрузка).

Извлечено из AISOIP_Status_plus_Otmeny_API_combined.ipynb (Этап 1):
1) выгрузка АИС ОИП выполняется через API-механику;
2) логика формирования итогового файла сохранена из Omega_status.py;
3) вход в АИС ОИП автоматический;
4) весь вывод stdout/stderr автоматически дублируется в run.log.

Этап 2 (Отмены) вынесен в отдельный скрипт scripts/otmeny.py.
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

from config import CREDENTIALS, _company_id

print(f"🏢 Запуск АИС ОИП статусов для компании ID={_company_id}")

# =========================
# ГЛОБАЛЬНЫЕ ПАРАМЕТРЫ
# =========================

list_not_data_for_view = []

# Пути — свои для каждой компании (папка на компании: load / load_finish /
# credentials v3.json), берутся из config.py по COMPANY_ID.
loading_path_part = CREDENTIALS['loading_path_part']
path_load_finish = CREDENTIALS['path_load_finish']
path_crm = CREDENTIALS['path_crm']
run_prod = True

# Логин/пароль АИС ОИП берутся из config.py (COMPANY_CREDENTIALS) по company_id,
# так же как в остальных скриптах — больше не захардкожены на один аккаунт.
name_login = f"company{_company_id}"
dict_log_pas = {
    name_login: {
        "log": CREDENTIALS['aisoip_login'],
        "pas": CREDENTIALS['aisoip_password'],
    }
}

# API — механика из первого скрипта
AISOIP_URL = "https://aisoip.adilet.gov.kz/cabinet/exec-productions"
API_SEARCH_URL = "https://aisoip.adilet.gov.kz/extperson/api/rest/execproc/search"
API_EXPORT_URL = "https://aisoip.adilet.gov.kz/extperson/api/rest/export/excel"

DATE_FROM = "2023-01-01"
EXPORT_REQUEST_SIZE = 50_000
API_RETRIES = 3
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
from io import BytesIO
from pathlib import Path
from datetime import date, datetime

import base64
import json
import math
import shutil
import time
import traceback

import numpy as np
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
    """Создает папку запуска и all_file для частей массового API-экспорта."""
    dt_string = datetime.now().strftime("%Y-%m-%d___%H-%M")
    loading_path = Path(loading_path_part) / f"{name_login} {dt_string}"
    loading_path_all = loading_path / "all_file"

    loading_path_all.mkdir(parents=True, exist_ok=False)

    return (
        str(loading_path),
        str(loading_path_all),
        str(loading_path_all),
        str(loading_path_all),
    )


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


def write_login(driver, name_login, force=False):
    """
    Автоматически вводит логин/пароль, если открыта страница авторизации.

    force=True пропускает проверку is_login_page() и сразу пытается
    залогиниться — используется, когда API вернул 401 несмотря на то,
    что is_login_page() посчитал сессию активной (DOM-проверка сразу
    после driver.get() иногда срабатывает раньше, чем долетает редирект
    через SSO, и ошибочно считает протухшую/ещё не готовую страницу
    авторизованной).
    """
    if not force and not is_login_page(driver):
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


def browser_fetch_bytes(driver, url):
    """Скачивает бинарный Excel через авторизованный API fetch внутри Chrome."""
    script = r"""
    const done = arguments[arguments.length - 1];
    const url = arguments[0];

    fetch(url, {
        method: 'GET',
        credentials: 'include',
        headers: {'Accept': 'application/vnd.ms-excel,application/octet-stream,*/*'}
    })
    .then(async response => {
        const blob = await response.blob();
        const reader = new FileReader();

        reader.onloadend = () => done({
            status: response.status,
            contentType: response.headers.get('content-type') || '',
            dataUrl: reader.result
        });

        reader.onerror = () => done({
            status: 0,
            contentType: '',
            dataUrl: ''
        });

        reader.readAsDataURL(blob);
    })
    .catch(error => done({
        status: 0,
        contentType: '',
        dataUrl: String(error)
    }));
    """

    result = driver.execute_async_script(script, url)

    if result.get("status") != 200:
        raise RuntimeError(
            f"API export error {result.get('status')} for {url}: {str(result)[:1000]}"
        )

    data_url = result.get("dataUrl") or ""
    if "," not in data_url:
        raise RuntimeError(
            f"Экспорт не вернул бинарный файл: {str(result)[:1000]}"
        )

    content = base64.b64decode(data_url.split(",", 1)[1])

    # XLSX = ZIP/PK, старый XLS = OLE D0 CF 11 E0
    if not (
        content.startswith(b"PK")
        or content.startswith(bytes.fromhex("D0CF11E0"))
    ):
        raise RuntimeError(
            f"Ответ экспорта не похож на Excel. "
            f"Content-Type={result.get('contentType')}, "
            f"первые байты={content[:20]!r}"
        )

    return content


def api_call_with_retry(callable_obj, description):
    """Автоматические повторные попытки API."""
    last_error = None

    for attempt in range(1, API_RETRIES + 1):
        try:
            print(f"API: {description}. Попытка {attempt}/{API_RETRIES}")
            result = callable_obj()
            print(f"API: {description} — успешно")
            return result

        except Exception as exc:
            last_error = exc
            print(
                f"API ERROR: {description}. "
                f"Попытка {attempt}/{API_RETRIES}: {exc}"
            )
            traceback.print_exc()

            if attempt < API_RETRIES:
                time.sleep(2 * attempt)

    raise RuntimeError(
        f"Не удалось выполнить: {description}"
    ) from last_error


def create_mass_search(driver, date_from, date_to):
    """Создает массовый поиск АИС ОИП и возвращает searchId + totalElements."""
    payload = {
        "fromDate": date_from,
        "toDate": date_to,
        "searchType": False,
    }

    url = f"{API_SEARCH_URL}?page=0&size=5"

    data = api_call_with_retry(
        lambda: browser_fetch_json(
            driver,
            url,
            method="POST",
            body=payload,
        ),
        "создание массового поиска АИС ОИП",
    )

    pagination = data.get("pagination") or {}
    search_id = pagination.get("searchId")
    total_elements = int(pagination.get("totalElements") or 0)

    if total_elements > 0 and not search_id:
        raise RuntimeError(
            "API вернул строки, но не вернул pagination.searchId"
        )

    print(f"Период API: {date_from} — {date_to}")
    print(f"Всего найдено в АИС ОИП: {total_elements}")
    print(f"searchId: {search_id}")

    return search_id, total_elements


def read_export_excel_bytes(content):
    """Читает Excel из памяти, ИИН сохраняет строкой."""
    return pd.read_excel(
        BytesIO(content),
        dtype={"ИИН/БИН должника": str},
    )


def export_mass_pages(driver, search_id, total_elements, output_dir):
    """
    Выгружает все страницы через /export/excel.

    Запрашивается size=50000, но фактический лимит сервера определяется
    автоматически по первой выгруженной странице.
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if total_elements == 0:
        return pd.DataFrame()

    def load_page(page):
        url = (
            f"{API_EXPORT_URL}?searchtype=false"
            f"&page={page}&size={EXPORT_REQUEST_SIZE}"
            f"&searchid={search_id}&lang=ru"
        )

        content = api_call_with_retry(
            lambda url=url: browser_fetch_bytes(driver, url),
            f"экспорт страницы {page}",
        )

        part_path = output_dir / f"aisoip_export_page_{page:04d}.xlsx"
        part_path.write_bytes(content)

        df_part = read_export_excel_bytes(content)

        print(
            f"Страница {page}: {len(df_part)} строк; "
            f"файл: {part_path}"
        )

        return df_part

    first_part = load_page(0)
    real_page_size = len(first_part)

    if real_page_size <= 0:
        raise RuntimeError(
            "Первая страница массового экспорта пуста, "
            f"хотя totalElements={total_elements}"
        )

    page_count = math.ceil(total_elements / real_page_size)

    print(f"Запрошенный размер страницы: {EXPORT_REQUEST_SIZE}")
    print(f"Фактический размер страницы: {real_page_size}")
    print(f"Количество частей: {page_count}")

    parts = [first_part]

    for page in tqdm(
        range(1, page_count),
        total=max(page_count - 1, 0),
        desc="Массовый экспорт АИС ОИП",
    ):
        df_part = load_page(page)

        if df_part.empty and sum(len(x) for x in parts) < total_elements:
            raise RuntimeError(
                f"Страница {page} оказалась пустой "
                "до получения всех строк"
            )

        parts.append(df_part)

    df_all = pd.concat(parts, ignore_index=True)

    print(f"Суммарно прочитано из API: {len(df_all)} строк")

    if len(df_all) < total_elements:
        raise RuntimeError(
            f"Получено меньше строк, чем сообщил API: "
            f"{len(df_all)} < {total_elements}"
        )

    if len(df_all) > total_elements:
        print(
            f"ВНИМАНИЕ: получено {len(df_all)} строк "
            f"при totalElements={total_elements}. "
            "Далее удаляются только точные дубли."
        )

    return df_all


# =========================
# ПОДГОТОВКА API-ВЫГРУЗКИ
# =========================

def parse_mixed_date_series(series):
    """
    Приводит даты к datetime.
    Никакого выбора одного ИП по ИИН здесь нет:
    выбор нужного производства выполняет исходная логика Omega_status.py.
    """
    result = pd.to_datetime(series, errors="coerce")

    missing = result.isna() & series.notna()

    if missing.any():
        result.loc[missing] = pd.to_datetime(
            series.loc[missing].astype(str),
            format="%d.%m.%Y",
            errors="coerce",
        )

    return result


def prepare_api_export_for_omega(df_raw):
    """
    Подготавливает API-выгрузку к исходной бизнес-логике Omega_status.py:
    - оставляет штатные колонки АИС ОИП;
    - ИИН хранит строкой;
    - приводит даты;
    - удаляет только точные дубли;
    - заново формирует №.
    """
    required = [
        "№",
        "Номер исполнительного документа",
        "Номер исполнительного производства",
        "Судебный исполнитель",
        "Дата возбуждения",
        "Дата выписки исполнительного документа",
        "Орган выдавший исполнительный документ",
        "Тип взыскания",
        "Должник",
        "ИИН/БИН должника",
        "Статус исполнительного производства",
        "Сумма",
    ]

    if df_raw.empty:
        return pd.DataFrame(columns=required)

    missing = [
        col for col in required
        if col not in df_raw.columns
    ]

    if missing:
        raise KeyError(
            f"В API Excel отсутствуют колонки: {missing}"
        )

    df = df_raw[required].copy()

    df["ИИН/БИН должника"] = (
        df["ИИН/БИН должника"]
        .fillna("")
        .astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .str.strip()
    )

    df["Сумма"] = pd.to_numeric(
        df["Сумма"],
        errors="coerce",
    )

    for column in [
        "Дата возбуждения",
        "Дата выписки исполнительного документа",
    ]:
        df[column] = parse_mixed_date_series(df[column])

    before = len(df)
    df = df.drop_duplicates().reset_index(drop=True)

    # Нумерация как в исходном Omega_status.py
    if "№" in df.columns:
        df = df.drop(columns=["№"])

    df.insert(0, "№", np.arange(1, len(df) + 1))

    print(f"Точных дублей удалено: {before - len(df)}")
    print(f"Строк перед логикой формирования Omega: {len(df)}")

    return df


# =========================
# ВЫГРУЗКА АИС ОИП ЧЕРЕЗ API
# =========================

driver = None
df_final = None
loading_path = None
loading_path_d = None
loading_path_all = None
dowlend_path = None
name_login = None

for name_login in dict_log_pas.keys():
    list_not_data_for_view = []

    loading_path, loading_path_d, loading_path_all, dowlend_path = make_load_dir(
        name_login,
        loading_path_part,
    )

    start_run_logging(loading_path)

    try:
        print("***", name_login, "***")
        print("Папка запуска:", loading_path)

        driver = to_site(loading_path_all)

        # Автоматическая авторизация.
        write_login(driver, name_login)

        # API fetch выполняется внутри авторизованной вкладки.
        driver.get(AISOIP_URL)

        WebDriverWait(driver, 60).until(
            lambda d: "aisoip.adilet.gov.kz" in (d.current_url or "").lower()
        )

        date_to = date.today().strftime("%Y-%m-%d")

        def _exception_chain_mentions_401(exc):
            # api_call_with_retry оборачивает исходную ошибку в новый
            # RuntimeError ("Не удалось выполнить: ..."), само "401"
            # остаётся только в исходном exc.__cause__ — проверяем всю цепочку.
            seen_ids = set()
            current = exc
            while current is not None and id(current) not in seen_ids:
                seen_ids.add(id(current))
                if "401" in str(current):
                    return True
                current = current.__cause__
            return False

        try:
            search_id, total_elements = create_mass_search(
                driver=driver,
                date_from=DATE_FROM,
                date_to=date_to,
            )
        except RuntimeError as exc:
            # is_login_page() иногда ошибочно решает, что сессия уже
            # активна (см. write_login), хотя реально она протухла —
            # тогда самый первый API-запрос падает с 401. Не сдаёмся
            # сразу: заново заходим на страницу и логинимся принудительно,
            # затем пробуем массовый поиск ещё раз.
            if not _exception_chain_mentions_401(exc):
                raise

            print("АИС ОИП: получен 401, сессия оказалась протухшей. Принудительный повторный вход.")
            driver.get(AISOIP_URL)
            WebDriverWait(driver, 60).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            write_login(driver, name_login, force=True)

            driver.get(AISOIP_URL)
            WebDriverWait(driver, 60).until(
                lambda d: "aisoip.adilet.gov.kz" in (d.current_url or "").lower()
            )

            search_id, total_elements = create_mass_search(
                driver=driver,
                date_from=DATE_FROM,
                date_to=date_to,
            )

        df_export_raw = export_mass_pages(
            driver=driver,
            search_id=search_id,
            total_elements=total_elements,
            output_dir=loading_path_all,
        )

        if df_export_raw.empty:
            raise RuntimeError(
                "Массовый API-экспорт АИС ОИП не вернул строк."
            )

        # ВАЖНО:
        # не применяем логику выбора одного ИП на ИИН из первого ноутбука,
        # т.к. выбор выполняется ниже исходной логикой Omega_status.py.
        df_final = prepare_api_export_for_omega(df_export_raw)

        now_str = datetime.now().strftime("%d.%m.%Y")
        full_export_path = (
            Path(loading_path)
            / f"{name_login} AIS_OIP {now_str}.xlsx"
        )

        df_final.to_excel(
            full_export_path,
            index=False,
        )

        print("Массовая API-выгрузка сохранена:", full_export_path)
        print("Итоговое количество строк API:", len(df_final))

    except Exception:
        print("КРИТИЧЕСКАЯ ОШИБКА НА ЭТАПЕ API-ВЫГРУЗКИ")
        traceback.print_exc()
        raise

    finally:
        # Этот скрипт больше не продолжается вторым этапом (отмены) —
        # тот вынесен в отдельный процесс scripts/otmeny.py со своим логином.
        # Chrome закрывается ниже, после формирования итогового файла.
        pass

# Ниже без изменения бизнес-логики идет формирование итогового файла
# из Omega_status.py.

try:
    import string
    from pprint import pprint
    import pandas as pd
    import gspread
    from gspread import Cell, Client, Spreadsheet, Worksheet
    from gspread.utils import rowcol_to_a1
    import requests


    # Своя Google-таблица (и свой лист) на каждую компанию — из config.py.
    SPREADSHEET_URL = CREDENTIALS['spreadsheet_url']
    SHEET_NAME = CREDENTIALS['sheet_name']



    def show_available_worksheets(sh: Spreadsheet):
        worksheets = sh.worksheets()

        for ws in worksheets:
            print("Worksheet with title", repr(ws.title), "and id", ws.id)


    def get_dataframe_from_sheet(sh: Spreadsheet, sheet_title: str) -> pd.DataFrame:
        # Получаем рабочий лист по названию
        worksheet = sh.worksheet(sheet_title)

        # Получаем все данные из листа
        data = worksheet.get_all_values()

        # Создаем DataFrame из полученных данных
        # Первая строка используется в качестве заголовков столбцов
        df = pd.DataFrame(data[1:], columns=data[0])

        return df


    def get_google_df():
        gc: Client = gspread.service_account(CREDENTIALS['credentials_path'])
        sh: Spreadsheet = gc.open_by_url(SPREADSHEET_URL)

        show_available_worksheets(sh)
        df = get_dataframe_from_sheet(sh, SHEET_NAME)
        return df

    df_google_t = get_google_df()

    # %% [markdown]
    # ### Редактирую данные из бд.

    # %%
    def to_int(v):
        try:
            return int(v)  # Преобразуем в int
        except (ValueError, TypeError):
            return 0  # Если ошибка — вернуть 0

        v = v.replace(' ', '')
        v = v.replace('-', '')
        v = v.replace(',', '.')
        v = v.replace('\xa0', '')
        v = v.split('.')[0]
        if len(v) == 0:
            return 0
        else:
            return int(v)


    df_google_t['ОД_claen'] = df_google_t['ОД'].apply(to_int)


    # В таблице компании 2 колонку "госпошина по суду" в какой-то момент
    # разбили на "госпошина по ауэз суду" и "госпошина по мед суду" (у
    # каждой строки заполнена только одна из двух). У остальных компаний
    # колонка осталась единой. Приводим к единой колонке "госпошина по
    # суду", которую использует остальной код (f1103, f1107 и т.д.),
    # независимо от того, какой вариант структуры сейчас в таблице.
    if 'госпошина по суду' not in df_google_t.columns:
        split_cols = ['госпошина по ауэз суду', 'госпошина по мед суду']
        found_split_cols = [c for c in split_cols if c in df_google_t.columns]
        if found_split_cols:
            df_google_t['госпошина по суду'] = sum(
                df_google_t[c].apply(to_int) for c in found_split_cols
            )

    columns = ['Сумма выдачи', 'ОД', 'Вознаграждение', 'Пеня', 'Всего задолженность при покупке', 'нотариальные услуги', 'сумма оплаты после цессии', 'остаток задолженности', 'Сумма оплат', 'госпошина по суду']
    for col in columns:
        print('col', col)
        df_google_t[col] = df_google_t[col].apply(to_int)
    # df = df['Всего задолженность при покупке']


    # %%
    df_google_t['ИИН'] = df_google_t['ИИН'].apply(lambda x: '0' * (12 - len(str(x))) + str(x) )

    # %% [markdown]
    # Добавить колонки.

    # %%
    # Остаток задолжности
    # Тотал по покупке + нат надпись.

    # адк названия. (Вроде так. Возможны расхождения.)
    # Value # ОД
    # f140   # Тотал (за вычетом погашений)
    # f148     # Тотал (по покупке) # Всего задолженность при покупке

    # f114,    # Тотал (на момент передачи нотариусу)
    # f118,    # Тотал (на момент иска в суд)
    # f147,    # Тотал (по решению суда)
    # f386    # Тотал (по медиации/мировому соглашению)

    # %%
    # omega названия

    # 1) f1101 - Всего задолженность при покупке
    # 2) f1102 - Всего задолженность при покупке + нотариальные услуги
    # 3) f1103 - Всего задолженность при покупке + госпошина по суду
    # 4) f1104 - Всего задолженность при покупке + госпошина по суду + нотариальные услуги

    # 5) f1105 - остаток задолженности
    # 6) f1106 - остаток задолженности + нотариальные услуги
    # 7) f1107 - остаток задолженности + госпошина по суду
    # 8) f1108 - остаток задолженности + госпошина по суду + нотариальные услуги

    # 9) f1109 - сумма АИСОИП.

    # %%
    df_google_t['f1101'] = df_google_t['Всего задолженность при покупке']
    df_google_t['f1102'] = df_google_t['Всего задолженность при покупке'] + df_google_t['нотариальные услуги']
    df_google_t['f1103'] = df_google_t['Всего задолженность при покупке'] + df_google_t['госпошина по суду']
    df_google_t['f1104'] = df_google_t['Всего задолженность при покупке'] + df_google_t['нотариальные услуги'] + df_google_t['нотариальные услуги']

    df_google_t['f1105'] = df_google_t['остаток задолженности']
    df_google_t['f1106'] = df_google_t['остаток задолженности'] + df_google_t['нотариальные услуги']
    df_google_t['f1107'] = df_google_t['остаток задолженности'] + df_google_t['госпошина по суду']
    df_google_t['f1108'] = df_google_t['остаток задолженности'] + df_google_t['нотариальные услуги'] + df_google_t['нотариальные услуги']

    df_google_t['f1109'] = df_google_t['Сумма по АИСОИП']

    df_google_t['f1110'] = df_google_t['остаток задолженности'] - 685

    # %%
    recode_old_new = [('ИИН', 'inn'),
    ('index', 'eid'),
    ('index', 'id'),
    ('f1101', 'f1101'),
    ('f1102', 'f1102'),
    ('f1103', 'f1103'),
    ('f1104', 'f1104'),
    ('f1105', 'f1105'),
    ('f1106', 'f1106'),
    ('f1107', 'f1107'),
    ('f1108', 'f1108'),
    ('f1109', 'f1109'),
    ('f1110', 'f1110'),

    ('Номер договора цессии', 'f258'), # Это не правильное поле !!! TODO. Надо добавить в место него "Номер испол. производ" из АИСОИП
    # ('ОД', 'Value'),  # 'ОД'
    # ('остаток задолженности', 'f140'),   # Тотал (за вычетом погашений)
    # ('f148_v1', 'f148'),     # Тотал (по покупке) # Всего задолженность при покупке

    # ('ОД_claen', 'f114'),    # Тотал (на момент передачи нотариусу)
    # ('ОД_claen', 'f118'),    # Тотал (на момент иска в суд)
    # ('ОД_claen', 'f147'),    # Тотал (по решению суда)
    # ('ОД_claen', 'f386'),    # Тотал (по медиации/мировому соглашению)
    ]

    df_google_t['index'] = df_google_t.index

    # %%
    # df.to_excel('df db.xlsx')

    # %%
    df_db_clients = pd.DataFrame() # df[['eid', ]] #.head()
    for old, new in recode_old_new:
        addition_col = df_google_t[old]
        df_db_clients[new] = addition_col
        print('old, new', old, new)

    # %% [markdown]
    # # Скачиваю данные AISOIP

    # %%
    # df_ais = pd.read_excel('Взыскатель-20240225-21-49.xlsx')


    df_ais = df_final.copy()

    # %%
    df_ais['ИИН/БИН должника'] = df_ais['ИИН/БИН должника'].apply(lambda x: '0' * (12 - len(str(x))) + str(x) )

    # %%
    ais_col = ['№',
    'Номер исполнительного документа',
    'Номер исполнительного производства',
    'Судебный исполнитель',
    'Дата возбуждения',
    'Дата выписки исполнительного документа',
    'Орган выдавший исполнительный документ',
    'Тип взыскания',
    'Должник',
    'ИИН/БИН должника',
    'Статус исполнительного производства',
    'Сумма']

    df_ais_oip_raw_data = df_ais[ais_col]

    # %% [markdown]
    # ### Загрузка данных в локальную БД SQLite

    # %%
    import sqlite3
    import pandas as pd

    # Предположим, что у вас уже есть DataFrame под названием df

    # Создание соединения с базой данных SQLite
    conn = sqlite3.connect('my_database.sqlite')

    # Загрузка данных из DataFrame в таблицу SQLite
    df_ais_oip_raw_data.to_sql('AKA_ais_oip_raw_data', conn, if_exists='replace', index=False)
    df_db_clients.to_sql('db_clients', conn, if_exists='replace', index=False)

    # Закрытие соединения с базой данных
    conn.close()


    # %% [markdown]
    # ### Запрос на SQLite

    # %%
    sql = '''
    with doubles as
    (
        select t.*
             , ROW_NUMBER() over(partition by "Номер исполнительного документа",
                                              "Номер исполнительного производства",
                                              "Судебный исполнитель",
                                              "ИИН/БИН должника",
                                              "Статус исполнительного производства",
                                              "Сумма",
                                              "Дата возбуждения"
                                 order by "Дата возбуждения" desc
                              ) as rn
          from AKA_ais_oip_raw_data t
    ),
    no_doubles as
    (
        select *
          from doubles
         where rn = 1
    ),
    clients as
    (
        select id, inn, f258, eid
             , IFNULL(round(f1101, 0), 0) as f1101
             , IFNULL(round(f1102, 0), 0) as f1102
             , IFNULL(round(f1103, 0), 0) as f1103
             , IFNULL(round(f1104, 0), 0) as f1104
             , IFNULL(round(f1105, 0), 0) as f1105
             , IFNULL(round(f1106, 0), 0) as f1106
             , IFNULL(round(f1107, 0), 0) as f1107
             , IFNULL(round(f1108, 0), 0) as f1108
             , IFNULL(round(f1109, 0), 0) as f1109
             , IFNULL(round(f1110, 0), 0) as f1110
          from db_clients c
          -- Примечание: Убедитесь, что соединения таблиц и условия соответствуют вашему запросу
          -- Этот блок может потребовать дополнительной адаптации в соответствии с вашей схемой БД
    ),
    right_data as
    (
        select d."№"
             , d."Номер исполнительного документа"
             , d."Номер исполнительного производства"
             , d."Судебный исполнитель"
             , d."Дата возбуждения"
             , d."Дата выписки исполнительного документа"
             , d."Орган выдавший исполнительный документ"
             , d."Тип взыскания"
             , d."Должник"
             , d."ИИН/БИН должника"
             , d."Статус исполнительного производства"
             , d."Сумма"
             , c.eid as "Уникальный идентификатор займа"
          from no_doubles d
          join clients c
            on c.inn = d."ИИН/БИН должника"
           and c.f258 = d."Номер исполнительного производства"
    ),



    wrong_data as
    (
        select d."№"
             , d."Номер исполнительного документа"
             , d."Номер исполнительного производства"
             , d."Судебный исполнитель"
             , d."Дата возбуждения"
             , d."Дата выписки исполнительного документа"
             , d."Орган выдавший исполнительный документ"
             , d."Тип взыскания"
             , d."Должник"
             , d."ИИН/БИН должника"
             , d."Статус исполнительного производства"
             , d."Сумма"
          from no_doubles d
        except
        select r."№"
             , r."Номер исполнительного документа"
             , r."Номер исполнительного производства"
             , r."Судебный исполнитель"
             , r."Дата возбуждения"
             , r."Дата выписки исполнительного документа"
             , r."Орган выдавший исполнительный документ"
             , r."Тип взыскания"
             , r."Должник"
             , r."ИИН/БИН должника"
             , r."Статус исполнительного производства"
             , r."Сумма"
          from right_data r
    ),
    right_data2 as
    (
        select d.*, c.eid as "Уникальный идентификатор займа"
          from wrong_data d
          join clients c
            on c.inn = d."ИИН/БИН должника"
           and (
                (d."Сумма" between c.f1101-3 and c.f1101+3 and c.f1101 != 0)
                or (d."Сумма" between c.f1102-3 and c.f1102+3 and c.f1102 != 0)
                or (d."Сумма" between c.f1103-3 and c.f1103+3 and c.f1103 != 0)
                or (d."Сумма" between c.f1104-3 and c.f1104+3 and c.f1104 != 0)
                or (d."Сумма" between c.f1105-3 and c.f1105+3 and c.f1105 != 0)
                or (d."Сумма" between c.f1106-3 and c.f1106+3 and c.f1106 != 0)
                or (d."Сумма" between c.f1107-3 and c.f1107+3 and c.f1107 != 0)
                or (d."Сумма" between c.f1108-3 and c.f1108+3 and c.f1108 != 0)
                or (d."Сумма" between c.f1109-3 and c.f1109+3 and c.f1109 != 0)
                or (d."Сумма" between c.f1110-3 and c.f1110+3 and c.f1110 != 0)

               )
    ),
    wrong2 as
    (
        select d.*, null as eid
          from wrong_data d
          left join clients c
            on c.inn = d."ИИН/БИН должника"
           and (
                (d."Сумма" between c.f1101-3 and c.f1101+3 and c.f1101 != 0)
                or (d."Сумма" between c.f1102-3 and c.f1102+3 and c.f1102 != 0)
                or (d."Сумма" between c.f1103-3 and c.f1103+3 and c.f1103 != 0)
                or (d."Сумма" between c.f1104-3 and c.f1104+3 and c.f1104 != 0)
                or (d."Сумма" between c.f1105-3 and c.f1105+3 and c.f1105 != 0)
                or (d."Сумма" between c.f1106-3 and c.f1106+3 and c.f1106 != 0)
                or (d."Сумма" between c.f1107-3 and c.f1107+3 and c.f1107 != 0)
                or (d."Сумма" between c.f1108-3 and c.f1108+3 and c.f1108 != 0)
                or (d."Сумма" between c.f1109-3 and c.f1109+3 and c.f1109 != 0)
                or (d."Сумма" between c.f1110-3 and c.f1110+3 and c.f1110 != 0)
               )
          where c.id is null
    ),
    right_data3 as
    (
        select d."№"
             , d."Номер исполнительного документа"
             , d."Номер исполнительного производства"
             , d."Судебный исполнитель"
             , d."Дата возбуждения"
             , d."Дата выписки исполнительного документа"
             , d."Орган выдавший исполнительный документ"
             , d."Тип взыскания"
             , d."Должник"
             , d."ИИН/БИН должника"
             , d."Статус исполнительного производства"
             , d."Сумма"
             , c.eid as eid
          from wrong2 d
          join (
                select c.inn, count(*) as cnt_iin, max(eid) as eid
                from clients c
                group by c.inn
               ) c
            on c.inn = d."ИИН/БИН должника"
           and c.cnt_iin = 1
    ),


    wrong3 as
    (
        select *
        from wrong2 w2
        where not exists (
            select 1 from right_data3 rd3 where rd3."ИИН/БИН должника" = w2."ИИН/БИН должника"
        )
    ),


    fin as
    (
        select *,
               case
                   when "Статус исполнительного производства" = 'На исполнении' then 1
                   else 2
               end as priority_id_status
          from right_data
        union all
        select *,
               case
                   when "Статус исполнительного производства" = 'На исполнении' then 1
                   else 2
               end as priority_id_status
          from right_data2
        union all
        select *,
               case
                   when "Статус исполнительного производства" = 'На исполнении' then 1
                   else 2
               end as priority_id_status
          from right_data3
    ),


    fin2 as
    (
        select *,
               ROW_NUMBER() over(
                   partition by "Уникальный идентификатор займа"
                   order by priority_id_status, "Дата возбуждения" desc, "Дата выписки исполнительного документа" desc, "Статус исполнительного производства"
               ) as rn
          from fin
    )
    select "№",
           "Номер исполнительного документа",
           "Номер исполнительного производства",
           "Судебный исполнитель",
           strftime('%d.%m.%Y', "Дата возбуждения") as "Дата возбуждения",
           strftime('%d.%m.%Y', "Дата выписки исполнительного документа") as "Дата выписки исполнительного документа",
           "Орган выдавший исполнительный документ",
           "Тип взыскания",
           "Должник",
           "ИИН/БИН должника",
           "Статус исполнительного производства",
           "Сумма",
           "Уникальный идентификатор займа"
    from fin2
    where rn = 1
    union
    select "№",
           "Номер исполнительного документа",
           "Номер исполнительного производства",
           "Судебный исполнитель",
           strftime('%d.%m.%Y', "Дата возбуждения") as "Дата возбуждения",
           strftime('%d.%m.%Y', "Дата выписки исполнительного документа") as "Дата выписки исполнительного документа",
           "Орган выдавший исполнительный документ",
           "Тип взыскания",
           "Должник",
           "ИИН/БИН должника",
           "Статус исполнительного производства",
           "Сумма",
           null
    from wrong3
    '''

    # %%
    # Открытие соединения с базой данных SQLite
    conn = sqlite3.connect('my_database.sqlite')

    # Пример: выбрать все данные из таблицы db_clients
    query = sql
    df_sql = pd.read_sql_query(query, conn)

    # Закрытие соединения с базой данных
    conn.close()


    # %%
    df_sql['Уникальный идентификатор займа'] = df_sql['Уникальный идентификатор займа'].astype('Int64')

    # %%
    # Проверки


    col = '№'
    cnt_load_check = df_sql.groupby(['Уникальный идентификатор займа'], as_index=False).agg({col:'count'}).rename(columns={col:'cnt_load'})
    cnt_duble = (cnt_load_check['cnt_load'] > 1).sum()

    print('Кол-во дублей', cnt_duble)
    print('Кол-во пустых значений', df_sql['Уникальный идентификатор займа'].isna().sum())

    is_finish = (df_google_t['статус кредита'] == 'погашен') | (df_google_t['статус кредита'] == 'обратный выкуп')

    df_google_t.loc[~is_finish, 'index'].isin(df_sql['Уникальный идентификатор займа'])
    cnt_not_pull = (~df_google_t.loc[~is_finish, 'index'].isin(df_sql['Уникальный идентификатор займа'])).sum()

    print('Кол-во дел по которым не подтянулось (из незавершонных)', cnt_not_pull)

    # %%
    # Подтягиваю под исходный формат
    df_google_t.rename(columns={'index':'Номер_строки_google_t'}, inplace=True)
    df_google_t['Номер_строки_google_t'] = df_google_t['Номер_строки_google_t'].astype('Int64')

    df_google_t['Номер_строки_google_t'] = df_google_t['Номер_строки_google_t'].astype('Int64')
    final_table = df_google_t[['Номер_строки_google_t', 'ИИН']].merge(df_sql, how='left', left_on=['Номер_строки_google_t'], right_on=['Уникальный идентификатор займа'])

    # %%
    final_table.sort_values('Номер_строки_google_t', inplace=True)

    # %%
    def uid_to_string(uid, first_num=1, cnt_num=6):
        """
        uid = unique_id
        suid = string_uid

        input 1
        output 100_001

        input 11
        output 100_011

        """

        suid = str(uid)
        cnt_zero =  cnt_num - 1 - len(suid)
        new_suid = str(first_num) + "0" * cnt_zero + suid
        return new_suid


    final_table['уникальный номер'] = final_table['Номер_строки_google_t'] + 1     # Старт не с "0" а с "1".
    # Первая цифра EID своя на компанию: Omega=1, KPI=2, Orion=3, InvestWay=3.
    final_table['уникальный номер'] = final_table['уникальный номер'].apply(
        lambda uid: uid_to_string(uid, first_num=CREDENTIALS['eid_prefix'])
    )
    final_table = final_table.drop('Уникальный идентификатор займа', axis=1)

    # %%
    now_str = datetime.now().strftime("%d.%m.%Y")
    final_table.to_excel(f'{loading_path}/{name_login} AIS_OIP_merge_google_table {now_str}.xlsx', index=False)

    # %%
    df_ais_not_processed = df_ais_oip_raw_data.loc[(~df_ais_oip_raw_data['Номер исполнительного производства'].isin(df_sql["Номер исполнительного производства"])) & (df_ais_oip_raw_data['Статус исполнительного производства'] == 'На исполнении')]
    df_ais_not_processed.to_excel(f'{loading_path}/{name_login} Производства не подтянулись .xlsx', index=False)

    # %% [markdown]
    # # Убираю из проверенных отмененных (не отменённые статусы)

    # %%
    def clean_cancel_list(new_clean_df, path_load_finish):
        '''
        Очитаю эскль в котором хранятся данные по отменам которыне не нужно проверять.
        '''
        new_clean_df['Дата возбуждения'] = pd.to_datetime(new_clean_df['Дата возбуждения'], format='%d.%m.%Y')
        new_clean_df = new_clean_df.sort_values(by=['Дата возбуждения'], ascending=False).drop_duplicates('ИИН/БИН должника', keep='first')


        path_check_iin = Path(path_load_finish).joinpath('Проверенные ранее.xlsx')
        df_old_ = pd.read_excel(path_check_iin, dtype='str')
        print(list(df_old_))
        df_old_['ИИН_and_id_Испол'] = df_old_['ИИН'] + '_' + df_old_['Номер исполнительного производства']
        new_clean_df['ИИН_and_id_Испол'] = new_clean_df['ИИН'] + '_' + new_clean_df['Номер исполнительного производства']


        df_drop = new_clean_df.loc[new_clean_df['Статус исполнительного производства'] != 'Окончено', 'ИИН_and_id_Испол']
        df_new_check = df_old_[~df_old_['ИИН_and_id_Испол'].isin(df_drop)].copy()
        print(list(df_new_check))
        df_new_check.to_excel(path_check_iin, index=False)

        print(f'Удалено из проверенных {df_old_.shape[0] - df_new_check.shape[0]}')


    clean_cancel_list(final_table, path_load_finish)

    # %%
    path_load_finish

    # %%
    final_table

    # %% [markdown]
    # # Создание дополнительного файла с форматами

    # %% [markdown]
    # ## Поиск отмен

    # %%
    # import os
    import re
    # from datetime import datetime
    import glob

    def get_last_cancel_path(load_path):
        # Регулярное выражение для поиска даты в названии папок
        date_pattern = re.compile(r"(\d{4}-\d{2}-\d{2})___\d{2}-\d{2}")

        latest_date = None
        latest_folder = None

        for folder in os.listdir(load_path):
            folder_path = os.path.join(load_path, folder)

            if not os.path.isdir(folder_path):  # Пропускаем файлы
                continue

            match = date_pattern.match(folder)
            if match:
                folder_date = datetime.strptime(match.group(1), "%Y-%m-%d")
                if latest_date is None or folder_date > latest_date:
                    latest_date = folder_date
                    latest_folder = folder

        print(f"Самая последняя папка: {latest_folder}")
        return latest_folder

    latest_folder_cancel = get_last_cancel_path(path_load_finish)

    if latest_folder_cancel is None:
        print("Папки с датой не найдены в path_load_finish. Пропускаем поиск файла отмен.")
        df_last_cancel = pd.DataFrame()  # пустой датафрейм как заглушка
    else:
        latest_folder_path = os.path.join(path_load_finish, latest_folder_cancel)

        file_pattern = os.path.join(latest_folder_path, "*Итог_union*")
        matching_files = glob.glob(file_pattern)

        if matching_files:
            target_file = matching_files[0]
            print(f"Найденный файл: {target_file}")
            df_last_cancel = pd.read_excel(target_file)
        else:
            print("Файл с 'Итог_union' не найден.")
            df_last_cancel = pd.DataFrame()

    # %%
    # Полный путь к последней папке
    #latest_folder_path = os.path.join(path_load_finish, latest_folder_cancel)

    # Ищем файл с "Итог_union" в названии
    #file_pattern = os.path.join(latest_folder_path, "*Итог_union*")
    #matching_files = glob.glob(file_pattern)

    # Выбираем первый найденный файл (если есть)
    #if matching_files:
    #    target_file = matching_files[0]
    #    print(f"Найденный файл: {target_file}")
    #    df_last_cancel = pd.read_excel(target_file)
    #else:
    #    print("Файл с 'Итог_union' не найден.")

    # %% [markdown]
    # ## Отмены + проверенные ранее

    # %%
    check_later = os.path.join(path_load_finish, 'Проверенные ранее.xlsx')
    df_check_later = pd.read_excel(check_later)

    # %%
    print(list(df_check_later))

    # %%
    #df_last_cancel = df_last_cancel.rename(columns={"inn":'ИИН',
    #                                                "f258":'Номер исполнительного производства',
    #                                                "Класс 1":"Класс 1",
    #                                                "Дата 1":"Дата 1"
    #                                               })

    # %%
    # 'Проверенные ранее.xlsx' ведётся вручную: в колонке "Дата 1" вперемешку
    # настоящие Excel-даты (пандас читает их как Timestamp) и даты, введённые
    # текстом (например "16.03.2026"). Без format="mixed" pandas пытается
    # угадать один формат на всю колонку и падает на первой же дате другого
    # вида — format="mixed" разбирает каждое значение отдельно.
    df_check_later["Дата 1"] = pd.to_datetime(df_check_later["Дата 1"], format="mixed", dayfirst=True)
    #df_last_cancel["Дата 1"] = pd.to_datetime(df_last_cancel["Дата 1"], format="%d.%m.%Y")

    # %%
    # df_cancel = pd.concat([df_last_cancel[['ИИН', 'Номер исполнительного производства', 'Класс 1']], df_check_later[['ИИН', 'Номер исполнительного производства', 'Класс 1'
    # ]]])

    # df_last_cancel больше не нужен — берём только текущие данные
    df_cancel = df_check_later[['ИИН', 'Номер исполнительного производства', 'Класс 1', 'Дата 1']].copy()

    # %%
    df_cancel = df_cancel.dropna()

    # %%
    df_cancel = df_cancel.rename(columns={"Класс 1":"Статус исполнительного документа", "Дата 1":"Дата Постановления о прекращении ИП"})
    df_cancel  = df_cancel.drop_duplicates(["ИИН", "Номер исполнительного производства"])

    # %% [markdown]
    # ## Создание нового файла.

    # %%
    "Номер_строки_google_t", "ИИН", "№", "Номер исполнительного документа", "Номер исполнительного производства", "Судебный исполнитель", "Дата возбуждения", "Дата выписки исполнительного документа", "Орган выдавший исполнительный документ", "Тип взыскания", "Должник", "ИИН/БИН должника", "Статус исполнительного производства", "Сумма", "Уникальный идентификатор займа"
    "Уникальный номер сделки", "Статус АИС ОИП", "ЧСИ", "Дата возбуждения судебного производства", "Сумма иска", "Номер исполнительного производства", "Номер исполнительного документа", "Статус исполнительного документа", "Дата выписки исполнительного документа", "Орган выдавший исполнительный документ"

    # %%
    "Номер_строки_google_t", "ИИН", "№", "Номер исполнительного документа", "Номер исполнительного производства", "Судебный исполнитель", "Дата возбуждения", "Дата выписки исполнительного документа", "Орган выдавший исполнительный документ", "Тип взыскания", "Должник", "ИИН/БИН должника", "Статус исполнительного производства", "Сумма", "Уникальный идентификатор займа"
    col_target = ["Уникальный номер сделки", "Статус АИС ОИП", "ЧСИ", "Дата возбуждения судебного производства", "Сумма иска", "Номер исполнительного производства", "Номер исполнительного документа", "Статус исполнительного документа", "Дата выписки исполнительного документа", "Орган выдавший исполнительный документ", "Дата Постановления о прекращении ИП"]

    # %%
    final_table_col_rename = final_table.copy()

    # %%
    rename_columns = {
    "Уникальный идентификатор займа":"Уникальный номер сделки",
    "Статус исполнительного производства":"Статус АИС ОИП",
    "Судебный исполнитель":"ЧСИ",
    "Дата возбуждения":"Дата возбуждения судебного производства",
    "Сумма":"Сумма иска",
    "Номер исполнительного производства":"Номер исполнительного производства",
    "Номер исполнительного документа":"Номер исполнительного документа",
    "Класс 1" : "Статус исполнительного документа",
    "Дата выписки исполнительного документа":"Дата выписки исполнительного документа",
    "Орган выдавший исполнительный документ":"Орган выдавший исполнительный документ"
    }

    final_table_col_rename = final_table.rename(columns=rename_columns)
    # final_table_col_rename["Уникальный номер сделки"] = final_table_col_rename["Уникальный номер сделки"] + 100000
    final_table_col_rename["Уникальный номер сделки"] = final_table_col_rename['уникальный номер'] # + 100000

    # %%
    #
    final_table_col_rename["ИИН"] = final_table_col_rename["ИИН"].astype(np.int64)
    df_cancel["ИИН"] = df_cancel["ИИН"].astype(np.int64)

    # %%
    df_cancel  = df_cancel.drop_duplicates(["ИИН", "Номер исполнительного производства"])

    # %%
    print("final_table_col_rename.shape", final_table_col_rename.shape)
    final_table_col_rename = pd.merge(final_table_col_rename, df_cancel, on=["ИИН", "Номер исполнительного производства"], how='left')
    print("final_table_col_rename.shape", final_table_col_rename.shape)

    # %%
    final_table_col_rename = final_table_col_rename[col_target]

    # %%
    final_table_col_rename

    # %%
    final_table_col_rename["Дата Постановления о прекращении ИП"] = pd.to_datetime(final_table_col_rename["Дата Постановления о прекращении ИП"]).dt.date

    # %%
    now_str = datetime.now().strftime("%d.%m.%Y")
    path_format1 = f'{loading_path}/AIS_OIP_Import.xlsx'

    final_table_col_rename.to_excel(path_format1, index=False)



    # %%
    template_path = f'{loading_path_part}/Шаблон1.xlsx'

    # %%
    final_table_col_rename.to_excel(path_format1, index=False)



    import win32com.client as win32

    def apply_format_and_header_from_template(template_path, target_path):
        import win32com.client as win32
        import pythoncom

        pythoncom.CoInitialize()
        excel = None
        template_wb = None
        target_wb = None

        try:
            excel = win32.DispatchEx('Excel.Application')  # ← НОВЫЙ процесс, не цепляется к открытому
            excel.Visible = False
            excel.DisplayAlerts = False

            template_wb = excel.Workbooks.Open(template_path)

            # Проверяем, не открыт ли уже файл в этом же экземпляре
            target_filename = os.path.basename(target_path)
            target_wb = None
            for wb in excel.Workbooks:
                if wb.Name == target_filename:
                    target_wb = wb
                    break
            if target_wb is None:
                target_wb = excel.Workbooks.Open(target_path)

            template_ws = template_wb.Sheets(1)
            target_ws = target_wb.Sheets(1)

            template_ws.Rows(1).Copy()
            target_ws.Rows(1).PasteSpecial(Paste=-4104)

            template_ws.Rows(2).Copy()
            target_last_row = target_ws.UsedRange.Rows.Count
            target_ws.Rows(f"2:{target_last_row}").PasteSpecial(Paste=-4122)

            for col in range(1, template_ws.UsedRange.Columns.Count + 1):
                target_ws.Columns(col).ColumnWidth = template_ws.Columns(col).ColumnWidth

            target_wb.Save()

        finally:
            if template_wb:
                try: template_wb.Close(SaveChanges=False)
                except: pass
            if target_wb:
                try: target_wb.Close(SaveChanges=True)
                except: pass
            if excel:
                try: excel.Quit()
                except: pass
            pythoncom.CoUninitialize()

    # # Укажите пути к файлу шаблона и целевому файлу
    # template_path = r"D:\Analyst_profession\Работа IDCollect\git\Форматирование.xlsx"
    # target_path = r"D:\Analyst_profession\Работа IDCollect\git\rows.xlsx"

    apply_format_and_header_from_template(template_path, path_format1)


    # %% [markdown]
    # # копирование файла в папку crm

    # %%
    path_format1 = Path(path_format1)
    path_crm = Path(path_crm)

    # Формирование полного пути для целевого файла
    path_crm_final = path_crm / path_format1.name


    # %%
    try:
        # Копирование файла, включая метаданные
        shutil.copy2(path_format1, path_crm_final)
        print(f"Файл успешно скопирован в: {path_crm_final}")
    except Exception as e:
        print(f"Ошибка при копировании файла: {e}")

except Exception:
    print("КРИТИЧЕСКАЯ ОШИБКА В СКРИПТЕ СТАТУСОВ")
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
