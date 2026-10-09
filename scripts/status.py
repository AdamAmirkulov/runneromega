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
    parser.add_argument(
        '--skip_crm_import', action='store_true',
        help='Тестовый прогон: не копировать итоговый xlsx в папку автоимпорта Дельты (path_crm). '
             'Все промежуточные файлы (в т.ч. AIS_OIP_Import.xlsx) всё равно пишутся в loading_path.',
    )
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
# --workdir переопределяет loading_path_part: в config.py этот путь обычно
# локальный для машины, на которой обычно запускают скрипт (например,
# 'C:\Users\Администратор\...' на сервере) — на другой машине его нужно
# подменить рабочей директорией явно (тестовые прогоны и т.п.).
loading_path_part = args.workdir or CREDENTIALS['loading_path_part']
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
    import re
    import pyodbc

    from config import CRM_DB, DB_COMPANY_FILTER

    # =========================================================
    # РЕЕСТР ЗАЙМОВ — ИЗ CRM (DBSRV), БЕЗ GOOGLE-ТАБЛИЦЫ
    # =========================================================
    # Раньше "Уникальный номер сделки" был не настоящим ID из CRM, а
    # синтетическим номером строки Google-таблицы компании (df_google_t ->
    # index -> uid_to_string с префиксом eid_prefix). Остальные скрипты
    # (otmeny.py, reestr_gp_import.py, sud_fio_zayavlenie.py и т.д.) уже
    # используют РЕАЛЬНЫЙ l.EID из Loans как "Уникальный номер сделки" —
    # здесь делаем так же, чтобы номер совпадал с остальной системой
    # (Дельта-импорт ждёт именно l.EID).
    #
    # Строка АИС ОИП (Должник + ИИН/БИН + номер производства/документа +
    # сумма) сопоставляется с займом в Loans каскадом. Привязка всегда
    # идёт по ИИН/БИН должника (c.F293), а затем — по убыванию надёжности:
    #   1) l.F146 — номер исполнительного производства;
    #   2) l.F149 — номер исполнительного документа (если по производству
    #      не нашлось — например ещё не занесли номер в CRM после
    #      перевозбуждения);
    #   3) l.F34  — сумма (остаток задолженности), если не помог ни номер
    #      производства, ни номер документа.
    # Так решается ситуация, когда у одного должника два и более займов
    # на исполнении: каждый займ отличается своим F146/F149/суммой. Этап
    # привязывает EID, только если по ключу в Loans находится РОВНО один
    # заём — иначе (например, два займа с одинаковой суммой у одного
    # должника) привязка неоднозначна и остаётся на ручную проверку,
    # вместо того чтобы угадывать.

    def get_crm_connection():
        conn_str = (
            f"DRIVER={{ODBC Driver 18 for SQL Server}};"
            f"SERVER={CRM_DB['server']};"
            f"DATABASE={CRM_DB['database']};"
            f"UID={CRM_DB['username']};"
            f"PWD={CRM_DB['password']};"
            f"TrustServerCertificate=yes;"
        )
        return pyodbc.connect(conn_str, timeout=30)

    LOANS_SQL = f"""
    SELECT
        l.EID         AS eid,
        c.F293        AS iin,
        l.F146        AS ip_number,
        l.F149        AS doc_number,
        CAST(l.F34 AS DECIMAL(18,2)) AS amount,
        s.Caption     AS loan_status,
        constF235.Caption AS doc_status
    FROM loans l
    JOIN clients c ON c.ID = l.CID
    LEFT JOIN states s ON s.ID = l.State
    LEFT JOIN Constants constF235 ON constF235.ID = l.F235
    WHERE l.F209 = N'{DB_COMPANY_FILTER}'
    """

    print("CRM: загрузка реестра займов (Loans) из DBSRV…")
    _crm_conn = get_crm_connection()
    try:
        df_loans = pd.read_sql(LOANS_SQL, _crm_conn)
    finally:
        _crm_conn.close()
    print(f"CRM: загружено займов: {len(df_loans)}")

    df_loans["eid"] = pd.to_numeric(df_loans["eid"], errors="coerce").astype("Int64")

    def _norm_iin(value):
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return ""
        digits = re.sub(r"\D", "", str(value))
        return digits.zfill(12) if digits else ""

    def _norm_key(value):
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return ""
        return re.sub(r"\s+", "", str(value)).strip().casefold()

    df_loans["iin_norm"] = df_loans["iin"].apply(_norm_iin)
    df_loans["ip_norm"] = df_loans["ip_number"].apply(_norm_key)
    df_loans["doc_norm"] = df_loans["doc_number"].apply(_norm_key)
    df_loans["amount_round"] = pd.to_numeric(df_loans["amount"], errors="coerce").round(0)

    def _unique_lookup(df, key_cols):
        """{key_cols -> eid}, только для комбинаций, которым в Loans
        соответствует ровно один заём — иначе привязка неоднозначна."""
        sub = df[key_cols + ["eid"]].dropna(subset=["eid"])
        for c in key_cols:
            sub = sub[sub[c] != ""]
        if sub.empty:
            return sub
        nunique = sub.groupby(key_cols)["eid"].transform("nunique")
        return sub[nunique == 1].drop_duplicates(key_cols)

    def _apply_stage(df_target, key_cols, stage_label):
        unmatched = df_target["eid"].isna()
        if not unmatched.any():
            return 0
        lookup = _unique_lookup(df_loans, key_cols).rename(columns={"eid": "_matched_eid"})
        merged = df_target.loc[unmatched, key_cols].merge(lookup, on=key_cols, how="left")
        merged.index = df_target.index[unmatched]
        found = merged["_matched_eid"].notna()
        idx = merged.index[found]
        df_target.loc[idx, "eid"] = merged.loc[found, "_matched_eid"].astype("Int64").values
        df_target.loc[idx, "match_stage"] = stage_label
        return int(found.sum())

    df_ais_match = df_final.copy()
    df_ais_match['ИИН/БИН должника'] = df_ais_match['ИИН/БИН должника'].apply(
        lambda x: '0' * (12 - len(str(x))) + str(x)
    )
    df_ais_match["iin_norm"] = df_ais_match["ИИН/БИН должника"].apply(_norm_iin)
    df_ais_match["ip_norm"] = df_ais_match["Номер исполнительного производства"].apply(_norm_key)
    df_ais_match["doc_norm"] = df_ais_match["Номер исполнительного документа"].apply(_norm_key)
    df_ais_match["amount_round"] = pd.to_numeric(df_ais_match["Сумма"], errors="coerce").round(0)
    df_ais_match["eid"] = pd.Series(pd.NA, index=df_ais_match.index, dtype="Int64")
    df_ais_match["match_stage"] = ""

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

    df_ais_oip_raw_data = df_ais_match[ais_col]

    n1 = _apply_stage(df_ais_match, ["iin_norm", "ip_norm"], "F146 (номер ИП)")
    print(f"Этап 1 — ИИН + номер исполнительного производства (F146): сопоставлено {n1}")

    n2 = _apply_stage(df_ais_match, ["iin_norm", "doc_norm"], "F149 (номер исп. документа)")
    print(f"Этап 2 — ИИН + номер исполнительного документа (F149): сопоставлено {n2}")

    n3 = _apply_stage(df_ais_match, ["iin_norm", "amount_round"], "F34 (сумма)")
    print(f"Этап 3 — ИИН + сумма (F34): сопоставлено {n3}")

    cnt_not_matched = int(df_ais_match["eid"].isna().sum())
    print(f"Не сопоставлено ни на одном этапе: {cnt_not_matched}")

    dup_counts = df_ais_match.dropna(subset=["eid"]).groupby("eid").size()
    cnt_duble = int((dup_counts > 1).sum())
    print(f"Займов, которым соответствует больше одной строки АИС ОИП: {cnt_duble}")

    df_ais_match["priority_id_status"] = np.where(
        df_ais_match["Статус исполнительного производства"] == "На исполнении", 1, 2
    )
    df_ais_matched = (
        df_ais_match.dropna(subset=["eid"])
        .sort_values(
            ["eid", "priority_id_status", "Дата возбуждения", "Дата выписки исполнительного документа"],
            ascending=[True, True, False, False],
        )
        .drop_duplicates("eid", keep="first")
    )

    # Справочно (НЕ влияет на итоговый файл): сколько незавершённых займов
    # CRM не встретились ни в одной строке этой выгрузки АИС ОИП вообще —
    # обычно это займы, ещё не поданные на исполнение (см. память), а не
    # брак сопоставления.
    FINISHED_LOAN_STATUSES = {"Погашен", "Обратный выкуп"}
    is_finish = df_loans["loan_status"].isin(FINISHED_LOAN_STATUSES)
    cnt_not_pull = int((~df_loans.loc[~is_finish, "eid"].isin(df_ais_matched["eid"])).sum())
    print(f"Справочно: незавершённых займов без строки в этой выгрузке АИС ОИП: {cnt_not_pull}")

    ais_cols_to_attach = ["eid"] + ais_col

    # Источник строк — АИС ОИП (производства), а не Loans: производство
    # существует только в АИС ОИП, поэтому итоговая таблица строится от
    # df_ais_matched (уже по одной строке на заём, самая приоритетная —
    # "На исполнении" + самая свежая), а к ней слева подтягивается ИИН и
    # статус исполнительного документа из CRM по EID. Займы CRM, которых
    # вообще нет в этой выгрузке АИС ОИП, в файл не попадают — по ним
    # нечего обновлять в Дельте (они и раньше сюда не попадали, просто
    # раньше это маскировалось строками с пустыми полями).
    final_table = df_ais_matched[ais_cols_to_attach].merge(
        df_loans[["eid", "iin", "doc_status"]].rename(
            columns={"iin": "ИИН", "doc_status": "Статус исполнительного документа"}
        ),
        on="eid", how="left",
    )

    now_str = datetime.now().strftime("%d.%m.%Y")
    final_table.to_excel(f'{loading_path}/{name_login} AIS_OIP_merge_crm {now_str}.xlsx', index=False)

    df_ais_not_processed = df_ais_oip_raw_data.loc[(~df_ais_oip_raw_data['Номер исполнительного производства'].isin(final_table["Номер исполнительного производства"])) & (df_ais_oip_raw_data['Статус исполнительного производства'] == 'На исполнении')]
    df_ais_not_processed.to_excel(f'{loading_path}/{name_login} Производства не подтянулись .xlsx', index=False)

    # =========================================================
    # "Проверенные ранее.xlsx" убран из пайплайна.
    # =========================================================
    # Раньше "Статус исполнительного документа" (и дата постановления о
    # прекращении ИП) брались из вручную ведущегося файла load_finish/
    # "Проверенные ранее.xlsx" — это была замена того, чего CRM не отдавала
    # напрямую. Теперь статус исполнительного документа тянется прямо из
    # DBSRV (l.F235 -> Constants.Caption, тот же constF235, что используют
    # reestr_gosposhliny.py/reestr_gp_import.py/sud.py) и уже сидит в
    # final_table. Отдельного поля "Дата Постановления о прекращении ИП" в
    # Loans не нашлось — колонка убрана из col_target; если найдётся нужное
    # поле в CRM, добавить туда же, где doc_status.

    # %% [markdown]
    # ## Создание нового файла.

    # %%
    col_target = ["Уникальный номер сделки", "Статус АИС ОИП", "ЧСИ", "Дата возбуждения судебного производства", "Сумма иска", "Номер исполнительного производства", "Номер исполнительного документа", "Статус исполнительного документа", "Дата выписки исполнительного документа", "Орган выдавший исполнительный документ"]

    # %%
    rename_columns = {
    "eid":"Уникальный номер сделки",
    "Статус исполнительного производства":"Статус АИС ОИП",
    "Судебный исполнитель":"ЧСИ",
    "Дата возбуждения":"Дата возбуждения судебного производства",
    "Сумма":"Сумма иска",
    "Номер исполнительного производства":"Номер исполнительного производства",
    "Номер исполнительного документа":"Номер исполнительного документа",
    "Дата выписки исполнительного документа":"Дата выписки исполнительного документа",
    "Орган выдавший исполнительный документ":"Орган выдавший исполнительный документ"
    }

    final_table_col_rename = final_table.rename(columns=rename_columns)
    final_table_col_rename["ИИН"] = final_table_col_rename["ИИН"].astype(np.int64)
    final_table_col_rename = final_table_col_rename[col_target]

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
    if args.skip_crm_import:
        print(f"--skip_crm_import: копирование в {path_crm_final} ПРОПУЩЕНО (тестовый прогон).")
    else:
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
