# ========= АРГУМЕНТЫ КОМАНДНОЙ СТРОКИ =========
import argparse
import os
import sys

def parse_args():
    parser = argparse.ArgumentParser(description="Скрипт подачи исков в office.sud.kz")
    parser.add_argument('--workdir',    type=str, default=None)
    parser.add_argument('--excel_path', type=str, default=None)
    parser.add_argument('--batch_dir',  type=str, default=None)
    parser.add_argument('--company_id', type=str, default=None)
    return parser.parse_args()

args = parse_args()

if not args.company_id or not args.company_id.strip():
    print("❌ ОШИБКА: не передан --company_id — компания не определена, запуск остановлен.")
    sys.exit(1)

# ДО импорта config — чтобы он увидел правильный COMPANY_ID
os.environ['COMPANY_ID'] = args.company_id.strip()

if args.workdir and args.workdir.strip():
    os.chdir(args.workdir.strip())

if args.excel_path and args.excel_path.strip():
    os.environ['OVERRIDE_EXCEL_PATH'] = args.excel_path.strip()

if args.batch_dir and args.batch_dir.strip():
    os.environ['OVERRIDE_BATCH_DIR'] = args.batch_dir.strip()
# ========= ОСТАЛЬНЫЕ ИМПОРТЫ =========

import copy
import json
import logging
import mimetypes
import re
import time
import socket
import subprocess
import xml.etree.ElementTree as ET
from datetime import datetime

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import parse_qsl, urlencode, urljoin, urlsplit, urlunsplit

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

from config import MAIN_EXCEL, CREDENTIALS, ROOT
from podacha_recipe import RECIPE_STEPS

BASE_URL = "https://office.sud.kz"

# Порт и профиль Chrome ОБЯЗАНЫ зависеть от company_id. Раньше оба были
# захардкожены на одно и то же значение для всех компаний: если два
# запуска (например, А-Омега и Orion) стартовали примерно в одно время,
# второй процесс видел, что порт 9222 уже занят первым, и просто
# подключался к ЕГО уже открытому и залогиненному браузеру — в результате
# иски второй компании подавались через аккаунт первой в Судебном
# кабинете. Порт и профиль ниже уникальны на компанию, поэтому
# параллельные запуски разных компаний больше не могут столкнуться.
_company_id_for_chrome = args.company_id.strip() or "1"
if _company_id_for_chrome.isdigit():
    _chrome_port_offset = int(_company_id_for_chrome)
else:
    import zlib as _zlib
    _chrome_port_offset = _zlib.crc32(_company_id_for_chrome.encode("utf-8")) % 1000
CHROME_DEBUG_PORT = 9222 + _chrome_port_offset
CHROME_DEBUG_ADDRESS = f"127.0.0.1:{CHROME_DEBUG_PORT}"

# ============================================================
# АВТОМАТИЧЕСКИЙ ВХОД
# ============================================================

AUTO_LOGIN = True

# Логин и пароль теперь берутся из config.py (по COMPANY_ID),
# как и в первом скрипте, а не хардкодятся здесь.
SUD_LOGIN    = CREDENTIALS['sk_login']
SUD_PASSWORD = CREDENTIALS['sk_password']

# Путь к Chrome. Скрипт сам запустит его в режиме remote debugging,
# если порт 9222 ещё не открыт.
CHROME_PATH = Path(
    r"C:\Program Files\Google\Chrome\Application\chrome.exe"
)

# Отдельный профиль Chrome для Судебного кабинета — свой на каждую
# компанию (см. комментарий у CHROME_DEBUG_PORT выше), чтобы разные
# компании никогда не делили один и тот же залогиненный браузер.
CHROME_USER_DATA_DIR = Path(
    rf"C:\Users\User\АИС ОИП1\chrome_sud_profile_company{_company_id_for_chrome}"
)

LOGIN_TIMEOUT = 90

# Эти файлы — статичные образцы, лежащие рядом со скриптом (в scripts/),
# а не в рабочей папке конкретного запуска (--workdir делает os.chdir выше,
# поэтому Path.cwd() здесь указывал бы на data/runs/<job_id> и не находил их).
SCRIPT_DIR = Path(__file__).resolve().parent

# Последовательность бизнес-запросов подачи иска заморожена в
# podacha_recipe.py (RECIPE_STEPS). Раньше она разбиралась на лету из
# бинарного HAR office.sud.kz_CURRENT_PARTICIPANTS.har плюс резервных
# JSON в SUD_API_CAPTURES/ — теперь эти файлы во время работы не нужны.
# Как регенерировать рецепт при смене JSF-id портала — см. podacha_recipe.py.

# Excel и документы — тоже через config, с возможностью переопределить
# через --excel_path.
EXCEL_FILE = Path(os.environ.get('OVERRIDE_EXCEL_PATH', MAIN_EXCEL))
EXCEL_SHEET = "Отмены"
# Датированные папки партий (каждая — результат отдельного запуска сбора
# документов) лежат не прямо в ROOT, а в ROOT\Документы для подачи Исков\,
# как и TARGET_BASE в config.py. Раньше CASES_BASE_DIR указывал на сам ROOT,
# из-за чего auto-выбор партии находил эту общую папку вместо конкретной
# датированной партии внутри неё.
CASES_BASE_DIR = Path(ROOT) / "Документы для подачи Исков"

# Папка конкретной партии, из которой берутся должники и документы.
# Берётся из --batch_dir, если передан; иначе None — selected_batch_folder()
# сама выберет самую свежую папку внутри CASES_BASE_DIR (ROOT текущей компании).
# Раньше здесь был захардкожен конкретный путь одной компании и конкретная
# дата, из-за чего для других компаний/после устаревания даты папка не находилась.
_override_batch_dir = os.environ.get('OVERRIDE_BATCH_DIR', '').strip()
CASES_BATCH_DIR = Path(_override_batch_dir) if _override_batch_dir else None

START_ROW = 2

# None = автоматически обработать все заполненные строки Excel.
END_ROW = None

# При ошибке по одному должнику перейти к следующему, а не останавливать пакет.
CONTINUE_ON_ROW_ERROR = True

# Ограничение портала: каждый файл не более 20 МБ.
MAX_UPLOAD_FILE_BYTES = 20 * 1024 * 1024

# Повтор временно оборвавшейся загрузки.
UPLOAD_RETRY_COUNT = 3
UPLOAD_RETRY_BASE_DELAY = 4

# Колонки Excel.
COL_FIO = "C"
COL_IIN = "D"
COL_SUM = "J"
COL_DUTY = "K"
COL_REGION = "P"
COL_COURT = "Q"

# Реквизиты организации/представителя — тоже из config по COMPANY_ID.
ORG_BIN  = CREDENTIALS['org_bin']
ORG_BANK = CREDENTIALS['org_bank']
REP_IIN  = CREDENTIALS['rep_iin']

# Плейсхолдеры, записанные в HAR вместо реальных персональных данных
# исходной записи (после очистки HAR от ПДн). Используются как маркеры
# для замены и для определения роли поля (ответчик/представитель)
# в динамической форме.
CAPTURED_ORG_BIN       = "100000000000"
CAPTURED_ORG_BANK      = "KZ00PLACEHOLDERBANK0"
CAPTURED_DEFENDANT_IIN = "200000000000"
CAPTURED_REP_IIN       = "300000000000"

HTTP_TIMEOUT = 120
PAUSE_BETWEEN_REQUESTS = 0.15
STOP_AT_SIGN_PAGE = True
VERIFY_TLS = True

# --- Устойчивость к недоступности портала office.sud.kz -------------------
# Портал регулярно отдаёт 502/504/404 или рвёт соединение во время
# деплоя/техработ. Такие ошибки не относятся к конкретному клиенту —
# запрос повторяется с нарастающей паузой, а если портал лежит устойчиво,
# весь пакет останавливается (см. главный цикл), чтобы не пометить сотни
# строк как «не поданы» и просто перезапустить позже.
HTTP_RETRY_ATTEMPTS = 4
HTTP_RETRY_BASE_DELAY = 5           # сек; удваивается на каждой попытке
HTTP_RETRY_STATUS = {404, 500, 502, 503, 504}
# Сколько клиентов подряд должны упасть на ошибке портала, чтобы
# остановить весь пакет.
PORTAL_ERROR_ABORT_THRESHOLD = 5
# Пауза между клиентами после единичной ошибки портала.
PORTAL_ERROR_COOLDOWN = 15

# Подробный технический лог (ViewState, component_id, каждый HAR-шаг и т.д.)
# пишется только в файл — он нужен для разбора проблем разработчиком.
# В консоль/веб-интерфейс идёт только короткий, понятный прогресс через
# print() ниже (см. функцию progress()) — его видят обычные сотрудники.
LOG_FILE = Path("sud_http_api.log")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[logging.FileHandler(LOG_FILE, encoding="utf-8")],
)
log = logging.getLogger("sud-http")


def progress(msg: str) -> None:
    """Короткое сообщение о ходе работы — всегда видно в консоли/веб-логе,
    в отличие от log.info/log.warning, которые теперь пишутся только в файл."""
    print(msg, flush=True)


def short_error(exc: Exception, limit: int = 200) -> str:
    """Сжимает сообщение об ошибке до одной читаемой строки для консоли
    (некоторые исключения несут внутри кусок сырого HTML/JSON ответа сайта —
    это остаётся в файле sud_http_api.log, а в консоль идёт короткая версия)."""
    text = re.sub(r"\s+", " ", str(exc)).strip()
    if len(text) > limit:
        text = text[:limit].rstrip() + "…"
    return text


# %% [markdown]
# ## 1. Получение cookies из открытого Chrome
# 
# Это единственный блок, где используется Selenium.

# %%

@dataclass
class BrowserAuth:
    cookies: list[dict[str, Any]]
    user_agent: str
    current_url: str


class PortalUnavailable(RuntimeError):
    """
    Портал office.sud.kz временно недоступен: 5xx, 404 на рабочем
    эндпоинте, обрыв соединения, SSL EOF или таймаут — после всех ретраев.

    Это НЕ проблема конкретного клиента: пакет нужно повторить позже.
    Главный цикл считает такие ошибки подряд и останавливает весь прогон,
    не трогая оставшиеся строки Excel.
    """


class SessionExpired(RuntimeError):
    """
    HTTP-сессия перестала быть авторизованной: портал отдаёт страницу
    входа вместо формы. Cookies, снятые из Chrome, больше не действуют.
    Лечится повторным входом (не повтором того же запроса), поэтому это
    НЕ PortalUnavailable.
    """


def _response_is_login_page(response: requests.Response) -> bool:
    """Портал не авторизовал запрос и вернул страницу входа /index.xhtml.

    Раньше проверялись только русские/английские маркеры ('войти в
    судебный кабинет', name="login"), а портал по умолчанию отдаёт
    страницу входа на казахском — из-за чего скрипт молча продолжал
    парсить логин-страницу и падал на 'Не найден селект Тип производства'.
    """
    try:
        body = response.text or ""
    except Exception:
        return False
    low = body[:40000].casefold()
    low_url = (response.url or "").casefold()

    # Поля формы входа портала: <input name="...:auth:xin"> / ":auth:password">
    if ":auth:xin" in low or ":auth:password" in low:
        return True
    if 'name="login"' in low or "войти в судебный кабинет" in low:
        return True
    # Казахская страница входа.
    if "құпия сөз" in low and ("сот кабинеті" in low or "эцқ" in low):
        return True
    # GET рабочей формы, но сервер увёл на /index.xhtml (корень = вход).
    if low_url.rstrip("/").endswith("/index.xhtml") and (
        "input[type='password']" in low or 'type="password"' in low
    ):
        return True
    return False


def _is_portal_outage_text(text: str) -> bool:
    """Распознаёт в тексте исключения признаки именно СЕТЕВОЙ недоступности
    портала (5xx, обрыв соединения, таймаут) — то, что имеет смысл
    повторить позже. Проблемы авторизации/сессии сюда НЕ входят."""
    low = (text or "").casefold()
    markers = (
        "портал недоступен", "портал отвечает",
        "bad gateway", "gateway time-out", "gateway timeout",
        "500 server error", "502 server error",
        "503 server error", "504 server error",
        "read timed out", "connection aborted", "connection reset",
        "max retries exceeded", "unexpected_eof", "eof occurred",
    )
    return any(m in low for m in markers)


def _verify_http_session(auth: BrowserAuth) -> bool:
    """Проверяет, что снятые из Chrome cookies дают авторизованную сессию:
    делает реальный GET рабочей формы и смотрит, не вернулась ли страница
    входа. Именно здесь ловится случай «вкладка Chrome показывает кабинет
    из кэша, а серверная сессия мертва»."""
    probe = make_http_session(auth)
    try:
        resp = probe.get(
            BASE_URL + "/form/send/index.xhtml",
            timeout=HTTP_TIMEOUT,
            verify=VERIFY_TLS,
            allow_redirects=True,
        )
    except Exception as exc:  # сеть/портал — не про авторизацию
        log.warning("Проверка HTTP-сессии не удалась (сеть): %s", exc)
        return False

    if _response_is_login_page(resp):
        return False
    body_low = (resp.text or "").casefold()
    url_low = (resp.url or "").casefold()
    # Положительный признак: селект «Тип производства» на форме, либо мы
    # реально остались на /form/send/... (сервер не увёл на страницу входа).
    return "case-type" in body_low or "/form/send" in url_low


def _debug_port_is_open(
    host: str = "127.0.0.1",
    port: int = CHROME_DEBUG_PORT,
) -> bool:
    try:
        with socket.create_connection((host, port), timeout=1):
            return True
    except OSError:
        return False


def _start_debug_chrome() -> None:
    """
    Запускает отдельный Chrome-профиль для Судебного кабинета —
    на СВОЁМ порту и профиле для текущей компании (CHROME_DEBUG_PORT /
    CHROME_USER_DATA_DIR). Если Chrome уже запущен на этом порту, ничего
    не делает.
    """
    if _debug_port_is_open():
        log.info("Chrome remote debugging уже запущен: %s", CHROME_DEBUG_ADDRESS)
        return

    if not CHROME_PATH.exists():
        raise FileNotFoundError(
            f"Chrome не найден: {CHROME_PATH}. "
            "Исправьте CHROME_PATH в настройках."
        )

    CHROME_USER_DATA_DIR.mkdir(parents=True, exist_ok=True)

    command = [
        str(CHROME_PATH),
        f"--remote-debugging-port={CHROME_DEBUG_PORT}",
        f"--user-data-dir={CHROME_USER_DATA_DIR}",
        "--start-maximized",
        "--disable-notifications",
        BASE_URL + "/form/proceedings/services.xhtml",
    ]

    subprocess.Popen(
        command,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

    deadline = time.time() + 30
    while time.time() < deadline:
        if _debug_port_is_open():
            log.info("Chrome запущен в режиме remote debugging.")
            return
        time.sleep(0.5)

    raise TimeoutError(
        f"Chrome запущен, но порт remote debugging {CHROME_DEBUG_PORT} не открылся."
    )


def _first_visible_element(driver, selectors: list[tuple[str, str]]):
    from selenium.webdriver.common.by import By

    by_map = {
        "css": By.CSS_SELECTOR,
        "id": By.ID,
        "name": By.NAME,
        "xpath": By.XPATH,
    }

    for selector_type, selector in selectors:
        try:
            elements = driver.find_elements(by_map[selector_type], selector)
        except Exception:
            continue

        for element in elements:
            try:
                if element.is_displayed() and element.is_enabled():
                    return element
            except Exception:
                continue

    return None


def _looks_authenticated(driver) -> bool:
    """
    Проверяет реальную авторизацию по ВИДИМОМУ DOM, а не по сырому HTML.

    Важно: страница входа тоже имеет URL /form/proceedings/services.xhtml,
    поэтому по одному URL авторизацию определять нельзя.

    Раньше проверка искала подстроки ("войти", "авторизация",
    'type="password"') в driver.page_source. На залогиненном портале эти
    строки почти всегда присутствуют в скрытой разметке/скриптах/меню,
    из-за чего функция вечно возвращала False и ожидание входа падало по
    таймауту даже после успешного входа (особенно на чистом профиле, где
    вход выполняется каждый запуск). Теперь смотрим только на видимые
    элементы и document.body.innerText.
    """
    # 1. Видимое поле пароля => это точно форма входа.
    try:
        for field in driver.find_elements(
            "css selector", "input[type='password']"
        ):
            try:
                if field.is_displayed():
                    return False
            except Exception:
                continue
    except Exception:
        pass

    # 2. Явный признак кабинета — видимая ссылка/кнопка выхода.
    try:
        has_logout = driver.execute_script(
            """
            return [...document.querySelectorAll('a,button')].some(el => {
                if (el.offsetParent === null) return false;
                const t = (el.textContent || '').trim().toLowerCase();
                const h = (el.getAttribute('href') || '').toLowerCase();
                const oc = (el.getAttribute('onclick') || '').toLowerCase();
                return t === 'выйти' || t === 'шығу'
                    || h.includes('logout') || oc.includes('logout');
            });
            """
        )
        if has_logout:
            return True
    except Exception:
        pass

    # 3. Видимый текст страницы (без скрытой разметки).
    try:
        visible = str(driver.execute_script(
            "return document.body ? document.body.innerText : '';"
        ) or "").casefold()
    except Exception:
        visible = ""

    login_markers = (
        "жсн/бсн",
        "құпия сөз",
        "войти",
        "кіру",
        "авторизация",
    )
    if any(marker in visible for marker in login_markers):
        return False

    auth_markers = (
        "шығу",
        "выйти",
        "менің істерім",
        "мои дела",
        "мои заявления",
        "құжаттар беру",
        "подача документов",
    )
    return any(marker in visible for marker in auth_markers)



def _set_controlled_input_value(
    driver,
    element,
    value: str,
) -> None:
    """
    Надёжно заполняет Vue/React/Angular-поле.

    Обычный send_keys может визуально написать текст, но модель страницы
    не получает новое значение. Native setter + input/change/blur
    обновляет и DOM, и JS-состояние формы.
    """
    driver.execute_script(
        """
        const el = arguments[0];
        const value = arguments[1];

        const prototype =
            el.tagName === 'TEXTAREA'
                ? window.HTMLTextAreaElement.prototype
                : window.HTMLInputElement.prototype;

        const descriptor = Object.getOwnPropertyDescriptor(
            prototype,
            'value'
        );

        descriptor.set.call(el, value);

        el.dispatchEvent(new Event('input', {
            bubbles: true,
            composed: true
        }));
        el.dispatchEvent(new Event('change', {
            bubbles: true,
            composed: true
        }));
        el.dispatchEvent(new Event('blur', {
            bubbles: true,
            composed: true
        }));
        """,
        element,
        value,
    )


# ============================================================
# ВХОД: устойчивый поиск формы (портал грузится на казахском,
# JSF-id динамические, placeholder'ы меняются) + диагностика
# ============================================================

def _visible_text(driver) -> str:
    try:
        return str(driver.execute_script(
            "return document.body ? document.body.innerText : '';"
        ) or "")
    except Exception:
        return ""


def _dump_login_debug(driver, tag: str) -> None:
    """Скриншот + HTML страницы входа — чтобы видеть, на чём застряли."""
    try:
        from datetime import datetime as _dt
        d = Path(r"C:\Users\User\АИС ОИП1\login_debug")
        d.mkdir(parents=True, exist_ok=True)
        stamp = _dt.now().strftime("%Y%m%d_%H%M%S")
        try:
            driver.save_screenshot(str(d / f"{tag}_{stamp}.png"))
        except Exception:
            pass
        try:
            (d / f"{tag}_{stamp}.html").write_text(
                driver.page_source, encoding="utf-8"
            )
        except Exception:
            pass
        log.error("Диагностика входа сохранена в %s (%s_%s.*)", d, tag, stamp)
    except Exception:
        pass


def _open_login_form(driver) -> None:
    """Открывает страницу входа, ждёт ЛИБО авторизации, ЛИБО появления поля
    пароля. Язык переключает на РУС по возможности (не критично)."""
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.common.by import By

    driver.get(BASE_URL + "/form/proceedings/services.xhtml")
    try:
        WebDriverWait(driver, 30).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
    except Exception:
        pass

    # РУС — best effort (портал по умолчанию на казахском).
    ru = _first_visible_element(driver, [
        ("css", "a[onclick*='selectLanguageRu']"),
        ("xpath", "//a[normalize-space()='РУС' or normalize-space()='Рус']"),
    ])
    if ru is not None:
        try:
            driver.execute_script("arguments[0].click();", ru)
            time.sleep(1.0)
        except Exception:
            pass

    # Ждём до 60 сек: либо уже вошли, либо появилось поле пароля.
    # (холодный профиль + цепочка редиректов IDP бывает медленной)
    deadline = time.time() + 60
    while time.time() < deadline:
        if _looks_authenticated(driver):
            return
        if driver.find_elements(By.CSS_SELECTOR, "input[type='password']"):
            return
        time.sleep(0.5)
    # ничего не дождались — пусть вызывающий сохранит диагностику
    raise TimeoutError("форма входа не появилась (нет поля пароля и не авторизованы)")


def _find_login_fields(driver):
    """Возвращает (login_input, password_input). Если уже авторизованы и формы
    нет — (None, None). Поля ищутся без опоры на конкретные JSF-id/placeholder."""
    from selenium.webdriver.common.by import By

    if _looks_authenticated(driver) and not driver.find_elements(
        By.CSS_SELECTOR, "input[type='password']"
    ):
        return None, None

    password_element = _first_visible_element(driver, [
        ("css", "input[id*=':auth:password']"),
        ("css", "input[name*=':auth:password']"),
        ("css", "input[type='password']"),
    ])
    if password_element is None:
        raise RuntimeError("поле пароля на странице входа не найдено")

    login_element = _first_visible_element(driver, [
        ("css", "input[id*=':auth:xin']"),
        ("css", "input[name*=':auth:xin']"),
        ("css", "input[placeholder*='ИИН']"),
        ("css", "input[placeholder*='БИН']"),
        ("css", "input[placeholder*='ЖСН']"),
        ("css", "input[placeholder*='БСН']"),
    ])
    if login_element is None:
        # запасной вариант: видимое текстовое поле прямо перед полем пароля
        login_element = driver.execute_script(
            """
            const pw = arguments[0];
            const inputs = [...document.querySelectorAll(
                "input[type='text'],input[type='tel'],input[type='email'],input:not([type])"
            )].filter(el => el.offsetParent !== null);
            let prev = null;
            for (const el of inputs) {
                if (el.compareDocumentPosition(pw) &
                    Node.DOCUMENT_POSITION_FOLLOWING) prev = el;
            }
            return prev;
            """,
            password_element,
        )
    if login_element is None:
        raise RuntimeError("поле логина (ИИН/БИН) на странице входа не найдено")

    return login_element, password_element


def _perform_login(driver, attempts: int = 3, force: bool = False) -> None:
    """
    Вход в Судебный кабинет с несколькими попытками.

    На чистом профиле Chrome (новый профиль на каждый запуск) полный вход
    выполняется каждый раз, поэтому единичный сбой формы/редиректа/AJAX
    портала не должен ронять весь пакет — повторяем открытие формы и ввод.

    force=True — не доверять состоянию вкладки Chrome (страница кабинета
    могла остаться из bfcache/кэша, а серверная сессия уже мертва) и
    выполнить полноценный вход в любом случае.
    """
    from selenium.common.exceptions import WebDriverException

    if not force and _looks_authenticated(driver):
        log.info("Пользователь уже авторизован — вход не требуется.")
        return

    last_exc: Exception | None = None
    for attempt in range(1, attempts + 1):
        try:
            log.info("Попытка входа %s/%s.", attempt, attempts)
            _perform_login_once(driver)
            log.info("Вход выполнен с попытки %s.", attempt)
            return
        except (RuntimeError, WebDriverException) as exc:
            last_exc = exc
            if _looks_authenticated(driver):
                log.info(
                    "После сбоя %r кабинет всё же открыт — считаем вход "
                    "выполненным.", exc,
                )
                return
            log.warning("Попытка входа %s не удалась: %s", attempt, exc)
            if attempt < attempts:
                time.sleep(3 * attempt)

    raise RuntimeError(
        f"Автоматический вход не выполнен за {attempts} попыток. "
        f"Последняя ошибка: {last_exc}"
    ) from last_exc


def _perform_login_once(driver) -> None:
    """
    Одна попытка входа: открывает форму, вводит ИИН/БИН и пароль
    посимвольно, жмёт «Войти». Поиск полей устойчив к смене языка и
    динамических JSF-id. При провале сохраняется скриншот+HTML.
    """
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.action_chains import ActionChains
    from selenium.common.exceptions import StaleElementReferenceException

    try:
        _open_login_form(driver)
        login_element, password_element = _find_login_fields(driver)
        if login_element is None:  # уже авторизованы — форма не появилась
            log.info("Пользователь уже авторизован в Chrome-профиле.")
            return
    except Exception as exc:
        _dump_login_debug(driver, "login_form_not_found")
        page = _visible_text(driver)[:400]
        raise RuntimeError(
            "Не удалось открыть форму входа в Судебный кабинет. "
            f"URL={getattr(driver, 'current_url', '?')}. "
            f"Видимый текст страницы: {page!r}"
        ) from exc

    if not SUD_LOGIN or not SUD_PASSWORD:
        raise RuntimeError("Не заполнены SUD_LOGIN и SUD_PASSWORD.")

    # --------------------------------------------------------
    # 3-4. Ввод логина и пароля (посимвольно; если не «прилипло» —
    #      подстраховка через native setter). Перепроверяем на свежем
    #      элементе — форма может перерисоваться при вводе.
    # --------------------------------------------------------
    def _fill_field(finders, value, label):
        got = ""
        for _ in range(3):
            el = WebDriverWait(driver, 20).until(
                lambda d: _first_visible_element(d, finders)
            )
            try:
                el.click()
            except Exception:
                pass
            try:
                driver.execute_script(
                    "arguments[0].value='';"
                    "arguments[0].dispatchEvent(new Event('input',{bubbles:true}));",
                    el,
                )
            except Exception:
                pass
            time.sleep(0.2)
            try:
                for ch in value:
                    el.send_keys(ch)
                    time.sleep(0.08)
            except Exception:
                pass
            time.sleep(0.3)
            got = str(driver.execute_script("return arguments[0].value;", el) or "")
            if got != value:
                # посимвольно не прилипло — ставим значение напрямую
                _set_controlled_input_value(driver, el, value)
                time.sleep(0.3)
                got = str(driver.execute_script("return arguments[0].value;", el) or "")
            if got == value:
                return el
            time.sleep(0.5)
        raise RuntimeError(
            f"{label} не удалось ввести в поле (сейчас там {got!r}, нужно {value!r})"
        )

    log.info("Ввожу логин.")
    login_element = _fill_field(
        [
            ("css", "input[id*=':auth:xin']"),
            ("css", "input[name*=':auth:xin']"),
            ("css", "input[placeholder*='ИИН']"),
            ("css", "input[placeholder*='ЖСН']"),
        ],
        SUD_LOGIN, "Логин",
    )

    log.info("Ввожу пароль.")
    password_element = _fill_field(
        [
            ("css", "input[id*=':auth:password']"),
            ("css", "input[name*=':auth:password']"),
            ("css", "input[type='password']"),
        ],
        SUD_PASSWORD, "Пароль",
    )

    # Добиваем события на обоих полях + снимаем фокус — на случай, если
    # кнопка/валидация RichFaces реагирует на keyup/change/blur.
    try:
        for el in (login_element, password_element):
            driver.execute_script(
                "for (const t of ['keydown','keyup','input','change','blur'])"
                " arguments[0].dispatchEvent(new Event(t,{bubbles:true}));",
                el,
            )
        driver.execute_script("document.activeElement && document.activeElement.blur();")
    except Exception:
        pass
    time.sleep(1.0)

    # --------------------------------------------------------
    # 5. Кнопка Войти
    # --------------------------------------------------------
    def find_login_button():
        el = _first_visible_element(
            driver,
            [
                ("css", "input[id*=':auth:'][type='submit']"),
                ("css", "input[name*=':auth:'][type='submit']"),
                ("css", "input[type='submit'][value='Войти']"),
                ("css", "input[type='submit'][value='Кіру']"),
                ("css", "button[type='submit']"),
            ],
        )
        if el is not None:
            return el
        return driver.execute_script(
            """
            return [...document.querySelectorAll("input[type='submit'],button")]
              .find(b => b.offsetParent !== null &&
                    /войти|кіру|кіріс/i.test((b.value||b.textContent||''))) || null;
            """
        )

    # Финальная проверка значения логина на том же элементе, что заполняли.
    try:
        actual_login = str(driver.execute_script(
            "return arguments[0].value;", login_element,
        ) or "")
    except Exception:
        actual_login = SUD_LOGIN  # элемент устарел — значение уже проверено в _fill_field
    if actual_login and actual_login != SUD_LOGIN:
        _fill_field(
            [
                ("css", "input[id*=':auth:xin']"),
                ("css", "input[name*=':auth:xin']"),
                ("css", "input[placeholder*='ИИН']"),
                ("css", "input[placeholder*='ЖСН']"),
            ],
            SUD_LOGIN, "Логин",
        )

    button = WebDriverWait(driver, 20).until(
        lambda d: find_login_button()
    )

    log.info("Логин/пароль введены. Нажимаю кнопку Войти.")

    clicked = False
    for attempt in range(1, 4):
        try:
            button = find_login_button()
            if button is None:
                raise RuntimeError("Кнопка Войти не найдена.")

            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});",
                button,
            )
            time.sleep(0.4)

            try:
                ActionChains(driver).move_to_element(
                    button
                ).pause(0.3).click().perform()
            except Exception:
                pass
            time.sleep(0.5)
            # onclick у кнопки — RichFaces.ajax(...); нативный click() надёжнее
            # прогоняет обработчик, чем синтетический ActionChains на некоторых
            # сборках Chrome.
            try:
                driver.execute_script("arguments[0].click();", button)
            except Exception:
                pass

            clicked = True
            log.info(
                "Кнопка Войти нажата, попытка %s/3.",
                attempt,
            )
            break

        except StaleElementReferenceException:
            log.warning(
                "Кнопка Войти перерисовалась. Повтор %s/3.",
                attempt,
            )
            time.sleep(0.8)

    if not clicked:
        raise RuntimeError(
            "Не удалось нажать кнопку Войти после трёх попыток."
        )

    # Ждём входа.
    try:
        WebDriverWait(driver, LOGIN_TIMEOUT).until(
            lambda d: _looks_authenticated(d)
        )
    except Exception as exc:
        debug_dir = Path(r"C:\Users\User\АИС ОИП1\login_debug")
        debug_dir.mkdir(parents=True, exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_path = debug_dir / f"login_error_{timestamp}.png"
        html_path = debug_dir / f"login_error_{timestamp}.html"

        try:
            driver.save_screenshot(str(screenshot_path))
            html_path.write_text(
                driver.page_source,
                encoding="utf-8",
            )
            log.error(
                "Диагностика входа сохранена: %s | %s",
                screenshot_path,
                html_path,
            )
        except Exception:
            pass

        page_text = str(driver.page_source or "").casefold()

        # видимый текст ошибок портала (rf-msgs / toast)
        portal_err = ""
        try:
            for el in driver.find_elements(
                "css selector", ".rf-msgs.error, .rf-msg-err, .ui-messages-error, .error"
            ):
                t = (el.text or "").strip()
                if t:
                    portal_err = t
                    break
        except Exception:
            pass

        # что реально в полях на момент провала
        fld = {}
        try:
            fld = driver.execute_script(
                """
                const g = s => { const e=[...document.querySelectorAll(s)]
                    .find(x=>x.offsetParent!==null); return e ? e.value : null; };
                return {xin: g("input[id*=':auth:xin']"),
                        pwd_len: (g("input[id*=':auth:password']")||'').length};
                """
            ) or {}
        except Exception:
            pass

        if portal_err:
            reason = f"портал сообщил: {portal_err}"
        elif "пользователь не найден" in page_text:
            reason = "портал сообщил: пользователь не найден"
        elif "проверьте правильность" in page_text or "неверн" in page_text:
            reason = "портал сообщил: проверьте логин или пароль"
        else:
            reason = (
                "кнопка Войти нажата, но кабинет не открылся и ошибки нет "
                f"(поле ИИН={fld.get('xin')!r}, длина пароля={fld.get('pwd_len')}); "
                f"URL={driver.current_url}. "
                "Проверьте логин/пароль вручную в этом же профиле Chrome — "
                "возможно, пароль в config.py устарел или аккаунту нужен вход по ЭЦП"
            )

        raise RuntimeError(
            f"Автоматический вход не выполнен: {reason}."
        ) from exc

    log.info(
        "Автоматический вход выполнен. URL: %s",
        driver.current_url,
    )



def get_auth_from_chrome_auto(
    debug_address: str = CHROME_DEBUG_ADDRESS,
) -> BrowserAuth:
    """
    Сам запускает Chrome, при необходимости выполняет вход,
    затем передаёт cookies в requests.
    """
    from selenium import webdriver

    _start_debug_chrome()

    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", debug_address)

    driver = webdriver.Chrome(options=options)
    try:
        def _snapshot() -> BrowserAuth:
            return BrowserAuth(
                cookies=driver.get_cookies(),
                user_agent=driver.execute_script("return navigator.userAgent"),
                current_url=driver.current_url,
            )

        if AUTO_LOGIN:
            _perform_login(driver)

        auth = _snapshot()
        log.info("Получено cookies из Chrome: %s шт.", len(auth.cookies))
        log.info("Текущая страница Chrome: %s", auth.current_url)

        # КЛЮЧЕВАЯ проверка: вкладка Chrome может показывать кабинет из
        # кэша, пока серверная сессия уже мертва. Проверяем реальным HTTP
        # GET рабочей формы; если это страница входа — принудительно
        # логинимся заново и берём cookies повторно.
        for attempt in range(1, 3):
            if not auth.cookies:
                log.warning("Chrome не отдал cookies (попытка %s).", attempt)
            elif _verify_http_session(auth):
                log.info("HTTP-сессия подтверждена (попытка %s).", attempt)
                return auth
            else:
                log.warning(
                    "HTTP-сессия НЕ авторизована (cookies=%s, попытка %s) — "
                    "принудительный повторный вход.",
                    len(auth.cookies), attempt,
                )

            if not AUTO_LOGIN:
                break
            _perform_login(driver, force=True)
            auth = _snapshot()
            log.info("После повторного входа cookies из Chrome: %s шт.",
                     len(auth.cookies))

        raise RuntimeError(
            "Не удалось получить рабочую авторизованную HTTP-сессию: портал "
            "отдаёт страницу входа даже после повторного входа. Проверьте "
            "логин/пароль (config.py, sk_login/sk_password для company_id) "
            "и вход по ЭЦП вручную в профиле Chrome "
            f"{CHROME_USER_DATA_DIR}."
        )
    finally:
        # Не закрываем Chrome: профиль остаётся доступным для проверки.
        try:
            driver.service.stop()
        except Exception:
            pass
        del driver



def make_http_session(auth: BrowserAuth) -> requests.Session:
    session = requests.Session()
    session.headers.update({
        "User-Agent": auth.user_agent,
        "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
        "Origin": BASE_URL,
    })

    for cookie in auth.cookies:
        kwargs = {
            "name": cookie["name"],
            "value": cookie["value"],
            "path": cookie.get("path", "/"),
        }
        domain = cookie.get("domain")
        if domain:
            kwargs["domain"] = domain
        session.cookies.set(**kwargs)

    return session


AUTH = get_auth_from_chrome_auto()
SESSION = make_http_session(AUTH)
print("Автоматический вход выполнен. HTTP session готова.")


# %% [markdown]
# ## 2. HTTP-клиент JSF / RichFaces

# %%

VIEWSTATE_NAMES = (
    "javax.faces.ViewState",
    "jakarta.faces.ViewState",
)


def clean_url(url: str) -> str:
    parts = urlsplit(url)
    return urlunsplit((parts.scheme, parts.netloc, parts.path, parts.query, ""))


def extract_viewstate_from_html(html: str) -> str | None:
    soup = BeautifulSoup(html, "lxml")
    for name in VIEWSTATE_NAMES:
        node = soup.find("input", attrs={"name": name})
        if node and node.get("value"):
            return str(node["value"])
    return None


def merge_partial_updates_into_html(
    current_html: str,
    updates: dict[str, str],
) -> str:
    """Применяет JSF partial-response updates к текущему полному DOM."""
    if not updates:
        return current_html

    for update_id, value in updates.items():
        if update_id.endswith("javax.faces.ViewRoot") or update_id == "javax.faces.ViewRoot":
            if value and "<" in value:
                return value

    if not current_html or "<html" not in current_html.lower():
        html_parts = [
            value for key, value in updates.items()
            if "ViewState" not in key and value and "<" in value
        ]
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


def extract_js_redirect_from_text(text: str) -> str | None:
    """Извлекает переходы из RichFaces eval/CDATA/скриптов."""
    if not text:
        return None

    patterns = [
        r"(?:window\.)?location(?:\.href)?\s*=\s*['\"]([^'\"]+)['\"]",
        r"document\.location(?:\.href)?\s*=\s*['\"]([^'\"]+)['\"]",
        r"(?:window\.)?location\.replace\(\s*['\"]([^'\"]+)['\"]\s*\)",
        r"(?:window\.)?location\.assign\(\s*['\"]([^'\"]+)['\"]\s*\)",
        r"window\.open\(\s*['\"]([^'\"]+)['\"]",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.I | re.S)
        if match:
            return match.group(1)
    return None


def parse_partial_response(xml_text: str) -> tuple[dict[str, str], str | None, str | None]:
    updates: dict[str, str] = {}
    redirect_url = None
    viewstate = None
    eval_scripts: list[str] = []

    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError:
        return updates, viewstate, redirect_url

    for node in root.iter():
        tag = node.tag.rsplit("}", 1)[-1]
        if tag == "update":
            update_id = node.attrib.get("id", "")
            value = node.text or ""
            updates[update_id] = value
            if "ViewState" in update_id and value.strip():
                viewstate = value.strip()
        elif tag == "redirect":
            redirect_url = node.attrib.get("url")
        elif tag == "eval":
            eval_scripts.append(node.text or "")

    # Некоторые версии JSF/RichFaces выполняют переход через <eval>,
    # а не через стандартный <redirect url="...">.
    if not redirect_url:
        for script in eval_scripts:
            patterns = [
                r"(?:window\.)?location(?:\.href)?\s*=\s*['\"]([^'\"]+)['\"]",
                r"document\.location(?:\.href)?\s*=\s*['\"]([^'\"]+)['\"]",
                r"location\.replace\(\s*['\"]([^'\"]+)['\"]\s*\)",
                r"location\.assign\(\s*['\"]([^'\"]+)['\"]\s*\)",
            ]
            for pattern in patterns:
                match = re.search(pattern, script, flags=re.I)
                if match:
                    redirect_url = match.group(1)
                    break
            if redirect_url:
                break

    if not redirect_url:
        redirect_url = extract_js_redirect_from_text(xml_text)

    return updates, viewstate, redirect_url


@dataclass
class HttpState:
    url: str
    html: str = ""
    viewstate: str | None = None
    updates: dict[str, str] = field(default_factory=dict)


class SudHttpClient:
    def __init__(self, session: requests.Session, base_url: str = BASE_URL):
        self.session = session
        self.base_url = base_url.rstrip("/")
        self.state = HttpState(url=BASE_URL + "/form/proceedings/services.xhtml")

    def _check_auth(self, response: requests.Response) -> None:
        if _response_is_login_page(response):
            raise SessionExpired(
                "Сессия не авторизована: портал вернул страницу входа "
                "(Сот кабинеті / ЖСН(БСН)/Құпия сөз). Cookies из Chrome "
                "недействительны — нужен повторный вход. "
                f"URL ответа: {response.url}"
            )

    def _check_unexpected_services_redirect(self, response: requests.Response, action: str) -> None:
        if "/form/proceedings/services.xhtml" in response.url.lower():
            raise RuntimeError(
                f"Сервер сбросил форму на список услуг во время действия: {action}. "
                "Обычно это означает неверную стартовую страницу, устаревший ViewState "
                "или пропущенный предыдущий шаг JSF. "
                f"URL ответа: {response.url}"
            )

    def _raw_request(
        self,
        method: str,
        full_url: str,
        **kwargs: Any,
    ) -> requests.Response:
        """
        Один HTTP-запрос с ретраями на транзиентных сбоях портала:
        обрыв соединения / SSL EOF / таймаут чтения и статусы
        HTTP_RETRY_STATUS (404/5xx). После всех попыток бросает
        PortalUnavailable — главный цикл ловит её отдельно.
        """
        last_reason = "неизвестно"
        for attempt in range(1, HTTP_RETRY_ATTEMPTS + 1):
            try:
                response = self.session.request(
                    method,
                    full_url,
                    timeout=HTTP_TIMEOUT,
                    verify=VERIFY_TLS,
                    **kwargs,
                )
            except (
                requests.exceptions.ConnectionError,
                requests.exceptions.Timeout,
                requests.exceptions.ChunkedEncodingError,
            ) as exc:
                last_reason = short_error(exc)
                if attempt >= HTTP_RETRY_ATTEMPTS:
                    raise PortalUnavailable(
                        f"{method} {full_url}: портал недоступен после "
                        f"{HTTP_RETRY_ATTEMPTS} попыток ({last_reason})"
                    ) from exc
            else:
                if response.status_code not in HTTP_RETRY_STATUS:
                    return response
                last_reason = f"HTTP {response.status_code}"
                if attempt >= HTTP_RETRY_ATTEMPTS:
                    raise PortalUnavailable(
                        f"{method} {full_url}: портал отвечает "
                        f"{response.status_code} после {HTTP_RETRY_ATTEMPTS} попыток"
                    )

            delay = HTTP_RETRY_BASE_DELAY * (2 ** (attempt - 1))
            log.warning(
                "Портал недоступен (%s), попытка %s/%s, жду %sс: %s",
                method, attempt, HTTP_RETRY_ATTEMPTS, delay, last_reason,
            )
            time.sleep(delay)

        raise PortalUnavailable(f"{method} {full_url}: {last_reason}")

    def get(self, url: str, *, referer: str | None = None) -> requests.Response:
        full_url = urljoin(self.base_url + "/", url)
        headers = {}
        if referer:
            headers["Referer"] = referer

        response = self._raw_request(
            "GET",
            full_url,
            headers=headers,
            allow_redirects=True,
        )
        response.raise_for_status()
        self._check_auth(response)

        self.state.url = response.url
        self.state.html = response.text
        html_viewstate = extract_viewstate_from_html(response.text)
        if html_viewstate:
            self.state.viewstate = html_viewstate

        log.info("GET %s -> %s", response.url, response.status_code)
        return response

    def post_form(
        self,
        url: str,
        data: dict[str, Any] | list[tuple[str, Any]],
        *,
        referer: str | None = None,
        ajax: bool = False,
        allow_redirects: bool = True,
        allow_services: bool = False,
    ) -> requests.Response:
        full_url = urljoin(self.base_url + "/", url)
        payload = list(data.items()) if isinstance(data, dict) else list(data)

        # Всегда заменяем старый ViewState актуальным.
        payload = [(k, v) for k, v in payload if k not in VIEWSTATE_NAMES]
        if self.state.viewstate:
            payload.append(("javax.faces.ViewState", self.state.viewstate))

        headers = {
            "Referer": referer or self.state.url,
        }
        if ajax:
            headers.update({
                "Accept": "application/xml, text/xml, */*; q=0.01",
                "Faces-Request": "partial/ajax",
                "X-Requested-With": "XMLHttpRequest",
            })

        response = self._raw_request(
            "POST",
            full_url,
            data=payload,
            headers=headers,
            allow_redirects=allow_redirects,
        )
        response.raise_for_status()
        self._check_auth(response)
        if not allow_services:
            self._check_unexpected_services_redirect(response, f"POST {full_url}")

        content_type = response.headers.get("Content-Type", "").lower()
        if "xml" in content_type or response.text.lstrip().startswith("<?xml"):
            updates, new_viewstate, redirect_url = parse_partial_response(response.text)
            self.state.updates = updates

            if new_viewstate:
                self.state.viewstate = new_viewstate

            self.state.html = merge_partial_updates_into_html(
                self.state.html,
                updates,
            )
            self.state.url = response.url

            if redirect_url:
                redirect_full = urljoin(response.url, redirect_url)
                log.info(
                    "JSF redirect после POST: %s -> %s",
                    response.url,
                    redirect_full,
                )
                return self.get(redirect_full, referer=response.url)
        else:
            self.state.url = response.url
            self.state.html = response.text
            html_viewstate = extract_viewstate_from_html(response.text)
            if html_viewstate:
                self.state.viewstate = html_viewstate

        log.info("POST %s -> %s", response.url, response.status_code)
        return response



    def upload(
        self,
        url: str,
        *,
        fields: list[tuple[str, str]],
        file_field: str,
        file_path: Path,
        referer: str | None = None,
    ) -> requests.Response:
        """
        Формирует multipart/form-data вручную, побайтно в стиле Chrome/WebKit.

        Тело запроса:
          1. ID JSF-формы;
          2. актуальный javax.faces.ViewState;
          3. бинарное содержимое файла.

        Стандартный requests(files=...) здесь намеренно не используется:
        старый RichFaces 4.5.17 чувствителен к формату multipart-заголовков,
        boundary и кодировке кириллического filename.
        """
        import random
        import secrets

        if not file_path.exists():
            raise FileNotFoundError(file_path)

        if not self.state.viewstate:
            raise RuntimeError(
                f"Перед загрузкой {file_path.name} отсутствует актуальный ViewState."
            )

        # Новый UID, как Math.random() в браузере.
        fresh_uid = f"0.{random.SystemRandom().randrange(10**15, 10**16)}"

        parts = urlsplit(urljoin(self.base_url + "/", url))
        query = parse_qsl(parts.query, keep_blank_values=True)

        new_query: list[tuple[str, str]] = []
        seen_uid = False
        seen_viewstate = False

        for key, value in query:
            if key == "rf_fu_uid":
                new_query.append((key, fresh_uid))
                seen_uid = True
            elif key in VIEWSTATE_NAMES:
                new_query.append(("javax.faces.ViewState", self.state.viewstate))
                seen_viewstate = True
            else:
                new_query.append((key, value))

        if not seen_uid:
            new_query.append(("rf_fu_uid", fresh_uid))
        if not seen_viewstate:
            new_query.append(("javax.faces.ViewState", self.state.viewstate))

        upload_url = urlunsplit(
            (
                parts.scheme,
                parts.netloc,
                parts.path,
                urlencode(new_query),
                "",
            )
        )

        # Chrome создаёт boundary с префиксом WebKitFormBoundary.
        boundary = "----WebKitFormBoundary" + secrets.token_urlsafe(12).replace("-", "A").replace("_", "B")[:16]
        boundary_bytes = boundary.encode("ascii")
        crlf = b"\r\n"

        mime = mimetypes.guess_type(file_path.name)[0] or "application/octet-stream"
        file_bytes = file_path.read_bytes()

        # Защита от кавычек и переводов строк в имени файла.
        safe_filename = (
            file_path.name
            .replace("\\", "_")
            .replace('"', "'")
            .replace("\r", "_")
            .replace("\n", "_")
        )

        body = bytearray()

        def add_text_part(name: str, value: str) -> None:
            body.extend(b"--" + boundary_bytes + crlf)
            header = (
                f'Content-Disposition: form-data; name="{name}"'
            ).encode("utf-8")
            body.extend(header + crlf + crlf)
            body.extend(str(value).encode("utf-8"))
            body.extend(crlf)

        # В capture браузера multipart содержит только форму и ViewState.
        for key, value in fields:
            if key not in VIEWSTATE_NAMES:
                add_text_part(str(key), str(value))

        add_text_part("javax.faces.ViewState", self.state.viewstate)

        # Файловая часть: точный порядок заголовков как у Chrome.
        body.extend(b"--" + boundary_bytes + crlf)
        disposition = (
            f'Content-Disposition: form-data; name="{file_field}"; '
            f'filename="{safe_filename}"'
        ).encode("utf-8")
        body.extend(disposition + crlf)
        body.extend(f"Content-Type: {mime}".encode("ascii") + crlf)
        body.extend(crlf)
        body.extend(file_bytes)
        body.extend(crlf)
        body.extend(b"--" + boundary_bytes + b"--" + crlf)

        raw_body = bytes(body)

        headers = {
            "Accept": "*/*",
            "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
            "Cache-Control": "no-cache",
            "Pragma": "no-cache",
            "Faces-Request": "partial/ajax",
            "Origin": self.base_url,
            "Referer": referer or self.state.url,
            "Content-Type": f"multipart/form-data; boundary={boundary}",
        }

        log.info(
            "RAW multipart: файл=%s | bytes=%s | boundary=%s | "
            "filename_utf8=%s",
            file_path.name,
            len(raw_body),
            boundary,
            safe_filename,
        )

        response = self.session.post(
            upload_url,
            data=raw_body,
            headers=headers,
            timeout=HTTP_TIMEOUT,
            verify=VERIFY_TLS,
            allow_redirects=False,
        )

        response.raise_for_status()
        self._check_auth(response)
        self._check_unexpected_services_redirect(
            response,
            f"UPLOAD {file_path.name}",
        )

        updates, new_viewstate, redirect_url = parse_partial_response(response.text)

        if new_viewstate:
            self.state.viewstate = new_viewstate
        if updates:
            self.state.updates = updates

        self.state.html = merge_partial_updates_into_html(
            self.state.html,
            updates,
        )
        self.state.url = response.url

        if not hasattr(self, "last_uploaded_files"):
            self.last_uploaded_files = {}
        self.last_uploaded_files[file_field] = file_path.name
        self.last_uploaded_file = file_path.name

        if redirect_url:
            redirect_full = urljoin(response.url, redirect_url)
            log.info(
                "JSF redirect после upload: %s -> %s",
                response.url,
                redirect_full,
            )
            return self.get(
                redirect_full,
                referer=response.url,
            )

        log.info(
            "UPLOAD %s -> %s | field=%s | uid=%s | body_bytes=%s | %s",
            file_path.name,
            response.status_code,
            file_field,
            fresh_uid,
            len(raw_body),
            response.url,
        )
        return response


CLIENT = SudHttpClient(SESSION)


# %% [markdown]
# ## 3. Чтение и нормализация сохранённых перехватов

# %%
@dataclass
class CapturedRequest:
    source_file: str
    order: int
    method: str
    url: str
    request_type: str
    headers: dict[str, str]
    post_data: str | None
    parsed_form_data: Any = None
    mime_type: str | None = None
    har_index: int | None = None
    upload_extra_index: int | None = None

    @property
    def is_http_action(self) -> bool:
        method = self.method.upper()
        if method == "POST":
            return True
        if method == "GET" and urlsplit(self.url).path.lower().endswith(
            "/form/requesttype2/sign.xhtml"
        ):
            return True
        return self.request_type.lower() in {"xhr", "fetch"} and method != "OPTIONS"


def load_recipe() -> list[CapturedRequest]:
    """Строит последовательность бизнес-запросов из podacha_recipe.RECIPE_STEPS.

    Раньше эта последовательность разбиралась на лету из бинарного HAR
    (office.sud.kz_CURRENT_PARTICIPANTS.har) плюс резервных JSON в
    SUD_API_CAPTURES/. Теперь она заморожена в podacha_recipe.py — тот же
    самый разобранный список, но без внешних файлов. См. модуль-рецепт,
    как его регенерировать, если портал сменит динамические JSF-id.
    """
    result: list[CapturedRequest] = []
    for step in RECIPE_STEPS:
        result.append(CapturedRequest(
            source_file=f"recipe#HAR-{step['har_index']:03d}",
            order=step["har_index"],
            har_index=step["har_index"],
            method=str(step["method"]).upper(),
            url=str(step["url"]),
            request_type=str(step.get("request_type") or ""),
            headers=dict(step.get("headers") or {}),
            post_data=step["post_data"],
            mime_type=str(step.get("mime_type") or ""),
        ))

    # Нумеруем обычные приложения строго в порядке multipart из рецепта.
    extra_no = 0
    for req in result:
        query = dict(parse_qsl(urlsplit(req.url).query, keep_blank_values=True))
        source_id = query.get("javax.faces.source", "")
        if "selectFileUploader" in source_id:
            extra_no += 1
            req.upload_extra_index = extra_no

    log.info(
        "Из рецепта загружено бизнес-запросов: %s; обычных файлов: %s",
        len(result),
        extra_no,
    )
    return result

def print_actions(captured: list[CapturedRequest]) -> None:
    """Список разобранных HAR-шагов — техническая информация для разработчика,
    в консоль/веб-лог не выводится, только в sud_http_api.log."""
    for req in captured:
        if req.is_http_action:
            suffix = (
                f" | EXTRA_{req.upload_extra_index}"
                if req.upload_extra_index is not None
                else ""
            )
            log.info(
                "%s | #%03d | %4s | %8s | %s%s",
                req.source_file, req.order, req.method, req.request_type, req.url, suffix,
            )

# %% [markdown]
# ## 4. Данные строки Excel и поиск файлов дела

# %%

def cell_text(ws, column: str, row: int) -> str:
    value = ws[f"{column}{row}"].value
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_iin(value: str) -> str:
    digits = re.sub(r"\D", "", value or "")
    return digits.zfill(12) if digits else ""


def normalize_money_cell(value: Any, *, field_name: str, row: int) -> str:
    """
    Приводит денежное значение Excel к виду, который принимает JSF:
    40168.76

    Поддерживает:
    40 168,76
    40168.76
    числовую ячейку Excel
    """
    if value is None:
        raise ValueError(
            f"Строка {row}: в Excel не заполнено поле {field_name}."
        )

    if isinstance(value, bool):
        raise ValueError(
            f"Строка {row}: некорректное значение {field_name}: {value!r}"
        )

    if isinstance(value, (int, float)):
        text = f"{value:.2f}"
    else:
        text = str(value).strip()
        text = (
            text.replace("\u00a0", "")
                .replace(" ", "")
                .replace("₸", "")
                .replace("тг", "")
                .replace(",", ".")
        )

    if not re.fullmatch(r"\d+(?:\.\d{1,2})?", text):
        raise ValueError(
            f"Строка {row}: некорректное значение {field_name}: {value!r}"
        )

    whole, dot, fraction = text.partition(".")
    fraction = (fraction + "00")[:2]
    return f"{int(whole)}.{fraction}"


def load_case_row(row: int) -> dict[str, str]:
    wb = load_workbook(EXCEL_FILE, data_only=True, read_only=True)
    try:
        ws = wb[EXCEL_SHEET]

        claim_sum = normalize_money_cell(
            ws[f"{COL_SUM}{row}"].value,
            field_name=f"Сумма иска ({COL_SUM})",
            row=row,
        )
        duty_sum = normalize_money_cell(
            ws[f"{COL_DUTY}{row}"].value,
            field_name=f"Госпошлина ({COL_DUTY})",
            row=row,
        )

        result = {
            "ROW": str(row),
            "FIO": cell_text(ws, COL_FIO, row),
            "IIN": normalize_iin(cell_text(ws, COL_IIN, row)),
            "CLAIM_SUM": claim_sum,
            "DUTY_SUM": duty_sum,
            "REGION": cell_text(ws, COL_REGION, row),
            "COURT": cell_text(ws, COL_COURT, row),
            "ORG_BIN": ORG_BIN,
            "ORG_BANK": ORG_BANK,
            "REP_IIN": REP_IIN,
        }

        log.info(
            "Денежные данные Excel, строка %s: "
            "Сумма иска=%s | Госпошлина=%s",
            row,
            claim_sum,
            duty_sum,
        )
        return result
    finally:
        wb.close()



def selected_batch_folder() -> Path:
    """
    Возвращает строго заданную папку партии.

    Автоматический выбор по времени изменения допускается только когда
    CASES_BATCH_DIR=None. Это защищает от случайного перехода в другую
    папку, где может находиться старый файл с тем же названием.
    """
    if CASES_BATCH_DIR is not None:
        batch = Path(CASES_BATCH_DIR).resolve()
        if not batch.exists():
            raise FileNotFoundError(
                f"Указанная папка партии не найдена: {batch}"
            )
        if not batch.is_dir():
            raise NotADirectoryError(batch)
        return batch

    # Настоящие папки партий называются строго по шаблону из config.py
    # (_make_target_base): "ДД.ММ.ГГГГ (ЧЧ-ММ)". В той же родительской
    # папке лежат и служебные каталоги вроде "Шаблоны документов" — их
    # нельзя пускать в автовыбор: если у служебной папки окажется более
    # свежий mtime (например, кто-то просто открыл файл из неё), автовыбор
    # ошибочно возьмёт её вместо реальной партии, и все дела в строке
    # "не найдена папка дела" посыпятся разом.
    batch_name_re = re.compile(r"^\d{2}\.\d{2}\.\d{4} \(\d{2}-\d{2}\)$")
    dirs = [
        p for p in CASES_BASE_DIR.iterdir()
        if p.is_dir() and batch_name_re.match(p.name)
    ]
    if not dirs:
        raise FileNotFoundError(
            f"Нет папок партии вида «ДД.ММ.ГГГГ (ЧЧ-ММ)» в {CASES_BASE_DIR}"
        )

    batch = max(dirs, key=lambda p: p.stat().st_mtime).resolve()
    log.warning(
        "CASES_BATCH_DIR не задана. Автоматически выбрана папка: %s",
        batch,
    )
    return batch


def find_case_folder(iin: str, fio: str) -> Path:
    batch = selected_batch_folder()
    normalized_fio = re.sub(r"\s+", " ", fio).strip().casefold()

    iin_matches: list[Path] = []
    fio_matches: list[Path] = []

    for path in batch.iterdir():
        if not path.is_dir():
            continue

        folder_digits = re.sub(r"\D", "", path.name)
        folder_name = re.sub(r"\s+", " ", path.name).casefold()

        if iin and iin in folder_digits:
            iin_matches.append(path.resolve())
        elif normalized_fio and normalized_fio in folder_name:
            fio_matches.append(path.resolve())

    matches = iin_matches or fio_matches

    if not matches:
        raise FileNotFoundError(
            f"Не найдена папка дела в партии {batch}: "
            f"ИИН={iin}, ФИО={fio}"
        )

    if len(matches) > 1:
        raise RuntimeError(
            "Найдено несколько папок текущего должника: "
            + "; ".join(str(path) for path in matches)
        )

    folder = matches[0]
    log.info("ТОЧНАЯ папка текущего должника: %s", folder)
    return folder


def classify_case_files(folder: Path) -> dict[str, Path]:
    """
    Создаёт независимый файловый набор текущего должника.

    HAR больше не определяет, какой бинарный файл загружается.
    Обычные приложения формируются как последовательная очередь:
    ATTACHMENT_001, ATTACHMENT_002 и т.д.
    """
    allowed_suffixes = {
        ".pdf", ".doc", ".docx", ".jpg", ".jpeg", ".png"
    }

    folder = folder.resolve()

    files = sorted(
        [
            path.resolve()
            for path in folder.iterdir()
            if path.is_file()
            and path.suffix.casefold() in allowed_suffixes
        ],
        key=lambda path: path.name.casefold(),
    )

    # Ограничение портала: каждый файл не более 20 МБ. Такие файлы просто
    # не участвуют в загрузке (кроме искового/госпошлины — для них ниже
    # по-прежнему сработает обязательная проверка на отсутствие).
    kept_files = []
    for path in files:
        size_mb = path.stat().st_size / (1024 * 1024)
        if path.stat().st_size > MAX_UPLOAD_FILE_BYTES:
            log.warning(
                "Пропускаю файл больше 20 МБ (портал не примет): "
                "%s | %.2f МБ | %s",
                path.name, size_mb, path,
            )
            continue
        kept_files.append(path)
    files = kept_files

    result: dict[str, Path] = {}
    attachments: list[Path] = []

    for path in files:
        low = path.name.casefold().replace("ё", "е")

        if path.suffix.casefold() == ".docx" and "исков" in low:
            if "CLAIM_DOCX" in result:
                raise RuntimeError(
                    "Найдено несколько исковых заявлений: "
                    f"{result['CLAIM_DOCX'].name}; {path.name}"
                )
            result["CLAIM_DOCX"] = path
            continue

        if path.suffix.casefold() == ".pdf" and "госпош" in low:
            if "DUTY_PDF" in result:
                raise RuntimeError(
                    "Найдено несколько файлов госпошлины: "
                    f"{result['DUTY_PDF'].name}; {path.name}"
                )
            result["DUTY_PDF"] = path
            continue

        attachments.append(path)

    if "CLAIM_DOCX" not in result:
        raise FileNotFoundError(
            f"В корне папки {folder} не найдено исковое заявление DOCX"
        )

    if "DUTY_PDF" not in result:
        raise FileNotFoundError(
            f"В корне папки {folder} не найдена госпошлина PDF"
        )

    for index, path in enumerate(attachments, start=1):
        result[f"ATTACHMENT_{index:03d}"] = path

    log.info(
        "ФАЙЛОВЫЙ МАНИФЕСТ текущего должника (%s):",
        folder,
    )
    for key, path in result.items():
        log.info(
            "  %s | %s | %.2f МБ | %s",
            key,
            path.name,
            path.stat().st_size / (1024 * 1024),
            path,
        )

    return result


def get_attachment_queue(case_files: dict[str, Path]) -> list[Path]:
    """Возвращает приложения строго в порядке ATTACHMENT_001..."""
    return [
        case_files[key]
        for key in sorted(case_files)
        if key.startswith("ATTACHMENT_")
    ]


def validate_case_files(
    folder: Path,
    case_files: dict[str, Path],
    context: dict[str, str],
) -> None:
    """Проверяет путь, размер и принадлежность каждого фактического файла."""
    root = folder.resolve()

    oversized: list[str] = []

    for key, original_path in case_files.items():
        path = Path(original_path).resolve()

        if not path.exists():
            raise FileNotFoundError(path)

        try:
            path.relative_to(root)
        except ValueError as exc:
            raise RuntimeError(
                f"Файл {key} находится вне папки текущего дела: {path}"
            ) from exc

        # Файл обязан находиться непосредственно в корне папки должника.
        if path.parent != root:
            raise RuntimeError(
                f"Файл найден во вложенной или чужой папке: {path}"
            )

        size_bytes = path.stat().st_size
        size_mb = size_bytes / (1024 * 1024)

        log.info(
            "Проверка файла: %s | %s | %.2f МБ | путь=%s",
            key,
            path.name,
            size_mb,
            path,
        )

        if size_bytes > MAX_UPLOAD_FILE_BYTES:
            oversized.append(
                f"{path.name} — {size_mb:.2f} МБ — {path}"
            )

    if oversized:
        raise RuntimeError(
            "Найдены файлы больше допустимых 20 МБ: "
            + "; ".join(oversized)
        )

    duty = case_files["DUTY_PDF"]
    stem = duty.stem.casefold().replace("ё", "е")
    surname = (
        (context.get("FIO", "").split() or [""])[0]
        .casefold()
        .replace("ё", "е")
    )
    iin = context.get("IIN", "")

    if "," in duty.stem and surname and surname not in stem and iin not in stem:
        raise RuntimeError(
            "Госпошлина, вероятно, относится к другому человеку: "
            f"{duty.name}. Текущая строка: "
            f"{context.get('FIO')} / {iin}"
        )


def read_docx_text(path: Path) -> str:
    from docx import Document
    doc = Document(path)
    parts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            text = " ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
            if text:
                parts.append(text)
    return "\n".join(parts)


# %% [markdown]
# ## 5. Подстановка значений и HTTP-воспроизведение перехватов

# %%

# Служебные поля браузера не должны слепо переноситься в requests.
DROP_HEADERS = {
    "content-length", "cookie", "host", "connection",
    "sec-fetch-dest", "sec-fetch-mode", "sec-fetch-site",
    "sec-ch-ua", "sec-ch-ua-mobile", "sec-ch-ua-platform",
}

# Поля, которые портал сам заполняет после AJAX-поиска по ИИН/БИН
# (ФИО/адрес участника, название и адрес организации). Их всегда нужно
# брать из свежего ответа сервера, а не из записанного HAR.
#
# org-factAddress (фактический адрес) сюда НЕ входит: по факту (проверено
# по реальному debug-снимку live_dom_fields) поиск по БИН его не
# возвращает вообще — портал отдаёт только org-name/org-jurAddress/
# org-bankDetails. Фактический адрес заполняется отдельно, см.
# _mirror_org_fact_address ниже.
LIVE_LOOKUP_FIELD_SUFFIXES = (
    "org-name", "org-juraddress", "org-bankdetails",
    "person-surname", "person-firstname", "person-patronymic", "person-liveplace",
)


def _mirror_org_fact_address(form_data: list[tuple[str, str]]) -> list[tuple[str, str]]:
    """
    Портал не возвращает фактический адрес истца через поиск по БИН —
    только юридический. В реальном образце (см. sud_participants_debug)
    фактический адрес совпадает с юридическим, поэтому копируем его,
    если поле факт.адреса осталось пустым.
    """
    jur_value = next(
        (v for k, v in form_data if k.lower().endswith("org-juraddress") and str(v or "").strip()),
        "",
    )
    if not jur_value:
        return form_data
    return [
        (k, jur_value) if k.lower().endswith("org-factaddress") and not str(v or "").strip() else (k, v)
        for k, v in form_data
    ]


def _require_live_lookup_fields(
    form_data: list[tuple[str, str]],
    *,
    role: str,
) -> None:
    """
    На "save"-шаге участника (истец/ответчик/представитель) убеждается,
    что портал реально прислал непустые ФИО/адрес/реквизиты после поиска
    по ИИН/БИН, а не оставил пустые заглушки из очищенного HAR.

    Раньше подстановка из live_values была best-effort: если поиск по
    ИИН/БИН не сработал (например, портал не успел ответить под нагрузкой),
    поле просто оставалось пустым и участник уходил в иск без данных —
    без единой ошибки в логе.
    """
    missing = [
        key for key, value in form_data
        if any(key.lower().endswith(suf) for suf in LIVE_LOOKUP_FIELD_SUFFIXES)
        and not str(value or "").strip()
    ]
    if missing:
        raise RuntimeError(
            f"Участник «{role}»: после поиска по ИИН/БИН поля остались "
            f"пустыми: {missing}. Портал не вернул данные (возможно, "
            "перегружен или сбоил поиск). Строка будет пропущена."
        )


def replace_tokens(value: str, context: dict[str, str]) -> str:
    result = value
    # Поддерживаются шаблоны {{IIN}}, ${IIN} и __IIN__.
    for key, replacement in context.items():
        replacement = "" if replacement is None else str(replacement)
        result = result.replace("{{" + key + "}}", replacement)
        result = result.replace("${" + key + "}", replacement)
        result = result.replace("__" + key + "__", replacement)
    return result


def replace_known_captured_values(value: str, context: dict[str, str]) -> str:
    """
    Дополнительная замена известных тестовых значений из перехватов.
    При необходимости добавьте сюда исходные значения из своих captures.
    """
    replacements = {
        # ИИН/БИН из тестовых перехватов:
        CAPTURED_ORG_BIN: context.get("ORG_BIN", ""),
        CAPTURED_DEFENDANT_IIN: context.get("IIN", ""),
        CAPTURED_REP_IIN: context.get("REP_IIN", ""),
        # Банковский счёт:
        CAPTURED_ORG_BANK: context.get("ORG_BANK", ""),

        # Тестовые суммы из исходного HAR.
        # Основная подстановка выполняется по DOM-подписям на payment.xhtml,
        # эти варианты оставлены как дополнительная страховка.
        "40168.76": context.get("CLAIM_SUM", ""),
        "40168,76": context.get("CLAIM_SUM", ""),
        "40 168,76": context.get("CLAIM_SUM", ""),
        "1205.06": context.get("DUTY_SUM", ""),
        "1205,06": context.get("DUTY_SUM", ""),
        "1 205,06": context.get("DUTY_SUM", ""),
    }
    result = value
    for old, new in replacements.items():
        if old and new:
            result = result.replace(old, new)
    return replace_tokens(result, context)


def parse_post_data(post_data: str | None) -> list[tuple[str, str]]:
    if not post_data:
        return []
    return [(k, v) for k, v in parse_qsl(post_data, keep_blank_values=True)]


def request_signature(req: CapturedRequest) -> tuple[str, str, str, str]:
    parts = urlsplit(req.url)
    content_type = next(
        (str(v).lower() for k, v in req.headers.items() if k.lower() == "content-type"),
        "",
    )

    # В v14 multipart-загрузки обычного документа и иска ошибочно считались
    # дублями: у них одинаковые method/path и пустой post_data. Учитываем query
    # и Content-Type, чтобы selectFileUploader и selectLawsuitScanUploader
    # остались двумя отдельными действиями.
    return (
        req.method,
        parts.path,
        parts.query,
        (req.post_data or "") + "|" + content_type,
    )


def deduplicate_actions(actions: list[CapturedRequest]) -> list[CapturedRequest]:
    """
    Удаляет точные повторы, появившиеся из-за отдельных перехватов одного шага.
    Порядок первых вхождений сохраняется.
    """
    seen = set()
    result = []
    for req in actions:
        sig = request_signature(req)
        if sig in seen:
            continue
        seen.add(sig)
        result.append(req)
    return result


def is_sign_page(url: str, text: str = "") -> bool:
    """
    Завершать строку можно только на фактической странице подписания.

    На createRequest.xhtml и других страницах в меню/скриптах также встречаются
    слова «Подписание» и «ЭЦП», поэтому анализ HTML здесь намеренно не используется.
    """
    path = urlsplit(url or "").path.lower().rstrip("/")
    return path.endswith("/form/requesttype2/sign.xhtml")


def is_service_request(req: CapturedRequest) -> bool:
    low = req.url.lower()
    return any(marker in low for marker in (
        "__richfaces_push",
        "pushresource.xhtml",
        "/rfres/",
        "javax.faces.resource",
        "org.richfaces.resources",
        ".css", ".js", ".png", ".jpg", ".jpeg", ".gif", ".woff", ".woff2",
    ))


def is_multipart_request(req: CapturedRequest) -> bool:
    content_type = ""
    for key, value in req.headers.items():
        if key.lower() == "content-type":
            content_type = value.lower()
            break
    return "multipart/form-data" in content_type


PARTICIPANTS_DEBUG_DIR = Path("sud_participants_debug")
PARTICIPANTS_DEBUG_DIR.mkdir(parents=True, exist_ok=True)


def _write_participant_debug(
    har_index: int,
    stage: str,
    payload: dict,
) -> None:
    """Сохраняет подробности первых createRequest POST без изменения запросов."""
    try:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        path = PARTICIPANTS_DEBUG_DIR / f"HAR_{har_index:03d}_{stage}_{stamp}.json"
        path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2, default=str),
            encoding="utf-8",
        )
        log.info("DEBUG HAR-%03d %s сохранён: %s", har_index, stage, path.resolve())
    except Exception:
        log.exception("Не удалось сохранить DEBUG HAR-%03d %s", har_index, stage)



REGION_MAP_HTTP = {
    "ГОРОД АСТАНА": "город Астана",
    "АСТАНА": "город Астана",
    "ГОРОД АЛМАТЫ": "город Алматы",
    "АЛМАТЫ": "город Алматы",
    "ГОРОД ШЫМКЕНТ": "город Шымкент",
    "ШЫМКЕНТ": "город Шымкент",
    "АКМОЛИНСКАЯ ОБЛАСТЬ": "Акмолинская область",
    "АКТЮБИНСКАЯ ОБЛАСТЬ": "Актюбинская область",
    "АЛМАТИНСКАЯ ОБЛАСТЬ": "Алматинская область",
    "АТЫРАУСКАЯ ОБЛАСТЬ": "Атырауская область",
    "ВОСТОЧНО-КАЗАХСТАНСКАЯ ОБЛАСТЬ": "Восточно-Казахстанская область",
    "ЖАМБЫЛСКАЯ ОБЛАСТЬ": "Жамбылская область",
    "ЗАПАДНО-КАЗАХСТАНСКАЯ ОБЛАСТЬ": "Западно-Казахстанская область",
    "КАРАГАНДИНСКАЯ ОБЛАСТЬ": "Карагандинская область",
    "КОСТАНАЙСКАЯ ОБЛАСТЬ": "Костанайская область",
    "КЫЗЫЛОРДИНСКАЯ ОБЛАСТЬ": "Кызылординская область",
    "МАНГИСТАУСКАЯ ОБЛАСТЬ": "Мангистауская область",
    "ПАВЛОДАРСКАЯ ОБЛАСТЬ": "Павлодарская область",
    "СЕВЕРО-КАЗАХСТАНСКАЯ ОБЛАСТЬ": "Северо-Казахстанская область",
    "ТУРКЕСТАНСКАЯ ОБЛАСТЬ": "Туркестанская область",
    "ОБЛАСТЬ ҰЛЫТАУ": "Область Ұлытау",
    "ОБЛАСТЬ АБАЙ": "Область Абай",
    "ОБЛАСТЬ ЖЕТІСУ": "Область Жетісу",
    "ВОЕННЫЙ СУД РЕСПУБЛИКИ КАЗАХСТАН": "Военный суд Республики Казахстан",
    "УПРАЗДНЁННЫЕ СУДЫ (С 2022 ГОДА)": "Упразднённые суды (c 2022 года)",
}


def _norm_option_text(value: str) -> str:
    value = str(value or "").replace("ё", "е").replace("Ё", "Е")
    return re.sub(r"\s+", " ", value).strip().casefold()


def _select_value_from_live_dom(
    soup: BeautifulSoup,
    field_suffix: str,
    wanted_text: str,
    *,
    partial_match: bool = False,
) -> tuple[str | None, str | None]:
    """
    Находит value опции по видимому тексту в актуальном DOM.
    Возвращает (value, фактический текст опции).
    """
    wanted_norm = _norm_option_text(wanted_text)
    if not wanted_norm:
        return None, None

    select = soup.find(
        "select",
        attrs={
            "name": lambda value: bool(value)
            and str(value).lower().endswith(field_suffix.lower())
        },
    )
    if select is None:
        select = soup.find(
            "select",
            id=lambda value: bool(value)
            and str(value).lower().endswith(field_suffix.lower()),
        )
    if select is None:
        return None, None

    options = []
    for option in select.find_all("option"):
        text = option.get_text(" ", strip=True)
        value = str(option.get("value") or "")
        if value:
            options.append((value, text, _norm_option_text(text)))

    # Сначала только точное совпадение.
    for value, text, normalized in options:
        if normalized == wanted_norm:
            return value, text

    # Для суда в исходном Selenium-скрипте был fallback на частичное совпадение.
    if partial_match:
        for value, text, normalized in options:
            if wanted_norm in normalized or normalized in wanted_norm:
                return value, text

    return None, None


ASTANA_DISTRICT_VALUE = "2"   # 'город Астана' — значение по умолчанию, зашитое в HAR
_HAR_DEFAULT_DISTRICT = "2"   # район из шаблона (город Астана)
_HAR_DEFAULT_COURT = "489"    # суд из шаблона (Астана) — согласован с районом 2

# Что делать, если суд из Excel не удалось сопоставить со списком портала:
#   True  (по умолчанию) — иск всё равно подаётся, но с судом из шаблона
#          (Астана); такие строки собираются в DEFAULT_COURT_ROWS и явно
#          перечисляются в итоге — их нужно проверить/переподать вручную.
#   False (OMEGA_STRICT_COURT=1) — строка останавливается с ошибкой.
ALLOW_DEFAULT_COURT_FALLBACK = os.environ.get(
    "OMEGA_STRICT_COURT", ""
).strip().lower() not in {"1", "true", "yes"}

# (row, court_name) для строк, поданных с судом по умолчанию.
DEFAULT_COURT_ROWS: list[tuple[str, str]] = []


def _dump_case_page(client: SudHttpClient, tag: str) -> None:
    try:
        d = Path(
            os.environ.get("OMEGA_WORKDIR")
            or os.environ.get("OMEGA_OUT")
            or "."
        )
        token = ""
        try:
            token = _current_request_token(client)
        except Exception:
            pass
        (d / f"{tag}_{token or 'nofile'}.html").write_text(
            client.state.html or "", encoding="utf-8"
        )
        log.error("HTML createRequest сохранён: %s/%s_%s.html", d, tag, token)
    except Exception:
        pass


# Подсказки «текст названия суда из Excel -> текст региона в портале».
# ВАЖНО: проверяются строго по порядку. Города-республиканского значения
# ('...районный суд города Алматы') идут ПЕРВЫМИ — иначе 'Жетысуский
# районный суд города Алматы' ошибочно уходит в 'Область Жетісу'.
_COURT_NAME_REGION_HINTS: tuple[tuple[tuple[str, ...], str], ...] = (
    (("города алматы", "город алматы", "г. алматы", "г.алматы"), "город Алматы"),
    (("города астаны", "город астана", "города нур-султан", "город нур-султан",
      "г. астана", "г.астана"), "город Астана"),
    (("города шымкента", "город шымкент", "г. шымкент", "г.шымкент"), "город Шымкент"),
    (("акмолинск", "кокшетау"), "Акмолинская область"),
    (("актюбинск", "актобе"), "Актюбинская область"),
    (("алматинск", "карасайск", "талгарск", "илийск", "енбекшиказахск",
      "капшагай", "конаев", "капчагай"), "Алматинская область"),
    (("атыраус", "атырау"), "Атырауская область"),
    (("восточно-казахстанск", "усть-каменогорск", "риддер"), "Восточно-Казахстанская область"),
    (("жамбылск", "тараз"), "Жамбылская область"),
    (("западно-казахстанск", "уральск"), "Западно-Казахстанская область"),
    (("карагандинск", "караганд", "темиртау", "балхаш", "сарань", "шахтинск"), "Карагандинская область"),
    (("костанайск", "костанай", "рудный", "лисаковск"), "Костанайская область"),
    (("кызылординск", "кызылорда"), "Кызылординская область"),
    (("мангистаус", "актау", "жанаозен"), "Мангистауская область"),
    (("павлодарск", "павлодар", "экибастуз", "аксу"), "Павлодарская область"),
    (("северо-казахстанск", "петропавловск"), "Северо-Казахстанская область"),
    (("туркестанск", "сайрам", "кентау", "сарыагаш", "туркестан"), "Туркестанская область"),
    (("улытау", "ұлытау", "жезказган", "жезказганск", "сатпаев"), "Область Ұлытау"),
    (("область абай", "абайск", "семей", "курчатов"), "Область Абай"),
    (("область жетісу", "жетісу", "жетысус", "талдыкорган", "текели"), "Область Жетісу"),
)

# region_value -> [(court_value, court_text)] — стабильно в рамках одного запуска.
_REGION_COURTS_CACHE: dict[str, list[tuple[str, str]]] = {}


def _district_options(soup: BeautifulSoup) -> list[tuple[str, str]]:
    sel = _find_select_by_suffix(soup, ":edit-district")
    if sel is None:
        return []
    return [
        (str(o.get("value") or ""), o.get_text(" ", strip=True))
        for o in sel.find_all("option")
        if str(o.get("value") or "").strip()
    ]


def _rank_region_candidates(court_name: str, districts: list[tuple[str, str]]) -> list[tuple[str, str]]:
    """Упорядочивает (region_value, region_text) — самый вероятный регион первым,
    остальные следом (для перебора, если подсказка не сработала)."""
    norm = _norm_option_text(court_name)
    by_text = {_norm_option_text(t): (v, t) for v, t in districts}

    ordered: list[tuple[str, str]] = []
    seen: set[str] = set()

    for needles, region_text in _COURT_NAME_REGION_HINTS:
        if any(n in norm for n in needles):
            hit = by_text.get(_norm_option_text(region_text))
            if hit and hit[0] not in seen:
                ordered.append(hit)
                seen.add(hit[0])

    # Затем — все остальные регионы (крупные города и области раньше «прочего»).
    for v, t in districts:
        if v and v not in seen:
            ordered.append((v, t))
            seen.add(v)
    return ordered


def _match_court(court_name: str, courts: list[tuple[str, str]]) -> tuple[str, str] | None:
    want = _norm_option_text(court_name)
    if not want:
        return None
    for v, t in courts:
        if _norm_option_text(t) == want:
            return v, t
    for v, t in courts:
        nt = _norm_option_text(t)
        if want and (want in nt or nt in want):
            return v, t
    return None


def _infer_region_text(court_name: str) -> str | None:
    """Определяет текст региона портала по названию суда из Excel (колонка P).

    Колонка P теперь содержит ПОЛНОЕ название суда ('Медеуский районный
    суд города Алматы ...'), а не область — регион вытаскиваем из текста
    ('...Актюбинской области', '...города Алматы')."""
    norm = _norm_option_text(court_name)
    if not norm:
        return None
    for needles, region_text in _COURT_NAME_REGION_HINTS:
        if any(n in norm for n in needles):
            return region_text
    return None


def _region_from_address(address: str) -> str | None:
    """Регион портала по адресу регистрации ответчика (person-livePlace),
    напр. 'АЛМАТИНСКАЯ ОБЛАСТЬ, ЖАМБЫЛСКИЙ РАЙОН, ...' -> 'Алматинская
    область'; 'АЛМАТЫ, БОСТАНДЫКСКИЙ, ...' -> 'город Алматы'."""
    a = _norm_option_text(address)
    if not a:
        return None
    if "астана" in a or "нур-султан" in a or "нур султан" in a:
        return "город Астана"
    if re.search(r"\bалматы\b", a) and "алматинск" not in a:
        return "город Алматы"
    if re.search(r"\bшымкент\b", a) and "шымкентск" not in a and "туркестан" not in a:
        return "город Шымкент"
    return _infer_region_text(a)


def _court_by_address(address: str, courts: list[tuple[str, str]]) -> tuple[str, str] | None:
    """Выбирает РАЙОННЫЙ суд по адресу регистрации ответчика.
    'ЖАМБЫЛСКИЙ РАЙОН' -> 'Жамбылский районный суд ...области'."""
    a = _norm_option_text(address)
    if not a or not courts:
        return None
    parts = [p.strip() for p in re.split(r"[,;]", a) if p.strip()]
    stems: list[str] = []
    for p in parts:
        if "област" in p:
            continue
        # берём только токены-районы (…ский / …цкий / со словом 'район')
        if not (re.search(r"(ский|ская|ского|ском|цкий|цкая)\b", p) or "район" in p):
            continue
        s = re.sub(r"\b(район|города|город|село|поселок|аул|пос|г\.|с\.)\b", " ", p)
        s = re.sub(r"(ский|ская|ского|ском|скому|скую)\b", "", s)
        w = (s.split() or [""])[0]
        if len(w) >= 4:
            stems.append(w[:9])
    for want_rayon in (True, False):
        for w in stems:
            for v, t in courts:
                nt = _norm_option_text(t)
                if w in nt and (not want_rayon or "районн" in nt):
                    return (v, t)
    return None


def _resolve_case_court(
    client: SudHttpClient,
    live_soup: BeautifulSoup,
    context: dict[str, str],
) -> None:
    """
    Инкрементально определяет район/суд текущего должника из Excel и
    кэширует на client. Колонка Q = «Судебный орган с Судебного кабинета»
    (название суда), колонка P = «Регион с Судебного кабинета» (область).
    Район определяется сразу, суд — когда портал (после выбора района)
    отдаёт непустой список судов в live_soup, либо в _cascade_district_court.
    """
    row = context.get("ROW", "?")
    court_name = str(context.get("COURT") or "").strip()
    # В колонке Q бывает не суд, а пометка оператору.
    if court_name and re.search(
        r"необходимо|вручную|несколько судов|уточн", court_name, re.I,
    ):
        log.warning(
            "Строка %s: в колонке Q не суд, а пометка: %r — попробуем "
            "определить суд по адресу должника.", row, court_name,
        )
        court_name = ""
    region_name = str(context.get("REGION") or "").strip()
    client.current_court_name = court_name
    client.current_region_name = region_name
    client.current_row = str(row)
    if not (court_name or region_name):
        return

    if not getattr(client, "current_region_value", None):
        region_text = (
            region_name
            or _infer_region_text(court_name)
            or _region_from_address(str(getattr(client, "defendant_livePlace", "")))
        )
        if region_text:
            rv, rc = _select_value_from_live_dom(
                live_soup, "edit-district", region_text, partial_match=False,
            )
            if rv:
                client.current_region_value = rv
                client.current_region_caption = rc
                log.info(
                    "Строка %s: область '%s' -> %s (%s)",
                    row, region_text, rv, rc,
                )

    if getattr(client, "current_court_value", None):
        return

    court_sel = _find_select_by_suffix(live_soup, ":edit-court")
    court_opts = [
        o for o in (court_sel.find_all("option") if court_sel is not None else [])
        if str(o.get("value") or "").strip()
    ]

    cv, cc = _select_value_from_live_dom(
        live_soup, "edit-court", court_name, partial_match=True,
    )
    if not cv:
        core = re.sub(r"\s*\(.*?\)\s*$", "", court_name).strip()
        if core and core != court_name:
            cv, cc = _select_value_from_live_dom(
                live_soup, "edit-court", core, partial_match=True,
            )
    if cv:
        client.current_court_value = cv
        client.current_court_caption = cc
        log.info("Строка %s: суд -> %s (%s)", row, cv, cc)
        return

    # Район выбран, портал отдал НЕПУСТОЙ список судов, но нашего там нет —
    # значит сопоставить не удалось. Нельзя оставлять район ≠ суд (портал
    # ругнётся «Поле обязательно для заполнения» на «Далее»). Откатываем
    # ОБЕ величины к дефолту шаблона (Астана), строку помечаем для ручной
    # проверки.
    if (
        court_name
        and getattr(client, "current_region_value", None)
        and court_opts
        and not getattr(client, "_court_giveup", False)
    ):
        log.warning(
            "Строка %s: суд '%s' НЕ найден в списке области '%s' (%s опций) — "
            "откат к суду по умолчанию (Астана). Проверьте вручную.",
            row, court_name, client.current_region_caption, len(court_opts),
        )
        DEFAULT_COURT_ROWS.append((str(row), court_name))
        client._court_giveup = True
        client.current_region_value = _HAR_DEFAULT_DISTRICT
        client.current_region_caption = "город Астана (по умолчанию)"
        client.current_court_value = _HAR_DEFAULT_COURT
        client.current_court_caption = "(суд по умолчанию из шаблона)"


def _patch_dynamic_case_fields(
    client: SudHttpClient,
    form_data: list[tuple[str, str]],
    live_soup: BeautifulSoup,
    context: dict[str, str],
) -> list[tuple[str, str]]:
    """
    Подставляет:
      - текст текущего DOCX в оба поля электронного бланка;
      - область из Excel P;
      - судебный орган из Excel Q.

    Область/суд определяются инкрементально по названию суда из Excel P
    и кэшируются на client.
    """
    claim_text = str(context.get("CLAIM_TEXT") or "")

    _resolve_case_court(client, live_soup, context)

    region_value = getattr(client, "current_region_value", None)
    region_caption = getattr(client, "current_region_caption", None)
    court_value = getattr(client, "current_court_value", None)
    court_caption = getattr(client, "current_court_caption", None)

    patched = []
    changed = []

    for key, value in form_data:
        low = key.lower()

        if low.endswith("edit-plaint-description") or low.endswith("edit-plaint-additional"):
            value = claim_text
            changed.append(f"{key}=CLAIM_TEXT({len(claim_text)} символов)")

        elif low.endswith("edit-district") or low.endswith("edit-district-hide"):
            if region_value:
                value = region_value
                changed.append(f"{key}={region_value} ({region_caption})")

        elif low.endswith("edit-court") or low.endswith("edit-court-hide"):
            if court_value:
                value = court_value
                changed.append(f"{key}={court_value} ({court_caption})")

        patched.append((key, value))

    # Медиация: портал требует явный отказ ('Доверяю суду') — по умолчанию
    # выбрано 'Не знаю', которое портал не принимает и блокирует «Далее».
    med_name = _find_mediation_radio_name(live_soup)
    if med_name and any(k == med_name for k, _ in patched):
        patched = [(k, v) for k, v in patched if k != med_name]
        patched.append((med_name, "TRUST_THE_COURT"))
        changed.append(f"{med_name}=TRUST_THE_COURT (медиация)")

    if changed:
        log.info("Динамические поля текущей строки: %s", " | ".join(changed))

    return patched



def _normalized_label_text(value: str) -> str:
    return re.sub(
        r"\s+",
        " ",
        str(value or "").replace("ё", "е").casefold(),
    ).strip()


def _find_input_name_by_nearby_text(
    soup: BeautifulSoup,
    phrases: tuple[str, ...],
) -> str | None:
    """
    Находит name поля по его видимой русской подписи.

    Это надёжнее жёсткого JSF id: префиксы j_idtXX могут меняться
    после каждого обновления портала.
    """
    normalized_phrases = tuple(
        _normalized_label_text(phrase)
        for phrase in phrases
    )

    # Самый точный вариант: label[for].
    for label in soup.find_all("label"):
        label_text = _normalized_label_text(
            label.get_text(" ", strip=True)
        )
        if not any(phrase in label_text for phrase in normalized_phrases):
            continue

        target_id = str(label.get("for") or "").strip()
        if target_id:
            target = soup.find(id=target_id)
            if target is not None:
                name = str(target.get("name") or target.get("id") or "").strip()
                if name:
                    return name

        nested = label.find(["input", "textarea"])
        if nested is not None:
            name = str(
                nested.get("name") or nested.get("id") or ""
            ).strip()
            if name:
                return name

    # Резерв: ищем поле внутри небольшого контейнера с нужной подписью.
    for node in soup.find_all(["input", "textarea"]):
        input_type = str(node.get("type") or "text").casefold()
        if input_type in {
            "hidden", "submit", "button", "file",
            "checkbox", "radio", "image", "reset",
        }:
            continue

        name = str(node.get("name") or node.get("id") or "").strip()
        if not name:
            continue

        current = node
        for _ in range(4):
            current = current.parent
            if current is None:
                break

            text = _normalized_label_text(
                current.get_text(" ", strip=True)
            )
            if len(text) > 500:
                continue

            if any(phrase in text for phrase in normalized_phrases):
                return name

    return None



def _patch_payment_amount_fields(
    form_data: list[tuple[str, str]],
    live_soup: BeautifulSoup,
    context: dict[str, str],
) -> list[tuple[str, str]]:
    """
    Подставляет суммы из Excel в точные JSF-поля страницы оплаты.

    По фактическому логу портала:
    - edit-totalSum = Сумма иска
    - edit-duty     = Государственная пошлина

    JSF-префикс j_idtXX может меняться, поэтому ищем по окончанию имени.
    """
    claim_name = None
    duty_name = None

    for key, _ in form_data:
        key_low = key.casefold()

        if key_low.endswith(":edit-totalsum"):
            claim_name = key

        elif key_low.endswith(":edit-duty"):
            duty_name = key

    if not claim_name or not duty_name:
        raise RuntimeError(
            "Не найдены точные поля оплаты: "
            f"edit-totalSum={claim_name!r}, edit-duty={duty_name!r}. "
            f"Поля формы={[key for key, _ in form_data]}"
        )

    patched: list[tuple[str, str]] = []

    for key, value in form_data:
        if key == claim_name:
            value = context["CLAIM_SUM"]
        elif key == duty_name:
            value = context["DUTY_SUM"]

        patched.append((key, value))

    log.info(
        "СУММЫ ИЗ EXCEL ПОДСТАВЛЕНЫ: "
        "%s=%s | %s=%s",
        claim_name,
        context["CLAIM_SUM"],
        duty_name,
        context["DUTY_SUM"],
    )

    return patched



def execute_captured_request(
    client: SudHttpClient,
    req: CapturedRequest,
    context: dict[str, str],
    debug_har_index: int | None = None,
) -> requests.Response | None:
    method = req.method.upper()

    # Служебные ресурсы и RichFaces push не воспроизводим.
    if is_service_request(req):
        return None

    # Multipart нельзя отправлять через parse_qsl: тело capture не содержит байты файла.
    if is_multipart_request(req):
        raise RuntimeError(
            "Обнаружен multipart-запрос загрузки, но он пытается выполняться как обычная форма. "
            f"Файл перехвата: {req.source_file}. "
            "Для этого шага требуется отдельный client.upload(...) с реальным файлом."
        )

    headers = {
        k: replace_known_captured_values(v, context)
        for k, v in req.headers.items()
        if k.lower() not in DROP_HEADERS
    }
    captured_referer = headers.get("Referer") or headers.get("referer") or ""
    if "/form/requestType2/" in req.url and "/form/requestType2/" in client.state.url:
        # HAR содержит старый параметр ?i=... от записанного черновика.
        # Все AJAX-запросы должны ссылаться на текущий открытый черновик.
        referer = client.state.url
    else:
        referer = captured_referer or client.state.url

    ajax = (
        req.request_type.lower() in {"xhr", "fetch"}
        or headers.get("Faces-Request") == "partial/ajax"
        or headers.get("X-Requested-With") == "XMLHttpRequest"
    )

    if method == "GET":
        response = client.get(
            replace_known_captured_values(req.url, context),
            referer=replace_known_captured_values(referer, context),
        )
    elif method == "POST":
        form_data = [
            (
                replace_known_captured_values(k, context),
                replace_known_captured_values(v, context),
            )
            for k, v in parse_post_data(req.post_data)
        ]

        # Подставляем значения, которые сервер только что вернул в AJAX DOM
        # (ФИО, адрес, реквизиты после поиска по ИИН/БИН). Это предотвращает
        # отправку персональных значений из тестового HAR.
        live_soup = BeautifulSoup(client.state.html or "", "lxml")
        live_values: dict[str, str] = {}
        for node in live_soup.find_all(["input", "select", "textarea"]):
            name = str(node.get("name") or "").strip()
            if not name or node.has_attr("disabled"):
                continue
            if node.name == "select":
                selected = node.find("option", selected=True)
                if selected is not None:
                    live_values[name] = str(selected.get("value") or "")
            elif node.name == "textarea":
                live_values[name] = node.get_text() or ""
            else:
                input_type = str(node.get("type") or "text").lower()
                if input_type not in {"button", "submit", "file"}:
                    live_values[name] = str(node.get("value") or "")

        # Данные конкретной строки Excel и текущего DOCX всегда имеют
        # приоритет над персональными значениями записанного HAR.
        form_data = _patch_dynamic_case_fields(
            client,
            form_data,
            live_soup,
            context,
        )

        # Характер заявления всегда = имущественный.
        # В исходном HAR записано значение 4 («прочее»), поэтому
        # принудительно заменяем его во всех POST на createRequest.xhtml.
        if "/createrequest.xhtml" in req.url.lower():
            patched_character = []
            for key, value in form_data:
                key_low = key.lower()
                if key_low.endswith("edit-character"):
                    value = "1"
                patched_character.append((key, value))
            form_data = patched_character

        # Для createRequest.xhtml нельзя безусловно подменять значения
        # из HAR значениями текущего DOM — это ломает сохранение участников
        # (скрытые id, ViewState и т.п.). Но ФИО/адрес/название организации
        # обязаны браться из свежего ответа сервера на поиск по ИИН/БИН —
        # иначе в иск уходят тестовые Гаппуров/Пернешева/«А-Омега» из HAR.
        if "/createRequest.xhtml" not in req.url:
            form_data = [
                (key, live_values.get(key, value))
                for key, value in form_data
            ]
        else:
            # save-шаг (клик "Сохранить") отличается от lookup-шага структурным
            # признаком протокола RichFaces/JSF, а не содержимым записанных полей —
            # это не зависит от того, есть ли в HAR реальные персональные данные.
            is_save_step = "javax.faces.partial.event=click" in (req.post_data or "")
            patched_live_lookup = []
            for key, value in form_data:
                key_low = key.lower()
                if is_save_step and any(key_low.endswith(suf) for suf in LIVE_LOOKUP_FIELD_SUFFIXES):
                    # На "save"-шаге подставляем актуальный результат поиска по
                    # ИИН/БИН текущего лица, а не значение из записанного HAR.
                    if key in live_values:
                        value = live_values[key]
                patched_live_lookup.append((key, value))
            form_data = patched_live_lookup

            if is_save_step:
                form_data = _mirror_org_fact_address(form_data)

                raw_post_data_probe = req.post_data or ""
                if any(k.lower().endswith("org-bin") for k, _ in form_data):
                    participant_role = "Истец (организация)"
                elif CAPTURED_REP_IIN in raw_post_data_probe:
                    participant_role = "Представитель"
                elif CAPTURED_DEFENDANT_IIN in raw_post_data_probe:
                    participant_role = "Ответчик"
                else:
                    participant_role = "участник"
                _require_live_lookup_fields(form_data, role=participant_role)

        # Важно: live DOM содержит суммы предыдущего/тестового HAR.
        # Поэтому переменные суммы из Excel подставляются ПОСЛЕ live DOM merge.
        if "/payment.xhtml" in req.url.casefold():
            form_data = _patch_payment_amount_fields(
                form_data,
                live_soup,
                context,
            )

        # Персональные hidden-поля нельзя брать из HAR:
        # в HAR сохранены имена файлов Ахметовой.
        # Для обязательных файлов всегда принудительно используем файлы
        # текущей строки Excel, независимо от состояния последнего upload.
        current_claim_name = str(
            context.get("CLAIM_FILENAME") or ""
        ).strip() or None

        current_duty_name = str(
            context.get("DUTY_FILENAME") or ""
        ).strip() or None

        last_uploaded_file = getattr(client, "last_uploaded_file", None)
        last_uploaded_kind = getattr(client, "last_uploaded_kind", None)

        patched_form_data = []
        for key, value in form_data:
            key_low = key.lower()

            if key_low.endswith("lawsuitscanhid") and current_claim_name:
                value = current_claim_name

            elif key_low.endswith("paymentscanhid") and current_duty_name:
                value = current_duty_name

            elif (
                last_uploaded_file
                and last_uploaded_kind
                and last_uploaded_kind == "NEXT_ATTACHMENT"
                and (
                    key_low.endswith("selectfilehid")
                    or key_low.endswith("attachfilehid")
                )
            ):
                value = last_uploaded_file

            patched_form_data.append((key, value))

        form_data = patched_form_data

        # Защита: если в обязательном hidden-поле осталось имя из HAR,
        # останавливаем выполнение до отправки запроса.
        for key, value in form_data:
            key_low = key.lower()
            value_text = str(value or "")

            if key_low.endswith("lawsuitscanhid"):
                if not current_claim_name:
                    raise RuntimeError(
                        "В context отсутствует CLAIM_FILENAME текущего должника."
                    )
                if value_text != current_claim_name:
                    raise RuntimeError(
                        "Некорректное имя файла в lawsuitScanHid: "
                        f"{value_text!r}; ожидалось {current_claim_name!r}"
                    )

            if key_low.endswith("paymentscanhid"):
                if not current_duty_name:
                    raise RuntimeError(
                        "В context отсутствует DUTY_FILENAME текущего должника."
                    )
                if value_text != current_duty_name:
                    raise RuntimeError(
                        "Некорректное имя файла в paymentScanHid: "
                        f"{value_text!r}; ожидалось {current_duty_name!r}"
                    )

        hidden_debug = [
            (key, value)
            for key, value in form_data
            if key.lower().endswith(("lawsuitscanhid", "paymentscanhid"))
        ]
        if hidden_debug:
            log.info("Обязательные hidden-файлы текущего дела: %s", hidden_debug)

        # Значения динамической строки.
        # У шага ответчика и шага представителя одно и то же имя поля
        # (j_idt272:person-iin) — различаем их по исходному ИИН, который
        # был записан в этом конкретном запросе HAR, иначе оба шага
        # получат один и тот же (ответчика) ИИН.
        raw_post_data = req.post_data or ""
        is_rep_step = CAPTURED_REP_IIN in raw_post_data
        is_defendant_step = CAPTURED_DEFENDANT_IIN in raw_post_data
        for idx, (key, value) in enumerate(form_data):
            key_low = key.lower()
            if key_low.endswith("person-iin"):
                if is_rep_step and context["REP_IIN"]:
                    form_data[idx] = (key, context["REP_IIN"])
                elif is_defendant_step and context["IIN"]:
                    form_data[idx] = (key, context["IIN"])
            elif "claim" in key_low and "sum" in key_low:
                form_data[idx] = (key, context["CLAIM_SUM"])
            elif ("duty" in key_low or "statefee" in key_low) and "sum" in key_low:
                form_data[idx] = (key, context["DUTY_SUM"])

        if (
            debug_har_index is not None
            and 0 <= debug_har_index <= 17
            and "/createRequest.xhtml" in req.url
        ):
            original_form_data = parse_post_data(req.post_data)
            original_map = {}
            final_map = {}
            for key, field_value in original_form_data:
                original_map.setdefault(key, []).append(field_value)
            for key, field_value in form_data:
                final_map.setdefault(key, []).append(field_value)

            changed_fields = {}
            for key in sorted(set(original_map) | set(final_map)):
                if original_map.get(key) != final_map.get(key):
                    changed_fields[key] = {
                        "captured": original_map.get(key),
                        "sent": final_map.get(key),
                    }

            _write_participant_debug(
                debug_har_index,
                "request",
                {
                    "har_index": debug_har_index,
                    "source_file": req.source_file,
                    "method": method,
                    "captured_url": req.url,
                    "sent_url": replace_known_captured_values(req.url, context),
                    "referer": replace_known_captured_values(referer, context),
                    "ajax": ajax,
                    "client_state_url": client.state.url,
                    "client_viewstate": client.state.viewstate,
                    "headers": headers,
                    "captured_fields": original_form_data,
                    "sent_fields": form_data,
                    "changed_fields": changed_fields,
                    "live_dom_field_count": len(live_values),
                    "live_dom_fields": live_values,
                },
            )

        hidden_upload_values = [
            (key, value)
            for key, value in form_data
            if key.lower().endswith(("scanhid", "filehid"))
        ]
        if hidden_upload_values:
            log.info(
                "Скрытые поля файла перед POST: %s",
                hidden_upload_values,
            )

        response = client.post_form(
            replace_known_captured_values(req.url, context),
            form_data,
            referer=replace_known_captured_values(referer, context),
            ajax=ajax,
        )

        # Шаг поиска ответчика по ИИН: портал возвращает адрес регистрации
        # (person-livePlace). По нему потом выбираем РАЙОННЫЙ суд, когда в
        # Excel указана только область (компания ORION).
        if is_defendant_step:
            _liv = re.search(
                r'person-livePlace"[^>]*value="([^"]*)"', response.text or "",
            )
            if _liv and _liv.group(1).strip():
                client.defendant_livePlace = _liv.group(1).strip()
                log.info(
                    "Адрес ответчика (из поиска по ИИН): %s",
                    client.defendant_livePlace,
                )

        if (
            debug_har_index is not None
            and 0 <= debug_har_index <= 17
            and "/createRequest.xhtml" in req.url
        ):
            updates = {}
            new_viewstate = None
            redirect_url = None
            parse_error = None
            try:
                updates, new_viewstate, redirect_url = parse_partial_response(response.text)
            except Exception as exc:
                parse_error = repr(exc)

            response_headers = dict(response.headers)
            body_text = response.text or ""
            _write_participant_debug(
                debug_har_index,
                "response",
                {
                    "har_index": debug_har_index,
                    "status_code": response.status_code,
                    "response_url": response.url,
                    "response_headers": response_headers,
                    "client_state_url_after": client.state.url,
                    "client_viewstate_after": client.state.viewstate,
                    "partial_update_ids": list(updates),
                    "new_viewstate_from_response": new_viewstate,
                    "redirect_url": redirect_url,
                    "parse_error": parse_error,
                    "response_length": len(body_text),
                    "response_preview": body_text[:5000],
                    "contains_exception": any(
                        token in body_text.lower()
                        for token in ("exception", "error", "validation", "facesmessage")
                    ),
                },
            )

            html_path = PARTICIPANTS_DEBUG_DIR / f"HAR_{debug_har_index:03d}_state_after.html"
            try:
                html_path.write_text(client.state.html or "", encoding="utf-8")
            except Exception:
                log.exception("Не удалось сохранить HTML после HAR-%03d", debug_har_index)
    else:
        log.info("Пропущен метод %s: %s", method, req.url)
        return None

    time.sleep(PAUSE_BETWEEN_REQUESTS)
    return response



def _header_value(headers: dict[str, str], name: str) -> str:
    for key, value in headers.items():
        if key.lower() == name.lower():
            return str(value or "")
    return ""


def bootstrap_from_har(
    client: SudHttpClient,
    captures: list[CapturedRequest],
) -> str:
    """
    Открывает исходную страницу, с которой начинается HAR-000.

    В текущем HAR первый бизнес-запрос уже относится к заполнению
    createRequest.xhtml, поэтому перед его воспроизведением нужен GET страницы
    черновика и актуальный ViewState.

    Сначала используется полный Referer HAR-000. Если его токен ?i=... больше
    недействителен, функция выдаёт отдельную понятную ошибку: тогда нужен новый
    HAR, записанный с момента нажатия на услугу/создания нового иска.
    """
    actions = [
        req for req in captures
        if req.is_http_action and not is_service_request(req)
    ]
    if not actions:
        raise RuntimeError("В HAR нет бизнес-запросов для запуска.")

    first = actions[0]
    first_referer = _header_value(first.headers, "Referer")
    if not first_referer:
        raise RuntimeError(
            f"У первого HAR-запроса {first.source_file} отсутствует Referer. "
            "Невозможно определить стартовую страницу createRequest."
        )

    if "/form/requestType2/createRequest.xhtml" not in first_referer:
        raise RuntimeError(
            "Первый HAR-запрос начинается не со страницы createRequest.xhtml: "
            f"{first_referer}"
        )

    log.info("HAR bootstrap: открываем стартовую страницу %s", first_referer)

    try:
        response = client.get(first_referer, referer=client.state.url)
    except requests.HTTPError as exc:
        status = getattr(exc.response, "status_code", None)
        raise RuntimeError(
            "Не удалось открыть стартовую страницу черновика из HAR. "
            f"HTTP={status}; URL={first_referer}. "
            "Токен ?i=... из HAR, вероятно, устарел. "
            "Нужен новый HAR, записанный начиная с клика по созданию нового иска, "
            "а не только с уже открытой страницы createRequest."
        ) from exc

    final_url = response.url or client.state.url
    soup = BeautifulSoup(response.text or "", "lxml")
    forms = soup.find_all("form")
    has_create_form = any(
        "/form/requestType2/createRequest.xhtml" in str(form.get("action") or "")
        for form in forms
    )

    if "/form/requesttype2/createrequest.xhtml" not in final_url.lower():
        raise RuntimeError(
            "HAR bootstrap не открыл createRequest.xhtml. "
            f"Получен URL: {final_url}. "
            "Сервер мог перенаправить устаревший черновик на другую страницу."
        )

    if not client.state.viewstate:
        raise RuntimeError(
            "Страница createRequest.xhtml открылась, но ViewState не найден."
        )

    if not has_create_form:
        available_actions = [
            str(form.get("action") or "")
            for form in forms[:20]
        ]
        raise RuntimeError(
            "На стартовой странице не найдена форма createRequest.xhtml. "
            f"Найденные action: {available_actions}"
        )

    _remember_request_token(client, final_url)
    log.info(
        "HAR bootstrap готов | URL=%s | ViewState=есть | forms=%s",
        final_url,
        len(forms),
    )
    return final_url


def run_capture_sequence(
    client: SudHttpClient,
    captures: list[CapturedRequest],
    context: dict[str, str],
    case_files: dict[str, Path],
) -> None:
    """
    Выполняет HAR-сценарий, но количество обычных приложений берёт
    из фактической папки должника.

    Если в HAR записано 14 обычных upload, а в папке найдено 15 файлов,
    дополнительный файл загружается повторным использованием последнего
    рабочего RichFaces upload-запроса до перехода на страницу подписания.
    """
    actions = [
        r for r in captures
        if r.is_http_action and not is_service_request(r)
    ]

    log.info("HTTP-действий после фильтрации: %s", len(actions))

    last_generic_upload_req: CapturedRequest | None = None

    for number, req in enumerate(actions, start=1):
        # После последнего записанного обычного upload HAR переходит к
        # завершающему POST. До этого перехода догружаем остаток очереди.
        if (
            not is_multipart_request(req)
            and last_generic_upload_req is not None
        ):
            queue = list(
                getattr(client, "current_attachment_queue", [])
            )
            queue_index = int(
                getattr(client, "current_attachment_index", 0)
            )

            while queue_index < len(queue):
                log.info(
                    "В HAR закончились upload-шаги, но в папке остались "
                    "приложения: %s. Выполняю дополнительный upload %s/%s.",
                    len(queue) - queue_index,
                    queue_index + 1,
                    len(queue),
                )

                execute_multipart_upload(
                    client,
                    last_generic_upload_req,
                    context,
                    case_files,
                )

                queue_index = int(
                    getattr(client, "current_attachment_index", 0)
                )

            # Дренирование выполняется только один раз перед финальными POST.
            last_generic_upload_req = None

        log.info(
            "[%s/%s] %s %s | %s",
            number, len(actions), req.method, req.url, req.source_file,
        )

        if is_multipart_request(req):
            rule = find_upload_rule(req)

            if rule.file_key == "NEXT_ATTACHMENT":
                last_generic_upload_req = req

            response = execute_multipart_upload(
                client,
                req,
                context,
                case_files,
            )
        else:
            response = execute_captured_request(
                client,
                req,
                context,
                debug_har_index=number - 1,
            )

        if response is None:
            continue

        if (
            is_sign_page(client.state.url, client.state.html)
            or is_sign_page(response.url, response.text)
        ):
            log.info(
                "Достигнута страница подписания: %s",
                client.state.url,
            )
            if STOP_AT_SIGN_PAGE:
                return

    if STOP_AT_SIGN_PAGE and not is_sign_page(
        client.state.url,
        client.state.html,
    ):
        raise RuntimeError(
            "Последовательность завершилась, но страница подписания "
            "не достигнута. "
            f"Последний URL: {client.state.url}"
        )


# %% [markdown]
# ## 6. Загрузки файлов
# 
# Multipart-загрузки квитанции, приложения и иска встроены непосредственно в HTTP-цепочку. Имя поля файла равно ID компонента RichFaces `fileUpload`; актуальный `ViewState` подставляется автоматически.
# 

# %%
@dataclass
class UploadRule:
    name: str
    file_key: str
    component_id: str


def find_upload_rule(req: CapturedRequest) -> UploadRule:
    query = dict(parse_qsl(urlsplit(req.url).query, keep_blank_values=True))
    component_id = str(
        query.get("javax.faces.source")
        or query.get("org.richfaces.ajax.component")
        or ""
    )

    if "personTableRows" in component_id:
        return UploadRule(
            name="Квитанция госпошлины",
            file_key="DUTY_PDF",
            component_id=component_id,
        )

    if "selectLawsuitScanUploader" in component_id:
        return UploadRule(
            name="Иск",
            file_key="CLAIM_DOCX",
            component_id=component_id,
        )

    if "selectFileUploader" in component_id:
        return UploadRule(
            name="Следующее приложение из папки должника",
            file_key="NEXT_ATTACHMENT",
            component_id=component_id,
        )

    raise RuntimeError(
        f"Не распознан RichFaces upload-компонент: {component_id or req.url}"
    )
def _remember_request_token(client: SudHttpClient, url: str) -> str | None:
    """Сохраняет параметр i текущего черновика до того, как AJAX POST уберёт его из URL."""
    query = dict(parse_qsl(urlsplit(url or "").query, keep_blank_values=True))
    token = str(query.get("i") or "").strip()
    if token:
        client.current_request_token = token
        client.current_request_url = url
        log.info("Сохранён токен текущего черновика: %s", token)
        return token
    return None


def _current_request_token(client: SudHttpClient) -> str:
    """
    Возвращает параметр i текущего черновика.

    После первого POST на createRequest.xhtml библиотека requests получает URL
    без query-параметра i, поэтому токен необходимо брать из сохранённого
    значения, а не только из client.state.url.
    """
    token = _remember_request_token(client, client.state.url or "")
    if token:
        return token

    saved = str(getattr(client, "current_request_token", "") or "").strip()
    if saved:
        return saved

    saved_url = str(getattr(client, "current_request_url", "") or "").strip()
    if saved_url:
        query = dict(parse_qsl(urlsplit(saved_url).query, keep_blank_values=True))
        token = str(query.get("i") or "").strip()
        if token:
            client.current_request_token = token
            return token

    raise RuntimeError(
        "Не найден параметр i текущего черновика. "
        f"Текущий URL={client.state.url}; "
        f"сохранённый URL={saved_url or 'нет'}"
    )

_MEDIATION_VALUES = {"DONT_KNOW", "NOT_TRUST", "TRUST_THE_COURT"}


def _find_mediation_radio_name(soup_or_html) -> str | None:
    """Имя группы радио-кнопок «отказ от медиации» (значения DONT_KNOW /
    NOT_TRUST / TRUST_THE_COURT). id динамический, ищем по значениям —
    сначала по DOM, потом регуляркой по сырому HTML (модалка может быть
    только в partial-ответе)."""
    try:
        for inp in soup_or_html.find_all("input", attrs={"type": "radio"}):
            if str(inp.get("value") or "") in _MEDIATION_VALUES:
                name = str(inp.get("name") or "").strip()
                if name:
                    return name
        raw = str(soup_or_html)
    except AttributeError:
        raw = str(soup_or_html or "")
    m = re.search(
        r'name="([^"]+)"[^>]*value="(?:DONT_KNOW|NOT_TRUST|TRUST_THE_COURT)"',
        raw,
    ) or re.search(
        r'value="(?:DONT_KNOW|NOT_TRUST|TRUST_THE_COURT)"[^>]*name="([^"]+)"',
        raw,
    )
    return m.group(1) if m else None


def _find_mediation_dialog_button(raw: str) -> str | None:
    """id/name кнопки «Далее» в модалке медиации
    (<div class="modal-footer">...<input ... value="Далее" ...>)."""
    m = re.search(
        r'modal-footer.*?<input[^>]*\bid="([^"]+)"[^>]*value="Далее"',
        raw, re.S,
    ) or re.search(
        r'<input[^>]*value="Далее"[^>]*RichFaces\.ajax\(\s*["\']([^"\']+)["\']',
        raw,
    )
    return m.group(1) if m else None


def _force_mediation_answer(
    soup,
    fields: list[tuple[str, str]],
    answer: str = "TRUST_THE_COURT",
) -> list[tuple[str, str]]:
    """Портал требует явно ответить на вопрос про медиацию перед «Далее».
    По умолчанию выбрана 'DONT_KNOW', которую портал не принимает как
    отказ — принудительно ставим 'Доверяю суду' (+ *-hide на всякий)."""
    name = _find_mediation_radio_name(soup)
    if not name:
        return fields
    drop = {name, name + "-hide"}
    out = [(k, v) for k, v in fields if k not in drop]
    out.append((name, answer))
    return out


def _serialize_jsf_form(form) -> list[tuple[str, str]]:
    """Собирает успешные поля HTML-формы примерно так же, как браузер."""
    result: list[tuple[str, str]] = []

    for node in form.find_all(["input", "select", "textarea"]):
        name = str(node.get("name") or "").strip()
        if not name or node.has_attr("disabled"):
            continue

        tag = node.name.lower()

        if tag == "input":
            input_type = str(node.get("type") or "text").lower()
            if input_type in {"submit", "button", "file", "reset", "image"}:
                continue
            if input_type in {"checkbox", "radio"} and not node.has_attr("checked"):
                continue
            value = str(node.get("value") or "")
            result.append((name, value))

        elif tag == "select":
            selected = node.find_all("option", selected=True)
            if not selected:
                first = node.find("option")
                selected = [first] if first is not None else []
            for option in selected:
                result.append((name, str(option.get("value") or "")))

        elif tag == "textarea":
            result.append((name, node.get_text() or ""))

    return result


def _confirm_mediation_dialog(
    client: SudHttpClient,
    response: requests.Response,
):
    """Портал при нажатии goNext показывает модалку «Вы уверены, что не
    хотите Медиацию?». По умолчанию отмечено 'Не знаю', которое портал НЕ
    принимает как отказ. Выбираем 'Доверяю суду' и жмём СОБСТВЕННУЮ кнопку
    «Далее» модалки (id вида j_idt552:j_idt562) — именно её нажатие
    переводит на страницу оплаты. Возвращает ответ портала или None.

    Раньше здесь искалась любая кнопка с текстом «Далее» — и находилась
    основная кнопка goNext формы (она в DOM раньше модалки), из-за чего
    goNext перезапускался по кругу и модалка всплывала снова."""
    raw = response.text or ""
    frag_m = re.search(
        r'<update id="[^"]*mediationAsnwerDialogPanel">\s*<!\[CDATA\[(.*?)\]\]>'
        r'</update>',
        raw, re.S | re.I,
    )
    frag = frag_m.group(1) if frag_m else raw

    med_name = _find_mediation_radio_name(frag)
    btn_id = _find_mediation_dialog_button(frag)
    if not (med_name and btn_id):
        full_html = client.state.html or ""
        med_name = med_name or _find_mediation_radio_name(
            BeautifulSoup(full_html, "lxml"))
        btn_id = btn_id or _find_mediation_dialog_button(full_html)
    if not (med_name and btn_id):
        log.warning(
            "Диалог медиации: не распознаны радио=%r / кнопка «Далее»=%r.",
            med_name, btn_id,
        )
        return None

    full = BeautifulSoup(client.state.html or "", "lxml")
    anchor = (
        full.find(id=btn_id)
        or full.find(attrs={"name": btn_id})
        or full.find(attrs={"name": med_name})
    )
    form = anchor.find_parent("form") if anchor is not None else full.find("form")
    if form is None:
        log.warning("Диалог медиации: форма для подтверждения не найдена.")
        return None
    form_id = str(form.get("id") or form.get("name") or "").strip()

    base = [
        (k, v)
        for k, v in _serialize_jsf_form(form)
        if k not in VIEWSTATE_NAMES
        and k not in {form_id, med_name, med_name + "-hide"}
    ]
    base.insert(0, (form_id, form_id))
    base.append((med_name, "TRUST_THE_COURT"))

    def _send(source_id: str, execute: str, render: str, *, extra=()):
        # ViewState post_form подставит сам (актуальный).
        fields = list(base)
        fields.extend([
            ("javax.faces.source", source_id),
            ("javax.faces.partial.execute", execute),
            ("javax.faces.partial.render", render),
            ("org.richfaces.ajax.component", source_id),
        ])
        fields.extend(extra)
        fields.extend([
            (source_id, source_id),
            ("rfExt", "null"),
            ("AJAX:EVENTS_COUNT", "1"),
            ("javax.faces.partial.ajax", "true"),
        ])
        return client.post_form(
            "/form/requestType2/createRequest.xhtml",
            fields, referer=response.url, ajax=True,
        )

    # 1) отмечаем «Доверяю суду» (valueChange радио-группы медиации)
    try:
        _send(med_name, med_name, med_name,
              extra=[("javax.faces.behavior.event", "valueChange")])
    except Exception as exc:
        log.warning("Диалог медиации: выбор 'Доверяю суду' не прошёл: %s", exc)

    # 2) жмём кнопку «Далее» самой модалки — это и есть переход к оплате
    log.info(
        "Диалог медиации: 'Доверяю суду' -> «Далее» модалки (%s).", btn_id,
    )
    try:
        return _send(btn_id, "@form", "@all", extra=[("incId", "1")])
    except Exception as exc:
        log.warning("Диалог медиации: подтверждение «Далее» не удалось: %s", exc)
        return None


def _cascade_district_court(client: SudHttpClient, response, soup: BeautifulSoup):
    """
    Каскад «выбрать район -> портал подгружает суды -> выбрать суд» через
    настоящий RichFaces AJAX (одного лишь значения edit-district в теле
    goNext-POST недостаточно: портал не грузит список судов и ругается
    «Поле обязательно для заполнения» на edit-court).

    Успех -> client.current_court_value заполнен, форма готова к goNext.
    Неудача -> откатываем район/суд к дефолту шаблона (Астана 2/489),
              строку помечаем в DEFAULT_COURT_ROWS.
    Возвращает (response, soup) — возможно обновлённые.
    """
    if getattr(client, "current_court_value", None):
        return response, soup

    court_name = str(getattr(client, "current_court_name", "") or "").strip()
    region_name = str(getattr(client, "current_region_name", "") or "").strip()
    row = str(getattr(client, "current_row", "?"))
    addr = str(getattr(client, "defendant_livePlace", "") or "")
    if not (court_name or region_name or addr):
        return response, soup

    region_text = (
        region_name
        or _infer_region_text(court_name)
        or _region_from_address(addr)
    )
    dsel = _find_select_by_suffix(soup, ":edit-district")
    region_value = None
    if region_text and dsel is not None:
        region_value, _ = _select_value_from_live_dom(
            soup, "edit-district", region_text, partial_match=False,
        )

    chosen: dict[str, str] = {}
    try:
        if not region_value:
            raise RuntimeError(
                f"регион не сопоставлен (Excel='{court_name}', адрес='{addr}')"
            )
        response, soup = _ajax_select_value_exact_har(
            client, response, soup, ":edit-district", region_value,
            f"Область «{region_text}»", chosen,
        )
        csel = _find_select_by_suffix(soup, ":edit-court")
        courts = [
            (str(o.get("value") or ""), o.get_text(" ", strip=True))
            for o in (csel.find_all("option") if csel is not None else [])
            if str(o.get("value") or "").strip()
        ]
        # 1) точное/частичное совпадение с названием суда из Excel (компания 1),
        # 2) районный суд по адресу регистрации ответчика (компания ORION).
        hit = _match_court(court_name, courts)
        how = "Excel"
        if not hit:
            core = re.sub(r"\s*\(.*?\)\s*$", "", court_name).strip()
            if core and core != court_name:
                hit = _match_court(core, courts)
        if not hit and addr:
            hit = _court_by_address(addr, courts)
            how = "адрес ответчика"
        if not hit:
            raise RuntimeError(
                f"суд не сопоставлен (Excel='{court_name}', адрес='{addr}', "
                f"судов в области: {len(courts)})"
            )
        response, soup = _ajax_select_value_exact_har(
            client, response, soup, ":edit-court", hit[0],
            f"Суд «{hit[1]}»", chosen,
        )
        client.current_region_value = region_value
        client.current_region_caption = region_text
        client.current_court_value = hit[0]
        client.current_court_caption = hit[1]
        log.info(
            "Строка %s: каскад район→суд (%s): область=%s | суд=%s (%s)",
            row, how, region_value, hit[0], hit[1],
        )
        return response, soup
    except Exception as exc:
        _dump_case_page(client, "region_court_fail")
        log.warning(
            "Строка %s: каскад район→суд не удался (%s) — иск уйдёт в суд "
            "по умолчанию (Астана). Проверьте вручную.", row, exc,
        )
        if not any(r == row for r, _ in DEFAULT_COURT_ROWS):
            DEFAULT_COURT_ROWS.append((row, f"{court_name} | {addr}"))
        client._court_giveup = True
        client.current_region_value = _HAR_DEFAULT_DISTRICT
        client.current_region_caption = "город Астана (по умолчанию)"
        client.current_court_value = _HAR_DEFAULT_COURT
        client.current_court_caption = "(суд по умолчанию из шаблона)"
        # Прерванный каскад оставил на сервере чужой район -> возвращаем
        # согласованную пару Астана 2/489 тем же AJAX (иначе goNext молча
        # не пустит из-за район≠суд).
        try:
            response, soup = _ajax_select_value_exact_har(
                client, response, soup, ":edit-district", _HAR_DEFAULT_DISTRICT,
                "Область «город Астана» (по умолчанию)", {},
            )
            response, soup = _ajax_select_value_exact_har(
                client, response, soup, ":edit-court", _HAR_DEFAULT_COURT,
                f"Суд по умолчанию {_HAR_DEFAULT_COURT}", {},
            )
        except Exception:
            try:
                token = _current_request_token(client)
                response = client.get(
                    client.base_url + "/form/requestType2/createRequest.xhtml?"
                    + urlencode({"i": token}),
                    referer=client.state.url,
                )
                soup = BeautifulSoup(client.state.html or "", "lxml")
            except Exception:
                pass
        return response, soup


def _transition_create_request_to_payment(client: SudHttpClient) -> None:
    """
    Выполняет штатный RichFaces AJAX-переход:
    1. Заполнение данных -> 2. Оплата.

    Прямой GET payment.xhtml?i=... сервер не принимает и отвечает 500,
    поэтому сначала вызывается goNext на createRequest.xhtml.
    """
    current_path = urlsplit(client.state.url or "").path.lower()
    if current_path.endswith("/payment.xhtml"):
        return

    token = _current_request_token(client)
    create_url = urljoin(
        client.base_url + "/",
        "form/requestType2/createRequest.xhtml?" + urlencode({"i": token}),
    )

    # Получаем полный актуальный DOM createRequest, поскольку после AJAX в
    # client.state.html могут оставаться только partial-update фрагменты.
    log.info(
        "Подготовка перехода Заполнение данных -> Оплата: GET %s",
        create_url,
    )
    _resp = client.get(
        create_url,
        referer=getattr(client, "current_request_url", None) or client.state.url,
    )

    soup = BeautifulSoup(client.state.html or "", "lxml")

    # Каскад район→суд (если суд ещё не выбран).
    _resp, soup = _cascade_district_court(client, _resp, soup)

    # В HTML функция имеет вид:
    # goNext=function(){RichFaces.ajax("...:j_idt122",null,{"incId":"1"} )}
    match = re.search(
        r'goNext\s*=\s*function\s*\(\)\s*\{\s*RichFaces\.ajax\(\s*["\']([^"\']+)["\']',
        client.state.html or "",
        flags=re.I | re.S,
    )

    if not match:
        # Резервный поиск по кнопке "Далее" и её onclick.
        next_node = soup.find(
            lambda tag: tag.name in {"a", "button", "input"}
            and "далее" in (
                (tag.get_text(" ", strip=True) if tag.name != "input" else str(tag.get("value") or ""))
                .lower()
            )
        )
        onclick = str(next_node.get("onclick") or "") if next_node else ""
        match = re.search(
            r'RichFaces\.ajax\(\s*["\']([^"\']+)["\']',
            onclick,
            flags=re.I | re.S,
        )

    if not match:
        raise RuntimeError(
            "На createRequest.xhtml не найден JSF-компонент функции goNext. "
            "Нужен HTML страницы перед нажатием кнопки «Далее»."
        )

    source_id = match.group(1)
    source_node = soup.find(id=source_id) or soup.find(attrs={"name": source_id})
    form = source_node.find_parent("form") if source_node is not None else soup.find("form")

    if form is None:
        raise RuntimeError("На createRequest.xhtml не найдена форма для перехода к оплате.")

    form_id = str(form.get("id") or form.get("name") or "").strip()
    if not form_id:
        raise RuntimeError("У формы createRequest отсутствует id/name.")

    payload = _serialize_jsf_form(form)

    # Удаляем служебные JSF-поля — добавим их заново с актуальными значениями.
    drop_names = {
        "javax.faces.source",
        "javax.faces.partial.execute",
        "javax.faces.partial.render",
        "org.richfaces.ajax.component",
        "javax.faces.partial.ajax",
        "AJAX:EVENTS_COUNT",
        "rfExt",
    }
    payload = [
        (key, value)
        for key, value in payload
        if key not in drop_names and key not in VIEWSTATE_NAMES
    ]

    # На повторном GET Judicial Cabinet иногда возвращает обязательные
    # select-поля пустыми, хотя они уже были выбраны на первом шаге.
    # Восстанавливаем их аккуратно: РАЙОН и СУД зависят от региона
    # должника (их разрешил _patch_dynamic_case_fields и положил в
    # client.current_*), поэтому жёстко «489» подставлять нельзя — для
    # должника не из Алматы это неверный суд, и портал молча не пускает
    # на шаг оплаты. Приоритет: значение из формы -> значение должника ->
    # дефолт категории.
    _form_vals: dict[str, str] = {}
    for _k, _v in payload:
        if ":" in _k and _v:
            _form_vals.setdefault(_k.rsplit(":", 1)[-1], _v)

    cur_region = getattr(client, "current_region_value", None)
    cur_court = getattr(client, "current_court_value", None)

    def _first(*vals: str | None) -> str:
        for v in vals:
            if v:
                return str(v)
        return ""

    required_suffix_values = {
        "edit-categoryGroup": _first(_form_vals.get("edit-categoryGroup"),
                                     _form_vals.get("edit-categoryGroup-hide"), "2"),
        "edit-categoryGroup-hide": _first(_form_vals.get("edit-categoryGroup-hide"),
                                          _form_vals.get("edit-categoryGroup"), "2"),
        "edit-category": _first(_form_vals.get("edit-category"),
                                _form_vals.get("edit-category-hide"), "27"),
        "edit-category-hide": _first(_form_vals.get("edit-category-hide"),
                                     _form_vals.get("edit-category"), "27"),
        "edit-character": _first(_form_vals.get("edit-character"), "1"),
        "edit-district": _first(_form_vals.get("edit-district"), cur_region,
                                _form_vals.get("edit-district-hide")),
        "edit-district-hide": _first(_form_vals.get("edit-district-hide"), cur_region,
                                     _form_vals.get("edit-district")),
        "edit-court": _first(_form_vals.get("edit-court"), cur_court,
                             _form_vals.get("edit-court-hide")),
        "edit-court-hide": _first(_form_vals.get("edit-court-hide"), cur_court,
                                  _form_vals.get("edit-court")),
        "processPersonListCode": _first(_form_vals.get("processPersonListCode"), "0"),
    }
    # Никогда не затираем поле пустой строкой.
    required_suffix_values = {
        k: v for k, v in required_suffix_values.items() if v
    }

    payload_map = {}
    payload_order = []
    for key, value in payload:
        if key not in payload_map:
            payload_order.append(key)
        payload_map[key] = value

    for suffix, value in required_suffix_values.items():
        full_name = f"{form_id}:{suffix}"
        payload_map[full_name] = value
        if full_name not in payload_order:
            payload_order.append(full_name)

    payload = [(key, payload_map[key]) for key in payload_order]

    # Медиация: портал требует явный отказ ('Доверяю суду').
    payload = _force_mediation_answer(soup, payload)

    # Форма и точные параметры подтверждены HAR/cURL.
    payload = [(key, value) for key, value in payload if key != form_id]
    payload.insert(0, (form_id, form_id))

    # ViewState должен быть именно текущим и передаваться в теле POST.
    payload.append(("javax.faces.ViewState", client.state.viewstate))

    payload.extend([
        ("javax.faces.source", source_id),
        ("javax.faces.partial.execute", f"{source_id} @component"),
        ("javax.faces.partial.render", "@component"),
        ("org.richfaces.ajax.component", source_id),
        (source_id, source_id),
        ("rfExt", "null"),
        ("AJAX:EVENTS_COUNT", "1"),
        ("javax.faces.partial.ajax", "true"),
    ])

    log.info(
        "Поля перехода к оплате: categoryGroup=%s | category=%s | "
        "character=%s | district=%s | court=%s",
        payload_map.get(f"{form_id}:edit-categoryGroup-hide"),
        payload_map.get(f"{form_id}:edit-category-hide"),
        payload_map.get(f"{form_id}:edit-character"),
        payload_map.get(f"{form_id}:edit-district-hide"),
        payload_map.get(f"{form_id}:edit-court-hide"),
    )

    referer = client.state.url
    log.info(
        "AJAX переход Заполнение данных -> Оплата: form=%s | source=%s | "
        "ViewState=%s | referer=%s",
        form_id,
        source_id,
        "есть" if client.state.viewstate else "НЕТ",
        referer,
    )

    response = client.post_form(
        "/form/requestType2/createRequest.xhtml",
        payload,
        referer=referer,
        ajax=True,
    )

    # ПРИМЕЧАНИЕ: модалка `mediationAsnwerDialogPanel` присутствует в КАЖДОМ
    # ответе портала как скрытая разметка и НИКОГДА не блокирует goNext
    # (подтверждено свежей записью HAR 2026-09-02: goNext -> сразу redirect
    # на payment.xhtml без шага медиации). Раньше здесь была попытка
    # «подтвердить» медиацию — она срабатывала на ложный признак и только
    # мешала. Если goNext не прошёл — причина в другом (обычно район≠суд).
    path = urlsplit(client.state.url or "").path.lower()

    # Обычно JSF возвращает redirect, который post_form уже откроет через GET.
    # Если redirect не распознан, пробуем URL из XML/eval и только затем
    # проверяем фактическое состояние.
    if not path.endswith("/payment.xhtml"):
        updates, _, redirect_url = parse_partial_response(response.text or "")
        if not redirect_url:
            # Резерв: любой payment.xhtml-URL, встретившийся в ответе.
            m = re.search(r'(/form/requestType2/payment\.xhtml[^"\'<> ]*)',
                          response.text or "")
            if m:
                redirect_url = m.group(1)
        if redirect_url:
            client.get(urljoin(response.url, redirect_url), referer=response.url)
            path = urlsplit(client.state.url or "").path.lower()

    if not path.endswith("/payment.xhtml"):
        # Диагностика: полный ответ портала + активные сообщения формы.
        full = response.text or ""
        try:
            _dbg = Path(
                os.environ.get("OMEGA_WORKDIR")
                or os.environ.get("OMEGA_OUT")
                or "."
            ) / f"gonext_fail_{_current_request_token(client)}.xml"
            _dbg.write_text(full, encoding="utf-8")
            log.error("Полный ответ goNext сохранён: %s", _dbg)
        except Exception:
            pass
        # Достаём ЧИСТЫЙ текст сообщений валидации портала (это ответ
        # сервера, а не CSS-селекторы — просто портал так помечает поля
        # с ошибкой классами rf-msg-err / rf-msgs-err).
        msgs: list[str] = []
        for upd in re.findall(
            r'<update id="([^"]*)"><!\[CDATA\[(.*?)\]\]></update>', full, re.S
        ):
            upd_id, frag = upd
            # Модалки (медиация и пр.) присутствуют в КАЖДОМ ответе —
            # это не ошибка, пропускаем их статический текст.
            if "modaldialogpanel" in upd_id.lower() or "modal-dialog" in frag.lower():
                continue
            if "rf-msg" not in frag:
                continue
            try:
                txt = BeautifulSoup(frag, "lxml").get_text(" ", strip=True)
            except Exception:
                txt = re.sub(r"<[^>]+>", " ", frag)
            txt = re.sub(r"\s+", " ", txt).strip()
            if txt and txt not in msgs:
                msgs.append(txt)

        if msgs:
            msg_part = " Портал сообщил: " + "; ".join(msgs[:5]) + "."
        else:
            msg_part = (
                " Портал не показал ни ошибок валидации, ни перехода — "
                "вероятно, временный сбой/перегрузка портала или неполные "
                "данные участников (поиск по ИИН не вернул ФИО)."
            )
        snippet = re.sub(r"\s+", " ", full)[:1200]
        raise RuntimeError(
            "AJAX goNext выполнен, но страница оплаты не достигнута."
            f"{msg_part} URL={client.state.url}. Ответ: {snippet}"
        )

    log.info("Переход на страницу оплаты подтверждён: %s", client.state.url)

def _prepare_actual_upload_page(
    client: SudHttpClient,
    rule: UploadRule,
    req: CapturedRequest,
) -> tuple[str, str, str]:
    """
    Подготавливает фактическую JSF-страницу перед RichFaces multipart upload.

    Логика:
    - квитанция госпошлины загружается на payment.xhtml;
    - иск и обычные документы загружаются на blankData/other.xhtml;
    - перед первой загрузкой документа выполняется GET страницы документов;
    - используется актуальный ViewState текущего черновика;
    - если после AJAX в client.state.html отсутствует полная HTML-форма,
      используется известный form_id страницы документов;
    - component_id берётся из DOM либо из HAR.

    Возвращает:
        form_id, component_id, upload_url
    """

    captured_parts = urlsplit(req.url)
    target_path = captured_parts.path

    current_url = client.state.url or ""
    current_path = urlsplit(current_url).path

    # =========================================================
    # 1. Переход на нужную фактическую страницу
    # =========================================================

    if rule.file_key == "DUTY_PDF":
        # Квитанция госпошлины загружается на payment.xhtml.
        if current_path != target_path:
            # Иногда переход «Заполнение данных -> Оплата» не долетел
            # (флап портала / потерянный JSF-redirect). Пробуем довыполнить
            # его ещё раз перед тем, как признать шаг проваленным.
            if current_path.endswith("/createRequest.xhtml"):
                log.warning(
                    "Перед загрузкой квитанции всё ещё на createRequest "
                    "(%s) — повторяю переход на payment.xhtml.",
                    client.state.url,
                )
                _transition_create_request_to_payment(client)
                current_url = client.state.url or ""
                current_path = urlsplit(current_url).path

        if current_path != target_path:
            raise RuntimeError(
                f"Перед загрузкой квитанции ожидалась страница "
                f"{target_path}, текущая страница: {client.state.url}"
            )

    else:
        # Иск и приложения загружаются на blankData/other.xhtml.
        # Если скрипт ещё находится на payment.xhtml, выполняем
        # полноценный GET страницы документов с токеном черновика.
        if current_path != target_path:
            token = _current_request_token(client)

            target_page_url = urljoin(
                client.base_url + "/",
                target_path.lstrip("/"),
            )

            target_page_url = (
                target_page_url
                + "?"
                + urlencode({"i": token})
            )

            referer = client.state.url

            log.info(
                "Переход на страницу документов перед upload: %s -> %s",
                referer,
                target_page_url,
            )

            client.get(
                target_page_url,
                referer=referer,
            )

            current_url = client.state.url or ""
            current_path = urlsplit(current_url).path

            if current_path != target_path:
                raise RuntimeError(
                    "Не удалось открыть страницу документов перед upload. "
                    f"Ожидалась: {target_path}; "
                    f"получена: {client.state.url}"
                )

    log.info(
        "Страница upload готова: %s | ViewState=%s",
        client.state.url,
        "получен" if client.state.viewstate else "НЕ НАЙДЕН",
    )

    if not client.state.viewstate:
        raise RuntimeError(
            f"На странице {client.state.url} отсутствует "
            f"javax.faces.ViewState для upload {rule.name}."
        )

    # =========================================================
    # 2. Поиск RichFaces upload-компонента и JSF-формы
    # =========================================================

    soup = BeautifulSoup(
        client.state.html or "",
        "lxml",
    )

    component_suffix = rule.component_id.split(":")[-1]

    component = soup.find(
        lambda tag: (
            tag.name is not None
            and (
                str(tag.get("id") or "") == rule.component_id
                or str(tag.get("name") or "") == rule.component_id
                or str(tag.get("id") or "").endswith(
                    ":" + component_suffix
                )
                or str(tag.get("name") or "").endswith(
                    ":" + component_suffix
                )
            )
        )
    )

    if component is not None:
        form = component.find_parent("form")
    else:
        form = soup.find("form")

    # =========================================================
    # 3. Определение form_id
    # =========================================================

    if form is not None:
        form_id = str(
            form.get("id")
            or form.get("name")
            or ""
        ).strip()
    else:
        form_id = ""

    if not form_id:
        if target_path.endswith("/blankData/other.xhtml"):
            # Подтвержденная форма страницы документов.
            form_id = "j_idt34:j_idt36:j_idt39"

            log.warning(
                "JSF-форма не найдена в текущем HTML страницы %s. "
                "Вероятно, сохранён AJAX partial-response. "
                "Используем form_id страницы документов: %s",
                client.state.url,
                form_id,
            )

        elif target_path.endswith("/payment.xhtml"):
            # Подтвержденная форма страницы оплаты.
            form_id = "j_idt15:j_idt16"

            log.warning(
                "JSF-форма не найдена в текущем HTML страницы %s. "
                "Используем form_id страницы оплаты: %s",
                client.state.url,
                form_id,
            )

        else:
            raise RuntimeError(
                f"На странице {client.state.url} не найдена JSF-форма "
                f"для upload {rule.name}."
            )

    # =========================================================
    # 4. Определение component_id
    # =========================================================

    if component is not None:
        component_id = str(
            component.get("id")
            or component.get("name")
            or rule.component_id
        ).strip()

    else:
        # Если DOM сохранён как AJAX partial-response, upload-компонент
        # может отсутствовать. Тогда используем точный component_id из HAR.
        component_id = rule.component_id

        log.warning(
            "Upload-компонент %s не найден в DOM страницы %s. "
            "Используем component_id из HAR.",
            component_id,
            client.state.url,
        )

    if not component_id:
        raise RuntimeError(
            f"Не определён component_id для upload {rule.name}."
        )

    # =========================================================
    # 5. Формирование upload URL с актуальными параметрами
    # =========================================================

    query_pairs = parse_qsl(
        captured_parts.query,
        keep_blank_values=True,
    )

    updated_query: list[tuple[str, str]] = []

    viewstate_found = False
    source_found = False
    execute_found = False
    component_found = False

    for key, value in query_pairs:

        if key == "javax.faces.ViewState":
            updated_query.append(
                (
                    key,
                    client.state.viewstate,
                )
            )
            viewstate_found = True

        elif key == "javax.faces.source":
            updated_query.append(
                (
                    key,
                    component_id,
                )
            )
            source_found = True

        elif key == "javax.faces.partial.execute":
            updated_query.append(
                (
                    key,
                    component_id,
                )
            )
            execute_found = True

        elif key == "org.richfaces.ajax.component":
            updated_query.append(
                (
                    key,
                    component_id,
                )
            )
            component_found = True

        else:
            # rf_fu_uid, javax.faces.partial.ajax и остальные
            # параметры сохраняются из HAR.
            updated_query.append(
                (
                    key,
                    value,
                )
            )

    # На случай, если какого-либо обязательного параметра нет в HAR.
    if not viewstate_found:
        updated_query.append(
            (
                "javax.faces.ViewState",
                client.state.viewstate,
            )
        )

    if not source_found:
        updated_query.append(
            (
                "javax.faces.source",
                component_id,
            )
        )

    if not execute_found:
        updated_query.append(
            (
                "javax.faces.partial.execute",
                component_id,
            )
        )

    if not component_found:
        updated_query.append(
            (
                "org.richfaces.ajax.component",
                component_id,
            )
        )

    # Для upload всегда должен присутствовать признак AJAX.
    if not any(
        key == "javax.faces.partial.ajax"
        for key, _ in updated_query
    ):
        updated_query.append(
            (
                "javax.faces.partial.ajax",
                "true",
            )
        )

    upload_url = urlunsplit(
        (
            captured_parts.scheme,
            captured_parts.netloc,
            captured_parts.path,
            urlencode(updated_query),
            captured_parts.fragment,
        )
    )

    log.info(
        "Параметры upload: %s | form=%s | component=%s | "
        "ViewState=%s | page=%s",
        rule.name,
        form_id,
        component_id,
        client.state.viewstate,
        client.state.url,
    )

    return form_id, component_id, upload_url
    
def execute_multipart_upload(
    client: SudHttpClient,
    req: CapturedRequest,
    context: dict[str, str],
    case_files: dict[str, Path],
) -> requests.Response | None:
    """
    Выполняет RichFaces multipart только после синхронизации с фактической
    страницей upload текущего черновика.
    """
    rule = find_upload_rule(req)

    if rule.file_key == "NEXT_ATTACHMENT":
        queue = list(getattr(client, "current_attachment_queue", []))
        queue_index = int(
            getattr(client, "current_attachment_index", 0)
        )

        if queue_index >= len(queue):
            log.warning(
                "Портал запросил ещё одно обычное приложение, но в папке "
                "должника их больше нет (загружено %s, найдено %s). "
                "Пропускаю шаг.",
                queue_index, len(queue),
            )
            return None

        file_path = Path(queue[queue_index])
        rule.name = (
            f"Приложение {queue_index + 1}/{len(queue)}"
        )

        log.info(
            "ОЧЕРЕДЬ приложений: позиция %s/%s | файл=%s | путь=%s",
            queue_index + 1,
            len(queue),
            file_path.name,
            file_path,
        )
    else:
        file_path = case_files.get(rule.file_key)

        if file_path is None:
            available = ", ".join(sorted(case_files))
            raise FileNotFoundError(
                f"Для шага {rule.name!r} отсутствует {rule.file_key}. "
                f"Найдены ключи: {available}"
            )

    file_path = Path(file_path).resolve()
    if not file_path.exists():
        raise FileNotFoundError(file_path)

    known_parents = {
        Path(path).resolve().parent
        for path in case_files.values()
        if path is not None
    }
    if file_path.parent not in known_parents:
        raise RuntimeError(
            f"Файл находится вне папки текущего дела: {file_path}"
        )

    form_id, component_id, upload_url = _prepare_actual_upload_page(
        client,
        rule,
        req,
    )

    fields = [(form_id, form_id)]
    referer = client.state.url

    log.info(
        "RichFaces upload: %s | файл=%s | component=%s | "
        "form=%s | ViewState=%s | referer=%s",
        rule.name,
        file_path.name,
        component_id,
        form_id,
        "есть" if client.state.viewstate else "НЕТ",
        referer,
    )

    # Размер проверяется ещё до отправки, чтобы портал не создавал
    # частично заполненный черновик с последующей ошибкой 413.
    file_size_bytes = file_path.stat().st_size
    file_size_mb = file_size_bytes / (1024 * 1024)

    if file_size_bytes > MAX_UPLOAD_FILE_BYTES:
        raise RuntimeError(
            f"Файл превышает лимит портала 20 МБ: "
            f"{file_path.name} — {file_size_mb:.2f} МБ"
        )

    retryable_errors = (
        requests.exceptions.ConnectionError,
        requests.exceptions.SSLError,
        requests.exceptions.Timeout,
    )

    response = None
    for attempt in range(1, UPLOAD_RETRY_COUNT + 1):
        try:
            log.info(
                "Попытка upload %s/%s: %s | %.2f МБ",
                attempt,
                UPLOAD_RETRY_COUNT,
                file_path.name,
                file_size_mb,
            )

            response = client.upload(
                upload_url,
                fields=fields,
                file_field=component_id,
                file_path=file_path,
                referer=referer,
            )
            break

        except requests.exceptions.HTTPError as exc:
            status = exc.response.status_code if exc.response is not None else None

            if status == 413:
                raise RuntimeError(
                    f"Портал отклонил файл как слишком большой: "
                    f"{file_path.name} — {file_size_mb:.2f} МБ"
                ) from exc

            raise

        except retryable_errors as exc:
            if attempt >= UPLOAD_RETRY_COUNT:
                raise RuntimeError(
                    f"Загрузка оборвалась после {UPLOAD_RETRY_COUNT} попыток: "
                    f"{file_path.name} — {exc}"
                ) from exc

            delay = UPLOAD_RETRY_BASE_DELAY * attempt
            log.warning(
                "Соединение оборвалось при загрузке %s. "
                "Повтор через %s сек. Ошибка: %s",
                file_path.name,
                delay,
                exc,
            )
            time.sleep(delay)

            # Возвращаемся на полную страницу документов, чтобы получить
            # свежий ViewState перед следующей попыткой.
            token = _current_request_token(client)
            refresh_url = urljoin(
                client.base_url + "/",
                "form/requestType2/blankData/other.xhtml?"
                + urlencode({"i": token}),
            )
            client.get(refresh_url, referer=referer)

            form_id, component_id, upload_url = _prepare_actual_upload_page(
                client,
                rule,
                req,
            )
            fields = [(form_id, form_id)]
            referer = client.state.url

    if response is None:
        raise RuntimeError(f"Не получен ответ upload для файла {file_path.name}")

    # Запоминаем не только имя, но и назначение файла.
    # Это не позволяет следующему обычному документу перезаписать поле «Иск».
    client.last_uploaded_kind = rule.file_key
    client.last_uploaded_file = file_path.name

    # Переходим к следующему приложению только после подтверждённой
    # успешной загрузки текущего файла.
    if rule.file_key == "NEXT_ATTACHMENT":
        client.current_attachment_index = (
            int(getattr(client, "current_attachment_index", 0)) + 1
        )

    body = response.text or ""
    body_low = body.lower()
    content_type = (response.headers.get("Content-Type") or "").lower()

    explicit_error_markers = (
        "<error>",
        "<error-name>",
        "<error-message>",
        "ошибка загрузки файла",
        "не удалось загрузить файл",
        "upload failed",
        "file upload failed",
    )

    if any(marker in body_low for marker in explicit_error_markers):
        snippet = re.sub(r"\s+", " ", body)[:1200]
        raise RuntimeError(
            f"Сервер сообщил об ошибке загрузки {rule.name}: "
            f"{file_path.name}. Ответ: {snippet}"
        )

    if response.status_code >= 400:
        raise RuntimeError(
            f"Ошибка HTTP {response.status_code} при загрузке "
            f"{rule.name}: {file_path.name}"
        )

    log.info(
        "UPLOAD подтверждён: %s | status=%s | "
        "content-type=%s | bytes=%s",
        rule.name,
        response.status_code,
        content_type or "не указан",
        len(response.content or b""),
    )
    log.info(
        "Загружено: %s | %s | component=%s",
        rule.name,
        file_path.name,
        component_id,
    )

    time.sleep(PAUSE_BETWEEN_REQUESTS)
    return response

# %% [markdown]
# ## 7. Полностью HTTP-создание нового иска для каждой строки
# 

# %%
import re
SEND_DOCS_URL = BASE_URL + "/form/send/index.xhtml"


def _control_name(node) -> str:
    return str(node.get("name") or node.get("id") or "").strip()


def _find_select_by_suffix(soup: BeautifulSoup, suffix: str):
    suffix = suffix.lower()
    for node in soup.find_all("select"):
        name = _control_name(node).lower()
        if name.endswith(suffix):
            return node
    return None


def _set_selected_value(select, value: str) -> None:
    found = False
    for option in select.find_all("option"):
        option.attrs.pop("selected", None)
        if str(option.get("value") or "") == value:
            option["selected"] = "selected"
            found = True
    if not found:
        available = [str(x.get("value") or "") for x in select.find_all("option")]
        raise RuntimeError(
            f"В селекте {_control_name(select)} нет значения {value!r}. "
            f"Доступные значения: {available[:30]}"
        )


def _successful_form_fields(form) -> list[tuple[str, str]]:
    fields: list[tuple[str, str]] = []
    for node in form.find_all("input"):
        name = _control_name(node)
        if not name or node.has_attr("disabled"):
            continue
        typ = str(node.get("type") or "text").lower()
        if typ in {"submit", "button", "image", "file", "reset"}:
            continue
        if typ in {"checkbox", "radio"} and not node.has_attr("checked"):
            continue
        fields.append((name, str(node.get("value") or "")))
    for node in form.find_all("select"):
        name = _control_name(node)
        if not name or node.has_attr("disabled"):
            continue
        selected = node.find("option", selected=True)
        if selected is not None:
            fields.append((name, str(selected.get("value") or "")))
    return fields


def _extract_jsf_click_params(node) -> list[tuple[str, str]]:
    """Извлекает параметры command-компонента из onclick/jsfcljs."""
    script = str(node.get("onclick") or "")
    params: list[tuple[str, str]] = []

    match = re.search(r"jsfcljs\s*\([^,]+,\s*\{(.*?)\}\s*,", script, flags=re.I | re.S)
    if match:
        for key, value in re.findall(r"['\"]([^'\"]+)['\"]\s*:\s*['\"]([^'\"]*)['\"]", match.group(1)):
            params.append((key, value))

    if not params:
        source = re.search(
            r"(?:jsf\.ajax\.request|RichFaces\.ajax)\s*\(\s*['\"]([^'\"]+)['\"]",
            script,
            flags=re.I,
        )
        if source:
            params.append((source.group(1), source.group(1)))

    return params


def _extract_ajax_option(script: str, option: str, default: str) -> str:
    patterns = [
        rf"{re.escape(option)}\s*:\s*['\"]([^'\"]+)['\"]",
        rf"{re.escape(option)}\s*:\s*\[([^\]]+)\]",
        rf"javax\.faces\.partial\.{re.escape(option)}['\"]?\s*[:,=]\s*['\"]([^'\"]+)['\"]",
    ]
    for pattern in patterns:
        match = re.search(pattern, script, flags=re.I | re.S)
        if match:
            raw = match.group(1)
            quoted = re.findall(r"['\"]([^'\"]+)['\"]", raw)
            return " ".join(quoted) if quoted else raw.strip()
    return default


def _choose_next_control(form) -> tuple[list[tuple[str, str]], bool, str, str, str]:
    """Находит стартовую кнопку и возвращает полные AJAX-настройки JSF."""
    candidates = []
    for node in form.find_all(["input", "button", "a"]):
        name = _control_name(node)
        value = str(node.get("value") or "")
        visible = " ".join(
            x for x in [value, node.get_text(" ", strip=True), str(node.get("title") or "")]
            if x
        ).strip()
        text = visible.lower().replace("ё", "е")
        if not any(token in text for token in ("далее", "отправить", "продолжить", "next")):
            continue

        typ = str(node.get("type") or ("submit" if node.name == "button" else "")).lower()
        onclick = str(node.get("onclick") or "")
        score = 0
        if text.strip() in {"далее", "отправить"}:
            score -= 20
        if typ == "submit":
            score -= 5
        if node.name == "a":
            score += 5
        candidates.append((score, node, name, value, onclick))

    if not candidates:
        available = []
        for node in form.find_all(["input", "button", "a"]):
            label = " ".join(
                x for x in [str(node.get("value") or ""), node.get_text(" ", strip=True)] if x
            ).strip()
            if label:
                available.append(label)
        raise RuntimeError(
            "На стартовой форме не найдена кнопка перехода «Далее»/«Отправить». "
            f"Доступные кнопки/ссылки: {available[:30]}"
        )

    _, node, name, value, onclick = sorted(candidates, key=lambda x: x[0])[0]
    params = _extract_jsf_click_params(node)
    if not params:
        if not name:
            raise RuntimeError("У найденной кнопки нет name/id и не разобран onclick.")
        params = [(name, value or name)]

    is_ajax = bool(re.search(
        r"jsf\.ajax\.request|RichFaces\.ajax|javax\.faces\.partial\.ajax",
        onclick,
        flags=re.I,
    ))
    execute = _extract_ajax_option(onclick, "execute", name or "@this")
    render = _extract_ajax_option(onclick, "render", "@all")

    log.info(
        "Найдена стартовая кнопка перехода: name=%s value=%s ajax=%s execute=%s render=%s params=%s onclick=%s",
        name, value, is_ajax, execute, render, params, onclick[:500],
    )
    return params, is_ajax, name, execute, render

def _extract_render_targets(node) -> str:
    script = str(node.get("onchange") or node.get("onblur") or "")
    patterns = [
        r"render\s*:\s*['\"]([^'\"]+)['\"]",
        r"render\s*:\s*\[([^\]]+)\]",
        r"javax\.faces\.partial\.render['\"]?\s*[:,=]\s*['\"]([^'\"]+)['\"]",
    ]
    for pattern in patterns:
        m = re.search(pattern, script, flags=re.I)
        if m:
            value = m.group(1)
            quoted = re.findall(r"['\"]([^'\"]+)['\"]", value)
            return " ".join(quoted) if quoted else value.strip()
    return "@all"


def _apply_partial_updates(current_soup: BeautifulSoup, updates: dict[str, str]) -> BeautifulSoup:
    """Применяет JSF partial-response к локальной копии HTML без нового GET."""
    for update_id, fragment_html in updates.items():
        if "ViewState" in update_id:
            continue

        # Некоторые JSF реализации присылают целое представление.
        if update_id in {"javax.faces.ViewRoot", "jakarta.faces.ViewRoot"}:
            parsed = BeautifulSoup(fragment_html, "lxml")
            if parsed.find("form"):
                current_soup = parsed
            continue

        target = current_soup.find(id=update_id)
        if target is None:
            target = current_soup.find(attrs={"name": update_id})

        fragment_soup = BeautifulSoup(fragment_html, "lxml")
        replacement = fragment_soup.find(id=update_id)
        if replacement is None:
            # CDATA может содержать один корневой элемент без ожидаемого id.
            replacement = fragment_soup.body
            if replacement is not None:
                children = [x for x in replacement.contents if getattr(x, "name", None)]
                replacement = children[0] if len(children) == 1 else replacement

        if target is not None and replacement is not None:
            target.replace_with(replacement)

    return current_soup


def _refresh_viewstate_input(soup: BeautifulSoup, viewstate: str | None) -> None:
    if not viewstate:
        return
    for name in VIEWSTATE_NAMES:
        node = soup.find("input", attrs={"name": name})
        if node is not None:
            node["value"] = viewstate
            return


def _ajax_select_value(
    client: SudHttpClient,
    response: requests.Response,
    soup: BeautifulSoup,
    suffix: str,
    value: str,
    label: str,
    chosen_values: dict[str, str],
) -> tuple[requests.Response, BeautifulSoup]:
    select = _find_select_by_suffix(soup, suffix)
    if select is None:
        available = [_control_name(x) for x in soup.find_all("select")]
        raise RuntimeError(f"Не найден селект {label} ({suffix}). Найдены: {available}")

    select_name = _control_name(select)
    form = select.find_parent("form")
    if form is None:
        raise RuntimeError(f"Селект {label} не находится внутри form.")
    form_id = _control_name(form)
    if not form_id:
        raise RuntimeError(f"У формы селекта {label} нет id/name.")

    # Локально отмечаем новое значение и сохраняем все предыдущие выборы.
    _set_selected_value(select, value)
    chosen_values[select_name] = value

    fields = _successful_form_fields(form)
    excluded = set(VIEWSTATE_NAMES) | set(chosen_values)
    fields = [(k, v) for k, v in fields if k not in excluded]
    fields.extend(chosen_values.items())
    fields.extend([
        (form_id, form_id),
        ("javax.faces.partial.ajax", "true"),
        ("javax.faces.source", select_name),
        ("javax.faces.partial.execute", "@this"),
        ("javax.faces.partial.render", _extract_render_targets(select)),
        ("javax.faces.behavior.event", "valueChange"),
        ("javax.faces.partial.event", "change"),
    ])

    action = urljoin(response.url, str(form.get("action") or response.url))
    ajax_response = client.post_form(
        action,
        fields,
        referer=response.url,
        ajax=True,
        allow_redirects=True,
        allow_services=True,
    )

    updates, new_viewstate, redirect_url = parse_partial_response(ajax_response.text)
    if redirect_url:
        raise RuntimeError(
            f"Во время AJAX-выбора {label} портал неожиданно прислал redirect: {redirect_url}"
        )
    if not updates:
        preview = ajax_response.text[:500].replace("\n", " ")
        raise RuntimeError(
            f"AJAX-выбор {label} вернул HTTP 200, но без JSF update. "
            f"Начало ответа: {preview}"
        )

    soup = _apply_partial_updates(soup, updates)
    _refresh_viewstate_input(soup, new_viewstate or client.state.viewstate)

    # Повторно закрепляем уже выбранные значения: update мог заменить сам select.
    for control_name, selected_value in chosen_values.items():
        node = soup.find("select", attrs={"name": control_name}) or soup.find("select", id=control_name)
        if node is not None:
            _set_selected_value(node, selected_value)

    log.info(
        "AJAX выбор %s=%s подтверждён | updates=%s | ViewState=%s",
        label,
        value,
        list(updates)[:8],
        "обновлён" if new_viewstate else "прежний",
    )
    return ajax_response, soup


def _extract_send_request_source(soup: BeautifulSoup, button_source: str) -> str | None:
    """Извлекает source предварительного RichFaces.ajax из sendRequest()."""
    page_text = str(soup)
    patterns = [
        r"sendRequest\s*=\s*function\s*\([^)]*\)\s*\{.*?RichFaces\.ajax\(\s*[\"']([^\"']+)[\"']",
        r"function\s+sendRequest\s*\([^)]*\)\s*\{.*?RichFaces\.ajax\(\s*[\"']([^\"']+)[\"']",
    ]
    for pattern in patterns:
        match = re.search(pattern, page_text, flags=re.I | re.S)
        if match:
            return match.group(1)

    # Резервный вариант для текущей формы портала: соседний component id.
    match = re.search(r":j_idt(\d+)$", button_source or "")
    if match:
        return re.sub(r":j_idt\d+$", f":j_idt{int(match.group(1)) + 1}", button_source)
    return None


def _build_start_form_fields(
    form,
    chosen_values: dict[str, str],
    *,
    source: str | None = None,
    execute: str | None = None,
    render: str | None = None,
    command_value: str | None = None,
) -> list[tuple[str, str]]:
    """Собирает актуальные поля стартовой JSF-формы."""
    fields = _successful_form_fields(form)
    excluded = set(VIEWSTATE_NAMES) | set(chosen_values)
    fields = [(k, v) for k, v in fields if k not in excluded]
    fields.extend(chosen_values.items())

    form_id = _control_name(form)
    if form_id and not any(k == form_id for k, _ in fields):
        fields.append((form_id, form_id))

    if source:
        remove_keys = {
            source,
            "javax.faces.partial.ajax",
            "javax.faces.source",
            "javax.faces.partial.execute",
            "javax.faces.partial.render",
            "org.richfaces.ajax.component",
        }
        fields = [(k, v) for k, v in fields if k not in remove_keys]
        fields.extend([
            ("javax.faces.partial.ajax", "true"),
            ("javax.faces.source", source),
            ("javax.faces.partial.execute", execute or source),
            ("javax.faces.partial.render", render or "@all"),
            ("org.richfaces.ajax.component", source),
            (source, command_value or source),
        ])
    return fields



def _ajax_select_value_exact_har(
    client: SudHttpClient,
    response: requests.Response,
    soup: BeautifulSoup,
    suffix: str,
    value: str,
    label: str,
    chosen_values: dict[str, str],
) -> tuple[requests.Response, BeautifulSoup]:
    """
    Точный RichFaces AJAX выбора значения по новому HAR office.sud.kz(2).har.

    В отличие от общего helper отправляет параметры именно как браузер:
      execute = source + @component
      render = @component
      behavior.event = change
      org.richfaces.ajax.component = source
      rfExt = null
      AJAX:EVENTS_COUNT = 1
    """
    select = _find_select_by_suffix(soup, suffix)
    if select is None:
        available = [_control_name(x) for x in soup.find_all("select")]
        raise RuntimeError(
            f"Не найден селект {label} ({suffix}). Найдены: {available}"
        )

    select_name = _control_name(select)
    if not select_name:
        raise RuntimeError(f"У селекта {label} отсутствует id/name.")

    form = select.find_parent("form")
    if form is None:
        raise RuntimeError(f"Селект {label} не находится внутри form.")

    form_id = _control_name(form)
    if not form_id:
        raise RuntimeError(f"У формы селекта {label} отсутствует id/name.")

    # До отправки убеждаемся, что нужное значение реально появилось в DOM.
    option_values = [
        str(option.get("value") if option.has_attr("value") else option.get_text(" ", strip=True))
        for option in select.find_all("option")
    ]
    if value not in option_values:
        raise RuntimeError(
            f"В селекте {label} нет значения {value!r}. "
            f"Доступные значения: {option_values}"
        )

    _set_selected_value(select, value)
    chosen_values[select_name] = value

    fields = _successful_form_fields(form)

    remove_keys = (
        set(VIEWSTATE_NAMES)
        | set(chosen_values)
        | {
            form_id,
            "javax.faces.partial.ajax",
            "javax.faces.source",
            "javax.faces.partial.event",
            "javax.faces.partial.execute",
            "javax.faces.partial.render",
            "javax.faces.behavior.event",
            "org.richfaces.ajax.component",
            "rfExt",
            "AJAX:EVENTS_COUNT",
        }
    )
    fields = [(k, v) for k, v in fields if k not in remove_keys]

    # Порядок полей соответствует новому успешному HAR.
    fields.append((form_id, form_id))
    fields.extend(chosen_values.items())
    fields.extend([
        ("javax.faces.source", select_name),
        ("javax.faces.partial.event", "change"),
        ("javax.faces.partial.execute", f"{select_name} @component"),
        ("javax.faces.partial.render", "@component"),
        ("javax.faces.behavior.event", "change"),
        ("org.richfaces.ajax.component", select_name),
        ("rfExt", "null"),
        ("AJAX:EVENTS_COUNT", "1"),
        ("javax.faces.partial.ajax", "true"),
    ])

    action = urljoin(response.url, str(form.get("action") or response.url))
    ajax_response = client.post_form(
        action,
        fields,
        referer=response.url,
        ajax=True,
        allow_redirects=True,
        allow_services=True,
    )

    updates, new_viewstate, redirect_url = parse_partial_response(ajax_response.text)
    if redirect_url:
        raise RuntimeError(
            f"Во время выбора {label} портал неожиданно вернул redirect: {redirect_url}"
        )
    if not updates:
        preview = ajax_response.text[:1000].replace("\n", " ")
        raise RuntimeError(
            f"AJAX-выбор {label} вернул HTTP {ajax_response.status_code}, "
            f"но без JSF update. Начало ответа: {preview}"
        )

    soup = _apply_partial_updates(soup, updates)
    _refresh_viewstate_input(soup, new_viewstate or client.state.viewstate)

    for control_name, selected_value in chosen_values.items():
        node = (
            soup.find("select", attrs={"name": control_name})
            or soup.find("select", id=control_name)
        )
        if node is not None:
            _set_selected_value(node, selected_value)

    log.info(
        "START HAR AJAX %s=%s | updates=%s | ViewState=%s",
        label,
        value,
        list(updates)[:10],
        "обновлён" if new_viewstate else "прежний",
    )
    return ajax_response, soup


def create_new_claim_http(client: SudHttpClient) -> str:
    """
    Создаёт новый черновик по точной последовательности нового HAR:

      1) GET /form/send/index.xhtml
      2) AJAX case-type=CIVIL
      3) AJAX instance=FIRSTINSTANCE
      4) AJAX request=3
      5) скрытый RichFaces command sendRequest
      6) JSF redirect на свежий createRequest.xhtml?i=...
      7) GET нового черновика
    """
    response = client.get(
        SEND_DOCS_URL,
        referer=client.state.url,
    )
    soup = BeautifulSoup(response.text, "lxml")

    debug_path = Path.cwd() / "first_page.html"
    debug_path.write_text(response.text, encoding="utf-8")
    log.info("Стартовая HTML-страница сохранена: %s", debug_path)

    chosen_values: dict[str, str] = {}

    response, soup = _ajax_select_value_exact_har(
        client,
        response,
        soup,
        ":case-type",
        "CIVIL",
        "Тип производства",
        chosen_values,
    )

    response, soup = _ajax_select_value_exact_har(
        client,
        response,
        soup,
        ":instance",
        "FIRSTINSTANCE",
        "Инстанция",
        chosen_values,
    )

    # Новый HAR подтвердил: после точного AJAX инстанции значение 3
    # появляется в обновлённом request-type-panel.
    response, soup = _ajax_select_value_exact_har(
        client,
        response,
        soup,
        ":request",
        "3",
        "Тип документа",
        chosen_values,
    )

    request_select = _find_select_by_suffix(soup, ":request")
    if request_select is None:
        raise RuntimeError("После выбора типа документа исчез селект request.")

    form = request_select.find_parent("form")
    if form is None:
        raise RuntimeError("Стартовый селект request не находится внутри form.")

    form_id = _control_name(form)
    if not form_id:
        raise RuntimeError("У стартовой формы отсутствует id/name.")

    # После AJAX request=3 сервер добавляет описание услуги и JS-функцию
    # sendRequest(), содержащую точный динамический source кнопки.
    source = _extract_send_request_source(soup, "")
    if not source:
        debug_after = Path.cwd() / "start_after_request_3.html"
        debug_after.write_text(str(soup), encoding="utf-8")
        raise RuntimeError(
            "После выбора request=3 не найден RichFaces source функции "
            f"sendRequest(). HTML сохранён: {debug_after}"
        )

    action = urljoin(response.url, str(form.get("action") or SEND_DOCS_URL))

    fields = _build_start_form_fields(
        form,
        chosen_values,
        source=source,
        execute=f"{source} @component",
        render="@component",
        command_value=source,
    )

    remove_keys = {"rfExt", "AJAX:EVENTS_COUNT"}
    fields = [(k, v) for k, v in fields if k not in remove_keys]
    fields.extend([
        ("rfExt", "null"),
        ("AJAX:EVENTS_COUNT", "1"),
    ])

    log.info(
        "START HAR COMMAND | form=%s | source=%s | request=3 | ViewState=%s",
        form_id,
        source,
        "есть" if client.state.viewstate else "не найден",
    )

    result = client.post_form(
        action,
        fields,
        referer=response.url,
        ajax=True,
        allow_redirects=True,
        allow_services=True,
    )

    current_url = result.url or client.state.url

    log.info(
        "START HAR COMMAND -> HTTP %s | итоговый URL=%s",
        result.status_code,
        current_url,
    )

    # post_form() уже обработал JSF redirect и вернул GET-ответ
    # открытой страницы нового заявления.
    if "/form/requesttype2/createrequest.xhtml" not in current_url.lower():
        preview = result.text[:1600].replace("\n", " ")
        raise RuntimeError(
            "Команда создания иска не открыла createRequest.xhtml. "
            f"HTTP={result.status_code}; URL={current_url}; "
            f"начало ответа={preview}"
        )

    if "/form/requesttype2/createrequest.xhtml" not in current_url.lower():
        raise RuntimeError(
            "После JSF redirect не открылась страница нового иска. "
            f"Ожидался createRequest.xhtml, получен URL: {current_url}"
        )

    if not client.state.viewstate:
        raise RuntimeError(
            "Новый черновик открыт, но на createRequest.xhtml не найден ViewState."
        )

    _remember_request_token(client, current_url)
    log.info("Создан новый черновик иска: %s", current_url)
    return current_url


# %% [markdown]
# ## 7. Запуск по строкам Excel

# %%

def get_excel_rows_to_process(
    start_row: int = START_ROW,
    end_row: int | None = END_ROW,
) -> list[int]:
    """
    Возвращает все строки Excel с заполненным ИИН в колонке COL_IIN.

    Если end_row=None, сканирует лист до последней используемой строки.
    Пустые строки внутри таблицы пропускаются.
    """
    wb = load_workbook(EXCEL_FILE, data_only=True, read_only=True)
    try:
        ws = wb[EXCEL_SHEET] if EXCEL_SHEET else wb.active
        last_row = end_row if end_row is not None else ws.max_row

        rows: list[int] = []
        for row in range(start_row, last_row + 1):
            raw_iin = ws[f"{COL_IIN}{row}"].value
            iin = str(raw_iin or "").strip()

            if not iin:
                log.info("Строка %s пропущена: ИИН пустой.", row)
                continue

            # Excel иногда превращает ИИН в число вида 30719500780.0.
            if iin.endswith(".0"):
                iin = iin[:-2]

            if not iin.isdigit():
                log.warning(
                    "Строка %s пропущена: некорректный ИИН %r.",
                    row,
                    raw_iin,
                )
                continue

            rows.append(row)

        return rows
    finally:
        wb.close()


def validate_http_only_notebook() -> None:
    """
    Контроль: после блока получения cookies код не должен обращаться к driver.
    """
    assert isinstance(CLIENT.session, requests.Session)
    log.info("Проверка HTTP-only пройдена.")


def run_row(row: int, captures: list[CapturedRequest], position: int = 0, total: int = 0) -> None:
    context = load_case_row(row)
    progress(f"[{position}/{total}] {context['FIO']} ({context['IIN']})")
    folder = find_case_folder(context["IIN"], context["FIO"])
    case_files = classify_case_files(folder)
    validate_case_files(folder, case_files, context)

    attachment_queue = get_attachment_queue(case_files)
    CLIENT.current_attachment_queue = attachment_queue
    CLIENT.current_attachment_index = 0

    log.info(
        "Очередь приложений текущего должника: %s файлов",
        len(attachment_queue),
    )
    for index, path in enumerate(attachment_queue, start=1):
        log.info(
            "  очередь %s/%s | %s | %.2f МБ | %s",
            index,
            len(attachment_queue),
            path.name,
            path.stat().st_size / (1024 * 1024),
            path,
        )

    context["CASE_FOLDER"] = str(folder)
    context["CLAIM_TEXT"] = read_docx_text(case_files["CLAIM_DOCX"])

    # Имена файлов текущего должника сохраняем в context.
    # context доступен во всех обычных HAR POST-запросах.
    context["CLAIM_FILENAME"] = case_files["CLAIM_DOCX"].name
    context["DUTY_FILENAME"] = case_files["DUTY_PDF"].name

    log.info(
        "Файлы текущего дела: CLAIM_FILENAME=%s | DUTY_FILENAME=%s",
        context["CLAIM_FILENAME"],
        context["DUTY_FILENAME"],
    )

    log.info("=" * 80)
    log.info("Строка %s | ИИН %s | %s", row, context["IIN"], context["FIO"])
    log.info("Папка дела: %s", folder)

    # Не переносим область, суд и состояние upload от предыдущего должника.
    for attr in (
        "current_region_value",
        "current_region_caption",
        "current_court_value",
        "current_court_caption",
        "current_court_name",
        "current_region_name",
        "current_row",
        "defendant_livePlace",
        "_court_giveup",
        "last_uploaded_kind",
        "last_uploaded_file",
    ):
        if hasattr(CLIENT, attr):
            delattr(CLIENT, attr)

    # Для каждой строки создаём новый свежий черновик через точную
    # стартовую последовательность из нового HAR.
    create_new_claim_http(CLIENT)

    # Затем выполняем основной HAR заполнения открытого черновика.
    run_capture_sequence(CLIENT, captures, context, case_files)

    # Проверка: если суд из Excel так и не удалось подставить (и это ещё не
    # отмечено внутри _resolve_case_court) — иск ушёл в суд по умолчанию.
    _cv = getattr(CLIENT, "current_court_value", None)
    if (not _cv or _cv == _HAR_DEFAULT_COURT) and not getattr(CLIENT, "_court_giveup", False):
        row_court = str(
            context.get("COURT") or context.get("REGION") or ""
        )
        if not any(r == str(row) for r, _ in DEFAULT_COURT_ROWS):
            DEFAULT_COURT_ROWS.append((str(row), row_court))
        log.warning(
            "Строка %s: суд из Excel не подставлен — иск подан с судом по "
            "умолчанию (Астана). Проверьте вручную.", row,
        )

    uploaded_attachment_count = int(
        getattr(CLIENT, "current_attachment_index", 0)
    )
    expected_attachment_count = len(attachment_queue)

    if uploaded_attachment_count != expected_attachment_count:
        raise RuntimeError(
            "Не все приложения из папки должника были загружены: "
            f"загружено {uploaded_attachment_count}, "
            f"найдено {expected_attachment_count}."
        )

    if STOP_AT_SIGN_PAGE and not is_sign_page(CLIENT.state.url, CLIENT.state.html):
        raise RuntimeError(
            f"Строка {row}: не достигнута страница подписания. URL={CLIENT.state.url}"
        )

    log.info("Строка %s завершена. Страница подписания: %s", row, CLIENT.state.url)
    progress("   ✅ Иск подготовлен, дошли до страницы подписания")
    # Следующая строка снова создаст отдельный новый черновик.


validate_http_only_notebook()
CAPTURES = load_recipe()
print_actions(CAPTURES)
progress(f"📄 Шаблон подачи загружен ({sum(1 for r in CAPTURES if r.is_http_action)} шагов)")

ROWS_TO_PROCESS = get_excel_rows_to_process()
if not ROWS_TO_PROCESS:
    raise RuntimeError(
        f"В Excel не найдено строк для обработки, начиная со строки {START_ROW}."
    )

log.info(
    "Найдено строк для обработки: %s | строки: %s",
    len(ROWS_TO_PROCESS),
    ROWS_TO_PROCESS,
)

MAX_SESSION_REFRESHES = 2
ROW_RETRY_ATTEMPTS = 2      # повторов строки при обычной (не портал/сессия) ошибке
ROW_RETRY_DELAY = 5


def _refresh_session_after_expiry() -> bool:
    """Повторно логинится в Chrome и пересобирает SESSION/CLIENT, если
    посреди пакета истекла авторизация портала."""
    global AUTH, SESSION, CLIENT
    try:
        AUTH = get_auth_from_chrome_auto()
        SESSION = make_http_session(AUTH)
        CLIENT = SudHttpClient(SESSION)
        return True
    except Exception as exc:
        log.error("Не удалось обновить авторизацию: %s", exc)
        return False


completed_rows: list[int] = []
failed_rows: list[tuple[int, str]] = []
portal_failed_rows: list[tuple[int, str]] = []
skipped_rows: list[int] = []
total_rows = len(ROWS_TO_PROCESS)
consecutive_portal_errors = 0
session_refreshes = 0
aborted_by_portal = False
aborted_by_session = False

progress(f"Всего клиентов для подачи: {total_rows}\n")

for position, excel_row in enumerate(ROWS_TO_PROCESS, start=1):
    log.info(
        "ПАКЕТ %s/%s | начинаю строку Excel %s",
        position,
        total_rows,
        excel_row,
    )

    try:
        run_row(excel_row, CAPTURES, position=position, total=total_rows)
        completed_rows.append(excel_row)
        consecutive_portal_errors = 0
    except Exception as exc:
        log.exception("Строка %s завершилась ошибкой: %s", excel_row, exc)
        msg_low = str(exc).casefold()
        is_session = (
            isinstance(exc, SessionExpired)
            or "сессия не авторизована" in msg_low
            or "страницу входа" in msg_low
            or "не найден селект тип производства" in msg_low
        )
        is_portal = (not is_session) and (
            isinstance(exc, PortalUnavailable) or _is_portal_outage_text(str(exc))
        )

        if is_session:
            portal_failed_rows.append((excel_row, str(exc)))
            progress(f"   ⚠️ Сессия истекла: {short_error(exc)}")
            if session_refreshes >= MAX_SESSION_REFRESHES:
                remaining = ROWS_TO_PROCESS[position:]
                skipped_rows.extend(remaining)
                aborted_by_session = True
                progress(
                    "\n⛔ Авторизация портала истекает повторно и не "
                    f"восстанавливается. Останавливаю пакет. Не тронуто "
                    f"строк: {len(remaining)} — перезапустите позже."
                )
                break
            session_refreshes += 1
            progress(f"   🔄 Обновляю авторизацию (попытка {session_refreshes})…")
            if _refresh_session_after_expiry():
                consecutive_portal_errors = 0
                progress(
                    "   ✅ Авторизация обновлена, продолжаю со следующей строки."
                )
            else:
                remaining = ROWS_TO_PROCESS[position:]
                skipped_rows.extend(remaining)
                aborted_by_session = True
                progress(
                    "\n⛔ Не удалось обновить авторизацию — останавливаю пакет. "
                    f"Не тронуто строк: {len(remaining)} — перезапустите позже."
                )
                break
        elif is_portal:
            portal_failed_rows.append((excel_row, str(exc)))
            consecutive_portal_errors += 1
            progress(
                f"   ⚠️ Портал недоступен ({consecutive_portal_errors}/"
                f"{PORTAL_ERROR_ABORT_THRESHOLD}): {short_error(exc)}"
            )
            if consecutive_portal_errors >= PORTAL_ERROR_ABORT_THRESHOLD:
                remaining = ROWS_TO_PROCESS[position:]
                skipped_rows.extend(remaining)
                aborted_by_portal = True
                progress(
                    f"\n⛔ {consecutive_portal_errors} ошибок портала подряд — "
                    f"office.sud.kz недоступен. Останавливаю пакет. "
                    f"Не тронуто строк: {len(remaining)} — перезапустите позже."
                )
                break
            time.sleep(PORTAL_ERROR_COOLDOWN)
        else:
            # Обычная ошибка по клиенту. Значительная часть таких сбоев
            # транзиентна (портал не вернул данные поиска по ИИН, goNext
            # не долетел) — пробуем строку ещё раз с нуля (новый черновик).
            retried_ok = False
            for row_try in range(1, ROW_RETRY_ATTEMPTS + 1):
                progress(
                    f"   ↻ Повтор строки {excel_row} "
                    f"({row_try}/{ROW_RETRY_ATTEMPTS}) после: {short_error(exc)}"
                )
                time.sleep(ROW_RETRY_DELAY)
                try:
                    run_row(excel_row, CAPTURES, position=position, total=total_rows)
                    completed_rows.append(excel_row)
                    consecutive_portal_errors = 0
                    retried_ok = True
                    progress("   ✅ Со второй попытки удалось.")
                    break
                except Exception as exc2:
                    log.exception("Повтор строки %s не удался: %s", excel_row, exc2)
                    exc = exc2

            if not retried_ok:
                failed_rows.append((excel_row, str(exc)))
                consecutive_portal_errors = 0
                progress(f"   ❌ Ошибка (после {ROW_RETRY_ATTEMPTS} повторов): "
                         f"{short_error(exc)}")
                if not CONTINUE_ON_ROW_ERROR:
                    raise

log.info("=" * 80)
log.info("ПАКЕТ ЗАВЕРШЁН")
log.info("Успешно обработаны строки: %s", completed_rows)
log.info("Строки с ошибками клиента: %s", failed_rows)
log.info("Строки с ошибкой портала: %s", portal_failed_rows)
log.info("Пропущено (пакет остановлен): %s", skipped_rows)

processed = len(completed_rows) + len(failed_rows) + len(portal_failed_rows)

progress("")
progress("=" * 60)
progress(
    f"ИТОГО: успешно {len(completed_rows)}/{total_rows}, "
    f"ошибок клиента {len(failed_rows)}, "
    f"портал недоступен {len(portal_failed_rows)}, "
    f"пропущено {len(skipped_rows)}"
)
if failed_rows:
    progress("\nНе поданы — проблема по клиенту (нужен разбор):")
    for row, error in failed_rows:
        progress(f"   строка {row}: {short_error(Exception(error))}")
if portal_failed_rows or skipped_rows:
    to_retry = sorted(
        {r for r, _ in portal_failed_rows} | set(skipped_rows)
    )
    progress(
        "\nНе поданы — портал был недоступен, ПРОСТО ПЕРЕЗАПУСТИТЕ ПОЗЖЕ."
    )
    progress(f"   строки для повторной подачи: {to_retry}")
if aborted_by_portal:
    progress(
        "\n⛔ Пакет остановлен из-за недоступности office.sud.kz — "
        "это не ошибка скрипта. Повторите запуск через 15–30 минут."
    )
if aborted_by_session:
    progress(
        "\n⛔ Пакет остановлен: авторизация портала не держится. "
        "Войдите вручную в Судебный кабинет в профиле Chrome "
        f"({CHROME_USER_DATA_DIR}), проверьте sk_login/sk_password в "
        "config.py, затем перезапустите."
    )
if DEFAULT_COURT_ROWS:
    progress(
        f"\n⚠️ ВНИМАНИЕ: {len(DEFAULT_COURT_ROWS)} иск(ов) поданы с судом "
        "ПО УМОЛЧАНИЮ (Астана) — суд из Excel не сопоставлен со списком "
        "портала. Проверьте и при необходимости переподайте вручную:"
    )
    for _r, _c in DEFAULT_COURT_ROWS:
        progress(f"   строка {_r}: {_c}")
progress("=" * 60)


# %% [markdown]
# ## Важные замечания
# 
# 1. Ноутбук не выполняет подписание ЭЦП и останавливается на странице подписания.
# 2. Cookies действуют ограниченное время. При возврате на вход повторно выполните блок получения cookies.
# 3. Если office.sud.kz изменит динамические JSF-id, пересоберите
#    `RECIPE_STEPS` в podacha_recipe.py из свежего HAR успешной подачи
#    (исходники и инструкция — в scripts/_legacy_har_capture/).
# 4. Chrome используется только для получения cookies, User-Agent и текущего URL.
# 5. Для `Обычный документ` сейчас берётся `EXTRA_1`; найденный файл выводится в журнал перед отправкой.
# 


