# aisoip_web.py
"""
Встроенный поиск/синхронизация АИС ОИП, портировано из
C:\\Users\\User\\Desktop\\ais_oip_internal_v3 (app.py + aisoip_sync.py) и
адаптировано для встраивания в основной FastAPI-интерфейс (app.py) с
привязкой данных и синхронизации к company_id (1..4, как в scripts/config.py).

Данные хранятся отдельно от users.db — в data/ais_oip.db, с колонкой
company_id и уникальностью (company_id, production_no, iin), чтобы компании
не видели данные друг друга.
"""
from __future__ import annotations

import base64
import json
import math
import re
import sqlite3
import threading
import time
import traceback
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, asdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Optional

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "data" / "ais_oip.db"
SYNC_TMP_DIR = BASE_DIR / "data" / "aisoip_sync_tmp"
LOGS_DIR = BASE_DIR / "logs" / "aisoip_sync"
SYNC_TMP_DIR.mkdir(parents=True, exist_ok=True)
LOGS_DIR.mkdir(parents=True, exist_ok=True)

# Логины/пароли АИС ОИП по company_id — те же значения, что в
# scripts/config.py (COMPANY_CREDENTIALS[<id>]['aisoip_login'/'aisoip_password']).
# Продублировано здесь намеренно: scripts/config.py при импорте делает glob по
# сетевым папкам и печатает служебные сообщения — импортировать его в живой
# веб-процесс небезопасно (может зависнуть, если сетевой диск недоступен).
COMPANY_AISOIP_CREDENTIALS = {
    1: {"login": "810813301334_230240016634", "password": "Qazaq123456*"},
    2: {"login": "901021350973_210540032049", "password": "Kk12345%"},
    3: {"login": "821029401373_251140014482", "password": "Aa12345%"},
    4: {"login": "840622302062_260140035543", "password": "Aa12345%"},
}

AISOIP_URL = "https://aisoip.adilet.gov.kz/cabinet/exec-productions"
API_SEARCH_URL = "https://aisoip.adilet.gov.kz/extperson/api/rest/execproc/search"
API_EXPORT_URL = "https://aisoip.adilet.gov.kz/extperson/api/rest/export/excel"
API_RETRIES = 3
API_TIMEOUT_SECONDS = 240
EXPORT_REQUEST_SIZE = 50_000
DEFAULT_DATE_FROM = "2023-01-01"

NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
# Атрибут r:id у <sheet> в workbook.xml живёт в неймспейсе officeDocument, а не
# package — "package/2006/relationships" используется в самих .rels-файлах для
# Relationship Type, а не как неймспейс атрибута r:id на элементах контента.
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

EXPECTED_HEADERS = {
    "Номер исполнительного документа": "document_no",
    "Номер исполнительного производства": "production_no",
    "Судебный исполнитель": "executor",
    "Дата возбуждения": "start_date",
    "Дата выписки исполнительного документа": "document_date",
    "Орган выдавший исполнительный документ": "issuer",
    "Тип взыскания": "collection_type",
    "Должник": "debtor",
    "ИИН/БИН должника": "iin",
    "Статус исполнительного производства": "status",
    "Сумма": "amount",
}


# ═══════════════════════════════════════════════════════════════
# БАЗА ДАННЫХ
# ═══════════════════════════════════════════════════════════════

def db_connect():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_aisoip_db():
    with db_connect() as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS productions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            document_no TEXT,
            production_no TEXT NOT NULL,
            executor TEXT,
            start_date TEXT,
            document_date TEXT,
            issuer TEXT,
            collection_type TEXT,
            debtor TEXT,
            iin TEXT,
            status TEXT,
            amount REAL,
            imported_at TEXT NOT NULL,
            source_file TEXT,
            UNIQUE(company_id, production_no, iin)
        );
        CREATE INDEX IF NOT EXISTS idx_productions_company ON productions(company_id);
        CREATE INDEX IF NOT EXISTS idx_productions_iin ON productions(company_id, iin);
        CREATE INDEX IF NOT EXISTS idx_productions_debtor ON productions(company_id, debtor);
        CREATE INDEX IF NOT EXISTS idx_productions_executor ON productions(company_id, executor);
        CREATE INDEX IF NOT EXISTS idx_productions_status ON productions(company_id, status);
        CREATE INDEX IF NOT EXISTS idx_productions_start_date ON productions(company_id, start_date);
        CREATE INDEX IF NOT EXISTS idx_productions_document_no ON productions(company_id, document_no);
        """)


# ═══════════════════════════════════════════════════════════════
# ПАРСИНГ / ИМПОРТ EXCEL (портировано из ais_oip_internal_v3/app.py)
# ═══════════════════════════════════════════════════════════════

def col_index(cell_ref: str) -> int:
    m = re.match(r"[A-Z]+", cell_ref.upper())
    letters = m.group(0) if m else "A"
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def excel_date(value):
    if value is None or value == "":
        return None
    try:
        serial = float(value)
        dt = datetime(1899, 12, 30) + timedelta(days=serial)
        return dt.strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        text = str(value).strip()
        for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
            try:
                return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
            except ValueError:
                pass
        return text


def parse_xlsx(path: str):
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        shared = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root:
                shared.append("".join((t.text or "") for t in si.iter(f"{{{NS['a']}}}t")))

        wb = ET.fromstring(z.read("xl/workbook.xml"))
        sheet = wb.find("a:sheets", NS)[0]
        rid = sheet.attrib[RID]
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        relmap = {r.attrib["Id"]: r.attrib["Target"] for r in rels}
        sheet_path = relmap[rid]
        # Target с ведущим "/" — абсолютный путь от корня пакета (уже включает "xl/",
        # если он там есть), без слэша — относительный от папки источника (xl/_rels/…,
        # т.е. база "xl/"). Раньше оба случая склеивались через "xl/" + lstrip("/"),
        # из-за чего абсолютный "/xl/worksheets/sheet1.xml" превращался в
        # "xl/xl/worksheets/sheet1.xml" и файл в архиве не находился.
        if sheet_path.startswith("/"):
            sheet_path = sheet_path.lstrip("/")
        else:
            sheet_path = "xl/" + sheet_path

        root = ET.fromstring(z.read(sheet_path))
        sheet_data = root.find("a:sheetData", NS)
        rows = list(sheet_data)
        if not rows:
            raise ValueError("Excel не содержит строк")

        def row_values(row):
            vals = {}
            for c in row.findall("a:c", NS):
                idx = col_index(c.attrib.get("r", "A1"))
                typ = c.attrib.get("t")
                v = c.find("a:v", NS)
                value = None if v is None else v.text
                if typ == "s" and value is not None:
                    value = shared[int(value)]
                elif typ == "inlineStr":
                    t = c.find(".//a:t", NS)
                    value = "" if t is None else (t.text or "")
                vals[idx] = value
            max_idx = max(vals.keys(), default=-1)
            return [vals.get(i) for i in range(max_idx + 1)]

        headers = row_values(rows[0])
        header_to_idx = {str(v).strip(): i for i, v in enumerate(headers) if v is not None}
        missing = [h for h in EXPECTED_HEADERS if h not in header_to_idx]
        if missing:
            raise ValueError("Не найдены обязательные колонки: " + ", ".join(missing))

        for row in rows[1:]:
            vals = row_values(row)

            def get(h):
                i = header_to_idx[h]
                return vals[i] if i < len(vals) else None

            production_no = str(get("Номер исполнительного производства") or "").strip()
            iin = re.sub(r"\D", "", str(get("ИИН/БИН должника") or ""))
            if not production_no:
                continue
            amount_raw = get("Сумма")
            try:
                amount = (
                    float(str(amount_raw).replace(" ", "").replace(",", "."))
                    if amount_raw not in (None, "") else None
                )
            except ValueError:
                amount = None
            yield {
                "document_no": str(get("Номер исполнительного документа") or "").strip(),
                "production_no": production_no,
                "executor": str(get("Судебный исполнитель") or "").strip(),
                "start_date": excel_date(get("Дата возбуждения")),
                "document_date": excel_date(get("Дата выписки исполнительного документа")),
                "issuer": str(get("Орган выдавший исполнительный документ") or "").strip(),
                "collection_type": str(get("Тип взыскания") or "").strip(),
                "debtor": str(get("Должник") or "").strip(),
                "iin": iin,
                "status": str(get("Статус исполнительного производства") or "").strip(),
                "amount": amount,
            }


def import_xlsx(path: str, source_file: str, company_id: int) -> int:
    now = datetime.now().isoformat(timespec="seconds")
    sql = """
    INSERT INTO productions(
        company_id, document_no, production_no, executor, start_date, document_date, issuer,
        collection_type, debtor, iin, status, amount, imported_at, source_file
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(company_id, production_no, iin) DO UPDATE SET
        document_no=excluded.document_no,
        executor=excluded.executor,
        start_date=excluded.start_date,
        document_date=excluded.document_date,
        issuer=excluded.issuer,
        collection_type=excluded.collection_type,
        debtor=excluded.debtor,
        status=excluded.status,
        amount=excluded.amount,
        imported_at=excluded.imported_at,
        source_file=excluded.source_file
    """
    batch, total = [], 0
    with db_connect() as conn:
        for rec in parse_xlsx(path):
            batch.append((
                company_id, rec["document_no"], rec["production_no"], rec["executor"],
                rec["start_date"], rec["document_date"], rec["issuer"], rec["collection_type"],
                rec["debtor"], rec["iin"], rec["status"], rec["amount"], now, source_file,
            ))
            if len(batch) >= 1000:
                conn.executemany(sql, batch)
                total += len(batch)
                batch.clear()
        if batch:
            conn.executemany(sql, batch)
            total += len(batch)
    return total


def build_where(company_id: int, qs: dict):
    clauses, params = ["company_id = ?"], [company_id]
    fields = {
        "iin": "iin", "bin": "iin", "debtor": "debtor", "document_no": "document_no",
        "production_no": "production_no", "executor": "executor", "status": "status", "issuer": "issuer",
    }
    for key, col in fields.items():
        val = (qs.get(key) or "").strip()
        if not val:
            continue
        if key in ("iin", "bin"):
            val = re.sub(r"\D", "", val)
            clauses.append(f"{col} = ?")
            params.append(val)
        else:
            clauses.append(f"UPPER({col}) LIKE UPPER(?)")
            params.append(f"%{val}%")
    date_from = (qs.get("date_from") or "").strip()
    date_to = (qs.get("date_to") or "").strip()
    if date_from:
        clauses.append("start_date >= ?")
        params.append(date_from)
    if date_to:
        clauses.append("start_date <= ?")
        params.append(date_to)
    return " WHERE " + " AND ".join(clauses), params


def get_stats(company_id: int):
    with db_connect() as conn:
        row = conn.execute(
            "SELECT COUNT(*) total, COUNT(DISTINCT iin) debtors, MAX(imported_at) updated "
            "FROM productions WHERE company_id = ?",
            (company_id,),
        ).fetchone()
        statuses = conn.execute(
            "SELECT status, COUNT(*) cnt FROM productions WHERE company_id = ? "
            "GROUP BY status ORDER BY cnt DESC",
            (company_id,),
        ).fetchall()
    return {
        "total": row["total"], "debtors": row["debtors"], "updated": row["updated"],
        "statuses": [dict(r) for r in statuses],
    }


def get_options(company_id: int):
    with db_connect() as conn:
        executors = [
            r[0] for r in conn.execute(
                "SELECT DISTINCT executor FROM productions WHERE company_id = ? AND executor<>'' ORDER BY executor",
                (company_id,),
            )
        ]
        statuses = [
            r[0] for r in conn.execute(
                "SELECT DISTINCT status FROM productions WHERE company_id = ? AND status<>'' ORDER BY status",
                (company_id,),
            )
        ]
    return {"executors": executors, "statuses": statuses}


def search(company_id: int, qs: dict, limit: int = 100):
    where, params = build_where(company_id, qs)
    with db_connect() as conn:
        rows = conn.execute(
            f"SELECT * FROM productions{where} ORDER BY start_date DESC, id DESC LIMIT ?",
            params + [limit],
        ).fetchall()
        count = conn.execute(f"SELECT COUNT(*) FROM productions{where}", params).fetchone()[0]
    return count, [dict(r) for r in rows]


def search_all(company_id: int, qs: dict):
    where, params = build_where(company_id, qs)
    with db_connect() as conn:
        rows = conn.execute(
            f"SELECT * FROM productions{where} ORDER BY start_date DESC, id DESC", params
        ).fetchall()
    return [dict(r) for r in rows]


# ═══════════════════════════════════════════════════════════════
# SELENIUM / АВТОРИЗАЦИЯ АИС ОИП (портировано из aisoip_sync.py)
# ═══════════════════════════════════════════════════════════════

def is_login_page(driver):
    url = (driver.current_url or "").lower()
    return (
        "authenticationendpoint" in url
        or "login.do" in url
        or len(driver.find_elements(By.CSS_SELECTOR, 'input[name="usernameUserInput"]')) > 0
    )


def to_site(download_dir: Path, log):
    """Запускает Chrome в headless-режиме — окно пользователю не показывается."""
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    prefs = {
        "profile.default_content_settings.popups": 0,
        "download.default_directory": str(download_dir.resolve()),
        "directory_upgrade": True,
        "profile.managed_default_content_settings.images": 2,
    }
    options.add_experimental_option("prefs", prefs)

    log("АИС ОИП: запускаю скрытую браузерную сессию (Chrome headless)…")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.implicitly_wait(5)
    driver.set_script_timeout(API_TIMEOUT_SECONDS)
    driver.get(AISOIP_URL)
    time.sleep(2)
    return driver


def write_login(driver, login: str, password: str, log):
    if not is_login_page(driver):
        log("АИС ОИП: активная авторизованная сессия уже есть.")
        return

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
    log("АИС ОИП: автоматическая авторизация выполнена.")


def reauthenticate(driver, login: str, password: str, log):
    """Переавторизация при истечении сессии (401 от API), как в otmeny.py."""
    log("АИС ОИП: сессия истекла (401) — повторная авторизация...")
    driver.get(AISOIP_URL)
    write_login(driver, login, password, log)
    driver.get(AISOIP_URL)
    WebDriverWait(driver, 60).until(
        lambda d: "aisoip.adilet.gov.kz" in (d.current_url or "").lower()
    )
    time.sleep(2)


def browser_fetch_json(driver, url, method="GET", body=None):
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
            f"API JSON error {result.get('status')} for {url}: {result.get('text', '')[:1000]}"
        )
    try:
        return json.loads(result["text"])
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"Ответ API не является JSON: {result['text'][:1000]}") from exc


def browser_fetch_bytes(driver, url):
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
        reader.onerror = () => done({status: 0, contentType: '', dataUrl: ''});
        reader.readAsDataURL(blob);
    })
    .catch(error => done({status: 0, contentType: '', dataUrl: String(error)}));
    """
    result = driver.execute_async_script(script, url)
    if result.get("status") != 200:
        raise RuntimeError(
            f"API export error {result.get('status')} for {url}: {str(result)[:1000]}"
        )

    data_url = result.get("dataUrl") or ""
    if "," not in data_url:
        raise RuntimeError(f"Экспорт не вернул бинарный файл: {str(result)[:1000]}")

    content = base64.b64decode(data_url.split(",", 1)[1])
    if not (content.startswith(b"PK") or content.startswith(bytes.fromhex("D0CF11E0"))):
        raise RuntimeError(
            f"Ответ экспорта не похож на Excel. "
            f"Content-Type={result.get('contentType')}, первые байты={content[:20]!r}"
        )
    return content


def api_call_with_retry(callable_obj, description: str, log, driver=None, login=None, password=None):
    last_error = None
    for attempt in range(1, API_RETRIES + 1):
        try:
            log(f"API: {description}. Попытка {attempt}/{API_RETRIES}")
            result = callable_obj()
            log(f"API: {description} — успешно")
            return result
        except Exception as exc:
            last_error = exc
            log(f"API ERROR: {description}. Попытка {attempt}/{API_RETRIES}: {exc}")

            if "401" in str(exc) and driver is not None and login and password:
                try:
                    reauthenticate(driver, login, password, log)
                except Exception as re_exc:
                    log(f"API: ошибка переавторизации: {re_exc}")
            elif attempt < API_RETRIES:
                time.sleep(2 * attempt)

    raise RuntimeError(f"Не удалось выполнить: {description}") from last_error


def create_mass_search(driver, date_from: str, date_to: str, log, login: str, password: str):
    payload = {"fromDate": date_from, "toDate": date_to, "searchType": False}
    url = f"{API_SEARCH_URL}?page=0&size=5"
    data = api_call_with_retry(
        lambda: browser_fetch_json(driver, url, method="POST", body=payload),
        "создание массового поиска АИС ОИП", log, driver=driver, login=login, password=password,
    )

    pagination = data.get("pagination") or {}
    search_id = pagination.get("searchId")
    total_elements = int(pagination.get("totalElements") or 0)
    if total_elements > 0 and not search_id:
        raise RuntimeError("API вернул строки, но не вернул pagination.searchId")

    log(f"Период API: {date_from} — {date_to}")
    log(f"Всего найдено в АИС ОИП: {total_elements}")
    log(f"searchId: {search_id}")
    return search_id, total_elements


def export_page(driver, search_id: str, page: int, log, login: str, password: str) -> bytes:
    url = (
        f"{API_EXPORT_URL}?searchtype=false"
        f"&page={page}&size={EXPORT_REQUEST_SIZE}"
        f"&searchid={search_id}&lang=ru"
    )
    return api_call_with_retry(
        lambda: browser_fetch_bytes(driver, url),
        f"экспорт страницы {page}", log, driver=driver, login=login, password=password,
    )


# ═══════════════════════════════════════════════════════════════
# МЕНЕДЖЕР СИНХРОНИЗАЦИИ (по одному на company_id, чтобы компании
# могли обновляться параллельно и независимо друг от друга)
# ═══════════════════════════════════════════════════════════════

@dataclass
class SyncState:
    running: bool = False
    started_at: Optional[str] = None
    finished_at: Optional[str] = None
    stage: str = "Готово"
    message: str = ""
    pages_total: int = 0
    pages_done: int = 0
    total_elements: int = 0
    rows_received: int = 0
    rows_imported: int = 0
    error: Optional[str] = None
    log_file: Optional[str] = None


class SyncManager:
    def __init__(self, company_id: int):
        self.company_id = company_id
        creds = COMPANY_AISOIP_CREDENTIALS.get(company_id)
        if not creds:
            raise RuntimeError(f"Нет учётных данных АИС ОИП для компании {company_id}")
        self.login = creds["login"]
        self.password = creds["password"]
        self.temp_dir = SYNC_TMP_DIR / f"company_{company_id}"
        self.temp_dir.mkdir(parents=True, exist_ok=True)
        self._lock = threading.Lock()
        self._state = SyncState()
        self._thread = None

    def state(self):
        with self._lock:
            return asdict(self._state)

    def _set(self, **kwargs):
        with self._lock:
            for k, v in kwargs.items():
                setattr(self._state, k, v)

    def start(self, date_from: str = DEFAULT_DATE_FROM):
        with self._lock:
            if self._state.running:
                return False, "Обновление уже выполняется"
            self._state = SyncState(
                running=True,
                started_at=datetime.now().isoformat(timespec="seconds"),
                stage="Подготовка",
            )
            self._thread = threading.Thread(
                target=self._run, args=(date_from,),
                name=f"aisoip-sync-company{self.company_id}", daemon=True,
            )
            self._thread.start()
        return True, "Обновление запущено"

    def _run(self, date_from: str):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_path = LOGS_DIR / f"sync_company{self.company_id}_{ts}.log"
        self._set(log_file=str(log_path.relative_to(BASE_DIR)))

        def log(msg):
            line = f"[{datetime.now().strftime('%d.%m.%Y %H:%M:%S')}] {msg}"
            print(f"[AISOIP company={self.company_id}] {line}")
            with log_path.open("a", encoding="utf-8") as f:
                f.write(line + "\n")
            self._set(message=msg)

        driver = None
        try:
            try:
                datetime.strptime(date_from, "%Y-%m-%d")
            except ValueError:
                date_from = DEFAULT_DATE_FROM
            date_to = date.today().strftime("%Y-%m-%d")

            self._set(stage="Авторизация")
            driver = to_site(self.temp_dir, log)
            write_login(driver, self.login, self.password, log)

            driver.get(AISOIP_URL)
            WebDriverWait(driver, 60).until(
                lambda d: "aisoip.adilet.gov.kz" in (d.current_url or "").lower()
            )

            self._set(stage="Массовый поиск")
            search_id, total_elements = create_mass_search(
                driver, date_from, date_to, log, self.login, self.password,
            )
            self._set(total_elements=total_elements)

            if total_elements == 0:
                self._set(
                    stage="Готово", running=False,
                    finished_at=datetime.now().isoformat(timespec="seconds"),
                )
                log("АИС ОИП не вернул строк за заданный период.")
                return

            self._set(stage="Экспорт страницы 1")
            first_content = export_page(driver, search_id, 0, log, self.login, self.password)
            first_path = self.temp_dir / f"sync_{ts}_page_0000.xlsx"
            first_path.write_bytes(first_content)
            first_count = import_xlsx(
                str(first_path), f"АИС ОИП API {date_from}—{date_to} page 0", self.company_id,
            )
            try:
                first_path.unlink()
            except OSError:
                pass

            if first_count <= 0:
                raise RuntimeError(
                    f"Первая страница массового экспорта пуста, хотя totalElements={total_elements}"
                )

            page_count = math.ceil(total_elements / first_count)
            rows_received = first_count
            rows_imported = first_count
            self._set(
                pages_total=page_count, pages_done=1,
                rows_received=rows_received, rows_imported=rows_imported,
            )
            log(f"Запрошенный размер страницы: {EXPORT_REQUEST_SIZE}")
            log(f"Фактический размер страницы: {first_count}")
            log(f"Количество частей: {page_count}")
            log(f"Страница 0: обработано {first_count} строк")

            for page in range(1, page_count):
                self._set(stage=f"Экспорт страницы {page + 1} из {page_count}")
                content = export_page(driver, search_id, page, log, self.login, self.password)
                part_path = self.temp_dir / f"sync_{ts}_page_{page:04d}.xlsx"
                part_path.write_bytes(content)
                imported = import_xlsx(
                    str(part_path),
                    f"АИС ОИП API {date_from}—{date_to} page {page}",
                    self.company_id,
                )
                try:
                    part_path.unlink()
                except OSError:
                    pass

                if imported <= 0 and rows_received < total_elements:
                    raise RuntimeError(
                        f"Страница {page} оказалась пустой до получения всех строк"
                    )

                rows_received += imported
                rows_imported += imported
                self._set(
                    pages_done=page + 1,
                    rows_received=rows_received, rows_imported=rows_imported,
                )
                log(
                    f"Страница {page}: обработано {imported} строк; "
                    f"суммарно {rows_received}/{total_elements}"
                )

            if rows_received < total_elements:
                raise RuntimeError(
                    f"Получено меньше строк, чем сообщил API: {rows_received} < {total_elements}"
                )

            self._set(
                stage="Готово", running=False,
                finished_at=datetime.now().isoformat(timespec="seconds"),
            )
            log(f"ОБНОВЛЕНИЕ ЗАВЕРШЕНО. Обработано строк: {rows_imported}")

        except Exception as exc:
            err = f"{type(exc).__name__}: {exc}"
            try:
                log("КРИТИЧЕСКАЯ ОШИБКА: " + err)
                with log_path.open("a", encoding="utf-8") as f:
                    f.write(traceback.format_exc() + "\n")
            except Exception:
                pass
            self._set(
                running=False,
                finished_at=datetime.now().isoformat(timespec="seconds"),
                stage="Ошибка", error=err,
            )
        finally:
            if driver is not None:
                try:
                    driver.quit()
                    log("Скрытая браузерная сессия закрыта.")
                except Exception:
                    pass


_SYNC_MANAGERS: dict[int, SyncManager] = {}
_SYNC_MANAGERS_LOCK = threading.Lock()


def get_sync_manager(company_id: int) -> SyncManager:
    with _SYNC_MANAGERS_LOCK:
        mgr = _SYNC_MANAGERS.get(company_id)
        if mgr is None:
            mgr = SyncManager(company_id)
            _SYNC_MANAGERS[company_id] = mgr
        return mgr
