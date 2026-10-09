# -*- coding: utf-8 -*-
"""
Подача исков в portal-sot.kz — подготовка БЕЗ ПОДПИСИ, по компаниям.

Порт ноутбука Portal_Sud_Podacha_Iska (PORTAL-SOT v54) в раннер вместо
scripts/podacha_iska_v2.py (office.sud.kz больше не работает).

Что делает по каждой строке листа «Отмены»:
  create (истец-ЮЛ, ответчик из ГБДФЛ, представитель) → загрузка госпошлины,
  иска и приложений → экран blank: текст иска в оба поля настоящим Ctrl+V →
  «Следующий шаг» до step=sign → «Мои дела» → «Ожидается отправка» → тот же
  черновик → текст в оба поля ещё раз → следующая сделка.
  ЭЦП НЕ накладывается — подписывает человек.

Сделки ведутся в реестре «Шаблоны документов/portal-sot обработанные сделки.xlsx»
(ключ — номер из колонки A): готовая сделка при повторном запуске пропускается,
недоделанная продолжается в том же declarationId — второй черновик не создаётся.

Нужны Word (текст иска берётся как Ctrl+A → Ctrl+C) и разблокированный рабочий
стол: окно браузера выводится на передний план, буфер обмена занят скриптом.

Всё, что в ноутбуке было захардкожено под Омегу, берётся по --company_id:
  COMPANY_ROOTS[id]              → папка «Документы для подачи Исков», Excel, партии
  COMPANY_CREDENTIALS[id]        → org_bin, org_bank, rep_iin
                                   (+ необяз. org_legal_address / org_display_address)
  PORTAL_SOT_BY_COMPANY[id]      → eds_password, portal_password, firefox_profile,
                                   cert_path (+ необяз. ncalayer_path)

Логика v24 сохранена: документы строго по 6-значному номеру из колонки A,
локальный справочник судов (без уголовных, Актау → CODE 194711), retry
загрузки файла, retry ГБДФЛ + перелогин, до 3 полных попыток строки, остановка
всей пачки при неудаче (NO-SKIP), пауза 30–60 сек между сделками.
"""

# ========= АРГУМЕНТЫ =========
import argparse
import os
import sys


def parse_args():
    p = argparse.ArgumentParser(description="Подача исков в portal-sot.kz (без подписи)")
    p.add_argument("--workdir",    type=str, default=None)
    p.add_argument("--company_id", type=str, default=None)
    p.add_argument("--excel_path", type=str, default=None)
    p.add_argument("--batch_dir",  type=str, default=None)
    p.add_argument("--start_row",  type=str, default=None)
    p.add_argument("--end_row",    type=str, default=None)
    return p.parse_known_args()[0]


args = parse_args()

if not args.company_id or not args.company_id.strip():
    print("❌ ОШИБКА: не передан --company_id — компания не определена, запуск остановлен.")
    sys.exit(1)

COMPANY_ID = args.company_id.strip()
# ДО импорта config — чтобы он выбрал ROOT/CREDENTIALS нужной компании.
os.environ["COMPANY_ID"] = COMPANY_ID
if args.excel_path and args.excel_path.strip():
    os.environ["OVERRIDE_EXCEL_PATH"] = args.excel_path.strip()

# ========= ИМПОРТЫ =========
import json
import mimetypes
import random
import re
import time
import traceback
import uuid
from datetime import datetime
from pathlib import Path

import ctypes
from ctypes import wintypes

import pythoncom
import requests
import win32clipboard
import win32com.client
import win32con
import win32gui
import win32process
from openpyxl import Workbook, load_workbook
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

sys.path.insert(0, str(Path(__file__).resolve().parent))
from portal_sot_login import PortalSotLogin, jwt_claims  # noqa: E402
from portal_sot_http import PortalBlocked, BLOCK_MESSAGE, is_waf_block, portal_lock  # noqa: E402
from config import (  # noqa: E402
    COMPANY_ROOTS, CREDENTIALS, MAIN_EXCEL, PORTAL_SOT_BY_COMPANY, ROOT,
)

# ============================================================
# НАСТРОЙКИ
# ============================================================

BASE_URL = "https://portal-sot.kz"

WORK_ROOT = Path(ROOT) / "Документы для подачи Исков"
EXCEL_SHEET = "Отмены"
COURTS_REL = Path("Документы для подачи Исков") / "Шаблоны документов" / "Справочник судов portal-sot.xlsx"
REGISTRY_FILE = WORK_ROOT / "Шаблоны документов" / "portal-sot обработанные сделки.xlsx"


def _int_arg(v, default):
    v = (v or "").strip()
    return int(v) if v.isdigit() else default


START_ROW = _int_arg(args.start_row, 2)
END_ROW = _int_arg(args.end_row, None)

# Данные иска, снятые из успешного HAR
CATEGORY_CODE = "142080004800000000"   # прочие исковые дела
CHARACTER_CODE = "62017001"            # имущественный
CASE_TYPE = "CIVIL"
INSTANCE_TYPE = "FIRSTINSTANCE"
SIMPLE_CASE = True

ORG_BIN = CREDENTIALS["org_bin"]
ORG_BANK = CREDENTIALS["org_bank"]
REP_IIN = CREDENTIALS["rep_iin"]

# Адреса истца: из COMPANY_CREDENTIALS (org_legal_address/org_display_address),
# иначе — известные значения (Омега из ноутбука), иначе — из ГБД ЮЛ.
_DEFAULT_ORG_ADDRESSES = {
    "1": (
        "город Алматы, Бостандыкский район, Проспект Аль-Фараби, дом 15, к 4В офис 1003",
        "КАЗАХСТАН, АЛМАТЫ, БОСТАНДЫКСКИЙ РАЙОН, ПРОСПЕКТ АЛЬ-ФАРАБИ, 15, к 4В офис 1003",
    ),
}
ORG_LEGAL_ADDRESS = CREDENTIALS.get("org_legal_address") or _DEFAULT_ORG_ADDRESSES.get(COMPANY_ID, ("", ""))[0]
ORG_DISPLAY_ADDRESS = CREDENTIALS.get("org_display_address") or _DEFAULT_ORG_ADDRESSES.get(COMPANY_ID, ("", ""))[1]

MAX_UPLOAD_FILE_BYTES = 20 * 1024 * 1024

# portal-sot: реквизиты входа по компании
PORTAL_CFG = PORTAL_SOT_BY_COMPANY.get(COMPANY_ID)
if not PORTAL_CFG:
    print(f"❌ Для компании {COMPANY_ID} нет записи в PORTAL_SOT_BY_COMPANY (scripts/config.py): "
          f"нужны ЭЦП-сертификат, пароли и Firefox-профиль с разрешением portal-sot.kz → NCALayer.")
    sys.exit(1)

FIREFOX_PROFILE_DIR = Path(PORTAL_CFG["firefox_profile"])

# ============================================================
# ЛОГИРОВАНИЕ
# ============================================================

OUT_DIR = Path(args.workdir) / "out" if args.workdir else WORK_ROOT
OUT_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = OUT_DIR / "portal-sot_run.log"
try:
    LOG_FILE.write_text("", encoding="utf-8")
except Exception:
    pass


def log(message="", level="INFO"):
    ts = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    line = f"[{ts}] [{level}] {message}"
    print(line, flush=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def log_exception(prefix="ОШИБКА"):
    text = traceback.format_exc()
    log(prefix, "ERROR")
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(text + "\n")
    except Exception:
        pass
    print(text, flush=True)


# ============================================================
# ОБЩИЕ ФУНКЦИИ
# ============================================================

def norm_text(v):
    return re.sub(r"\s+", " ", str(v or "").replace("ё", "е").replace("Ё", "Е")).strip().casefold()


def money(v):
    if v is None or v == "":
        return ""
    x = float(v)
    return f"{x:.2f}".rstrip("0").rstrip(".")


def normalize_iin(v):
    s = str(v or "").strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = re.sub(r"\D", "", s)
    return s.zfill(12) if s else ""


def read_docx_text(path: Path) -> str:
    """Текст иска ровно таким, каким его даёт ручное Ctrl+A → Ctrl+C в Word.

    python-docx отдавал другой текст (на одном иске 6561 символ против 6618 из
    Word), поэтому документ открывается в Word (невидимо, ReadOnly, без
    сохранения), основной текст копируется в буфер обмена и забирается оттуда.
    Если Word/буфер недоступны — ошибка, без отката на python-docx.
    """
    word = None
    doc = None
    pythoncom.CoInitialize()
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(
            str(Path(path).resolve()),
            ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False,
        )
        doc.StoryRanges(1).Copy()  # wdMainTextStory — то, что выделяет Ctrl+A в теле документа

        text = None
        last_error = None
        for _ in range(30):
            try:
                win32clipboard.OpenClipboard()
                try:
                    if win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
                        text = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
                        break
                finally:
                    win32clipboard.CloseClipboard()
            except Exception as e:
                last_error = e
            time.sleep(0.1)
        if text is None:
            raise RuntimeError(f"Word скопировал документ, но текст не получен из буфера обмена: {last_error}")

        # textarea в браузере использует LF; остальное (пустые строки, табы) не трогаем.
        return str(text).replace("\r\n", "\n").replace("\r", "\n").rstrip("\n")
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def api_data(payload):
    if not isinstance(payload, dict):
        return payload
    if payload.get("status") == 0 and "data" in payload:
        return payload["data"]
    if "result" in payload:
        return payload["result"]
    if "organization" in payload:
        return payload["organization"]
    return payload


# ============================================================
# РЕЕСТР ПОДГОТОВЛЕННЫХ СДЕЛОК
# Ключ — уникальный номер сделки (Excel, колонка A). По нему повторный запуск
# не создаёт второй черновик, а пропускает сделку или продолжает тот же.
# ============================================================

REGISTRY_HEADERS = [
    "UNIQUE_NO", "EXCEL_ROW", "IIN", "FIO",
    "DECLARATION_ID", "DOCUMENT_NUMBER",
    "STATUS", "TEXT_VERIFIED", "UPDATED_AT",
]
REGISTRY_READY_STATUSES = {"READY_TO_SIGN", "SIGNED", "SENT"}
REG_STAGE_FILES_DONE = "FILES_DONE"
REG_STAGE_FIRST_SIGN = "FIRST_SIGN_TRUSTED_PASTE_OK"
REG_STAGE_DONE = "MY_CASES_TRUSTED_PASTE_SAVED"
# READY_TO_SIGN, но проход 2 («Мои дела») ещё не сделан — доделываем только его.
REGISTRY_SECOND_PASS_PENDING = {
    REG_STAGE_FIRST_SIGN, "PASS1_DONE_MY_CASES_PASS2_PENDING", "SERVER_OK_PASS2_PENDING",
}


def _norm_unique_no(value):
    s = str(value or "").strip()
    return s[:-2] if s.endswith(".0") else s


def _registry_load():
    REGISTRY_FILE.parent.mkdir(parents=True, exist_ok=True)
    if not REGISTRY_FILE.exists():
        wb = Workbook()
        ws = wb.active
        ws.title = "Сделки"
        ws.append(REGISTRY_HEADERS)
        ws.freeze_panes = "A2"
        for col, width in zip("ABCDEFGHI", (16, 12, 18, 42, 18, 24, 20, 30, 22)):
            ws.column_dimensions[col].width = width
        wb.save(REGISTRY_FILE)
    wb = load_workbook(REGISTRY_FILE)
    ws = wb["Сделки"] if "Сделки" in wb.sheetnames else wb.active
    return wb, ws


def registry_check_writable():
    """До первого /create: реестр должен открываться на запись (не открыт в Excel)."""
    wb, _ws = _registry_load()
    wb.close()
    with open(REGISTRY_FILE, "r+b"):
        pass


def registry_get(unique_no):
    key = _norm_unique_no(unique_no)
    wb, ws = _registry_load()
    try:
        for r in range(2, ws.max_row + 1):
            if _norm_unique_no(ws.cell(r, 1).value) == key:
                return {
                    "UNIQUE_NO": key,
                    "EXCEL_ROW": ws.cell(r, 2).value,
                    "IIN": ws.cell(r, 3).value,
                    "FIO": ws.cell(r, 4).value,
                    "DECLARATION_ID": ws.cell(r, 5).value,
                    "DOCUMENT_NUMBER": ws.cell(r, 6).value,
                    "STATUS": str(ws.cell(r, 7).value or "").strip(),
                    "TEXT_VERIFIED": str(ws.cell(r, 8).value or "").strip(),
                    "UPDATED_AT": ws.cell(r, 9).value,
                }
        return None
    finally:
        wb.close()


def _registry_declaration_id(rec):
    s = _norm_unique_no((rec or {}).get("DECLARATION_ID"))
    return int(s) if s.isdigit() else None


def registry_upsert(unique_no, excel_row, iin, fio, declaration_id,
                    document_number="", status="READY_TO_SIGN", text_verified=""):
    key = _norm_unique_no(unique_no)
    wb, ws = _registry_load()
    try:
        target = next((r for r in range(2, ws.max_row + 1)
                       if _norm_unique_no(ws.cell(r, 1).value) == key), ws.max_row + 1)
        values = [
            key, excel_row, str(iin or ""), str(fio or ""),
            declaration_id, str(document_number or ""),
            status, str(text_verified or ""),
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        ]
        for c, value in enumerate(values, 1):
            ws.cell(target, c).value = value
        # Сначала временный файл, затем replace — чтобы не оставить битый реестр.
        tmp = REGISTRY_FILE.with_name(REGISTRY_FILE.stem + ".__tmp__.xlsx")
        wb.save(tmp)
    finally:
        wb.close()
    tmp.replace(REGISTRY_FILE)
    log(f"РЕЕСТР: №{key} | declarationId={declaration_id} | "
        f"номер документа={document_number or '-'} | {status} | {text_verified or '-'}")


# ============================================================
# 1. CHROME + АВТОЛОГИН ЧЕРЕЗ NCALAYER (как в poiskvsk.py — portal_sot_login.py)
# ============================================================

driver = None
session = None
tokens = {"access_token": None, "refresh_token": None}
org = {}
portal_login = None  # PortalSotLogin, создаётся в main()


def portal_sign_in(force_fresh=False):
    """Вход (или проверка живой сессии) + пересоздание API-сессии."""
    access, refresh = portal_login.login(driver, force_fresh=force_fresh)
    reload_api_session_from_browser(access, refresh)


# ============================================================
# 2. ТОКЕН ТЕКУЩЕЙ АВТОРИЗОВАННОЙ СЕССИИ
# ============================================================

def decode_storage():
    return driver.execute_script("""
        const out = {};
        for (let i=0; i<localStorage.length; i++) {
            const k = localStorage.key(i);
            out[k] = localStorage.getItem(k);
        }
        return out;
    """)


def find_tokens(obj):
    found = {"access_token": None, "refresh_token": None}

    def walk(x):
        if isinstance(x, dict):
            for k, v in x.items():
                lk = str(k).lower()
                if lk in ("access_token", "accesstoken", "access-token") and isinstance(v, str):
                    found["access_token"] = v
                if lk in ("refresh_token", "refreshtoken", "refresh-token") and isinstance(v, str):
                    found["refresh_token"] = v
                walk(v)
        elif isinstance(x, list):
            for v in x:
                walk(v)
        elif isinstance(x, str):
            st = x.strip()
            if st.startswith("{") or st.startswith("["):
                try:
                    walk(json.loads(st))
                except Exception:
                    pass
    walk(obj)
    return found


def reload_api_session_from_browser(access=None, refresh=None):
    """Берёт access/refresh token (из входа или localStorage) и пересоздаёт requests-сессию."""
    global tokens, session
    new_tokens = {"access_token": access, "refresh_token": refresh}
    if not access:
        storage = decode_storage()
        new_tokens = find_tokens(storage)
        # Дополнительный поиск JWT в значениях localStorage.
        if not new_tokens.get("access_token"):
            for _k, _v in storage.items():
                if isinstance(_v, str):
                    cand = re.findall(r"eyJ[a-zA-Z0-9_-]+\.[a-zA-Z0-9_-]+\.[a-zA-Z0-9_-]+", _v)
                    if cand:
                        new_tokens["access_token"] = cand[0]
                        if len(cand) > 1:
                            new_tokens["refresh_token"] = cand[-1]
                        break

    if not new_tokens.get("access_token"):
        raise RuntimeError("Access token не найден после входа на portal-sot.kz")

    tokens = new_tokens
    session = requests.Session()
    session.headers.update({
        "Accept": "*/*",
        "accept-language": "ru",
        "User-Agent": driver.execute_script("return navigator.userAgent"),
        "Authorization": "Bearer " + tokens["access_token"],
        "Origin": BASE_URL,
        "Referer": BASE_URL + "/cabinet",
    })


def refresh_access():
    rt = tokens.get("refresh_token")
    if not rt:
        return False
    try:
        _api_throttle()
        r = requests.post(
            BASE_URL + "/api/public/auth/refresh",
            json={"refreshToken": rt},
            headers={"Origin": BASE_URL, "Referer": BASE_URL + "/cabinet"},
            timeout=60,
        )
        if not r.ok:
            return False
        data = api_data(r.json())
    except Exception:
        return False
    at = data.get("access_token") if isinstance(data, dict) else None
    if not at:
        return False
    tokens["access_token"] = at
    tokens["refresh_token"] = data.get("refresh_token") or rt
    session.headers["Authorization"] = "Bearer " + at
    return True


def ensure_token_fresh(margin=90):
    """access_token живёт ~15 мин, пачка идёт дольше. Заранее: refresh по API,
    иначе — вход заново (живая сессия SPA → без ЭЦП, иначе полный ЭЦП-вход)."""
    exp = jwt_claims(tokens.get("access_token") or "").get("exp", 0) or 0
    if not exp or exp > time.time() + margin:
        return
    if refresh_access():
        log("access_token обновлён через refresh")
        return
    log("access_token истекает — обновляю сессию через браузер")
    portal_sign_in()


API_MIN_GAP_SECONDS = (1.0, 2.0)   # пауза между любыми двумя API-вызовами
API_BACKOFF_429 = (60, 180, 600)   # ожидания при 429; после последнего — стоп
_last_api_call_at = 0.0


def _api_throttle():
    global _last_api_call_at
    wait = _last_api_call_at + random.uniform(*API_MIN_GAP_SECONDS) - time.time()
    if wait > 0:
        time.sleep(wait)
    _last_api_call_at = time.time()


def _send(method, url, **kwargs):
    """Один HTTP-вызов с темпом: WAF-403 → стоп пачки, 429 → ждём и повторяем."""
    backoff = list(API_BACKOFF_429)
    while True:
        _api_throttle()
        r = session.request(method, url, timeout=120, **kwargs)
        if is_waf_block(r):
            raise PortalBlocked(BLOCK_MESSAGE)
        if r.status_code != 429:
            return r
        if not backoff:
            raise PortalBlocked("portal-sot.kz продолжает отвечать 429 — пачка остановлена")
        wait = backoff.pop(0)
        log(f"portal-sot.kz ответил 429 — жду {wait} сек.", "WARNING")
        time.sleep(wait)
        # файлы в multipart надо перемотать, иначе повтор уйдёт пустым
        for f in (kwargs.get("files") or {}).values():
            try:
                f[1].seek(0)
            except Exception:
                pass


def api_request(method, path, **kwargs):
    ensure_token_fresh()
    url = path if path.startswith("http") else BASE_URL + path
    r = _send(method, url, **kwargs)
    if r.status_code == 401:
        if not refresh_access():
            log("HTTP 401 — обновляю сессию через браузер")
            portal_sign_in()
        r = _send(method, url, **kwargs)
    if not r.ok:
        raise RuntimeError(f"{method} {url} -> HTTP {r.status_code}: {r.text[:1000]}")
    try:
        payload = r.json()
    except Exception:
        return r
    if isinstance(payload, dict):
        st = payload.get("status")
        if isinstance(st, dict):
            pass  # у ГБД ответов другая структура status
        elif st not in (None, 0):
            raise RuntimeError(f"API error: {payload}")
    return payload


def check_org():
    global org
    data = api_data(api_request("GET", f"/api/secure/gbdul/v2/byBin/{ORG_BIN}"))
    if not isinstance(data, dict):
        raise RuntimeError(f"ГБД ЮЛ не вернул данные организации по БИН {ORG_BIN}: {data}")
    org = data
    return org


MAX_FULL_RELOGINS = 2  # частые входы через ЭЦП — главный повод для блокировки
_full_relogins = 0


def force_relogin_portal():
    """Сначала — обновление токена по API (без ЭЦП). Полный вход через ЭЦП —
    только если refresh не сработал, и не больше MAX_FULL_RELOGINS раз за прогон."""
    global _full_relogins
    if refresh_access():
        check_org()
        log("ПОВТОРНЫЙ ВХОД не понадобился: токен обновлён через refresh — OK")
        return
    if _full_relogins >= MAX_FULL_RELOGINS:
        raise PortalBlocked(
            f"уже было {MAX_FULL_RELOGINS} полных входа через ЭЦП за прогон — "
            f"пачка остановлена, чтобы не вызвать блокировку"
        )
    _full_relogins += 1
    log(f"ПОВТОРНЫЙ ВХОД ({_full_relogins}/{MAX_FULL_RELOGINS}): сбрасываю web-сессию portal-sot и вхожу через ЭЦП...")
    portal_sign_in(force_fresh=True)
    check_org()
    log("ПОВТОРНЫЙ ВХОД: кабинет и API снова авторизованы — OK")


# ============================================================
# 3. EXCEL + ПАПКИ ДОЛЖНИКОВ
# ============================================================

EXCEL_FILE = Path(MAIN_EXCEL)
CASES_BATCH_DIR = None

COL_UNIQUE = "A"
COL_FIO = "C"
COL_IIN = "D"
COL_SUM = "J"
COL_DUTY = "K"
COL_REGION = "P"
COL_COURT = "Q"


def _batch_dt(name):
    m = re.match(r"^\s*(\d{2}\.\d{2}\.\d{4})(?:\s*\((\d{1,2})-(\d{2})\))?\s*$", name)
    if not m:
        return None
    d = datetime.strptime(m.group(1), "%d.%m.%Y")
    return d.replace(hour=int(m.group(2) or 0), minute=int(m.group(3) or 0))


def discover_batch_dir():
    if args.batch_dir and args.batch_dir.strip():
        p = Path(args.batch_dir.strip())
        if not p.is_dir():
            raise FileNotFoundError(f"Указанная папка партии не найдена: {p}")
        return p
    if not WORK_ROOT.exists():
        raise FileNotFoundError(f"Нет папки: {WORK_ROOT}")
    batches = [(_batch_dt(p.name), p) for p in WORK_ROOT.iterdir()
               if p.is_dir() and _batch_dt(p.name) is not None]
    if not batches:
        raise FileNotFoundError(f"Не найдены папки формата DD.MM.YYYY (HH-MM) в {WORK_ROOT}")
    return max(batches, key=lambda x: x[0])[1]


def load_case_row(row):
    wb = load_workbook(EXCEL_FILE, data_only=True, read_only=True)
    try:
        ws = wb[EXCEL_SHEET]
        unique_raw = ws[f"{COL_UNIQUE}{row}"].value
        unique_no = str(unique_raw or "").strip()
        if unique_no.endswith(".0"):
            unique_no = unique_no[:-2]
        unique_no = re.sub(r"\D", "", unique_no)
        if not re.fullmatch(r"\d{6}", unique_no):
            raise ValueError(
                f"Excel строка {row}: колонка A должна содержать уникальный номер из 6 цифр, получено {unique_raw!r}"
            )
        return {
            "ROW": row,
            "UNIQUE_NO": unique_no,
            "FIO": str(ws[f"{COL_FIO}{row}"].value or "").strip(),
            "IIN": normalize_iin(ws[f"{COL_IIN}{row}"].value),
            "CLAIM_SUM": money(ws[f"{COL_SUM}{row}"].value),
            "DUTY_SUM": money(ws[f"{COL_DUTY}{row}"].value),
            "REGION": str(ws[f"{COL_REGION}{row}"].value or "").strip(),
            "COURT": str(ws[f"{COL_COURT}{row}"].value or "").strip(),
        }
    finally:
        wb.close()


def rows_to_process():
    wb = load_workbook(EXCEL_FILE, data_only=True, read_only=True)
    try:
        ws = wb[EXCEL_SHEET]
        last = END_ROW or ws.max_row
        return [row for row in range(START_ROW, last + 1)
                if normalize_iin(ws[f"{COL_IIN}{row}"].value)]
    finally:
        wb.close()


def find_case_folder(unique_no, fio="", iin=""):
    """Папка сделки — по номеру из колонки A: имя оканчивается на №123456.
    Партии, собранные до «папки на каждый займ», называются 'ФИО, ИИН' без
    номера — для них берём папку по ИИН, но только если она ровно одна."""
    target = re.sub(r"\D", "", str(unique_no or ""))
    if not re.fullmatch(r"\d{6}", target):
        raise ValueError(f"Некорректный уникальный номер сделки: {unique_no!r}")
    hits = []
    by_iin = []
    for p in CASES_BATCH_DIR.iterdir():
        if not p.is_dir():
            continue
        m = re.search(r"№\s*(\d{6})\s*$", p.name)
        if m and m.group(1) == target:
            hits.append(p)
        elif iin and re.search(rf"(?<!\d){iin}(?!\d)", p.name):
            by_iin.append(p)
    if not hits:
        if len(by_iin) == 1 and "№" not in by_iin[0].name:
            log(f"Папка по ИИН {iin} (старая партия, без №): {by_iin[0].name}")
            return by_iin[0]
        if len(by_iin) > 1:
            raise ValueError(
                f"В {CASES_BATCH_DIR.name} нет папки №{target}, а по ИИН {iin} их несколько "
                f"(несколько займов): " + "; ".join(p.name for p in by_iin)
            )
        raise FileNotFoundError(f"В {CASES_BATCH_DIR.name} нет папки, оканчивающейся на №{target} ({fio}, {iin})")
    if len(hits) > 1:
        raise ValueError(f"Несколько папок с №{target}: " + "; ".join(p.name for p in hits))
    return hits[0]


class MissingDocsError(FileNotFoundError):
    """В папке займа нет обязательного документа — строку пропускаем, а не стопаем пачку."""
    def __init__(self, folder, missing):
        self.missing = missing
        super().__init__(f"Не найдено в {folder.name}: {', '.join(missing)}")


def classify_case_files(folder):
    allowed = {".pdf", ".doc", ".docx", ".jpg", ".jpeg", ".png"}
    files = sorted([p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in allowed],
                   key=lambda p: p.name.casefold())
    claim = None
    duty = None
    extras = []
    for p in files:
        low = norm_text(p.name)
        if p.suffix.lower() == ".docx" and "исков" in low:
            if claim:
                raise ValueError(f"Несколько исковых DOCX в {folder.name}")
            claim = p
        elif p.suffix.lower() == ".pdf" and "госпош" in low:
            if duty:
                raise ValueError(f"Несколько PDF госпошлины в {folder.name}")
            duty = p
        else:
            extras.append(p)
    missing = []
    if not claim:
        missing.append("исковое заявление (DOCX)")
    if not duty:
        missing.append("госпошлина (PDF)")
    if missing:
        raise MissingDocsError(folder, missing)
    for p in [claim, duty, *extras]:
        if p.stat().st_size > MAX_UPLOAD_FILE_BYTES:
            raise ValueError(f"Файл >20 МБ: {p.name}")
    return {"claim": claim, "duty": duty, "extras": extras}


# ============================================================
# 4. ЛОКАЛЬНЫЙ СПРАВОЧНИК СУДОВ — БЕЗ API СПРАВОЧНИКОВ PORTAL-SOT
# ============================================================
# - чисто уголовные суды исключаются;
# - Военный суд (18), Кассационный суд (23), упразднённые (99) исключаются;
# - для Актау гражданские иски принудительно идут в CODE 194711.

EXCLUDED_REGION_IDS = {18, 23, 99}
AKTAU_CIVIL_CODE = "194711"
LOCAL_COURTS = []


class CourtDataError(ValueError):
    pass


def find_courts_file():
    """Справочник судов общий для всех компаний: сначала в папке компании,
    иначе — в папке любой другой компании, где он есть."""
    roots = [ROOT] + [r for cid, r in COMPANY_ROOTS.items() if r != ROOT]
    for r in roots:
        p = Path(r) / COURTS_REL
        try:
            if p.exists():
                return p
        except OSError:
            continue
    raise FileNotFoundError(f"Не найден локальный справочник судов: {Path(ROOT) / COURTS_REL}")


def _court_kind(name):
    n = norm_text(name)
    if "уголовные дела" in n and "гражданские" not in n:
        return "criminal"
    if "гражданские дела" in n and "уголовные" not in n:
        return "civil"
    if "гражданские" in n and "уголовные" in n:
        return "mixed"
    if "общая юрисдикция" in n:
        return "general"
    return "other"


def _match_norm(v):
    s = str(v or "")
    # Убираем специализацию, включая оборванное "(Гражданские," из исходного Excel.
    s = s.split("(", 1)[0]
    s = norm_text(s)
    # Казахские варианты букв → русское написание (только для сопоставления).
    trans = str.maketrans({"ұ": "у", "ү": "у", "қ": "к", "ғ": "г", "ң": "н", "ө": "о", "ә": "а", "і": "и"})
    s = s.translate(trans)
    s = re.sub(r"(?:№|nº|no|n°)\s*", "n", s, flags=re.IGNORECASE)
    return s


def load_local_courts(courts_file):
    wb = load_workbook(courts_file, data_only=True, read_only=True)
    try:
        if "Все суды" not in wb.sheetnames:
            raise RuntimeError("В справочнике нет листа 'Все суды'")
        ws = wb["Все суды"]
        headers = {str(c.value or "").strip(): i for i, c in enumerate(ws[1], 1)}
        required = ["Регион", "ID региона", "CODE суда", "Наименование суда"]
        missing = [x for x in required if x not in headers]
        if missing:
            raise RuntimeError(f"В справочнике отсутствуют колонки: {missing}")

        rows = []
        for rr in ws.iter_rows(min_row=2, values_only=True):
            def val(col):
                return rr[headers[col] - 1] if headers[col] - 1 < len(rr) else None
            try:
                rid_int = int(val("ID региона"))
            except Exception:
                continue
            code = val("CODE суда")
            name = val("Наименование суда")
            if rid_int in EXCLUDED_REGION_IDS or not code or not name:
                continue
            if _court_kind(name) == "criminal":
                continue
            rows.append({
                "REGION": str(val("Регион") or "").strip(),
                "REGION_ID": rid_int,
                "CODE": str(code).strip(),
                "NAME": str(name).strip(),
            })
    finally:
        wb.close()

    if not rows:
        raise RuntimeError("После фильтрации локальный справочник судов пуст")
    return rows


def resolve_court(region_name, court_name):
    nr = _match_norm(region_name)
    source = _match_norm(court_name)

    region_rows = [c for c in LOCAL_COURTS if _match_norm(c["REGION"]) == nr]
    if not region_rows:
        raise CourtDataError(f"В локальном справочнике не найден регион: {region_name!r}")

    # Подтверждённое соответствие нового портала для гражданских дел Актау.
    if nr == _match_norm("Мангистауская область") and "актау" in source:
        hits = [c for c in region_rows if c["CODE"] == AKTAU_CIVIL_CODE]
        if len(hits) != 1:
            raise CourtDataError(f"В справочнике не найден гражданский суд Актау CODE={AKTAU_CIVIL_CODE}")
        court = hits[0]
        log(f"СУД АКТАУ: {court['NAME']} | CODE={court['CODE']}")
        return {"ID": court["REGION_ID"], "VALUE_RU": court["REGION"]}, court

    # 1) Полное совпадение названия без специализации.
    exact = [c for c in region_rows if _match_norm(c["NAME"]) == source]
    # 2) В старом Excel название иногда короче: без "... области".
    if not exact:
        exact = [c for c in region_rows
                 if source and (source in _match_norm(c["NAME"]) or _match_norm(c["NAME"]) in source)]

    priority = {"civil": 0, "general": 1, "mixed": 2, "other": 3}
    exact.sort(key=lambda c: priority.get(_court_kind(c["NAME"]), 99))

    if not exact:
        raise CourtDataError(
            f"СТОП ПО СУДУ: не найдено безопасное соответствие в локальном справочнике. "
            f"Регион={region_name!r}; суд из Excel={court_name!r}"
        )

    best_priority = priority.get(_court_kind(exact[0]["NAME"]), 99)
    best = [c for c in exact if priority.get(_court_kind(c["NAME"]), 99) == best_priority]
    if len(best) > 1:
        raise CourtDataError(
            f"СТОП ПО СУДУ: найдено несколько вариантов: {[(c['CODE'], c['NAME']) for c in best]}"
        )

    court = exact[0]
    log(f"СУД: {court['NAME']} | CODE={court['CODE']}")
    return {"ID": court["REGION_ID"], "VALUE_RU": court["REGION"]}, court


# ============================================================
# 5. ДАННЫЕ УЧАСТНИКОВ
# ============================================================

def person_from_gbdfl(iin):
    """ГБДФЛ: до 3 попыток; затем GBDFL_RELOGIN_REQUIRED → цикл перелогинится
    и повторит ЭТУ ЖЕ строку (declarationId на этом этапе ещё не создан)."""
    _last = None
    for _attempt in range(1, 4):
        try:
            payload = api_request("GET", f"/api/secure/gbdfl/v2/byIin/{iin}")
            p = api_data(payload)
            if not isinstance(p, dict) or not p.get("iin"):
                raise RuntimeError(f"ГБДФЛ не вернул данные по {iin}: {payload}")

            addr_parts = [
                p.get("regCountryNameRu"), p.get("regDistrictNameRu"),
                p.get("regRegionNameRu"), p.get("regCity"),
                p.get("regStreet"), p.get("regBuilding"),
            ]
            address = ", ".join(str(x).strip() for x in addr_parts if x not in (None, ""))
            if _attempt > 1:
                log(f"ГБДФЛ {iin}: попытка {_attempt}/3 — OK")
            return p, address

        except PortalBlocked:
            raise
        except Exception as _e:
            _last = _e
            _msg = str(_e)
            _low = _msg.lower()
            _temporary = (
                "timed out" in _low or "timeout" in _low or "epso_fl" in _low
                or re.search(r"HTTP\s+(500|502|503|504)\b", _msg, re.I) is not None
                or "internal server error" in _low
            )
            if not _temporary:
                raise
            log(f"ГБДФЛ {iin}: временная ошибка, попытка {_attempt}/3: {type(_e).__name__}: {_e}", "ERROR")
            if _attempt >= 3:
                raise RuntimeError(
                    f"GBDFL_RELOGIN_REQUIRED: ГБДФЛ {iin} не ответил после 3 попыток: {_e}"
                ) from _e
            _pause = random.randint(30, 45) if _attempt == 1 else random.randint(60, 90)
            log(f"ГБДФЛ {iin}: пауза {_pause} сек. перед повтором...")
            time.sleep(_pause)

    raise _last


def participant_fl(iin, process_side):
    p, address = person_from_gbdfl(iin)
    return {
        "processSide": process_side,
        "resident": True,
        "type": "FL",
        "iin": iin,
        "patronymic": p.get("patronymic") or "",
        "surname": p.get("lastName") or "",
        "firstname": p.get("firstName") or "",
        "address": address,
        "workplace": "",
        "phone": "",
        "email": "",
    }


def _org_address_from_gbdul():
    """Адрес истца из ГБД ЮЛ (organization.address — структурный), если в config
    не задан. Возвращает (legalAddress, address) в формате, как у Омеги:
      'город Алматы, Алмалинский район, Проспект АБЫЛАЙ ХАНА, дом 58'
      'КАЗАХСТАН, АЛМАТЫ, АЛМАЛИНСКИЙ РАЙОН, ПРОСПЕКТ АБЫЛАЙ ХАНА, 58'"""
    a = org.get("address") if isinstance(org.get("address"), dict) else {}

    def v(key):
        s = str(a.get(key) or "").strip()
        return "" if s in ("-", "—") else s

    district, region, rural, city, street = (v("districtRu"), v("regionRu"), v("ruralRu"),
                                             v("cityRu"), v("streetRu"))
    building = v("buildingNumber")
    btype = str(((a.get("buildingType") or {}).get("nameRu")) or "дом").strip()
    tail = []  # корпус/блок/офис/квартира
    if v("corpus"):
        tail.append(f"корпус {v('corpus')}")
    if v("block"):
        tail.append(f"блок {v('block')}")
    if v("officeNumber"):
        tail.append(f"офис {v('officeNumber')}")
    elif v("appartmentNumber"):
        tail.append(f"кв. {v('appartmentNumber')}")
    tail_s = " ".join(tail)

    if district and street and building:
        legal = ", ".join(x for x in [district, region, rural, city, street,
                                      f"{btype} {building}", tail_s] if x)
        short_district = re.sub(r"^(город|г\.)\s+", "", district, flags=re.I)
        display = ", ".join(x for x in ["КАЗАХСТАН", short_district.upper(), region.upper(),
                                        rural.upper(), city.upper(), street.upper(), building,
                                        tail_s] if x)
        return legal, display

    # Запасной вариант — фактический адрес одной строкой.
    fact = (((org.get("statCommInfo") or {}).get("addressFact") or {}).get("nameRu") or "").strip()
    return fact, fact


def plaintiff_participant():
    name = org.get("fullNameRu")
    if not name:
        raise RuntimeError(f"ГБД ЮЛ не вернул fullNameRu для БИН {ORG_BIN}")
    legal, display = ORG_LEGAL_ADDRESS, ORG_DISPLAY_ADDRESS
    if not legal:
        legal, gb_display = _org_address_from_gbdul()
        display = display or gb_display
    display = display or legal
    if not legal:
        log("Ответ ГБД ЮЛ (адреса не нашлось): "
            + json.dumps(org, ensure_ascii=False)[:3000], "ERROR")
        raise ValueError(
            f"Не задан адрес истца для компании {COMPANY_ID}: добавьте "
            f"'org_legal_address' / 'org_display_address' в COMPANY_CREDENTIALS['{COMPANY_ID}']"
        )
    return {
        "processSide": "61090001",
        "resident": True,
        "type": "UL",
        "bin": ORG_BIN,
        "name": name,
        "legalAddress": legal,
        "address": display,
        "bankDetails": ORG_BANK,
        "kbk": "",
        "kno": "",
        "knp": "",
        "govOrganization": False,
    }


# ============================================================
# 6. API СОЗДАНИЯ + ЗАГРУЗКИ + BLANK
# ============================================================

def create_declaration(ctx):
    _, court = resolve_court(ctx["REGION"], ctx["COURT"])
    debtor = participant_fl(ctx["IIN"], "61090002")
    representative = participant_fl(REP_IIN, "61090005")

    payload = {
        "declarationId": None,
        "categoryCode": CATEGORY_CODE,
        "characterCode": CHARACTER_CODE,
        "courtCode": str(court["CODE"]),
        "simpleCase": SIMPLE_CASE,
        "participants": [plaintiff_participant(), debtor, representative],
        "caseType": CASE_TYPE,
        "instanceType": INSTANCE_TYPE,
    }
    data = api_data(api_request("POST", "/api/secure/declaration/REQUEST_TYPE2/create", json=payload))
    if not isinstance(data, dict) or not data.get("id"):
        raise RuntimeError(f"Не получен declarationId: {data}")
    return data


def upload_file(declaration_id, path, file_type, participant_id=None):
    """Повторяет загрузку только текущего файла до 3 раз при временных отказах portal-sot."""
    def _one_attempt():
        endpoint = f"/api/secure/declaration/attachment/saveAttachment/{declaration_id}"
        if participant_id is not None:
            endpoint += f"/{participant_id}"
        mime = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        with open(path, "rb") as f:
            files = {"file": (path.name, f, mime)}
            form = {"type": file_type, "filename": path.name}
            return api_request("POST", endpoint, files=files, data=form)

    _last_error = None
    for _attempt in range(1, 4):
        try:
            if _attempt > 1:
                log(f"ПОВТОР ЗАГРУЗКИ: попытка {_attempt}/3")
            return _one_attempt()
        except PortalBlocked:
            raise
        except Exception as _e:
            _last_error = _e
            _msg = str(_e)
            _low = _msg.lower()
            _retryable = (
                isinstance(_e, (requests.exceptions.Timeout, requests.exceptions.ConnectionError))
                or re.search(r"HTTP\s+(500|502|503|504)\b", _msg, re.I) is not None
                or "internal server error" in _low
                or "пустой или поврежд" in _low
            )
            if not _retryable:
                raise
            log(f"ВРЕМЕННАЯ ОШИБКА ЗАГРУЗКИ, попытка {_attempt}/3: {type(_e).__name__}: {_e}")
            if _attempt >= 3:
                log("Файл не принят после 3 попыток.")
                raise RuntimeError(f"UPLOAD_RETRY_EXHAUSTED: файл не загружен после 3 попыток: {_e}") from _e
            _pause = random.randint(20, 30) if _attempt == 1 else random.randint(45, 60)
            log(f"Пауза {_pause} сек. перед повтором ЭТОГО ЖЕ файла...")
            time.sleep(_pause)

    raise _last_error


def get_declaration(declaration_id):
    return api_data(api_request(
        "GET", f"/api/secure/declaration/REQUEST_TYPE2/{declaration_id}",
        headers={"Content-Type": "application/json"},
    ))


def open_blank_page(declaration_id):
    url = (f"{BASE_URL}/cabinet/declarations/REQUEST_TYPE2"
           f"?step=blank&caseType={CASE_TYPE}&instanceType={INSTANCE_TYPE}"
           f"&uuid={uuid.uuid4()}&id={declaration_id}")
    driver.get(url)
    return url


# ============================================================
# 7. ЭКРАН BLANK: «НАСТОЯЩАЯ» ВСТАВКА CTRL+V (v54)
# ============================================================
# Портал сохраняет текст двух полей только после доверенного (trusted) ввода:
# значение, выставленное через JS, и прямой POST /blank из Python он теряет.
# Поэтому текст кладём в буфер Windows и жмём Ctrl+A / Ctrl+V через SendInput
# в окне браузера — как человек. Окно браузера при этом должно быть на экране,
# а буфером обмена на этой машине во время пачки пользоваться нельзя.

BLANK_FIELDS = ("circumstances", "requirement")
BROWSER_WINDOW_CLASSES = ("MozillaWindowClass", "Chrome_WidgetWin_1")

_INPUT_KEYBOARD = 1
_KEYEVENTF_KEYUP = 0x0002
_VK_CONTROL, _VK_A, _VK_V = 0x11, 0x41, 0x56
_ULONG_PTR = ctypes.c_ulonglong if ctypes.sizeof(ctypes.c_void_p) == 8 else ctypes.c_ulong


class _MOUSEINPUT(ctypes.Structure):
    _fields_ = [("dx", wintypes.LONG), ("dy", wintypes.LONG), ("mouseData", wintypes.DWORD),
                ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD), ("dwExtraInfo", _ULONG_PTR)]


class _KEYBDINPUT(ctypes.Structure):
    _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD), ("dwFlags", wintypes.DWORD),
                ("time", wintypes.DWORD), ("dwExtraInfo", _ULONG_PTR)]


class _HARDWAREINPUT(ctypes.Structure):
    _fields_ = [("uMsg", wintypes.DWORD), ("wParamL", wintypes.WORD), ("wParamH", wintypes.WORD)]


class _INPUT_UNION(ctypes.Union):
    _fields_ = [("mi", _MOUSEINPUT), ("ki", _KEYBDINPUT), ("hi", _HARDWAREINPUT)]


class _INPUT(ctypes.Structure):
    _anonymous_ = ("u",)
    _fields_ = [("type", wintypes.DWORD), ("u", _INPUT_UNION)]


def _send_ctrl_combo(vk, label):
    user32 = ctypes.windll.user32
    user32.SendInput.argtypes = (wintypes.UINT, ctypes.POINTER(_INPUT), ctypes.c_int)
    user32.SendInput.restype = wintypes.UINT

    def key(code, up=False):
        x = _INPUT()
        x.type = _INPUT_KEYBOARD
        x.ki.wVk = code
        x.ki.dwFlags = _KEYEVENTF_KEYUP if up else 0
        return x

    seq = (_INPUT * 4)(key(_VK_CONTROL), key(vk), key(vk, True), key(_VK_CONTROL, True))
    n = user32.SendInput(4, seq, ctypes.sizeof(_INPUT))
    if n != 4:
        raise RuntimeError(f"{label}: SendInput {n}/4")


def _set_clipboard_text(value):
    last = None
    for _ in range(20):
        try:
            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardText(value, win32con.CF_UNICODETEXT)
            finally:
                win32clipboard.CloseClipboard()
            return
        except Exception as e:
            last = e
            time.sleep(0.1)
    raise RuntimeError(f"Не удалось записать текст в буфер обмена Windows: {last}")


def _browser_hwnd(label):
    """Окно браузера, которым управляет driver: по PID процесса, иначе по заголовку."""
    wins = []

    def enum_cb(hwnd, _):
        try:
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetClassName(hwnd) in BROWSER_WINDOW_CLASSES:
                title = win32gui.GetWindowText(hwnd)
                if title:
                    wins.append((hwnd, title, win32process.GetWindowThreadProcessId(hwnd)[1]))
        except Exception:
            pass

    win32gui.EnumWindows(enum_cb, None)
    if not wins:
        raise RuntimeError(f"{label}: окно браузера не найдено")
    pid = (driver.capabilities or {}).get("moz:processID")
    title = driver.title or ""
    return (next((h for h, _t, p in wins if pid and p == pid), None)
            or next((h for h, t, _p in wins if title and title in t), None)
            or wins[0][0])


def _activate_browser_window(label):
    user32 = ctypes.windll.user32
    target = _browser_hwnd(label)
    cur = ctypes.windll.kernel32.GetCurrentThreadId()
    fg_tid = user32.GetWindowThreadProcessId(user32.GetForegroundWindow(), None)
    target_tid = user32.GetWindowThreadProcessId(target, None)
    a = b = False
    try:
        if fg_tid and fg_tid != cur:
            a = bool(user32.AttachThreadInput(cur, fg_tid, True))
        if target_tid and target_tid != cur:
            b = bool(user32.AttachThreadInput(cur, target_tid, True))
        win32gui.ShowWindow(target, win32con.SW_RESTORE)
        user32.BringWindowToTop(target)
        user32.SetForegroundWindow(target)
        user32.SetActiveWindow(target)
        user32.SetFocus(target)
    finally:
        if b:
            user32.AttachThreadInput(cur, target_tid, False)
        if a:
            user32.AttachThreadInput(cur, fg_tid, False)
    time.sleep(0.45)
    if user32.GetForegroundWindow() != target:
        raise RuntimeError(f"{label}: окно браузера не стало активным (сеанс рабочего стола заблокирован/свёрнут?)")


def _blank_field(name, timeout=60):
    def finder(d):
        els = d.find_elements(By.CSS_SELECTOR, f'textarea[name="{name}"]')
        return els[0] if els and els[0].is_displayed() and els[0].is_enabled() else False
    return WebDriverWait(driver, timeout).until(finder)


def _blank_field_state(name):
    return driver.execute_script("""
      const el=document.querySelector(`textarea[name="${arguments[0]}"]`);
      const k=Object.keys(el).find(k=>k.startsWith('__reactProps'));
      const p=k?el[k]:null;
      const tr=el._valueTracker&&el._valueTracker.getValue?el._valueTracker.getValue():null;
      return {valueLen:el.value.length,
              trackerLen:tr==null?null:String(tr).length,
              reactLen:p&&typeof p.value==='string'?p.value.length:null};
    """, name)


def trusted_paste(name, value, label):
    """Ctrl+A, Ctrl+V в textarea[name] и проверка, что текст дошёл до DOM и до состояния React."""
    _set_clipboard_text(value)
    _activate_browser_window(label)

    el = _blank_field(name)
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.3)
    el.click()
    time.sleep(0.25)
    if not driver.execute_script("return document.activeElement===arguments[0];", el):
        raise RuntimeError(f"{label}: поле {name} не получило фокус")

    _send_ctrl_combo(_VK_A, label)
    time.sleep(0.15)
    _send_ctrl_combo(_VK_V, label)

    expected = len(value)
    state = {}
    deadline = time.time() + 5
    while time.time() < deadline:
        time.sleep(0.4)
        state = _blank_field_state(name)
        if state["valueLen"] == state["reactLen"] == state["trackerLen"] == expected:
            break
    log(f"{label}: ВСТАВКА {name} | DOM={state['valueLen']} | React={state['reactLen']} | tracker={state['trackerLen']}")
    if not (state["valueLen"] == state["reactLen"] == state["trackerLen"] == expected):
        raise RuntimeError(f"{label}: {name} — вставилось не то: ожидалось {expected} символов, получено {state}")


def paste_both_fields(claim_text, label):
    for name in BLANK_FIELDS:
        trusted_paste(name, claim_text, label)
    vals = [_blank_field(name).get_attribute("value") or "" for name in BLANK_FIELDS]
    if any(v != claim_text for v in vals):
        raise RuntimeError(f"{label}: после вставки текст в полях не совпал с текстом иска")


def click_next_to_sign(declaration_id, label):
    """«Следующий шаг» → ждём step=sign. НЕ подписывает."""
    def find_next_button(d):
        for b in d.find_elements(
            By.XPATH,
            "//button[normalize-space(.)='Следующий шаг'] | //a[normalize-space(.)='Следующий шаг']",
        ):
            try:
                if b.is_displayed() and b.is_enabled():
                    return b
            except Exception:
                pass
        return False

    def scroll_and_click_next():
        b = WebDriverWait(driver, 20).until(find_next_button)
        driver.execute_script(
            "arguments[0].scrollIntoView({behavior:'auto', block:'center', inline:'nearest'});", b
        )
        time.sleep(1.2)
        b = WebDriverWait(driver, 10).until(find_next_button)
        try:
            b.click()
        except Exception:
            driver.execute_script("arguments[0].click();", b)

    # Портал после «Следующий шаг» сохраняет текст и проверяет вложения — с
    # 15+ файлами это бывает дольше минуты. Раньше ждали 10 сек и кликали
    # снова: переход на sign случался позже, а скрипт уже считал шаг
    # проваленным. Поэтому ждём долго и повторный клик — только если кнопка
    # всё ещё на экране (запрос явно не ушёл).
    MAX_NEXT_CLICKS = 3
    NEXT_WAIT_SEC = 120
    for click_attempt in range(1, MAX_NEXT_CLICKS + 1):
        if "step=sign" in driver.current_url:
            break
        if click_attempt > 1 and not find_next_button(driver):
            log(f"{label}: кнопки «Следующий шаг» нет, но и step=sign нет — жду ещё...")
        else:
            log(f"{label}: «Следующий шаг», попытка {click_attempt}/{MAX_NEXT_CLICKS} | declarationId={declaration_id}")
            scroll_and_click_next()
        try:
            WebDriverWait(driver, NEXT_WAIT_SEC).until(lambda d: "step=sign" in d.current_url)
            break
        except TimeoutException:
            log(f"{label}: переход на step=sign не подтверждён за {NEXT_WAIT_SEC} сек.")

    if "step=sign" not in driver.current_url:
        portal_login.dump_page(driver, f"no_sign_{declaration_id}")
        raise RuntimeError(
            f"NEXT_BUTTON_RETRY_EXHAUSTED: declarationId={declaration_id} остался на шаге blank"
        )
    log(f"{label}: step=sign подтверждён")


def wait_sign_page_loaded():
    """С экрана подписи нельзя уходить сразу: ждём, пока бланк реально отрисуется."""
    WebDriverWait(driver, 60).until(
        lambda d: "step=sign" in d.current_url
        and d.execute_script("return document.readyState") == "complete"
    )
    WebDriverWait(driver, 90).until(lambda d: d.execute_script("""
        const visible = el => {
          if(!el) return false;
          const s=getComputedStyle(el), r=el.getBoundingClientRect();
          return s.display!=='none' && s.visibility!=='hidden' && r.width>0 && r.height>0;
        };
        const loaders=[...document.querySelectorAll(
          '.ant-spin-spinning,.ant-spin,.spinner,.loader,[class*="loading"],[class*="spinner"]'
        )].filter(visible);
        const body=(document.body?.innerText||'').trim();
        const contentReady = body.length > 150 ||
          !!document.querySelector('iframe,embed,object,canvas,[class*="pdf"],[class*="document"],[class*="blank"]');
        return loaders.length===0 && contentReady;
    """))
    time.sleep(2.0)


def open_draft_from_my_cases(declaration_id):
    """«Мои дела» → «Ожидается отправка» → этот же черновик на шаге blank."""
    driver.get(BASE_URL)
    WebDriverWait(driver, 30).until(lambda d: d.execute_script("return document.readyState") == "complete")
    my = WebDriverWait(driver, 30).until(lambda d: next(
        (x for x in d.find_elements(
            By.XPATH, "//*[self::a or self::button][contains(normalize-space(.),'Мои дела')]")
         if x.is_displayed()), None))
    driver.execute_script("arguments[0].click();", my)
    WebDriverWait(driver, 40).until(lambda d: "Мои дела" in (d.execute_script(
        "return document.body ? document.body.innerText : ''") or ""))
    time.sleep(1)

    wait_tab = WebDriverWait(driver, 30).until(lambda d: next(
        (x for x in d.find_elements(
            By.XPATH,
            "//*[self::a or self::button or @role='tab'][contains(normalize-space(.),'Ожидается отправка')]")
         if x.is_displayed()), None))
    driver.execute_script("arguments[0].click();", wait_tab)
    time.sleep(1.5)
    log("Открыт раздел «Мои дела» → «Ожидается отправка»")

    # Ровно тот URL, который портал открывает по клику на черновик в списке
    # (id в кавычках) — искать строку в таблице по номеру документа не нужно.
    driver.get(
        f"{BASE_URL}/cabinet/declarations/REQUEST_TYPE2"
        f"?step=blank&caseType={CASE_TYPE}&instanceType={INSTANCE_TYPE}"
        f"&id=%22{declaration_id}%22"
    )
    WebDriverWait(driver, 60).until(
        lambda d: str(declaration_id) in d.current_url
        and all(d.find_elements(By.CSS_SELECTOR, f'textarea[name="{n}"]') for n in BLANK_FIELDS)
    )
    WebDriverWait(driver, 30).until(lambda d: d.execute_script("return document.readyState") == "complete")
    time.sleep(1.5)
    log(f"Черновик открыт из «Мои дела» | declarationId={declaration_id}")


# ============================================================
# 8. ОДНА СДЕЛКА: СОЗДАТЬ ИСК ДО ЭКРАНА ПОДПИСАНИЯ
# ============================================================

_uploaded_files = {}  # declarationId -> имена файлов, загруженных в ЭТОМ запуске


def _names_in_declaration(obj, names):
    """Какие из имён файлов уже встречаются в JSON заявления (структура вложений не фиксирована)."""
    found = set()

    def walk(x):
        if isinstance(x, dict):
            for v in x.values():
                walk(v)
        elif isinstance(x, list):
            for v in x:
                walk(v)
        elif isinstance(x, str) and x in names:
            found.add(x)
    walk(obj)
    return found


def upload_case_files(declaration_id, plaintiff_id, files, already=frozenset()):
    done = _uploaded_files.setdefault(declaration_id, set())
    extras = files["extras"]
    plan = [(files["duty"], "OFFLINE_PAYMENT_CHECK", plaintiff_id, "госпошлина (OFFLINE_PAYMENT_CHECK)"),
            (files["claim"], "MAIN_DECLARATION_FILE", None, "иск")]
    plan += [(p, "ADDITIONAL_FILE", None, f"приложение {i}/{len(extras)}") for i, p in enumerate(extras, 1)]
    for p, file_type, participant_id, label in plan:
        if p.name in done or p.name in already:
            log(f"УЖЕ ЗАГРУЖЕН, пропускаю: {label}: {p.name}")
            continue
        size_bytes = p.stat().st_size if p.exists() else -1
        log(f"ЗАГРУЗКА {label}: {p.name} | {size_bytes} байт ({size_bytes / 1024:.1f} КБ)")
        try:
            upload_file(declaration_id, p, file_type, participant_id)
        except Exception as e:
            log(f"ОШИБКА НА ФАЙЛЕ {label}: {p.name} | путь: {p} | "
                f"размер: {size_bytes} байт | {type(e).__name__}: {e}", "ERROR")
            raise
        done.add(p.name)
        log(f"OK: {label}: {p.name}")


def _plaintiff_id(decl):
    plaintiffs = [p for p in decl.get("participants", []) if str(p.get("type")) == "61090001"]
    if len(plaintiffs) != 1:
        raise RuntimeError("Не удалось определить participantId истца")
    return plaintiffs[0]["id"]


def run_one(row):
    ctx = load_case_row(row)
    unique_no = ctx["UNIQUE_NO"]

    rec = registry_get(unique_no) or {}
    status = rec.get("STATUS", "").upper()
    stage = rec.get("TEXT_VERIFIED", "")
    declaration_id = _registry_declaration_id(rec)
    document_number = str(rec.get("DOCUMENT_NUMBER") or "")
    is_ready = status in REGISTRY_READY_STATUSES
    second_pass_only = bool(is_ready and declaration_id and stage in REGISTRY_SECOND_PASS_PENDING)

    if is_ready and not second_pass_only:
        log(f"ПРОПУСК №{unique_no}: уже {status} | declarationId={declaration_id or '-'} | "
            f"номер документа={document_number or '-'}")
        return {"declaration_id": declaration_id, "document_number": document_number,
                "unique_no": unique_no, "skipped": True}

    folder = find_case_folder(ctx["UNIQUE_NO"], ctx["FIO"], ctx["IIN"])
    files = classify_case_files(folder)

    log("=" * 80)
    log(f"СТРОКА EXCEL: {row}")
    log(f"УНИКАЛЬНЫЙ № СДЕЛКИ: {unique_no}")
    log(f"ФИО: {ctx['FIO']}")
    log(f"ИИН: {ctx['IIN']}")
    log(f"Суд: {ctx['COURT']}")
    log(f"Сумма иска: {ctx['CLAIM_SUM']}")
    log(f"Госпошлина: {ctx['DUTY_SUM']}")
    log(f"Папка должника: {folder}")
    log(f"Иск: {files['claim'].name}")
    log(f"Файл госпошлины: {files['duty'].name}")
    log(f"Приложений: {len(files['extras'])}")

    # Текст иска — до /create: если Word не отдал текст, черновик не плодим.
    claim_text = read_docx_text(files["claim"])
    if not claim_text.strip():
        raise ValueError(f"Текст DOCX иска пустой: {files['claim'].name}")
    log(f"Текст иска (как Ctrl+A → Ctrl+C в Word): {len(claim_text)} символов | "
        f"строк={claim_text.count(chr(10)) + 1}")

    def reg(new_status, new_stage):
        registry_upsert(unique_no, row, ctx["IIN"], ctx["FIO"], declaration_id,
                        document_number, new_status, new_stage)

    resumed = bool(declaration_id)
    if not second_pass_only:
        if resumed:
            log(f"ПРОДОЛЖАЮ существующий черновик №{unique_no}: declarationId={declaration_id} "
                f"({status or '-'} / {stage or '-'}). Новый /create НЕ делаю.")
            if stage != REG_STAGE_FILES_DONE:
                decl = get_declaration(declaration_id)
                names = {p.name for p in [files["duty"], files["claim"], *files["extras"]]}
                already = _names_in_declaration(decl, names)
                if not already and declaration_id not in _uploaded_files:
                    raise DraftStuckError(
                        f"declarationId={declaration_id} (№{unique_no}) создан раньше, но неизвестно, какие "
                        f"файлы в него загружены. Проверьте черновик в кабинете: либо удалите его и строку "
                        f"№{unique_no} из реестра «{REGISTRY_FILE.name}», либо догрузите файлы вручную."
                    )
                upload_case_files(declaration_id, _plaintiff_id(decl), files, already)
                reg("CREATED_IN_PROGRESS", REG_STAGE_FILES_DONE)
        else:
            # 1. Создание
            decl = create_declaration(ctx)
            declaration_id = decl["id"]
            log(f"Создан declarationId = {declaration_id}")
            # ID фиксируем СРАЗУ: если дальше что-то упадёт, повтор этой сделки
            # продолжит этот declarationId и не создаст второй черновик.
            try:
                reg("CREATED_IN_PROGRESS", "FILES_PENDING")
            except Exception as e:
                raise DraftStuckError(
                    f"declarationId={declaration_id} (№{unique_no}) создан, но не записан в реестр "
                    f"«{REGISTRY_FILE}» ({type(e).__name__}: {e}) — повтор создал бы дубль"
                ) from e
            # 2. Госпошлина (к истцу), иск, приложения
            upload_case_files(declaration_id, _plaintiff_id(decl), files)
            reg("CREATED_IN_PROGRESS", REG_STAGE_FILES_DONE)

        # 3. Проход 1: blank → текст в оба поля → «Следующий шаг» → step=sign.
        log(f"ПРОХОД 1: открываю step=blank | declarationId={declaration_id}")
        open_blank_page(declaration_id)
        current = [_blank_field(name).get_attribute("value") or "" for name in BLANK_FIELDS]
        if resumed and all(current):
            log(f"ПРОХОД 1: текст уже в черновике ({len(current[0])} / {len(current[1])} символов)")
        else:
            paste_both_fields(claim_text, "ПРОХОД 1")
        click_next_to_sign(declaration_id, "ПРОХОД 1")
        wait_sign_page_loaded()
        log("ПРОХОД 1: бланк подписи полностью загружен")

        try:
            document_number = str((get_declaration(declaration_id) or {}).get("requestUID") or "").strip()
        except PortalBlocked:
            raise
        except Exception as e:
            log(f"Номер документа (requestUID) не получен: {type(e).__name__}: {e}", "WARNING")
        log(f"Номер документа: {document_number or '-'}")
        reg("READY_TO_SIGN", REG_STAGE_FIRST_SIGN)
    else:
        log(f"ПРОДОЛЖАЮ №{unique_no}: declarationId={declaration_id} уже на подписи, "
            f"остался только проход 2 («Мои дела»).")

    # 4. Проход 2: «Мои дела» → «Ожидается отправка» → тот же черновик → текст ещё раз.
    #    После этой вставки портал текст сохраняет сам; второй «Следующий шаг» не нужен.
    open_draft_from_my_cases(declaration_id)
    paste_both_fields(claim_text, "ПРОХОД 2")
    reg("READY_TO_SIGN", REG_STAGE_DONE)

    # 5. НЕ ПОДПИСЫВАЕМ И НЕ ОТПРАВЛЯЕМ.
    log("=" * 80)
    log(f"ГОТОВО ДО ПОДПИСАНИЯ. №{unique_no} | declarationId={declaration_id}")
    log("ЭЦП НЕ НАКЛАДЫВАЛАСЬ. Заявление оставлено в «Ожидается отправка» для ручной проверки и подписи.")
    return {"declaration_id": declaration_id, "document_number": document_number,
            "unique_no": unique_no, "skipped": False}


class DraftStuckError(RuntimeError):
    """Черновик уже создан, но продолжить его автоматически нельзя (не записан
    в реестр / неизвестно, какие файлы загружены). Повтор строки создал бы
    ДУБЛЬ или неполное заявление — поэтому пачку останавливаем."""


def _is_data_error(e):
    """Повтор строки не поможет (или навредит): ошибка исходных данных
    (папка/файлы/номер/суд) либо черновик, который нужно разобрать вручную."""
    return isinstance(e, (FileNotFoundError, ValueError, DraftStuckError))


def log_skipped(skipped):
    if not skipped:
        return
    log("-" * 80)
    log(f"НЕ ПОДАНЫ — НЕ ХВАТАЕТ ДОКУМЕНТОВ: {len(skipped)}", "WARNING")
    for excel_row, fio, iin, uid, missing in skipped:
        log(f"НЕТ ДОКОВ | Excel {excel_row} | №{uid} | {iin} | {fio} | нет: {', '.join(missing)}", "WARNING")


def _is_gbdfl_error(e):
    """ГБДФЛ не отдал данные по ИИН после собственных повторов (таймаут на стороне портала)."""
    return "GBDFL_RELOGIN_REQUIRED" in str(e)


def write_results(success, errors, rows_left, skipped=(), gbdfl_failed=()):
    wb = Workbook()
    ws = wb.active
    ws.title = "Результат"
    ws.append(["Excel строка", "ИИН", "ФИО", "Статус", "declarationId / ошибка"])
    for excel_row, fio, iin, uid, did, doc_no, was_ready in success:
        ws.append([excel_row, iin, fio,
                   "УЖЕ БЫЛА ПОДГОТОВЛЕНА (реестр)" if was_ready else "ГОТОВО ДО ПОДПИСИ",
                   f"№{uid}: declarationId={did}" + (f", номер документа {doc_no}" if doc_no else "")])
    for excel_row, fio, iin, err in errors:
        ws.append([excel_row, iin, fio, "ОШИБКА", err])
    for excel_row, fio, iin, uid, missing in skipped:
        ws.append([excel_row, iin, fio, "ПРОПУЩЕНА — НЕТ ДОКУМЕНТОВ", f"№{uid}: нет {', '.join(missing)}"])
    for excel_row, fio, iin in gbdfl_failed:
        ws.append([excel_row, iin, fio, "ПРОПУЩЕНА — ГБДФЛ НЕ ОТВЕТИЛ", "перезапустить позже"])
    for excel_row in rows_left:
        ws.append([excel_row, "", "", "НЕ ОБРАБОТАНА (пачка остановлена)", ""])
    for col, width in zip("ABCDE", (12, 14, 40, 34, 100)):
        ws.column_dimensions[col].width = width
    path = OUT_DIR / "portal_sot_results.xlsx"
    wb.save(path)
    log(f"Результат: {path}")


# ============================================================
# MAIN
# ============================================================

def main():
    global driver, portal_login, CASES_BATCH_DIR, LOCAL_COURTS

    log("=" * 90)
    log(f"СТАРТ portal-sot.kz — подготовка исков БЕЗ ПОДПИСИ | компания {COMPANY_ID}")
    log(f"Корневая папка: {WORK_ROOT}")
    log(f"Excel: {EXCEL_FILE}")
    log(f"Реестр подготовленных сделок: {REGISTRY_FILE}")
    log(f"Firefox-профиль: {FIREFOX_PROFILE_DIR}")
    log(f"Лог: {LOG_FILE}")
    log("=" * 90)

    if not EXCEL_FILE.exists():
        raise FileNotFoundError(f"Нет Excel-файла: {EXCEL_FILE}")
    CASES_BATCH_DIR = discover_batch_dir()
    log(f"Папка документов (партия): {CASES_BATCH_DIR}")
    try:
        registry_check_writable()
    except OSError as e:
        raise RuntimeError(f"Реестр недоступен на запись (открыт в Excel?): {REGISTRY_FILE} — {e}") from e

    courts_file = find_courts_file()
    LOCAL_COURTS = load_local_courts(courts_file)
    log(f"Локальный справочник судов: {courts_file} — судов: {len(LOCAL_COURTS)} "
        f"(уголовные и ID 18/23/99 исключены)")

    rows = rows_to_process()
    if not rows:
        raise RuntimeError(f"В Excel нет строк с ИИН начиная со строки {START_ROW}")

    # Предпроверка ДО входа: папки и файлы по всем строкам.
    # Нет иска/госпошлины — строку пропускаем и перечисляем в конце, остальное — стоп.
    problems = []
    skipped = []  # (excel_row, fio, iin, unique_no, [чего нет])
    for r in rows:
        try:
            c = load_case_row(r)
            f = find_case_folder(c["UNIQUE_NO"], c["FIO"], c["IIN"])
            try:
                ff = classify_case_files(f)
            except MissingDocsError as e:
                skipped.append((r, c["FIO"], c["IIN"], c["UNIQUE_NO"], e.missing))
                log(f"  {r} | №{c['UNIQUE_NO']} | {c['FIO']} | ПРОПУСК — нет: {', '.join(e.missing)}", "WARNING")
                continue
            if not c["REGION"] or not c["COURT"]:
                raise CourtDataError(
                    "пустые регион (P) / суд (Q) — сначала прогоните «Поиск адреса в СК» "
                    "(poiskvsk) по этому отчёту"
                )
            resolve_court(c["REGION"], c["COURT"])
            log(f"  {r} | №{c['UNIQUE_NO']} | {c['FIO']} | {c['IIN']} -> {f.name} | файлов: {2 + len(ff['extras'])}")
        except Exception as e:
            problems.append(f"строка {r}: {e}")
    if problems:
        for p in problems:
            log(f"ПРЕДПРОВЕРКА: {p}", "ERROR")
        raise RuntimeError(f"Предпроверка не пройдена ({len(problems)} строк) — исправьте данные и перезапустите")
    if skipped:
        skipped_rows = {s[0] for s in skipped}
        rows = [r for r in rows if r not in skipped_rows]
        log(f"Пропущено из-за отсутствующих документов: {len(skipped)} (список — в конце)", "WARNING")
    if not rows:
        log_skipped(skipped)
        write_results([], [], [], skipped)
        raise RuntimeError("Нет строк с полным комплектом документов — подавать нечего")

    global _lock
    _lock = portal_lock(log)
    _lock.acquire()  # другой скрипт уже на portal-sot.kz → ждём его

    portal_login = PortalSotLogin(COMPANY_ID, PORTAL_CFG, log=log, dump_dir=str(OUT_DIR))
    log("🔐 Вход на portal-sot.kz через ЭЦП...")
    driver = portal_login.init_driver()
    portal_sign_in()
    check_org()
    log(f"Авторизация portal-sot.kz: OK | Организация: {org.get('fullNameRu') or ORG_BIN}")
    pl = plaintiff_participant()  # адрес/наименование истца — проверка до первой сделки
    log(f"Истец: {pl['name']} | {pl['legalAddress']}")

    log("=" * 80)
    log(f"ПОДГОТОВКА БЕЗ ПОДПИСИ. Строк к обработке: {len(rows)}. Начинаем с Excel №{rows[0]}")
    log("Сделки из реестра со статусом READY_TO_SIGN повторно не создаются.")
    log("Пауза 30–60 секунд между сделками. При неудаче строки после 3 попыток — стоп пачки.")
    log("=" * 80)

    success, errors = [], []
    gbdfl_failed = []  # (excel_row, fio, iin) — ГБДФЛ не ответил, строка пропущена
    gbdfl_streak = 0
    MAX_GBDFL_STREAK = 3  # столько сбоев ГБДФЛ подряд — сервис лежит, стоп пачки
    stop_row = None
    MAX_ROW_ATTEMPTS = 3

    for n, excel_row in enumerate(rows, 1):
        # Следующая строка только после успеха текущей до step=sign
        # (исключение — сбой ГБДФЛ по ИИН: строку пропускаем и перечисляем в конце).
        row_ctx = load_case_row(excel_row)
        fio, iin = row_ctx["FIO"], row_ctx["IIN"]
        row_success = False
        row_gbdfl_skip = False
        was_ready = False

        for row_attempt in range(1, MAX_ROW_ATTEMPTS + 1):
            try:
                if row_attempt > 1:
                    log(f"ПОВТОР ТЕКУЩЕЙ СДЕЛКИ: Excel строка {excel_row}, попытка {row_attempt}/{MAX_ROW_ATTEMPTS}")
                result = run_one(excel_row)
                declaration_id = result["declaration_id"]
                was_ready = result["skipped"]
                success.append((excel_row, fio, iin, result["unique_no"], declaration_id,
                                result["document_number"], was_ready))
                log(f"СДЕЛКА {n}/{len(rows)} "
                    f"{'УЖЕ БЫЛА ПОДГОТОВЛЕНА — НОВЫЙ ИСК НЕ СОЗДАВАЛСЯ' if was_ready else 'ПОДГОТОВЛЕНА БЕЗ ПОДПИСИ'}"
                    f" | №{result['unique_no']} | declarationId={declaration_id}"
                    f" | номер документа={result['document_number'] or '-'}")
                row_success = True
                gbdfl_streak = 0
                break

            except Exception as e:
                log(f"СДЕЛКА {n}/{len(rows)} | Excel строка {excel_row} | "
                    f"попытка {row_attempt}/{MAX_ROW_ATTEMPTS} ОШИБКА: {type(e).__name__}: {e}", "ERROR")
                log_exception("TRACEBACK")

                if isinstance(e, PortalBlocked):
                    errors.append((excel_row, fio, iin, f"{type(e).__name__}: {e}"))
                    stop_row = excel_row
                    log(f"СТОП ВСЕЙ ПАЧКИ на Excel строке {excel_row}: {e}", "ERROR")
                    break

                if _is_gbdfl_error(e):
                    # person_from_gbdfl уже сделал 3 попытки с паузами — повтор строки
                    # и перелогин не помогают, черновик ещё не создан.
                    gbdfl_streak += 1
                    if gbdfl_streak >= MAX_GBDFL_STREAK:
                        errors.append((excel_row, fio, iin, f"{type(e).__name__}: {e}"))
                        stop_row = excel_row
                        log(f"СТОП ВСЕЙ ПАЧКИ на Excel строке {excel_row}: ГБДФЛ не отвечает "
                            f"{gbdfl_streak} займа подряд — похоже, сервис недоступен.", "ERROR")
                    else:
                        gbdfl_failed.append((excel_row, fio, iin))
                        row_gbdfl_skip = True
                        log(f"ПРОПУСК Excel строки {excel_row}: ГБДФЛ не ответил по {iin} "
                            f"(сбоев подряд: {gbdfl_streak}/{MAX_GBDFL_STREAK}). Идём дальше.", "WARNING")
                    break

                if _is_data_error(e):
                    errors.append((excel_row, fio, iin, f"{type(e).__name__}: {e}"))
                    stop_row = excel_row
                    log(f"СТОП ВСЕЙ ПАЧКИ на Excel строке {excel_row}: повтор не поможет "
                        f"({'черновик нужно разобрать вручную' if isinstance(e, DraftStuckError) else 'ошибка исходных данных'}).",
                        "ERROR")
                    break

                if row_attempt >= MAX_ROW_ATTEMPTS:
                    errors.append((excel_row, fio, iin, f"{type(e).__name__}: {e}"))
                    stop_row = excel_row
                    log(f"СТОП ВСЕЙ ПАЧКИ на Excel строке {excel_row}: не удалось довести сделку до "
                        f"step=sign после {MAX_ROW_ATTEMPTS} попыток.", "ERROR")
                    break

                if row_attempt > 1:
                    log(f"Excel строка {excel_row} всё ещё не завершена. Повторный вход и снова ЭТА ЖЕ сделка.")
                    try:
                        force_relogin_portal()
                    except Exception as relog_e:
                        log(f"Повторный вход завершился ошибкой: {type(relog_e).__name__}: {relog_e}", "ERROR")
                pause_sec = random.randint(45, 75)
                log(f"Excel строка {excel_row} НЕ ПРОПУЩЕНА. Пауза {pause_sec} сек. и повтор...")
                time.sleep(pause_sec)

        if not row_success and not row_gbdfl_skip:
            break

        if n < len(rows) and not was_ready:
            pause_sec = random.randint(30, 60)
            log(f"Пауза {pause_sec} секунд перед следующей сделкой...")
            time.sleep(pause_sec)

    done_rows = {r for r, *_ in success} | {r for r, *_ in errors} | {r for r, *_ in gbdfl_failed}
    rows_left = [r for r in rows if r not in done_rows]

    log("=" * 80)
    log("ПОДГОТОВКА БЕЗ ПОДПИСИ ЗАВЕРШЕНА")
    log(f"ПОДГОТОВЛЕНО БЕЗ ПОДПИСИ: {len(success)}")
    log(f"ОШИБОК: {len(errors)}")
    if stop_row is not None:
        log(f"STOP: для продолжения запустите с «Начальная строка» = {stop_row}", "ERROR")
    for excel_row, fio, iin, uid, did, doc_no, was_ready in success:
        log(f"{'УЖЕ ГОТОВА' if was_ready else 'OK'} | Excel {excel_row} | №{uid} | {iin} | {fio} | "
            f"declarationId={did} | номер документа={doc_no or '-'}")
    for excel_row, fio, iin, err in errors:
        log(f"ERROR | Excel {excel_row} | {iin} | {fio} | {err}", "ERROR")
    log_skipped(skipped)
    if gbdfl_failed:
        log("-" * 80)
        log(f"НЕ ПОДАНЫ — ГБДФЛ НЕ ОТВЕТИЛ: {len(gbdfl_failed)} (перезапустите позже по одной строке)", "WARNING")
        for excel_row, fio, iin in gbdfl_failed:
            log(f"ГБДФЛ | Excel {excel_row} | {iin} | {fio}", "WARNING")
    log("=" * 80)
    write_results(success, errors, rows_left, skipped, gbdfl_failed)
    return 1 if errors or gbdfl_failed else 0


_lock = None

if __name__ == "__main__":
    try:
        code = main()
    except Exception as e:
        log(f"ФАТАЛЬНАЯ ОШИБКА: {type(e).__name__}: {e}", "ERROR")
        log_exception("TRACEBACK")
        code = 1
    finally:
        if _lock:
            _lock.release()
    sys.exit(code)
