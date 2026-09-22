# blocks/postobotmenee.py
# -*- coding: utf-8 -*-
"""
Блок 13: Постановление об отмене исполнительной надписи

Схема — как в AISOIP restrictions checker (API version): браузер логинится
один раз через Selenium и ОСТАЁТСЯ ОТКРЫТЫМ на всё время работы блока.
Все запросы к API (поиск, список документов, скачивание) идут через
fetch() внутри самой страницы (driver.execute_async_script), с
credentials: 'include' — браузер сам подставляет все свои cookies,
включая httpOnly. Никакого ручного переноса cookies в requests.Session:
именно это было причиной постоянных 401/403 в предыдущей версии.
"""

import io
import os
import re
import json
import time
import base64
from pathlib import Path
from datetime import datetime
from urllib.parse import urlencode

import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException

from pypdf import PdfReader

from config import CREDENTIALS, MAIN_EXCEL, TARGET_BASE
from utils import safe_log, safe_update_summary

# ═══════════════════════════════════════════════════════════════
# НАСТРОЙКИ
# ═══════════════════════════════════════════════════════════════

AISOIP_URL = "https://aisoip.adilet.gov.kz/cabinet/exec-productions"
API_BASE   = "https://aisoip.adilet.gov.kz/extperson/api/rest/execproc"
LOGIN    = CREDENTIALS['aisoip_login']
PASSWORD = CREDENTIALS['aisoip_password']

DEBUG_DIR = str(Path(os.getenv("LOCALAPPDATA", r"C:\Users\User\AppData\Local")) / "Temp" / "AISOIP_debug")
os.makedirs(DEBUG_DIR, exist_ok=True)

EXCLUDE_TITLES = [
    "Постановление о прекращении исполнительного производства",
    "Отчет о доставке", "Отчёт о доставке",
    "Извещение",
    "Инкассовое распоряжение",
    "Постановление о запрете должнику совершать определенные действия",
    "Постановление об истребовании информации",
    "Извещение должника о временном ограничении на выезд из РК",
    "Постановление о возбуждении исполнительного производства",
    "Постановление о наложении ареста на транспортные средства",
]

RE_CANCEL_LIST = [
    re.compile(r"\bоб\s+отмен[еёы]\b.{0,60}\bисполн\w*\s+надпис\w*", re.IGNORECASE),
    re.compile(r"\bотменить\b.{0,60}\bисполн\w*\s+надпис\w*", re.IGNORECASE),
    re.compile(r"\bпризнать\b.{0,60}\bисполн\w*\s+надпис\w*\s+недействител\w*", re.IGNORECASE),
    re.compile(r"\bоб\s+отмен[еёы]\s+ин\b", re.IGNORECASE),
]

RE_NEGATION = re.compile(
    r"(отказать(?:\s+в\s+удовлетворении)?|в\s+удовлетворении\s+.*?отказать|оставить\s+без\s+изменения)",
    re.IGNORECASE
)

ALLOWED_EXTS = {".pdf", ".jpg", ".jpeg", ".png"}

# ═══════════════════════════════════════════════════════════════
# ЛОГИРОВАНИЕ
# ═══════════════════════════════════════════════════════════════

_t0 = time.time()

def log(msg):
    dt = time.time() - _t0
    print(f"[{dt:7.2f}s] {msg}")

def dbg_snapshot(driver, tag):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    png = os.path.join(DEBUG_DIR, f"{ts}_{tag}.png")
    html = os.path.join(DEBUG_DIR, f"{ts}_{tag}.html")
    try:
        driver.save_screenshot(png)
        with open(html, "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        log(f"[DBG] Снимки сохранены: {png}")
    except Exception as e:
        log(f"[DBG] Не удалось сохранить снимки: {e}")

# ═══════════════════════════════════════════════════════════════
# SELENIUM — ЛОГИН (браузер остаётся открытым после этого)
# ═══════════════════════════════════════════════════════════════

def make_driver():
    opts = webdriver.ChromeOptions()
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--start-maximized")
    service = Service(ChromeDriverManager().install())
    drv = webdriver.Chrome(service=service, options=opts)
    drv.set_script_timeout(90)
    log("[INIT] Запущен Chrome WebDriver")
    return drv

def fill_login_form(driver):
    wait = WebDriverWait(driver, 20)
    user = wait.until(EC.presence_of_element_located((
        By.CSS_SELECTOR,
        "input[name='usernameUserInput'], input#usernameUserInput, input[name='username'], input#username, form#loginForm input[type='text']"
    )))
    pwd = wait.until(EC.presence_of_element_located((
        By.CSS_SELECTOR,
        "input#password, input[name='password'], form#loginForm input[type='password']"
    )))
    driver.execute_script("""
        const u = arguments[0], p = arguments[1], lu = arguments[2], lp = arguments[3];
        u.value = lu; u.dispatchEvent(new Event('input', {bubbles:true})); u.dispatchEvent(new Event('change', {bubbles:true}));
        p.value = lp; p.dispatchEvent(new Event('input', {bubbles:true})); p.dispatchEvent(new Event('change', {bubbles:true}));
    """, user, pwd, LOGIN, PASSWORD)
    try:
        user.clear(); user.send_keys(LOGIN)
    except Exception:
        pass
    try:
        pwd.clear(); pwd.send_keys(PASSWORD)
    except Exception:
        pass

def click_sign_in(driver):
    wait = WebDriverWait(driver, 10)
    btn = wait.until(EC.element_to_be_clickable((
        By.CSS_SELECTOR, "#signInButton, form#loginForm button.ui.primary.button.fluid, button[type='submit']"
    )))
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    except Exception:
        pass
    for _ in range(3):
        try:
            btn.click()
        except Exception:
            pass
        time.sleep(0.15)
        try:
            driver.execute_script("arguments[0].click();", btn)
        except Exception:
            pass
        time.sleep(0.2)

def is_logged_in(driver):
    try:
        if "/cabinet/" in (driver.current_url or "") and \
           driver.find_elements(By.XPATH, "//div[contains(@class,'navigation') or contains(@class,'v-tabs') or contains(@class,'v-tab')]"):
            return True
    except Exception:
        pass
    return False

def is_login_page(driver) -> bool:
    try:
        url = (driver.current_url or "").lower()
        if "login" in url:
            return True
        return len(driver.find_elements(By.CSS_SELECTOR, "form#loginForm, #signInButton")) > 0
    except Exception:
        return False

def login_with_retries(driver, max_rounds=6):
    log("[LOGIN] Начинаю авторизацию…")
    driver.get(AISOIP_URL)
    for i in range(1, max_rounds + 1):
        log(f"[LOGIN] попытка {i}/{max_rounds}")
        try:
            WebDriverWait(driver, 25).until(EC.presence_of_element_located((
                By.CSS_SELECTOR,
                "input[name='usernameUserInput'], input#usernameUserInput, input[name='username'], input#username, form#loginForm input[type='text']"
            )))
        except TimeoutException:
            dbg_snapshot(driver, f"login_form_not_found_{i}")
            continue

        fill_login_form(driver)
        time.sleep(0.2)
        click_sign_in(driver)

        try:
            WebDriverWait(driver, 12).until(lambda d: is_logged_in(d))
            if is_logged_in(driver):
                log("[OK] Авторизация выполнена.")
                return True
        except TimeoutException:
            pass

        dbg_snapshot(driver, f"login_retry_{i}")
        try:
            driver.refresh()
        except Exception:
            pass
        time.sleep(0.8)

    dbg_snapshot(driver, "login_failed_after_retries")
    raise TimeoutException("Не удалось войти после серии повторов.")

def start_driver_and_login(max_browser_restarts: int = 3):
    last_err = None
    for attempt in range(1, max_browser_restarts + 1):
        drv = None
        try:
            log(f"[SESSION] Старт браузера, попытка {attempt}/{max_browser_restarts}")
            drv = make_driver()
            login_with_retries(drv)
            log("[SESSION] Браузер и авторизация успешны")
            return drv
        except Exception as e:
            last_err = e
            log(f"[SESSION] Ошибка при запуске/авторизации: {e}")
            try:
                if drv is not None:
                    drv.quit()
            except Exception:
                pass
            time.sleep(2)
    raise TimeoutException(f"Не удалось инициализировать браузер и войти в систему: {last_err}")

def ensure_driver_and_auth(driver, max_browser_restarts: int = 3):
    """Проверяет, что браузер жив и сессия активна. Если разлогинило —
    логинится заново В ТОМ ЖЕ браузере (без пересоздания драйвера).
    Только если сам браузер умер — пересоздаёт его."""
    try:
        _ = driver.current_url
    except WebDriverException as e:
        log(f"[SESSION] Браузер закрыт, перезапуск: {e}")
        try:
            driver.quit()
        except Exception:
            pass
        return start_driver_and_login(max_browser_restarts=max_browser_restarts)

    try:
        if is_login_page(driver) or not is_logged_in(driver):
            log("[SESSION] Разлогинило. Повторная авторизация в этом же браузере...")
            login_with_retries(driver)
    except Exception as e:
        log(f"[SESSION] Не удалось восстановить сессию: {e}. Перезапуск браузера...")
        try:
            driver.quit()
        except Exception:
            pass
        return start_driver_and_login(max_browser_restarts=max_browser_restarts)

    return driver

# ═══════════════════════════════════════════════════════════════
# API ЧЕРЕЗ fetch() ВНУТРИ СТРАНИЦЫ (как во втором скрипте)
# ═══════════════════════════════════════════════════════════════

class ExecProcAuthError(Exception):
    """401/403 от API — сессия в браузере протухла, нужен повторный логин."""

_FETCH_JSON_SCRIPT = """
const callback = arguments[arguments.length - 1];
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
    body: body ? JSON.stringify(body) : undefined
})
.then(async r => {
    const text = await r.text();
    callback({status: r.status, text: text});
})
.catch(e => callback({status: 0, text: String(e)}));
"""

_FETCH_BINARY_SCRIPT = """
const callback = arguments[arguments.length - 1];
const url = arguments[0];

fetch(url, {
    method: 'GET',
    credentials: 'include',
    headers: { 'Accept': '*/*' }
})
.then(async r => {
    const blob = await r.blob();
    const reader = new FileReader();
    reader.onloadend = () => callback({
        status: r.status,
        contentType: r.headers.get('content-type') || '',
        contentDisposition: r.headers.get('content-disposition') || '',
        dataUrl: reader.result
    });
    reader.readAsDataURL(blob);
})
.catch(e => callback({status: 0, contentType: '', contentDisposition: '', dataUrl: String(e)}));
"""

def browser_fetch_json(driver, url, method="GET", body=None):
    res = driver.execute_async_script(_FETCH_JSON_SCRIPT, url, method, body)
    if res["status"] in (401, 403):
        raise ExecProcAuthError(f"HTTP {res['status']} на {url}")
    if res["status"] != 200:
        raise RuntimeError(f"fetch failed {res['status']}: {res['text'][:500]}")
    return json.loads(res["text"])

def browser_fetch_binary(driver, url):
    """Возвращает (content_bytes, content_type)."""
    res = driver.execute_async_script(_FETCH_BINARY_SCRIPT, url)
    if res["status"] in (401, 403):
        raise ExecProcAuthError(f"HTTP {res['status']} на {url}")
    if res["status"] != 200:
        raise RuntimeError(f"fetch failed {res['status']}: {str(res)[:500]}")

    data_url = res["dataUrl"]
    if "," not in data_url:
        raise RuntimeError(f"Некорректный dataUrl: {data_url[:200]}")

    b64 = data_url.split(",", 1)[1]
    content = base64.b64decode(b64)
    return content, res.get("contentType", "")

def api_search_page(driver, iin: str, page: int, size: int) -> dict:
    url = f"{API_BASE}/search?{urlencode({'page': page, 'size': size})}"
    return browser_fetch_json(driver, url, method="POST", body={"iin": iin, "searchType": False})

def api_search_all(driver, iin: str, size=50) -> list:
    results, page = [], 0
    while True:
        data = api_search_page(driver, iin, page, size)
        content = data.get("content", []) if isinstance(data, dict) else (data or [])
        results.extend(content)
        pagination = data.get("pagination") or {} if isinstance(data, dict) else {}
        if pagination.get("last", True) or not content:
            break
        page += 1
    return results

def api_get_docs(driver, exec_proc_id) -> list:
    url = f"{API_BASE}/doc/{exec_proc_id}?{urlencode({'searchType': 'false'})}"
    data = browser_fetch_json(driver, url, method="GET")
    return data if isinstance(data, list) else []

def api_download_doc(driver, did: str):
    """Возвращает (content_bytes, ext)."""
    url = f"{API_BASE}/doc/{did}/{did}?{urlencode({'lang': 'ru'})}"
    content, ctype = browser_fetch_binary(driver, url)
    ext = ".pdf" if "pdf" in ctype else (".jpg" if "jpeg" in ctype or "jpg" in ctype else ".png" if "png" in ctype else ".pdf")
    return content, ext

# ═══════════════════════════════════════════════════════════════
# КЛАССИФИКАЦИЯ ДОКУМЕНТА
# ═══════════════════════════════════════════════════════════════

def is_cancel_doc_bytes(content: bytes) -> bool:
    try:
        reader = PdfReader(io.BytesIO(content))
        text_chunks = []
        for page in reader.pages[:6]:
            try:
                text_chunks.append(page.extract_text() or "")
            except Exception:
                pass

        raw = "\n".join(text_chunks)
        if not raw.strip():
            return False

        txt = raw.replace("Ё", "Е").replace("ё", "е")
        txt = re.sub(r"\s+", " ", txt).strip().lower()

        for ex in (e.lower() for e in EXCLUDE_TITLES):
            if ex in txt:
                return False

        for pat in RE_CANCEL_LIST:
            for m in pat.finditer(txt):
                start = max(0, m.start() - 120)
                end = min(len(txt), m.end() + 120)
                window = txt[start:end]
                if RE_NEGATION.search(window):
                    continue
                return True

        return False
    except Exception as e:
        log(f"[PDF] Не удалось распарсить содержимое: {e}")
        return False

def pick_and_download_cancel_doc_api(driver, exec_proc_id, max_checks=4):
    docs = api_get_docs(driver, exec_proc_id)
    checks = 0

    for doc in docs:
        if checks >= max_checks:
            break
        title = (doc.get("ddocTitle") or "").strip()
        did = doc.get("did")
        if not did:
            continue
        if any(ex.lower() in title.lower() for ex in EXCLUDE_TITLES):
            log(f"[DL] Пропуск по исключению: {title}")
            continue

        try:
            content, ext = api_download_doc(driver, did)
        except ExecProcAuthError:
            raise
        except Exception as e:
            log(f"[DL] Ошибка скачивания '{title}': {e}")
            continue

        checks += 1

        if ext == ".pdf":
            if is_cancel_doc_bytes(content):
                log(f"[DL] Найден документ об отмене: {title}")
                return content, ext
            log(f"[DL] Не подходит: {title}")
        else:
            log(f"[DL] Получено изображение, считаю документом об отмене: {title}")
            return content, ext

    return None

def sort_exec_procs_by_date_desc(proceedings: list) -> list:
    """Сортирует производства должника от самого свежего к самому старому."""
    if not proceedings:
        return []

    def _key(p):
        try:
            return datetime.fromisoformat(p.get("startDate") or "1900-01-01")
        except Exception:
            return datetime(1900, 1, 1)

    return sorted(proceedings, key=_key, reverse=True)

# ═══════════════════════════════════════════════════════════════
# ФАЙЛЫ / ПАПКИ КЛИЕНТОВ
# ═══════════════════════════════════════════════════════════════

def find_client_folder_by_iin(iin: str):
    for name in os.listdir(TARGET_BASE):
        full = os.path.join(TARGET_BASE, name)
        if os.path.isdir(full) and iin in name:
            return full
    return None

def get_fio_from_folder(iin: str) -> str:
    folder = find_client_folder_by_iin(iin)
    if not folder:
        return "Неизвестный"
    base = os.path.basename(folder)
    fio = base.split(",")[0].strip()
    return fio or "Неизвестный"

def sanitize_filename_strict(s: str) -> str:
    s = re.sub(r'[<>:"/\\|?*\x00-\x1F]', " ", s)
    s = re.sub(r"\s+", " ", s, flags=re.UNICODE).strip(" \t\u00A0")
    s = s.rstrip(".").rstrip(" \t\u00A0")
    return s[:180] if len(s) > 180 else s

def save_doc_bytes(content: bytes, ext: str, fio: str, iin: str) -> bool:
    folder = find_client_folder_by_iin(iin)
    if not folder:
        msg = f"[FATAL] Не найдена папка клиента по ИИН {iin}"
        log(msg)
        safe_log(f"[ОТМЕНА ИН] {msg}")
        return False

    fio_safe = sanitize_filename_strict(fio or "Неизвестный")
    if ext not in ALLOWED_EXTS:
        ext = ".pdf"

    base_name = sanitize_filename_strict(f"Постановление об отмене ИН, {fio_safe}, {iin}{ext}")
    dst = os.path.join(folder, base_name)

    try:
        if os.path.exists(dst):
            try:
                os.remove(dst)
                log(f"[OK] Перезаписываю существующий файл: {dst}")
            except Exception as e:
                log(f"[WARN] Не удалось удалить старый файл: {e}")

        with open(dst, "wb") as f:
            f.write(content)
        log(f"[OK] Сохранён: {dst}")
        return True
    except Exception as e:
        msg = f"[FATAL] Не удалось сохранить '{dst}': {e}"
        log(msg)
        safe_log(f"[ОТМЕНА ИН] {msg}")
        return False

# ═══════════════════════════════════════════════════════════════
# ГЛАВНАЯ ФУНКЦИЯ БЛОКА
# ═══════════════════════════════════════════════════════════════

def run(df_main=None):
    if df_main is None:
        df_main = pd.read_excel(MAIN_EXCEL, usecols=[1, 2, 3], header=0)
        df_main.columns = ['Product', 'FIO', 'IIN']
        df_main['IIN'] = df_main['IIN'].astype(str).str.zfill(12)

    print("\n=== БЛОК 13: Постановление об отмене ИН ===\n")

    count_success = 0
    count_failed = 0
    not_found_list = []

    driver = start_driver_and_login()

    try:
        for idx, row in df_main.iterrows():
            iin = str(row['IIN']).strip().zfill(12)

            if len(iin) != 12:
                log(f"[{idx+1}] Пропуск — некорректный ИИН: {iin}")
                safe_log(f"[ОТМЕНА ИН] Пропуск некорректного ИИН: {iin}")
                count_failed += 1
                not_found_list.append(f"Неизвестный, {iin}")
                continue

            log(f"\n[{idx+1}] === ИИН: {iin} ===")
            fio_from_folder = get_fio_from_folder(iin)

            for retry in range(2):
                try:
                    proceedings = api_search_all(driver, iin)
                    ordered = sort_exec_procs_by_date_desc(proceedings)

                    if not ordered:
                        msg = f"[WARN] Для {iin} не найдено исполнительных производств"
                        log(msg)
                        count_failed += 1
                        not_found_list.append(f"{fio_from_folder}, {iin}")
                        safe_log(f"[ОТМЕНА ИН] {msg}")
                        break

                    # У должника может быть несколько производств — документ об
                    # отмене ИН может лежать не в самом свежем. Перебираем все,
                    # от свежего к старому, пока не найдём.
                    result = None
                    for i, proc in enumerate(ordered, start=1):
                        exec_proc_id = proc.get("execProcId")
                        if not exec_proc_id:
                            continue
                        log(f"[{idx+1}] Проверяю производство {i}/{len(ordered)} (execProcId={exec_proc_id})")
                        result = pick_and_download_cancel_doc_api(driver, exec_proc_id, max_checks=4)
                        if result:
                            break

                    if not result:
                        msg = f"[WARN] Для {iin} не найден документ об отмене ИН"
                        log(msg)
                        count_failed += 1
                        not_found_list.append(f"{fio_from_folder}, {iin}")
                        safe_log(f"[ОТМЕНА ИН] Не найден документ об отмене для ФИО: {fio_from_folder}, ИИН: {iin}")
                    else:
                        content, ext = result
                        ok = save_doc_bytes(content, ext, fio_from_folder, iin)
                        if ok:
                            count_success += 1
                            log(f"[OK] Документ сохранён для {fio_from_folder}, {iin}")
                        else:
                            count_failed += 1
                            not_found_list.append(f"{fio_from_folder}, {iin}")
                    break

                except ExecProcAuthError:
                    if retry == 0:
                        log("[SESSION] Сессия протухла — повторный логин в этом же браузере...")
                        driver = ensure_driver_and_auth(driver)
                        continue
                    msg = f"[WARN] ИИН {iin}: не удалось восстановить сессию"
                    log(msg)
                    count_failed += 1
                    not_found_list.append(f"{fio_from_folder}, {iin}")
                    safe_log(f"[ОТМЕНА ИН] {msg}")

                except Exception as e:
                    msg = f"[WARN] ИИН {iin}: ошибка обработки: {e}"
                    log(msg)
                    count_failed += 1
                    not_found_list.append(f"{fio_from_folder}, {iin}")
                    safe_log(f"[ОТМЕНА ИН] Ошибка обработки для ФИО: {fio_from_folder}, ИИН: {iin}, ошибка: {e}")
                    dbg_snapshot(driver, f"error_{iin}")
                    break

        log("\n=== ГОТОВО ===")
        log(f"[ИТОГ ОТМЕН ИН] Найдено {count_success} из {len(df_main)}")

    finally:
        try:
            driver.quit()
            log("[CLEANUP] Браузер закрыт")
        except Exception:
            pass

    print(f"\n--- ИТОГ ПОСТАНОВЛЕНИЙ ОБ ОТМЕНЕ ИН ---")
    print(f"✔️ Скопировано: {count_success}")
    print(f"❌ Не найдено: {count_failed}")

    safe_update_summary("ПОСТАНОВЛЕНИЯ ОБ ОТМЕНЕ ИН", {
        "found": count_success,
        "total": len(df_main),
        "not_found": not_found_list,
    })

    return count_success, count_failed