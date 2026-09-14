# -*- coding: utf-8 -*-
"""
СК — подготовка заявлений на выдачу исполнительного листа (ИЛ).

Портирован из ноутбука SK_IL_Omega_DECISION_v27_NO_POA_NS_FIX.ipynb без
изменения логики подачи. Отличия от ноутбука:
  * параметры --workdir / --company_id / --max_rows (запуск из веб-интерфейса);
  * реквизиты (project_name, director) и шаблон .docx берутся из
    config.IL_DECISION_CONFIG по company_id;
  * SK-логин/пароль — из config.CREDENTIALS (sk_login/sk_password);
  * фильтр компании в SQL — config.DB_COMPANY_FILTER (l.F209), а не '%Омега%';
  * справочник судов СК — scripts/data/sk_courts_directory.json;
  * результаты (DOCX/PDF/xlsx-лог) пишутся в <workdir>/out/, DEBUG_* — в <workdir>.

Поведение на финале не изменено: скрипт доводит каждое заявление до
sign.xhtml (номер присвоен) и ОСТАНАВЛИВАЕТСЯ до подписания ЭЦП.
"""
import argparse
import os
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass


def _parse_args():
    p = argparse.ArgumentParser(description="СК — заявления на выдачу ИЛ")
    p.add_argument("--workdir", default=None, help="рабочая папка задачи (из runner)")
    p.add_argument("--company_id", default=None, help="ID компании (config)")
    p.add_argument("--max_rows", default=None, help="ограничить число сделок (тест)")
    return p.parse_args()


_args = _parse_args()

if not _args.company_id or not _args.company_id.strip():
    print("❌ ОШИБКА: не передан --company_id — компания не определена.")
    sys.exit(1)

# COMPANY_ID нужен config.py ещё на этапе импорта
os.environ["COMPANY_ID"] = _args.company_id.strip()

_WORKDIR = Path(_args.workdir).resolve() if (_args.workdir and _args.workdir.strip()) else Path.cwd()
_WORKDIR.mkdir(parents=True, exist_ok=True)

from config import CREDENTIALS, DB_COMPANY_FILTER  # noqa: E402

try:
    from config import IL_DECISION_CONFIG  # noqa: E402
except ImportError:
    print("❌ В scripts/config.py нет IL_DECISION_CONFIG — добавьте по образцу config.example.py")
    sys.exit(1)

# ============================================================
# НАСТРОЙКИ
# ============================================================

from config import CRM_DB
DB_SERVER   = CRM_DB["server"]
DB_DATABASE = CRM_DB["database"]
DB_USERNAME = CRM_DB["username"]
DB_PASSWORD = CRM_DB["password"]

AUTO_LOGIN = True
USER_AUTH     = CREDENTIALS["sk_login"]
USER_PASSWORD = CREDENTIALS["sk_password"]

WORK_DIR = _WORKDIR
OUT_DIR  = WORK_DIR / "out"
DOCX_DIR = OUT_DIR / "DOCX"
PDF_DIR  = OUT_DIR / "PDF"
LOG_FILE = OUT_DIR / "SK_IL_results.xlsx"

_SCRIPT_DIR = Path(__file__).resolve().parent
COURTS_JSON = _SCRIPT_DIR / "data" / "sk_courts_directory.json"

_ILCFG = IL_DECISION_CONFIG.get(os.environ["COMPANY_ID"])
if not _ILCFG:
    print(f"❌ Нет записи для компании {os.environ['COMPANY_ID']} в IL_DECISION_CONFIG "
          f"(config.py): нужны project_name, director, template.")
    sys.exit(1)
_missing = [k for k in ("project_name", "director", "template") if not _ILCFG.get(k)]
if _missing:
    print(f"❌ В IL_DECISION_CONFIG['{os.environ['COMPANY_ID']}'] не заполнено: {', '.join(_missing)}")
    sys.exit(1)

PROJECT_CFG = {
    "project_name": _ILCFG["project_name"],
    "director": _ILCFG["director"],
}
# Совместимость с кодом ноутбука (обращался к PROJECT_CONFIG["Омега"])
PROJECT_CONFIG = {"Омега": PROJECT_CFG}

TEMPLATE_PATH = _SCRIPT_DIR / "templates" / _ILCFG["template"]
if not TEMPLATE_PATH.exists():
    print(f"❌ Шаблон не найден: {TEMPLATE_PATH}")
    sys.exit(1)

if _args.max_rows and str(_args.max_rows).strip():
    try:
        MAX_ROWS = int(str(_args.max_rows).strip())
    except ValueError:
        MAX_ROWS = None
else:
    MAX_ROWS = None

SKIP_ALREADY_CREATED = True

HTTP_TIMEOUT = 180
HTTP_RETRIES = 3
PAUSE_BETWEEN_ROWS = 0.5

LOADING_TIME = 60
PAGELOAD_TIMEOUT = 120

print(f"Компания: {os.environ['COMPANY_ID']} | {PROJECT_CFG['project_name']}")
print(f"Фильтр F209: {DB_COMPANY_FILTER}")
print(f"Шаблон: {TEMPLATE_PATH.name}")
print(f"Рабочая папка: {WORK_DIR}")

# ============================================================
# SQL — ОТБОР СДЕЛОК ДЛЯ ЗАЯВЛЕНИЯ НА ВЫДАЧУ ИЛ
# (условия отбора не изменены; фильтр компании — DB_COMPANY_FILTER)
# ============================================================

_SQL_TEMPLATE = r"""
WITH LoanCounts AS (
    SELECT 
        c.F293 AS ИИН,
        COUNT(*) AS LoanCount
    FROM loans l WITH (NOLOCK)
    JOIN clients c WITH (NOLOCK) ON l.CID = c.ID
    WHERE l.F209 = N'{F209}'
    GROUP BY c.F293
)
SELECT
    l.ID                                  AS LoanID,
    l.F209                                AS Project,
    l.EID                                 AS EID,
    dF246.F249                            AS Product,
    c.FIO                                 AS FIO,
    c.F293                                AS IIN,
    s.Caption                             AS CreditStatus,
    NULLIF(constF236.Caption, N'')        AS F236,
    psel.F186                             AS DecisionDate,
    constF238.Caption                     AS Region,
    LTRIM(RTRIM(l.F157))                  AS Sud,
    l.F65                                 AS CaseNumber,
    l.F49                                 AS Judge,
    ISNULL(lc.LoanCount, 0)               AS LoanCount
FROM loans l WITH (NOLOCK)

OUTER APPLY (
    SELECT TOP (1)
        p.F186,
        p.F318
    FROM ProcessCreatedInLoans pcl WITH (NOLOCK)
    JOIN dbo.Process p WITH (NOLOCK) ON p.Id = pcl.ProcessId
    WHERE pcl.LoanId = l.Id
    ORDER BY p.Id ASC
) psel

JOIN clients c WITH (NOLOCK) ON l.CID = c.ID
JOIN Constants constF238 WITH (NOLOCK) ON constF238.ID = c.F238
JOIN states s WITH (NOLOCK) ON s.ID = l.State
JOIN Dictionary dF246 WITH (NOLOCK) ON dF246.ID = l.F246
LEFT JOIN Constants constF236 WITH (NOLOCK) ON constF236.ID = l.F236
LEFT JOIN LoanCounts lc ON c.F293 = lc.ИИН

WHERE 
    s.Caption NOT IN (N'Погашен', N'Обратный выкуп', N'В графике', N'Умерший')
    AND l.F209 = N'{F209}'
    AND (constF236.ID IS NULL OR constF236.Caption NOT IN (N'На исполнении'))
    AND psel.F186 IS NOT NULL
    AND (l.F220 = 0 OR l.F220 IS NULL)
    AND psel.F186 < DATEADD(DAY, -33, GETDATE())
    AND psel.F318 IS NULL
    AND NULLIF(LTRIM(RTRIM(l.F157)), N'') IS NOT NULL
    AND c.F293 IS NOT NULL

ORDER BY psel.F186 ASC;
"""
SQL_QUERY = _SQL_TEMPLATE.replace("{F209}", DB_COMPANY_FILTER)


# ============================================================
# ОСНОВНОЙ КОД
# ============================================================

import io
import json
import re
import html
import time
import uuid
import mimetypes
import shutil
import subprocess
import zipfile
import base64
from pathlib import Path
from urllib.parse import urljoin

import pandas as pd
import pyodbc
import requests
from bs4 import BeautifulSoup
from lxml import etree

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


BASE_URL = "https://office.sud.kz"
LOGIN_URL = BASE_URL + "/index.xhtml"
LETTER_INFO_URL = BASE_URL + "/form/letter/info.xhtml"
LETTER_WAIT_URL = BASE_URL + "/form/letter/wait.xhtml"


# ============================================================
# ВСТРОЕННЫЙ СПРАВОЧНИК СУДОВ СУДЕБНОГО КАБИНЕТА
# ============================================================
# Справочник выгружен непосредственно из office.sud.kz.
# Внешний JSON-файл для работы основного скрипта НЕ нужен.


def load_embedded_court_directory():
    return json.loads(Path(COURTS_JSON).read_text(encoding="utf-8"))


SK_COURTS_DIRECTORY = load_embedded_court_directory()

SK_COURTS_FLAT = []
for _region in SK_COURTS_DIRECTORY:
    for _court in _region.get("courts", []):
        SK_COURTS_FLAT.append({
            "region_id": str(_region["region_id"]),
            "region_name": _region["region_name"],
            "court_id": str(_court["court_id"]),
            "court_name": _court["court_name"],
        })

print(
    f"Справочник СК загружен: "
    f"{len(SK_COURTS_DIRECTORY)} регионов/разделов, "
    f"{len(SK_COURTS_FLAT)} судов"
)



# ============================================================
# 1. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================

def ensure_dirs():
    for p in [WORK_DIR, DOCX_DIR, PDF_DIR]:
        p.mkdir(parents=True, exist_ok=True)


def write_embedded_template():
    ensure_dirs()
    return TEMPLATE_PATH


def normalize_text(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\xa0", " ")).strip()


def normalize_match(value):
    value = normalize_text(value).lower()
    value = value.replace("ё", "е")
    value = re.sub(r"[«»\"'`]", "", value)
    value = re.sub(r"[^0-9a-zа-яәіңғүұқөһ\s()-]+", " ", value, flags=re.I)
    return re.sub(r"\s+", " ", value).strip()


def format_money_kz(value):
    if value is None or pd.isna(value):
        return "0,00"
    x = float(value)
    s = f"{x:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", " ")


def safe_filename(value):
    s = normalize_text(value)
    s = re.sub(r'[<>:"/\\|?*]+', "_", s)
    return s[:120].strip(" ._")


def build_letter_text(row, project_cfg):
    sud = normalize_text(row["Sud"])
    fio = normalize_text(row["FIO"])
    iin = normalize_text(row["IIN"])
    judge = normalize_text(row.get("Judge", ""))
    project_name = project_cfg["project_name"]
    defendants = f"1. {fio}, ИИН {iin}"

    return (
        f"В {sud}\n"
        f"От: Директора {project_name}\n"
        f"{project_cfg['director']}\n"
        f"Судье: {judge}\n\n"
        f"ЗАЯВЛЕНИЕ\n\n"
        f"В производстве судебного органа – {sud} находятся гражданские дела по иску "
        f"{project_name} о взыскании задолженности к следующим ответчикам:\n"
        f"{defendants}\n\n"
        f"На данный момент, в отношении указанных должников были вынесены решения суда. "
        f"В связи с этим, просим предоставить исполнительные листы по указанным ответчикам, "
        f"в соответствии с ранее вынесенными решениями суда.\n\n"
        f"ПРОШУ:\n\n"
        f"Выписать исполнительный лист по Решению суда о взыскании задолженности в пользу "
        f"{project_name} и закрепить исполнительный документ за региональной палатой "
        f"частных судебных исполнителей посредством информационной системы «Төрелік», "
        f"по месту регистрации Ответчика.\n\n"
        f"Директор {project_name} {project_cfg['director']}"
    )

# WordprocessingML namespace для OOXML-шаблона.
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
}


def _replace_placeholder_across_text_nodes(nodes, placeholder, replacement):
    """
    Заменяет placeholder даже если Word разорвал его на несколько <w:t>.
    Форматирование первого текстового узла сохраняется.
    """
    if not nodes:
        return False

    full_text = "".join((node.text or "") for node in nodes)
    if placeholder not in full_text:
        return False

    new_text = full_text.replace(placeholder, replacement)

    # Весь новый текст кладём в первый w:t, остальные очищаем.
    nodes[0].text = new_text
    for node in nodes[1:]:
        node.text = ""

    return True


def fill_docx_ooxml(template_path, output_path, replacements):
    """
    Меняет переменные непосредственно в OOXML.
    Важно: python-docx здесь НЕ используется, потому что он может потерять
    плавающий графический объект подписи/печати.
    """
    with zipfile.ZipFile(template_path, "r") as zin, \
         zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:

        for item in zin.infolist():
            data = zin.read(item.filename)

            if item.filename == "word/document.xml":
                root = etree.fromstring(data)

                for p in root.xpath(".//w:p", namespaces=NS):
                    nodes = p.xpath(".//w:t", namespaces=NS)
                    if not nodes:
                        continue

                    for key, value in replacements.items():
                        _replace_placeholder_across_text_nodes(
                            nodes, key, str(value)
                        )

                data = etree.tostring(
                    root,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone="yes"
                )

            zout.writestr(item, data)


def convert_docx_to_pdf(docx_path, pdf_path):
    """
    Сначала Microsoft Word COM (лучше сохраняет верстку).
    Если pywin32/Word недоступны — пробуем LibreOffice.
    """
    docx_path = Path(docx_path).resolve()
    pdf_path = Path(pdf_path).resolve()
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    # Вариант 1: MS Word
    try:
        import win32com.client

        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        try:
            doc = word.Documents.Open(str(docx_path))
            # 17 = wdFormatPDF
            doc.SaveAs(str(pdf_path), FileFormat=17)
            doc.Close(False)
        finally:
            word.Quit()

        if pdf_path.exists() and pdf_path.stat().st_size > 0:
            return pdf_path

    except Exception as e:
        print(f"Word PDF fallback: {e}")

    # Вариант 2: LibreOffice
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError(
            "Не удалось конвертировать DOCX в PDF. "
            "Нужен Microsoft Word + pywin32 или LibreOffice."
        )

    subprocess.run(
        [
            soffice,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(pdf_path.parent),
            str(docx_path)
        ],
        check=True,
        capture_output=True
    )

    generated = pdf_path.parent / (docx_path.stem + ".pdf")
    if generated != pdf_path and generated.exists():
        generated.replace(pdf_path)

    if not pdf_path.exists():
        raise RuntimeError(f"PDF не создан: {pdf_path}")

    return pdf_path


def make_documents(row):
    project_cfg = PROJECT_CFG
    template_path = write_embedded_template()

    fio = normalize_text(row["FIO"])
    iin_text = normalize_text(row["IIN"])
    fio_filename = safe_filename(fio)
    iin_filename = safe_filename(iin_text)

    base_name = f"Заявление_на_выдачу_ИЛ, {fio_filename}, {iin_filename}"
    docx_path = DOCX_DIR / f"{base_name}.docx"
    pdf_path = PDF_DIR / f"{base_name}.pdf"

    replacements = {
        "{Sud}": normalize_text(row["Sud"]),
        "{ProjectName}": project_cfg["project_name"],
        "{ProjectDirector}": project_cfg["director"],
        "{Judge}": normalize_text(row.get("Judge", "")),
        "{DefendantsData}": f"1. {fio}, ИИН {iin_text}",
        "{Pechat}": "",
    }

    fill_docx_ooxml(template_path, docx_path, replacements)
    convert_docx_to_pdf(docx_path, pdf_path)

    text_for_sk = build_letter_text(row, project_cfg)
    return docx_path, pdf_path, text_for_sk


# ============================================================
# 3. БД
# ============================================================

def get_db_connection():
    # На разных машинах установлен разный ODBC-драйвер (17 или 18) — пробуем оба,
    # как в otmeny.py / reestr_gosposhliny.py.
    last_err = None
    for driver in ("ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server"):
        conn_str = (
            f"DRIVER={{{driver}}};"
            f"SERVER={DB_SERVER};"
            f"DATABASE={DB_DATABASE};"
            f"UID={DB_USERNAME};"
            f"PWD={DB_PASSWORD};"
            "TrustServerCertificate=yes;"
            "Encrypt=no;"
        )
        try:
            return pyodbc.connect(conn_str, timeout=30)
        except pyodbc.Error as e:
            last_err = e
    raise last_err


def load_source():
    print("Подключение к БД...")
    conn = get_db_connection()

    try:
        df = pd.read_sql(SQL_QUERY, conn)
    finally:
        conn.close()

    df["EID"] = df["EID"].astype(str)
    df["IIN"] = df["IIN"].astype(str).str.replace(r"\.0$", "", regex=True)
    df["Sud"] = df["Sud"].apply(normalize_text)
    df["FIO"] = df["FIO"].apply(normalize_text)
    df["DebtRestWithoutStateDuty"] = ""

    df = df.drop_duplicates(subset=["LoanID"]).reset_index(drop=True)

    if SKIP_ALREADY_CREATED and LOG_FILE.exists():
        try:
            old = pd.read_excel(LOG_FILE)
            done = set(
                old.loc[
                    (
                        old["Status"].isin([
                            "SAVED_FOR_SIGNING",
                            "SAVED_FOR_SIGNING_VERIFIED"
                        ])
                        |
                        (
                            old["LetterNumber"].fillna("").astype(str).str.strip().ne("")
                            &
                            old["SignURL"].fillna("").astype(str).str.strip().ne("")
                        )
                    ),
                    "LoanID"
                ].astype(str)
            )
            before = len(df)
            df = df[~df["LoanID"].astype(str).isin(done)].reset_index(drop=True)
            print(f"Уже подготовленных сделок пропущено: {before - len(df)}")
        except Exception as e:
            print(f"Не удалось прочитать старый лог: {e}")

    if MAX_ROWS is not None:
        df = df.head(int(MAX_ROWS)).copy()

    print(f"К обработке: {len(df)}")
    return df


# ============================================================
# 4. АВТОРИЗАЦИЯ: SELENIUM ТОЛЬКО ДЛЯ ВХОДА
# ============================================================

def build_login_driver():
    opts = Options()
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-notifications")
    opts.add_argument("--disable-popup-blocking")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(options=opts)
    driver.set_page_load_timeout(PAGELOAD_TIMEOUT)
    return driver


def switch_to_ru(driver):
    try:
        els = driver.find_elements(By.LINK_TEXT, "РУС")
        if els:
            els[0].click()
            time.sleep(1)
    except Exception:
        pass


_ATTACHED_DRIVER = None


def _looks_like_login_page(html_text):
    """Определяет, что сервер вернул страницу входа, даже если URL не изменился."""
    sp = BeautifulSoup(html_text or "", "html.parser")

    # Самый надёжный признак — поле password.
    if sp.find("input", attrs={"type": "password"}) is not None:
        return True

    # Формы вида ...:auth
    for form in sp.find_all("form"):
        fid = (form.get("id") or form.get("name") or "").lower()
        if ":auth" in fid or fid.endswith("auth"):
            return True

    low = (html_text or "").lower()
    markers = (
        'placeholder="пароль"',
        "placeholder='пароль'",
        'placeholder="құпия сөз"',
        "placeholder='құпия сөз'",
    )
    return any(m in low for m in markers)


def _browser_is_logged_in(driver):
    """
    Проверяет авторизацию прямо внутри Chrome, не по requests.
    """
    html_text = driver.page_source or ""
    if _looks_like_login_page(html_text):
        return False

    # Дополнительный позитивный признак: мы на office.sud.kz и нет auth-формы.
    return (driver.current_url or "").startswith(BASE_URL)


def _get_all_office_cookies_from_chrome(driver):
    """
    Получает ВСЕ cookies office.sud.kz через CDP.
    Это надёжнее driver.get_cookies(), который может вернуть только cookies
    текущего контекста/пути.
    """
    cookies = []

    # 1) Основной способ — CDP Network.getAllCookies
    try:
        data = driver.execute_cdp_cmd("Network.getAllCookies", {})
        for c in data.get("cookies", []):
            domain = (c.get("domain") or "").lstrip(".").lower()
            if domain == "office.sud.kz" or domain.endswith(".office.sud.kz"):
                cookies.append(c)
    except Exception as e:
        print(f"  ! CDP Network.getAllCookies не сработал: {e!r}")

    # 2) Fallback/дополнение — Selenium cookies текущей вкладки
    try:
        seen = {(c.get("name"), c.get("domain"), c.get("path")) for c in cookies}
        for c in driver.get_cookies():
            domain = (c.get("domain") or "").lstrip(".").lower()
            if domain == "office.sud.kz" or domain.endswith(".office.sud.kz"):
                key = (c.get("name"), c.get("domain"), c.get("path"))
                if key not in seen:
                    cookies.append(c)
                    seen.add(key)
    except Exception as e:
        print(f"  ! driver.get_cookies() не сработал: {e!r}")

    return cookies


def _put_chrome_cookies_into_requests(session, cookies):
    """
    Копирует cookies Chrome в requests.Session.
    Добавляет как доменную версию, так и host-only fallback для office.sud.kz.
    """
    session.cookies.clear()

    for c in cookies:
        name = c.get("name")
        value = c.get("value")
        if not name:
            continue

        path = c.get("path") or "/"
        domain = c.get("domain") or "office.sud.kz"

        # Нормальная доменная cookie.
        try:
            session.cookies.set(
                name,
                value,
                domain=domain,
                path=path,
            )
        except Exception:
            pass

        # Host-only fallback. Нужен на случай, если браузерная domain-cookie
        # в requests трактуется иначе.
        try:
            session.cookies.set(
                name,
                value,
                domain="office.sud.kz",
                path=path,
            )
        except Exception:
            pass


def _find_login_elements(driver):
    """
    Надёжно находит логин, пароль и кнопку входа на рус/каз интерфейсе СК.
    Поддерживает поле ИИН с type=email, text, tel, number и без type.
    """
    wait = WebDriverWait(driver, LOADING_TIME)

    password_input = wait.until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "input[type='password']"))
    )

    form = password_input.find_element(By.XPATH, "./ancestor::form[1]")

    # 1. Сначала ищем по смысловым атрибутам.
    login_input = None
    semantic_selectors = [
        "input[placeholder*='ИИН' i]",
        "input[placeholder*='ЖСН' i]",
        "input[placeholder*='БСН' i]",
        "input[name*='iin' i]",
        "input[id*='iin' i]",
        "input[name*='login' i]",
        "input[id*='login' i]",
        "input[type='email']",
    ]

    for selector in semantic_selectors:
        try:
            for el in form.find_elements(By.CSS_SELECTOR, selector):
                if el.is_displayed() and el.is_enabled() and el != password_input:
                    login_input = el
                    break
        except Exception:
            pass
        if login_input is not None:
            break

    # 2. Fallback — любой подходящий видимый input той же формы.
    if login_input is None:
        inputs = form.find_elements(By.CSS_SELECTOR, "input")
        allowed_types = ("", "text", "email", "tel", "number")
        for el in inputs:
            try:
                typ = (el.get_attribute("type") or "").lower()
                if (
                    el.is_displayed()
                    and el.is_enabled()
                    and el != password_input
                    and typ in allowed_types
                ):
                    login_input = el
                    break
            except Exception:
                pass

    if login_input is None:
        raise RuntimeError(
            "Не найдено поле ИИН/логина в форме авторизации "
            "(проверены ИИН/ЖСН/БСН/login и type=email/text/tel/number)."
        )

    # Кнопка входа.
    submit_element = None
    buttons = form.find_elements(
        By.CSS_SELECTOR,
        "button, input[type='submit'], input[type='button']"
    )

    preferred_words = ("кіру", "войти", "вход", "login", "sign in")
    for el in buttons:
        try:
            label = (
                (el.text or "")
                or (el.get_attribute("value") or "")
                or (el.get_attribute("aria-label") or "")
            ).strip().lower()
            if el.is_displayed() and el.is_enabled() and any(w in label for w in preferred_words):
                submit_element = el
                break
        except Exception:
            pass

    if submit_element is None:
        for el in buttons:
            try:
                if el.is_displayed() and el.is_enabled():
                    submit_element = el
                    break
            except Exception:
                pass

    if submit_element is None:
        raise RuntimeError("Не найдена кнопка входа в форме авторизации.")

    print(
        "  ✓ Поле логина найдено: "
        f"type={login_input.get_attribute('type')!r}, "
        f"id={login_input.get_attribute('id')!r}, "
        f"placeholder={login_input.get_attribute('placeholder')!r}"
    )

    return login_input, password_input, submit_element

def login_and_make_session():
    """
    v24 FULL AUTO:
    1) сам запускает новый Chrome;
    2) сам открывает Судебный кабинет;
    3) сам вводит USER_AUTH / USER_PASSWORD и выполняет вход;
    4) открывает letter/info.xhtml;
    5) переносит cookies в requests.Session;
    6) проверяет, что requests действительно авторизован;
    7) закрывает Chrome — дальше весь процесс идёт через requests.
    """
    global _ATTACHED_DRIVER

    print("Автоматический запуск Chrome и вход в Судебный кабинет...")
    driver = build_login_driver()
    _ATTACHED_DRIVER = driver

    try:
        driver.get(LOGIN_URL)

        # Если сайт уже каким-то образом открыл авторизованную страницу —
        # повторный вход не нужен.
        if _looks_like_login_page(driver.page_source or ""):
            print("→ Выполняю автоматический вход...")

            login_input, password_input, submit_element = _find_login_elements(driver)

            login_input.clear()
            login_input.send_keys(USER_AUTH)

            password_input.clear()
            password_input.send_keys(USER_PASSWORD)

            submit_element.click()

            # Ждём исчезновения формы входа.
            WebDriverWait(driver, LOADING_TIME).until(
                lambda d: not _looks_like_login_page(d.page_source or "")
            )

        if not _browser_is_logged_in(driver):
            raise RuntimeError(
                "После автоматического ввода логина и пароля Судебный кабинет "
                "не подтвердил авторизацию."
            )

        print("✓ Автоматическая авторизация выполнена.")

        # Язык не критичен для v23/v24, но стараемся переключить на русский.
        switch_to_ru(driver)

        print("→ Открываю раздел «Отправка писем»...")
        driver.get(LETTER_INFO_URL)
        WebDriverWait(driver, LOADING_TIME).until(
            lambda d: d.execute_script("return document.readyState") in ("interactive", "complete")
        )
        time.sleep(1.0)

        browser_html = driver.page_source or ""
        if _looks_like_login_page(browser_html):
            raise RuntimeError(
                "После перехода в «Отправка писем» Судебный кабинет снова показал форму входа."
            )

        print(f"✓ Авторизованный раздел открыт: {driver.current_url}")

        cookies = _get_all_office_cookies_from_chrome(driver)
        if not cookies:
            raise RuntimeError("После входа не удалось получить cookies office.sud.kz.")

        print(f"✓ Cookies получены: {len(cookies)}")
        print("  Имена cookies:", ", ".join(sorted({c.get("name", "?") for c in cookies})))

        ua = driver.execute_script("return navigator.userAgent")
        lang = driver.execute_script("return navigator.language || 'ru-RU'")

        session = requests.Session()
        session.headers.update({
            "User-Agent": ua,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
            "Accept-Language": f"{lang},ru;q=0.9,kk;q=0.8,en;q=0.7",
            "Referer": driver.current_url,
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
        })
        _put_chrome_cookies_into_requests(session, cookies)

        print("→ Проверяю перенесённую HTTP-сессию...")
        r = session.get(
            LETTER_INFO_URL,
            timeout=HTTP_TIMEOUT,
            allow_redirects=True,
        )
        r.raise_for_status()

        if _looks_like_login_page(r.text or ""):
            debug_auth = WORK_DIR / "DEBUG_AUTH_requests_got_login.html"
            debug_auth.write_text(r.text or "", encoding="utf-8")
            raise RuntimeError(
                "Chrome вошёл успешно, но requests.Session получил страницу входа. "
                f"Ответ сохранён: {debug_auth}"
            )

        print(f"✓ HTTP-сессия действительно авторизована: {r.url}")
        print("✓ Дальше скрипт работает через requests.Session.")
        return session

    finally:
        # Selenium нужен только для входа. После переноса сессии браузер закрываем.
        try:
            driver.quit()
            print("✓ Chrome после авторизации закрыт.")
        except Exception:
            pass
        _ATTACHED_DRIVER = None

# ============================================================
# 5. JSF / RICHFACES
# ============================================================

AJAX_HEADERS = {
    "Faces-Request": "partial/ajax",
    "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
}


RETRYABLE_EXCEPTIONS = (
    requests.exceptions.ReadTimeout,
    requests.exceptions.ConnectTimeout,
    requests.exceptions.ConnectionError,
)

def http_get(session, url, *, timeout=None, retry=True, **kwargs):
    attempts = HTTP_RETRIES if retry else 1
    timeout = timeout or HTTP_TIMEOUT
    last_error = None
    for attempt in range(1, attempts + 1):
        try:
            return session.get(url, timeout=timeout, **kwargs)
        except RETRYABLE_EXCEPTIONS as e:
            last_error = e
            if attempt >= attempts:
                raise
            wait = 2 * attempt
            print(f"  GET timeout/connection error, попытка {attempt}/{attempts}. Повтор через {wait} сек...")
            time.sleep(wait)
    raise last_error

def http_post_ajax(session, url, *, data=None, headers=None, timeout=None, retry_safe=False, description="POST", **kwargs):
    attempts = HTTP_RETRIES if retry_safe else 1
    timeout = timeout or HTTP_TIMEOUT
    last_error = None
    for attempt in range(1, attempts + 1):
        try:
            return session.post(url, data=data, headers=headers, timeout=timeout, **kwargs)
        except RETRYABLE_EXCEPTIONS as e:
            last_error = e
            if attempt >= attempts:
                raise
            wait = 2 * attempt
            print(f"  {description}: timeout, попытка {attempt}/{attempts}. Повтор через {wait} сек...")
            time.sleep(wait)
    raise last_error


def soup(text):
    return BeautifulSoup(text or "", "html.parser")


def normalize_js_identifier(value):
    if not value:
        return ""
    value = html.unescape(str(value))
    value = re.sub(
        r"\\u([0-9a-fA-F]{4})",
        lambda m: chr(int(m.group(1), 16)),
        value
    )
    return value.replace("\\/", "/")


def extract_redirect(text):
    m = re.search(
        r'<redirect\s+url=["\']([^"\']+)["\']',
        text or "",
        flags=re.I
    )
    if not m:
        return ""
    return html.unescape(m.group(1))


def extract_viewstate_from_partial(text, fallback=""):
    patterns = [
        r'<update[^>]+id=["\'][^"\']*javax\.faces\.ViewState[^"\']*["\'][^>]*>'
        r'<!\[CDATA\[(.*?)\]\]>',
        r'<update[^>]+id=["\'][^"\']*javax\.faces\.ViewState[^"\']*["\'][^>]*>'
        r'(.*?)</update>',
    ]

    for pat in patterns:
        m = re.search(pat, text or "", flags=re.S | re.I)
        if m:
            return html.unescape(m.group(1)).strip()

    return fallback


def parse_partial_update(text, id_suffix):
    """
    Возвращает CDATA/HTML из <update id="...suffix">...</update>.
    """
    for m in re.finditer(
        r'<update\s+id=["\']([^"\']+)["\']>(.*?)</update>',
        text or "",
        flags=re.S | re.I
    ):
        upd_id = html.unescape(m.group(1))
        body = m.group(2)

        if upd_id.endswith(id_suffix) or id_suffix in upd_id:
            cdata = re.search(r'<!\[CDATA\[(.*?)\]\]>', body, flags=re.S)
            return cdata.group(1) if cdata else html.unescape(body)

    return ""


def get_main_form(html_text, marker=None):
    sp = soup(html_text)

    if marker:
        el = sp.find(attrs={"name": re.compile(re.escape(marker) + r"$")})
        if not el:
            el = sp.find(id=re.compile(re.escape(marker) + r"$"))

        if el:
            form = el.find_parent("form")
            if form:
                return form

    return None


def get_viewstate_from_form(form):
    el = form.find("input", attrs={"name": "javax.faces.ViewState"})
    if not el:
        raise RuntimeError("В форме не найден javax.faces.ViewState")
    return el.get("value", "")


def make_richfaces_payload(form_id, viewstate, source, extra_fields=None,
                           partial_event="click", behavior_event=None):
    data = {
        form_id: form_id,
        "javax.faces.ViewState": viewstate,
        "javax.faces.source": source,
        "javax.faces.partial.execute": f"{source} @component",
        "javax.faces.partial.render": "@component",
        "org.richfaces.ajax.component": source,
        source: source,
        "rfExt": "null",
        "AJAX:EVENTS_COUNT": "1",
        "javax.faces.partial.ajax": "true",
    }

    if partial_event:
        data["javax.faces.partial.event"] = partial_event

    if behavior_event:
        data["javax.faces.behavior.event"] = behavior_event

    if extra_fields:
        # Поля формы должны идти вместе с событием.
        d = {}
        d.update(extra_fields)
        d.update(data)
        data = d

    return data



def _control_label(el):
    """Читаемый текст/подпись JSF-кнопки или ссылки."""
    if el is None:
        return ""
    parts = [
        el.get("value", ""),
        el.get("title", ""),
        el.get("aria-label", ""),
        el.get_text(" ", strip=True),
    ]
    return normalize_text(" ".join(str(x) for x in parts if x))


def _has_viewstate(form):
    return form.find(
        "input",
        attrs={"name": "javax.faces.ViewState"}
    ) is not None


def _extract_ajax_source(el):
    """Извлекает JSF/RichFaces source из id/name/onclick/href."""
    if el is None:
        return ""

    direct_id = el.get("id") or el.get("name") or ""
    js_blob = " ".join([
        str(el.get("onclick") or ""),
        str(el.get("href") or ""),
        str(el.get("onmousedown") or ""),
    ])

    patterns = [
        r"""source\s*:\s*['"]([^'"]+)['"]""",
        r"""javax\.faces\.source['"]?\s*[:=]\s*['"]([^'"]+)['"]""",
        r"""RichFaces\.ajax\(\s*['"]([^'"]+)['"]""",
        r"""mojarra\.ab\(\s*['"]([^'"]+)['"]""",
        r"""jsf\.ajax\.request\(\s*['"]([^'"]+)['"]""",
    ]

    for pattern in patterns:
        m = re.search(pattern, js_blob, flags=re.I)
        if m:
            return m.group(1)

    return direct_id


def _is_auth_like(form_id, source, label):
    text = " ".join([
        str(form_id or ""),
        str(source or ""),
        str(label or ""),
    ]).lower()

    return any(marker in text for marker in (
        ":auth", "auth:", "login", "войти", "кіру",
        "авториза", "sign in",
    ))


def _find_js_function_ajax_source(sp, function_name):
    """
    Ищет определение вида:
      send=function(){RichFaces.ajax("FORM:SOURCE", ...)}
    и возвращает (form, source).

    Это именно текущая архитектура Судебного кабинета:
    видимая кнопка вызывает send(), а реальный RichFaces source
    находится в скрытом span/script внутри отдельной JSF-формы.
    """
    if not function_name:
        return None, ""

    fname = re.escape(function_name)

    # Ищем по всем script-тегам.
    for script in sp.find_all("script"):
        text = script.get_text(" ", strip=False) or ""

        # functionName = function(...) { ... RichFaces.ajax("source" ...)
        patterns = [
            rf"""{fname}\s*=\s*function\s*\([^)]*\)\s*\{{.*?RichFaces\.ajax\(\s*["']([^"']+)["']""",
            rf"""function\s+{fname}\s*\([^)]*\)\s*\{{.*?RichFaces\.ajax\(\s*["']([^"']+)["']""",
        ]

        source = ""
        for pat in patterns:
            m = re.search(pat, text, flags=re.I | re.S)
            if m:
                source = m.group(1)
                break

        if not source:
            continue

        # Скрипт обычно лежит в span внутри нужной JSF-формы.
        form = script.find_parent("form")
        if form is not None:
            return form, source

        # Fallback: найдём форму по префиксу source.
        for f in sp.find_all("form"):
            fid = f.get("id") or f.get("name") or ""
            if fid and source.startswith(fid + ":"):
                return f, source

    return None, ""


def _find_create_letter_control(sp):
    """
    v22: определяет создание письма через реальную структуру страницы.

    На letter/info.xhtml видимая кнопка:
        <button onclick="send();">Хат жіберу</button>

    А реальный JSF source находится отдельно:
        send=function(){RichFaces.ajax("j_idt33:j_idt36:j_idt37",...)}

    Поэтому больше НЕ пытаемся угадывать source среди обычных submit-кнопок
    модальных окон.
    """
    candidates = []

    # 1) Главный способ: видимая кнопка, вызывающая JS-функцию.
    for el in sp.find_all(["button", "a", "input"]):
        label = _control_label(el)
        onclick = str(el.get("onclick") or "").strip()

        # Для input допускаем только реальные кнопки.
        if el.name == "input":
            typ = (el.get("type") or "").lower()
            if typ not in ("button", "submit", "image"):
                continue

        # Из onclick="send();" извлекаем имя функции.
        m = re.match(r"""^\s*([A-Za-z_$][\w$]*)\s*\(\s*\)\s*;?\s*(?:return\s+false\s*;?)?\s*$""", onclick)
        if not m:
            continue

        function_name = m.group(1)

        # Отсекаем явно служебные функции.
        low_fn = function_name.lower()
        if any(x in low_fn for x in (
            "hide", "showfeedback", "logout", "language",
            "modal", "video", "error", "qr",
        )):
            continue

        form, source = _find_js_function_ajax_source(sp, function_name)
        if form is None or not source:
            continue

        form_id = form.get("id") or form.get("name") or ""

        if _is_auth_like(form_id, source, label):
            continue

        score = 0
        low_label = label.lower()
        if any(k in low_label for k in (
            "хат жіберу",       # казахский
            "отправить письмо", # русский
            "отправка письма",
            "создать письмо",
        )):
            score += 1000

        if function_name.lower() in ("send", "sendletter"):
            score += 500

        if form_id and source.startswith(form_id + ":"):
            score += 200

        candidates.append({
            "score": score,
            "form": form,
            "control": el,
            "form_id": form_id,
            "source": source,
            "label": label,
            "ajax": True,
            "tag": el.name,
            "type": (el.get("type") or "").lower(),
            "function": function_name,
        })

    if candidates:
        candidates.sort(key=lambda x: x["score"], reverse=True)
        best = candidates[0]
        return (
            best["form"],
            best["control"],
            best["source"],
            candidates,
        )

    # 2) Специальный fallback для текущего СК:
    # функция send() может быть определена, даже если button распознан необычно.
    form, source = _find_js_function_ajax_source(sp, "send")
    if form is not None and source:
        form_id = form.get("id") or form.get("name") or ""
        fake_control = sp.find(
            lambda tag: (
                tag.name in ("button", "a", "input")
                and "send()" in str(tag.get("onclick") or "").replace(" ", "")
            )
        )
        candidates = [{
            "score": 900,
            "form": form,
            "control": fake_control,
            "form_id": form_id,
            "source": source,
            "label": _control_label(fake_control) if fake_control else "send()",
            "ajax": True,
            "tag": fake_control.name if fake_control else "script",
            "type": (fake_control.get("type") or "") if fake_control else "",
            "function": "send",
        }]
        return form, fake_control, source, candidates

    return None, None, "", []

def open_new_letter(session):
    """
    Создаёт новое письмо через текущие JSF/RichFaces IDs.
    Фиксированные j_idtXX не используются.
    """
    print("  [1/3] GET letter/info.xhtml ...")

    r = http_get(
        session,
        LETTER_INFO_URL,
        timeout=60,
        retry=True
    )
    r.raise_for_status()

    sp = soup(r.text)
    form, control, source, candidates = _find_create_letter_control(sp)

    if form is None or control is None or not source:
        print("  ! Динамический поиск не нашёл безопасную кнопку создания письма.")
        if candidates:
            print("  Кандидаты (без auth), TOP-10:")
            for cand in candidates[:10]:
                print(
                    f"    score={cand.get('score')} "
                    f"tag={cand.get('tag')} type={cand.get('type')} "
                    f"form={cand.get('form_id')} "
                    f"source={cand.get('source')} "
                    f"label={cand.get('label')!r} "
                    f"function={cand.get('function')!r} "
                    f"ajax={cand.get('ajax')}"
                )

        debug_path = WORK_DIR / "DEBUG_letter_info.xhtml.html"
        debug_path.write_text(r.text, encoding="utf-8")

        debug_candidates = WORK_DIR / "DEBUG_letter_info_controls.txt"
        lines = []
        for f in sp.find_all("form"):
            fid = f.get("id") or f.get("name") or ""
            has_vs = _has_viewstate(f)
            lines.append(f"FORM: {fid!r} | ViewState={has_vs}")
            for el in f.find_all(["input", "button", "a"]):
                eid = el.get("id") or el.get("name") or ""
                if not eid:
                    continue
                lines.append(
                    f"  {el.name} id={eid!r} type={el.get('type')!r} "
                    f"label={_control_label(el)!r} "
                    f"onclick={(el.get('onclick') or '')[:300]!r}"
                )
        debug_candidates.write_text("\n".join(lines), encoding="utf-8")

        raise RuntimeError(
            "На letter/info.xhtml не удалось динамически определить "
            "кнопку создания письма. HTML и список контролов сохранены:\n"
            f"  {debug_path}\n"
            f"  {debug_candidates}"
        )

    form_id = form.get("id") or form.get("name")
    viewstate = get_viewstate_from_form(form)

    if not form_id or not source or not viewstate:
        raise RuntimeError(
            "Форма создания письма найдена, но не удалось получить "
            "form_id/source/ViewState."
        )

    print(f"  ✓ JSF форма определена динамически: {form_id}")
    print(f"  ✓ JSF source определён динамически: {source}")
    label = _control_label(control)
    if label:
        print(f"    Кнопка: {label}")

    payload = {
        form_id: form_id,
        "javax.faces.ViewState": viewstate,
        "javax.faces.source": source,
        "javax.faces.partial.execute": f"{source} @component",
        "javax.faces.partial.render": "@component",
        "org.richfaces.ajax.component": source,
        source: source,
        "rfExt": "null",
        "AJAX:EVENTS_COUNT": "1",
        "javax.faces.partial.ajax": "true",
    }

    print(
        f"  [2/3] POST create letter "
        f"(form={form_id}, source={source}) ..."
    )

    try:
        rr = session.post(
            LETTER_INFO_URL,
            data=payload,
            headers={
                **AJAX_HEADERS,
                "Referer": LETTER_INFO_URL,
            },
            timeout=(20, 45)
        )
    except requests.exceptions.ReadTimeout as e:
        raise RuntimeError(
            "Судебный кабинет не ответил на POST создания нового письма "
            "за 45 секунд. Автоповтор отключён, чтобы не создать дубль."
        ) from e

    rr.raise_for_status()

    ctype = rr.headers.get("Content-Type", "")
    print(
        f"        status={rr.status_code}, "
        f"Content-Type={ctype}, len={len(rr.text)}"
    )

    redirect = extract_redirect(rr.text)
    if not redirect:
        redirect = extract_redirect(html.unescape(rr.text or ""))

    if not redirect:
        debug_resp = WORK_DIR / "DEBUG_create_letter_response.txt"
        debug_resp.write_text(rr.text or "", encoding="utf-8")
        raise RuntimeError(
            "POST создания письма выполнен, но redirect на send.xhtml "
            "не найден. Ответ сохранён:\n"
            f"  {debug_resp}"
        )

    send_url = urljoin(BASE_URL, redirect)

    print("  [3/3] GET send.xhtml ...")
    rs = http_get(
        session,
        send_url,
        timeout=60,
        retry=True
    )
    rs.raise_for_status()

    if not re.search(r"district-field", rs.text or "", flags=re.I):
        debug_send = WORK_DIR / "DEBUG_send_after_create.xhtml.html"
        debug_send.write_text(rs.text or "", encoding="utf-8")
        raise RuntimeError(
            "Страница send.xhtml открылась, но поле региона не найдено. "
            f"HTML сохранён: {debug_send}"
        )

    print("  ✓ Новый черновик открыт.")
    return send_url, rs.text

def _find_by_name_suffix(form, tag, suffix):
    return form.find(
        tag,
        attrs={"name": re.compile(re.escape(suffix) + r"$")}
    )


def _find_component_by_semantic_id(form, semantic_name):
    """
    Находит RichFaces-компонент по стабильной семантической части имени,
    не завязываясь на родительские j_idtXX.
    """
    for tag in ("div", "span"):
        el = form.find(
            tag,
            id=re.compile(re.escape(semantic_name) + r"$")
        )
        if el is not None:
            return el

    el = form.find(
        attrs={"id": re.compile(re.escape(semantic_name), re.I)}
    )
    if el is not None:
        cur = el
        while cur is not None and cur is not form:
            cid = cur.get("id") or ""
            if cid.endswith(semantic_name):
                return cur
            cur = cur.parent
        return el

    return form.find(
        attrs={"name": re.compile(re.escape(semantic_name), re.I)}
    )


def _find_next_control(form):
    """
    Ищет кнопку перехода дальше на русском и казахском языках.

    Рус:  Далее
    Каз:  Ары қарай
    Доп. fallback: Келесі / Продолжить
    """
    accepted = (
        "далее",
        "ары қарай",
        "келесі",
        "продолжить",
    )

    candidates = []

    for el in form.find_all(["input", "button", "a"]):
        label = _control_label(el)
        if not label:
            continue

        low = normalize_text(label).lower()
        eid = el.get("id") or el.get("name") or ""
        if not eid:
            continue

        # Точное/содержательное совпадение.
        if any(word in low for word in accepted):
            candidates.append(el)

    if not candidates:
        return None

    # Предпочитаем submit/button, затем ссылку.
    candidates.sort(
        key=lambda el: (
            0 if el.name in ("input", "button") else 1,
            0 if (el.get("type") or "").lower() in ("submit", "button") else 1,
        )
    )
    return candidates[0]

def parse_send_context(send_url, html_text):
    sp = soup(html_text)

    district = sp.find(
        "select",
        attrs={"name": re.compile(r"district-field$", re.I)}
    )
    court = sp.find(
        "select",
        attrs={"name": re.compile(r"court-field$", re.I)}
    )
    text_field = sp.find(
        "textarea",
        attrs={"name": re.compile(r"text-field$", re.I)}
    )

    if not district or not court or not text_field:
        debug_path = WORK_DIR / "DEBUG_send_missing_fields.xhtml.html"
        debug_path.write_text(html_text or "", encoding="utf-8")
        raise RuntimeError(
            "Не удалось определить district/court/text поля формы. "
            f"HTML сохранён: {debug_path}"
        )

    form = district.find_parent("form")
    if form is None:
        raise RuntimeError(
            "Для district-field не найдена родительская JSF-форма."
        )

    form_id = form.get("id") or form.get("name")
    if not form_id:
        raise RuntimeError(
            "У JSF-формы send.xhtml отсутствует id/name."
        )

    request_hid = form.find(
        "input",
        attrs={"name": re.compile(r"requestScanHid$", re.I)}
    )

    next_btn = _find_next_control(form)
    req_uploader = _find_component_by_semantic_id(
        form, "selectRequestScanUploader"
    )
    attach_uploader = _find_component_by_semantic_id(
        form, "selectFileUploader"
    )

    missing = []
    if next_btn is None:
        missing.append("Далее")
    if req_uploader is None:
        missing.append("selectRequestScanUploader")

    if missing:
        debug_path = WORK_DIR / "DEBUG_send_components.xhtml.html"
        debug_path.write_text(html_text or "", encoding="utf-8")

        debug_txt = WORK_DIR / "DEBUG_send_components.txt"
        details = [
            f"form_id={form_id}",
            f"missing={missing}",
            "",
            "FILE INPUTS:",
        ]
        for el in form.find_all("input", attrs={"type": "file"}):
            details.append(
                f"id={el.get('id')!r} name={el.get('name')!r}"
            )

        details.append("")
        details.append("BUTTONS/LINKS:")
        for el in form.find_all(["input", "button", "a"]):
            eid = el.get("id") or el.get("name") or ""
            if eid:
                details.append(
                    f"{el.name} id={eid!r} "
                    f"label={_control_label(el)!r}"
                )

        debug_txt.write_text("\n".join(details), encoding="utf-8")

        raise RuntimeError(
            "Не найдены компоненты формы: "
            + ", ".join(missing)
            + ". Диагностика сохранена:\n"
            + f"  {debug_path}\n"
            + f"  {debug_txt}"
        )

    next_source = next_btn.get("id") or next_btn.get("name")
    req_uploader_id = req_uploader.get("id") or req_uploader.get("name")
    attach_uploader_id = (
        (attach_uploader.get("id") or attach_uploader.get("name"))
        if attach_uploader is not None
        else ""
    )

    ctx = {
        "current_url": send_url,
        "post_url": urljoin(
            BASE_URL,
            form.get("action", "/form/letter/send.xhtml")
        ),
        "form_id": form_id,
        "viewstate": get_viewstate_from_form(form),
        "district_name": district.get("name"),
        "court_name": court.get("name"),
        "text_name": text_field.get("name"),
        "request_hid_name": request_hid.get("name") if request_hid else "",
        "request_hid_value": request_hid.get("value", "") if request_hid else "",
        "next_source": next_source,
        "request_uploader": req_uploader_id,
        "attach_uploader": attach_uploader_id,
        "district_options": [
            (o.get("value", ""), normalize_text(o.get_text(" ", strip=True)))
            for o in district.find_all("option")
            if o.get("value", "")
        ],
        "court_options": [
            (o.get("value", ""), normalize_text(o.get_text(" ", strip=True)))
            for o in court.find_all("option")
            if o.get("value", "")
        ],
        "district_value": district.get("value", "") or "",
        "court_value": court.get("value", "") or "",
        "text_value": text_field.get_text() or "",
    }

    print("  ✓ send.xhtml разобран динамически:")
    print(f"    form={ctx['form_id']}")
    print(f"    district={ctx['district_name']}")
    print(f"    court={ctx['court_name']}")
    print(f"    requestUploader={ctx['request_uploader']}")
    print(f"    attachUploader={ctx['attach_uploader']}")
    print(f"    next={ctx['next_source']}")

    return ctx

def current_form_fields(ctx):
    fields = {
        ctx["form_id"]: ctx["form_id"],
        ctx["district_name"]: ctx.get("district_value", ""),
        ctx["court_name"]: ctx.get("court_value", ""),
        ctx["text_name"]: ctx.get("text_value", ""),
        "javax.faces.ViewState": ctx["viewstate"],
    }

    if ctx.get("request_hid_name"):
        fields[ctx["request_hid_name"]] = ctx.get(
            "request_hid_value", ""
        )

    return fields


def ajax_change_select(session, ctx, field_name, value):
    """
    JSF/RichFaces change для selectOneMenu.

    ВАЖНО:
    selected field должен содержать РЕАЛЬНОЕ значение:
        district-field = 8
        court-field    = 154

    Раньше make_richfaces_payload() добавлял source=source и тем самым
    перезаписывал 8/154 строкой имени компонента. Из-за этого JSF не менял
    регион/суд на сервере.
    """
    value = str(value)

    if field_name == ctx["district_name"]:
        ctx["district_value"] = value
        ctx["court_value"] = ""
    elif field_name == ctx["court_name"]:
        ctx["court_value"] = value

    fields = current_form_fields(ctx)

    payload = make_richfaces_payload(
        form_id=ctx["form_id"],
        viewstate=ctx["viewstate"],
        source=field_name,
        extra_fields=fields,
        partial_event="change",
        behavior_event="change"
    )

    # КРИТИЧЕСКИЙ FIX:
    # Для select source-компонент одновременно является полем формы.
    # Его значение должно быть выбранным ID, а не именем компонента.
    payload[field_name] = value
    payload["org.richfaces.ajax.component"] = field_name

    print(
        f"  JSF change: {field_name.split(':')[-1]}={value}"
    )

    r = http_post_ajax(
        session,
        ctx["post_url"],
        data=payload,
        headers=AJAX_HEADERS,
        retry_safe=True,
        description=f"JSF change {field_name}"
    )
    r.raise_for_status()

    # Диагностика JSF error
    if "<error>" in (r.text or "").lower():
        raise RuntimeError(
            f"JSF вернул <error> при выборе значения {value}:\n"
            + (r.text or "")[:2000]
        )

    ctx["viewstate"] = extract_viewstate_from_partial(
        r.text, ctx["viewstate"]
    )

    return r.text


def parse_court_options_from_partial(partial_text, court_field_name):
    update = parse_partial_update(partial_text, court_field_name)

    # Иногда сервер обновляет контейнер court, а не сам select.
    if not update:
        update = parse_partial_update(partial_text, "court")

    if not update:
        return []

    sp = soup(update)
    sel = sp.find("select")

    if not sel:
        return []

    return [
        (o.get("value", ""), normalize_text(o.get_text(" ", strip=True)))
        for o in sel.find_all("option")
        if o.get("value", "")
    ]


COURT_CACHE = {}

REGION_ALIASES = [
    ("город астана", "город Астана"), ("г. астана", "город Астана"),
    ("город алматы", "город Алматы"), ("г. алматы", "город Алматы"),
    ("город шымкент", "город Шымкент"), ("г. шымкент", "город Шымкент"),
    ("акмолин", "Акмолинская область"), ("актюбин", "Актюбинская область"),
    ("алматинск", "Алматинская область"), ("атырауск", "Атырауская область"),
    ("восточно-казахстан", "Восточно-Казахстанская область"),
    ("восточноказахстан", "Восточно-Казахстанская область"),
    ("жамбыл", "Жамбылская область"), ("западно-казахстан", "Западно-Казахстанская область"),
    ("западноказахстан", "Западно-Казахстанская область"), ("караганд", "Карагандинская область"),
    ("костанай", "Костанайская область"), ("кызылордин", "Кызылординская область"),
    ("мангиста", "Мангистауская область"), ("павлодар", "Павлодарская область"),
    ("северо-казахстан", "Северо-Казахстанская область"),
    ("североказахстан", "Северо-Казахстанская область"), ("туркестан", "Туркестанская область"),
    ("ұлытау", "Область Ұлытау"), ("улытау", "Область Ұлытау"),
    ("абай", "Область Абай"), ("жетісу", "Область Жетісу"), ("жетысу", "Область Жетісу"),
    ("военный суд", "Военный суд Республики Казахстан"),
]

def infer_region_caption_from_court(court_name):
    norm = normalize_match(court_name)
    for needle, caption in REGION_ALIASES:
        if normalize_match(needle) in norm:
            if caption == "город Алматы" and "алматинск" in norm:
                continue
            return caption
    return ""

def find_district_id(ctx, district_caption):
    target = normalize_match(district_caption)
    for district_id, caption in ctx["district_options"]:
        if normalize_match(caption) == target:
            return district_id, caption
    for district_id, caption in ctx["district_options"]:
        cand = normalize_match(caption)
        if target and (target in cand or cand in target):
            return district_id, caption
    return None, None



def strip_court_service_parts(name):
    """
    Для сравнения CRM F157 с полным справочником СК.
    Убираем только служебные детали, которые не меняют идентичность суда.
    """
    s = normalize_text(name)

    # Удаляем служебный хвост в скобках целиком.
    # Пример: "(Общая юрисдикция)" / "(Гражданские дела)".
    s = re.sub(r"\s*\([^)]*\)\s*$", "", s).strip()

    return s


def court_match_key(name):
    s = strip_court_service_parts(name)
    s = normalize_match(s)

    # Нормализуем отдельные символы/варианты написания.
    s = s.replace("№ ", "№")
    s = s.replace("г ", "город ")
    s = re.sub(r"\s+", " ", s).strip()

    return s


def court_word_score(target, candidate):
    """
    Используется только для диагностического сообщения.
    Автовыбор по нечёткому совпадению НЕ делаем.
    """
    t = set(court_match_key(target).split())
    c = set(court_match_key(candidate).split())
    if not t or not c:
        return 0
    return int(round(len(t & c) / len(t | c) * 100))


def resolve_court_from_directory(target_court):
    """
    F157 -> точные RegionID/CourtID из встроенного справочника СК.

    Правила безопасности:
    1) сначала полное точное совпадение;
    2) затем точное совпадение после удаления служебного хвоста в скобках;
    3) автоматически выбираем ТОЛЬКО если совпадение однозначное;
    4) fuzzy matching используется лишь для подсказки в ошибке.
    """
    original = normalize_text(target_court)
    full_key = normalize_match(original)
    short_key = court_match_key(original)

    # 1. Полное точное совпадение.
    exact_full = [
        x for x in SK_COURTS_FLAT
        if normalize_match(x["court_name"]) == full_key
    ]

    if len(exact_full) == 1:
        result = dict(exact_full[0])
        result["match_type"] = "EXACT_FULL"
        return result

    if len(exact_full) > 1:
        raise RuntimeError(
            f"В справочнике СК найдено несколько одинаковых полных названий суда: "
            f"{original}"
        )

    # 2. Совпадение после удаления хвоста "(...)"
    exact_short = [
        x for x in SK_COURTS_FLAT
        if court_match_key(x["court_name"]) == short_key
    ]

    if len(exact_short) == 1:
        result = dict(exact_short[0])
        result["match_type"] = "EXACT_NORMALIZED"
        return result

    if len(exact_short) > 1:
        variants = "; ".join(
            f'{x["region_name"]} / {x["court_name"]}'
            for x in exact_short[:10]
        )
        raise RuntimeError(
            f"После нормализации суд '{original}' совпал с несколькими судами СК. "
            f"Автовыбор запрещён. Варианты: {variants}"
        )

    # 3. Ничего однозначного не найдено.
    scored = sorted(
        [
            (court_word_score(original, x["court_name"]), x)
            for x in SK_COURTS_FLAT
        ],
        key=lambda z: z[0],
        reverse=True
    )

    hints = "; ".join(
        f'{score}%: {x["region_name"]} / {x["court_name"]}'
        for score, x in scored[:5]
    )

    raise RuntimeError(
        f"Суд из F157 не найден однозначно во встроенном справочнике СК: "
        f"'{original}'. Лучшие варианты: {hints}"
    )


def select_region_and_court_by_ids(session, ctx, mapping):
    """
    Устанавливаем RegionID и CourtID в серверной JSF-модели.
    """
    region_id = str(mapping["region_id"])
    court_id = str(mapping["court_id"])

    print(
        f'→ СК IDs: RegionID={region_id} '
        f'({mapping["region_name"]}), '
        f'CourtID={court_id} ({mapping["court_name"]})'
    )

    # 1. Регион
    region_response = ajax_change_select(
        session,
        ctx,
        ctx["district_name"],
        region_id
    )
    print("  ✓ RegionID принят JSF")

    # После change региона сервер формирует список допустимых судов.
    # Теперь выбираем конкретный суд.
    court_response = ajax_change_select(
        session,
        ctx,
        ctx["court_name"],
        court_id
    )
    print("  ✓ CourtID принят JSF")

    ctx["district_value"] = region_id
    ctx["court_value"] = court_id

    return region_response, court_response


def select_region_and_court(session, ctx, mapping):
    # Backward-compatible alias.
    return select_region_and_court_by_ids(session, ctx, mapping)


def _refresh_send_context_after_upload_error(session, ctx):
    """Безопасно перечитывает текущую send.xhtml после сетевого обрыва."""
    r = http_get(session, ctx["current_url"])
    r.raise_for_status()
    fresh = parse_send_context(ctx["current_url"], r.text)

    # Сохраняем уже выбранные/введённые значения, которые GET может не показать.
    fresh["district_value"] = ctx.get("district_value", "")
    fresh["court_value"] = ctx.get("court_value", "")
    fresh["text_value"] = ctx.get("text_value", "")
    if ctx.get("request_hid_value"):
        fresh["request_hid_value"] = ctx.get("request_hid_value", "")
    ctx.clear()
    ctx.update(fresh)
    return r.text


def _filename_visible_in_send(html_text, file_name):
    """Проверка, успел ли СК зарегистрировать файл до обрыва ответа."""
    if not html_text:
        return False
    decoded = html.unescape(html_text)
    name = Path(file_name).name
    return name in decoded or name.replace(" ", "&nbsp;") in html_text


def richfaces_upload(session, ctx, uploader_id, file_path):
    """
    RichFaces 4.5 FileUpload с безопасным восстановлением после SSL EOF/
    RemoteDisconnected.

    ВАЖНО: слепого повторного POST нет. После обрыва сначала перечитываем
    send.xhtml. Если имя файла уже видно серверу — считаем upload принятым.
    Если файла нет — обновляем ViewState/ID компонентов и делаем только ОДИН
    повтор. Это уменьшает риск дублирования вложений.
    """
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(file_path)

    is_request = uploader_id == ctx.get("request_uploader")

    def do_one_upload(actual_uploader_id):
        params = {
            "rf_fu_uid": str(time.time()).replace(".", ""),
            "javax.faces.partial.ajax": "true",
            "javax.faces.source": actual_uploader_id,
            "javax.faces.partial.execute": actual_uploader_id,
            "org.richfaces.ajax.component": actual_uploader_id,
            "javax.faces.ViewState": ctx["viewstate"],
        }
        mime = mimetypes.guess_type(file_path.name)[0] or "application/octet-stream"
        data = {
            ctx["form_id"]: ctx["form_id"],
            "javax.faces.ViewState": ctx["viewstate"],
        }
        headers = {
            "Faces-Request": "partial/ajax",
            "Referer": ctx["current_url"],
        }
        with file_path.open("rb") as f:
            files = {actual_uploader_id: (file_path.name, f, mime)}
            return session.post(
                ctx["post_url"], params=params, data=data, files=files,
                headers=headers, timeout=HTTP_TIMEOUT
            )

    try:
        r = do_one_upload(uploader_id)
    except RETRYABLE_EXCEPTIONS as first_error:
        print(f"  ! Соединение оборвалось при upload: {type(first_error).__name__}")
        print("  → Проверяю состояние send.xhtml перед возможным повтором...")

        try:
            fresh_html = _refresh_send_context_after_upload_error(session, ctx)
        except Exception as check_error:
            raise RuntimeError(
                f"Upload '{file_path.name}' оборвался, а безопасно проверить "
                f"состояние формы не удалось. Повтор НЕ выполнен. "
                f"Upload error: {first_error}; check error: {check_error}"
            ) from first_error

        if _filename_visible_in_send(fresh_html, file_path.name):
            print("  ✓ Файл уже виден в форме после обрыва — повтор POST не нужен")
            return "RECOVERED_ALREADY_UPLOADED"

        # После свежего GET IDs могли измениться.
        retry_uploader = ctx["request_uploader"] if is_request else ctx["attach_uploader"]
        print("  → Файл в форме не найден. Один безопасный повтор upload с новым ViewState...")
        try:
            r = do_one_upload(retry_uploader)
        except RETRYABLE_EXCEPTIONS as second_error:
            raise RuntimeError(
                f"Повторная загрузка файла '{file_path.name}' также оборвалась. "
                f"Дальнейшие повторы отключены. Ошибка: {second_error}"
            ) from second_error

    r.raise_for_status()
    if "<partial-response" not in r.text:
        raise RuntimeError(
            f"Неожиданный ответ загрузки {file_path.name}:\n" + r.text[:1200]
        )

    ctx["viewstate"] = extract_viewstate_from_partial(r.text, ctx["viewstate"])
    return r.text


def upload_complete(session, ctx, uploader_id):
    fields = current_form_fields(ctx)

    payload = make_richfaces_payload(
        form_id=ctx["form_id"],
        viewstate=ctx["viewstate"],
        source=uploader_id,
        extra_fields=fields,
        partial_event="onuploadcomplete",
        behavior_event="uploadcomplete"
    )

    payload["org.richfaces.ajax.component"] = uploader_id

    r = http_post_ajax(
        session, ctx["post_url"], data=payload, headers=AJAX_HEADERS,
        retry_safe=False, description="uploadcomplete"
    )
    r.raise_for_status()

    ctx["viewstate"] = extract_viewstate_from_partial(
        r.text, ctx["viewstate"]
    )

    # После загрузки основного письма сервер отдаёт requestScanHid
    panel = parse_partial_update(r.text, "requestScanPanel")

    if panel:
        sp = soup(panel)
        hidden = sp.find(
            "input",
            attrs={"name": re.compile(r"requestScanHid$")}
        )
        if hidden:
            ctx["request_hid_name"] = hidden.get("name")
            ctx["request_hid_value"] = hidden.get("value", "")

    return r.text


def upload_main_pdf(session, ctx, pdf_path):
    richfaces_upload(
        session,
        ctx,
        ctx["request_uploader"],
        pdf_path
    )
    upload_complete(
        session,
        ctx,
        ctx["request_uploader"]
    )

    if not ctx.get("request_hid_value"):
        # В перехваченном трафике на финальном POST это поле содержит имя PDF.
        # Если сервер не вернул hidden value, подставляем имя файла.
        ctx["request_hid_value"] = Path(pdf_path).name


def upload_attachment(session, ctx, attachment_path):
    richfaces_upload(
        session,
        ctx,
        ctx["attach_uploader"],
        attachment_path
    )
    upload_complete(
        session,
        ctx,
        ctx["attach_uploader"]
    )


def extract_jsf_messages(partial_text):
    """
    Вынимает непустые сообщения из JSF partial-response.
    """
    messages = []
    try:
        sp = soup(partial_text)
        for el in sp.find_all(["span", "div", "li"]):
            classes = " ".join(el.get("class", []))
            if "rf-msg" in classes or "message" in classes.lower():
                txt = normalize_text(el.get_text(" ", strip=True))
                if txt and txt not in messages:
                    messages.append(txt)
    except Exception:
        pass

    for pat in [
        r'(?i)(обязатель[^<]{0,200})',
        r'(?i)(необходимо[^<]{0,200})',
        r'(?i)(некоррект[^<]{0,200})',
        r'(?i)(ошибк[^<]{0,200})',
        r'(?i)(выберите[^<]{0,200})',
    ]:
        for mm in re.finditer(pat, partial_text or ""):
            txt = normalize_text(re.sub(r"<[^>]+>", " ", mm.group(1)))
            if txt and txt not in messages:
                messages.append(txt)

    return messages[:20]


def extract_letter_number_from_sign_html(sign_html):
    """
    На sign.xhtml номер заявления находится в xmlToSign0:
        <f1>196368L15657387</f1>
    """
    if not sign_html:
        return ""

    decoded = html.unescape(sign_html)

    patterns = [
        r"<f1>\s*([^<]+?)\s*</f1>",
        r"&lt;f1&gt;\s*([^&<]+?)\s*&lt;/f1&gt;",
    ]

    for pat in patterns:
        mm = re.search(pat, decoded, flags=re.I | re.S)
        if mm:
            return normalize_text(mm.group(1))

    return ""


def extract_visible_wait_numbers(html_text):
    """
    Берём номера только из реально отображаемых карточек wait.xhtml.

    Не ищем номер по всему HTML, потому что он теоретически может
    присутствовать в hidden/JS/служебных данных.
    """
    sp = soup(html_text)

    numbers = []

    # Карточки списка имеют class case-item-container.
    for card in sp.select(".case-item-container"):
        h = card.find(["h1", "h2", "h3", "h4"])
        if not h:
            continue

        txt = normalize_text(h.get_text(" ", strip=True))
        txt = txt.replace("№", "").strip()

        if txt and txt not in numbers:
            numbers.append(txt)

    return numbers


def get_wait_main_form(html_text):
    sp = soup(html_text)

    # Ищем форму, внутри которой находятся пагинационные ссылки thisPage.
    for form in sp.find_all("form"):
        if form.find(
            attrs={
                "onclick": re.compile(r'thisPage')
            }
        ):
            return form

    # fallback по action
    for form in sp.find_all("form"):
        action = form.get("action", "")
        if "wait.xhtml" in action and form.find("input", attrs={"name": "javax.faces.ViewState"}):
            # Не берём поисковую форму, если в ней нет pagination.
            if form.find("a", onclick=re.compile(r'thisPage')):
                return form

    return None


def extract_jsf_form_fields(form):
    """
    Собираем значения служебной JSF-формы.
    На wait.xhtml пагинационные RichFaces-компоненты находятся вне <form>,
    но браузер отправляет первую JSF-форму страницы с её ViewState.
    """
    fields = {}

    form_id = form.get("id") or form.get("name")
    if form_id:
        fields[form_id] = form_id

    for inp in form.find_all("input"):
        name = inp.get("name")
        if not name:
            continue

        typ = (inp.get("type") or "").lower()

        if typ in {"submit", "button", "file"}:
            continue

        if typ in {"checkbox", "radio"} and not inp.has_attr("checked"):
            continue

        fields[name] = inp.get("value", "")

    for sel in form.find_all("select"):
        name = sel.get("name")
        if not name:
            continue

        selected = sel.find("option", selected=True)
        if selected is None:
            selected = sel.find("option")

        fields[name] = selected.get("value", "") if selected else ""

    for ta in form.find_all("textarea"):
        name = ta.get("name")
        if name:
            fields[name] = ta.get_text() or ""

    return fields


def get_wait_ajax_form(html_text):
    """
    Пагинация wait.xhtml не вложена в свою форму.

    По реальному перехваченному трафику RichFaces использует служебную
    JSF-форму вверху страницы (поиск по сайту), содержащую актуальный
    javax.faces.ViewState.
    """
    sp = soup(html_text)

    # Сначала форма, action которой явно wait.xhtml.
    candidates = []
    for form in sp.find_all("form"):
        if not form.find("input", attrs={"name": "javax.faces.ViewState"}):
            continue

        action = form.get("action", "")
        if "wait.xhtml" in action:
            candidates.append(form)

    if candidates:
        return candidates[0]

    # Fallback: первая форма с ViewState.
    form = sp.find(
        "form",
        lambda tag: False
    )

    for f in sp.find_all("form"):
        if f.find("input", attrs={"name": "javax.faces.ViewState"}):
            return f

    return None


def extract_wait_html_fragments(partial_text):
    """
    RichFaces partial-response хранит обновлённые куски страницы внутри CDATA.
    Возвращаем их как единый HTML для поиска новой пагинации/карточек.
    """
    chunks = []

    for mm in re.finditer(
        r'<update\s+id=["\']([^"\']+)["\'][^>]*>(.*?)</update>',
        partial_text or "",
        flags=re.I | re.S
    ):
        upd_id = html.unescape(mm.group(1))
        body = mm.group(2)

        cm = re.search(r'<!\[CDATA\[(.*?)\]\]>', body, flags=re.S)
        content = cm.group(1) if cm else html.unescape(body)

        # Нас интересуют список и пагинация, но добавляем всё HTML-похожее.
        if (
            "filter" in upd_id.lower()
            or "list" in upd_id.lower()
            or "page" in upd_id.lower()
            or "case-item-container" in content
            or "thisPage" in content
        ):
            chunks.append(content)

    return "\n".join(chunks)


def find_wait_page_link(html_text, page_number):
    """
    Ищет RichFaces source ссылки страницы N.
    Родительская form для самой ссылки НЕ требуется.
    """
    sp = soup(html_text)
    target = str(page_number)

    for a in sp.find_all("a"):
        onclick = html.unescape(a.get("onclick", "") or "")
        text = normalize_text(a.get_text(" ", strip=True))

        if text != target:
            continue

        if (
            f'"thisPage":"{target}"' in onclick
            or f"'thisPage':'{target}'" in onclick
            or re.search(
                r'thisPage[^0-9]{0,20}' + re.escape(target),
                onclick,
                flags=re.I
            )
        ):
            source_id = a.get("id") or a.get("name")

            if source_id:
                return {
                    "source_id": normalize_js_identifier(source_id),
                    "page": target,
                }

    return None


def wait_ajax_page(session, base_html, page_number, link_html=None,
                   form_fields=None, viewstate=None):
    """
    Повторяет реальный RichFaces.ajax пагинации wait.xhtml.

    Важная особенность СК:
    pagination source расположен вне <form>, поэтому используем
    служебную JSF-форму страницы + её ViewState.
    """
    html_for_link = link_html if link_html is not None else base_html
    meta = find_wait_page_link(html_for_link, page_number)

    if meta is None:
        return None, base_html, form_fields, viewstate

    source = meta["source_id"]
    target = meta["page"]

    if form_fields is None or viewstate is None:
        form = get_wait_ajax_form(base_html)

        if form is None:
            raise RuntimeError(
                "На wait.xhtml не найдена служебная JSF-форма с ViewState."
            )

        form_fields = extract_jsf_form_fields(form)
        form_id = form.get("id") or form.get("name")
        viewstate = get_viewstate_from_form(form)
    else:
        form_id = next(
            (
                k for k, v in form_fields.items()
                if k == v and ":" in k
            ),
            None
        )

        if not form_id:
            # fallback: форма из исходного HTML
            form = get_wait_ajax_form(base_html)
            if form is None:
                raise RuntimeError("Не удалось восстановить JSF form_id wait.xhtml")
            form_id = form.get("id") or form.get("name")

    payload = dict(form_fields)

    payload.update({
        form_id: form_id,
        "javax.faces.ViewState": viewstate,
        "javax.faces.source": source,
        "javax.faces.partial.execute": f"{source} @component",
        "javax.faces.partial.render": "@component",
        "thisPage": target,
        "org.richfaces.ajax.component": source,
        source: source,
        "rfExt": "null",
        "AJAX:EVENTS_COUNT": "1",
        "javax.faces.partial.ajax": "true",
    })

    print(
        f"  JSF wait pagination → страница {target} "
        f"(form={form_id}, source={source})"
    )

    r = http_post_ajax(
        session,
        LETTER_WAIT_URL,
        data=payload,
        headers={
            **AJAX_HEADERS,
            "Referer": LETTER_WAIT_URL,
        },
        retry_safe=True,
        description=f"wait page {target}"
    )
    r.raise_for_status()

    if "<error>" in (r.text or "").lower():
        raise RuntimeError(
            f"JSF error при переключении wait.xhtml на страницу {target}:\n"
            + (r.text or "")[:1500]
        )

    new_viewstate = extract_viewstate_from_partial(
        r.text,
        viewstate
    )

    form_fields["javax.faces.ViewState"] = new_viewstate

    fragment = extract_wait_html_fragments(r.text)

    return r.text, fragment, form_fields, new_viewstate


def force_wait_refresh_2_to_1(session):
    """
    Реально повторяем пользовательское действие:
      wait page 1 -> AJAX page 2 -> AJAX page 1.

    После этого делаем свежий GET первой страницы и возвращаем её HTML.
    """
    nonce = int(time.time() * 1000)

    r0 = http_get(
        session,
        LETTER_WAIT_URL,
        params={"_refresh": nonce},
        headers={
            "Cache-Control": "no-cache, no-store, max-age=0",
            "Pragma": "no-cache",
        },
        timeout=60,
        retry=True
    )
    r0.raise_for_status()

    html0 = r0.text or ""

    visible_before = extract_visible_wait_numbers(html0)
    if visible_before:
        print(
            "  До refresh первая карточка: "
            f"№{visible_before[0]}"
        )

    page2 = find_wait_page_link(html0, 2)

    if page2 is None:
        print("  Страницы 2 нет — принудительная пагинация не требуется")
        return html0

    # Инициализируем JSF form один раз.
    form = get_wait_ajax_form(html0)
    if form is None:
        raise RuntimeError(
            "На wait.xhtml не найдена служебная JSF-форма с ViewState."
        )

    fields = extract_jsf_form_fields(form)
    viewstate = get_viewstate_from_form(form)

    # 1 -> 2
    _, fragment2, fields, viewstate = wait_ajax_page(
        session,
        base_html=html0,
        page_number=2,
        link_html=html0,
        form_fields=fields,
        viewstate=viewstate
    )

    if not fragment2:
        raise RuntimeError(
            "СК выполнил переход на страницу 2, "
            "но partial-response не содержит обновлённую пагинацию."
        )

    # На обновлённой странице 2 должна быть ссылка обратно на 1.
    page1 = find_wait_page_link(fragment2, 1)

    if page1 is None:
        # Сохраняем для диагностики.
        debug = WORK_DIR / "LAST_WAIT_PAGE2_RESPONSE.html"
        try:
            debug.write_text(fragment2, encoding="utf-8")
        except Exception:
            pass

        raise RuntimeError(
            "После AJAX-перехода на страницу 2 не найдена ссылка страницы 1. "
            f"Фрагмент сохранён: {debug}"
        )

    # 2 -> 1
    _, fragment1, fields, viewstate = wait_ajax_page(
        session,
        base_html=html0,
        page_number=1,
        link_html=fragment2,
        form_fields=fields,
        viewstate=viewstate
    )

    print("  ✓ JSF refresh wait.xhtml выполнен: 1 → 2 → 1")

    # В fragment1 уже должен быть обновлённый список первой страницы.
    # Сначала используем его — это точнее, чем обычный GET.
    if fragment1:
        nums = extract_visible_wait_numbers(fragment1)
        if nums:
            print(f"  После refresh первая карточка: №{nums[0]}")
            return fragment1

    # Fallback: свежий GET после server-side refresh.
    nonce2 = int(time.time() * 1000)

    rf = http_get(
        session,
        LETTER_WAIT_URL,
        params={"_refresh": nonce2},
        headers={
            "Cache-Control": "no-cache, no-store, max-age=0",
            "Pragma": "no-cache",
        },
        timeout=60,
        retry=True
    )
    rf.raise_for_status()

    return rf.text or ""


def verify_in_wait(session, letter_number, retries=6, delay=2):
    """
    Финальная проверка v10:
    - НЕ переключает JSF-страницы 1 -> 2 -> 1;
    - на каждой попытке делает новый GET wait.xhtml с cache-busting;
    - проверяет номер только среди карточек .case-item-container.

    Диагностика показала, что сервер уже отдаёт новый номер в свежем GET,
    даже если старая открытая вкладка браузера визуально ещё не обновилась.
    """
    if not letter_number:
        raise RuntimeError(
            "Не удалось определить номер заявления на sign.xhtml, "
            "поэтому проверить wait.xhtml невозможно."
        )

    print(f"→ Проверяю «Ожидает отправки» свежим GET: №{letter_number}")

    last_numbers = []

    for attempt in range(1, retries + 1):
        nonce = int(time.time() * 1000)
        r = http_get(
            session,
            LETTER_WAIT_URL,
            params={"_refresh": nonce},
            headers={
                "Cache-Control": "no-cache, no-store, max-age=0",
                "Pragma": "no-cache",
            },
            timeout=60,
            retry=True
        )
        r.raise_for_status()

        body = r.text or ""
        visible_numbers = extract_visible_wait_numbers(body)
        last_numbers = visible_numbers

        if visible_numbers:
            print(f"  Свежий GET: первая карточка №{visible_numbers[0]}")
        else:
            print("  Свежий GET: видимые карточки с номерами не найдены")

        if letter_number in visible_numbers:
            print(
                f"  ✓ №{letter_number} найден среди ВИДИМЫХ карточек "
                f"свежего wait.xhtml (попытка {attempt}/{retries})"
            )
            return True

        preview = ", ".join(visible_numbers[:5]) or "<нет номеров>"
        print(
            f"  Не найден, попытка {attempt}/{retries}. "
            f"Первые номера: {preview}"
        )

        if attempt < retries:
            print(f"  Повтор свежего GET через {delay} сек...")
            time.sleep(delay)

    raise RuntimeError(
        f"Заявление №{letter_number} получило sign.xhtml, "
        f"но не найдено среди видимых карточек свежего GET wait.xhtml "
        f"после {retries} попыток. "
        f"Последние видимые номера: {last_numbers[:10]}"
    )

def go_next_without_signing(session, ctx):
    fields = current_form_fields(ctx)

    print(
        "  Финальные поля: "
        f"RegionID={ctx.get('district_value')}, "
        f"CourtID={ctx.get('court_value')}, "
        f"PDF={ctx.get('request_hid_value')!r}, "
        f"TextLen={len(ctx.get('text_value',''))}"
    )

    payload = make_richfaces_payload(
        form_id=ctx["form_id"],
        viewstate=ctx["viewstate"],
        source=ctx["next_source"],
        extra_fields=fields,
        partial_event="click"
    )

    r = http_post_ajax(
        session,
        ctx["post_url"],
        data=payload,
        headers=AJAX_HEADERS,
        retry_safe=False,
        description="Кнопка Далее"
    )
    r.raise_for_status()

    redirect = extract_redirect(r.text)

    if not redirect:
        debug_path = WORK_DIR / "LAST_NEXT_RESPONSE.xml"
        try:
            debug_path.write_text(r.text or "", encoding="utf-8")
        except Exception:
            pass

        messages = extract_jsf_messages(r.text)
        msg_text = "\n".join(f"  - {x}" for x in messages)

        extra = ""
        if msg_text:
            extra = "\nСообщения JSF:\n" + msg_text

        raise RuntimeError(
            "После кнопки 'Далее' не получен redirect на sign.xhtml."
            + extra
            + f"\nПолный ответ сохранён: {debug_path}"
        )

    sign_url = urljoin(BASE_URL, redirect)

    if "/form/letter/sign.xhtml" not in sign_url:
        raise RuntimeError(
            f"Ожидался sign.xhtml, получено: {sign_url}"
        )

    # GET sign.xhtml — только чтение страницы подписи.
    rs = http_get(session, sign_url)
    rs.raise_for_status()

    letter_number = extract_letter_number_from_sign_html(rs.text)

    if letter_number:
        print(f"  ✓ СК присвоил номер: №{letter_number}")
    else:
        print("  ! Номер заявления на sign.xhtml не распознан")

    return sign_url, letter_number


# ============================================================
# 6. ЛОГ
# ============================================================

RESULT_COLUMNS = [
    "LoanID",
    "EID",
    "IIN",
    "FIO",
    "Sud",
    "DebtRestWithoutStateDuty",
    "PDF",
    "District",
    "CourtInSK",
    "LetterNumber",
    "SignURL",
    "Status",
    "Error",
]


def save_results(rows):
    ensure_dirs()
    df_new = pd.DataFrame(rows, columns=RESULT_COLUMNS)

    if LOG_FILE.exists():
        try:
            df_old = pd.read_excel(LOG_FILE)
            df_all = pd.concat([df_old, df_new], ignore_index=True)
        except Exception:
            df_all = df_new
    else:
        df_all = df_new

    df_all.to_excel(LOG_FILE, index=False)
    return df_all


# ============================================================
# 7. ОСНОВНОЙ ЗАПУСК
# ============================================================



def run():
    ensure_dirs()

    df = load_source()

    if df.empty:
        print("Нет сделок для обработки.")
        return pd.DataFrame()

    session = login_and_make_session()
    results = []

    try:
        for idx, row in df.iterrows():
            print("\n" + "=" * 100)
            print(
                f"[{idx + 1}/{len(df)}] "
                f"EID={row['EID']} | ИИН={row['IIN']}"
            )
            print(f"Суд: {row['Sud']}")

            item = {
                "LoanID": row["LoanID"],
                "EID": row["EID"],
                "IIN": row["IIN"],
                "FIO": row["FIO"],
                "Sud": row["Sud"],
                "DebtRestWithoutStateDuty": row["DebtRestWithoutStateDuty"],
                "PDF": "",
                "District": "",
                "CourtInSK": "",
                "LetterNumber": "",
                "SignURL": "",
                "Status": "",
                "Error": "",
            }

            try:
                # 1. DOCX + PDF
                docx_path, pdf_path, letter_text = make_documents(row)
                item["PDF"] = str(pdf_path)
                print(f"✓ PDF: {pdf_path.name}")

                # 2. Сопоставить F157 с зашитым справочником СК.
                print("→ Сопоставляю суд F157 со встроенным справочником СК")
                mapping = resolve_court_from_directory(row["Sud"])

                item["District"] = mapping["region_name"]
                item["CourtInSK"] = mapping["court_name"]

                print(
                    f'✓ Суд найден локально [{mapping["match_type"]}]: '
                    f'{mapping["region_name"]} → {mapping["court_name"]}'
                )
                print(
                    f'  RegionID={mapping["region_id"]}, '
                    f'CourtID={mapping["court_id"]}'
                )

                # 3. Создать ОДИН чистый черновик для самой подачи.
                print("→ Создание черновика для подачи")
                send_url, send_html = open_new_letter(session)

                # ВРЕМЕННАЯ ДИАГНОСТИКА: всегда сохраняем фактически полученный send.xhtml
                # до разбора формы. Если структура страницы изменилась, файл останется
                # на диске даже при последующей ошибке parse_send_context().
                debug_send_path = WORK_DIR / "DEBUG_send.xhtml.html"
                debug_send_path.write_text(send_html, encoding="utf-8")
                print(f"  DEBUG: send.xhtml сохранён: {debug_send_path}")
                print(f"  DEBUG: HTML len={len(send_html)}")
                print(f"  DEBUG: district-field={'district-field' in send_html}")
                print(f"  DEBUG: selectRequestScanUploader={'selectRequestScanUploader' in send_html}")
                print(f"  DEBUG: selectFileUploader={'selectFileUploader' in send_html}")
                print(f"  DEBUG: Далее={'Далее' in send_html}")

                try:
                    ctx = parse_send_context(send_url, send_html)
                except Exception:
                    # Дополнительно сохраняем первые признаки того, что реально вернул сервер.
                    debug_meta_path = WORK_DIR / "DEBUG_send.xhtml_meta.txt"
                    sp_debug = soup(send_html)
                    debug_meta_path.write_text(
                        "URL: " + str(send_url) + "\n"
                        + "HTML_LEN: " + str(len(send_html)) + "\n"
                        + "TITLE: " + (normalize_text(sp_debug.title.get_text(" ", strip=True)) if sp_debug.title else "") + "\n"
                        + "district-field: " + str("district-field" in send_html) + "\n"
                        + "court-field: " + str("court-field" in send_html) + "\n"
                        + "text-field: " + str("text-field" in send_html) + "\n"
                        + "selectRequestScanUploader: " + str("selectRequestScanUploader" in send_html) + "\n"
                        + "selectFileUploader: " + str("selectFileUploader" in send_html) + "\n"
                        + "Далее: " + str("Далее" in send_html) + "\n",
                        encoding="utf-8"
                    )
                    print(f"  DEBUG: метаданные сохранены: {debug_meta_path}")
                    raise

                # 4. Регион + суд
                print("→ Устанавливаю регион и суд по готовым ID")
                select_region_and_court_by_ids(
                    session,
                    ctx,
                    mapping
                )
                print("✓ RegionID и CourtID отправлены в форму")

                # 5. Текст
                ctx["text_value"] = letter_text

                # 6. Основной PDF заявления
                print("→ Загружаю PDF заявления")
                upload_main_pdf(
                    session,
                    ctx,
                    pdf_path
                )
                print("✓ PDF заявления загружен")

                # 7. Доверенность НЕ прикладываем.
                # В Судебный кабинет загружается только PDF заявления.

                # 8. Далее -> sign.xhtml -> STOP
                print("→ Нажимаю «Далее» (без подписи)")
                sign_url, letter_number = go_next_without_signing(
                    session,
                    ctx
                )

                item["SignURL"] = sign_url
                item["LetterNumber"] = letter_number

                # В массовом режиме номер заявления + sign.xhtml
                # являются достаточным подтверждением создания черновика.
                # wait.xhtml НЕ проверяем здесь: его JSF-список может
                # возвращать устаревшее состояние внутри текущей HTTP-сессии.
                if not letter_number:
                    raise RuntimeError(
                        "Получен sign.xhtml, но номер заявления не распознан."
                    )

                item["Status"] = "SAVED_FOR_SIGNING"

                print("✓ Заявление создано")
                print("✓ СК присвоил номер")
                print("✓ Остановлено ДО подписания")
                print(f"  №{letter_number}")
                print(f"  {sign_url}")

            except Exception as e:
                item["Status"] = "ERROR"
                item["Error"] = repr(e)

                print(f"✗ ОШИБКА: {e}")

            results.append(item)
            save_results([item])

            time.sleep(PAUSE_BETWEEN_ROWS)

    finally:
        session.close()

    print("\n" + "=" * 100)
    print("ГОТОВО")
    print(f"Лог: {LOG_FILE}")
    print("=" * 100)

    return pd.DataFrame(results)


if __name__ == "__main__":
    run()
