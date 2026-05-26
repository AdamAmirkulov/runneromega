# -*- coding: utf-8 -*-
import os
import re
import time
import difflib
from pathlib import Path
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

from pypdf import PdfReader
from openpyxl import load_workbook
from docx import Document as DocxDocument


# ================= НАСТРОЙКИ =================
AISOIP_URL = "https://aisoip.adilet.gov.kz/cabinet/exec-productions"
STATEMENTS_URL = "https://aisoip.adilet.gov.kz/cabinet/statements"

LOGIN = "810813301334_230240016634"
PASSWORD = "Qazaq123456*"

IL_FOLDER = r"C:\Users\user\Desktop\Подача по АИС ОИП\ИЛ"
STATEMENTS_FOLDER = r"C:\Users\user\Desktop\Подача по АИС ОИП\Заявления"

DEALS_XLSX = r"C:\Users\user\Desktop\Подача по АИС ОИП\Список для подачи заявления ЧСИ через АИС ОИП_Omega.xlsx"
COURTS_XLSX = r"C:\Users\user\Desktop\Подача по АИС ОИП\Суды в АИС ОИП.xlsx"
CHSI_XLSX   = r"C:\Users\user\Desktop\Подача по АИС ОИП\ЧСИ в АИС ОИП.xlsx"

DEAL_CHSI_COL_LETTER = "H"

REGION_MATCH_THRESHOLD = 0.3
COURT_MATCH_THRESHOLD  = 0.3
CHSI_MATCH_THRESHOLD   = 0.25
CHSI_NAME_MATCH_THRESHOLD = 0.55

DEBUG_CHSI = True
DEBUG_CHSI_DIR = r"C:\Users\user\Desktop\Подача по АИС ОИП\debug_chsi"


_t0 = time.time()
def log(msg):
    dt = time.time() - _t0
    print(f"[{dt:7.2f}s] {msg}")


def chsi_dbg_save(driver, name: str):
    if not DEBUG_CHSI:
        return
    try:
        os.makedirs(DEBUG_CHSI_DIR, exist_ok=True)
    except Exception:
        return

    ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    base = os.path.join(DEBUG_CHSI_DIR, f"{ts}_{name}")

    try:
        driver.save_screenshot(base + ".png")
    except Exception as e:
        log(f"[CHSI-DBG] screenshot error: {type(e).__name__}: {e}")

    try:
        with open(base + ".html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
    except Exception as e:
        log(f"[CHSI-DBG] page_source error: {type(e).__name__}: {e}")


class AlreadyRegisteredException(Exception):
    pass


STOP_WORDS = {
    "суд", "суда", "суды", "районный", "районного", "городской",
    "городского", "города", "области", "обл", "обл.", "республики",
    "рк", "республика", "административный", "межрайонный",
    "специализированный", "специализированного", "район", "область",
    "в", "частные", "судебные", "исполнители", "частных",
    "сот", "соты", "соттары", "облысы", "облыс", "аудандық",
    "қалалық", "қаласы", "г", "г.", "город",
}

def stem_token(token: str) -> str:
    endings = [
        "ая", "яя", "ой", "ый", "ий", "ое", "ее", "ые", "ие",
        "ого", "его", "ому", "ему", "ыми", "ими", "ых", "их",
        "ую", "юю", "ой", "ей", "ом", "ем", "ам", "ям", "ах", "ях",
    ]
    for e in endings:
        if token.endswith(e) and len(token) > len(e) + 2:
            return token[:-len(e)]
    return token

def normalize_text(s: str) -> str:
    if not s:
        return ""
    s = str(s).lower()
    s = s.replace("ё", "е").replace("’", "'").replace("`", "'")
    s = re.sub(r"[^0-9a-zа-яәіңғүұқөһі\s]", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def tokens_without_stopwords(norm: str):
    if not norm:
        return set()
    result = set()
    for t in norm.split():
        if not t:
            continue
        if t in STOP_WORDS:
            continue
        result.add(stem_token(t))
    return result

def extract_12_digits_from_key(key: str):
    if not key:
        return None
    m = re.search(r"\b(\d{12})\b", str(key))
    return m.group(1) if m else None

def excel_col_letter_to_index(letter: str) -> int:
    letter = (letter or "").strip().upper()
    if not letter:
        raise ValueError("Пустая буква колонки Excel")
    n = 0
    for ch in letter:
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"Некорректная буква колонки: {letter}")
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


RU_MONTHS = {
    "января": 1, "февраля": 2, "марта": 3, "апреля": 4, "мая": 5, "июня": 6,
    "июля": 7, "августа": 8, "сентября": 9, "октября": 10, "ноября": 11, "декабря": 12
}

def parse_ru_date_words(s: str):
    if not s:
        return None
    m = re.search(r"\b(\d{1,2})\s+([а-яё]+)\s+(\d{4})\s*(?:года|г\.)?\b", s.lower())
    if not m:
        return None
    day = int(m.group(1))
    mon_word = m.group(2)
    year = int(m.group(3))
    mon = RU_MONTHS.get(mon_word)
    if not mon:
        return None
    try:
        dt = datetime(year, mon, day)
        return dt.strftime("%d.%m.%Y")
    except Exception:
        return None


def load_courts_and_regions(xlsx_path: str):
    regions = []
    courts = []

    if not os.path.exists(xlsx_path):
        log(f"[COURTS] Файл судов не найден: {xlsx_path}")
        return regions, courts

    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb.active

    current_region = None
    region_seen = {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        region_cell, court_cell = row[0], row[1]
        if court_cell is None:
            continue

        if region_cell:
            current_region = str(region_cell).strip()
        if not current_region:
            continue

        if current_region not in region_seen:
            norm_reg = normalize_text(current_region)
            tokens_core_reg = tokens_without_stopwords(norm_reg)
            reg_obj = {
                "region": current_region,
                "norm": norm_reg,
                "tokens_core": tokens_core_reg,
            }
            regions.append(reg_obj)
            region_seen[current_region] = reg_obj

        court_name = str(court_cell).strip()
        norm_court = normalize_text(court_name)
        tokens_core_court = tokens_without_stopwords(norm_court)
        courts.append({
            "court": court_name,
            "region": current_region,
            "norm": norm_court,
            "tokens_core": tokens_core_court,
        })

    wb.close()
    log(f"[COURTS] Регионов: {len(regions)}, судов: {len(courts)}")
    return regions, courts


REGIONS_LIST, COURTS_LIST = load_courts_and_regions(COURTS_XLSX)


def load_chsi_table(xlsx_path: str):
    rows = []
    if not os.path.exists(xlsx_path):
        log(f"[CHSI] Файл ЧСИ не найден: {xlsx_path}")
        return rows

    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb.active

    for row in ws.iter_rows(min_row=2, values_only=True):
        to_cell       = row[0] if len(row) > 0 else None
        district_cell = row[1] if len(row) > 1 else None
        chsi_cell     = row[2] if len(row) > 2 else None

        if not to_cell:
            continue

        to       = str(to_cell).strip()
        district = str(district_cell).strip() if district_cell else ""
        chsi     = str(chsi_cell).strip() if chsi_cell else ""

        rows.append({
            "to": to,
            "district": district,
            "chsi": chsi,
            "norm_to": normalize_text(to),
            "tokens_to": tokens_without_stopwords(normalize_text(to)),
            "norm_district": normalize_text(district) if district else "",
            "tokens_district": tokens_without_stopwords(normalize_text(district)) if district else set(),
            "norm_chsi": normalize_text(chsi),
            "tokens_chsi": tokens_without_stopwords(normalize_text(chsi)) if chsi else set(),
        })

    wb.close()
    log(f"[CHSI] Строк в таблице ЧСИ: {len(rows)}")
    return rows


CHSI_ROWS = load_chsi_table(CHSI_XLSX)


def load_deals_list(xlsx_path: str):
    deals = []
    if not os.path.exists(xlsx_path):
        log(f"[LIST] Файл списка сделок не найден: {xlsx_path}")
        return deals

    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb.active

    chsi_col_idx = excel_col_letter_to_index(DEAL_CHSI_COL_LETTER)

    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) < 9:
            continue

        key  = row[4]   # E
        addr = row[8]   # I
        deal_chsi = row[chsi_col_idx] if len(row) > chsi_col_idx else None

        if not key:
            continue

        deals.append({
            "key": str(key).strip(),
            "excel_address": str(addr).strip() if addr else "",
            "deal_chsi": str(deal_chsi).strip() if deal_chsi else "",
        })

    wb.close()
    log(f"[LIST] Кол-во записей в списке для подачи: {len(deals)}")
    return deals


DEALS = load_deals_list(DEALS_XLSX)


def read_pdf_text(pdf_path: str):
    reader = PdfReader(pdf_path)
    parts = []
    for page in reader.pages:
        try:
            t = page.extract_text() or ""
        except Exception:
            t = ""
        if t:
            parts.append(t)
    full_text = "\n".join(parts).replace("\xa0", " ")
    lines = [ln.strip() for ln in full_text.splitlines() if ln.strip()]
    joined = "\n".join(lines)
    return joined, lines

# ==========================
# ✅ ДОБАВЛЕНО: чтение ТОЛЬКО 1-й страницы PDF
# (Нужно, чтобы суд определялся по шапке 1-й страницы, а не по перемешанным строкам из других страниц)
# ==========================
def read_pdf_first_page_text(pdf_path: str):
    reader = PdfReader(pdf_path)
    if not reader.pages:
        return "", []
    try:
        t = reader.pages[0].extract_text() or ""
    except Exception:
        t = ""
    t = t.replace("\xa0", " ")
    lines = [ln.strip() for ln in t.splitlines() if ln.strip()]
    joined = "\n".join(lines)
    return joined, lines
# ==========================

def read_docx_text(docx_path: str):
    doc = DocxDocument(docx_path)
    parts = []
    for p in doc.paragraphs:
        txt = (p.text or "").strip()
        if txt:
            parts.append(txt)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                txt = (cell.text or "").strip()
                if txt:
                    parts.append(txt)
    joined = "\n".join(parts).replace("\xa0", " ")
    lines = [ln.strip() for ln in joined.splitlines() if ln.strip()]
    joined = "\n".join(lines)
    return joined, lines

def extract_iin_or_zsn_from_text(joined_text: str):
    if not joined_text:
        return None
    t = joined_text

    is_kz_exec_sheet = re.search(r"АТҚАРУ\s+ПАРАҒЫ", t, flags=re.IGNORECASE) is not None

    def find_in_block(block_title_regex: str, id_word_regex: str):
        m = re.search(block_title_regex, t, flags=re.IGNORECASE)
        if not m:
            return None
        start = m.end()
        chunk = t[start:start + 2500]
        m_id = re.search(rf"\b({id_word_regex})\b\s*[:\-]?\s*(\d{{12}})\b", chunk, flags=re.IGNORECASE)
        return m_id.group(2) if m_id else None

    if is_kz_exec_sheet:
        iin = find_in_block(r"Жауапкердің\s+толық\s+атауы\s+және\s+мекенжайы", r"ЖСН")
        if iin:
            return iin

    iin = find_in_block(r"Полное\s+наименование\s+ответчика\s+и\s+его\s+адрес", r"ИИН")
    if iin:
        return iin

    iin = find_in_block(r"Ответчик\s*[:\-]", r"ИИН")
    if iin:
        return iin

    m = re.search(r"\bИИН\b\s*[:\-]?\s*(\d{12})\b", t, flags=re.IGNORECASE)
    if m:
        return m.group(1)
    m = re.search(r"\bЖСН\b\s*[:\-]?\s*(\d{12})\b", t, flags=re.IGNORECASE)
    if m:
        return m.group(1)

    return None

def extract_debtor_fio_from_text(joined_text: str):
    if not joined_text:
        return None

    m = re.search(r"\bФИО\s*[:\-]\s*(.+)", joined_text, flags=re.IGNORECASE)
    if m:
        fio_line = m.group(1).splitlines()[0].strip()
        fio_line = re.sub(r"\s{2,}", " ", fio_line).strip(" ,.;")
        if fio_line:
            return fio_line

    m = re.search(r"\bТ\.?\s*А\.?\s*Ә\.?\s*[:\-]\s*(.+)", joined_text, flags=re.IGNORECASE)
    if m:
        fio_line = m.group(1).splitlines()[0].strip()
        fio_line = re.sub(r"\s{2,}", " ", fio_line).strip(" ,.;")
        if fio_line:
            return fio_line

    m = re.search(r"\bОтветчик\s*[:\-]\s*([^\n\r]+)", joined_text, flags=re.IGNORECASE)
    if m:
        fio_line = m.group(1).strip()
        fio_line = re.split(r",\s*ИИН\b|;\s*ИИН\b|\s+ИИН\b", fio_line, maxsplit=1, flags=re.IGNORECASE)[0].strip(" ,.;")
        fio_line = re.sub(r"\s{2,}", " ", fio_line).strip()
        if fio_line:
            return fio_line

    return None


def build_docs_index_from_text(folder: str):
    by_iin = {}
    all_docs = []

    if not os.path.exists(folder):
        log(f"[DOCS] Папка не найдена: {folder}")
        return by_iin, all_docs

    pdfs  = list(Path(folder).rglob("*.pdf"))
    docxs = list(Path(folder).rglob("*.docx"))
    log(f"[DOCS] Найдено PDF: {len(pdfs)}, DOCX: {len(docxs)} (включая подпапки)")

    def add_doc(path: str, kind: str, joined_text: str, fio: str, iin: str):
        doc = {
            "path": path,
            "kind": kind,
            "iin": iin,
            "fio_raw": fio,
            "fio_norm": normalize_text(fio) if fio else "",
            "text": joined_text,
        }
        all_docs.append(doc)
        if iin:
            by_iin.setdefault(iin, []).append(doc)

    for p in pdfs:
        path = str(p)
        try:
            joined, _lines = read_pdf_text(path)
        except Exception as e:
            log(f"[DOCS] Не смог прочитать PDF: {path} ({type(e).__name__}: {e})")
            continue
        iin = extract_iin_or_zsn_from_text(joined)
        fio = extract_debtor_fio_from_text(joined)
        add_doc(path, "pdf", joined, fio, iin)

    for p in docxs:
        path = str(p)
        try:
            joined, _lines = read_docx_text(path)
        except Exception as e:
            log(f"[DOCS] Не смог прочитать DOCX: {path} ({type(e).__name__}: {e})")
            continue
        iin = extract_iin_or_zsn_from_text(joined)
        fio = extract_debtor_fio_from_text(joined)
        add_doc(path, "docx", joined, fio, iin)

    log(f"[DOCS] В индексе ИИН/ЖСН: {len(by_iin)} (уникальных)")
    return by_iin, all_docs


DOCS_BY_IIN, DOCS_ALL = build_docs_index_from_text(IL_FOLDER)


def find_doc_for_key(key: str):
    if not key:
        return None

    key_str = str(key).strip()
    if not key_str:
        return None

    key_iin = extract_12_digits_from_key(key_str)
    if key_iin:
        candidates = DOCS_BY_IIN.get(key_iin, [])
        if not candidates:
            return None
        return candidates[0]["path"]

    norm_key = normalize_text(key_str)
    if not norm_key:
        return None

    for doc in DOCS_ALL:
        if doc["fio_norm"] and doc["fio_norm"] == norm_key:
            return doc["path"]

    for doc in DOCS_ALL:
        if doc["fio_norm"] and (norm_key in doc["fio_norm"] or doc["fio_norm"] in norm_key):
            return doc["path"]

    return None


def find_region_in_text(text: str):
    if not REGIONS_LIST:
        return None, 0.0

    norm_text = normalize_text(text)
    if not norm_text:
        return None, 0.0

    text_tokens_core = tokens_without_stopwords(norm_text)

    best = None
    best_score = 0.0

    for r in REGIONS_LIST:
        tokens_core = r["tokens_core"]
        token_score = (len(tokens_core & text_tokens_core) / float(len(tokens_core))) if tokens_core else 0.0
        char_score = difflib.SequenceMatcher(None, r["norm"], norm_text).ratio()
        score = 0.8 * token_score + 0.2 * char_score
        if score > best_score:
            best = r
            best_score = score

    if best and best_score >= REGION_MATCH_THRESHOLD:
        return best["region"], best_score
    return None, best_score


def extract_court_header_line(header_lines):
    STAFF_MARKERS = ["секретарь", "заседания", "заседание", "администратора", "администратор", "подписи", "подпись"]
    IGNORE_MARKERS = ["вступлен", "законн", "акт"]

    def has_court_word(text_norm: str) -> bool:
        tokens = text_norm.split()
        court_tokens = {"суд", "суда", "суды", "суде", "суду", "судом", "сот", "соты"}
        return any(t in court_tokens for t in tokens)

    raw_candidates = []
    for i, ln in enumerate(header_lines):
        if not ln.strip():
            continue
        combos = [ln]
        if i > 0:
            combos.append(header_lines[i-1] + " " + ln)
        for combo in combos:
            norm_combo = normalize_text(combo)
            if not norm_combo:
                continue
            if not has_court_word(norm_combo):
                continue
            raw_candidates.append(combo)

    if not raw_candidates:
        return None

    good_candidates = []
    for c in raw_candidates:
        low = c.lower()
        if any(mark in low for mark in STAFF_MARKERS):
            continue
        if any(mark in low for mark in IGNORE_MARKERS):
            continue
        good_candidates.append(c)

    candidates = good_candidates if good_candidates else raw_candidates

    uniq = []
    seen = set()
    for c in candidates:
        if c not in seen:
            uniq.append(c)
            seen.add(c)

    best_line = None
    best_score = -1
    for ln in uniq:
        norm = normalize_text(ln)
        core = tokens_without_stopwords(norm)
        score = len(core)
        if score > best_score:
            best_score = score
            best_line = ln

    return best_line


def find_court_for_region(court_text: str, region_name: str):
    if not court_text or not region_name:
        return None, 0.0

    candidates = [c for c in COURTS_LIST if c["region"] == region_name]
    if not candidates:
        return None, 0.0

    norm_text = normalize_text(court_text)
    text_tokens_core = tokens_without_stopwords(norm_text)

    scored = []
    for c in candidates:
        tokens_core = c["tokens_core"]
        overlap = len(tokens_core & text_tokens_core)
        token_score = (overlap / float(len(tokens_core))) if tokens_core else 0.0
        char_score = difflib.SequenceMatcher(None, c["norm"], norm_text).ratio()
        score = char_score if overlap == 0 else (0.7 * token_score + 0.3 * char_score)
        scored.append({"court": c["court"], "score": score})

    scored.sort(key=lambda x: x["score"], reverse=True)
    best = scored[0] if scored else None

    if best and best["score"] >= COURT_MATCH_THRESHOLD:
        return best["court"], best["score"]
    return None, best["score"] if best else 0.0


def find_to_chsi_for_address(fact_address: str):
    if not fact_address or not CHSI_ROWS:
        return None, 0.0

    norm_addr = normalize_text(fact_address)
    tokens_addr = tokens_without_stopwords(norm_addr)

    scored = []
    for r in CHSI_ROWS:
        to_tokens   = r["tokens_to"]
        dist_tokens = r["tokens_district"]

        overlap_to   = len(to_tokens & tokens_addr) / float(len(to_tokens)) if to_tokens else 0.0
        overlap_dist = len(dist_tokens & tokens_addr) / float(len(dist_tokens)) if dist_tokens else 0.0

        char_to   = difflib.SequenceMatcher(None, r["norm_to"], norm_addr).ratio()
        char_dist = difflib.SequenceMatcher(None, r["norm_district"], norm_addr).ratio() if r["norm_district"] else 0.0

        score = 0.35 * overlap_to + 0.45 * overlap_dist + 0.10 * char_to + 0.10 * char_dist
        scored.append({"row": r, "score": score})

    scored.sort(key=lambda x: x["score"], reverse=True)
    best = scored[0] if scored else None

    if best and best["score"] >= CHSI_MATCH_THRESHOLD:
        return best["row"], best["score"]
    return None, best["score"] if best else 0.0


def find_chsi_row_by_name(chsi_name: str):
    if not chsi_name or not CHSI_ROWS:
        return None, 0.0

    target_norm = normalize_text(chsi_name)
    if not target_norm:
        return None, 0.0

    best = None
    best_score = 0.0

    for r in CHSI_ROWS:
        cand_norm = r.get("norm_chsi") or ""
        if not cand_norm:
            continue

        if cand_norm == target_norm:
            return r, 1.0

        if target_norm in cand_norm or cand_norm in target_norm:
            score = 0.90
        else:
            score = difflib.SequenceMatcher(None, cand_norm, target_norm).ratio()

        if score > best_score:
            best_score = score
            best = r

    if best and best_score >= CHSI_NAME_MATCH_THRESHOLD:
        return best, best_score

    return None, best_score


def extract_fact_address_from_pdf(joined: str):
    if not joined:
        return None

    block_match = re.search(
        r"Полное наименование ответчика и его адрес:(.+?)(?:Полное наименование|ИСПОЛНИТЕЛЬНЫЙ ЛИСТ|Дата выписки|$)",
        joined,
        flags=re.DOTALL | re.IGNORECASE
    )
    if block_match:
        block_text = block_match.group(1)
        m_fact = re.search(r"Фактический адрес:\s*(.+)", block_text, flags=re.IGNORECASE)
        if m_fact:
            return m_fact.group(1).splitlines()[0].strip()

    m_kz = re.search(r"Нақты\s+мекенжайы\s*:\s*(.+)", joined, flags=re.IGNORECASE)
    if m_kz:
        return m_kz.group(1).splitlines()[0].strip()

    m_fact2 = re.search(r"Фактический адрес:\s*(.+)", joined, flags=re.IGNORECASE)
    if m_fact2:
        return m_fact2.group(1).splitlines()[0].strip()

    return None

def extract_fact_address_from_docx(joined: str):
    if not joined:
        return None

    m = re.search(r"Ответчик\s*:\s*([^\n\r]+)", joined, flags=re.IGNORECASE)
    if m:
        line = m.group(1).strip()
        m_i = re.search(r"\bИИН\b[:\s]*\d{12}\s*,?\s*(.+)$", line, flags=re.IGNORECASE)
        if m_i:
            addr = m_i.group(1).strip()
            if addr:
                return addr

    m2 = re.search(r"Фактический адрес:\s*(.+)", joined, flags=re.IGNORECASE)
    if m2:
        return m2.group(1).splitlines()[0].strip()

    m3 = re.search(r"Юридический адрес:\s*(.+)", joined, flags=re.IGNORECASE)
    if m3:
        return m3.group(1).splitlines()[0].strip()

    return None


def extract_data_from_pdf(pdf_path: str):
    log(f"[PDF] Читаю файл: {pdf_path}")
    joined, lines = read_pdf_text(pdf_path)

    # ==========================
    # ✅ ИЗМЕНЕНО ТОЛЬКО ЭТО МЕСТО:
    # Суд/регион определяем по 1-й странице, чтобы не ловить "область ..." с других страниц
    # ==========================
    _p1_joined, p1_lines = read_pdf_first_page_text(pdf_path)
    header_lines = p1_lines[:60]   # чуть шире, но только 1-я страница
    header_text  = "\n".join(header_lines)
    # ==========================

    fact_address = extract_fact_address_from_pdf(joined)

    num = None
    m_num = re.search(r"^[^\S\r\n]*№\s*([0-9][0-9\-–/]*\/[0-9\-–/]+)", joined, flags=re.MULTILINE)
    if not m_num:
        m_num = re.search(r"ИСПОЛНИТЕЛЬНЫЙ\s+ЛИСТ.*?№\s*([0-9][0-9\-–/]+)", joined, flags=re.DOTALL | re.IGNORECASE)
    if not m_num:
        m_num = re.search(r"№\s*([0-9][0-9\-–/]+)", joined)
    if m_num:
        num = m_num.group(1).strip()

    date_str = None
    m_date = re.search(r"Дата выписки\s+(\d{2}\.\d{2}\.\d{4})", joined, flags=re.IGNORECASE)
    if m_date:
        date_str = m_date.group(1)
    else:
        m_date2 = re.search(r"(Шығарылу\s+күні|Шыгарылу\s+күні)\s+(\d{2}\.\d{2}\.\d{4})", joined, flags=re.IGNORECASE)
        if m_date2:
            date_str = m_date2.group(2)

    iin = extract_iin_or_zsn_from_text(joined)
    fio = extract_debtor_fio_from_text(joined)

    # ==========================
    # ✅ ИЗМЕНЕНО ТОЛЬКО ЭТО МЕСТО:
    # регион ищем по строке суда (приоритет), а не по всему header_text
    # ==========================
    court_header_line = extract_court_header_line(header_lines)

    # ✅ регион ищем по header_text (там есть "город Алматы")
    court_region, _ = find_region_in_text(header_text)

    # ✅ суд ищем по строке суда, но уже в найденном регионе
    court_name = None
    if court_header_line and court_region:
        court_name, _ = find_court_for_region(court_header_line, court_region)

    # ✅ дополнительный fallback: если суд не найден, пробуем по header_text
    if not court_name and court_region:
        court_name, _ = find_court_for_region(header_text, court_region)

    # ==========================

    log(f"[PDF] Найдено: номер={num}, дата={date_str}, ИИН/ЖСН={iin}, ФИО={fio or 'не найдено'}")
    return {
        "number": num, "date": date_str, "iin": iin, "fio": fio,
        "court_name": court_name, "court_region": court_region,
        "fact_address": fact_address, "doc_kind": "pdf"
    }

def extract_data_from_docx(docx_path: str):
    log(f"[DOCX] Читаю файл: {docx_path}")
    joined, lines = read_docx_text(docx_path)
    header_lines = lines[:50]
    header_text  = "\n".join(header_lines)

    fact_address = extract_fact_address_from_docx(joined)

    num = None
    m_num = re.search(r"№\s*([0-9][0-9\-–/]+)", joined)
    if m_num:
        num = m_num.group(1).strip()

    date_str = None
    m_date = re.search(r"\b(\d{2}\.\d{2}\.\d{4})\b", header_text)
    if m_date:
        date_str = m_date.group(1)
    else:
        date_str = parse_ru_date_words(header_text) or parse_ru_date_words(joined)

    iin = extract_iin_or_zsn_from_text(joined)
    fio = extract_debtor_fio_from_text(joined)

    court_region, _ = find_region_in_text(header_text)
    court_header_line = extract_court_header_line(header_lines)
    court_name = None
    if court_header_line and court_region:
        court_name, _ = find_court_for_region(court_header_line, court_region)

    log(f"[DOCX] Найдено: номер={num}, дата={date_str}, ИИН={iin}, ФИО={fio or 'не найдено'}")
    return {
        "number": num, "date": date_str, "iin": iin, "fio": fio,
        "court_name": court_name, "court_region": court_region,
        "fact_address": fact_address, "doc_kind": "docx"
    }

def extract_data_from_document(path: str):
    ext = Path(path).suffix.lower()
    if ext == ".pdf":
        return extract_data_from_pdf(path)
    if ext == ".docx":
        return extract_data_from_docx(path)
    raise ValueError(f"Неподдерживаемый тип файла: {ext}")


def make_driver():
    opts = webdriver.ChromeOptions()
    prefs = {"safebrowsing.enabled": True, "profile.managed_default_content_settings.images": 2}
    opts.add_experimental_option("prefs", prefs)
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--start-maximized")
    service = Service(ChromeDriverManager().install())
    drv = webdriver.Chrome(service=service, options=opts)
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
        By.CSS_SELECTOR,
        "#signInButton, form#loginForm button.ui.primary.button.fluid, button[type='submit']"
    )))
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    except Exception:
        pass
    for _ in range(2):
        try:
            btn.click()
        except Exception:
            pass
        time.sleep(0.2)
        try:
            driver.execute_script("arguments[0].click();", btn)
        except Exception:
            pass
        time.sleep(0.3)

def is_logged_in(driver):
    try:
        return "/cabinet/" in (driver.current_url or "")
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
            try:
                driver.refresh()
            except Exception:
                pass
            time.sleep(1.0)
            continue

        fill_login_form(driver)
        time.sleep(0.2)
        click_sign_in(driver)

        try:
            WebDriverWait(driver, 15).until(lambda d: is_logged_in(d))
            log("[OK] Авторизация выполнена.")
            return
        except TimeoutException:
            try:
                driver.refresh()
            except Exception:
                pass
            time.sleep(1.0)

    raise TimeoutException("Не удалось войти после серии повторов.")


def open_statements_tab(driver):
    log("[NAV] Открываю вкладку 'Подача заявления ЧСИ'")
    driver.get(STATEMENTS_URL)
    wait = WebDriverWait(driver, 30)
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(.,'Сформированные заявления')]")))
    except Exception:
        wait.until(EC.url_contains("/cabinet/statements"))
    log("[NAV] Вкладка 'Подача заявления ЧСИ' открыта.")

def click_podat_zayavlenie(driver):
    log("[NAV] Жму кнопку 'Подать заявление'")
    wait = WebDriverWait(driver, 30)
    xpath_btn = (
        "//span[contains(@class,'v-btn__content') and normalize-space()='Подать заявление']"
        "/ancestor::*[self::a or self::button][1]"
    )
    btn = wait.until(EC.presence_of_element_located((By.XPATH, xpath_btn)))
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    except Exception:
        pass
    try:
        driver.execute_script("arguments[0].click();", btn)
    except Exception:
        btn.click()
    wait.until(EC.url_contains("/cabinet/statements/new"))
    log("[NAV] Форма 'Подать заявление' открыта.")


def fill_statement_form(driver, data):
    wait = WebDriverWait(driver, 20)
    actions = ActionChains(driver)

    log("[FORM] Выбираю 'Источник ИД' = АИАС «Төрелік»")
    src_input = wait.until(EC.element_to_be_clickable((
        By.XPATH, "//label[contains(.,'Источник ИД')]/following::input[1]"
    )))
    src_input.click()
    time.sleep(0.3)

    src_item = wait.until(EC.element_to_be_clickable((
        By.XPATH, "//div[contains(@class,'v-list-item__title') and normalize-space()='АИАС «Төрелік»']"
    )))
    try:
        src_item.click()
    except Exception:
        driver.execute_script("arguments[0].click();", src_item)

    if data.get("number"):
        log("[FORM] Заполняю 'Номер исполнительного документа'")
        num_input = wait.until(EC.element_to_be_clickable((
            By.XPATH, "//label[contains(.,'Номер исполнительного документа')]/following::input[1]"
        )))
        num_input.click()
        actions.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).send_keys(Keys.BACKSPACE).perform()
        num_input.send_keys(data["number"])

    if data.get("date"):
        log("[FORM] Заполняю 'Дата выписки исполнительного документа'")
        date_input = wait.until(EC.element_to_be_clickable((
            By.XPATH, "//label[contains(.,'Дата выписки исполнительного документа')]/following::input[1]"
        )))
        date_input.click()
        actions.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).send_keys(Keys.BACKSPACE).perform()
        date_input.send_keys(data["date"])

    court_region = data.get("court_region")
    court_name   = data.get("court_name")
    if court_region and court_name:
        log(f"[FORM] Выбираю суд: {court_region} → {court_name}")
        court_input = wait.until(EC.presence_of_element_located((
            By.XPATH, "//label[contains(.,'Орган, выдавший исполнительный документ')]/following::input[1]"
        )))
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", court_input)
        except Exception:
            pass
        try:
            driver.execute_script("arguments[0].click();", court_input)
        except Exception:
            court_input.click()
        time.sleep(0.5)

        region_xpath = "//div[contains(@class,'v-list-item__title') and contains(normalize-space(), '{}')]".format(court_region)
        region_el = wait.until(EC.element_to_be_clickable((By.XPATH, region_xpath)))
        try:
            driver.execute_script("arguments[0].click();", region_el)
        except Exception:
            region_el.click()
        time.sleep(0.4)

        court_xpath = "//div[contains(@class,'v-list-item__title') and contains(normalize-space(), '{}')]".format(court_name)
        court_el = wait.until(EC.element_to_be_clickable((By.XPATH, court_xpath)))
        try:
            driver.execute_script("arguments[0].click();", court_el)
        except Exception:
            court_el.click()
    else:
        log("[FORM] Суд не определён, выбор суда пропускаю!")

    if data.get("iin"):
        log("[FORM] Заполняю 'ИИН/БИН должника'")
        iin_input = wait.until(EC.element_to_be_clickable((
            By.XPATH, "//label[contains(.,'ИИН/БИН должника')]/following::input[1]"
        )))
        iin_input.click()
        actions.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).send_keys(Keys.BACKSPACE).perform()
        iin_input.send_keys(data["iin"])

    if data.get("fio"):
        log("[FORM] Заполняю 'ФИО/Наименование должника'")
        fio_input = wait.until(EC.element_to_be_clickable((
            By.XPATH, "//label[contains(.,'ФИО/Наименование должника')]/following::input[1]"
        )))
        fio_input.click()
        actions.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).send_keys(Keys.BACKSPACE).perform()
        fio_input.send_keys(data["fio"])


def click_next(driver, label="Далее", timeout=25):
    log(f"[FORM] Жму кнопку '{label}'")
    btn = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.XPATH, "//button[@id='next-btn' or .//span[normalize-space()='Далее']]"))
    )
    try:
        btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", btn)
    time.sleep(2.0)


def click_final_dalee(driver, timeout=25) -> bool:
    """
    Финальная кнопка 'Далее' после выбора ТО/район/ЧСИ.
    """
    btn_xpath = (
        "//span[contains(@class,'v-btn__content') and normalize-space()='Далее']"
        "/ancestor::*[self::button or self::a][1]"
    )
    btn = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH, btn_xpath)))
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    except Exception:
        pass

    try:
        btn.click()
        return True
    except Exception:
        pass

    try:
        driver.execute_script("arguments[0].click();", btn)
        return True
    except Exception:
        return False


def find_zayavlenie_file(iin: str):
    if not iin:
        return None
    folder = Path(STATEMENTS_FOLDER)
    if not folder.exists():
        log(f"[UPLOAD] Папка с заявлениями не найдена: {STATEMENTS_FOLDER}")
        return None

    for f in folder.glob("*.pdf"):
        if iin in f.name:
            return str(f)
    return None


def upload_files(driver, doc_path: str, iin: str):
    wait = WebDriverWait(driver, 20)
    log("[UPLOAD] Прикрепляю исполнительный документ (ИЛ/СП)")
    try:
        exec_input = wait.until(EC.presence_of_element_located((By.XPATH, "(//input[@type='file'])[1]")))
    except TimeoutException:
        try:
            banners = driver.find_elements(By.XPATH, "//*[contains(.,'Данный исполнительный документ уже зарегистрирован в Системе')]")
        except Exception:
            banners = []
        if banners:
            log("[UPLOAD] ИД уже зарегистрирован — пропускаю.")
            raise AlreadyRegisteredException()
        raise

    exec_input.send_keys(doc_path)
    time.sleep(1.0)

    other_path = find_zayavlenie_file(iin)
    if not other_path:
        log(f"[UPLOAD] Не найдено 'Заявление ...' по ИИН {iin} — второе поле пропускаю.")
        return

    log(f"[UPLOAD] Прикрепляю 'Заявление ЧСИ': {other_path}")
    other_input = wait.until(EC.presence_of_element_located((By.XPATH, "(//input[@type='file'])[2]")))
    other_input.send_keys(other_path)
    time.sleep(1.0)


# ==========================================================
# !!! ВСТАВЛЕН ТВОЙ БЛОК fill_to_and_chsi БЕЗ ИЗМЕНЕНИЙ !!!
# ==========================================================
def fill_to_and_chsi(
    driver,
    to_name: str,
    district_name: str,
    chsi_name: str,
    timeout=40,
    log=print,
    DEBUG_CHSI=False,
    chsi_dbg_save=None,
):
    wait = WebDriverWait(driver, timeout)

    def norm(s):
        return " ".join((s or "").replace("\u00a0", " ").split()).strip().upper()

    def js_click(el):
        try:
            driver.execute_script("arguments[0].click();", el)
            return True
        except Exception:
            try:
                el.click()
                return True
            except Exception:
                return False

    def wrapper_by_label(label):
        return wait.until(EC.presence_of_element_located((
            By.XPATH,
            f"//label[contains(normalize-space(),'{label}')]/ancestor::div[contains(@class,'v-input')][1]"
        )))

    # ✅ ВАЖНО: читаем value не из input, а из .v-select__selection
    def read_value(label):
        try:
            wrap = wrapper_by_label(label)

            # 1) то, что реально отображается пользователю
            sels = wrap.find_elements(By.CSS_SELECTOR, ".v-select__selection")
            if sels:
                txt = " ".join([s.text.strip() for s in sels if s.text.strip()]).strip()
                if txt:
                    return txt

            # 2) иногда текст прямо в selections
            try:
                cont = wrap.find_element(By.CSS_SELECTOR, ".v-select__selections")
                txt2 = cont.text.strip()
                if txt2:
                    return txt2
            except Exception:
                pass

            # 3) fallback: input.value (часто пустой)
            inp = wrap.find_element(By.CSS_SELECTOR, "input")
            return (inp.get_attribute("value") or "").strip()
        except Exception:
            return ""

    def open_menu(wrapper):
        try:
            slot = wrapper.find_element(By.CSS_SELECTOR, ".v-input__slot")
            js_click(slot)
        except Exception:
            js_click(wrapper)
        time.sleep(0.35)

    def get_menu():
        return wait.until(EC.presence_of_element_located((
            By.CSS_SELECTOR,
            "div.v-menu__content.menuable__content__active"
        )))

    def close_menu():
        # клик по body обычно закрывает dropdown
        try:
            driver.find_element(By.TAG_NAME, "body").click()
        except Exception:
            pass
        time.sleep(0.2)

    def wait_not_disabled(label, seconds=15):
        t_end = time.time() + seconds
        while time.time() < t_end:
            try:
                wrap = wrapper_by_label(label)
                cls = (wrap.get_attribute("class") or "")
                if "v-input--is-disabled" not in cls:
                    return True
            except Exception:
                pass
            time.sleep(0.2)
        return False

    def find_option_exact(menu, target_text):
        target = norm(target_text)
        opts = menu.find_elements(By.CSS_SELECTOR, "div[role='option']")
        for o in opts:
            try:
                title = o.find_element(By.CSS_SELECTOR, ".v-list-item__title").text
            except Exception:
                title = o.text
            if norm(title) == target:
                return o, title
        return None, None

    def scroll_menu_down(menu, step=3.5):
        # Скроллим именно меню
        try:
            driver.execute_script(
                "arguments[0].scrollTop = arguments[0].scrollTop + arguments[0].clientHeight * arguments[1];",
                menu, step
            )
        except Exception:
            try:
                driver.execute_script("""
                    const el = arguments[0];
                    el.dispatchEvent(new WheelEvent('wheel', {deltaY: 800, bubbles: true}));
                """, menu)
            except Exception:
                pass

    def select_from_dropdown(label, value, max_steps=90):
        if not value or not value.strip():
            return True

        target = value.strip()

        if not wait_not_disabled(label, seconds=20):
            log(f"[FORM] '{label}': поле disabled, не могу выбрать.")
            return False

        wrap = wrapper_by_label(label)

        # лог до
        before = read_value(label)
        log(f"[FORM-DBG] '{label}': before='{before}' target='{target}'")

        open_menu(wrap)
        menu = get_menu()

        # ждём, что реально появились options
        t_end = time.time() + 8
        while time.time() < t_end:
            if menu.find_elements(By.CSS_SELECTOR, "div[role='option']"):
                break
            time.sleep(0.2)

        last_scroll = None

        for step in range(1, max_steps + 1):
            el, title = find_option_exact(menu, target)
            if el:
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                except Exception:
                    pass
                time.sleep(0.15)
                js_click(el)
                time.sleep(0.3)

                # закрыть меню (важно для Vuetify, иногда value обновляется после blur)
                close_menu()

                now = read_value(label)
                ok = (norm(now) == norm(target))
                log(f"[FORM-DBG] '{label}': click '{title}' -> now='{now}' ok={ok}")
                return ok

            # debug: видимые первые пункты и scroll
            try:
                titles = menu.find_elements(By.CSS_SELECTOR, ".v-list-item__title")
                visible = [t.text.strip() for t in titles[:12]]
            except Exception:
                visible = []
            try:
                st = driver.execute_script("return arguments[0].scrollTop;", menu)
                ch = driver.execute_script("return arguments[0].clientHeight;", menu)
                sh = driver.execute_script("return arguments[0].scrollHeight;", menu)
            except Exception:
                st, ch, sh = None, None, None

            log(f"[FORM-DBG] '{label}' step={step} visible={visible} scrollTop={st} clientH={ch} scrollH={sh}")

            scroll_menu_down(menu, step=3.5)
            time.sleep(0.10)

            try:
                st2 = driver.execute_script("return arguments[0].scrollTop;", menu)
            except Exception:
                st2 = None

            if last_scroll is not None and st2 == last_scroll:
                break
            last_scroll = st2

        close_menu()
        log(f"[FORM] '{label}': '{target}' НЕ найден в списке.")
        if DEBUG_CHSI and chsi_dbg_save:
            chsi_dbg_save(driver, f"not_found_{label}")
        return False

    # ====== ТО ======
    if to_name:
        log(f"[FORM] Выбираю ТО: {to_name}")
        ok = select_from_dropdown("Выберите ТО", to_name, max_steps=80)
        if not ok:
            log("[FORM] ТО: НЕ выбран.")
            return False
        time.sleep(0.6)

    # ====== Район ======
    if district_name:
        dn = district_name.strip()
        if dn in {"-", "—"}:
            log("[FORM] Район не требуется (городское ТО) — пропускаю выбор района.")
        else:
            log(f"[FORM] Выбираю Район: {district_name}")
            ok = select_from_dropdown("Выберите Район", district_name, max_steps=120)
            if not ok:
                log("[FORM] Район: НЕ выбран.")
                return False
            time.sleep(0.6)


    # ====== ЧСИ ======
    if chsi_name:
        log(f"[FORM] Выбираю ЧСИ: {chsi_name}")
        time.sleep(0.8)  # дождаться обновления списка ЧСИ после района
        ok = select_from_dropdown("ФИО ЧСИ", chsi_name, max_steps=180)
        if not ok:
            log("[FORM] ЧСИ: НЕ выбран.")
            return False

        now = read_value("ФИО ЧСИ")
        if norm(now) != norm(chsi_name):
            log(f"[FORM] ЧСИ: значение не совпало. expected='{chsi_name}', now='{now}'")
            if DEBUG_CHSI and chsi_dbg_save:
                chsi_dbg_save(driver, "chsi_value_mismatch")
            return False

        log(f"[FORM] ЧСИ выбран: '{now}'")

    log("[FORM] ТО / Район / ЧСИ выбраны успешно")
    return True
# ==========================================================


def run():
    driver = make_driver()
    try:
        login_with_retries(driver)
        open_statements_tab(driver)

        log(f"[MAIN] В списке для подачи: {len(DEALS)} записей")

        for deal in DEALS:
            key = deal["key"]
            excel_address = (deal.get("excel_address") or "").strip()
            deal_chsi = (deal.get("deal_chsi") or "").strip()

            doc_path = find_doc_for_key(key)
            if not doc_path:
                log(f"[MAIN] Для ключа '{key}' не найден документ (PDF/DOCX) в папке ИЛ — пропускаю.")
                continue

            log(f"[MAIN] Обрабатываю ключ '{key}', Документ: {doc_path}")

            try:
                data = extract_data_from_document(doc_path)
                if not (data.get("number") and data.get("date") and data.get("iin")):
                    log("[MAIN] Нет номера/даты/ИИН(ЖСН) — пропускаю.")
                    continue

                if excel_address:
                    data["fact_address"] = excel_address
                    log(f"[MAIN] Адрес для fallback взят из Excel (колонка I): {excel_address}")
                else:
                    if data.get("fact_address"):
                        log(f"[MAIN] Адрес для fallback взят из документа: {data['fact_address']}")
                    else:
                        log("[MAIN] ВНИМАНИЕ: адрес не найден ни в Excel, ни в документе.")

                open_statements_tab(driver)
                click_podat_zayavlenie(driver)

                fill_statement_form(driver, data)
                click_next(driver, label="Далее (страница 1)")

                upload_files(driver, doc_path, data["iin"])
                click_next(driver, label="Далее (страница 2)")

                # ТО/Район/ЧСИ: сначала по ФИО ЧСИ из сделок, иначе fallback по адресу
                row_chsi = None
                score = 0.0
                source = ""

                if deal_chsi:
                    row_chsi, score = find_chsi_row_by_name(deal_chsi)
                    if row_chsi:
                        source = "by_name"
                        log(f"[CHSI] Найден по ФИО ЧСИ из списка сделок (score={score:.2f}): '{deal_chsi}'")
                    else:
                        log(f"[CHSI] НЕ найден по ФИО ЧСИ из списка сделок (score={score:.2f}): '{deal_chsi}'")

                if not row_chsi:
                    row_chsi, score = find_to_chsi_for_address(data.get("fact_address") or "")
                    if not row_chsi:
                        raise Exception(f"Не удалось определить ТО/район/ЧСИ (score={score:.2f}).")
                    source = "by_address"
                    log(f"[CHSI] Fallback по адресу (score={score:.2f})")

                to_name = row_chsi.get("to", "")
                district_name = row_chsi.get("district", "")
                chsi_name = row_chsi.get("chsi", "")

                log(f"[CHSI] Итог ({source}, score={score:.2f}): ТО='{to_name}', район='{district_name}', ЧСИ='{chsi_name}'")

                ok = fill_to_and_chsi(
                    driver,
                    to_name,
                    district_name,
                    chsi_name,
                    timeout=40,
                    log=log,
                    DEBUG_CHSI=DEBUG_CHSI,
                    chsi_dbg_save=chsi_dbg_save,
                )
                if not ok:
                    raise Exception("Не удалось выбрать ТО/район/ЧСИ на странице")

                # ==========================================================
                # ✅ ЕДИНСТВЕННОЕ ДОБАВЛЕНИЕ: финальная кнопка "Далее"
                # (не трогаем твой fill_to_and_chsi, жмём после него)
                # ==========================================================
                log("[FORM] Жму финальную кнопку 'Далее' (после ТО/район/ЧСИ)")
                if not click_final_dalee(driver, timeout=25):
                    raise Exception("Не удалось нажать финальную кнопку 'Далее'")
                time.sleep(0.8)
                # ==========================================================

                log(f"[MAIN] Готово: {data['iin']} обработан.\n")

            except AlreadyRegisteredException:
                log(f"[MAIN] ИД по ключу '{key}' уже зарегистрирован. Следующая запись.\n")
                continue
            except Exception as e:
                log(f"[MAIN] ОШИБКА при обработке ключа '{key}': {type(e).__name__}: {e}. Пропускаю.\n")
                continue

        log("[MAIN] Все записи обработаны. Оставляю браузер открытым.")
        input("Нажмите Enter, чтобы закрыть браузер и завершить скрипт...")

    finally:
        driver.quit()
        log("[EXIT] Браузер закрыт.")



