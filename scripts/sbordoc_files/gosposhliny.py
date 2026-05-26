# -*- coding: utf-8 -*-
import os
import re
import shutil
from datetime import datetime

from pypdf import PdfReader, PdfWriter

# нестрогое сравнение ФИО
try:
    from rapidfuzz.fuzz import token_set_ratio as fuzz_ratio
    HAVE_RF = True
except Exception:
    from difflib import SequenceMatcher
    HAVE_RF = False

from config import MAIN_EXCEL, ROOT, TARGET_BASE
from utils import safe_log, safe_update_summary

# ===== ПУТИ =====
SOURCE_ROOT = rf"{ROOT}\Госпошлины"  # внутри: 2023, 2024, 2025 ...
DEST_ROOT   = TARGET_BASE
log_file_path = 'log_not_copied.txt'

try:
    LOG_SUMMARY
except NameError:
    LOG_SUMMARY = {}


# ===== УТИЛИТЫ =====
def normalize_name(s: str) -> str:
    s = (s or "").lower()
    s = s.replace('\xa0', ' ').replace('\u202f', ' ').replace('\u200b', ' ')
    s = s.replace('ё', 'е')
    s = re.sub(r'[.,;:()"\'`]+', ' ', s)
    s = re.sub(r'[-_]+', ' ', s)
    s = re.sub(r'\s{2,}', ' ', s).strip()
    return s

def similarity(a: str, b: str) -> float:
    a, b = normalize_name(a), normalize_name(b)
    if HAVE_RF:
        return float(fuzz_ratio(a, b))
    return SequenceMatcher(None, a, b).ratio() * 100.0

def sanitize_filename(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', '', name)

def _clean_text(text: str) -> str:
    if not text:
        return ""
    t = (text.replace('\xa0', ' ')
              .replace('\u202f', ' ')
              .replace('\u200b', ' '))
    t = re.sub(r'[\n\r\f]+', ' ', t)
    t = re.sub(r'\s{2,}', ' ', t).strip()
    return t


# ===== ИЗВЛЕЧЕНИЕ ФИО =====
# маркеры, на которых обрезаем хвост после ФИО
BUDGET_STOP_RE = r'(?:Бюджеттік|Бюджетке|Код\s+бюджетной|бюджетной\s+классификации|жіктеу\s+коды|классификация)'

# слова, которые не являются частью ФИО, но могут стоять
# перед именем в назначении платежа
FILLER_WORDS = {
    'госпошлина', 'госпошлины', 'алым', 'алымы',
    'мемлекеттік', 'оплата', 'плата', 'взнос',
    'за', 'үшін', 'бойынша', 'толемі', 'төлемі',
}

def _strip_filler(words: list[str]) -> list[str]:
    while words and words[0].lower().strip('.,') in FILLER_WORDS:
        words = words[1:]
    return words


def _words_to_fio(raw: str) -> str | None:
    """Общий постпроцессинг сырого куска текста -> красиво оформленное ФИО (до 3 слов)."""
    words = raw.strip().split()
    words = _strip_filler(words)
    words = words[:3]  # Фамилия Имя Отчество

    if len(words) >= 2:
        return " ".join(w[0].upper() + w[1:] for w in words)
    return None


def _extract_fio_old_format(t: str) -> str | None:
    """
    Старый формат платёжки: ФИО стоит МЕЖДУ кодом '911' и словом 'Сумма'.
    Пример: '... 911 Иванов Иван Иванович Сумма прописью ...'
    """
    m = re.search(r'\b911\s+(.+?)\s+Сумма', t)
    if not m:
        m = re.search(
            r'(?:Төлемнің\s+мақсаты|Назначение\s+платежа)\s*:.+?911\s+(.+?)\s+Сумма',
            t, flags=re.IGNORECASE
        )
    if not m:
        return None
    return _words_to_fio(m.group(1))


def _extract_fio_new_format(t: str) -> str | None:
    """
    Новый формат платёжки (например, Freedom Bank): ФИО стоит СРАЗУ ПОСЛЕ
    строки 'Назначение платежа' / 'Төлем мақсаты' и ПЕРЕД пояснительной
    скобкой '(тауар атауын...)'. Код '911' в этом формате идёт значительно
    позже (в отдельном поле "Код назначения платежа"), поэтому старый
    паттерн '911 ... Сумма' здесь не срабатывает.

    Извлекаем 2-3 идущих подряд слова с заглавной буквы, стоящих
    непосредственно перед '(тауар' — это устойчиво к тому, где именно
    в тексте оказался код 911 после извлечения текста из PDF.
    """
    m = re.search(
        r'([А-ЯЁ][а-яёА-ЯЁ\-]+(?:\s+[А-ЯЁ][а-яёА-ЯЁ\-]+){1,2})\s*\(\s*тауар',
        t
    )
    if not m:
        return None
    return _words_to_fio(m.group(1))


def extract_fio(page_text: str) -> str | None:
    t = _clean_text(page_text)
    if not t:
        return None

    fio = _extract_fio_old_format(t)
    if fio:
        return fio

    fio = _extract_fio_new_format(t)
    if fio:
        return fio

    return None

def _extract_fio_between_markers(text: str) -> str | None:
    # оставляем для обратной совместимости, но делегируем в extract_fio
    return extract_fio(text)


# ===== ПАПКИ =====
def latest_year_dir(root: str) -> str | None:
    now_year = str(datetime.now().year)
    candidates = [d for d in os.listdir(root) if os.path.isdir(os.path.join(root, d))]
    if not candidates:
        return None
    if now_year in candidates:
        return os.path.join(root, now_year)
    candidates_num = sorted(candidates, key=lambda x: int(re.sub(r'\D', '', x) or 0), reverse=True)
    return os.path.join(root, candidates_num[0])

DATE_RE = re.compile(r'(\d{2})\.(\d{2})\.(\d{4})')  # DD.MM.YYYY

def parse_date_from_foldername(name: str) -> datetime | None:
    m = DATE_RE.search(name)
    if not m:
        return None
    d, mth, y = map(int, m.groups())
    try:
        return datetime(y, mth, d)
    except ValueError:
        return None

def pick_latest_day_folder(year_dir: str) -> str | None:
    subdirs = [os.path.join(year_dir, d) for d in os.listdir(year_dir)
               if os.path.isdir(os.path.join(year_dir, d))]
    if not subdirs:
        return None

    dated, undated = [], []
    for p in subdirs:
        dt = parse_date_from_foldername(os.path.basename(p))
        if dt:
            dated.append((dt, p))
        else:
            undated.append(p)

    if dated:
        dated.sort(key=lambda x: x[0], reverse=True)
        return dated[0][1]

    return max(undated, key=lambda p: os.path.getmtime(p))

def list_client_folders(dest_root: str):
    res = []
    if not os.path.isdir(dest_root):
        return res
    for name in os.listdir(dest_root):
        p = os.path.join(dest_root, name)
        if os.path.isdir(p):
            fio_part = name.split(',', 1)[0].strip()
            res.append((p, fio_part))
    return res

def find_best_client_folder(fio: str, candidates: list, threshold: float = 90.0):
    best_path, best_score = None, -1.0
    for path, folder_fio in candidates:
        sc = similarity(fio, folder_fio)
        if sc > best_score:
            best_score, best_path = sc, path
    return (best_path if best_score >= threshold else None), best_score


# ===== ОСНОВНОЙ СЦЕНАРИЙ =====
def run(df_main):
    global LOG_SUMMARY
     # ✅ ДОБАВЬ ЭТИ СТРОКИ В НАЧАЛО run()
    print(f"\n[DEBUG] DEST_ROOT = {DEST_ROOT}")
    print(f"[DEBUG] Папка существует: {os.path.isdir(DEST_ROOT)}")
    
    client_dirs = list_client_folders(DEST_ROOT)
    print(f"[DEBUG] Найдено клиентских папок: {len(client_dirs)}")
    

    total_gos = 0
    found_gos = 0
    not_found_gos = []

    year_dir = latest_year_dir(SOURCE_ROOT)
    if not year_dir or not os.path.isdir(year_dir):
        print(f"❌ Не найдена папка года в '{SOURCE_ROOT}'")
        return

    day_dir = pick_latest_day_folder(year_dir)
    if not day_dir:
        print(f"❌ В '{year_dir}' нет папок с датами.")
        return

    out_dir = os.path.join(day_dir, "Готовые")
    os.makedirs(out_dir, exist_ok=True)

    print(f"📂 Год: {year_dir}")
    print(f"📂 Папка даты: {day_dir}")
    print(f"📁 Готовые: {out_dir}")
    print(f"📦 Папки клиентов: {DEST_ROOT}")

    client_dirs = list_client_folders(DEST_ROOT)
    print(f"👥 Найдено клиентских папок: {len(client_dirs)}")

    pdf_files = [f for f in os.listdir(day_dir)
                 if f.lower().endswith(".pdf") and os.path.isfile(os.path.join(day_dir, f))]
    if not pdf_files:
        print("⚠️ В выбранной папке PDF не найдены.")
        return

    for pdf_name in pdf_files:
        src_path = os.path.join(day_dir, pdf_name)
        try:
            reader = PdfReader(src_path)
        except Exception as e:
            print(f"⛔ Не удалось открыть '{pdf_name}': {e}")
            with open(log_file_path, 'a', encoding='utf-8') as log:
                log.write(f"[ГОСПОШЛИНА] Не удалось открыть файл '{pdf_name}': {e}\n")
            continue

        num_pages = len(reader.pages)
        print(f"\n===== Обработка: {pdf_name} | страниц: {num_pages} =====")

        for i in range(num_pages):
            try:
                page = reader.pages[i]
                page_text = page.extract_text() or ""

                fio = extract_fio(page_text)

                if fio:
                    total_gos += 1
                    file_name = f"Госпошлина, {fio}.pdf"
                else:
                    file_name = f"Госпошлина, page-{i+1:03d}.pdf"

                file_name = sanitize_filename(file_name)
                page_pdf_path = os.path.join(out_dir, file_name)

                writer = PdfWriter()
                writer.add_page(page)
                with open(page_pdf_path, "wb") as f:
                    writer.write(f)

                print(f"✅ Стр. {i+1}/{num_pages}: {file_name} — сохранён в 'Готовые'")

                if fio:
                    best_path, score = find_best_client_folder(fio, client_dirs, threshold=90.0)
                    if best_path:
                        dst_path = os.path.join(best_path, file_name)
                        try:
                            shutil.copy2(page_pdf_path, dst_path)
                            found_gos += 1
                            print(f"   ➤ 📤 Скопирован в клиентскую папку [{score:.0f}%]: {best_path}")
                        except Exception as e:
                            print(f"   ➤ ⛔ Ошибка копирования в '{best_path}': {e}")
                            not_found_gos.append(fio)
                            with open(log_file_path, 'a', encoding='utf-8') as log:
                                log.write(
                                    f"[ГОСПОШЛИНА] Ошибка копирования для ФИО: {fio}. "
                                    f"Файл: {pdf_name}, стр. {i+1}, папка: {best_path}, ошибка: {e}\n"
                                )
                    else:
                        print(f"   ➤ ⚠️ Папка клиента не найдена (score < 90) для ФИО: {fio}")
                        not_found_gos.append(fio)
                        with open(log_file_path, 'a', encoding='utf-8') as log:
                            log.write(
                                f"[ГОСПОШЛИНА] Не найдена клиентская папка для ФИО: {fio}. "
                                f"Файл: {pdf_name}, стр. {i+1}\n"
                            )
                else:
                    print("   ➤ ⚠️ ФИО не найдено — пропущено копирование в клиентскую папку.")
                    with open(log_file_path, 'a', encoding='utf-8') as log:
                        log.write(
                            f"[ГОСПОШЛИНА] Не удалось извлечь ФИО. "
                            f"Файл: {pdf_name}, стр. {i+1}\n"
                        )

            except Exception as e:
                print(f"⛔ Ошибка на странице {i+1}: {e}")
                with open(log_file_path, 'a', encoding='utf-8') as log:
                    log.write(
                        f"[ГОСПОШЛИНА] Ошибка обработки страницы. "
                        f"Файл: {pdf_name}, стр. {i+1}, ошибка: {e}\n"
                    )

    print("\n🎉 Готово.")
    print(f"Госпошлины: найдено {found_gos} из {total_gos} (страницы с распознанным ФИО).")
    safe_update_summary("ГОСПОШЛИНЫ", {
        "found": found_gos,
        "total": total_gos,
        "not_found": not_found_gos,
    })
   
    return found_gos, total_gos - found_gos