# -*- coding: utf-8 -*-
"""
Блок 11: Формирование досудебных претензий из шаблона Word
"""
import os
from pathlib import Path
from datetime import datetime, date
import pythoncom
from openpyxl import load_workbook
from docx import Document
  # pip install docx2pdf

from config import MAIN_EXCEL, ROOT, TARGET_BASE,LOG_SUMMARY
from utils import ensure_client_folder, safe_log, safe_update_summary

# ═══════════════════════════════════════════════════════════════
# НАСТРОЙКИ
# ═══════════════════════════════════════════════════════════════

SHEET_NAME = "Данные для шаблонов"

TEMPLATE_DOC = rf"{ROOT}\Документы для подачи Исков\Шаблоны документов\Шаблоны по Досудебной претензии.docx"
# -*- coding: utf-8 -*-
import os
from pathlib import Path
from datetime import datetime, date

from openpyxl import load_workbook
from docx import Document
 # pip install docx2pdf

# ---------------- НАСТРОЙКИ ----------------


# Корневая папка, где лежат подпапки по клиентам

log_file_path = 'log_not_copied.txt'
# ---------- ЛОГИ (как в первом скрипте) ----------
# общий лог по всем блокам (если не задан в верхней ячейке — задаём здесь)
try:
    log_file_path
except NameError:
    os.makedirs(TARGET_BASE, exist_ok=True)
    log_file_path = os.path.join(TARGET_BASE, "лог_сбор_документов.txt")

# общий словарь сводки
try:
    LOG_SUMMARY
except NameError:
    LOG_SUMMARY = {}

# счётчики конкретно для досудебных претензий (по шаблону)
  # список "ФИО, ИИН" для не сформированных / с ошибками


# -------------------------------------------------
# Вспомогательные функции
# -------------------------------------------------



import threading
import pythoncom
import win32com.client

# Глобальный Lock чтобы Word не вызывался параллельно
_word_lock = threading.Lock()

def convert_to_pdf(docx_path: str, pdf_path: str):
    """Потокобезопасная конвертация docx → pdf через Word COM."""
    with _word_lock:
        pythoncom.CoInitialize()
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            try:
                doc = word.Documents.Open(docx_path)
                doc.SaveAs(pdf_path, FileFormat=17)  # 17 = wdFormatPDF
                doc.Close()
            finally:
                word.Quit()
        finally:
            pythoncom.CoUninitialize()

def nice_case(s: str) -> str:
    """Первая буква каждого слова заглавная, остальные строчные."""
    if not s:
        return ""
    parts = str(s).split()
    return " ".join(p.capitalize() for p in parts)


def replace_placeholders_in_paragraph(paragraph, mapping: dict):
    """
    Работаем с целым paragraph.text, затем пересобираем абзац так,
    чтобы плейсхолдеры, разрезанные на несколько run'ов, тоже заменялись.
    При этом формат (шрифт, размер) берём из ПЕРВОГО run абзаца.

    Для абзацев, где были {Email} или {Дата минус 7 дней}, снимаем жирность.
    """
    original_text = paragraph.text
    text = original_text
    changed = False

    for key, value in mapping.items():
        if key in text:
            text = text.replace(key, value)
            changed = True

    if not changed:
        return  # ничего не меняем

    if paragraph.runs:
        # сохраняем формат первого run (шрифт, размер и т.д.)
        paragraph.runs[0].text = text
        for r in paragraph.runs[1:]:
            r.text = ""
        target_run = paragraph.runs[0]
    else:
        target_run = paragraph.add_run(text)

    # для абзацев с Email и Датой минус 7 дней делаем обычный шрифт (не жирный)
    lower_orig = original_text.lower()
    if "{email}" in lower_orig or "{дата минус 7 дней}" in lower_orig:
        target_run.bold = False


def fill_document(row_dict: dict) -> Document:
    doc = Document(TEMPLATE_DOC)

    mapping = {}

    for col_name, val in row_dict.items():
        if not col_name:  # пустой заголовок
            continue

        placeholder = "{" + col_name + "}"
        lower_name = col_name.lower()

        # --- формируем строку значения ---
        if val is None or val != val:
            value_str = ""
        else:
            # 1) если это реальный datetime/date → всегда дд.мм.гггг
            if isinstance(val, (datetime, date)):
                value_str = val.strftime("%d.%m.%Y")
            else:
                # приводим к строке
                s = str(val).strip()

                # 2) если в названии колонки есть "дата" → обрезаем время
                if "дата" in lower_name:
                    # отрезаем то, что после пробела (время)
                    # 2025-12-04 09:41:55 → 2025-12-04
                    if " " in s:
                        s = s.split()[0]

                    # 2025-12-04 → 04.12.2025
                    parts = s.replace("/", "-").split("-")
                    if len(parts) == 3 and len(parts[0]) == 4:  # yyyy-mm-dd
                        y, m, d = parts
                        s = f"{d.zfill(2)}.{m.zfill(2)}.{y}"

                    value_str = s

                # 3) ФИО/ЧСИ → красивый регистр
                elif "фио" in lower_name or "чси" in lower_name:
                    value_str = nice_case(s)
                else:
                    value_str = s

        mapping[placeholder] = value_str

    # Дополнительно текущая дата
    mapping["{CurrDate}"] = datetime.now().strftime("%d.%m.%Y")

    # дальше всё как было...
    for p in doc.paragraphs:
        replace_placeholders_in_paragraph(p, mapping)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_placeholders_in_paragraph(p, mapping)

    return doc


def find_debtor_folder(root_folder: Path, iin: str) -> Path:
    """
    Ищем подпапку в root_folder, в имени которой встречается ИИН.
    Если не нашли — создаём новую папку с именем ИИН.
    """
    for sub in root_folder.iterdir():
        if sub.is_dir() and iin in sub.name:
            return sub

    new_folder = root_folder / iin
    new_folder.mkdir(parents=True, exist_ok=True)
    return new_folder


# -------------------------------------------------
# Основная логика
# -------------------------------------------------
def run(main_df):
    pythoncom.CoInitialize()  # ← инициализация COM для текущего потока
    try:
        global pret_found, pret_total, pret_not_found, LOG_SUMMARY
        pret_found = 0        # сколько претензий удалось сформировать
        pret_total = 0        # сколько записей обработали из Excel
        pret_not_found = [] 

        root_path = Path(TARGET_BASE)
        root_path.mkdir(parents=True, exist_ok=True)

        wb = load_workbook(MAIN_EXCEL, read_only=True, data_only=True)
        ws = wb[SHEET_NAME]

        # Заголовки колонок (первая строка)
        headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

        # Обход строк, начиная со 2-й
        for row in ws.iter_rows(min_row=2, values_only=True):
            if all(v is None for v in row):
                continue

            row_dict = dict(zip(headers, row))

            # ФИО - столбец C (индекс 2), ИИН - столбец D (индекс 3)
            fio_value = row[2] if len(row) > 2 else None
            iin_value = row[3] if len(row) > 3 else None

            fio = nice_case(fio_value)
            iin = str(iin_value).strip() if iin_value is not None else ""

            if not fio or not iin:
                continue

            pret_total += 1
            print("\n=======================================")
            print(f"[STEP DOCX] {pret_total} — {fio}, {iin}")

            try:
                # Обновляем ФИО/ИИН в словаре, чтобы плейсхолдеры с любыми названиями колонок,
                # содержащими "фио" / "иин", брали уже нормализованные значения
                for key in list(row_dict.keys()):
                    if key and "фио" in key.lower():
                        row_dict[key] = fio
                    if key and "иин" in key.lower():
                        row_dict[key] = iin

                # 1) Собираем docx по шаблону
                doc = fill_document(row_dict)

                # Находим/создаём папку должника по ИИН
                debtor_folder = find_debtor_folder(root_path, iin)

                base_name = f"Досудебная претензия, {fio}, {iin}"
                docx_path = debtor_folder / (base_name + ".docx")
                pdf_path  = debtor_folder / (base_name + ".pdf")

                doc.save(docx_path)

                # 2) Конвертируем в PDF
                convert_to_pdf(str(docx_path), str(pdf_path))

                # 3) Удаляем промежуточный DOCX, если не нужен
                try:
                    os.remove(docx_path)
                except OSError:
                    pass

                print(f"Сформирован файл: {pdf_path}")
                pret_found += 1

            except Exception as e:
                label = f"{fio or 'Неизвестный'}, {iin or 'ИИН не указан'}"
                pret_not_found.append(label)
                print(f"[ERROR DOCX] Ошибка при формировании для {label}: {e}")
                try:
                    with open(log_file_path, "a", encoding="utf-8") as lf:
                        lf.write(f"[ДОСУДЕБНАЯ ПРЕТЕНЗИЯ DOCX] Ошибка для {label}: {e}\n")
                except Exception:
                    pass
                continue

        wb.close()

        print("\n=== ГОТОВО: все записи обработаны (DOCX) ===")
        print(f"[ИТОГ ДОСУДЕБНЫЕ ПРЕТЕНЗИИ DOCX] сформировано {pret_found} из {pret_total}")

        # записываем в общий сводный словарь
        LOG_SUMMARY["ДОСУДЕБНЫЕ ПРЕТЕНЗИИ DOCX"] = {
            "found": pret_found,
            "total": pret_total,
            "not_found": pret_not_found,
        }

        print("Готово.")
        return pret_found, pret_total

    finally:
        pythoncom.CoUninitialize()  # освобождаем COM по завершении

    