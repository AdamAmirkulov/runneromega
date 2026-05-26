# -*- coding: utf-8 -*-
# -*- coding: utf-8 -*-
import re
import sys
import argparse
import fitz  # PyMuPDF
from pathlib import Path

# Принудительно UTF-8 для вывода в консоль Windows
sys.stdout = open(sys.stdout.fileno(), mode="w", encoding="utf-8", buffering=1)
sys.stderr = open(sys.stderr.fileno(), mode="w", encoding="utf-8", buffering=1)

# ========= АРГУМЕНТЫ КОМАНДНОЙ СТРОКИ =========
parser = argparse.ArgumentParser()
parser.add_argument("--source_folder", default=None, help="Путь к папке с PDF-файлами")
parser.add_argument("--folder",        default=None, help="Альтернативное имя аргумента папки")
parser.add_argument("--workdir",       default=None, help="Рабочая директория задачи (передаётся runner'ом)")
args, _ = parser.parse_known_args()

def find_folder() -> Path:
    if args.source_folder:
        return Path(args.source_folder)
    if args.folder:
        return Path(args.folder)
    if args.workdir:
        return Path(args.workdir)
    path = input("Введите путь к папке с PDF: ").strip().strip('"')
    return Path(path)

# --- Пробелы/переносы ---
NBSP = "\u00A0"                                                      # неразрывный пробел
SPACES_CLASS = r"[ \t" + NBSP + r"]+"                                # как и раньше
WSP = rf"(?:[ \t{NBSP}]|\r?\n)+"                                     # пробел/таб/nbps/перенос(ы)

# --- Служебные слова/обращения для чистки ФИО ---
HONORIFICS = r"(?:г\-ну|г\-же|господину|госпоже|мырзаға|мырза|ханым|құрметті)"
SERVICE_WORDS = r"(?:руководителю|мфо|жшс|жкпс|товарищество|электронн|email|адрес|gmail|yandex|mail\.ru)"

# --- Заголовки/триггеры для распознавания документа ---
TITLE_BASE_RE = re.compile(
    rf"\bЗаявлени[ея]{WSP}на{WSP}(?:получени[ея]|предоставлени[ея]){WSP}микрокредит[ауы]\b",
    re.IGNORECASE
)
TITLE_GET_ONLY_RE = re.compile(rf"\bЗаявлени[ея]{WSP}на{WSP}получени[ея]\b", re.IGNORECASE)
TITLE_PROVIDE_RE = re.compile(rf"\bЗаявлени[ея]{WSP}на{WSP}предоставлени[ея]{WSP}микрокредит[ауы]\b", re.IGNORECASE)
MICROCREDIT_WORD_RE = re.compile(rf"\bмикрокредит[ауы]\b", re.IGNORECASE)

def contains_target_title(text: str) -> bool:
    T = normalize(text)
    return bool(
        TITLE_BASE_RE.search(T)
        or TITLE_GET_ONLY_RE.search(T)
        or TITLE_PROVIDE_RE.search(T)
        or MICROCREDIT_WORD_RE.search(T)
    )

# --- ИИН/ЖСН: «рваные» цифры 5..40, нормализуем до 12 ---
IIN_CHUNK_RE = re.compile(rf"((?:\d[ \t{NBSP}\-]?){5,40})")

# --- Метки полей (строго по метке) ---
FIO_LABEL_PTRN = re.compile(
    r"(Аты[\-\s]?жөні|ТАӘ|ФИО)(?:\s*/\s*(?:ФИО|Аты[\-\s]?жөні|ТАӘ))?\s*[:\-]?\s*",
    re.IGNORECASE
)
IIN_LABEL_PTRN = re.compile(
    r"(ЖСН|ИИН)(?:\s*/\s*(?:ИИН|ЖСН))?\s*[:\-]?\s*",
    re.IGNORECASE
)

# -------------------- Нормализация --------------------
def normalize(text: str) -> str:
    text = text.replace("\r", "\n").replace("\f", "\n")
    text = re.sub(SPACES_CLASS, " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()

def collapse_spaced_letters(s: str) -> str:
    """ 'Т о х т а х у н о в а' -> 'Тохтахунова' """
    def _fix_word(w: str) -> str:
        if len(w) >= 3 and " " in w.strip():
            parts = w.split()
            if all(len(p) == 1 or (len(p) == 2 and p.endswith("-")) for p in parts):
                return "".join(p.replace("-", "") for p in parts)
        return w
    s = re.sub(SPACES_CLASS, " ", s).strip()
    chunks = re.split(r"([,.;()])", s)
    for i in range(0, len(chunks), 2):
        chunks[i] = " ".join(_fix_word(w) for w in chunks[i].split())
    return "".join(chunks)

def titlecase_keep_hyphen(word: str) -> str:
    return "-".join(p.capitalize() for p in word.split("-") if p)

def clean_person(s: str) -> str:
    s = collapse_spaced_letters(s)
    s = re.sub(fr"\b{HONORIFICS}\b", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"[^\w \-’ʼ'`ӘІҢҮҰҚӨҺәіңүұқөһЁёА-Яа-я]", " ", s)
    s = re.sub(SPACES_CLASS, " ", s).strip()
    toks = [t for t in s.split() if t.upper() not in {"ИИН", "ЖСН", "Ж", "С", "Н"}]
    return " ".join(titlecase_keep_hyphen(t) for t in toks)

# -------------------- Получение текста --------------------
def get_right_column_text(doc: fitz.Document) -> str:
    right_text_parts = []
    for page in doc:
        page_w = page.rect.width
        tol = 5
        blocks = page.get_text("blocks")
        blocks = sorted(blocks, key=lambda b: (round(b[1], 1), round(b[0], 1)))
        for (x0, y0, x1, y1, text, *_rest) in blocks:
            if x0 >= page_w / 2 - tol:
                right_text_parts.append(text)
    return "\n".join(right_text_parts)

def get_full_text(doc: fitz.Document) -> str:
    return "\n".join(page.get_text("text") for page in doc)

# -------------------- Извлечение ПО МЕТКАМ --------------------
def extract_fio_after_label(text: str) -> str | None:
    """
    Строго берём ФИО после метки «Аты-жөні/ФИО …».
    Поддержка: перенос ФИО на следующую строку; отсечение хвостов ';', '(' и пр.
    """
    lines = normalize(text).split("\n")
    for i, line in enumerate(lines):
        m = FIO_LABEL_PTRN.search(line)
        if not m:
            continue
        rest = line[m.end():].strip()
        # Если на этой строке пусто/слишком коротко — захватим следующую строку
        if len(rest.split()) < 2 and i + 1 < len(lines):
            rest = (rest + " " + lines[i + 1].strip()).strip()

        # Отсекаем комментарии/подписи
        rest = rest.split("(")[0].split(";")[0].strip()
        # Убираем хвосты с цифрами, если случайно попали
        rest = re.sub(r"\s+\d{2,}\s*$", "", rest).strip()

        fio = clean_person(rest)
        # Минимальная проверка: не меньше 2 слов, без цифр
        if fio and len(fio.split()) >= 2 and not re.search(r"\d", fio):
            return fio
    return None

def extract_iin_after_label(text: str) -> str | None:
    """
    Строго берём ИИН после метки «ЖСН/ИИН …».
    Разрешаем «рваные» цифры и переносы; нормализуем до 12 цифр.
    """
    lines = normalize(text).split("\n")
    for i, line in enumerate(lines):
        m = IIN_LABEL_PTRN.search(line)
        if not m:
            continue
        # Ищем цифры в остатке строки + 2 строки вниз (на случай переноса)
        window = line[m.end():]
        if i + 1 < len(lines):
            window += " " + lines[i + 1]
        if i + 2 < len(lines):
            window += " " + lines[i + 2]
        m2 = IIN_CHUNK_RE.search(window)
        if m2:
            digits = re.sub(r"\D+", "", m2.group(1))
            if digits:
                return digits.zfill(12)[:12]
    return None

# -------------------- (опциональные) Fallback-правила --------------------
# Используем только если строго по меткам ничего не нашли.
FIO_KARYZ_ZAEM_RE = re.compile(
    rf"(?:Қарыз{WSP}алушы|Заемщик)(?:{WSP}\/{WSP}(?:Заемщик|Қарыз{WSP}алушы))?"
    rf"{WSP}[:\-]?\s*([A-Za-zА-Яа-яЁёӘІҢҮҰҚӨҺәіңүұқөһ\-\s’ʼ'`]+)",
    re.IGNORECASE
)
FIO_BEFORE_IIN_RE = re.compile(
    rf"([A-ZА-ЯЁӘІҢҮҰҚӨҺ][A-Za-zА-Яа-яЁёӘІҢҮҰҚӨҺ’ʼ'`\-\s]{{5,80}}?)"
    rf"{WSP},?\s*(?:ИИН|ЖСН)\s*[:\-]?\s*((?:\d[ \t{NBSP}\-]?)+)",
    re.IGNORECASE
)

def extract_fallbacks(text: str, iin_hint: str | None) -> tuple[str | None, str | None]:
    T = normalize(text)

    m = FIO_BEFORE_IIN_RE.search(T)
    if m:
        fio = clean_person(m.group(1))
        iin = re.sub(r"\D+", "", m.group(2)).zfill(12)[:12]
        if fio and iin:
            return fio, iin

    m = FIO_KARYZ_ZAEM_RE.search(T)
    if m and iin_hint:
        return clean_person(m.group(1)), iin_hint

    return None, iin_hint

# -------------------- Переименование --------------------
def safe_rename(old: Path, new: Path) -> Path:
    if not new.exists():
        old.rename(new)
        return new
    k = 1
    while True:
        cand = new.with_name(f"{new.stem} ({k}){new.suffix}")
        if not cand.exists():
            old.rename(cand)
            return cand
        k += 1

# -------------------- Обработка одного PDF --------------------
def process_pdf(pdf_path: Path):
    try:
        with fitz.open(pdf_path) as doc:
            right_text = get_right_column_text(doc)
            full_text = get_full_text(doc)

        # Заголовок: и справа, и по всему документу
        has_title = contains_target_title(right_text) or contains_target_title(full_text)
        if not has_title:
            print(f"• Пропуск: {pdf_path.name} (нет заголовка заявления)")
            return

        # --- Приоритет: строго по меткам ---
        iin = extract_iin_after_label(full_text)
        fio = extract_fio_after_label(full_text)

        # --- Мягкий fallback (если метки не сработали) ---
        if (not fio or not iin):
            fb_fio, fb_iin = extract_fallbacks(full_text, iin)
            fio = fio or fb_fio
            iin = iin or fb_iin

        print(f"📄 {pdf_path.name} → ФИО: {fio or '—'} | ИИН: {iin or '—'}")

        if not iin or not fio:
            print("⚠️ Не хватает реквизитов (ИИН или ФИО) — файл не переименован.\n")
            return

        new_name = f"Договор о предоставление микрокредита, {fio}, {iin}.pdf"
        new_name = re.sub(r'[<>:"/\\|?*]', "", new_name)
        final_path = safe_rename(pdf_path, pdf_path.with_name(new_name))
        print(f"✅ Переименовано: {final_path.name}\n")

    except Exception as e:
        print(f"⛔ Ошибка в '{pdf_path.name}': {e}\n")

# -------------------- Точка входа --------------------
def main():
    base = find_folder()
    if not base.exists():
        print(f"Папка не найдена: {base}")
        return
    total = renamed = 0
    for pdf in base.rglob("*.pdf"):
        total += 1
        existed = pdf.exists()
        process_pdf(pdf)
        if existed and not pdf.exists():
            renamed += 1
    print(f"🎉 Готово. Найдено PDF: {total}. Переименовано: {renamed}.")

if __name__ == "__main__":
    main()
