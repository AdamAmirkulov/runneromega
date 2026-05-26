# -*- coding: utf-8 -*-
import re
import sys
import argparse
from pathlib import Path

# Принудительно UTF-8 для вывода в консоль Windows
sys.stdout = open(sys.stdout.fileno(), mode="w", encoding="utf-8", buffering=1)
sys.stderr = open(sys.stderr.fileno(), mode="w", encoding="utf-8", buffering=1)

import fitz  # pip install --upgrade pymupdf

# ========= АРГУМЕНТЫ КОМАНДНОЙ СТРОКИ =========
parser = argparse.ArgumentParser()
parser.add_argument("--source_folder", default=None, help="Путь к папке с PDF-файлами")
parser.add_argument("--folder",        default=None, help="Альтернативное имя аргумента папки")
parser.add_argument("--workdir",       default=None, help="Рабочая директория задачи (передаётся runner'ом)")
args, _ = parser.parse_known_args()
# --- Утилиты ---
def normalize(text: str) -> str:
    text = text.replace("\r", "\n").replace("\f", "\n")
    text = re.sub(r"[ \t\xa0]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()

def clean_person(s: str) -> str:
    # убираем лишние слова, запятые и т.п.
    s = re.split(r"[,()]", s)[0]
    s = re.sub(r"[^\w \-ЁёӘІҢҮҰҚӨҺәіңүұқөһА-Яа-я]", " ", s)
    s = re.sub(r"\s{2,}", " ", s).strip()
    # нормализуем регистр
    def tc(w): return "-".join(p.capitalize() for p in w.split("-"))
    return " ".join(tc(t) for t in s.split())

def extract_from_contract(text: str):
    """Ищем ФИО и ИИН в блоке 'Заемщик:'"""
    T = normalize(text)

    # Проверка что это именно "Договор о предоставлении микрокредита"
    if not re.search(r"Договор о предоставлении\s+микрокредита", T, re.IGNORECASE):
        return None, None

    # Ищем блок "Заемщик:" ближе к концу документа
    m = re.search(r"Заемщик:(.+?)(?:\nМФО:|\Z)", T, re.IGNORECASE | re.DOTALL)
    if not m:
        return None, None
    block = m.group(1)

    # ФИО
    fio = None
    fio_m = re.search(r"ФИО\s+([^\n,]+)", block, re.IGNORECASE)
    if fio_m:
        fio = clean_person(fio_m.group(1))

    # ИИН
    iin = None
    iin_m = re.search(r"ИИН\s*(\d{12})", block)
    if iin_m:
        iin = iin_m.group(1)

    return fio, iin

def safe_rename(old: Path, new: Path) -> Path:
    if not new.exists():
        old.rename(new)
        return new
    k = 1
    while True:
        cand = new.with_stem(f"{new.stem} ({k})")
        if not cand.exists():
            old.rename(cand)
            return cand
        k += 1

def process_pdf(pdf_path: Path):
    try:
        with fitz.open(pdf_path) as doc:
            text = "".join(page.get_text() for page in doc)

        fio, iin = extract_from_contract(text)
        if fio and iin:
            new_name = f"Договор о предоставлении микрокредита, {fio}, {iin}.pdf"
            new_name = re.sub(r'[<>:"/\\|?*]', "", new_name)
            new_path = pdf_path.with_name(new_name)
            final_path = safe_rename(pdf_path, new_path)
            print(f"✅ {pdf_path.name} → {final_path.name}")
        else:
            print(f"⚠️ Пропуск: {pdf_path.name} (ФИО/ИИН не найдены)")

    except Exception as e:
        print(f"⛔ Ошибка в '{pdf_path.name}': {e}")

def main():
    base = Path(BASE_FOLDER)
    if not base.exists():
        print(f"Папка не найдена: {BASE_FOLDER}")
        return
    count = 0
    for pdf in base.rglob("*.pdf"):
        process_pdf(pdf)
        count += 1
    print(f"🎉 Готово. Обработано файлов: {count}")

if __name__ == "__main__":
    main()
