# -*- coding: utf-8 -*-
import os
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


# ========= ОПРЕДЕЛЯЕМ ПАПКУ =========
def find_folder() -> Path:
    if args.source_folder:
        return Path(args.source_folder)
    if args.folder:
        return Path(args.folder)
    if args.workdir:
        return Path(args.workdir)
    path = input("Введите путь к папке с PDF: ").strip().strip('"')
    return Path(path)


# ========= ИЗВЛЕЧЕНИЕ ФИО И ИИН =========
def extract_fio_iin(full_text: str):
    text = full_text.replace('\r', '\n')
    text_clean = re.sub(r'[ \t]+', ' ', text)

    # ---- ФИО ----
    fio = None

    fio_matches = list(re.finditer(
        r"(?:Аты.?жөні\s*/\s*ФИО|ТАӘ\s*/\s*ФИО)\s*[:\-–]?\s*([^\n]+)",
        text_clean,
        flags=re.IGNORECASE
    ))

    if fio_matches:
        fio_raw = fio_matches[-1].group(1).strip()
        fio_raw = re.split(
            r"\s*(?:ЖСН|ИИН|Дата|E[-–]mail|Туған|Телефон|Номер|№|\d{12})",
            fio_raw, flags=re.IGNORECASE
        )[0].strip()
        fio_raw = re.sub(r'\s{2,}', ' ', fio_raw).strip()

        def normalize_word(w):
            return w.capitalize() if w.isupper() else w

        if len(fio_raw.split()) >= 2:
            fio = " ".join(normalize_word(w) for w in fio_raw.split())

    # ---- ИИН ----
    iin = None
    iin_match = re.search(
        r"(?:ЖСН/ИИН|ЖСН|ИИН)\s*[:\-–]?\s*(\d[\d\s]{10,13})",
        text_clean,
        flags=re.IGNORECASE
    )
    if iin_match:
        iin_digits = re.sub(r"\D", "", iin_match.group(1))
        if len(iin_digits) == 12:
            iin = iin_digits

    return fio, iin


# ========= ОБРАБОТКА ПАПКИ =========
def rename_pdfs_in_folder(folder: Path):
    renamed = failed = skipped = 0

    for pdf_path in folder.rglob("*.pdf"):
        try:
            doc = fitz.open(pdf_path)
            full_text = "".join(page.get_text() for page in doc)
            doc.close()

            fio, iin = extract_fio_iin(full_text)

            print(f"  Файл: {pdf_path.name}")
            print(f"    ➤ ФИО: {fio}")
            print(f"    ➤ ИИН: {iin}")

            if fio and iin:
                new_name = f"Договор о предоставление микрокредита, {fio}, {iin}.pdf"
                safe_name = re.sub(r'[<>:"/\\|?*]', '', new_name)
                new_path = pdf_path.with_name(safe_name)

                if new_path != pdf_path:
                    pdf_path.rename(new_path)
                    print(f"    ✅ Переименовано -> {safe_name}\n")
                    renamed += 1
                else:
                    print(f"    ⏩ Уже в нужном формате\n")
                    skipped += 1
            else:
                print("    ⚠ Не смог извлечь ФИО или ИИН. Пропускаю.\n")
                skipped += 1

        except Exception as e:
            print(f"    ⛔ Ошибка при обработке {pdf_path.name}: {e}\n")
            failed += 1

    return renamed, skipped, failed


# ========= ОСНОВНАЯ ЛОГИКА =========
base = find_folder()

print(f"\n{'='*60}")
print(f"📁 Папка: {base}")
print(f"{'='*60}")

if not base.exists():
    print(f"❌ Папка не найдена: {base}")
    raise SystemExit(1)

renamed, skipped, failed = rename_pdfs_in_folder(base)

print(f"\n{'='*60}")
print(f"🏁 ИТОГ:")
print(f"   ✅ Переименовано: {renamed}")
print(f"   ⏩ Пропущено:     {skipped}")
print(f"   ⛔ Ошибок:        {failed}")
print(f"{'='*60}")