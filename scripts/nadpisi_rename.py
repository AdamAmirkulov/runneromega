import os
import re
import fitz  # pymupdf
import sys
import io
import argparse

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# ═══════════════════════════════════════════════════════════════
# ОБРАБОТКА АРГУМЕНТОВ
# ═══════════════════════════════════════════════════════════════

def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('--workdir', type=str, default=None)
    parser.add_argument('--source_folder', type=str, default=None)
    return parser.parse_args()

args = parse_args()

# ✅ Если передан --source_folder — используем его, иначе fallback на список
if args.source_folder:
    base_folders = [args.source_folder.strip()]
else:
    base_folders = [
        r"\\KAMILLA\work folder\Надписи\Bereke Bank",
        # добавляйте сколько угодно папок...
    ]

print("📂 Папки для обработки:")
for f in base_folders:
    print(f"   {f}")


def extract_fio_iin(text):
    match = re.search(r'с\s+([А-ЯЁӘІҢҮҰҚӨҺҒ\s-]{5,}),\s+\d{2}\.\d{2}\.\d{4}г\.р\.,\s+ИИН\s+(\d{12})', text, re.IGNORECASE)
    if match:
        fio = match.group(1).strip().replace('\n', ' ')
        iin = match.group(2).strip()
        return fio, iin
    return None, None

def process_folder(base_folder):
    print(f"\n{'='*60}")
    print(f"📁 Обработка папки: {base_folder}")
    print(f"{'='*60}")

    renamed = 0
    failed = 0
    skipped = 0

    for root, dirs, files in os.walk(base_folder):
        for file in files:
            if file.lower().endswith('.pdf'):
                file_path = os.path.join(root, file)
                try:
                    doc = fitz.open(file_path)
                    text = doc[0].get_text()
                    doc.close()

                    fio, iin = extract_fio_iin(text)
                    if fio and iin:
                        new_name = f"Исполнительная надпись, {fio}, {iin}.pdf"
                        new_name_clean = re.sub(r'[<>:"/\\|?*]', '', new_name)
                        new_path = os.path.join(root, new_name_clean)

                        try:
                            os.rename(file_path, new_path)
                            print(f"  ✅ Переименовано: {new_name_clean}")
                            renamed += 1
                        except Exception as rename_error:
                            print(f"  ⛔ Ошибка при переименовании: {file}")
                            print(f"     ➤ Новое имя: {new_name_clean}")
                            print(f"     ➤ Причина: {rename_error}")
                            failed += 1
                    else:
                        print(f"  ⚠ Не удалось извлечь ФИО/ИИН: {file}")
                        skipped += 1
                except Exception as e:
                    print(f"  ⛔ Ошибка при обработке {file}: {e}")
                    failed += 1

    print(f"\n  📊 Итог: ✅ {renamed} переименовано | ⚠ {skipped} пропущено | ⛔ {failed} ошибок")
    return renamed, skipped, failed

# === Запуск по всем папкам ===
total_renamed = total_skipped = total_failed = 0

for folder in base_folders:
    if os.path.exists(folder):
        r, s, f = process_folder(folder)
        total_renamed += r
        total_skipped += s
        total_failed += f
    else:
        print(f"\n❌ Папка не найдена: {folder}")

print(f"\n{'='*60}")
print(f"🏁 ОБЩИЙ ИТОГ по всем папкам:")
print(f"   ✅ Переименовано: {total_renamed}")
print(f"   ⚠  Пропущено:    {total_skipped}")
print(f"   ⛔ Ошибок:       {total_failed}")
print(f"{'='*60}")