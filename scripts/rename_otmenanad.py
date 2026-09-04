import os
import re
import argparse

# Папка с PDF-файлами: приоритет — аргумент из веб-формы (--source_folder),
# затем рабочая папка задачи (--workdir), затем локальный дефолт.
parser = argparse.ArgumentParser()
parser.add_argument("--source_folder", default=None)
parser.add_argument("--workdir", default=None)
args, _ = parser.parse_known_args()

base_folder = args.source_folder or args.workdir or r'C:\Users\User\Desktop\Отмены'

# Обход всех PDF-файлов в папке (без подпапок)
for root, dirs, files in os.walk(base_folder):
    for file in files:
        if file.lower().endswith('.pdf'):
            old_path = os.path.join(root, file)

            # Удаляем расширение и лишние пробелы
            file_name_only = os.path.splitext(file)[0].strip()

            # Новое имя с префиксом
            new_name = f"Постановление об отмене надписи, {file_name_only}.pdf"
            new_name_clean = re.sub(r'[<>:"/\\|?*]', '', new_name)
            new_path = os.path.join(root, new_name_clean)

            # Пропускаем, если уже переименовано
            if os.path.abspath(old_path) == os.path.abspath(new_path):
                continue

            try:
                os.rename(old_path, new_path)
                print(f"✅ Переименовано: {new_name_clean}")
            except Exception as e:
                print(f"⛔ Ошибка при переименовании {file}: {e}")

print("\n🎉 Готово! Все файлы обработаны.")
