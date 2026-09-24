"""
Блок 4: Копирование оферт
"""
import os
import shutil
import pandas as pd
from config import ROOT, TARGET_BASE, MAIN_EXCEL, LOG_SUMMARY
from utils import (
    normalize,
    filename_contains_iin,
    ensure_client_folder,
    ensure_row_folder,
    safe_log,
    safe_update_summary
)

# ═══════════════════════════════════════════════════════════════
# НАСТРОЙКИ
# ═══════════════════════════════════════════════════════════════



# === ПУТИ ===


# === КОНФИГ ===
base_cessii   = rf"{ROOT}\Цессии"   # ИСТОЧНИК, ТОЛЬКО ЧТЕНИЕ
log_file_path = 'log_not_copied.txt'

# если сводный словарь ещё не создан в другом блоке – создадим


def normalize(text: str) -> str:
    """
    Нормализуем строку:
    - нижний регистр,
    - убираем пробелы / дефисы / подчёркивания,
    - приводим похожие кириллические буквы к латинице,
      чтобы 'ММ-1' и 'мм 1' считались одинаковыми.
    """
    trans = str.maketrans(
        "аеорсухкмвнтзм",    # кириллица, похожая на латиницу
        "aercsyxkmvntzm"     # латиница
    )
    return (
        str(text).lower()
        .translate(trans)
        .replace('-', '')
        .replace('_', '')
        .replace(' ', '')
    )

def build_client_folder(base_dir: str, fio: str, iin: str) -> str:
    """
    Создаём (если нет) и возвращаем путь к папке клиента:
    'ФИО, ИИН', с заменой запрещённых для Windows символов.
    """
    safe_name = f"{fio}, {iin}"
    for bad, good in [
        ('/', '_'), ('\\', '_'), (':', '_'), ('*', '_'),
        ('?', '_'), ('"', '_'), ('<', '_'), ('>', '_'), ('|', '_')
    ]:
        safe_name = safe_name.replace(bad, good)
    folder_path = os.path.join(base_dir, safe_name)
    os.makedirs(folder_path, exist_ok=True)
    return folder_path

def filename_matches_iin(file_name_no_ext: str, iin: str) -> bool:
    """
    Проверяем: встречается ли ИИН внутри имени файла (без расширения),
    если убрать пробелы, дефисы, подчёркивания.
    """
    candidate = (
        file_name_no_ext
        .replace(' ', '')
        .replace('-', '')
        .replace('_', '')
    )
    return iin in candidate

def run(df_main):
    count_success = 0
    count_failed  = 0
    found_offers = 0
    total_offers = len(df_main)
    not_found_offers = []
    print("\n=== ПОИСК ОФЕРТ В ЦЕССИЯХ (ЕДИНЫЕ ПРАВИЛА ДЛЯ ВСЕХ ПРОДУКТОВ, ТОЛЬКО ПО ИИН) ===\n")

    for _, row in df_main.iterrows():
        product = str(row['Product']).strip()
        fio     = str(row['FIO']).strip()
        iin     = str(row['IIN']).strip().zfill(12)

        # Папка клиента
        target_folder = ensure_row_folder(row, TARGET_BASE)

        copied = False
        found_product_match = False

        # 1. Ищем ПАПКУ ПРОДУКТА с полным совпадением
        for product_dir in os.listdir(base_cessii):
            product_path = os.path.join(base_cessii, product_dir)

            if not os.path.isdir(product_path):
                continue

            # Строго: совпадает ли продукт?
            if normalize(product) != normalize(product_dir):
                continue

            found_product_match = True
            print(f"[🔍] Продукт совпал 1-в-1: '{product}' == '{product_dir}'")

            # 2. Находим все подпапки, имя которых содержит 'оферт'
            offer_roots = []
            for root, dirs, files in os.walk(product_path):
                base = os.path.basename(root).lower()
                if 'оферт' in base:
                    offer_roots.append(root)

            # 3. Обходим каждую "офертную" папку и её подпапки, ищем pdf по ИИН
            def deep_walk(start_dir):
                for r, dnames, fnames in os.walk(start_dir):
                    yield r, fnames

            for offer_root in offer_roots:
                for r, fnames in deep_walk(offer_root):
                    for file in fnames:
                        if not file.lower().endswith('.pdf'):
                            continue

                        file_no_ext = os.path.splitext(file)[0]

                        if filename_matches_iin(file_no_ext, iin):
                            src = os.path.join(r, file)

                            # Формат итогового файла в папке клиента:
                            new_filename = f"Оферта, {fio}, {iin}.pdf"
                            dst = os.path.join(target_folder, new_filename)

                            shutil.copy2(src, dst)
                            copied = True
                            count_success += 1
                            print(f"[✔] Найдено по ИИН: {file} → {new_filename} в {target_folder}")
                            break
                    if copied:
                        break
                if copied:
                    break

            # Если нашли и скопировали — нет смысла смотреть другие product_dir
            if copied:
                break

        # 4. Логирование результатов по клиенту
        if not found_product_match:
            print(f"[⚠] Нет папки продукта '{product}' для {fio} ({iin})")

        if not copied:
            count_failed += 1
            # добавляем в список для сводного лога
            not_found_offers.append(f"{fio}, {iin}")

            with open(log_file_path, 'a', encoding='utf-8') as log_file:
                log_file.write(
                    f"[ОФЕРТА] Не найдена оферта по ИИН для: {fio}, {iin}, продукт: {product}\n"
                )
        else:
            # оферта найдена — учитываем в сводном логе
            found_offers += 1

    print("\n--- ИТОГ ОФЕРТ ---")
    print(f"✔️ Скопировано: {count_success}")
    print(f"❌ Не найдено (в лог): {count_failed}")
    print(f"📄 Лог: {os.path.abspath(log_file_path)}")

    # сохраняем результат блока «ОФЕРТЫ» в общий сводный лог
    LOG_SUMMARY["ОФЕРТЫ"] = {
        "found": found_offers,
        "total": total_offers,
        "not_found": not_found_offers,
    }
    return count_success, count_failed
