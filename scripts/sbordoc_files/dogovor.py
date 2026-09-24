"""
Блок 6: Копирование договоров цессии
"""
import os
import shutil

from config import ROOT, TARGET_BASE, BASE_CESSII
from utils import (
    normalize,
    filename_contains_iin,
    ensure_client_folder,
    ensure_row_folder,
    safe_log,
    safe_update_summary
)
import os
import shutil
import pandas as pd

# === ПУТИ ===

base_cessii   = BASE_CESSII  # ТОЛЬКО ЧТЕНИЕ
log_file_path = 'log_not_copied.txt'

# если сводный словарь ещё не создан – создаём


# === ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ===

# === ЧИТАЕМ EXCEL ===

def run(df_main):
    print("\n=== БЛОК 4: Договор цессии по продукту (файл с 'обезлич...') ===\n")

    # детальные счётчики
    count_success_cession = 0
    count_failed_cession  = 0

    # счётчики для сводного отчёта
    total_cession      = len(df_main)
    found_cession      = 0
    not_found_cession  = []

    for _, row in df_main.iterrows():
        product = str(row['Product']).strip()
        fio     = str(row['FIO']).strip()
        iin     = str(row['IIN']).strip().zfill(12)

        # Папка клиента
        target_folder = ensure_row_folder(row, TARGET_BASE)

        copied = False
        found_product_match = False

        # Перебираем продуктовые папки в Цессии
        for product_dir in os.listdir(base_cessii):
            product_path = os.path.join(base_cessii, product_dir)

            # оставляем только папки
            if not os.path.isdir(product_path):
                continue

            # продукт должен совпасть 100% (после normalize)
            if normalize(product) != normalize(product_dir):
                continue

            found_product_match = True
            print(f"[ЦЕССИЯ-ДОГОВОР] Продукт совпал: '{product}' == '{product_dir}'")

            # Ищем договор цессии: PDF, в имени которого есть 'обезлич' (в любом регистре)
            for root, dirs, files in os.walk(product_path):
                for file in files:
                    if not file.lower().endswith('.pdf'):
                        continue

                    filename_lower = file.lower()

                    # условие: имя файла содержит корень "обезлич"
                    if 'обезлич' in filename_lower:
                        src = os.path.join(root, file)

                        # целевое имя в папке клиента:
                        # "Договор цессии, <наименование продукта>.pdf"
                        new_filename = f"Договор цессии, {product}.pdf"
                        dst = os.path.join(target_folder, new_filename)

                        shutil.copy2(src, dst)
                        copied = True
                        count_success_cession += 1
                        found_cession += 1
                        print(f"[✔ ДОГОВОР] {file} → '{new_filename}' в '{target_folder}'")
                        break
                if copied:
                    break
            if copied:
                break

        if not found_product_match:
            print(f"[⚠ ДОГОВОР] Не найдена папка продукта '{product}' для {fio} ({iin})")

        if not copied:
            count_failed_cession += 1
            not_found_cession.append(f"{fio}, {iin}")  # для сводного лога

            # пишем в лог, но не затираем существующее содержимое
            with open(log_file_path, 'a', encoding='utf-8') as log_file:
                log_file.write(
                    f"[ДОГОВОР ЦЕССИИ] Не найден файл с 'обезлич...' для продукта '{product}' "
                    f"у клиента {fio}, {iin}\n"
                )

    print("\n--- ИТОГ ДОГОВОРОВ ЦЕССИИ ---")
    print(f"✔️ Скопировано договоров цессии: {count_success_cession}")
    print(f"❌ Не найдено договоров цессии: {count_failed_cession}")
    print(f"📄 Лог: {os.path.abspath(log_file_path)}")

    # сохраняем результат этого блока в общий сводный лог
    safe_update_summary("ДОГОВОР ЦЕССИИ", {
        "found": found_cession,
        "total": total_cession,
        "not_found": not_found_cession,
    })
    return count_success_cession, count_failed_cession
