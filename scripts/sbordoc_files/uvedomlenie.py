"""
Блок 3: Копирование уведомлений об уступке права требования из EMAIL
"""
import os
import shutil
import pandas as pd

from config import (
    MAIN_EXCEL,
    TARGET_BASE,
    EMAIL_SOURCE,
    LOG_FILE
)

from utils import (
    normalize,
    filename_contains_iin,
    ensure_client_folder,
    safe_log,
    safe_update_summary
)

# ═══════════════════════════════════════════════════════════════
# ОСНОВНАЯ ФУНКЦИЯ
# ═══════════════════════════════════════════════════════════════

def run(df_main):
    """
    Копирует уведомления об уступке права требования из папки EMAIL.
    
    Логика:
    1. Для каждого клиента находит соответствующую папку продукта в EMAIL_SOURCE
    2. Ищет PDF файл, содержащий ИИН клиента
    3. Копирует найденный файл в папку клиента
    
    Args:
        df_main: DataFrame с колонками ['Product', 'FIO', 'IIN']
    
    Returns:
        tuple: (количество успешно скопированных, количество не найденных)
    """
    print("\n" + "="*70)
    print("  📧 БЛОК 3: УВЕДОМЛЕНИЯ ОБ УСТУПКЕ ПРАВА (EMAIL)")
    print("="*70 + "\n")

    # Счётчики
    count_success = 0
    count_failed = 0
    total = len(df_main)
    not_found_list = []

    # Проверка существования источника
    if not os.path.exists(EMAIL_SOURCE):
        print(f"❌ ОШИБКА: Папка EMAIL_SOURCE не найдена: {EMAIL_SOURCE}")
        safe_log(f"[УВЕДОМЛЕНИЯ] EMAIL_SOURCE не найден: {EMAIL_SOURCE}")
        return 0, total

    # Обработка каждого клиента
    for idx, row in df_main.iterrows():
        product = str(row['Product']).strip()
        fio = str(row['FIO']).strip()
        iin = str(row['IIN']).strip().zfill(12)

        # Создаём папку клиента
        target_folder = ensure_client_folder(iin, fio, TARGET_BASE)

        copied = False
        found_product = False

        # Поиск папки продукта в EMAIL_SOURCE
        for product_dir in os.listdir(EMAIL_SOURCE):
            product_path = os.path.join(EMAIL_SOURCE, product_dir)

            # Пропускаем файлы
            if not os.path.isdir(product_path):
                continue

            # Сравниваем нормализованные названия продуктов
            if normalize(product) != normalize(product_dir):
                continue

            found_product = True
            print(f"[{idx+1}/{total}] Продукт найден: '{product}' → '{product_dir}'")

            # Рекурсивный поиск PDF с ИИН
            for root, dirs, files in os.walk(product_path):
                for file in files:
                    if not file.lower().endswith('.pdf'):
                        continue

                    file_no_ext = os.path.splitext(file)[0]

                    # Проверяем наличие ИИН в имени файла
                    if filename_contains_iin(file_no_ext, iin):
                        src = os.path.join(root, file)
                        dst = os.path.join(target_folder, file)

                        try:
                            shutil.copy2(src, dst)
                            copied = True
                            count_success += 1
                            print(f"   ✅ Скопирован: {file}")
                        except Exception as e:
                            print(f"   ⚠️  Ошибка копирования: {e}")
                            safe_log(f"[УВЕДОМЛЕНИЯ] Ошибка копирования для {fio} ({iin}): {e}")
                        
                        break  # Нашли файл для этого клиента

                if copied:
                    break  # Выходим из os.walk

            if copied:
                break  # Выходим из цикла по папкам продуктов

        # Если продукт не найден
        if not found_product:
            print(f"[{idx+1}/{total}] ⚠️  Папка продукта '{product}' не найдена для {fio} ({iin})")

        # Если файл не скопирован
        if not copied:
            count_failed += 1
            not_found_list.append(f"{fio}, {iin}")
            safe_log(f"[УВЕДОМЛЕНИЯ] Не найден PDF для: {fio}, {iin}, продукт: {product}")
            print(f"[{idx+1}/{total}] ❌ Файл не найден: {fio} ({iin})")

    # Итоговая статистика
    print("\n" + "-"*70)
    print(f"📊 ИТОГО:")
    print(f"   ✅ Успешно скопировано: {count_success}/{total}")
    print(f"   ❌ Не найдено: {count_failed}/{total}")
    print("-"*70 + "\n")

    # Сохраняем статистику в LOG_SUMMARY
    safe_update_summary("УВЕДОМЛЕНИЯ ОБ УСТУПКЕ ПРАВА ТРЕБОВАНИЯ", {
        "found": count_success,
        "total": total,
        "not_found": not_found_list
    })

    return count_success, count_failed
