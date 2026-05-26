"""
Блок 8: Копирование исполнительных (нотариальных) надписей
"""
import os
import shutil

from config import ROOT, TARGET_BASE
from utils import (
    normalize,
    filename_contains_iin,
    ensure_client_folder,
    safe_log,
    safe_update_summary
)

# ═══════════════════════════════════════════════════════════════
# НАСТРОЙКИ
# ═══════════════════════════════════════════════════════════════

NOTARY_SOURCE = rf"{ROOT}\Исполнительные надписи" # ТОЛЬКО ЧТЕНИЕ

# ═══════════════════════════════════════════════════════════════
# ОСНОВНАЯ ФУНКЦИЯ
# ════════════════════════════════════════════════  
def run(df_main):
    """
    Копирует исполнительные (нотариальные) надписи из папки NOTARY_SOURCE.
    
    Логика:
    1. Для каждого клиента ищет папку продукта в NOTARY_SOURCE
    2. Рекурсивно обходит все подпапки продукта
    3. Ищет PDF файлы, содержащие ИИН в имени
    4. Копирует БЕЗ переименования (сохраняет оригинальное имя)
    
    Args:
        df_main: DataFrame с колонками ['Product', 'FIO', 'IIN']
    
    Returns:
        tuple: (count_success, count_failed)
    """
    print("\n" + "="*70)
    print("  📝 БЛОК 8: ИСПОЛНИТЕЛЬНЫЕ (НОТАРИАЛЬНЫЕ) НАДПИСИ")
    print("="*70 + "\n")

    count_success = 0
    count_failed = 0
    total = len(df_main)
    not_found_list = []

    # Проверка существования источника
    if not os.path.exists(NOTARY_SOURCE):
        print(f"❌ ОШИБКА: NOTARY_SOURCE не найден: {NOTARY_SOURCE}")
        safe_log(f"[НАДПИСИ] NOTARY_SOURCE не найден: {NOTARY_SOURCE}")
        return 0, total

    print(f"📂 Источник: {NOTARY_SOURCE}")
    print(f"📦 Целевая папка: {TARGET_BASE}")
    print(f"👥 Клиентов для обработки: {total}\n")

    # === ОБРАБОТКА КАЖДОГО КЛИЕНТА ===
    for idx, row in df_main.iterrows():
        product = str(row['Product']).strip()
        fio = str(row['FIO']).strip()
        iin = str(row['IIN']).strip().zfill(12)

        # Создаём/получаем папку клиента
        target_folder = ensure_client_folder(iin, fio, TARGET_BASE)

        copied = False
        found_product_match = False

        # Перебираем подпапки продуктов
        for product_dir in os.listdir(NOTARY_SOURCE):
            product_path = os.path.join(NOTARY_SOURCE, product_dir)

            # Пропускаем файлы
            if not os.path.isdir(product_path):
                continue

            # Требуем ПОЛНОЕ совпадение продукта
            if normalize(product) != normalize(product_dir):
                continue

            found_product_match = True
            print(f"[{idx+1}/{total}] 🔍 Продукт совпал: '{product}' == '{product_dir}'")

            # Рекурсивно обходим ВСЕ подпапки продукта
            for root, dirs, files in os.walk(product_path):
                for file in files:
                    # Только PDF
                    if not file.lower().endswith('.pdf'):
                        continue

                    file_no_ext = os.path.splitext(file)[0]

                    # Проверяем наличие ИИН в имени файла
                    if filename_contains_iin(file_no_ext, iin):
                        src = os.path.join(root, file)

                        # ВАЖНО: Надпись НЕ переименовываем, копируем как есть
                        dst = os.path.join(target_folder, file)

                        try:
                            shutil.copy2(src, dst)
                            copied = True
                            count_success += 1
                            print(f"        ✅ {file}")
                        except Exception as e:
                            print(f"        ⚠️  Ошибка копирования: {e}")
                            safe_log(f"[НАДПИСИ] Ошибка копирования для {fio} ({iin}): {e}")

                        break  # Нашли файл для этого клиента
                
                if copied:
                    break  # Выходим из os.walk

            if copied:
                break  # Выходим из цикла по product_dir

        # Если продукт не найден
        if not found_product_match:
            print(f"[{idx+1}/{total}] ⚠️  Нет папки продукта '{product}'")
            safe_log(f"[НАДПИСИ] Нет папки продукта '{product}' для {fio} ({iin})")

        # Если надпись не найдена
        if not copied:
            count_failed += 1
            not_found_list.append(f"{fio}, {iin}")
            safe_log(f"[НАДПИСИ] Не найдена надпись по ИИН: {fio}, {iin}, продукт: {product}")
            
            if found_product_match:
                print(f"[{idx+1}/{total}] ❌ Надпись не найдена (продукт есть, но нет файла с ИИН)")
            else:
                print(f"[{idx+1}/{total}] ❌ Надпись не найдена (нет папки продукта)")

    # === ИТОГ ===
    print("\n" + "="*70)
    print("📊 ИТОГО:")
    print(f"   ✅ Успешно скопировано: {count_success}/{total}")
    print(f"   ❌ Не найдено: {count_failed}/{total}")
    print("="*70 + "\n")

    safe_update_summary("ИСПОЛНИТЕЛЬНЫЕ НАДПИСИ", {
        "found": count_success,
        "total": total,
        "not_found": not_found_list
    })

    return count_success, count_failed