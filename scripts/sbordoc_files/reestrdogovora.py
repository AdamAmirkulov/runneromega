"""
Блок 7: Реестр договора цессии
"""
from utils import safe_update_summary
from config import MAIN_EXCEL, TARGET_BASE, LOG_SUMMARY,BASE_CESSII
import os
import shutil
import pandas as pd

# === ПУТИ ===

log_file_path = 'log_not_copied.txt'

# если сводный словарь ещё не создан – создаём
try:
    LOG_SUMMARY
except NameError:
    LOG_SUMMARY = {}

# === ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ===

def normalize(text: str) -> str:
    """
    Нормализуем названия продуктов и имена папок для строгого сравнения.
    Убираем пробелы, дефисы, подчёркивания и заменяем похожие кириллические буквы.
    """
    trans = str.maketrans(
        "аеорсухкмвнтзм",    # кириллические символы, похожие на латиницу
        "aercsyxkmvntzm"     # латиница
    )
    return (
        str(text).lower()
        .translate(trans)
        .replace('-', '')
        .replace('_', '')
        .replace(' ', '')
    )

def safe_folder_name(fio: str, iin: str) -> str:
    """
    Формируем имя папки клиента 'ФИО, ИИН' и заменяем запрещённые символы.
    """
    raw = f"{fio}, {iin}"
    for bad, good in [
        ('/', '_'), ('\\', '_'), (':', '_'), ('*', '_'),
        ('?', '_'), ('"', '_'), ('<', '_'), ('>', '_'), ('|', '_')
    ]:
        raw = raw.replace(bad, good)
    return raw

def ensure_client_folder(iin: str, fio: str, base_dir: str) -> str:
    """
    Убеждаемся, что папка клиента существует, возвращаем путь.
    """
    folder_name = safe_folder_name(fio, iin)
    target_folder_path = os.path.join(base_dir, folder_name)
    os.makedirs(target_folder_path, exist_ok=True)
    return target_folder_path


# === ЧИТАЕМ EXCEL ===
def run(main_df):
    print("\n=== БЛОК 4: Реестр договора цессии по продукту ===\n")

        
    df_main = pd.read_excel(MAIN_EXCEL, usecols=[1, 2, 3], header=0)
    df_main.columns = ['Product', 'FIO', 'IIN']
    df_main['IIN'] = df_main['IIN'].astype(str).str.zfill(12)

    print("\n=== БЛОК 4: Реестр договора цессии по продукту ===\n")

    # детальные счётчики
    count_success_registry = 0
    count_failed_registry  = 0

    # счётчики для сводного отчёта
    total_registry     = len(df_main)
    found_registry     = 0
    not_found_registry = []

    for _, row in df_main.iterrows():
        product = str(row['Product']).strip()
        fio     = str(row['FIO']).strip()
        iin     = str(row['IIN']).strip().zfill(12)

        # Папка клиента
        target_folder = ensure_client_folder(iin, fio, TARGET_BASE)

        copied = False
        found_product_match = False

        # Перебираем продуктовые папки в Цессии
        for product_dir in os.listdir(BASE_CESSII):
            product_path = os.path.join(BASE_CESSII, product_dir)

            # оставляем только папки
            if not os.path.isdir(product_path):
                continue

            # продукт должен совпасть 100% (после normalize)
            if normalize(product) != normalize(product_dir):
                continue

            found_product_match = True
            print(f"[ЦЕССИЯ-РЕЕСТР] Продукт совпал: '{product}' == '{product_dir}'")

            # Ищем «Реестр договора цессии»: PDF, имя содержит слова реестр/договора/цессии
            for root, dirs, files in os.walk(product_path):
                for file in files:
                    if not file.lower().endswith('.pdf'):
                        continue

                    filename_lower = file.lower()
                    if ('реестр'  in filename_lower and
                        'договора' in filename_lower and
                        'цессии'   in filename_lower):
                        src = os.path.join(root, file)

                        # целевое имя в папке клиента:
                        # "Реестр договора цессии, <наименование продукта>.pdf"
                        new_filename = f"Реестр договора цессии, {product}.pdf"
                        dst = os.path.join(target_folder, new_filename)

                        shutil.copy2(src, dst)
                        copied = True
                        count_success_registry += 1
                        found_registry += 1
                        print(f"[✔ РЕЕСТР] {file} → '{new_filename}' в '{target_folder}'")
                        break
                if copied:
                    break
            if copied:
                break

        if not found_product_match:
            print(f"[⚠ РЕЕСТР] Не найдена папка продукта '{product}' для {fio} ({iin})")

        if not copied:
            count_failed_registry += 1
            not_found_registry.append(f"{fio}, {iin}")  # для сводного лога

            # пишем в лог, но не затираем существующее содержимое
            with open(log_file_path, 'a', encoding='utf-8') as log_file:
                log_file.write(
                    f"[РЕЕСТР ДОГОВОРА ЦЕССИИ] Не найден файл для продукта '{product}' "
                    f"у клиента {fio}, {iin}\n"
                )

    print("\n--- ИТОГ РЕЕСТРОВ ДОГОВОРА ЦЕССИИ ---")
    print(f"✔️ Скопировано реестров: {count_success_registry}")
    print(f"❌ Не найдено реестров: {count_failed_registry}")
    print(f"📄 Лог: {os.path.abspath(log_file_path)}")

    # сохраняем результат этого блока в общий сводный лог
    safe_update_summary("РЕЕСТР ДОГОВОРА ЦЕССИИ", {
        "found": found_registry,
        "total": total_registry,
        "not_found": not_found_registry,
    })
    return count_success_registry, count_failed_registry

