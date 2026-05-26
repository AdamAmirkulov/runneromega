"""
Блок: formirovaniepool
"""
from utils import safe_update_summary
# -*- coding: utf-8 -*-
# 1) Открывает Excel:
#    C:\Users\User\Desktop\Документы для подачи Исков\Отчёт по отменам_Omega.xlsx
# 2) Берёт строки, где в колонке B НЕТ слова "Vivus" (регистр неважен).
# 3) Берёт по порядку ИИН из этих строк (колонка D) и разбивает на группы по 10.
# 4) Для каждой группы создаёт папку:
#    C:\Users\User\Desktop\Документы для подачи Исков\Папки_для_исков\Пул 1
#    C:\Users\User\Desktop\Документы для подачи Исков\Папки_для_исков\Пул 2
#    и т.д.
# 5) В каждой группе по ИИН ищет «исходные» папки в
#    C:\Users\User\Desktop\Документы для подачи Исков\Папки_для_исков
#    (по вхождению ИИН в название папки, например "ФИО, 990101300000").
#    Из найденных папок копирует все файлы в соответствующий "Пул N".
#    Если файл с таким же именем в папке Пула уже есть — не копирует (оставляет один).

import os
import re
import shutil
from openpyxl import load_workbook
from config import MAIN_EXCEL, TARGET_BASE
# ===== НАСТРОЙКИ ПУТЕЙ =====
EXCEL_PATH = MAIN_EXCEL
FOLDERS_ROOT = TARGET_BASE 
# Колонки в Excel (1-based, как в Excel)
COL_PRODUCT = 2   # B - продукт (для фильтра по "Vivus")
COL_IIN = 4       # D - ИИН

# Максимум ИИН в одном пуле
POOL_SIZE = 10


def read_iin_groups_from_excel(path, pool_size=10):
    """
    Возвращает список групп ИИН:
    [[iin1, iin2, ... upto 10], [iin11, ...], ...]
    """
    if not os.path.isfile(path):
        raise FileNotFoundError(f"Excel файл не найден: {path}")

    wb = load_workbook(path, data_only=True)
    ws = wb.active  # если нужен конкретный лист — можно указать имя

    filtered_iins = []

    # пропускаем первую строку (заголовок)
    for row_idx in range(2, ws.max_row + 1):
        product = ws.cell(row=row_idx, column=COL_PRODUCT).value
        iin = ws.cell(row=row_idx, column=COL_IIN).value

        if iin is None:
            continue

        # строка продукта → строка, чтобы искать "Vivus"
        product_str = str(product).strip() if product is not None else ""

        # пропускаем все, где встречается Vivus (в любом регистре)
        if "vivus" in product_str.lower():
            continue

        # сохраняем ИИН как строку
        filtered_iins.append(str(iin).strip())

    # Разбиваем на группы не более pool_size
    groups = []
    for i in range(0, len(filtered_iins), pool_size):
        groups.append(filtered_iins[i:i + pool_size])

    return groups


def build_iin_to_folder_map(root_dir):
    """
    Проходит по всем папкам в root_dir (кроме уже созданных 'Пул N')
    и строит словарь: IIN -> [список папок с этим ИИН].
    ИИН ищется как последовательность из 11-12 цифр в названии папки.
    """
    iin_to_folders = {}

    if not os.path.isdir(root_dir):
        raise FileNotFoundError(f"Каталог с папками не найден: {root_dir}")

    for name in os.listdir(root_dir):
        full = os.path.join(root_dir, name)
        if not os.path.isdir(full):
            continue

        # пропускаем уже созданные пулы
        if name.lower().startswith("пул "):
            continue

        # ищем все последовательности цифр длиной 11–12 (IIN)
        for m in re.finditer(r"\d{11,12}", name):
            iin = m.group(0)
            iin_to_folders.setdefault(iin, []).append(full)

    return iin_to_folders


def copy_unique_files(src_folder, dst_folder):
    """
    Копирует файлы из src_folder в dst_folder.
    Если в dst_folder уже есть файл с таким именем — пропускает его.
    """
    os.makedirs(dst_folder, exist_ok=True)

    for entry in os.listdir(src_folder):
        src_path = os.path.join(src_folder, entry)
        if not os.path.isfile(src_path):
            continue

        dst_path = os.path.join(dst_folder, entry)

        # если файл уже существует — не перезаписываем
        if os.path.exists(dst_path):
            continue

        shutil.copy2(src_path, dst_path)


def run(df_main):
    # 1. читаем группы ИИН из Excel
    iin_groups = read_iin_groups_from_excel(EXCEL_PATH, POOL_SIZE)
    print(f"Найдено ИИН (после фильтра по 'Vivus'): {sum(len(g) for g in iin_groups)}")
    print(f"Всего пулов: {len(iin_groups)}")
    
    count_success = 0
    count_failed = 0
    # 2. строим карту: ИИН -> папки
    iin_to_folders = build_iin_to_folder_map(FOLDERS_ROOT)

    # 3. создаём пулы и копируем файлы
    for idx, group in enumerate(iin_groups, start=1):
        pool_name = f"Пул {idx}"
        pool_dir = os.path.join(FOLDERS_ROOT, pool_name)
        os.makedirs(pool_dir, exist_ok=True)

        print(f"\n=== {pool_name} ===")
        for iin in group:
            folders = iin_to_folders.get(iin)
            count_success += 1
            if not folders:
                print(f"  [!] Не найдена папка для ИИН {iin}")
                count_failed += 1   
                continue

            for folder in folders:
                print(f"  ИИН {iin}: копирую из {folder}")
                copy_unique_files(folder, pool_dir)

    print("\nГотово. Папки 'Пул 1', 'Пул 2', ... созданы в:")
    print(FOLDERS_ROOT)
    return count_success, count_failed 


