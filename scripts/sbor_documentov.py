#!/usr/bin/env python
# coding: utf-8

# ## Оглавление
# - [ЛОГИ](#ЛОГИ)
# - [Создавать папки и добавлять учдоки](#Создавать-папки-и-добавлять-учдоки)
# - [Оферты](#Оферты)
# - [Уведомления об уступки права требования](#Уведомления-об-уступки-права-требования)
# - [Договор цессии](#Договор-цессии)
# - [Реестр договора цессии](#Реестр-договора-цессии)
# - [Исполнительные надписи](#Исполнительные-надписи)
# - [Госпошлины](#Госпошлины)
# - [Поиск в СК адреса + Привязка к Суду + Импорт Суда в Delta-M](#Поиск-в-СК-адреса-+-Привязка-к-Суду-+-Импорт-Суда-в-Delta-M)
# - [Досудебная претензия](#Досудебная-претензия)
# - [Скрин отправки Досудебной претензии](#Скрин-отправки-Досудебной-претензии)
# - [Расчет задолженности](#Расчет-задолженности)
# - [Исковое заявление](#Исковое-заявление)
# - [Постановление об отмене надписи](#Постановление-об-отмене-надписи)
# - [Формирование по Пулам](#Формирование-по-Пулам)
# - [Формирование Массового заявления в Медеуский суд](#Формирование-Массового-заявления-в-Медеуский-суд)
# - [Логирование Сбора документов](#Логирование-Сбора-документов)
# - [Дублирование папки](#Дублирование-папки)

# # ЛОГИ

import os

# === ОБЩИЙ ЛОГ ПО БЛОКАМ ===

# Папка, где находятся папки клиентов
BASE_FOLDER = r"\\PC002\work folder\Документы для подачи ИсковПапки_для_исков"

# На всякий случай — создаём её, если вдруг нет
os.makedirs(BASE_FOLDER, exist_ok=True)

# Один ОДИН общий файл лога для всего скрипта — СРАЗУ внутри Папки_для_исков
log_file_path = os.path.join(BASE_FOLDER, "лог_сбор_документов.txt")

# Сюда каждый блок будет складывать свои итоги (ОФЕРТЫ, УВЕДОМЛЕНИЯ и т.д.)
LOG_SUMMARY = {}

print("Лог будет писаться сюда:", log_file_path)

# # Создавать папки и добавлять учдоки

import os
import shutil
import pandas as pd

# Путь к Excel-файлу
excel_path = r'C:\Users\User\Desktop\Документы для подачи Исков\Отчёт по отменам_Omega.xlsx'

# Папка с учредительными документами
docs_source_folder = r'C:\Users\User\Desktop\Документы для подачи Исков\Учредительные документы'

# Список файлов, которые нужно скопировать
docs_to_copy = [
    'Доверенность представителя.pdf',
    'Приказ на директора.pdf',
    'Устав ТОО.pdf'
]

# Чтение Excel, берём ФИО (столбец C = индекс 2), ИИН (D = индекс 3)
df = pd.read_excel(excel_path, usecols=[2, 3], header=0)

# Папка, куда будут создаваться папки
base_dir = os.path.join(os.path.dirname(excel_path), 'Папки_для_исков')
os.makedirs(base_dir, exist_ok=True)

# Создание папок и копирование документов
for index, row in df.iterrows():
    fio = str(row.iloc[0]).strip()
    raw_iin = str(row.iloc[1]).strip()
    
    if fio and raw_iin:
        iin = raw_iin.zfill(12)
        safe_name = (
            f"{fio}, {iin}"
            .replace("/", "_")
            .replace("\\", "_")
            .replace(":", "_")
            .replace("*", "_")
            .replace("?", "_")
            .replace('"', "_")
            .replace("<", "_")
            .replace(">", "_")
            .replace("|", "_")
        )
        
        folder_path = os.path.join(base_dir, safe_name)
        os.makedirs(folder_path, exist_ok=True)
        
        # Копируем нужные документы
        for doc_name in docs_to_copy:
            source_path = os.path.join(docs_source_folder, doc_name)
            target_path = os.path.join(folder_path, doc_name)
            if os.path.exists(source_path):
                shutil.copy2(source_path, target_path)
            else:
                print(f"[!] Файл не найден: {source_path}")

print(f"Создано {len(df)} папок и скопированы документы в: {base_dir}")

# # Оферты

import os
import shutil
import pandas as pd

# === КОНФИГ ===
main_excel    = r'C:\Users\User\Desktop\Документы для подачи Исков\Отчёт по отменам_Omega.xlsx'
base_cessii   = r'D:\work folder\Цессии'   # ИСТОЧНИК, ТОЛЬКО ЧТЕНИЕ
target_base   = r'C:\Users\User\Desktop\Документы для подачи Исков\Папки_для_исков'
log_file_path = 'log_not_copied.txt'

# если сводный словарь ещё не создан в другом блоке – создадим
try:
    LOG_SUMMARY
except NameError:
    LOG_SUMMARY = {}

# счётчики для детального лога
count_success = 0
count_failed  = 0

# читаем Excel со списком клиентов
df_main = pd.read_excel(main_excel, usecols=[1, 2, 3], header=0)
df_main.columns = ['Product', 'FIO', 'IIN']
df_main['IIN'] = df_main['IIN'].astype(str).str.zfill(12)

# счётчики для сводного отчёта по блоку "ОФЕРТЫ"
total_offers    = len(df_main)   # всего клиентов
found_offers    = 0              # сколько оферт нашли
not_found_offers = []            # список "ФИО, ИИН" для ненайденных

# при необходимости можно очистить лог:
with open(log_file_path, 'w', encoding='utf-8') as log_file:
    log_file.write("Файлы, которые не удалось найти и скопировать:\n\n")

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

print("\n=== ПОИСК ОФЕРТ В ЦЕССИЯХ (ЕДИНЫЕ ПРАВИЛА ДЛЯ ВСЕХ ПРОДУКТОВ, ТОЛЬКО ПО ИИН) ===\n")

for _, row in df_main.iterrows():
    product = str(row['Product']).strip()
    fio     = str(row['FIO']).strip()
    iin     = str(row['IIN']).strip().zfill(12)
    
    # Папка клиента
    target_folder = build_client_folder(target_base, fio, iin)
    
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

# # Уведомления об уступки права требования

import os
import shutil
import pandas as pd

# === ПУТИ ===
main_excel  = r'C:\Users\User\Desktop\Документы для подачи Исков\Отчёт по отменам_Omega.xlsx'
target_base = r'C:\Users\User\Desktop\Документы для подачи Исков\Папки_для_исков'
email_source = r'D:\work folder\email'  # ТОЛЬКО ЧТЕНИЕ, НИЧЕГО НЕ МЕНЯЕМ

log_file_path = 'log_not_copied.txt'

# если сводный словарь ещё не создан (например, блок оферт не запускали) – создаём
try:
    LOG_SUMMARY
except NameError:
    LOG_SUMMARY = {}

# === ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ===

def normalize(text: str) -> str:
    """
    Нормализуем названия продуктов (Excel vs имя папки продукта).
    Убираем пробелы, дефисы, подчёркивания и выравниваем похожие буквы (а/a, с/c, х/x и т.д.).
    """
    trans = str.maketrans(
        "аеорсухкмвнтзм",    # похожие кириллические буквы
        "aercsyxkmvntzm"     # латиница
    )
    return (
        str(text).lower()
        .translate(trans)
        .replace('-', '')
        .replace('_', '')
        .replace(' ', '')
    )

def filename_contains_iin(file_name_no_ext: str, iin: str) -> bool:
    """
    Проверяем, есть ли ИИН внутри имени файла (без пробелов, дефисов и подчёркиваний).
    """
    candidate = (
        file_name_no_ext
        .replace(' ', '')
        .replace('-', '')
        .replace('_', '')
    )
    return iin in candidate

def safe_folder_name(fio: str, iin: str) -> str:
    """
    Формируем имя папки клиента 'ФИО, ИИН' и чистим запрещённые для Windows символы.
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
    Гарантируем, что папка клиента существует, и возвращаем путь к ней.
    """
    folder_name = safe_folder_name(fio, iin)
    target_folder_path = os.path.join(base_dir, folder_name)
    os.makedirs(target_folder_path, exist_ok=True)
    return target_folder_path

# === ГОТОВИМ ДАННЫЕ ИЗ EXCEL ===

df_main = pd.read_excel(main_excel, usecols=[1, 2, 3], header=0)
df_main.columns = ['Product', 'FIO', 'IIN']
df_main['IIN'] = df_main['IIN'].astype(str).str.zfill(12)

print("\n=== БЛОК 3: уведомления об уступке права требования из EMAIL ===\n")

# детальные счётчики этого блока
count_success_email = 0
count_failed_email  = 0

# счётчики для сводного отчёта
total_notify     = len(df_main)
found_notify     = 0
not_found_notify = []

for _, row in df_main.iterrows():
    product = str(row['Product']).strip()
    fio     = str(row['FIO']).strip()
    iin     = str(row['IIN']).strip().zfill(12)
    
    # создаём / получаем папку клиента "ФИО, ИИН"
    target_folder = ensure_client_folder(iin, fio, target_base)
    
    copied = False
    found_product_match = False
    
    # листаем подпапки продуктов внутри email_source
    for product_dir in os.listdir(email_source):
        product_path = os.path.join(email_source, product_dir)
        
        # пропускаем файлы, оставляем только папки
        if not os.path.isdir(product_path):
            continue
        
        # продукт должен совпасть 1-в-1 по нормализованной форме
        if normalize(product) != normalize(product_dir):
            continue
        
        found_product_match = True
        print(f"[EMAIL] Продукт совпал: '{product}' == '{product_dir}'")
        
        # рекурсивно обходим ВСЕ подпапки и ищем PDF, в имени которого есть ИИН
        for root, dirs, files in os.walk(product_path):
            for file in files:
                if not file.lower().endswith('.pdf'):
                    continue
                
                file_no_ext = os.path.splitext(file)[0]
                
                if filename_contains_iin(file_no_ext, iin):
                    src = os.path.join(root, file)
                    
                    # ВАЖНО: не переименовываем файл, копируем как есть
                    dst = os.path.join(target_folder, file)
                    
                    shutil.copy2(src, dst)
                    copied = True
                    count_success_email += 1
                    found_notify += 1
                    print(f"[✔ EMAIL] {file} → {target_folder}")
                    break  # нашли pdf для этого клиента
            if copied:
                break  # выходим из os.walk
        if copied:
            break  # выходим из цикла по product_dir
    
    if not found_product_match:
        print(f"[⚠ EMAIL] Не найдена папка продукта '{product}' для {fio} ({iin})")
    
    if not copied:
        count_failed_email += 1
        not_found_notify.append(f"{fio}, {iin}")  # для сводного лога
        
        # логируем, чтобы потом руками проверить
        with open(log_file_path, 'a', encoding='utf-8') as log_file:
            log_file.write(
                f"[EMAIL / уведомление] Не найден pdf по ИИН: {fio}, {iin}, продукт: {product}\n"
            )

print("\n--- ИТОГ EMAIL ---")
print(f"✔️ Скопировано из email (уведомления): {count_success_email}")
print(f"❌ Не найдено из email (уведомления): {count_failed_email}")
print(f"📄 Лог: {os.path.abspath(log_file_path)}")

# сохраняем результат блока «Уведомления» в общий сводный лог
LOG_SUMMARY["УВЕДОМЛЕНИЯ ОБ УСТУПКЕ ПРАВА ТРЕБОВАНИЯ"] = {
    "found": found_notify,
    "total": total_notify,
    "not_found": not_found_notify,
}

# # Договор цессии

import os
import shutil
import pandas as pd

# === ПУТИ ===
main_excel    = r'C:\Users\User\Desktop\Документы для подачи Исков\Отчёт по отменам_Omega.xlsx'
target_base   = r'C:\Users\User\Desktop\Документы для подачи Исков\Папки_для_исков'
base_cessii   = r'D:\work folder\Цессии'  # ТОЛЬКО ЧТЕНИЕ
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

df_main = pd.read_excel(main_excel, usecols=[1, 2, 3], header=0)
df_main.columns = ['Product', 'FIO', 'IIN']
df_main['IIN'] = df_main['IIN'].astype(str).str.zfill(12)

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
    target_folder = ensure_client_folder(iin, fio, target_base)
    
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
LOG_SUMMARY["ДОГОВОРЫ ЦЕССИИ"] = {
    "found": found_cession,
    "total": total_cession,
    "not_found": not_found_cession,
}

# # Реестр договора цессии

import os
import shutil
import pandas as pd

# === ПУТИ ===
main_excel    = r'C:\Users\User\Desktop\Документы для подачи Исков\Отчёт по отменам_Omega.xlsx'
target_base   = r'C:\Users\User\Desktop\Документы для подачи Исков\Папки_для_исков'
base_cessii   = r'D:\work folder\Цессии'  # ТОЛЬКО ЧТЕНИЕ
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

df_main = pd.read_excel(main_excel, usecols=[1, 2, 3], header=0)
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
    target_folder = ensure_client_folder(iin, fio, target_base)
    
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
LOG_SUMMARY["РЕЕСТРЫ ДОГОВОРА ЦЕССИИ"] = {
    "found": found_registry,
    "total": total_registry,
    "not_found": not_found_registry,
}

# # Исполнительные надписи

import os
import shutil
import pandas as pd

# === ПУТИ ДЛЯ БЛОКА 5 ===
main_excel    = r'C:\Users\User\Desktop\Документы для подачи Исков\Отчёт по отменам_Omega.xlsx'
target_base   = r'C:\Users\User\Desktop\Документы для подачи Исков\Папки_для_исков'
notary_source = r'D:\work folder\Надписи - Рассортированные_v2'  # ТОЛЬКО ЧТЕНИЕ
log_file_path = 'log_not_copied.txt'

# если сводный словарь ещё не создан – создаём
try:
    LOG_SUMMARY
except NameError:
    LOG_SUMMARY = {}

# === ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ===

def normalize(text: str) -> str:
    """
    Нормализуем названия продуктов и папок:
    - приводим к нижнему регистру,
    - убираем пробелы, дефисы, подчёркивания,
    - заменяем похожие кириллические буквы на латинские.
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
    Формат имени папки клиента: 'ФИО, ИИН', плюс очистка от запрещённых символов Windows.
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
    Гарантируем, что папка клиента существует, и возвращаем путь к ней.
    """
    folder_name = safe_folder_name(fio, iin)
    target_folder_path = os.path.join(base_dir, folder_name)
    os.makedirs(target_folder_path, exist_ok=True)
    return target_folder_path

def filename_contains_iin(file_name_no_ext: str, iin: str) -> bool:
    """
    Проверяем, есть ли ИИН внутри имени файла, допускаем варианты:
    '010101123456-Надпись.pdf', 'Нотариальная надпись 010101123456.pdf',
    '010101123456_НАДПИСЬ.pdf', и т.п.
    """
    candidate = (
        file_name_no_ext
        .replace(' ', '')
        .replace('-', '')
        .replace('_', '')
    )
    return iin in candidate

# === ЧТЕНИЕ EXCEL ===

df_main = pd.read_excel(main_excel, usecols=[1, 2, 3], header=0)
df_main.columns = ['Product', 'FIO', 'IIN']
df_main['IIN'] = df_main['IIN'].astype(str).str.zfill(12)

print("\n=== БЛОК 5: Нотариальные надписи ===\n")

# детальные счётчики
count_success_notary = 0
count_failed_notary  = 0

# счётчики для сводного отчёта
total_notary     = len(df_main)
found_notary     = 0
not_found_notary = []

for _, row in df_main.iterrows():
    product = str(row['Product']).strip()
    fio     = str(row['FIO']).strip()
    iin     = str(row['IIN']).strip().zfill(12)
    
    # Папка клиента (ФИО, ИИН)
    target_folder = ensure_client_folder(iin, fio, target_base)
    
    copied = False
    found_product_match = False
    
    # Перебираем подпапки продуктов внутри notary_source
    for product_dir in os.listdir(notary_source):
        product_path = os.path.join(notary_source, product_dir)
        
        if not os.path.isdir(product_path):
            continue
        
        # Требуем ПОЛНОЕ совпадение продукта
        if normalize(product) != normalize(product_dir):
            continue
        
        found_product_match = True
        print(f"[НАДПИСЬ] Продукт совпал 1-в-1: '{product}' == '{product_dir}'")
        
        # Рекурсивно обходим ВСЕ подпапки продукта, ищем PDF где есть ИИН
        for root, dirs, files in os.walk(product_path):
            for file in files:
                # Только PDF
                if not file.lower().endswith('.pdf'):
                    continue
                
                file_no_ext = os.path.splitext(file)[0]
                
                if filename_contains_iin(file_no_ext, iin):
                    src = os.path.join(root, file)
                    
                    # Нотариальную надпись НЕ переименовываем.
                    dst = os.path.join(target_folder, file)
                    
                    shutil.copy2(src, dst)
                    copied = True
                    count_success_notary += 1
                    found_notary += 1
                    print(f"[✔ НАДПИСЬ] {file} → {target_folder}")
                    break
            if copied:
                break
        if copied:
            break
    
    if not found_product_match:
        print(f"[⚠ НАДПИСЬ] Не найдена папка продукта '{product}' для {fio} ({iin})")
    
    if not copied:
        count_failed_notary += 1
        not_found_notary.append(f"{fio}, {iin}")  # для сводного лога
        
        # добавим запись в лог, не перетирая лог предыдущих блоков
        with open(log_file_path, 'a', encoding='utf-8') as log_file:
            log_file.write(
                f"[НАДПИСЬ] Не найдена нотариальная надпись по ИИН: {fio}, {iin}, продукт: {product}\n"
            )

print("\n--- ИТОГ НОТАРИАЛЬНЫЕ НАДПИСИ ---")
print(f"✔️ Скопировано нотариальных надписей: {count_success_notary}")
print(f"❌ Не найдено нотариальных надписей: {count_failed_notary}")
print(f"📄 Лог: {os.path.abspath(log_file_path)}")

# сохраняем результат этого блока в общий сводный лог
LOG_SUMMARY["НОТАРИАЛЬНЫЕ НАДПИСИ"] = {
    "found": found_notary,
    "total": total_notary,
    "not_found": not_found_notary,
}

# # Госпошлины

# -*- coding: utf-8 -*-
import os
import re
import shutil
from datetime import datetime

from pypdf import PdfReader, PdfWriter

# нестрогое сравнение ФИО
try:
    from rapidfuzz.fuzz import token_set_ratio as fuzz_ratio
    HAVE_RF = True
except Exception:
    from difflib import SequenceMatcher
    HAVE_RF = False

# ===== ПУТИ =====
SOURCE_ROOT = r"D:\work folder\Госпошлины"  # внутри: 2023, 2024, 2025 ...
DEST_ROOT   = r"C:\Users\User\Desktop\Документы для подачи Исков\Папки_для_исков"
log_file_path = 'log_not_copied.txt'

try:
    LOG_SUMMARY
except NameError:
    LOG_SUMMARY = {}

# ===== УТИЛИТЫ =====
def normalize_name(s: str) -> str:
    s = (s or "").lower()
    s = s.replace('\xa0', ' ').replace('\u202f', ' ').replace('\u200b', ' ')
    s = s.replace('ё', 'е')
    s = re.sub(r'[.,;:()"\']+', ' ', s)
    s = re.sub(r'[-_]+', ' ', s)
    s = re.sub(r'\s{2,}', ' ', s).strip()
    return s

def similarity(a: str, b: str) -> float:
    a, b = normalize_name(a), normalize_name(b)
    if HAVE_RF:
        return float(fuzz_ratio(a, b))
    return SequenceMatcher(None, a, b).ratio() * 100.0

def sanitize_filename(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', '', name)

def _clean_text(text: str) -> str:
    if not text:
        return ""
    t = (text.replace('\xa0', ' ')
              .replace('\u202f', ' ')
              .replace('\u200b', ' '))
    t = re.sub(r'[\n\r\f]+', ' ', t)
    t = re.sub(r'\s{2,}', ' ', t).strip()
    return t

# ===== ИЗВЛЕЧЕНИЕ ФИО =====
# маркеры, на которых обрезаем хвост после ФИО
BUDGET_STOP_RE = r'(?:Бюджеттік|Бюджетке|Код\s+бюджетной|бюджетной\s+классификации|жіктеу\s+коды|классификация)'

def extract_fio(page_text: str) -> str | None:
    """
    Надёжно вытаскивает ФИО:
    1) пытается искать после блока 'Назначение платежа' (но не режет по 'Код назначения', потому что порядок в тексте плавает)
    2) затем fallback по всей странице
    ФИО берём как 2–4 слова между 'госпошлина' и 'Бюджеттік/Код бюджетной...'
    """
    t = _clean_text(page_text)
    if not t:
        return None
    
    # 1) Пытаемся ограничиться частью после "Назначение платежа"
    # (в твоих PDF "госпошлина ..." идёт рядом с этим блоком, но иногда после правой колонки)
    purpose_anchor = re.search(r'(?:Төлемнің\s+мақсаты|Назначение\s+платежа)\s*:', t, flags=re.IGNORECASE)
    search_zone = t[purpose_anchor.start():] if purpose_anchor else t
    
    fio = _extract_fio_between_markers(search_zone)
    if fio:
        return fio
    
    # 2) fallback — по всей странице
    return _extract_fio_between_markers(t)

def _extract_fio_between_markers(text: str) -> str | None:
    # Ищем: госпошлина <ФИО> Бюджеттік...
    m = re.search(
        rf'госпошлина\s+(.+?)\s+{BUDGET_STOP_RE}',
        text,
        flags=re.IGNORECASE
    )
    if not m:
        return None
    
    candidate = m.group(1).strip()
    # чистим пунктуацию вокруг
    candidate = re.sub(r'[,:;()"]+', ' ', candidate)
    candidate = re.sub(r'\s{2,}', ' ', candidate).strip()
    words = candidate.split()
    
    # ФИО обычно 2-4 слова (встречается 4 слова как у тебя на скрине)
    if len(words) < 2:
        return None
    
    words = words[:4]
    fio = " ".join(words)
    fio = " ".join(w[:1].upper() + w[1:] for w in fio.split())
    return fio

# ===== ПАПКИ =====
def latest_year_dir(root: str) -> str | None:
    now_year = str(datetime.now().year)
    candidates = [d for d in os.listdir(root) if os.path.isdir(os.path.join(root, d))]
    if not candidates:
        return None
    if now_year in candidates:
        return os.path.join(root, now_year)
    candidates_num = sorted(candidates, key=lambda x: int(re.sub(r'\D', '', x) or 0), reverse=True)
    return os.path.join(root, candidates_num[0])

DATE_RE = re.compile(r'(\d{2})\.(\d{2})\.(\d{4})')  # DD.MM.YYYY

def parse_date_from_foldername(name: str) -> datetime | None:
    m = DATE_RE.search(name)
    if not m:
        return None
    d, mth, y = map(int, m.groups())
    try:
        return datetime(y, mth, d)
    except ValueError:
        return None

def pick_latest_day_folder(year_dir: str) -> str | None:
    subdirs = [os.path.join(year_dir, d) for d in os.listdir(year_dir)
               if os.path.isdir(os.path.join(year_dir, d))]
    if not subdirs:
        return None
    
    dated, undated = [], []
    for p in subdirs:
        dt = parse_date_from_foldername(os.path.basename(p))
        if dt:
            dated.append((dt, p))
        else:
            undated.append(p)
    
    if dated:
        dated.sort(key=lambda x: x[0], reverse=True)
        return dated[0][1]
    
    return max(undated, key=lambda p: os.path.getmtime(p))

def list_client_folders(dest_root: str):
    res = []
    if not os.path.isdir(dest_root):
        return res
    for name in os.listdir(dest_root):
        p = os.path.join(dest_root, name)
        if os.path.isdir(p):
            fio_part = name.split(',', 1)[0].strip()
            res.append((p, fio_part))
    return res

def find_best_client_folder(fio: str, candidates: list, threshold: float = 90.0):
    best_path, best_score = None, -1.0
    for path, folder_fio in candidates:
        sc = similarity(fio, folder_fio)
        if sc > best_score:
            best_score, best_path = sc, path
    return (best_path if best_score >= threshold else None), best_score

# ===== ОСНОВНОЙ СЦЕНАРИЙ =====
def main():
    global LOG_SUMMARY
    
    total_gos = 0
    found_gos = 0
    not_found_gos = []
    
    year_dir = latest_year_dir(SOURCE_ROOT)
    if not year_dir or not os.path.isdir(year_dir):
        print(f"❌ Не найдена папка года в '{SOURCE_ROOT}'")
        return
    
    day_dir = pick_latest_day_folder(year_dir)
    if not day_dir:
        print(f"❌ В '{year_dir}' нет папок с датами.")
        return
    
    out_dir = os.path.join(day_dir, "Готовые")
    os.makedirs(out_dir, exist_ok=True)
    
    print(f"📂 Год: {year_dir}")
    print(f"📂 Папка даты: {day_dir}")
    print(f"📁 Готовые: {out_dir}")
    print(f"📦 Папки клиентов: {DEST_ROOT}")
    
    client_dirs = list_client_folders(DEST_ROOT)
    print(f"👥 Найдено клиентских папок: {len(client_dirs)}")
    
    pdf_files = [f for f in os.listdir(day_dir)
                 if f.lower().endswith(".pdf") and os.path.isfile(os.path.join(day_dir, f))]
    if not pdf_files:
        print("⚠️ В выбранной папке PDF не найдены.")
        return
    
    for pdf_name in pdf_files:
        src_path = os.path.join(day_dir, pdf_name)
        try:
            reader = PdfReader(src_path)
        except Exception as e:
            print(f"⛔ Не удалось открыть '{pdf_name}': {e}")
            with open(log_file_path, 'a', encoding='utf-8') as log:
                log.write(f"[ГОСПОШЛИНА] Не удалось открыть файл '{pdf_name}': {e}\n")
            continue
        
        num_pages = len(reader.pages)
        print(f"\n===== Обработка: {pdf_name} | страниц: {num_pages} =====")
        
        for i in range(num_pages):
            try:
                page = reader.pages[i]
                page_text = page.extract_text() or ""
                
                fio = extract_fio(page_text)
                
                if fio:
                    total_gos += 1
                    file_name = f"Госпошлина, {fio}.pdf"
                else:
                    file_name = f"Госпошлина, page-{i+1:03d}.pdf"
                
                file_name = sanitize_filename(file_name)
                page_pdf_path = os.path.join(out_dir, file_name)
                
                writer = PdfWriter()
                writer.add_page(page)
                with open(page_pdf_path, "wb") as f:
                    writer.write(f)
                
                print(f"✅ Стр. {i+1}/{num_pages}: {file_name} — сохранён в 'Готовые'")
                
                if fio:
                    best_path, score = find_best_client_folder(fio, client_dirs, threshold=90.0)
                    if best_path:
                        dst_path = os.path.join(best_path, file_name)
                        try:
                            shutil.copy2(page_pdf_path, dst_path)
                            found_gos += 1
                            print(f"   ➤ 📤 Скопирован в клиентскую папку [{score:.0f}%]: {best_path}")
                        except Exception as e:
                            print(f"   ➤ ⛔ Ошибка копирования в '{best_path}': {e}")
                            not_found_gos.append(fio)
                            with open(log_file_path, 'a', encoding='utf-8') as log:
                                log.write(
                                    f"[ГОСПОШЛИНА] Ошибка копирования для ФИО: {fio}. "
                                    f"Файл: {pdf_name}, стр. {i+1}, папка: {best_path}, ошибка: {e}\n"
                                )
                    else:
                        print(f"   ➤ ⚠️ Папка клиента не найдена (score < 90) для ФИО: {fio}")
                        not_found_gos.append(fio)
                        with open(log_file_path, 'a', encoding='utf-8') as log:
                            log.write(
                                f"[ГОСПОШЛИНА] Не найдена клиентская папка для ФИО: {fio}. "
                                f"Файл: {pdf_name}, стр. {i+1}\n"
                            )
                else:
                    print("   ➤ ⚠️ ФИО не найдено — пропущено копирование в клиентскую папку.")
                    with open(log_file_path, 'a', encoding='utf-8') as log:
                        log.write(
                            f"[ГОСПОШЛИНА] Не удалось извлечь ФИО. "
                            f"Файл: {pdf_name}, стр. {i+1}\n"
                        )
            
            except Exception as e:
                print(f"⛔ Ошибка на странице {i+1}: {e}")
                with open(log_file_path, 'a', encoding='utf-8') as log:
                    log.write(
                        f"[ГОСПОШЛИНА] Ошибка обработки страницы. "
                        f"Файл: {pdf_name}, стр. {i+1}, ошибка: {e}\n"
                    )
    
    print("\n🎉 Готово.")
    print(f"Госпошлины: найдено {found_gos} из {total_gos} (страницы с распознанным ФИО).")
    
    LOG_SUMMARY["ГОСПОШЛИНЫ"] = {
        "found": found_gos,
        "total": total_gos,
        "not_found": not_found_gos,
    }

if __name__ == "__main__":
    main()

# ## Поиск в СК адреса + Привязка к Суду + Импорт Суда в Delta-M

# -*- coding: utf-8 -*-
"""
Полный объединённый скрипт:

1) Логинится в office.sud.kz
2) Переходит в "Подача документов"
3) Выбирает CIVIL / FIRSTINSTANCE / Иск и нажимает «Отправить»
4) На createRequest.xhtml заполняет поля и открывает форму участника
5) Читает ИИН из Excel, по каждому ИИН тянет адрес и записывает в тот же Excel:
      - ИИН берётся из колонки D;
      - обрабатываются только строки, где в колонке B есть "Vivus";
      - в ту же строку, в колонку O пишется полный адрес (Место жительства);
      - заголовок O1 ставится только если пустой;
      - в колонку P пишется Регион, определённый по адресу из O (заголовок также ставится только если пустой).

   При «залипании» ИИН:
      - нажимает «Закрыть»
      - заново «Добавить участника процесса» → «Далее»
      - снова вводит тот же ИИН (несколько попыток)

   В конце:
      - по тем строкам, где не удалось получить адрес,
        снова заходит в «Подача документов» и ещё раз пытается найти данные;
      - при успехе перезаписывает адрес в колонке O и Регион в колонке P.

6) После Selenium-части:
      - по адресам (O) и регионам (P) заполняет колонку Q
        "Судебный орган с Судебного кабинета" по справочнику
        "Суды по гражданским делам.xlsx" (лист "Возврат"):
          * если продукт (B) НЕ содержит "Vivus" — в Q ставится
            "Медеуский районный суд города Алматы (Гражданские дела)";
          * если продукт (B) содержит "Vivus" — суд подбирается по адресу/региону.

7) Формирует файл ProcessImport.xlsx в сетевой папке \\192.168.1.251\A-Omega
   (структура как в шаблоне ProcessImport):
      - колонка A  "Уникальный номер сделки" — из отчёта "Отчёт по отменам_Omega.xlsx"
        (столбец, в заголовке которого есть слова "Уникальный номер");
      - колонка L  "Судебный орган":
            * если продукт (колонка B) содержит "Vivus" — берётся значение
              из столбца отчёта, где в заголовке есть "Судебный орган"
              (туда скрипт записал данные Судебного органа);
            * иначе — по умолчанию:
              "Медеуский районный суд города Алматы (Гражданские дела)";
      - колонка AD "Тип процесса"    — "1. Упрощённое производство";
      - колонка AE "Статус процесса" — "Рассмотрение дела".

Остальные колонки в ProcessImport остаются пустыми.
"""

import os
import time
import logging
import datetime as _dt
import re
import unicodedata

import pandas as pd
from difflib import SequenceMatcher

from openpyxl import load_workbook, Workbook

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    StaleElementReferenceException,
    ElementNotInteractableException,
)

# ========= КОНФИГ =========
USER_AUTH     = "230240016634"        # твой ИИН/БИН
USER_PASSWORD = "n3y&pAM5mD&4zKZ"     # твой пароль

BASE      = "https://office.sud.kz"
LOGIN_URL = f"{BASE}/index.xhtml"
HOME_URL  = f"{BASE}/form/proceedings/services.xhtml"
SEND_DOCS_URL = f"{BASE}/form/send/index.xhtml"
SEARCH_FALLBACK_URL = f"{BASE}/lawsuit/document.xhtml"   # пока не используем

# Паузы
WAIT   = 1.0
RETRY  = 8

# ========= ПУТИ =========
BASE_DIR   = r"C:\Users\User\Desktop\Документы для подачи Исков"
INPUT_XLSX = os.path.join(BASE_DIR, "Отчёт по отменам_Omega.xlsx")  # отчёт
OUT_XLSX   = INPUT_XLSX                                            # пишем туда же

FILE_PEOPLE = INPUT_XLSX
FILE_COURTS = os.path.join(BASE_DIR, "Суды по гражданским делам.xlsx")

# Папка на сервере
NETWORK_DIR = r"\\192.168.1.251\A-Omega"

# На всякий случай: создаём, если нет (если нет прав/папки — будет ошибка)
os.makedirs(NETWORK_DIR, exist_ok=True)
os.makedirs(BASE_DIR, exist_ok=True)

# Куда сохраняем ProcessImport
PROCESSIMPORT_PATH = os.path.join(NETWORK_DIR, "ProcessImport.xlsx")

# ========= ЛОГИ =========
logging.basicConfig(
    filename="sud_script.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8",
)

def log_step(msg):
    print(msg)
    logging.info(msg)

def log_warn(msg):
    print("WARN:", msg)
    logging.warning(msg)

def log_err(msg):
    print("ERROR:", msg)
    logging.error(msg)

# короткий лог с временем
def log4(msg):
    print(time.strftime("[%H:%M:%S]"), msg, flush=True)
    logging.info(msg)

# ========= DRIVER =========
def init_driver() -> webdriver.Chrome:
    opts = webdriver.ChromeOptions()
    opts.add_argument("--start-maximized")
    opts.page_load_strategy = "eager"

    prefs = {"profile.managed_default_content_settings.images": 2}
    opts.add_experimental_option("prefs", prefs)
    opts.add_argument("--blink-settings=imagesEnabled=false")
    opts.add_argument("--disable-extensions")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--disable-background-networking")
    opts.add_argument("--disable-renderer-backgrounding")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=opts)
    return driver

# ========= ВСПОМОГАТЕЛЬНЫЕ =========
def find_with_fallbacks(driver, variants, desc, tries=8, delay=0.4):
    """Пробует несколько локаторов с коротким ретраем."""
    last = None
    for _ in range(tries):
        for by, sel in variants:
            try:
                el = driver.find_element(by, sel)
                logging.info(f"Нашёл {desc} по {by}={sel}")
                return el
            except Exception as e:
                last = e
        time.sleep(delay)
    raise NoSuchElementException(f"Не удалось найти {desc}. Последняя ошибка: {last}")

# ========= LOGIN =========
def login(driver):
    log_step("Открываю страницу логина…")
    driver.get(LOGIN_URL)
    time.sleep(WAIT)

    # язык — РУС
    try:
        ru = driver.find_elements(By.XPATH, "//a[contains(.,'РУС') and not(contains(@class,'active'))]")
        if ru:
            ru[0].click()
            log_step("Переключил язык на РУС")
            time.sleep(WAIT)
    except Exception:
        log_warn("Не удалось переключить язык на РУС")

    # ЛОГИН
    login_el = find_with_fallbacks(
        driver,
        variants=[
            (By.ID, "j_idt78:auth:xin"),
            (By.XPATH, "//input[contains(@id,':auth:xin')]"),
            (By.NAME, "j_idt78:auth:xin"),
            (By.XPATH, "//input[contains(@placeholder,'ИИН') or contains(@placeholder,'ЖСН') or contains(@placeholder,'ИИН/БИН')]"),
            (By.CSS_SELECTOR, "input[type='email']"),
        ],
        desc="поле ИИН/БИН"
    )
    pass_el = find_with_fallbacks(
        driver,
        variants=[
            (By.ID, "j_idt78:auth:password"),
            (By.XPATH, "//input[contains(@id,':auth:password')]"),
            (By.XPATH, "//input[@type='password']"),
            (By.XPATH, "//input[contains(@placeholder,'Пароль') or contains(@placeholder,'Құпия')]"),
        ],
        desc="поле Пароль"
    )
    submit = find_with_fallbacks(
        driver,
        variants=[
            (By.CSS_SELECTOR, "input.button-primary[type='submit']"),
            (By.XPATH, "//input[@type='submit' and (contains(@value,'Войти') or contains(@class,'button-primary'))]"),
            (By.XPATH, "//button[contains(.,'Войти')]"),
        ],
        desc="кнопка Войти"
    )

    login_el.clear(); login_el.send_keys(USER_AUTH)
    pass_el.clear();  pass_el.send_keys(USER_PASSWORD)
    log_step("Ввёл ИИН/БИН и Пароль")
    submit.click()
    log_step("Нажал 'Войти'")
    time.sleep(WAIT * 2)

# ========= ПЕРЕХОД В «Подача документов» =========
def go_to_send_docs(driver):
    log_step("Открываю главную страницу (плитки)…")
    driver.get(HOME_URL)
    time.sleep(WAIT)

    link = None
    try:
        link = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, '/form/send/index.xhtml')]"))
        )
    except TimeoutException:
        link = None

    if not link:
        try:
            link = driver.find_element(By.XPATH, "//p[normalize-space()='Подача документов']/ancestor::a")
        except Exception:
            try:
                link = driver.find_element(By.XPATH, "//a[.//p[contains(normalize-space(),'Подача документ')]]")
            except Exception:
                link = None

    if link:
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", link)
        except Exception:
            pass
        try:
            link.click()
        except Exception:
            driver.execute_script("arguments[0].click();", link)
        log_step("Кликнул по плитке 'Подача документов'")
    else:
        log_warn("Плитка не найдена, открываю 'Подача документов' по прямому URL")
        driver.get(SEND_DOCS_URL)

    loaded = False
    try:
        WebDriverWait(driver, 10).until(EC.url_contains("/form/send"))
        loaded = True
    except TimeoutException:
        pass

    if not loaded:
        try:
            WebDriverWait(driver, 6).until(
                EC.presence_of_element_located((
                    By.XPATH,
                    "//h1[contains(.,'Подача документов')] | "
                    "//h2[contains(.,'Подача документов')] | "
                    "//form[contains(@action,'/form/send')]"
                ))
            )
            loaded = True
        except TimeoutException:
            pass

    if not loaded:
        log_warn("Не удалось явно подтвердить загрузку страницы 'Подача документов' — продолжаю по текущему состоянию")

# ========= БЛОК 2: выбор CIVIL/FIRSTINSTANCE/Иск =========
def _css_from_locator(locator):
    by, sel = locator
    if by == By.CSS_SELECTOR:
        return sel
    if by == By.ID:
        css_id = sel.replace(":", r"\:")
        return f"#{css_id}"
    return sel

def _select_by_value_robust(driver, locator, value, desc, timeout=20, tries=3):
    w = WebDriverWait(driver, timeout)

    # ждём <select>
    w.until(EC.presence_of_element_located(locator))

    # ждём наличия нужной опции
    def _opt_present(d):
        try:
            el = d.find_element(*locator)
            return any((o.get_attribute("value") or "") == value
                       for o in el.find_elements(By.TAG_NAME, "option"))
        except StaleElementReferenceException:
            return False
    w.until(_opt_present)

    last_err = None
    for _ in range(tries):
        try:
            el = driver.find_element(*locator)
            Select(el).select_by_value(value)
            w.until(lambda d: d.find_element(*locator).get_attribute("value") == value)
            print(f"✔ {desc} → {value}")
            return
        except (StaleElementReferenceException, TimeoutException) as e:
            last_err = e

    # JS-фолбэк
    ok = driver.execute_script("""
        var el = document.querySelector(arguments[0]);
        if(!el) return false;
        var idx = Array.from(el.options).findIndex(o => (o.value||'') === arguments[1]);
        if(idx < 0) return false;
        el.selectedIndex = idx;
        el.dispatchEvent(new Event('change', {bubbles:true}));
        return true;
    """, _css_from_locator(locator), value)

    if not ok:
        raise TimeoutException(f"Не удалось выбрать '{desc}' со значением {value}. Последняя ошибка: {last_err}")
    WebDriverWait(driver, 10).until(lambda d: d.find_element(*locator).get_attribute("value") == value)
    print(f"✔ {desc} → {value} (через JS)")

def send_claim(driver, timeout=20):
    """
    На странице /form/send/index.xhtml:
      1) Тип производства = CIVIL
      2) Инстанция = FIRSTINSTANCE
      3) Тип документа = Иск (3)
      4) Нажать «Отправить»
    """
    w = WebDriverWait(driver, timeout)

    _select_by_value_robust(
        driver, (By.CSS_SELECTOR, "select[id$=':case-type']"),
        "CIVIL", "Тип производства", timeout=timeout
    )

    _select_by_value_robust(
        driver, (By.CSS_SELECTOR, "select[id$=':instance']"),
        "FIRSTINSTANCE", "Инстанция", timeout=timeout
    )

    _select_by_value_robust(
        driver, (By.CSS_SELECTOR, "select[id$=':request']"),
        "3", "Тип документа", timeout=timeout
    )

    submit_locators = [
        (By.XPATH, "//input[@type='submit' and @value='Отправить']"),
        (By.XPATH, "//input[contains(@id,':j_idt') and @type='submit']"),
        (By.XPATH, "//button[normalize-space()='Отправить']"),
    ]
    btn = None
    for loc in submit_locators:
        try:
            btn = w.until(EC.element_to_be_clickable(loc))
            break
        except TimeoutException:
            continue
    if not btn:
        raise TimeoutException("Кнопка «Отправить» не найдена.")

    try:
        btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", btn)
    print("✔ Нажал «Отправить».")

# ========= БЛОК 3: createRequest + форма участника =========
def _wait(drv, cond, t=25):
    return WebDriverWait(drv, t, ignored_exceptions=(StaleElementReferenceException,)).until(cond)

def _select_value(drv, css, value, desc, t=25):
    sel = _wait(drv, EC.presence_of_element_located((By.CSS_SELECTOR, css)), t)
    for _ in range(3):
        try:
            Select(sel).select_by_value(value)
            _wait(drv, lambda d: d.find_element(By.CSS_SELECTOR, css).get_attribute("value") == value, 10)
            print(f"✔ {desc} → value={value}")
            return
        except StaleElementReferenceException:
            sel = drv.find_element(By.CSS_SELECTOR, css)
        except Exception:
            pass
    # JS-фолбэк
    drv.execute_script("""
        var s = document.querySelector(arguments[0]);
        var val = arguments[1];
        if(!s) return false;
        s.value = val;
        s.dispatchEvent(new Event('change', {bubbles:true}));
        return true;
    """, css, value)
    _wait(drv, lambda d: d.find_element(By.CSS_SELECTOR, css).get_attribute("value") == value, 10)
    print(f"✔ {desc} → value={value} (JS)")

def _click(drv, locator, desc, t=25):
    btn = _wait(drv, EC.element_to_be_clickable(locator), t)
    drv.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    try:
        btn.click()
    except Exception:
        drv.execute_script("arguments[0].click();", btn)
    print(f"✔ Нажал {desc}")

def block3_fill_case_and_open_modal(driver, timeout=40):
    w = WebDriverWait(driver, timeout)

    # убеждаемся, что на createRequest.xhtml
    try:
        w.until(
            lambda d: "/form/requestType2/createRequest.xhtml" in d.current_url
            or d.find_element(By.CSS_SELECTOR, "select[id$=':edit-category']")
        )
    except Exception:
        raise TimeoutException("Страница createRequest.xhtml не открыта")

    # 1) Вид производства по делу: value=2
    _select_value(
        driver,
        "select[id$=':edit-categoryGroup']",
        "2",
        "Вид производства по делу",
        t=timeout,
    )

    # 2) Категория дела: value=27
    _wait(
        driver,
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "select[id$=':edit-category'] option[value='27']")
        ),
        30,
    )
    _select_value(
        driver,
        "select[id$=':edit-category']",
        "27",
        "Категория дела",
        t=timeout,
    )

    # 3) Характер заявления: value=1
    _wait(
        driver,
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "select[id$=':edit-character'] option[value='1']")
        ),
        20,
    )
    _select_value(
        driver,
        "select[id$=':edit-character']",
        "1",
        "Характер заявления",
        t=timeout,
    )

    # 4) Галочка «Дело упрощенного производства»
    try:
        cb = w.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "input[id$=':edit-simpleProcess'][type='checkbox']")
            )
        )
        if not cb.is_selected():
            driver.execute_script("arguments[0].click();", cb)
        print("✔ Галочка «Дело упрощенного производства» установлена")
    except Exception:
        print("⚠️ Не удалось установить галочку «Дело упрощенного производства» (продолжаю)")

    # 5) Открыть модал «Добавить участника процесса»
    add_locators = [
        (By.XPATH, "//button[contains(.,'Добавить участника процесса')]"),
        (By.CSS_SELECTOR, "button[id*='addPerson']"),
    ]
    next_xpath = (
        "//div[contains(@class,'modal') or contains(@class,'modal-dialog')]"
        "//input[@type='button' or @type='submit'][@value='Далее']"
    )
    for loc in add_locators:
        try:
            _click(driver, loc, "«Добавить участника процесса»", t=timeout)
            _wait(driver, EC.presence_of_element_located((By.XPATH, next_xpath)), 20)
            break
        except TimeoutException:
            continue
    else:
        raise TimeoutException("Кнопка «Добавить участника процесса» не найдена")

    # 6) В модалке жмём «Далее» и ждём поле ИИН
    person_iin_css = "input[id$=':person-iin']"

    def modal_switched(d):
        try:
            return d.find_element(By.CSS_SELECTOR, person_iin_css)
        except Exception:
            return False

    for attempt in range(3):
        try:
            btn_next = _wait(
                driver,
                EC.element_to_be_clickable((By.XPATH, next_xpath)),
                15,
            )

            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", btn_next
            )

            try:
                btn_next.click()
            except Exception:
                driver.execute_script("arguments[0].click();", btn_next)

            _wait(driver, modal_switched, 25)
            print("✅ Блок 3 завершён: открыта форма участника с полем ИИН.")
            break

        except StaleElementReferenceException:
            print("⚠ Stale element на кнопке «Далее», пробую ещё раз…")
            if attempt == 2:
                print("❌ Не удалось нажать «Далее» — кнопка всё время устаревает")
                raise
            time.sleep(1)

# ========= БЛОК 4: парсинг ИИН из Excel =========
print("=== BLOCK-4 v4 loaded (Vivus only, write address to column O, region to P) ===")

AUTOSAVE_EVERY = 25  # каждые N обработанных ИИН автосейв

# профиль по времени
PERSON_TIMEOUT         = 8
RETRIES_PER_PERSON     = 2
PAUSE_BETWEEN_TRIES    = 0.4
COOLDOWN_BETWEEN_ROWS  = 1.1
MIN_SECONDS_PER_ROW    = 1.2
POLL                   = 0.10
MIN_STABLE_SEC         = 0.35
HARD_WAIT_AFTER_SEARCH = 0.20

# локаторы формы «Добавить участника процесса»
IIN  = (By.CSS_SELECTOR, 'input[id$=":person-iin"]')
SUR  = (By.CSS_SELECTOR, 'input[id$=":person-surname"]')
NAM  = (By.CSS_SELECTOR, 'input[id$=":person-firstname"]')
PATR = (By.CSS_SELECTOR, 'input[id$=":person-patronymic"]')
LIVE = (By.CSS_SELECTOR, 'textarea[id$=":person-livePlace"], input[id$=":person-livePlace"]')

# кнопки/диалоги для перезапуска формы
BTN_CLOSE_PERSON = (
    By.XPATH,
    "//input[@value='Закрыть' and contains(@onclick,'hideFizModalDialog')]"
    " | //input[@id='j_idt278:j_idt328']"
)
BTN_ADD_PERSON = (
    By.XPATH,
    "//button[contains(@onclick,'renderAddPersonModalDialog') or contains(.,'Добавить участника процесса')]"
)
BTN_NEXT_SIDE = (
    By.XPATH,
    "//div[contains(@class,'modal')]//input[@type='button' and @value='Далее'] "
    "| //input[@id='j_idt185:j_idt205']"
)

def _norm_iin(val) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    try:
        if "e" in s.lower():
            s = str(int(float(s)))
    except Exception:
        pass
    s = "".join(ch for ch in s if ch.isdigit())
    return s[:12]

def _get_val(drv, loc) -> str:
    try:
        el = drv.find_element(*loc)
        return (el.get_attribute("value") or "").strip()
    except Exception:
        return ""

def _rf_queue_size(drv):
    try:
        return drv.execute_script("""
            try {
              if (window.RichFaces && RichFaces.Queue && typeof RichFaces.Queue.getSize==='function'){
                return RichFaces.Queue.getSize();
              }
            } catch(e){}
            return 0;
        """) or 0
    except Exception:
        return 0

def _wait_queue_quiet(drv, min_quiet=MIN_STABLE_SEC, total=6):
    t_end = time.time() + total
    quiet_start = None
    last = -1
    while time.time() < t_end:
        q = _rf_queue_size(drv)
        if q != last:
            last = q
            log4(f"   ajax queue size = {q}")
        if q == 0:
            if quiet_start is None:
                quiet_start = time.time()
            if time.time() - quiet_start >= min_quiet:
                return True
        else:
            quiet_start = None
        time.sleep(0.15)
    return False

def _values_snapshot(drv):
    return {
        "sur":  _get_val(drv, SUR),
        "name": _get_val(drv, NAM),
        "patr": _get_val(drv, PATR),
        "live": _get_val(drv, LIVE),
    }

def _stable_after_change(drv, before: dict, timeout=PERSON_TIMEOUT):
    w = WebDriverWait(
        drv,
        timeout,
        poll_frequency=POLL,
        ignored_exceptions=(StaleElementReferenceException,),
    )

    def _changed(_):
        now = _values_snapshot(drv)
        changed = any(now[k] != before.get(k, "") for k in now) and any(now.values())
        return now if changed else False

    try:
        changed_vals = w.until(_changed)
        t0 = time.time()
        last = changed_vals
        while time.time() - t0 < MIN_STABLE_SEC:
            time.sleep(0.15)
            now = _values_snapshot(drv)
            if now != last:
                t0 = time.time()
                last = now
        return last
    except TimeoutException:
        return _values_snapshot(drv)

def _click_magnifier(drv):
    # сначала пробуем js-функцию
    try:
        drv.execute_script("if (typeof fillPersonData==='function'){fillPersonData('j_idt278:person-iin');}")
        log4("   вызвал fillPersonData()")
        return True
    except Exception:
        pass

    # потом ищем иконку-лупу
    try:
        span = drv.find_element(
            By.XPATH,
            "//input[contains(@id,':person-iin')]/following-sibling::span[contains(@class,'gbdSearch')]",
        )
        drv.execute_script("arguments[0].scrollIntoView({block:'center'});", span)
        span.click()
        log4("   клик по иконке-лупе")
        return True
    except Exception:
        try:
            span = drv.find_element(By.CSS_SELECTOR, "span.gbdSearch")
            drv.execute_script("arguments[0].scrollIntoView({block:'center'});", span)
            span.click()
            log4("   клик по .gbdSearch (fallback)")
            return True
        except Exception:
            log4("   ⚠ не смог нажать лупу")
            return False

def _trigger_and_wait(drv) -> dict:
    before = _values_snapshot(drv)
    for attempt in range(1, RETRIES_PER_PERSON + 1):
        log4(f"   попытка #{attempt} — запуск поиска")
        _click_magnifier(drv)
        time.sleep(HARD_WAIT_AFTER_SEARCH)
        _wait_queue_quiet(drv, min_quiet=MIN_STABLE_SEC, total=PERSON_TIMEOUT)
        after = _stable_after_change(drv, before, timeout=PERSON_TIMEOUT)
        changed = any(after[k] != before.get(k, "") for k in after) and any(after.values())
        log4(f"   снэпшот: {after}  | changed={changed}")
        if changed:
            return after
        log4("   ↺ не изменилось — подожду и попробую ещё раз")
        time.sleep(PAUSE_BETWEEN_TRIES)
    return {"sur": "", "name": "", "patr": "", "live": ""}

def _safe_click(drv, loc, desc=""):
    try:
        el = WebDriverWait(drv, 6).until(EC.element_to_be_clickable(loc))
    except TimeoutException:
        try:
            el = drv.find_element(*loc)
        except Exception:
            return False
    try:
        drv.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        el.click()
        return True
    except Exception:
        try:
            drv.execute_script("arguments[0].click();", el)
            return True
        except Exception:
            return False

def _reopen_person_modal(drv):
    """
    Закрыть форму участника, снова нажать «Добавить участника процесса»
    и (если нужно) «Далее» в выборе стороны.
    """
    log4("   ↻ перезапуск формы участника…")

    _safe_click(drv, BTN_CLOSE_PERSON, "Закрыть")
    time.sleep(0.7)

    if not _safe_click(drv, BTN_ADD_PERSON, "Добавить участника"):
        log4("   ⚠ не нашёл кнопку «Добавить участника процесса»")
        return False

    clicked_next = _safe_click(drv, BTN_NEXT_SIDE, "Далее")

    try:
        WebDriverWait(drv, 12).until(EC.presence_of_element_located(IIN))
        if clicked_next:
            log4("   форма участника открыта заново (после выбора стороны)")
        else:
            log4("   форма участника открыта заново (без выбора стороны)")
        return True
    except TimeoutException:
        log4("   ⚠ не дождался поля ИИН после перезапуска")
        return False

def ensure_person_form_open(drv, timeout=15):
    """
    Гарантируем, что открыта форма участника с полем ИИН.
    Если поля нет — пробуем заново открыть модалку.
    """
    try:
        return WebDriverWait(drv, timeout).until(
            EC.presence_of_element_located(IIN)
        )
    except TimeoutException:
        if not _reopen_person_modal(drv):
            raise RuntimeError(
                "Не удалось найти форму участника (поле ИИН). "
                "Возможно, заявка закрылась или сессия слетела."
            )
        return WebDriverWait(drv, timeout).until(
            EC.presence_of_element_located(IIN)
        )

def _safe_save(wb, path):
    try:
        wb.save(path)
        log4(f"💾 autosave to {path}")
        return True
    except Exception as e:
        log4(f"⚠ не удалось сохранить файл: {e}")
        return False

# ===== РЕГИОНЫ И ФУНКЦИЯ ОПРЕДЕЛЕНИЯ РЕГИОНА =====
REGION_NAMES = [
    "Акмолинская область",
    "Актюбинская область",
    "Алматинская область",
    "город Алматы",
    "город Астана",
    "Атырауская область",
    "Восточно-Казахстанская область",
    "Жамбылская область",
    "Западно-Казахстанская область",
    "Карагандинская область",
    "Костанайская область",
    "Кызылординская область",
    "Мангистауская область",
    "Область Абай",
    "Область Жетісу",
    "Область Ұлытау",
    "Павлодарская область",
    "Северо-Казахстанская область",
    "Туркестанская область",
    "город Шымкент",
]

def extract_region_from_address(address: str) -> str:
    """
    Определяет регион по строке адреса.
    - Если встречается 'НУР-СУЛТАН' → 'город Астана'
    - Если 'КАЗАХСТАН, АЛМАТЫ,' или 'Г. АЛМАТЫ' → 'город Алматы'
    - Если 'ШЫМКЕНТ' → 'город Шымкент'
    - Иначе ищет любой регион из REGION_NAMES как подстроку.
    """
    if not address:
        return ""

    up = str(address).upper()

    # Нур-Султан → город Астана
    # Нур-Султан / Астана → город Астана
    if (
        "НУР-СУЛТАН" in up
        or "НУР СУЛТАН" in up
        or "ASTANA" in up
        or "КАЗАХСТАН, АСТАНА," in up
        or " АСТАНА," in up
        or " Г. АСТАНА" in up
    ):
        return "город Астана"

    # Алматы (адрес обычно "КАЗАХСТАН, АЛМАТЫ, ...")
    if "КАЗАХСТАН, АЛМАТЫ," in up or " Г. АЛМАТЫ" in up or ", АЛМАТЫ," in up:
        return "город Алматы"

    # Шымкент, на будущее
    if "КАЗАХСТАН, ШЫМКЕНТ" in up or " Г. ШЫМКЕНТ" in up:
        return "город Шымкент"

    addr_low = up.lower()
    for reg in REGION_NAMES:
        if reg.lower() in addr_low:
            return reg

    return ""

def parse_people_from_excel_and_save_v2(drv):
    """
    Основной цикл по ИИН:
    - ОБРАБАТЫВАЕМ ТОЛЬКО ТЕ СТРОКИ, ГДЕ В КОЛОНКЕ B ЕСТЬ "Vivus" (любым регистром);
    - ИИН берём из колонки D исходного файла;
    - по каждому ИИН вытаскиваем данные с сайта;
    - в колонку O (15) записываем полный адрес (поле 'Место жительства');
    - в колонку P (16) записываем Регион, определённый по адресу.
    """
    MAX_RESTARTS_PER_ROW = 2  # сколько раз полностью перезапускать форму для одного ИИН

    log4("=== START parse_people_from_excel_and_save_v2 ===")
    log4(f"Файл-источник: {INPUT_XLSX}")
    log4(f"Файл-выгрузка (тот же): {OUT_XLSX}")
    log4(f"Паузы: COOLDOWN={COOLDOWN_BETWEEN_ROWS}s, MIN_ROW={MIN_SECONDS_PER_ROW}s, "
         f"AUTOSAVE_EVERY={AUTOSAVE_EVERY}")

    wb = load_workbook(INPUT_XLSX, data_only=True)
    try:
        ws = wb["Отмены"]
    except KeyError:
        # на всякий случай, если лист вдруг будет называться иначе
        ws = wb.active

    # Заголовок колонки O — только если пустой
    header_o = ws.cell(row=1, column=15)
    if not header_o.value or str(header_o.value).strip() == "":
        header_o.value = "Актуальный адрес с СК"

    # Заголовок колонки P — только если пустой
    header_p = ws.cell(row=1, column=16)
    if not header_p.value or str(header_p.value).strip() == "":
        header_p.value = "Регион"

    WebDriverWait(drv, 20).until(EC.presence_of_element_located(IIN))

    failed_rows = []   # номера строк, где не получилось получить адрес
    processed = 0

    try:
        for r in range(2, ws.max_row + 1):
            # --- фильтр по колонке B (только Vivus) ---
            cell_b = ws.cell(row=r, column=2).value
            text_b = (str(cell_b) or "").lower()
            if "vivus" not in text_b:
                continue
            # ------------------------------------------

            # ИИН берём из колонки D (4)
            src_iin = _norm_iin(ws.cell(row=r, column=4).value)
            if not src_iin or len(src_iin) != 12:
                continue

            processed += 1
            t_row_start = time.time()
            log4(f"#{processed} (Excel row {r}) → ИИН: {src_iin}")

            success = False

            # несколько полных перезапусков формы на один ИИН
            for restart in range(MAX_RESTARTS_PER_ROW + 1):
                if restart > 0:
                    log4(f"   ↻ полный перезапуск формы для этого ИИН (рестарт {restart})")
                    if not _reopen_person_modal(drv):
                        log4("   ❌ не удалось перезапустить форму, выхожу из попыток по этому ИИН")
                        break

                try:
                    iin_el = ensure_person_form_open(drv, timeout=20)
                except RuntimeError as e:
                    log4(f"   ❌ форма участника недоступна: {e}")
                    break

                drv.execute_script("arguments[0].scrollIntoView({block:'center'});", iin_el)

                # очищаем поле
                try:
                    iin_el.clear()
                except Exception:
                    try:
                        drv.execute_script("arguments[0].value='';", iin_el)
                    except Exception:
                        pass

                # вводим ИИН
                try:
                    iin_el.send_keys(src_iin)
                except ElementNotInteractableException:
                    log4("   ⚠ элемент ИИН не интерактивен — ставлю значение через JS")
                    drv.execute_script("arguments[0].value = arguments[1];", iin_el, src_iin)

                # триггеры change/blur
                try:
                    drv.execute_script(
                        "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
                        iin_el,
                    )
                    drv.execute_script(
                        "arguments[0].dispatchEvent(new Event('blur',{bubbles:true}));",
                        iin_el,
                    )
                except Exception:
                    pass

                # ждём результат
                before = _values_snapshot(drv)
                data = _trigger_and_wait(drv)
                changed = any(data.values()) and any(data[k] != before.get(k, "") for k in data)
                log4(f"   changed={changed}, data={data}")

                if changed:
                    live = data.get("live", "")
                    # пишем полный адрес в колонку O текущей строки
                    ws.cell(row=r, column=15, value=live)

                    # определяем регион и пишем в колонку P
                    region = extract_region_from_address(live)
                    ws.cell(row=r, column=16, value=region)

                    log4(f"   → SAVE: row {r}, address='{live}', region='{region}'")
                    success = True
                    break
                else:
                    log4("   данных нет / не обновились, пробую перезапустить форму…")

            if not success:
                failed_rows.append(r)
                log4(f"   → SAVE (error): row {r}, address='' (ERROR_NO_DATA)")

            # паузы и очередь ajax
            _wait_queue_quiet(drv, min_quiet=MIN_STABLE_SEC, total=4)
            time.sleep(COOLDOWN_BETWEEN_ROWS)
            elapsed = time.time() - t_row_start
            if elapsed < MIN_SECONDS_PER_ROW:
                time.sleep(MIN_SECONDS_PER_ROW - elapsed)

            if processed % AUTOSAVE_EVERY == 0:
                _safe_save(wb, OUT_XLSX)

        # первичный сейв после первого прохода
        _safe_save(wb, OUT_XLSX)
        log4(f"✅ Готово (первый проход). Сохранено в: {OUT_XLSX}")

        # === ВТОРОЙ ПРОХОД ПО ПРОБЛЕМНЫМ СТРОКАМ ===
        if failed_rows:
            log4(f"=== SECOND PASS: найдено {len(failed_rows)} строк с ERROR_NO_DATA, пробуем ещё раз ===")
            try:
                # заново заходим в Подачу документов и открываем форму участника
                go_to_send_docs(drv)
                send_claim(drv)
                block3_fill_case_and_open_modal(drv)
                WebDriverWait(drv, 20).until(EC.presence_of_element_located(IIN))
            except Exception as e:
                log4(f"⚠ не удалось заново открыть форму участника для второго прохода: {e}")
            else:
                for r in failed_rows:
                    # доп. проверка Vivus (на всякий случай)
                    cell_b = ws.cell(row=r, column=2).value
                    text_b = (str(cell_b) or "").lower()
                    if "vivus" not in text_b:
                        continue

                    src_iin = _norm_iin(ws.cell(row=r, column=4).value)
                    if not src_iin:
                        continue

                    log4(f"[2ND PASS] row {r} → ИИН: {src_iin}")

                    try:
                        iin_el = ensure_person_form_open(drv, timeout=20)
                    except RuntimeError as e:
                        log4(f"   ❌ форма участника недоступна при втором проходе: {e}")
                        break  # дальше смысла нет

                    drv.execute_script("arguments[0].scrollIntoView({block:'center'});", iin_el)

                    # чистим поле
                    try:
                        iin_el.clear()
                    except Exception:
                        try:
                            drv.execute_script("arguments[0].value='';", iin_el)
                        except Exception:
                            pass

                    # вводим ИИН
                    try:
                        iin_el.send_keys(src_iin)
                    except ElementNotInteractableException:
                        log4("   ⚠ [2ND PASS] элемент ИИН не интерактивен — ставлю значение через JS")
                        drv.execute_script("arguments[0].value = arguments[1];", iin_el, src_iin)

                    # триггеры change/blur
                    try:
                        drv.execute_script(
                            "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
                            iin_el,
                        )
                        drv.execute_script(
                            "arguments[0].dispatchEvent(new Event('blur',{bubbles:true}));",
                            iin_el,
                        )
                    except Exception:
                        pass

                    before = _values_snapshot(drv)
                    data = _trigger_and_wait(drv)
                    changed = any(data.values()) and any(data[k] != before.get(k, "") for k in data)
                    log4(f"   [2ND PASS] changed={changed}, data={data}")

                    if changed:
                        live = data.get("live", "")
                        ws.cell(row=r, column=15, value=live)

                        region = extract_region_from_address(live)
                        ws.cell(row=r, column=16, value=region)

                        log4(f"   [2ND PASS] → OVERWRITE row {r}: address='{live}', region='{region}'")
                    else:
                        log4("   [2ND PASS] снова нет данных, оставляю ячейку пустой")

                    _wait_queue_quiet(drv, min_quiet=MIN_STABLE_SEC, total=4)
                    time.sleep(COOLDOWN_BETWEEN_ROWS)

                _safe_save(wb, OUT_XLSX)
                log4(f"✅ Второй проход завершён. Итог сохранён: {OUT_XLSX}")

    finally:
        try:
            _safe_save(wb, OUT_XLSX)
        except Exception:
            pass

    # <<< ДОБАВИТЬ ЭТО >>>
    return processed, failed_rows

# ========= БЛОК 5: логика выбора суда =========
ADDR_COL_NAME   = "Актуальный адрес с Судебного кабинета"       # O
REGION_COL_NAME = "Регион с Судебного кабинета"                 # P
COURT_COL_NAME  = "Судебный орган с Судебного кабинета"         # Q

def normalize_kz_text(s: str) -> str:
    """Унифицируем русские/казахские буквы для сравнения."""
    s = unicodedata.normalize("NFC", str(s))
    s = s.lower()

    replace_map = str.maketrans({
        "қ": "к",
        "ғ": "г",
        "ң": "н",
        "ү": "у",
        "ұ": "у",
        "ө": "о",
        "ә": "а",
        "һ": "х",
        "і": "и",
        "ё": "е",
    })
    s = s.translate(replace_map)
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def norm_region(s: str) -> str:
    """Привести название региона к единому виду."""
    s = normalize_kz_text(s)
    s = s.replace("город ", "").replace("г. ", "").replace("г.", "")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def norm_locality_raw(s: str) -> str:
    """Нормализуем кусок адреса (район/город) для сравнения."""
    s = normalize_kz_text(s)
    s = s.replace("р-н", "район").replace("р н", "район")
    s = re.sub(r"\s+", " ", s)

    remove_words = [
        "район", "аудан", "область", "облысы",
        "город", "қала", "поселок", "посёлок",
        "село", "ауыл"
    ]
    for w in remove_words:
        s = s.replace(w, " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def locality_stem(loc_key: str) -> str:
    """Выделяем корень из локалити: 'самарский' -> 'самар' и т.п."""
    if not loc_key:
        return ""

    s = loc_key
    s = re.sub(r"(ский|ская|ское|скии|скій)$", "", s)
    s = re.sub(r"(ый|ий)$", "", s)

    s = s.strip()
    return s if s else loc_key

def simplify_court_name(name: str) -> str:
    """Упростить название суда до 'ядра' для сравнения."""
    s = normalize_kz_text(name)

    s = re.sub(r"\(.*?\)", " ", s)  # убираем текст в скобках

    remove_words = [
        "районный суд", "районный  суд", "городской суд",
        "район", "района",
        "районный", "городской",
        "суд", "соты",
        "области", "область", "облысы",
        "города",
        "республики казахстан", "рк",
    ]
    for w in remove_words:
        s = s.replace(w, " ")

    s = re.sub(r"\s+", " ", s)
    return s.strip()

def str_similarity(a: str, b: str) -> float:
    """Оценка похожести двух строк (0..100)."""
    a = str(a)
    b = str(b)
    if not a or not b:
        return 0.0
    try:
        from rapidfuzz import fuzz
        return float(fuzz.token_set_ratio(a, b))
    except Exception:
        return SequenceMatcher(None, a, b).ratio() * 100.0

def extract_locality_from_address(address: str):
    """
    'КАЗАХСТАН, <ОБЛАСТЬ>, <РАЙОН/ГОРОД>, ...' -> берём 3-й элемент.
    """
    if pd.isna(address):
        return None
    parts = [p.strip() for p in str(address).split(",") if p.strip()]
    if len(parts) >= 3:
        return parts[2]
    elif len(parts) >= 2:
        return parts[-1]
    else:
        return None

def fill_courts_column():
    """
    Заполняет колонку Q 'Судебный орган с Судебного кабинета' в FILE_PEOPLE
    по файлу FILE_COURTS (лист 'Возврат').

    Если продукт (колонка B) НЕ содержит "Vivus" — в Q ставится
    "Медеуский районный суд города Алматы (Гражданские дела)"
    вне зависимости от адреса/региона.
    """
    print("Читаю файл отчёта:", FILE_PEOPLE)
    df_people = pd.read_excel(FILE_PEOPLE, sheet_name="Отмены")

    print("Столбцы в отчёте:", list(df_people.columns))

    for col in (ADDR_COL_NAME, REGION_COL_NAME, COURT_COL_NAME):
        if col not in df_people.columns:
            raise ValueError(f"В файле нет столбца '{col}'")

    print("Читаю файл судов:", FILE_COURTS)
    df_courts = pd.read_excel(FILE_COURTS, sheet_name="Возврат")

    df_courts["Области"] = df_courts["Области"].ffill()
    df_courts = df_courts.dropna(subset=["Суды"]).copy()
    df_courts["region_key"] = df_courts["Области"].apply(norm_region)
    df_courts["court_key"] = df_courts["Суды"].apply(simplify_court_name)

    courts_by_region = {
        reg_key: grp.reset_index(drop=True)
        for reg_key, grp in df_courts.groupby("region_key")
    }

    def pick_court(region_value, address_value):
        """Выбрать лучший суд по региону и адресу с учётом доп. правил."""
        addr_norm = ""
        if not pd.isna(address_value):
            addr_norm = normalize_kz_text(address_value).upper()

            # Жёсткие правила
            if "КАЗАХСТАН, КОСТАНАЙСКАЯ ОБЛАСТЬ, КОСТАНАЙСКИЙ РАЙОН, ТОБЫЛ" in addr_norm:
                return "Костанайский межрайонный суд Костанайской области"
            if "КАЗАХСТАН, КОСТАНАЙСКАЯ ОБЛАСТЬ, КОСТАНАЙ," in addr_norm:
                return "Костанайский городской суд Костанайской области (Гражданские дела)"
            if "КАЗАХСТАН, АТЫРАУСКАЯ ОБЛАСТЬ, АТЫРАУ" in addr_norm:
                return "Атырауский городской суд Атырауской области (Гражданские дела)"
            if "КАЗАХСТАН, ТУРКЕСТАНСКАЯ ОБЛАСТЬ, САУРАНСКИЙ РАЙОН" in addr_norm:
                return "Кентауский городской суд Туркестанской области (Общая юрисдикция)"
            if "КАЗАХСТАН, ПАВЛОДАРСКАЯ ОБЛАСТЬ, ПАВЛОДАР," in addr_norm:
                return "Межрайонный суд по гражданским делам города Павлодара"
            if "КАЗАХСТАН, ПАВЛОДАРСКАЯ ОБЛАСТЬ, ПАВЛОДАРСКИЙ РАЙОН" in addr_norm:
                return "Межрайонный суд по гражданским делам города Павлодара"
            if "КАЗАХСТАН, КЫЗЫЛОРДИНСКАЯ ОБЛАСТЬ, КЫЗЫЛОРДА" in addr_norm:
                return "Кызылординский городской суд Кызылординской области (Гражданские дела)"
            if "КАЗАХСТАН, КАРАГАНДИНСКАЯ ОБЛАСТЬ, САРАНЬ" in addr_norm:
                return "Саранский городской суд Карагандинской области (Общая юрисдикция)"
            if "КАЗАХСТАН, КОСТАНАЙСКАЯ ОБЛАСТЬ, АРКАЛЫК" in addr_norm:
                return "Аркалыкский городской суд Костанайской области (Общая юрисдикция)"
            if "КАЗАХСТАН, КАРАГАНДИНСКАЯ ОБЛАСТЬ, КАРАГАНДА" in addr_norm:
                return "Необходимо вручную проставить суд! Так как в данном городе несколько судов."
            if "КАЗАХСТАН, АКТЮБИНСКАЯ ОБЛАСТЬ, АКТОБЕ" in addr_norm:
                return "Необходимо вручную проставить суд! Так как в данном городе несколько судов."

        if pd.isna(region_value):
            return None

        reg_key = norm_region(region_value)
        candidates_df = courts_by_region.get(reg_key)
        if candidates_df is None or candidates_df.empty:
            return None

        raw_loc = extract_locality_from_address(address_value)
        if not raw_loc:
            return candidates_df["Суды"].iloc[0]

        loc_key = norm_locality_raw(raw_loc)
        if not loc_key:
            return candidates_df["Суды"].iloc[0]

        stem = locality_stem(loc_key)

        filtered = candidates_df
        if stem:
            mask = filtered["court_key"].apply(
                lambda ck: stem in ck if isinstance(ck, str) else False
            )
            if mask.any():
                filtered = filtered[mask]

        best_score = -1.0
        best_court = None
        for _, row in filtered.iterrows():
            court_name = row["Суды"]
            court_key = row["court_key"]
            score = str_similarity(loc_key, court_key)
            if score > best_score:
                best_score = score
                best_court = court_name

        return best_court

    print("Определяю суды для каждой строки...")
    courts_series = df_people.apply(
        lambda row: pick_court(row[REGION_COL_NAME], row[ADDR_COL_NAME]),
        axis=1
    ).reset_index(drop=True)

    wb = load_workbook(FILE_PEOPLE)

    # --- лист "Отмены" ---
    try:
        ws = wb["Отмены"]
    except KeyError:
        ws = wb.active

    # --- лист "Данные для шаблонов" (может не быть) ---
    try:
        ws_templates = wb["Данные для шаблонов"]
    except KeyError:
        ws_templates = None
        print("⚠ Лист 'Данные для шаблонов' не найден — писать туда не буду.")

    # ищем на листе "Данные для шаблонов" колонку, где в заголовке есть "Судебный орган"
    templates_court_col = None
    if ws_templates is not None:
        for c in range(1, ws_templates.max_column + 1):
            h = ws_templates.cell(row=1, column=c).value
            if not h:
                continue
            if "судебный орган" in str(h).lower():
                templates_court_col = c
                break

        if templates_court_col is None:
            print("⚠ На листе 'Данные для шаблонов' не найден столбец с заголовком, содержащим 'Судебный орган'.")

    # заголовок Q1 в "Отмены"
    ws["Q1"].value = COURT_COL_NAME

    DEFAULT_NON_VIVUS_COURT = "Медеуский районный суд города Алматы (Гражданские дела)"

    # данные начинаются со 2-й строки
    for i, court_name in enumerate(courts_series, start=2):
        # продукт в колонке B на листе "Отмены"
        prod_val = ws.cell(row=i, column=2).value
        prod_text = (str(prod_val) or "").lower()

        if "vivus" not in prod_text:
            court_to_write = DEFAULT_NON_VIVUS_COURT
        else:
            court_to_write = court_name

        # 1) пишем в лист "Отмены", колонка Q
        ws[f"Q{i}"].value = court_to_write

        # 2) параллельно пишем в лист "Данные для шаблонов", колонка "Судебный орган"
        if ws_templates is not None and templates_court_col is not None:
            # предполагаем, что строки совпадают по номеру
            ws_templates.cell(row=i, column=templates_court_col, value=court_to_write)

    wb.save(FILE_PEOPLE)
    print("Готово! Значения судов записаны в:")
    print("  - лист 'Отмены', колонка Q")
    if ws_templates is not None and templates_court_col is not None:
        print("  - лист 'Данные для шаблонов', колонка с заголовком 'Судебный орган'")

# ========= БЛОК 6: формирование ProcessImport.xlsx =========
PROCESS_IMPORT_HEADERS = [
    'Уникальный номер сделки',
    'Дата подачи заявления на выписку ИЛ',
    'Дата получения ИЛ',
    'Дата передачи ИЛ ЧСИ',
    'Комментарии по суду',
    'Дата подачи заявления на выдазу СП',
    'Дата получения СП',
    'Дата передачи СП ЧСИ',
    'Комментарий',
    'Представитель истца',
    'Номер судебного дела',
    'Судебный орган',
    'Категория дела',
    'Сумма иска',
    'Сумма государственной пошлины',
    'Дата отправки искового заявления',
    'Отклонено',
    'Причина отклонения заявления',
    'Зарегистрировано',
    'Судья',
    'Вынесено определение о возврате искового заявления',
    'Вынесено определение об утверждении соглашения об урегулировании спора',
    'Вынесено определение о рассмотрении дела в порядке упрощенного производства',
    'Вынесено решение первой инстанции',
    'Вынесен судебный приказ',
    'Определение об отмене решения в порядке упрощенного производства',
    'Ответственный',
    'Дата создания',
    'Название процесса',
    'Тип процесса',
    'Статус процесса',
]

def create_process_import_file():
    """
    Формирует файл ProcessImport.xlsx в сетевой папке NETWORK_DIR
    по данным из отчёта INPUT_XLSX.

    Логика:
      - берём только строки, где ИИН (колонка D) = 12 цифр;
      - A  "Уникальный номер сделки" — из столбца отчёта,
         в заголовке которого есть 'Уникальный номер';
      - L  "Судебный орган":
            * если продукт (B) содержит 'Vivus', и в строке есть значение
              в столбце, заголовок которого содержит 'Судебный орган' —
              копируем его;
            * иначе — 'Медеуский районный суд города Алматы (Гражданские дела)';
      - AD "Тип процесса"    — '1. Упрощённое производство';
      - AE "Статус процесса" — 'Рассмотрение дела'.
    """
    os.makedirs(NETWORK_DIR, exist_ok=True)

    wb_src = load_workbook(INPUT_XLSX, data_only=True)
    try:
        ws_src = wb_src["Отмены"]
    except KeyError:
        ws_src = wb_src.active

    wb_out = Workbook()
    ws_out = wb_out.active
    ws_out.title = "Sheet1"

    # заголовки
    for col_idx, header in enumerate(PROCESS_IMPORT_HEADERS, start=1):
        ws_out.cell(row=1, column=col_idx, value=header)

    # ищем колонку с "Уникальный номер"
    uniq_col = None
    # и колонку с "Судебный орган"
    court_col = None

    for c in range(1, ws_src.max_column + 1):
        h = ws_src.cell(row=1, column=c).value
        if not h:
            continue
        h_low = str(h).strip().lower()

        if uniq_col is None and "уникальный номер" in h_low:
            uniq_col = c

        if court_col is None and "судебный орган" in h_low:
            court_col = c

    if uniq_col is None:
        log4("⚠ не найден столбец с 'Уникальный номер' в исходном файле.")
    if court_col is None:
        log4("⚠ не найден столбец с 'Судебный орган' в исходном файле, Vivus будет с судом по умолчанию.")

    out_row = 2
    default_court = "Медеуский районный суд города Алматы (Гражданские дела)"

    for r in range(2, ws_src.max_row + 1):
        # ИИН в колонке D — берём только строки с валидным ИИН
        iin_val = _norm_iin(ws_src.cell(row=r, column=4).value)
        if not iin_val or len(iin_val) != 12:
            continue

        # продукт из колонки B
        product_val = ws_src.cell(row=r, column=2).value
        product_text = (str(product_val) or "").lower()

        # "Уникальный номер сделки" из исходного файла
        uniq_val = ws_src.cell(row=r, column=uniq_col).value if uniq_col else None

        # судебный орган из найденного столбца
        court_from_ws = None
        if court_col is not None:
            court_from_ws = ws_src.cell(row=r, column=court_col).value

        court_from_ws_str = (str(court_from_ws).strip()
                             if court_from_ws is not None else "")

        if "vivus" in product_text and court_from_ws_str:
            court_value = court_from_ws_str
        else:
            court_value = default_court

        ws_out.cell(row=out_row, column=1,  value=uniq_val)      # A "Уникальный номер сделки"
        ws_out.cell(row=out_row, column=12, value=court_value)   # L "Судебный орган"
        ws_out.cell(row=out_row, column=30, value="1. Упрощённое производство")  # AD "Тип процесса"
        ws_out.cell(row=out_row, column=31, value="Рассмотрение дела")          # AE "Статус процесса"

        out_row += 1

    wb_out.save(PROCESSIMPORT_PATH)
    log4(f"✅ Файл ProcessImport сформирован: {PROCESSIMPORT_PATH} (строк: {out_row - 2})")

# ========= ORCHESTRATOR =========
def main():
    global LOG_SUMMARY   # чтобы писать в общий сводный лог

    # 1. Selenium: тянем адреса/регионы
    drv = init_driver()
    try:
        login(drv)
        go_to_send_docs(drv)
        send_claim(drv)
        block3_fill_case_and_open_modal(drv)
        # получаем статистику: сколько обработано и какие строки не удалось
        processed, failed_rows = parse_people_from_excel_and_save_v2(drv)
    finally:
        # если нужно, можно раскомментировать закрытие браузера
        # drv.quit()
        pass

    # 2. Заполняем судебный орган в колонку Q
    fill_courts_column()

    # 3. Формируем ProcessImport.xlsx в сетевой папке NETWORK_DIR
    create_process_import_file()

    # 4. Добавляем раздел в сводный LOG_SUMMARY
    LOG_SUMMARY["АДРЕСА И РЕГИОНЫ (VIVUS)"] = {
        "found": processed - len(failed_rows),   # у кого адрес удалось получить
        "total": processed,                      # всего попыток (Vivus-строк)
        "not_found": [f"строка {r}" for r in failed_rows],  # можно заменить на ФИО+ИИН, если захочешь
    }

if __name__ == "__main__":
    main()

# # Досудебная претензия

# -*- coding: utf-8 -*-
import os
from pathlib import Path
from datetime import datetime, date

from openpyxl import load_workbook
from docx import Document
from docx2pdf import convert   # pip install docx2pdf

# ---------------- НАСТРОЙКИ ----------------
EXCEL_PATH   = r"C:\Users\User\Desktop\Документы для подачи Исков\Отчёт по отменам_Omega.xlsx"
SHEET_NAME   = "Данные для шаблонов"

TEMPLATE_DOC = r"C:\Users\User\Desktop\Документы для подачи Исков\Шаблоны документов\Шаблоны по Досудебной претензии.docx"

# Корневая папка, где лежат подпапки по клиентам
ROOT_OUT_FOLDER = r"C:\Users\User\Desktop\Документы для подачи Исков\Папки_для_исков"

# ---------- ЛОГИ (как в первом скрипте) ----------
# общий лог по всем блокам (если не задан в верхней ячейке — задаём здесь)
try:
    log_file_path
except NameError:
    os.makedirs(ROOT_OUT_FOLDER, exist_ok=True)
    log_file_path = os.path.join(ROOT_OUT_FOLDER, "лог_сбор_документов.txt")

# общий словарь сводки
try:
    LOG_SUMMARY
except NameError:
    LOG_SUMMARY = {}

# счётчики конкретно для досудебных претензий (по шаблону)
pret_found = 0        # сколько претензий удалось сформировать
pret_total = 0        # сколько записей обработали из Excel
pret_not_found = []   # список "ФИО, ИИН" для не сформированных / с ошибками

# -------------------------------------------------
# Вспомогательные функции
# -------------------------------------------------
def nice_case(s: str) -> str:
    """Первая буква каждого слова заглавная, остальные строчные."""
    if not s:
        return ""
    parts = str(s).split()
    return " ".join(p.capitalize() for p in parts)

def replace_placeholders_in_paragraph(paragraph, mapping: dict):
    """
    Работаем с целым paragraph.text, затем пересобираем абзац так,
    чтобы плейсхолдеры, разрезанные на несколько run'ов, тоже заменялись.
    При этом формат (шрифт, размер) берём из ПЕРВОГО run абзаца.

    Для абзацев, где были {Email} или {Дата минус 7 дней}, снимаем жирность.
    """
    original_text = paragraph.text
    text = original_text
    changed = False

    for key, value in mapping.items():
        if key in text:
            text = text.replace(key, value)
            changed = True

    if not changed:
        return  # ничего не меняем

    if paragraph.runs:
        # сохраняем формат первого run (шрифт, размер и т.д.)
        paragraph.runs[0].text = text
        for r in paragraph.runs[1:]:
            r.text = ""
        target_run = paragraph.runs[0]
    else:
        target_run = paragraph.add_run(text)

    # для абзацев с Email и Датой минус 7 дней делаем обычный шрифт (не жирный)
    lower_orig = original_text.lower()
    if "{email}" in lower_orig or "{дата минус 7 дней}" in lower_orig:
        target_run.bold = False

def fill_document(row_dict: dict) -> Document:
    doc = Document(TEMPLATE_DOC)

    mapping = {}

    for col_name, val in row_dict.items():
        if not col_name:  # пустой заголовок
            continue

        placeholder = "{" + col_name + "}"
        lower_name = col_name.lower()

        # --- формируем строку значения ---
        if val is None:
            value_str = ""
        else:
            # 1) если это реальный datetime/date → всегда дд.мм.гггг
            if isinstance(val, (datetime, date)):
                value_str = val.strftime("%d.%m.%Y")
            else:
                # приводим к строке
                s = str(val).strip()

                # 2) если в названии колонки есть "дата" → обрезаем время
                if "дата" in lower_name:
                    # отрезаем то, что после пробела (время)
                    # 2025-12-04 09:41:55 → 2025-12-04
                    if " " in s:
                        s = s.split()[0]

                    # 2025-12-04 → 04.12.2025
                    parts = s.replace("/", "-").split("-")
                    if len(parts) == 3 and len(parts[0]) == 4:  # yyyy-mm-dd
                        y, m, d = parts
                        s = f"{d.zfill(2)}.{m.zfill(2)}.{y}"

                    value_str = s

                # 3) ФИО/ЧСИ → красивый регистр
                elif "фио" in lower_name or "чси" in lower_name:
                    value_str = nice_case(s)
                else:
                    value_str = s

        mapping[placeholder] = value_str

    # Дополнительно текущая дата
    mapping["{CurrDate}"] = datetime.now().strftime("%d.%m.%Y")

    # дальше всё как было...
    for p in doc.paragraphs:
        replace_placeholders_in_paragraph(p, mapping)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_placeholders_in_paragraph(p, mapping)

    return doc

def find_debtor_folder(root_folder: Path, iin: str) -> Path:
    """
    Ищем подпапку в root_folder, в имени которой встречается ИИН.
    Если не нашли — создаём новую папку с именем ИИН.
    """
    for sub in root_folder.iterdir():
        if sub.is_dir() and iin in sub.name:
            return sub

    new_folder = root_folder / iin
    new_folder.mkdir(parents=True, exist_ok=True)
    return new_folder

# -------------------------------------------------
# Основная логика
# -------------------------------------------------
def main():
    global pret_found, pret_total, pret_not_found, LOG_SUMMARY

    root_path = Path(ROOT_OUT_FOLDER)
    root_path.mkdir(parents=True, exist_ok=True)

    wb = load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    ws = wb[SHEET_NAME]

    # Заголовки колонок (первая строка)
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

    # Обход строк, начиная со 2-й
    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(v is None for v in row):
            continue

        row_dict = dict(zip(headers, row))

        # ФИО - столбец C (индекс 2), ИИН - столбец D (индекс 3)
        fio_value = row[2] if len(row) > 2 else None
        iin_value = row[3] if len(row) > 3 else None

        fio = nice_case(str(fio_value)) if fio_value is not None else ""
        iin = str(iin_value).strip() if iin_value is not None else ""

        if not fio or not iin:
            continue

        pret_total += 1
        print("\n=======================================")
        print(f"[STEP DOCX] {pret_total} — {fio}, {iin}")

        try:
            # Обновляем ФИО/ИИН в словаре, чтобы плейсхолдеры с любыми названиями колонок,
            # содержащими "фио" / "иин", брали уже нормализованные значения
            for key in list(row_dict.keys()):
                if key and "фио" in key.lower():
                    row_dict[key] = fio
                if key and "иин" in key.lower():
                    row_dict[key] = iin

            # 1) Собираем docx по шаблону
            doc = fill_document(row_dict)

            # Находим/создаём папку должника по ИИН
            debtor_folder = find_debtor_folder(root_path, iin)

            base_name = f"Досудебная претензия, {fio}, {iin}"
            docx_path = debtor_folder / (base_name + ".docx")
            pdf_path  = debtor_folder / (base_name + ".pdf")

            doc.save(docx_path)

            # 2) Конвертируем в PDF
            convert(str(docx_path), str(pdf_path))

            # 3) Удаляем промежуточный DOCX, если не нужен
            try:
                os.remove(docx_path)
            except OSError:
                pass

            print(f"Сформирован файл: {pdf_path}")
            pret_found += 1

        except Exception as e:
            label = f"{fio or 'Неизвестный'}, {iin or 'ИИН не указан'}"
            pret_not_found.append(label)
            print(f"[ERROR DOCX] Ошибка при формировании для {label}: {e}")
            try:
                with open(log_file_path, "a", encoding="utf-8") as lf:
                    lf.write(f"[ДОСУДЕБНАЯ ПРЕТЕНЗИЯ DOCX] Ошибка для {label}: {e}\n")
            except Exception:
                pass
            continue

    wb.close()

    print("\n=== ГОТОВО: все записи обработаны (DOCX) ===")
    print(f"[ИТОГ ДОСУДЕБНЫЕ ПРЕТЕНЗИИ DOCX] сформировано {pret_found} из {pret_total}")

    # записываем в общий сводный словарь
    LOG_SUMMARY["ДОСУДЕБНЫЕ ПРЕТЕНЗИИ DOCX"] = {
        "found": pret_found,
        "total": pret_total,
        "not_found": pret_not_found,
    }

    print("Готово.")

if __name__ == "__main__":
    main()

# # Скрин отправки Досудебной претензии

import os
import random
from pathlib import Path
from datetime import datetime

import pandas as pd
from jinja2 import Template

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

from PIL import Image

# ---------- НАСТРОЙКИ ----------

EXCEL_PATH = r"C:\Users\User\Desktop\Документы для подачи Исков\Отчёт по отменам_Omega.xlsx"
SHEET_NAME = "Данные для шаблонов"

TEMPLATE_PATH = Path(
    r"C:\Users\User\Desktop\Документы для подачи Исков\Шаблоны документов\Скрин (Досудебная претензия).html"
)

ROOT_OUT_FOLDER = Path(
    r"C:\Users\User\Desktop\Документы для подачи Исков\Папки_для_исков"
)

# ---------- ЛОГИ ----------

try:
    log_file_path
except NameError:
    ROOT_OUT_FOLDER.mkdir(parents=True, exist_ok=True)
    log_file_path = ROOT_OUT_FOLDER / "лог_сбор_документов.txt"

try:
    LOG_SUMMARY
except NameError:
    LOG_SUMMARY = {}

screen_found = 0
screen_total = 0
screen_not_found = []

# ---------- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ----------

def nice_case(s: str) -> str:
    if not s:
        return ""
    parts = str(s).split()
    return " ".join(p.capitalize() for p in parts)

def load_template():
    text = TEMPLATE_PATH.read_text(encoding="utf-8")
    return Template(text)

def html_to_jpeg(html_path: Path, jpeg_path: Path):
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1200,1600")
    options.add_argument("--hide-scrollbars")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )

    try:
        url = html_path.absolute().as_uri()
        driver.get(url)
        driver.implicitly_wait(2)

        tmp_png = jpeg_path.with_suffix(".png")
        driver.save_screenshot(str(tmp_png))
    finally:
        driver.quit()

    img = Image.open(tmp_png).convert("RGB")
    img.save(jpeg_path, "JPEG", quality=95)
    tmp_png.unlink(missing_ok=True)

def random_time_str():
    hour = random.randint(9, 17)
    minute = random.randint(0, 59)
    return f"{hour:02d}:{minute:02d}"

def find_debtor_folder(root_folder: Path, iin: str):
    for sub in root_folder.iterdir():
        if sub.is_dir() and iin in sub.name:
            return sub
    return None

def get_first_existing_value(row: pd.Series, candidates: list[str], default: str = "") -> str:
    """Берёт первое непустое значение из списка возможных названий колонок."""
    for col in candidates:
        if col in row.index:
            val = row.get(col)
            if val is not None:
                s = str(val).strip()
                if s and s.lower() != "nan":
                    return s
    return default

# ---------- ОСНОВНОЙ КОД ----------

def main():
    global screen_found, screen_total, screen_not_found, LOG_SUMMARY

    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
    template = load_template()

    for idx, row in df.iterrows():
        fio = ""
        iin = ""

        try:
            fio_raw = row.get("ФИО")
            iin_raw = row.get("ИИН")

            fio = nice_case(str(fio_raw)) if fio_raw is not None else ""

            # --- ИИН: всегда 12 символов, с нулями слева ---
            if iin_raw is not None:
                raw_iin_str = str(iin_raw).strip()
                if "." in raw_iin_str:
                    raw_iin_str = raw_iin_str.split(".")[0]
                iin = raw_iin_str.zfill(12)
            else:
                iin = ""

            if not fio or not iin:
                continue

            screen_total += 1
            print("\n=======================================")
            print(f"[STEP SCREEN] {screen_total} — {fio}, {iin}")

            debtor_folder = find_debtor_folder(ROOT_OUT_FOLDER, iin)
            if debtor_folder is None:
                label = f"{fio}, {iin}"
                screen_not_found.append(label + " (папка не найдена)")
                print(f"[SKIP] Папка для {label} не найдена в {ROOT_OUT_FOLDER}")
                with open(log_file_path, "a", encoding="utf-8") as lf:
                    lf.write(f"[СКРИН ДП] Папка не найдена для {label}\n")
                continue

            email = str(row.get("Email", ""))

            loan_date = row.get("Дата выдачи займа")
            if isinstance(loan_date, (datetime, pd.Timestamp)) and pd.notna(loan_date):
                loan_date_str = loan_date.strftime("%d.%m.%Y")
            else:
                loan_date_str = str(loan_date) if loan_date is not None and pd.notna(loan_date) else ""

            cession_date = row.get("Дата договора цессии")
            if isinstance(cession_date, (datetime, pd.Timestamp)) and pd.notna(cession_date):
                cession_date_str = cession_date.strftime("%d.%m.%Y")
            else:
                cession_date_str = str(cession_date) if cession_date is not None and pd.notna(cession_date) else ""

            email_date = row.get("Дата минус 7 дней")
            dt = None
            if isinstance(email_date, (datetime, pd.Timestamp)) and pd.notna(email_date):
                dt = email_date
            else:
                try:
                    parsed = pd.to_datetime(str(email_date))
                    dt = parsed if pd.notna(parsed) else None
                except Exception:
                    dt = None

            if dt is not None:
                email_date_str = dt.strftime("%d.%m.%Y")
            else:
                email_date_str = str(email_date) if email_date is not None else ""

            email_datetime_str = f"{email_date_str} {random_time_str()}"

            # --- КРЕДИТОР: берём из Excel ---
            creditor = get_first_existing_value(
                row,
                candidates=[
                    "Кредитор",
                    "Кредитор (наименование)",
                    "Первоначальный кредитор",
                    "МФО",
                    "Микрофинансовая организация",
                ],
                default=""
            )

            data = {
                "fio": fio,
                "iin": iin,
                "mobile": str(row.get("Мобильный", "")),
                "loan_number": str(row.get("Номер договора займа", "")),
                "loan_date": loan_date_str,
                "loan_amount": str(row.get("Сумма выдачи займа", "")),
                "cession_number": str(row.get("Номер договора цессии", "")),
                "cession_date": cession_date_str,
                "debt_amount": str(row.get("Остаток задолженности", "")),
                "email": email,
                "email_datetime": email_datetime_str,

                # ✅ вот это важно для шаблона {{ creditor }}
                "creditor": creditor,
            }

            base_name = f"Скрин отправки Досудебной претензии, {fio}, {iin}"
            html_name = debtor_folder / (base_name + ".html")
            img_name = debtor_folder / (base_name + ".jpg")

            html_text = template.render(**data)
            html_name.write_text(html_text, encoding="utf-8")

            html_to_jpeg(html_name, img_name)

            try:
                html_name.unlink()
            except OSError:
                pass

            print(f"[OK] {fio} ({iin}) -> {img_name}")
            screen_found += 1

        except Exception as e:
            label = f"{fio or 'Неизвестный'}, {iin or 'ИИН не указан'}"
            screen_not_found.append(label)
            print(f"[ERROR] строка {idx}: {e}")
            try:
                with open(log_file_path, "a", encoding="utf-8") as lf:
                    lf.write(f"[СКРИН ДП] Ошибка для {label}: {e}\n")
            except Exception:
                pass
            continue

    print("\n=== ГОТОВО: все записи обработаны (СКРИНЫ) ===")
    print(f"[ИТОГ СКРИНЫ ДП] сформировано {screen_found} из {screen_total}")

    LOG_SUMMARY["СКРИНЫ ДОСУДЕБНЫЕ ПРЕТЕНЗИИ"] = {
        "found": screen_found,
        "total": screen_total,
        "not_found": screen_not_found,
    }

    print("Готово.")

if __name__ == "__main__":
    main()

# # Расчет задолженности

# -*- coding: utf-8 -*-
import os
from pathlib import Path
from datetime import datetime, date

from openpyxl import load_workbook
from docx import Document
from docx2pdf import convert

# ---------------- НАСТРОЙКИ ----------------
EXCEL_PATH   = r"C:\Users\User\Desktop\Документы для подачи Исков\Отчёт по отменам_Omega.xlsx"
SHEET_NAME   = "Данные для шаблонов"

TEMPLATE_DOC = r"C:\Users\User\Desktop\Документы для подачи Исков\Шаблоны документов\Шаблоны по Расчетам задолженности.docx"

ROOT_OUT_FOLDER = r"C:\Users\User\Desktop\Документы для подачи Исков\Папки_для_исков"

# ---------- ЛОГИ ----------
try:
    log_file_path
except NameError:
    os.makedirs(ROOT_OUT_FOLDER, exist_ok=True)
    log_file_path = os.path.join(ROOT_OUT_FOLDER, "лог_сбор_документов.txt")

# счётчики для расчёта задолженности
calc_found = 0
calc_total = 0
calc_not_found = []

# -------------------------------------------------
# Вспомогательные функции
# -------------------------------------------------
def nice_case(s: str) -> str:
    if not s:
        return ""
    return " ".join(word.capitalize() for word in str(s).split())

def replace_placeholders_in_paragraph(paragraph, mapping: dict):
    original_text = paragraph.text
    text = original_text
    changed = False

    for key, value in mapping.items():
        if key in text:
            text = text.replace(key, value)
            changed = True

    if not changed:
        return

    if paragraph.runs:
        paragraph.runs[0].text = text
        for r in paragraph.runs[1:]:
            r.text = ""
        target_run = paragraph.runs[0]
    else:
        target_run = paragraph.add_run(text)

def fill_document(row_dict: dict) -> Document:
    doc = Document(TEMPLATE_DOC)
    mapping = {}

    for col_name, val in row_dict.items():
        if not col_name:
            continue

        placeholder = "{" + col_name + "}"
        lower = col_name.lower()

        if val is None or val != val:
            value_str = ""
        else:
            if isinstance(val, (datetime, date)):
                value_str = val.strftime("%d.%m.%Y")
            else:
                if "фио" in lower:
                    value_str = nice_case(val)
                else:
                    value_str = str(val)

        mapping[placeholder] = value_str

    mapping["{CurrDate}"] = datetime.now().strftime("%d.%m.%Y")

    for p in doc.paragraphs:
        replace_placeholders_in_paragraph(p, mapping)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_placeholders_in_paragraph(p, mapping)

    return doc

def find_debtor_folder(root_folder: Path, iin: str) -> Path:
    for sub in root_folder.iterdir():
        if sub.is_dir() and iin in sub.name:
            return sub

    new_folder = root_folder / iin
    new_folder.mkdir(parents=True, exist_ok=True)
    return new_folder

# -------------------------------------------------
# Основная логика
# -------------------------------------------------
def main():
    global calc_found, calc_total, calc_not_found, LOG_SUMMARY

    root_path = Path(ROOT_OUT_FOLDER)
    root_path.mkdir(parents=True, exist_ok=True)

    wb = load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    ws = wb[SHEET_NAME]

    headers = [str(c.value).strip() if c.value else "" for c in ws[1]]

    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(v is None for v in row):
            continue

        row_dict = dict(zip(headers, row))

        fio = nice_case(row[2]) if len(row) > 2 else ""
        iin = str(row[3]).strip() if len(row) > 3 else ""

        if not fio or not iin:
            continue

        calc_total += 1
        print(f"\n[STEP CALC] {calc_total} — {fio}, {iin}")

        try:
            for key in row_dict.keys():
                if "фио" in key.lower():
                    row_dict[key] = fio
                if "иин" in key.lower():
                    row_dict[key] = iin

            doc = fill_document(row_dict)

            debtor_folder = find_debtor_folder(root_path, iin)

            base = f"Расчёт задолженности, {fio}, {iin}"
            docx_path = debtor_folder / (base + ".docx")
            pdf_path  = debtor_folder / (base + ".pdf")

            doc.save(docx_path)
            convert(str(docx_path), str(pdf_path))

            try:
                os.remove(docx_path)
            except OSError:
                pass

            print(f"Сформирован файл: {pdf_path}")
            calc_found += 1

        except Exception as e:
            label = f"{fio}, {iin}"
            calc_not_found.append(label)

            print(f"[ERROR CALC] Ошибка для {label}: {e}")
            try:
                with open(log_file_path, "a", encoding="utf-8") as lf:
                    lf.write(f"[РАСЧЁТ ЗАДОЛЖЕННОСТИ] Ошибка для {label}: {e}\n")
            except:
                pass
            continue

    wb.close()

    print("\n=== ГОТОВО: расчёты задолженности сформированы ===")
    print(f"[ИТОГ] {calc_found} из {calc_total}")

    LOG_SUMMARY["РАСЧЁТ ЗАДОЛЖЕННОСТИ"] = {
        "found": calc_found,
        "total": calc_total,
        "not_found": calc_not_found,
    }

if __name__ == "__main__":
    main()

# # Исковое заявление

# -*- coding: utf-8 -*-
import os
from pathlib import Path
from datetime import datetime, date

from openpyxl import load_workbook
from docx import Document

# ---------------- НАСТРОЙКИ ----------------
EXCEL_PATH   = r"C:\Users\User\Desktop\Документы для подачи Исков\Отчёт по отменам_Omega.xlsx"
SHEET_NAME   = "Данные для шаблонов"

# Шаблон ИСКОВОГО ЗАЯВЛЕНИЯ (.docx!)
TEMPLATE_DOC_ISK = r"C:\Users\User\Desktop\Документы для подачи Исков\Шаблоны документов\Шаблоны по Исковому заявлению в суд.docx"

ROOT_OUT_FOLDER = r"C:\Users\User\Desktop\Документы для подачи Исков\Папки_для_исков"

# ---------- ЛОГИ ----------
try:
    log_file_path
except NameError:
    os.makedirs(ROOT_OUT_FOLDER, exist_ok=True)
    log_file_path = os.path.join(ROOT_OUT_FOLDER, "лог_сбор_документов.txt")

try:
    LOG_SUMMARY
except NameError:
    LOG_SUMMARY = {}

# счётчики для исковых заявлений по шаблону
isk_found = 0
isk_total = 0
isk_not_found = []

# -------------------------------------------------
# Вспомогательные функции
# -------------------------------------------------
def nice_case(s: str) -> str:
    """Приводим ФИО к виду 'Фамилия Имя Отчество'."""
    if not s:
        return ""
    return " ".join(word.capitalize() for word in str(s).split())

def replace_placeholders_in_paragraph(paragraph, mapping: dict):
    """
    Заменяет плейсхолдеры в абзаце, стараясь сохранить форматирование.
    Работает и когда плейсхолдер разбит на несколько run’ов вида:
    '{', 'ОД+проценты+штрафы', '} тенге.'
    """

    # 1) простой случай — плейсхолдер целиком внутри одного run
    for run in paragraph.runs:
        text = run.text
        changed = False
        for key, value in mapping.items():
            if key in text:
                text = text.replace(key, value)
                changed = True
        if changed:
            run.text = text

    # 2) случай, когда плейсхолдер растянут на несколько run’ов
    runs = paragraph.runs
    i = 0
    while i < len(runs):
        replaced_here = False

        for key, value in mapping.items():
            L = len(key)
            accum = ""
            j = i

            # копим текст из нескольких run’ов, пока не увидим '}' или сильно не выйдем за длину ключа
            while j < len(runs) and len(accum) < L + 10 and "}" not in accum:
                accum += runs[j].text
                j += 1

            idx = accum.find(key)
            if idx == 0:
                # плейсхолдер начинается сразу с первого символа накопленного текста
                tail = accum[L:]          # хвост после плейсхолдера (например, " тенге.")
                new_text = value + tail   # подставляем значение и хвост

                runs[i].text = new_text
                # остальные run’ы, которые участвовали в накоплении, очищаем
                for k in range(i + 1, j):
                    runs[k].text = ""
                replaced_here = True
                break

        if not replaced_here:
            i += 1

def fill_document_isk(row_dict: dict) -> Document:
    """Заполнение шаблона искового заявления по данным одной строки."""
    doc = Document(TEMPLATE_DOC_ISK)
    mapping = {}

    for col_name, val in row_dict.items():
        if not col_name:
            continue

        placeholder = "{" + col_name + "}"
        lower = col_name.lower()

        if val is None or val != val:
            value_str = ""
        else:
            if isinstance(val, (datetime, date)):
                value_str = val.strftime("%d.%m.%Y")
            else:
                if "фио" in lower:
                    value_str = nice_case(val)
                else:
                    value_str = str(val)

        mapping[placeholder] = value_str

    # текущая дата
    mapping["{CurrDate}"] = datetime.now().strftime("%d.%m.%Y")

    # абзацы
    for p in doc.paragraphs:
        replace_placeholders_in_paragraph(p, mapping)

    # таблицы
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_placeholders_in_paragraph(p, mapping)

    return doc

def find_debtor_folder(root_folder: Path, iin: str) -> Path:
    """
    Ищем уже существующую папку, в имени которой есть ИИН.
    Если нет — создаём новую папку с именем ИИН.
    """
    for sub in root_folder.iterdir():
        if sub.is_dir() and iin in sub.name:
            return sub

    new_folder = root_folder / iin
    new_folder.mkdir(parents=True, exist_ok=True)
    return new_folder

# -------------------------------------------------
# Основная логика
# -------------------------------------------------
def main():
    global isk_found, isk_total, isk_not_found, LOG_SUMMARY

    root_path = Path(ROOT_OUT_FOLDER)
    root_path.mkdir(parents=True, exist_ok=True)

    wb = load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    ws = wb[SHEET_NAME]

    # Заголовки (первая строка)
    headers = [str(c.value).strip() if c.value else "" for c in ws[1]]

    # Идём по строкам, начиная со 2-й
    for row in ws.iter_rows(min_row=2, values_only=True):
        # пустая строка — пропускаем
        if all(v is None for v in row):
            continue

        # ФИЛЬТР ПО VIVUS ИЗ КОЛОНКИ B
        prod_value = row[1] if len(row) > 1 else None  # B-колонка
        if not (prod_value and "vivus" in str(prod_value).lower()):
            continue

        row_dict = dict(zip(headers, row))
        # предполагаем, что ФИО в колонке C (индекс 2), ИИН в колонке D (индекс 3)
        fio = nice_case(row[2]) if len(row) > 2 else ""
        iin = str(row[3]).strip() if len(row) > 3 else ""

        if not fio or not iin:
            continue

        isk_total += 1
        print(f"\n[STEP ISK] {isk_total} — {fio}, {iin}")

        try:
            # нормализуем ФИО/ИИН во всех колонках, где они упоминаются
            for key in row_dict.keys():
                if "фио" in key.lower():
                    row_dict[key] = fio
                if "иин" in key.lower():
                    row_dict[key] = iin

            doc = fill_document_isk(row_dict)

            debtor_folder = find_debtor_folder(root_path, iin)

            base = f"Исковое заявление, {fio}, {iin}"
            docx_path = debtor_folder / (base + ".docx")

            doc.save(docx_path)

            print(f"Сформирован файл: {docx_path}")
            isk_found += 1

        except Exception as e:

            label = f"{fio}, {iin}"
            isk_not_found.append(label)

            print(f"[ERROR ISK] Ошибка для {label}: {e}")
            try:

                with open(log_file_path, "a", encoding="utf-8") as lf:
                    lf.write(f"[ИСКОВОЕ ЗАЯВЛЕНИЕ ПО ШАБЛОНУ] Ошибка для {label}: {e}\n")
            except:
                pass
            continue

    wb.close()

    print("\n=== ГОТОВО: исковые заявления по шаблону сформированы ===")
    print(f"[ИТОГ ИСК] {isk_found} из {isk_total}")

    LOG_SUMMARY["ИСКОВОЕ ЗАЯВЛЕНИЕ ПО ШАБЛОНУ"] = {
        "found": isk_found,
        "total": isk_total,
        "not_found": isk_not_found,
    }

if __name__ == "__main__":
    main()

# # Постановление об отмене надписи

# -*- coding: utf-8 -*-
import os
import re
import time
import shutil
import pandas as pd
from pathlib import Path
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException

from pypdf import PdfReader

# ================= НАСТРОЙКИ =================
AISOIP_URL = "https://aisoip.adilet.gov.kz/cabinet/exec-productions"
LOGIN = "810813301334_230240016634"
PASSWORD = "Qazaq123456*"

EXCEL_PATH = r"C:\Users\User\Desktop\Документы для подачи Исков\Отчёт по отменам_Omega.xlsx"
IIN_COLUMN_INDEX = 3  # колонка D (0-индексация)

BASE_FOLDER = r"C:\Users\User\Desktop\Документы для подачи Исков\Папки_для_исков"
DOWNLOAD_DIR = str(Path(os.getenv("LOCALAPPDATA", r"C:\Users\User\AppData\Local")) / "Temp" / "AISOIP_downloads")
DEBUG_DIR   = str(Path(os.getenv("LOCALAPPDATA", r"C:\Users\User\AppData\Local")) / "Temp" / "AISOIP_debug")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(DEBUG_DIR, exist_ok=True)

log_file_path = "log_not_copied.txt"

# если сводный словарь ещё не создан – создаём
try:
    LOG_SUMMARY
except NameError:
    LOG_SUMMARY = {}

EXCLUDE_TITLES = [
    "Постановление о прекращении исполнительного производства",
    "Отчет о доставке", "Отчёт о доставке",
    "Извещение",
    "Инкассовое распоряжение",

    "Постановление о запрете должнику совершать определенные действия",
    "Постановление об истребовании информации",
    "Извещение должника о временном ограничении на выезд из РК",
    "Постановление о возбуждении исполнительного производства",
    "Постановление о наложении ареста на транспортные средства",
]

RE_CANCEL_LIST = [
    re.compile(r"\bоб\s+отмен[еёы]\b.{0,60}\bисполн\w*\s+надпис\w*", re.IGNORECASE),
    re.compile(r"\bотменить\b.{0,60}\bисполн\w*\s+надпис\w*", re.IGNORECASE),
    re.compile(r"\bпризнать\b.{0,60}\bисполн\w*\s+надпис\w*\s+недействител\w*", re.IGNORECASE),
    re.compile(r"\bоб\s+отмен[еёы]\s+ин\b", re.IGNORECASE),
]
RE_NEGATION = re.compile(
    r"(отказать(?:\s+в\s+удовлетворении)?|в\s+удовлетворении\s+.*?отказать|оставить\s+без\s+изменения)",
    re.IGNORECASE
)
RE_IIN = re.compile(r"(?<!\d)(\d{12})(?!\d)")

ALLOWED_EXTS = {".pdf", ".jpg", ".jpeg", ".png"}

# ================= ЛОГИ / ОТЛАДКА =================
_t0 = time.time()
def log(msg):
    dt = time.time() - _t0
    print(f"[{dt:7.2f}s] {msg}")

def dbg_snapshot(driver, tag):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    png = os.path.join(DEBUG_DIR, f"{ts}_{tag}.png")
    html = os.path.join(DEBUG_DIR, f"{ts}_{tag}.html")
    try:
        driver.save_screenshot(png)
        with open(html, "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        log(f"[DBG] Снимки сохранены: {png} | {html}")
    except Exception as e:
        log(f"[DBG] Не удалось сохранить снимки: {e}")

# ================= ВСПОМОГАТЕЛЬНЫЕ =================
def make_driver():
    opts = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
        "safebrowsing.enabled": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "download.directory_upgrade": True,
        # чуть ускоряем: отключаем изображения на странице логина, это не мешает
        "profile.managed_default_content_settings.images": 2,
    }
    opts.add_experimental_option("prefs", prefs)
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--start-maximized")
    service = Service(ChromeDriverManager().install())  # обычный ChromeDriverManager
    drv = webdriver.Chrome(service=service, options=opts)
    log("[INIT] Запущен Chrome WebDriver")
    return drv

def wait_new_file(before_files, timeout=25):
    t0 = time.time()
    while time.time() - t0 < timeout:
        after = set(os.listdir(DOWNLOAD_DIR))
        new_files = list(after - before_files)
        if new_files:
            files = [Path(DOWNLOAD_DIR) / nf for nf in new_files]
            files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
            p = files[0]
            if p.suffix.lower() == ".crdownload":
                time.sleep(0.4)
                continue
            return str(p)
        time.sleep(0.4)
    return None

def is_supported_file(path: str) -> bool:
    """Поддерживаем PDF и JPG/PNG."""
    ext = Path(path).suffix.lower()
    if ext not in ALLOWED_EXTS:
        return False

    # Для PDF дополнительно проверяем сигнатуру
    if ext == ".pdf":
        try:
            with open(path, "rb") as f:
                return f.read(5).startswith(b"%PDF")
        except Exception:
            return False

    # Для изображений достаточно расширения
    return True

def is_cancel_doc(path: str) -> bool:
    """Проверка, что PDF содержит текст об отмене ИН."""
    try:
        if not is_supported_file(path):
            return False

        # Обрабатываем только PDF, картинки здесь не трогаем
        if Path(path).suffix.lower() != ".pdf":
            return False

        reader = PdfReader(path)
        text_chunks = []
        for page in reader.pages[:6]:
            try:
                text_chunks.append(page.extract_text() or "")
            except Exception:
                pass
        raw = "\n".join(text_chunks)
        if not raw.strip():
            return False

        txt = raw.replace("Ё", "Е").replace("ё", "е")
        txt = re.sub(r"\s+", " ", txt).strip().lower()

        for ex in (e.lower() for e in EXCLUDE_TITLES):
            if ex in txt:
                return False

        for pat in RE_CANCEL_LIST:
            for m in pat.finditer(txt):
                start = max(0, m.start() - 120)
                end   = min(len(txt), m.end() + 120)
                window = txt[start:end]
                if RE_NEGATION.search(window):
                    continue
                if "постановлен" in window or "решени" in window or "определени" in window:
                    return True
                return True
        return False
    except Exception as e:
        log(f"[PDF] Не удалось распарсить {path}: {e}")
        return False

def find_client_folder_by_iin(iin: str):
    for name in os.listdir(BASE_FOLDER):
        full = os.path.join(BASE_FOLDER, name)
        if os.path.isdir(full) and iin in name:
            return full
    return None

def get_fio_from_folder(iin: str) -> str:
    folder = find_client_folder_by_iin(iin)
    if not folder:
        return "Неизвестный"
    base = os.path.basename(folder)
    fio = base.split(",")[0].strip()
    return fio or "Неизвестный"

def sanitize_filename_strict(s: str) -> str:
    s = re.sub(r'[<>:"/\\|?*\x00-\x1F]', " ", s)
    s = re.sub(r"\s+", " ", s, flags=re.UNICODE).strip(" \t\u00A0")
    s = s.rstrip(".").rstrip(" \t\u00A0")
    return s[:180] if len(s) > 180 else s

def save_pdf(pdf_path, fio, iin) -> bool:
    """
    Сохраняет PDF/изображение в папку клиента.
    Возвращает True при успехе, False при ошибке.
    """
    folder = find_client_folder_by_iin(iin)
    if not folder:
        msg = f"[FATAL] Не найдена папка клиента по ИИН {iin}"
        log(msg)
        with open(log_file_path, "a", encoding="utf-8") as lf:
            lf.write(f"[ОТМЕНА ИН] {msg}\n")
        return False

    fio_safe = sanitize_filename_strict(fio or "Неизвестный")
    fio_safe = re.sub(
        r"(arrow_drop_down|chevron_left|chevron_right|строк на странице\s*\d+.*$)",
        "", fio_safe, flags=re.IGNORECASE
    )
    fio_safe = sanitize_filename_strict(fio_safe)

    ext = Path(pdf_path).suffix.lower()
    if ext not in ALLOWED_EXTS:
        ext = ".pdf"

    base_name = sanitize_filename_strict(f"Постановление об отмене ИН, {fio_safe}, {iin}{ext}")
    dst = os.path.join(folder, base_name)

    try:
        # 🔁 Если файл уже есть — удаляем и записываем заново
        if os.path.exists(dst):
            try:
                os.remove(dst)
                log(f"[OK] Перезаписываю существующий файл: {dst}")
            except Exception as e:
                log(f"[WARN] Не удалось удалить старый файл: {e}. Попробую поверх записать.")

        shutil.copy2(pdf_path, dst)
        log(f"[OK] Сохранён (перезаписан): {dst}")
        return True

    except Exception as e:
        msg = f"[FATAL] Не удалось сохранить '{dst}': {e}"
        log(msg)
        with open(log_file_path, "a", encoding="utf-8") as lf:
            lf.write(f"[ОТМЕНА ИН] {msg}\n")
        return False

# ================= ВХОД (устойчиво, с повторами) =================
def fill_login_form(driver):
    """Ввод логина/пароля максимально «жёстко», чтобы сайт не стирал значения."""
    wait = WebDriverWait(driver, 20)

    user = wait.until(EC.presence_of_element_located((
        By.CSS_SELECTOR,
        "input[name='usernameUserInput'], input#usernameUserInput, input[name='username'], input#username, form#loginForm input[type='text']"
    )))
    pwd = wait.until(EC.presence_of_element_located((
        By.CSS_SELECTOR,
        "input#password, input[name='password'], form#loginForm input[type='password']"
    )))

    driver.execute_script("""
        const u = arguments[0], p = arguments[1], lu = arguments[2], lp = arguments[3];
        u.value = lu; u.dispatchEvent(new Event('input', {bubbles:true})); u.dispatchEvent(new Event('change', {bubbles:true}));
        p.value = lp; p.dispatchEvent(new Event('input', {bubbles:true})); p.dispatchEvent(new Event('change', {bubbles:true}));
    """, user, pwd, LOGIN, PASSWORD)
    try:
        user.clear(); user.send_keys(LOGIN)
    except Exception:
        pass
    try:
        pwd.clear(); pwd.send_keys(PASSWORD)
    except Exception:
        pass

def click_sign_in(driver):
    """Многократные клики по кнопке «Войти» (+ JS-клик)."""
    wait = WebDriverWait(driver, 10)
    btn = wait.until(EC.element_to_be_clickable((
        By.CSS_SELECTOR, "#signInButton, form#loginForm button.ui.primary.button.fluid, button[type='submit']"
    )))
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    except Exception:
        pass

    for _ in range(3):
        try:
            btn.click()
        except Exception:
            pass
        time.sleep(0.15)
        try:
            driver.execute_script("arguments[0].click();", btn)
        except Exception:
            pass
        time.sleep(0.2)

def is_logged_in(driver):
    """Признаки того, что мы прошли авторизацию и видим оболочку кабинета."""
    try:
        if (
            "/cabinet/" in (driver.current_url or "")
            and driver.find_elements(
                By.XPATH,
                "//div[contains(@class,'navigation') or contains(@class,'v-tabs') or contains(@class,'v-tab')]",
            )
        ):
            return True
    except Exception:
        pass
    return False

def login_with_retries(driver, max_rounds=6):
    """Открываем форму, вводим данные, кликаем «Войти». Если сбросило — повторяем цикл."""

    log("[LOGIN] Начинаю авторизацию…")
    driver.get(AISOIP_URL)

    for i in range(1, max_rounds + 1):
        log(f"[LOGIN] попытка {i}/{max_rounds}")
        try:

            WebDriverWait(driver, 25).until(EC.presence_of_element_located((

                By.CSS_SELECTOR,

                "input[name='usernameUserInput'], input#usernameUserInput, input[name='username'], input#username, form#loginForm input[type='text']"

            )))
        except TimeoutException:
            dbg_snapshot(driver, f"login_form_not_found_{i}")
            continue

        fill_login_form(driver)
        time.sleep(0.2)
        click_sign_in(driver)

        try:
            WebDriverWait(driver, 12).until(lambda d: is_logged_in(d))
            if is_logged_in(driver):
                log("[OK] Авторизация выполнена.")
                return True
        except TimeoutException:
            pass

        dbg_snapshot(driver, f"login_retry_{i}")
        try:
            driver.refresh()
        except Exception:
            pass
        time.sleep(0.8)

    dbg_snapshot(driver, "login_failed_after_retries")
    raise TimeoutException("Не удалось войти после серии повторов.")

# ==== НОВОЕ: перезапуск браузера и логина с нуля при неудаче ====
def start_driver_and_login(max_browser_restarts: int = 3):
    """
    Запускает браузер и выполняет login_with_retries.
    При неудаче несколько раз перезапускает браузер.
    """
    last_err = None
    for attempt in range(1, max_browser_restarts + 1):
        drv = None
        try:
            log(f"[SESSION] Старт браузера, попытка {attempt}/{max_browser_restarts}")
            drv = make_driver()
            login_with_retries(drv)
            log("[SESSION] Браузер и авторизация успешны")
            return drv
        except Exception as e:
            last_err = e
            log(f"[SESSION] Ошибка при запуске/авторизации: {e}")
            try:
                if drv is not None:
                    drv.quit()
            except Exception:
                pass
            time.sleep(2)

    raise TimeoutException(f"Не удалось инициализировать браузер и войти в систему: {last_err}")

def ensure_driver_alive(driver, max_browser_restarts: int = 3):
    """
    Проверяет, жива ли сессия браузера.
    Если браузер закрыт/сессия недоступна — поднимает новый и логинится.
    """
    if driver is None:
        return start_driver_and_login(max_browser_restarts=max_browser_restarts)

    try:
        _ = driver.current_url
        return driver
    except WebDriverException as e:
        log(f"[SESSION] Обнаружено, что браузер закрыт или сессия потеряна: {e}. Перезапуск...")
        try:
            driver.quit()
        except Exception:
            pass
        return start_driver_and_login(max_browser_restarts=max_browser_restarts)

# ===================== ИСПРАВЛЕНИЕ #1: РЕЛОГИН ПРИ АНЛОГИНЕ =====================

def is_login_page(driver) -> bool:
    """Определяем, что нас выкинуло на логин (сессия умерла, но браузер жив)."""
    try:
        url = (driver.current_url or "").lower()
        if "login" in url:
            return True
        return len(driver.find_elements(By.CSS_SELECTOR, "form#loginForm, #signInButton")) > 0
    except Exception:
        return False

def ensure_driver_and_auth(driver, max_browser_restarts: int = 3):
    """
    1) Драйвер жив?
    2) Авторизация жива? Если нет — перелогиниваемся.
    3) Если не получилось — перезапуск браузера.
    """
    driver = ensure_driver_alive(driver, max_browser_restarts=max_browser_restarts)

    try:
        if is_login_page(driver) or (not is_logged_in(driver)):
            log("[SESSION] Разлогинило. Повторная авторизация...")
            try:
                driver.delete_all_cookies()
                driver.execute_script("window.localStorage.clear(); window.sessionStorage.clear();")
            except Exception:
                pass
            login_with_retries(driver)
    except Exception as e:
        log(f"[SESSION] Не удалось восстановить сессию: {e}. Перезапуск браузера...")
        try:
            driver.quit()
        except Exception:
            pass
        driver = start_driver_and_login(max_browser_restarts=max_browser_restarts)

    return driver

# ===================== ИСПРАВЛЕНИЕ #2: ЖДАТЬ ЗАГРУЗКУ ПОСЛЕ ПОИСКА =====================

def wait_table_refresh_after_search(driver, timeout=25):
    """
    Ждём, пока результаты поиска реально обновятся (а не sleep).
    Под Vuetify: таблица/карточки/сообщение "нет данных", плюс возможные оверлеи/прогресс.
    """
    wait = WebDriverWait(driver, timeout)

    def count_rows_cards(d):
        rows = d.find_elements(By.XPATH, "//table//tbody//tr")
        cards = d.find_elements(By.XPATH, "//div[contains(@class,'v-card')]")
        return len(rows) + len(cards)

    before_cnt = count_rows_cards(driver)

    def loading_started(d):
        if is_login_page(d):
            return True
        # overlay/spinner (часто в Vuetify)
        if d.find_elements(By.XPATH, "//*[contains(@class,'v-overlay') and contains(@class,'active')]"):
            return True
        src = (d.page_source or "").lower()
        if "v-progress" in src or "loading" in src:
            return True
        # или изменились результаты
        return count_rows_cards(d) != before_cnt

    try:
        wait.until(lambda d: loading_started(d))
    except TimeoutException:
        # Иногда мгновенно выдаёт результат без прогресса — ок, пойдём дальше
        pass

    def loading_finished(d):
        if is_login_page(d):
            return True
        if d.find_elements(By.XPATH, "//*[contains(@class,'v-overlay') and contains(@class,'active')]"):
            return False

        src = (d.page_source or "").lower()
        if "нет данных" in src or "ничего не найдено" in src or "no data" in src:
            return True

        now_cnt = count_rows_cards(d)
        if now_cnt != before_cnt:
            return True

        # базово: есть таблица/контент
        if d.find_elements(By.XPATH, "//div[contains(@class,'v-data-table')]") or d.find_elements(By.XPATH, "//table//tbody"):
            return True

        return False

    wait.until(lambda d: loading_finished(d))
    time.sleep(0.25)

# ================= НАВИГАЦИЯ =================
def open_exec_productions(driver, max_attempts=4):
    def _on_page():
        url_ok = "/cabinet/exec-productions" in (driver.current_url or "")
        if not url_ok:
            return False
        try:
            driver.find_element(By.XPATH, "//div[contains(.,'Должник') or contains(.,'Взыскатель')]")
            return True
        except NoSuchElementException:
            return False

    for attempt in range(1, max_attempts + 1):
        log(f"[NAV] 'Исполнительные производства' попытка {attempt}")
        try:
            driver.get(AISOIP_URL)
            WebDriverWait(driver, 12).until(lambda d: "/cabinet/exec-productions" in d.current_url)
        except Exception:
            pass
        if _on_page():
            return True
        try:
            driver.refresh()
        except Exception:
            pass
        time.sleep(1.2)

    log("[ERR] Не удалось открыть раздел")
    dbg_snapshot(driver, "exec_not_loaded")
    return False

def open_vzyskatel_tab(driver, max_attempts=6):
    def _loaded():
        src = driver.page_source.lower()
        return ("иин взыскателя" in src) or ("иин должника" in src)

    for attempt in range(1, max_attempts + 1):
        log(f"[TAB] Переключение на 'Взыскатель' (попытка {attempt})")
        try:
            groups = driver.find_elements(By.XPATH,
                "//div[contains(@class,'v-tabs') or contains(@class,'v-tabs-bar') or @role='tablist']"
            )
            target_group = None
            for g in groups:
                txt = g.text.upper().replace("Ё","Е")
                if ("ДОЛЖНИК" in txt) and ("ВЗЫСКАТЕЛЬ" in txt):
                    target_group = g
                    break

            if not target_group:
                try:
                    target_group = driver.find_element(By.XPATH,
                        "//*[@class][contains(.,'ДОЛЖНИК')]/ancestor::div[contains(@class,'v-sheet') or contains(@class,'container')][1]"
                        "//div[contains(@class,'v-tabs') or @role='tablist']"
                    )
                except Exception:
                    pass

            if not target_group:
                log("[TAB] Группа вкладок не найдена")
                dbg_snapshot(driver, f"tab_group_not_found_{attempt}")
                time.sleep(0.8)
                continue

            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", target_group)
            time.sleep(0.15)

            vz = None
            for xp in [
                ".//div[@role='tab' and contains(.,'Взыскатель')]",
                ".//div[contains(@class,'v-tab') and contains(.,'Взыскатель')]",
                ".//*[self::div or self::a][contains(@class,'v-tab') and contains(.,'Взыскатель')]",
            ]:
                try:
                    vz = target_group.find_element(By.XPATH, xp)
                    break
                except Exception:
                    continue

            ActionChains(driver).move_to_element(vz).pause(0.05).click(vz).perform()
            driver.execute_script("arguments[0].click();", vz)
            log("[TAB] Клик по 'Взыскатель' отправлен")

            WebDriverWait(driver, 7).until(lambda d: _loaded())
            if _loaded():
                log("[OK] Контент 'Взыскатель' загружен")
                return True

        except Exception as e:
            log(f"[TAB] Ошибка: {e}")

        dbg_snapshot(driver, f"tab_vzyskatel_attempt_{attempt}")
        time.sleep(0.9)

    log("[ERR] Не удалось активировать вкладку 'Взыскатель'")
    return False

def clear_all_filters(driver):
    chips = driver.find_elements(
        By.XPATH,
        "//*[contains(@class,'chip')][contains(.,'ИИН должника') or contains(.,'БИН должника') or "
        "contains(.,'ИИН взыскателя') or contains(.,'БИН взыскателя')]"
    )
    for chip in chips:
        try:
            close_btn = chip.find_element(By.XPATH, ".//*[contains(@class,'close') or @role='button' or contains(.,'×')]")
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", close_btn)
            try:
                close_btn.click()
            except Exception:
                driver.execute_script("arguments[0].click();", close_btn)
            log("[CLR] Удалён чип фильтра")
            time.sleep(0.05)
        except Exception:
            pass

    for lbl in ["ИИН должника", "ИИН взыскателя", "БИН должника", "БИН взыскателя"]:
        try:
            inp = driver.find_element(By.XPATH, f"//label[contains(.,'{lbl}')]/following::input[1]")
            if inp.is_displayed():
                inp.click()
                ActionChains(driver).key_down('\ue009').send_keys('a').key_up('\ue009').send_keys('\ue003').perform()
                log(f"[CLR] Очищено поле '{lbl}'")
        except Exception:
            pass

def search_by_iin_dolzhnik_on_vzyskatel(driver, iin: str):
    log(f"[SEARCH] Ввожу ИИН ДОЛЖНИКА во вкладке 'Взыскатель': {iin}")
    clear_all_filters(driver)
    wait = WebDriverWait(driver, 15)
    inp = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//label[contains(.,'ИИН должника')]/following::input[1]")
    ))
    inp.click()
    ActionChains(driver).key_down('\ue009').send_keys('a').key_up('\ue009').send_keys('\ue003').perform()
    inp.send_keys(iin)
    time.sleep(0.15)
    for xp in [
        "//button[.//span[contains(.,'ПОИСК')] or contains(.,'Поиск')]",
        "//button[contains(.,'Поиск') or contains(.,'ПОИСК')]",
    ]:
        try:
            driver.find_element(By.XPATH, xp).click()
            log("[SEARCH] Нажал 'Поиск'")
            break
        except Exception:
            continue

    # ✅ ВАЖНО: вместо sleep ждём обновления результатов
    log("[SEARCH] Жду загрузку результатов после поиска...")
    wait_table_refresh_after_search(driver, timeout=25)

def open_latest_execution(driver):
    time.sleep(0.6)
    log("[OPEN] Ищу карточки ИП...")

    links = driver.find_elements(
        By.XPATH,
        "//div[contains(@class,'v-card') or contains(@class,'v-sheet') or contains(@class,'v-data-table')]/descendant::a[contains(.,'/')][contains(@style,'text-decoration')]"
    )
    if not links:
        links = driver.find_elements(By.XPATH, "//a[contains(@href,'exec-productions') or contains(@href,'exec/')]")
    if not links:
        dbg_snapshot(driver, "no_exec_links")
        raise RuntimeError("Нет ссылок на исполнительные производства.")

    best = None  # (date, link, card)
    for link in links:
        try:
            card = link.find_element(
                By.XPATH,
                "./ancestor::*[self::div[contains(@class,'v-card') or contains(@class,'v-row') or contains(@class,'container')] or self::tr][1]"
            )
            txt = card.text
            m = re.search(r"от\s+(\d{2}\.\d{2}\.\d{4})", txt)
            d = datetime.strptime(m.group(1), "%d.%m.%Y") if m else datetime(1900,1,1)
            if (best is None) or (d > best[0]):
                best = (d, link, card)
        except Exception:
            continue

    if best is None:
        dbg_snapshot(driver, "no_card_parsed")
        raise RuntimeError("Не удалось выбрать карточку ИП.")

    _, chosen_link, _ = best
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", chosen_link)
    time.sleep(0.1)
    try:
        chosen_link.click()
    except Exception:
        driver.execute_script("arguments[0].click();", chosen_link)
    log("[OPEN] Клик по номеру ИП выполнен")
    time.sleep(1.0)

def pick_and_download_cancel_doc(driver, max_checks=4):
    time.sleep(0.4)
    pdf_icons = driver.find_elements(By.XPATH,
        "//i[contains(@class,'material-icons') and normalize-space(text())='get_app']"
    )
    if not pdf_icons:
        pdf_icons = driver.find_elements(By.XPATH,
            "//*[contains(@class,'material-icons') and normalize-space(text())='get_app']"
        )
    log(f"[DL] Иконок скачивания: {len(pdf_icons)}")

    checks = 0
    for icon in pdf_icons:
        if checks >= max_checks:
            break
        path = None
        try:
            row = icon.find_element(By.XPATH, "./ancestor::*[self::tr or contains(@class,'row') or self::li][1]")
            title = row.text.strip()

            if any(ex.lower() in title.lower() for ex in EXCLUDE_TITLES):
                log(f"[DL] Пропуск по исключению: {(title.splitlines()[-1] if title else title)}")
                continue

            before = set(os.listdir(DOWNLOAD_DIR))
            for attempt in range(2):
                try:
                    driver.execute_script("arguments[0].scrollIntoView(true);", icon)
                except Exception:
                    pass
                try:
                    icon.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", icon)

                path = wait_new_file(before, timeout=18 if attempt == 0 else 30)
                if path:
                    break
                time.sleep(0.4)

            if not path:
                log("[DL] Не дождались файла")
                continue

            if not is_supported_file(path):
                log(f"[DL] Скачан файл неподдерживаемого типа, пропуск: {path}")
                continue

            ext = Path(path).suffix.lower()
            checks += 1

            if ext == ".pdf":
                if is_cancel_doc(path):
                    log(f"[DL] Найден документ об отмене: {path}")
                    return path
                else:
                    log(f"[DL] Не подходит: {path}")
            elif ext in (".jpg", ".jpeg", ".png"):
                log(f"[DL] Скачано изображение, считаю его документом об отмене: {path}")
                return path

        except Exception as e:
            log(f"[DL] Ошибка при скачивании: {e}")
            continue
    return None

# ================= ОСНОВНОЙ ПРОГОН =================
def main():
    global LOG_SUMMARY

    total_cancel     = 0
    found_cancel     = 0
    not_found_cancel = []

    driver = start_driver_and_login()
    try:
        if not open_exec_productions(driver): return
        if not open_vzyskatel_tab(driver):   return

        df = pd.read_excel(EXCEL_PATH, dtype=str, header=0)
        if IIN_COLUMN_INDEX >= df.shape[1]:
            raise ValueError(f"Нет колонки с индексом {IIN_COLUMN_INDEX}. Всего: {df.shape[1]}")
        iin_series = df.iloc[:, IIN_COLUMN_INDEX].astype(str)

        iin_list = []
        for v in iin_series:
            m = RE_IIN.search(v)
            iin = m.group(1) if m else re.sub(r"\D", "", v)
            iin_list.append(iin)

        for idx, iin in enumerate(iin_list, start=1):
            # ✅ ВАЖНО: теперь перед каждым ИИН проверяем и драйвер, и авторизацию (анлогин лечится)
            driver = ensure_driver_and_auth(driver)

            if len(iin) != 12:
                log(f"[{idx}] Пропуск — некорректный ИИН: {iin}")
                with open(log_file_path, "a", encoding="utf-8") as lf:
                    lf.write(f"[ОТМЕНА ИН] Пропуск некорректного ИИН: {iin}\n")
                continue

            total_cancel += 1
            log(f"\n[{idx}] === ИИН: {iin} ===")

            if not open_exec_productions(driver) or not open_vzyskatel_tab(driver):
                msg = "[WARN] Раздел/вкладка недоступны — пропуск"
                log(msg)
                not_found_cancel.append(f"Неизвестный, {iin}")
                with open(log_file_path, "a", encoding="utf-8") as lf:
                    lf.write(f"[ОТМЕНА ИН] {msg} для ИИН {iin}\n")
                continue

            fio_from_folder = get_fio_from_folder(iin)

            try:
                search_by_iin_dolzhnik_on_vzyskatel(driver, iin)
                open_latest_execution(driver)

                pdf_path = pick_and_download_cancel_doc(driver, max_checks=4)

                if not pdf_path:
                    msg = f"[WARN] Для {iin} не найден PDF/изображение об отмене ИН"
                    log(msg)
                    not_found_cancel.append(f"{fio_from_folder}, {iin}")
                    with open(log_file_path, "a", encoding="utf-8") as lf:
                        lf.write(f"[ОТМЕНА ИН] Не найден документ об отмене для ФИО: {fio_from_folder}, ИИН: {iin}\n")
                else:
                    ok = save_pdf(pdf_path, fio_from_folder, iin)
                    if ok:
                        found_cancel += 1
                    else:
                        not_found_cancel.append(f"{fio_from_folder}, {iin}")

            except Exception as e:
                msg = f"[WARN] ИИН {iin}: ошибка обработки: {e}"
                log(msg)
                not_found_cancel.append(f"{fio_from_folder}, {iin}")
                with open(log_file_path, "a", encoding="utf-8") as lf:
                    lf.write(f"[ОТМЕНА ИН] Ошибка обработки для ФИО: {fio_from_folder}, ИИН: {iin}, ошибка: {e}\n")
                dbg_snapshot(driver, f"error_{iin}")

            try:
                driver.get(AISOIP_URL)
            except Exception:
                pass
            time.sleep(1.2)

        log("\n=== ГОТОВО ===")
        log(f"[ИТОГ ОТМЕН ИН] Найдено {found_cancel} из {total_cancel}")

        LOG_SUMMARY["ПОСТАНОВЛЕНИЯ ОБ ОТМЕНЕ ИН"] = {
            "found": found_cancel,
            "total": total_cancel,
            "not_found": not_found_cancel,
        }

    finally:
        # driver.quit()  # включи при необходимости автозакрытие браузера
        pass

if __name__ == "__main__":
    main()

# # Формирование по Пулам

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

# ===== НАСТРОЙКИ ПУТЕЙ =====
EXCEL_PATH = r"C:\Users\User\Desktop\Документы для подачи Исков\Отчёт по отменам_Omega.xlsx"
FOLDERS_ROOT = r"C:\Users\User\Desktop\Документы для подачи Исков\Папки_для_исков"

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

def main():
    # 1. читаем группы ИИН из Excel
    iin_groups = read_iin_groups_from_excel(EXCEL_PATH, POOL_SIZE)
    print(f"Найдено ИИН (после фильтра по 'Vivus'): {sum(len(g) for g in iin_groups)}")
    print(f"Всего пулов: {len(iin_groups)}")

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

            if not folders:
                print(f"  [!] Не найдена папка для ИИН {iin}")
                continue

            for folder in folders:
                print(f"  ИИН {iin}: копирую из {folder}")
                copy_unique_files(folder, pool_dir)

    print("\nГотово. Папки 'Пул 1', 'Пул 2', ... созданы в:")
    print(FOLDERS_ROOT)

if __name__ == "__main__":
    main()

# # Формирование Массового заявления в Медеуский суд

# -*- coding: utf-8 -*-
import os
import re
from pathlib import Path
from datetime import datetime, date

from openpyxl import load_workbook
from docx import Document

# ===================== НАСТРОЙКИ =====================

EXCEL_PATH   = r"C:\Users\User\Desktop\Документы для подачи Исков\Отчёт по отменам_Omega.xlsx"
SHEET_NAME   = "Данные для шаблонов"

# Шаблон массового иска
TEMPLATE_DOC_MASS = r"C:\Users\User\Desktop\Документы для подачи Исков\Шаблоны документов\Шаблон массового иска.docx"

# Базовая папка, где лежат папки должников и Пул 1, Пул 2, ...
FOLDERS_ROOT = r"C:\Users\User\Desktop\Документы для подачи Исков\Папки_для_исков"

# Для совместимости: используем ту же папку для лога
ROOT_OUT_FOLDER = FOLDERS_ROOT

# Сколько максимум ответчиков в одном иске
GROUP_SIZE = 10

# Лог-файл
os.makedirs(ROOT_OUT_FOLDER, exist_ok=True)
log_file_path = os.path.join(ROOT_OUT_FOLDER, "лог_массовые_иски.txt")

# ===================== КАРТА ПЕРЕИМЕНОВАНИЯ ЗАГОЛОВКОВ =====================

HEADERS_REMAP = {
    # если когда-нибудь в Excel поменяются названия колонок —
    # сюда можно добавить соответствия
    # "Номер договора займа (ММ)": "Номер договора займа",
    # "Номер договора уступки (цессии)": "Номер договора цессии",
    # "Сумма выдачи займа, тг": "Сумма выдачи займа",
}

# Денежные колонки (форматируем как '16 487,00')
MONEY_COLUMNS = {
    "Остаток задолженности",
    "Сумма выдачи займа",
    "Основной долг",
    "Вознаграждение",
    "Штраф",
    "Исполнительная надпись (расходы)",
    "Госпошлина (расходы)",
    "ОД+проценты+штрафы",
}

# ===================== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ =====================

def nice_case(s: str) -> str:
    """ФИО -> 'Фамилия Имя Отчество'."""
    if not s:
        return ""
    return " ".join(word.capitalize() for word in str(s).split())

def format_money(value) -> str:
    """Формат денежной суммы: 16487 -> '16 487,00'."""
    try:
        f = float(str(value).replace(" ", "").replace(",", "."))
    except Exception:
        return str(value)
    f = round(f + 1e-9, 2)
    s = f"{f:,.2f}"          # '16,487.00'
    s = s.replace(",", " ")  # '16 487.00'
    s = s.replace(".", ",")  # '16 487,00'
    return s

def num_to_words_ru(n: int) -> str:
    """Число в русские слова (для тенге, без слова 'тенге')."""
    n = int(n)
    if n == 0:
        return "ноль"

    units = {
        0: ("ноль", "ноль"),
        1: ("один", "одна"),
        2: ("два", "две"),
        3: ("три", "три"),
        4: ("четыре", "четыре"),
        5: ("пять", "пять"),
        6: ("шесть", "шесть"),
        7: ("семь", "семь"),
        8: ("восемь", "восемь"),
        9: ("девять", "девять"),
    }
    teens = {
        10: "десять",
        11: "одиннадцать",
        12: "двенадцать",
        13: "тринадцать",
        14: "четырнадцать",
        15: "пятнадцать",
        16: "шестнадцать",
        17: "семнадцать",
        18: "восемнадцать",
        19: "девятнадцать",
    }
    tens = {
        2: "двадцать",
        3: "тридцать",
        4: "сорок",
        5: "пятьдесят",
        6: "шестьдесят",
        7: "семьдесят",
        8: "восемьдесят",
        9: "девяносто",
    }
    hundreds = {
        1: "сто",
        2: "двести",
        3: "триста",
        4: "четыреста",
        5: "пятьсот",
        6: "шестьсот",
        7: "семьсот",
        8: "восемьсот",
        9: "девятьсот",
    }
    # формы: 1, 2-4, 5+
    orders = [
        ("", "", ""),
        ("тысяча", "тысячи", "тысяч"),
        ("миллион", "миллиона", "миллионов"),
        ("миллиард", "миллиарда", "миллиардов"),
    ]

    words = []
    order = 0
    while n > 0:
        n, rem = divmod(n, 1000)
        if rem == 0:
            order += 1
            continue

        part = []
        h = rem // 100
        t_u = rem % 100

        if h:
            part.append(hundreds[h])

        if 10 <= t_u <= 19:
            part.append(teens[t_u])
            u = 0
        else:
            t = t_u // 10
            u = t_u % 10
            if t:
                part.append(tens[t])
            if u:
                fem = (order == 1)  # тысячи — женский род
                part.append(units[u][1 if fem else 0])

        if order > 0:
            if u == 1 and t_u != 11:
                form = orders[order][0]
            elif u in (2, 3, 4) and not 12 <= t_u <= 14:
                form = orders[order][1]
            else:
                form = orders[order][2]
            if form:
                part.append(form)

        words.append(" ".join(part))
        order += 1

    return " ".join(reversed(words))

# ---------- хелпер для замены подстроки в runs БЕЗ потери форматирования ----------

def replace_substring_in_runs(runs, index_map, start, end, new_text):
    """
    Заменяет символы с позиций [start, end) в склеенном тексте
    на new_text, сохраняя форматирование остальных частей абзаца.
    """
    if start >= end:
        return

    start_pos = start
    end_pos = end - 1

    start_run_idx, start_off = index_map[start_pos]
    end_run_idx, end_off = index_map[end_pos]

    start_run = runs[start_run_idx]
    end_run = runs[end_run_idx]

    start_text = start_run.text
    end_text = end_run.text

    prefix = start_text[:start_off]
    suffix = end_text[end_off + 1:]

    if start_run_idx == end_run_idx:
        start_run.text = prefix + new_text + suffix
    else:
        start_run.text = prefix + new_text
        end_run.text = suffix
        for ri in range(start_run_idx + 1, end_run_idx):
            runs[ri].text = ""

# ---------- ЗАМЕНА ПЛЕЙСХОЛДЕРОВ ----------

def replace_placeholders_in_paragraph(paragraph, mapping: dict):
    """
    Надёжная замена плейсхолдеров вида {Колонка_N} во всём абзаце,
    с учётом того, что плейсхолдер может быть разбит на несколько run'ов.
    """
    runs = paragraph.runs
    if not runs:
        return

    full_text = "".join(r.text for r in runs)
    if not full_text:
        return

    if not any(k in full_text for k in mapping.keys()):
        return

    index_map = []
    for ri, r in enumerate(runs):
        for ci in range(len(r.text)):
            index_map.append((ri, ci))

    replacements = []
    for key, value in mapping.items():
        if not key:
            continue
        search_from = 0
        while True:
            pos = full_text.find(key, search_from)
            if pos == -1:
                break
            replacements.append((pos, pos + len(key), str(value)))
            search_from = pos + len(key)

    if not replacements:
        return

    replacements.sort(key=lambda x: x[0], reverse=True)

    for start, end, new_text in replacements:
        replace_substring_in_runs(runs, index_map, start, end, new_text)

# ---------- {сумма прописью} БЕЗ потери форматирования ----------

def fill_sum_in_words(doc: Document):
    pattern = re.compile(r'(\d[\d\s]*,\d{2})\s*\(\{сумма прописью\}\)')

    for p in doc.paragraphs:
        runs = p.runs
        if not runs:
            continue

        full_text = "".join(r.text for r in runs)
        if "{сумма прописью}" not in full_text:
            continue

        matches = list(pattern.finditer(full_text))
        if not matches:
            continue

        index_map = []
        for ri, r in enumerate(runs):
            for ci in range(len(r.text)):
                index_map.append((ri, ci))

        for m in reversed(matches):
            start, end = m.span()
            num_str = m.group(1)
            digits = num_str.replace(" ", "").replace("\u00A0", "")
            if "," in digits:
                int_part = digits.split(",")[0]
            else:
                int_part = digits
            try:
                n = int(int_part)
            except Exception:
                n = 0
            words = num_to_words_ru(n)
            new_sub = f"{num_str} ({words})"

            replace_substring_in_runs(runs, index_map, start, end, new_sub)

# ---------- ОЧИСТКА ЛИШНИХ ОТВЕТЧИКОВ ----------

def cleanup_unused_respondents(doc: Document, count: int, max_count: int = 10):
    """
    Удаляем блоки по Ответчикам, которых нет в текущем пуле.
    """

    # 1. Шапка "Ответчики"
    for n in range(count + 1, max_count + 1):
        i = 0
        while i < len(doc.paragraphs):
            p = doc.paragraphs[i]
            if f"Ответчик №{n}:" in p.text:
                for _ in range(5):
                    if i < len(doc.paragraphs):
                        p_del = doc.paragraphs[i]
                        p_del._element.getparent().remove(p_del._element)
                    else:
                        break
                continue
            i += 1

    # 2. Фактическая часть до ПРОШУ
    i = 0
    while i < len(doc.paragraphs):
        paras = doc.paragraphs
        if i >= len(paras):
            break

        p = paras[i]
        text = p.text or ""

        m = re.search(r"Ответчик\s*№\s*(\d+)", text)
        if not m:
            i += 1
            continue

        num = int(m.group(1))

        if num <= count or num > max_count:
            i += 1
            continue

        start = i
        j = i + 1
        while j < len(paras):
            t2 = paras[j].text or ""
            t2_strip = t2.strip()

            if re.search(r"^\s*\d+\.\s*Ответчик\s*№\s*\d+", t2_strip):
                break

            if (t2_strip.startswith("Согласно ст. 339") or
                t2_strip.startswith("На основании вышеизложенного") or
                t2_strip.startswith("ПРОШУ:")):
                break

            j += 1

        for _ in range(j - start):
            p_del = doc.paragraphs[start]
            p_del._element.getparent().remove(p_del._element)

    # 3. Хвост в разделе "ПРОШУ"
    for n in range(count + 1, max_count + 1):
        mark1 = f"Ответчик №{n}"
        mark2 = f"Ответчика №{n}"
        to_delete = []
        for p in doc.paragraphs:
            txt = p.text or ""
            if mark1 in txt or mark2 in txt:
                to_delete.append(p)
        for p in to_delete:
            try:
                p._element.getparent().remove(p._element)
            except Exception:
                pass

# ---------- ЗАПОЛНЕНИЕ ДОКУМЕНТА ----------

def fill_document_mass(group_rows: list, norm_headers: list) -> Document:
    doc = Document(TEMPLATE_DOC_MASS)
    mapping = {}

    for idx, row_dict in enumerate(group_rows, start=1):
        for col_name in norm_headers:
            if not col_name:
                continue

            value = row_dict.get(col_name, "")

            if value is None or value != value:
                value_str = ""
            elif isinstance(value, (datetime, date)):
                value_str = value.strftime("%d.%m.%Y")
            else:
                col_lower = col_name.lower()
                if "фио" in col_lower:
                    value_str = nice_case(value)
                elif col_name in MONEY_COLUMNS:
                    value_str = format_money(value)
                else:
                    value_str = str(value)

            placeholder = "{" + f"{col_name}_{idx}" + "}"
            mapping[placeholder] = value_str

    for idx in range(len(group_rows) + 1, GROUP_SIZE + 1):
        for col_name in norm_headers:
            if not col_name:
                continue
            placeholder = "{" + f"{col_name}_{idx}" + "}"
            mapping.setdefault(placeholder, "")

    mapping["{CurrDate}"] = datetime.now().strftime("%d.%m.%Y")
    mapping["{RespondentCount}"] = str(len(group_rows))

    for p in doc.paragraphs:
        replace_placeholders_in_paragraph(p, mapping)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_placeholders_in_paragraph(p, mapping)

    fill_sum_in_words(doc)
    cleanup_unused_respondents(doc, len(group_rows), max_count=GROUP_SIZE)

    return doc

# ===================== ОСНОВНАЯ ЛОГИКА =====================

def main():
    global LOG_SUMMARY

    base_path = Path(FOLDERS_ROOT)
    base_path.mkdir(parents=True, exist_ok=True)

    wb = load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    ws = wb[SHEET_NAME]

    raw_headers = [str(c.value).strip() if c.value else "" for c in ws[1]]

    print("[DEBUG] Заголовки из Excel (сырые):")
    for i, h in enumerate(raw_headers, start=1):
        print(f"{i}: '{h}'")

    norm_headers = [HEADERS_REMAP.get(h, h) for h in raw_headers]

    print("\n[DEBUG] Нормализованные заголовки (по ним делаем плейсхолдеры {Колонка_N}):")
    for i, h in enumerate(norm_headers, start=1):
        print(f"{i}: '{h}'")

    fio_header = norm_headers[2] if len(norm_headers) > 2 else "ФИО"
    iin_header = norm_headers[3] if len(norm_headers) > 3 else "ИИН"

    rows_for_mass = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(v is None for v in row):
            continue

        prod_value = row[1] if len(row) > 1 else None
        if prod_value and "vivus" in str(prod_value).lower():
            continue

        raw_row_dict = dict(zip(raw_headers, row))
        row_dict = {}
        for k, v in raw_row_dict.items():
            if not k:
                continue
            new_k = HEADERS_REMAP.get(k, k)
            row_dict[new_k] = v

        fio = nice_case(row[2]) if len(row) > 2 and row[2] else ""
        iin = str(row[3]).strip() if len(row) > 3 and row[3] else ""

        if not fio or not iin:
            continue

        for key in list(row_dict.keys()):
            if not key:
                continue
            lk = key.lower()
            if "фио" in lk:
                row_dict[key] = fio
            if "иин" in lk:
                row_dict[key] = iin

        row_dict[fio_header] = fio
        row_dict[iin_header] = iin

        rows_for_mass.append(row_dict)

    wb.close()

    if not rows_for_mass:
        print("[INFO] Не найдено строк для массовых исков (кроме Vivus).")
        return

    print(f"\n[INFO] Всего должников (кроме Vivus): {len(rows_for_mass)}")

    mass_docs_created = 0

    # Разбиваем на группы по GROUP_SIZE (10 ответчиков)
    for start in range(0, len(rows_for_mass), GROUP_SIZE):
        group_rows = rows_for_mass[start:start + GROUP_SIZE]
        group_index = start // GROUP_SIZE + 1

        print(f"\n[STEP MASS] Группа {group_index}: ответчиков {len(group_rows)}")

        try:
            doc = fill_document_mass(group_rows, norm_headers)

            first_row = group_rows[0]
            first_fio = nice_case(first_row.get(fio_header, "")) or "БезФИО"
            first_iin = str(first_row.get(iin_header, "")).strip() or "БезИИН"

            base_name = f"Массовый иск в Медеуски суд, пул {group_index}"
            for ch in r'\/:*?"<>|':
                base_name = base_name.replace(ch, "_")

            # <<< главное изменение: сохраняем в соответствующий Пул >>>
            pool_dir = base_path / f"Пул {group_index}"
            pool_dir.mkdir(parents=True, exist_ok=True)

            docx_path = pool_dir / (base_name + ".docx")
            doc.save(docx_path)

            print(f"[OK] Сформирован массовый иск: {docx_path}")
            mass_docs_created += 1

        except Exception as e:
            msg = f"[ERROR MASS] Ошибка при формировании группы {group_index}: {e}"
            print(msg)
            try:
                with open(log_file_path, "a", encoding="utf-8") as lf:
                    lf.write(msg + "\n")
            except Exception:
                pass
            continue

    print("\n=== ГОТОВО: массовые иски (кроме Vivus) сформированы ===")
    print(f"[ИТОГ] документов: {mass_docs_created}")

    LOG_SUMMARY["МАССОВЫЕ ИСКИ (КРОМЕ VIVUS)"] = {
        "found": mass_docs_created,       # сколько реально создано документов
        "total": len(rows_for_mass),      # сколько групп/дел по плану
        "not_found": []                   # или список проблемных, если есть
    }

if __name__ == "__main__":
    main()

# # Логирование Сбора документов

# === ФИНАЛЬНАЯ СВОДКА ПО ВСЕМ БЛОКАМ ===

import os

# Папка, где находятся папки клиентов
BASE_FOLDER = r"C:\Users\User\Desktop\Документы для подачи Исков\Папки_для_исков"
os.makedirs(BASE_FOLDER, exist_ok=True)

# Файл, куда сохранить именно СВОДНЫЙ отчёт
SUMMARY_LOG_PATH = os.path.join(BASE_FOLDER, "лог_сбор_документов.txt")

def write_final_summary(log_summary_dict, summary_path=SUMMARY_LOG_PATH):
    """
    log_summary_dict:
        {
          "ОФЕРТЫ": {
              "found": 57,
              "total": 57,
              "not_found": ["ФИО1, ИИН", "ФИО2, ИИН"]
          },
          ...
        }
    """
    lines = []
    lines.append("")
    lines.append("======================================")
    lines.append("=== СВОДНЫЙ ОТЧЁТ ПО ДОКУМЕНТАМ ===")
    lines.append("======================================")
    lines.append("")

    # порядок разделов
    preferred_order = [
        "ОФЕРТЫ",
        "УВЕДОМЛЕНИЯ ОБ УСТУПКЕ ПРАВА ТРЕБОВАНИЯ",
        "ДОГОВОРЫ ЦЕССИИ",
        "РЕЕСТРЫ ДОГОВОРА ЦЕССИИ",
        "НОТАРИАЛЬНЫЕ НАДПИСИ",
        "ГОСПОШЛИНЫ",
        "ПОСТАНОВЛЕНИЯ ОБ ОТМЕНЕ ИН",
    ]

    used_keys = set()
    for key in preferred_order:
        if key in log_summary_dict:
            used_keys.add(key)
            data = log_summary_dict.get(key, {})
            found = int(data.get("found", 0))
            total = int(data.get("total", 0))
            not_found_list = list(data.get("not_found", []) or [])
            missed = total - found

            lines.append(key)
            lines.append(f"найдены {found} из {total}")
            if missed <= 0:
                lines.append(f"не найдены 0 из {total}")
            else:
                if not_found_list:
                    nf_join = "; ".join(not_found_list)
                    lines.append(f"не найдены {missed} из {total}: {nf_join}")
                else:
                    lines.append(f"не найдены {missed} из {total}")
            lines.append("")

    # остальные ключи (на будущее)
    for key, data in log_summary_dict.items():
        if key in used_keys:
            continue
        found = int(data.get("found", 0))
        total = int(data.get("total", 0))
        not_found_list = list(data.get("not_found", []) or [])
        missed = total - found

        lines.append(key)
        lines.append(f"найдены {found} из {total}")
        if missed <= 0:
            lines.append(f"не найдены 0 из {total}")
        else:
            if not_found_list:
                nf_join = "; ".join(not_found_list)
                lines.append(f"не найдены {missed} из {total}: {nf_join}")
            else:
                lines.append(f"не найдены {missed} из {total}")
        lines.append("")

    text = "\n".join(lines)

    # 1) Печатаем в консоль (как раньше)
    print(text)

    # 2) Записываем в текстовый файл в Папки_для_исков
    with open(summary_path, "w", encoding="utf-8") as f:
        f.write(text + "\n")

    print("\n[OK] Сводный отчёт сохранён в файл:")
    print(summary_path)

# ВЫЗОВ В САМОМ КОНЦЕ СКРИПТА / НОУТБУКА:
try:
    write_final_summary(LOG_SUMMARY)
except NameError:
    print("LOG_SUMMARY не найден — нет данных для сводного отчёта.")

# # Дублирование папки

import os
import shutil
from datetime import datetime

# Папка, которую бэкапим
SRC_DIR = r"C:\Users\User\Desktop\Документы для подачи Исков\Папки_для_исков"

# Куда складываем резервные копии
DST_ROOT = r"D:\work folder\Документы для подачи Исков"

# Папка, где лежит Excel-файл "Отчёт по отменам_Omega..."
EXCEL_DIR = r"C:\Users\User\Desktop\Документы для подачи Исков"

# Часть имени Excel-файла
EXCEL_NAME_PART = "Отчёт по отменам_Omega"

def make_timestamp_foldername() -> str:
    # Был запрос "дд.мм.гггг (чч:мм)", но ":" в Windows нельзя — заменяем на "-"
    return datetime.now().strftime("%d.%m.%Y (%H-%M)")

def copy_report_excel(dst_folder: str):
    """
    Ищет в EXCEL_DIR все Excel-файлы, в имени которых есть EXCEL_NAME_PART,
    и копирует их в папку dst_folder.
    """
    if not os.path.isdir(EXCEL_DIR):
        print(f"[WARN] Папка с Excel не найдена: {EXCEL_DIR}")
        return

    found = False
    for fname in os.listdir(EXCEL_DIR):
        full_path = os.path.join(EXCEL_DIR, fname)
        if not os.path.isfile(full_path):
            continue

        # Проверяем, что в названии есть нужный фрагмент и расширение Excel
        if (EXCEL_NAME_PART in fname and
                fname.lower().endswith((".xlsx", ".xls", ".xlsm"))):
            found = True
            dst_path = os.path.join(dst_folder, fname)
            print(f"[INFO] Копирую Excel-файл:\n  из: {full_path}\n  в : {dst_path}")
            shutil.copy2(full_path, dst_path)

    if not found:
        print(f"[WARN] В папке {EXCEL_DIR} не найдено Excel-файлов с именем, содержащим '{EXCEL_NAME_PART}'.")

def backup_folder(src: str = SRC_DIR, dst_root: str = DST_ROOT) -> str:
    if not os.path.isdir(src):
        raise FileNotFoundError(f"Источник не найден: {src}")

    os.makedirs(dst_root, exist_ok=True)

    base_name = make_timestamp_foldername()
    dst = os.path.join(dst_root, base_name)

    # Если на ту же минуту уже делали копию — добавим счетчик
    counter = 1
    final_dst = dst
    while os.path.exists(final_dst):
        final_dst = f"{dst} #{counter}"
        counter += 1

    print(f"[INFO] Копируем папку:\n  из: {src}\n  в : {final_dst}")
    shutil.copytree(src, final_dst)
    print("[OK] Папка скопирована.")

    # Дополнительно копируем Excel "Отчёт по отменам_Omega..." в эту же папку
    copy_report_excel(final_dst)

    print("[OK] Резервное копирование завершено.")
    return final_dst

if __name__ == "__main__":
    backup_path = backup_folder()
    print(f"[RESULT] Резервная копия: {backup_path}")