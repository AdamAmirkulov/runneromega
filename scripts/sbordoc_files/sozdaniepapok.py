"""
Блок 2: Создание папок для клиентов + копирование учредительных документов
"""
import os
import shutil

from config import ROOT, TARGET_BASE
from utils import ensure_row_folder, safe_update_summary

# ═══════════════════════════════════════════════════════════════
# НАСТРОЙКИ
# ═══════════════════════════════════════════════════════════════

DOCS_SOURCE_FOLDER = rf'{ROOT}\Документы для подачи Исков\Учредительные документы'

DOCS_TO_COPY = [
    #'Доверенность представителя.pdf',
    'Приказ на директора.pdf',
    'Устав ТОО.pdf',
    'Разъяснение ВС касательно подсудности.pdf',
]

# ═══════════════════════════════════════════════════════════════
# ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ
# ═══════════════════════════════════════════════════════════════

def copy_item(source_path: str, target_path: str, doc_name: str) -> bool:
    """
    Копирует файл или папку из source_path в target_path.
    Returns True при успехе, False при ошибке.
    """
    try:
        if os.path.isdir(source_path):
            if os.path.exists(target_path):
                shutil.rmtree(target_path)
            shutil.copytree(source_path, target_path)
        else:
            shutil.copy2(source_path, target_path)
        return True
    except Exception as e:
        print(f"     Ошибка копирования '{doc_name}': {e}")
        return False


def get_prikaz_sotrudnika_files(source_folder: str) -> list:
    """
    Возвращает список всех файлов/папок с 'Приказ на сотрудника' в названии.
    """
    try:
        return [
            f for f in os.listdir(source_folder)
            if 'Приказ на сотрудника' in f
        ]
    except Exception as e:
        print(f"     Ошибка при поиске приказов на сотрудника: {e}")
        return []
    
def get_doverennost_files(source_folder: str) -> list:
    """
    Возвращает список всех файлов/папок с 'Приказ на сотрудника' в названии.
    """
    try:
        return [
            f for f in os.listdir(source_folder)
            if 'Доверенность представителя' in f
        ]
    except Exception as e:
        print(f"     Ошибка при поиске приказов на сотрудника: {e}")
        return []

# ═══════════════════════════════════════════════════════════════
# ОСНОВНАЯ ФУНКЦИЯ
# ═══════════════════════════════════════════════════════════════

def run(df_main):
    print(f"Строк в df_main: {len(df_main)}")
    print(df_main[['FIO', 'IIN']].to_string())

    dupes = df_main[df_main.duplicated(subset=['IIN'], keep=False)]
    if not dupes.empty:
        print(f"Дубликаты ИИН:\n{dupes[['FIO','IIN']]}")

    print("\n" + "="*70)
    print("   БЛОК 2: СОЗДАНИЕ ПАПОК КЛИЕНТОВ")
    print("="*70 + "\n")

    count_created = 0
    total = len(df_main)

    docs_available = os.path.exists(DOCS_SOURCE_FOLDER)
    if not docs_available:
        print(f"  ВНИМАНИЕ: Папка с учредительными документами не найдена:")
        print(f"   {DOCS_SOURCE_FOLDER}")
        print(f"   Папки будут созданы БЕЗ учредительных документов\n")
    else:
        # Заранее находим все «Приказ на сотрудника» файлы/папки
        prikaz_files = get_prikaz_sotrudnika_files(DOCS_SOURCE_FOLDER)
        doverennost_files = get_doverennost_files(DOCS_SOURCE_FOLDER)
        if prikaz_files:
            print(f" Найдено 'Приказ на сотрудника': {len(prikaz_files)} шт. → {prikaz_files}")
        else:
            print(f" ВНИМАНИЕ: Файлы 'Приказ на сотрудника' не найдены в источнике")

        print(f" Источник учредительных документов: {DOCS_SOURCE_FOLDER}")
        print(f" Документов для копирования: {len(DOCS_TO_COPY)} + {len(prikaz_files)} приказов на сотрудника\n")

    for idx, row in df_main.iterrows():
        fio = str(row['FIO']).strip()
        iin = str(row['IIN']).strip().zfill(12)

        if not fio or not iin:
            print(f"[{idx+1}/{total}]   Пропущено: пустое ФИО или ИИН")
            continue

        folder_path = ensure_row_folder(row, TARGET_BASE)
        count_created += 1

        if docs_available:
            # 1. Копируем стандартные документы из DOCS_TO_COPY
            for doc_name in DOCS_TO_COPY:
                source_path = os.path.join(DOCS_SOURCE_FOLDER, doc_name)
                target_path = os.path.join(folder_path, doc_name)

                if os.path.exists(target_path):
                    continue

                if os.path.exists(source_path):
                    copy_item(source_path, target_path, doc_name)
                else:
                    print(f"     Не найден: '{doc_name}' — проверьте источник")

            # 2. Копируем ВСЕ файлы/папки «Приказ на сотрудника»
            for prikaz_name in prikaz_files:
                actual_source = os.path.join(DOCS_SOURCE_FOLDER, prikaz_name)
                actual_target = os.path.join(folder_path, prikaz_name)

                if os.path.exists(actual_target):
                    continue

                copy_item(actual_source, actual_target, prikaz_name)

            for doverennost_name in doverennost_files:
                actual_source = os.path.join(DOCS_SOURCE_FOLDER, doverennost_name)
                actual_target = os.path.join(folder_path, doverennost_name)

                if os.path.exists(actual_target):
                    continue

                copy_item(actual_source, actual_target, doverennost_name)

        if (idx + 1) % 10 == 0 or (idx + 1) % 50 == 0:
            print(f"[{idx+1}/{total}] Создано папок: {count_created}")

    print(f"\n Создано папок: {count_created}/{total}")

    if docs_available:
        print(f" В каждую папку скопировано до {len(DOCS_TO_COPY)} учредительных документов")
        print(f" + {len(prikaz_files)} приказов на сотрудника")

    print()

    safe_update_summary("СОЗДАНИЕ ПАПОК", {
        "found": count_created,
        "total": total,
        "not_found": []
    })

    return count_created, 0