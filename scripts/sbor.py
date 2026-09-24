# sbor.py
import os
import sys
import io
import argparse

#from scripts import poiskvsk
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# ─────────────── Добавляем корень проекта в sys.path ───────────────
# Этот код гарантирует, что Python увидит пакет `scripts`
project_root = os.path.dirname(os.path.abspath(__file__))  # путь к папке scripts
project_root = os.path.dirname(project_root)               # идем на уровень выше (OmegaRunner)
if project_root not in sys.path:
    sys.path.insert(0, project_root)
# ────────────────────────────────────────────────────────────────

# ═══════════════════════════════════════════════════════════════
# ОБРАБОТКА АРГУМЕНТОВ ОТ FASTAPI
# ═══════════════════════════════════════════════════════════════

def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('--workdir', type=str, default=None)
    parser.add_argument('--excel_path', type=str, default=None)
    parser.add_argument('--company_id', type=str, default=None)
    return parser.parse_args()

args = parse_args()

if not args.company_id or not args.company_id.strip():
    print("❌ ОШИБКА: не передан --company_id — компания не определена, запуск остановлен.")
    sys.exit(1)

if args.excel_path and args.excel_path.strip():
    os.environ['OVERRIDE_EXCEL_PATH'] = args.excel_path.strip()

os.environ['COMPANY_ID'] = args.company_id.strip()  # до import config!
# ✅ ПЕРЕОПРЕДЕЛЯЕМ ПУТЬ К EXCEL ПЕРЕД ИМПОРТОМ config
if args.excel_path and args.excel_path.strip():
    os.environ['OVERRIDE_EXCEL_PATH'] = args.excel_path.strip()



import pandas as pd
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

# ─────────────── Импорт конфигурации и модулей ───────────────
from config import MAIN_EXCEL, LOG_SUMMARY
from utils import norm_uid, add_folder_names
from sbordoc_files import (
    logi,
    sozdaniepapok,
    oferty,
    uvedomlenie,
    dogovor,
    reestrdogovora,
    ispolnadpisi,
    gosposhliny,
    dosudebnaya,
    screen,
    raschet,
    iskovoezayav,
    postobotmenee,
    formirovaniepool,
    formirovaniemassovogo,
    logirovanie,
)
# ────────────────────────────────────────────────────────────────

# Дальше оставляем твой основной код main() без изменений
def print_separator(title: str):
    print("\n" + "="*70)
    print(f"  {title}")
    print("="*70 + "\n")

def main():
    start_time = time.time()
    
    print("""
            СИСТЕМА ОБРАБОТКИ ДОКУМЕНТОВ ДЛЯ ИСКОВ
    """)
    
    print_separator("ЗАГРУЗКА ДАННЫХ ИЗ EXCEL")
    
    try:
        # ✅ ЧИТАЕМ ИЗ ПЕРЕОПРЕДЕЛЁННОГО ПУТИ
        df = pd.read_excel(MAIN_EXCEL, usecols=[0, 1, 2, 3], header=0)
        # Колонка A — «Уникальный номер» (EID займа). В старых отчётах её
        # может не быть — тогда A уже «Продукт», и работаем как раньше.
        if str(df.columns[0]).strip().lower().startswith('уникальн'):
            df.columns = ['UID', 'Product', 'FIO', 'IIN']
        else:
            df = df.iloc[:, :3]
            df.columns = ['Product', 'FIO', 'IIN']
            df['UID'] = ""
        df['IIN'] = df['IIN'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(12)
        df['UID'] = df['UID'].apply(norm_uid)
        
        print(f"✅ Загружено записей: {len(df)}")
        print(f"📋 Уникальных продуктов: {df['Product'].nunique()}")
        print(f"👥 Уникальных клиентов: {df['IIN'].nunique()}\n")
    
    except Exception as e:
        print(f"❌ КРИТИЧЕСКАЯ ОШИБКА загрузки Excel: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

    # ───────────── ОЧИСТКА DataFrame ─────────────
    # Удаляем пустые строки по ключевым столбцам
    df = df.dropna(subset=['FIO', 'IIN'])

    # Убираем пробелы в строках
    df['FIO'] = df['FIO'].str.strip()
    df['IIN'] = df['IIN'].str.strip()
    df['Product'] = df['Product'].astype(str).str.strip()

    # Одна строка = один займ. Дубликаты убираем по Уникальному номеру,
    # а не по ФИО+ИИН: у должника может быть несколько займов, и на каждый
    # нужна своя папка (иначе документы займов смешиваются в одной).
    if (df['UID'] != "").all():
        df = df.drop_duplicates(subset=['UID'])
    else:
        df = df.drop_duplicates(subset=['FIO', 'IIN'])
    df = add_folder_names(df).reset_index(drop=True)

    multi = df[df['IIN'].duplicated(keep=False)]
    if not multi.empty:
        print(f"Должников с несколькими займами — отдельная папка на каждый займ: {len(multi)}")
        for name in multi['FolderName']:
            print(f"   {name}")

    print(f"Строк после очистки: {len(df)}")
    
    print(f"Загружено записей: {len(df)}")
    print(f"Уникальных продуктов: {df['Product'].nunique()}")
    print(f"Уникальных клиентов: {df['IIN'].nunique()}\n")
    
    # ───────────── Этап 1: Подготовка ─────────────
    print_separator("ЭТАП 1: ПОДГОТОВКА")
   
    preparation_blocks = [
        ("Инициализация логов", logi.run),
        ("Создание папок клиентов", sozdaniepapok.run),
        #("Поиск адресов в СК + Привязка к суду", poiskvsk.run),
    ]
    
    for name, block_func in preparation_blocks:
        print(f" {name}...")
        try:
            success, failed = block_func(df)
            print(f"   Успешно: {success} |  Ошибок: {failed}\n")
        except Exception as e:
            print(f"   ОШИБКА: {e}\n")
            import traceback
            traceback.print_exc()
    
    # ───────────── Этап 2: Сбор документов ─────────────
    print_separator(" ЭТАП 2: СБОР ДОКУМЕНТОВ (параллельная обработка)")
    
    document_blocks = [
        ("Постановление об отмене надписи", postobotmenee.run),
        ("Договор цессии", dogovor.run),
        ("Исполнительные надписи", ispolnadpisi.run),
        ("Реестр договора цессии", reestrdogovora.run),
        ("Госпошлины", gosposhliny.run),        
        ("Скрин отправки претензии", screen.run),    
        ("Исковое заявление", iskovoezayav.run),
        ("Уведомление об уступке права требования", uvedomlenie.run),
        ("Оферта", oferty.run),
        
    ]
    
    print(f"Запуск {len(document_blocks)} блоков параллельно (макс. 8 потоков)...\n")
    
    document_results = []
    
    with ThreadPoolExecutor(max_workers=8) as executor:
        futures = {executor.submit(block_func, df): name for name, block_func in document_blocks}
        
        for future in as_completed(futures):
            block_name = futures[future]
            try:
                success, failed = future.result()
                document_results.append({"name": block_name, "success": success, "failed": failed})
                print(f"{block_name}: успешно={success}, ошибок={failed}")
            except Exception as e:
                print(f"{block_name}: ОШИБКА - {e}")
                document_results.append({"name": block_name, "success": 0, "failed": len(df), "error": str(e)})
    print_separator("ЭТАП 2б: ФОРМИРОВАНИЕ ДОКУМЕНТОВ WORD (последовательно)")

    word_blocks = [
        ("Досудебная претензия", dosudebnaya.run),
        ("Расчет задолженности", raschet.run),
    ]

    for name, block_func in word_blocks:
        print(f"⏳ {name}...")
        try:
            success, failed = block_func(df)
            print(f"   Успешно: {success} | Ошибок: {failed}\n")
        except Exception as e:
            print(f"   ОШИБКА: {e}\n")
    
    # ───────────── Этап 3: Финализация ─────────────
    print_separator(" ЭТАП 3: ФИНАЛИЗАЦИЯ")
    
    final_blocks = [
        #("Формирование по Пулам", formirovaniepool.run),
        #("Массовое заявление в Медеуский суд", formirovaniemassovogo.run),
        ("Логирование сбора документов", logirovanie.run),
    ]
    
    for name, block_func in final_blocks:
        print(f"⏳ {name}...")
        try:
            success, failed = block_func(df)
            print(f"    Успешно: {success} | ❌ Ошибок: {failed}\n")
        except Exception as e:
            print(f"   ОШИБКА: {e}\n")
            import traceback
            traceback.print_exc()
    
    end_time = time.time()
    total_time = end_time - start_time
    
    print("\n" + "="*70)
    print("   ИТОГОВЫЙ ОТЧЁТ")
    print("="*70 + "\n")
    
    if LOG_SUMMARY:
        total_found = 0
        total_not_found = 0
        
        for block_name, stats in LOG_SUMMARY.items():
            found = stats.get('found', 0)
            total = stats.get('total', 0)
            not_found_count = len(stats.get('not_found', []))
            
            total_found += found
            total_not_found += not_found_count
            
            percentage = (found / total * 100) if total > 0 else 0
        
            print(f" {block_name}:")
            print(f"    Найдено: {found}/{total} ({percentage:.1f}%)")
            if not_found_count > 0:
                print(f"    Не найдено: {not_found_count}")
            print()
        
        print("-" * 70)
        print(f" ОБЩАЯ СТАТИСТИКА:")
        print(f"    Всего успешно обработано: {total_found}")
        print(f"    Всего не найдено: {total_not_found}")
        print(f"     Время выполнения: {total_time:.2f} секунд ({total_time/60:.1f} минут)")
        print("-" * 70)
    
    print("""
    ВСЕ БЛОКИ ЗАВЕРШЕНЫ!
    """)
    
    print(f" Подробный лог ошибок: log_not_copied.txt\n")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n  Обработка прервана пользователем!")
    except Exception as e:
        print(f"\n\n КРИТИЧЕСКАЯ ОШИБКА: {e}")
        import traceback
        traceback.print_exc()
