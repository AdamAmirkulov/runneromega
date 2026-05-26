# sbordoc/oferty.py
import os
import shutil
import pandas as pd

from scripts.config import (
    MAIN_EXCEL,
    TARGET_BASE,
    OFFERS_SOURCE,
)

from scripts.utils import (
    normalize,
    ensure_client_folder,
    safe_log,
    safe_update_summary,
)

def run(df_main=None):
    """
    Блок: Оферты
    Потокобезопасен, готов к ThreadPoolExecutor
    """

    # Если DataFrame не передан — читаем сами (fallback)
    if df_main is None:
        df_main = pd.read_excel(MAIN_EXCEL, usecols=[1, 2, 3], header=0)
        df_main.columns = ['Product', 'FIO', 'IIN']
        df_main['IIN'] = df_main['IIN'].astype(str).str.zfill(12)

    print("📄 [ОФЕРТЫ] Старт обработки")

    success = 0
    failed = 0
    not_found = []

    for _, row in df_main.iterrows():
        product = str(row['Product']).strip()
        fio = str(row['FIO']).strip()
        iin = str(row['IIN']).strip().zfill(12)

        try:
            # Папка клиента (безопасно)
            client_folder = ensure_client_folder(iin, fio, TARGET_BASE)

            normalized_fio = normalize(fio)

            copied = False

            # Поиск оферты
            for root, _, files in os.walk(OFFERS_SOURCE):
                for file in files:
                    if normalized_fio in normalize(file):
                        src = os.path.join(root, file)
                        dst = os.path.join(client_folder, file)

                        shutil.copy2(src, dst)
                        copied = True
                        break
                if copied:
                    break

            if copied:
                success += 1
            else:
                failed += 1
                not_found.append(f"{fio} | {iin}")
                safe_log(
                    f"[ОФЕРТЫ] Не найден файл | {fio} | {iin} | продукт: {product}"
                )

        except Exception as e:
            failed += 1
            not_found.append(f"{fio} | {iin}")
            safe_log(
                f"[ОФЕРТЫ][ОШИБКА] {fio} | {iin} | {e}"
            )

    # Обновляем общий summary (под замком)
    safe_update_summary("ОФЕРТЫ", {
        "found": success,
        "total": len(df_main),
        "not_found": not_found,
    })

    print(f"📄 [ОФЕРТЫ] Готово: ✅ {success} | ❌ {failed}")

    return success, failed
