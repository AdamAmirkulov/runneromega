# === ОБЩИЙ ЛОГ ПО БЛОКАМ ===

import os
from config import TARGET_BASE

# Папка, где находятся папки клиентов
BASE_FOLDER = TARGET_BASE


LOG_SUMMARY = {}
def run(df_main):
    # На всякий случай — создаём её, если вдруг нет
    os.makedirs(BASE_FOLDER, exist_ok=True)

# Один ОДИН общий файл лога для всего скрипта — СРАЗУ внутри Папки_для_исков
    log_file_path = os.path.join(BASE_FOLDER, "лог_сбор_документов.txt")

# Сюда каждый блок будет складывать свои итоги (ОФЕРТЫ, УВЕДОМЛЕНИЯ и т.д.)
    print("Лог будет писаться сюда:", log_file_path)
    return 0, 0