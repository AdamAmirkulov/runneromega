# === ФИНАЛЬНАЯ СВОДКА ПО ВСЕМ БЛОКАМ ===

import os
from config import TARGET_BASE, LOG_SUMMARY
from pathlib import Path
# Папка, где находятся папки клиентов

os.makedirs(TARGET_BASE, exist_ok=True)
BASE_FOLDER = TARGET_BASE
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


def run(df_main):
    write_final_summary(LOG_SUMMARY)
    # sbor.py ждёт (успешно, ошибок) от каждого блока
    return len(LOG_SUMMARY), 0