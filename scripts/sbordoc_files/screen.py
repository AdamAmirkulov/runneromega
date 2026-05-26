"""
Блок 10: Формирование скринов отправки досудебных претензий
"""
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

from config import MAIN_EXCEL, ROOT, TARGET_BASE
from utils import ensure_client_folder, safe_log, safe_update_summary

# ═══════════════════════════════════════════════════════════════
# НАСТРОЙКИ
# ═══════════════════════════════════════════════════════════════

SHEET_NAME = "Данные для шаблонов"

TEMPLATE_PATH = Path(
    rf"{ROOT}\Документы для подачи Исков\Шаблоны документов\Скрин (Досудебная претензия).html"
)

# ═══════════════════════════════════════════════════════════════
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ══════════════════════════════════════════

def nice_case(s: str) -> str:
    """Приводит строку к формату 'Слово Слово'"""
    if not s:
        return ""
    parts = str(s).split()
    return " ".join(p.capitalize() for p in parts)


def load_template():
    """Загружает HTML шаблон"""
    text = TEMPLATE_PATH.read_text(encoding="utf-8")
    return Template(text)


def html_to_jpeg(html_path: Path, jpeg_path: Path):
    """Рендерит HTML в JPEG через Selenium"""
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

    # Конвертируем PNG в JPEG
    img = Image.open(tmp_png).convert("RGB")
    img.save(jpeg_path, "JPEG", quality=95)
    tmp_png.unlink(missing_ok=True)


def random_time_str():
    """Генерирует случайное время 09:00 - 17:59"""
    hour = random.randint(9, 17)
    minute = random.randint(0, 59)
    return f"{hour:02d}:{minute:02d}"


def get_first_existing_value(row: pd.Series, candidates: list[str], default: str = "") -> str:
    """Берёт первое непустое значение из списка возможных названий колонок"""
    for col in candidates:
        if col in row.index:
            val = row.get(col)
            if val is not None:
                s = str(val).strip()
                if s and s.lower() != "nan":
                    return s
    return default


# ═══════════════════════════════════════════════════════════════
# ОСНОВНАЯ ФУНКЦИЯ
# ═══════════════════════════════════════════════════════════════

def run(df_main):
    """
    Формирует скрины отправки досудебных претензий из HTML-шаблона.
    
    Логика:
    1. Загружает данные из листа "Данные для шаблонов"
    2. Для каждого клиента:
       - Заполняет HTML шаблон данными
       - Рендерит HTML в JPEG через Selenium
       - Сохраняет в папку клиента
    
    Args:
        df_main: DataFrame с колонками ['Product', 'FIO', 'IIN']
    
    Returns:
        tuple: (count_success, count_failed)
    """
    print("\n" + "="*70)
    print("  📸 БЛОК 10: СКРИНЫ ОТПРАВКИ ДОСУДЕБНЫХ ПРЕТЕНЗИЙ")
    print("="*70 + "\n")

    count_success = 0
    count_failed = 0
    not_found_list = []

    # Проверка существования файлов
    if not os.path.exists(MAIN_EXCEL):
        print(f"❌ ОШИБКА: Excel файл не найден: {MAIN_EXCEL}")
        safe_log(f"[СКРИНЫ] Excel файл не найден: {MAIN_EXCEL}")
        return 0, len(df_main)

    if not TEMPLATE_PATH.exists():
        print(f"❌ ОШИБКА: Шаблон не найден: {TEMPLATE_PATH}")
        safe_log(f"[СКРИНЫ] Шаблон не найден: {TEMPLATE_PATH}")
        return 0, len(df_main)

    # Загружаем шаблон
    try:
        template = load_template()
    except Exception as e:
        print(f"❌ ОШИБКА при загрузке шаблона: {e}")
        safe_log(f"[СКРИНЫ] Ошибка загрузки шаблона: {e}")
        return 0, len(df_main)

    # Загружаем данные из Excel
    try:
        df = pd.read_excel(MAIN_EXCEL, sheet_name=SHEET_NAME)
    except Exception as e:
        print(f"❌ ОШИБКА при загрузке листа '{SHEET_NAME}': {e}")
        safe_log(f"[СКРИНЫ] Ошибка загрузки листа: {e}")
        return 0, len(df_main)

    print(f"📊 Загружено данных из Excel: {len(df)} строк")
    print(f"📋 Клиентов для обработки: {len(df_main)}\n")

    total = len(df_main)

    # Обрабатываем каждого клиента
    for idx, row in df.iterrows():
        fio = ""
        iin = ""

        try:
            # Получаем ФИО и ИИН
            fio_raw = row.get("ФИО")
            iin_raw = row.get("ИИН")

            fio = nice_case(str(fio_raw)) if fio_raw is not None else ""

            # ИИН: всегда 12 символов, с нулями слева
            if iin_raw is not None:
                raw_iin_str = str(iin_raw).strip()
                if "." in raw_iin_str:
                    raw_iin_str = raw_iin_str.split(".")[0]
                iin = raw_iin_str.zfill(12)
            else:
                iin = ""

            if not fio or not iin:
                continue

            print(f"[{idx+1}/{len(df)}] {fio} ({iin})")

            # Получаем папку клиента
            target_folder = ensure_client_folder(iin, fio, TARGET_BASE)

            # Получаем email
            email = str(row.get("Email", ""))

            # Дата выдачи займа
            loan_date = row.get("Дата выдачи займа")
            if isinstance(loan_date, (datetime, pd.Timestamp)) and pd.notna(loan_date):
                loan_date_str = loan_date.strftime("%d.%m.%Y")
            else:
                loan_date_str = str(loan_date) if loan_date is not None and pd.notna(loan_date) else ""

            # Дата договора цессии
            cession_date = row.get("Дата договора цессии")
            if isinstance(cession_date, (datetime, pd.Timestamp)) and pd.notna(cession_date):
                cession_date_str = cession_date.strftime("%d.%m.%Y")
            else:
                cession_date_str = str(cession_date) if cession_date is not None and pd.notna(cession_date) else ""

            # Дата email (минус 7 дней)
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
                email_date_str = str(email_date) if email_date is not None and pd.notna(email_date) else ""

            # Добавляем случайное время
            email_datetime_str = f"{email_date_str} {random_time_str()}"

            # Кредитор: берём из Excel
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

            # Подготовка данных для шаблона
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
                "creditor": creditor,
            }

            # Формируем имена файлов
            base_name = f"Скрин отправки Досудебной претензии, {fio}, {iin}"
            html_path = Path(target_folder) / (base_name + ".html")
            img_path = Path(target_folder) / (base_name + ".jpg")

            # Заполняем шаблон
            html_text = template.render(**data)
            html_path.write_text(html_text, encoding="utf-8")

            # Рендерим в JPEG
            html_to_jpeg(html_path, img_path)

            # Удаляем промежуточный HTML
            try:
                html_path.unlink()
            except OSError:
                pass

            count_success += 1
            print(f"   ✅ Создан: {img_path.name}")

        except Exception as e:
            count_failed += 1
            label = f"{fio or 'Неизвестный'}, {iin or 'ИИН не указан'}"
            not_found_list.append(label)
            print(f"   ❌ Ошибка: {e}")
            safe_log(f"[СКРИНЫ] Ошибка для {label}: {e}")
            continue

    # Итоговая статистика
    print("\n" + "-"*70)
    print(f"📊 ИТОГО:")
    print(f"   ✅ Успешно создано: {count_success}/{len(df)}")
    print(f"   ❌ Ошибок: {count_failed}/{len(df)}")
    print("-"*70 + "\n")

    # Сохраняем в LOG_SUMMARY
    safe_update_summary("СКРИНЫ ДОСУДЕБНЫЕ ПРЕТЕНЗИИ", {
        "found": count_success,
        "total": len(df),
        "not_found": not_found_list
    })

    return count_success, count_failed