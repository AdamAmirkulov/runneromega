#!/usr/bin/env python
# coding: utf-8

# ============================================================
#  НАСТРОЙКИ
# ============================================================
err_not_final = 'в последнем скаченном АСИОИП файле нет "итога"'

dict_log_pas = {
    "kpi": {
        "log": "810813301334_230240016634",
        "pas": "Qazaq123456*"
    }
}

# --- БД ---
DB_SERVER   = "DBSRV"
DB_DATABASE = "crm"
DB_USERNAME = "user"
DB_PASSWORD = "*****"

# --- Папка для временных скачанных файлов ---
download_dir = r"C:/Users/User/Desktop/py scripts/downloads"

# --- Куда сохранить итоговый Excel ---
OUTPUT_EXCEL = r"C:\Users\User\Desktop\Omega Scripts\Результат_АСИОИП.xlsx"

# ============================================================
#  ИМПОРТЫ
# ============================================================
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

import time
import pandas as pd
import pyodbc
from pathlib import Path
from datetime import datetime
import os
import traceback

# ============================================================
#  ПОЛУЧЕНИЕ ДАННЫХ ИЗ БД
# ============================================================
SQL_QUERY = """
SELECT 
    l.EID AS [Уникальный номер],
    CAST(LEFT(dF246.F249, 50) AS VARCHAR(50)) AS [Продукт],
    CAST(l.F287 AS VARCHAR(50)) AS [Номер договора],
    c.FIO  AS [ФИО],
    c.F293 AS [ИИН],
    s.Caption AS [Статус кредита],
    l.F34 AS [Остаток задолженности],
    l.F62 AS [ЧСИ],
    l.F146 AS [Номер исполнительного производства],
    NULLIF(constF236.Caption, '') AS [Статус АИС ОИП], 
    l.F352 AS [Извещение о временном ограничении на выезд],
    l.F351 AS [Постановление об ограничении на выезд],
    l.F353 AS [Снятие ограничения на выезд]
FROM loans l
LEFT JOIN clients   c         ON c.ID = l.CID
LEFT JOIN states    s         ON s.ID = l.State
LEFT JOIN Dictionary dF246    ON dF246.ID = l.F246
LEFT JOIN Constants constF236 ON constF236.ID = l.F236
WHERE 
    s.Caption NOT IN (N'Погашен', N'Обратный выкуп', N'В графике',
                      N'Армия', N'Умерший', N'Банкрот', N'ВосПС', N'Мошенничество')
    AND l.F209 LIKE N'%Омега%'
    AND constF236.Caption = N'На исполнении'
    AND l.F34 > 173000
ORDER BY l.EID ASC;
"""


def load_data_from_db():
    """Подключается к SQL Server и возвращает DataFrame."""
    print("🔌 Подключение к БД...")
    conn_str = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={DB_SERVER};"
        f"DATABASE={DB_DATABASE};"
        f"UID={DB_USERNAME};"
        f"PWD={DB_PASSWORD};"
        f"TrustServerCertificate=yes;"
    )
    try:
        conn = pyodbc.connect(conn_str, timeout=30)
        df = pd.read_sql(SQL_QUERY, conn)
        conn.close()
        print(f"✅ Загружено {len(df)} записей из БД")
        return df
    except Exception as e:
        print(f"❌ Ошибка подключения к БД: {e}")
        raise


# ============================================================
#  SELENIUM — ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================
def to_site(loading_path_all):
    url = "https://aisoip.adilet.gov.kz/cabinet/exec-productions"

    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")

    prefs = {
        "profile.default_content_settings.popups": 0,
        "download.default_directory": loading_path_all.replace("/", "\\"),
        "directory_upgrade": True,
        "download.prompt_for_download": False,
        "safebrowsing.enabled": False,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "profile.content_settings.exceptions.automatic_downloads.*.setting": 1,
        "profile.managed_default_content_settings.images": 2,
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-logging")
    options.add_argument("--disable-web-security")
    options.add_argument("--disk-cache-size=0")
    options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    options.add_experimental_option("useAutomationExtension", False)

    driver = webdriver.Chrome(options=options)
    driver.execute_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    driver.implicitly_wait(5)
    driver.get(url)
    time.sleep(1.5)
    return driver


def check_if_logged_in(driver):
    try:
        driver.current_url
    except Exception:
        return None
    try:
        driver.find_element(By.CLASS_NAME, "field")
        return False
    except NoSuchElementException:
        try:
            driver.find_element(By.CLASS_NAME, "cabinet-header")
            return True
        except Exception:
            return None
    except Exception:
        return None


def write_login(driver, name_login):
    try:
        login_global = dict_log_pas[name_login]["log"]
        pasword_global = dict_log_pas[name_login]["pas"]

        WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((By.CLASS_NAME, "field"))
        )
        fields = driver.find_elements(By.CLASS_NAME, "field")

        login_field = fields[0].find_element(By.CSS_SELECTOR, "input")
        login_field.clear()
        login_field.send_keys(login_global)
        time.sleep(0.3)

        password_field = fields[1].find_element(By.CSS_SELECTOR, "input")
        password_field.clear()
        password_field.send_keys(pasword_global)
        time.sleep(0.3)

        class_field = "ui.primary.button.fluid"
        driver.find_elements(By.CLASS_NAME, class_field)[0].click()

        WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((By.CLASS_NAME, "cabinet-header"))
        )
        print("✅ Успешный вход в систему")
        time.sleep(1)
        return True
    except Exception as e:
        print(f"❌ Ошибка при логине: {e}")
        return False


def relogin_if_needed(driver, dl_dir, name_login="kpi"):
    try:
        status = check_if_logged_in(driver)

        if status is None:
            print("⚠️ Драйвер не отвечает! Полный перезапуск...")
            try:
                driver.quit()
            except Exception:
                pass
            time.sleep(3)
            driver = to_site(dl_dir)
            if not write_login(driver, name_login):
                raise Exception("Не удалось залогиниться после перезапуска")
            to_targer_sheet_1(driver)
            to_targer_sheet_2(driver)
            return driver, True

        if status is False:
            print("⚠️ Обнаружен разлогин! Переподключаемся...")
            if not write_login(driver, name_login):
                raise Exception("Не удалось залогиниться")
            to_targer_sheet_1(driver)
            to_targer_sheet_2(driver)
            return driver, True

        return driver, False

    except Exception as e:
        print(f"❌ Критическая ошибка при перелогине: {e}")
        try:
            driver.quit()
        except Exception:
            pass
        time.sleep(3)
        driver = to_site(dl_dir)
        if write_login(driver, name_login):
            to_targer_sheet_1(driver)
            to_targer_sheet_2(driver)
            return driver, True
        raise


def to_targer_sheet_1(driver):
    WebDriverWait(driver, 8).until(
        EC.presence_of_element_located((By.CLASS_NAME, "cabinet-header"))
    )
    field = driver.find_element(By.CLASS_NAME, "cabinet-header")
    field.find_elements(By.CLASS_NAME, "navigation-item")[1].click()
    time.sleep(1)


def to_targer_sheet_2(driver):
    class_field = "v-slide-group__content.v-tabs-bar__content"
    field = driver.find_element(By.CLASS_NAME, class_field)
    field.find_elements(By.CSS_SELECTOR, "div")[3].click()
    time.sleep(0.7)


def iin_insert(driver, iin_iter):
    field = driver.find_element(By.CLASS_NAME, "row.mt-8.mt-md-12")
    slot = field.find_elements(By.CLASS_NAME, "v-input__slot")[0]
    inp = slot.find_element(By.CSS_SELECTOR, "input")
    inp.clear()
    inp.send_keys(iin_iter)
    time.sleep(0.2)


def buttom_find(driver):
    btn_css = ".mb-2.mr-2.mr-md-5.v-btn.v-btn--is-elevated.v-btn--has-bg.theme--light.v-size--default.primary"
    driver.find_element(By.CSS_SELECTOR, btn_css).click()
    time.sleep(2)


def open_production_in_execution(driver):
    time.sleep(1.5)
    productions = driver.find_elements(By.CSS_SELECTOR, ".exec-productions-item")
    if not productions:
        print("❌ Производства не найдены")
        return False

    for prod in productions:
        try:
            labels = prod.find_elements(By.CSS_SELECTOR, "p.label")
            for label in labels:
                if "Статус" in label.text:
                    status_value = label.find_element(
                        By.XPATH, "following-sibling::p[@class='value']"
                    )
                    if "На исполнении" in status_value.text.strip():
                        print("✅ Найдено производство со статусом 'На исполнении'")
                        prod.find_element(
                            By.CSS_SELECTOR, "a[style='text-decoration: underline;']"
                        ).click()
                        time.sleep(1.5)
                        return True
        except Exception:
            continue

    print("⚠️ Производство 'На исполнении' не найдено")
    return False


def click_export_excel_in_documents(driver, timeout=12):
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, ".debtor-documents.info-row-item")
            )
        )
        block = driver.find_element(By.CSS_SELECTOR, ".debtor-documents.info-row-item")
        btn = block.find_element(By.XPATH, ".//button[@title='экспорт в Excel']")

        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", block)
        time.sleep(0.3)
        driver.execute_script(
            "document.querySelectorAll('.v-overlay__scrim').forEach(o => o.style.display='none');"
        )
        driver.execute_script("arguments[0].click();", btn)
        print("✅ Кнопка экспорта нажата")
        return True
    except Exception as e:
        print(f"❌ Ошибка экспорта: {e}")
        try:
            all_btns = driver.find_elements(
                By.XPATH, "//button[@title='экспорт в Excel']"
            )
            if len(all_btns) >= 2:
                driver.execute_script("arguments[0].click();", all_btns[1])
                print("✅ Альтернативная кнопка нажата")
                return True
        except Exception as e2:
            print(f"❌ Альтернатива не сработала: {e2}")
        return False


def close_production_card(driver):
    try:
        css = ".d-print-none.v-btn.v-btn--absolute.v-btn--icon.v-btn--right.v-btn--round.v-btn--top.theme--light.v-size--small"
        driver.find_element(By.CSS_SELECTOR, css).click()
        time.sleep(0.3)
    except Exception as e:
        print(f"⚠️ Ошибка закрытия карточки: {e}")


def delite_filter(driver):
    try:
        field = WebDriverWait(driver, 4).until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, ".v-btn.v-btn--outlined.theme--light.v-size--default")
            )
        )
        field.click()
        time.sleep(0.2)
    except Exception as e:
        print(f"⚠️ Ошибка удаления фильтра: {e}")


def wait_excel_download(dl_dir, timeout=30):
    start = time.time()
    dl_dir = Path(dl_dir)
    last_size = 0
    stable_count = 0

    while time.time() - start < timeout:
        files = list(dl_dir.glob("*.xls*"))
        if files:
            current_size = files[0].stat().st_size
            if current_size == last_size and current_size > 0:
                stable_count += 1
                if stable_count >= 2:
                    print(f"✅ Файл загружен ({current_size} байт)")
                    return files[0]
            else:
                stable_count = 0
            last_size = current_size
        time.sleep(0.5)

    return None


# ============================================================
#  ПАРСИНГ СКАЧАННОГО EXCEL ИЗ АСИОИП
# ============================================================
def extract_dates_from_downloaded_excel(excel_file_path):
    """Извлекает даты постановления, извещения и снятия из файла АСИОИП."""
    try:
        print(f"📄 Обработка: {excel_file_path.name}")
        df = pd.read_excel(excel_file_path)
        df.columns = df.columns.str.strip()

        if (
            "Наименование документа" not in df.columns
            or "Дата подписания" not in df.columns
        ):
            print("❌ Не найдены нужные колонки")
            return None, None, None

        postanovlenie_dates = []
        izveshenie_dates = []
        removal_dates = []

        for _, row in df.iterrows():
            doc_name = str(row["Наименование документа"]).lower()
            date_str = str(row["Дата подписания"])

            date_obj = None
            for fmt in ["%d.%m.%Y %H:%M:%S", "%d.%m.%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"]:
                try:
                    date_obj = datetime.strptime(date_str, fmt)
                    break
                except Exception:
                    continue

            if not date_obj:
                continue

            date_formatted = date_obj.strftime("%d.%m.%Y")

            # 1. СНЯТИЕ
            if (
                ("снят" in doc_name or "сняти" in doc_name or "отмен" in doc_name)
                and "огранич" in doc_name
                and ("выезд" in doc_name or "vyezd" in doc_name)
            ):
                removal_dates.append(date_formatted)
                print(f"  ✅ СНЯТИЕ: {date_formatted}")
                continue

            # 2. ИЗВЕЩЕНИЕ
            if "извещ" in doc_name and (
                "огранич" in doc_name or "ogranichenii" in doc_name
            ) and ("выезд" in doc_name or "vyezd" in doc_name):
                izveshenie_dates.append(date_formatted)
                print(f"  ✅ ИЗВЕЩЕНИЕ: {date_formatted}")
                continue

            # 3. ПОСТАНОВЛЕНИЕ
            if (
                ("постановлен" in doc_name or "postanovlenie" in doc_name or "kauly boryshkerdin" in doc_name)
                and ("огранич" in doc_name or "ogranichenii" in doc_name or "shekteu" in doc_name)
                and ("выезд" in doc_name or "vyezd" in doc_name or "shyguyn" in doc_name)
            ):
                postanovlenie_dates.append(date_formatted)
                print(f"  ✅ ПОСТАНОВЛЕНИЕ: {date_formatted}")

        date_postanovlenie = "; ".join(sorted(set(postanovlenie_dates))) if postanovlenie_dates else None
        date_izveshenie    = "; ".join(sorted(set(izveshenie_dates)))    if izveshenie_dates    else None
        date_removal       = "; ".join(sorted(set(removal_dates)))       if removal_dates       else None

        return date_postanovlenie, date_izveshenie, date_removal

    except Exception as e:
        print(f"❌ Ошибка обработки Excel: {e}")
        traceback.print_exc()
        return None, None, None


# ============================================================
#  ГЛАВНЫЙ ЦИКЛ
# ============================================================
def main():
    # 1. Загружаем данные из БД
    df = load_data_from_db()

    # Нормализуем ИИН (ведущие нули, 12 символов)
    df["ИИН"] = df["ИИН"].apply(
        lambda x: str(x).strip().zfill(12) if pd.notna(x) and str(x).strip() else None
    )

    # Фильтруем только строки без постановления (пустые / NaN)
    df_to_process = df[
        df["Постановление об ограничении на выезд"].isna()
        | (df["Постановление об ограничении на выезд"].astype(str).str.strip() == "")
    ].copy()

    iin_list = (
        df_to_process["ИИН"]
        .dropna()
        .astype(str)
        .str.strip()
        .str.zfill(12)
        .tolist()
    )

    print("📊 СТАТИСТИКА:")
    print(f"   ВСЕГО В БД:          {len(df)}")
    print(f"   БЕЗ ПОСТАНОВЛЕНИЯ:   {len(df_to_process)}")
    print(f"   К ОБРАБОТКЕ:         {len(iin_list)}")
    print("=" * 60)

    Path(download_dir).mkdir(parents=True, exist_ok=True)

    # 2. Запускаем браузер
    print("🚀 ЗАПУСК SELENIUM")
    driver = to_site(download_dir)
    write_login(driver, "kpi")
    to_targer_sheet_1(driver)
    to_targer_sheet_2(driver)

    processed = 0
    successful = 0
    errors = []
    start_time = time.time()

    START_FROM = 1
    start_index = START_FROM - 1

    for idx, iin in enumerate(iin_list[start_index:], start=START_FROM):
        iteration_start = time.time()
        iin = str(iin).strip().zfill(12)
        print(f"\n[{idx}/{len(iin_list)}] 🔍 ИИН: {iin}")

        max_retries = 3
        retry_count = 0

        while retry_count < max_retries:
            try:
                driver, was_relogin = relogin_if_needed(driver, download_dir)
                if was_relogin:
                    time.sleep(2)
                    to_targer_sheet_1(driver)
                    to_targer_sheet_2(driver)

                delite_filter(driver)
                iin_insert(driver, iin)
                buttom_find(driver)

                if not open_production_in_execution(driver):
                    time.sleep(2)
                    driver, was_relogin = relogin_if_needed(driver, download_dir)
                    if was_relogin:
                        time.sleep(2)
                        to_targer_sheet_1(driver)
                        to_targer_sheet_2(driver)
                        delite_filter(driver)
                        iin_insert(driver, iin)
                        buttom_find(driver)
                        if not open_production_in_execution(driver):
                            print("❌ Нет производства (после релогина)")
                            break
                    else:
                        print("❌ Нет производства")
                        break

                print("✅ Карточка открыта")
                time.sleep(1)

                # Очищаем старые файлы
                for old_file in Path(download_dir).glob("*.xls*"):
                    try:
                        old_file.unlink()
                    except Exception:
                        pass

                if not click_export_excel_in_documents(driver):
                    raise Exception("Не удалось экспортировать")

                excel_file = wait_excel_download(download_dir, timeout=35)

                if excel_file:
                    date_post, date_izv, date_rem = extract_dates_from_downloaded_excel(
                        excel_file
                    )

                    # Записываем в df по ИИН
                    mask = df["ИИН"] == iin
                    if mask.any():
                        if date_post:
                            df.loc[mask, "Постановление об ограничении на выезд"] = date_post
                        if date_izv:
                            df.loc[mask, "Извещение о временном ограничении на выезд"] = date_izv
                        if date_rem:
                            df.loc[mask, "Снятие ограничения на выезд"] = date_rem

                    if date_post or date_izv or date_rem:
                        successful += 1
                    else:
                        print("⚠️ Нет дат в документе")
                else:
                    raise Exception("Файл не скачан")

                close_production_card(driver)
                time.sleep(0.5)
                delite_filter(driver)

                if excel_file and os.path.exists(excel_file):
                    try:
                        os.remove(excel_file)
                    except Exception:
                        pass

                processed += 1
                print(f"⏱️ Время: {time.time() - iteration_start:.1f}с")
                break

            except Exception as e:
                retry_count += 1
                err_msg = str(e)

                if any(
                    x in err_msg.lower()
                    for x in ["invalid session", "disconnected", "timed out", "connection"]
                ):
                    print(f"❌ Критическая ошибка драйвера (попытка {retry_count}/{max_retries})")
                    try:
                        driver.quit()
                    except Exception:
                        pass
                    time.sleep(3)
                    try:
                        driver = to_site(download_dir)
                        write_login(driver, "kpi")
                        to_targer_sheet_1(driver)
                        to_targer_sheet_2(driver)
                        print("✅ Драйвер восстановлен!")
                    except Exception as e2:
                        print(f"❌ Не удалось восстановить драйвер: {e2}")
                        if retry_count >= max_retries:
                            errors.append(iin)
                            break
                else:
                    print(f"❌ Попытка {retry_count}/{max_retries}: {e}")

                if retry_count < max_retries:
                    print("🔄 Повтор...")
                    time.sleep(3)
                    try:
                        close_production_card(driver)
                        delite_filter(driver)
                    except Exception:
                        pass
                else:
                    print(f"❌ Пропускаем {iin}")
                    errors.append(iin)

        time.sleep(0.5)

    # ============================================================
    #  СОХРАНЯЕМ ИТОГОВЫЙ EXCEL СО ВСЕМИ КОЛОНКАМИ ИЗ ЗАПРОСА
    # ============================================================
    print("\n💾 Сохраняем итоговый Excel...")

    output_columns = [
        "Уникальный номер",
        "Продукт",
        "Номер договора",
        "ФИО",
        "ИИН",
        "Статус кредита",
        "Остаток задолженности",
        "ЧСИ",
        "Номер исполнительного производства",
        "Статус АИС ОИП",
        "Извещение о временном ограничении на выезд",
        "Постановление об ограничении на выезд",
        "Снятие ограничения на выезд",
    ]

    # Оставляем только колонки, которые реально есть в df
    existing_cols = [c for c in output_columns if c in df.columns]
    df_out = df[existing_cols].copy()

    with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False, sheet_name="Результат")

        ws = writer.sheets["Результат"]

        # Форматируем ИИН как текст с ведущими нулями
        if "ИИН" in existing_cols:
            iin_col_idx = existing_cols.index("ИИН") + 1  # openpyxl — 1-based
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=iin_col_idx)
                if cell.value:
                    cell.value = str(cell.value).strip().zfill(12)
                    cell.number_format = "@"

        # Автоширина колонок
        for col_cells in ws.columns:
            max_len = max(
                (len(str(c.value)) if c.value is not None else 0) for c in col_cells
            )
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 50)

    print(f"✅ Файл сохранён: {OUTPUT_EXCEL}")

    # ============================================================
    #  ИТОГИ
    # ============================================================
    total_time = time.time() - start_time
    avg_time   = total_time / processed if processed > 0 else 0

    print(f"\n{'='*60}")
    print(f"📊 ИТОГИ")
    print(f"   Обработано:          {processed}/{len(iin_list)}")
    print(f"   Успешно:             {successful}")
    print(f"   Ошибки:              {len(errors)}")
    print(f"   ⏱️ Общее время:      {total_time/60:.1f} мин")
    print(f"   ⚡ Среднее на ИИН:   {avg_time:.1f}с")
    if errors:
        print(f"   ❌ Список ошибок:   {errors[:10]}")
    print("=" * 60)

    try:
        driver.quit()
    except Exception:
        pass


if __name__ == "__main__":
    main()