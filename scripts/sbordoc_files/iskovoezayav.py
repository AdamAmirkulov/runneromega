# -*- coding: utf-8 -*-
"""
Блок 10: Формирование исковых заявлений из шаблона Word
"""
import os
import re
from pathlib import Path
from datetime import datetime, date

import pandas as pd
import pyodbc
from openpyxl import load_workbook
from docx import Document

from config import ROOT, TARGET_BASE, MAIN_EXCEL
from utils import ensure_client_folder, safe_log, safe_update_summary

# ═══════════════════════════════════════════════════════════════
# НАСТРОЙКИ
# ═══════════════════════════════════════════════════════════════

# Шаблон искового заявления
TEMPLATE_DOC_ISK = rf"{ROOT}\Документы для подачи Исков\Шаблоны документов\Шаблоны по Исковому заявлению в суд.docx"

# Адрес регистрации и Судебный орган по-прежнему считаются отдельным
# офлайн-пайплайном (poiskvsk.py: Selenium-адрес с Судебного кабинета +
# справочник "Суды по гражданским делам.xlsx") и синхронизируются им в
# лист "Данные для шаблонов" отчёта. В БД crm этих значений нет —
# берём их оттуда же, откуда брали раньше.
TEMPLATES_SHEET_NAME = "Данные для шаблонов"

# ═══════════════════════════════════════════════════════════════
# БД: ДАННЫЕ ДЛЯ ШАБЛОНОВ (напрямую из crm, без «Отчёта по отменам»)
# ═══════════════════════════════════════════════════════════════

DB_CONFIG = {
    "server":   "DBSRV",
    "database": "crm",
    "username": "sa",
    "password": "QazWsxEdc123!@#",
}

ISKOVOE_SQL_QUERY = r"""
-- база: дата сегодня минус 7 дней (без времени)
DECLARE @base datetime2 = DATEADD(day, -7, CAST(GETDATE() AS date));

WITH LoanCounts AS (
    SELECT
        c.F293 AS ИИН,
        COUNT(*) AS LoanCount
    FROM loans l
    JOIN clients c ON l.CID = c.ID
    GROUP BY c.F293
)
SELECT
    l.EID AS [Уникальный номер],
    dF246.F249 AS [Продукт],
    (
        SELECT STRING_AGG(
            CASE
                WHEN LEN(value) > 1
                    THEN UPPER(LEFT(value, 1)) + LOWER(SUBSTRING(value, 2, LEN(value) - 1))
                ELSE UPPER(value)
            END, ' '
        )
        FROM STRING_SPLIT(c.FIO, ' ')
    ) AS [ФИО],
    c.F293 AS [ИИН],
    CAST(l.F34 AS DECIMAL(18,2)) AS [Остаток задолженности],

    -- один email по клиенту
    MAX(e.Email) AS [Email],

    -- адрес регистрации
    MAX(
        LTRIM(RTRIM(
            CONCAT(
                -- область/регион
                NULLIF(LTRIM(RTRIM(a.Region)), ''),

                -- район (без лишней запятой, если Region пустой)
                CASE
                    WHEN NULLIF(LTRIM(RTRIM(a.Distric)), '') IS NOT NULL THEN
                        CASE
                            WHEN NULLIF(LTRIM(RTRIM(a.Region)), '') IS NOT NULL
                                THEN N', ' + LTRIM(RTRIM(a.Distric))
                            ELSE LTRIM(RTRIM(a.Distric))
                        END
                    ELSE N''
                END,

                -- город / населённый пункт
                CASE
                    WHEN NULLIF(LTRIM(RTRIM(a.City)), '') IS NOT NULL
                        THEN N', ' + LTRIM(RTRIM(a.City))
                    ELSE N''
                END,

                -- улица
                CASE
                    WHEN NULLIF(LTRIM(RTRIM(a.Street)), '') IS NOT NULL
                        THEN N', ' + LTRIM(RTRIM(a.Street))
                    ELSE N''
                END,

                -- дом (игнорим пустые и '-')
                CASE
                    WHEN NULLIF(LTRIM(RTRIM(a.House)), '') IS NOT NULL
                         AND NULLIF(LTRIM(RTRIM(a.House)), '-') IS NOT NULL
                        THEN N', д. ' + LTRIM(RTRIM(a.House))
                    ELSE N''
                END,

                -- квартира (игнорим пустые и '-')
                CASE
                    WHEN NULLIF(LTRIM(RTRIM(a.Flat)), '') IS NOT NULL
                         AND NULLIF(LTRIM(RTRIM(a.Flat)), '-') IS NOT NULL
                        THEN N', кв. ' + LTRIM(RTRIM(a.Flat))
                    ELSE N''
                END
            )
        ))
    ) AS [Адрес регистрации],

    -- случайная дата/время: 7 дней назад, в интервале 09:00:00..18:59:59 (как строка)
    CONVERT(varchar(19),
        DATEADD(
            SECOND,
            32400 + CAST(RAND(CHECKSUM(NEWID())) * 36000 AS int),  -- 32400 = 9*3600, 36000 секунд диапазон
            @base
        ),
        120
    ) AS [Дата минус 7 дней],

    -- мобильный: берём один номер по клиенту
    MAX(ph.PhoneNumber) AS [Мобильный],

    constF234.Caption AS [Кредитор],
    l.F287 AS [Номер договора займа],
    l.F15  AS [Дата выдачи займа],
    l.F200 AS [Сумма выдачи займа],
    l.F69  AS [Номер договора цессии],
    l.F68  AS [Дата договора цессии],
    CAST(
        ( ISNULL(l.F34, 0)
        - ISNULL(l.F315, 0)
        - ISNULL(l.F307, 0)
        - ISNULL(l.F38, 0)
        - ISNULL(l.F8, 0)
        ) AS DECIMAL(18,2)
    ) AS [Основной долг],
    l.F8  AS [Вознаграждение],
    l.F38 AS [Штраф],
    CAST(ISNULL(l.F315, 0) AS DECIMAL(18,2)) AS [Исполнительная надпись (расходы)],
    CAST(ISNULL(l.F307, 0) AS DECIMAL(18,2)) AS [Госпошлина (расходы)],
    l.F43 AS [Срок займа],
    DATEADD(DAY, l.F43, l.F15) AS [Дата окончания займа],
    null AS [Судебный орган],
    DATEADD(DAY, 5, l.F68) AS [Дата договора цессии + 5 дней],
    (l.F32 + l.F8 + l.F38) AS [ОД+проценты+штрафы],
    CAST(
        (ISNULL(l.F34,0) - ISNULL(l.F315,0) - ISNULL(l.F307,0))
        AS DECIMAL(18,2)
        ) AS [Остаток без расходов за надпись и без госпошлины]

FROM loans l (NOLOCK)
JOIN clients   c   (NOLOCK) ON l.CID = c.ID
JOIN states    s             ON s.ID = l.State
JOIN Dictionary dF246        ON dF246.ID = l.F246
JOIN Constants constF234     ON constF234.ID = l.F234
LEFT JOIN Constants constF235 ON constF235.ID = l.F235
JOIN      Constants constF236 ON constF236.ID = l.F236
LEFT JOIN LoanCounts lc        ON c.F293 = lc.ИИН
LEFT JOIN addresses a (NOLOCK)
       ON a.CID = l.CID
      AND a.FT = 153           -- только нужный тип адреса
LEFT JOIN Emails e (NOLOCK)    -- таблица с email
       ON e.CID = l.CID        -- или e.ClientID = c.ID, если так в схеме

-- берём один телефон клиента: сначала предпочитаем State=165 ("Актуальный"),
-- если такого нет — берём самый последний телефон независимо от статуса
OUTER APPLY (
    SELECT TOP 1 p.PhoneNumber
    FROM phones p
    WHERE p.CID = c.ID
    ORDER BY
        CASE WHEN p.State = 165 THEN 0 ELSE 1 END,
        p.ID DESC
) ph

WHERE
    RIGHT('000000000000' + LTRIM(RTRIM(CAST(c.F293 AS varchar(20)))), 12) IN ({iin_placeholders})

GROUP BY
    l.EID,
    dF246.F249,
    c.FIO,
    c.F293,
    l.F34,
    constF234.Caption,
    constF235.Caption,
    constF236.Caption,
    s.Caption,
    lc.LoanCount,
    l.F315,
    l.F307,
    l.F287,
    l.F15,
    l.F200,
    l.F69,
    l.F68,
    l.F32,
    l.F8,
    l.F38,
    l.F60,
    l.F13,
    l.F43,
    l.F157;
"""


def _norm_iin(value) -> str:
    s = re.sub(r"\D", "", str(value) if value is not None else "")
    return s.zfill(12) if s else ""


def load_data_from_db(iins: list) -> dict:
    """
    Загружает данные для шаблонов искового заявления напрямую из БД crm
    (вместо листа 'Данные для шаблонов' в «Отчёте по отменам»),
    только по ИИН, переданным из df_main (файла клиентов).
    Возвращает словарь ИИН (12 цифр) -> строка данных.
    """
    if not iins:
        return {}

    conn_str = (
        "DRIVER={ODBC Driver 18 for SQL Server};"
        f"SERVER={DB_CONFIG['server']};"
        f"DATABASE={DB_CONFIG['database']};"
        f"UID={DB_CONFIG['username']};"
        f"PWD={DB_CONFIG['password']};"
        "TrustServerCertificate=yes;"
        "Encrypt=no;"
    )

    query = ISKOVOE_SQL_QUERY.format(
        iin_placeholders=", ".join(["?"] * len(iins))
    )
    # Без бизнес-фильтров (статус, отмена, "уже отправлено" и т.д.) —
    # просто вытягиваем данные по переданным ИИН как есть.
    params = list(iins)

    conn = pyodbc.connect(conn_str, timeout=30)
    try:
        df = pd.read_sql(query, conn, params=params)
    finally:
        conn.close()

    db_data = {}
    for row in df.to_dict("records"):
        iin = _norm_iin(row.get("ИИН"))
        if not iin:
            continue
        # УГД / БИН УГД этим запросом не поставляются — оставляем пустыми,
        # чтобы плейсхолдеры {УГД}/{БИН УГД} в шаблоне гарантированно заменялись.
        row.setdefault("УГД", "")
        row.setdefault("БИН УГД", "")
        db_data[iin] = row

    return db_data


def load_address_and_court_from_excel() -> dict:
    """
    Адрес регистрации и Судебный орган по-прежнему берутся так же, как
    и раньше — из листа "Данные для шаблонов" в «Отчёте по отменам»
    (эти два значения считает отдельный офлайн-пайплайн poiskvsk.py,
    а не БД). Остальные поля из этого файла больше не используются.

    Возвращает словарь ИИН (12 цифр) -> {"Адрес регистрации": ..., "Судебный орган": ...}.
    Если файл/лист недоступны — возвращает {} (не должно останавливать блок,
    т.к. основные данные уже пришли из БД).
    """
    if not os.path.exists(MAIN_EXCEL):
        print(f"⚠️  Файл отчёта не найден (адрес/суд не будут заполнены): {MAIN_EXCEL}")
        safe_log(f"[ИСКОВОЕ] Файл отчёта не найден для адреса/суда: {MAIN_EXCEL}")
        return {}

    try:
        wb = load_workbook(MAIN_EXCEL, read_only=True, data_only=True)
    except Exception as e:
        print(f"⚠️  Не удалось открыть отчёт для адреса/суда: {e}")
        safe_log(f"[ИСКОВОЕ] Ошибка открытия отчёта для адреса/суда: {e}")
        return {}

    if TEMPLATES_SHEET_NAME not in wb.sheetnames:
        print(f"⚠️  В отчёте нет листа '{TEMPLATES_SHEET_NAME}' — адрес/суд не будут заполнены")
        wb.close()
        return {}

    ws = wb[TEMPLATES_SHEET_NAME]
    headers = [str(c.value).strip() if c.value else "" for c in ws[1]]

    try:
        iin_idx = headers.index("ИИН")
        addr_idx = headers.index("Адрес регистрации")
        court_idx = headers.index("Судебный орган")
    except ValueError as e:
        print(f"⚠️  В листе '{TEMPLATES_SHEET_NAME}' нет нужной колонки: {e}")
        wb.close()
        return {}

    extra = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) <= max(iin_idx, addr_idx, court_idx):
            continue
        iin = _norm_iin(row[iin_idx])
        if not iin:
            continue
        extra[iin] = {
            "Адрес регистрации": row[addr_idx],
            "Судебный орган": row[court_idx],
        }

    wb.close()
    return extra

# ═══════════════════════════════════════════════════════════════
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ══════════════════════════════════════════════════════
def nice_case(s: str) -> str:
    """Приводим ФИО к виду 'Фамилия Имя Отчество'."""
    if not s:
        return ""
    return " ".join(word.capitalize() for word in str(s).split())


def replace_placeholders_in_paragraph(paragraph, mapping: dict):
    """
    Заменяет плейсхолдеры в абзаце, стараясь сохранить форматирование.
    Работает и когда плейсхолдер разбит на несколько run'ов.
    """
    # 1) Простой случай — плейсхолдер целиком внутри одного run
    for run in paragraph.runs:
        text = run.text
        changed = False
        for key, value in mapping.items():
            if key in text:
                text = text.replace(key, value)
                changed = True
        if changed:
            run.text = text

    # 2) Случай, когда плейсхолдер растянут на несколько run'ов
    runs = paragraph.runs
    i = 0
    while i < len(runs):
        replaced_here = False

        for key, value in mapping.items():
            L = len(key)
            accum = ""
            j = i

            # Копим текст из нескольких run'ов
            while j < len(runs) and len(accum) < L + 10 and "}" not in accum:
                accum += runs[j].text
                j += 1

            idx = accum.find(key)
            if idx == 0:
                # Плейсхолдер начинается сразу
                tail = accum[L:]
                new_text = value + tail

                runs[i].text = new_text
                # Очищаем остальные run'ы
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

    # Текущая дата
    mapping["{CurrDate}"] = datetime.now().strftime("%d.%m.%Y")

    # Абзацы
    for p in doc.paragraphs:
        replace_placeholders_in_paragraph(p, mapping)

    # Таблицы
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_placeholders_in_paragraph(p, mapping)

    return doc


# ═══════════════════════════════════════════════════════════════
# ОСНОВНАЯ ФУНКЦИЯ
# ═══════════════════════════════════════════════════════════════

def run(df_main):
    """
    Формирует исковые заявления для каждого клиента из Excel-шаблона.
    
    Args:
        df_main: DataFrame с колонками ['Product', 'FIO', 'IIN']
    
    Returns:
        tuple: (count_success, count_failed)
    """
    print("\n" + "="*70)
    print("  ⚖️  БЛОК 10: ИСКОВОЕ ЗАЯВЛЕНИЕ")
    print("="*70 + "\n")

    count_success = 0
    count_failed = 0
    not_found_list = []

    if not os.path.exists(TEMPLATE_DOC_ISK):
        print(f"❌ ОШИБКА: Шаблон не найден: {TEMPLATE_DOC_ISK}")
        safe_log(f"[ИСКОВОЕ] Шаблон не найден: {TEMPLATE_DOC_ISK}")
        return 0, len(df_main)

    # Данные для шаблонов — напрямую из БД crm (вместо «Отчёта по отменам»),
    # только по ИИН клиентов, пришедших из файла (df_main)
    iins_needed = sorted({
        _norm_iin(iin) for iin in df_main['IIN'] if _norm_iin(iin)
    })

    try:
        excel_data = load_data_from_db(iins_needed)
    except Exception as e:
        print(f"❌ ОШИБКА при загрузке данных из БД: {e}")
        safe_log(f"[ИСКОВОЕ] Ошибка загрузки данных из БД: {e}")
        return 0, len(df_main)

    print(f"📊 Загружено данных из БД: {len(excel_data)} строк")

    # Адрес регистрации и Судебный орган — как и раньше, из отчёта
    # (poiskvsk.py), а не из БД
    addr_court_extra = load_address_and_court_from_excel()
    filled_addr_court = 0
    for iin, row_dict in excel_data.items():
        extra = addr_court_extra.get(iin)
        if not extra:
            continue
        if extra.get("Адрес регистрации"):
            row_dict["Адрес регистрации"] = extra["Адрес регистрации"]
        if extra.get("Судебный орган"):
            row_dict["Судебный орган"] = extra["Судебный орган"]
        filled_addr_court += 1
    print(f"📍 Адрес/суд из отчёта подставлены для: {filled_addr_court}/{len(excel_data)}")

    print(f"📋 Клиентов для обработки: {len(df_main)}\n")

    # Обрабатываем каждого клиента из df_main
    total = len(df_main)
    
    for idx, row in df_main.iterrows():
        fio = str(row['FIO']).strip()
        iin = str(row['IIN']).strip()
        
        print(f"[{idx+1}/{total}] {fio} ({iin})")

        # Ищем данные клиента в результатах запроса к БД
        if iin not in excel_data:
            print(f"   ⚠️  Данные не найдены в БД")
            count_failed += 1
            not_found_list.append(f"{fio}, {iin}")
            safe_log(f"[ИСКОВОЕ] Данные не найдены для: {fio}, {iin}")
            continue

        try:
            row_dict = excel_data[iin]
            
            # Нормализуем ФИО и ИИН
            for key in row_dict.keys():
                if "фио" in key.lower():
                    row_dict[key] = nice_case(fio)
                if "иин" in key.lower():
                    row_dict[key] = iin

            # Создаём документ
            doc = fill_document_isk(row_dict)

            # Получаем папку клиента
            target_folder = ensure_client_folder(iin, fio, TARGET_BASE)

            # Формируем имя файла
            base_name = f"Исковое заявление, {fio}, {iin}"
            docx_path = Path(target_folder) / (base_name + ".docx")

            # Сохраняем Word
            doc.save(str(docx_path))

            count_success += 1
            print(f"   ✅ Создан: {docx_path.name}")

        except Exception as e:
            count_failed += 1
            not_found_list.append(f"{fio}, {iin}")
            print(f"   ❌ Ошибка: {e}")
            safe_log(f"[ИСКОВОЕ] Ошибка для {fio} ({iin}): {e}")

    # Итоговая статистика
    print("\n" + "-"*70)
    print(f"📊 ИТОГО:")
    print(f"   ✅ Успешно создано: {count_success}/{total}")
    print(f"   ❌ Ошибок: {count_failed}/{total}")
    print("-"*70 + "\n")

    # Сохраняем в LOG_SUMMARY
    safe_update_summary("ИСКОВОЕ ЗАЯВЛЕНИЕ", {
        "found": count_success,
        "total": total,
        "not_found": not_found_list
    })

    return count_success, count_failed