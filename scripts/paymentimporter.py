import os
import re
import sys
import argparse
from copy import copy
from collections import defaultdict

import pandas as pd
import pyodbc
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font


# ========= АРГУМЕНТЫ КОМАНДНОЙ СТРОКИ =========
parser = argparse.ArgumentParser()
parser.add_argument("--excel_file",    default=None, help="Путь к файлу выписки (загруженный через веб)")
parser.add_argument("--output_folder", default=None, help="Папка для результата (опционально)")
parser.add_argument("--workdir",       default=None, help="Рабочая директория задачи (передаётся runner'ом)")
parser.add_argument("--company_id", default=None, type=int, help="ID компании")
args, _ = parser.parse_known_args()

if not args.company_id:
    print("❌ ОШИБКА: не передан --company_id — без него get_eid_map() ищет EID БЕЗ фильтра "
          "по компании (l.F209) и подтягивает сделки других компаний по тому же ИИН. Запуск остановлен.")
    sys.exit(1)


# ========= НАСТРОЙКИ =========
DESKTOP_DIR = r"C:\Users\User\Desktop"

DB_SERVER   = "DBSRV"
DB_DATABASE = "crm"
DB_USERNAME = "user"
DB_PASSWORD = "Log1cF"

# Папка для результата: если передана --output_folder — берём её,
# иначе --workdir/out (папка задачи в runner'е), иначе рабочий стол
if args.output_folder:
    OUT_DIR = args.output_folder
elif args.workdir:
    OUT_DIR = os.path.join(args.workdir, "out")
else:
    OUT_DIR = DESKTOP_DIR

os.makedirs(OUT_DIR, exist_ok=True)
OUTPUT_FILE = os.path.join(OUT_DIR, "PaymentsImport_draft.xlsx")


# ========= ВСПОМОГАТЕЛЬНЫЕ =========
def normalize(x):
    return "" if pd.isna(x) else str(x).strip().replace("\n", " ").replace("\r", " ").replace("\xa0", " ")


def find_source_file():
    # Если файл передан через веб-интерфейс — используем его
    if args.excel_file:
        if not os.path.exists(args.excel_file):
            raise FileNotFoundError(f"Переданный файл не найден: {args.excel_file}")
        return args.excel_file

    # Иначе — старое поведение: ищем на рабочем столе
    candidates = []
    for f in os.listdir(DESKTOP_DIR):
        low = f.lower()
        if "выписка по счету" in low and (low.endswith(".xls") or low.endswith(".xlsx")):
            candidates.append(os.path.join(DESKTOP_DIR, f))

    if not candidates:
        raise FileNotFoundError("На рабочем столе не найден файл, содержащий 'Выписка по счету'.")

    candidates.sort()
    return candidates[0]

def get_company_name(company_id):
    import sqlite3
    # SQLite база лежит рядом со скриптом (в корне проекта)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    db_path = os.path.join(script_dir, "..", "users.db")  # укажи правильное имя файла
    
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("SELECT name FROM companies WHERE id = ?", [company_id])
    row = cur.fetchone()
    cur.close()
    conn.close()
    
    if not row:
        raise ValueError(f"Компания с ID={company_id} не найдена в БД")
    return row[0]

def find_template_file():
    # Шаблон всегда ищем рядом со скриптом или на рабочем столе
    script_dir = os.path.dirname(os.path.abspath(__file__))

    for search_dir in [script_dir, DESKTOP_DIR]:
        candidates = []
        for f in os.listdir(search_dir):
            low = f.lower()
            if low.endswith(".xlsx") and low.startswith("paymentsimport_draft"):
                full_path = os.path.join(search_dir, f)
                # исключаем итоговый файл
                if os.path.normcase(full_path) != os.path.normcase(OUTPUT_FILE):
                    candidates.append(full_path)
        if candidates:
            candidates.sort()
            return candidates[0]

    raise FileNotFoundError(
        "Не найден шаблон вида 'PaymentsImport_draft*.xlsx'.\n"
        "Положи шаблон рядом со скриптом или на рабочий стол, например:\n"
        "PaymentsImport_draft(1).xlsx"
    )


def extract_iin(text):
    if pd.isna(text):
        return "ошибка"

    found = re.findall(r'(?<!\d)\d{12}(?!\d)', str(text))

    uniq = []
    seen = set()
    for x in found:
        if x not in seen:
            uniq.append(x)
            seen.add(x)

    return ", ".join(uniq) if uniq else "ошибка"


def is_chsi(text):
    """
    True, если в колонке 'Контрагент' указан частный судебный исполнитель
    (ЧСИ), а не просто ФИО физлица. "ЧСИ" у разных банков попадается не
    только в начале строки (например, "Бугунаев (ЧСИ)"), поэтому ищем
    по всей строке, а не только как префикс.
    """
    if pd.isna(text):
        return False
    s = str(text).strip()
    return bool(re.search(r'частный\s+судебный\s+исполнитель|\bчси\b', s, flags=re.I))


def extract_iin_from_bin(value):
    """
    Достаёт ИИН из колонки 'БИН контрагента'.
    Используется для платежей СУБП: Homebank, где в тексте назначения
    платежа ИИН может отсутствовать/дублироваться ненадёжно, а в БИН
    контрагента приходит напрямую из банка (иногда как float, из-за
    чего теряется ведущий ноль — это тут компенсируется).
    """
    if pd.isna(value):
        return "ошибка"

    s = str(value).strip()
    if s.endswith(".0"):
        s = s[:-2]

    s = re.sub(r"\D", "", s)
    if not s:
        return "ошибка"

    s = s.zfill(12)
    if len(s) != 12:
        return "ошибка"

    return s


def clean_fio(text):
    if pd.isna(text):
        return ""

    s = str(text).strip()
    # "ЧСИ"/"частный судебный исполнитель" может стоять не только в начале
    # строки (например, "Бугунаев (ЧСИ)"), поэтому вырезаем по всей строке.
    s = re.sub(r'частный\s+судебный\s+исполнитель', '', s, flags=re.I)
    s = re.sub(r'\bчси\b', '', s, flags=re.I)
    # Убираем пустые скобки, оставшиеся после вырезания ("(ЧСИ)" -> "()").
    s = re.sub(r'\(\s*\)', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s.upper()


def is_trash_row(row):
    txt = " ".join([normalize(x).lower() for x in row.tolist()])
    bad = [
        "исходящий остаток",
        "входящий остаток",
        "итого обороты",
        "итого ндс"
    ]
    return any(x in txt for x in bad)


def get_conn():
    conn_str = (
        "DRIVER={ODBC Driver 18 for SQL Server};"
        f"SERVER={DB_SERVER};"
        f"DATABASE={DB_DATABASE};"
        f"UID={DB_USERNAME};"
        f"PWD={DB_PASSWORD};"
        "TrustServerCertificate=yes;"
    )
    return pyodbc.connect(conn_str)


def get_eid_map(iins, company_name=None):
    clean_iins = []
    for x in iins:
        x = str(x).strip()
        if re.fullmatch(r"\d{12}", x):
            clean_iins.append(x)

    clean_iins = list(dict.fromkeys(clean_iins))

    if not clean_iins:
        return {}

    conn = get_conn()
    cur = conn.cursor()
    result = defaultdict(list)

    chunk_size = 500
    for i in range(0, len(clean_iins), chunk_size):
        chunk = clean_iins[i:i + chunk_size]
        placeholders = ",".join("?" for _ in chunk)
        company_filter = "AND l.F209 LIKE ?" if company_name else ""
        
        sql = f"""
        SELECT c.F293 AS IIN, l.EID AS EID
        FROM Loans l
        JOIN Clients c ON l.CID = c.ID
        WHERE c.F293 IN ({placeholders})
        AND c.F293 IS NOT NULL
        AND l.EID IS NOT NULL
        {company_filter}
        ORDER BY c.F293, l.EID
        """
        
        params = chunk + ([f"%{company_name}%"] if company_name else [])
        cur.execute(sql, params)
        

        for row in cur.fetchall():
            iin = str(row.IIN).strip()
            eid = str(row.EID).strip()

            if eid not in result[iin]:
                result[iin].append(eid)

    cur.close()
    conn.close()

    return {k: ", ".join(v) for k, v in result.items()}


def find_header_row(raw_df):
    for i in range(len(raw_df)):
        row_text = " ".join([normalize(x).lower() for x in raw_df.iloc[i].tolist()])

        if (
            "дата валютирования" in row_text and
            "контрагент" in row_text and
            "кредит" in row_text and
            "назначение платежа" in row_text
        ):
            return i

    raise ValueError(
        "Не удалось найти строку заголовков. "
        "В файле должны быть колонки: Дата валютирования, Контрагент, Кредит, Назначение платежа."
    )


def read_statement_file(source_file):
    if source_file.lower().endswith(".xls"):
        raw_df = pd.read_excel(source_file, header=None, engine="xlrd")
        header_row = find_header_row(raw_df)
        df = pd.read_excel(source_file, header=header_row, engine="xlrd")
    else:
        raw_df = pd.read_excel(source_file, header=None)
        header_row = find_header_row(raw_df)
        df = pd.read_excel(source_file, header=header_row)

    df.columns = [normalize(c) for c in df.columns]
    return df, header_row


def parse_amount(series):
    return pd.to_numeric(
        series.astype(str)
        .str.replace("\xa0", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False),
        errors="coerce"
    )


def copy_row_style(ws, source_row, target_row, start_col=1, end_col=7):
    for col in range(start_col, end_col + 1):
        src = ws.cell(source_row, col)
        dst = ws.cell(target_row, col)

        if src.has_style:
            dst._style = copy(src._style)

        if src.number_format:
            dst.number_format = src.number_format

        if src.font:
            dst.font = copy(src.font)

        if src.fill:
            dst.fill = copy(src.fill)

        if src.border:
            dst.border = copy(src.border)

        if src.alignment:
            dst.alignment = copy(src.alignment)

        if src.protection:
            dst.protection = copy(src.protection)


# ========= ОСНОВНАЯ ЛОГИКА =========
SOURCE_FILE   = find_source_file()
TEMPLATE_FILE = find_template_file()
# После find_source_file() / find_template_file()
company_name = get_company_name(args.company_id) if args.company_id else None
print(f"Компания: {company_name if company_name else 'не указана'}")

print("Файл выписки:", SOURCE_FILE)
print("Шаблон:", TEMPLATE_FILE)
print("Выходной файл:", OUTPUT_FILE)

df, header_row = read_statement_file(SOURCE_FILE)
print(f"Строка заголовков: {header_row + 1}")
print("Колонки:", df.columns.tolist())

required_cols = [
    "Кредит",
    "Дата валютирования",
    "Контрагент",
    "Назначение платежа"
]
missing = [c for c in required_cols if c not in df.columns]
if missing:
    raise ValueError(f"Не найдены колонки: {', '.join(missing)}")

has_bin_col = "БИН контрагента" in df.columns
if not has_bin_col:
    print("ВНИМАНИЕ: колонка 'БИН контрагента' не найдена — правило для Homebank применяться не будет.")

# Убираем мусорные/итоговые строки
df = df[~df.apply(is_trash_row, axis=1)].copy()

# Убираем полностью пустые строки по ключевым полям
df = df[
    ~(
        df["Кредит"].isna() &
        df["Дата валютирования"].isna() &
        df["Контрагент"].isna() &
        df["Назначение платежа"].isna()
    )
].copy()

# Число
df["Кредит_num"] = parse_amount(df["Кредит"])

# Оставляем строки, где сумма есть
df = df[df["Кредит_num"].notna()].copy()

# Формируем результат
result = pd.DataFrame()
result["Сумма погашения"]    = df["Кредит_num"]
result["Дата погашения"]     = pd.to_datetime(df["Дата валютирования"], errors="coerce", dayfirst=True)
result["Источник погашения"] = df["Контрагент"].apply(clean_fio)
result["ИИН"]                = df["Назначение платежа"].apply(extract_iin)
result["Назначение платежа"] = df["Назначение платежа"].fillna("").astype(str)

# Для платежей СУБП: Homebank, где в "Контрагент" указано просто ФИО
# физлица (НЕ ЧСИ/частный судебный исполнитель) — если в тексте
# назначения платежа НЕ нашлось валидного 12-значного ИИН, берём его
# из колонки "БИН контрагента". Если в тексте ИИН уже есть — не трогаем.
# Для ЧСИ БИН контрагента — это реквизиты самого судебного исполнителя,
# а не должника, поэтому в этом случае продолжаем брать ИИН только из текста.
if has_bin_col:
    homebank_mask     = df["Назначение платежа"].astype(str).str.contains("Homebank", case=False, na=False)
    plain_fio_mask    = ~df["Контрагент"].apply(is_chsi)
    text_failed_mask  = (result["ИИН"] == "ошибка")
    use_bin_mask      = homebank_mask & plain_fio_mask & text_failed_mask

    if use_bin_mask.any():
        result.loc[use_bin_mask, "ИИН"] = df.loc[use_bin_mask, "БИН контрагента"].apply(extract_iin_from_bin)
        print(f"Homebank-платежей (ФИО, не ЧСИ, ИИН в тексте не найден) обработано по БИН контрагента: {int(use_bin_mask.sum())}")

    skipped_chsi_mask = homebank_mask & ~plain_fio_mask & text_failed_mask
    if skipped_chsi_mask.any():
        print(f"Homebank-платежей от ЧСИ без ИИН в тексте — ИИН НЕ взят из БИН контрагента (осталась 'ошибка'): {int(skipped_chsi_mask.sum())}")

# Получаем список ИИН
all_iins = []
for x in result["ИИН"]:
    if x != "ошибка":
        all_iins.extend([v.strip() for v in x.split(",") if re.fullmatch(r"\d{12}", v.strip())])

eid_map = get_eid_map(all_iins, company_name=company_name)


def map_eid(iin_text):
    if iin_text == "ошибка":
        return "ошибка"

    res = []
    for i in [x.strip() for x in iin_text.split(",") if x.strip()]:
        if i in eid_map:
            for eid in eid_map[i].split(","):
                eid = eid.strip()
                if eid and eid not in res:
                    res.append(eid)

    return ", ".join(res) if res else "ошибка"


result["Уникальный номер сделки"]    = result["ИИН"].apply(map_eid)
result["Уникальный номер погашения"] = ""

# Проверка суммы
source_sum = round(df["Кредит_num"].sum(), 2)
result_sum = round(result["Сумма погашения"].sum(), 2)

# ========= ЗАПИСЬ В ШАБЛОН =========
wb = load_workbook(TEMPLATE_FILE)
ws = wb.active

start_row          = 2
template_style_row = 2

# Очищаем старые данные
for r in range(start_row, ws.max_row + 1):
    for c in range(1, 8):
        ws.cell(r, c).value = None

# Заготовки выравнивания
center_align    = Alignment(horizontal="center", vertical="center")
left_align      = Alignment(horizontal="left",   vertical="center")
left_wrap_align = Alignment(horizontal="left",   vertical="center", wrap_text=True)

# Красный шрифт для строк с ошибкой
red_font = Font(color="FF0000")

# Записываем новые данные
for i, row in result.iterrows():
    r = start_row + i

    copy_row_style(ws, template_style_row, r, 1, 7)

    ws.cell(r, 1).value = float(row["Сумма погашения"]) if pd.notna(row["Сумма погашения"]) else None
    ws.cell(r, 2).value = row["Дата погашения"].to_pydatetime() if pd.notna(row["Дата погашения"]) else None
    ws.cell(r, 3).value = row["Источник погашения"]
    ws.cell(r, 4).value = row["Уникальный номер сделки"]
    ws.cell(r, 5).value = row["Уникальный номер погашения"]
    ws.cell(r, 6).value = row["ИИН"]
    ws.cell(r, 7).value = row["Назначение платежа"]

    ws.cell(r, 1).number_format = '# ##0.00'
    ws.cell(r, 2).number_format = 'DD.MM.YYYY'

    ws.cell(r, 1).alignment = center_align
    ws.cell(r, 2).alignment = center_align
    ws.cell(r, 4).alignment = center_align
    ws.cell(r, 5).alignment = center_align
    ws.cell(r, 6).alignment = center_align
    ws.cell(r, 3).alignment = left_align
    ws.cell(r, 7).alignment = left_wrap_align

    if row["ИИН"] == "ошибка" or row["Уникальный номер сделки"] == "ошибка":
        for col in range(1, 8):
            ws.cell(r, col).font = copy(red_font)

ws.freeze_panes = "A2"
wb.save(OUTPUT_FILE)

print("\n=== ГОТОВО ===")
print("Создан файл:", OUTPUT_FILE)
print(f"Сумма в выписке:   {source_sum:,.2f}")
print(f"Сумма в результате: {result_sum:,.2f}")
print("Проверка суммы:", "OK" if source_sum == result_sum else "НЕ СОВПАДАЕТ")