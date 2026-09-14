# mailing_common.py
"""
Общие утилиты для механизма «Рассылки по отчётам» (ЧСИ).

Используются и веб-приложением (app.py — при сохранении справочника ЧСИ),
и скриптом-исполнителем (scripts/mailing_send.py).
"""
import re

# Фиксированный набор автоподстановок в теме/теле письма.
# Ключ плейсхолдера -> человекочитаемое пояснение (показывается в форме админки).
PLACEHOLDERS = {
    "{{ФИО_ЧСИ}}":        "ФИО частного судебного исполнителя",
    "{{кол_во_записей}}": "сколько строк отчёта относится к этому ЧСИ",
    "{{компания}}":       "короткое название компании-отправителя",
    "{{компания_ru}}":    "полное название компании на русском (config.MAILING_SIGNATURE)",
    "{{компания_kz}}":    "полное название компании на казахском (config.MAILING_SIGNATURE)",
    "{{телефон}}":        "контактный телефон компании (config.MAILING_SIGNATURE)",
    "{{дата}}":           "текущая дата (ДД.ММ.ГГГГ)",
}


def normalize_fio(s) -> str:
    """Приводит ФИО к каноничному виду для сопоставления результата SQL со
    справочником: нижний регистр, ё→е, пунктуация → пробел, лишние пробелы убраны.
    Портировано 1:1 из ноутбука «Рассылка ORION пункт 7» (ячейка 1)."""
    s = str(s or "").strip().lower().replace("ё", "е")
    s = re.sub(r"[.,;:()\"'`]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def safe_filename(name: str) -> str:
    """Безопасное имя файла-вложения из ФИО ЧСИ."""
    name = re.sub(r'[\\/:*?"<>|]+', "_", str(name or ""))
    name = re.sub(r"\s+", " ", name).strip()
    return (name[:150] if len(name) > 150 else name) or "ЧСИ"


def render_placeholders(text: str, ctx: dict) -> str:
    """Заменяет плейсхолдеры из PLACEHOLDERS на значения из ctx.
    ctx: {"ФИО_ЧСИ": ..., "кол_во_записей": ..., "компания": ..., "дата": ...}"""
    if not text:
        return text or ""
    out = text
    for key in PLACEHOLDERS:
        var = key[2:-2]  # {{ФИО_ЧСИ}} -> ФИО_ЧСИ
        out = out.replace(key, str(ctx.get(var, "")))
    return out


_FORBIDDEN_SQL = re.compile(
    r"\b(INSERT|UPDATE|DELETE|DROP|ALTER|TRUNCATE|EXEC|EXECUTE|MERGE|GRANT|REVOKE|"
    r"CREATE|BACKUP|RESTORE|INTO|OPENROWSET|OPENQUERY|OPENDATASOURCE)\b",
    re.IGNORECASE,
)
_FORBIDDEN_PROC = re.compile(r"\b(sp_|xp_)\w+", re.IGNORECASE)
_COMMENT_LINE = re.compile(r"--[^\n]*")
_COMMENT_BLOCK = re.compile(r"/\*.*?\*/", re.DOTALL)


def check_select_only(sql: str) -> None:
    """Грубая защита: разрешаем только один SELECT/WITH-запрос.
    Бросает ValueError с понятным текстом, если запрос выглядит опасным."""
    if not sql or not sql.strip():
        raise ValueError("Пустой SQL-запрос.")

    stripped = _COMMENT_BLOCK.sub(" ", sql)
    stripped = _COMMENT_LINE.sub(" ", stripped)
    stripped = stripped.strip()

    low = stripped.lstrip("(").lstrip().lower()
    if not (low.startswith("select") or low.startswith("with")):
        raise ValueError("SQL должен начинаться с SELECT или WITH.")

    # Точка с запятой допустима только в самом конце (один запрос).
    if ";" in stripped.rstrip().rstrip(";"):
        raise ValueError("Разрешён только один SQL-запрос (лишняя «;»).")

    m = _FORBIDDEN_SQL.search(stripped) or _FORBIDDEN_PROC.search(stripped)
    if m:
        raise ValueError(
            f"В запросе запрещённая конструкция: {m.group(0).upper()}. "
            "Разрешены только SELECT-запросы (чтение)."
        )
