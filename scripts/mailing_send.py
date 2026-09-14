# -*- coding: utf-8 -*-
"""
Рассылка по отчёту (ЧСИ).

Настраивается в веб-интерфейсе: Админка → «Рассылки по отчётам». Здесь —
исполнитель, который раннер запускает подпроцессом:

    python -u scripts/mailing_send.py --workdir <dir> --company_id <id[,id2,...]>
                                      --mailing_id <id> --mode <test|real>

--company_id может быть списком через запятую (мульти-компанийная рассылка,
mailing.company_id_list() в database.py) — тогда весь цикл (SQL с подстановкой
{company_filter}, свой SMTP-ящик, свой справочник ЧСИ) прогоняется по очереди
для каждой компании, письма собираются в один общий журнал с колонкой «Компания».

Алгоритм:
  1. Берём рассылку из users.db, её эффективный SQL (свой SQL приоритетнее
     выбранного отчёта), подставляем {company_filter} и проверяем, что это
     SELECT (только чтение).
  2. Выполняем запрос в CRM (MS SQL) → DataFrame.
  3. Группируем строки по колонке-ЧСИ (mailing.group_column).
  4. Для каждого ЧСИ ищем email в справочнике ЧСИ этой компании (по ФИО,
     нормализация + точное совпадение) и отправляем письмо с его строками
     в xlsx-вложении. Тема/тело — из рассылки, с автоподстановками.

Режимы:
  * real — реальная отправка каждому ЧСИ (+ BCC из рассылки и из config);
  * test — одно письмо на test_email по данным ПЕРВОЙ группы, тема «[ТЕСТ] …»,
    реальные адреса ЧСИ не затрагиваются.

Результат: <workdir>/out/Рассылка_<дата>.xlsx (журнал: ЧСИ, email, статус,
строк, ошибка). Прогресс печатается в stdout (→ logs.txt).

Код выхода: 1 — фатальная ошибка (нет SMTP для компании, плохой SQL, нет
колонки-ЧСИ, пустой результат, test без test_email). 0 — прошло (даже если
часть писем не ушла — смотри журнал).
"""
import argparse
import os
import sys
import time
import smtplib
from datetime import date, datetime
from decimal import Decimal
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

# Корень репозитория — чтобы импортировать database / mailing_common
_BASE = Path(__file__).resolve().parent.parent
if str(_BASE) not in sys.path:
    sys.path.insert(0, str(_BASE))


def _parse_args():
    p = argparse.ArgumentParser(description="Рассылка по отчёту (ЧСИ)")
    p.add_argument("--workdir", default=None)
    p.add_argument("--company_id", default=None)
    p.add_argument("--mailing_id", default=None)
    p.add_argument("--mode", default="test", choices=["test", "real"])
    p.add_argument("--test_email", default=None)
    return p.parse_args()


_args = _parse_args()


def die(msg: str):
    print(f"❌ {msg}", flush=True)
    sys.exit(1)


if not _args.company_id or not _args.company_id.strip():
    die("не передан --company_id")
if not _args.mailing_id or not _args.mailing_id.strip():
    die("не передан --mailing_id")

COMPANY_IDS = [x.strip() for x in _args.company_id.split(",") if x.strip()]
os.environ["COMPANY_ID"] = COMPANY_IDS[0]      # нужен config.py на этапе импорта
MODE = _args.mode

WORKDIR = Path(_args.workdir).resolve() if _args.workdir else Path.cwd()
OUT_DIR = WORKDIR / "out"
OUT_DIR.mkdir(parents=True, exist_ok=True)

# ── импорты, зависящие от окружения ───────────────────────────
import pandas as pd                                   # noqa: E402
import pyodbc                                          # noqa: E402
from openpyxl import Workbook                          # noqa: E402

import database                                        # noqa: E402
from mailing_common import (                           # noqa: E402
    normalize_fio, safe_filename, render_placeholders, check_select_only,
)

try:
    from config import MAILING_SMTP, COMPANY_DB_FILTER  # noqa: E402
except ImportError as e:
    die(f"не удалось импортировать config: {e}")

try:
    from config import MAILING_SIGNATURE  # noqa: E402
except ImportError:
    MAILING_SIGNATURE = {}

try:
    from config import CRM_DB as DB_CONFIG  # noqa: E402
except ImportError as e:
    die(f"не удалось импортировать config.CRM_DB: {e}")


def crm_connect():
    conn_str = (
        "DRIVER={ODBC Driver 18 for SQL Server};"
        f"SERVER={DB_CONFIG['server']};"
        f"DATABASE={DB_CONFIG['database']};"
        f"UID={DB_CONFIG['username']};"
        f"PWD={DB_CONFIG['password']};"
        "TrustServerCertificate=yes;Encrypt=no;"
    )
    return pyodbc.connect(conn_str, timeout=30)


def cell(v):
    """Приводит значение из БД к тому, что умеет писать openpyxl."""
    if v is None or isinstance(v, (str, int, float, bool, datetime, date)):
        return v
    if isinstance(v, Decimal):
        return float(v)
    return str(v)


def build_attachment(path: Path, columns, rows):
    wb = Workbook()
    ws = wb.active
    ws.append([str(c) for c in columns])
    for r in rows:
        ws.append([cell(x) for x in r])
    wb.save(path)


def send_one(smtp_cfg, to_email, bcc_list, subject, body_html, attach_path):
    msg = MIMEMultipart()
    msg["From"] = formataddr((smtp_cfg.get("from_name") or "", smtp_cfg["user"]))
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body_html or "", "html", "utf-8"))

    if attach_path:
        fname = Path(attach_path).name
        with open(attach_path, "rb") as f:
            part = MIMEBase(
                "application",
                "vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", "attachment", filename=("utf-8", "", fname))
        msg.attach(part)

    recipients = [to_email] + [b for b in bcc_list if b]

    host, port = smtp_cfg["host"], int(smtp_cfg.get("port", 587))
    with smtplib.SMTP(host, port, timeout=30) as server:
        if smtp_cfg.get("use_tls", True):
            server.starttls()
        server.login(smtp_cfg["user"], smtp_cfg["password"])
        server.sendmail(smtp_cfg["user"], recipients, msg.as_string())


def run_for_company(cid: str, mailing, base_sql: str):
    """Прогоняет весь цикл (SQL → группировка → отправка) для одной компании.
    Возвращает список строк журнала (fio, email, статус, строк, ошибка) с
    добавленным именем компании первым элементом; фатальные для ЭТОЙ компании
    проблемы (нет SMTP, ошибка SQL, пустой результат…) не останавливают
    остальные компании — печатаются как ошибка и company пропускается."""
    smtp_cfg = MAILING_SMTP.get(cid)
    if not smtp_cfg or any(not smtp_cfg.get(k) for k in ("host", "user", "password")):
        print(f"❌ [{cid}] в config.MAILING_SMTP нет полной записи для компании {cid} — пропуск", flush=True)
        return []

    _db = database.SessionLocal()
    _row = _db.query(database.Company).filter(database.Company.id == int(cid)).first()
    company_name = _row.name if _row else (smtp_cfg.get("from_name") or cid)
    _db.close()

    print(f"\n=== Рассылка «{mailing.name}» | компания {company_name} | режим {MODE.upper()} ===", flush=True)

    company_filter = COMPANY_DB_FILTER.get(cid)
    if not company_filter:
        print(f"❌ [{company_name}] в config.COMPANY_DB_FILTER нет записи для компании {cid} — пропуск", flush=True)
        return []
    sql = base_sql.replace("{company_filter}", "N'" + company_filter.replace("'", "''") + "'")
    sql = sql.replace("{company}", company_filter.replace("'", "''"))

    try:
        check_select_only(sql)
    except ValueError as e:
        print(f"❌ [{company_name}] SQL отклонён: {e}", flush=True)
        return []

    print(f"🔌 [{company_name}] Подключение к CRM…", flush=True)
    conn = crm_connect()
    try:
        df = pd.read_sql(sql, conn)
    except Exception as e:
        conn.close()
        print(f"❌ [{company_name}] ошибка выполнения SQL: {e}", flush=True)
        return []
    conn.close()
    print(f"📊 [{company_name}] Строк в отчёте: {len(df)} | колонки: {', '.join(map(str, df.columns))}", flush=True)

    if df.empty:
        print(f"❌ [{company_name}] SQL вернул 0 строк — рассылать нечего", flush=True)
        return []

    # ── определяем колонку-ЧСИ ────────────────────────────────
    gcol = (mailing.group_column or "").strip()
    if gcol not in df.columns:
        matches = [c for c in df.columns if str(c).strip().lower() == gcol.lower()]
        if matches:
            gcol = matches[0]
        else:
            print(f"❌ [{company_name}] колонка-ЧСИ «{mailing.group_column}» не найдена в результате. "
                  f"Доступные колонки: {', '.join(map(str, df.columns))}", flush=True)
            return []

    # ── справочник ЧСИ этой компании ──────────────────────────
    contacts = database.get_chsi_contacts(int(cid), only_active=True)
    fio_to_email = {}
    for c in contacts:
        key = c.fio_norm or normalize_fio(c.fio)
        if key and key not in fio_to_email:
            fio_to_email[key] = (c.email.strip(), c.fio)
    print(f"📇 [{company_name}] Активных ЧСИ в справочнике: {len(fio_to_email)}", flush=True)

    # ── группировка ──────────────────────────────────────────
    groups = []  # (fio_raw, DataFrame)
    for fio_val, gdf in df.groupby(gcol, dropna=True):
        if fio_val is None or str(fio_val).strip() == "":
            continue
        groups.append((str(fio_val).strip(), gdf))
    groups.sort(key=lambda x: x[0].lower())
    print(f"👥 [{company_name}] Уникальных ЧСИ в отчёте: {len(groups)}", flush=True)

    delay = float(smtp_cfg.get("send_delay_sec", 5) or 0)
    today_str = datetime.now().strftime("%d.%m.%Y")
    signature = MAILING_SIGNATURE.get(cid, {})

    def make_ctx(fio, n):
        return {
            "ФИО_ЧСИ": fio,
            "кол_во_записей": n,
            "компания": company_name,
            "компания_ru": signature.get("full_name_ru") or company_name,
            "компания_kz": signature.get("full_name_kz") or company_name,
            "телефон": signature.get("phone") or "",
            "дата": today_str,
        }

    journal = []  # (компания, fio, email, статус, строк, ошибка)

    # ── ТЕСТОВЫЙ РЕЖИМ ───────────────────────────────────────
    if MODE == "test":
        test_to = (_args.test_email or mailing.test_email or "").strip()
        if "@" not in test_to:
            print(f"❌ [{company_name}] режим test: не задан тестовый email "
                  f"(ни --test_email, ни в рассылке)", flush=True)
            return []
        if not groups:
            print(f"❌ [{company_name}] в отчёте нет ни одного непустого ЧСИ для теста", flush=True)
            return []

        fio, gdf = groups[0]
        ctx = make_ctx(fio, len(gdf))
        subject = "[ТЕСТ] " + render_placeholders(mailing.subject, ctx)
        body = render_placeholders(mailing.body_html, ctx)

        attach_path = None
        if mailing.attach_enabled:
            attach_path = OUT_DIR / (safe_filename(f"{company_name}_{fio}") + ".xlsx")
            build_attachment(attach_path, gdf.columns, gdf.itertuples(index=False, name=None))

        print(f"✉️  [{company_name}] ТЕСТ → {test_to} (данные ЧСИ «{fio}», строк {len(gdf)})", flush=True)
        try:
            send_one(smtp_cfg, test_to, [], subject, body, attach_path)
            journal.append((company_name, fio, test_to, "ОК (тест)", len(gdf), ""))
            print("✅ Тестовое письмо отправлено", flush=True)
        except Exception as e:
            journal.append((company_name, fio, test_to, "Ошибка", len(gdf), str(e)))
            print(f"❌ Ошибка отправки: {e}", flush=True)
        return journal

    # ── БОЕВОЙ РЕЖИМ ─────────────────────────────────────────
    bcc_list = _split_emails(mailing.bcc) + _split_emails(smtp_cfg.get("bcc"))
    bcc_list = list(dict.fromkeys(bcc_list))  # уникальные, порядок сохранён
    if bcc_list:
        print(f"📑 [{company_name}] BCC на каждое письмо: {', '.join(bcc_list)}", flush=True)

    sent = skipped = errors = 0
    for i, (fio, gdf) in enumerate(groups, 1):
        key = normalize_fio(fio)
        rec = fio_to_email.get(key)
        if not rec:
            skipped += 1
            journal.append((company_name, fio, "", "Без email", len(gdf), "нет в справочнике ЧСИ"))
            print(f"[{company_name} {i}/{len(groups)}] ⏭  {fio} — нет email в справочнике", flush=True)
            continue

        to_email, _ = rec
        ctx = make_ctx(fio, len(gdf))
        subject = render_placeholders(mailing.subject, ctx)
        body = render_placeholders(mailing.body_html, ctx)

        attach_path = None
        if mailing.attach_enabled:
            attach_path = OUT_DIR / (safe_filename(f"{company_name}_{fio}") + ".xlsx")
            build_attachment(attach_path, gdf.columns, gdf.itertuples(index=False, name=None))

        try:
            send_one(smtp_cfg, to_email, bcc_list, subject, body, attach_path)
            sent += 1
            journal.append((company_name, fio, to_email, "ОК", len(gdf), ""))
            print(f"[{company_name} {i}/{len(groups)}] ✅ {fio} → {to_email}", flush=True)
        except Exception as e:
            errors += 1
            journal.append((company_name, fio, to_email, "Ошибка", len(gdf), str(e)))
            print(f"[{company_name} {i}/{len(groups)}] ❌ {fio} → {to_email} | {e}", flush=True)
        finally:
            if delay and i < len(groups):
                time.sleep(delay)

    print(f"=== ИТОГ [{company_name}] === отправлено: {sent} | без email: {skipped} | ошибок: {errors}", flush=True)
    return journal


def main():
    mailing = database.get_mailing(int(_args.mailing_id))
    if not mailing:
        die(f"рассылка id={_args.mailing_id} не найдена")

    allowed = set(str(x) for x in mailing.company_id_list())
    bad = [cid for cid in COMPANY_IDS if cid not in allowed]
    if bad:
        die(f"рассылка id={mailing.id} не настроена на компани{'ю' if len(bad) == 1 else 'и'} "
            f"{', '.join(bad)} (настроено: {', '.join(sorted(allowed))})")

    sql = (mailing.effective_sql() or "").strip()
    if not sql:
        die("у рассылки не задан ни свой SQL, ни отчёт")

    journal = []
    for cid in COMPANY_IDS:
        journal.extend(run_for_company(cid, mailing, sql))

    if not journal:
        die("ни для одной компании не удалось отправить ни одного письма — см. лог выше")

    write_journal(mailing, journal)


def _split_emails(s):
    if not s:
        return []
    return [x.strip() for x in str(s).replace(";", ",").split(",") if "@" in x]


def write_journal(mailing, journal):
    path = OUT_DIR / f"Рассылка_{safe_filename(mailing.name)}_{datetime.now():%Y%m%d_%H%M}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Журнал"
    ws.append(["Компания", "ЧСИ", "Email", "Статус", "Строк", "Ошибка"])
    for row in journal:
        ws.append(list(row))
    wb.save(path)
    print(f"📄 Журнал: {path.name}", flush=True)


if __name__ == "__main__":
    main()
