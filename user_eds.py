# user_eds.py
"""
Личные ЭЦП сотрудников для подачи исков на portal-sot.kz.

Сотрудник (или админ за него) загружает в интерфейсе свой .p12, пароль от него
и пароль от portal-sot.kz. При загрузке файл сразу проверяется (пароль, ИИН,
срок) — scripts/eds_cert.py. Хранение:
  data/eds_store/<user_id>.p12     — файл ЭЦП
  data/eds_store/secret.key        — ключ Fernet (создаётся сам, НЕ в git)
  users.db / user_eds              — ИИН, ФИО, срок, зашифрованные пароли

Скрипту подачи данные уходят через переменные окружения процесса (не через
командную строку — она пишется в logs.txt задачи), см. job_env().
"""
from __future__ import annotations

import importlib.util
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

from cryptography.fernet import Fernet, InvalidToken

from database import User, UserEds

BASE_DIR = Path(__file__).resolve().parent
STORE_DIR = BASE_DIR / "data" / "eds_store"
KEY_PATH = STORE_DIR / "secret.key"
MAX_P12_BYTES = 64 * 1024

# scripts/ в sys.path веб-процесса не добавляем (там свой config.py) — грузим файл напрямую
_spec = importlib.util.spec_from_file_location("eds_cert", BASE_DIR / "scripts" / "eds_cert.py")
eds_cert = importlib.util.module_from_spec(_spec)
sys.modules["eds_cert"] = eds_cert  # dataclass ищет модуль в sys.modules
_spec.loader.exec_module(eds_cert)
EdsError = eds_cert.EdsError


class EdsMissing(Exception):
    """У сотрудника нет годной ЭЦП — подачу запускать нельзя."""


def _fernet() -> Fernet:
    STORE_DIR.mkdir(parents=True, exist_ok=True)
    if not KEY_PATH.exists():
        KEY_PATH.write_bytes(Fernet.generate_key())
    return Fernet(KEY_PATH.read_bytes().strip())


def _enc(s: str) -> str:
    return _fernet().encrypt(s.encode("utf-8")).decode("ascii")


def _dec(s: str) -> str:
    return _fernet().decrypt(s.encode("ascii")).decode("utf-8")


def cert_path(user_id: int) -> Path:
    return STORE_DIR / f"{int(user_id)}.p12"


def get(db, user_id: int) -> Optional[UserEds]:
    return db.query(UserEds).filter(UserEds.user_id == user_id).first()


def save(db, user: User, data: bytes, file_name: str, eds_password: str,
         portal_password: str, uploaded_by: str) -> UserEds:
    """Проверяет и сохраняет ЭЦП. Ошибки — EdsError с понятным текстом."""
    if not data:
        raise EdsError("Файл ЭЦП не выбран")
    if len(data) > MAX_P12_BYTES:
        raise EdsError("Слишком большой файл — это не файл ЭЦП (.p12)")
    if not eds_password:
        raise EdsError("Укажите пароль от ЭЦП")
    if not portal_password:
        raise EdsError("Укажите пароль от portal-sot.kz")

    info = eds_cert.read_p12(data, eds_password)
    if info.valid_to < datetime.now():
        raise EdsError(f"Срок действия ЭЦП истёк {info.valid_to:%d.%m.%Y}")

    path = cert_path(user.id)
    STORE_DIR.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_bytes(data)
    os.replace(tmp, path)

    row = get(db, user.id) or UserEds(user_id=user.id)
    row.iin = info.iin
    row.full_name = info.full_name
    row.org_bin = info.bin
    row.valid_from = info.valid_from
    row.valid_to = info.valid_to
    row.file_name = (file_name or "")[:200]
    row.eds_password_enc = _enc(eds_password)
    row.portal_password_enc = _enc(portal_password)
    row.uploaded_at = datetime.now()
    row.uploaded_by = uploaded_by
    db.add(row)
    db.commit()
    return row


def delete(db, user_id: int) -> None:
    row = get(db, user_id)
    if row:
        db.delete(row)
        db.commit()
    try:
        cert_path(user_id).unlink()
    except FileNotFoundError:
        pass


def job_env(db, username: str) -> dict:
    """Переменные окружения для скрипта, которому нужна личная ЭЦП того, кто
    запустил задачу. Нет / просрочена / битая — EdsMissing."""
    user = db.query(User).filter(User.username == username).first() if username else None
    if not user:
        raise EdsMissing("Задача запущена не сотрудником (автозапуск?) — подача по личной ЭЦП невозможна")
    row = get(db, user.id)
    path = cert_path(user.id)
    if not row or not path.exists():
        raise EdsMissing(f"У пользователя {username} не загружена личная ЭЦП (раздел «🔑 Моя ЭЦП»)")
    if row.valid_to and row.valid_to < datetime.now():
        raise EdsMissing(f"Личная ЭЦП пользователя {username} просрочена ({row.valid_to:%d.%m.%Y}) — загрузите новую")
    try:
        eds_password = _dec(row.eds_password_enc)
        portal_password = _dec(row.portal_password_enc)
    except InvalidToken:
        raise EdsMissing(
            f"Не удалось расшифровать пароли ЭЦП пользователя {username} (сменился {KEY_PATH.name}?) — загрузите ЭЦП заново"
        )
    return {
        "PORTAL_PERSONAL_EDS_CERT": str(path),
        "PORTAL_PERSONAL_EDS_PASSWORD": eds_password,
        "PORTAL_PERSONAL_PORTAL_PASSWORD": portal_password,
        "PORTAL_PERSONAL_IIN": row.iin,
        "PORTAL_PERSONAL_FIO": row.full_name or "",
        "PORTAL_PERSONAL_BIN": row.org_bin or "",
        "PORTAL_PERSONAL_USERNAME": username,
    }
