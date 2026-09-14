# database.py
from sqlalchemy import create_engine, Column, Integer, String, Boolean, DateTime, ForeignKey, Table, Text, text
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship
from datetime import datetime, timedelta
import bcrypt

DATABASE_URL = "sqlite:///./users.db"

engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

# ═══════════════════════════════════════════════════════════════
# ТАБЛИЦА СВЯЗЕЙ: Пользователь <-> Скрипты (многие-ко-многим)
# ═══════════════════════════════════════════════════════════════

user_scripts = Table(
    'user_scripts',
    Base.metadata,
    Column('user_id', Integer, ForeignKey('users.id'), primary_key=True),
    Column('script_key', String, primary_key=True)
)

# ═══════════════════════════════════════════════════════════════
# ТАБЛИЦА СВЯЗЕЙ: Пользователь <-> Компании (многие-ко-многим, для роли "manager")
# ═══════════════════════════════════════════════════════════════

user_companies = Table(
    'user_companies',
    Base.metadata,
    Column('user_id', Integer, ForeignKey('users.id'), primary_key=True),
    Column('company_id', Integer, ForeignKey('companies.id'), primary_key=True)
)

# ═══════════════════════════════════════════════════════════════
# МОДЕЛИ
# ═══════════════════════════════════════════════════════════════

class Company(Base):
    __tablename__ = "companies"
    
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, unique=True, index=True)
    created_at = Column(DateTime, default=datetime.now)
    
    users = relationship("User", back_populates="company")

class User(Base):
    __tablename__ = "users"
    
    id = Column(Integer, primary_key=True, index=True)
    username = Column(String, unique=True, index=True)
    password_hash = Column(String)
    full_name = Column(String)
    role = Column(String)
    company_id = Column(Integer, ForeignKey("companies.id"), nullable=True)
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=datetime.now)
    
    company = relationship("Company", back_populates="users")
    
    # ✅ НОВОЕ: Связь со скриптами (разрешённые скрипты)
    # Храним ключи скриптов как строки в таблице user_scripts
    # Для доступа используем: user.allowed_scripts (list строк)
    
    @property
    def allowed_scripts(self):
        """Список разрешённых скриптов для пользователя"""
        db = SessionLocal()
        result = db.execute(
            user_scripts.select().where(user_scripts.c.user_id == self.id)
        ).fetchall()
        db.close()
        return [row.script_key for row in result]
    
    def has_script_access(self, script_key: str) -> bool:
        """Проверка доступа к скрипту"""
        # Админ имеет доступ ко всем скриптам
        if self.role == "admin":
            return True
        
        # Остальные - только к разрешённым
        return script_key in self.allowed_scripts
    
class Schedule(Base):
    __tablename__ = "schedules"

    id = Column(Integer, primary_key=True, index=True)
    script_key = Column(String, nullable=False)
    hour = Column(Integer, nullable=False)                # 0-23
    minute = Column(Integer, nullable=False)               # 0-59
    weekdays = Column(String, nullable=False, default="0,1,2,3,4,5,6")  # 0=пн ... 6=вс
    params_json = Column(Text, default="{}")
    company_id = Column(Integer, ForeignKey("companies.id"), nullable=True)  # None = без привязки к компании
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=datetime.now)

    company = relationship("Company")

    def weekday_list(self):
        return [int(x) for x in self.weekdays.split(",") if x != ""]

    def params(self):
        import json
        try:
            p = json.loads(self.params_json or "{}")
        except Exception:
            p = {}
        # company_id из отдельной колонки всегда подставляется в параметры
        # запуска скрипта (--company_id), чтобы автозапуск шёл под нужной компанией.
        if self.company_id:
            p["company_id"] = str(self.company_id)
        return p

class JobRun(Base):
    __tablename__ = "job_runs"

    id = Column(String, primary_key=True)          # job_id (uuid hex)
    script_key = Column(String, nullable=False)
    status = Column(String, default="queued")       # queued/running/success/error
    source = Column(String, default="manual")        # manual/scheduler
    username = Column(String, default="—")
    company_id = Column(Integer, nullable=True)
    company_name = Column(String, nullable=True)
    created_at = Column(DateTime, default=datetime.now)
    started_at = Column(DateTime, nullable=True)
    finished_at = Column(DateTime, nullable=True)
    error_text = Column(Text, nullable=True)


# ═══════════════════════════════════════════════════════════════
# РАССЫЛКИ ПО ОТЧЁТАМ (ЧСИ)
# ───────────────────────────────────────────────────────────────
# MailingReport — именованный SQL-отчёт (готовые из списка).
# ChsiContact   — справочник получателей (ФИО ЧСИ -> email), по компании.
# Mailing       — сама рассылка: отчёт/свой SQL + текст письма + расписание.
# ═══════════════════════════════════════════════════════════════

class MailingReport(Base):
    __tablename__ = "mailing_reports"

    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, nullable=False)
    company_id = Column(Integer, ForeignKey("companies.id"), nullable=True)  # None = общий
    sql_text = Column(Text, nullable=False, default="")
    created_at = Column(DateTime, default=datetime.now)

    company = relationship("Company")


class ChsiContact(Base):
    __tablename__ = "chsi_contacts"

    id = Column(Integer, primary_key=True, index=True)
    company_id = Column(Integer, ForeignKey("companies.id"), nullable=False)
    fio = Column(String, nullable=False)
    fio_norm = Column(String, index=True, default="")
    email = Column(String, nullable=False)
    is_active = Column(Boolean, default=True)
    note = Column(String, nullable=True)
    created_at = Column(DateTime, default=datetime.now)

    company = relationship("Company")


class Mailing(Base):
    __tablename__ = "mailings"

    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, nullable=False)
    company_id = Column(Integer, ForeignKey("companies.id"), nullable=False)  # первая/основная компания
    extra_company_ids = Column(String, nullable=True)  # ост. компании через запятую (мульти-рассылка)
    report_id = Column(Integer, ForeignKey("mailing_reports.id"), nullable=True)
    sql_text = Column(Text, nullable=True)          # «свой SQL» (приоритетнее report)
    group_column = Column(String, nullable=False, default="")  # колонка результата = ЧСИ
    subject = Column(String, nullable=False, default="")
    body_html = Column(Text, nullable=False, default="")
    attach_enabled = Column(Boolean, default=True)
    test_email = Column(String, nullable=True)
    bcc = Column(String, nullable=True)             # адреса через запятую (скрытая копия)
    is_active = Column(Boolean, default=True)
    schedule_enabled = Column(Boolean, default=False)
    schedule_hour = Column(Integer, nullable=True)
    schedule_minute = Column(Integer, nullable=True)
    schedule_weekdays = Column(String, default="0,1,2,3,4,5,6")
    created_at = Column(DateTime, default=datetime.now)

    company = relationship("Company")
    report = relationship("MailingReport")

    def weekday_list(self):
        return [int(x) for x in (self.schedule_weekdays or "").split(",") if x != ""]

    def company_id_list(self):
        """Все компании рассылки: основная (company_id) + доп. (extra_company_ids),
        без дублей, порядок сохранён."""
        ids = [self.company_id]
        for x in (self.extra_company_ids or "").split(","):
            x = x.strip()
            if x and x.isdigit() and int(x) not in ids:
                ids.append(int(x))
        return ids

    def effective_sql(self) -> str:
        if self.sql_text and self.sql_text.strip():
            return self.sql_text
        if self.report and self.report.sql_text:
            return self.report.sql_text
        return ""





# ═══════════════════════════════════════════════════════════════
# ФУНКЦИИ ДЛЯ РАБОТЫ С РАСПИСАНИЕМ
# ═══════════════════════════════════════════════════════════════

def get_all_schedules() -> list:
    db = SessionLocal()
    result = db.query(Schedule).order_by(Schedule.hour, Schedule.minute).all()
    db.close()
    return result

def get_active_schedules() -> list:
    db = SessionLocal()
    result = db.query(Schedule).filter(Schedule.is_active == True).all()
    db.close()
    return result

def get_schedule(schedule_id: int):
    db = SessionLocal()
    s = db.query(Schedule).filter(Schedule.id == schedule_id).first()
    db.close()
    return s

def create_schedule(script_key: str, hour: int, minute: int, weekdays: str, params: dict = None, company_id: int = None):
    import json
    db = SessionLocal()
    s = Schedule(
        script_key=script_key,
        hour=hour,
        minute=minute,
        weekdays=weekdays,
        params_json=json.dumps(params or {}),
        company_id=company_id,
        is_active=True,
    )
    db.add(s)
    db.commit()
    db.refresh(s)
    db.close()
    return s

def update_schedule(schedule_id: int, script_key: str, hour: int, minute: int, weekdays: str, params: dict = None, company_id: int = None):
    import json
    db = SessionLocal()
    s = db.query(Schedule).filter(Schedule.id == schedule_id).first()
    if s:
        s.script_key = script_key
        s.hour = hour
        s.minute = minute
        s.weekdays = weekdays
        s.params_json = json.dumps(params or {})
        s.company_id = company_id
        db.commit()
    db.close()

def delete_schedule(schedule_id: int):
    db = SessionLocal()
    s = db.query(Schedule).filter(Schedule.id == schedule_id).first()
    if s:
        db.delete(s)
        db.commit()
    db.close()

def toggle_schedule(schedule_id: int):
    db = SessionLocal()
    s = db.query(Schedule).filter(Schedule.id == schedule_id).first()
    if s:
        s.is_active = not s.is_active
        db.commit()
    db.close()


# ═══════════════════════════════════════════════════════════════
# ФУНКЦИИ ДЛЯ ИСТОРИИ ЗАПУСКОВ (JobRun)
# ═══════════════════════════════════════════════════════════════

def create_job_run(job_id: str, script_key: str, source: str, username: str, company_id: int = None, company_name: str = None):
    db = SessionLocal()
    jr = JobRun(
        id=job_id,
        script_key=script_key,
        status="queued",
        source=source,
        username=username,
        company_id=company_id,
        company_name=company_name,
    )
    db.add(jr)
    db.commit()
    db.close()

def update_job_run_status(job_id: str, status: str, started_at=None, finished_at=None, error_text=None):
    db = SessionLocal()
    jr = db.query(JobRun).filter(JobRun.id == job_id).first()
    if jr:
        jr.status = status
        if started_at:
            jr.started_at = started_at
        if finished_at:
            jr.finished_at = finished_at
        if error_text:
            jr.error_text = error_text
        db.commit()
    db.close()

def _job_runs_query(db, source=None, status=None, only_today=False,
                     date_from=None, date_to=None, script_key=None, company_id=None):
    q = db.query(JobRun)
    if source:
        q = q.filter(JobRun.source == source)
    if status:
        q = q.filter(JobRun.status == status)
    if company_id:
        q = q.filter(JobRun.company_id == company_id)
    if script_key:
        q = q.filter(JobRun.script_key == script_key)
    if only_today:
        today = datetime.now().strftime("%Y-%m-%d")
        q = q.filter(JobRun.created_at >= datetime.strptime(today, "%Y-%m-%d"))
    if date_from:
        q = q.filter(JobRun.created_at >= datetime.combine(date_from, datetime.min.time()))
    if date_to:
        q = q.filter(JobRun.created_at < datetime.combine(date_to, datetime.min.time()) + timedelta(days=1))
    return q


def get_job_runs(source: str = None, status: str = None, only_today: bool = False,
                  date_from=None, date_to=None, script_key: str = None,
                  limit: int = 200, offset: int = 0, company_id: int = None) -> list:
    """date_from/date_to — date или datetime (включительно по дням, по created_at)."""
    db = SessionLocal()
    q = _job_runs_query(db, source, status, only_today, date_from, date_to, script_key, company_id)
    q = q.order_by(JobRun.created_at.desc()).offset(offset).limit(limit)
    result = q.all()
    db.close()
    return result


def count_job_runs(source: str = None, status: str = None, only_today: bool = False,
                    date_from=None, date_to=None, script_key: str = None, company_id: int = None) -> int:
    db = SessionLocal()
    q = _job_runs_query(db, source, status, only_today, date_from, date_to, script_key, company_id)
    n = q.count()
    db.close()
    return n


def get_job_run_script_keys() -> list:
    db = SessionLocal()
    rows = db.query(JobRun.script_key).distinct().all()
    db.close()
    return sorted({r[0] for r in rows if r[0]})

def _ensure_columns():
    """
    Base.metadata.create_all() создаёт только отсутствующие таблицы —
    для уже существующего users.db новые колонки (company_id у schedules,
    company_id/company_name у job_runs) сами не появятся. Добавляем их
    вручную; если колонка уже есть — просто игнорируем ошибку.
    """
    statements = [
        "ALTER TABLE schedules ADD COLUMN company_id INTEGER",
        "ALTER TABLE job_runs ADD COLUMN company_id INTEGER",
        "ALTER TABLE job_runs ADD COLUMN company_name VARCHAR",
        "ALTER TABLE mailings ADD COLUMN extra_company_ids VARCHAR",
    ]
    with engine.connect() as conn:
        for stmt in statements:
            try:
                conn.execute(text(stmt))
                conn.commit()
            except Exception:
                pass  # колонка уже существует


def init_db():
    """Инициализация базы данных"""
    Base.metadata.create_all(bind=engine)
    _ensure_columns()

    db = SessionLocal()
    
    # Создаём компании
    companies_data = ["А-Омега", "KPI", "Orion", "InvestWay"
    ""]
    for company_name in companies_data:
        company = db.query(Company).filter(Company.name == company_name).first()
        if not company:
            company = Company(name=company_name)
            db.add(company)
            print(f"✅ Создана компания: {company_name}")
    
    db.commit()
    
    # Создаём админа
    admin = db.query(User).filter(User.username == "admin").first()
    if not admin:
        admin = User(
            username="admin",
            password_hash=bcrypt.hashpw("admin123".encode(), bcrypt.gensalt()).decode(),
            full_name="Администратор",
            role="admin",
            company_id=None,
            is_active=True
        )
        db.add(admin)
        db.commit()
        print("✅ Создан админ: username='admin', password='admin123'")
    
    db.close()

def get_db():
    """Получить сессию БД"""
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

# ═══════════════════════════════════════════════════════════════
# ФУНКЦИИ ДЛЯ РАБОТЫ С РАЗРЕШЕНИЯМИ
# ═══════════════════════════════════════════════════════════════

def add_script_permission(user_id: int, script_key: str):
    """Добавить разрешение на скрипт"""
    db = SessionLocal()
    # Проверяем что такого разрешения ещё нет
    exists = db.execute(
        user_scripts.select().where(
            user_scripts.c.user_id == user_id,
            user_scripts.c.script_key == script_key
        )
    ).fetchone()
    
    if not exists:
        db.execute(user_scripts.insert().values(user_id=user_id, script_key=script_key))
        db.commit()
    db.close()

def remove_script_permission(user_id: int, script_key: str):
    """Удалить разрешение на скрипт"""
    db = SessionLocal()
    db.execute(
        user_scripts.delete().where(
            user_scripts.c.user_id == user_id,
            user_scripts.c.script_key == script_key
        )
    )
    db.commit()
    db.close()

def set_user_scripts(user_id: int, script_keys: list):
    """Установить список разрешённых скриптов для пользователя"""
    db = SessionLocal()
    
    # Удаляем все старые разрешения
    db.execute(user_scripts.delete().where(user_scripts.c.user_id == user_id))
    
    # Добавляем новые
    for script_key in script_keys:
        db.execute(user_scripts.insert().values(user_id=user_id, script_key=script_key))
    
    db.commit()
    db.close()

def get_user_scripts(user_id: int) -> list:
    """Получить список разрешённых скриптов"""
    db = SessionLocal()
    result = db.execute(
        user_scripts.select().where(user_scripts.c.user_id == user_id)
    ).fetchall()
    db.close()
    return [row.script_key for row in result]


def set_user_companies(user_id: int, company_ids: list):
    """Установить список компаний, доступных пользователю (роль 'manager')"""
    db = SessionLocal()

    db.execute(user_companies.delete().where(user_companies.c.user_id == user_id))

    for company_id in company_ids:
        db.execute(user_companies.insert().values(user_id=user_id, company_id=company_id))

    db.commit()
    db.close()


def get_user_companies(user_id: int) -> list:
    """Получить список ID компаний, доступных пользователю (роль 'manager')"""
    db = SessionLocal()
    result = db.execute(
        user_companies.select().where(user_companies.c.user_id == user_id)
    ).fetchall()
    db.close()
    return [row.company_id for row in result]


# ═══════════════════════════════════════════════════════════════
# РАССЫЛКИ ПО ОТЧЁТАМ — CRUD
# ═══════════════════════════════════════════════════════════════

def _fio_norm(fio: str) -> str:
    try:
        from mailing_common import normalize_fio
        return normalize_fio(fio)
    except Exception:
        return (fio or "").strip().lower()


# ── Именованные SQL-отчёты ─────────────────────────────────────

def get_mailing_reports(company_id: int = None) -> list:
    db = SessionLocal()
    q = db.query(MailingReport)
    if company_id is not None:
        # общие (company_id IS NULL) + отчёты этой компании
        q = q.filter((MailingReport.company_id == company_id) | (MailingReport.company_id.is_(None)))
    result = q.order_by(MailingReport.name).all()
    db.close()
    return result


def get_mailing_report(report_id: int):
    db = SessionLocal()
    r = db.query(MailingReport).filter(MailingReport.id == report_id).first()
    db.close()
    return r


def create_mailing_report(name: str, sql_text: str, company_id: int = None):
    db = SessionLocal()
    r = MailingReport(name=name, sql_text=sql_text or "", company_id=company_id)
    db.add(r)
    db.commit()
    db.refresh(r)
    db.close()
    return r


def update_mailing_report(report_id: int, name: str, sql_text: str, company_id: int = None):
    db = SessionLocal()
    r = db.query(MailingReport).filter(MailingReport.id == report_id).first()
    if r:
        r.name = name
        r.sql_text = sql_text or ""
        r.company_id = company_id
        db.commit()
    db.close()


def delete_mailing_report(report_id: int):
    db = SessionLocal()
    r = db.query(MailingReport).filter(MailingReport.id == report_id).first()
    if r:
        db.delete(r)
        db.commit()
    db.close()


# ── Справочник ЧСИ ────────────────────────────────────────────

def get_chsi_contacts(company_id: int, only_active: bool = False) -> list:
    db = SessionLocal()
    q = db.query(ChsiContact).filter(ChsiContact.company_id == company_id)
    if only_active:
        q = q.filter(ChsiContact.is_active == True)  # noqa: E712
    result = q.order_by(ChsiContact.fio).all()
    db.close()
    return result


def get_chsi_contact(contact_id: int):
    db = SessionLocal()
    c = db.query(ChsiContact).filter(ChsiContact.id == contact_id).first()
    db.close()
    return c


def create_chsi_contact(company_id: int, fio: str, email: str, is_active: bool = True, note: str = None):
    db = SessionLocal()
    c = ChsiContact(
        company_id=company_id, fio=fio, fio_norm=_fio_norm(fio),
        email=(email or "").strip(), is_active=is_active, note=note,
    )
    db.add(c)
    db.commit()
    db.refresh(c)
    db.close()
    return c


def update_chsi_contact(contact_id: int, fio: str, email: str, is_active: bool = True, note: str = None):
    db = SessionLocal()
    c = db.query(ChsiContact).filter(ChsiContact.id == contact_id).first()
    if c:
        c.fio = fio
        c.fio_norm = _fio_norm(fio)
        c.email = (email or "").strip()
        c.is_active = is_active
        c.note = note
        db.commit()
    db.close()


def delete_chsi_contact(contact_id: int):
    db = SessionLocal()
    c = db.query(ChsiContact).filter(ChsiContact.id == contact_id).first()
    if c:
        db.delete(c)
        db.commit()
    db.close()


def bulk_add_chsi_contacts(company_id: int, rows: list) -> int:
    """rows: list[(fio, email)]. Пропускает пустые и уже существующие (по fio_norm+email)."""
    db = SessionLocal()
    added = 0
    existing = {
        (c.fio_norm, (c.email or "").strip().lower())
        for c in db.query(ChsiContact).filter(ChsiContact.company_id == company_id).all()
    }
    for fio, email in rows:
        fio = (fio or "").strip()
        email = (email or "").strip()
        if not fio or "@" not in email:
            continue
        key = (_fio_norm(fio), email.lower())
        if key in existing:
            continue
        existing.add(key)
        db.add(ChsiContact(
            company_id=company_id, fio=fio, fio_norm=_fio_norm(fio),
            email=email, is_active=True,
        ))
        added += 1
    db.commit()
    db.close()
    return added


# ── Рассылки ─────────────────────────────────────────────────

def get_mailings(company_id: int = None) -> list:
    db = SessionLocal()
    q = db.query(Mailing)
    if company_id is not None:
        q = q.filter(Mailing.company_id == company_id)
    result = q.order_by(Mailing.name).all()
    db.close()
    return result


def get_mailing(mailing_id: int):
    db = SessionLocal()
    m = db.query(Mailing).filter(Mailing.id == mailing_id).first()
    if m:
        _ = m.report  # подгрузить связь до закрытия сессии
    db.close()
    return m


def create_mailing(**kw):
    db = SessionLocal()
    m = Mailing(**kw)
    db.add(m)
    db.commit()
    db.refresh(m)
    db.close()
    return m


def update_mailing(mailing_id: int, **kw):
    db = SessionLocal()
    m = db.query(Mailing).filter(Mailing.id == mailing_id).first()
    if m:
        for k, v in kw.items():
            setattr(m, k, v)
        db.commit()
    db.close()


def delete_mailing(mailing_id: int):
    db = SessionLocal()
    m = db.query(Mailing).filter(Mailing.id == mailing_id).first()
    if m:
        db.delete(m)
        db.commit()
    db.close()


def toggle_mailing(mailing_id: int):
    db = SessionLocal()
    m = db.query(Mailing).filter(Mailing.id == mailing_id).first()
    if m:
        m.is_active = not m.is_active
        db.commit()
    db.close()


def get_scheduled_mailings() -> list:
    db = SessionLocal()
    result = (
        db.query(Mailing)
        .filter(Mailing.is_active == True, Mailing.schedule_enabled == True)  # noqa: E712
        .all()
    )
    db.close()
    return result