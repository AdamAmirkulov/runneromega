# database.py
from sqlalchemy import create_engine, Column, Integer, String, Boolean, DateTime, ForeignKey, Table, Text, text
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship
from datetime import datetime
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

def get_job_runs(source: str = None, status: str = None, only_today: bool = False, limit: int = 200, company_id: int = None) -> list:
    db = SessionLocal()
    q = db.query(JobRun)
    if source:
        q = q.filter(JobRun.source == source)
    if status:
        q = q.filter(JobRun.status == status)
    if company_id:
        q = q.filter(JobRun.company_id == company_id)
    if only_today:
        today = datetime.now().strftime("%Y-%m-%d")
        q = q.filter(JobRun.created_at >= datetime.strptime(today, "%Y-%m-%d"))
    q = q.order_by(JobRun.created_at.desc()).limit(limit)
    result = q.all()
    db.close()
    return result

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