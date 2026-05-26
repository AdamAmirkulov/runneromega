# auth.py
from fastapi import Depends, HTTPException, status, Cookie, Request
from fastapi.responses import RedirectResponse
from sqlalchemy.orm import Session
from database import get_db, User, get_user_companies
import bcrypt
import secrets

# Хранилище сессий (в памяти)
# Структура: {session_id: {"username": "...", "user_id": ..., "role": "...", "company_id": ...}}
SESSIONS = {}

def create_session(user_id: int, username: str, role: str, company_id: int = None) -> str:
    """Создать сессию с данными пользователя"""
    session_id = secrets.token_urlsafe(32)
    SESSIONS[session_id] = {
        "user_id": user_id,
        "username": username,
        "role": role,
        "company_id": company_id
    }
    return session_id

def get_current_user(
    request: Request,
    session_id: str = Cookie(None, alias="session"),
    db: Session = Depends(get_db)
) -> User:
    """Получить текущего пользователя"""
    if not session_id or session_id not in SESSIONS:
        return RedirectResponse("/login", status_code=303)
    
    session_data = SESSIONS[session_id]
    user = db.query(User).filter(User.id == session_data["user_id"]).first()
    
    if not user or not user.is_active:
        # Удаляем невалидную сессию
        if session_id in SESSIONS:
            del SESSIONS[session_id]
        return RedirectResponse("/login", status_code=303)
    
    return user

def get_admin_user(current_user: User = Depends(get_current_user)) -> User:
    """Проверка что пользователь - админ"""
    if isinstance(current_user, RedirectResponse):
        return current_user
    
    if current_user.role != "admin":
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Доступ запрещён"
        )
    return current_user

def verify_password(plain_password: str, hashed_password: str) -> bool:
    """Проверка пароля"""
    return bcrypt.checkpw(plain_password.encode(), hashed_password.encode())

def hash_password(password: str) -> str:
    """Хеширование пароля"""
    return bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()

def user_can_access_company(user: User, company_id: int) -> bool:
    """Проверка доступа к компании"""
    # Админ имеет доступ ко всем компаниям
    if user.role == "admin":
        return True

    # Менеджер - к нескольким компаниям, назначенным ему (таблица user_companies)
    if user.role == "manager":
        return company_id in get_user_companies(user.id)

    # Оператор - только к своей единственной компании
    return user.company_id == company_id


def get_user_allowed_company_ids(user: User) -> list | None:
    """
    Список ID компаний, доступных пользователю.
    None означает "все компании" (админ) - в отличие от [] ("ни одной").
    """
    if user.role == "admin":
        return None
    if user.role == "manager":
        return get_user_companies(user.id)
    return [user.company_id] if user.company_id else []


def get_active_company_id(session_id: str | None) -> int | None:
    """
    Компания, выбранная admin/manager в шапке сайта ("Компания:" над списком
    скриптов). Хранится в сессии, чтобы не выбирать её заново на каждой странице —
    именно она автоматически подставляется как --company_id при запуске скриптов.
    """
    if not session_id or session_id not in SESSIONS:
        return None
    return SESSIONS[session_id].get("active_company_id")


def set_active_company_id(session_id: str | None, company_id: int | None) -> None:
    """Запоминает выбор компании из шапки сайта для текущей сессии."""
    if session_id and session_id in SESSIONS:
        SESSIONS[session_id]["active_company_id"] = company_id