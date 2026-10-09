# -*- coding: utf-8 -*-
"""
portal-sot.kz: браузер ТОЛЬКО для входа, дальше — чистый HTTP с бережным темпом.

Схема:
  1. Токены (access/refresh + cookies + User-Agent) хранятся на диске по компании:
     data/portal_sot_tokens/<company_id>.json.
  2. access_token живёт ~15 минут, refresh_token ~24 часа. Когда access истёк —
     POST /api/public/auth/refresh (без браузера, без ЭЦП). Браузер + ЭЦП
     поднимаются, если refresh истёк/отклонён ИЛИ если токен после refresh всё
     равно получает 401: когда сессия на портале умерла, /auth/refresh отвечает
     200 и выдаёт новый access_token, но API его не принимает (сам сайт в этой
     ситуации тоже уходит на вход по ЭЦП — см. data/portal_debug).
  3. Браузер открывается функцией browser_login() из скрипта и СРАЗУ закрывается
     после получения токенов.
  4. Все запросы идут через get(): случайная пауза между запросами, длинная
     пауза каждые N запросов, отступ на 429/5xx, немедленная остановка на 403
     (это блок WAF — долбить дальше = продлевать бан).
  5. Межпроцессный лок: к portal-sot.kz с этой машины одновременно ходит только
     один скрипт (все компании идут с одного IP).
"""
import base64
import json
import os
import random
import time
from pathlib import Path

import requests

PORTAL_SOT_BASE = "https://portal-sot.kz"

TOKEN_DIR = Path(__file__).resolve().parent.parent / "data" / "portal_sot_tokens"
LOCK_PATH = TOKEN_DIR / "portal_sot.lock"

# Темп. Меняйте здесь, а не в скриптах.
MIN_GAP_SECONDS = (1.5, 3.0)        # случайная пауза между любыми двумя запросами
LONG_PAUSE_EVERY = 50               # каждые N запросов ...
LONG_PAUSE_SECONDS = (30.0, 60.0)   # ... передышка
BACKOFF_429_5XX = (60, 180, 600)    # ожидания при 429/5xx; после последнего — стоп
MAX_REQUESTS_PER_RUN = 3000         # предохранитель на один запуск
MAX_BROWSER_LOGINS_PER_RUN = 3      # входов по ЭЦП за запуск; больше — портал явно не пускает


class PortalBlocked(RuntimeError):
    """Портал явно нас режет (403 / 429 после всех ожиданий) — прогон надо остановить."""


def _jwt_exp(token: str) -> float:
    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        return float(json.loads(base64.urlsafe_b64decode(payload)).get("exp") or 0)
    except Exception:
        return 0.0


def is_waf_block(r) -> bool:
    """403 от nginx/WAF (HTML/пустое тело) — это блок. 403 с JSON — обычный
    ответ API (нет прав на конкретный объект), его обрабатывает вызывающий код."""
    if r.status_code != 403:
        return False
    ctype = (r.headers.get("Content-Type") or "").lower()
    return "json" not in ctype


BLOCK_MESSAGE = ("portal-sot.kz ответил 403 от WAF — похоже на блокировку. Прогон остановлен, "
                 "чтобы не продлевать бан. Подождите несколько часов, прежде чем запускать снова.")


def portal_lock(log=print) -> "_ProcessLock":
    """Общий лок для ВСЕХ скриптов, которые ходят на portal-sot.kz с этой машины."""
    return _ProcessLock(LOCK_PATH, log)


class _ProcessLock:
    """Файловый лок через msvcrt: ОС сама снимает его, если процесс умер."""

    def __init__(self, path: Path, log):
        self.path, self.log, self.fh = path, log, None

    def acquire(self):
        import msvcrt
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.fh = open(self.path, "a+")
        warned = False
        while True:
            try:
                self.fh.seek(0)
                msvcrt.locking(self.fh.fileno(), msvcrt.LK_NBLCK, 1)
                return
            except OSError:
                if not warned:
                    self.log("⏳ К portal-sot.kz уже ходит другой скрипт — жду, пока он закончит...")
                    warned = True
                time.sleep(10)

    def release(self):
        if not self.fh:
            return
        try:
            import msvcrt
            self.fh.seek(0)
            msvcrt.locking(self.fh.fileno(), msvcrt.LK_UNLCK, 1)
        except Exception:
            pass
        self.fh.close()
        self.fh = None


class PortalSotHttp:
    """
    browser_login() -> dict(access_token, refresh_token, cookies=[{name,value,domain,path}], user_agent)
    Должна открыть браузер, войти через ЭЦП, прочитать токены и ЗАКРЫТЬ браузер.

    Использование:
        with PortalSotHttp(company_id, browser_login, log) as portal:
            r = portal.get("/api/secure/gbdfl/v2/byIin/123456789012")
    """

    def __init__(self, company_id: str, browser_login, log=print):
        self.company_id = str(company_id)
        self.browser_login = browser_login
        self.log = log
        self.token_path = TOKEN_DIR / f"{self.company_id}.json"
        self.tokens = {}
        self.session = None
        self._last_request_at = 0.0
        self._count = 0
        self._browser_logins = 0
        self._lock = _ProcessLock(LOCK_PATH, log)

    # ---------- жизненный цикл ----------
    def __enter__(self):
        self._lock.acquire()
        try:
            self._ensure_session()
        except Exception:
            self._lock.release()
            raise
        return self

    def __exit__(self, *exc):
        self._lock.release()

    # ---------- токены ----------
    def _load_tokens(self):
        try:
            return json.loads(self.token_path.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def _save_tokens(self):
        TOKEN_DIR.mkdir(parents=True, exist_ok=True)
        tmp = self.token_path.with_suffix(".tmp")
        tmp.write_text(json.dumps(self.tokens, ensure_ascii=False), encoding="utf-8")
        os.replace(tmp, self.token_path)

    def _build_session(self):
        s = requests.Session()
        s.headers.update({
            "Authorization": f"Bearer {self.tokens['access_token']}",
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "ru",
            "Accept-Encoding": "gzip, deflate",
            "Origin": PORTAL_SOT_BASE,
            "Referer": PORTAL_SOT_BASE + "/",
        })
        if self.tokens.get("user_agent"):
            s.headers["User-Agent"] = self.tokens["user_agent"]
        for c in self.tokens.get("cookies") or []:
            try:
                s.cookies.set(c["name"], c["value"],
                              domain=c.get("domain") or "portal-sot.kz",
                              path=c.get("path") or "/")
            except Exception:
                pass
        self.session = s

    def _ensure_session(self):
        self.tokens = self._load_tokens()
        now = time.time()
        if self.tokens.get("access_token") and _jwt_exp(self.tokens["access_token"]) > now + 60:
            self.log("✅ portal-sot.kz: сохранённый access_token ещё действует — браузер не нужен")
        elif self.tokens.get("refresh_token") and _jwt_exp(self.tokens["refresh_token"]) > now + 60 \
                and self._refresh():
            pass
        else:
            self._login_via_browser()
        self._build_session()

    def _refresh(self) -> bool:
        """POST /api/public/auth/refresh — без браузера и без ЭЦП."""
        self._throttle()
        try:
            r = requests.post(
                PORTAL_SOT_BASE + "/api/public/auth/refresh",
                json={"refreshToken": self.tokens.get("refresh_token", "")},
                headers={
                    "Accept": "*/*", "Accept-Language": "ru",
                    "Origin": PORTAL_SOT_BASE, "Referer": PORTAL_SOT_BASE + "/cabinet",
                    **({"User-Agent": self.tokens["user_agent"]} if self.tokens.get("user_agent") else {}),
                },
                timeout=30,
            )
            js = r.json()
        except Exception as e:
            self.log(f"⚠ portal-sot.kz refresh не удался: {e}")
            return False
        data = js.get("data") or {}
        if js.get("status") != 0 or not data.get("access_token"):
            self.log(f"⚠ portal-sot.kz refresh отклонён: {js.get('message') or js}")
            return False
        self.tokens["access_token"] = data["access_token"]
        if data.get("refresh_token"):
            self.tokens["refresh_token"] = data["refresh_token"]
        self._save_tokens()
        self.log("🔄 portal-sot.kz: access_token обновлён через refresh (без браузера)")
        return True

    def _login_via_browser(self, reason="refresh_token нет/истёк"):
        if self._browser_logins >= MAX_BROWSER_LOGINS_PER_RUN:
            raise PortalBlocked(
                f"portal-sot.kz не принимает токен даже после {self._browser_logins} входов по ЭЦП "
                f"за этот запуск — прогон остановлен. Проверьте вход на портал вручную."
            )
        self._browser_logins += 1
        self.log(f"🔐 portal-sot.kz: {reason} — вход через браузер + ЭЦП (браузер закроется сразу после входа)")
        got = self.browser_login()
        if not got or not got.get("access_token"):
            raise RuntimeError("Вход на portal-sot.kz через браузер не вернул access_token")
        self.tokens = {
            "access_token": got["access_token"],
            "refresh_token": got.get("refresh_token") or "",
            "cookies": got.get("cookies") or [],
            "user_agent": got.get("user_agent") or "",
            "saved_at": time.time(),
        }
        self._save_tokens()

    def renew(self, full=False):
        """Токен не принят (401 / пустой 200): сначала refresh, потом браузер.
        full=True — сразу вход по ЭЦП (refresh уже пробовали, и он не помог)."""
        if full:
            self._login_via_browser("токен после refresh всё равно не принят (401)")
        elif not self._refresh():
            self._login_via_browser()
        self._build_session()

    # ---------- темп ----------
    def _throttle(self):
        if self._count >= MAX_REQUESTS_PER_RUN:
            raise PortalBlocked(f"Достигнут лимит {MAX_REQUESTS_PER_RUN} запросов за запуск — остановка")
        self._count += 1
        if self._count > 1 and self._count % LONG_PAUSE_EVERY == 0:
            pause = random.uniform(*LONG_PAUSE_SECONDS)
            self.log(f"☕ {self._count} запросов к portal-sot.kz — передышка {pause:.0f} с")
            time.sleep(pause)
        wait = self._last_request_at + random.uniform(*MIN_GAP_SECONDS) - time.time()
        if wait > 0:
            time.sleep(wait)
        self._last_request_at = time.time()

    # ---------- запросы ----------
    def get(self, url, **kwargs):
        """GET с темпом, авто-обновлением токена и отступом. Совместим с requests.Session.get."""
        return self.request("GET", url, **kwargs)

    def post(self, url, **kwargs):
        """POST с теми же правилами, что и get()."""
        return self.request("POST", url, **kwargs)

    def request(self, method, url, backoff=BACKOFF_429_5XX, **kwargs):
        """backoff — ожидания при 429/5xx. Свой короткий список нужен там, где
        5xx — обычный ответ на конкретный объект, а не признак перегрузки."""
        if url.startswith("/"):
            url = PORTAL_SOT_BASE + url
        kwargs.setdefault("timeout", (10, 60))
        if _jwt_exp(self.tokens.get("access_token", "")) < time.time() + 30:
            self.renew()
        renew_stage = 0   # 0 — ещё не обновляли, 1 — был refresh, 2 — был вход по ЭЦП
        backoff = list(backoff)
        while True:
            self._throttle()
            r = self.session.request(method, url, **kwargs)

            if r.status_code == 401:
                if renew_stage == 0:
                    renew_stage = 1
                    self.renew()
                    continue
                if renew_stage == 1:
                    renew_stage = 2
                    self.renew(full=True)
                    continue
                raise PortalBlocked(
                    "portal-sot.kz отвечает 401 даже сразу после свежего входа по ЭЦП — "
                    "прогон остановлен. Проверьте вход на портал вручную."
                )

            if is_waf_block(r):
                raise PortalBlocked(BLOCK_MESSAGE)

            if r.status_code == 429 or r.status_code >= 500:
                if not backoff:
                    if r.status_code == 429:
                        raise PortalBlocked("portal-sot.kz продолжает отвечать 429 — прогон остановлен")
                    return r
                wait = backoff.pop(0)
                ra = r.headers.get("Retry-After", "")
                if ra.isdigit():
                    wait = max(wait, int(ra))
                self.log(f"⏸ portal-sot.kz ответил {r.status_code} — жду {wait} с")
                time.sleep(wait)
                continue

            return r
