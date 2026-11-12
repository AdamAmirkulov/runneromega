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

try:
    from curl_cffi import requests as cf_requests
    from curl_cffi import CurlMime
except ImportError:
    cf_requests = None
    CurlMime = None

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

# HTTP-транспорты. WAF portal-sot.kz то режет обычный requests и пропускает
# curl_cffi с отпечатком Firefox, то наоборот (2026-10-06 и 2026-10-16). Поэтому
# транспорт, через который прошёл вход, запоминается в токенах, а на 403 от WAF
# один раз пробуем второй и только потом считаем это блокировкой.
CF_IMPERSONATE = "firefox135"       # тот же профиль, что portal_sot_auth.IMPERSONATE
TRANSPORTS = ("requests", "curl_cffi") if cf_requests is not None else ("requests",)


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


TRACE_DIR = Path(__file__).resolve().parent.parent / "data" / "portal_trace"


def _mask_token(v) -> str:
    """Токен в трассировке: хвост + exp, целиком не пишем."""
    v = str(v or "")
    bearer = v.lower().startswith("bearer ")
    tok = v.split(" ", 1)[1] if bearer else v
    if len(tok) < 40:
        return v
    exp = _jwt_exp(tok)
    left = int(exp - time.time()) if exp else None
    return f"{'Bearer ' if bearer else ''}...{tok[-10:]} (exp={exp:.0f}, осталось {left} с)"


class _Tracer:
    """Полная трассировка запросов к portal-sot.kz: одна JSON-строка на запрос
    (заголовки и тело запроса, статус, заголовки и ПОЛНОЕ тело ответа, время,
    транспорт, исключение) + события (refresh, вход по ЭЦП, смена транспорта...).
    Файл: data/portal_trace/<компания>_<дата-время>.jsonl."""

    def __init__(self, path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.n = 0

    def event(self, kind, **data):
        self.n += 1
        rec = {"n": self.n, "ts": time.strftime("%Y-%m-%d %H:%M:%S"),
               "t": round(time.time(), 3), "kind": kind, **data}
        try:
            with open(self.path, "a", encoding="utf-8") as fh:
                fh.write(json.dumps(rec, ensure_ascii=False, default=str) + "\n")
        except Exception:
            pass

    @staticmethod
    def _hdrs(h):
        out = {}
        try:
            items = dict(h or {}).items()
        except Exception:
            return {}
        for k, v in items:
            out[str(k)] = _mask_token(v) if str(k).lower() == "authorization" else str(v)
        return out

    def request(self, transport, method, url, req_headers, kwargs, r=None, exc=None, elapsed=0.0):
        body = kwargs.get("json", kwargs.get("data"))
        if isinstance(body, dict) and "refreshToken" in body:
            body = {**body, "refreshToken": _mask_token(body["refreshToken"])}
        rec = {"transport": transport, "method": method, "url": url,
               "req_headers": self._hdrs(req_headers), "req_body": body,
               "elapsed_s": round(elapsed, 3)}
        if r is not None:
            try:
                text = r.text or ""
            except Exception as e:
                text = f"<не прочитать тело: {e}>"
            try:
                cookies = {k: str(v) for k, v in dict(r.cookies).items()}
            except Exception:
                cookies = {}
            rec.update({
                "status": r.status_code, "reason": getattr(r, "reason", ""),
                "resp_url": str(getattr(r, "url", "")),
                "resp_headers": {str(k): str(v) for k, v in dict(r.headers).items()},
                "resp_cookies": cookies,
                "body_len": len(text), "body": text,
            })
        if exc is not None:
            rec.update({"exc_type": type(exc).__name__, "exc": repr(exc)})
        self.event("http", **rec)


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

    def __init__(self, company_id: str, browser_login, log=print, trace=False):
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
        self.tracer = None
        if trace or os.environ.get("PORTAL_SOT_TRACE"):
            self.tracer = _Tracer(TRACE_DIR / f"{self.company_id}_{time.strftime('%Y%m%d_%H%M%S')}.jsonl")
            self.log(f"🧾 portal-sot.kz: полная трассировка запросов → {self.tracer.path}")

    def trace(self, kind, **data):
        if self.tracer:
            self.tracer.event(kind, **data)

    def _token_state(self):
        now = time.time()
        a = _jwt_exp(self.tokens.get("access_token", ""))
        rt = _jwt_exp(self.tokens.get("refresh_token", ""))
        return {"transport": self.transport,
                "access_left_s": int(a - now) if a else None,
                "refresh_left_s": int(rt - now) if rt else None,
                "cookies": [c.get("name") for c in self.tokens.get("cookies") or []],
                "user_agent": self.tokens.get("user_agent", "")}

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

    # ---------- транспорт ----------
    @property
    def transport(self):
        t = self.tokens.get("transport")
        return t if t in TRANSPORTS else TRANSPORTS[0]

    def _new_http(self):
        if self.transport == "curl_cffi":
            return cf_requests.Session(impersonate=CF_IMPERSONATE)
        s = requests.Session()
        s.headers["Accept-Encoding"] = "gzip, deflate"
        return s

    def _switch_transport(self, reason="403 от WAF") -> bool:
        """403 от WAF / обрыв соединения: перейти на второй транспорт.
        False — переключаться некуда."""
        others = [t for t in TRANSPORTS if t != self.transport]
        if not others:
            return False
        self.log(f"⚠ portal-sot.kz: {reason} на {self.transport} — пробую {others[0]}")
        self.trace("switch_transport", reason=reason, frm=self.transport, to=others[0])
        self.tokens["transport"] = others[0]
        self._save_tokens()
        if self.tokens.get("access_token"):
            self._build_session()
        return True

    def _build_session(self):
        s = self._new_http()
        s.headers.update({
            "Authorization": f"Bearer {self.tokens['access_token']}",
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "ru",
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
        self.trace("session_start", **self._token_state())
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
        switched = False
        try:
            while True:
                self._throttle()
                rurl = PORTAL_SOT_BASE + "/api/public/auth/refresh"
                rkw = dict(
                    json={"refreshToken": self.tokens.get("refresh_token", "")},
                    headers={
                        "Accept": "*/*", "Accept-Language": "ru",
                        "Origin": PORTAL_SOT_BASE, "Referer": PORTAL_SOT_BASE + "/cabinet",
                        **({"User-Agent": self.tokens["user_agent"]} if self.tokens.get("user_agent") else {}),
                    },
                    timeout=30,
                )
                t0 = time.time()
                try:
                    r = self._new_http().post(rurl, **rkw)
                except Exception as e:
                    if self.tracer:
                        self.tracer.request(self.transport, "POST", rurl, rkw["headers"], rkw,
                                            exc=e, elapsed=time.time() - t0)
                    raise
                if self.tracer:
                    self.tracer.request(self.transport, "POST", rurl, rkw["headers"], rkw,
                                        r=r, elapsed=time.time() - t0)
                if is_waf_block(r) and not switched and self._switch_transport():
                    switched = True
                    continue
                break
            js = r.json()
        except Exception as e:
            self.log(f"⚠ portal-sot.kz refresh не удался: {e}")
            self.trace("refresh_failed", exc=repr(e))
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
        self.trace("browser_login_start", reason=reason, n=self._browser_logins, **self._token_state())
        self.log(f"🔐 portal-sot.kz: {reason} — вход через браузер + ЭЦП (браузер закроется сразу после входа)")
        got = self.browser_login()
        if not got or not got.get("access_token"):
            raise RuntimeError("Вход на portal-sot.kz через браузер не вернул access_token")
        self.tokens = {
            "access_token": got["access_token"],
            "refresh_token": got.get("refresh_token") or "",
            "cookies": got.get("cookies") or [],
            "user_agent": got.get("user_agent") or "",
            # вход без браузера сообщает, какой транспорт прошёл WAF; после
            # входа браузером оставляем тот, что работал раньше
            "transport": got.get("transport") or self.tokens.get("transport") or "",
            "saved_at": time.time(),
        }
        self._save_tokens()
        self.trace("browser_login_done", **self._token_state())

    def renew(self, full=False):
        """Токен не принят (401 / пустой 200): сначала refresh, потом браузер.
        full=True — сразу вход по ЭЦП (refresh уже пробовали, и он не помог)."""
        self.trace("renew", full=full, **self._token_state())
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
            self.trace("long_pause", count=self._count, pause_s=round(pause, 1))
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

    def _send(self, method, url, kwargs):
        t0 = time.time()
        try:
            r = self._send_raw(method, url, kwargs)
        except Exception as e:
            if self.tracer:
                self.tracer.request(self.transport, method, url, getattr(self.session, "headers", {}),
                                    kwargs, exc=e, elapsed=time.time() - t0)
            raise
        if self.tracer:
            try:
                hdrs = dict(r.request.headers)
            except Exception:
                hdrs = getattr(self.session, "headers", {})
            self.tracer.request(self.transport, method, url, hdrs, kwargs, r=r, elapsed=time.time() - t0)
        return r

    def _send_raw(self, method, url, kwargs):
        """Один запрос текущей сессией. curl_cffi не принимает files= (как requests) —
        файлы и поля формы уходят через его multipart (CurlMime), собранный заново
        на каждую попытку."""
        if self.transport != "curl_cffi" or not kwargs.get("files"):
            return self.session.request(method, url, **kwargs)
        kw = dict(kwargs)
        files = kw.pop("files")
        form = kw.pop("data", None) or {}
        mime = CurlMime()
        try:
            for name, value in form.items():
                mime.addpart(name=name, data=str(value).encode("utf-8"))
            for name, spec in files.items():
                filename, content, content_type = (tuple(spec) + (None, None))[:3]
                if hasattr(content, "read"):
                    # файл могли уже прочитать прошлой попыткой — перематываем
                    if hasattr(content, "seek"):
                        content.seek(0)
                    content = content.read()
                if isinstance(content, str):
                    content = content.encode("utf-8")
                mime.addpart(name=name, filename=filename, content_type=content_type, data=content)
            return self.session.request(method, url, multipart=mime, **kw)
        finally:
            mime.close()

    def request(self, method, url, backoff=BACKOFF_429_5XX, **kwargs):
        """backoff — ожидания при 429/5xx. Свой короткий список нужен там, где
        5xx — обычный ответ на конкретный объект, а не признак перегрузки."""
        if url.startswith("/"):
            url = PORTAL_SOT_BASE + url
        kwargs.setdefault("timeout", (10, 60))
        if _jwt_exp(self.tokens.get("access_token", "")) < time.time() + 30:
            self.renew()
        renew_stage = 0   # 0 — ещё не обновляли, 1 — был refresh, 2 — был вход по ЭЦП
        switched = False  # второй транспорт после 403 от WAF уже пробовали
        backoff = list(backoff)
        while True:
            self._throttle()
            try:
                r = self._send(method, url, kwargs)
            except Exception as e:
                # Сервер рвёт соединение одному транспорту (curl: (56) Connection
                # closed abruptly у curl_cffi, 09.11 InvestWay) — как и на 403,
                # один раз пробуем второй.
                if not switched and "timeout" not in type(e).__name__.lower()                         and self._switch_transport(f"обрыв соединения ({type(e).__name__})"):
                    switched = True
                    continue
                raise

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
                if not switched and self._switch_transport():
                    switched = True
                    continue
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
                self.trace("backoff", status=r.status_code, wait_s=wait)
                time.sleep(wait)
                continue

            return r
