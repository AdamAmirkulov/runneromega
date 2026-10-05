# -*- coding: utf-8 -*-
"""
Вход на portal-sot.kz через ЭЦП — БЕЗ браузера (без Selenium/Firefox).

Почему так: с 2026-09 portal-sot.kz на уровне nginx/WAF блокирует 403-м
ЛЮБОЙ настоящий браузерный движок (проверено: и Chrome, и Firefox, и вручную,
и через Selenium — все получают plain "403 Forbidden" от nginx). При этом
обычные HTTP-инструменты (curl, requests с TLS-отпечатком браузера через
curl_cffi) проходят нормально — похоже на фильтрацию по TLS-фингерпринту
(JA3), которая ловит реальные браузерные движки. Подробности и как это было
установлено — см. историю чата/git log на дату появления этого файла.

Протокол восстановлен из HAR настоящего успешного входа (не угадан):
  1. GET  /api/public/auth/xml   -> XML-челлендж <sud><authentication .../></sud>
  2. WSS  127.0.0.1:13579 (NCALayer, тот самый протокол kz.gov.pki.knca.basics,
     Origin: https://portal-sot.kz) -> просим подписать XML (format="xml").
     NCALayer поднимает СВОЁ нативное окно (выбор/пароль сертификата +
     «Открыть» + «Подписать») — это автоматизируется через pywinauto, как и
     раньше при Selenium-варианте, только теперь без браузера вообще.
  3. POST /api/public/auth/eds   {"signedXml": base64(подписанный XML)}
     -> {"status":0,"data":{"user":..., "access_token":..., "refresh_token":...}}
Портального пароля (второй фактор) в этом потоке НЕ было — одной подписи ЭЦП
достаточно (наблюдение из реального входа, а не документация).

HTTP-часть идёт через curl_cffi (импersonate="firefox135") — обходит
TLS-фингерпринт-блокировку. WebSocket к NCALayer — локальный (127.0.0.1),
через блокировку portal-sot.kz вообще не идёт, поэтому обычный
websocket-client тут ни при чём не мешает.
"""
import base64
import json
import os
import socket
import ssl
import subprocess
import time

from curl_cffi import requests as cf_requests
import websocket

PORTAL_SOT_BASE = "https://portal-sot.kz"
NCALAYER_WS_URL = "wss://127.0.0.1:13579/"
NCALAYER_WS_PORT = 13579

# Профиль TLS-отпечатка для curl_cffi — подтверждено тестами, что "firefox135"
# и "firefox133" проходят WAF portal-sot.kz, а "chrome131" — нет.
IMPERSONATE = "firefox135"


def _default_log(msg):
    print(msg, flush=True)


def _ncalayer_ws_open() -> bool:
    try:
        with socket.create_connection(("127.0.0.1", NCALAYER_WS_PORT), timeout=1):
            return True
    except OSError:
        return False


def ensure_ncalayer_running(company_id: str, ncalayer_path: str, timeout: int = 70, log=_default_log):
    """NCALayer должен быть запущен и настроен на сертификат нужной компании
    ДО отправки запроса на подпись. Переиспользует config.ensure_ncalayer_cert
    (переключает recentPath + перезапускает NCALayer, если сертификат не тот)."""
    try:
        from config import ensure_ncalayer_cert
        ensure_ncalayer_cert(str(company_id))
    except Exception as e:
        log(f"⚠ ensure_ncalayer_cert: {e}")

    if _ncalayer_ws_open():
        log("NCALayer уже запущен (порт 13579)")
        return

    if os.path.isfile(ncalayer_path):
        log(f"Запускаю NCALayer: {ncalayer_path}")
        try:
            subprocess.Popen([ncalayer_path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception as e:
            log(f"⚠ Не смог запустить NCALayer ({e}) — запустите его вручную")
    else:
        log(f"⚠ NCALayer.exe не найден: {ncalayer_path}")

    deadline = time.time() + timeout
    while time.time() < deadline:
        if _ncalayer_ws_open():
            log("✅ NCALayer готов (порт 13579)")
            time.sleep(3)  # даём модулям подгрузиться
            return
        time.sleep(1)
    raise RuntimeError(
        f"NCALayer не поднялся за {timeout}s (порт {NCALAYER_WS_PORT} закрыт). "
        f"Запустите NCALayer вручную и повторите."
    )


def _handle_ncalayer_dialog(eds_password: str, log=_default_log, timeout: int = 60):
    """Нативное окно NCALayer (не веб-страница!) — ввод пароля ЭЦП, «Открыть»,
    затем «Подписать». Портировано 1:1 из прежнего Selenium-варианта
    (scripts/poiskvsk.py _portal_login) — само окно и его поля от способа
    инициации подписи (браузер vs. прямой WS) не зависят."""
    from pywinauto import Desktop

    t0 = time.time()
    nca = None
    while time.time() - t0 < timeout:
        try:
            cand = Desktop(backend="uia").window(title="NCALayer")
            if cand.exists() and cand.is_visible():
                nca = cand
                break
        except Exception:
            pass
        time.sleep(0.4)
    if nca is None:
        raise RuntimeError(
            "Окно NCALayer не появилось после запроса на подпись. Убедитесь, "
            "что для portal-sot.kz в NCALayer нажато «Разрешить»."
        )
    nca.set_focus()

    pwd = nca.child_window(title="Пароль", control_type="Edit")
    if not pwd.exists():
        pwd = nca.child_window(auto_id="JavaFX39", control_type="Edit")
    open_btn = nca.child_window(title="Открыть", control_type="Button")
    if not open_btn.exists():
        open_btn = nca.child_window(auto_id="JavaFX47", control_type="Button")

    pwd.click_input()
    time.sleep(0.2)
    pwd.type_keys("^a{BACKSPACE}")
    pwd.type_keys(eds_password, with_spaces=True, pause=0.03)
    time.sleep(0.3)
    open_btn.click_input()
    log("NCALayer: пароль ЭЦП введён, файл открыт")

    t0 = time.time()
    nca2 = sign_btn = None
    while time.time() - t0 < timeout:
        try:
            cand = Desktop(backend="uia").window(title="NCALayer")
            if cand.exists() and cand.is_visible():
                btn = cand.child_window(title="Подписать", control_type="Button")
                if btn.exists():
                    nca2, sign_btn = cand, btn
                    break
        except Exception:
            pass
        time.sleep(0.4)
    if sign_btn is None:
        raise RuntimeError("Кнопка «Подписать» не появилась в окне NCALayer")
    nca2.set_focus()
    time.sleep(0.3)
    sign_btn.click_input()
    log("NCALayer: нажата «Подписать»")


def _ncalayer_sign_xml(xml_data: str, eds_password: str, log=_default_log) -> str:
    """Протокол kz.gov.pki.knca.basics/sign (format=xml), как в реальном
    браузерном запросе (восстановлено из HAR, не из документации)."""
    ws = websocket.create_connection(
        NCALAYER_WS_URL,
        sslopt={"cert_reqs": ssl.CERT_NONE},
        header=["Origin: https://portal-sot.kz"],
        timeout=15,
    )
    try:
        greeting = ws.recv()
        log(f"NCALayer greeting: {greeting}")

        req = {
            "module": "kz.gov.pki.knca.basics",
            "method": "sign",
            "args": {
                "format": "xml",
                "data": xml_data,
                "signingParams": {
                    "decode": "false", "encapsulate": "true",
                    "digested": "false", "tsaProfile": None,
                },
                "signerParams": {
                    "extKeyUsageOids": [], "iin": "", "bin": "",
                    "serialNumber": "", "chain": None,
                },
                "locale": "ru",
            },
        }
        ws.send(json.dumps(req))

        # ws.recv() ниже блокируется, пока в нативном окне NCALayer не нажмут
        # «Подписать» — поэтому сначала обрабатываем окно, потом читаем ответ.
        _handle_ncalayer_dialog(eds_password, log=log)

        raw = ws.recv()
        try:
            data = json.loads(raw)
        except Exception as e:
            raise RuntimeError(f"NCALayer: ответ не JSON: {raw[:500]!r}") from e

        if not data.get("status"):
            raise RuntimeError(f"NCALayer вернул ошибку: {raw[:800]!r}")

        result = (data.get("body") or {}).get("result")
        if not result:
            raise RuntimeError(f"NCALayer: не удалось получить подписанный XML: {raw[:800]!r}")
        return result[0] if isinstance(result, list) else result
    finally:
        try:
            ws.close()
        except Exception:
            pass


def login(company_id: str, eds_password: str, ncalayer_path: str, log=_default_log):
    """Полный вход на portal-sot.kz через ЭЦП без браузера.
    Возвращает (access_token, refresh_token, curl_cffi.requests.Session)."""
    ensure_ncalayer_running(company_id, ncalayer_path, log=log)

    session = cf_requests.Session(impersonate=IMPERSONATE)
    session.headers.update({
        "Referer": PORTAL_SOT_BASE + "/",
        "Origin": PORTAL_SOT_BASE,
        "Accept-Language": "ru",
    })

    log("🔐 Запрашиваю XML-челлендж для входа...")
    r = session.get(PORTAL_SOT_BASE + "/api/public/auth/xml", timeout=30)
    r.raise_for_status()
    xml_challenge = r.text
    log(f"Получен челлендж: {xml_challenge.strip()[:150]}")

    log("✍️  Прошу NCALayer подписать челлендж (ждите нативное окно ЭЦП)...")
    signed_xml = _ncalayer_sign_xml(xml_challenge, eds_password, log=log)

    signed_b64 = base64.b64encode(signed_xml.encode("utf-8")).decode("ascii")
    log("📤 Отправляю подписанный XML на portal-sot.kz...")
    r2 = session.post(
        PORTAL_SOT_BASE + "/api/public/auth/eds",
        json={"signedXml": signed_b64},
        timeout=30,
    )
    r2.raise_for_status()
    data = r2.json()
    if data.get("status") != 0:
        raise RuntimeError(f"portal-sot.kz /auth/eds отклонил вход: {data.get('message') or data}")

    tok = data["data"]
    access_token = tok["access_token"]
    refresh_token = tok["refresh_token"]
    user = tok.get("user") or {}
    log(f"✅ Вход выполнен без браузера: {user.get('fio', '')} ({user.get('organizationName', '')})")

    session.headers["Authorization"] = f"Bearer {access_token}"
    return access_token, refresh_token, session


def make_session(access_token: str) -> "cf_requests.Session":
    """Новая curl_cffi-сессия с уже готовым access_token (когда логиниться
    заново не нужно — например, повторное использование сохранённого токена)."""
    s = cf_requests.Session(impersonate=IMPERSONATE)
    s.headers.update({
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "ru",
        "Origin": PORTAL_SOT_BASE,
        "Referer": PORTAL_SOT_BASE + "/",
    })
    return s
