# -*- coding: utf-8 -*-
"""
Вход на portal-sot.kz через ЭЦП (NCALayer) — общая механика по компаниям.

Логика 1:1 из poiskvsk.py (_portal_init_driver / _portal_ensure_ncalayer /
_portal_login), вынесена сюда, чтобы подача исков (podacha_portal_sot.py)
входила так же надёжно. poiskvsk.py пока держит свою копию — не трогаем
работающий скрипт.

    from portal_sot_login import PortalSotLogin
    pl = PortalSotLogin(company_id, cfg, log=log, dump_dir=out_dir)
    driver = pl.init_driver()
    access, refresh = pl.login(driver)            # живая сессия → без ЭЦП
    access, refresh = pl.login(driver, force_fresh=True)
"""

import base64
import json
import os
import socket
import subprocess
import time
from datetime import datetime

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.firefox.service import Service as FirefoxService
from webdriver_manager.firefox import GeckoDriverManager

BASE_URL = "https://portal-sot.kz"
NCALAYER_WS_PORT = 13579

LOGIN_BTN_XPATH = (
    "//*[self::button or self::a or @role='button']"
    "[contains(normalize-space(.),'Войти') or contains(normalize-space(.),'Вход')"
    " or contains(normalize-space(.),'Кіру')]"
)


def _port_open(port: int) -> bool:
    try:
        with socket.create_connection(("127.0.0.1", port), timeout=1):
            return True
    except OSError:
        return False


def jwt_claims(token: str) -> dict:
    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        return json.loads(base64.urlsafe_b64decode(payload))
    except Exception:
        return {}


class PortalSotLogin:
    def __init__(self, company_id: str, cfg: dict, log=print, dump_dir=None):
        """cfg — PORTAL_SOT_BY_COMPANY[company_id] из config.py."""
        self.company_id = str(company_id)
        self.cfg = cfg
        self.log = log
        self.dump_dir = dump_dir or os.getcwd()
        self.profile = cfg.get("firefox_profile")
        if not self.profile:
            raise RuntimeError(
                f"Для компании {company_id} не задан firefox_profile в PORTAL_SOT_BY_COMPANY (scripts/config.py)"
            )
        self.ncalayer_path = cfg.get("ncalayer_path") or os.path.expandvars(
            r"%LOCALAPPDATA%\Programs\NCALayer\NCALayer.exe"
        )

    # ---------- Firefox ----------

    def _release_stale_profile(self):
        """Как в poiskvsk.py: если прошлый запуск не закрыл Firefox, профиль
        остаётся залоченным (.parentlock) и новый Firefox на нём не стартует."""
        try:
            subprocess.run(["taskkill", "/F", "/IM", "firefox.exe"],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            pass
        time.sleep(1)
        for lock_name in (".parentlock", "lock"):
            try:
                lock_path = os.path.join(self.profile, lock_name)
                if os.path.exists(lock_path):
                    os.remove(lock_path)
            except Exception:
                pass

    def init_driver(self):
        # Постоянный профиль компании (не копия!), где уже нажато «Разрешить»
        # для portal-sot.kz → NCALayer и принят self-signed сертификат NCALayer.
        os.makedirs(self.profile, exist_ok=True)
        opts = FirefoxOptions()
        opts.add_argument("-profile")
        opts.add_argument(self.profile)
        opts.set_preference("dom.disable_beforeunload", True)
        try:
            drv = webdriver.Firefox(service=FirefoxService(GeckoDriverManager().install()), options=opts)
        except Exception as e:
            self.log(f"Прямой запуск Firefox не удался ({str(e).splitlines()[0]}); снимаю лок профиля и повторяю")
            self._release_stale_profile()
            drv = webdriver.Firefox(service=FirefoxService(GeckoDriverManager().install()), options=opts)
        caps = drv.capabilities or {}
        self.log(f"Браузер: {caps.get('browserName')} {caps.get('browserVersion')}, профиль: {self.profile}")
        drv.set_page_load_timeout(120)
        drv.maximize_window()
        return drv

    # ---------- NCALayer ----------

    def ensure_ncalayer(self, timeout=70):
        """NCALayer обязан работать с сертификатом ЭТОЙ компании ДО клика «Войти»."""
        try:
            from config import ensure_ncalayer_cert
            ensure_ncalayer_cert(self.company_id)  # при смене сертификата перезапускает NCALayer
        except Exception as e:
            self.log(f"ensure_ncalayer_cert: {e}")

        if _port_open(NCALAYER_WS_PORT):
            self.log("NCALayer уже запущен (порт 13579)")
            return
        if os.path.isfile(self.ncalayer_path):
            self.log(f"Запускаю NCALayer: {self.ncalayer_path}")
            try:
                subprocess.Popen([self.ncalayer_path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except Exception as e:
                self.log(f"Не смог запустить NCALayer ({e}) — запустите его вручную")
        else:
            self.log(f"NCALayer.exe не найден: {self.ncalayer_path}. Задайте "
                     f"PORTAL_SOT_BY_COMPANY['{self.company_id}']['ncalayer_path'] или запустите вручную.")
        deadline = time.time() + timeout
        while time.time() < deadline:
            if _port_open(NCALAYER_WS_PORT):
                self.log("✅ NCALayer готов (порт 13579)")
                time.sleep(3)  # даём модулям подгрузиться
                return
            time.sleep(1)
        raise RuntimeError(f"NCALayer не поднялся за {timeout}s (порт {NCALAYER_WS_PORT} закрыт).")

    @staticmethod
    def _wait_ncalayer(timeout=60):
        from pywinauto import Desktop
        desktop = Desktop(backend="uia")
        t0 = time.time()
        while time.time() - t0 < timeout:
            try:
                nca = desktop.window(title="NCALayer")
                if nca.exists() and nca.is_visible():
                    return nca
            except Exception:
                pass
            time.sleep(0.4)
        raise RuntimeError("NCALayer не появился")

    @staticmethod
    def _wait_nca_button(title, timeout=60):
        from pywinauto import Desktop
        t0 = time.time()
        while time.time() - t0 < timeout:
            try:
                nca = Desktop(backend="uia").window(title="NCALayer")
                if nca.exists() and nca.is_visible():
                    btn = nca.child_window(title=title, control_type="Button")
                    if btn.exists():
                        return nca, btn
            except Exception:
                pass
            time.sleep(0.4)
        raise RuntimeError(f"Кнопка NCALayer «{title}» не появилась")

    # ---------- портал ----------

    def dump_page(self, driver, tag: str):
        try:
            os.makedirs(self.dump_dir, exist_ok=True)
            stem = os.path.join(self.dump_dir, f"portal_{tag}_{time.strftime('%H%M%S')}")
            driver.save_screenshot(stem + ".png")
            with open(stem + ".html", "w", encoding="utf-8") as f:
                f.write(driver.page_source or "")
            self.log(f"Снимок страницы портала: {stem}.png / .html")
        except Exception as e:
            self.log(f"Не смог сохранить снимок страницы портала: {e}")

    def _click_login(self, driver):
        # WebDriverWait при таймауте даёт пустой «Message:» со стектрейсом
        # NoSuchElement — ловим и объясняем сами.
        for attempt in range(1, 3):
            try:
                btn = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, LOGIN_BTN_XPATH)))
                break
            except TimeoutException:
                self.log(f"Кнопка «Войти» на {driver.current_url} не найдена за 30s "
                         f"(попытка {attempt}, title='{driver.title}')")
                self.dump_page(driver, f"no_login_btn_{attempt}")
                if attempt == 2:
                    raise RuntimeError(
                        f"На portal-sot.kz нет кнопки «Войти» (URL: {driver.current_url}) — см. снимок в out/."
                    )
                # Профиль «полузалогинен» (кнопки нет, годного токена тоже) — сброс.
                self.clear_session(driver)
        try:
            btn.click()
        except Exception:
            driver.execute_script("arguments[0].click();", btn)

    @staticmethod
    def clear_session(driver):
        try:
            driver.get(BASE_URL + "/")
            time.sleep(1)
            driver.execute_script("try{localStorage.clear()}catch(e){} try{sessionStorage.clear()}catch(e){}")
            driver.delete_all_cookies()
        except Exception:
            pass
        driver.get(BASE_URL + "/")
        time.sleep(3)

    @staticmethod
    def read_tokens(driver):
        access = driver.execute_script("return localStorage.getItem('access_token');")
        refresh = driver.execute_script("return localStorage.getItem('refresh_token');")
        return access, refresh

    def login(self, driver, force_fresh=False):
        """Возвращает (access_token, refresh_token). Живая сессия с неистёкшим
        JWT → без ЭЦП; иначе полный вход: «Войти» → NCALayer (пароль ЭЦП,
        «Открыть», «Подписать») → пароль портала → «Войти» → /cabinet."""
        # NCALayer обязан работать ДО клика «Войти».
        self.ensure_ncalayer()

        if force_fresh:
            self.clear_session(driver)
        else:
            driver.get(BASE_URL + "/cabinet")
            time.sleep(3)
            access = driver.execute_script("return localStorage.getItem('access_token');")
            has_pwd_field = driver.execute_script("return !!document.querySelector(\"input[type='password']\");")
            if access and "/cabinet" in driver.current_url.lower() and not has_pwd_field:
                exp = jwt_claims(access).get("exp", 0) or 0
                if exp and exp > time.time() + 30:
                    self.log(f"✅ Сессия portal-sot.kz жива (токен действует до "
                             f"{datetime.fromtimestamp(exp):%H:%M:%S})")
                    return self.read_tokens(driver)
                self.log("access_token в localStorage просрочен/битый — полный вход через ЭЦП")
            driver.get(BASE_URL)
            time.sleep(2)

        nca = None
        for attempt in range(1, 3):
            self._click_login(driver)
            self.log(f"Нажал «Войти» (попытка {attempt}), жду NCALayer…")
            try:
                nca = self._wait_ncalayer(45)
                break
            except RuntimeError:
                self.log("Окно NCALayer не появилось — проверяю службу и пробую ещё раз")
                self.ensure_ncalayer()
                time.sleep(2)
        if nca is None:
            raise RuntimeError(
                "Окно подписи NCALayer так и не появилось после клика «Войти». Убедитесь, что NCALayer "
                "запущен и в Firefox-профиле компании для portal-sot.kz нажато «Разрешить»."
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
        pwd.type_keys(self.cfg["eds_password"], with_spaces=True, pause=0.03)
        time.sleep(0.3)
        open_btn.click_input()

        nca, sign_btn = self._wait_nca_button("Подписать", 60)
        nca.set_focus()
        time.sleep(0.3)
        sign_btn.click_input()
        self.log("✅ ЭЦП подписана")

        try:
            portal_pwd = WebDriverWait(driver, 60).until(
                EC.visibility_of_element_located((By.XPATH, "//input[@type='password']"))
            )
        except TimeoutException:
            self.dump_page(driver, "no_portal_password")
            raise RuntimeError("После подписи ЭЦП не появилось поле пароля портала — см. снимок в out/.")
        portal_pwd.clear()
        portal_pwd.send_keys(self.cfg["portal_password"])

        final_btn = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((
            By.XPATH, "//button[contains(normalize-space(.),'Войти')]",
        )))
        final_btn.click()

        try:
            WebDriverWait(driver, 60).until(lambda d: "/cabinet" in d.current_url.lower())
        except TimeoutException:
            self.dump_page(driver, "no_cabinet")
            raise RuntimeError("Нет перехода в кабинет после пароля портала — см. снимок в out/.")

        access, refresh = self.read_tokens(driver)
        if not access:
            raise RuntimeError("access_token не найден после входа")
        self.log("✅ Вход на portal-sot.kz выполнен, access_token получен")
        return access, refresh
