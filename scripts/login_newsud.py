# ============================================================
# PORTAL-SOT.KZ — ПОЛНЫЙ АВТОВХОД ЧЕРЕЗ ЭЦП
# ============================================================
#
# Автономный скрипт входа (двойной клик / python scripts/login_newsud.py
# [company_id]). Держит Chrome открытым после входа. Рабочая логика того же
# входа продублирована внутри scripts/poiskvsk.py (блок 4, функции _portal_*),
# который для поиска адреса ей и пользуется — этот файл на poiskvsk не влияет.
#
# Реквизиты (пароль ЭЦП / пароль портала / профиль Chrome) НЕ хранятся
# здесь в коде — берутся из PORTAL_SOT_BY_COMPANY в scripts/config.py
# (тот же гитигнорнутый конфиг, что и у остальных скриптов репозитория).

import os
import sys
import time

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from pywinauto import Desktop

_BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _BASE not in sys.path:
    sys.path.insert(0, _BASE)


# ============================================================
# НАСТРОЙКИ
# ============================================================

PORTAL_URL = "https://portal-sot.kz"

_COMPANY_ID = (sys.argv[1] if len(sys.argv) > 1 else "1").strip()

from config import PORTAL_SOT_BY_COMPANY  # noqa: E402

_cfg = PORTAL_SOT_BY_COMPANY.get(_COMPANY_ID)
if not _cfg:
    raise SystemExit(
        f"Для компании {_COMPANY_ID} нет записи в PORTAL_SOT_BY_COMPANY "
        f"(scripts/config.py)."
    )

# Пароль от файла ЭЦП
EDS_PASSWORD = _cfg["eds_password"]

# Пароль от portal-sot.kz,
# который появляется после подписания ЭЦП
PORTAL_PASSWORD = _cfg["portal_password"]

# Постоянный профиль Chrome.
# В нём уже сохранено разрешение portal-sot.kz на NCALayer.
CHROME_PROFILE = _cfg["chrome_profile"]


# ============================================================
# 1. ЗАПУСК CHROME
# ============================================================

def build_driver():

    options = Options()

    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")

    # Не закрывать Chrome после завершения скрипта
    options.add_experimental_option("detach", True)

    # Постоянный профиль
    options.add_argument(
        rf"--user-data-dir={CHROME_PROFILE}"
    )

    driver = webdriver.Chrome(options=options)

    driver.set_page_load_timeout(120)

    return driver


# ============================================================
# 2. ЖДЁМ NCALAYER
# ============================================================

def wait_ncalayer(timeout=60):

    desktop = Desktop(backend="uia")

    started = time.time()

    while time.time() - started < timeout:

        try:

            nca = desktop.window(title="NCALayer")

            if nca.exists() and nca.is_visible():
                return nca

        except Exception:
            pass

        time.sleep(0.5)

    raise RuntimeError(
        f"NCALayer не появился за {timeout} секунд"
    )


# ============================================================
# 3. ЖДЁМ КОНКРЕТНУЮ КНОПКУ NCALAYER
# ============================================================

def wait_nca_button(title, timeout=60):

    started = time.time()

    while time.time() - started < timeout:

        try:

            nca = Desktop(
                backend="uia"
            ).window(
                title="NCALayer"
            )

            if nca.exists() and nca.is_visible():

                button = nca.child_window(
                    title=title,
                    control_type="Button"
                )

                if button.exists():
                    return nca, button

        except Exception:
            pass

        time.sleep(0.5)

    raise RuntimeError(
        f"Кнопка NCALayer «{title}» "
        f"не появилась за {timeout} секунд"
    )


# ============================================================
# 4. ПОЛНЫЙ ВХОД
# ============================================================

def login_portal_eds():

    print("=" * 80)
    print("PORTAL-SOT.KZ — АВТОМАТИЧЕСКИЙ ВХОД ЧЕРЕЗ ЭЦП")
    print("=" * 80)

    # --------------------------------------------------------
    # Chrome
    # --------------------------------------------------------

    print("\n[1/10] Запускаем Chrome...")

    driver = build_driver()

    print("       ✅ Chrome запущен")


    # --------------------------------------------------------
    # Открываем портал
    # --------------------------------------------------------

    print("\n[2/10] Открываем portal-sot.kz...")

    driver.get(PORTAL_URL)

    wait = WebDriverWait(driver, 60)

    print("       ✅ Портал открыт")


    # --------------------------------------------------------
    # Если вдруг профиль уже авторизован
    # --------------------------------------------------------

    time.sleep(2)

    if "/cabinet" in driver.current_url.lower():

        print("\n✅ Уже авторизованы в кабинете")

    else:

        # ----------------------------------------------------
        # Первая кнопка ВОЙТИ
        # ----------------------------------------------------

        print("\n[3/10] Нажимаем первую кнопку «Войти»...")

        login_button = wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//*[self::button or self::a]"
                "[contains(normalize-space(.),'Войти')]"
            ))
        )

        login_button.click()

        print("       ✅ Нажато")


        # ----------------------------------------------------
        # Первое окно NCALayer
        # ----------------------------------------------------

        print("\n[4/10] Ждём NCALayer...")

        nca = wait_ncalayer(60)

        nca.set_focus()

        print("       ✅ NCALayer открыт")


        # ----------------------------------------------------
        # Пароль ЭЦП
        # ----------------------------------------------------

        print("\n[5/10] Вводим пароль ЭЦП...")

        password_edit = nca.child_window(
            title="Пароль",
            control_type="Edit"
        )

        if not password_edit.exists():

            # запасной вариант по известному auto_id
            password_edit = nca.child_window(
                auto_id="JavaFX39",
                control_type="Edit"
            )

        if not password_edit.exists():
            raise RuntimeError(
                "Не найдено поле пароля ЭЦП"
            )


        # JavaFX плохо работает с set_edit_text(),
        # поэтому вставляем пароль через clipboard.

        password_edit.click_input()

        time.sleep(0.3)

        password_edit.type_keys("^a{BACKSPACE}")

        time.sleep(0.2)

        password_edit.type_keys(
            EDS_PASSWORD,
            with_spaces=True,
            pause=0.05
        )

        time.sleep(0.5)


        # ----------------------------------------------------
        # Открыть
        # ----------------------------------------------------

        open_button = nca.child_window(
            title="Открыть",
            control_type="Button"
        )

        if not open_button.exists():

            open_button = nca.child_window(
                auto_id="JavaFX47",
                control_type="Button"
            )

        if not open_button.exists():
            raise RuntimeError(
                "Кнопка «Открыть» не найдена"
            )

        open_button.click_input()

        print("       ✅ Нажата «Открыть»")


        # ----------------------------------------------------
        # Второе окно — ПОДПИСАТЬ
        # ----------------------------------------------------

        print("\n[6/10] Ждём выбор сертификата...")

        nca, sign_button = wait_nca_button(
            "Подписать",
            timeout=60
        )

        nca.set_focus()

        print("       ✅ Сертификат загружен")

        time.sleep(0.5)

        sign_button.click_input()

        print("       ✅ Нажата «Подписать»")


        # ----------------------------------------------------
        # Возвращаемся в браузер
        # ----------------------------------------------------

        print(
            "\n[7/10] Ждём поле пароля portal-sot.kz..."
        )

        portal_password_input = WebDriverWait(
            driver,
            60
        ).until(
            EC.visibility_of_element_located((
                By.XPATH,
                "//input[@type='password']"
            ))
        )

        print("       ✅ Поле пароля появилось")


        # ----------------------------------------------------
        # Пароль портала
        # ----------------------------------------------------

        print(
            "\n[8/10] Вводим пароль портала..."
        )

        portal_password_input.clear()

        portal_password_input.send_keys(
            PORTAL_PASSWORD
        )

        print("       ✅ Пароль введён")


        # ----------------------------------------------------
        # Финальная кнопка ВОЙТИ
        # ----------------------------------------------------

        print(
            "\n[9/10] Нажимаем финальную кнопку «Войти»..."
        )

        final_login_button = WebDriverWait(
            driver,
            30
        ).until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//button[contains(normalize-space(.),'Войти')]"
            ))
        )

        final_login_button.click()

        print("       ✅ Нажато")


        # ----------------------------------------------------
        # Ждём кабинет
        # ----------------------------------------------------

        print(
            "\n[10/10] Ждём вход в кабинет..."
        )

        WebDriverWait(
            driver,
            60
        ).until(
            lambda d:
                "/cabinet" in d.current_url.lower()
        )

        print("       ✅ Кабинет открыт")


    # ========================================================
    # ТОКЕНЫ
    # ========================================================

    time.sleep(1)

    access_token = driver.execute_script(
        """
        return localStorage.getItem('access_token');
        """
    )

    refresh_token = driver.execute_script(
        """
        return localStorage.getItem('refresh_token');
        """
    )


    print()
    print("=" * 80)
    print("✅ АВТОРИЗАЦИЯ ЗАВЕРШЕНА")
    print("=" * 80)

    print("URL:", driver.current_url)

    print(
        "access_token:",
        "✅ получен"
        if access_token
        else "❌ не найден"
    )

    print(
        "refresh_token:",
        "✅ получен"
        if refresh_token
        else "❌ не найден"
    )

    if access_token:

        print(
            "Длина access_token:",
            len(access_token)
        )

    print("=" * 80)


    return (
        driver,
        access_token,
        refresh_token
    )


# ============================================================
# ЗАПУСК
# ============================================================

if __name__ == "__main__":

    driver, ACCESS_TOKEN, REFRESH_TOKEN = login_portal_eds()

    # Держим процесс живым, чтобы окно Chrome не закрывалось.
    input("\nНажмите Enter, чтобы завершить скрипт (Chrome останется открытым)...")
