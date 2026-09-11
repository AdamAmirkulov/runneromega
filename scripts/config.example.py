# config.example.py
"""
ШАБЛОН конфигурации. НЕ содержит реальных данных.

Первый запуск на новой машине:
    copy scripts\config.example.py scripts\config.py
затем заполнить реальными значениями. Файл scripts/config.py в .gitignore
и на каждой машине (локальная / сервер) свой — пути и учётки отличаются.
"""
from threading import Lock
import os
from datetime import datetime
import glob
import sys
sys.stdout.reconfigure(encoding='utf-8')

# ═══════════════════════════════════════════════════════════════
# КОРЕНЬ ДЛЯ СКРИПТОВ ПЕРЕИМЕНОВАНИЯ / КОНВЕРТАЦИИ
# ───────────────────────────────────────────────────────────────
# Внутри этой папки лежат подпапки-партии, которые сотрудник выбирает
# в веб-форме из выпадающего списка (Vivus / ММ / Lime / Solva /
# Надписи / Отмены надписей / Госпошлины / MSG→PDF …).
# ВАЖНО: путь должен быть таким, каким его видит СЕРВЕР приложения
# (у сотрудников на компах шара смонтирована по-разному — это не важно).
# Должен быть простой строковой литерал (его читает app.py без импорта).
# ═══════════════════════════════════════════════════════════════
RENAME_ROOT = r'\\SERVER\share\work folder'

# ═══════════════════════════════════════════════════════════════
# КОРНЕВЫЕ ПАПКИ КОМПАНИЙ  (COMPANY_ID -> корень рабочих документов)
# ═══════════════════════════════════════════════════════════════
COMPANY_ROOTS = {
    '1': r'\\SERVER\share\company1',
    '2': r'\\SERVER\share\company2',
    '3': r'C:\path\to\company3',
    '4': r'C:\path\to\company4',
}

# ═══════════════════════════════════════════════════════════════
# УЧЁТНЫЕ ДАННЫЕ КОМПАНИЙ
# aisoip_* — АИС ОИП; sk_* — Судебный кабинет office.sud.kz
# ═══════════════════════════════════════════════════════════════
COMPANY_CREDENTIALS = {
    '1': {
        'aisoip_login':      'AISOIP_LOGIN',
        'aisoip_password':   'AISOIP_PASSWORD',
        'sk_login':          'SUDCABINET_LOGIN',
        'sk_password':       'SUDCABINET_PASSWORD',
        'loading_path_part': r'C:\path\to\company1\load',
        'path_load_finish':  r'C:\path\to\company1\load_finish',
        'path_crm':          r'C:\DeltaM\AutoImport\Company1',
        'name_login':        'company1',
        'spreadsheet_url':   'https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit',
        'credentials_path':  r'C:\path\to\google_credentials.json',
        'sheet_name':        'SHEET',
        'eid_prefix':        '1',
        'rep_iin':           'REP_IIN',
        'org_bin':           'ORG_BIN',
        'org_bank':          'ORG_IBAN',
    },
    # '2': { ... }, '3': { ... }, '4': { ... }  — по образцу выше
}

# Имя компании в CRM (поле l.F209) — для otmeny.py и SQL-фильтров
COMPANY_DB_FILTER = {
    '1': 'ТОО "..."',
    '2': 'ТОО "..."',
    '3': 'ТОО "..."',
    '4': 'ТОО "..."',
}

# Вход на portal-sot.kz через ЭЦП — для scripts/sud_fio_zayavlenie.py.
# Требует физического ЭЦП-сертификата компании на этой машине и отдельного
# Chrome-профиля с уже нажатым «Разрешить» для portal-sot.kz → NCALayer —
# см. подробный комментарий в реальном scripts/config.py. Компанию без
# записи здесь скрипт не запустит (явная ошибка вместо чужого сертификата).
PORTAL_SOT_BY_COMPANY = {
    '1': {
        'eds_password': '...',
        'portal_password': COMPANY_CREDENTIALS['1']['sk_password'],
        'chrome_profile': r'C:\...\ChromePortalSot',
    },
    # '2': { ... }, '3': { ... }, '4': { ... } — добавить по мере настройки
}

# ═══════════════════════════════════════════════════════════════
# SMTP для рассылок по отчётам (scripts/mailing_send.py) — свой ящик
# отправителя на каждую компанию (ключ = ID компании в users.db).
# ═══════════════════════════════════════════════════════════════
MAILING_SMTP = {
    '1': {
        'host': 'smtp.example.kz', 'port': 587,
        'user': 'info@example.kz', 'password': 'SMTP_PASSWORD',
        'from_name': 'ТОО «...»',
        'use_tls': True, 'bcc': '', 'send_delay_sec': 5,
    },
    # '2': { ... }, '3': { ... }, '4': { ... }  — по образцу выше
}

# ═══════════════════════════════════════════════════════════════
# Заявления на выдачу ИЛ (scripts/sk_zayavlenie_il.py)
# По каждой компании: истец в тексте заявления (project_name), ФИО
# директора (director) и имя .docx-шаблона в scripts/templates/.
# Компании без записи здесь скрипт для ИЛ запускать не даст.
# ═══════════════════════════════════════════════════════════════
IL_DECISION_CONFIG = {
    '1': {
        'project_name': 'ТОО «...»',
        'director': 'Фамилия И.',
        'template': 'OMEGA_IL_DECISION_TEMPLATE.docx',
    },
    # '2': { ... }, '3': { ... }, '4': { ... }
}

# ═══════════════════════════════════════════════════════════════
# PORTAL-SOT.KZ — новый Судебный кабинет (вход через ЭЦП)
# scripts/poiskvsk.py, блок «Поиск адреса в СК» — ПОКА ТОЛЬКО ОМЕГА
# (company_id=1). Вход не по компаниям — один аккаунт-представитель.
# chrome_profile — постоянный профиль Chrome, в котором ОДИН РАЗ вручную
# нажато «Разрешить» для portal-sot.kz → NCALayer (иначе запрос всплывает
# при каждом входе и автоматика застревает). Отдельный автономный скрипт
# входа — scripts/login_newsud.py (там же env PORTAL_EDS_PASSWORD /
# PORTAL_PASSWORD / PORTAL_CHROME_PROFILE, приоритетнее этого словаря).
# ═══════════════════════════════════════════════════════════════
PORTAL_SOT = {
    'eds_password':    'EDS_FILE_PASSWORD',      # пароль от файла ЭЦП (ключа)
    'portal_password': 'PORTAL_SOT_PASSWORD',    # пароль от portal-sot.kz
    'chrome_profile':  r'C:\Users\User\Documents\ChromePortalSot',
    # необязательно:
    # 'chrome_path':    r'C:\Program Files\Google\Chrome\Application\chrome.exe',
    # 'ncalayer_path':  r'C:\Users\User\AppData\Local\Programs\NCALayer\NCALayer.exe',
}

# ═══════════════════════════════════════════════════════════════
# WhatsApp (Wamm Chat) — теги при отправке реестра на возврат ГП
# ═══════════════════════════════════════════════════════════════
REESTR_GP_WA_TAGS = [
    '+70000000000',
]

# ═══════════════════════════════════════════════════════════════
# Ниже — производные значения, менять обычно не нужно
# ═══════════════════════════════════════════════════════════════
_company_id = os.environ.get('COMPANY_ID', '1')
CREDENTIALS = COMPANY_CREDENTIALS.get(_company_id, COMPANY_CREDENTIALS['1'])
DB_COMPANY_FILTER = COMPANY_DB_FILTER.get(_company_id, COMPANY_DB_FILTER['1'])
ROOT = COMPANY_ROOTS.get(_company_id, COMPANY_ROOTS['1'])
print(f" Компания ID={_company_id} | Корень: {ROOT}")


def _find_latest_excel() -> str:
    override = os.environ.get('OVERRIDE_EXCEL_PATH')
    if override:
        print(f" Используется указанный файл: {override}")
        return override

    search_dir = rf"{ROOT}\Документы для подачи Исков"
    pattern = os.path.join(search_dir, "Отчёт по отменам*.xlsx")
    files = glob.glob(pattern)

    if not files:
        print(f"⚠️  ВНИМАНИЕ: Файлы не найдены в {search_dir}")
        return os.path.join(search_dir, "Отчёт по отменам.xlsx")

    latest_file = max(files, key=os.path.getmtime)
    print(f"📊 Автоматически выбран файл: {os.path.basename(latest_file)}")
    return latest_file


MAIN_EXCEL = _find_latest_excel()


def _make_target_base() -> str:
    now = datetime.now()
    folder_name = now.strftime("%d.%m.%Y (%H-%M)")
    return rf"{ROOT}\Документы для подачи Исков\{folder_name}"


TARGET_BASE = _make_target_base()
print(f"TARGET_BASE: {TARGET_BASE}")

# ═══════════════════════════════════════════════════════════════
# ИСТОЧНИКИ ДОКУМЕНТОВ (ТОЛЬКО ЧТЕНИЕ)
# ═══════════════════════════════════════════════════════════════
BASE_CESSII        = rf"{ROOT}\Цессии"
EMAIL_SOURCE       = rf"{ROOT}\Уведомления об уступки права требования"
OFFERS_SOURCE      = rf"{ROOT}\Цессии\ММ-5\Оферты"
EXEC_NOTES_SOURCE  = rf"{ROOT}\Надписи (переименованные)"
GOV_FEES_SOURCE    = rf"{ROOT}\Госпошлины"
PRETRIAL_SOURCE    = rf"{ROOT}\Досудебные претензии"
SCREENSHOTS_SOURCE = rf"{ROOT}\Скрины"
RESOLUTIONS_SOURCE = rf"{ROOT}\Постановления"

# ═══════════════════════════════════════════════════════════════
# СЛУЖЕБНЫЕ ФАЙЛЫ / ГЛОБАЛЬНЫЕ НАСТРОЙКИ
# ═══════════════════════════════════════════════════════════════
LOG_FILE = 'log_not_copied.txt'
TEMP_DIR = 'temp'
LOG_SUMMARY = {}
LOG_LOCK = Lock()
SUMMARY_LOCK = Lock()
FOLDER_LOCK = Lock()
MAX_WORKERS = 8
COPY_TIMEOUT = 30
LOG_ENCODING = 'utf-8'


def ensure_directories():
    for directory in [TARGET_BASE, TEMP_DIR]:
        if not os.path.exists(directory):
            try:
                os.makedirs(directory, exist_ok=True)
                print(f" Создана директория: {directory}")
            except Exception as e:
                print(f"  Не удалось создать директорию {directory}: {e}")


def validate_source_paths():
    sources = {
        'MAIN_EXCEL': MAIN_EXCEL,
        'BASE_CESSII': BASE_CESSII,
        'EMAIL_SOURCE': EMAIL_SOURCE,
    }
    missing = [f"{n}: {p}" for n, p in sources.items() if not os.path.exists(p)]
    if missing:
        print("  ВНИМАНИЕ! Не найдены следующие пути:")
        for item in missing:
            print(f"    {item}")
        print("\nПроверьте config.py и исправьте пути!\n")
        return False
    return True
