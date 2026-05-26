# scripts_registry.py
from dataclasses import dataclass
from typing import List, Dict, Any




@dataclass
class ScriptDef:
    key: str
    title: str
    description: str
    command: list
    params: list
    allow_scheduled_start: bool = False   # ✅ новое поле, по умолчанию выключено

SCRIPTS: Dict[str, ScriptDef] = {
    
    # ═══════════════════════════════════════════════════════════════
    # ОСНОВНЫЕ СКРИПТЫ
    # ═══════════════════════════════════════════════════════════════
    
    
    "status": ScriptDef(
        key="status",
        title="📊 АИС ОИП - Статусы",
        description="Проверка статусов дел и документов",
        command=["python", "-u", "scripts/status.py"],
        params=[
            {"name": "company_id", "label": "Компания", "type": "select"},
        ]
    ),

        "otmeny": ScriptDef(
        key="otmeny",
        title="🔄 АИС ОИП - Отмены",
        description="Обработка отмен исполнительных производств",
        command=["python", "-u", "scripts/otmeny.py"],
        params=[
            {"name": "company_id", "label": "Компания", "type": "select"},
            {"name": "date_from", "label": "Дата с", "type": "date"},
            {"name": "date_to", "label": "Дата по", "type": "date"},
            {"name": "login", "label": "Логин системы", "type": "text"},
            {"name": "password", "label": "Пароль системы", "type": "password"},
        ],
    ),
    
    "ogranichenie": ScriptDef(
        key="ogranichenie",
        title="🔍АИС ОИП - Ограничение на выезд",
        description="Проверка постановлений на ограничение выезда в системе АИС ОИП",
        command=["python", "-u", "scripts/ogranichenie.py"],
        params=[
            {"name": "iin", "label": "ИИН (если нужен)", "type": "text"},
            {"name": "login", "label": "Логин системы", "type": "text"},
            {"name": "password", "label": "Пароль системы", "type": "password"},
        ],
    ),
    
    "sud": ScriptDef(
        key="sud",
        title="⚖️ Выгрузка с СК",
        description="Выгрузка статусов дел из Судебного кабинета (office.sud.kz) по талонам компании",
        command=["python", "-u", "scripts/sud.py"],
        params=[
            {"name": "company_id", "label": "Компания", "type": "select"},
        ],
    ),
    "sudFIO": ScriptDef(
        key="sudFIO",
        title="⚖️ Судебный кабинет - ФИО, адрес",
        description="Выгрузка ФИО и адреса через СК",
        # Раньше этот пункт был захардкожен на scripts/sud.py — тот файл
        # раньше содержал именно эту логику (подача иска + добавление
        # участника + автоподстановка ФИО/адреса по ИИН), а не выгрузку
        # статусов дел. scripts/sud.py теперь заменён на новую выгрузку
        # с СК по талонам (см. ключ "sud" выше), поэтому старое содержимое
        # перенесено без изменений в scripts/sud_fio_zayavlenie.py, чтобы
        # эта кнопка продолжала работать как раньше.
        command=["python", "-u", "scripts/sud_fio_zayavlenie.py"],
        params=[
            {"name": "date_from", "label": "Дата с", "type": "date"},
            {"name": "date_to", "label": "Дата по", "type": "date"},
        ],
    ),

    "poiskvsk": ScriptDef(
        key="poiskvsk",
        title="🔍 Поиск Адреса в СК + Привязка к Суду",
        description="Поиск адреса в страховой компании и привязка к соответствующему суду",
        command=["python", "-u", "scripts/poiskvsk.py"],
        params=[
            {
                "name": "excel_file",
                "label": "Excel файл",
                "type": "file",
                "accept": ".xlsx,.xls",
                "save_to": "uploads/poiskvsk_input.xlsx",
            },
            {"name": "company_id", "label": "Компания", "type": "select"},
        ],
    ),

        "sbor": ScriptDef(
        key="sbor",
        title="📦 Сбор документов для Исков",
        description="Полный автоматический сбор всех документов для подачи исков",
        command=["python", "-u", "scripts/sbor.py"],
        params=[
            {"name": "company_id", "label": "Компания", "type": "select"},
            {"name": "excel_path", "label": "Путь к Excel файлу (опционально)", "type": "text"},
        ],
    ),

        "chsi_podacha": ScriptDef(
            key="chsi_podacha",
            title="📨 Подача иска в СК",
            description="Подача иска в Судебном Кабинете",
            command=["python", "-u", "scripts/chsi_podacha_runner.py"],
            allow_scheduled_start=True,
            params=[
                {"name": "company_id", "label": "Компания", "type": "select"},
            ],
        ),

        "podacha_iska_v2": ScriptDef(
            key="podacha_iska_v2",
            title="📨 Подача иска в СК(ЧЕРЕЗ API)",
            description="Подача иска в Судебном Кабинете через канал API",
            command=["python", "-u", "scripts/podacha_iska_v2.py"],
            allow_scheduled_start=True,
            params=[
                {"name": "company_id", "label": "Компания", "type": "select"},
            ],
        ),

        "reestr_gosposhliny": ScriptDef(
            key="reestr_gosposhliny",
            title="📋 Реестр на возврат госпошлины (СК + WhatsApp)",
            description=(
                "Выгружает из БД должников по условиям отбора, формирует реестр по шаблону, "
                "ищет актуальный адрес в Судебном кабинете по ИИН и отправляет сводку в WhatsApp-группу"
            ),
            command=["python", "-u", "scripts/reestr_gosposhliny.py"],
            params=[
                {"name": "company_id", "label": "Компания", "type": "select"},
            ],
        ),

        "reestr_gp_import": ScriptDef(
            key="reestr_gp_import",
            title="📥 Возврат ГП — импорт в Дельту (+ сбор/подача)",
            description=(
                "Принимает PDF с оплаченными госпошлинами и ранее отправленный "
                "в WhatsApp реестр. Сопоставляет платежи с реестром по ФИО, формирует "
                "LoansImport.xlsx в папке автоимпорта Дельты, затем тянет данные из БД "
                "по EID в формате «Отчёта по отменам» и прогоняет сбор документов и "
                "подачу иска. Этап: 1 — только LoansImport; 2 — + сбор; 3 — + подача "
                "(по умолчанию 3). Старые sbor.py / podacha_iska_v2.py не затрагиваются."
            ),
            command=["python", "-u", "scripts/reestr_gp_import.py"],
            params=[
                {"name": "company_id", "label": "Компания", "type": "select"},
                {
                    "name": "pdf_file",
                    "label": "PDF с оплаченными госпошлинами (Реквизиты …)",
                    "type": "file",
                    "accept": ".pdf",
                    "save_to": "uploads/gp_rekvizity_input.pdf",
                },
                {
                    "name": "reestr_file",
                    "label": "Реестр на возврат ГП (xlsx из WhatsApp)",
                    "type": "file",
                    "accept": ".xlsx,.xls",
                    "save_to": "uploads/gp_reestr_input.xlsx",
                },
                {"name": "stage", "label": "Этап (1 / 2 / 3, по умолчанию 3)", "type": "text"},
            ],
        ),

        "sud_download_mediation": ScriptDef(
            key="sud_download_mediation",
            title="📄 Скачивание определений о медиации (СК)",
            description=(
                "Ищет дела из отчёта по Возврату Госпошлин в Судебном кабинете и скачивает "
                "определения об утверждении медиации/мирового соглашения (RU и KZ, DOC/DOCX)"
            ),
            command=["python", "-u", "scripts/sud_download_mediation.py"],
            params=[
                {"name": "company_id", "label": "Компания", "type": "select"},
                {
                    "name": "excel_file",
                    "label": "Excel «Отчёт по Возврату Госпошлин»",
                    "type": "file",
                    "accept": ".xlsx,.xls",
                },
            ],
        ),

        "podacha_iska": ScriptDef(
        key="podacha_iska",
        title="📤 Подача заявления в АИС ОИП ",
        description="Подготовка и подача искового заявления",
        command=["python", "-u", "scripts/podacha_iska_runner.py"],
        params=[
            {"name": "work_mode", "label": "Режим работы", "type": "text"},
            {"name": "login", "label": "Логин системы", "type": "text"},
            {"name": "password", "label": "Пароль системы", "type": "password"},
        ],
    ),

    # ═══════════════════════════════════════════════════════════════
    # ПАРСИНГ ФИО / АДРЕСОВ
    # ═══════════════════════════════════════════════════════════════

    "sud_parser": ScriptDef(
        key="sud_parser",
        title="🔎 Парсинг ФИО и адресов по ИИН (sud.kz)",
        description=(
            "Загружает Excel с ИИН, авторизуется на office.sud.kz и выгружает "
            "ФИО, место жительства, страну и регион для каждого физлица. "
            "Поддерживает автоперезапуск браузера (до 20 раз) и продолжение с контрольной точки."
        ),
        command=["python", "-u", "scripts/sud_parser.py"],
        params=[
            {
                "name": "input_file",
                "label": "Excel-файл с ИИН (колонка A, с 2-й строки)",
                "type": "file",
                "accept": ".xlsx,.xls",
                "save_to": "uploads/sud_parser_input.xlsx",
            },
            {
                "name": "login",
                "label": "Логин (ИИН/БИН для office.sud.kz)",
                "type": "text",
            },
            {
                "name": "password",
                "label": "Пароль для office.sud.kz",
                "type": "password",
            },
        ],
    ),
    
    # ═══════════════════════════════════════════════════════════════
    # РАБОТА С СИСТЕМАМИ
    # ═══════════════════════════════════════════════════════════════
    

    
    # ═══════════════════════════════════════════════════════════════
    # ПОДАЧА ДОКУМЕНТОВ
    # ═══════════════════════════════════════════════════════════════
    "frsp_omega": ScriptDef(
        key="frsp_omega",
        title="📈 Отчёт по ФРСП Омега",
        description="Формирование отчёта по ФРСП для Омега",
        command=["python", "-u", "scripts/run_frsp_omega.py"],
        params=[],
    ),

    "frsp_kpi": ScriptDef(
        key="frsp_kpi",
        title="📊 Отчёт по ФРСП KPI",
        description="Формирование отчёта КПИ по ФРСП для KPI",
        command=["python", "-u", "scripts/run_frsp_kpi.py"],
        params=[],
    ),

    "frsp_orion": ScriptDef(
        key="frsp_orion",
        title="📉 Отчёт по ФРСП Орион",
        description="Формирование отчёта по ФРСП для Орион",
        command=["python", "-u", "scripts/run_frsp_orion.py"],
        params=[],
    ),

    

    
    # ═══════════════════════════════════════════════════════════════
    # ПЕРЕИМЕНОВАНИЕ И КОНВЕРТАЦИЯ
    # ═══════════════════════════════════════════════════════════════
    
    "rename_mm": ScriptDef(
        key="rename_mm",
        title="🏷️ Переименование ММ",
        description="Переименование файлов для продукта ММ",
        command=["python", "-u", "scripts/rename_mm.py"],
        params=[
            {"name": "source_folder", "label": "Исходная папка", "type": "text"},
        ],
    ),
    
    "rename_lime": ScriptDef(
        key="rename_lime",
        title="🏷️ Переименование Lime",
        description="Переименование файлов для продукта Lime",
        command=["python", "-u", "scripts/rename_lime.py"],
        params=[
            {"name": "source_folder", "label": "Исходная папка", "type": "text"},
        ],
    ),
    
    "rename_vivus": ScriptDef(
        key="rename_vivus",
        title="🏷️ Переименование Vivus",
        description="Переименование файлов для продукта Vivus",
        command=["python", "-u", "scripts/rename_vivus.py"],
        params=[
            {"name": "source_folder", "label": "Исходная папка", "type": "text"},
        ],
    ),
    
    "rename_solva": ScriptDef(
        key="rename_solva",
        title="🏷️ Переименование Solva",
        description="Переименование файлов для продукта Solva",
        command=["python", "-u", "scripts/rename_solva.py"],
        params=[
            {"name": "source_folder", "label": "Исходная папка", "type": "text"},
        ],
    ),
    
    "bereke_split_rename": ScriptDef(
        key="bereke_split_rename",
        title="🏷️ Bereke Bank - разделение и переименование",
        description="Разделяет PDF на Досудебную претензию и Уведомление об уступки права, переименовывает по ФИО и ИИН",
        command=["python", "-u", "scripts/bereke_split_rename.py"],
        params=[
            {"name": "source_folder", "label": "Исходная папка", "type": "text"},
        ],
    ),

    "rename_gp": ScriptDef(
        key="rename_gp",
        title="🏷️ Переименование Госпошлин",
        description="Переименование файлов госпошлин",
        command=["python", "-u", "scripts/gp_rename.py"],
        params=[
            {"name": "source_folder", "label": "Исходная папка", "type": "text"},
        ],
    ),
    
    "rename_nadpisi": ScriptDef(
        key="rename_nadpisi",
        title="🏷️ Переименование Надписей",
        description="Переименование исполнительных надписей",
        command=["python", "-u", "scripts/nadpisi_rename.py"],
        params=[
            {"name": "source_folder", "label": "Исходная папка", "type": "text"},
        ],
    ),
    
    "rename_otmenanad": ScriptDef(
        key="rename_otmenanad",
        title="🏷️ Переименование Отмен Надписей",
        description="Переименование постановлений об отмене исполнительных надписей",
        command=["python", "-u", "scripts/rename_otmenanad.py"],
        params=[
            {"name": "source_folder", "label": "Исходная папка", "type": "text"},
        ],
    ),
    
    "convert_msg": ScriptDef(
        key="convert_msg",
        title="📧 Конвертация MSG в PDF",
        description="Конвертация email файлов .msg в PDF формат",
        command=["python", "-u", "scripts/convert_msg.py"],
        params=[
            {"name": "company_id", "label": "Компания", "type": "select"},
            {"name": "source_folder", "label": "Папка с MSG файлами", "type": "text"},
            {"name": "output_folder", "label": "Папка для PDF (опционально)", "type": "text"},
            {"name": "workers", "label": "Параллельных Chrome (по умолч. 4)", "type": "text"},
        ],
    ),
    
    # ═══════════════════════════════════════════════════════════════
    # ИМПОРТ И ОБРАБОТКА ПЛАТЕЖЕЙ
    # ═══════════════════════════════════════════════════════════════
    
    "paymentimporter": ScriptDef(
        key="paymentimporter",
        title="💳 Payment Importer",
        description="Импорт и обработка платёжных данных из Excel файла",
        command=["python", "-u", "scripts/paymentimporter.py"],
        params=[
            {"name": "company_id", "label": "Компания", "type": "select"},
            {
                "name": "excel_file",
                "label": "Excel файл с выпиской",
                "type": "file",
                "accept": ".xlsx,.xls",
                "save_to": "uploads/payments_input.xlsx",
            },
            {
                "name": "output_folder",
                "label": "Папка для результатов (опционально)",
                "type": "text",
            },
        ],
    ),
    
    # ═══════════════════════════════════════════════════════════════
    # УСТАРЕВШИЕ / РЕЗЕРВНЫЕ
    # ═══════════════════════════════════════════════════════════════
    
    "sbor_documentov": ScriptDef(
        key="sbor_documentov",
        title="📦 Сбор документов (старая версия)",
        description="Устаревшая версия сбора документов (используйте основной 'Сбор документов')",
        command=["python", "-u", "scripts/sbor_documentov.py"],
        params=[],
    ),
}