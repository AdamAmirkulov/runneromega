# blocks/__init__.py
"""
Модуль с блоками обработки документов для исков
"""

__version__ = "1.0.0"

# ═══════════════════════════════════════════════════════════════
# ИМПОРТЫ БЛОКОВ
# ═══════════════════════════════════════════════════════════════

# Этап 1: Подготовка
try:
    from . import logi
except ImportError:
    print("⚠️  logi.py не найден, создайте файл")

try:
    from . import sozdaniepapok
except ImportError:
    print("⚠️  sozdaniepapok.py не найден")

try:
    from . import poiskvsk
except ImportError:
    print("⚠️  poiskvsk.py не найден")

# Этап 2: Сбор документов
try:
    from . import oferty
except ImportError:
    print("⚠️  oferty.py не найден")

try:
    from . import uvedomlenie
except ImportError:
    print("⚠️  uvedomlenie.py не найден")

try:
    from . import dogovor
except ImportError:
    print("⚠️  dogovor.py не найден")

try:
    from . import reestrdogovora
except ImportError:
    print("⚠️  reestrdogovora.py не найден")

try:
    from . import ispolnadpisi
except ImportError:
    print("⚠️  ispolnadpisi.py не найден")

try:
    from . import gosposhliny
except ImportError:
    print("⚠️  gosposhliny.py не найден")

try:
    from . import dosudebnaya
except ImportError:
    print("⚠️  dosudebnaya.py не найден")

try:
    from . import screen
except ImportError:
    print("⚠️  screen.py не найден")

try:
    from . import raschet
except ImportError:
    print("⚠️  raschet.py не найден")

try:
    from . import iskovoezayav
except ImportError:
    print("⚠️  iskovoezayav.py не найден")

try:
    from . import postobotmenee
except ImportError:
    print("⚠️  postobotmenee.py не найден")

# Этап 3: Финализация
try:
    from . import formirovaniepool
except ImportError:
    print("⚠️  formirovaniepool.py не найден")

try:
    from . import formirovaniemassovogo
except ImportError:
    print("⚠️  formirovaniemassovogo.py не найден")

try:
    from . import logirovanie
except ImportError:
    print("⚠️  logirovanie.py не найден")

# ═══════════════════════════════════════════════════════════════
# ЭКСПОРТ
# ═══════════════════════════════════════════════════════════════

__all__ = [
    'logi',
    'sozdaniepapok',
    'poiskvsk',
    'oferty',
    'uvedomlenie',
    'dogovor',
    'reestrdogovora',
    'ispolnadpisi',
    'gosposhliny',
    'dosudebnaya',
    'screen',
    'raschet',
    'iskovoezayav',
    'postobotmenee',
    'formirovaniepool',
    'formirovaniemassovogo',
    'logirovanie',
]