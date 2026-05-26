# utils.py
"""
Вспомогательные функции для обработки документов
"""
import os
from config import LOG_LOCK, SUMMARY_LOCK, LOG_SUMMARY, LOG_FILE, FOLDER_LOCK

# ═══════════════════════════════════════════════════════════════
# ФУНКЦИИ НОРМАЛИЗАЦИИ
# ═══════════════════════════════════════════════════════════════

def normalize(text: str) -> str:
    """
    Нормализация названий продуктов для сравнения.
    Убирает пробелы, дефисы, подчёркивания и заменяет похожие кириллические буквы на латиницу.
    
    Args:
        text: Текст для нормализации
    
    Returns:
        Нормализованный текст (строчные буквы, без пробелов/дефисов)
    
    Example:
        >>> normalize("Cash-Credit КК")
        'cashcreditkk'
    """
    trans = str.maketrans(
        "аеорсухкмвнтзм",    # кириллические символы, похожие на латиницу
        "aercsyxkmvntzm"     # латинские эквиваленты
    )
    return (
        str(text).lower()
        .translate(trans)
        .replace('-', '')
        .replace('_', '')
        .replace(' ', '')
    )

# ═══════════════════════════════════════════════════════════════
# РАБОТА С ИМЕНАМИ ФАЙЛОВ И ПАПОК
# ═══════════════════════════════════════════════════════════════

def safe_folder_name(fio: str, iin: str) -> str:
    """
    Создаёт безопасное имя папки для Windows.
    Формат: 'ФИО, ИИН'
    
    Args:
        fio: ФИО клиента
        iin: ИИН клиента (12 цифр)
    
    Returns:
        Безопасное имя папки
    
    Example:
        >>> safe_folder_name("Иванов Иван", "123456789012")
        'Иванов Иван, 123456789012'
    """
    raw = f"{fio}, {iin}"
    
    # Заменяем запрещённые в Windows символы
    forbidden_chars = [
        ('/', '_'), ('\\', '_'), (':', '_'), ('*', '_'),
        ('?', '_'), ('"', '_'), ('<', '_'), ('>', '_'), ('|', '_')
    ]
    
    for bad, good in forbidden_chars:
        raw = raw.replace(bad, good)
    
    return raw

def filename_contains_iin(file_name: str, iin: str) -> bool:
    """
    Проверяет, содержит ли имя файла ИИН (игнорируя пробелы, дефисы, подчёркивания).
    
    Args:
        file_name: Имя файла (без расширения)
        iin: ИИН для поиска
    
    Returns:
        True если ИИН найден в имени файла
    
    Example:
        >>> filename_contains_iin("Doc_123456789012_scan", "123456789012")
        True
    """
    # Убираем все разделители из имени файла
    clean_filename = (
        file_name
        .replace(' ', '')
        .replace('-', '')
        .replace('_', '')
    )
    
    return iin in clean_filename

# ═══════════════════════════════════════════════════════════════
# РАБОТА С ПАПКАМИ КЛИЕНТОВ
# ═══════════════════════════════════════════════════════════════

def ensure_client_folder(iin: str, fio: str, base_dir: str) -> str:
    """
    Гарантирует существование папки клиента (потокобезопасно).
    
    Args:
        iin: ИИН клиента
        fio: ФИО клиента
        base_dir: Базовая директория для создания папки
    
    Returns:
        Полный путь к папке клиента
    
    Example:
        >>> ensure_client_folder("123456789012", "Иванов И.И.", "C:\\Docs")
        'C:\\Docs\\Иванов И.И., 123456789012'
    """
    folder_name = safe_folder_name(fio, iin)
    target_folder_path = os.path.join(base_dir, folder_name)
    
    # Используем блокировку для безопасного создания папки в многопоточной среде
    with FOLDER_LOCK:
        os.makedirs(target_folder_path, exist_ok=True)
    
    return target_folder_path

# ═══════════════════════════════════════════════════════════════
# ПОТОКОБЕЗОПАСНОЕ ЛОГИРОВАНИЕ
# ═══════════════════════════════════════════════════════════════

def safe_log(message: str, log_file: str = LOG_FILE):
    """
    Потокобезопасная запись в лог файл.
    
    Args:
        message: Сообщение для записи
        log_file: Путь к лог файлу (по умолчанию из config)
    
    Example:
        >>> safe_log("[ОФЕРТЫ] Не найден файл для клиента Иванов И.И.")
    """
    with LOG_LOCK:
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(message + '\n')

def safe_update_summary(block_name: str, stats: dict):
    """
    Потокобезопасное обновление сводной статистики LOG_SUMMARY.
    
    Args:
        block_name: Название блока (например, "ОФЕРТЫ")
        stats: Словарь со статистикой:
               {
                   'found': int,
                   'total': int,
                   'not_found': list[str]
               }
    
    Example:
        >>> safe_update_summary("ОФЕРТЫ", {
        ...     'found': 100,
        ...     'total': 120,
        ...     'not_found': ['Иванов, 123', 'Петров, 456']
        ... })
    """
    with SUMMARY_LOCK:
        LOG_SUMMARY[block_name] = stats

def clear_log():
    """
    Очищает лог файл (вызывается в начале обработки).
    """
    with LOG_LOCK:
        with open(LOG_FILE, 'w', encoding='utf-8') as f:
            f.write("=== ЛОГ ОБРАБОТКИ ДОКУМЕНТОВ ===\n")
            f.write(f"Дата запуска: {get_current_datetime()}\n\n")

# ═══════════════════════════════════════════════════════════════
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ═══════════════════════════════════════════════════════════════

def get_current_datetime() -> str:
    """
    Возвращает текущую дату и время в читаемом формате.
    
    Returns:
        Строка с датой и временем
    
    Example:
        >>> get_current_datetime()
        '2025-02-10 15:30:45'
    """
    from datetime import datetime
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')

def count_files_in_folder(folder_path: str, extension: str = None) -> int:
    """
    Подсчитывает количество файлов в папке.
    
    Args:
        folder_path: Путь к папке
        extension: Расширение файла для фильтрации (например, '.pdf')
    
    Returns:
        Количество файлов
    """
    if not os.path.exists(folder_path):
        return 0
    
    count = 0
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if extension is None or file.lower().endswith(extension.lower()):
                count += 1
    
    return count

def format_file_size(size_bytes: int) -> str:
    """
    Форматирует размер файла в человекочитаемый вид.
    
    Args:
        size_bytes: Размер в байтах
    
    Returns:
        Отформатированная строка
    
    Example:
        >>> format_file_size(1536)
        '1.5 KB'
    """
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.1f} TB"

# ═══════════════════════════════════════════════════════════════
# ФУНКЦИИ ДЛЯ ПОИСКА ФАЙЛОВ
# ═══════════════════════════════════════════════════════════════

def find_pdf_by_keyword(
    folder_path: str, 
    keyword: str, 
    case_sensitive: bool = False
) -> list:
    """
    Ищет PDF файлы по ключевому слову в имени.
    
    Args:
        folder_path: Путь к папке для поиска
        keyword: Ключевое слово для поиска
        case_sensitive: Учитывать регистр
    
    Returns:
        Список путей к найденным файлам
    """
    found_files = []
    
    if not os.path.exists(folder_path):
        return found_files
    
    search_keyword = keyword if case_sensitive else keyword.lower()
    
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if not file.lower().endswith('.pdf'):
                continue
            
            file_to_check = file if case_sensitive else file.lower()
            
            if search_keyword in file_to_check:
                found_files.append(os.path.join(root, file))
    
    return found_files