import os
import re
import fitz  # PyMuPDF
from datetime import datetime
import sys
import io


if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
# Путь к папке
base_folder = r"\\PC002\work folder\Госпошлины"
def latest_year_dir(root: str) -> str | None:
    now_year = str(datetime.now().year)
    candidates = [d for d in os.listdir(root) if os.path.isdir(os.path.join(root, d))]
    if not candidates:
        return None
    if now_year in candidates:
        return os.path.join(root, now_year)
    candidates_num = sorted(candidates, key=lambda x: int(re.sub(r'\D', '', x) or 0), reverse=True)
    return os.path.join(root, candidates_num[0])

DATE_RE = re.compile(r'(\d{2})\.(\d{2})\.(\d{4})')  # DD.MM.YYYY

def parse_date_from_foldername(name: str) -> datetime | None:
    m = DATE_RE.search(name)
    if not m:
        return None
    d, mth, y = map(int, m.groups())
    try:
        return datetime(y, mth, d)
    except ValueError:
        return None

def pick_latest_day_folder(year_dir: str) -> str | None:
    subdirs = [os.path.join(year_dir, d) for d in os.listdir(year_dir)
               if os.path.isdir(os.path.join(year_dir, d))]
    if not subdirs:
        return None

    dated, undated = [], []
    for p in subdirs:
        dt = parse_date_from_foldername(os.path.basename(p))
        if dt:
            dated.append((dt, p))
        else:
            undated.append(p)

    if dated:
        dated.sort(key=lambda x: x[0], reverse=True)
        return dated[0][1]

    return max(undated, key=lambda p: os.path.getmtime(p))
def extract_fio(text):
    """
    Ищет 3 слова с заглавной буквы сразу после 'госпошлина'
    """
    # Удаляем всё мешающее: переносы, неразрывные пробелы и лишние пробелы
    text = text.replace('\xa0', ' ').replace('\u202f', ' ').replace('\u200b', '')
    text = re.sub(r'[\n\r\f]+', ' ', text)
    text = re.sub(r'\s{2,}', ' ', text).strip()

    # Упрощённый универсальный шаблон
    match = re.search(r'госпошлина\s+([^\s]+)\s+([^\s]+)\s+([^\s]+)', text, re.IGNORECASE)
    if match:
        fio = f"{match.group(1)} {match.group(2)} {match.group(3)}"
        return ' '.join([w.capitalize() for w in fio.split()])
    return None

# Обход всех файлов
for root, dirs, files in os.walk(base_folder):
    for file in files:
        if file.lower().endswith('.pdf'):
            file_path = os.path.join(root, file)
            try:
                # Загружаем документ
                doc = fitz.open(file_path)
                # Используем режим "text" (весь текст целиком)
                full_text = doc.get_toc(simple=False)
                text = doc.get_page_text(0, "text")
                for i in range(1, len(doc)):
                    text += "\n" + doc.get_page_text(i, "text")
                doc.close()

                fio = extract_fio(text)
                print(f"📄 Файл: {file}")
                print(f"   ➤ ФИО: {fio}")

                if fio:
                    new_name = f"Госпошлина, {fio}.pdf"
                    new_name_clean = re.sub(r'[<>:"/\\|?*]', '', new_name)
                    new_path = os.path.join(root, new_name_clean)
                    os.rename(file_path, new_path)
                    print(f"✅ Переименовано в: {new_name_clean}\n")
                else:
                    print(f"⚠️ ФИО не найдено: {file}\n")

            except Exception as e:
                print(f"⛔ Ошибка: {file}: {e}\n")

print("🎉 Готово.")
