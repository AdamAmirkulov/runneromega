import sys
import os

# Показываем путь поиска модулей
print("sys.path:")
for p in sys.path:
    print(" ", p)

# Проверяем существование папки sbordoc и файлов внутри
sbordoc_path = os.path.join(os.path.dirname(__file__), "scripts", "sbordoc")
print("\nПроверяем папку sbordoc:", sbordoc_path)
print("Существует?", os.path.exists(sbordoc_path))

if os.path.exists(sbordoc_path):
    files = os.listdir(sbordoc_path)
    print("Содержимое sbordoc:")
    for f in files:
        print(" ", f)

# Проверяем наличие конкретного файла, который не найден
file_to_check = os.path.join(sbordoc_path, "logi.py")
print("\nПроверяем файл logi.py:", file_to_check)
print("Существует?", os.path.exists(file_to_check))
