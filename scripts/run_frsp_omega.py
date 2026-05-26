import subprocess
import sys
import os
import argparse

parser = argparse.ArgumentParser()
parser.add_argument("--workdir", default="")
args, _ = parser.parse_known_args()

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
nb_path = os.path.join(BASE_DIR, "scripts", "Отчет ФРСП (А-Омега).ipynb")

# Запускаем с cwd=BASE_DIR чтобы service_account.json нашёлся
result = subprocess.run(
    [sys.executable, "-m", "jupyter", "nbconvert",
     "--to", "notebook",
     "--execute",
     "--inplace",
     "--ExecutePreprocessor.timeout=600",
     nb_path],
    cwd=BASE_DIR  # <-- вот главное исправление
)

sys.exit(result.returncode)