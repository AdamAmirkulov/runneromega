import subprocess
import sys
import os
import argparse

parser = argparse.ArgumentParser()
parser.add_argument("--workdir", default="")
args, _ = parser.parse_known_args()

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
nb_path = os.path.join(BASE_DIR, "scripts", "Отчет ФРСП (Orion).ipynb")

env = os.environ.copy()
env["JUPYTER_RUNTIME_DIR"] = BASE_DIR  # ядро стартует из BASE_DIR

result = subprocess.run(
    [sys.executable, "-m", "jupyter", "nbconvert",
     "--to", "notebook",
     "--execute",
     "--inplace",
     "--ExecutePreprocessor.timeout=600",
     "--ExecutePreprocessor.kernel_name=python3",
     nb_path],
    cwd=BASE_DIR,
    env=env
)

sys.exit(result.returncode)