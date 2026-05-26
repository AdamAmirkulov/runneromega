import argparse
import os
import runpy
from pathlib import Path


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--workdir", required=True)
    p.add_argument("--work_mode", default="")
    p.add_argument("--login", default="")
    p.add_argument("--password", default="")
    args, _unknown = p.parse_known_args()

    # --- рабочая папка job (для скачиваний Selenium, временных файлов, логов панели) ---
    workdir = Path(args.workdir)
    (workdir / "downloads").mkdir(parents=True, exist_ok=True)
    (workdir / "out").mkdir(parents=True, exist_ok=True)
    (workdir / "chrome_profile").mkdir(parents=True, exist_ok=True)

    # --- ищем входной Excel в последней папке по дате/времени изменения ---
    BASE_DOCS = Path(r"D:\work folder\Документы для подачи Исков")
    REPORT_NAME = "Отчёт по отменам_Omega.xlsx"

    if not BASE_DOCS.exists():
        raise FileNotFoundError(f"Базовая папка не найдена: {BASE_DOCS}")

    folders = [p for p in BASE_DOCS.iterdir() if p.is_dir()]
    if not folders:
        raise FileNotFoundError(f"Нет подпапок в: {BASE_DOCS}")

    # последняя папка по времени изменения
    latest_folder = max(folders, key=lambda p: p.stat().st_mtime)

    report_path = latest_folder / REPORT_NAME
    if not report_path.exists():
        raise FileNotFoundError(
            f"Не найден файл '{REPORT_NAME}' в последней папке: {latest_folder}\n"
            f"Ожидали: {report_path}"
        )

    output_folder = latest_folder  # сохраняем туда же

    # --- ENV для большого скрипта ---
    os.environ["OMEGA_WORKDIR"] = str(workdir)
    os.environ["OMEGA_DOWNLOADS"] = str(workdir / "downloads")
    os.environ["OMEGA_OUT"] = str(workdir / "out")
    os.environ["OMEGA_CHROME_PROFILE"] = str(workdir / "chrome_profile")

    os.environ["OMEGA_WORK_MODE"] = args.work_mode
    os.environ["OMEGA_LOGIN"] = args.login
    os.environ["OMEGA_PASSWORD"] = args.password

    # Входной отчёт и папка сохранения (по вашей бизнес-логике)
    os.environ["OMEGA_REPORT_XLSX"] = str(report_path)
    os.environ["OMEGA_OUTPUT_DIR"] = str(output_folder)

    # --- информативные логи ---
    script_path = Path(__file__).resolve().parent / "podacha_iska_v2.py"
    print("RUNNER: starting       =", script_path)
    print("RUNNER: WORKDIR        =", os.environ["OMEGA_WORKDIR"])
    print("RUNNER: latest_folder  =", latest_folder)
    print("RUNNER: report_path    =", report_path)
    print("RUNNER: output_folder  =", output_folder)

    # --- запуск большого скрипта ---
    runpy.run_path(str(script_path), run_name="__main__")


if __name__ == "__main__":
    main()
