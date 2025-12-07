    # all_reports_run.py
from pathlib import Path
import subprocess
import sys

# какие скрипты запускать и в каком порядке
SCRIPTS = [
    "extract_vo2max.py",
    "extract_weight_and_vo2.py",
    "walks_by_week.py",
    "walks_total.py",
    "weekly_from_daily.py",
    "health_last_walk.py",
]

def main() -> None:
    project_root = Path(__file__).resolve().parent
    reports_dir = project_root / "reports"
    reports_dir.mkdir(exist_ok=True)

    # какие xlsx были до запуска
    before = {p.resolve() for p in project_root.glob("*.xlsx")}

    # запускаем все скрипты тем же интерпретатором (venv)
    for script_name in SCRIPTS:
        script_path = project_root / script_name
        if not script_path.exists():
            print(f"[WARN] Скрипт {script_name} не найден, пропускаю")
            continue

        print(f"=== Запуск {script_name} ===")
        subprocess.run([sys.executable, str(script_path)], check=True)

    # какие xlsx появились после
    after = {p.resolve() for p in project_root.glob("*.xlsx")}
    new_files = after - before

    # переносим только новые xlsx в reports/
    for path in new_files:
        target = reports_dir / path.name
        if target.exists():
            target.unlink()  # перезаписать, если файл с таким именем уже был
        path.rename(target)
        print(f"Переместил {path.name} -> {reports_dir.name}/")

    print("Готово.")

if __name__ == "__main__":
    main()