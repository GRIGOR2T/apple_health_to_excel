from pathlib import Path
from lxml import etree as ET
import datetime as dt
import pandas as pd
from tqdm import tqdm
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font

XML_FILE = Path("export.xml")

DISTANCE_TYPES = {
    "HKQuantityTypeIdentifierDistanceWalkingRunning",
    "HKQuantityTypeIdentifierDistanceWalking",
    "HKQuantityTypeIdentifierDistanceHiking",
}

WORKOUT_TYPES = {
    "HKWorkoutActivityTypeWalking",
}

# берём только часы
PRIMARY_DEVICE_MARK = "Watch"

# отсечка по дате
CUTOFF_DATE = dt.date(2025, 8, 1)


def parse_dt(s: str) -> dt.datetime:
    return dt.datetime.strptime(s[:19], "%Y-%m-%d %H:%M:%S")


def distance_value_to_km(value: float, unit: str | None) -> float:
    if unit in ("m", "meter", "meters"):
        return value / 1000.0
    if unit in ("km", "kilometer", "kilometers", None):
        return value
    if unit in ("mi", "mile", "miles"):
        return value * 1.60934
    return value


def is_watch_source(elem: ET._Element) -> bool:
    """
    Смотрим и sourceName, и device, нормализуем NBSP → ' '.
    """
    src = (elem.get("sourceName") or "") + " " + (elem.get("device") or "")
    src = src.replace("\u00a0", " ")
    return PRIMARY_DEVICE_MARK in src


def parse_workouts(xml_path: Path) -> pd.DataFrame:
    """
    Тренировки-ходьба только с часов.
    Возвращает: date, distance_workouts_km
    """
    rows = []

    print("Читаю XML (Workout):", xml_path.name)
    context = ET.iterparse(str(xml_path), events=("end",), tag="Workout")

    for _, elem in tqdm(context, desc="Парсинг тренировок", unit="rec"):
        w_type = elem.get("workoutActivityType")
        if w_type not in WORKOUT_TYPES:
            elem.clear()
            continue

        if not is_watch_source(elem):
            elem.clear()
            continue

        start_str = elem.get("startDate")
        if not start_str:
            elem.clear()
            continue

        try:
            start_dt = parse_dt(start_str)
        except Exception:
            elem.clear()
            continue

        date = start_dt.date()
        if date < CUTOFF_DATE:
            elem.clear()
            continue

        # дистанция
        dist = None
        dist_str = elem.get("totalDistance")
        unit = elem.get("totalDistanceUnit")

        if dist_str:
            try:
                dist = float(dist_str)
            except Exception:
                dist = None

        if dist is None:
            for ws in elem.findall("WorkoutStatistics"):
                t = ws.get("type")
                if t in DISTANCE_TYPES:
                    s = ws.get("sum")
                    u = ws.get("unit")
                    try:
                        dist = float(s)
                        unit = u
                        break
                    except Exception:
                        continue

        if dist is None:
            dist_km = 0.0
        else:
            dist_km = distance_value_to_km(dist, unit)

        rows.append({"date": date, "distance_workouts_km": dist_km})
        elem.clear()

    del context

    if not rows:
        return pd.DataFrame(columns=["date", "distance_workouts_km"])

    df = pd.DataFrame(rows)
    df = df.groupby("date", as_index=False)["distance_workouts_km"].sum()
    return df


def parse_daily_walking(xml_path: Path) -> pd.DataFrame:
    """
    Сутки ходьбы по Record (только часы).
    Возвращает: date, distance_km
    """
    rows = []

    print("Читаю XML (Record walking):", xml_path.name)
    context = ET.iterparse(str(xml_path), events=("end",), tag="Record")

    for _, elem in tqdm(context, desc="Парсинг дистанций", unit="rec"):
        r_type = elem.get("type")
        if r_type not in DISTANCE_TYPES:
            elem.clear()
            continue

        if not is_watch_source(elem):
            elem.clear()
            continue

        value_str = elem.get("value")
        unit = elem.get("unit")
        start = elem.get("startDate")

        if not value_str or not start:
            elem.clear()
            continue

        try:
            value = float(value_str)
        except Exception:
            elem.clear()
            continue

        try:
            dt_start = parse_dt(start)
        except Exception:
            elem.clear()
            continue

        date = dt_start.date()
        if date < CUTOFF_DATE:
            elem.clear()
            continue

        dist_km = distance_value_to_km(value, unit)
        rows.append({"date": date, "distance_km": dist_km})
        elem.clear()

    del context

    if not rows:
        return pd.DataFrame(columns=["date", "distance_km"])

    df = pd.DataFrame(rows)
    df = df.groupby("date", as_index=False)["distance_km"].sum()
    return df


def build_daily_walk_table(df_daily_dist: pd.DataFrame, df_workouts: pd.DataFrame) -> pd.DataFrame:
    """
    df_daily_dist: date, distance_km (вся ходьба, часы)
    df_workouts:   date, distance_workouts_km (ходьба в тренировках)
    """
    # если по тренировкам пусто — создаём пустую колонку заранее
    if df_workouts.empty:
        df_workouts = pd.DataFrame(columns=["date", "distance_workouts_km"])

    df = pd.merge(df_daily_dist, df_workouts, on="date", how="left")

    if "distance_workouts_km" not in df.columns:
        df["distance_workouts_km"] = 0.0

    df["distance_workouts_km"] = df["distance_workouts_km"].fillna(0.0)
    df["distance_non_workout_km"] = df["distance_km"] - df["distance_workouts_km"]
    df.loc[df["distance_non_workout_km"] < 0, "distance_non_workout_km"] = 0.0

    for col in ["distance_km", "distance_workouts_km", "distance_non_workout_km"]:
        df[col] = df[col].round(2)

    df = df.sort_values("date", ascending=False).reset_index(drop=True)

    df["Дата"] = df["date"].apply(lambda d: d.strftime("%m/%d/%Y"))
    df = df.rename(
        columns={
            "distance_km": "Всего ходьба, км",
            "distance_workouts_km": "Ходьба в тренировках, км",
            "distance_non_workout_km": "Ходьба вне тренировок, км",
        }
    )

    df = df[
        [
            "Дата",
            "Всего ходьба, км",
            "Ходьба в тренировках, км",
            "Ходьба вне тренировок, км",
        ]
    ]

    return df


def output_filename() -> str:
    now = dt.datetime.now().strftime("%Y%m%d_%H%M")
    return f"daily_walk.xlsx"


def autoformat_sheet(ws):
    for col_idx, column_cells in enumerate(ws.columns, start=1):
        max_len = 0
        for cell in column_cells:
            cell.font = Font(size=28)
            val = cell.value
            if val is None:
                continue
            text = str(val)
            if len(text) > max_len:
                max_len = len(text)
        ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 4


def save_excel(daily_df: pd.DataFrame) -> None:
    out_name = output_filename()
    print("Сохраняю в Excel:", out_name)

    with pd.ExcelWriter(out_name, engine="openpyxl") as writer:
        daily_df.to_excel(writer, sheet_name="Daily_walk", index=False)
        ws = writer.sheets["Daily_walk"]
        ws.freeze_panes = "A2"
        autoformat_sheet(ws)


def main():
    df_daily_dist = parse_daily_walking(XML_FILE)
    df_workouts = parse_workouts(XML_FILE)

    print("Дней с данными ходьбы:", len(df_daily_dist))
    print("Дней с тренировками ходьбы:", len(df_workouts))

    daily_table = build_daily_walk_table(df_daily_dist, df_workouts)
    save_excel(daily_table)
    print("Готово.")


if __name__ == "__main__":
    main()