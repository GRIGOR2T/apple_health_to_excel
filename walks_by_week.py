from pathlib import Path
from lxml import etree as ET
import datetime as dt
import pandas as pd
from tqdm import tqdm
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
import math

XML_FILE = Path("export.xml")
START_DATE = dt.date(2025, 8, 1)  # учитывать только тренировки с августа 2025

# только ходьба
WORKOUT_TYPES = {
    "HKWorkoutActivityTypeWalking",
}

MONTHS_RU = {
    1: "января",
    2: "февраля",
    3: "марта",
    4: "апреля",
    5: "мая",
    6: "июня",
    7: "июля",
    8: "августа",
    9: "сентября",
    10: "октября",
    11: "ноября",
    12: "декабря",
}


def parse_dt(s: str) -> dt.datetime:
    # формат в export.xml: "2025-11-30 09:48:55 +0100"
    return dt.datetime.strptime(s[:19], "%Y-%m-%d %H:%M:%S")


def distance_value_to_km(value: float, unit: str | None) -> float:
    if unit in ("m", "meter", "meters"):
        return value / 1000.0
    if unit in ("km", "kilometer", "kilometers", None):
        return value
    if unit in ("mi", "mile", "miles"):
        return value * 1.60934
    return value


def get_distance_km_from_workout(elem: ET._Element) -> float:
    """
    Дистанция из Workout:
      1) totalDistance / totalDistanceUnit
      2) либо из WorkoutStatistics.
    Возвращаем в км.
    """
    dist = None
    unit = None

    # 1) атрибуты Workout
    dist_str = elem.get("totalDistance")
    unit = elem.get("totalDistanceUnit")

    if dist_str:
        try:
            dist = float(dist_str)
        except Exception:
            dist = None

    # 2) если нет — смотрим WorkoutStatistics
    if dist is None:
        for ws in elem.findall("WorkoutStatistics"):
            t = ws.get("type")
            if t == "HKQuantityTypeIdentifierDistanceWalkingRunning":
                s = ws.get("sum")
                u = ws.get("unit")
                try:
                    dist = float(s)
                    unit = u
                    break
                except Exception:
                    continue

    if dist is None:
        return 0.0

    return distance_value_to_km(dist, unit)


def get_duration_min_from_workout(elem: ET._Element) -> float:
    """
    Длительность тренировки в минутах.
    """
    dur_str = elem.get("duration")
    dur_unit = elem.get("durationUnit")

    if not dur_str:
        return 0.0

    try:
        dur = float(dur_str)
    except Exception:
        return 0.0

    if dur_unit in (None, "min", "minute", "minutes"):
        return dur
    if dur_unit in ("hr", "hour", "hours"):
        return dur * 60.0
    if dur_unit in ("sec", "second", "seconds"):
        return dur / 60.0

    return dur


def parse_workouts(xml_path: Path) -> pd.DataFrame:
    """
    Достаём все тренировки-ходьбы (только с START_DATE и позже).
    """
    rows = []

    print("Читаю XML (Workout):", xml_path.name)
    context = ET.iterparse(str(xml_path), events=("end",), tag="Workout")

    for _, elem in tqdm(context, desc="Парсинг тренировок", unit="rec"):
        w_type = elem.get("workoutActivityType")
        if w_type not in WORKOUT_TYPES:
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
        if date < START_DATE:
            elem.clear()
            continue

        distance_km = get_distance_km_from_workout(elem)
        duration_min = get_duration_min_from_workout(elem)

        year, week, _ = date.isocalendar()

        rows.append(
            {
                "year": year,
                "week": week,
                "date": date,
                "distance_km": distance_km,
                "duration_min": duration_min,
            }
        )

        elem.clear()

    del context

    if not rows:
        return pd.DataFrame(
            columns=["year", "week", "date", "distance_km", "duration_min"]
        )

    return pd.DataFrame(rows)


def aggregate_weekly(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Недельная агрегация по тренировкам ходьбы.
    """
    if df_raw.empty:
        return pd.DataFrame(
            columns=[
                "Год",
                "Неделя",
                "Дистанция, км",
                "Время, мин",
                "Тренировок",
                "Время, ч",
                "Средний темп, мин/км",
                "Начало недели",
                "Конец недели",
            ]
        )

    grp = (
        df_raw.groupby(["year", "week"])
        .agg(
            distance_km=("distance_km", "sum"),
            duration_min=("duration_min", "sum"),
            workouts=("date", "count"),
            week_start=("date", "min"),
            week_end=("date", "max"),
        )
        .reset_index()
    )

    grp["time_hours"] = grp["duration_min"] / 60.0
    grp["avg_pace_min_per_km"] = grp.apply(
        lambda row: row["duration_min"] / row["distance_km"]
        if row["distance_km"] > 0
        else None,
        axis=1,
    )

    grp = grp[
        [
            "year",
            "week",
            "distance_km",
            "duration_min",
            "workouts",
            "time_hours",
            "avg_pace_min_per_km",
            "week_start",
            "week_end",
        ]
    ]

    grp = grp.rename(
        columns={
            "year": "Год",
            "week": "Неделя",
            "distance_km": "Дистанция, км",
            "duration_min": "Время, мин",
            "workouts": "Тренировок",
            "time_hours": "Время, ч",
            "avg_pace_min_per_km": "Средний темп, мин/км",
            "week_start": "Начало недели",
            "week_end": "Конец недели",
        }
    )

    # сортировка от новых недель к старым
    grp = grp.sort_values(["Год", "Неделя"], ascending=False).reset_index(drop=True)
    return grp


# ---------- форматирование листа 1 ----------

def minutes_to_h_m(minutes: float) -> str:
    if minutes is None or (isinstance(minutes, float) and math.isnan(minutes)):
        return ""
    total_min = int(round(minutes))
    h = total_min // 60
    m = total_min % 60
    return f"{h} ч {m:02d} мин"


def hours_to_h_m(hours: float) -> str:
    if hours is None or (isinstance(hours, float) and math.isnan(hours)):
        return ""
    return minutes_to_h_m(hours * 60.0)


def pace_to_str(pace_min_per_km: float) -> str:
    if pace_min_per_km is None or (isinstance(pace_min_per_km, float) and math.isnan(pace_min_per_km)):
        return ""
    total_sec = int(round(pace_min_per_km * 60))
    m = total_sec // 60
    s = total_sec % 60
    return f"{m}:{s:02d} мин/км"


def distance_to_km_m(dist_km: float) -> str:
    if dist_km is None or (isinstance(dist_km, float) and math.isnan(dist_km)):
        return ""
    dist_km = round(dist_km, 2)
    km_int = int(dist_km)
    meters = int(round((dist_km - km_int) * 1000))
    if meters == 1000:
        km_int += 1
        meters = 0
    return f"{km_int} км {meters} м"


def date_to_ru_str(d: dt.date | dt.datetime | pd.Timestamp) -> str:
    if pd.isna(d):
        return ""
    if isinstance(d, pd.Timestamp):
        d = d.date()
    if isinstance(d, dt.datetime):
        d = d.date()
    month_name = MONTHS_RU.get(d.month, "")
    # формат дд/месяц/год
    return f"{d.day:02d}/{month_name}/{d.year}"


def format_weekly(df: pd.DataFrame) -> pd.DataFrame:
    """
    Формат под требования + порядок колонок.
    """
    df = df.copy()

    # убираем Год и Неделя из финальной таблицы
    if "Год" in df.columns:
        df = df.drop(columns=["Год"])
    if "Неделя" in df.columns:
        df = df.drop(columns=["Неделя"])

    # преобразуем поля
    df["Дистанция, км"] = df["Дистанция, км"].apply(distance_to_km_m)
    df["Время, мин"] = df["Время, мин"].apply(minutes_to_h_m)
    df["Время, ч"] = df["Время, ч"].apply(hours_to_h_m)
    df["Средний темп, мин/км"] = df["Средний темп, мин/км"].apply(pace_to_str)

    df["Начало недели"] = df["Начало недели"].apply(date_to_ru_str)
    df["Конец недели"] = df["Конец недели"].apply(date_to_ru_str)

    # порядок колонок: сначала даты
    cols_order = [
        "Начало недели",
        "Конец недели",
        "Дистанция, км",
        "Время, мин",
        "Тренировок",
        "Время, ч",
        "Средний темп, мин/км",
    ]
    df = df[cols_order]

    # плавающие (если ещё остались) — до 2 знаков
    for col in df.columns:
        if pd.api.types.is_float_dtype(df[col]):
            df[col] = df[col].round(2)

    return df


def output_filename() -> str:
    now = dt.datetime.now().strftime("%Y%m%d_%H%M")
    return f"weekly_walk_summary_{now}.xlsx"


def autoformat_sheet(ws):
    """
    Общий автоформат: крупный шрифт, автоширина.
    """
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


def save_excel(weekly_df: pd.DataFrame) -> None:
    out_name = output_filename()
    print("Сохраняю в Excel:", out_name)

    with pd.ExcelWriter(out_name, engine="openpyxl") as writer:
        weekly_df.to_excel(writer, sheet_name="Weekly_walks", index=False)
        ws = writer.sheets["Weekly_walks"]
        ws.freeze_panes = "A2"
        autoformat_sheet(ws)


def main():
    df_workouts = parse_workouts(XML_FILE)
    print("Найдено тренировок ходьбы:", len(df_workouts))

    weekly = aggregate_weekly(df_workouts)
    weekly_formatted = format_weekly(weekly)

    print("Недель в сводке:", len(weekly_formatted))

    save_excel(weekly_formatted)
    print("Готово.")


if __name__ == "__main__":
    main()