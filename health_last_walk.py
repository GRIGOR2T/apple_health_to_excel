from pathlib import Path
import datetime as dt

import pandas as pd
from lxml import etree as ET
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font


EXPORT_XML = Path("export.xml")

# Зоны пульса – как у тебя
ZONES = [
    ("Zone 1", 0, 114),
    ("Zone 2", 115, 134),
    ("Zone 3", 135, 149),
    ("Zone 4", 150, 169),
    ("Zone 5", 170, 10_000),
]

# Берём только записи с часов (как в других скриптах)
PRIMARY_DEVICE_MARK = "Watch"


def is_watch_source(elem: ET._Element) -> bool:
    """
    Источник — Apple Watch (по sourceName/device, с учётом NBSP).
    """
    src = (elem.get("sourceName") or "") + " " + (elem.get("device") or "")
    src = src.replace("\u00a0", " ")
    return PRIMARY_DEVICE_MARK in src


def parse_apple_datetime(s: str) -> dt.datetime:
    # формат: '2021-09-13 14:40:55 +0100'
    return dt.datetime.strptime(s[:19], "%Y-%m-%d %H:%M:%S")


def find_last_walking_workout():
    """
    Находим последнюю Walking-тренировку и
    сразу берём duration, kcal, дистанцию, набор высоты.
    """
    last_start = None
    best = None

    context = ET.iterparse(str(EXPORT_XML), events=("end",), tag="Workout")
    for _, elem in context:
        wtype = elem.get("workoutActivityType")
        if wtype == "HKWorkoutActivityTypeWalking":
            start = parse_apple_datetime(elem.get("startDate"))
            end = parse_apple_datetime(elem.get("endDate"))

            if last_start is None or start > last_start:
                last_start = start

                duration_min = float(elem.get("duration", "0"))

                active_kcal = None
                basal_kcal = None
                distance = None
                distance_unit = None
                elevation_m = None

                for child in elem:
                    if child.tag == "WorkoutStatistics":
                        stype = child.get("type")
                        if stype == "HKQuantityTypeIdentifierActiveEnergyBurned":
                            active_kcal = float(child.get("sum", "0"))
                        elif stype == "HKQuantityTypeIdentifierBasalEnergyBurned":
                            basal_kcal = float(child.get("sum", "0"))
                        elif stype == "HKQuantityTypeIdentifierDistanceWalkingRunning":
                            distance = float(child.get("sum", "0"))
                            distance_unit = child.get("unit")
                    elif child.tag == "MetadataEntry":
                        key = child.get("key")
                        if key == "HKElevationAscended":
                            val = child.get("value") or "0"
                            num = val.split()[0]
                            try:
                                elevation_m = float(num) / 100.0  # см → м
                            except ValueError:
                                elevation_m = None

                best = {
                    "start": start,
                    "end": end,
                    "duration_min": duration_min,
                    "active_kcal": active_kcal,
                    "basal_kcal": basal_kcal,
                    "distance": distance,
                    "distance_unit": distance_unit,
                    "elevation_m": elevation_m,
                }

        elem.clear()
    del context

    if best is None:
        raise RuntimeError("Walking-тренировок в export.xml не найдено")

    if best["basal_kcal"] is None:
        best["basal_kcal"] = 0.0
    if best["active_kcal"] is None:
        best["active_kcal"] = 0.0

    # Нормализуем дистанцию к км
    dist = best["distance"] or 0.0
    unit = best["distance_unit"]
    if unit == "mi":
        dist = dist * 1.60934
    best["distance_km"] = dist

    return best


def collect_hr_and_dist(start: dt.datetime, end: dt.datetime, total_dist_km: float):
    """
    Собираем:
    - пульс во время тренировки
    - расстояние во времени (для сплитов)

    Теперь берём только записи с WATCH,
    чтобы не было дублей с айфона.
    """
    hr_samples = []              # (time, bpm)
    dist_samples_raw = []        # (time, value_km)

    context = ET.iterparse(str(EXPORT_XML), events=("end",), tag="Record")
    for _, elem in context:
        # фильтруем источник
        if not is_watch_source(elem):
            elem.clear()
            continue

        rtype = elem.get("type")

        s = parse_apple_datetime(elem.get("startDate"))
        e = parse_apple_datetime(elem.get("endDate"))

        if e <= start or s >= end:
            elem.clear()
            continue

        if rtype == "HKQuantityTypeIdentifierHeartRate":
            v = float(elem.get("value"))
            hr_samples.append((s, v))

        elif rtype in (
            "HKQuantityTypeIdentifierDistanceWalkingRunning",
            "HKQuantityTypeIdentifierDistanceWalking",
        ):
            v = float(elem.get("value"))
            unit = elem.get("unit")
            # Приводим к км
            if unit == "mi":
                v *= 1.60934
            dist_samples_raw.append((e, v))

        elem.clear()
    del context

    if not hr_samples:
        raise RuntimeError("Не найдено данных по пульсу в интервале тренировки")

    hr_samples.sort(key=lambda x: x[0])
    dist_samples_raw.sort(key=lambda x: x[0])

    # Если вдруг нет расстояний – вернём пустой список
    if not dist_samples_raw:
        return hr_samples, []

    # Авто-определение delta / cumulative
    vals = [v for _, v in dist_samples_raw]
    max_val = max(vals)
    sum_val = sum(vals)

    diff_if_delta = abs(sum_val - total_dist_km)
    diff_if_cum = abs(max_val - total_dist_km)

    mode = "delta" if diff_if_delta <= diff_if_cum else "cumulative"

    dist_samples = []
    if mode == "delta":
        cum = 0.0
        for t, v in dist_samples_raw:
            cum += v
            dist_samples.append((t, cum))
    else:  # cumulative
        dist_samples = [(t, v) for t, v in dist_samples_raw]

    return hr_samples, dist_samples


def compute_avg_hr(hr_samples):
    vals = [v for _, v in hr_samples]
    return sum(vals) / len(vals)


def compute_zones(hr_samples):
    zone_seconds = {name: 0 for name, _, _ in ZONES}

    for i in range(len(hr_samples) - 1):
        t0, hr = hr_samples[i]
        t1, _ = hr_samples[i + 1]
        dt_sec = (t1 - t0).total_seconds()
        for name, lo, hi in ZONES:
            if lo <= hr <= hi:
                zone_seconds[name] += dt_sec
                break

    zone_str = {}
    for name, secs in zone_seconds.items():
        secs = int(secs)
        m, s = divmod(secs, 60)
        zone_str[name] = f"{m:02d}:{s:02d}"
    return zone_str


def format_timedelta(td: dt.timedelta) -> str:
    total = int(td.total_seconds())
    h = total // 3600
    m = (total % 3600) // 60
    s = total % 60
    return f"{h}:{m:02d}:{s:02d}"


def format_pace(duration: dt.timedelta, distance_km: float) -> str:
    if not distance_km or duration.total_seconds() <= 0:
        return ""
    sec_per_km = duration.total_seconds() / distance_km
    m = int(sec_per_km // 60)
    s = int(sec_per_km % 60)
    return f"{m}'{s:02d}\"/KM"


def autoformat_sheet(ws):
    for col_idx, col in enumerate(ws.columns, start=1):
        max_len = 0
        for cell in col:
            cell.font = Font(size=28)
            if cell.value is not None:
                val_str = str(cell.value)
                if len(val_str) > max_len:
                    max_len = len(val_str)
        # Делаем колонки заметно шире: небольшой запас + минимальная ширина
        padding = 10
        min_width = 26
        width = max(max_len + padding, min_width)
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def compute_splits(hr_samples, dist_samples, start: dt.datetime, total_dist_km: float):
    """
    Строим сплиты:
      - по каждому полному километру (1,2,3,...)
      - + при необходимости последний неполный кусок
    """
    if not dist_samples:
        return []

    EPS = 0.05  # допуск по длине отрезка (50 м)

    # сколько полных километров считаем сплитами
    num_full = int(total_dist_km + EPS)  # для 12.01 -> 12

    bounds = [k for k in range(1, num_full + 1)]

    # если после этого остался заметный хвост (больше ~50 м) — добавим его как отдельный сплит
    leftover = total_dist_km - num_full
    if leftover > EPS:
        bounds.append(total_dist_km)

    if not bounds:
        return []

    times = [t for t, _ in dist_samples]
    dists = [d for _, d in dist_samples]

    splits = []
    prev_t = start
    hr_samples_sorted = sorted(hr_samples, key=lambda x: x[0])

    idx = 0
    n = len(times)
    prev_dist = 0.0

    for i, target in enumerate(bounds, start=1):
        # ищем участок, где cum_dist пересекает target
        while idx < n and dists[idx] < target:
            prev_t = times[idx]
            prev_dist = dists[idx]
            idx += 1
        if idx == n:
            break

        cur_t = times[idx]
        cur_dist = dists[idx]

        # линейная интерполяция момента достижения нужной дистанции
        if cur_dist == prev_dist:
            t_target = cur_t
        else:
            frac = (target - prev_dist) / (cur_dist - prev_dist)
            dt_sec = (cur_t - prev_t).total_seconds()
            t_target = prev_t + dt.timedelta(seconds=dt_sec * frac)

        # границы сплита по времени
        t_start_split = splits[-1]["_t_end"] if splits else start
        t_end_split = t_target

        split_duration = t_end_split - t_start_split
        sec = int(split_duration.total_seconds())
        pace_min = sec // 60
        pace_sec = sec % 60

        # средний пульс внутри сплита
        hr_vals = [v for (t, v) in hr_samples_sorted if t_start_split <= t <= t_end_split]
        avg_hr = int(round(sum(hr_vals) / len(hr_vals))) if hr_vals else 0

        splits.append(
            {
                "KM": i,
                "Time": t_target.strftime("%H:%M"),
                "Pace": f"{pace_min}'{pace_sec:02d}\"/KM",
                "Heart Rate": f"{avg_hr} BPM",
                "_t_end": t_target,
            }
        )

    for sp in splits:
        sp.pop("_t_end", None)

    return splits


def main():
    print("Поиск последней Walking-тренировки…")
    info = find_last_walking_workout()

    start = info["start"]
    end = info["end"]
    duration_td = dt.timedelta(minutes=info["duration_min"])
    elapsed_td = end - start

    distance_km = info["distance_km"]
    active_kcal = info["active_kcal"]
    basal_kcal = info["basal_kcal"]
    total_kcal = active_kcal + basal_kcal
    elevation_m = info["elevation_m"]

    print(f"Нашёл тренировку: {start} — {end}")

    print("Сбор пульса и дистанции…")
    hr_samples, dist_samples = collect_hr_and_dist(start, end, distance_km)

    avg_hr = compute_avg_hr(hr_samples)
    zones = compute_zones(hr_samples)
    splits = compute_splits(hr_samples, dist_samples, start, distance_km)

    avg_pace_str = format_pace(duration_td, distance_km)

    date_str = start.strftime("%A, %B %d")

    # Левый блок
    left_rows = [
        ["Workout Time",        format_timedelta(duration_td),             ""],
        ["Elapsed Time",        format_timedelta(elapsed_td),              ""],
        ["Active Kilocalories", f"{int(round(active_kcal))} KCAL",        ""],
        ["Total Kilocalories",  f"{int(round(total_kcal))} KCAL",         ""],
        ["Avg. Pace",           avg_pace_str,                             ""],
        ["Distance",            f"{distance_km:.2f} KM" if distance_km else "", ""],
        ["Elevation Gain",      f"{elevation_m:.0f} M" if elevation_m is not None else "", ""],
        ["Avg. Heart Rate",     f"{int(round(avg_hr))} BPM",              ""],
    ]

    for name, lo, hi in ZONES:
        if name == "Zone 1":
            rng = f"<{hi} BPM"
        elif name == "Zone 5":
            rng = f"{lo}+ BPM"
        else:
            rng = f"{lo}–{hi} BPM"
        left_rows.append([name, zones[name], rng])

    left_df = pd.DataFrame(left_rows, columns=[date_str, "", ""])

    # Правый блок – сплиты
    if splits:
        right_df = pd.DataFrame(splits, columns=["KM", "Time", "Pace", "Heart Rate"])
    else:
        right_df = pd.DataFrame(columns=["KM", "Time", "Pace", "Heart Rate"])

    out_name = "full_last_walk.xlsx"
    with pd.ExcelWriter(out_name, engine="openpyxl") as writer:
        left_df.to_excel(writer, sheet_name="Sheet1", index=False, startrow=0, startcol=0)
        right_df.to_excel(writer, sheet_name="Sheet1", index=False, startrow=0, startcol=3)
        ws = writer.sheets["Sheet1"]
        ws.freeze_panes = "A2"
        autoformat_sheet(ws)

    print(f"Готово: {out_name}")


if __name__ == "__main__":
    main()