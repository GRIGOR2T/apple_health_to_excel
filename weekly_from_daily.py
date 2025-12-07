from pathlib import Path
import datetime as dt
import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font

INPUT_FILE = Path("daily_walk.xlsx")      # <-- имя твоего файла с листом Daily_walk


def autoformat_sheet(ws):
    for col_idx, column_cells in enumerate(ws.columns, start=1):
        max_len = 0
        for cell in column_cells:
            cell.font = Font(size=28)
            val = cell.value
            if val is None:
                continue
            text = str(val)
            max_len = max(max_len, len(text))
        ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 4


def main():
    # читаем лист Daily_walk
    df = pd.read_excel(INPUT_FILE, sheet_name="Daily_walk")

    # дата
    df["date"] = pd.to_datetime(df["Дата"], format="%m/%d/%Y")

    # фильтр: только с 1 октября 2025
    df = df[df["date"] >= dt.datetime(2025, 10, 1)]

    # год + ISO-неделя
    isocal = df["date"].dt.isocalendar()
    df["year"] = isocal["year"]
    df["week"] = isocal["week"]

    # агрегируем по неделям
    weekly = (
        df.groupby(["year", "week"])
        .agg(
            week_start=("date", "min"),
            week_end=("date", "max"),
            days_with_data=("date", "nunique"),
            total_walk_km=("Всего ходьба, км", "sum"),
            workout_walk_km=("Ходьба в тренировках, км", "sum"),
            nonwork_walk_km=("Ходьба вне тренировок, км", "sum"),
        )
        .reset_index()
    )

    # среднее в день по каждой неделе
    weekly["avg_total_km_per_day"] = weekly["total_walk_km"] / weekly["days_with_data"]
    weekly["avg_nonwork_km_per_day"] = weekly["nonwork_walk_km"] / weekly["days_with_data"]

    # доля «тихой» ходьбы
    weekly["nonwork_share_%"] = (
        weekly["nonwork_walk_km"] / weekly["total_walk_km"] * 100
    ).round(1)

    # сортировка от новых недель к старым
    weekly = weekly.sort_values(["year", "week"], ascending=False).reset_index(drop=True)

    # красивый формат дат
    weekly["Начало недели"] = weekly["week_start"].dt.strftime("%d.%m.%Y")
    weekly["Конец недели"] = weekly["week_end"].dt.strftime("%d.%m.%Y")

    # округление километров
    for col in ["total_walk_km", "workout_walk_km", "nonwork_walk_km",
                "avg_total_km_per_day", "avg_nonwork_km_per_day"]:
        weekly[col] = weekly[col].round(2)

    # финальный порядок колонок
    weekly = weekly[
        [
            "Начало недели",
            "Конец недели",
            "days_with_data",
            "total_walk_km",
            "workout_walk_km",
            "nonwork_walk_km",
            "avg_total_km_per_day",
            "avg_nonwork_km_per_day",
            "nonwork_share_%",
        ]
    ]

    weekly = weekly.rename(
        columns={
            "days_with_data": "Дней с данными",
            "total_walk_km": "Всего ходьба, км",
            "workout_walk_km": "Ходьба в тренировках, км",
            "nonwork_walk_km": "Ходьба вне тренировок, км",
            "avg_total_km_per_day": "Ср. всего км/день",
            "avg_nonwork_km_per_day": "Ср. вне тренировок км/день",
            "nonwork_share_%": "Доля тихой ходьбы, %",
        }
    )

    out_name = "weekly_from_daily_walk.xlsx"
    print("Сохраняю:", out_name)

    with pd.ExcelWriter(out_name, engine="openpyxl") as writer:
        weekly.to_excel(writer, sheet_name="Weekly_from_daily", index=False)
        ws = writer.sheets["Weekly_from_daily"]
        ws.freeze_panes = "A2"
        autoformat_sheet(ws)

    print("Готово. Недель:", len(weekly))


if __name__ == "__main__":
    main()