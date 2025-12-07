# Apple Health Walk / VO₂max Tools

Набор утилит для разбора `export.xml` из Apple Health и получения отчётов по:
- VO₂max
- весу + VO₂max
- ходьбе по тренировка́м
- общей дневной ходьбе (включая «тихую» вне тренировок)
- недельным объёмам из дневных данных
- последней тренировке ходьбы

Все скрипты ожидают файл `export.xml` в корне проекта.

## Зависимости

Python 3.x и пакеты:

- `lxml`
- `pandas`
- `openpyxl`
- `tqdm`

Установка (пример):

```bash
pip install lxml pandas openpyxl tqdm