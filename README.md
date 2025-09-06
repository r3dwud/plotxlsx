# plotxlsx

Приложение на Python для визуализации данных из Excel.

## Функционал
- Загрузка `sample_data.xlsx`
- Рисует график
- Клик по графику — выделяет ближайшую точку и показывает её в списке ближайших

## Запуск
1. Python 3.11+
2. Установить зависимости: `pip install -r requirements.txt`
3. Запустить: `python src/app.py`
4. Или использовать .exe

## Требования
- dearpygui==2.1.0
- pandas
- numpy
- openpyxl
