"""
Добавляет лист с диаграммой в существующий xlsx-файл через openpyxl.
Принимает: xlsx_b64 (оригинальный файл), chart_sheet (имя листа с данными), chart_type.
Возвращает base64-encoded xlsx с добавленным графиком.
"""

import json
import base64
import io
from openpyxl import load_workbook
from openpyxl.chart import PieChart, BarChart, LineChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.utils import get_column_letter


CHART_COLORS = [
    "34D399", "FBBF24", "60A5FA", "A78BFA",
    "F87171", "2DD4BF", "FB923C", "818CF8",
]


def handler(event: dict, context) -> dict:
    """Добавляет диаграмму в существующий xlsx не трогая другие листы и стили"""

    if event.get("httpMethod") == "OPTIONS":
        return {
            "statusCode": 200,
            "headers": {
                "Access-Control-Allow-Origin": "*",
                "Access-Control-Allow-Methods": "POST, OPTIONS",
                "Access-Control-Allow-Headers": "Content-Type",
                "Access-Control-Max-Age": "86400",
            },
            "body": "",
        }

    body = json.loads(event.get("body") or "{}")
    xlsx_b64 = body.get("xlsx_b64")          # base64 оригинального файла
    chart_sheet_name = body.get("chart_sheet")  # имя листа с данными для графика
    chart_type = body.get("chart_type", "pie")  # pie | bar | line
    chart_title = body.get("chart_title", "График")

    if not xlsx_b64 or not chart_sheet_name:
        return {
            "statusCode": 400,
            "headers": {"Access-Control-Allow-Origin": "*"},
            "body": json.dumps({"error": "Нужны xlsx_b64 и chart_sheet"}, ensure_ascii=False),
        }

    # Загружаем оригинальный файл — все стили, форматы, объединения сохраняются
    try:
        raw = base64.b64decode(xlsx_b64)
        wb = load_workbook(io.BytesIO(raw))
    except Exception as e:
        return {
            "statusCode": 400,
            "headers": {"Access-Control-Allow-Origin": "*"},
            "body": json.dumps({"error": f"Ошибка чтения файла: {e}"}, ensure_ascii=False),
        }

    # Находим лист с данными для графика
    if chart_sheet_name not in wb.sheetnames:
        return {
            "statusCode": 400,
            "headers": {"Access-Control-Allow-Origin": "*"},
            "body": json.dumps({"error": f"Лист «{chart_sheet_name}» не найден"}, ensure_ascii=False),
        }

    chart_ws = wb[chart_sheet_name]
    data_rows = chart_ws.max_row
    data_cols = chart_ws.max_column

    if data_rows < 2 or data_cols < 2:
        return {
            "statusCode": 400,
            "headers": {"Access-Control-Allow-Origin": "*"},
            "body": json.dumps({"error": "Недостаточно данных для графика (нужно минимум 2 строки и 2 колонки)"}, ensure_ascii=False),
        }

    # Метки (первая колонка, строки 2..N)
    labels = Reference(chart_ws, min_col=1, min_row=2, max_row=data_rows)
    # Значения (вторая колонка, строки 1..N — заголовок + данные)
    values = Reference(chart_ws, min_col=2, min_row=1, max_row=data_rows)

    if chart_type == "pie":
        chart = PieChart()
        chart.title = chart_title
        chart.style = 10
        chart.add_data(values, titles_from_data=True)
        chart.set_categories(labels)
        # Раскрашиваем сегменты
        series = chart.series[0]
        for idx in range(data_rows - 1):
            pt = DataPoint(idx=idx)
            pt.graphicalProperties.solidFill = CHART_COLORS[idx % len(CHART_COLORS)]
            series.dPt.append(pt)

    elif chart_type == "bar":
        chart = BarChart()
        chart.type = "col"
        chart.title = chart_title
        chart.style = 10
        chart.add_data(values, titles_from_data=True)
        chart.set_categories(labels)

    else:
        chart = LineChart()
        chart.title = chart_title
        chart.style = 10
        chart.add_data(values, titles_from_data=True)
        chart.set_categories(labels)

    chart.width = 18
    chart.height = 12

    # Добавляем график правее данных на том же листе
    anchor_col = get_column_letter(data_cols + 2)
    chart_ws.add_chart(chart, f"{anchor_col}2")

    # Сохраняем — оригинальные листы и стили не тронуты
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    b64 = base64.b64encode(buf.read()).decode()

    return {
        "statusCode": 200,
        "headers": {
            "Access-Control-Allow-Origin": "*",
            "Content-Type": "application/json",
        },
        "body": json.dumps({"xlsx_b64": b64}, ensure_ascii=False),
    }
