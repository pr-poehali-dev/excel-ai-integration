"""
ИИ-аналитик для Excel: принимает контекст файлов и задание,
возвращает текстовый ответ + (опционально) данные для нового листа.
"""

import os
import json
from openai import OpenAI

client = OpenAI(
    api_key=os.environ["OPENAI_API_KEY"],
    base_url="https://routerai.ru/api/v1",
)

SYSTEM_PROMPT = """Ты — профессиональный аналитик данных и эксперт по Excel.
Тебе передают содержимое одного или нескольких Excel-файлов в виде текста (TSV-формат), а также задание пользователя.

Твоя задача:
1. Выполнить задание, используя данные из файлов
2. Вернуть JSON-ответ строго в следующем формате:

{
  "text": "Объяснение что сделано и краткий итог (на русском)",
  "new_sheet": {
    "file_index": 0,
    "sheet_name": "Название_листа",
    "data": [["заголовок1","заголовок2",...],[значение1,значение2,...],...]
  }
}

Правила:
- new_sheet включай ТОЛЬКО если нужно создать новый лист с данными
- file_index — индекс файла из списка (0 = первый файл), обычно основной файл
- data — двумерный массив, первая строка — заголовки
- В text используй ** для выделения ключевых цифр и «кавычки» для названий
- Если задание только аналитическое (без создания листа) — new_sheet не включай
- Максимум 500 строк в new_sheet.data
- Отвечай ТОЛЬКО JSON, без markdown-обёртки"""


def handler(event: dict, context) -> dict:
    """Обработка ИИ-запроса к Excel-данным"""

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
    prompt = body.get("prompt", "").strip()
    files_context = body.get("files_context", [])

    if not prompt:
        return {
            "statusCode": 400,
            "headers": {"Access-Control-Allow-Origin": "*"},
            "body": json.dumps({"error": "Пустой запрос"}, ensure_ascii=False),
        }

    # Формируем контекст из файлов
    context_parts = []
    for i, fc in enumerate(files_context):
        role_label = ""
        if fc.get("role") == "main":
            role_label = " [ОСНОВНОЙ ФАЙЛ]"
        elif fc.get("role") == "reference":
            role_label = " [ОБРАЗЕЦ/ШАБЛОН]"

        context_parts.append(f"=== Файл {i} «{fc['name']}»{role_label} ===")
        for sheet in fc.get("sheets", []):
            context_parts.append(f"--- Лист «{sheet['name']}» ---")
            context_parts.append(sheet.get("preview", ""))

    user_message = f"ДАННЫЕ ФАЙЛОВ:\n{chr(10).join(context_parts)}\n\nЗАДАНИЕ ПОЛЬЗОВАТЕЛЯ:\n{prompt}"

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_message},
        ],
        max_tokens=4000,
        temperature=0.3,
        response_format={"type": "json_object"},
    )

    raw = response.choices[0].message.content or "{}"
    result = json.loads(raw)

    return {
        "statusCode": 200,
        "headers": {"Access-Control-Allow-Origin": "*", "Content-Type": "application/json"},
        "body": json.dumps(result, ensure_ascii=False),
    }