"""
ИИ-аналитик для Excel: принимает контекст файлов и задание,
возвращает текстовый ответ + данные нового листа.
"""

import os
import json
import re
from openai import OpenAI

SYSTEM_PROMPT = """Ты — профессиональный аналитик данных и эксперт по Excel.
Тебе передают содержимое одного или нескольких Excel-файлов в виде текста (TSV), задание пользователя, и (опционально) изображения — скриншоты графиков, примеры оформления.

КРИТИЧЕСКИ ВАЖНО: ты ВСЕГДА отвечаешь ТОЛЬКО валидным JSON и НИЧЕМ КРОМЕ JSON. Никакого текста до или после. Никаких ```json``` обёрток.

Формат ответа:
{
  "text": "Краткое объяснение что сделано (на русском, 1-3 предложения)",
  "new_sheet": {
    "file_index": 0,
    "sheet_name": "Название_листа",
    "data": [["Заголовок1","Заголовок2",...],["значение1","значение2",...],...]
  },
  "cell_styles": [
    {
      "file_index": 0,
      "sheet_name": "Лист1",
      "changes": [
        {"row": 2, "col": 0, "bgColor": "FFFFFF00", "fontColor": "FF000000", "bold": true}
      ]
    }
  ]
}

ПРАВИЛА:
1. Поле "text" — ОБЯЗАТЕЛЬНО всегда
2. Поле "new_sheet" — включай когда нужно создать новый лист с данными/расчётами/графиком
3. Поле "cell_styles" — включай когда нужно покрасить/форматировать ячейки в СУЩЕСТВУЮЩЕМ листе, НЕ создавая новый лист
4. "cell_styles[].changes[].row" и "col" — 0-based индексы строки и столбца
5. "bgColor" — цвет заливки в формате AARRGGBB (например FF FFFF00 = жёлтый, FFFFA500 = оранжевый, FF92D050 = зелёный, FFFF0000 = красный)
6. "fontColor" — цвет текста в формате AARRGGBB
7. "bold": true/false — жирный шрифт
8. В "data" первая строка — заголовки, остальные — данные. Максимум 500 строк
9. НЕ пиши текст вне JSON. НЕ используй markdown
10. Все числа в data передавай как числа, текст — как строки
11. Если просят покрасить строки/ячейки — используй cell_styles, НЕ создавай новый лист
12. Если пользователь прикрепил изображение — используй его как образец при формировании данных"""

DEFAULT_BASE_URL = "https://routerai.ru/api/v1"
DEFAULT_MODEL = "deepseek/deepseek-chat"


def extract_json(raw: str) -> dict:
    """Извлекает JSON из ответа даже если модель добавила лишний текст"""
    raw = raw.strip()

    # Убираем markdown-обёртку
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    raw = raw.strip()

    # Пробуем напрямую
    try:
        return json.loads(raw)
    except Exception:
        pass

    # Ищем первый JSON-объект в тексте
    match = re.search(r'\{.*\}', raw, re.DOTALL)
    if match:
        try:
            return json.loads(match.group())
        except Exception:
            pass

    return {"text": raw or "ИИ вернул пустой ответ"}


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
    images = body.get("images", [])  # [{name, data (base64), mime}]

    api_key = body.get("api_key") or os.environ.get("OPENAI_API_KEY", "")
    base_url = body.get("base_url") or DEFAULT_BASE_URL
    model = body.get("model") or DEFAULT_MODEL

    if not prompt and not images:
        return {
            "statusCode": 400,
            "headers": {"Access-Control-Allow-Origin": "*"},
            "body": json.dumps({"error": "Пустой запрос"}, ensure_ascii=False),
        }

    if not api_key:
        return {
            "statusCode": 400,
            "headers": {"Access-Control-Allow-Origin": "*"},
            "body": json.dumps({"error": "API ключ не указан"}, ensure_ascii=False),
        }

    client = OpenAI(api_key=api_key, base_url=base_url)

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

    text_block = f"ДАННЫЕ ФАЙЛОВ:\n{chr(10).join(context_parts)}\n\nЗАДАНИЕ: {prompt or '(см. изображения)'}\n\nОтветь ТОЛЬКО JSON."

    # Строим content: текст + изображения (vision)
    if images:
        user_content = [{"type": "text", "text": text_block}]
        for img in images:
            mime = img.get("mime", "image/png")
            data = img.get("data", "")
            user_content.append({
                "type": "image_url",
                "image_url": {"url": f"data:{mime};base64,{data}"},
            })
    else:
        user_content = text_block

    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_content},
        ],
        max_tokens=6000,
        temperature=0.1,
    )

    raw = response.choices[0].message.content or "{}"
    print(f"[AI RAW RESPONSE]: {raw[:500]}")

    result = extract_json(raw)

    return {
        "statusCode": 200,
        "headers": {"Access-Control-Allow-Origin": "*", "Content-Type": "application/json"},
        "body": json.dumps(result, ensure_ascii=False),
    }