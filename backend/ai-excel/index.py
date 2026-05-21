"""
ИИ-аналитик для Excel: принимает контекст файлов и задание,
возвращает текстовый ответ + данные нового листа.
"""

import os
import json
import re
from openai import OpenAI

SYSTEM_PROMPT = """Ты — профессиональный аналитик данных и эксперт по Excel.
Тебе передают содержимое одного или нескольких Excel-файлов в виде текста (TSV), а также задание пользователя.

КРИТИЧЕСКИ ВАЖНО: ты ВСЕГДА отвечаешь ТОЛЬКО валидным JSON и НИЧЕМ КРОМЕ JSON. Никакого текста до или после. Никаких ```json``` обёрток.

Формат ответа:
{
  "text": "Краткое объяснение что сделано (на русском, 1-3 предложения)",
  "new_sheet": {
    "file_index": 0,
    "sheet_name": "Название_листа",
    "data": [["Заголовок1","Заголовок2",...],["значение1","значение2",...],...]
  }
}

ПРАВИЛА:
1. Поле "text" — ОБЯЗАТЕЛЬНО всегда
2. Поле "new_sheet" — включай ВСЕГДА когда задание подразумевает создание таблицы, расчётов, итогов, фильтрации, сортировки, графика (данные для графика), преобразования данных
3. "new_sheet" НЕ включай только если вопрос чисто аналитический без создания данных (например "сколько строк?" или "какие колонки есть?")
4. В "data" первая строка — заголовки, остальные — данные. Максимум 500 строк
5. "file_index" — индекс файла (0 = первый), куда добавить лист
6. НЕ пиши текст вне JSON. НЕ используй markdown. НЕ объясняй что ты собираешься сделать — просто сделай и верни JSON с результатом
7. Если просят "график" — создай лист с данными для графика (две колонки: метка и значение)
8. Все числа в data передавай как числа (не строки), текст — как строки"""

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

    api_key = body.get("api_key") or os.environ.get("OPENAI_API_KEY", "")
    base_url = body.get("base_url") or DEFAULT_BASE_URL
    model = body.get("model") or DEFAULT_MODEL

    if not prompt:
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

    user_message = f"ДАННЫЕ ФАЙЛОВ:\n{chr(10).join(context_parts)}\n\nЗАДАНИЕ: {prompt}\n\nОтветь ТОЛЬКО JSON."

    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_message},
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
