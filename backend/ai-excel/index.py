"""
ИИ-аналитик для Excel: принимает контекст файлов и задание,
возвращает текстовый ответ + данные нового листа.
Также конвертирует PDF в изображения страниц (action=pdf_to_images).
"""

import os
import json
import re
import base64
import uuid

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
4. "cell_styles[].changes[].row" и "col" — 0-based индексы строки и столбца. Данные файла передаются с колонкой ROW — это и есть row-индекс для cell_styles
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

CORS = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Access-Control-Max-Age": "86400",
}


def extract_json(raw: str) -> dict:
    """Извлекает JSON из ответа даже если модель добавила лишний текст"""
    raw = raw.strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    raw = raw.strip()
    try:
        return json.loads(raw)
    except Exception:
        pass
    match = re.search(r'\{.*\}', raw, re.DOTALL)
    if match:
        try:
            return json.loads(match.group())
        except Exception:
            pass
    return {"text": raw or "ИИ вернул пустой ответ"}


def handle_pdf_to_images(body: dict) -> dict:
    """Конвертирует PDF (base64) в PNG-изображения страниц, сохраняет в S3, возвращает URL."""
    import fitz  # PyMuPDF
    import boto3

    pdf_b64 = body.get("pdf_b64", "")
    if not pdf_b64:
        return {"statusCode": 400, "headers": {**CORS}, "body": json.dumps({"error": "pdf_b64 required"})}

    dpi = min(int(body.get("dpi", 150)), 200)
    max_pages = min(int(body.get("max_pages", 50)), 100)

    pdf_bytes = base64.b64decode(pdf_b64)
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    total_pages = len(doc)
    pages_to_render = min(total_pages, max_pages)

    s3 = boto3.client(
        "s3",
        endpoint_url="https://bucket.poehali.dev",
        aws_access_key_id=os.environ["AWS_ACCESS_KEY_ID"],
        aws_secret_access_key=os.environ["AWS_SECRET_ACCESS_KEY"],
    )
    access_key = os.environ["AWS_ACCESS_KEY_ID"]
    session_id = str(uuid.uuid4())[:8]
    matrix = fitz.Matrix(dpi / 72, dpi / 72)
    page_urls = []

    for page_num in range(pages_to_render):
        page = doc[page_num]
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        img_bytes = pix.tobytes("png")
        key = f"pdf_pages/{session_id}/page_{page_num + 1:03d}.png"
        s3.put_object(Bucket="files", Key=key, Body=img_bytes, ContentType="image/png")
        url = f"https://cdn.poehali.dev/projects/{access_key}/bucket/{key}"
        page_urls.append(url)

    doc.close()

    return {
        "statusCode": 200,
        "headers": {**CORS},
        "body": json.dumps({
            "page_urls": page_urls,
            "total_pages": total_pages,
            "rendered_pages": pages_to_render,
            "session_id": session_id,
        }),
    }


def handler(event: dict, context) -> dict:
    """Обработка ИИ-запроса к Excel-данным или конвертация PDF в изображения"""

    if event.get("httpMethod") == "OPTIONS":
        return {"statusCode": 200, "headers": {**CORS}, "body": ""}

    body = json.loads(event.get("body") or "{}")

    # Роутинг по action
    if body.get("action") == "pdf_to_images":
        return handle_pdf_to_images(body)

    # ── Основная логика: ИИ по Excel ──
    prompt = body.get("prompt", "").strip()
    files_context = body.get("files_context", [])
    images = body.get("images", [])  # [{name, data (base64), mime}]

    api_key = body.get("api_key") or os.environ.get("OPENAI_API_KEY", "")
    base_url = body.get("base_url") or DEFAULT_BASE_URL
    model = body.get("model") or DEFAULT_MODEL

    if not prompt and not images:
        return {
            "statusCode": 400,
            "headers": {**CORS},
            "body": json.dumps({"error": "Пустой запрос"}, ensure_ascii=False),
        }

    if not api_key:
        return {
            "statusCode": 400,
            "headers": {**CORS},
            "body": json.dumps({"error": "API ключ не указан"}, ensure_ascii=False),
        }

    client = OpenAI(api_key=api_key, base_url=base_url, timeout=25.0)

    context_parts = []
    for i, fc in enumerate(files_context):
        role_label = ""
        if fc.get("role") == "main":
            role_label = " [ОСНОВНОЙ ФАЙЛ]"
        elif fc.get("role") == "reference":
            role_label = " [ОБРАЗЕЦ/ШАБЛОН]"

        context_parts.append(f"=== Файл {i} «{fc['name']}»{role_label} ===")
        for sheet in fc.get("sheets", []):
            is_active = sheet.get("active", False)
            total = sheet.get("total_rows", "?")
            marker = " [АКТИВНЫЙ ЛИСТ — полные данные]" if is_active else f" [краткий просмотр, всего {total} строк]"
            context_parts.append(f"--- Лист «{sheet['name']}»{marker} ---")
            context_parts.append(sheet.get("preview", ""))

    text_block = f"ДАННЫЕ ФАЙЛОВ:\n{chr(10).join(context_parts)}\n\nЗАДАНИЕ: {prompt or '(см. изображения)'}\n\nОтветь ТОЛЬКО JSON."

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

    try:
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_content},
            ],
            max_tokens=2000,
            temperature=0.1,
        )
    except Exception as e:
        err_msg = str(e)
        if "timeout" in err_msg.lower() or "timed out" in err_msg.lower():
            err_msg = "Модель не успела ответить за отведённое время. Попробуй ещё раз или выбери более быструю модель (например GPT-4o mini или Gemini Flash)."
        return {
            "statusCode": 200,
            "headers": {**CORS, "Content-Type": "application/json"},
            "body": json.dumps({"text": f"⚠️ {err_msg}"}, ensure_ascii=False),
        }

    raw = response.choices[0].message.content or "{}"
    print(f"[AI RAW RESPONSE]: {raw[:500]}")
    result = extract_json(raw)

    return {
        "statusCode": 200,
        "headers": {**CORS, "Content-Type": "application/json"},
        "body": json.dumps(result, ensure_ascii=False),
    }
