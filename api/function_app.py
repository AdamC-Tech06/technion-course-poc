import azure.functions as func
import csv
import json
import logging
import os
import re
import requests
from io import StringIO
from datetime import datetime, timedelta
from openai import OpenAI

app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)

CORS_HEADERS = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
}

DATE_PATTERN = re.compile(r'\b(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})\b')


def clean_env(val: str) -> str:
    if not val:
        return ""
    return re.sub(r'[^\x21-\x7E]', '', val)


def extract_dates_from_text(text: str) -> list:
    if not text:
        return []
    parsed_dates = []
    for match in DATE_PATTERN.finditer(str(text)):
        day, month, year = int(match.group(1)), int(match.group(2)), int(match.group(3))
        if year < 100:
            year += 2000
        try:
            parsed_dates.append(datetime(year, month, day))
        except ValueError:
            continue
    return parsed_dates


def check_course_date_status(row: dict, date_columns: list, cutoff_date: datetime) -> str:
    """
    בודק האם הקורס בתוקף מבחינת מועדי פתיחה:
    - מחזיר 'EXPIRED' אם עברו מעל 14 יום
    - מחזיר 'NO_SCHEDULE' אם רשום ללא צפי
    - מחזיר 'ACTIVE' אם התאריך עתידי או בטווח המותר
    """
    for col in date_columns:
        val = str(row.get(col, "")).strip()
        if "ללא צפי" in val or "לא ידוע" in val:
            return "NO_SCHEDULE"

    found_dates = []
    for col in date_columns:
        cell_val = str(row.get(col, "")).strip()
        if cell_val:
            found_dates.extend(extract_dates_from_text(cell_val))

    if not found_dates:
        for col, val in row.items():
            if any(k in col for k in ["פתיחה", "התחלה", "מועד", "תאריך"]):
                found_dates.extend(extract_dates_from_text(str(val)))

    if found_dates:
        if all(d < cutoff_date for d in found_dates):
            return "EXPIRED"

    all_text = " ".join(str(v) for v in row.values())
    deadline_match = re.search(r'הרשמה.*?עד\s*(\d{1,2})[/.](\d{1,2})', all_text)
    if deadline_match:
        d_day, d_month = int(deadline_match.group(1)), int(deadline_match.group(2))
        try:
            deadline_date = datetime(cutoff_date.year, d_month, d_day)
            if deadline_date < cutoff_date:
                return "EXPIRED"
        except ValueError:
            pass

    return "ACTIVE"


def load_courses_catalog() -> str:
    sheet_id = "1emkWYVjlrhC0m67gz7b3MUXzsG3RtlX-kbg8eNvzBtc"
    gid = "1525331068"

    csv_url = (
        f"https://docs.google.com/spreadsheets/d/"
        f"{sheet_id}/export?format=csv&gid={gid}"
    )

    response = requests.get(csv_url, timeout=30)
    response.raise_for_status()
    response.encoding = "utf-8-sig"

    reader = csv.DictReader(StringIO(response.text))
    if not reader.fieldnames:
        return ""

    clean_fieldnames = [col.strip() for col in reader.fieldnames if col]
    date_columns = [
        col for col in clean_fieldnames
        if any(keyword in col for keyword in ["פתיחה", "התחלה", "מועד", "תאריך", "ת."])
    ]

    today = datetime.now()
    cutoff_date = today - timedelta(days=14)

    catalog_blocks = []

    for raw_row in reader:
        row = {k.strip(): (v.strip() if isinstance(v, str) else "") for k, v in raw_row.items() if k}
        if not any(row.values()):
            continue

        status = check_course_date_status(row, date_columns, cutoff_date)

        row_lines = []
        # תיוג סטטוס תאריך עבור מודל ה-AI
        if status == "EXPIRED":
            row_lines.append("סטטוס פתיחה: פג תוקף (עברו מעל 14 יום ממועד הפתיחה - חובה לתת ציון 0 ולפסול)")
        elif status == "NO_SCHEDULE":
            row_lines.append("סטטוס פתיחה: ללא צפי פתיחה (חובה לתת ציון 0 ולפסול)")
        else:
            row_lines.append("סטטוס פתיחה: פעיל ורלוונטי להרשמה")

        for col in clean_fieldnames:
            val = row.get(col, "")
            if val:
                if ("לינק" in col or "קישור" in col or "דפי מידע" in col or "אתר" in col) and val != "לא זמין":
                    if not val.startswith("http://") and not val.startswith("https://"):
                        val = "https://" + val
                row_lines.append(f"{col}: {val}")

        if row_lines:
            catalog_blocks.append("\n".join(row_lines))

    return "\n\n" + ("\n" + "=" * 80 + "\n\n").join(catalog_blocks)


@app.route(route="match-courses", methods=["POST", "OPTIONS"])
def match_courses(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("Course matcher API received a request.")

    if req.method == "OPTIONS":
        return func.HttpResponse(status_code=200, headers=CORS_HEADERS)

    try:
        req_body = req.get_json()
    except Exception:
        return func.HttpResponse(
            json.dumps({"error": "Invalid JSON format in request body"}),
            status_code=400,
            headers=CORS_HEADERS,
            mimetype="application/json"
        )

    api_key = clean_env(os.environ.get("OPENAI_API_KEY", ""))
    base_url = clean_env(os.environ.get("OPENAI_BASE_URL", "https://api.groq.com/openai/v1"))
    model_name = clean_env(os.environ.get("OPENAI_MODEL", "llama-3.1-8b-instant"))

    if not api_key:
        return func.HttpResponse(
            json.dumps({"error": "Missing OPENAI_API_KEY configuration on Azure"}),
            status_code=500,
            headers=CORS_HEADERS,
            mimetype="application/json"
        )

    try:
        courses_catalog = load_courses_catalog()
        today_str = datetime.now().strftime("%d/%m/%Y")
        cutoff_str = (datetime.now() - timedelta(days=14)).strftime("%d/%m/%Y")

        client = OpenAI(api_key=api_key, base_url=base_url)

        system_prompt = f"""
אתה יועץ לימודים ומנהל אקדמי מומחה של בית הספר ללימודי המשך של הטכניון.
התאריך של היום: {today_str}.
תאריך הסף להרשמה (14 יום אחורה מהיום): {cutoff_str}.

תפקידך להעריך את התאמת המועמד מול כל הקורסים המופיעים בקטלוג שלמטה ולהחזיר את כל הקורסים שהוערכו בתוך אובייקט JSON.

קטלוג הקורסים:
{courses_catalog}

הנחיות קריטיות לקביעת ציונים ופסילות:
1. חוק 14 הימים וקורסים ללא צפי (חוק פסילה קשיח):
   - כל קורס שמסומן ב'סטטוס פתיחה: פג תוקף' או שתאריך פתיחתו עבר את {cutoff_str} — חייב לקבל ציון התאמה 0!
   - כל קורס שמסומן ב'סטטוס פתיחה: ללא צפי פתיחה' — חייב לקבל ציון התאמה 0!
   - עבור קורסים אלה יש לכתוב נימוק ברור בעברית שמתחיל במילים: "נפסל: מועד הפתיחה חלף וההרשמה הסתיימה" או "נפסל: אין כרגע מועד פתיחה צפוי לקורס".
2. חוק תנאי קבלה:
   - אם מועמד לא עומד בדרישות מהותיות של קורס (למשל נדרש תואר הנדסי ראשון ולמועמד אין) — הקורס יקבל ציון 0 עם נימוק מתאים.
3. קורסים מומלצים:
   - קורסים שמועד פתיחתם עתידי/פעיל והמועמד מתאים להם יקבלו ציון התאמה גבוה (בין 50 ל-100).
4. חובה להחזיר JSON תקין ללא Markdown:
{{
  "results": [
    {{
      "שם הקורס": "שם הקורס מתוך הקטלוג",
      "ציון התאמה": 0,
      "תאריך התחלה": "תאריך או לא צוין",
      "קמפוס": "חיפה / שרונה",
      "נימוק": "נימוק בעברית",
      "לינק לעמוד התכנית": "https://..."
    }}
  ]
}}
"""

        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": "נתוני השאלון של המועמד:\n" + json.dumps(req_body, ensure_ascii=False)}
        ]

        try:
            response = client.chat.completions.create(
                model=model_name,
                messages=messages,
                temperature=0.2,
                response_format={"type": "json_object"}
            )
        except Exception as model_err:
            logging.warning(f"Default model {model_name} failed. Attempting fallback: {model_err}")
            available_models = [
                m.id for m in client.models.list().data
                if not any(x in m.id.lower() for x in ["whisper", "guard", "tts", "vision", "embed"])
            ]
            if not available_models:
                raise model_err
            fallback_model = available_models[0]
            logging.info(f"Falling back to model: {fallback_model}")
            response = client.chat.completions.create(
                model=fallback_model,
                messages=messages,
                temperature=0.2,
                response_format={"type": "json_object"}
            )

        ai_content = response.choices[0].message.content

        return func.HttpResponse(
            ai_content,
            status_code=200,
            headers=CORS_HEADERS,
            mimetype="application/json"
        )

    except Exception as e:
        cause = repr(e.__cause__) if e.__cause__ else "None"
        error_details = f"{type(e).__name__}: {str(e)} | Underlying cause: {cause}"
        logging.error(f"Error communicating with AI model: {error_details}")
        return func.HttpResponse(
            json.dumps({"error": error_details}),
            status_code=500,
            headers=CORS_HEADERS,
            mimetype="application/json"
        )
