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

# ביטוי רגולרי שתופס תאריכים בכל פורמט: DD/MM/YYYY, DD.MM.YYYY, DD/MM/YY, DD-MM-YYYY
DATE_PATTERN = re.compile(r'\b(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})\b')


def clean_env(val: str) -> str:
    if not val:
        return ""
    return re.sub(r'[^\x21-\x7E]', '', val)


def extract_dates_from_text(text: str) -> list:
    """מחלץ את כל התאריכים מתוך טקסט חופשי וממיר אותם לאובייקטי datetime"""
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


def is_course_expired(row: dict, date_columns: list, cutoff_date: datetime) -> bool:
    """
    בודק האם הקורס עבר את חוק ה-14 יום.
    מחזיר True אם נמצאו תאריכי פתיחה בשורה וכולם ישנים יותר מ-cutoff_date.
    """
    found_dates = []

    # 1. בדיקה בעמודות התאריך המזוהות
    for col in date_columns:
        cell_val = str(row.get(col, "")).strip()
        if cell_val:
            found_dates.extend(extract_dates_from_text(cell_val))

    # 2. אם לא נמצאו תאריכים בעמודת התאריך, סריקת גיבוי בכל תאי השורה
    if not found_dates:
        for col, val in row.items():
            if any(k in col for k in ["פתיחה", "התחלה", "מועד", "תאריך"]):
                found_dates.extend(extract_dates_from_text(str(val)))

    # אם לקורס יש תאריכים מוגדרים, וכל התאריכים שלו קטנים מתאריך הסף (לפני יותר מ-14 יום) -> הקורס פג תוקף!
    if found_dates:
        if all(d < cutoff_date for d in found_dates):
            return True

    return False


def load_courses_catalog() -> str:
    sheet_id = "1emkWYVjlrhC0m67gz7b3MUXzsG3RtlX-kbg8eNvzBtc"
    gid = "1525331068"

    csv_url = (
        f"https://docs.google.com/spreadsheets/d/"
        f"{sheet_id}/export?format=csv&gid={gid}"
    )

    response = requests.get(csv_url, timeout=30)
    response.raise_for_status()
    # הבטחת קידוד UTF-8 תקין לעברית וניקוי BOM
    response.encoding = "utf-8-sig"

    reader = csv.DictReader(StringIO(response.text))
    if not reader.fieldnames:
        return ""

    # ניקוי רווחים משמות העמודות
    clean_fieldnames = [col.strip() for col in reader.fieldnames if col]

    # זיהוי אוטומטי של כל עמודה שמייצגת תאריך פתיחה/התחלה
    date_columns = [
        col for col in clean_fieldnames
        if any(keyword in col for keyword in ["פתיחה", "התחלה", "מועד", "תאריך", "ת."])
    ]

    today = datetime.now()
    cutoff_date = today - timedelta(days=14)

    filtered_rows = []
    skipped_count = 0

    for raw_row in reader:
        # יצירת מילון נקי עם שמות עמודות ללא רווחים מיותרים
        row = {k.strip(): (v.strip() if isinstance(v, str) else "") for k, v in raw_row.items() if k}

        # דילוג על שורות ריקות לחלוטין
        if not any(row.values()):
            continue

        if is_course_expired(row, date_columns, cutoff_date):
            skipped_count += 1
            logging.info(f"Filtered out expired course (>14 days): {row.get('שם הקורס', '')}")
            continue

        filtered_rows.append(row)

    logging.info(f"Courses kept: {len(filtered_rows)} | Courses filtered out (>14 days): {skipped_count}")

    catalog_blocks = []
    for row in filtered_rows:
        row_lines = []
        for col in clean_fieldnames:
            val = row.get(col, "")
            if val:
                # וידוא שלינקים תמיד יתחילו ב-https:// כדי למנוע שגיאות 404 באתר
                if ("לינק" in col or "קישור" in col or "אתר" in col) and val != "לא זמין":
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
התאריך של היום הוא: {today_str}.
תאריך הסף להרשמה (14 יום אחורה מהיום) הוא: {cutoff_str}.

תפקידך לנתח את פרטי השאלון שמילא המועמד, להצליב אותם מול קטלוג הקורסים הפעילים שלהלן, ולספק התאמה מדויקת ומנומקת.

קטלוג הקורסים הפעילים:
{courses_catalog}

הנחיות קריטיות (חובה להקפיד ב-100%):
1. החזר אך ורק אובייקט JSON תקין, ללא Markdown וללא שום טקסט מחוץ ל-JSON.
2. חוק 14 הימים (קריטי!): חל איסור מוחלט להציג או להמליץ על קורס שתאריך ההתחלה/פתיחה שלו מוקדם יותר מ-{cutoff_str} (כלומר עברו יותר מ-14 יום מאז שנפתח). אם קורס כזה מופיע, התעלם ממנו לחלוטין ואל תכלול אותו בפלט.
3. החזר לפחות 3 קורסים מומלצים מבין הקורסים הפעילים.
4. ניתן להחזיר 1-2 קורסים שנפסלו עם ציון 0 (בגלל אי-עמידה בתנאי קבלה או חוסר התאמה לתחום).
5. "ציון התאמה" חייב להיות מספר שלם (Integer) בין 0 ל-100.
6. השתמש אך ורק בקורסים ובלינקים המופיעים בקטלוג שלמעלה. ודא שכל לינק מתחיל ב-https://.
7. כל הנימוקים בעברית מקצועית, רהוטה ובגוף שלישי.

מבנה הפלט הנדרש:
{{
  "results": [
    {{
      "שם הקורס": "שם הקורס",
      "ציון התאמה": 95,
      "תאריך התחלה": "תאריך",
      "קמפוס": "קמפוס",
      "נימוק": "הסבר מפורט",
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
