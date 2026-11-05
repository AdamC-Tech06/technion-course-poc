import azure.functions as func
import json
import logging
import os
import re
import requests
from io import BytesIO
from datetime import datetime, timedelta
import openpyxl
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
    # הורדת כל הגיליון במלואו כקובץ אקסל בקריאה אחת
    excel_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"

    today = datetime.now()
    cutoff_date = today - timedelta(days=14)
    all_blocks = []

    try:
        response = requests.get(excel_url, timeout=25)
        response.raise_for_status()

        wb = openpyxl.load_workbook(filename=BytesIO(response.content), data_only=True)

        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            rows = list(sheet.iter_rows(values_only=True))
            if not rows or len(rows) < 2:
                continue

            headers = [str(cell).strip() if cell is not None else "" for cell in rows[0]]
            clean_headers = [h for h in headers if h]

            date_columns = [
                h for h in clean_headers
                if any(k in h for k in ["פתיחה", "התחלה", "מועד", "תאריך", "ת."])
            ]

            for row_vals in rows[1:]:
                if not any(row_vals):
                    continue

                row = {}
                for idx, col_name in enumerate(headers):
                    if col_name and idx < len(row_vals):
                        val = row_vals[idx]
                        row[col_name] = str(val).strip() if val is not None else ""

                course_name = row.get("שם הקורס") or row.get("שם תוכנית") or row.get("שם התכנית") or ""
                if not course_name:
                    continue

                status = check_course_date_status(row, date_columns, cutoff_date)

                row_lines = [f"תחום/גיליון: {sheet_name}"]
                if status == "EXPIRED":
                    row_lines.append("סטטוס פתיחה: פג תוקף (עברו מעל 14 יום ממועד הפתיחה - חובה לתת ציון 0 ולפסול)")
                elif status == "NO_SCHEDULE":
                    row_lines.append("סטטוס פתיחה: ללא צפי פתיחה (חובה לתת ציון 0 ולפסול)")
                else:
                    row_lines.append("סטטוס פתיחה: פעיל ורלוונטי להרשמה")

                for col in clean_headers:
                    val = row.get(col, "")
                    if val:
                        if ("לינק" in col or "קישור" in col or "דפי מידע" in col or "אתר" in col) and val != "לא זמין":
                            if not val.startswith("http://") and not val.startswith("https://"):
                                val = "https://" + val
                        row_lines.append(f"{col}: {val}")

                if row_lines:
                    all_blocks.append("\n".join(row_lines))

    except Exception as e:
        logging.error(f"Failed to load courses workbook: {e}")

    return "\n\n" + ("\n" + "=" * 80 + "\n\n").join(all_blocks)


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

תפקידך לנתח את פרטי השאלון שמילא המועמד, להצליב אותם מול קטלוג הקורסים הפעילים שלהלן, ולספק התאמה מדויקת ומנומקת.

קטלוג הקורסים:
{courses_catalog}

הנחיות קריטיות לקביעת ציונים ופסילות:
1. חוק 14 הימים וקורסים ללא צפי:
   - כל קורס שמסומן ב'סטטוס פתיחה: פג תוקף' או שתאריך פתיחתו עבר את {cutoff_str} חייב לקבל ציון 0 ולהיפסל.
   - כל קורס שמסומן ב'סטטוס פתיחה: ללא צפי פתיחה' חייב לקבל ציון 0 ולהיפסל.
2. קורסים מומלצים:
   - המלץ על 3-4 קורסים שהם ב'סטטוס פתיחה: פעיל ורלוונטי להרשמה' ותואמים ביותר לתחום ולרקע של המועמד (ציונים 60-100).
   - החזר עד 3 קורסים שנפסלו (ציון 0).
3. "ציון התאמה" חייב להיות מספר שלם בין 0 ל-100.
4. השתמש בלינקים מהקטלוג (תמיד עם קידומת https://).
5. החזר JSON תקין בלבד:
{{
  "results": [
    {{
      "שם הקורס": "שם הקורס",
      "ציון התאמה": 95,
      "תאריך התחלה": "תאריך",
      "קמפוס": "קמפוס",
      "נימוק": "נימוק בעברית מקצועית בגוף שלישי",
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
