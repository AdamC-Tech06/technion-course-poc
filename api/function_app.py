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


def load_courses_catalog() -> tuple[str, list]:
    """
    מחזיר:
    1. טקסט מרוכז של הקורסים הפעילים בלבד (לחסכון בטוקנים ומניעת חריגת מגבלה ב-Groq)
    2. רשימה של עד 3 קורסים שנפסלו אוטומטית בגלל תאריך/ללא צפי להצגה בטבלת הנפסלים
    """
    sheet_id = "1emkWYVjlrhC0m67gz7b3MUXzsG3RtlX-kbg8eNvzBtc"
    excel_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"

    today = datetime.now()
    cutoff_date = today - timedelta(days=14)
    active_blocks = []
    expired_samples = []

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

                # חילוץ קישור וקמפוס ותאריך גם עבור קורסים שנפסלו
                link_val = "לא זמין"
                campus_val = "לא צוין"
                date_val = "לא צוין"
                for col in clean_headers:
                    v = row.get(col, "")
                    if not v:
                        continue
                    if any(k in col for k in ["לינק", "קישור", "דפי מידע", "אתר"]) and v != "לא זמין":
                        link_val = v if (v.startswith("http://") or v.startswith("https://")) else "https://" + v
                    elif "קמפוס" in col or "מיקום" in col:
                        campus_val = v
                    elif any(k in col for k in ["פתיחה", "התחלה", "מועד"]):
                        date_val = v

                if status in ("EXPIRED", "NO_SCHEDULE"):
                    if len(expired_samples) < 3:
                        reason = (
                            "נפסל: מועד הפתיחה חלף (עברו מעל 14 יום) וההרשמה הסתיימה"
                            if status == "EXPIRED"
                            else "נפסל: אין כרגע מועד פתיחה צפוי לקורס"
                        )
                        expired_samples.append({
                            "שם הקורס": course_name,
                            "ציון התאמה": 0,
                            "תאריך התחלה": date_val,
                            "קמפוס": campus_val,
                            "נימוק": reason,
                            "לינק לעמוד התכנית": link_val
                        })
                    continue

                # עבור קורסים פעילים בלבד - בניית בלוק תמציתי וברור
                row_lines = [f"גיליון: {sheet_name}"]
                for col in clean_headers:
                    val = row.get(col, "")
                    if val:
                        if any(k in col for k in ["לינק", "קישור", "דפי מידע", "אתר"]) and val != "לא זמין":
                            if not val.startswith("http://") and not val.startswith("https://"):
                                val = "https://" + val
                        # קיצור שדות טקסט ארוכים מדי (כמו עלויות/תקנונים) כדי לא לחרוג ממכסת טוקנים
                        if len(val) > 350:
                            val = val[:350] + "..."
                        row_lines.append(f"{col}: {val}")

                if row_lines:
                    active_blocks.append("\n".join(row_lines))

    except Exception as e:
        logging.error(f"Failed to load courses workbook: {e}")

    catalog_text = "\n\n" + ("\n" + "-" * 40 + "\n\n").join(active_blocks)
    return catalog_text, expired_samples


def extract_json_object(raw_text: str) -> dict:
    """מחלץ אובייקט JSON תקין גם אם מודל Reasoning הוסיף טקסט או בלוק קוד מסביב"""
    cleaned = raw_text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r'^```(?:json)?\s*', '', cleaned)
        cleaned = re.sub(r'\s*```$', '', cleaned)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        match = re.search(r'\{[\s\S]*\}', cleaned)
        if match:
            return json.loads(match.group(0))
        raise


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
    base_url = clean_env(os.environ.get("OPENAI_BASE_URL", "[https://api.groq.com/openai/v1](https://api.groq.com/openai/v1)"))
    configured_model = clean_env(os.environ.get("OPENAI_MODEL", ""))
    primary_model = configured_model if configured_model else "openai/gpt-oss-120b"

    if not api_key:
        return func.HttpResponse(
            json.dumps({"error": "Missing OPENAI_API_KEY configuration on Azure"}),
            status_code=500,
            headers=CORS_HEADERS,
            mimetype="application/json"
        )

    try:
        courses_catalog, expired_samples = load_courses_catalog()
        today_str = datetime.now().strftime("%d/%m/%Y")

        client = OpenAI(api_key=api_key, base_url=base_url)

        system_prompt = f"""
אתה יועץ לימודים ומנהל אקדמי מומחה של בית הספר ללימודי המשך של הטכניון.
התאריך של היום: {today_str}.

תפקידך לנתח את פרטי השאלון שמילא המועמד, להצליב אותם מול קטלוג הקורסים הפעילים שלהלן, ולהחזיר התאמה מדויקת ומנומקת.

קטלוג הקורסים הפעילים (כל הקורסים כאן בתוקף):
{courses_catalog}

הנחיות קריטיות (חובה להקפיד ב-100%):
1. החזר אך ורק אובייקט JSON חוקי ותקני, ללא שום טקסט לפניו או אחריו.
2. בחר 3 עד 5 קורסים מתאימים ביותר מתוך הקטלוג שלמעלה עבור המועמד (עם ציוני התאמה 65–100).
3. אם מועמד אינו עומד בתנאי קבלה מהותי של קורס שבחרת לבחון — תן ציון 0 וכתוב נימוק שמתחיל ב"נפסל: ...".
4. "ציון התאמה" חייב להיות מספר שלם (Integer) בין 0 ל-100.
5. השתמש אך ורק בשמות הקורסים, תאריכי ההתחלה והלינקים שמופיעים בקטלוג שלמעלה (ודא קידומת https://).
6. כתוב נימוק מקצועי, רהוט וחם בעברית בגוף שלישי ("המועמד").

מבנה הפלט הנדרש:
{{
  "results": [
    {{
      "שם הקורס": "שם הקורס",
      "ציון התאמה": 95,
      "תאריך התחלה": "תאריך",
      "קמפוס": "קמפוס",
      "נימוק": "נימוק מקצועי בעברית",
      "לינק לעמוד התכנית": "https://..."
    }}
  ]
}}
"""

        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": "נתוני השאלון של המועמד:\n" + json.dumps(req_body, ensure_ascii=False)}
        ]

        # רק המודלים שבאמת זמינים בחשבון Groq שלך
        candidate_models = list(dict.fromkeys([
            primary_model,
            "openai/gpt-oss-120b",
            "openai/gpt-oss-20b",
            "qwen/qwen3.8-27b"
        ]))

        last_error = None
        response = None

        for m_name in candidate_models:
            try:
                logging.info(f"Trying Groq model: {m_name}")
                response = client.chat.completions.create(
                    model=m_name,
                    messages=messages,
                    temperature=0.3
                )
                if response and response.choices and response.choices[0].message.content:
                    break
            except Exception as err:
                logging.warning(f"Model {m_name} failed: {err}")
                last_error = err

        if not response or not response.choices:
            raise last_error if last_error else Exception("Failed to get response from Groq models")

        raw_content = response.choices[0].message.content
        parsed = extract_json_object(raw_content)
        results_list = parsed.get("results", [])

        # צירוף אוטומטי של עד 3 קורסים שנפסלו בגלל תאריך/ללא צפי (ללא בזבוז טוקנים במודל)
        existing_names = {r.get("שם הקורס") for r in results_list}
        for exp in expired_samples:
            if exp["שם הקורס"] not in existing_names:
                results_list.append(exp)

        final_payload = {"results": results_list}

        return func.HttpResponse(
            json.dumps(final_payload, ensure_ascii=False),
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
