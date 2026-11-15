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

DATE_DMY_PATTERN = re.compile(r'\b(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})\b')
DATE_YMD_PATTERN = re.compile(r'\b(\d{4})[/.\-](\d{1,2})[/.\-](\d{1,2})\b')

# שדות הבחירה הספציפיים בטופס שמייצגים את תחום ההתעניינות המרכזי של המועמד
INTEREST_FIELDS = [
    "field",
    "data_programming",
    "data_interest",
    "ai_interest",
    "cyber_technical",
    "cyber_interest",
    "pm_experience",
    "digital_interest",
    "business_interest",
]

SYNONYM_MAP = {
    "בניית": {"full", "stack", "פיתוח", "אתרי", "אתרים", "web", "tcwd", "אינטרנט"},
    "אתרים": {"full", "stack", "פיתוח", "אתרי", "בניית", "web", "tcwd", "אינטרנט"},
    "אתרי": {"full", "stack", "פיתוח", "אתרים", "בניית", "web", "tcwd", "אינטרנט"},
    "full": {"stack", "פיתוח", "אתרים", "אתרי", "web", "תוכנה", "tcwd", "אינטרנט"},
    "stack": {"full", "פיתוח", "אתרים", "אתרי", "web", "תוכנה", "tcwd", "אינטרנט"},
    "פיתוח": {"full", "stack", "web", "תוכנה", "מפתחי", "תכנות", "אתרי", "אינטרנט"},
    "תכנות": {"full", "stack", "פיתוח", "תוכנה", "python", "web"},
    "ai": {"בינה", "מלאכותית", "generative", "llm", "machine", "learning"},
    "בינה": {"ai", "llm", "generative", "דאטה", "data", "עסקית", "bi"},
    "data": {"דאטה", "נתונים", "bi", "analyst", "אנליסט", "בינה"},
    "דאטה": {"data", "נתונים", "bi", "analyst", "אנליסט"},
    "אנליסט": {"analyst", "דאטה", "data", "bi", "נתונים"},
    "bi": {"בינה", "עסקית", "דאטה", "data", "analyst"},
    "סייבר": {"cyber", "אבטחת", "מידע", "רשתות", "מיישמי"},
    "cyber": {"סייבר", "אבטחת", "מידע", "מיישמי"},
    "qa": {"בדיקות", "בודקי", "תוכנה", "איכות", "tcqa"},
    "בדיקות": {"qa", "בודקי", "תוכנה", "איכות", "tcqa"},
    "מוצר": {"product", "management", "מאפיינים", "ניתוח", "מערכות"},
    "ניהול": {"מנהלים", "פרויקטים", "מערכות", "mba", "מנמ", "דירקטורים", "pmp"},
    "פרויקטים": {"ניהול", "pmp", "מערכות", "agile"},
    "תואר": {"שני", "me", "mba", "הנדסה", "אקדמי", "מוסמך"},
    "הנדסה": {"מהנדסים", "me", "תואר", "שני", "מערכות", "אזרחית"},
    "שיווק": {"דיגיטלי", "digital", "מכירות", "עסקי"},
}


def clean_env(val: str) -> str:
    if not val:
        return ""
    return re.sub(r'[^\x21-\x7E]', '', val)


def extract_dates_from_value(val) -> list:
    if not val:
        return []
    if isinstance(val, datetime):
        return [val]

    text = str(val)
    parsed_dates = []

    for match in DATE_YMD_PATTERN.finditer(text):
        year, month, day = int(match.group(1)), int(match.group(2)), int(match.group(3))
        try:
            parsed_dates.append(datetime(year, month, day))
        except ValueError:
            continue

    if parsed_dates:
        return parsed_dates

    for match in DATE_DMY_PATTERN.finditer(text):
        day, month, year = int(match.group(1)), int(match.group(2)), int(match.group(3))
        if year < 100:
            year += 2000
        try:
            parsed_dates.append(datetime(year, month, day))
        except ValueError:
            continue

    return parsed_dates


def format_cell_display(val) -> str:
    if val is None:
        return ""
    if isinstance(val, datetime):
        return val.strftime("%d/%m/%Y")
    text = str(val).strip()
    text = re.sub(r'(\d{4}-\d{2}-\d{2})\s+00:00:00', r'\1', text)
    return text


def check_course_date_status(raw_row: dict, date_columns: list, cutoff_date: datetime) -> str:
    for col in date_columns:
        val = str(raw_row.get(col, "")).strip()
        if any(phrase in val for phrase in ["ללא צפי", "לא צפי", "לא ידוע", "טרם נקבע"]):
            return "NO_SCHEDULE"

    found_dates = []
    for col in date_columns:
        cell_val = raw_row.get(col)
        if cell_val:
            found_dates.extend(extract_dates_from_value(cell_val))

    if not found_dates:
        for col, val in raw_row.items():
            if any(k in col for k in ["פתיחה", "התחלה", "מועד", "תאריך"]):
                found_dates.extend(extract_dates_from_value(val))

    if found_dates:
        if all(d < cutoff_date for d in found_dates):
            return "EXPIRED"

    all_text = " ".join(str(v) for v in raw_row.values() if v is not None)
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


def expand_words(words: set) -> set:
    expanded = set(words)
    for w in words:
        if w in SYNONYM_MAP:
            expanded.update(SYNONYM_MAP[w])
    return expanded


def extract_user_keywords(req_body: dict) -> tuple[set, set]:
    """
    מחזיר שני סטים של מילות מפתח:
    1. primary_keywords: מתוך שדות הבחירה של תחום העניין וההתמחות (בעלי משקל גבוה)
    2. general_keywords: מתוך שאר שדות הטופס (רקע, יעדים, השכלה)
    """
    form_data = req_body.get("body", req_body) if isinstance(req_body, dict) else {}
    primary_tokens = []
    general_tokens = []

    for k, v in form_data.items():
        if not v or str(v).strip().lower() in ("null", "none", ""):
            continue
        val_str = str(v).lower()
        if k in INTEREST_FIELDS:
            primary_tokens.append(val_str)
        else:
            general_tokens.append(val_str)

    primary_words = set(re.findall(r'[a-zA-Z\u0590-\u05FF]{2,}', " ".join(primary_tokens)))
    general_words = set(re.findall(r'[a-zA-Z\u0590-\u05FF]{2,}', " ".join(general_tokens)))

    return expand_words(primary_words), expand_words(general_words)


def get_course_name_from_row(row: dict, clean_headers: list) -> str:
    priority_keys = ["שם הקורס", "שם התוכנית", "שם התכנית", "שם תוכנית", "שם קורס", "הכשרה", "תוכנית"]
    for pk in priority_keys:
        if row.get(pk):
            return row[pk]
    for h in clean_headers:
        if ("שם" in h or "קורס" in h or "תוכנית" in h or "תכנית" in h) and row.get(h):
            val = row[h]
            if not val.startswith("http") and not val.isdigit():
                return val
    for h in clean_headers:
        val = row.get(h, "")
        if val and len(val) > 3 and not val.startswith("http") and not val.replace("-", "").isdigit():
            return val
    return ""


def load_filtered_courses_catalog(req_body: dict) -> tuple[str, list]:
    sheet_id = "1emkWYVjlrhC0m67gz7b3MUXzsG3RtlX-kbg8eNvzBtc"
    excel_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"

    today = datetime.now()
    cutoff_date = today - timedelta(days=14)
    primary_kw, general_kw = extract_user_keywords(req_body)

    active_courses = []
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

                raw_row = {}
                str_row = {}
                for idx, col_name in enumerate(headers):
                    if col_name and idx < len(row_vals):
                        raw_val = row_vals[idx]
                        raw_row[col_name] = raw_val
                        str_row[col_name] = format_cell_display(raw_val)

                course_name = get_course_name_from_row(str_row, clean_headers)
                if not course_name:
                    continue

                status = check_course_date_status(raw_row, date_columns, cutoff_date)

                link_val = "לא זמין"
                campus_val = "לא צוין"
                date_val = "לא צוין"
                for col in clean_headers:
                    v = str_row.get(col, "")
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

                title_lower = f"{sheet_name} {course_name}".lower()
                desc_lower = " ".join(str_row.values()).lower()

                # ניקוד אוניברסלי: התאמה ישירה לבחירת תחום ההתעניינות של המועמד מקבלת עדיפות עליונה בכל גיליון
                score = (
                    sum(15 for kw in primary_kw if kw in title_lower)
                    + sum(4 for kw in primary_kw if kw in desc_lower)
                    + sum(3 for kw in general_kw if kw in title_lower)
                    + sum(1 for kw in general_kw if kw in desc_lower)
                )

                compact_lines = [
                    f"תחום/גיליון: {sheet_name}",
                    f"שם הקורס: {course_name}",
                    f"תאריך פתיחה: {date_val}",
                    f"קמפוס: {campus_val}",
                    f"לינק: {link_val}",
                ]
                for col in clean_headers:
                    if any(k in col for k in ["תיאור", "אודות", "תנאי קבלה", "מתכונת"]):
                        val = str_row.get(col, "")
                        if val:
                            compact_lines.append(f"{col}: {val[:200]}")

                active_courses.append((score, "\n".join(compact_lines)))

    except Exception as e:
        logging.error(f"Failed to load courses workbook: {e}")

    active_courses.sort(key=lambda x: x[0], reverse=True)
    # לקיחת 12 המועמדים המובילים (נשאר בבטחה מתחת למגבלת ה-7,000 טוקנים של Groq)
    top_blocks = [item[1] for item in active_courses[:12]]

    catalog_text = "\n\n" + ("\n" + "-" * 30 + "\n\n").join(top_blocks)
    return catalog_text, expired_samples


def extract_json_object(raw_text: str) -> dict:
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
        courses_catalog, expired_samples = load_filtered_courses_catalog(req_body)
        today_str = datetime.now().strftime("%d/%m/%Y")

        client = OpenAI(api_key=api_key, base_url=base_url)

        system_prompt = f"""אתה יועץ לימודים מומחה של בית הספר ללימודי המשך של הטכניון. התאריך היום: {today_str}.
התאם את פרטי המועמד לרשימת הקורסים הפעילים הבאה:
{courses_catalog}

חוקים:
1. החזר אך ורק JSON תקין במבנה:
{{"results": [{{"שם הקורס": "...", "ציון התאמה": 90, "תאריך התחלה": "...", "קמפוס": "...", "נימוק": "...", "לינק לעמוד התכנית": "https://..."}}]}}
2. בחר 3-5 קורסים מתאימים (ציון 65-100) בהתאם לתחום העניין ולרקע של המועמד. אם מועמד לא עומד בתנאי קבלה מהותי של קורס שבחרת לבחון, תן ציון 0 וכתוב "נפסל: ...".
3. השתמש רק בקורסים ובלינקים מהרשימה למעלה (תמיד עם https://). כתוב נימוק תמציתי ומקצועי בעברית בגוף שלישי."""

        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": "פרטי המועמד:\n" + json.dumps(req_body, ensure_ascii=False)}
        ]

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
