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

HEADER_KEYWORDS = [
    "שם הקורס", "שם התוכנית", "שם התכנית", "שם תוכנית",
    "פתיחה", "התחלה", "מועד", "תנאי קבלה", "אודות", "תיאור",
    "קמפוס", "מק\"ט", "מקט", "דפי מידע", "לינק", "רכזת", "עלות"
]

NON_TOPIC_KEYS = {
    "firstName", "lastName", "first_name", "last_name", "name",
    "email", "phone", "tel", "mobile", "id", "timestamp", "submittedAt",
    "education", "experience", "english_level", "study_mode",
    "start_time", "preferred_campus", "comments"
}

SYNONYM_MAP = {
    "בניית": {"full", "stack", "פיתוח", "אתרי", "אתרים", "web", "tcwd", "אינטרנט", "פול", "סטאק"},
    "אתרים": {"full", "stack", "פיתוח", "אתרי", "בניית", "web", "tcwd", "אינטרנט", "פול", "סטאק"},
    "אתרי": {"full", "stack", "פיתוח", "אתרים", "בניית", "web", "tcwd", "אינטרנט"},
    "full": {"stack", "פיתוח", "אתרים", "אתרי", "web", "תוכנה", "tcwd", "אינטרנט", "פול", "סטאק"},
    "stack": {"full", "פיתוח", "אתרים", "אתרי", "web", "תוכנה", "tcwd", "אינטרנט", "פול", "סטאק"},
    "פיתוח": {"full", "stack", "web", "תוכנה", "מפתחי", "תכנות", "אתרי", "אינטרנט", "tcwd"},
    "תכנות": {"full", "stack", "פיתוח", "תוכנה", "python", "web", "מידע"},
    "ai": {"בינה", "מלאכותית", "generative", "llm", "machine", "learning"},
    "בינה": {"ai", "llm", "generative", "דאטה", "data", "עסקית", "bi", "מלאכותית"},
    "מלאכותית": {"ai", "llm", "generative", "בינה"},
    "data": {"דאטה", "נתונים", "bi", "analyst", "אנליסט", "בינה", "big"},
    "דאטה": {"data", "נתונים", "bi", "analyst", "אנליסט", "big"},
    "אנליסט": {"analyst", "דאטה", "data", "bi", "נתונים", "ניתוח"},
    "bi": {"בינה", "עסקית", "דאטה", "data", "analyst", "נתונים"},
    "סייבר": {"cyber", "אבטחת", "מידע", "רשתות", "מיישמי", "soc", "הגנה"},
    "cyber": {"סייבר", "אבטחת", "מידע", "מיישמי"},
    "אבטחת": {"סייבר", "cyber", "מידע", "מיישמי"},
    "qa": {"בדיקות", "בודקי", "תוכנה", "איכות", "tcqa", "אוטומציה"},
    "בדיקות": {"qa", "בודקי", "תוכנה", "איכות", "tcqa"},
    "בודקי": {"qa", "בדיקות", "תוכנה", "tcqa"},
    "מוצר": {"product", "management", "מאפיינים", "ניתוח", "מערכות"},
    "product": {"מוצר", "management", "ניהול"},
    "ניתוח": {"מערכות", "מידע", "אנליסט", "דאטה", "data"},
    "מערכות": {"ניתוח", "מידע", "הנדסת", "מנמ", "cio"},
    "ניהול": {"מנהלים", "פרויקטים", "מערכות", "mba", "מנמ", "דירקטורים", "pmp", "בכיר"},
    "פרויקטים": {"ניהול", "pmp", "מערכות", "agile", "scrum"},
    "תואר": {"שני", "me", "mba", "הנדסה", "אקדמי", "מוסמך", "תארים"},
    "הנדסה": {"מהנדסים", "me", "תואר", "שני", "מערכות", "אזרחית", "עירונית", "ביו"},
    "עיצוב": {"גרפי", "תעשייתי", "פנים", "אדריכלות", "ui", "ux", "דיגיטלי"},
    "שיווק": {"דיגיטלי", "digital", "מכירות", "עסקי", "מסחר"},
    "נדלן": {"מקרקעין", "שמאות", "בנייה", "mre"},
    "מקרקעין": {"נדלן", "שמאות", "mre"},
}


def clean_env(val: str) -> str:
    if not val:
        return ""
    return re.sub(r'[^\x21-\x7E]', '', val)


def find_header_row_index(rows: list) -> int:
    best_idx = 0
    best_hits = 0
    for idx, r in enumerate(rows[:10]):
        if not r:
            continue
        cells_text = [str(c).strip() for c in r if c is not None and str(c).strip()]
        hits = sum(
            1 for cell in cells_text
            if any(kw in cell for kw in HEADER_KEYWORDS)
        )
        if hits > best_hits:
            best_hits = hits
            best_idx = idx
    return best_idx if best_hits >= 2 else 0


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
        if any(phrase in val for phrase in ["ללא צפי", "לא צפי", "לא ידוע", "טרם נקבע", "אין צפי", "מבוטל"]):
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


def expand_keywords(words: set) -> set:
    expanded = set(words)
    for w in words:
        if w in SYNONYM_MAP:
            expanded.update(SYNONYM_MAP[w])
    return expanded


def extract_topic_and_general_keywords(req_body: dict) -> tuple[set, set, str]:
    form_data = req_body.get("body", req_body) if isinstance(req_body, dict) else {}
    topic_tokens = []
    general_tokens = []

    for k, v in form_data.items():
        if not v or str(v).strip().lower() in ("null", "none", ""):
            continue
        val_str = str(v).lower()
        if k not in NON_TOPIC_KEYS:
            topic_tokens.append(val_str)
        elif k in ("current_field", "goals"):
            general_tokens.append(val_str)

    topic_text = " ".join(topic_tokens)
    topic_words = set(re.findall(r'[a-zA-Z\u0590-\u05FF]{2,}', topic_text))
    general_words = set(re.findall(r'[a-zA-Z\u0590-\u05FF]{2,}', " ".join(general_tokens)))

    return expand_keywords(topic_words), expand_keywords(general_words), topic_text


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
    topic_kw, general_kw, topic_text = extract_topic_and_general_keywords(req_body)

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

            header_idx = find_header_row_index(rows)
            headers = [str(cell).strip() if cell is not None else "" for cell in rows[header_idx]]
            clean_headers = [h for h in headers if h]

            date_columns = [
                h for h in clean_headers
                if any(k in h for k in ["פתיחה", "התחלה", "מועד", "תאריך", "ת."])
            ]

            for row_vals in rows[header_idx + 1:]:
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
                if not course_name or course_name in ("שם הקורס", "שם התוכנית", "שם התכנית", "שם תוכנית"):
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
                    elif any(k in col for k in ["קמפוס", "מיקום", "שלוחה"]):
                        campus_val = v
                    elif any(k in col for k in ["פתיחה", "התחלה", "מועד"]):
                        date_val = v

                course_title_lower = course_name.lower()
                sheet_lower = sheet_name.lower()

                topic_match_score = (
                    sum(30 for kw in topic_kw if kw in course_title_lower)
                    + sum(12 for kw in topic_kw if kw in sheet_lower)
                )

                if status in ("EXPIRED", "NO_SCHEDULE"):
                    # ניקוי הערות פנימיות של רכזות כדי שלא ידלפו לתצוגת המועמד
                    if status == "NO_SCHEDULE":
                        clean_date_display = "אין צפי פתיחה"
                    else:
                        extracted = extract_dates_from_value(date_val)
                        clean_date_display = extracted[0].strftime("%d/%m/%Y") if extracted else "הרשמה נסגרה"

                    reason = (
                        "נפסל: מועד הפתיחה חלף (עברו מעל 14 יום) וההרשמה הסתיימה"
                        if status == "EXPIRED"
                        else "נפסל: אין כרגע מועד פתיחה צפוי לקורס"
                    )
                    expired_samples.append((
                        topic_match_score,
                        {
                            "שם הקורס": course_name,
                            "ציון התאמה": 0,
                            "תאריך התחלה": clean_date_display,
                            "קמפוס": campus_val,
                            "נימוק": reason,
                            "לינק לעמוד התכנית": link_val
                        }
                    ))
                    continue

                desc_lower = " ".join(str_row.values()).lower()

                total_score = (
                    topic_match_score
                    + sum(3 for kw in topic_kw if kw in desc_lower)
                    + sum(4 for kw in general_kw if kw in course_title_lower)
                    + sum(1 for kw in general_kw if kw in desc_lower)
                )

                compact_lines = [
                    f"קורס: {course_name} | קטגוריה: {sheet_name} | פתיחה: {date_val} | קמפוס: {campus_val} | לינק: {link_val}"
                ]
                for col in clean_headers:
                    if any(k in col for k in ["תנאי קבלה", "קהל יעד", "דרישות"]):
                        val = str_row.get(col, "")
                        if val:
                            compact_lines.append(f"תנאי קבלה: {val[:180]}")
                    elif any(k in col for k in ["תיאור", "אודות", "מתכונת"]):
                        val = str_row.get(col, "")
                        if val:
                            compact_lines.append(f"{col}: {val[:130]}")

                active_courses.append((total_score, "\n".join(compact_lines)))

    except Exception as e:
        logging.error(f"Failed to load courses workbook: {e}")

    active_courses.sort(key=lambda x: x[0], reverse=True)
    top_blocks = [item[1] for item in active_courses[:24]]

    expired_samples.sort(key=lambda x: x[0], reverse=True)
    top_expired = [item[1] for item in expired_samples[:3]]

    catalog_text = "\n\n".join(top_blocks)
    return catalog_text, top_expired


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

        system_prompt = f"""אתה יועץ לימודים אקדמי של בית הספר ללימודי המשך של הטכניון. התאריך היום: {today_str}.
התאם את פרטי המועמד לרשימת הקורסים הפעילים הבאה:

{courses_catalog}

הנחיות קריטיות (חובה להקפיד ב-100%):
1. החזר אך ורק אובייקט JSON תקין במבנה:
{{"results": [{{"שם הקורס": "...", "ציון התאמה": 92, "תאריך התחלה": "...", "קמפוס": "...", "נימוק": "...", "לינק לעמוד התכנית": "https://..."}}]}}
2. בדיקת רקע תעסוקתי ותנאי קבלה (קריטי!):
   - שים לב לשדה תחום העיסוק הנוכחי (current_field) ולשנות הניסיון (experience). אל תניח ששנות הניסיון של המועמד הן בתחום הטכנולוגי אם תחום העיסוק שלו הוא תחום אחר (כגון "חינוך והוראה", "בריאות", "פיננסים" וכו')!
   - אם קורס מיועד למנהלים בכירים או דורש ניסיון קודם בהייטק/אבטחת מידע (כגון קורסי CISO, מנמ"רים, Executives או קורסים מתקדמים) ולמועמד אין רקע תעסוקתי רלוונטי — **חובה לפסול את הקורס עם `"ציון התאמה": 0`** ונימוק שמתחיל ב`"נפסל: הקורס דורש ניסיון קודם בתחום..."`.
   - למועמדים ללא רקע טכנולוגי קודם, המלץ רק על מסלולי כניסה והסבה (כגון "מיישמי סייבר ואבטחת מידע") עם ציון 80–95.
3. התעלם משמות המפתחות באנגלית ב-JSON והתייחס רק לערכים בעברית שבחרו המועמדים (אל תכתוב "תחום ה-AI" אם המועמד בחר סייבר או דאטה).
4. השתמש אך ורק בשמות הקורסים, התאריכים והלינקים המדויקים מהרשימה למעלה. כתוב נימוק מקצועי בעברית בגוף שלישי."""

        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": "נתוני השאלון של המועמד:\n" + json.dumps(req_body, ensure_ascii=False)}
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
                    temperature=0.2
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
        raw_results = parsed.get("results", [])

        validated_results = []
        existing_names = set()

        for item in raw_results:
            if not isinstance(item, dict):
                continue
            c_name = str(item.get("שם הקורס", "")).strip()
            if not c_name or c_name in existing_names:
                continue

            reason_str = str(item.get("נימוק", "")).strip()
            try:
                score_val = int(float(item.get("ציון התאמה", 0)))
            except Exception:
                score_val = 0

            if reason_str.startswith("נפסל") or "אינו מתאים" in reason_str:
                score_val = 0

            link_str = str(item.get("לינק לעמוד התכנית", "לא זמין")).strip()
            if link_str and link_str != "לא זמין" and not link_str.startswith("http"):
                link_str = "https://" + link_str

            validated_results.append({
                "שם הקורס": c_name,
                "ציון התאמה": score_val,
                "תאריך התחלה": str(item.get("תאריך התחלה", "לא צוין")).strip(),
                "קמפוס": str(item.get("קמפוס", "לא צוין")).strip(),
                "נימוק": reason_str,
                "לינק לעמוד התכנית": link_str
            })
            existing_names.add(c_name)

        for exp in expired_samples:
            if exp["שם הקורס"] not in existing_names:
                validated_results.append(exp)
                existing_names.add(exp["שם הקורס"])

        validated_results.sort(key=lambda x: x["ציון התאמה"], reverse=True)

        return func.HttpResponse(
            json.dumps({"results": validated_results}, ensure_ascii=False),
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
