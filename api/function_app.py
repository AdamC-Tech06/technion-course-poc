import azure.functions as func
import csv
import json
import logging
import os
import re
import requests
from io import StringIO
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor
from openai import OpenAI

app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)

CORS_HEADERS = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
}

DATE_PATTERN = re.compile(r'\b(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})\b')

# רשימת ה-GIDs של כל הטאבים שהגדרת
SHEET_GIDS = [
    "567882636",
    "729949464",
    "1615067993",
    "1468731794",
    "1794759740",
    "1968987360",
    "340965249",
    "1734632274",
    "1525331068",
]


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


def fetch_single_tab(gid: str, sheet_id: str, cutoff_date: datetime) -> list:
    """מושך טאב בודד ומחזיר את רשימת השורות המעובדות שלו"""
    csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    blocks = []
    try:
        res = requests.get(csv_url, timeout=10)
        if res.status_code != 200:
            return blocks
        res.encoding = "utf-8-sig"

        reader = csv.DictReader(StringIO(res.text))
        if not reader.fieldnames:
            return blocks

        clean_fieldnames = [col.strip() for col in reader.fieldnames if col]
        date_columns = [
            col for col in clean_fieldnames
            if any(k in col for k in ["פתיחה", "התחלה", "מועד", "תאריך", "ת."])
        ]

        for raw_row in reader:
            row = {k.strip(): (v.strip() if isinstance(v, str) else "") for k, v in raw_row.items() if k}
            # חובה ששם הקורס יהיה קיים
            course_name = row.get("שם הקורס") or row.get("שם תוכנית") or row.get("שם התכנית") or ""
            if not course_name:
                continue

            status = check_course_date_status(row, date_columns, cutoff_date)

            row_lines = []
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
                blocks.append("\n".join(row_lines))

    except Exception as e:
        logging.warning(f"Error reading gid={gid}: {e}")

    return blocks


def load_courses_catalog() -> str:
    sheet_id = "1emkWYVjlrhC0m67gz7b3MUXzsG3RtlX-kbg8eNvzBtc"
    today = datetime.now()
    cutoff_date = today - timedelta(days=14)

    all_blocks = []
    # הורדה במקביל של כל 9 הטאבים יחד כדי למנוע Timeout
    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = [executor.submit(fetch_single_tab, gid, sheet_id, cutoff_date) for gid in SHEET_GIDS]
        for f in futures:
            all_blocks.extend(f.result())

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
