import azure.functions as func
import json
import logging
import os
import re
import pandas as pd
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


def clean_env(val: str) -> str:
    if not val:
        return ""

    return re.sub(r'[^\x21-\x7E]', '', val)


def load_courses_catalog():

    sheet_id = "1emkWYVjlrhC0m67gz7b3MUXzsG3RtlX-kbg8eNvzBtc"
    gid = "1525331068"

    csv_url = (
        f"https://docs.google.com/spreadsheets/d/"
        f"{sheet_id}/export?format=csv&gid={gid}"
    )

    response = requests.get(csv_url, timeout=30)
    response.raise_for_status()

    df = pd.read_csv(StringIO(response.text))
    df = df.fillna("")

    logging.info(f"Columns found: {list(df.columns)}")

    date_column = None

    possible_columns = [
        "ת. פתיחה",
        "תאריך פתיחה",
        "מועד פתיחה",
        "תאריך התחלה"
    ]

    for col in possible_columns:
        if col in df.columns:
            date_column = col
            break

    logging.info(f"Date column found: {date_column}")

    today = datetime.today()
    cutoff_date = today - timedelta(days=14)

    filtered_rows = []

    for _, row in df.iterrows():

        keep_course = True

        if date_column:

            raw_date = str(
                row.get(date_column, "")
            ).strip()

            logging.info(
                f"Checking course date: {raw_date}"
            )

            if raw_date:

                parsed = None

                for fmt in [
                    "%d/%m/%Y",
                    "%d/%m/%y",
                    "%Y-%m-%d",
                    "%d.%m.%Y"
                \]:

                    try:
                        parsed = datetime.strptime(
                            raw_date,
                            fmt
                        )
                        break

                    except Exception:
                        pass

                if parsed:

                    if parsed < cutoff_date:

                        logging.info(
                            f"Excluded course. Date={raw_date}"
                        )

                        keep_course = False

                else:

                    logging.warning(
                        f"Could not parse date: {raw_date}"
                    )

        if keep_course:
            filtered_rows.append(row)

    logging.info(
        f"Courses after filter: {len(filtered_rows)}"
    )

    catalog = ""

    for row in filtered_rows:

        row_data = []

        for col in df.columns:

            value = str(row[col]).strip()

            if value:

                row_data.append(
                    f"{col}: {value}"
                )

        if row_data:

            catalog += "\n\n"
            catalog += "\n".join(row_data)
            catalog += "\n"
            catalog += "=" * 80

    return catalog

@app.route(route="match-courses", methods=["POST", "OPTIONS"])
def match_courses(req: func.HttpRequest) -> func.HttpResponse:

    logging.info(
        "Course matcher API received a request."
    )

    if req.method == "OPTIONS":

        return func.HttpResponse(
            status_code=200,
            headers=CORS_HEADERS
        )

    try:

        req_body = req.get_json()

    except Exception:

        return func.HttpResponse(
            json.dumps(
                {
                    "error": "Invalid JSON format in request body"
                }
            ),
            status_code=400,
            headers=CORS_HEADERS,
            mimetype="application/json"
        )

    api_key = clean_env(
        os.environ.get(
            "OPENAI_API_KEY",
            ""
        )
    )

    base_url = clean_env(
        os.environ.get(
            "OPENAI_BASE_URL",
            "https://api.groq.com/openai/v1"
        )
    )

    model_name = clean_env(
        os.environ.get(
            "OPENAI_MODEL",
            "llama-3.1-8b-instant"
        )
    )

    if not api_key:

        return func.HttpResponse(
            json.dumps(
                {
                    "error": "Missing OPENAI_API_KEY configuration on Azure"
                }
            ),
            status_code=500,
            headers=CORS_HEADERS,
            mimetype="application/json"
        )

    try:

        courses_catalog = load_courses_catalog()

        client = OpenAI(
            api_key=api_key,
            base_url=base_url
        )

        system_prompt = f"""
אתה יועץ לימודים ומנהל אקדמי מומחה של בית הספר ללימודי המשך של הטכניון.

תפקידך לנתח את פרטי השאלון שמילא המועמד,
להצליב אותם מול קטלוג הקורסים,
ולספק התאמה מדויקת ומנומקת.

קטלוג הקורסים:

{courses_catalog}

הנחיות קריטיות:

1. החזר אך ורק JSON תקין.
2. אין להחזיר Markdown.
3. אין להחזיר טקסט מחוץ ל-JSON.
4. החזר לפחות 3 קורסים מומלצים.
5. ניתן להחזיר 1-2 קורסים שנפסלו עם ציון 0.
6. ציון התאמה חייב להיות מספר שלם בין 0 ל-100.
7. השתמש אך ורק בקורסים המופיעים בקטלוג.
8. השתמש בלינקים כפי שמופיעים בקטלוג.
9. אל תמליץ על קורסים שההרשמה אליהם הסתיימה.
10. אל תמליץ על קורסים שאינם פעילים.
11. המלץ רק על קורסים הרלוונטיים להרשמה.
12. כל הנימוקים בעברית מקצועית.
13. השתמש במידע מכל העמודות בטבלה לצורך קבלת ההחלטה.
14. קורסים שתאריך הפתיחה שלהם ישן ביותר מ-14 יום כבר סוננו מראש ואין להתייחס אליהם.

מבנה הפלט:

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
            {
                "role": "system",
                "content": system_prompt
            },
            {
                "role": "user",
                "content":
                    "נתוני השאלון של המועמד:\n"
                    + json.dumps(
                        req_body,
                        ensure_ascii=False
                    )
            }
        ]

        try:

            response = client.chat.completions.create(
                model=model_name,
                messages=messages,
                temperature=0.3,
                response_format={
                    "type": "json_object"
                }
            )

        except Exception as model_err:

            logging.warning(
                f"Default model {model_name} failed. "
                f"Attempting fallback: {model_err}"
            )

            available_models = [
                m.id
                for m in client.models.list().data
                if not any(
                    x in m.id.lower()
                    for x in [
                        "whisper",
                        "guard",
                        "tts",
                        "vision",
                        "embed"
                    ]
                )
            ]

            if not available_models:
                raise model_err

            fallback_model = available_models[0]

            logging.info(
                f"Falling back to model: {fallback_model}"
            )

            response = client.chat.completions.create(
                model=fallback_model,
                messages=messages,
                temperature=0.3,
                response_format={
                    "type": "json_object"
                }
            )

        ai_content = response.choices[0].message.content

        return func.HttpResponse(
            ai_content,
            status_code=200,
            headers=CORS_HEADERS,
            mimetype="application/json"
        )

    except Exception as e:

        cause = (
            repr(e.__cause__)
            if e.__cause__
            else "None"
        )

        error_details = (
            f"{type(e).__name__}: {str(e)} | "
            f"Underlying cause: {cause}"
        )

        logging.error(
            f"Error communicating with AI model: {error_details}"
        )

        return func.HttpResponse(
            json.dumps(
                {
                    "error": error_details
                }
            ),
            status_code=500,
            headers=CORS_HEADERS,
            mimetype="application/json"
        )
