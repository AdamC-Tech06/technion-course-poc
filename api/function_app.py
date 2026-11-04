import azure.functions as func
import json
import logging
import os
import re
import pandas as pd
from pathlib import Path
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
    excel_path = Path(__file__).parent / "Technion Course - Limor.xlsx"

    df = pd.read_excel(excel_path)
    df = df.fillna("")

    catalog = ""

    for _, row in df.iterrows():
        catalog += f"""

שם הקורס: {row.get('שם הקורס', '')}

תיאור:
{row.get('תיאור', '')}

תחום:
{row.get('תחום', '')}

תנאי קבלה:
{row.get('תנאי קבלה', '')}

מתכונת לימודים:
{row.get('מתכונת לימודים', '')}

קמפוס:
{row.get('קמפוס', '')}

מועד התחלה:
{row.get('מועד התחלה', '')}

לינק לעמוד התכנית:
{row.get('לינק לעמוד התכנית', '')}

==================================================
"""

    return catalog


@app.route(route="match-courses", methods=["POST", "OPTIONS"])
def match_courses(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("Course matcher API received a request.")

    if req.method == "OPTIONS":
        return func.HttpResponse(
            status_code=200,
            headers=CORS_HEADERS
        )

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
                {"error": "Missing OPENAI_API_KEY configuration on Azure"}
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
8. השתמש בלינק לעמוד התכנית כפי שמופיע בקטלוג.
9. כל הנימוקים בעברית מקצועית.

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
                "content": (
                    "נתוני השאלון של המועמד:\n"
                    + json.dumps(req_body, ensure_ascii=False)
                )
            }
        ]

        try:
            response = client.chat.completions.create(
                model=model_name,
                messages=messages,
                temperature=0.3,
                response_format={"type": "json_object"}
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

        error_details = (
            f"{type(e).__name__}: {str(e)} | "
            f"Underlying cause: {cause}"
        )

        logging.error(
            f"Error communicating with AI model: {error_details}"
        )

        return func.HttpResponse(
            json.dumps({"error": error_details}),
            status_code=500,
            headers=CORS_HEADERS,
            mimetype="application/json"
        )
