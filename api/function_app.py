import azure.functions as func
import json
import logging
import os
import re
from openai import OpenAI

app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)

CORS_HEADERS = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
}

COURSES_CATALOG = """
1. פיתוח Full Stack ו-Web (React, Node.js, Python, ענן, מסדי נתונים, Git) - קמפוס חיפה ושרונה | מועד קרוב: בחודש הקרוב | מתאים לחסרי רקע תכנותי או למעוניינים בהסבה מהירה | לינק: https://cont-ed.technion.ac.il
2. בינה מלאכותית ולמידת מכונה (AI, Deep Learning, Data Science, Python, LLMs) - קמפוס חיפה ושרונה | מועד קרוב: בשלושה חודשים | דורש רקע כמותי/תכנותי ואנגלית טובה מאוד | לינק: https://cont-ed.technion.ac.il
3. ניתוח נתונים ו-BI (Data Analyst, SQL, Power BI, Python, סטטיסטיקה) - קמפוס חיפה ושרונה | מועד קרוב: בחודש הקרוב | מתאים לבעלי רקע עסקי המעוניינים בתפקידי דאטה | לינק: https://cont-ed.technion.ac.il
4. סייבר ואבטחת מידע (Cyber Security, רשתות, הגנה ותקיפה SOC) - קמפוס חיפה ושרונה | מועד קרוב: בהקדם | מתאים לבעלי זיקה לתשתיות ו-IT | לינק: https://cont-ed.technion.ac.il
5. ניהול פרויקטים טכנולוגיים ומתודולוגיות Agile/Scrum - קמפוס שרונה וחיפה | מועד קרוב: גמיש | מתאים לבעלי ניסיון ניהולי או ארגוני | לינק: https://cont-ed.technion.ac.il
6. DevOps וארכיטקטורת ענן (Docker, Kubernetes, CI/CD, Azure/AWS) - קמפוס שרונה וחיפה | מועד קרוב: בחודש הקרוב | מיועד למפתחים או אנשי סיסטם | לינק: https://cont-ed.technion.ac.il
7. בדיקות תוכנה ואבטחת איכות (QA ידני ואוטומציה ב-Python) - קמפוס חיפה ושרונה | מועד קרוב: בהקדם | מתאים לחסרי רקע שרוצים להיכנס לתחום | לינק: https://cont-ed.technion.ac.il
8. ניהול מוצר (Product Management - אפיון, UI/UX ואסטרטגיה) - קמפוס שרונה | מועד קרוב: בשלושה חודשים | מתאים לבעלי רקע עסקי, שיווקי או עיצובי | לינק: https://cont-ed.technion.ac.il
"""

def clean_env(val: str) -> str:
    if not val:
        return ""
    return re.sub(r'[^\x21-\x7E]', '', val)

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
        client = OpenAI(api_key=api_key, base_url=base_url)

        system_prompt = f"""
אתה יועץ לימודים ומנהל אקדמי מומחה של בית הספר ללימודי המשך של הטכניון.
תפקידך לנתח את פרטי השאלון שמילא המועמד, להצליב אותם מול קטלוג הקורסים של היחידה, ולספק התאמה מדויקת ומנומקת.

קטלוג הקורסים:
{COURSES_CATALOG}

הנחיות קריטיות למבנה הפלט (חובה להקפיד ב-100%):
1. החזר תמיד אך ורק אובייקט JSON חוקי ותקני, ללא שום טקסט מקדים, ללא הסברים מסביב, וללא תגיות Markdown של ```json.
2. ה-JSON חייב להכיל מפתח ראשי יחיד בשם "results" המכיל מערך של קורסים.
3. שדות כל אובייקט במערך חייבים להיות בעברית בדיוק לפי השמות הבאים:
{{
  "results": [
    {{
      "שם הקורס": "שם הקורס מתוך הקטלוג",
      "ציון התאמה": 95,
      "תאריך התחלה": "מועד התחלה מותאם",
      "קמפוס": "חיפה / שרונה",
      "נימוק": "נימוק אישי מפורט בעברית המסביר מדוע הקורס מתאים לרקע, להשכלה ולמטרות שלו",
      "לינק לעמוד התכנית": "[https://cont-ed.technion.ac.il](https://cont-ed.technion.ac.il)"
    }}
  ]
}}
4. דרג לפחות 3 קורסים מומלצים עם ציון התאמה גבוה (מעל 70), ו-1 עד 2 קורסים עם ציון 0 עבור קורסים שאינם מתאימים בעליל לרקע המועמד (כדי שהממשק יציג גם קורסים שנפסלו).
5. "ציון התאמה" חייב להיות מספר שלם (Integer בין 0 ל-100).
6. כל הנימוקים והטקסטים חייבים להיכתב בעברית תקנית, רהוטה ומקצועית.
"""

        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": f"נתוני השאלון של המועמד:\n{json.dumps(req_body, ensure_ascii=False)}"}
        ]

        try:
            response = client.chat.completions.create(
                model=model_name,
                messages=messages,
                temperature=0.3,
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
        error_details = f"{type(e).__name__}: {str(e)} | Underlying cause: {cause}"
        logging.error(f"Error communicating with AI model: {error_details}")
        return func.HttpResponse(
            json.dumps({"error": error_details}),
            status_code=500,
            headers=CORS_HEADERS,
            mimetype="application/json"
        )
