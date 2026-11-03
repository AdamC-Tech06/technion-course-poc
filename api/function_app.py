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
1. פיתוח Full Stack ו-Web (React, Node.js, Python, ענן, מסדי נתונים, Git) - מתאים למי שרוצה להיכנס לפיתוח תוכנה מאפס או להסב מקצוע.
2. בינה מלאכותית ולמידת מכונה (AI & Machine Learning, Data Science, Python, Deep Learning, LLMs) - מיועד לבעלי רקע כמותי, תכנותי או אקדמי שרוצים להתמחות ב-AI.
3. ניתוח נתונים ו-BI (Data Analyst, SQL, Power BI, Python, סטטיסטיקה יישומית) - מתאים לחסרי רקע טכנולוגי עמוק המעוניינים בתפקידי דאטה עסקיים.
4. סייבר ואבטחת מידע (Cyber Security, רשתות, הגנה ותקיפה, SOC, ענן) - מיועד לבעלי זיקה לתשתיות, IT ורשתות תקשורת.
5. ניהול פרויקטים טכנולוגיים ומתודולוגיות Agile/Scrum - מתאים לבעלי ניסיון ניהולי או ארגוני המעוניינים להוביל פרויקטים בעולמות הטק.
6. DevOps וארכיטקטורת ענן (Cloud Engineering, Docker, Kubernetes, CI/CD, Azure/AWS) - מיועד למפתחים או אנשי סיסטם המעוניינים בהתמחות ענן.
7. בדיקות תוכנה ואבטחת איכות (QA ידני ואוטומציה ב-Python/Selenium) - כרטיס כניסה מהיר לעולם ההייטק ללא צורך בידע מוקדם.
8. ניהול מוצר (Product Management) - אפיון, חוויית משתמש (UX), אסטרטגיה עסקית והובלת מוצרים דיגיטליים.
"""

def clean_env(val: str) -> str:
    """מסיר רווחים ותווי כיווניות נסתרים (RTL/LTR) שהועתקו מהדפדפן"""
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
    model_name = clean_env(os.environ.get("OPENAI_MODEL", "llama-3.3-70b-versatile"))

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

הנחיות קריטיות למענה:
1. החזר תמיד אך ורק אובייקט JSON חוקי ותקני, ללא שום טקסט מקדים, ללא הסברים מסביב, וללא תגיות Markdown של ```json.
2. ה-JSON חייב להיות במבנה הבא בדיוק:
{{
  "candidateName": "שם המועמד (אם סופק)",
  "topRecommendation": "שם הקורס בעל ההתאמה הגבוהה ביותר",
  "summary": "סיכום קצר, מקצועי ומעודד (בעברית) המסביר מדוע קורס זה נבחר עבורו",
  "matches": [
    {{
      "courseName": "שם הקורס",
      "matchScore": 95,
      "reason": "נימוק מפורט ומותאם אישית להעדפותיו ולרקע שלו (בעברית)",
      "fitLevel": "גבוהה מאוד / גבוהה / בינונית"
    }}
  ]
}}
3. דרג לפחות 3 קורסים מתוך הקטלוג לפי אחוזי התאמה יורדים (0 עד 100).
4. כל הטקסטים והנימוקים חייבים להיכתב בעברית תקנית, רהוטה ומזמינה.
"""

        response = client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"נתוני השאלון של המועמד:\n{json.dumps(req_body, ensure_ascii=False)}"}
            ],
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
