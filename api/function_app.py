# רשימת ה-GIDs של כל הגיליונות שרוצים לטעון:
# החלף או הוסף כאן את ה-gid של גיליון ההייטק (לרוב 0 או מספר אחר מה-URL)
SHEET_GIDS = [
    "567882636", "729949464", "1615067993", "1468731794", "1794759740", "1968987360", "340965249", "1734632274"           # הטאב הראשי / קורסי תעודה והייטק (ודא את ה-gid המדויק מהדפדפן)
    "1525331068",  # טאב תארים שניים
]

def load_courses_catalog() -> str:
    sheet_id = "1emkWYVjlrhC0m67gz7b3MUXzsG3RtlX-kbg8eNvzBtc"
    today = datetime.now()
    cutoff_date = today - timedelta(days=14)

    all_catalog_blocks = []

    for gid in SHEET_GIDS:
        csv_url = (
            f"https://docs.google.com/spreadsheets/d/"
            f"{sheet_id}/export?format=csv&gid={gid}"
        )

        try:
            response = requests.get(csv_url, timeout=30)
            if response.status_code != 200:
                continue
            response.encoding = "utf-8-sig"

            reader = csv.DictReader(StringIO(response.text))
            if not reader.fieldnames:
                continue

            clean_fieldnames = [col.strip() for col in reader.fieldnames if col]
            date_columns = [
                col for col in clean_fieldnames
                if any(keyword in col for keyword in ["פתיחה", "התחלה", "מועד", "תאריך", "ת."])
            ]

            for raw_row in reader:
                row = {k.strip(): (v.strip() if isinstance(v, str) else "") for k, v in raw_row.items() if k}
                if not any(row.values()):
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
                    all_catalog_blocks.append("\n".join(row_lines))

        except Exception as e:
            logging.warning(f"Failed to load sheet gid={gid}: {e}")

    return "\n\n" + ("\n" + "=" * 80 + "\n\n").join(all_catalog_blocks)
