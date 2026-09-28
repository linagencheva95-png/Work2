import re
import pandas as pd

from config import INPUT_FILE, OUTPUT_FILE, OUTPUT_TEXT_DIR, SIGNATURE


def clean_value(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def extract_first_name(submitted_by):
    """
    Example:
    'Smith, John (50594949)' -> 'John'
    """
    submitted_by = clean_value(submitted_by)

    if not submitted_by:
        return ""

    if "," in submitted_by:
        part_after_comma = submitted_by.split(",", 1)[1].strip()
    else:
        part_after_comma = submitted_by

    first_name = part_after_comma.split("(", 1)[0].strip()
    return first_name


def format_created_date(value):
    if pd.isna(value):
        return ""

    if hasattr(value, "strftime"):
        return value.strftime("%Y-%m-%d")

    return str(value).strip()


def extract_kb_number(article):
    """
    Extract KB number from article text.
    Examples:
    'KB0021856 v3 - Sample Article' -> 'KB0021856'
    'KB0021856 version 5' -> 'KB0021856'
    'Article KB0021856 v2' -> 'KB0021856'
    """
    article = clean_value(article)

    match = re.search(r'\bKB\d+\b', article, re.IGNORECASE)
    if match:
        return match.group(0).upper()

    return ""


def build_kb_link(kb_number):
    if not kb_number:
        return ""
    return f"https://geaerospace.service-now.com/kb_view.do?sysparm_article={kb_number}"


def build_email_body(first_name, number, article, description, created_date, kb_number, kb_link):
    article_reference = kb_number if kb_number else article

    body = f"""Hello {first_name},

Regarding your request {number} regarding {article_reference}, created on {created_date}. You have left this feedback: {description}

Could you please further elaborate on how we can improve {article_reference}?"""

    if kb_link:
        body += f"""

Article link: {kb_link}"""

    body += f"""

{SIGNATURE}"""

    return body


def generate_emails():
    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"Input file not found: {INPUT_FILE}")

    OUTPUT_TEXT_DIR.mkdir(parents=True, exist_ok=True)

    df = pd.read_excel(INPUT_FILE)

    print("Columns found in Excel:")
    print(list(df.columns))

    submitted_by_col = "Submitted By"
    number_col = "Number"
    article_col = "Article"
    description_col = "Description"
    created_date_col = "Created"

    required_cols = [submitted_by_col, number_col, article_col, description_col, created_date_col]
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise ValueError(
            f"Missing required columns: {missing}\n"
            f"Available columns: {list(df.columns)}"
        )

    output_rows = []

    for idx, row in df.iterrows():
        submitted_by = clean_value(row[submitted_by_col])
        first_name = extract_first_name(submitted_by)
        number = clean_value(row[number_col])
        article = clean_value(row[article_col])
        description = clean_value(row[description_col])
        created_date = format_created_date(row[created_date_col])

        kb_number = extract_kb_number(article)
        kb_link = build_kb_link(kb_number)

        email_subject = f"Follow-up on your feedback for {kb_number if kb_number else article}"
        email_body = build_email_body(
            first_name,
            number,
            article,
            description,
            created_date,
            kb_number,
            kb_link
        )

        output_rows.append({
            "Submitted By": submitted_by,
            "First Name": first_name,
            "Number": number,
            "Article": article,
            "KB Number": kb_number,
            "KB Link": kb_link,
            "Description": description,
            "Created": created_date,
            "Email Subject": email_subject,
            "Email Body": email_body
        })

        safe_number = number.replace("/", "_").replace("\\", "_").replace(" ", "_")
        if not safe_number:
            safe_number = f"row_{idx + 1}"

        output_file = OUTPUT_TEXT_DIR / f"{safe_number}.txt"
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(f"Subject: {email_subject}\n\n{email_body}")

    output_df = pd.DataFrame(output_rows)
    output_df.to_csv(OUTPUT_FILE, index=False, encoding="utf-8-sig")

    print(f"Done. Output CSV saved to: {OUTPUT_FILE}")
    print(f"Individual email text files saved to: {OUTPUT_TEXT_DIR}")
