import re
from pathlib import Path

import pandas as pd
import win32com.client as win32


BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_FILE = BASE_DIR / "data" / "input" / "kb_feedback_task.xlsx"

SIGNATURE_HTML = """
<p>Thank you.</p>
<p>Kind regards,<br>
Lina Gencheva<br>
Knowledge Management Team</p>
"""


def clean_value(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def extract_first_name(submitted_by):
    """
    Example:
    Smith, John (50594949) -> John
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


def extract_kb_number(article):
    """
    Extract KB number from article text.
    Example:
    KB0021856 v4 - PTO Policy -> KB0021856
    """
    article = clean_value(article)
    match = re.search(r"\bKB\d+\b", article, re.IGNORECASE)
    if match:
        return match.group(0).upper()
    return ""


def build_kb_link(kb_number):
    if not kb_number:
        return ""
    return f"https://geaerospace.service-now.com/kb_view.do?sysparm_article={kb_number}"


def format_created_date(value):
    if pd.isna(value):
        return ""
    if hasattr(value, "strftime"):
        return value.strftime("%Y-%m-%d")
    return str(value).strip()


def build_html_body(first_name, number, description, created_date, kb_number, kb_link):
    if kb_number and kb_link:
        article_html = f'<a href="{kb_link}">{kb_number}</a>'
    else:
        article_html = kb_number

    return f"""
    <html>
      <body>
        <p>Hello {first_name},</p>

        <p>
          Regarding your request {number} regarding {article_html}, created on {created_date}.
          You have left this feedback: {description}
        </p>

        <p>
          Could you please further elaborate on how we can improve {article_html}?
        </p>

        {SIGNATURE_HTML}
      </body>
    </html>
    """


def main():
    print("INPUT_FILE:", INPUT_FILE)
    print("Exists:", INPUT_FILE.exists())

    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"Input file not found: {INPUT_FILE}")

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
            f"Missing required columns: {missing}\nAvailable columns: {list(df.columns)}"
        )

    outlook = win32.Dispatch("Outlook.Application")

    for _, row in df.iterrows():
        submitted_by = clean_value(row[submitted_by_col])
        first_name = extract_first_name(submitted_by)
        number = clean_value(row[number_col])
        article = clean_value(row[article_col])
        description = clean_value(row[description_col])
        created_date = format_created_date(row[created_date_col])

        kb_number = extract_kb_number(article)
        kb_link = build_kb_link(kb_number)

        subject = f"Follow-up on your feedback for {kb_number if kb_number else article}"
        html_body = build_html_body(
            first_name=first_name,
            number=number,
            description=description,
            created_date=created_date,
            kb_number=kb_number,
            kb_link=kb_link
        )

        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.HTMLBody = html_body
        mail.Save()   # saves to Drafts, does NOT send

        print(f"Draft created: {subject}")

    print("All drafts created successfully.")


if __name__ == "__main__":
    main()
