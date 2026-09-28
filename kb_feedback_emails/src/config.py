from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_FILE = BASE_DIR / "data" / "input" / "kb_feedback_task.xlsx"
OUTPUT_FILE = BASE_DIR / "data" / "output" / "kb_feedback_emails.csv"
OUTPUT_TEXT_DIR = BASE_DIR / "data" / "output" / "emails"

SIGNATURE = """Thank you.

Kind regards,

Lina Gencheva

Knowledge Management Team"""