import os
import shutil
from docx import Document
from dotenv import load_dotenv

# Load .env variables
load_dotenv()

# Settings
INPUT_DIR = os.getenv("INPUT_DIR")
OUTPUT_DIR = os.getenv("OUTPUT_DIR")
FIND_TEXT = os.getenv("FIND_TEXT")
FIND_TEXTS = os.getenv("FIND_TEXTS").split(",")
REPLACE_TEXT = os.getenv("REPLACE_TEXT")
LOG_FILE = os.path.join(OUTPUT_DIR, "log.txt")

os.makedirs(OUTPUT_DIR, exist_ok=True)


def log(message: str):
    with open(LOG_FILE, "a") as f:
        f.write(message + "\n")
    print(message)


def replace_text_in_docx(docx_path, output_path, find_texts, replace_text):
    doc = Document(docx_path)
    replaced = False

    def replace_in_runs(runs):
        nonlocal replaced
        for run in runs:
            for find_text in find_texts:
                if find_text in run.text:
                    run.text = run.text.replace(find_text, replace_text)
                    replaced = True

    # Paragraphs
    for para in doc.paragraphs:
        replace_in_runs(para.runs)

    # Tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    replace_in_runs(para.runs)

    # Headers and Footers
    for section in doc.sections:
        for container in [section.header, section.footer]:
            for para in container.paragraphs:
                replace_in_runs(para.runs)

    doc.save(output_path)
    return replaced


# Clear log at start
with open(LOG_FILE, "w") as f:
    f.write("=== Processing Log ===\n")

# Main loop
for filename in os.listdir(INPUT_DIR):
    input_path = os.path.join(INPUT_DIR, filename)
    output_path = os.path.join(OUTPUT_DIR, filename)

    try:
        if filename.lower().endswith(".docx"):
            replaced = replace_text_in_docx(
                input_path, output_path, FIND_TEXTS, REPLACE_TEXT)
            if replaced:
                log(f"[MODIFIED] {filename}")
            else:
                shutil.copy(input_path, output_path)
                log(f"[COPIED] {filename} (no changes)")

        else:
            # Copy non-docx files unchanged
            shutil.copy(input_path, output_path)
            log(f"[SKIPPED] {filename} (not .docx)")

    except Exception as e:
        log(f"[ERROR] Unexpected error on {filename}: {e}")

log("=== Done! ===")
