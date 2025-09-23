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
REPLACE_TEXT = os.getenv("REPLACE_TEXT")
LOG_FILE = os.path.join(OUTPUT_DIR, "log.txt")

os.makedirs(OUTPUT_DIR, exist_ok=True)


def log(message: str):
    with open(LOG_FILE, "a") as f:
        f.write(message + "\n")
    print(message)


def replace_text_in_docx(docx_path, output_path, find_text, replace_text):
    try:
        doc = Document(docx_path)
        replaced = False

        # Replace in paragraphs
        for para in doc.paragraphs:
            for run in para.runs:
                if find_text in run.text:
                    run.text = run.text.replace(find_text, replace_text)
                    replaced = True
                    
        # Replace in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            if find_text in run.text:
                                run.text = run.text.replace(find_text, replace_text)
                                replaced = True
                                
        # Replace in headers and footers
        for section in doc.sections:
            header = section.header
            footer = section.footer

            for container in [header, footer]:
                for para in container.paragraphs:  # fix typo: paragraphs not paragraph
                    for run in para.runs:
                        if find_text in run.text:
                            run.text = run.text.replace(find_text, replace_text)
                            replaced = True

        # Always save the output file, even if nothing changed
        doc.save(output_path)
        return replaced

    except Exception as e:
        print(f"[ERROR] Failed on {docx_path}: {e}")
        return False

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
                input_path, output_path, FIND_TEXT, REPLACE_TEXT)
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
