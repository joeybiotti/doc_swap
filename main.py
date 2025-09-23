import os
import argparse
from dotenv import load_dotenv
from src.process_docs import replace_text_in_docx, batch_process_docs

def main():
    parser = argparse.ArgumentParser(description="Replace text in .docx files.")
    parser.add_argument("--dry-run", action="store_true", help="Preview changes without saving files")
    args = parser.parse_args()

    batch_process_docs(dry_run=args.dry_run)

    # Load .env variables
    load_dotenv()
    input_dir = os.getenv("INPUT_DIR")
    output_dir = os.getenv("OUTPUT_DIR")
    find_texts = [t.strip() for t in os.getenv("FIND_TEXTS", "").split(",")]
    replace_text = os.getenv("REPLACE_TEXT")

    os.makedirs(output_dir, exist_ok=True)

    for file in os.listdir(input_dir):
        if file.endswith(".docx"):
            input_path = os.path.join(input_dir, file)
            output_path = os.path.join(output_dir, file)
            replace_text_in_docx(input_path, output_path, find_texts, replace_text, dry_run=args.dry_run)
            print(f"{'[DRY RUN] ' if args.dry_run else ''}Processed {file}")

if __name__ == "__main__":
    main()
