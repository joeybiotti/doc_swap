# DocSwap

A Python tool for batch search-and-replace in `.docx` files, with logging and robust handling of multiple phrases.

## Features

- Recursively scans an input directory for `.docx` files.
- Supports replacing multiple phrases at once (comma-separated via environment variable).
- Processes paragraphs, tables, headers, and footers.
- Copies non-`.docx` files unchanged to the output directory.
- Logs all actions to `log.txt` in the output directory.
- Never modifies original files.

## Requirements

- Python 3.8+
- [`python-docx`](https://python-docx.readthedocs.io/en/latest/)
- [`python-dotenv`](https://pypi.org/project/python-dotenv/)

Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

1. Create a `.env` file in the project root with:

   ```
   INPUT_DIR=path/to/input
   OUTPUT_DIR=path/to/output
   FIND_TEXTS=phrase1,phrase2
   REPLACE_TEXT=replacement
   ```

2. Run the script:

   ```bash
   python src/process_docs.py
   ```

3. Check the output directory for processed files and `log.txt` for a summary.

## Project Structure

```
doc_swap/
├── main.py
├── src/
│   └── process_docs.py
├── tests/
│   └── test_process_docs.py
├── requirements.txt
├── README.md
└── .env
```

## Notes

- Only `.docx` files are modified; others are copied as-is.
- All changes and actions are logged.
- See `tests/test_process_docs.py` for usage examples and test cases.
