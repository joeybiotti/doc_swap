
# DocSwap

A Python tool for batch search-and-replace in `.docx` files, with logging, dry-run support, and robust handling of multiple phrases.

## Features

- Recursively scans an input directory for `.docx` files.
- Supports replacing multiple phrases at once (comma-separated via environment variable).
- Processes paragraphs, tables, headers, and footers.
- Copies non-`.docx` files unchanged to the output directory.
- Logs all actions to `log.txt` in the output directory.
- Never modifies original files.
- Command-line interface with `--dry-run` option for previewing changes.
- Includes automated tests for core functionality.

## Requirements

- Python 3.8+
- [`python-docx`](https://python-docx.readthedocs.io/en/latest/)
- [`python-dotenv`](https://pypi.org/project/python-dotenv/)
- [`pytest`](https://docs.pytest.org/en/stable/)

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

2. Run the script via CLI:

   ```bash
   python main.py
   ```

   To preview changes without saving files, use:

   ```bash
   python main.py --dry-run
   ```

3. Check the output directory for processed files and `log.txt` for a summary.

## Project Structure

```
doc_swap/
├── main.py                  # CLI entry point
├── src/
│   └── process_docs.py      # Core logic for processing documents
├── tests/
│   └── test_process_docs.py # Pytest-based unit tests
├── requirements.txt
├── README.md
└── .env                     # Environment configuration
```

## Testing

Run all tests with:

```bash
pytest
```

## Notes

- Only `.docx` files are modified; others are copied as-is.
- All changes and actions are logged.
- See `tests/test_process_docs.py` for usage examples and test cases.
- The `--dry-run` option allows you to preview which files would be changed without making any modifications.
