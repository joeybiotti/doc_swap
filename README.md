# DocxTextReplacer

A Python script to search for a specific phrase in `.docx` files, replace it with new text, and save the modified documents to a separate directory.

## Features

* Recursively scans a folder for `.docx` files.
* Replaces a given phrase with another phrase.
* Saves updated files to a specified output directory without modifying the originals.
* Easy to configure and run.

## Requirements

* Python 3.8+
* [`python-docx`](https://python-docx.readthedocs.io/en/latest/) library

Install dependencies:

```bash
pip install -r requirements.txt
```

Or manually:

```bash
pip install python-docx
```

## Usage

1. Activate your virtual environment:

```bash
python -m venv venv
source venv/bin/activate  # macOS/Linux
venv\Scripts\activate     # Windows
```

2. Edit `main.py` to specify:

   * The input directory containing `.docx` files
   * The phrase to search for
   * The replacement phrase
   * The output directory

3. Run the script:

```bash
python main.py
```

4. Check the output directory for modified files.

## Project Structure

```
DocxTextReplacer/
├── main.py
├── README.md
├── .gitignore
├── requirements.txt
└── venv/  # optional, virtual environment
```

## Notes

* Original files are never modified.
* Works only with `.docx` files, not legacy `.doc` files.
* Suitable for batch processing many documents quickly.
