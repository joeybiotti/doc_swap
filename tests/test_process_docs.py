import pytest
from pathlib import Path
from docx import Document
from src.process_docs import replace_text_in_docx

# Globals to match your function
FIND_TEXT = "World"
REPLACE_TEXT = "Universe"


@pytest.fixture
def tmp_docx(tmp_path):
    """Create a temporary .docx file for testing."""
    def _make_docx(text: str, filename: str = "test.docx") -> Path:
        file_path = tmp_path / filename
        doc = Document()
        doc.add_paragraph(text)
        doc.save(file_path)
        return file_path
    return _make_docx


def get_docx_text(path: Path) -> str:
    doc = Document(path)
    return "\n".join([p.text for p in doc.paragraphs])


def test_replacement(tmp_docx, monkeypatch):
    infile = tmp_docx("Hello World!")
    outfile = infile.parent / "out.docx"

    # Monkeypatch the globals in process_docs
    monkeypatch.setattr("src.process_docs.find_text", FIND_TEXT)
    monkeypatch.setattr("src.process_docs.replace_text", REPLACE_TEXT)

    replace_text_in_docx(infile, outfile)
    assert REPLACE_TEXT in get_docx_text(outfile)


def test_no_match(tmp_docx, monkeypatch):
    infile = tmp_docx("Hello again")
    outfile = infile.parent / "out.docx"

    monkeypatch.setattr("src.process_docs.find_text", "Missing")
    monkeypatch.setattr("src.process_docs.replace_text", "Replaced")

    replace_text_in_docx(infile, outfile)
    assert "Replaced" not in get_docx_text(outfile)
    assert get_docx_text(infile) == get_docx_text(outfile)


def test_empty_file(tmp_docx, monkeypatch):
    infile = tmp_docx("")
    outfile = infile.parent / "out.docx"

    monkeypatch.setattr("src.process_docs.find_text", "Anything")
    monkeypatch.setattr("src.process_docs.replace_text", "Replacement")

    replace_text_in_docx(infile, outfile)
    assert get_docx_text(outfile) == ""
