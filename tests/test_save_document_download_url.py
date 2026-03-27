import asyncio
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from docx import Document

from word_document_server.tools.document_tools import save_document


def _create_docx(path: Path) -> None:
    doc = Document()
    doc.add_paragraph("Sample document for download URL tests.")
    doc.save(path)


def test_save_document_returns_plain_download_url_without_secret(tmp_path: Path, monkeypatch):
    source_path = tmp_path / "source.docx"
    output_dir = tmp_path / "out"
    output_dir.mkdir(parents=True, exist_ok=True)
    _create_docx(source_path)

    monkeypatch.setenv("DOC_OUTPUT_DIR", str(output_dir))
    monkeypatch.setenv("DOC_DOWNLOAD_BASE_URL", "https://example.com/files")
    monkeypatch.delenv("DOC_DOWNLOAD_SIGNING_SECRET", raising=False)
    monkeypatch.delenv("WORD_MCP_API_KEY", raising=False)

    result = asyncio.run(save_document("saved.docx", str(source_path)))

    assert "error" not in result
    assert result["download_url"] == "https://example.com/files/saved.docx"


def test_save_document_returns_signed_download_url_with_secret(tmp_path: Path, monkeypatch):
    source_path = tmp_path / "source.docx"
    output_dir = tmp_path / "out"
    output_dir.mkdir(parents=True, exist_ok=True)
    _create_docx(source_path)

    monkeypatch.setenv("DOC_OUTPUT_DIR", str(output_dir))
    monkeypatch.setenv("DOC_DOWNLOAD_BASE_URL", "https://example.com/files")
    monkeypatch.setenv("DOC_DOWNLOAD_SIGNING_SECRET", "signing-secret")

    result = asyncio.run(save_document("saved.docx", str(source_path)))

    assert "error" not in result
    parsed = urlparse(result["download_url"])
    params = parse_qs(parsed.query)
    assert parsed.path == "/files/saved.docx"
    assert "exp" in params
    assert "sig" in params
