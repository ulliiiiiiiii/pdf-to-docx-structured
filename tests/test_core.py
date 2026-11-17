from pathlib import Path

import pytest

from pdfdocx import core, office
from pdfdocx.core import Block, Line, Page


def word(text, x0, top, size=10.0, font="Helvetica"):
    return {"text": text, "x0": x0, "x1": x0 + 6 * len(text), "top": top, "bottom": top + size,
            "size": size, "fontname": font}


def line(text, top, size=10.0, x0=50.0, bold=False):
    return Line(text, x0, top, top + size, size, bold)


@pytest.mark.parametrize(("front", "back", "joined"), [
    ("Vertrags-", "beginn am Montag", "Vertragsbeginn am Montag"),
    ("Online-", "Shop", "Online-Shop"),
    ("Haupt-", "und Nebenkosten", "Haupt- und Nebenkosten"),
    ("plain", "text", "plain text"),
])
def test_join_lines(front, back, joined):
    assert core.join_lines(front, back) == joined


def test_words_become_lines_left_to_right():
    lines = core._lines([word("world", 80, 100), word("Hello", 50, 100.5), word("next", 50, 120)])
    assert [ln.text for ln in lines] == ["Hello world", "next"]
    assert not lines[0].bold


def test_bold_line_needs_every_word_bold():
    bold = core._lines([word("Title", 50, 100, font="Helvetica-Bold"), word("Here", 90, 100, font="Arial,Bold")])
    mixed = core._lines([word("Half", 50, 100, font="Helvetica-Bold"), word("bold", 90, 100)])
    assert bold[0].bold and not mixed[0].bold


def test_paragraphs_lists_and_gaps():
    blocks = core._paragraphs([
        line("First line of a paragraph", 100),
        line("that goes on here.", 111),
        line("• a bullet point.", 140),
        line("1. a numbered point", 151),
        line("After a gap.", 190),
    ])
    assert [(b.kind, b.text) for b in blocks] == [
        ("p", "First line of a paragraph that goes on here."),
        ("ul", "a bullet point."),
        ("ol", "a numbered point"),
        ("p", "After a gap."),
    ]
    assert blocks[2].marker == "1."


def test_numbered_items_without_full_stop_stay_apart():
    blocks = core._paragraphs([
        line("1. Setup of the server", 100),
        line("2. Operation and monitoring", 111),
        line("3. Quartal is still part of the item", 122, x0=62),
    ])
    assert [(b.kind, b.text) for b in blocks] == [
        ("ol", "Setup of the server"),
        ("ol", "Operation and monitoring 3. Quartal is still part of the item"),
    ]


def test_page_numbers_and_running_headers_are_dropped_and_paragraphs_continue():
    pages = [
        Page(1000, [line("ACME report", 10), line("A sentence that does not", 100), line("1", 960)], []),
        Page(1000, [line("ACME report", 10), line("end here but goes on.", 100), line("Page 2 of 2", 960)], []),
    ]
    assert [b.text for b in core.assemble(pages)] == ["A sentence that does not end here but goes on."]


def test_headings_by_size_and_numbered_heading():
    blocks = [
        Block("p", "Big title", size=20), Block("p", "Smaller title", size=14),
        Block("p", "Body text " * 20, size=10), Block("ol", "Scope", size=14, marker="1."),
        Block("p", "Short bold", size=10, bold=True),
    ]
    core._headings(blocks)
    assert [(b.kind, b.text.strip()) for b in blocks] == [
        ("h1", "Big title"), ("h2", "Smaller title"), ("p", ("Body text " * 20).strip()),
        ("h2", "1. Scope"), ("h3", "Short bold"),
    ]


def test_html_and_text_output():
    blocks = [Block("h1", "Title"), Block("ul", "one"), Block("ul", "two"), Block("pagebreak"),
              Block("table", rows=[["a", "b"], ["c", "d & e"]]), Block("ol", "first"), Block("ol", "second")]
    page_html = core.to_html(blocks)
    assert "<h1>Title</h1>" in page_html
    assert "<ul><li>one</li><li>two</li></ul>" in page_html
    assert '<table border="1" cellpadding="4" cellspacing="0" style="page-break-before: always">' in page_html
    assert "<td>d &amp; e</td>" in page_html
    assert "<ol><li>first</li><li>second</li></ol>" in page_html
    assert core.to_text(blocks) == "Title\n\n- one\n\n- two\n\na\tb\nc\td & e\n\n1. first\n\n2. second\n"


def test_convert_rejects_pdf_without_text(monkeypatch, tmp_path):
    monkeypatch.setattr(core, "read", lambda path: [])
    with pytest.raises(ValueError, match="no text"):
        core.convert(tmp_path / "scan.pdf", "docx", tmp_path)


def test_office_call(monkeypatch, tmp_path):
    calls = []

    def fake_run(cmd, **kwargs):
        calls.append(cmd)
        (tmp_path / "doc.docx").write_bytes(b"x")

    monkeypatch.setattr(office, "SOFFICE", "soffice")
    result = office.convert(Path("/in/doc.html"), "docx:MS Word 2007 XML", tmp_path, infilter="HTML (StarWriter)",
                            run=fake_run)
    assert result == tmp_path / "doc.docx"
    assert calls[0][-5:] == ["--convert-to", "docx:MS Word 2007 XML", "--outdir", str(tmp_path), "/in/doc.html"]
    assert "--infilter=HTML (StarWriter)" in calls[0]
