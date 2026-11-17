"""Turn a PDF into an editable document. Paragraphs, headings, lists and tables are rebuilt from where
the text sits on the page (pdfplumber), written as HTML and converted to .docx with LibreOffice.
LibreOffice's own PDF import puts every line into a separate text frame, which nobody can edit in Word.

Limits: scanned PDFs (no text layer) are rejected, multi-column layouts are read line by line across
the columns, and tables without ruling lines come out as paragraphs.
"""

from __future__ import annotations

import html
import re
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

from pdfdocx import office

PAGE_NUMBER = re.compile(
    r"^(?:seite|page|стр\.?)?\s*[-\u2013]?\s*\d{1,4}\s*[-\u2013]?(?:\s*(?:von|of|/|из)\s*\d{1,4})?$", re.I
)
GLYPH = re.compile(r"^[\u2022\u00b7\u25aa\u25e6\u2023\u25cf\u25a0\u25a1]\s*")  # a bullet glyph always starts a list item
NUMBER = re.compile(r"^\d{1,2}[.)]\s+")  # "1. " only counts at the start of a paragraph
SENTENCE_END = (".", ":", ";", "!", "?")
CONJUNCTIONS = {"und", "oder", "bzw", "sowie", "bis", "and", "or", "и", "или"}  # "Haupt-" + "und" keeps its hyphen
MARGIN = 0.07  # top and bottom margin of a page, where page numbers, headers and footers live
GAP = 0.6  # more space than this many line heights between two lines starts a new paragraph


@dataclass
class Line:
    text: str
    x0: float
    top: float
    bottom: float
    size: float
    bold: bool


@dataclass
class Page:
    height: float
    lines: list[Line]
    tables: list[Block]


@dataclass
class Block:
    kind: str  # p | h1 | h2 | h3 | ul | ol | table | pagebreak
    text: str = ""
    top: float = 0
    x0: float = 0
    size: float = 0
    bold: bool = False
    rows: list[list[str]] = field(default_factory=list)  # tables only
    marker: str = ""  # the number of a list item ("1."), kept when the item turns out to be a heading


def _lines(words: list[dict]) -> list[Line]:
    """Words to lines: same height (centre within half a font size), left to right."""
    groups: list[tuple[float, list[dict]]] = []
    for w in sorted(words, key=lambda w: ((w["top"] + w["bottom"]) / 2, w["x0"])):
        middle = (w["top"] + w["bottom"]) / 2
        if groups and abs(groups[-1][0] - middle) <= 0.5 * max(w["size"], 1):
            groups[-1][1].append(w)
        else:
            groups.append((middle, [w]))
    out = []
    for _, ws in groups:
        ws.sort(key=lambda w: w["x0"])
        size = Counter({round(w["size"], 1): len(w["text"]) for w in ws}).most_common(1)[0][0]
        bold = all("bold" in (w.get("fontname") or "").lower() for w in ws if not GLYPH.fullmatch(w["text"]))
        out.append(
            Line(" ".join(w["text"] for w in ws), ws[0]["x0"], min(w["top"] for w in ws),
                 max(w["bottom"] for w in ws), size, bold)
        )
    return out


def join_lines(front: str, back: str) -> str:
    """Join two lines of a paragraph. A hyphen at the end of a line splits a word ("Vertrags-" + "beginn")
    and disappears; before a capitalised part ("Online-" + "Shop") it stays without a space; before
    "und", "oder", "and", "or" ("Haupt- und Nebenkosten") it stays with one."""
    first = back.split(" ", 1)[0]
    if re.search(r"\w-$", front):
        if first[:1].isupper():
            return front + back
        if first[:1].islower() and first.strip(",.;").lower() not in CONJUNCTIONS:
            return front[:-1] + back
    return f"{front} {back}"


def _paragraphs(lines: list[Line]) -> list[Block]:
    """Lines of a page to paragraphs and list items: a new block starts after extra space, with a
    different font or at a bullet; indented follow-up lines belong to the list item."""
    blocks: list[Block] = []
    previous: Line | None = None
    for line in lines:
        gap = previous is not None and line.top - previous.bottom > GAP * (previous.bottom - previous.top)
        different = previous is not None and (abs(line.size - previous.size) > 0.8 or line.bold != previous.bold)
        start = previous is None or gap or different or previous.text.endswith(SENTENCE_END)
        # a number aligned with the numbered item above starts the next item, even without a full stop
        after_item = bool(blocks) and blocks[-1].kind == "ol" and abs(line.x0 - blocks[-1].x0) <= 2
        item = "ul" if GLYPH.match(line.text) else "ol" if NUMBER.match(line.text) and (start or after_item) else None
        follows = blocks and blocks[-1].kind in ("ul", "ol") and not gap and line.x0 > blocks[-1].x0 + 2
        if item or (previous is None or gap or different) and not follows:
            text = (GLYPH if item == "ul" else NUMBER).sub("", line.text) if item else line.text
            marker = NUMBER.match(line.text).group(0).strip() if item == "ol" else ""
            blocks.append(Block(item or "p", text, line.top, line.x0, line.size, line.bold, marker=marker))
        else:
            blocks[-1].text = join_lines(blocks[-1].text, line.text)
        previous = line
    return blocks


def _inside(w: dict, box: tuple) -> bool:
    x, y = (w["x0"] + w["x1"]) / 2, (w["top"] + w["bottom"]) / 2
    return box[0] <= x <= box[2] and box[1] <= y <= box[3]


def read(path: Path) -> list[Block]:
    """All blocks of the PDF in reading order; empty when it has no text (scanned)."""
    import pdfplumber

    pages = []
    try:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                tables = page.find_tables()
                words = [
                    w for w in page.extract_words(extra_attrs=["size", "fontname"])
                    if not any(_inside(w, t.bbox) for t in tables)
                ]
                pages.append(Page(page.height, _lines(words), [
                    Block("table", top=t.bbox[1],
                          rows=[[" ".join((c or "").split()) for c in r] for r in t.extract() if r and any(r)])
                    for t in tables
                ]))
    except Exception as exc:  # broken or encrypted: pdfplumber has its own error types
        raise ValueError(f"PDF not readable ({type(exc).__name__})") from None
    return assemble(pages)


def assemble(pages: list[Page]) -> list[Block]:
    """Pages to blocks: without page numbers and running headers or footers (in the margin and the same
    on at least half of the pages, two or more), and a paragraph across a page break stays one."""

    def in_margin(line: Line, page: Page) -> bool:
        return line.top < MARGIN * page.height or line.bottom > (1 - MARGIN) * page.height

    counts = Counter(t for p in pages for t in {line.text for line in p.lines if in_margin(line, p)})
    repeated = {t for t, n in counts.items() if n >= 2 and n >= len(pages) / 2}
    out: list[Block] = []
    for p in pages:
        lines = [
            line for line in p.lines
            if not (in_margin(line, p) and (PAGE_NUMBER.fullmatch(line.text) or line.text in repeated))
        ]
        page_blocks = sorted(_paragraphs(lines) + [t for t in p.tables if t.rows], key=lambda b: b.top)
        if out and page_blocks and _continues(out[-1], page_blocks[0]):
            out[-1].text = join_lines(out[-1].text, page_blocks.pop(0).text)
        out += page_blocks
    _headings(out)
    return out


def _continues(front: Block, back: Block) -> bool:
    return (
        front.kind == back.kind == "p" and not front.text.endswith(SENTENCE_END) and back.text[:1].islower()
        and abs(front.size - back.size) <= 0.8
    )


def _headings(blocks: list[Block]) -> None:
    """A larger font than the body text, or bold, short and without a full stop, makes a heading; the
    largest size is h1, the next h2, everything else h3."""
    paragraphs = [b for b in blocks if b.kind == "p"]
    if not paragraphs:
        return
    body = Counter({b.size: 0 for b in paragraphs})
    for b in paragraphs:
        body[b.size] += len(b.text)
    body_size = body.most_common(1)[0][0]
    candidates = [
        b for b in paragraphs
        if b.size >= body_size + 1.5 or b.bold and len(b.text.split()) <= 12 and not b.text.endswith(".")
    ]
    # A numbered heading ("1. Scope" in heading size) is not a list item. Only the size decides here,
    # so bold list items stay a list.
    numbered = [b for b in blocks if b.kind == "ol" and b.size >= body_size + 1.5]
    for b in numbered:
        b.text = f"{b.marker} {b.text}".strip()
    candidates += numbered
    levels = sorted({b.size for b in candidates}, reverse=True)
    for b in candidates:
        b.kind = f"h{min(levels.index(b.size) + 1, 3)}"


PAGE_BREAK = ' style="page-break-before: always"'  # after a pagebreak block the next element starts a new page


def to_html(blocks: list[Block]) -> str:
    parts: list[str] = []
    items: list[Block] = []
    style = {"next": "", "list": ""}

    def close_list() -> None:
        if items:
            tag = items[0].kind
            parts.append(
                f"<{tag}{style['list']}>" + "".join(f"<li>{html.escape(b.text)}</li>" for b in items) + f"</{tag}>"
            )
            items.clear()

    def take_style() -> str:
        s, style["next"] = style["next"], ""
        return s

    for b in blocks:
        if b.kind == "pagebreak":
            close_list()
            style["next"] = PAGE_BREAK
            continue
        if b.kind in ("ul", "ol"):
            if items and items[0].kind != b.kind:
                close_list()
            if not items:
                style["list"] = take_style()
            items.append(b)
            continue
        close_list()
        if b.kind == "table":
            rows = "".join("<tr>" + "".join(f"<td>{html.escape(c)}</td>" for c in r) + "</tr>" for r in b.rows)
            parts.append(f'<table border="1" cellpadding="4" cellspacing="0"{take_style()}>{rows}</table>')
        else:
            tag = b.kind if b.kind.startswith("h") else "p"
            parts.append(f"<{tag}{take_style()}>{html.escape(b.text)}</{tag}>")
    close_list()
    return '<!doctype html><html><head><meta charset="utf-8"></head><body>' + "\n".join(parts) + "</body></html>"


def to_text(blocks: list[Block]) -> str:
    parts, number = [], 0
    for b in blocks:
        if b.kind == "pagebreak":
            continue
        number = number + 1 if b.kind == "ol" else 0
        if b.kind == "table":
            parts.append("\n".join("\t".join(r) for r in b.rows))
        else:
            parts.append({"ul": "- ", "ol": f"{number}. "}.get(b.kind, "") + b.text)
    return "\n\n".join(parts) + "\n"


def to_docx(blocks: list[Block], out: Path, name: str) -> Path:
    """The blocks as `out/<name>.docx`, through HTML so headings, lists and tables become real ones."""
    source = out / ".tmp" / f"{name}.html"
    source.parent.mkdir(parents=True, exist_ok=True)
    source.write_text(to_html(blocks), encoding="utf-8")
    return office.convert(source, "docx:MS Word 2007 XML", out, infilter="HTML (StarWriter)")


def convert(src: Path, target: str, out: Path) -> Path:
    """`src` (PDF) as a Word file (`docx`) or plain text (`txt`) in `out`."""
    blocks = read(src)
    if not blocks:
        raise ValueError("no text in the PDF (scanned?)")
    out.mkdir(parents=True, exist_ok=True)
    if target == "txt":
        path = out / f"{src.stem}.txt"
        path.write_text(to_text(blocks), encoding="utf-8")
        return path
    return to_docx(blocks, out, src.stem)


def combine(sources: list[Path], out: Path, name: str = "combined") -> Path:
    """Several PDFs in one Word file, in the given order, each starting on a new page."""
    everything: list[Block] = []
    for src in sources:
        blocks = read(src)
        if not blocks:
            raise ValueError(f"{src.name}: no text in the PDF (scanned?)")
        everything += [Block("pagebreak"), *blocks] if everything else blocks
    if not everything:
        raise ValueError("no PDF given")
    out.mkdir(parents=True, exist_ok=True)
    return to_docx(everything, out, name)
