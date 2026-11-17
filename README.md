# pdf-to-docx-structured

Turns a PDF into a Word file you can actually edit. LibreOffice's own PDF import, like many
converters, puts every line into a separate text frame. This tool reads where the text sits on the
page and rebuilds the structure instead: paragraphs, headings, bullet and numbered lists, and tables.

## How it works

1. [pdfplumber](https://github.com/jsvine/pdfplumber) extracts every word with its position, size
   and font, plus ruled tables.
2. Words become lines, lines become paragraphs and list items, based on spacing, font changes,
   bullets and indentation. Words split at a line break are joined back ("Vertrags-" + "beginn"),
   compound words keep their hyphen ("Online-Shop"), and "Haupt- und Nebenkosten" stays as it is.
3. Page numbers and running headers or footers are dropped, and a paragraph that continues on the
   next page stays one paragraph.
4. Lines in a larger font, or short bold lines, become headings (three levels).
5. The result is written as HTML and converted to .docx by LibreOffice, so headings, lists and
   tables are real Word elements.

## Install

```sh
pip install git+https://github.com/ulliiiiiiiii/pdf-to-docx-structured
```

`.docx` output needs LibreOffice (`soffice` on the PATH). Plain text works without it.

## Usage

```sh
pdfdocx contract.pdf                    # contract.docx in the current folder
pdfdocx contract.pdf --txt -o out/      # out/contract.txt
pdfdocx a.pdf b.pdf --combine all       # all.docx, each PDF starting on a new page
```

```python
from pathlib import Path
from pdfdocx import core

core.convert(Path("contract.pdf"), "docx", Path("out"))
blocks = core.read(Path("contract.pdf"))  # paragraphs, headings, lists and tables as data
print(core.to_text(blocks))
```

## Limits

- Scanned PDFs without a text layer are rejected; run OCR first.
- Multi-column layouts are read line by line across the columns.
- Tables without ruling lines come out as paragraphs.

## License

MIT
