"""Command line: pdfdocx FILE.pdf [FILE.pdf ...] [-o DIR] [--txt] [--combine NAME]"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pdfdocx import core
from pdfdocx.office import OfficeError


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="pdfdocx", description="PDF to an editable Word file or plain text.")
    parser.add_argument("pdf", nargs="+", type=Path, help="PDF file(s)")
    parser.add_argument("-o", "--out", type=Path, default=Path("."), help="output folder (default: current folder)")
    parser.add_argument("--txt", action="store_true", help="write plain text instead of .docx")
    parser.add_argument("--combine", metavar="NAME", help="put all PDFs into one NAME.docx, each on a new page")
    args = parser.parse_args(argv)
    try:
        if args.combine:
            print(core.combine(args.pdf, args.out, args.combine))
        else:
            for src in args.pdf:
                print(core.convert(src, "txt" if args.txt else "docx", args.out))
    except (ValueError, OfficeError) as exc:
        print(f"pdfdocx: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
