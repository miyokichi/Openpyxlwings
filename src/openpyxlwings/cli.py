"""Small command line helpers for openpyxlwings."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from openpyxlwings import ExcelWorkbook


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="openpyxlwings",
        description="Read Excel values quickly with openpyxl.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    sheets_parser = subparsers.add_parser("sheets", help="Print sheet names.")
    sheets_parser.add_argument("path", type=Path)

    range_parser = subparsers.add_parser("read", help="Read a range as JSON.")
    range_parser.add_argument("path", type=Path)
    range_parser.add_argument("sheet")
    range_parser.add_argument("range")

    args = parser.parse_args()

    with ExcelWorkbook(args.path) as workbook:
        if args.command == "sheets":
            print(json.dumps(workbook.sheet_names(), ensure_ascii=False))
        elif args.command == "read":
            print(json.dumps(workbook.read_range(args.sheet, args.range), ensure_ascii=False))
