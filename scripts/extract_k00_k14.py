#!/usr/bin/env python3
"""Extract K00-K14 ICD-10 hierarchy rows from the 医保2.0 PDF."""

from __future__ import annotations

import argparse
import csv
import re
from collections import Counter
from pathlib import Path

import pdfplumber


DEFAULT_PDF = Path("ICD-10医保2.0版.pdf")
DEFAULT_TSV = Path("k00_k14_full_hierarchy.tsv")

# PDF pages are 1-based in the document/UI, but pdfplumber uses 0-based indexes.
START_PAGE = 1579
END_PAGE = 1629

PDF_COLUMNS = [
    "章",
    "章代码范围",
    "章的名称",
    "节代码范围",
    "节名称",
    "类目代码",
    "类目名称",
    "亚目代码",
    "亚目名称",
    "诊断代码",
    "诊断名称",
]

OUTPUT_COLUMNS = [
    "章",
    "节",
    "类目代码",
    "类目名称",
    "亚目代码",
    "亚目名称",
    "诊断代码",
    "诊断名称",
]

K00_K14_RE = re.compile(r"^K(0[0-9]|1[0-4])")


def clean_cell(value: str | None) -> str:
    """Remove PDF layout line breaks while preserving the cell text."""
    return (value or "").replace("\n", "").strip()


def joined_label(*parts: str) -> str:
    return " ".join(part for part in parts if part)


def extract_rows(pdf_path: Path) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    seen_codes: set[str] = set()

    with pdfplumber.open(pdf_path) as pdf:
        for page_number in range(START_PAGE, END_PAGE + 1):
            page = pdf.pages[page_number - 1]
            for table in page.extract_tables():
                if not table:
                    continue

                header = [clean_cell(cell) for cell in table[0]]
                if not all(column in header for column in PDF_COLUMNS):
                    continue

                indexes = {column: header.index(column) for column in PDF_COLUMNS}
                for raw_row in table[1:]:
                    source = {
                        column: clean_cell(
                            raw_row[indexes[column]]
                            if indexes[column] < len(raw_row)
                            else ""
                        )
                        for column in PDF_COLUMNS
                    }
                    diagnosis_code = source["诊断代码"]
                    if not K00_K14_RE.match(diagnosis_code):
                        continue
                    if diagnosis_code in seen_codes:
                        continue

                    rows.append(
                        {
                            "章": joined_label(
                                source["章"],
                                source["章代码范围"],
                                source["章的名称"],
                            ),
                            "节": joined_label(source["节代码范围"], source["节名称"]),
                            "类目代码": source["类目代码"],
                            "类目名称": source["类目名称"],
                            "亚目代码": source["亚目代码"],
                            "亚目名称": source["亚目名称"],
                            "诊断代码": diagnosis_code,
                            "诊断名称": source["诊断名称"],
                        }
                    )
                    seen_codes.add(diagnosis_code)

    return rows


def write_tsv(rows: list[dict[str, str]], output_path: Path) -> None:
    with output_path.open("w", encoding="utf-8", newline="") as output_file:
        writer = csv.DictWriter(
            output_file,
            fieldnames=OUTPUT_COLUMNS,
            delimiter="\t",
            lineterminator="\n",
        )
        writer.writeheader()
        writer.writerows(rows)


def validate(rows: list[dict[str, str]]) -> None:
    missing_codes = [row for row in rows if not row["诊断代码"]]
    out_of_range = [
        row["诊断代码"]
        for row in rows
        if not K00_K14_RE.match(row["诊断代码"])
    ]
    duplicate_codes = [
        code
        for code, count in Counter(row["诊断代码"] for row in rows).items()
        if count > 1
    ]
    subtype = next(
        (row for row in rows if row["诊断代码"] == "K00.100x001"),
        None,
    )
    counts = Counter(row["诊断代码"][:3] for row in rows)

    print(f"Total row count: {len(rows)}")
    print(f"Missing 诊断代码: {len(missing_codes)}")
    print(f"Rows outside K00-K14: {len(out_of_range)}")
    print(f"Duplicated 诊断代码 removed/remaining: {len(duplicate_codes)}")
    if subtype:
        print(
            "Subtype K00.100x001 hierarchy: "
            f"章={subtype['章']} | 节={subtype['节']} | "
            f"类目代码={subtype['类目代码']} | 亚目代码={subtype['亚目代码']}"
        )
    else:
        print("Subtype K00.100x001 hierarchy: MISSING")

    print("Chapter-level count by K00-K14:")
    for number in range(15):
        prefix = f"K{number:02d}"
        print(f"{prefix}\t{counts[prefix]}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract K00-K14 full hierarchy rows from ICD-10医保2.0版.pdf."
    )
    parser.add_argument("--pdf", type=Path, default=DEFAULT_PDF)
    parser.add_argument("--tsv", type=Path, default=DEFAULT_TSV)
    args = parser.parse_args()

    rows = extract_rows(args.pdf)
    write_tsv(rows, args.tsv)
    print(f"Wrote {len(rows)} rows to {args.tsv}")
    validate(rows)


if __name__ == "__main__":
    main()
