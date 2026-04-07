"""Extract tabular data from PDF using pdfplumber."""

import logging
import re
from typing import Any

import pdfplumber

from app.utils.chinese_utils import extract_number, normalize_fullwidth

logger = logging.getLogger(__name__)

# Common header patterns for engineering quotation tables
HEADER_KEYWORDS = [
    "項次", "項目", "品名", "規格", "說明", "單位", "數量",
    "單價", "複價", "金額", "合計", "備註", "工程內容", "名稱",
]


def extract_tables_from_pdf(pdf_bytes: bytes) -> list[dict[str, Any]]:
    """
    Extract all tables from a PDF file.

    Returns a list of dicts, each containing:
      - page: int
      - headers: list[str]
      - rows: list[list[str]]
      - raw_table: the original pdfplumber table
    """
    import io

    results = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables(
                table_settings={
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "snap_tolerance": 5,
                    "join_tolerance": 5,
                }
            )

            if not tables:
                # Retry with text strategy if line-based extraction fails
                tables = page.extract_tables(
                    table_settings={
                        "vertical_strategy": "text",
                        "horizontal_strategy": "text",
                        "snap_tolerance": 8,
                        "join_tolerance": 8,
                    }
                )

            for table in tables:
                if not table or len(table) < 2:
                    continue

                cleaned = _clean_table(table)
                if not cleaned:
                    continue

                headers, data_rows = _identify_header(cleaned)
                results.append({
                    "page": page_num,
                    "headers": headers,
                    "rows": data_rows,
                    "raw_table": table,
                })

    return results


def _clean_table(table: list[list]) -> list[list[str]]:
    """Clean and normalize a raw pdfplumber table."""
    cleaned = []
    for row in table:
        cleaned_row = []
        for cell in row:
            if cell is None:
                cleaned_row.append("")
            else:
                text = normalize_fullwidth(str(cell).strip())
                text = re.sub(r"\s+", " ", text)
                cleaned_row.append(text)
        # Skip completely empty rows
        if any(c for c in cleaned_row):
            cleaned.append(cleaned_row)
    return cleaned


def _identify_header(table: list[list[str]]) -> tuple[list[str], list[list[str]]]:
    """Identify the header row and separate from data rows."""
    best_idx = 0
    best_score = 0

    for i, row in enumerate(table[:5]):  # Only check first 5 rows
        score = sum(
            1 for cell in row
            if any(kw in cell for kw in HEADER_KEYWORDS)
        )
        if score > best_score:
            best_score = score
            best_idx = i

    if best_score >= 2:
        return table[best_idx], table[best_idx + 1:]
    # No clear header found, use first row
    return table[0], table[1:]


def map_columns(headers: list[str]) -> dict[str, int]:
    """Map semantic column names to indices based on header text."""
    mapping = {}
    column_patterns = {
        "seq": ["項次", "項目", "序號", "No"],
        "description": ["品名", "名稱", "工程內容", "說明", "品名/規格", "工程項目"],
        "specification": ["規格", "型號", "規格說明"],
        "unit": ["單位"],
        "quantity": ["數量"],
        "unit_price": ["單價"],
        "total_price": ["複價", "金額", "合計", "小計"],
        "remark": ["備註"],
    }

    for i, header in enumerate(headers):
        header_clean = header.strip()
        for field, patterns in column_patterns.items():
            if field in mapping:
                continue
            for pat in patterns:
                if pat in header_clean:
                    mapping[field] = i
                    break

    return mapping
