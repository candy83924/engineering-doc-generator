"""PDF parsing orchestrator: combines table extraction, text extraction, and normalization."""

import logging
import re

from app.models.quote_data import (
    QuoteData,
    QuoteLineItem,
    QuoteMetadata,
    QuoteSection,
    QuoteSummary,
)
from app.parsing.normalizer import normalize_tables_to_quote_data
from app.parsing.table_extractor import extract_tables_from_pdf
from app.parsing.text_extractor import extract_metadata_fields, extract_text_from_pdf
from app.utils.chinese_utils import extract_number, normalize_unit

logger = logging.getLogger(__name__)


class PDFParser:
    """Orchestrates PDF parsing pipeline."""

    async def parse(self, pdf_bytes: bytes) -> QuoteData:
        """
        Parse a quotation PDF into structured QuoteData.

        Pipeline:
          1. Extract full text (with OCR fallback for scanned PDFs)
          2. Extract tables for line items
          3. Parse metadata from text
          4. If no tables found, parse items from OCR text
          5. Normalize and assemble into QuoteData
        """
        # Step 1: Extract raw text (includes OCR for scanned PDFs)
        raw_text = extract_text_from_pdf(pdf_bytes)
        if not raw_text:
            logger.warning("No text extracted from PDF")
            return QuoteData(
                metadata=_empty_metadata(),
                raw_text="",
                parse_confidence=0.0,
                parse_warnings=["無法從 PDF 擷取任何文字內容"],
            )

        # Step 2: Try table extraction
        tables = extract_tables_from_pdf(pdf_bytes)

        # Step 3: Parse metadata from text
        metadata_fields = extract_metadata_fields(raw_text)

        # Step 4: Normalize and assemble
        quote_data = None
        if tables:
            quote_data = normalize_tables_to_quote_data(
                tables, metadata_fields, raw_text
            )

        # If table normalization produced no items, try OCR spatial parser
        if not quote_data or not quote_data.get_real_items():
            logger.info("Table extraction produced no items, trying OCR spatial parser...")
            try:
                from app.parsing.ocr_parser import ocr_parse_pdf
                ocr_result = ocr_parse_pdf(pdf_bytes)
                if ocr_result and ocr_result.get_real_items():
                    quote_data = ocr_result
                    logger.info("OCR parser found %d items", len(ocr_result.get_real_items()))
            except Exception as e:
                logger.warning("OCR parser failed: %s", e)

        # If OCR also failed, try parsing from raw text
        if not quote_data or not quote_data.get_real_items():
            logger.info("OCR parser produced no items, trying text-based parsing...")
            quote_data = _parse_from_ocr_text(raw_text, metadata_fields)

        logger.info(
            "PDF parsed: %d items, confidence=%.2f, warnings=%d",
            len(quote_data.all_items),
            quote_data.parse_confidence,
            len(quote_data.parse_warnings),
        )
        return quote_data


def _empty_metadata():
    return QuoteMetadata(project_name="未命名工程")


def _parse_from_ocr_text(raw_text: str, metadata_fields: dict) -> QuoteData:
    """Parse structured quote data from OCR-extracted text lines."""
    lines = raw_text.strip().splitlines()
    warnings = []
    all_items = []
    sections = []
    item_seq = 0

    # Build metadata
    metadata = QuoteMetadata(
        project_name=metadata_fields.get("project_name", "未命名工程"),
        document_number=metadata_fields.get("document_number"),
        revision=metadata_fields.get("revision"),
        vendor_name=metadata_fields.get("vendor_name"),
        vendor_phone=metadata_fields.get("vendor_phone"),
        vendor_contact=metadata_fields.get("vendor_contact"),
        project_location=metadata_fields.get("project_location"),
        client_name=metadata_fields.get("client_name"),
        quote_date=metadata_fields.get("quote_date"),
    )

    # Strategy: scan lines looking for patterns that indicate line items
    # OCR text is grouped into lines. A typical item line contains:
    #   - A number or description
    #   - A unit (式/組/m/片/車/工...)
    #   - Quantities and prices (comma-separated numbers)
    #
    # We look for lines with price-like numbers and try to match them.

    i = 0
    current_section = None

    while i < len(lines):
        line = lines[i].strip()
        i += 1

        if not line:
            continue

        # Check for section headers (Chinese numerals alone)
        section_match = re.match(r"^([壹貳參肆伍陸柒捌玖拾一二三四五六七八九十]+)$", line)
        if section_match:
            # Next line might be the section name
            if i < len(lines):
                next_line = lines[i].strip()
                # Check if next line has prices — if yes, it's an item not a section name
                if not re.search(r"[\d,]+\.\d+|[\d,]{4,}", next_line):
                    if current_section and current_section.items:
                        sections.append(current_section)
                    current_section = QuoteSection(
                        section_number=section_match.group(1),
                        section_name=next_line,
                    )
                    i += 1
                    continue

        # Check for subtotal line
        if re.match(r"^(小計|合計|總\s*計|總計)", line):
            # Try to find the amount
            amounts = re.findall(r"[\d,]+(?:\.\d+)?", line)
            # Also check the next few tokens on the same conceptual line
            combined = line
            if i < len(lines):
                next_line = lines[i].strip()
                if re.match(r"^[\d,]+$", next_line) or "未稅" in next_line or "式" in next_line:
                    combined += " " + next_line
                    amounts = re.findall(r"[\d,]+(?:\.\d+)?", combined)

            total_val = 0.0
            if amounts:
                # Use the largest number as the subtotal
                nums = [float(a.replace(",", "")) for a in amounts]
                total_val = max(nums) if nums else 0.0

            subtotal_item = QuoteLineItem(
                seq=0,
                description=line,
                total_price=total_val,
                is_subtotal=True,
                section=current_section.section_name if current_section else None,
            )
            all_items.append(subtotal_item)
            if current_section:
                current_section.subtotal = total_val
            continue

        # Try to parse as a line item
        item = _try_parse_item_line(line, lines, i, current_section)
        if item:
            item_seq += 1
            item.seq = item_seq
            all_items.append(item)
            if current_section:
                current_section.items.append(item)

    if current_section and current_section.items:
        sections.append(current_section)

    # Calculate summary
    real_items = [it for it in all_items if not it.is_subtotal and not it.is_header]
    subtotal = sum(it.total_price for it in real_items)

    # Try to find grand total from text
    grand_total_match = re.search(r"總\s*計.*?([\d,]+)", raw_text)
    if grand_total_match:
        gt = float(grand_total_match.group(1).replace(",", ""))
        if gt > subtotal:
            subtotal = gt  # Use the stated total

    summary = QuoteSummary(
        subtotal=subtotal,
        tax_amount=round(subtotal * 0.05, 0),
        grand_total=round(subtotal * 1.05, 0),
        warranty_terms=metadata_fields.get("warranty_terms"),
        insurance_terms=metadata_fields.get("insurance_terms"),
        safety_terms=metadata_fields.get("safety_terms"),
        cleanup_terms=metadata_fields.get("cleanup_terms"),
        management_fee=metadata_fields.get("management_fee"),
        miscellaneous_fee=metadata_fields.get("miscellaneous_fee"),
        notes=metadata_fields.get("notes", []),
    )

    confidence = 0.7 if real_items else 0.3
    if not metadata.project_name or metadata.project_name == "未命名工程":
        confidence -= 0.1
    if len(real_items) < 2:
        confidence -= 0.1

    if not real_items:
        warnings.append("未能從 OCR 文字中解析出報價項目，建議人工確認")

    return QuoteData(
        metadata=metadata,
        sections=sections,
        all_items=all_items,
        summary=summary,
        raw_text=raw_text,
        parse_confidence=max(0.1, confidence),
        parse_warnings=warnings,
    )


def _try_parse_item_line(
    line: str, all_lines: list[str], next_idx: int,
    current_section: QuoteSection | None,
) -> QuoteLineItem | None:
    """Try to parse a line as a quotation item."""

    # Pattern: description followed by unit, quantity, price data
    # Common OCR patterns:
    #   "PVC無縫地磚,t=2mm(顏色送審)"  (description only)
    #   "m"  (unit on next line)
    #   "490"  (quantity)
    #   "1,800"  (unit price)
    #   "882,000"  (total price)
    #
    # Or combined:
    #   "PVC無縫地磚安裝 工 25 8,000 200,000"

    # Skip header/note lines
    skip_patterns = [
        r"^報價", r"^客戶", r"^工程名稱", r"^承辦", r"^TEL", r"^FA",
        r"^項目", r"^名\s*稱", r"^單位", r"^單價", r"^金額", r"^備註",
        r"^QUOTAT", r"^台灣", r"^Taiwan", r"^桃園", r"^TF\.", r"^333",
        r"^\d+\.\s*本工程", r"^得標", r"^本工程",
        r"^報價單號", r"^報價日期", r"^$",
    ]
    for pat in skip_patterns:
        if re.match(pat, line, re.IGNORECASE):
            return None

    # Skip very short lines that are just noise
    if len(line) <= 1 and not line.isdigit():
        return None

    # Try to find a description + numbers pattern
    # Look for lines that contain Chinese text AND numbers
    # Or lines that are purely descriptive (with numbers on nearby lines)

    # Extract all numbers from the line
    numbers_in_line = re.findall(r"[\d,]+(?:\.\d+)?", line)
    numbers_cleaned = []
    for n in numbers_in_line:
        n_str = n.replace(",", "").strip()
        if n_str:
            try:
                val = float(n_str)
                if val > 0:
                    numbers_cleaned.append(val)
            except ValueError:
                pass

    # Remove numbers to get the description part
    desc_part = re.sub(r"[\d,]+(?:\.\d+)?", "", line).strip()
    desc_part = re.sub(r"\s+", " ", desc_part).strip("., ")

    # Common unit patterns
    unit_pattern = r"(式|組|台|臺|個|支|片|只|車|工|批|次|趟|m²|m|cm|kg|L|坪|才|SET|PCS|EA|LOT)"
    unit_match = re.search(unit_pattern, line)

    # If line has a substantial description and at least one large number
    has_price = any(n >= 100 for n in numbers_cleaned)
    has_description = len(desc_part) >= 2 and re.search(r"[\u4e00-\u9fff]", desc_part)

    if not has_description and not has_price:
        return None

    # If this looks like a description-only line, skip it for now
    # (it might be part of a multi-line item handled elsewhere)
    if has_description and not has_price and not unit_match:
        return None

    # Try to parse unit, quantity, unit_price, total_price
    unit = ""
    quantity = 1.0
    unit_price = 0.0
    total_price = 0.0

    if unit_match:
        unit = unit_match.group(1)
        # Remove unit from description
        desc_part = line[:unit_match.start()].strip()
        desc_part = re.sub(r"[\d,]+(?:\.\d+)?", "", desc_part).strip("., ")
        # Numbers after unit are likely qty, unit_price, total_price
        after_unit = line[unit_match.end():]
        after_nums = re.findall(r"[\d,]+(?:\.\d+)?", after_unit)
        after_cleaned = _safe_parse_numbers(after_nums)

        # Also include numbers between description and unit
        before_unit = line[:unit_match.start()]
        before_nums = re.findall(r"[\d,]+(?:\.\d+)?", before_unit)
        before_cleaned = _safe_parse_numbers(before_nums)

        all_nums = before_cleaned + after_cleaned

        if len(all_nums) >= 3:
            quantity = all_nums[-3]
            unit_price = all_nums[-2]
            total_price = all_nums[-1]
        elif len(all_nums) == 2:
            # Could be qty+total or unit_price+total
            if all_nums[0] * 100 < all_nums[1]:  # qty is small, total is large
                quantity = all_nums[0]
                total_price = all_nums[1]
            else:
                unit_price = all_nums[0]
                total_price = all_nums[1]
        elif len(all_nums) == 1:
            total_price = all_nums[0]
    elif has_price:
        # No unit found, try to parse numbers
        if len(numbers_cleaned) >= 2:
            total_price = numbers_cleaned[-1]
            unit_price = numbers_cleaned[-2] if len(numbers_cleaned) >= 2 else 0
            quantity = numbers_cleaned[-3] if len(numbers_cleaned) >= 3 else 1
        elif len(numbers_cleaned) == 1:
            total_price = numbers_cleaned[0]

    if not desc_part or len(desc_part) < 2:
        return None

    # Skip if it looks like a note/condition rather than an item
    if desc_part.startswith("本工程") or desc_part.startswith("得標"):
        return None

    return QuoteLineItem(
        seq=0,  # Will be set by caller
        description=desc_part,
        unit=normalize_unit(unit) if unit else "式",
        quantity=quantity,
        unit_price=unit_price,
        total_price=total_price,
        section=current_section.section_name if current_section else None,
        confidence=0.7,  # OCR-parsed items have lower confidence
    )


def _safe_parse_numbers(num_strings: list[str]) -> list[float]:
    """Safely parse a list of number strings, skipping invalid ones."""
    result = []
    for n in num_strings:
        n_str = n.replace(",", "").strip()
        if n_str:
            try:
                result.append(float(n_str))
            except ValueError:
                pass
    return result
