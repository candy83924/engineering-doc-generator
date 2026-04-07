"""Normalize and assemble extracted data into a QuoteData model."""

import logging

from app.models.quote_data import (
    QuoteData,
    QuoteLineItem,
    QuoteMetadata,
    QuoteSection,
    QuoteSummary,
)
from app.parsing.table_extractor import map_columns
from app.utils.chinese_utils import (
    detect_section_marker,
    extract_number,
    is_chinese_number_prefix,
    normalize_unit,
)

logger = logging.getLogger(__name__)


def normalize_tables_to_quote_data(
    tables: list[dict],
    metadata_fields: dict,
    raw_text: str,
) -> QuoteData:
    """
    Assemble extracted tables and metadata into a unified QuoteData model.
    """
    warnings: list[str] = []
    all_items: list[QuoteLineItem] = []
    sections: list[QuoteSection] = []
    current_section: QuoteSection | None = None
    item_seq = 0

    for table_info in tables:
        headers = table_info["headers"]
        rows = table_info["rows"]
        col_map = map_columns(headers)

        if "description" not in col_map:
            warnings.append(
                f"第 {table_info['page']} 頁表格缺少品名/工程內容欄位，已跳過"
            )
            continue

        for row in rows:
            if len(row) <= col_map["description"]:
                continue

            desc = row[col_map["description"]].strip()
            if not desc:
                continue

            # Check if this is a section header
            section_marker = detect_section_marker(desc)
            if section_marker and _looks_like_section_header(row, col_map):
                if current_section and current_section.items:
                    sections.append(current_section)
                current_section = QuoteSection(
                    section_number=section_marker,
                    section_name=desc,
                )
                header_item = QuoteLineItem(
                    seq=0,
                    description=desc,
                    is_header=True,
                    section=desc,
                )
                all_items.append(header_item)
                continue

            # Check for subtotal rows
            if _is_subtotal_row(desc):
                total_val = _get_cell_number(row, col_map.get("total_price"))
                subtotal_item = QuoteLineItem(
                    seq=0,
                    description=desc,
                    total_price=total_val or 0.0,
                    is_subtotal=True,
                    section=current_section.section_name if current_section else None,
                )
                all_items.append(subtotal_item)
                if current_section:
                    current_section.subtotal = total_val
                continue

            # Regular line item
            item_seq += 1
            seq_val = _get_cell_number(row, col_map.get("seq"))
            unit = _get_cell(row, col_map.get("unit"))
            qty = _get_cell_number(row, col_map.get("quantity"))
            up = _get_cell_number(row, col_map.get("unit_price"))
            tp = _get_cell_number(row, col_map.get("total_price"))
            spec = _get_cell(row, col_map.get("specification"))
            remark = _get_cell(row, col_map.get("remark"))

            # Try to extract brand/model from description or spec
            brand, model = _extract_brand_model(desc, spec)

            item = QuoteLineItem(
                seq=int(seq_val) if seq_val else item_seq,
                description=desc,
                specification=spec if spec else None,
                brand=brand,
                model=model,
                unit=normalize_unit(unit) if unit else "式",
                quantity=qty or 1.0,
                unit_price=up or 0.0,
                total_price=tp or 0.0,
                remark=remark if remark else None,
                section=current_section.section_name if current_section else None,
            )

            # Validate price consistency
            if up and qty and tp:
                expected = round(up * qty, 0)
                if abs(expected - tp) > 1:
                    item.confidence = 0.8
                    warnings.append(
                        f"項次 {item.seq}「{desc[:20]}」單價×數量={expected} ≠ 複價{tp}"
                    )

            all_items.append(item)
            if current_section:
                current_section.items.append(item)

    # Don't forget the last section
    if current_section and current_section.items:
        sections.append(current_section)

    # Build metadata
    metadata = QuoteMetadata(
        document_number=metadata_fields.get("document_number"),
        revision=metadata_fields.get("revision"),
        vendor_name=metadata_fields.get("vendor_name"),
        vendor_phone=metadata_fields.get("vendor_phone"),
        project_name=metadata_fields.get("project_name", "未命名工程"),
        project_location=metadata_fields.get("project_location"),
        client_name=metadata_fields.get("client_name"),
        quote_date=metadata_fields.get("quote_date"),
    )

    # Build summary
    real_items = [i for i in all_items if not i.is_subtotal and not i.is_header]
    subtotal = sum(i.total_price for i in real_items)
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

    # Calculate parse confidence
    confidence = _calculate_confidence(all_items, warnings, raw_text)

    return QuoteData(
        metadata=metadata,
        sections=sections,
        all_items=all_items,
        summary=summary,
        raw_text=raw_text,
        parse_confidence=confidence,
        parse_warnings=warnings,
    )


def _looks_like_section_header(row: list[str], col_map: dict) -> bool:
    """A section header typically has no unit, quantity, or price."""
    unit = _get_cell(row, col_map.get("unit"))
    qty = _get_cell_number(row, col_map.get("quantity"))
    price = _get_cell_number(row, col_map.get("unit_price"))
    return not unit and not qty and not price


def _is_subtotal_row(desc: str) -> bool:
    subtotal_keywords = ["小計", "合計", "總計", "total", "subtotal"]
    desc_lower = desc.lower().strip()
    return any(kw in desc_lower for kw in subtotal_keywords)


def _get_cell(row: list[str], idx: int | None) -> str:
    if idx is None or idx >= len(row):
        return ""
    return row[idx].strip()


def _get_cell_number(row: list[str], idx: int | None) -> float | None:
    cell = _get_cell(row, idx)
    return extract_number(cell)


def _extract_brand_model(desc: str, spec: str | None) -> tuple[str | None, str | None]:
    """Try to extract brand and model from description or specification text."""
    combined = f"{desc} {spec or ''}"
    brand = None
    model = None

    # Common brand patterns in Taiwan engineering
    import re
    brand_match = re.search(
        r"(?:品牌|廠牌)[：:]\s*(\S+)", combined
    )
    if brand_match:
        brand = brand_match.group(1)

    model_match = re.search(
        r"(?:型號|型式|Model)[：:]\s*(\S+)", combined, re.IGNORECASE
    )
    if model_match:
        model = model_match.group(1)

    return brand, model


def _calculate_confidence(
    items: list[QuoteLineItem], warnings: list[str], raw_text: str
) -> float:
    """Calculate an overall parse confidence score."""
    if not items:
        return 0.1 if raw_text else 0.0

    real_items = [i for i in items if not i.is_subtotal and not i.is_header]
    if not real_items:
        return 0.3

    score = 1.0

    # Deduct for warnings
    score -= len(warnings) * 0.05

    # Deduct for items with low confidence
    low_conf = sum(1 for i in real_items if i.confidence < 0.9)
    score -= low_conf * 0.03

    # Deduct for items missing prices
    no_price = sum(1 for i in real_items if i.total_price == 0)
    if real_items:
        score -= (no_price / len(real_items)) * 0.2

    return max(0.1, min(1.0, score))
