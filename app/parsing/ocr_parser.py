"""Parse structured quote data from OCR results using spatial position."""

import io
import logging
import re

from app.models.quote_data import (
    QuoteLineItem,
    QuoteSection,
    QuoteSummary,
    QuoteMetadata,
    QuoteData,
)
from app.utils.chinese_utils import normalize_unit

logger = logging.getLogger(__name__)

# Lines containing these keywords are notes/conditions, not items
SKIP_KEYWORDS = [
    "本工程", "得標廠商", "填寫清楚", "塗改", "否則無效",
    "僅供參考", "自行評估", "結算數量", "以平面量計價",
]


def ocr_parse_pdf(pdf_bytes: bytes) -> QuoteData:
    """Parse a scanned PDF using OCR with spatial awareness."""
    try:
        import easyocr
        import pdfplumber
    except ImportError:
        logger.error("easyocr or pdfplumber not installed")
        return QuoteData(
            metadata=QuoteMetadata(project_name="未命名工程"),
            parse_confidence=0.0,
            parse_warnings=["OCR 模組未安裝"],
        )

    reader = easyocr.Reader(["ch_tra", "en"], gpu=False, verbose=False)
    all_blocks = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages):
            img = page.to_image(resolution=300)
            img_bytes = io.BytesIO()
            img.save(img_bytes, format="PNG")
            img_bytes.seek(0)

            results = reader.readtext(img_bytes.getvalue(), detail=1, paragraph=False)
            for bbox, text, conf in results:
                if conf < 0.05:
                    continue
                top_y = (bbox[0][1] + bbox[1][1]) / 2
                left_x = (bbox[0][0] + bbox[3][0]) / 2
                all_blocks.append({
                    "text": text.strip(),
                    "conf": conf,
                    "x": left_x,
                    "y": top_y,
                    "page": page_num,
                })

    if not all_blocks:
        return QuoteData(
            metadata=QuoteMetadata(project_name="未命名工程"),
            parse_confidence=0.0,
            parse_warnings=["OCR 未偵測到任何文字"],
        )

    all_blocks.sort(key=lambda b: (b["page"], b["y"], b["x"]))
    rows = _group_into_rows(all_blocks)
    raw_text = "\n".join(" ".join(b["text"] for b in row) for row in rows)

    metadata = _extract_metadata_from_rows(rows)
    header_idx, col_boundaries = _find_table_structure(rows)

    if header_idx is None:
        return QuoteData(
            metadata=metadata,
            raw_text=raw_text,
            parse_confidence=0.3,
            parse_warnings=["無法辨識表格結構，請確認 PDF 格式"],
        )

    items, sections = _parse_table_rows(rows, header_idx, col_boundaries)

    real_items = [it for it in items if not it.is_subtotal and not it.is_header]
    item_total = sum(it.total_price for it in real_items)

    gt_match = re.search(r"總\s*計.*?([\d,]+)", raw_text)
    grand_total_raw = 0.0
    if gt_match:
        try:
            grand_total_raw = float(gt_match.group(1).replace(",", ""))
        except ValueError:
            pass

    subtotal = grand_total_raw if grand_total_raw > item_total else item_total

    summary = QuoteSummary(
        subtotal=subtotal,
        tax_amount=round(subtotal * 0.05, 0),
        grand_total=round(subtotal * 1.05, 0),
    )

    confidence = 0.75 if len(real_items) >= 3 else 0.5
    warnings = []
    if not metadata.project_name or metadata.project_name == "未命名工程":
        warnings.append("未能辨識工程名稱")

    return QuoteData(
        metadata=metadata,
        sections=sections,
        all_items=items,
        summary=summary,
        raw_text=raw_text,
        parse_confidence=confidence,
        parse_warnings=warnings,
    )


def _group_into_rows(blocks, y_threshold=18):
    if not blocks:
        return []
    rows = []
    cur_row = [blocks[0]]
    cur_y = blocks[0]["y"]
    for block in blocks[1:]:
        if abs(block["y"] - cur_y) <= y_threshold and block["page"] == cur_row[0]["page"]:
            cur_row.append(block)
        else:
            cur_row.sort(key=lambda b: b["x"])
            rows.append(cur_row)
            cur_row = [block]
            cur_y = block["y"]
    if cur_row:
        cur_row.sort(key=lambda b: b["x"])
        rows.append(cur_row)
    return rows


def _extract_metadata_from_rows(rows):
    meta = QuoteMetadata(project_name="未命名工程")
    for row in rows[:15]:
        texts = [b["text"] for b in row]
        line = " ".join(texts)
        for i, t in enumerate(texts):
            if "客戶名稱" in t and i + 1 < len(texts):
                meta.client_name = texts[i + 1]
            if "工程名稱" in t and i + 1 < len(texts):
                meta.project_name = texts[i + 1]
        for t in texts:
            dm = re.search(r"(\d{4})/(\d{1,2})/(\d{1,2})", t)
            if dm:
                from datetime import date
                try:
                    meta.quote_date = date(int(dm.group(1)), int(dm.group(2)), int(dm.group(3)))
                except ValueError:
                    pass
            if re.search(r"(科技|工程|有限公司)", t) and "客戶" not in line and "欣興" not in t:
                if not meta.vendor_name:
                    meta.vendor_name = t
            doc_m = re.search(r"([A-Z]{1,6}\d{6,12})", t)
            if doc_m:
                meta.document_number = doc_m.group(1)
            phone_m = re.search(r"TEL[：:]?\s*([\d\-() ]+)", t, re.IGNORECASE)
            if phone_m:
                meta.vendor_phone = phone_m.group(1).strip()
    return meta


def _find_table_structure(rows):
    header_keywords = {"項目", "名稱", "規格", "規範", "單位", "數量", "單價", "金額", "備註", "長度"}
    best_idx = None
    best_score = 0
    for i, row in enumerate(rows):
        texts = [b["text"] for b in row]
        score = sum(1 for t in texts if any(kw in t for kw in header_keywords))
        if score > best_score:
            best_score = score
            best_idx = i
    if best_score < 2 or best_idx is None:
        return None, {}

    header_row = rows[best_idx]
    col_boundaries = {}
    for block in header_row:
        text = block["text"]
        x = block["x"]
        if "項目" in text or "項次" in text:
            col_boundaries["seq"] = x
        elif "名稱" in text or "規格" in text or "規範" in text:
            col_boundaries["description"] = x
        elif "單位" in text or "長度" in text:
            col_boundaries["unit"] = x
        elif "數量" in text:
            col_boundaries["quantity"] = x
        elif "單價" in text:
            col_boundaries["unit_price"] = x
        elif "金額" in text or "複價" in text:
            col_boundaries["total_price"] = x
        elif "備註" in text:
            col_boundaries["remark"] = x
    return best_idx, col_boundaries


def _parse_table_rows(rows, header_idx, col_boundaries):
    """Parse data rows using X-position zone assignment."""
    items = []
    sections = []
    current_section = None
    item_seq = 0

    desc_x = col_boundaries.get("description", 400)
    unit_x = col_boundaries.get("unit", 1000)

    data_rows = rows[header_idx + 1:]
    merged = _merge_multiline_rows(data_rows, unit_x)

    for row in merged:
        texts = [b["text"] for b in row]
        line = " ".join(texts)

        if not line.strip() or len(line.strip()) <= 1:
            continue
        if any(kw in line for kw in SKIP_KEYWORDS):
            continue

        # Subtotal / grand total
        if any(kw in line for kw in ["小計", "合計"]) or re.search(r"總\s*計", line):
            nums = re.findall(r"[\d,]+", line)
            total = 0.0
            for n in nums:
                try:
                    v = float(n.replace(",", ""))
                    if v > total:
                        total = v
                except ValueError:
                    pass
            items.append(QuoteLineItem(
                seq=0,
                description=re.sub(r"\s+", "", line.strip())[:20],
                total_price=total,
                is_subtotal=True,
            ))
            continue

        # Split blocks into zones
        seq_blocks = []
        desc_blocks = []
        numeric_blocks = []
        for b in row:
            if b["x"] < desc_x - 50:
                seq_blocks.append(b)
            elif b["x"] < unit_x - 50:
                desc_blocks.append(b)
            else:
                numeric_blocks.append(b)

        desc = " ".join(b["text"] for b in desc_blocks).strip()
        if not desc:
            desc = " ".join(b["text"] for b in seq_blocks + desc_blocks).strip()
        if not desc:
            continue

        if any(kw in desc for kw in ["客戶", "報價日期", "報價單號", "工程名稱", "承辦", "TEL", "QUOTAT"]):
            continue

        seq_text = " ".join(b["text"] for b in seq_blocks).strip()
        seq_num = _parse_num(seq_text)

        numeric_blocks.sort(key=lambda b: b["x"])
        unit_text, quantity, unit_price, total_price = _parse_numeric_columns(
            numeric_blocks, col_boundaries
        )

        # Description-only row with no numeric data — skip (already merged)
        if not numeric_blocks and not total_price:
            continue

        item_seq += 1
        item = QuoteLineItem(
            seq=int(seq_num) if seq_num and seq_num < 1000 else item_seq,
            description=desc,
            unit=normalize_unit(unit_text) if unit_text else "式",
            quantity=quantity or 1.0,
            unit_price=unit_price or 0.0,
            total_price=total_price or 0.0,
            section=current_section.section_name if current_section else None,
            confidence=0.7,
        )
        items.append(item)
        if current_section:
            current_section.items.append(item)

    if current_section and current_section.items:
        sections.append(current_section)

    return items, sections


def _merge_multiline_rows(rows, unit_x):
    """Merge description-only continuation rows into their parent item row."""
    if not rows:
        return []
    merged = []
    i = 0
    while i < len(rows):
        row = list(rows[i])
        while i + 1 < len(rows):
            next_row = rows[i + 1]
            next_line = " ".join(b["text"] for b in next_row)
            all_in_desc = all(b["x"] < unit_x - 50 for b in next_row)
            has_prices = bool(re.search(r"[\d,]{4,}", next_line))
            if any(kw in next_line for kw in SKIP_KEYWORDS):
                i += 1
                continue
            if all_in_desc and not has_prices and len(next_line.strip()) > 1:
                row.extend(next_row)
                i += 1
            else:
                break
        merged.append(row)
        i += 1
    return merged


def _parse_numeric_columns(num_blocks, col_boundaries):
    """Assign numeric blocks to unit/qty/price columns by closest X match."""
    unit = ""
    quantity = 0.0
    unit_price = 0.0
    total_price = 0.0

    if not num_blocks:
        return unit, quantity, unit_price, total_price

    col_x = {
        "unit": col_boundaries.get("unit", 0),
        "quantity": col_boundaries.get("quantity", 0),
        "unit_price": col_boundaries.get("unit_price", 0),
        "total_price": col_boundaries.get("total_price", 0),
    }

    for block in num_blocks:
        x = block["x"]
        text = block["text"].strip()

        # Find closest column
        dists = {col: abs(x - cx) for col, cx in col_x.items() if cx > 0}
        if not dists:
            continue
        closest = min(dists, key=dists.get)

        if closest == "unit":
            m = re.match(r"([^\d]+)(\d+)?$", text)
            if m:
                unit = m.group(1)
                if m.group(2):
                    quantity = float(m.group(2))
            else:
                unit = text
        elif closest == "quantity":
            v = _parse_num(text)
            if v is not None:
                quantity = v
        elif closest == "unit_price":
            v = _parse_num(text)
            if v is not None:
                unit_price = v
        elif closest == "total_price":
            v = _parse_num(text)
            if v is not None:
                total_price = v

    return unit, quantity, unit_price, total_price


def _parse_num(text):
    if not text:
        return None
    text = text.replace(",", "").replace(" ", "").strip()
    m = re.search(r"[\d.]+", text)
    if m:
        try:
            return float(m.group())
        except ValueError:
            pass
    return None
