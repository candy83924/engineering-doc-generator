"""Extract free-text metadata from PDF pages (headers, footers, notes)."""

import io
import logging
import re
from datetime import date

import pdfplumber

from app.utils.chinese_utils import normalize_fullwidth

logger = logging.getLogger(__name__)


def extract_text_from_pdf(pdf_bytes: bytes) -> str:
    """Extract full text from all pages of a PDF. Falls back to OCR for scanned PDFs."""
    full_text = []
    needs_ocr_pages = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text()
            if text and len(text.strip()) > 20:
                full_text.append(normalize_fullwidth(text))
            else:
                needs_ocr_pages.append(page_num)

    if needs_ocr_pages or not full_text:
        logger.info("PDF has %d pages needing OCR, running easyocr...",
                     len(needs_ocr_pages) or len(full_text) == 0)
        ocr_text = _ocr_pdf_pages(pdf_bytes, needs_ocr_pages or None)
        if ocr_text:
            full_text.append(normalize_fullwidth(ocr_text))

    return "\n".join(full_text)


def _ocr_pdf_pages(pdf_bytes: bytes, page_indices: list[int] | None = None) -> str:
    """OCR specific pages of a PDF using easyocr."""
    try:
        import easyocr
        from PIL import Image

        reader = easyocr.Reader(["ch_tra", "en"], gpu=False, verbose=False)
        all_text_blocks = []

        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            pages_to_process = page_indices if page_indices else range(len(pdf.pages))
            for page_num in pages_to_process:
                if page_num >= len(pdf.pages):
                    continue
                page = pdf.pages[page_num]
                # Convert page to image
                img = page.to_image(resolution=300)
                img_bytes = io.BytesIO()
                img.save(img_bytes, format="PNG")
                img_bytes.seek(0)
                img_pil = Image.open(img_bytes)

                # Run OCR
                results = reader.readtext(
                    img_bytes.getvalue(),
                    detail=1,
                    paragraph=False,
                )
                # Sort by vertical position (top to bottom), then horizontal
                results.sort(key=lambda r: (r[0][0][1], r[0][0][0]))

                page_lines = _group_ocr_into_lines(results)
                all_text_blocks.extend(page_lines)

        return "\n".join(all_text_blocks)

    except ImportError:
        logger.warning("easyocr not installed, cannot OCR scanned PDF")
        return ""
    except Exception as e:
        logger.error("OCR failed: %s", str(e))
        return ""


def _group_ocr_into_lines(results: list, y_threshold: float = 15) -> list[str]:
    """Group OCR text blocks into lines based on vertical position."""
    if not results:
        return []

    lines = []
    current_line = []
    current_y = None

    for bbox, text, conf in results:
        if conf < 0.1:  # Skip very low confidence
            continue
        top_y = bbox[0][1]  # Top-left Y coordinate

        if current_y is None or abs(top_y - current_y) > y_threshold:
            if current_line:
                lines.append(" ".join(current_line))
            current_line = [text]
            current_y = top_y
        else:
            current_line.append(text)

    if current_line:
        lines.append(" ".join(current_line))

    return lines


def extract_metadata_fields(full_text: str) -> dict:
    """Parse metadata fields from the extracted text."""
    meta = {}

    # Document number patterns (e.g., LIQP2603015)
    doc_num = re.search(r"([A-Z]{2,6}\d{5,10}(?:-?\w+)?)", full_text)
    if doc_num:
        meta["document_number"] = doc_num.group(1)

    # Version/revision
    rev = re.search(r"[Vv](\d+)", full_text)
    if rev:
        meta["revision"] = f"V{rev.group(1)}"

    # Project name patterns
    project_patterns = [
        r"工程名稱[：:]\s*(.+?)(?:\n|$)",
        r"案名[：:]\s*(.+?)(?:\n|$)",
        r"專案名稱[：:]\s*(.+?)(?:\n|$)",
        r"工程案號[：:]\s*(.+?)(?:\n|$)",
    ]
    for pat in project_patterns:
        m = re.search(pat, full_text)
        if m:
            meta["project_name"] = m.group(1).strip()
            break

    # Vendor name
    vendor_patterns = [
        r"(?:廠商|承攬商|施工廠商|報價廠商)[名稱]*[：:]\s*(.+?)(?:\n|$)",
        r"(?:公司名稱|供應商)[：:]\s*(.+?)(?:\n|$)",
    ]
    for pat in vendor_patterns:
        m = re.search(pat, full_text)
        if m:
            meta["vendor_name"] = m.group(1).strip()
            break

    # Location
    loc_patterns = [
        r"(?:施工地點|工程地點|施工位置|地點)[：:]\s*(.+?)(?:\n|$)",
        r"(?:廠區|位置)[：:]\s*(.+?)(?:\n|$)",
    ]
    for pat in loc_patterns:
        m = re.search(pat, full_text)
        if m:
            meta["project_location"] = m.group(1).strip()
            break

    # Client name
    client_patterns = [
        r"(?:業主|甲方|客戶)[名稱]*[：:]\s*(.+?)(?:\n|$)",
    ]
    for pat in client_patterns:
        m = re.search(pat, full_text)
        if m:
            meta["client_name"] = m.group(1).strip()
            break

    # Date
    date_patterns = [
        r"(?:報價日期|日期)[：:]\s*(\d{2,4})[./年-](\d{1,2})[./月-](\d{1,2})",
        r"(\d{4})[./年-](\d{1,2})[./月-](\d{1,2})",
        r"中華民國\s*(\d{2,3})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日",
    ]
    for pat in date_patterns:
        m = re.search(pat, full_text)
        if m:
            y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
            # ROC year conversion
            if y < 200:
                y += 1911
            try:
                meta["quote_date"] = date(y, mo, d)
            except ValueError:
                pass
            break

    # Contact info
    phone = re.search(r"(?:電話|TEL)[：:]\s*([\d\-() ]+)", full_text, re.IGNORECASE)
    if phone:
        meta["vendor_phone"] = phone.group(1).strip()

    # Terms extraction
    meta["notes"] = _extract_notes(full_text)
    meta["warranty_terms"] = _extract_term(full_text, ["保固", "warranty"])
    meta["insurance_terms"] = _extract_term(full_text, ["保險", "insurance"])
    meta["safety_terms"] = _extract_term(full_text, ["安全", "safety"])
    meta["cleanup_terms"] = _extract_term(full_text, ["清運", "清潔", "復原"])
    meta["management_fee"] = _extract_term(full_text, ["管理費"])
    meta["miscellaneous_fee"] = _extract_term(full_text, ["運雜費", "雜費"])

    return meta


def _extract_notes(text: str) -> list[str]:
    """Extract footer notes and conditions."""
    notes = []
    note_patterns = [
        r"備註[：:]\s*(.+?)(?:\n|$)",
        r"附註[：:]\s*(.+?)(?:\n|$)",
        r"說明[：:]\s*(.+?)(?:\n|$)",
        r"注意事項[：:]\s*(.+?)(?:\n|$)",
    ]
    for pat in note_patterns:
        for m in re.finditer(pat, text):
            note = m.group(1).strip()
            if note and len(note) > 2:
                notes.append(note)

    # Numbered notes: 1. xxx  2. xxx
    numbered = re.findall(r"\d+[.)]\s*(.+?)(?=\d+[.)]|\Z)", text, re.DOTALL)
    # Only add if they look like notes (near the end of text)
    return notes


def _extract_term(text: str, keywords: list[str]) -> str | None:
    """Extract a specific term/condition by keywords."""
    for kw in keywords:
        patterns = [
            rf"{kw}[：:]\s*(.+?)(?:\n|$)",
            rf"{kw}\s*[:]\s*(.+?)(?:\n|$)",
        ]
        for pat in patterns:
            m = re.search(pat, text, re.IGNORECASE)
            if m:
                return m.group(1).strip()
    return None
