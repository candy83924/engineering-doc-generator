"""CJK text utilities for normalizing Chinese text from PDF extraction."""

import re
import unicodedata


def normalize_fullwidth(text: str) -> str:
    """Convert full-width ASCII characters to half-width equivalents."""
    result = []
    for ch in text:
        code = ord(ch)
        # Full-width ASCII variants (！ to ～)
        if 0xFF01 <= code <= 0xFF5E:
            result.append(chr(code - 0xFEE0))
        # Full-width space
        elif code == 0x3000:
            result.append(" ")
        else:
            result.append(ch)
    return "".join(result)


def clean_whitespace(text: str) -> str:
    """Normalize whitespace: collapse multiple spaces, strip lines."""
    text = re.sub(r"[ \t]+", " ", text)
    lines = [line.strip() for line in text.splitlines()]
    return "\n".join(line for line in lines if line)


def extract_number(text: str) -> float | None:
    """Extract a numeric value from text, handling commas and full-width digits."""
    if not text:
        return None
    text = normalize_fullwidth(str(text).strip())
    text = text.replace(",", "").replace("，", "")
    # Remove currency symbols
    text = re.sub(r"[NT$元]", "", text)
    text = text.strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def is_chinese_number_prefix(text: str) -> bool:
    """Check if text starts with a Chinese ordinal (壹貳參肆伍陸柒捌玖拾)."""
    chinese_nums = "壹貳參肆伍陸柒捌玖拾零一二三四五六七八九十"
    text = text.strip()
    return bool(text) and text[0] in chinese_nums


def normalize_unit(unit: str) -> str:
    """Normalize common engineering units."""
    unit = normalize_fullwidth(unit.strip())
    unit_map = {
        "式": "式", "組": "組", "台": "台", "臺": "台",
        "只": "只", "個": "個", "支": "支", "片": "片",
        "公尺": "m", "米": "m", "M": "m", "m": "m",
        "公分": "cm", "CM": "cm", "cm": "cm",
        "平方公尺": "m²", "㎡": "m²", "M2": "m²", "m2": "m²",
        "才": "才", "坪": "坪",
        "公斤": "kg", "KG": "kg", "kg": "kg",
        "公升": "L", "L": "L",
        "天": "天", "日": "天", "月": "月",
        "批": "批", "趟": "趟", "次": "次",
        "SET": "組", "set": "組", "Set": "組",
        "PCS": "個", "pcs": "個", "EA": "個", "ea": "個",
        "LOT": "批", "lot": "批",
    }
    return unit_map.get(unit, unit)


def safe_filename(name: str, max_length: int = 80) -> str:
    """Create a filesystem-safe filename from a Chinese project name."""
    # Remove characters invalid in Windows filenames
    name = re.sub(r'[<>:"/\\|?*]', "", name)
    name = name.strip(". ")
    if len(name) > max_length:
        name = name[:max_length]
    return name or "unnamed"


def detect_section_marker(text: str) -> str | None:
    """Detect Chinese section numbering patterns like 壹、貳、or (一)、(二)."""
    text = text.strip()
    patterns = [
        r"^([壹貳參肆伍陸柒捌玖拾]+)[、\.\s]",
        r"^[\(（]([一二三四五六七八九十]+)[\)）]",
        r"^([一二三四五六七八九十]+)[、\.\s]",
        r"^([A-Z])[、\.\s]",
        r"^(\d+)[、\.\s]",
    ]
    for pattern in patterns:
        m = re.match(pattern, text)
        if m:
            return m.group(1)
    return None
