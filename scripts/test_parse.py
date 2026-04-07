"""CLI utility for testing PDF parsing without starting the API server."""

import asyncio
import json
import sys
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from app.parsing.pdf_parser import PDFParser


async def main():
    if len(sys.argv) < 2:
        print("用法: python test_parse.py <報價單.pdf>")
        print("  解析 PDF 報價單並輸出結構化 JSON 資料")
        sys.exit(1)

    pdf_path = Path(sys.argv[1])
    if not pdf_path.exists():
        print(f"找不到檔案：{pdf_path}")
        sys.exit(1)

    print(f"正在解析：{pdf_path.name}")
    print("-" * 50)

    pdf_bytes = pdf_path.read_bytes()
    parser = PDFParser()
    quote_data = await parser.parse(pdf_bytes)

    # Print summary
    meta = quote_data.metadata
    print(f"工程名稱：{meta.project_name}")
    print(f"文件編號：{meta.document_number or '未偵測'}")
    print(f"廠商名稱：{meta.vendor_name or '未偵測'}")
    print(f"施工位置：{meta.project_location or '未偵測'}")
    print(f"報價日期：{meta.quote_date or '未偵測'}")
    print()
    print(f"解析信心度：{quote_data.parse_confidence:.0%}")
    print(f"總項目數：{len(quote_data.all_items)}")
    print(f"實際項目：{len(quote_data.get_real_items())}")
    print(f"分類數：{len(quote_data.sections)}")
    print()

    if quote_data.parse_warnings:
        print("⚠ 解析警告：")
        for w in quote_data.parse_warnings:
            print(f"  - {w}")
        print()

    # Print items
    real_items = quote_data.get_real_items()
    if real_items:
        print("報價項目：")
        for item in real_items:
            price_str = f"${item.total_price:,.0f}" if item.total_price else "N/A"
            print(f"  {item.seq:3d}. {item.description[:40]:<40s} {item.quantity:>6.1f} {item.unit:<4s} {price_str:>12s}")
    print()

    # Summary
    s = quote_data.summary
    print(f"小計：${s.subtotal:,.0f}")
    print(f"稅額：${s.tax_amount:,.0f}")
    print(f"合計：${s.grand_total:,.0f}")

    # Output full JSON
    output_path = pdf_path.with_suffix(".parsed.json")
    json_data = quote_data.model_dump(mode="json")
    output_path.write_text(json.dumps(json_data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n完整 JSON 已輸出至：{output_path}")


if __name__ == "__main__":
    asyncio.run(main())
