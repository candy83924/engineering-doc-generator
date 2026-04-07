"""CLI utility for testing full document generation pipeline."""

import asyncio
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from app.services.generation_service import GenerationService


async def main():
    if len(sys.argv) < 2:
        print("用法: python test_generate.py <報價單.pdf> [文件類型...]")
        print()
        print("文件類型（可多選）：")
        print("  acceptance_spec    - 驗收規範 (Word)")
        print("  common_bid         - 空白標單 (Excel)")
        print("  quote_fill         - 報價內容 (Excel)")
        print("  compliance_filter  - 規格篩選表 (Excel)")
        print()
        print("範例：")
        print("  python test_generate.py quote.pdf")
        print("  python test_generate.py quote.pdf common_bid quote_fill")
        sys.exit(1)

    pdf_path = Path(sys.argv[1])
    if not pdf_path.exists():
        print(f"找不到檔案：{pdf_path}")
        sys.exit(1)

    # Parse document types
    if len(sys.argv) > 2:
        doc_types = sys.argv[2:]
    else:
        doc_types = ["acceptance_spec", "common_bid", "quote_fill", "compliance_filter"]

    print(f"正在處理：{pdf_path.name}")
    print(f"要生成的文件：{', '.join(doc_types)}")
    print("-" * 50)

    pdf_bytes = pdf_path.read_bytes()
    service = GenerationService()

    try:
        zip_bytes, response = await service.generate(
            quote_pdf_bytes=pdf_bytes,
            requested_docs=doc_types,
        )

        print(f"\n生成結果：{'成功' if response.success else '失敗'}")
        print(f"工程名稱：{response.project_name}")
        print(f"解析信心度：{response.parse_confidence:.0%}")

        if response.documents:
            print(f"\n已生成 {len(response.documents)} 份文件：")
            for doc in response.documents:
                size_kb = doc.size_bytes / 1024
                print(f"  📄 {doc.filename} ({size_kb:.1f} KB)")
                for w in doc.warnings:
                    print(f"     ⚠ {w}")

        if response.errors:
            print("\n❌ 錯誤：")
            for e in response.errors:
                print(f"  - {e}")

        if response.parse_warnings:
            print("\n⚠ 解析警告：")
            for w in response.parse_warnings:
                print(f"  - {w}")

        # Save ZIP
        if zip_bytes:
            output_path = pdf_path.with_name(f"output_{pdf_path.stem}.zip")
            output_path.write_bytes(zip_bytes)
            print(f"\n✅ ZIP 已輸出至：{output_path}")

    finally:
        await service.close()


if __name__ == "__main__":
    asyncio.run(main())
