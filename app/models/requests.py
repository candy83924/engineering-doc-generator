from typing import Optional

from pydantic import BaseModel, Field


class GenerateRequest(BaseModel):
    """Request model for document generation (used with JSON body mode)."""

    documents: list[str] = Field(
        default=["acceptance_spec", "common_bid", "quote_fill", "compliance_filter"],
        description="要生成的文件類型清單",
    )
    project_name_override: Optional[str] = Field(
        None, description="手動指定工程名稱（覆蓋 PDF 解析結果）"
    )
    output_format: str = Field("zip", description="輸出格式：zip 或 individual")


VALID_DOCUMENT_TYPES = {
    "acceptance_spec",
    "common_bid",
    "quote_fill",
    "compliance_filter",
}
