from __future__ import annotations

from datetime import date
from enum import Enum
from typing import Optional

from pydantic import BaseModel, Field


class WorkCategory(str, Enum):
    CLEANROOM = "無塵室裝修"
    HVAC = "空調工程"
    ELECTRICAL = "電力工程"
    PIPING = "管路工程"
    MONITORING = "中央監控"
    FIRE_PROTECTION = "消防工程"
    RAISED_FLOOR = "高架地板"
    DOOR = "門窗工程"
    WATERPROOF = "防水工程"
    PAINTING = "油漆工程"
    PLUMBING = "給排水工程"
    CIVIL = "土建工程"
    EQUIPMENT = "設備維修"
    CLEANING = "清潔工程"
    OTHER = "其他"


class QuoteLineItem(BaseModel):
    """Single line item from the quotation."""

    seq: int = Field(..., description="項次序號")
    item_code: Optional[str] = Field(None, description="項目編號")
    description: str = Field(..., description="品名/工程內容")
    specification: Optional[str] = Field(None, description="規格說明")
    brand: Optional[str] = Field(None, description="品牌")
    model: Optional[str] = Field(None, description="型號")
    unit: str = Field("式", description="單位")
    quantity: float = Field(1.0, description="數量")
    unit_price: float = Field(0.0, description="單價")
    total_price: float = Field(0.0, description="複價")
    remark: Optional[str] = Field(None, description="備註")
    is_subtotal: bool = Field(False, description="是否為小計列")
    is_header: bool = Field(False, description="是否為大項標題列")
    section: Optional[str] = Field(None, description="所屬大項名稱")
    confidence: float = Field(1.0, description="解析信心度 0-1")


class QuoteSection(BaseModel):
    """A major section/category within the quotation."""

    section_number: str = Field(..., description="大項編號，如 壹、貳、A、1")
    section_name: str = Field(..., description="大項名稱")
    items: list[QuoteLineItem] = Field(default_factory=list)
    subtotal: Optional[float] = None


class QuoteMetadata(BaseModel):
    """Header/metadata extracted from the quotation PDF."""

    document_number: Optional[str] = Field(None, description="文件編號")
    revision: Optional[str] = Field(None, description="版本")
    vendor_name: Optional[str] = Field(None, description="廠商名稱")
    vendor_contact: Optional[str] = Field(None, description="廠商聯絡人")
    vendor_phone: Optional[str] = Field(None, description="廠商電話")
    vendor_address: Optional[str] = Field(None, description="廠商地址")
    project_name: str = Field("未命名工程", description="工程名稱")
    project_location: Optional[str] = Field(None, description="施工位置")
    client_name: Optional[str] = Field(None, description="業主名稱")
    quote_date: Optional[date] = None
    valid_until: Optional[date] = None
    currency: str = Field("TWD", description="幣別")
    work_categories: list[WorkCategory] = Field(default_factory=list)


class QuoteSummary(BaseModel):
    """Financial summary of the quotation."""

    subtotal: float = Field(0.0, description="未稅小計")
    tax_rate: float = Field(0.05, description="稅率")
    tax_amount: float = Field(0.0, description="稅額")
    grand_total: float = Field(0.0, description="含稅總計")
    discount: Optional[float] = None
    warranty_terms: Optional[str] = Field(None, description="保固條件")
    insurance_terms: Optional[str] = Field(None, description="保險條件")
    safety_terms: Optional[str] = Field(None, description="安全條件")
    cleanup_terms: Optional[str] = Field(None, description="清運條件")
    management_fee: Optional[str] = Field(None, description="管理費")
    miscellaneous_fee: Optional[str] = Field(None, description="運雜費")
    notes: list[str] = Field(default_factory=list, description="備註/附加條件")


class QuoteData(BaseModel):
    """
    Complete structured representation of a construction/engineering quotation.
    This is the shared data model consumed by all four generators.
    """

    metadata: QuoteMetadata
    sections: list[QuoteSection] = Field(default_factory=list)
    all_items: list[QuoteLineItem] = Field(default_factory=list)
    summary: QuoteSummary = Field(default_factory=QuoteSummary)
    raw_text: Optional[str] = Field(None, description="完整原始文字（LLM 備援用）")
    parse_confidence: float = Field(1.0, description="整體解析信心度 0-1")
    parse_warnings: list[str] = Field(default_factory=list)
    project_brief: Optional[str] = Field(None, description="案子概略說明")

    def get_real_items(self) -> list[QuoteLineItem]:
        """Return only actual line items (exclude subtotals and headers)."""
        return [i for i in self.all_items if not i.is_subtotal and not i.is_header]

    def get_project_description(self) -> str:
        """Build a concise project description from metadata and items."""
        parts = [self.metadata.project_name]
        if self.metadata.project_location:
            parts.append(f"位置：{self.metadata.project_location}")
        items_desc = [i.description for i in self.get_real_items()[:10]]
        if items_desc:
            parts.append(f"主要項目：{'、'.join(items_desc)}")
        return "；".join(parts)


# ---------- LLM output models ----------

class AcceptanceCriterion(BaseModel):
    """Single acceptance criterion for the 驗收標準表."""

    category: str = Field(..., description="驗收項目 (e.g. 材料驗收, 結構固定驗收)")
    standard: str = Field(..., description="驗收標準說明 (detailed multi-point text)")


class AcceptanceContent(BaseModel):
    """LLM-generated content for acceptance specification (5-section format)."""

    # Section 1: 維修設備基本資訊 (table)
    equipment_location: str = Field("", description="設備位置")
    equipment_type: str = Field("", description="設備類型")
    spec_info: str = Field("", description="馬力規格")
    problem_description: str = Field("", description="問題說明")
    repair_needs: str = Field("", description="維修需求")

    # Section 2: 施工處置與安全規範 (sub-headed paragraphs)
    contractor_requirements: list[str] = Field(default_factory=list, description="施工廠商需具備")
    preparation_steps: list[str] = Field(default_factory=list, description="作業前準備")
    construction_content: list[str] = Field(default_factory=list, description="施工內容")
    acceptance_procedures: list[str] = Field(default_factory=list, description="驗收作業")

    # Section 3: 施工與驗收標準表
    acceptance_criteria: list[AcceptanceCriterion] = Field(
        default_factory=list, description="驗收標準表"
    )

    # Section 4: 施工保險要求
    insurance_text: str = Field("", description="施工保險要求")

    # Section 5: 施工區域
    work_area_text: str = Field("", description="施工區域")


class QuoteFillContent(BaseModel):
    """LLM-generated content for 欣興 internal requisition form."""

    purchase_name: str = Field("", description="請購名稱")
    equipment_name: str = Field("", description="設備名稱")
    station: str = Field("", description="站別")
    location: str = Field("", description="位置")
    cost_type: str = Field("非修繕", description="費用別")
    purchase_category_repair: str = Field("NA", description="請購類別(修繕)")
    purchase_category_non_repair: str = Field("", description="請購類別(非修繕)")
    work_reason: str = Field("其他", description="施作原因")

    # Narrative block 1
    situation_desc_1: str = Field("", description="現況說明:1")
    photo_location_1a: str = Field("", description="照片/施作位置 1a")
    photo_location_1b: str = Field("", description="照片/施作位置 1b")
    problem_risk_1: str = Field("", description="現況問題及風險 1")
    countermeasure_1: str = Field("", description="對策評估說明 1")
    execution_1: str = Field("", description="執行改善/維修對策 1")

    # Narrative block 2
    situation_desc_2: str = Field("", description="現況說明:2")
    photo_location_2a: str = Field("", description="照片/施作位置 2a")
    photo_location_2b: str = Field("", description="照片/施作位置 2b")
    problem_risk_2: str = Field("", description="現況問題及風險 2")
    improvement_desc: str = Field("", description="改善/維修說明")
    execution_2: str = Field("", description="執行改善/維修對策 2")

    # Supplementary
    supplementary_notes: str = Field("", description="補充說明")


class ComplianceResult(BaseModel):
    """LLM judgment on whether a spec item is relevant to the project."""

    row_index: int = Field(..., description="原始表格列索引")
    is_relevant: bool = Field(..., description="是否適用於本案")
    reason: str = Field("", description="判定理由")
