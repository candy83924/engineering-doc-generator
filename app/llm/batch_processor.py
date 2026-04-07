"""Batched LLM call orchestrator: one call per generator type."""

import logging

from app.llm.client import LLMClient
from app.llm.prompts import (
    build_acceptance_prompt,
    build_compliance_prompt,
    build_project_brief_prompt,
    build_quote_fill_prompt,
)
from app.models.quote_data import (
    AcceptanceContent,
    AcceptanceCriterion,
    ComplianceResult,
    QuoteData,
    QuoteFillContent,
)

logger = logging.getLogger(__name__)


class LLMBatchProcessor:
    """Orchestrates LLM calls for document generation."""

    def __init__(self, client: LLMClient):
        self.client = client

    async def generate_project_brief(self, quote_data: QuoteData) -> str:
        """Auto-generate a project brief from quote content."""
        system_prompt, user_prompt = build_project_brief_prompt(quote_data)
        try:
            return await self.client.generate_text(
                system_prompt=system_prompt,
                user_prompt=user_prompt,
                max_tokens=1024,
            )
        except Exception as e:
            logger.error("Failed to generate project brief: %s", str(e))
            return quote_data.get_project_description()

    async def generate_acceptance_content(
        self, quote_data: QuoteData, project_brief: str = ""
    ) -> AcceptanceContent:
        """Generate acceptance specification content via a single LLM call."""
        system_prompt, user_prompt = build_acceptance_prompt(quote_data, project_brief)

        try:
            result = await self.client.generate_json(
                system_prompt=system_prompt,
                user_prompt=user_prompt,
                max_tokens=8192,
            )

            criteria = []
            for c in result.get("acceptance_criteria", []):
                criteria.append(AcceptanceCriterion(
                    category=c.get("category", ""),
                    standard=c.get("standard", ""),
                ))

            return AcceptanceContent(
                equipment_location=result.get("equipment_location", ""),
                equipment_type=result.get("equipment_type", ""),
                spec_info=result.get("spec_info", ""),
                problem_description=result.get("problem_description", ""),
                repair_needs=result.get("repair_needs", ""),
                contractor_requirements=result.get("contractor_requirements", []),
                preparation_steps=result.get("preparation_steps", []),
                construction_content=result.get("construction_content", []),
                acceptance_procedures=result.get("acceptance_procedures", []),
                acceptance_criteria=criteria,
                insurance_text=result.get("insurance_text", ""),
                work_area_text=result.get("work_area_text", ""),
            )

        except Exception as e:
            logger.error("Failed to generate acceptance content: %s", str(e))
            return _fallback_acceptance_content(quote_data)

    async def generate_quote_fill_content(
        self, quote_data: QuoteData, project_brief: str = ""
    ) -> QuoteFillContent:
        """Generate 欣興 internal form content via a single LLM call."""
        system_prompt, user_prompt = build_quote_fill_prompt(quote_data, project_brief)

        try:
            result = await self.client.generate_json(
                system_prompt=system_prompt,
                user_prompt=user_prompt,
                max_tokens=4096,
            )

            return QuoteFillContent(
                purchase_name=result.get("purchase_name", ""),
                equipment_name=result.get("equipment_name", ""),
                station=result.get("station", ""),
                location=result.get("location", ""),
                cost_type=result.get("cost_type", "非修繕"),
                purchase_category_repair=result.get("purchase_category_repair", "NA"),
                purchase_category_non_repair=result.get("purchase_category_non_repair", ""),
                work_reason=result.get("work_reason", "其他"),
                situation_desc_1=result.get("situation_desc_1", ""),
                photo_location_1a=result.get("photo_location_1a", ""),
                photo_location_1b=result.get("photo_location_1b", ""),
                problem_risk_1=result.get("problem_risk_1", ""),
                countermeasure_1=result.get("countermeasure_1", ""),
                execution_1=result.get("execution_1", ""),
                situation_desc_2=result.get("situation_desc_2", ""),
                photo_location_2a=result.get("photo_location_2a", ""),
                photo_location_2b=result.get("photo_location_2b", ""),
                problem_risk_2=result.get("problem_risk_2", ""),
                improvement_desc=result.get("improvement_desc", ""),
                execution_2=result.get("execution_2", ""),
                supplementary_notes=result.get("supplementary_notes", ""),
            )

        except Exception as e:
            logger.error("Failed to generate quote fill content: %s", str(e))
            return _fallback_quote_fill_content(quote_data)

    async def generate_compliance_judgments(
        self,
        quote_data: QuoteData,
        spec_items: list[dict],
        project_brief: str = "",
    ) -> list[ComplianceResult]:
        """Judge which spec items are relevant via a single LLM call."""
        if not spec_items:
            return []

        system_prompt, user_prompt = build_compliance_prompt(
            quote_data, spec_items, project_brief
        )

        try:
            results = await self.client.generate_json(
                system_prompt=system_prompt,
                user_prompt=user_prompt,
                max_tokens=4096,
            )

            if not isinstance(results, list):
                results = results.get("results", results.get("items", []))

            return [
                ComplianceResult(
                    row_index=r.get("row_index", 0),
                    is_relevant=r.get("is_relevant", True),
                    reason=r.get("reason", ""),
                )
                for r in results
            ]

        except Exception as e:
            logger.error("Failed to generate compliance judgments: %s", str(e))
            return [
                ComplianceResult(
                    row_index=item["row_index"],
                    is_relevant=True,
                    reason="LLM 判定失敗，保守保留",
                )
                for item in spec_items
            ]


def _fallback_acceptance_content(quote_data: QuoteData) -> AcceptanceContent:
    """Generate minimal acceptance content when LLM fails."""
    items_desc = "、".join(
        item.description for item in quote_data.get_real_items()[:5]
    )
    meta = quote_data.metadata

    return AcceptanceContent(
        equipment_location=meta.project_location or meta.project_name,
        equipment_type=items_desc[:100] if items_desc else "依報價單內容",
        spec_info="依現場既設設備及本次報價內容辦理",
        problem_description=f"因{meta.project_name}相關設施或設備需進行修改、更新或改善，以符合使用需求。",
        repair_needs=f"依報價單內容執行：{items_desc}等工程項目。",
        contractor_requirements=[
            "具備相關工程施工經驗與技術能力，並應指派專責監工人員。",
            "進場人員須熟悉相關施工規範與安全要求。",
            "持有合法營業登記與相關證照。",
        ],
        preparation_steps=[
            "施工前應確認修改範圍、施工時段及防護需求。",
            "依報價內容逐項核對材料、設備與數量。",
            "完成現況丈量、放樣與拍照存查。",
            "完成施工區域之警示標示與落塵防護。",
            "提出施工工序與工期甘特圖。",
        ],
        construction_content=[
            f"依報價單執行{meta.project_name}相關施工作業。",
            "各工項完工後應依區域辦理清潔與現場復原。",
        ],
        acceptance_procedures=[
            "驗收應依各系統分項辦理，確認內容與報價單一致。",
            "應填寫驗收表單，檢附施工照片與測試紀錄。",
        ],
        acceptance_criteria=[
            AcceptanceCriterion(
                category="通用驗收",
                standard="1. 依報價內容確認施工完成度。\n2. 提供施工照片與測試紀錄。",
            ),
        ],
        insurance_text="所有進場人員須投保工程意外保險，並於開工前提供有效保險證明文件。",
        work_area_text=f"本次施工區域為{meta.project_location or meta.project_name}範圍。",
    )


def _fallback_quote_fill_content(quote_data: QuoteData) -> QuoteFillContent:
    """Generate minimal quote fill content when LLM fails."""
    meta = quote_data.metadata
    items_desc = "、".join(
        item.description for item in quote_data.get_real_items()[:5]
    )

    return QuoteFillContent(
        purchase_name=f"{meta.project_name}（{meta.document_number or ''}）",
        equipment_name=meta.project_name,
        station="",
        location=meta.project_location or "",
        situation_desc_1=f"本案為{meta.project_name}，\n內容含{items_desc}。",
        problem_risk_1="依報價單內容施作。",
        countermeasure_1="依報價分項執行各工程項目。",
        execution_1="依報價內容完成施工。",
        supplementary_notes=f"1. 報價單號：{meta.document_number or '未指定'}。\n2. 工程名稱：{meta.project_name}。\n3. 未稅總計：{quote_data.summary.subtotal:,.0f}，含稅合計：{quote_data.summary.grand_total:,.0f}。",
    )
