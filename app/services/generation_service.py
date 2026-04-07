"""Orchestration service: ties parsing, LLM, and generation together."""

import asyncio
import logging
from datetime import datetime
from pathlib import Path

from app.config import settings
from app.generators.acceptance_spec import AcceptanceSpecGenerator
from app.generators.common_bid import CommonBidGenerator
from app.generators.compliance_filter import ComplianceFilterGenerator
from app.generators.quote_fill import QuoteFillGenerator
from app.llm.batch_processor import LLMBatchProcessor
from app.llm.client import LLMClient
from app.models.quote_data import QuoteData
from app.models.responses import GenerationResponse, GenerationResult
from app.parsing.pdf_parser import PDFParser
from app.services.packaging_service import PackagingService

logger = logging.getLogger(__name__)


class GenerationService:
    """Orchestrates the full document generation pipeline."""

    def __init__(self):
        self.parser = PDFParser()
        self.llm_client = LLMClient()
        self.llm_processor = LLMBatchProcessor(self.llm_client)
        self.packager = PackagingService()

    async def generate(
        self,
        quote_pdf_bytes: bytes,
        requested_docs: list[str],
        templates: dict[str, Path] | None = None,
        project_name_override: str | None = None,
        project_brief: str | None = None,
    ) -> tuple[bytes, GenerationResponse]:
        """
        Run the full generation pipeline.

        Returns (zip_bytes, response_metadata).
        """
        templates = templates or {}
        errors = []

        # Step 1: Parse PDF
        logger.info("Parsing quotation PDF...")
        quote_data = await self.parser.parse(quote_pdf_bytes)

        if project_name_override:
            quote_data.metadata.project_name = project_name_override

        logger.info(
            "Parsed: project=%s, items=%d, confidence=%.2f",
            quote_data.metadata.project_name,
            len(quote_data.all_items),
            quote_data.parse_confidence,
        )

        # Step 1.5: Auto-generate project brief if not provided
        if project_brief:
            quote_data.project_brief = project_brief
        elif not quote_data.project_brief:
            try:
                logger.info("Auto-generating project brief...")
                quote_data.project_brief = await self.llm_processor.generate_project_brief(quote_data)
                logger.info("Project brief: %s", quote_data.project_brief[:100])
            except Exception as e:
                logger.warning("Failed to auto-generate project brief: %s", e)
                quote_data.project_brief = quote_data.get_project_description()

        # Step 2: Pre-generate LLM content in parallel
        llm_results = {}
        llm_tasks = {}

        if "acceptance_spec" in requested_docs:
            llm_tasks["acceptance_content"] = (
                self.llm_processor.generate_acceptance_content(
                    quote_data, quote_data.project_brief or ""
                )
            )

        if "quote_fill" in requested_docs:
            llm_tasks["quote_fill_content"] = (
                self.llm_processor.generate_quote_fill_content(
                    quote_data, quote_data.project_brief or ""
                )
            )

        if llm_tasks:
            logger.info("Running LLM calls for: %s", list(llm_tasks.keys()))
            results = await asyncio.gather(
                *llm_tasks.values(), return_exceptions=True
            )
            for key, result in zip(llm_tasks.keys(), results):
                if isinstance(result, Exception):
                    logger.error("LLM call %s failed: %s", key, result)
                    errors.append(f"LLM {key} 呼叫失敗：{str(result)}")
                else:
                    llm_results[key] = result

        # Step 3: Initialize generators
        generators = {}
        if "acceptance_spec" in requested_docs:
            generators["acceptance_spec"] = AcceptanceSpecGenerator(
                template_path=templates.get("acceptance_spec"),
                llm_processor=self.llm_processor,
            )
        if "common_bid" in requested_docs:
            generators["common_bid"] = CommonBidGenerator(
                template_path=templates.get("common_bid"),
            )
        if "quote_fill" in requested_docs:
            generators["quote_fill"] = QuoteFillGenerator(
                template_path=templates.get("quote_fill"),
                llm_processor=self.llm_processor,
            )
        if "compliance_filter" in requested_docs:
            generators["compliance_filter"] = ComplianceFilterGenerator(
                template_path=templates.get("compliance_filter"),
                llm_processor=self.llm_processor,
            )

        # Step 4: Run generators in parallel
        logger.info("Running %d generators...", len(generators))
        gen_tasks = {}
        for doc_type, generator in generators.items():
            kwargs = {}
            if doc_type == "acceptance_spec" and "acceptance_content" in llm_results:
                kwargs["acceptance_content"] = llm_results["acceptance_content"]
            if doc_type == "quote_fill" and "quote_fill_content" in llm_results:
                kwargs["quote_fill_content"] = llm_results["quote_fill_content"]
            gen_tasks[doc_type] = generator.generate(quote_data, **kwargs)

        gen_results = await asyncio.gather(
            *gen_tasks.values(), return_exceptions=True
        )

        # Step 5: Collect results and package
        files: dict[str, bytes] = {}
        doc_results: list[GenerationResult] = []

        for doc_type, result in zip(gen_tasks.keys(), gen_results):
            generator = generators[doc_type]
            if isinstance(result, Exception):
                logger.error("Generator %s failed: %s", doc_type, result)
                errors.append(f"{doc_type} 生成失敗：{str(result)}")
                continue

            filename = generator.get_output_filename(quote_data)
            files[filename] = result
            doc_results.append(GenerationResult(
                doc_type=doc_type,
                filename=filename,
                size_bytes=len(result),
                warnings=generator.validate_input(quote_data),
            ))

        # Step 6: Create metadata
        metadata_text = _build_metadata_text(quote_data, doc_results, errors)

        # Step 7: Package
        if files:
            zip_bytes = self.packager.package_with_metadata(files, metadata_text)
        else:
            zip_bytes = b""

        response = GenerationResponse(
            success=len(errors) == 0,
            project_name=quote_data.metadata.project_name,
            documents=doc_results,
            parse_confidence=quote_data.parse_confidence,
            parse_warnings=quote_data.parse_warnings,
            errors=errors,
        )

        logger.info(
            "Generation complete: %d docs, %d errors", len(doc_results), len(errors)
        )
        return zip_bytes, response

    async def parse_only(self, quote_pdf_bytes: bytes) -> QuoteData:
        """Parse a PDF without generating documents (for debugging)."""
        return await self.parser.parse(quote_pdf_bytes)

    async def close(self):
        """Clean up resources."""
        await self.llm_client.close()


def _build_metadata_text(
    quote_data: QuoteData,
    results: list[GenerationResult],
    errors: list[str],
) -> str:
    """Build a metadata summary text file for the ZIP."""
    lines = [
        "工程文件生成資訊",
        "=" * 40,
        f"生成時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"工程名稱：{quote_data.metadata.project_name}",
        f"文件編號：{quote_data.metadata.document_number or '未知'}",
        f"廠商名稱：{quote_data.metadata.vendor_name or '未知'}",
        f"解析信心度：{quote_data.parse_confidence:.0%}",
        f"案子概略說明：{quote_data.project_brief or '未提供'}",
        "",
        "生成文件：",
    ]

    for r in results:
        size_kb = r.size_bytes / 1024
        lines.append(f"  - {r.filename} ({size_kb:.1f} KB)")
        if r.warnings:
            for w in r.warnings:
                lines.append(f"    ⚠ {w}")

    if errors:
        lines.append("")
        lines.append("錯誤：")
        for e in errors:
            lines.append(f"  ✗ {e}")

    if quote_data.parse_warnings:
        lines.append("")
        lines.append("解析警告：")
        for w in quote_data.parse_warnings:
            lines.append(f"  ⚠ {w}")

    return "\n".join(lines)
