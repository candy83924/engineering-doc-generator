"""FastAPI application entry point."""

import logging
import tempfile
import traceback
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, Header, Query, Request, UploadFile
from fastapi.responses import JSONResponse, Response
from fastapi.staticfiles import StaticFiles

from app.config import settings
from app.models.requests import VALID_DOCUMENT_TYPES
from app.models.responses import GenerationResponse
from app.services.generation_service import GenerationService

logging.basicConfig(
    level=getattr(logging, settings.log_level),
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger(__name__)

generation_service: GenerationService | None = None

# Default template file mapping
DEFAULT_TEMPLATES = {
    "acceptance_spec": "驗收規範範本.docx",
    "common_bid": "共同標單範本.xlsx",
    "quote_fill": "填寫報價單內容範本.xlsx",
    "compliance_filter": "工程規範符合性篩選表範本.xlsx",
}


def _get_default_templates(requested_docs: list[str]) -> dict[str, Path]:
    """Load default templates for requested doc types."""
    templates = {}
    default_dir = settings.default_templates_dir
    for doc_type in requested_docs:
        filename = DEFAULT_TEMPLATES.get(doc_type)
        if filename:
            path = default_dir / filename
            if path.exists():
                templates[doc_type] = path
                logger.info("Using default template for %s: %s", doc_type, filename)
    return templates


def _verify_password(password: str | None) -> bool:
    """Check if API password is correct. If no password is configured, allow all."""
    if not settings.api_password:
        return True
    return password == settings.api_password


@asynccontextmanager
async def lifespan(app: FastAPI):
    global generation_service
    generation_service = GenerationService()
    logger.info("Engineering Document Generator started")
    # Log default templates status
    default_dir = settings.default_templates_dir
    for doc_type, filename in DEFAULT_TEMPLATES.items():
        path = default_dir / filename
        status = "OK" if path.exists() else "MISSING"
        logger.info("Default template [%s] %s: %s", doc_type, filename, status)
    yield
    if generation_service:
        await generation_service.close()
    logger.info("Engineering Document Generator stopped")


app = FastAPI(
    title="工程文件生成系統",
    description="根據報價單 PDF 與範本自動生成驗收規範、空白標單、報價內容、規格篩選表",
    version="1.0.0",
    lifespan=lifespan,
)

# Serve static frontend
_static_dir = Path(__file__).resolve().parent / "static"
app.mount("/static", StaticFiles(directory=str(_static_dir), html=True), name="static")


@app.get("/", include_in_schema=False)
async def root():
    """Serve the web UI."""
    from fastapi.responses import FileResponse
    return FileResponse(str(_static_dir / "index.html"))


@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    tb = traceback.format_exc()
    logger.error("Unhandled exception:\n%s", tb)
    return JSONResponse(
        status_code=500,
        content={"error": str(exc), "traceback": tb},
    )


@app.post("/api/verify-password")
async def verify_password(x_api_password: Optional[str] = Header(None)):
    """Check if password is correct (or not required)."""
    if not settings.api_password:
        return {"required": False, "valid": True}
    if _verify_password(x_api_password):
        return {"required": True, "valid": True}
    return JSONResponse(status_code=401, content={"required": True, "valid": False})


@app.post(
    "/api/generate",
    response_class=Response,
    summary="生成工程文件",
    description="上傳報價單 PDF 與範本，生成指定的工程文件。回傳 ZIP 壓縮檔。",
)
async def generate_documents(
    quote_pdf: UploadFile = File(..., description="報價單 PDF 檔案"),
    acceptance_template: Optional[UploadFile] = File(
        None, description="驗收規範 Word 範本 (.docx)"
    ),
    bid_template: Optional[UploadFile] = File(
        None, description="空白標單 Excel 範本 (.xlsx)"
    ),
    quote_template: Optional[UploadFile] = File(
        None, description="報價內容 Excel 範本 (.xlsx)"
    ),
    compliance_template: Optional[UploadFile] = File(
        None, description="規格篩選 Excel 範本 (.xlsx)"
    ),
    documents: list[str] = Query(
        default=["acceptance_spec", "common_bid", "quote_fill", "compliance_filter"],
        description="要生成的文件類型",
    ),
    project_name: Optional[str] = Query(
        None, description="手動指定工程名稱（覆蓋 PDF 解析結果）"
    ),
    project_brief: Optional[str] = Query(
        None, description="案子概略說明（不填則由 AI 自動根據報價內容生成）"
    ),
    x_api_password: Optional[str] = Header(None),
):
    """Generate engineering documents from a quotation PDF."""
    # Auth check
    if not _verify_password(x_api_password):
        return JSONResponse(status_code=401, content={"error": "密碼錯誤"})

    # Validate document types
    invalid = set(documents) - VALID_DOCUMENT_TYPES
    if invalid:
        return Response(
            content=f"無效的文件類型：{invalid}。有效類型：{VALID_DOCUMENT_TYPES}",
            status_code=400,
        )

    # Read PDF
    pdf_bytes = await quote_pdf.read()

    # Start with default templates, then override with user uploads
    templates = _get_default_templates(documents)

    template_files = {
        "acceptance_spec": acceptance_template,
        "common_bid": bid_template,
        "quote_fill": quote_template,
        "compliance_filter": compliance_template,
    }

    temp_paths = []
    for doc_type, upload in template_files.items():
        if upload and doc_type in documents:
            content = await upload.read()
            if len(content) > 0:  # Only override if user actually uploaded a file
                suffix = Path(upload.filename).suffix if upload.filename else ".tmp"
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
                tmp.write(content)
                tmp.close()
                path = Path(tmp.name)
                templates[doc_type] = path  # Override default
                temp_paths.append(path)
                logger.info("User template override for %s: %s", doc_type, upload.filename)

    try:
        zip_bytes, response_meta = await generation_service.generate(
            quote_pdf_bytes=pdf_bytes,
            requested_docs=documents,
            templates=templates,
            project_name_override=project_name,
            project_brief=project_brief,
        )

        if not zip_bytes:
            return Response(
                content=response_meta.model_dump_json(indent=2),
                status_code=500,
                media_type="application/json",
            )

        # Build filename for the ZIP (RFC 5987 for non-ASCII)
        from urllib.parse import quote
        safe_name = response_meta.project_name.replace("/", "_").replace("\\", "_")
        zip_filename = f"output_{safe_name}.zip"
        encoded_filename = quote(zip_filename)

        return Response(
            content=zip_bytes,
            media_type="application/zip",
            headers={
                "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}",
                "X-Generation-Success": str(response_meta.success),
                "X-Parse-Confidence": str(response_meta.parse_confidence),
                "X-Documents-Count": str(len(response_meta.documents)),
            },
        )

    finally:
        for path in temp_paths:
            try:
                path.unlink(missing_ok=True)
            except OSError:
                pass


@app.post(
    "/api/parse",
    summary="解析報價單 PDF",
    description="僅解析報價單 PDF 並回傳結構化資料，不生成文件。用於偵錯與預覽。",
)
async def parse_quote(
    quote_pdf: UploadFile = File(..., description="報價單 PDF 檔案"),
    x_api_password: Optional[str] = Header(None),
):
    """Parse a quotation PDF and return structured data."""
    if not _verify_password(x_api_password):
        return JSONResponse(status_code=401, content={"error": "密碼錯誤"})
    pdf_bytes = await quote_pdf.read()
    quote_data = await generation_service.parse_only(pdf_bytes)
    return quote_data.model_dump()


@app.post(
    "/api/generate/single/{doc_type}",
    response_class=Response,
    summary="生成單一文件",
    description="生成指定類型的單一工程文件。",
)
async def generate_single(
    doc_type: str,
    quote_pdf: UploadFile = File(..., description="報價單 PDF 檔案"),
    template: Optional[UploadFile] = File(None, description="範本檔案"),
    project_name: Optional[str] = Query(None, description="手動指定工程名稱"),
    project_brief: Optional[str] = Query(None, description="案子概略說明"),
    x_api_password: Optional[str] = Header(None),
):
    """Generate a single document type."""
    if not _verify_password(x_api_password):
        return JSONResponse(status_code=401, content={"error": "密碼錯誤"})

    if doc_type not in VALID_DOCUMENT_TYPES:
        return Response(
            content=f"無效的文件類型：{doc_type}。有效類型：{VALID_DOCUMENT_TYPES}",
            status_code=400,
        )

    pdf_bytes = await quote_pdf.read()

    # Start with default template
    templates = _get_default_templates([doc_type])
    temp_path = None

    if template:
        content = await template.read()
        if len(content) > 0:
            suffix = Path(template.filename).suffix if template.filename else ".tmp"
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
            tmp.write(content)
            tmp.close()
            temp_path = Path(tmp.name)
            templates[doc_type] = temp_path

    try:
        zip_bytes, response_meta = await generation_service.generate(
            quote_pdf_bytes=pdf_bytes,
            requested_docs=[doc_type],
            templates=templates,
            project_name_override=project_name,
            project_brief=project_brief,
        )

        if not response_meta.documents:
            return Response(
                content=response_meta.model_dump_json(indent=2),
                status_code=500,
                media_type="application/json",
            )

        ext_map = {"acceptance_spec": "docx", "common_bid": "xlsx",
                    "quote_fill": "xlsx", "compliance_filter": "xlsx"}
        media_map = {
            "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }
        ext = ext_map[doc_type]

        import zipfile
        import io
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            for name in zf.namelist():
                if name.endswith(f".{ext}"):
                    file_bytes = zf.read(name)
                    return Response(
                        content=file_bytes,
                        media_type=media_map[ext],
                        headers={
                            "Content-Disposition": f'attachment; filename="{name}"',
                        },
                    )

        return Response(content="文件生成失敗", status_code=500)

    finally:
        if temp_path:
            try:
                temp_path.unlink(missing_ok=True)
            except OSError:
                pass


@app.get("/api/health", summary="健康檢查")
async def health_check():
    return {
        "status": "ok",
        "service": "engineering-doc-generator",
        "version": "1.0.0",
    }


@app.get("/api/doc-types", summary="取得支援的文件類型")
async def get_doc_types():
    return {
        "document_types": {
            "acceptance_spec": {
                "name": "驗收規範",
                "output_format": "docx",
                "requires_llm": True,
                "template_type": "Word (.docx)",
                "description": "依報價單內容生成詳細的工程驗收規範文件",
            },
            "common_bid": {
                "name": "空白標單",
                "output_format": "xlsx",
                "requires_llm": False,
                "template_type": "Excel (.xlsx)",
                "description": "將報價項目轉為可供廠商填寫的空白標單",
            },
            "quote_fill": {
                "name": "填寫報價單內容",
                "output_format": "xlsx",
                "requires_llm": True,
                "template_type": "Excel (.xlsx)",
                "description": "生成欣興內部保養/維修/改善/工程說明表單，含 AI 生成敘述",
            },
            "compliance_filter": {
                "name": "規格篩選表",
                "output_format": "xlsx",
                "requires_llm": True,
                "template_type": "Excel (.xlsx)",
                "description": "篩選與本案工程相關的規範項目",
            },
        }
    }
