"""Base generator abstract class."""

from abc import ABC, abstractmethod
from pathlib import Path

from app.llm.batch_processor import LLMBatchProcessor
from app.models.quote_data import QuoteData


class BaseGenerator(ABC):
    """Abstract base for all document generators."""

    doc_type: str = ""
    output_extension: str = ""

    def __init__(
        self,
        template_path: Path | None = None,
        llm_processor: LLMBatchProcessor | None = None,
    ):
        self.template_path = template_path
        self.llm = llm_processor

    @abstractmethod
    async def generate(self, quote_data: QuoteData, **kwargs) -> bytes:
        """Generate the document and return as bytes."""
        ...

    def validate_input(self, quote_data: QuoteData) -> list[str]:
        """Return a list of warnings for missing or questionable data."""
        warnings = []
        if not quote_data.metadata.project_name or quote_data.metadata.project_name == "未命名工程":
            warnings.append("未偵測到工程名稱")
        if not quote_data.get_real_items():
            warnings.append("未偵測到任何報價項目")
        return warnings

    def get_output_filename(self, quote_data: QuoteData) -> str:
        """Generate the output filename."""
        from app.utils.file_utils import get_output_filename
        return get_output_filename(
            project_name=quote_data.metadata.project_name,
            doc_type=self.doc_type,
            extension=self.output_extension,
            document_number=quote_data.metadata.document_number,
        )
