"""Temporary file and path management utilities."""

import tempfile
from pathlib import Path


def create_temp_dir(prefix: str = "engdoc_") -> Path:
    """Create a temporary directory and return its path."""
    return Path(tempfile.mkdtemp(prefix=prefix))


def ensure_dir(path: Path) -> Path:
    """Ensure a directory exists, creating it if necessary."""
    path.mkdir(parents=True, exist_ok=True)
    return path


def get_output_filename(
    project_name: str,
    doc_type: str,
    extension: str,
    document_number: str | None = None,
) -> str:
    """Generate an output filename following the naming convention."""
    from app.utils.chinese_utils import safe_filename

    doc_type_names = {
        "acceptance_spec": "驗收規範",
        "common_bid": "空白標單",
        "quote_fill": "報價內容",
        "compliance_filter": "規格篩選",
    }
    type_label = doc_type_names.get(doc_type, doc_type)
    base = safe_filename(project_name)

    if document_number:
        return f"{type_label}_{base}_{document_number}.{extension}"
    return f"{type_label}_{base}.{extension}"
