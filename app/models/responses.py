from pydantic import BaseModel, Field


class GenerationResult(BaseModel):
    """Metadata about a single generated document."""

    doc_type: str
    filename: str
    size_bytes: int
    warnings: list[str] = Field(default_factory=list)


class GenerationResponse(BaseModel):
    """Response metadata for the generation endpoint."""

    success: bool = True
    project_name: str = ""
    documents: list[GenerationResult] = Field(default_factory=list)
    parse_confidence: float = 1.0
    parse_warnings: list[str] = Field(default_factory=list)
    errors: list[str] = Field(default_factory=list)
