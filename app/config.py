from pathlib import Path
from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    # LLM Provider: "anthropic" or "google"
    llm_provider: str = "anthropic"

    # Anthropic settings
    anthropic_api_key: str = ""
    llm_model: str = "claude-sonnet-4-20250514"
    llm_model_complex: str = "claude-opus-4-20250514"

    # Google Gemini settings
    google_api_key: str = ""
    google_model: str = "gemini-2.0-flash"

    # API access password (protect against unauthorized use)
    api_password: str = ""

    # General settings
    templates_dir: Path = Path("./templates")
    output_dir: Path = Path("./output")
    log_level: str = "INFO"
    max_retries: int = 3
    parse_confidence_threshold: float = 0.6

    @property
    def default_templates_dir(self) -> Path:
        return Path(__file__).resolve().parent.parent / "templates" / "default"

    model_config = {
        "env_file": str(Path(__file__).resolve().parent.parent / ".env"),
        "env_file_encoding": "utf-8",
        "extra": "ignore",
    }


settings = Settings()
