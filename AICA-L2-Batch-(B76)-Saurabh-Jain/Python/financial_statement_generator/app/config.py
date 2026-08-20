"""Application configuration for paths and runtime settings."""

from pathlib import Path

from pydantic_settings import BaseSettings, SettingsConfigDict

BASE_DIR = Path(__file__).resolve().parent.parent


class Settings(BaseSettings):
    """Runtime settings loaded from environment variables and `.env`."""

    model_config = SettingsConfigDict(
        env_file=BASE_DIR / ".env",
        env_file_encoding="utf-8",
        extra="ignore",
        case_sensitive=False,
    )

    app_name: str = "Financial Statement Generator"
    app_version: str = "0.1.0"
    debug: bool = False

    input_dir: Path = BASE_DIR / "input"
    reference_dir: Path = BASE_DIR / "reference"
    template_dir: Path = BASE_DIR / "templates"
    output_dir: Path = BASE_DIR / "output"

    template_filename: str = "Financial Statements_Sample.xlsx"
    reference_filename: str = "ICAI_GN_Div_I_Sch_III.pdf"
    max_upload_bytes: int = 25 * 1024 * 1024
    ocr_enabled: bool = True
    ocr_language: str = "eng"
    ocr_dpi: int = 200
    mapping_config_path: Path = BASE_DIR / "app" / "mapping" / "excel_field_map.json"
    review_confidence_threshold: float = 0.85

    @property
    def template_path(self) -> Path:
        return self.template_dir / self.template_filename

    @property
    def reference_path(self) -> Path:
        return self.reference_dir / self.reference_filename

    @property
    def upload_dir(self) -> Path:
        return self.input_dir / "uploads"

    @property
    def jobs_dir(self) -> Path:
        return self.output_dir / "jobs"

    def ensure_directories(self) -> None:
        """Create configured directories if they do not already exist."""
        for directory in (
            self.input_dir,
            self.upload_dir,
            self.reference_dir,
            self.template_dir,
            self.output_dir,
            self.jobs_dir,
        ):
            directory.mkdir(parents=True, exist_ok=True)


settings = Settings()
settings.ensure_directories()
