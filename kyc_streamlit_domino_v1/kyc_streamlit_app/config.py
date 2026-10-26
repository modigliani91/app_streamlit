"""Seul point de configuration de l'application; règles métier dans notebook_config."""
from dataclasses import dataclass, replace
from pathlib import Path
import os

APP_VERSION = "1.0.0"

@dataclass(frozen=True)
class AppConfig:
    app_root: Path = Path(__file__).resolve().parent
    input_dir: Path = Path("/mnt/kyc_app/input")
    extract_dir: Path = Path("/mnt/kyc_app/extracted")
    reference_csv: Path = Path("/mnt/recertification_lissage_updated.csv")
    output_dir: Path = Path("/mnt/output_qwen38")
    model_path: Path = Path(os.environ.get("KYC_MODEL_PATH", "/domino/edv/modelhub/ModelHub-model-huggingface-Qwen/Qwen3.8-27B-FP8/main"))
    zip_customer_root: str | None = os.environ.get("KYC_ZIP_CUSTOMER_ROOT") or None
    max_upload_mb: int = int(os.environ.get("KYC_MAX_UPLOAD_MB", "1024"))
    input_filename: str = "kyc_documents_test.zip"
    report_filename: str = "kyc_recertification_control.xlsx"

    @property
    def input_path(self): return self.input_dir / self.input_filename
    @property
    def report_path(self): return self.output_dir / "reports" / self.report_filename
    @property
    def metadata_dir(self): return self.output_dir / "run_metadata"
    @property
    def lock_path(self): return self.output_dir / ".kyc_processing.lock"
    @property
    def report_receipt(self): return self.metadata_dir / "current_report.json"

    def notebook_config(self, archive_sha256):
        from notebook_config import StructuredConfig
        return replace(StructuredConfig(), RUN_MODE="FULL_PIPELINE", INPUT_SOURCE="ZIP",
            ZIP_PATH=self.input_path, EXTRACT_DIR=self.extract_dir,
            OUTPUT_DIR=self.output_dir / "datasets" / archive_sha256,
            REFERENCE_CSV=self.reference_csv, MODEL_PATH=self.model_path,
            ZIP_CUSTOMER_ROOT=self.zip_customer_root,
            MAX_QWEN_CALLS=None, PRINT_RAW_OCR=False)

SETTINGS = AppConfig()
