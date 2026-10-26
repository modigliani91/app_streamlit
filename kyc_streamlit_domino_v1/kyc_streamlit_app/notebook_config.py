# Configuration originale, classes copiées sans modification.
from __future__ import annotations
from pathlib import Path
from dataclasses import dataclass, field, asdict
import sys
import os
os.environ["HF_HUB_OFFLINE"]="1"
os.environ["TRANSFORMERS_OFFLINE"]="1"
from importlib import metadata

@dataclass(frozen=True)
class Config:
    MODEL_PATH: Path = Path('/domino/edv/modelhub/ModelHub-model-huggingface-Qwen/Qwen3.8-27B-FP8/main')
    ZIP_PATH: Path = Path.cwd() / 'kyc_documents.zip'
    EXTRACT_DIR: Path = Path.cwd() / 'kyc_extracted'
    OUTPUT_DIR: Path = Path.cwd() / 'outputs'
    PIPELINE_VERSION: str = 'KYC_QWEN38_PRODUCTION_1'
    MODEL_REVISION_NOTE: str = 'local-checkpoint'
    MAX_QWEN_CALLS: int = 20
    DUPLICATE_POLICY: str = 'all'
    ZIP_CUSTOMER_ROOT: str | None = None
    MAX_ZIP_MEMBERS: int = 200000
    MAX_EXTRACT_BYTES: int = 50 * 1024 ** 3
    MAX_PDF_BYTES: int = 1024 ** 3
    MAX_COMPRESSION_RATIO: float = 2000.0
    TARGET_DOCUMENTS: tuple[str, ...] = ('JUSTIFICATIF IDENTITE.PDF', 'JUSTIFICATIF DOMICILE.PDF', 'CONVENTION COMPTE.PDF', 'FATCA.PDF', 'CARTON SIGNATUTE.PDF')
    DOCUMENT_ALIASES: dict[str, str] = field(default_factory=lambda: {'CARTON SIGNATURE.PDF': 'CARTON SIGNATUTE.PDF'})
    RECOGNIZE_DUPLICATE_SUFFIX: bool = True
    RENDER_ZOOM_STANDARD: float = 3.0
    IMAGE_MAX_SIZE_STANDARD: int = 1400
    RENDER_ZOOM_HD: float = 4.5
    IMAGE_MAX_SIZE_HD: int = 2200
    MAX_RENDER_PIXELS: int = 35000000
    MIN_PIXELS: int = 4 * 32 * 32
    MAX_PIXELS: int = 2600 * 32 * 32
    MAX_NEW_TOKENS_RAW_OCR: int = 2048
    MAX_GENERATION_SECONDS: float = 120.0
    SLOW_GENERATION_SECONDS: float = 60.0
    STOP_ON_BAD_BENCHMARK: bool = True
    STOP_AFTER_CONSECUTIVE_BAD_PAGES: int = 2
    ENABLE_PREPROCESSING: bool = True
    ENABLE_AUTO_CROP: bool = False
    ENABLE_DESKEW: bool = False
    ENABLE_CONTRAST: bool = False
    ENABLE_SHARPEN: bool = False
    ENABLE_DENOISE: bool = False
    ROTATION_OVERRIDES: dict[str, int] = field(default_factory=dict)
    CROP_WHITE_LEVEL: int = 250
    CROP_PADDING_FRACTION: float = 0.025
    CROP_MIN_AREA_REDUCTION: float = 0.2
    DESKEW_MAX_DEGREES: float = 2.0
    DESKEW_MIN_SCORE_GAIN: float = 0.15
    CONTRAST_FACTOR: float = 1.06
    ENABLE_HD_FALLBACK: bool = False
    SKIP_BLANK_PAGES: bool = False
    BLANK_WHITE_RATIO: float = 0.9995
    BLANK_MAX_DARK_PIXELS: int = 12
    BLANK_MAX_STD: float = 1.5
    DENSE_PAGE_INK_RATIO: float = 0.015
    SHORT_OCR_CHAR_LIMIT: int = 25
    SAVE_PAGE_IMAGES: bool = True
    PRINT_RAW_OCR: bool = True

@dataclass(frozen=True)
class StructuredConfig(Config):
    PIPELINE_VERSION: str = "KYC_QWEN38_PRODUCTION_1"
    SCHEMA_VERSION: str = "KYC_Q38_EVIDENCE_1"
    CUSTOMER_JSON_SCHEMA_VERSION: str = "Q38-PROD-1"
    DOCUMENT_JSON_SCHEMA_VERSION: str = "Q38-PROD-1"
    MATCHING_VERSION: str = "KYC_MATCHING_HANDWRITTEN_NAME_SIM_75_V1"
    HANDWRITTEN_NAME_SIMILARITY_THRESHOLD: float = 75.0
    OUTPUT_DIR: Path = Path.cwd() / "outputs_qwen38"
    REFERENCE_CSV: Path = Path.cwd() / "recertication_lissage_updated.csv"
    RUN_MODE: str = "MATCH_ONLY"  # Recalcul des JSON existants sans Qwen. FULL_PIPELINE reste disponible.
    MAX_NEW_TOKENS_STRUCTURED: int = 3072
    ENABLE_STRUCTURED_HD_FALLBACK: bool = True  # Au plus un repli image, sans répéter l'OCR.
    RAW_CACHE_REVISION: str = "validated-raw-v1"
    EXTRACTION_REVISION: str = "v3-preserved-with-staged-extraction"
    CONFIDENCE_VERSION: str = "OPERATIONAL_SCORE_3"
    CONFIDENCE_HIGH_THRESHOLD: float = 0.85
    CONFIDENCE_MEDIUM_THRESHOLD: float = 0.65
    CONFIDENCE_REVIEW_THRESHOLD: float = 0.85
    CONFIDENCE_CONFLICT_CAP: float = 0.35
    RECOGNITION_CLEAR_SCORE: float = 0.90
    RECOGNITION_PARTIAL_SCORE: float = 0.65
    RECOGNITION_UNCERTAIN_SCORE: float = 0.40
    RECOGNITION_TEXT_ONLY_SCORE: float = 0.85
    REVIEW_UNKNOWN_WRITING_TYPE: bool = True
    REVIEW_CROSS_DOCUMENT_CONFLICT: bool = True
    REVIEW_HIGH_CONFIDENCE_REFERENCE_MISMATCH: bool = False
    FAIL_ON_REFERENCE_SCHEMA_ERROR: bool = True
    REFERENCE_ENCODING: str = "utf-8-sig"  # ou cp1252 / latin1 / utf-8-sig explicitement.
    REFERENCE_SEPARATOR: str | None = None  # None : ; , tabulation |, contrôlés contre le schéma.
    REPAIR_LEGACY_DECIMAL_IDS: bool = True  # Réparation textuelle de ^[0-9]+\.0+$, journalisée.
    DATE_ORDER: str = "DMY"
    ALLOW_IMAGE_ONLY_EVIDENCE: bool = True
    REQUIRE_COMPLETE_OCR_FOR_EXTRACTION: bool = True
    HEADER_SCAN_MAX_LINES: int = 22
    DUPLICATE_POLICY: str = "family_and_size"
    ENABLE_HD_FALLBACK: bool = False  # OCR : préserver une passe, sans escalade implicite.
    PRINT_RAW_OCR: bool = False
    SAVE_PAGE_IMAGES: bool = True

    RAW_OCR_VERSION: str = "Q38_RAW_PRODUCTION_1"
    STRUCTURED_EXTRACTION_VERSION: str = "Q38_STRUCTURE_PRODUCTION_1"
    MAX_QWEN_CALLS: int | None = None  # None : pas de limite cachée sur la population complète.
    MAX_NEW_TOKENS_STRUCTURED_TEXT: int = 1536
    MAX_STRUCTURED_INPUT_TOKENS: int = 32768
    ENABLE_QWEN_TEXT: bool = True
    ENABLE_MULTIMODAL_FALLBACK: bool = True
    INPUT_SOURCE: str = "AUTO"  # AUTO | ZIP | DIRECTORY
    CUSTOMER_ROOT: Path | None = None  # Racine contenant directement les dossiers clients si déjà dézippés.
    DIAGNOSTIC_CUSTOMER_ID: str | None = None
    DIAGNOSTIC_CUSTOMER_COUNT: int = 1  # Mettre 5 pour la validation structurée limitée.
    RAW_PROGRESS_EVERY_CUSTOMERS: int = 10
    RAW_EXPORT_EVERY_DOCUMENTS: int = 50
    MAX_CONSECUTIVE_RAW_FAILURES: int = 3
    REUSE_RAW_CACHE: bool = True
    REUSE_STRUCTURED_CACHE: bool = True
    RETRY_PARTIAL_STRUCTURED: bool = False  # Mettre True pour retenter les champs manquants, jamais l'OCR.
    FORCE_RAW_RERUN: bool = False  # Explicite ; les caches Qwen3.6 ne sont jamais importés.
    ALLOW_RECORDED_RAW_ERRORS: bool = True  # Pages réussies exploitables ; erreurs RAW restent non structurées.
    HASH_MODEL_WEIGHT_FILES: bool = False  # True : SHA-256 de chaque shard, coût I/O important.
    COMBINED_NAME_ORDER: str = "SURNAME_THEN_GIVEN"  # Uniquement 2 mots et libellé explicite Nom + Prénom.
    VERBOSE_DIAGNOSTICS: bool = True
    RUN_PREFLIGHT_BENCHMARK: bool = True
    MODEL_PROBABILITY_THRESHOLD: float = 0.85
    COLLECT_STRUCTURED_TOKEN_SCORES: bool = True
    REQUIRE_MODEL_PROBABILITY: bool = False  # True : une vérification textuelle groupée, jamais un appel par champ.
    ENABLE_WRITING_METADATA: bool = True  # Un appel visuel ciblé pour les champs déjà extraits, si nécessaire.
    STRICT_INSTANCE_CONFLICT: bool = False
    ALLOW_RELIABLE_ANCHORED_RULE_MATCH: bool = False  # Par défaut une probabilité absente exige revue.


    @property
    def DIAGNOSTIC_MODE(self) -> bool:
        return self.RUN_MODE == "DIAGNOSTIC"

