from pathlib import Path


# ── Project paths ────────────────────────────────────────
PROJECT_ROOT = Path(__file__).resolve().parent.parent
OUTPUT_DIR = PROJECT_ROOT / "output"
LOGS_DIR = PROJECT_ROOT / "logs"
SAMPLE_DATA_DIR = PROJECT_ROOT / "sample_data"

# ── File handling ────────────────────────────────────────
ALLOWED_EXTENSIONS = {".xlsx", ".csv"}
MAX_FILE_SIZE_MB = 50

# ── Logging ──────────────────────────────────────────────
LOG_FILE = LOGS_DIR / "app.log"
LOG_FORMAT = "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s"
LOG_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"

# ── Data processing defaults ─────────────────────────────
DEFAULT_ENCODING = "utf-8"
FALLBACK_ENCODING = "latin-1"
SHEET_READ_ENGINE = "openpyxl"

# ── Schema compatibility ─────────────────────────────────
SCHEMA_THRESHOLD = 0.3
