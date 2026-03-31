import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# API Keys
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")

# Models
EXTRACTION_MODEL = "claude-sonnet-4-6"
REPORT_MODEL = "claude-sonnet-4-6"

# Thresholds
CONFIDENCE_THRESHOLD = 0.9

# Paths
DATA_DIR = Path(__file__).parent / "data"
RAW_DIR = DATA_DIR / "raw"
EXTRACTED_DIR = DATA_DIR / "extracted"
POPPLER_PATH = r"C:\Users\alexi\AppData\Local\Programs\Poppler\poppler-24.08.0\Library\bin"

SURVEY_TYPES = {
    "IL": {
        "extracted_dir": "data/extracted",
        "output_dir": "output/IL",
        "prompt_file": "prompts/extraction_prompt_IL.txt",
    },
    "412YZ": {
        "extracted_dir": "data/extracted",
        "output_dir": "output/412YZ",
        "prompt_file": "prompts/extraction_prompt_412YZ.txt",
    },
}

