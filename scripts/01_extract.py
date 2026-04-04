"""
01_extract.py
For each PDF in data/raw/, convert pages to images, send to Claude Vision for
field extraction, and save the result to data/extracted/{survey_id}.json.

Usage:
    python scripts/01_extract.py                    ← process all PDFs (auto-detect type)
    python scripts/01_extract.py s001               ← single file, auto-detect type
    python scripts/01_extract.py s011 s020          ← range of files, auto-detect type
    python scripts/01_extract.py s012 412YZ         ← single file, explicit type override
    python scripts/01_extract.py s011 s020 412YZ    ← range + explicit type override
"""

import base64
import io
import json
import re
import sys
import traceback
from pathlib import Path

import anthropic
from dotenv import load_dotenv
from PIL import Image
from pdf2image import convert_from_path

# Load .env before importing config so ANTHROPIC_API_KEY is available
load_dotenv(Path(__file__).parent.parent / ".env")

sys.path.insert(0, str(Path(__file__).parent.parent))
import config

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def get_survey_type(pdf_path: Path) -> str:
    """Infer survey type from the filename (e.g. s001 → IL, s012 → 412YZ)."""
    match = re.search(r's(\d+)', pdf_path.stem, re.IGNORECASE)
    if match:
        number = int(match.group(1))
        return "IL" if number <= 11 else "412YZ"
    raise ValueError(f"Cannot determine survey type from filename: {pdf_path.stem}")


def pdf_to_base64_jpegs(pdf_path: Path, dpi: int = 150) -> list[str]:
    """Convert every page of a PDF to a base64-encoded JPEG string."""
    pages = convert_from_path(str(pdf_path), dpi=dpi, poppler_path=str(config.POPPLER_PATH))
    encoded = []
    for page in pages:
        # Convert to RGB so JPEG encoding never complains about alpha channel
        rgb = page.convert("RGB")
        buffer = io.BytesIO()
        rgb.save(buffer, format="JPEG", quality=85)
        encoded.append(base64.standard_b64encode(buffer.getvalue()).decode("utf-8"))
    return encoded


def load_prompt(survey_type: str) -> str:
    prompt_file = config.SURVEY_TYPES[survey_type]["prompt_file"]
    prompt_path = Path(__file__).parent.parent / prompt_file
    return prompt_path.read_text(encoding="utf-8")


def strip_markdown_fences(text: str) -> str:
    """Remove leading/trailing ```json ... ``` or ``` ... ``` wrappers if present."""
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    return text.strip()


def build_user_message(survey_id: str, b64_jpegs: list[str]) -> list[dict]:
    """
    Build the content block list for the user turn:
      - one image block per page
      - a closing text instruction that includes the survey_id
    """
    content = []
    for b64 in b64_jpegs:
        content.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": "image/jpeg",
                "data": b64,
            },
        })
    content.append({
        "type": "text",
        "text": (
            f'Extract all fields from this survey. '
            f'The survey_id is "{survey_id}".'
        ),
    })
    return content


def count_low_confidence(result: dict, threshold: float) -> int:
    confidence = result.get("confidence", {})
    return sum(1 for v in confidence.values() if isinstance(v, (int, float)) and v < threshold)


# ---------------------------------------------------------------------------
# Per-file extraction
# ---------------------------------------------------------------------------

def extract_survey(pdf_path: Path, client: anthropic.Anthropic, system_prompt: str) -> dict:
    survey_id = pdf_path.stem

    b64_jpegs = pdf_to_base64_jpegs(pdf_path)

    response = client.messages.create(
        model=config.EXTRACTION_MODEL,
        max_tokens=4096,
        system=system_prompt,
        messages=[
            {
                "role": "user",
                "content": build_user_message(survey_id, b64_jpegs),
            }
        ],
    )

    raw_text = response.content[0].text
    clean_text = strip_markdown_fences(raw_text)
    result = json.loads(clean_text)

    # Ensure survey_id is always set to the filename, not whatever the model returned
    result["survey_id"] = survey_id
    return result


def save_result(result: dict, survey_type: str) -> Path:
    extracted_dir = Path(__file__).parent.parent / config.SURVEY_TYPES[survey_type]["extracted_dir"]
    extracted_dir.mkdir(parents=True, exist_ok=True)
    out_path = extracted_dir / f"{result['survey_id']}.json"
    out_path.write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")
    return out_path


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    if not config.ANTHROPIC_API_KEY:
        print("ERROR: ANTHROPIC_API_KEY is not set. Add it to .env and retry.")
        sys.exit(1)

    client = anthropic.Anthropic(api_key=config.ANTHROPIC_API_KEY)

    # Determine PDF list and optional explicit type override
    explicit_type = None
    args = sys.argv[1:]

    # Check if last arg is a survey type override
    if args and args[-1] in config.SURVEY_TYPES:
        explicit_type = args[-1]
        args = args[:-1]

    def resolve_pdf(sid: str) -> Path:
        p = config.RAW_DIR / f"{sid.removesuffix('.pdf')}.pdf"
        if not p.exists():
            print(f"ERROR: {p} not found.")
            sys.exit(1)
        return p

    if len(args) == 2:
        # Range mode: s011 s020 [optional type]
        all_pdfs = sorted(config.RAW_DIR.glob("*.pdf"))
        names = [p.stem for p in all_pdfs]
        start_id = args[0].removesuffix(".pdf")
        end_id   = args[1].removesuffix(".pdf")
        for sid in (start_id, end_id):
            if sid not in names:
                print(f"ERROR: {sid}.pdf not found in {config.RAW_DIR}")
                sys.exit(1)
        start_i = names.index(start_id)
        end_i   = names.index(end_id)
        if start_i > end_i:
            print(f"ERROR: {start_id} comes after {end_id} in sort order.")
            sys.exit(1)
        pdf_files = all_pdfs[start_i : end_i + 1]
    elif len(args) == 1:
        # Single file: s001 [optional type]
        pdf_files = [resolve_pdf(args[0])]
    else:
        # All files mode
        pdf_files = sorted(config.RAW_DIR.glob("*.pdf"))
        if not pdf_files:
            print(f"No PDF files found in {config.RAW_DIR}")
            sys.exit(0)

    print(f"Found {len(pdf_files)} PDF(s) in {config.RAW_DIR}\n")

    successes = 0
    failures = 0

    # Cache prompts by type to avoid re-reading the file for every survey
    prompt_cache: dict[str, str] = {}

    for pdf_path in pdf_files:
        try:
            survey_type = explicit_type if explicit_type else get_survey_type(pdf_path)

            if survey_type not in prompt_cache:
                prompt_cache[survey_type] = load_prompt(survey_type)
            system_prompt = prompt_cache[survey_type]

            result = extract_survey(pdf_path, client, system_prompt)
            result["survey_type"] = survey_type
            out_path = save_result(result, survey_type)

            total_fields = len(result.get("fields", {}))
            low_conf = count_low_confidence(result, config.CONFIDENCE_THRESHOLD)

            print(
                f"[OK]   {pdf_path.name:<40} "
                f"type={survey_type:<6} "
                f"fields={total_fields}  "
                f"low_confidence={low_conf}  "
                f"-> {out_path.relative_to(Path(__file__).parent.parent)}"
            )
            successes += 1

        except json.JSONDecodeError as exc:
            print(f"[FAIL] {pdf_path.name}: JSON parse error — {exc}")
            failures += 1
        except anthropic.APIError as exc:
            print(f"[FAIL] {pdf_path.name}: Anthropic API error — {exc}")
            failures += 1
        except Exception:
            print(f"[FAIL] {pdf_path.name}: unexpected error")
            traceback.print_exc()
            failures += 1

    print(f"\nDone. {successes} succeeded, {failures} failed.")


if __name__ == "__main__":
    main()
