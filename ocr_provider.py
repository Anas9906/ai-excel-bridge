"""
ocr_provider.py — Pluggable OCR Provider Interface for AI Excel Bridge

Defines a stable BaseOCRProvider abstraction so additional vision backends
(OpenRouter, Mistral Pixtral, Azure Vision, etc.) can be added as a simple
subclass without touching the rest of the pipeline.

Currently implemented providers:
    GeminiVisionProvider  — Google Gemini Vision AI (primary, free-tier aware)
    WinRTOCRProvider      — Windows-native WinRT OCR (offline fallback)

Usage:
    from ocr_provider import GeminiVisionProvider, DailyQuotaExhaustedError
    provider = GeminiVisionProvider()
    results = provider.extract_batch(["cert1.jpeg", "cert2.jpeg", "cert3.jpeg"])
    # returns List[Optional[dict]], one entry per image path
"""

import json
import logging
import os
import re
import time
import random
from abc import ABC, abstractmethod
from pathlib import Path
from typing import List, Optional

logger = logging.getLogger(__name__)

# ── Configuration (overridable via .env) ─────────────────────────────────────
GEMINI_BATCH_SIZE = int(os.environ.get("GEMINI_BATCH_SIZE", 3))
GEMINI_RPM        = int(os.environ.get("GEMINI_RPM", 15))
GEMINI_MODEL      = os.environ.get("GEMINI_MODEL", "gemini-1.5-flash")

# 429 sub-type identifiers (inspected from Gemini error messages)
_RPD_SIGNALS = [
    "quota exceeded per day",
    "resource_exhausted",
    "daily",
    "per_day",
    "rpd",
]


# ── Custom Exceptions ────────────────────────────────────────────────────────

class DailyQuotaExhaustedError(Exception):
    """Raised when Gemini signals the *daily* quota is exhausted for this key.
    The caller should cease API calls, persist remaining work, and stop cleanly.
    """


class BatchExtractionError(Exception):
    """Raised when a batch fails after all local repairs and retries."""


# ── Abstract Provider Interface ───────────────────────────────────────────────

class BaseOCRProvider(ABC):
    """
    Pluggable seam for all OCR/Vision backends.

    Implementors receive a list of image paths and MUST return a list of the
    same length, where each entry is either:
        - A dict containing the extracted certificate fields
        - None if extraction for that image failed
    """

    @abstractmethod
    def extract_batch(self, image_paths: List[str]) -> List[Optional[dict]]:
        """
        Extract certificate data from a batch of image paths.

        Args:
            image_paths: List of absolute paths to certificate images.

        Returns:
            List[Optional[dict]]: Length == len(image_paths).
            None entries indicate per-image extraction failure.

        Raises:
            DailyQuotaExhaustedError: When the daily API quota is exhausted.
        """
        ...


# ── Local JSON Repair Utility ─────────────────────────────────────────────────

def _repair_json(raw: str) -> Optional[object]:
    """
    Attempt to parse possibly-malformed JSON without an API call.

    Steps applied in order:
        1. Strip leading/trailing whitespace.
        2. Remove markdown code fences (```json ... ```).
        3. Remove trailing commas before ] or }.
        4. Attempt standard json.loads.
        5. Extract first {...} or [...] block if bare parse fails.

    Returns the parsed Python object or None if all attempts fail.
    """
    if not raw or not raw.strip():
        return None

    text = raw.strip()

    # Strip markdown code fences
    text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\s*```$", "", text)
    text = text.strip()

    # Remove trailing commas before closing brackets/braces
    text = re.sub(r",(\s*[}\]])", r"\1", text)

    # First attempt: direct parse
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Second attempt: extract outermost JSON block
    for pattern in (r"(\[.*\])", r"(\{.*\})"):
        match = re.search(pattern, text, re.DOTALL)
        if match:
            try:
                return json.loads(match.group(1))
            except json.JSONDecodeError:
                continue

    logger.warning("[JSON REPAIR] All local repair attempts failed.")
    return None


# ── Gemini Vision Provider ─────────────────────────────────────────────────────

class GeminiVisionProvider(BaseOCRProvider):
    """
    Multi-image batch extractor using Google Gemini Vision AI.

    Key design decisions:
    - Sends GEMINI_BATCH_SIZE images per request (default 3).
    - Validates returned array length matches batch size.
    - Runs local zero-cost JSON repair before consuming an API retry.
    - Distinguishes RPM/TPM 429s (backoff + retry) from RPD quota exhaustion
      (DailyQuotaExhaustedError — caller must stop and save state).
    - Exposes a `request_count` property for daily quota monitoring.
    """

    _FIELD_PROMPT = """\
You are an expert Arabic OCR parser for Egyptian Vehicle Insurance Certificates
(شهادة سداد اشتراكات التأمين الاجتماعي عن سيارة).

I am sending you {n} certificate images. Return ONLY a JSON array of exactly
{n} objects — one per image — in the same order as the images provided.
Each object must use this exact schema (use null for missing fields):

{{
  "image_index": <int, 0-based>,
  "service_no":  <str|null>,
  "office_code": <str|null>,
  "plate_digits": <str|null>,
  "plate_letters": <str|null>,
  "plate_full": <str|null>,
  "insurance_no": <str|null>,
  "chassis_no": <str|null>,
  "driver_name": <str|null>,
  "date_from": <YYYY-MM-DD|null>,
  "date_to": <YYYY-MM-DD|null>,
  "cert_end_date": <YYYY-MM-DD|null>
}}

Return ONLY the JSON array. No explanation, no markdown, no extra text.
The array MUST have exactly {n} elements."""

    _STRICT_RETRY_SUFFIX = (
        "\n\nIMPORTANT: Your previous response was not valid JSON or had the wrong"
        " array length. Return ONLY the raw JSON array with no other text."
    )

    def __init__(self, api_key: Optional[str] = None, batch_size: int = GEMINI_BATCH_SIZE):
        self._api_key = api_key or self._load_api_key()
        self.batch_size = batch_size
        self.request_count = 0
        self._client = None  # lazy-loaded

    @staticmethod
    def _load_api_key() -> Optional[str]:
        key = os.environ.get("GEMINI_API_KEY", "").strip()
        if key:
            return key
        for candidate in [Path(__file__).parent / ".env",
                           Path(__file__).parent.parent / ".env"]:
            if candidate.exists():
                for line in candidate.read_text(encoding="utf-8").splitlines():
                    if line.startswith("GEMINI_API_KEY="):
                        return line.split("=", 1)[1].strip().strip('"\'')
        return None

    def _get_client(self):
        if self._client is None:
            from google import genai
            self._client = genai.Client(api_key=self._api_key)
        return self._client

    def extract_batch(self, image_paths: List[str]) -> List[Optional[dict]]:
        if not self._api_key:
            logger.warning("[GeminiProvider] No API key — returning None for all images.")
            return [None] * len(image_paths)

        results: List[Optional[dict]] = [None] * len(image_paths)

        # Chunk into sub-batches of self.batch_size
        for chunk_start in range(0, len(image_paths), self.batch_size):
            chunk = image_paths[chunk_start: chunk_start + self.batch_size]
            chunk_results = self._extract_chunk(chunk)
            for local_idx, data in enumerate(chunk_results):
                results[chunk_start + local_idx] = data

        return results

    def _extract_chunk(self, image_paths: List[str]) -> List[Optional[dict]]:
        """Process a single sub-batch of up to self.batch_size images."""
        from PIL import Image
        from google.genai import types

        n = len(image_paths)
        prompt = self._FIELD_PROMPT.format(n=n)

        images = []
        for p in image_paths:
            try:
                img = Image.open(p)
                MAX_DIM = 1600
                if img.width > MAX_DIM or img.height > MAX_DIM:
                    scale = min(MAX_DIM / img.width, MAX_DIM / img.height)
                    img = img.resize(
                        (int(img.width * scale), int(img.height * scale)),
                        Image.Resampling.LANCZOS
                    )
                images.append(img)
            except Exception as e:
                logger.error("[GeminiProvider] Cannot open image %s: %s", p, e)
                images.append(None)

        if all(img is None for img in images):
            return [None] * n

        # Build contents list: interleave images + prompt
        contents = [img for img in images if img is not None] + [prompt]

        raw_response = self._call_gemini_with_backoff(contents, types)
        if raw_response is None:
            return [None] * n

        # Local JSON repair before any API retry
        parsed = _repair_json(raw_response)
        if not isinstance(parsed, list) or len(parsed) != n:
            logger.warning(
                "[GeminiProvider] Response length mismatch or invalid JSON (%s items, expected %d). "
                "Attempting local repair, then one strict retry.",
                len(parsed) if isinstance(parsed, list) else "N/A", n,
            )
            # One strict API retry
            strict_contents = contents[:-1] + [prompt + self._STRICT_RETRY_SUFFIX]
            raw_response = self._call_gemini_with_backoff(strict_contents, types)
            parsed = _repair_json(raw_response) if raw_response else None
            if not isinstance(parsed, list) or len(parsed) != n:
                logger.error("[GeminiProvider] Strict retry also failed — returning None for batch.")
                return [None] * n

        # Normalize each item to the standard field dict
        return [self._normalize(item) if isinstance(item, dict) else None for item in parsed]

    def _call_gemini_with_backoff(self, contents, types, max_retries: int = 3) -> Optional[str]:
        """Call Gemini with exponential backoff on RPM 429s. Raises on RPD exhaustion."""
        client = self._get_client()
        delay = 2.0

        for attempt in range(max_retries + 1):
            try:
                response = client.models.generate_content(
                    model=GEMINI_MODEL,
                    contents=contents,
                    config=types.GenerateContentConfig(response_mime_type="application/json"),
                )
                self.request_count += 1
                logger.info("[GeminiProvider] API request #%d succeeded (attempt %d).",
                            self.request_count, attempt + 1)
                return response.text if response and response.text else None

            except Exception as e:
                err_str = str(e).lower()

                # --- Daily quota exhaustion: do NOT retry ---
                if any(sig in err_str for sig in _RPD_SIGNALS):
                    logger.error(
                        "[GeminiProvider] Daily quota exhausted: %s. "
                        "Stopping to preserve remaining quota.",
                        e,
                    )
                    raise DailyQuotaExhaustedError(str(e)) from e

                # --- RPM / TPM 429: backoff and retry ---
                if "429" in err_str or "rate" in err_str or "quota" in err_str:
                    if attempt < max_retries:
                        jitter = random.uniform(0, delay * 0.3)
                        wait = delay + jitter
                        logger.warning(
                            "[GeminiProvider] 429 rate-limit (attempt %d/%d). "
                            "Backing off for %.1fs.",
                            attempt + 1, max_retries, wait,
                        )
                        time.sleep(wait)
                        delay *= 2
                        continue
                    logger.error("[GeminiProvider] Max retries exceeded on 429.")
                    return None

                # --- Other errors ---
                logger.error("[GeminiProvider] API error: %s", e)
                return None

        return None

    @staticmethod
    def _normalize(item: dict) -> dict:
        """Convert a raw Gemini JSON object to the standard pipeline schema."""
        def _s(v) -> Optional[str]:
            return str(v).strip() or None if v is not None else None

        return {
            "service_no":     _s(item.get("service_no")),
            "office_code":    _s(item.get("office_code")),
            "plate_digits":   _s(item.get("plate_digits")),
            "plate_letters":  _s(item.get("plate_letters")),
            "plate_full":     _s(item.get("plate_full")),
            "insurance_no":   _s(item.get("insurance_no")),
            "chassis_no":     _s(item.get("chassis_no")),
            "driver_name":    _s(item.get("driver_name")),
            "date_from":      _s(item.get("date_from")),
            "date_to":        _s(item.get("date_to")),
            "print_date":     _s(item.get("cert_end_date")),
            "engine_used":    "Gemini Vision AI",
            "confidence": {
                "office_code":  0.99 if item.get("office_code") else 0.0,
                "plate_digits": 0.99 if item.get("plate_digits") else 0.0,
                "insurance_no": 0.99 if item.get("insurance_no") else 0.0,
                "chassis_no":   0.99 if item.get("chassis_no") else 0.0,
                "license_line": 0.99 if (item.get("date_from") and item.get("date_to")) else 0.0,
                "service_no":   0.99 if item.get("service_no") else 0.0,
            },
            "raw_ocr": {"ai_json": item},
        }


# ── WinRT OCR Provider (Offline Fallback) ────────────────────────────────────

class WinRTOCRProvider(BaseOCRProvider):
    """
    Windows-native WinRT OCR fallback — used when Gemini is unavailable
    or the daily quota is exhausted.

    Processes images one-by-one (WinRT is not batch-capable) and delegates
    to the existing stamp_remover + WinRT pipeline.
    """

    def extract_batch(self, image_paths: List[str]) -> List[Optional[dict]]:
        # Import here to avoid loading Windows-only modules on startup
        from ocr_engine import _extract_single_via_winrt
        results = []
        for path in image_paths:
            try:
                data = _extract_single_via_winrt(path)
                results.append(data)
            except Exception as e:
                logger.error("[WinRTProvider] Extraction failed for %s: %s", path, e)
                results.append(None)
        return results
