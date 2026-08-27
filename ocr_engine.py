"""
ocr_engine.py — OCR Orchestrator for GIECO Insurance Sync

Acts as the single entry point for certificate extraction. Coordinates:
    1. ImageCacheManager   — skip API calls for already-processed images
    2. GeminiVisionProvider — multi-image batch Gemini calls (primary)
    3. WinRTOCRProvider     — Windows-native offline fallback

Public API:
    extract_certificate_data(image_path)           -> dict   (single image, legacy compat)
    extract_certificates_batch(image_paths)        -> List[dict]   (preferred)
    _extract_single_via_winrt(image_path)          -> dict   (internal, used by WinRTOCRProvider)
"""

import logging
import os
import re
import sys
import asyncio
import tempfile
import cv2
import numpy as np

from pathlib import Path
from typing import Optional, Any, Dict, List

os.environ.setdefault("OPENBLAS_NUM_THREADS", "1")
os.environ.setdefault("MKL_NUM_THREADS", "1")
os.environ.setdefault("OMP_NUM_THREADS", "1")
os.environ.setdefault("NUMEXPR_NUM_THREADS", "1")

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

from stamp_remover import remove_stamp, crop_region
from cache_manager import ImageCacheManager
from ocr_provider import (
    GeminiVisionProvider,
    WinRTOCRProvider,
    DailyQuotaExhaustedError,
)

logger = logging.getLogger(__name__)

# Module-level singletons — reused across calls
_cache = ImageCacheManager()
_gemini_provider: Optional[GeminiVisionProvider] = None
_winrt_provider: Optional[WinRTOCRProvider] = None

# Arabic-Indic digit transliteration
ARABIC_INDIC_MAP = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")

# Bounding box regions calibrated on 960×1280 reference image
REGIONS_RAPID = {
    "office_code":  (340, 420, 50, 500),
    "plate_digits": (460, 540, 50, 450),
    "insurance_no": (480, 560, 380, 900),
    "chassis_no":   (520, 590, 380, 900),
}
REGIONS_EASYOCR = {
    "license_line":    (720, 755, 30, 950),
    "cert_end_line":   (770, 810, 30, 950),
    "print_date_line": (1210, 1260, 30, 700),
}


# ── Provider access helpers ───────────────────────────────────────────────────

def _get_gemini_provider() -> GeminiVisionProvider:
    global _gemini_provider
    if _gemini_provider is None:
        _gemini_provider = GeminiVisionProvider()
    return _gemini_provider


def _get_winrt_provider() -> WinRTOCRProvider:
    global _winrt_provider
    if _winrt_provider is None:
        _winrt_provider = WinRTOCRProvider()
    return _winrt_provider


# ── Public Batch API ──────────────────────────────────────────────────────────

def extract_certificates_batch(image_paths: List[str]) -> List[dict]:
    """
    Extract certificate data from a list of images using cache + batch Gemini.

    This is the preferred entry point for the GUI batch worker. It will:
        1. Serve cached results instantly.
        2. Send uncached images to Gemini in groups of GEMINI_BATCH_SIZE.
        3. Fall back to WinRT OCR if Gemini is unavailable.
        4. Re-raise DailyQuotaExhaustedError so the GUI can handle it gracefully.

    Args:
        image_paths: List of absolute paths to certificate images.

    Returns:
        List[dict] of the same length — each dict contains extracted fields.
        On failure for a given image, returns a dict with all fields None.

    Raises:
        DailyQuotaExhaustedError: When Gemini signals the daily quota is gone.
    """
    results: List[Optional[dict]] = [None] * len(image_paths)
    uncached_indices: List[int] = []

    # Step 1: Check cache
    for i, path in enumerate(image_paths):
        cached = _cache.get(path)
        if cached is not None:
            results[i] = cached
        else:
            uncached_indices.append(i)

    stats = _cache.get_stats()
    if uncached_indices:
        logger.info(
            "[OCR] %d cache hits, %d to extract via API.",
            len(image_paths) - len(uncached_indices), len(uncached_indices),
        )
    else:
        logger.info("[OCR] All %d images served from cache (hit rate: %.0f%%).",
                    len(image_paths), stats["hit_rate"] * 100)

    if not uncached_indices:
        return [r or _empty_result() for r in results]

    # Step 2: Extract via Gemini (may raise DailyQuotaExhaustedError)
    uncached_paths = [image_paths[i] for i in uncached_indices]
    provider = _get_gemini_provider()
    batch_results = provider.extract_batch(uncached_paths)

    # Step 3: Fill cache + fall back to WinRT for None slots
    winrt_fallback_indices: List[int] = []
    for local_idx, global_idx in enumerate(uncached_indices):
        data = batch_results[local_idx]
        if data is not None:
            _cache.set(image_paths[global_idx], data)
            results[global_idx] = data
        else:
            winrt_fallback_indices.append(global_idx)

    # Step 4: WinRT fallback for images Gemini couldn't parse
    if winrt_fallback_indices:
        logger.info("[OCR] %d images falling back to WinRT OCR.", len(winrt_fallback_indices))
        winrt = _get_winrt_provider()
        fallback_paths = [image_paths[i] for i in winrt_fallback_indices]
        fallback_results = winrt.extract_batch(fallback_paths)
        for local_idx, global_idx in enumerate(winrt_fallback_indices):
            data = fallback_results[local_idx]
            if data is not None:
                _cache.set(image_paths[global_idx], data)
                results[global_idx] = data
            else:
                results[global_idx] = _empty_result()

    # Log final cache stats
    final_stats = _cache.get_stats()
    logger.info(
        "[OCR] Run complete. Cache hit rate: %.0f%% (%d hits / %d total). "
        "Gemini API calls this session: %d.",
        final_stats["hit_rate"] * 100,
        final_stats["hits"],
        final_stats["total"],
        provider.request_count,
    )

    return [r or _empty_result() for r in results]


# ── Single-image legacy API (preserves backward compatibility) ────────────────

def extract_certificate_data(image_path: str) -> dict:
    """
    Single-image extraction. Maintained for backward compatibility with
    test_pipeline.py and any single-item callers.

    Internally calls extract_certificates_batch with a list of one.
    """
    try:
        results = extract_certificates_batch([image_path])
        return results[0]
    except DailyQuotaExhaustedError:
        logger.warning("[OCR] Daily quota exhausted — falling back to WinRT for single image.")
        try:
            return _extract_single_via_winrt(image_path)
        except Exception as e:
            logger.error("[OCR] WinRT fallback also failed: %s", e)
            return _empty_result()


def _empty_result() -> dict:
    """Return a null-field result dict so the GUI always gets a well-typed object."""
    return {
        "service_no": None, "office_code": None, "plate_digits": None,
        "plate_letters": None, "plate_full": None, "insurance_no": None,
        "chassis_no": None, "driver_name": None, "date_from": None,
        "date_to": None, "print_date": None,
        "engine_used": "None",
        "confidence": {k: 0.0 for k in
                       ("office_code", "plate_digits", "insurance_no",
                        "chassis_no", "license_line", "service_no")},
        "raw_ocr": {},
    }


# ── WinRT OCR helpers (used by WinRTOCRProvider) ─────────────────────────────

def _extract_single_via_winrt(image_path: str) -> dict:
    """
    Full WinRT OCR pipeline for a single image: stamp removal → region OCR → field assembly.
    Called by WinRTOCRProvider.extract_batch().
    """
    logger.info("[WinRT] Processing: %s", os.path.basename(image_path))
    clean_image = remove_stamp(image_path)

    MAX_DIM = 1600
    h, w = clean_image.shape[:2]
    if h > MAX_DIM or w > MAX_DIM:
        r = min(MAX_DIM / h, MAX_DIM / w)
        clean_image = cv2.resize(
            clean_image,
            (int(w * r), int(h * r)),
            interpolation=cv2.INTER_AREA,
        )
    img_h, img_w = clean_image.shape[:2]

    engine = _init_winrt_engine()

    full_results, full_text = _ocr_fullimage(engine, clean_image)
    region_results = {
        name: _ocr_region(engine, clean_image, region, img_h, img_w, scale=3.0, region_name=name)
        for name, region in REGIONS_RAPID.items()
    }
    easyocr_results = {
        name: _ocr_region(engine, clean_image, region, img_h, img_w, scale=4.0, region_name=name)
        for name, region in REGIONS_EASYOCR.items()
    }

    service_no   = _extract_service_no(full_results)
    office_code  = _extract_number(region_results.get("office_code", []), min_digits=5)
    plate_digits = _extract_number(region_results.get("plate_digits", []), min_digits=3)
    insurance_no = _extract_number(region_results.get("insurance_no", []), min_digits=5)
    chassis_no   = _extract_number(region_results.get("chassis_no", []), min_digits=3)

    dates      = _extract_dates_from_license_line(easyocr_results.get("license_line", []))
    cert_end   = _extract_single_date(easyocr_results.get("cert_end_line", []))
    print_date = _extract_single_date(easyocr_results.get("print_date_line", []))

    # Robust full-page date fallback
    if not dates.get("date_from") or not dates.get("date_to"):
        clean_full = transliterate_arabic_digits(full_text)
        all_dates = re.findall(r"(\d{4})[\s/\-.]* (\d{1,2})[\s/\-.]*(\d{1,2})", clean_full)
        parsed = sorted({
            f"{y}-{m.zfill(2)}-{d.zfill(2)}"
            for y, m, d in all_dates
            if 2020 <= int(y) <= 2035 and 1 <= int(m) <= 12 and 1 <= int(d) <= 31
        })
        if len(parsed) >= 2:
            dates.setdefault("date_from", parsed[0])
            dates.setdefault("date_to", parsed[-1])
        elif len(parsed) == 1:
            dates.setdefault("date_from", parsed[0])

    if not dates.get("date_to") and cert_end:
        dates["date_to"] = cert_end

    confidence = {}
    for field, res in region_results.items():
        confidence[field] = max((c for _, c in res), default=0.0)
    confidence["service_no"] = 0.95 if service_no else 0.0
    for field, res in easyocr_results.items():
        confidence[field] = max((c for _, c in res), default=0.0)

    raw_ocr = {
        "full_scan": [f"y={y} x={x}: {t}" for t, _, y, x in full_results],
        **{f: [t for t, _ in res] for f, res in region_results.items()},
        **{f: [t for t, _ in res] for f, res in easyocr_results.items()},
    }

    return {
        "service_no":   service_no,
        "office_code":  office_code,
        "plate_digits": plate_digits,
        "plate_letters": None,
        "plate_full":   None,
        "insurance_no": insurance_no,
        "chassis_no":   chassis_no,
        "driver_name":  None,
        "date_from":    dates.get("date_from"),
        "date_to":      dates.get("date_to"),
        "print_date":   print_date,
        "engine_used":  "WinRT OCR",
        "confidence":   confidence,
        "raw_ocr":      raw_ocr,
    }


# ── WinRT helpers ─────────────────────────────────────────────────────────────

def _init_winrt_engine():
    try:
        from winsdk.windows.media.ocr import OcrEngine
        import winsdk.windows.globalization as wg
        lang = wg.Language("ar")
        engine = OcrEngine.try_create_from_language(lang)
        if engine is None:
            engine = OcrEngine.try_create_from_user_profile_languages()
        return engine
    except Exception as e:
        logger.error("[WinRT] Engine init failed: %s", e)
        return None


def _scale_region(region: tuple, img_h: int, img_w: int) -> tuple:
    y1, y2, x1, x2 = region
    return int(y1 * img_h / 1280), int(y2 * img_h / 1280), \
           int(x1 * img_w / 960),  int(x2 * img_w / 960)


def _ocr_region(engine, image, region, img_h, img_w, scale=3.0, region_name="") -> list:
    y1, y2, x1, x2 = _scale_region(region, img_h, img_w)
    y1, y2 = max(0, y1), max(1, min(y2, img_h))
    x1, x2 = max(0, x1), max(1, min(x2, img_w))
    if y2 <= y1 or x2 <= x1:
        return []
    cropped = crop_region(image, y1, y2, x1, x2, scale=scale)
    if cropped is None or cropped.size == 0:
        return []
    result = _run_ocr_numpy(cropped, engine)
    if not result or not result.text:
        return []
    return [(line.text.strip(), 0.9) for line in result.lines if line.text.strip()]


def _ocr_fullimage(engine, image) -> tuple:
    h, w = image.shape[:2]
    if h > 2000 or w > 2000:
        r = min(2000 / h, 2000 / w)
        image = cv2.resize(image, (int(w * r), int(h * r)), interpolation=cv2.INTER_AREA)
    result = _run_ocr_numpy(image, engine)
    if not result or not result.text:
        return [], ""
    parsed = []
    lines = []
    for line in result.lines:
        lines.append(line.text)
        for word in line.words:
            t = word.text.strip()
            if t:
                rect = word.bounding_rect
                parsed.append((t, 0.95, rect.y + rect.height // 2, rect.x + rect.width // 2))
    return parsed, "\n".join(lines)


async def _ocr_image_async(image: np.ndarray, engine: Any):
    if engine is None:
        return None
    tmp = os.path.join(tempfile.gettempdir(), "temp_ocr_bridge.png")
    cv2.imwrite(tmp, image)
    try:
        from winsdk.windows.graphics.imaging import BitmapDecoder
        from winsdk.windows.storage import StorageFile, FileAccessMode
        f = await StorageFile.get_file_from_path_async(tmp)
        stream = await f.open_async(FileAccessMode.READ)
        decoder = await BitmapDecoder.create_async(stream)
        bmp = await decoder.get_software_bitmap_async()
        result = await engine.recognize_async(bmp)
        stream.close()
        return result
    except Exception as e:
        logger.warning("[WinRT] OCR failed: %s", e)
        return None


def _run_ocr_numpy(image: np.ndarray, engine: Any):
    try:
        loop = asyncio.get_event_loop()
        if loop.is_running():
            return None
    except RuntimeError:
        pass
    return asyncio.run(_ocr_image_async(image, engine))


# ── Field Extraction Helpers (WinRT path) ─────────────────────────────────────

def transliterate_arabic_digits(text: str) -> str:
    return text.translate(ARABIC_INDIC_MAP).replace("\\", "/").replace("|", "/")


def _extract_service_no(full_results: list) -> Optional[str]:
    full_text = " ".join(w for w, _, __, ___ in full_results)
    m = re.search(r"cs[-\s]?\d{8}[-\s]?\d+", full_text, re.IGNORECASE)
    if m:
        code = m.group().replace(" ", "").upper()
        code = re.sub(r"CS(\d{8})(\d+)", r"CS-\1-\2", code)
        return code
    return None


def _extract_number(ocr_results: list, min_digits: int = 3) -> Optional[str]:
    best, best_conf = None, 0.0
    for text, conf in ocr_results:
        for num in re.findall(r"\d+", text):
            if len(num) >= min_digits and conf > best_conf:
                best, best_conf = num, conf
        for an in re.findall(r"[A-Z0-9]{5,}", text.upper()):
            if conf > best_conf:
                best, best_conf = an, conf
    return best


def _reconstruct_rtl_date(digit_groups: list) -> Optional[str]:
    if not digit_groups:
        return None
    all_digits = "".join(digit_groups)
    for candidate in (all_digits, "".join(reversed(digit_groups)), all_digits[::-1]):
        if len(candidate) >= 8:
            try:
                y, m, d = candidate[:4], candidate[4:6], candidate[6:8]
                if 2020 <= int(y) <= 2035 and 1 <= int(m) <= 12 and 1 <= int(d) <= 31:
                    return f"{y}-{m}-{d}"
            except ValueError:
                pass
    return None


def _extract_dates_from_license_line(ocr_results: list) -> dict:
    dates: dict = {"date_from": None, "date_to": None}
    for text, _ in ocr_results:
        clean = transliterate_arabic_digits(text)
        for y, m, d in re.findall(r"(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})", clean):
            s = f"{y}-{m.zfill(2)}-{d.zfill(2)}"
            if not dates["date_from"]:
                dates["date_from"] = s
            elif not dates["date_to"]:
                dates["date_to"] = s
    if dates["date_from"] and dates["date_to"] and dates["date_from"] > dates["date_to"]:
        dates["date_from"], dates["date_to"] = dates["date_to"], dates["date_from"]
    return dates


def _extract_single_date(ocr_results: list) -> Optional[str]:
    for text, _ in ocr_results:
        clean = transliterate_arabic_digits(text)
        m = re.search(r"(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})", clean)
        if m:
            y, mo, d = m.groups()
            return f"{y}-{mo.zfill(2)}-{d.zfill(2)}"
    return None


if __name__ == "__main__":
    import sys as _sys
    logging.basicConfig(level=logging.INFO)
    if len(_sys.argv) < 2:
        print("Usage: python ocr_engine.py <image_path>")
        _sys.exit(1)
    result = extract_certificate_data(_sys.argv[1])
    print("\n" + "=" * 60)
    for k, v in result.items():
        if k not in ("confidence", "raw_ocr"):
            print(f"  {'✅' if v else '❌'} {k:15s}: {v}")
    print("\nConfidence:")
    for k, c in result.get("confidence", {}).items():
        print(f"  {'✅' if c >= 0.7 else '⚠️' if c >= 0.5 else '❌'} {k:15s}: {c:.2f}")
    print(f"\nCache stats: {_cache.get_stats()}")
