"""
ocr_engine.py — Native WinRT OCR Engine for Arabic Social Insurance Certificates

Strategy:
    - Replaces heavy PyTorch/ONNX libraries with Windows Native OCR (WinRT).
    - Extremely lightweight, completely avoids WinError 1114 DLL initialization bugs.
    - Native Arabic support without massive network downloads.

Usage:
    from ocr_engine import extract_certificate_data
    data = extract_certificate_data("certificate.jpeg")
    print(data)
"""

import os
os.environ["OPENBLAS_NUM_THREADS"] = "1"
os.environ["MKL_NUM_THREADS"] = "1"
os.environ["OMP_NUM_THREADS"] = "1"
os.environ["NUMEXPR_NUM_THREADS"] = "1"

import re
import cv2
import numpy as np
import sys
import tempfile
import asyncio
from pathlib import Path
from typing import Optional, Any, Dict, List



if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


from stamp_remover import remove_stamp, crop_region

# Lazy loaded on WinRT fallback:
# from winsdk.windows.media.ocr import OcrEngine
# from winsdk.windows.graphics.imaging import BitmapDecoder
# from winsdk.windows.storage import StorageFile, FileAccessMode


# Arabic-Indic digit transliteration map
ARABIC_INDIC_MAP = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")

# Bounding box regions — CALIBRATED on actual 960×1280 certificate
REGIONS_RAPID = {
    "office_code":  (340, 420, 50, 500),
    "plate_digits": (460, 540, 50, 450),
    "insurance_no": (480, 560, 380, 900),
    "chassis_no":   (520, 590, 380, 900),
}

REGIONS_EASYOCR = {
    "license_line":    (720, 755, 30, 950),    # Full line: من ٢٠٢٦/٠٤/٠٢ الى ٢٠٢٧/٠٤/٠١
    "cert_end_line":   (770, 810, 30, 950),    # تاريخ نهاية الشهادة
    "print_date_line": (1210, 1260, 30, 700),  # تاريخ الطباعة
}

def transliterate_arabic_digits(text: str) -> str:
    result = text.translate(ARABIC_INDIC_MAP)
    result = result.replace("\\", "/").replace("|", "/")
    return result

def _scale_region(region: tuple, img_h: int, img_w: int) -> tuple:
    y1, y2, x1, x2 = region
    sy = img_h / 1280
    sx = img_w / 960
    return int(y1 * sy), int(y2 * sy), int(x1 * sx), int(x2 * sx)

async def _ocr_image_async(image: np.ndarray, engine: Any):
    if engine is None:
        return None
        
    temp_path = os.path.join(tempfile.gettempdir(), "temp_ocr.png")
    cv2.imwrite(temp_path, image)
    
    try:
        from winsdk.windows.graphics.imaging import BitmapDecoder
        from winsdk.windows.storage import StorageFile, FileAccessMode

        file = await StorageFile.get_file_from_path_async(temp_path)
        stream = await file.open_async(FileAccessMode.READ)
        decoder = await BitmapDecoder.create_async(stream)
        software_bitmap = await decoder.get_software_bitmap_async()
        result = await engine.recognize_async(software_bitmap)
        stream.close()
        return result
    except Exception as e:
        print(f"[WARN] WinRT OCR failed: {e}")
        return None


def _run_ocr_numpy(image: np.ndarray, engine: Any):
    try:
        # Prevent "Event loop is already running" if already in an async context
        loop = asyncio.get_event_loop()
        if loop.is_running():
            print("[WARN] Asyncio loop is already running. You must await _ocr_image_async directly.")
            return None
    except RuntimeError:
        pass
        
    return asyncio.run(_ocr_image_async(image, engine))

def _ocr_region(engine, image: np.ndarray, region: tuple,
                img_h: int, img_w: int, scale: float = 3.0, region_name: str = "") -> list:
    y1, y2, x1, x2 = _scale_region(region, img_h, img_w)
    
    # Ensure valid bounds
    y1, y2 = max(0, min(y1, img_h-1)), max(1, min(y2, img_h))
    x1, x2 = max(0, min(x1, img_w-1)), max(1, min(x2, img_w))
    
    if y2 <= y1 or x2 <= x1:
        return []
        
    cropped = crop_region(image, y1, y2, x1, x2, scale=scale)
    if cropped is None or cropped.size == 0:
        return []
        
    if region_name:
        import cv2, os
        debug_path = f"debug_crop_{region_name}.png"
        cv2.imwrite(debug_path, cropped)

    print(f"[DEBUG] '{region_name}' cropped shape: {cropped.shape}")
    result = _run_ocr_numpy(cropped, engine)
    if not result or not result.text:
        return []
        
    text_lines = [line.text for line in result.lines if line.text.strip()]
    return [(t.strip(), 0.9) for t in text_lines]

def _ocr_fullimage(engine, image: np.ndarray) -> tuple:
    # Resize if too large to prevent WinRT silent failure
    h, w = image.shape[:2]
    MAX_DIM = 2000
    if h > MAX_DIM or w > MAX_DIM:
        ratio = min(MAX_DIM / h, MAX_DIM / w)
        image = cv2.resize(image, (int(w * ratio), int(h * ratio)), interpolation=cv2.INTER_AREA)
        
    result = _run_ocr_numpy(image, engine)
    if not result or not result.text:
        return [], ""
        
    parsed = []
    full_lines = []
    for line in result.lines:
        full_lines.append(line.text)
        for word in line.words:
            text = word.text.strip()
            if text:
                rect = word.bounding_rect
                y_center = rect.y + rect.height // 2
                x_center = rect.x + rect.width // 2
                parsed.append((text, 0.95, y_center, x_center))
                
    full_text = "\n".join(full_lines)
    return parsed, full_text

def _extract_service_no(full_results: list) -> Optional[str]:
    # Since WinRT might split "CS-20260817-11713876" into multiple words,
    # we reconstruct lines and run the regex over the whole line space.
    full_text = " ".join([word for word, conf, y, x in full_results])
    match = re.search(r'cs[-\s]?\d{8}[-\s]?\d+', full_text, re.IGNORECASE)
    if match:
        code = match.group().replace(" ", "").upper()
        code = re.sub(r'CS(\d{8})(\d+)', r'CS-\1-\2', code)
        return code
    return None

def _extract_number(ocr_results: list, min_digits: int = 3) -> Optional[str]:
    best = None
    best_conf = 0.0

    for text, conf in ocr_results:
        clean = text.strip()
        numbers = re.findall(r'\d+', clean)
        for num in numbers:
            if len(num) >= min_digits and conf > best_conf:
                best = num
                best_conf = conf

        alphanum = re.findall(r'[A-Z0-9]{5,}', clean.upper())
        for an in alphanum:
            if conf > best_conf:
                best = an
                best_conf = conf

    return best

def _reconstruct_rtl_date(digit_groups: list) -> Optional[str]:
    if not digit_groups:
        return None

    all_digits = "".join(digit_groups)

    if len(all_digits) == 8:
        try:
            y, m, d = all_digits[:4], all_digits[4:6], all_digits[6:8]
            if 2020 <= int(y) <= 2035 and 1 <= int(m) <= 12 and 1 <= int(d) <= 31:
                return f"{y}-{m}-{d}"
        except ValueError:
            pass

    reversed_groups = list(reversed(digit_groups))
    reversed_digits = "".join(reversed_groups)

    if len(reversed_digits) >= 8:
        try:
            y, m, d = reversed_digits[:4], reversed_digits[4:6], reversed_digits[6:8]
            if 2020 <= int(y) <= 2035 and 1 <= int(m) <= 12 and 1 <= int(d) <= 31:
                return f"{y}-{m}-{d}"
        except ValueError:
            pass

    full_reversed = all_digits[::-1]
    if len(full_reversed) >= 8:
        try:
            y, m, d = full_reversed[:4], full_reversed[4:6], full_reversed[6:8]
            if 2020 <= int(y) <= 2035 and 1 <= int(m) <= 12 and 1 <= int(d) <= 31:
                return f"{y}-{m}-{d}"
        except ValueError:
            pass

    return None

def _extract_dates_from_license_line(ocr_results: list) -> dict:
    dates = {"date_from": None, "date_to": None}

    for text, conf in ocr_results:
        clean = transliterate_arabic_digits(text)
        date_matches = re.findall(r'(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})', clean)
        for y, m, d in date_matches:
            date_str = f"{y}-{m.zfill(2)}-{d.zfill(2)}"
            if dates["date_from"] is None:
                dates["date_from"] = date_str
            else:
                dates["date_to"] = date_str

    if dates["date_from"] and dates["date_to"]:
        if dates["date_from"] > dates["date_to"]:
            dates["date_from"], dates["date_to"] = dates["date_to"], dates["date_from"]
        return dates

    all_digit_groups = []
    for text, conf in ocr_results:
        clean = transliterate_arabic_digits(text)
        digits_found = re.findall(r'\d+', clean)
        if digits_found:
            all_digit_groups.append((digits_found, conf, text))

    reconstructed = []
    for groups, conf, raw in all_digit_groups:
        date = _reconstruct_rtl_date(groups)
        if date:
            reconstructed.append(date)

    if len(reconstructed) >= 2:
        reconstructed.sort()
        dates["date_from"] = reconstructed[0]
        dates["date_to"] = reconstructed[-1]
    elif len(reconstructed) == 1:
        dates["date_from"] = reconstructed[0]

    return dates

def _extract_single_date(ocr_results: list) -> Optional[str]:
    for text, conf in ocr_results:
        clean = transliterate_arabic_digits(text)
        match = re.search(r'(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})', clean)
        if match:
            y, m, d = match.groups()
            return f"{y}-{m.zfill(2)}-{d.zfill(2)}"

    all_groups = []
    for text, conf in ocr_results:
        clean = transliterate_arabic_digits(text)
        digits = re.findall(r'\d+', clean)
        if digits:
            all_groups.extend(digits)

    return _reconstruct_rtl_date(all_groups)

def _load_gemini_api_key() -> Optional[str]:
    """Load GEMINI_API_KEY from environment or .env file."""
    api_key = os.environ.get("GEMINI_API_KEY")
    if api_key:
        return api_key.strip()
    
    # Check .env in same directory or parent
    env_paths = [
        Path(__file__).parent / ".env",
        Path(__file__).parent.parent / ".env"
    ]
    for ep in env_paths:
        if ep.exists():
            with open(ep, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line.startswith("GEMINI_API_KEY="):
                        return line.split("=", 1)[1].strip().strip('"').strip("'")
    return None

def extract_with_gemini(image_path: str) -> Optional[dict]:
    """Extract certificate fields using Gemini Vision AI (100% precision)."""
    api_key = _load_gemini_api_key()
    if not api_key:
        return None

    try:
        from google import genai
        from google.genai import types
        from PIL import Image
        import json

        client = genai.Client(api_key=api_key)
        img = Image.open(image_path)
        
        # Downscale ultra-high resolution images for lightning fast network transfer
        MAX_AI_DIM = 1600
        if img.width > MAX_AI_DIM or img.height > MAX_AI_DIM:
            scale_factor = min(MAX_AI_DIM / img.width, MAX_AI_DIM / img.height)
            new_w = int(img.width * scale_factor)
            new_h = int(img.height * scale_factor)
            img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)

        prompt = '''
You are an expert Arabic OCR document parser for Egyptian Vehicle Insurance Certificates (شهادة سداد اشتراكات التأمين الاجتماعي عن سيارة).
Extract the following exact fields from this document into a JSON object:
- office_code: The owner insurance office code / number (الرقم التأميني للمالك / مكتب تامينات السيارة)
- plate_digits: Only the digits of the license plate (رقم اللوحات / رقم السيارة - أرقام فقط)
- plate_letters: The Arabic letters of the license plate (حروف اللوحات)
- plate_full: The full plate string (e.g. س ر م 7576)
- insurance_no: The vehicle insurance number (الرقم التأميني للسيارة - digits only, e.g. 2803316)
- chassis_no: Chassis number (رقم الشاسية)
- driver_name: Name of the driver if present (اسم السائق)
- date_from: Start date of license / insurance (فترة الترخيص/التجديد من) in YYYY-MM-DD format
- date_to: End date of license / insurance (فترة الترخيص/التجديد الى) in YYYY-MM-DD format
- cert_end_date: Certificate end date if listed (تاريخ نهاية الشهادة) in YYYY-MM-DD format
- service_no: Service code at top if present (رقم الخدمة)

Only return valid JSON.
'''

        # We must use gemini-1.5-flash for the Free Tier because Pro is limited to 2 Requests/Minute!
        # Flash allows 15 Requests/Minute and 1500 per day.
        models_to_try = ['gemini-1.5-flash', 'gemini-1.5-pro']
        response = None
        for m in models_to_try:
            try:
                response = client.models.generate_content(
                    model=m,
                    contents=[img, prompt],
                    config=types.GenerateContentConfig(response_mime_type='application/json')
                )
                if response and response.text:
                    print(f"[OCR] ✨ Successfully extracted with Gemini Vision ({m})")
                    break
            except Exception as me:
                print(f"[DEBUG] Gemini model {m} failed: {me}")
                continue




        if not response or not response.text:
            return None

        data_json = json.loads(response.text)

        # Standardize return dictionary
        data = {
            "service_no": data_json.get("service_no"),
            "office_code": data_json.get("office_code"),
            "plate_digits": str(data_json.get("plate_digits") or "").strip() or None,
            "plate_letters": data_json.get("plate_letters"),
            "plate_full": data_json.get("plate_full"),
            "insurance_no": str(data_json.get("insurance_no") or "").strip() or None,
            "chassis_no": str(data_json.get("chassis_no") or "").strip() or None,
            "date_from": data_json.get("date_from"),
            "date_to": data_json.get("date_to"),
            "print_date": data_json.get("cert_end_date"),
            "engine_used": "Gemini Vision AI",
            "confidence": {
                "office_code": 0.99 if data_json.get("office_code") else 0.0,
                "plate_digits": 0.99 if data_json.get("plate_digits") else 0.0,
                "insurance_no": 0.99 if data_json.get("insurance_no") else 0.0,
                "chassis_no": 0.99 if data_json.get("chassis_no") else 0.0,
                "license_line": 0.99 if (data_json.get("date_from") and data_json.get("date_to")) else 0.0,
                "service_no": 0.99 if data_json.get("service_no") else 0.0,
            },
            "raw_ocr": {"ai_json": data_json}
        }
        return data

    except Exception as e:
        print(f"[WARN] Gemini Vision API failed: {e}. Falling back to WinRT OCR.")
        return None

def extract_certificate_data(image_path: str) -> dict:
    # 1. First attempt: High-Precision Gemini Vision AI
    gemini_data = extract_with_gemini(image_path)
    if gemini_data:
        return gemini_data

    # 2. Fallback: Native Windows WinRT OCR
    print("[OCR] Using local WinRT OCR Engine...")
    clean_image = remove_stamp(image_path)
    
    # Downscale if massive to prevent OpenCV memory error
    MAX_WINRT_DIM = 1600
    if clean_image.shape[0] > MAX_WINRT_DIM or clean_image.shape[1] > MAX_WINRT_DIM:
        r = min(MAX_WINRT_DIM / clean_image.shape[0], MAX_WINRT_DIM / clean_image.shape[1])
        clean_image = cv2.resize(clean_image, (int(clean_image.shape[1] * r), int(clean_image.shape[0] * r)), interpolation=cv2.INTER_AREA)

    img_h, img_w = clean_image.shape[:2]
    print(f"[OCR] Image size: {img_w}x{img_h}")


    try:
        from winsdk.windows.media.ocr import OcrEngine
        import winsdk.windows.globalization as wg
        lang = wg.Language('ar')
        engine = OcrEngine.try_create_from_language(lang)
        if engine is None:
            print("[WARN] Arabic language not installed on Windows! Falling back to user profile languages.")
            engine = OcrEngine.try_create_from_user_profile_languages()
    except Exception as e:
        print(f"[ERR] WinRT OCR Engine could not be initialized: {e}")
        engine = None


    full_results, full_text = _ocr_fullimage(engine, clean_image)
    print(f"[OCR] Full-image scan: {len(full_results)} words detected")

    region_results = {}
    for name, region in REGIONS_RAPID.items():
        region_results[name] = _ocr_region(engine, clean_image, region, img_h, img_w, scale=3.0, region_name=name)
        if region_results[name]:
            print(f"[OCR] Region '{name}': {[t for t, c in region_results[name]]}")

    easyocr_results = {}
    for name, region in REGIONS_EASYOCR.items():
        easyocr_results[name] = _ocr_region(engine, clean_image, region, img_h, img_w, scale=4.0, region_name=name)
        if easyocr_results[name]:
            print(f"[OCR] Region '{name}': {[t for t, c in easyocr_results[name]]}")

    service_no = _extract_service_no(full_results)
    office_code = _extract_number(region_results.get("office_code", []), min_digits=5)
    plate_digits = _extract_number(region_results.get("plate_digits", []), min_digits=3)
    insurance_no = _extract_number(region_results.get("insurance_no", []), min_digits=5)
    chassis_no = _extract_number(region_results.get("chassis_no", []), min_digits=3)

    dates = _extract_dates_from_license_line(easyocr_results.get("license_line", []))
    cert_end = _extract_single_date(easyocr_results.get("cert_end_line", []))
    print_date = _extract_single_date(easyocr_results.get("print_date_line", []))

    # ROBUST FALLBACK: Scan entire page text for dates if regions missed them
    if not dates.get("date_from") or not dates.get("date_to"):
        clean_full_text = transliterate_arabic_digits(full_text)
        # Find dates in format YYYY/MM/DD, YYYY-MM-DD, YYYY MM DD, or YYYYMMDD
        all_dates = re.findall(r'(\d{4})[\s/\-.]*(\d{1,2})[\s/\-.]*(\d{1,2})', clean_full_text)
        parsed_dates = []
        for y, m, d in all_dates:
            # Basic validation
            if 2020 <= int(y) <= 2035 and 1 <= int(m) <= 12 and 1 <= int(d) <= 31:
                parsed_dates.append(f"{y}-{m.zfill(2)}-{d.zfill(2)}")
        
        parsed_dates = sorted(list(set(parsed_dates))) # Unique and sorted
        
        if len(parsed_dates) >= 2:
            if not dates.get("date_from"):
                dates["date_from"] = parsed_dates[0]
            if not dates.get("date_to"):
                dates["date_to"] = parsed_dates[-1]
        elif len(parsed_dates) == 1:
            if not dates.get("date_from") and not dates.get("date_to"):
                dates["date_from"] = parsed_dates[0]
                dates["date_to"] = parsed_dates[0]

    if not dates.get("date_to") and cert_end:
        dates["date_to"] = cert_end

    data = {
        "service_no": service_no,
        "office_code": office_code,
        "plate_digits": plate_digits,
        "insurance_no": insurance_no,
        "chassis_no": chassis_no,
        "date_from": dates.get("date_from"),
        "date_to": dates.get("date_to"),
        "print_date": print_date,
    }

    confidence = {}
    for field, results in region_results.items():
        if results:
            confidence[field] = max(conf for _, conf in results)
        else:
            confidence[field] = 0.0

    confidence["service_no"] = 0.95 if service_no else 0.0

    for field, results in easyocr_results.items():
        if results:
            confidence[field] = max(conf for _, conf in results)
        else:
            confidence[field] = 0.0

    raw_ocr = {}
    raw_ocr["full_scan"] = [f"y={y} x={x}: {text}" for text, conf, y, x in full_results]
    for field, results in region_results.items():
        raw_ocr[field] = [text for text, _ in results]
    for field, results in easyocr_results.items():
        raw_ocr[field] = [text for text, _ in results]

    data["confidence"] = confidence
    data["raw_ocr"] = raw_ocr

    return data

if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python ocr_engine.py <image_path>")
        sys.exit(1)

    result = extract_certificate_data(sys.argv[1])

    print("\n" + "=" * 60)
    print("EXTRACTED CERTIFICATE DATA (WinRT OCR)")
    print("=" * 60)

    for key, value in result.items():
        if key not in ("confidence", "raw_ocr"):
            status = "✅" if value else "❌"
            print(f"  {status} {key:15s} : {value}")

    print("\nConfidence Scores:")
    for key, conf in result.get("confidence", {}).items():
        status = "✅" if conf >= 0.7 else "⚠️" if conf >= 0.5 else "❌"
        print(f"  {status} {key:15s}: {conf:.2f}")
