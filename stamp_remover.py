"""
stamp_remover.py — OpenCV HSV Blue Stamp Masking

Detects and removes official blue/purple government seal ink from
scanned Social Insurance certificates. Produces a clean, high-contrast
black-on-white image optimized for OCR extraction.

Usage:
    from stamp_remover import remove_stamp, preprocess_for_ocr
    clean = remove_stamp("certificate.jpeg")
    binary = preprocess_for_ocr(clean)
"""

import cv2
import numpy as np
from pathlib import Path


def remove_stamp(image_path: str) -> np.ndarray:
    """
    Remove blue/purple official stamp ink from a scanned certificate image.

    Strategy:
        1. Convert to HSV color space
        2. Create mask for blue/purple hue range (stamp ink color)
        3. Dilate mask to cover stamp edges
        4. Inpaint masked regions with white background

    Args:
        image_path: Path to the scanned certificate image (JPEG/PNG).

    Returns:
        np.ndarray: BGR image with blue stamp regions replaced by white.
    """
    img = cv2.imread(str(image_path))
    if img is None:
        raise FileNotFoundError(f"Cannot read image: {image_path}")

    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)

    # Blue stamp ink typically falls in these HSV ranges:
    # Hue: 90-140 (blue to blue-purple)
    # Saturation: 40+ (colored, not gray)
    # Value: 50-255 (not too dark)
    lower_blue = np.array([90, 40, 50])
    upper_blue = np.array([140, 255, 255])
    mask_blue = cv2.inRange(hsv, lower_blue, upper_blue)

    # Also catch purple/violet stamp ink
    lower_purple = np.array([120, 30, 50])
    upper_purple = np.array([160, 255, 255])
    mask_purple = cv2.inRange(hsv, lower_purple, upper_purple)

    # Combine masks
    stamp_mask = cv2.bitwise_or(mask_blue, mask_purple)

    # Dilate to cover stamp edges and bleed-through
    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (5, 5))
    stamp_mask = cv2.dilate(stamp_mask, kernel, iterations=2)

    # Replace stamped regions with white
    result = img.copy()
    result[stamp_mask > 0] = [255, 255, 255]

    return result


def preprocess_for_ocr(image: np.ndarray) -> np.ndarray:
    """
    Convert a clean (stamp-removed) image to high-contrast binary
    for optimal OCR accuracy.

    Steps:
        1. Convert to grayscale
        2. Apply adaptive thresholding (handles uneven lighting)
        3. Light morphological closing to connect broken Arabic characters

    Args:
        image: BGR image (typically output of remove_stamp).

    Returns:
        np.ndarray: Binary (black text on white background) image.
    """
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    # Adaptive threshold handles uneven scan lighting
    binary = cv2.adaptiveThreshold(
        gray, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        blockSize=31,
        C=10
    )

    # Light morphological closing to reconnect Arabic letter fragments
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
    binary = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, kernel)

    return binary


def crop_region(image: np.ndarray, y_start: int, y_end: int,
                x_start: int, x_end: int, scale: float = 2.0) -> np.ndarray:
    cropped = image[y_start:y_end, x_start:x_end]

    if scale != 1.0:
        h, w = cropped.shape[:2]
        new_w = int(w * scale)
        new_h = int(h * scale)
        
        # WinRT OCR has limits (~4000px max dimension). Cap to 3000.
        MAX_DIM = 3000
        if new_w > MAX_DIM or new_h > MAX_DIM:
            ratio = min(MAX_DIM / new_w, MAX_DIM / new_h)
            new_w = int(new_w * ratio)
            new_h = int(new_h * ratio)
            
        cropped = cv2.resize(cropped, (max(1, new_w), max(1, new_h)),
                             interpolation=cv2.INTER_CUBIC)

    return cropped


def save_debug_images(image_path: str, output_dir: str = "debug_output"):
    """
    Save intermediate processing steps for visual debugging.

    Creates:
        - original.png
        - stamp_removed.png
        - binary.png
        - cropped regions for each bounding box

    Args:
        image_path: Path to the original scanned certificate.
        output_dir: Directory to save debug images.
    """
    out = Path(output_dir)
    out.mkdir(exist_ok=True)

    img = cv2.imread(str(image_path))
    if img is None:
        raise FileNotFoundError(f"Cannot read image: {image_path}")

    cv2.imwrite(str(out / "original.png"), img)

    clean = remove_stamp(image_path)
    cv2.imwrite(str(out / "stamp_removed.png"), clean)

    binary = preprocess_for_ocr(clean)
    cv2.imwrite(str(out / "binary.png"), binary)

    # Crop known bounding box regions (calibrated on 960×1280 images)
    regions = {
        "service_no":   (195, 235, 200, 960),
        "office_code":  (365, 405, 80, 450),
        "plate_digits": (480, 520, 80, 420),
        "insurance_no": (498, 538, 420, 850),
        "chassis_no":   (533, 573, 420, 850),
        "date_from":    (715, 755, 80, 500),
        "date_to":      (715, 755, 500, 850),
        "print_date":   (763, 803, 200, 850),
    }

    h, w = clean.shape[:2]
    for name, (y1, y2, x1, x2) in regions.items():
        # Scale bounding boxes if image is not exactly 960×1280
        sy = h / 1280
        sx = w / 960
        ry1, ry2 = int(y1 * sy), int(y2 * sy)
        rx1, rx2 = int(x1 * sx), int(x2 * sx)

        cropped = crop_region(clean, ry1, ry2, rx1, rx2, scale=3.0)
        cv2.imwrite(str(out / f"region_{name}.png"), cropped)

    print(f"Debug images saved to: {out.resolve()}")


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python stamp_remover.py <image_path>")
        sys.exit(1)

    save_debug_images(sys.argv[1])
