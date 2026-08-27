#!/usr/bin/env python3
"""
verify_batch_accuracy.py — Batch=3 Accuracy Verification Script

Runs Gemini batch extraction (GEMINI_BATCH_SIZE=3) against a folder of
known-correct certificates, then diffs each extracted field against a
ground-truth CSV you provide.

Usage:
    python verify_batch_accuracy.py \\
        --images  path/to/cert_images/ \\
        --ground  path/to/ground_truth.csv \\
        --out     accuracy_report.html

Ground truth CSV format (one row per certificate image):
    filename,chassis_no,insurance_no,plate_digits,date_from,date_to

Output:
    - Console summary: field-level accuracy %, misattribution matrix
    - HTML report:     side-by-side diff per image (open in browser)

Exit code:
    0 if ALL field accuracy >= 95%
    1 if any field falls below 95% (treat as pipeline gate failure)
"""

import argparse
import csv
import json
import os
import sys
from pathlib import Path

# Resolve project root so imports work regardless of working directory
ROOT = Path(__file__).parent
sys.path.insert(0, str(ROOT))

from ocr_engine import extract_certificates_batch


# ── Fields we track accuracy for ─────────────────────────────────────────────
TRACKED_FIELDS = ["chassis_no", "insurance_no", "plate_digits", "date_from", "date_to"]
ACCURACY_GATE = 0.95   # 95% per field required to pass


def _norm(v) -> str:
    """Normalize a field value for comparison — strip whitespace, lowercase."""
    if v is None:
        return ""
    return str(v).strip().lower()


def load_ground_truth(csv_path: str) -> dict:
    """Load ground truth CSV → {filename: {field: value}}."""
    gt = {}
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            fname = os.path.basename(row["filename"].strip())
            gt[fname] = {k: v.strip() for k, v in row.items() if k != "filename"}
    return gt


def collect_images(image_dir: str, ground_truth: dict) -> list:
    """Return sorted list of image paths that have a ground truth entry."""
    valid_exts = {".jpg", ".jpeg", ".png", ".bmp", ".tiff"}
    all_images = sorted(
        p for p in Path(image_dir).iterdir()
        if p.suffix.lower() in valid_exts
    )
    missing_gt = [p.name for p in all_images if p.name not in ground_truth]
    if missing_gt:
        print(f"[WARN] {len(missing_gt)} images have no ground truth entry and will be skipped:")
        for n in missing_gt[:5]:
            print(f"       {n}")
        if len(missing_gt) > 5:
            print(f"       ...and {len(missing_gt) - 5} more")

    return [p for p in all_images if p.name in ground_truth]


def run_extraction(image_paths: list) -> dict:
    """Run extract_certificates_batch and return {filename: result_dict}."""
    print(f"\n[VERIFY] Extracting {len(image_paths)} images via batch OCR (BATCH_SIZE={os.environ.get('GEMINI_BATCH_SIZE', 3)})...")
    results = extract_certificates_batch([str(p) for p in image_paths])
    return {image_paths[i].name: results[i] for i in range(len(image_paths))}


def diff_results(image_paths, extracted, ground_truth):
    """
    Build per-image diff records and aggregate field accuracy stats.

    Returns:
        diffs: list of {filename, field, expected, got, match}
        stats: {field: {correct, total, accuracy}}
    """
    diffs = []
    stats = {f: {"correct": 0, "total": 0} for f in TRACKED_FIELDS}

    for img_path in image_paths:
        fname = img_path.name
        ext = extracted.get(fname, {})
        gt = ground_truth[fname]

        for field in TRACKED_FIELDS:
            expected = _norm(gt.get(field, ""))
            got = _norm(ext.get(field, ""))
            match = (expected == got) or (expected and got and expected in got)
            diffs.append({
                "filename": fname,
                "field": field,
                "expected": gt.get(field, ""),
                "got": ext.get(field, ""),
                "match": match,
            })
            stats[field]["total"] += 1
            if match:
                stats[field]["correct"] += 1

    for f in TRACKED_FIELDS:
        t = stats[f]["total"]
        stats[f]["accuracy"] = stats[f]["correct"] / t if t else 0.0

    return diffs, stats


def print_summary(stats: dict, image_count: int):
    """Print a console accuracy table."""
    print("\n" + "=" * 65)
    print(f"  ACCURACY REPORT  ({image_count} certificates, batch_size={os.environ.get('GEMINI_BATCH_SIZE', 3)})")
    print("=" * 65)
    print(f"  {'Field':<20}  {'Correct':>8}  {'Total':>6}  {'Accuracy':>9}  {'Status':>8}")
    print("  " + "─" * 60)

    all_pass = True
    for field in TRACKED_FIELDS:
        s = stats[field]
        acc = s["accuracy"]
        status = "✅ PASS" if acc >= ACCURACY_GATE else "❌ FAIL"
        if acc < ACCURACY_GATE:
            all_pass = False
        print(f"  {field:<20}  {s['correct']:>8}  {s['total']:>6}  {acc:>8.1%}  {status:>8}")

    print("=" * 65)
    if all_pass:
        print("  ✅ ALL FIELDS PASS — batch=3 is safe to use in production.")
        print(f"  ✅ You can try GEMINI_BATCH_SIZE=6 and re-run to verify higher throughput.")
    else:
        print("  ❌ SOME FIELDS BELOW 95% — DO NOT widen batch size yet.")
        print("  ❌ Investigate misattribution cases in the HTML report.")
    print("=" * 65)

    return all_pass


def write_html_report(diffs: list, stats: dict, out_path: str, image_count: int):
    """Write a side-by-side HTML diff report."""
    batch_size = os.environ.get("GEMINI_BATCH_SIZE", 3)

    rows = ""
    for d in diffs:
        bg = "#d4edda" if d["match"] else "#f8d7da"
        icon = "✅" if d["match"] else "❌"
        rows += (
            f"<tr style='background:{bg}'>"
            f"<td>{d['filename']}</td>"
            f"<td><code>{d['field']}</code></td>"
            f"<td>{d['expected'] or '<em>empty</em>'}</td>"
            f"<td>{d['got'] or '<em>empty</em>'}</td>"
            f"<td style='text-align:center'>{icon}</td>"
            f"</tr>\n"
        )

    summary_rows = ""
    for field in TRACKED_FIELDS:
        s = stats[field]
        acc = s["accuracy"]
        bg = "#d4edda" if acc >= ACCURACY_GATE else "#f8d7da"
        summary_rows += (
            f"<tr style='background:{bg}'>"
            f"<td><code>{field}</code></td>"
            f"<td>{s['correct']}/{s['total']}</td>"
            f"<td><strong>{acc:.1%}</strong></td>"
            f"<td>{'PASS ✅' if acc >= ACCURACY_GATE else 'FAIL ❌'}</td>"
            f"</tr>\n"
        )

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Batch=3 Accuracy Report</title>
<style>
  body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; margin: 2rem; }}
  h1 {{ color: #2c3e50; }}
  table {{ border-collapse: collapse; width: 100%; margin-bottom: 2rem; }}
  th, td {{ padding: 0.5rem 0.75rem; border: 1px solid #dee2e6; text-align: left; font-size: 0.9rem; }}
  th {{ background: #343a40; color: #fff; }}
  code {{ background: #f8f9fa; padding: 2px 5px; border-radius: 3px; }}
  .meta {{ color: #6c757d; margin-bottom: 1rem; }}
</style>
</head>
<body>
<h1>🧪 Batch OCR Accuracy Report</h1>
<p class="meta">Certificates: {image_count} &nbsp;|&nbsp; GEMINI_BATCH_SIZE: {batch_size} &nbsp;|&nbsp; Gate: {ACCURACY_GATE:.0%}</p>

<h2>Field-Level Summary</h2>
<table>
<tr><th>Field</th><th>Correct/Total</th><th>Accuracy</th><th>Status</th></tr>
{summary_rows}
</table>

<h2>Per-Image Diff</h2>
<table>
<tr><th>Filename</th><th>Field</th><th>Expected</th><th>Extracted</th><th>Match</th></tr>
{rows}
</table>
</body>
</html>"""

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"\n[VERIFY] HTML report written: {out_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Verify Gemini batch OCR accuracy against known-correct certificates."
    )
    parser.add_argument("--images", required=True, help="Directory of certificate images to test")
    parser.add_argument("--ground", required=True, help="Ground truth CSV file (see format in docstring)")
    parser.add_argument("--out", default="accuracy_report.html", help="Output HTML report path")
    parser.add_argument("--batch-size", type=int, default=3,
                        help="GEMINI_BATCH_SIZE to test (sets env var, default 3)")
    args = parser.parse_args()

    # Apply batch size override
    os.environ["GEMINI_BATCH_SIZE"] = str(args.batch_size)

    if not os.path.isdir(args.images):
        print(f"[ERROR] Images directory not found: {args.images}")
        sys.exit(1)
    if not os.path.isfile(args.ground):
        print(f"[ERROR] Ground truth CSV not found: {args.ground}")
        sys.exit(1)

    gt = load_ground_truth(args.ground)
    print(f"[VERIFY] Loaded {len(gt)} ground truth entries from {args.ground}")

    image_paths = collect_images(args.images, gt)
    if not image_paths:
        print("[ERROR] No testable images found (check filenames match CSV 'filename' column).")
        sys.exit(1)
    print(f"[VERIFY] Found {len(image_paths)} testable images.")

    extracted = run_extraction(image_paths)
    diffs, stats = diff_results(image_paths, extracted, gt)
    passed = print_summary(stats, len(image_paths))
    write_html_report(diffs, stats, args.out, len(image_paths))

    # Print worst misses for quick manual review
    misses = [d for d in diffs if not d["match"]]
    if misses:
        print(f"\n[VERIFY] {len(misses)} mismatches to investigate:")
        for d in misses[:10]:
            print(f"  ❌  {d['filename']} | {d['field']:15s} | expected='{d['expected']}' got='{d['got']}'")
        if len(misses) > 10:
            print(f"  ...and {len(misses) - 10} more (see HTML report)")

    sys.exit(0 if passed else 1)


if __name__ == "__main__":
    main()
