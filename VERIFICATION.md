# Pre-Production Validation Checklist

This document tracks the **two hard gates** that must pass before treating the
`~120–160 docs/day` throughput number as a production commitment.
"19/19 unit tests passing" and "actually works on the desk" are different bars.

---

## Gate 1 — Batch=3 Field Attribution Accuracy

**Status: ⏳ Not yet verified**

**Why it matters:** Sending 3 certificate images in a single Gemini call and
asking for a JSON array back is efficient, but vision models can cross-contaminate
fields between images — e.g., writing Image 2's chassis number into Image 1's row.
This would produce silent ERP corruption that is harder to detect than a crash.

### How to run the verification

1. Collect 15–20 real certificates whose values you know are correct (from a
   previous manual run or the master Excel sheet).

2. Create a `ground_truth.csv` with this format:
   ```csv
   filename,chassis_no,insurance_no,plate_digits,date_from,date_to
   cert_001.jpg,34271,2521672,2938,2026-04-02,2027-04-01
   cert_002.jpg,...
   ```

3. Run the verification script (requires your `.env` with `GEMINI_API_KEY`):
   ```bash
   python verify_batch_accuracy.py \
     --images  path/to/sample_certs/ \
     --ground  ground_truth.csv \
     --out     accuracy_report.html
   ```

4. Open `accuracy_report.html` in a browser.

### Acceptance criteria

| Field | Required accuracy |
|---|---|
| `chassis_no` | ≥ 95% |
| `insurance_no` | ≥ 95% |
| `plate_digits` | ≥ 95% |
| `date_from` | ≥ 95% |
| `date_to` | ≥ 95% |

- If all fields pass → set `GEMINI_BATCH_SIZE=6` and repeat to verify higher throughput.
- If any field fails → investigate misattribution patterns in the HTML diff before going live.

> [!CAUTION]
> Do NOT set `GEMINI_BATCH_SIZE` above 3 until this step passes. Every
> mis-attributed field goes directly into the ERP master sheet.

---

## Gate 2 — Real Windows/Excel COM Validation

**Status: ⏳ Not yet verified (all testing done on Mac with mocked win32com)**

**Why it matters:** Mocks cannot catch real COM edge cases:
- Does `wb.Close()` behave as expected when Excel throws mid-write?
- Does killing Excel during a batch correctly trigger the guarded `finally` cleanup?
- Does the timestamped backup survive a hard crash?

### Steps

1. On a **Windows laptop with Microsoft Excel installed**, run:
   ```bash
   git clone https://github.com/Anas9906/ai-excel-bridge.git
   cd ai-excel-bridge
   pip install -r requirements.txt
   copy .env.example .env
   # Add your GEMINI_API_KEY to .env
   ```

2. Load a **copy** (not the production file) of the master workbook.

3. Run a small batch (3–5 images) in **dry-run mode** first to confirm OCR matches.

4. Run with **write mode** on a copy to confirm:
   - Yellow audit highlighting appears on the correct rows.
   - Formulas in surrounding cells are 100% intact.
   - Timestamped backup file is created in the same folder.

5. **Deliberately force a failure**: mid-batch, kill the Excel process via Task
   Manager. Verify that:
   - Python does not crash or hang indefinitely.
   - The `finally` block cleans up without `AttributeError`.
   - No orphaned `EXCEL.EXE` process remains in Task Manager.

### Acceptance criteria

- [ ] Dry-run completes without errors
- [ ] Write-run produces correct yellow highlighting
- [ ] No formula corruption in the master sheet
- [ ] Timestamped backup present
- [ ] Forced Excel kill does NOT leave orphaned process or Python crash

---

## Gate 3 — Partial Match Review Process (Operational)

**Status: 📋 Process documentation needed**

The application flags partial chassis matches with `⚠️ Review Required` in the
inspector panel. This is only useful if the operator actually checks these rows
before closing the app.

**Recommended desk procedure:**
1. After batch extraction, scan the queue for any `⚠️ Review Required` items.
2. Click each flagged item and verify the matched Excel row visually.
3. If the partial match is wrong — clear the item and manually locate the correct row.
4. Only write items you have reviewed.

---

## Token Security Note

The GitHub push in the initial setup used a token embedded directly in the
remote URL (`https://Anas9906:{TOKEN}@github.com/...`). The URL was reset
immediately after, but **rotate the GitHub token** as a precaution if it was
used in a shared or recorded environment.

**Going forward**, use `gh` CLI's built-in credential helper instead:
```bash
gh auth login   # one-time setup
git push origin master   # no manual token handling needed
```
