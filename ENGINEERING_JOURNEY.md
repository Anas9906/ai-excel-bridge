# Engineering Journey: AI Excel Bridge

This document is a transparent record of the problem-solving process, architectural decisions, and iterative improvements made throughout this project. It is written for other developers and hiring managers to understand not just *what* was built, but *why* each decision was made.

---

## The Problem: 45 Minutes of Manual Data Entry Every Morning

The HR department at a construction company receives a batch of physical vehicle insurance certificates from their insurance provider every week. Each certificate is a photographed document containing ~10 critical data points (dates, vehicle plate numbers, chassis numbers, insurance codes) that must be manually copied into a master Excel spreadsheet with thousands of rows and complex ERP-linked formulas.

This was taking **45+ minutes per batch** and was prone to human transcription errors that could cause payroll and insurance compliance issues.

**Objective:** Build a tool that can process a batch of 15 images and update the master Excel file in under 60 seconds, with zero formula corruption.

---

## Iteration 1: The MVP (Day 1) — Naive openpyxl Approach

**Hypothesis:** We can read the master Excel file with Python's openpyxl library and write the new values directly.

**What happened:** We built a working prototype that could parse the images and write to the Excel file. However, when the file was opened, all ERP-linked formulas (used for payroll calculations) had been **silently wiped**. The openpyxl library, when saving the file, strips out any Excel features it does not understand — including complex macros, named ranges, and cross-sheet formula dependencies.

**The Failure:** The tool was technically writing the right data, but producing an unusable file.

**Lesson Learned:** For production ERP systems, you cannot use file-level manipulation. You need a native integration that talks *to* Excel itself, not the file on disk.

---

## Iteration 2: The Pivot (Day 2) — Native Windows COM Interop

**The Insight:** Windows provides a technology called **COM (Component Object Model)** that allows any program to control a running Office application directly, just like a macro would — but from Python.

**The Solution:** We replaced the entire Excel-writing module with win32com.client, which drives the actual installed Microsoft Excel application silently in the background. This means:

- Excel itself opens the file, applies changes, and saves it.
- No formulas, macros, or formatting are ever touched by our code.
- The result is **byte-for-byte identical** to a user manually typing the values.

**Technical Implementation:**
`python
import win32com.client
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False  # Run silently in background
wb = excel.Workbooks.Open(filepath)
ws = wb.ActiveSheet
ws.Cells(row, col).Value = new_value
wb.Save()
`

**Result:** Zero formula corruption. The ERP system remained fully functional.

---

## Iteration 3: The Bottleneck (Day 3) — Employee Feedback Loop

**User Testing:** After delivering the working v1.0 to the employee, real-world testing revealed a critical performance bottleneck. Processing a batch of 15 certificates was taking **3-5 minutes** depending on network conditions.

**Root Cause Analysis:**
1. The batch worker was processing images **sequentially** — one image would fully complete before the next one started.
2. Each Gemini API call takes approximately 5-15 seconds of network round-trip time.
3. For 15 images: 15 images * 10 seconds avg = 150 seconds (2.5 minutes).

**The Lesson:** A real-world user will never accept a 3-minute wait for a task they could do in 2 minutes manually. Performance is a feature.

---

## Iteration 4: The Concurrency Solution (Day 4) — ThreadPoolExecutor

**The Solution:** The Gemini Vision API calls are **I/O-bound** (the program is waiting for the network, not computing). This is the exact use case for Python's concurrent.futures.ThreadPoolExecutor.

We split the batch worker into two distinct phases:
- **Phase 1 (Parallel):** Fire all 15 images to the Gemini API simultaneously using a 10-thread pool. All network waiting happens concurrently.
- **Phase 2 (Sequential):** Once all results are back, match them against the Excel data one by one (safe for COM objects, which are not thread-safe).

`python
with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
    futures = {executor.submit(extract_certificate_data, item["path"]): idx
               for idx, item in enumerate(self.batch_items)}
`

**Result:** A batch of 15 images now completes in **5-15 seconds** — a **10x speed improvement**.

**Model Selection Note:** We also discovered that the Google AI gemini-1.5-pro model (highest accuracy) is rate-limited to only 2 requests/minute on the Free Tier. Firing 15 simultaneous requests would cause immediate failures. gemini-1.5-flash allows 15 requests/minute, perfectly matching our concurrent batch size, and remains **100% free**.

---

## Iteration 5: UX Polish (Day 5) — Receipt Date Global Sync

**User Feedback:** The employee noted that the "Receipt Date" field (تاريخ الاستلام) is a **batch-level** attribute — it represents the date the company physically received the documents from the insurance provider, which is the same date for all 20 documents in a given delivery.

**The Problem:** Each image was loading with a separate receipt date field, requiring the employee to update it 15 times per batch.

**The Fix:** We changed the behavior so that when the eceipt_date field is modified on any image, the change is immediately propagated to the in-memory store of every item in the batch:

`python
# Global sync for Receipt Date across the entire batch
new_receipt = self.receipt_date_var.get().strip() or None
for b_item in self.batch_items:
    b_item["receipt_date"] = new_receipt
`

**Result:** The employee enters the receipt date exactly once per batch delivery, regardless of batch size.

---

## Key Technical Decisions Summary

| Decision | Why |
|---|---|
| win32com over openpyxl | COM preserves ERP formulas; openpyxl silently destroys them |
| gemini-1.5-flash over pro | Flash has a 15 req/min free tier limit; Pro has only 2 req/min |
| ThreadPoolExecutor | API calls are I/O-bound; concurrency gives a 10x speed boost |
| PyInstaller --onedir | Faster startup than --onefile; easier for IT to deploy and update |
| Receipt Date as global batch field | Real-world workflow: all documents in one delivery share the same receipt date |
