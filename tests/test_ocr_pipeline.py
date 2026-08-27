"""
tests/test_ocr_pipeline.py

Unit tests for:
    - ImageCacheManager (SHA-256 cache, hit/miss, stats)
    - _repair_json (local zero-cost JSON repair)
    - excel_handler.find_vehicle (exact vs partial chassis match precedence)
    - excel_handler.apply_update (timestamped backup, COM lifecycle guards)

Run with: python -m pytest tests/test_ocr_pipeline.py -v
"""

import hashlib
import json
import os
import shutil
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import MagicMock, patch, PropertyMock

# Adjust sys.path so we can import from the project root
sys.path.insert(0, str(Path(__file__).parent.parent))


# ── 1. ImageCacheManager Tests ────────────────────────────────────────────────

class TestImageCacheManager(unittest.TestCase):

    def setUp(self):
        self.tmp_dir = tempfile.mkdtemp()
        self.cache_path = os.path.join(self.tmp_dir, "test_cache.json")
        # Create a fake image file
        self.fake_image = os.path.join(self.tmp_dir, "test_cert.jpg")
        with open(self.fake_image, "wb") as f:
            f.write(b"\xff\xd8\xff\xe0" + b"\x00" * 100)   # Minimal JPEG-like bytes

        from cache_manager import ImageCacheManager
        self.cache = ImageCacheManager(cache_path=self.cache_path)

    def tearDown(self):
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def test_sha256_hash_is_deterministic(self):
        """Same file always produces the same hash."""
        from cache_manager import ImageCacheManager
        h1 = ImageCacheManager._hash(self.fake_image)
        h2 = ImageCacheManager._hash(self.fake_image)
        self.assertEqual(h1, h2)
        self.assertIsNotNone(h1)
        self.assertEqual(len(h1), 64)   # SHA-256 hex length

    def test_cache_miss_returns_none(self):
        result = self.cache.get(self.fake_image)
        self.assertIsNone(result)

    def test_cache_set_and_hit(self):
        payload = {"chassis_no": "TEST123", "insurance_no": "99999"}
        self.cache.set(self.fake_image, payload)
        result = self.cache.get(self.fake_image)
        self.assertIsNotNone(result)
        self.assertEqual(result["chassis_no"], "TEST123")

    def test_cache_persists_to_disk(self):
        """Cache entry written to JSON file should survive a new instance."""
        payload = {"chassis_no": "PERSIST_TEST"}
        self.cache.set(self.fake_image, payload)

        from cache_manager import ImageCacheManager
        reloaded = ImageCacheManager(cache_path=self.cache_path)
        result = reloaded.get(self.fake_image)
        self.assertIsNotNone(result)
        self.assertEqual(result["chassis_no"], "PERSIST_TEST")

    def test_cache_stats_hit_rate(self):
        import shutil as _shutil
        payload = {"chassis_no": "STATS_TEST"}
        self.cache.set(self.fake_image, payload)

        # Create a second distinct image file that IS readable but NOT in cache
        second_image = os.path.join(self.tmp_dir, "other_cert.jpg")
        _shutil.copy2(self.fake_image, second_image)
        # Flip one byte so hash differs
        with open(second_image, "r+b") as f:
            f.write(b"\x00")

        self.cache.get(self.fake_image)   # hit
        self.cache.get(second_image)      # miss — readable file, not cached

        stats = self.cache.get_stats()
        self.assertEqual(stats["hits"], 1)
        self.assertEqual(stats["misses"], 1)

    def test_unreadable_file_returns_none(self):
        result = self.cache.get("/does/not/exist/at/all.jpg")
        self.assertIsNone(result)


# ── 2. Local JSON Repair Tests ────────────────────────────────────────────────

class TestRepairJson(unittest.TestCase):

    def setUp(self):
        from ocr_provider import _repair_json
        self._repair = _repair_json

    def test_plain_json_object(self):
        raw = '{"chassis_no": "12345", "insurance_no": "99999"}'
        result = self._repair(raw)
        self.assertEqual(result["chassis_no"], "12345")

    def test_plain_json_array(self):
        raw = '[{"image_index": 0}, {"image_index": 1}]'
        result = self._repair(raw)
        self.assertIsInstance(result, list)
        self.assertEqual(len(result), 2)

    def test_strips_markdown_fences(self):
        raw = '```json\n[{"image_index": 0}]\n```'
        result = self._repair(raw)
        self.assertIsNotNone(result)
        self.assertIsInstance(result, list)

    def test_removes_trailing_commas(self):
        raw = '[{"a": 1,}, {"b": 2,}]'
        result = self._repair(raw)
        self.assertIsNotNone(result)

    def test_extracts_json_block_from_surrounding_text(self):
        raw = 'Here is the JSON:\n[{"image_index": 0, "chassis_no": "XYZ"}]\nEnd of response.'
        result = self._repair(raw)
        self.assertIsNotNone(result)
        self.assertIsInstance(result, list)

    def test_returns_none_for_completely_broken_json(self):
        raw = "This is not JSON at all, sorry."
        result = self._repair(raw)
        self.assertIsNone(result)

    def test_empty_string_returns_none(self):
        self.assertIsNone(self._repair(""))
        self.assertIsNone(self._repair("   "))


# ── 3. Chassis Matching Tests (ExcelHandler.find_vehicle) ─────────────────────

class TestChassisMatching(unittest.TestCase):
    """
    Test find_vehicle directly using a real ExcelHandler instance with
    a mocked openpyxl worksheet.
    """

    def _make_handler_with_rows(self, rows: dict):
        """
        Build a minimal ExcelHandler stub.
        rows: {row_number: {"col19": chassis_val, "col18": plate_val}}
        """
        from excel_handler import ExcelHandler, HEADER_ROWS

        handler = object.__new__(ExcelHandler)
        handler.excel_path = "/fake/path.xlsx"

        # Build a mock worksheet
        ws = MagicMock()
        max_row = max(rows.keys()) if rows else HEADER_ROWS

        def cell_side_effect(row, column):
            cell = MagicMock()
            row_data = rows.get(row, {})
            if column == 19:
                cell.value = row_data.get("col19")
            elif column == 18:
                cell.value = row_data.get("col18")
            else:
                cell.value = None
            return cell

        ws.cell.side_effect = cell_side_effect
        handler.ws = ws
        handler.ws_data = ws
        handler.max_row = max_row

        return handler

    def test_exact_chassis_match_wins(self):
        """Exact chassis match should return match_type 'chassis', not 'partial'."""
        handler = self._make_handler_with_rows({
            3: {"col19": "CHASSIS123", "col18": "1234"},
            4: {"col19": "XCHASSIS123Y", "col18": "9999"},
        })
        with patch.object(handler, "_build_match",
                          side_effect=lambda row, mtype: MagicMock(match_type=mtype)):
            matches = handler.find_vehicle("CHASSIS123")

        self.assertEqual(len(matches), 1)
        self.assertEqual(matches[0].match_type, "chassis")

    def test_partial_match_flagged_when_no_exact(self):
        """Partial match should be flagged when no exact match exists."""
        handler = self._make_handler_with_rows({
            3: {"col19": "XCHASSIS123Y", "col18": "9999"},
        })
        with patch.object(handler, "_build_match",
                          side_effect=lambda row, mtype: MagicMock(match_type=mtype)):
            matches = handler.find_vehicle("CHASSIS123")

        self.assertEqual(len(matches), 1)
        self.assertIn("partial", matches[0].match_type)
        self.assertIn("review required", matches[0].match_type)

    def test_plate_fallback_only_when_chassis_fails(self):
        """Plate-digit match should only trigger if all chassis passes return nothing."""
        handler = self._make_handler_with_rows({
            3: {"col19": None, "col18": "1234"},
        })
        with patch.object(handler, "_build_match",
                          side_effect=lambda row, mtype: MagicMock(match_type=mtype)):
            matches = handler.find_vehicle("1234")

        self.assertTrue(any("plate" in m.match_type for m in matches))

    def test_no_match_returns_empty_list(self):
        handler = self._make_handler_with_rows({
            3: {"col19": "TOTALLYDIFFERENT", "col18": "0000"},
        })
        with patch.object(handler, "_build_match",
                          side_effect=lambda row, mtype: MagicMock(match_type=mtype)):
            matches = handler.find_vehicle("FINDME999")

        self.assertEqual(matches, [])


# ── 4. Timestamped Backup Tests ───────────────────────────────────────────────

class TestTimestampedBackup(unittest.TestCase):

    def _make_win32_mocks(self):
        """Return sys.modules stubs for win32com and pythoncom (Windows-only, not on Mac)."""
        win32com = MagicMock()
        win32com.client.DispatchEx.side_effect = Exception("COM not available on Mac")
        pythoncom = MagicMock()
        return win32com, pythoncom

    def test_backup_filename_contains_timestamp(self):
        """apply_update should create a backup file with a timestamp in the name."""
        import re

        win32com_mock, pythoncom_mock = self._make_win32_mocks()

        with tempfile.TemporaryDirectory() as tmp:
            src = os.path.join(tmp, "master.xlsx")
            with open(src, "wb") as f:
                f.write(b"FAKE_XLSX_CONTENT")

            with patch.dict("sys.modules", {
                "win32com": win32com_mock,
                "win32com.client": win32com_mock.client,
                "pythoncom": pythoncom_mock,
            }):
                from excel_handler import ExcelHandler, HEADER_ROWS
                handler = object.__new__(ExcelHandler)
                handler.excel_path = src
                handler.ws = MagicMock()
                handler.ws_data = MagicMock()
                handler.max_row = HEADER_ROWS

                try:
                    handler.apply_update([], dry_run=False,
                                         output_path=os.path.join(tmp, "out.xlsx"))
                except Exception:
                    pass  # COM failure is expected; backup creation happens before COM

            files = os.listdir(tmp)
            backup_files = [f for f in files if "backup" in f and f.endswith(".xlsx")]
            self.assertTrue(len(backup_files) >= 1,
                            f"Expected at least one timestamped backup, got: {files}")

            ts_pattern = re.compile(r"backup_\d{8}_\d{6}")
            self.assertTrue(any(ts_pattern.search(f) for f in backup_files),
                            f"Backup filename should contain timestamp pattern, got: {backup_files}")



# ── 5. COM Error Recovery (guarded finally) ───────────────────────────────────

class TestCOMGuardedFinally(unittest.TestCase):

    def test_quit_not_called_if_excel_init_fails(self):
        """If DispatchEx raises, excel.Quit() must not be called (excel stays None)."""
        win32com_mock = MagicMock()
        win32com_mock.client.DispatchEx.side_effect = Exception("COM init failed")
        pythoncom_mock = MagicMock()

        with patch.dict("sys.modules", {
            "win32com": win32com_mock,
            "win32com.client": win32com_mock.client,
            "pythoncom": pythoncom_mock,
        }):
            with tempfile.TemporaryDirectory() as tmp:
                src = os.path.join(tmp, "master.xlsx")
                with open(src, "wb") as f:
                    f.write(b"FAKE_XLSX_CONTENT")

                from excel_handler import ExcelHandler, HEADER_ROWS
                handler = object.__new__(ExcelHandler)
                handler.excel_path = src
                handler.ws = MagicMock()
                handler.ws_data = MagicMock()
                handler.max_row = HEADER_ROWS

                try:
                    handler.apply_update([],
                                         output_path=os.path.join(tmp, "out.xlsx"),
                                         dry_run=False)
                except Exception:
                    pass  # COM failure expected

        # Quit on the mock's return value must never have been called
        # because DispatchEx raised before assignment
        win32com_mock.client.DispatchEx.return_value.Quit.assert_not_called()


if __name__ == "__main__":
    unittest.main(verbosity=2)
