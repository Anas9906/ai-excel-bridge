"""
cache_manager.py — SHA-256 Image Hash Cache for OCR Results

Avoids duplicate Gemini API calls for images that have already been processed.
Results are stored in a local JSON cache file alongside the workbook.

Usage:
    from cache_manager import ImageCacheManager
    cache = ImageCacheManager()
    result = cache.get(image_path)  # None on miss
    if result is None:
        result = ocr_provider.extract(image_path)
        cache.set(image_path, result)
"""

import hashlib
import json
import logging
import os
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

# Default cache file location — sits next to the application
_DEFAULT_CACHE_PATH = Path(__file__).parent / ".ocr_cache.json"


class ImageCacheManager:
    """
    Persistent, file-backed cache keyed by SHA-256 hash of image bytes.

    Thread-safety note: reads are safe concurrently; writes are protected
    by a per-instance lock but the JSON file is not safe for multiple
    processes writing simultaneously. For this single-user desktop app
    that constraint is acceptable.
    """

    def __init__(self, cache_path: Optional[str] = None):
        import threading
        self._lock = threading.Lock()
        self.cache_path = Path(cache_path) if cache_path else _DEFAULT_CACHE_PATH
        self._cache: dict = {}
        self._hits = 0
        self._misses = 0
        self._load()

    # ── Public API ──────────────────────────────────────────────────────────

    def get(self, image_path: str) -> Optional[dict]:
        """Return cached OCR result for image_path, or None on miss."""
        key = self._hash(image_path)
        if key is None:
            return None
        with self._lock:
            result = self._cache.get(key)
        if result is not None:
            self._hits += 1
            logger.info("[CACHE HIT]  hash=%s  file=%s", key[:8], os.path.basename(image_path))
            return result
        self._misses += 1
        logger.debug("[CACHE MISS] hash=%s  file=%s", key[:8], os.path.basename(image_path))
        return None

    def set(self, image_path: str, data: dict) -> None:
        """Store OCR result under the SHA-256 hash of image_path."""
        key = self._hash(image_path)
        if key is None:
            return
        with self._lock:
            self._cache[key] = data
            self._persist()

    def get_stats(self) -> dict:
        """Return cache statistics for logging / UI display."""
        total = self._hits + self._misses
        return {
            "hits": self._hits,
            "misses": self._misses,
            "total": total,
            "hit_rate": round(self._hits / total, 3) if total else 0.0,
            "cached_entries": len(self._cache),
            "cache_path": str(self.cache_path),
        }

    def clear(self) -> None:
        """Wipe the entire cache (used for testing or forced re-scan)."""
        with self._lock:
            self._cache.clear()
            self._persist()
        logger.warning("[CACHE] Cache cleared — all entries removed.")

    # ── Internal helpers ────────────────────────────────────────────────────

    @staticmethod
    def _hash(image_path: str) -> Optional[str]:
        """Compute SHA-256 of raw image bytes. Returns None if file unreadable."""
        try:
            with open(image_path, "rb") as f:
                return hashlib.sha256(f.read()).hexdigest()
        except (OSError, IOError) as e:
            logger.warning("[CACHE] Cannot hash file %s: %s", image_path, e)
            return None

    def _load(self) -> None:
        """Load cache from disk on startup."""
        if self.cache_path.exists():
            try:
                with open(self.cache_path, "r", encoding="utf-8") as f:
                    self._cache = json.load(f)
                logger.info("[CACHE] Loaded %d cached entries from %s",
                            len(self._cache), self.cache_path)
            except (json.JSONDecodeError, OSError) as e:
                logger.warning("[CACHE] Failed to load cache file (%s) — starting fresh.", e)
                self._cache = {}
        else:
            logger.info("[CACHE] No existing cache file — starting fresh at %s", self.cache_path)

    def _persist(self) -> None:
        """Write current cache state to disk. Must be called inside self._lock."""
        try:
            with open(self.cache_path, "w", encoding="utf-8") as f:
                json.dump(self._cache, f, ensure_ascii=False, indent=2)
        except OSError as e:
            logger.error("[CACHE] Failed to persist cache to %s: %s", self.cache_path, e)
