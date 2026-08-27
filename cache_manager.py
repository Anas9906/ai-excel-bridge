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
import time
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

_DEFAULT_CACHE_PATH = Path(__file__).parent / ".ocr_cache.json"

# Entries older than this many days are evicted at startup.
# Set to 0 to disable eviction. Override via environment variable.
CACHE_TTL_DAYS = int(os.environ.get("OCR_CACHE_TTL_DAYS", 90))


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
            entry = self._cache.get(key)
        if entry is not None:
            # Unwrap timestamped envelope
            result = entry.get("data") if isinstance(entry, dict) and "data" in entry else entry
            self._hits += 1
            logger.info("[CACHE HIT]  hash=%s  file=%s", key[:8], os.path.basename(image_path))
            return result
        self._misses += 1
        logger.debug("[CACHE MISS] hash=%s  file=%s", key[:8], os.path.basename(image_path))
        return None

    def set(self, image_path: str, data: dict) -> None:
        """Store OCR result under the SHA-256 hash of image_path, with timestamp."""
        key = self._hash(image_path)
        if key is None:
            return
        with self._lock:
            self._cache[key] = {"data": data, "ts": time.time()}
            self._persist()

    def get_stats(self) -> dict:
        """Return cache statistics for logging / UI display."""
        total = self._hits + self._misses
        size_bytes = self.cache_path.stat().st_size if self.cache_path.exists() else 0
        return {
            "hits": self._hits,
            "misses": self._misses,
            "total": total,
            "hit_rate": round(self._hits / total, 3) if total else 0.0,
            "cached_entries": len(self._cache),
            "cache_size_kb": round(size_bytes / 1024, 1),
            "cache_ttl_days": CACHE_TTL_DAYS,
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
        """Load cache from disk on startup and evict stale entries."""
        if self.cache_path.exists():
            try:
                with open(self.cache_path, "r", encoding="utf-8") as f:
                    raw = json.load(f)

                # Evict entries older than CACHE_TTL_DAYS (if TTL is enabled)
                if CACHE_TTL_DAYS > 0:
                    cutoff = time.time() - (CACHE_TTL_DAYS * 86400)
                    before = len(raw)
                    raw = {
                        k: v for k, v in raw.items()
                        if not (isinstance(v, dict) and "ts" in v and v["ts"] < cutoff)
                    }
                    evicted = before - len(raw)
                    if evicted:
                        logger.info("[CACHE] Evicted %d stale entries (>%d days old).",
                                    evicted, CACHE_TTL_DAYS)

                self._cache = raw
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
