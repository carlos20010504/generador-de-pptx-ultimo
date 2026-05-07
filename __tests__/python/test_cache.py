# __tests__/python/test_cache.py
import json
import hashlib
from pathlib import Path
import pytest
from socya_pipeline.ai_cache import PlanCache, compute_cache_key

PLANNER_VERSION = "v1"

def test_compute_cache_key_deterministic(tmp_path):
    file_bytes = b"hello"
    key_a = compute_cache_key(file_bytes, "prompt", "ejecutivos", "es", PLANNER_VERSION)
    key_b = compute_cache_key(file_bytes, "prompt", "ejecutivos", "es", PLANNER_VERSION)
    assert key_a == key_b
    assert len(key_a) == 64  # sha256 hex

def test_cache_miss_then_hit(tmp_path):
    cache = PlanCache(cache_dir=tmp_path)
    key = "abc123"
    assert cache.get(key) is None
    cache.set(key, {"slides": [{"type": "title"}]})
    got = cache.get(key)
    assert got == {"slides": [{"type": "title"}]}

def test_cache_invalidation_by_planner_version(tmp_path):
    bytes_ = b"x"
    k1 = compute_cache_key(bytes_, "p", "a", "es", "v1")
    k2 = compute_cache_key(bytes_, "p", "a", "es", "v2")
    assert k1 != k2
