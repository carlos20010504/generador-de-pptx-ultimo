"""Content-hash persistent cache for IA-generated plans."""
import hashlib
import json
from pathlib import Path
from typing import Optional

DEFAULT_CACHE_DIR = Path(__file__).resolve().parent.parent / ".cache" / "plans"

def compute_cache_key(file_bytes: bytes, prompt: str, audience: str,
                       language: str, planner_version: str) -> str:
    h = hashlib.sha256()
    h.update(file_bytes)
    h.update(b"\x00")
    h.update(prompt.strip().encode("utf-8"))
    h.update(b"\x00")
    h.update(audience.strip().encode("utf-8"))
    h.update(b"\x00")
    h.update(language.strip().encode("utf-8"))
    h.update(b"\x00")
    h.update(planner_version.encode("utf-8"))
    return h.hexdigest()

class PlanCache:
    def __init__(self, cache_dir: Optional[Path] = None):
        self.cache_dir = Path(cache_dir) if cache_dir else DEFAULT_CACHE_DIR
        self.cache_dir.mkdir(parents=True, exist_ok=True)

    def _path(self, key: str) -> Path:
        return self.cache_dir / f"{key}.json"

    def get(self, key: str) -> Optional[dict]:
        path = self._path(key)
        if not path.exists():
            return None
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return None

    def set(self, key: str, value: dict) -> None:
        path = self._path(key)
        try:
            path.write_text(json.dumps(value, ensure_ascii=False, indent=2),
                            encoding="utf-8")
        except OSError:
            pass  # cache failures are non-fatal
