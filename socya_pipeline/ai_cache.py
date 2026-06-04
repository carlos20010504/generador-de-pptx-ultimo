"""Content-hash persistent cache for IA-generated plans."""
import hashlib
import json
import os
import tempfile
from pathlib import Path
from typing import Optional


def _default_cache_dir() -> Path:
    """Ubicación del cache de planes.

    IMPORTANTE: vive FUERA del árbol del proyecto. Antes era
    `<project_root>/.cache/plans`, lo que metía escrituras JSON dentro del
    directorio que el dev server de Next observa. Aunque el watcher de webpack
    no reacciona a archivos fuera del module graph, mantenerlo afuera elimina
    de raíz cualquier riesgo de que un file-watcher (Turbopack, una versión
    futura, un editor) dispare un reload del dev server justo mientras la IA
    planifica — el síntoma de "el programa se cierra durante el plan".

    Override explícito vía env SOCYA_CACHE_DIR. Default: subcarpeta estable en
    el temp del sistema (persiste entre corridas; el SO la limpia eventualmente,
    lo cual es aceptable para un cache).
    """
    override = (os.environ.get("SOCYA_CACHE_DIR") or "").strip()
    if override:
        return Path(override)
    return Path(tempfile.gettempdir()) / "socya_plan_cache"


DEFAULT_CACHE_DIR = _default_cache_dir()

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
