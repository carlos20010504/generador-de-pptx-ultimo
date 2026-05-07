# Hermes-Driven Pipeline — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Refactor Excel→PPTX pipeline so Hermes (with free-model fallback chain) acts as the single slide planner. Every datum in the PPT must be traceable to actual Excel cells. Eliminate generic templates. Quota-resistant via single-call architecture, content-hash cache, and 4-model fallback.

**Architecture:** 5-step pipeline (parser → inventory → AI planner → validator+extractor → renderer). AI is invoked exactly once per presentation. Pure-Python parser/validator/renderer with no template fallback logic. Frontend gets SSE-streamed progress and uniform error contracts.

**Tech Stack:** Next.js 16 (TypeScript), React 19, Python 3 (pandas, openpyxl, python-pptx, matplotlib, requests, rapidfuzz, pytest), OpenRouter API (Hermes 3 405B + 3 fallbacks), Node.js built-in test runner.

**Spec reference:** `docs/superpowers/specs/2026-05-07-hermes-driven-pptx-design.md`

---

## File Structure

### New Python modules (under repo root, alongside `organizer.py`)

| File | Responsibility |
|---|---|
| `socya_pipeline/__init__.py` | Package marker |
| `socya_pipeline/parser.py` | Excel parsing (no inference). Returns raw `WorkbookData`. |
| `socya_pipeline/inventory.py` | Builds `Block[]` catalog with provenance + quality_flags |
| `socya_pipeline/ai_cache.py` | sha256 content-hash cache for IA plans (`.cache/plans/`) |
| `socya_pipeline/ai_chain.py` | OpenRouter call with 4-model fallback, retry, profile-aware |
| `socya_pipeline/planner.py` | Builds payload, calls IA chain, parses plan JSON |
| `socya_pipeline/validator.py` | Validates plan against inventory, applies drop rules |
| `socya_pipeline/extractor.py` | Pulls real cell data per validated slide, attaches provenance |
| `socya_pipeline/renderer.py` | Pure execute renderer (uses python-pptx + matplotlib) |
| `socya_pipeline/errors.py` | Typed error codes used in stderr → API contract |
| `socya_pipeline/cli.py` | CLI entry: `python -m socya_pipeline {analyze|generate} <args>` |

### New TypeScript modules

| File | Responsibility |
|---|---|
| `utils/error-codes.ts` | Uniform error contract type + helper |
| `utils/ai-status.ts` | AI status badge formatting helpers |
| `utils/sse-stream.ts` | SSE response helper for Next.js 16 |
| `components/GenerationProgress.tsx` | SSE-driven progress UI (5 fases) |
| `components/AuditModal.tsx` | Audit details modal in plain Spanish |

### Modified files

| File | What changes |
|---|---|
| `organizer.py` | Becomes thin shim that calls `socya_pipeline.cli analyze` |
| `generate_template_presentation.py` | Becomes thin shim that calls `socya_pipeline.cli generate` |
| `app/api/excel-intelligence/route.ts` | Calls new pipeline with profile=`fast`, uses error contract |
| `app/api/advanced-generate/route.ts` | Calls new pipeline, returns plan |
| `app/api/generate-pptx/route.ts` | Converts to SSE stream, profile=`patient` |
| `components/AIControlPanel.tsx` | Dynamic AI status indicator, dynamic suggestions |
| `components/ExcelUploader.tsx` | Consumes SSE, retry button, audit modal trigger |
| `utils/excel-ai-panel.cjs` | Strip local fallback logic, keep only profile/processing helpers |
| `utils/server-runtime.ts` | New env vars for profiles, model chain config |

### New tests

| File | Coverage |
|---|---|
| `__tests__/python/conftest.py` | Pytest fixtures (sample workbooks, mocked AI responses) |
| `__tests__/python/test_parser.py` | Parser handles 7 fixtures |
| `__tests__/python/test_inventory.py` | Inventory produces expected blocks with provenance |
| `__tests__/python/test_validator.py` | Drop rules trigger correctly |
| `__tests__/python/test_provenance.py` | Bullet provenance enforcement |
| `__tests__/python/test_renderer.py` | Validated plan renders to valid pptx |
| `__tests__/python/test_ai_chain.py` | Fallback chain on 429/timeout |
| `__tests__/python/test_cache.py` | Content-hash cache hit/miss |
| `__tests__/error-contract.test.js` | API responses use uniform error shape |
| `__tests__/sse-stream.test.js` | `/api/generate-pptx` emits expected SSE events |
| `__tests__/fixtures/` (folder) | 7 representative Excels |
| `scripts/generate-fixtures.py` | Generates fixtures programmatically |

### Files deleted in F5

- `app/api/ai-suggestions/route.ts` (suggestions now bundled in planner output)
- `utils/excel-ai-panel.cjs` legacy fallback functions (`buildSemanticSummary`, `inferExcelTopic`, `inferInformationType`, `buildKeyFindings`, `buildPatterns`, `buildTrends`, `buildPromptHints`, `inferPalette`)
- `organizer.py` legacy code: `build_local_executive_summary`, `build_local_briefing_payload`, all `enrich_*_with_ai` variants except the final unified one (which moves to `socya_pipeline/planner.py`)

---

## Phase 0 — Setup (must run before F1)

### Task 0.1: Commit baseline + create branch

**Files:**
- Modify: `.gitignore` (already done in 674ff15)
- New branch: `feature/hermes-driven-pipeline`

- [ ] **Step 1: Inspect pending changes one more time**

```bash
git status --short
git diff --stat HEAD
```

Expected: ~11 files modified totaling ~2000 lines (continuation of "strengthen hermes ppt pipeline" work).

- [ ] **Step 2: Commit pending changes as baseline on master**

```bash
git add GEMINI.md app/api/advanced-generate/route.ts app/api/excel-intelligence/route.ts app/api/generate-pptx/route.ts components/AIControlPanel.tsx components/ExcelUploader.tsx generate_template_presentation.py organizer.py utils/excel-ai-panel.cjs utils/pptx-helper.ts
git commit -m "$(cat <<'EOF'
chore: snapshot baseline before hermes-driven redesign

Continuation of strengthen-hermes / improve-data-fidelity work.
Captured as baseline for the upcoming pipeline redesign.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

Do NOT add the `.pyc` file (`.agents/skills/pptx/scripts/office/__pycache__/soffice.cpython-314.pyc`).

- [ ] **Step 3: Create feature branch**

```bash
git checkout -b feature/hermes-driven-pipeline
git status
```

Expected: `On branch feature/hermes-driven-pipeline`, working tree clean except `.pyc`.

- [ ] **Step 4: Snapshot current behavior with 3 fixtures**

```bash
mkdir -p __tests__/baseline-snapshots
cp "Comisiones V1.xlsx" __tests__/baseline-snapshots/baseline_comisiones.xlsx
cp test_ventas.xlsx __tests__/baseline-snapshots/baseline_ventas.xlsx
```

Run dev server and produce reference PPTs (manual):
```bash
npm run dev &
# wait ~10s for server
curl -X POST http://localhost:3001/api/generate-pptx \
  -F "file=@__tests__/baseline-snapshots/baseline_ventas.xlsx" \
  -F "userPrompt=Genera el PPT" \
  --output __tests__/baseline-snapshots/baseline_ventas.pptx
```

(Repeat for Comisiones.) Commit only the input xlsx (pptx is gitignored).

- [ ] **Step 5: Commit setup**

```bash
git add __tests__/baseline-snapshots/
git commit -m "test: add baseline xlsx snapshots for regression comparison

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task 0.2: Add Python test infrastructure

**Files:**
- Create: `requirements-dev.txt`
- Create: `__tests__/python/__init__.py`
- Create: `__tests__/python/conftest.py`
- Create: `pytest.ini`

- [ ] **Step 1: Write `requirements-dev.txt`**

```
pytest==8.3.4
pytest-mock==3.14.0
rapidfuzz==3.10.1
```

(`rapidfuzz` is needed at runtime in Phase 2; declaring here so dev install covers it.)

- [ ] **Step 2: Write `pytest.ini`**

```ini
[pytest]
testpaths = __tests__/python
python_files = test_*.py
python_classes = Test*
python_functions = test_*
addopts = -ra -q --strict-markers
```

- [ ] **Step 3: Write `__tests__/python/__init__.py`** (empty file).

- [ ] **Step 4: Write `__tests__/python/conftest.py` with shared fixtures**

```python
import pytest
import pandas as pd
from pathlib import Path

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures"

@pytest.fixture
def fixtures_dir():
    return FIXTURES_DIR

@pytest.fixture
def ventas_simple_path():
    return FIXTURES_DIR / "ventas_simple.xlsx"

@pytest.fixture
def casi_vacio_path():
    return FIXTURES_DIR / "casi_vacio.xlsx"

@pytest.fixture
def stub_planner_response():
    return {
        "presentation_meta": {"title": "Ventas 2024", "subtitle": "Resumen"},
        "slides": [
            {"type": "title", "title": "Ventas 2024", "subtitle": "Resumen"},
            {"type": "kpi_row", "title": "KPIs principales", "block_refs": ["K1"]},
            {"type": "chart", "chart_type": "bar", "title": "Por ciudad",
             "block_ref": "C1", "narrative": "Bogotá lidera con 60 ventas."},
        ],
        "prompt_suggestions": ["Compara Bogotá vs Medellín"],
    }
```

- [ ] **Step 5: Install and smoke-test pytest**

```bash
pip install -r requirements-dev.txt
pytest --collect-only
```

Expected: pytest collects 0 tests (no test files yet) but exits 0.

- [ ] **Step 6: Commit**

```bash
git add requirements-dev.txt pytest.ini __tests__/python/
git commit -m "test: add pytest infrastructure for python pipeline tests

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task 0.3: Create test fixture generator

**Files:**
- Create: `scripts/generate-fixtures.py`
- Create (via running script): `__tests__/fixtures/ventas_simple.xlsx`, `casi_vacio.xlsx`, `enorme.xlsx`, `corrupto.xlsx`, `dominio_raro.xlsx`

- [ ] **Step 1: Write `scripts/generate-fixtures.py`**

```python
"""Generate deterministic test fixtures for the Excel→PPTX pipeline."""
from pathlib import Path
import random
import pandas as pd
import numpy as np

OUT = Path(__file__).resolve().parent.parent / "__tests__" / "fixtures"
OUT.mkdir(parents=True, exist_ok=True)

random.seed(42)
np.random.seed(42)

def ventas_simple():
    months = pd.date_range("2024-01-01", periods=100, freq="D")
    cities = np.random.choice(["Bogotá", "Medellín", "Cali", "Cartagena", "Barranquilla"],
                               size=100, p=[0.5, 0.25, 0.15, 0.07, 0.03])
    totals = np.random.randint(50_000, 900_000, size=100)
    df = pd.DataFrame({
        "Mes": months,
        "Total": totals,
        "Ciudad": cities,
        "Vendedor": np.random.choice(["Ana","Luis","Marta","Pedro"], 100),
        "Producto": np.random.choice(["A","B","C"], 100),
    })
    df.to_excel(OUT / "ventas_simple.xlsx", sheet_name="Ventas", index=False)

def casi_vacio():
    df = pd.DataFrame({
        "Col1": [None, None, "X"],
        "Col2": [None, "Y", None],
        "Col3": [None, None, None],
    })
    df.to_excel(OUT / "casi_vacio.xlsx", sheet_name="Hoja1", index=False)

def enorme():
    n = 50_000
    df = pd.DataFrame({
        "Fecha": pd.date_range("2020-01-01", periods=n, freq="h"),
        "Valor": np.random.randn(n) * 1000 + 50_000,
        "Categoria": np.random.choice([f"C{i}" for i in range(20)], n),
    })
    df.to_excel(OUT / "enorme.xlsx", sheet_name="Datos", index=False)

def corrupto():
    (OUT / "corrupto.xlsx").write_bytes(b"PK\x03\x04 not a real xlsx file")

def dominio_raro():
    df = pd.DataFrame({
        "Empleado": [f"Emp_{i:03d}" for i in range(60)],
        "Departamento": np.random.choice(["RRHH","IT","Logistica","Salud","Educacion"], 60),
        "Antiguedad_anios": np.random.randint(1, 25, 60),
        "Salario": np.random.randint(1_500_000, 12_000_000, 60),
        "Genero": np.random.choice(["F","M","Otro"], 60, p=[0.45,0.5,0.05]),
    })
    df.to_excel(OUT / "dominio_raro.xlsx", sheet_name="Plantilla", index=False)

if __name__ == "__main__":
    ventas_simple()
    casi_vacio()
    enorme()
    corrupto()
    dominio_raro()
    print("Fixtures generated in", OUT)
```

- [ ] **Step 2: Run generator**

```bash
python scripts/generate-fixtures.py
ls __tests__/fixtures/
```

Expected output: 5 .xlsx files. (`comisiones_real.xlsx` and `auditoria_compleja.xlsx` will be added in Step 3 by copying.)

- [ ] **Step 3: Add the two real-data fixtures**

```bash
cp "Comisiones V1.xlsx" __tests__/fixtures/comisiones_real.xlsx
# auditoria_compleja: if you have one, copy. Else the AI test will skip this fixture.
```

- [ ] **Step 4: Whitelist the fixture xlsx in `.gitignore`**

Edit `.gitignore`, add after the `*.xlsx` line:

```
!__tests__/fixtures/*.xlsx
!__tests__/baseline-snapshots/*.xlsx
```

- [ ] **Step 5: Commit**

```bash
git add scripts/generate-fixtures.py .gitignore __tests__/fixtures/
git commit -m "test: add fixture generator and 5 representative excel fixtures

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

## Phase 1 — Anti-Quota Foundation

**Goal of F1:** Hermes (and 3 free fallbacks) actually get used. Cache works. Existing pipeline still runs and produces same PPTs as before, but now with real AI.

**Deployable after F1:** Yes — observable improvement in IA usage, no UI/output regression.

### Task 1.1: Create `socya_pipeline` package skeleton

**Files:**
- Create: `socya_pipeline/__init__.py`
- Create: `socya_pipeline/errors.py`

- [ ] **Step 1: Create `socya_pipeline/__init__.py`**

```python
"""Socya PPTX pipeline — Excel→PPTX with Hermes-driven planning."""
__version__ = "0.1.0"
```

- [ ] **Step 2: Create `socya_pipeline/errors.py`**

```python
"""Typed error codes used by the pipeline. Stable contract for the API layer."""
from enum import Enum

class ErrorCode(str, Enum):
    EXCEL_INVALID = "EXCEL_INVALID"
    EXCEL_EMPTY = "EXCEL_EMPTY"
    EXCEL_INSUFFICIENT_DATA = "EXCEL_INSUFFICIENT_DATA"
    AI_SATURATED = "AI_SATURATED"
    AI_RESPONSE_INVALID = "AI_RESPONSE_INVALID"
    PLANNER_REJECTED_PROMPT = "PLANNER_REJECTED_PROMPT"
    PYTHON_RUNTIME_ERROR = "PYTHON_RUNTIME_ERROR"
    TIMEOUT = "TIMEOUT"

class PipelineError(Exception):
    def __init__(self, code: ErrorCode, message: str, details: str = "",
                 user_action: str = "report_bug", retry_after_seconds: int = 0):
        self.code = code
        self.message = message
        self.details = details
        self.user_action = user_action
        self.retry_after_seconds = retry_after_seconds
        super().__init__(f"[{code.value}] {message}")

    def to_dict(self) -> dict:
        return {
            "code": self.code.value,
            "message": self.message,
            "details": self.details,
            "user_action": self.user_action,
            "retry_after_seconds": self.retry_after_seconds,
        }
```

- [ ] **Step 3: Commit**

```bash
git add socya_pipeline/
git commit -m "feat: scaffold socya_pipeline package with typed error codes

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task 1.2: Implement content-hash cache

**Files:**
- Create: `socya_pipeline/ai_cache.py`
- Test: `__tests__/python/test_cache.py`

- [ ] **Step 1: Write the failing test**

```python
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
```

- [ ] **Step 2: Run test, expect failure**

```bash
pytest __tests__/python/test_cache.py -v
```

Expected: ImportError on `socya_pipeline.ai_cache`.

- [ ] **Step 3: Implement `socya_pipeline/ai_cache.py`**

```python
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
```

- [ ] **Step 4: Run test, expect pass**

```bash
pytest __tests__/python/test_cache.py -v
```

Expected: 3 passed.

- [ ] **Step 5: Commit**

```bash
git add socya_pipeline/ai_cache.py __tests__/python/test_cache.py
git commit -m "feat(pipeline): add content-hash plan cache

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task 1.3: Implement model fallback chain

**Files:**
- Create: `socya_pipeline/ai_chain.py`
- Test: `__tests__/python/test_ai_chain.py`

- [ ] **Step 1: Write the failing test**

```python
# __tests__/python/test_ai_chain.py
import pytest
from unittest.mock import MagicMock, patch
from socya_pipeline.ai_chain import AIChain, AIProfile, AIChainResult, MODEL_CHAIN
from socya_pipeline.errors import PipelineError, ErrorCode

@pytest.fixture
def fast_chain():
    return AIChain(api_key="test-key", profile=AIProfile.FAST)

def _mock_response(status_code, json_body=None, headers=None):
    resp = MagicMock()
    resp.status_code = status_code
    resp.ok = 200 <= status_code < 300
    resp.text = ""
    resp.headers = headers or {}
    resp.json.return_value = json_body or {}
    return resp

def test_first_model_success(fast_chain):
    happy = _mock_response(200, {
        "choices": [{"message": {"content": '{"plan":"ok"}'}}]
    })
    with patch("socya_pipeline.ai_chain.requests.post", return_value=happy):
        result = fast_chain.call("prompt")
    assert result.model == MODEL_CHAIN[0]
    assert result.content == '{"plan":"ok"}'
    assert result.fallback_steps == []

def test_falls_back_on_429(fast_chain):
    rate_limited = _mock_response(429, {"error": {"message": "rate limit"}})
    happy = _mock_response(200, {"choices":[{"message":{"content":"ok"}}]})
    with patch("socya_pipeline.ai_chain.requests.post",
               side_effect=[rate_limited, happy]):
        result = fast_chain.call("prompt")
    assert result.model == MODEL_CHAIN[1]
    assert len(result.fallback_steps) == 1
    assert result.fallback_steps[0]["from"] == MODEL_CHAIN[0]
    assert result.fallback_steps[0]["reason"] == "rate_limited"

def test_all_saturated_raises_ai_saturated(fast_chain):
    rate_limited = _mock_response(429, {"error": {"message": "rate limit"}})
    with patch("socya_pipeline.ai_chain.requests.post", return_value=rate_limited):
        with pytest.raises(PipelineError) as exc:
            fast_chain.call("prompt")
    assert exc.value.code == ErrorCode.AI_SATURATED

def test_no_api_key_raises():
    chain = AIChain(api_key="", profile=AIProfile.FAST)
    with pytest.raises(PipelineError) as exc:
        chain.call("prompt")
    assert exc.value.code == ErrorCode.AI_SATURATED
```

- [ ] **Step 2: Run test, expect failure**

```bash
pytest __tests__/python/test_ai_chain.py -v
```

Expected: ImportError.

- [ ] **Step 3: Implement `socya_pipeline/ai_chain.py`**

```python
"""OpenRouter call layer with 4-model fallback chain and profile-aware retry."""
import os
import time
from dataclasses import dataclass, field
from enum import Enum
from typing import List, Optional
import requests
from socya_pipeline.errors import PipelineError, ErrorCode

MODEL_CHAIN = [
    "nousresearch/hermes-3-llama-3.1-405b:free",
    "meta-llama/llama-3.3-70b-instruct:free",
    "google/gemma-2-9b-it:free",
    "qwen/qwen-2.5-72b-instruct:free",
]

OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
APP_NAME = os.environ.get("OPENROUTER_APP_NAME", "Socya PPTX Generator")
SITE_URL = os.environ.get("OPENROUTER_SITE_URL", "http://localhost")

class AIProfile(str, Enum):
    FAST = "fast"      # 25s timeout, 1 model only, raise if it fails
    PATIENT = "patient"  # up to 4 min, full chain, multiple cycles

PROFILE_SETTINGS = {
    AIProfile.FAST: {
        "timeout_per_call": 25,
        "max_models_to_try": 1,
        "max_cycles": 1,
        "retry_within_model": 0,
    },
    AIProfile.PATIENT: {
        "timeout_per_call": 60,
        "max_models_to_try": len(MODEL_CHAIN),
        "max_cycles": 4,
        "retry_within_model": 1,
    },
}

RATE_LIMIT_TOKENS = ("rate limit", "high demand", "limit_rpm",
                     "429", "limited to", "temporarily rate-limited")
TRANSIENT_TOKENS = ("upstream error", "timed out", "timeout",
                    "service unavailable", "overloaded", "bad gateway")

@dataclass
class AIChainResult:
    model: str
    content: str
    fallback_steps: List[dict] = field(default_factory=list)
    cache_hit: bool = False

class AIChain:
    def __init__(self, api_key: str, profile: AIProfile = AIProfile.FAST):
        self.api_key = (api_key or "").strip().strip('"').strip("'")
        self.profile = profile
        self.settings = PROFILE_SETTINGS[profile]

    def call(self, prompt: str, system_msg: str = "You must output strictly valid JSON.",
             temperature: float = 0.2) -> AIChainResult:
        if not self.api_key:
            raise PipelineError(
                ErrorCode.AI_SATURATED,
                "Falta OPENROUTER_API_KEY en el entorno.",
                user_action="report_bug",
            )

        fallback_steps: List[dict] = []
        last_error = "unknown"
        models_to_try = MODEL_CHAIN[: self.settings["max_models_to_try"]]

        for cycle in range(self.settings["max_cycles"]):
            for model in models_to_try:
                try:
                    content = self._call_one(model, prompt, system_msg, temperature)
                    return AIChainResult(model=model, content=content,
                                         fallback_steps=fallback_steps)
                except _Retryable as exc:
                    fallback_steps.append({
                        "from": model, "reason": exc.reason,
                        "message": str(exc)[:200],
                    })
                    last_error = exc.reason
                    if exc.retry_after and self.profile == AIProfile.PATIENT:
                        time.sleep(min(exc.retry_after, 90))
                except _Fatal as exc:
                    raise PipelineError(
                        ErrorCode.AI_RESPONSE_INVALID,
                        f"El modelo {model} devolvió una respuesta inválida.",
                        details=str(exc)[:300],
                        user_action="retry",
                    )
            if cycle < self.settings["max_cycles"] - 1:
                time.sleep(min(30 * (cycle + 1), 60))

        raise PipelineError(
            ErrorCode.AI_SATURATED,
            "Todos los modelos IA disponibles están saturados ahora mismo.",
            details=f"Último error: {last_error}. Intentos: {len(fallback_steps)}.",
            user_action="retry_later",
            retry_after_seconds=300,
        )

    def _call_one(self, model: str, prompt: str, system_msg: str,
                   temperature: float) -> str:
        payload = {
            "model": model,
            "messages": [
                {"role": "system", "content": system_msg},
                {"role": "user", "content": prompt},
            ],
            "temperature": temperature,
        }
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
            "HTTP-Referer": SITE_URL,
            "X-Title": APP_NAME,
        }
        try:
            resp = requests.post(OPENROUTER_URL, headers=headers, json=payload,
                                  timeout=self.settings["timeout_per_call"])
        except requests.Timeout:
            raise _Retryable("timeout", "request timed out")
        except requests.RequestException as e:
            raise _Retryable("network_error", str(e))

        text = resp.text or ""
        lower = text.lower()

        if resp.status_code == 429:
            retry_after = _parse_retry_after(resp.headers, lower)
            raise _Retryable("rate_limited", text[:200], retry_after=retry_after)
        if not resp.ok:
            if any(t in lower for t in RATE_LIMIT_TOKENS):
                raise _Retryable("rate_limited", text[:200])
            if any(t in lower for t in TRANSIENT_TOKENS):
                raise _Retryable("transient", text[:200])
            raise _Fatal(f"http_{resp.status_code}: {text[:200]}")

        try:
            data = resp.json()
            content = data["choices"][0]["message"]["content"]
            if not content:
                raise _Fatal("empty content")
            return content
        except (KeyError, IndexError, ValueError) as e:
            raise _Fatal(f"malformed response: {e}")


def _parse_retry_after(headers, body_lower: str) -> int:
    h = headers.get("Retry-After") or headers.get("retry-after")
    if h:
        try:
            return int(float(h))
        except (TypeError, ValueError):
            pass
    import re
    m = re.search(r"retry in\s+([0-9]+(?:\.[0-9]+)?)\s*s", body_lower)
    if m:
        return int(float(m.group(1)))
    return 0


class _Retryable(Exception):
    def __init__(self, reason: str, message: str = "", retry_after: int = 0):
        self.reason = reason
        self.retry_after = retry_after
        super().__init__(message)

class _Fatal(Exception):
    pass
```

- [ ] **Step 4: Run tests, expect pass**

```bash
pytest __tests__/python/test_ai_chain.py -v
```

Expected: 4 passed.

- [ ] **Step 5: Commit**

```bash
git add socya_pipeline/ai_chain.py __tests__/python/test_ai_chain.py
git commit -m "feat(pipeline): add 4-model openrouter fallback chain with profiles

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task 1.4: Wire `ai_chain` + `ai_cache` into existing organizer

**Goal:** Make existing `organizer.py` use the new chain when calling Hermes, while keeping the existing pipeline structure. This proves F1 works end-to-end without breaking outputs.

**Files:**
- Modify: `organizer.py` (replace `call_ai_api` body)
- Modify: `app/api/excel-intelligence/route.ts` (remove `SOCYA_AI_HARD_DEADLINE_SECONDS=35`, set profile)
- Modify: `app/api/advanced-generate/route.ts` (same)
- Modify: `app/api/generate-pptx/route.ts` (set profile=patient)

- [ ] **Step 1: Read current `call_ai_api` in `organizer.py`**

Re-read lines 1920-2025 of `organizer.py` to capture exact current behavior so the replacement is drop-in compatible (same return type: string content or None).

- [ ] **Step 2: Replace `call_ai_api` to delegate to `AIChain`**

In `organizer.py`, replace the function body of `call_ai_api(prompt, response_mime_type="application/json")` with:

```python
def call_ai_api(prompt, response_mime_type="application/json"):
    """Compat shim: delegates to socya_pipeline.AIChain."""
    from socya_pipeline.ai_chain import AIChain, AIProfile
    from socya_pipeline.errors import PipelineError

    api_key = get_openrouter_api_key()
    if not api_key:
        return None

    profile_name = (os.environ.get("SOCYA_AI_PROFILE") or "fast").strip().lower()
    profile = AIProfile.PATIENT if profile_name == "patient" else AIProfile.FAST

    system_msg = ("You must output strictly valid JSON."
                  if response_mime_type == "application/json"
                  else "You are a helpful assistant.")
    temperature = 0.2 if response_mime_type == "application/json" else 0.4

    try:
        result = AIChain(api_key=api_key, profile=profile).call(
            prompt, system_msg=system_msg, temperature=temperature
        )
    except PipelineError as exc:
        print(f"INFO: AIChain falló: {exc.code.value} {exc.message}", file=sys.stderr)
        return None

    content = result.content
    if response_mime_type == "application/json":
        return normalize_ai_json_response_text(content)
    return content
```

Keep the helper `normalize_ai_json_response_text` as is.

- [ ] **Step 3: Remove the old AI deadline overrides in API routes**

In all three route files (`app/api/excel-intelligence/route.ts`, `app/api/advanced-generate/route.ts`, `app/api/generate-pptx/route.ts`), update the `env:` block in `execFileAsync`. Replace:

```typescript
env: {
  ...process.env,
  PYTHONUTF8: '1',
  SOCYA_AI_WAIT_ON_RATE_LIMIT: '0',
  SOCYA_AI_MAX_WAIT_SECONDS: '0',
  SOCYA_AI_WAIT_POLL_SECONDS: '15',
  SOCYA_AI_EXECUTION_MODE: 'best_effort',
  SOCYA_AI_HARD_DEADLINE_SECONDS: '35',
},
```

With (for `excel-intelligence` and `advanced-generate`):

```typescript
env: {
  ...process.env,
  PYTHONUTF8: '1',
  SOCYA_AI_PROFILE: 'fast',
},
```

And for `generate-pptx`:

```typescript
env: {
  ...process.env,
  PYTHONUTF8: '1',
  SOCYA_AI_PROFILE: 'patient',
  SOCYA_PRESENTATION_MODE: visualMode,
},
```

- [ ] **Step 4: Manual smoke test — confirm Hermes is now actually called**

```bash
npm run dev &
# wait
curl -X POST http://localhost:3001/api/excel-intelligence \
  -F "file=@__tests__/fixtures/ventas_simple.xlsx" \
  -F "userPrompt=Resume las ventas por ciudad" \
  -F "audience=ejecutivos" -F "language=Español" \
  -o /tmp/intel.json
cat /tmp/intel.json | python -m json.tool | head -50
```

Expected: response includes `executiveSummary` and `keyFindings` with **real content from the IA** (not the local fallback templates). Check stderr of the dev server: should see no `AI_SATURATED` errors if your `OPENROUTER_API_KEY` is valid.

- [ ] **Step 5: Verify existing PPT generation still works (regression check)**

```bash
curl -X POST http://localhost:3001/api/generate-pptx \
  -F "file=@__tests__/baseline-snapshots/baseline_ventas.xlsx" \
  -F "userPrompt=Genera el PPT" \
  --output /tmp/post_F1_ventas.pptx
ls -la /tmp/post_F1_ventas.pptx __tests__/baseline-snapshots/baseline_ventas.pptx
```

Expected: File exists, size in same order of magnitude as baseline. Open both in PowerPoint side-by-side and confirm structure is comparable.

- [ ] **Step 6: Commit**

```bash
git add organizer.py app/api/
git commit -m "feat(pipeline): wire new ai_chain into existing organizer

call_ai_api now delegates to socya_pipeline.AIChain with fast/patient
profiles. Old SOCYA_AI_* env vars removed in favor of SOCYA_AI_PROFILE.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task 1.5: Add cache integration in `organizer.py`'s unified AI flow

This wires the cache into the existing unified AI cache (`organizer_unified_ai_cache.json`). It uses the new content-hash key.

**Files:**
- Modify: `organizer.py` (find the unified cache helper, add content-hash version)

- [ ] **Step 1: Locate the unified cache helper**

Grep for `UNIFIED_AI_CACHE_FILE` in `organizer.py`. Read the surrounding `load_*` / `save_*` / `build_unified_cache_key` functions.

- [ ] **Step 2: Add content-hash key alongside existing key**

In `organizer.py`, near where `build_unified_cache_key` is defined, add:

```python
def build_content_hash_cache_key(file_path, user_prompt, audience, language,
                                   planner_version="v1"):
    from socya_pipeline.ai_cache import compute_cache_key
    try:
        with open(file_path, "rb") as f:
            file_bytes = f.read()
    except OSError:
        return None
    return compute_cache_key(file_bytes, user_prompt or "",
                              audience or "ejecutivos",
                              language or "es", planner_version)
```

This is preparation — actual usage happens in F2 when planner.py replaces the multi-call AI flow.

- [ ] **Step 3: Commit**

```bash
git add organizer.py
git commit -m "feat(pipeline): add content-hash cache key helper for upcoming planner

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task 1.6: F1 acceptance check

- [ ] **Step 1: Run all Python tests**

```bash
pytest __tests__/python/ -v
```

Expected: all tests pass.

- [ ] **Step 2: Run all Node tests**

```bash
npm test
```

Expected: existing tests still pass.

- [ ] **Step 3: Run lint**

```bash
npm run lint
```

Expected: no errors.

- [ ] **Step 4: Manual end-to-end with each fixture**

For each of `ventas_simple.xlsx`, `comisiones_real.xlsx`, `dominio_raro.xlsx`:
1. Upload via UI at `http://localhost:3001`.
2. Confirm AI panel shows real semantic summary (not template).
3. Click Generate, confirm PPT downloads.

- [ ] **Step 5: Tag F1 milestone**

```bash
git tag -a f1-anti-quota -m "F1 complete: anti-quota foundation"
```

---

## Phase 2 — Planner + Provenance

**Goal of F2:** Single-call AI plan with strict block references. Validator drops invalid slides. Pipeline runs in parallel to old one (feature-flagged).

**Deployable after F2:** Yes — new pipeline behind `SOCYA_USE_NEW_PIPELINE=1` env flag; old one is default.

### Task 2.1: Implement `parser.py`

**Files:**
- Create: `socya_pipeline/parser.py`
- Test: `__tests__/python/test_parser.py`

- [ ] **Step 1: Write the failing test**

```python
# __tests__/python/test_parser.py
import pytest
from pathlib import Path
from socya_pipeline.parser import parse_workbook, WorkbookData
from socya_pipeline.errors import PipelineError, ErrorCode

def test_parse_ventas_simple(ventas_simple_path):
    wb = parse_workbook(ventas_simple_path)
    assert isinstance(wb, WorkbookData)
    assert wb.filename == "ventas_simple.xlsx"
    assert "Ventas" in [s.name for s in wb.sheets]
    ventas = next(s for s in wb.sheets if s.name == "Ventas")
    assert ventas.shape == (100, 5)
    cols = {c.name for c in ventas.columns}
    assert {"Mes", "Total", "Ciudad", "Vendedor", "Producto"}.issubset(cols)

def test_parse_corrupto_raises(fixtures_dir):
    with pytest.raises(PipelineError) as exc:
        parse_workbook(fixtures_dir / "corrupto.xlsx")
    assert exc.value.code == ErrorCode.EXCEL_INVALID

def test_parse_casi_vacio_succeeds(casi_vacio_path):
    wb = parse_workbook(casi_vacio_path)
    assert wb.sheets
    sheet = wb.sheets[0]
    assert sheet.fill_ratio < 0.5  # mostly empty
```

- [ ] **Step 2: Run test, expect failure (ImportError)**

```bash
pytest __tests__/python/test_parser.py -v
```

- [ ] **Step 3: Implement `socya_pipeline/parser.py`**

```python
"""Pure-data Excel parser. No semantic inference."""
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Any
import pandas as pd
from socya_pipeline.errors import PipelineError, ErrorCode

@dataclass
class ColumnData:
    name: str
    dtype: str            # "numeric" | "currency" | "date" | "categorical" | "text" | "bool"
    n_unique: int
    fill_ratio: float
    samples: List[Any] = field(default_factory=list)
    min: Any = None
    max: Any = None
    mean: Any = None
    sum: Any = None
    top_values: List[List[Any]] = field(default_factory=list)  # [[value, count], ...]

@dataclass
class SheetData:
    name: str
    shape: tuple              # (rows, cols)
    fill_ratio: float
    columns: List[ColumnData]
    first_rows: List[List[Any]]   # up to 8 rows in compact form

@dataclass
class WorkbookData:
    filename: str
    sheets: List[SheetData]

def parse_workbook(path) -> WorkbookData:
    p = Path(path)
    try:
        xls = pd.ExcelFile(p, engine=None)  # auto-detect openpyxl/xlrd
    except Exception as e:
        raise PipelineError(
            ErrorCode.EXCEL_INVALID,
            "El archivo no se pudo leer como Excel.",
            details=str(e)[:200],
            user_action="upload_again",
        )

    sheets: List[SheetData] = []
    for sheet_name in xls.sheet_names:
        try:
            df = xls.parse(sheet_name)
        except Exception:
            continue
        sheets.append(_summarize_sheet(sheet_name, df))

    if not sheets:
        raise PipelineError(
            ErrorCode.EXCEL_EMPTY,
            "El Excel se abrió pero no tiene hojas legibles.",
            user_action="upload_again",
        )

    return WorkbookData(filename=p.name, sheets=sheets)


def _summarize_sheet(name: str, df: pd.DataFrame) -> SheetData:
    rows, cols = df.shape
    if rows == 0 or cols == 0:
        return SheetData(name=name, shape=(rows, cols), fill_ratio=0.0,
                         columns=[], first_rows=[])
    fill_ratio = float(df.notna().sum().sum()) / max(1, rows * cols)
    columns = [_summarize_column(c, df[c]) for c in df.columns]
    first_rows = df.head(8).where(df.head(8).notna(), None).values.tolist()
    return SheetData(name=name, shape=(rows, cols), fill_ratio=fill_ratio,
                     columns=columns, first_rows=first_rows)


def _summarize_column(name: Any, series: pd.Series) -> ColumnData:
    name_str = str(name)
    fill = float(series.notna().sum()) / max(1, len(series))
    n_unique = int(series.nunique(dropna=True))
    dtype = _infer_dtype(name_str, series)
    samples = series.dropna().head(8).tolist()

    col = ColumnData(name=name_str, dtype=dtype, n_unique=n_unique,
                     fill_ratio=fill, samples=_jsonify(samples))

    if dtype in ("numeric", "currency"):
        nums = pd.to_numeric(series, errors="coerce").dropna()
        if not nums.empty:
            col.min = float(nums.min())
            col.max = float(nums.max())
            col.mean = float(nums.mean())
            col.sum = float(nums.sum())
    elif dtype in ("categorical", "text"):
        vc = series.dropna().astype(str).value_counts().head(8)
        col.top_values = [[v, int(c)] for v, c in vc.items()]

    return col


def _infer_dtype(name: str, series: pd.Series) -> str:
    if pd.api.types.is_bool_dtype(series):
        return "bool"
    if pd.api.types.is_datetime64_any_dtype(series):
        return "date"
    if pd.api.types.is_numeric_dtype(series):
        lower = name.lower()
        if any(t in lower for t in ("total", "valor", "precio", "costo",
                                      "salario", "monto", "ingreso")):
            return "currency"
        return "numeric"
    n_unique = series.nunique(dropna=True)
    if n_unique > 0 and n_unique <= max(20, len(series) * 0.05):
        return "categorical"
    return "text"


def _jsonify(values):
    out = []
    for v in values:
        if pd.isna(v):
            out.append(None)
        elif hasattr(v, "isoformat"):
            out.append(v.isoformat())
        elif isinstance(v, (int, float, str, bool)):
            out.append(v)
        else:
            out.append(str(v))
    return out
```

- [ ] **Step 4: Run test, expect pass**

```bash
pytest __tests__/python/test_parser.py -v
```

Expected: 3 passed.

- [ ] **Step 5: Commit**

```bash
git add socya_pipeline/parser.py __tests__/python/test_parser.py
git commit -m "feat(pipeline): add pure-data excel parser

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task 2.2: Implement `inventory.py`

**Files:**
- Create: `socya_pipeline/inventory.py`
- Test: `__tests__/python/test_inventory.py`

- [ ] **Step 1: Write the failing test**

```python
# __tests__/python/test_inventory.py
from socya_pipeline.parser import parse_workbook
from socya_pipeline.inventory import build_inventory, Block

def test_inventory_basic(ventas_simple_path):
    wb = parse_workbook(ventas_simple_path)
    inv = build_inventory(wb)
    ids = {b.id for b in inv}
    kinds = {b.kind for b in inv}
    assert any(b.kind == "kpi_candidate" for b in inv)
    assert any(b.kind == "categorical_distribution" for b in inv)
    assert any(b.kind == "table" for b in inv)
    # IDs must be unique
    assert len(ids) == len(inv)

def test_inventory_provenance(ventas_simple_path):
    wb = parse_workbook(ventas_simple_path)
    inv = build_inventory(wb)
    for b in inv:
        assert b.provenance.sheet
        assert b.provenance.columns
        assert b.provenance.rows[1] >= b.provenance.rows[0]

def test_inventory_marks_quality_flags(casi_vacio_path):
    wb = parse_workbook(casi_vacio_path)
    inv = build_inventory(wb)
    for b in inv:
        if b.kind == "table":
            assert "low_fill_ratio" in b.quality_flags
```

- [ ] **Step 2: Run test, expect failure**

```bash
pytest __tests__/python/test_inventory.py -v
```

- [ ] **Step 3: Implement `socya_pipeline/inventory.py`**

```python
"""Catalog data blocks with provenance + quality flags."""
from dataclasses import dataclass, field
from typing import List, Set
from socya_pipeline.parser import WorkbookData, SheetData, ColumnData

@dataclass
class Provenance:
    sheet: str
    columns: List[str]
    rows: tuple              # (start, end), 0-indexed inclusive

@dataclass
class Block:
    id: str
    kind: str                # "table" | "kpi_candidate" | "categorical_distribution" | "time_series_candidate"
    label: str
    provenance: Provenance
    quality_flags: Set[str] = field(default_factory=set)
    extra: dict = field(default_factory=dict)

def build_inventory(wb: WorkbookData) -> List[Block]:
    blocks: List[Block] = []
    counter = {"T": 0, "K": 0, "C": 0, "S": 0}

    for sheet in wb.sheets:
        # 1. Whole-sheet table block
        counter["T"] += 1
        bid = f"T{counter['T']}"
        flags = set()
        if sheet.fill_ratio < 0.3:
            flags.add("low_fill_ratio")
        if sheet.shape[0] < 2:
            flags.add("too_few_rows")
        blocks.append(Block(
            id=bid, kind="table",
            label=f"{sheet.name} (tabla completa)",
            provenance=Provenance(sheet=sheet.name,
                                  columns=[c.name for c in sheet.columns],
                                  rows=(0, max(0, sheet.shape[0] - 1))),
            quality_flags=flags,
            extra={"shape": list(sheet.shape), "fill_ratio": sheet.fill_ratio},
        ))

        # 2. KPI candidates (numeric/currency aggregations)
        for col in sheet.columns:
            if col.dtype in ("numeric", "currency") and col.fill_ratio >= 0.5:
                counter["K"] += 1
                bid = f"K{counter['K']}"
                flags = set()
                if col.sum is not None and col.sum == 0:
                    flags.add("all_zero")
                blocks.append(Block(
                    id=bid, kind="kpi_candidate",
                    label=f"Suma de {col.name}",
                    provenance=Provenance(sheet=sheet.name, columns=[col.name],
                                          rows=(0, sheet.shape[0] - 1)),
                    quality_flags=flags,
                    extra={"agg": "sum", "value": col.sum,
                            "min": col.min, "max": col.max, "mean": col.mean},
                ))

        # 3. Categorical distributions
        for col in sheet.columns:
            if col.dtype == "categorical" and 2 <= col.n_unique <= 20:
                counter["C"] += 1
                bid = f"C{counter['C']}"
                flags = set()
                if col.top_values:
                    top_pct = col.top_values[0][1] / max(1, sum(c[1] for c in col.top_values))
                    if top_pct > 0.85:
                        flags.add("single_dominant_category")
                blocks.append(Block(
                    id=bid, kind="categorical_distribution",
                    label=f"Distribución por {col.name}",
                    provenance=Provenance(sheet=sheet.name, columns=[col.name],
                                          rows=(0, sheet.shape[0] - 1)),
                    quality_flags=flags,
                    extra={"top_values": col.top_values, "n_unique": col.n_unique},
                ))

        # 4. Time series candidates (date column + numeric column pair)
        date_cols = [c for c in sheet.columns if c.dtype == "date"]
        num_cols = [c for c in sheet.columns if c.dtype in ("numeric", "currency")]
        for d in date_cols[:1]:
            for n in num_cols[:2]:
                counter["S"] += 1
                bid = f"S{counter['S']}"
                blocks.append(Block(
                    id=bid, kind="time_series_candidate",
                    label=f"{n.name} a lo largo de {d.name}",
                    provenance=Provenance(sheet=sheet.name, columns=[d.name, n.name],
                                          rows=(0, sheet.shape[0] - 1)),
                    extra={"x": d.name, "y": n.name},
                ))

    return blocks
```

- [ ] **Step 4: Run test, expect pass**

```bash
pytest __tests__/python/test_inventory.py -v
```

Expected: 3 passed.

- [ ] **Step 5: Commit**

```bash
git add socya_pipeline/inventory.py __tests__/python/test_inventory.py
git commit -m "feat(pipeline): add data inventory with provenance and quality flags

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task 2.3: Implement `planner.py`

**Files:**
- Create: `socya_pipeline/planner.py`
- Test: `__tests__/python/test_planner.py`

- [ ] **Step 1: Write the failing test**

```python
# __tests__/python/test_planner.py
import json
import pytest
from unittest.mock import patch, MagicMock
from socya_pipeline.parser import parse_workbook
from socya_pipeline.inventory import build_inventory
from socya_pipeline.planner import build_payload, plan_presentation, PLANNER_VERSION

def test_build_payload_under_token_budget(ventas_simple_path):
    wb = parse_workbook(ventas_simple_path)
    blocks = build_inventory(wb)
    payload = build_payload(wb, blocks, user_prompt="resume", audience="ejecutivos",
                              language="es")
    serialized = json.dumps(payload, ensure_ascii=False)
    # Heuristic: ~4 chars per token. 6K tokens ≈ 24K chars.
    assert len(serialized) < 30_000

def test_payload_includes_block_ids(ventas_simple_path):
    wb = parse_workbook(ventas_simple_path)
    blocks = build_inventory(wb)
    payload = build_payload(wb, blocks, "p", "a", "es")
    payload_block_ids = {b["id"] for b in payload["workbook"]["blocks"]}
    inventory_ids = {b.id for b in blocks}
    assert payload_block_ids == inventory_ids

def test_plan_calls_ai_chain(ventas_simple_path, stub_planner_response):
    wb = parse_workbook(ventas_simple_path)
    blocks = build_inventory(wb)
    fake_result = MagicMock(model="hermes", content=json.dumps(stub_planner_response),
                             fallback_steps=[], cache_hit=False)
    with patch("socya_pipeline.planner.AIChain") as MockChain:
        MockChain.return_value.call.return_value = fake_result
        plan = plan_presentation(wb, blocks, user_prompt="p", audience="a",
                                  language="es", api_key="key")
    assert plan["slides"][0]["type"] == "title"
    assert plan["_meta"]["model"] == "hermes"

def test_plan_uses_cache_on_second_call(ventas_simple_path, stub_planner_response,
                                          tmp_path):
    wb = parse_workbook(ventas_simple_path)
    blocks = build_inventory(wb)
    fake_result = MagicMock(model="hermes", content=json.dumps(stub_planner_response),
                             fallback_steps=[], cache_hit=False)
    with patch("socya_pipeline.planner.AIChain") as MockChain:
        MockChain.return_value.call.return_value = fake_result
        # First call hits AI
        plan_presentation(wb, blocks, "p", "a", "es", "key", cache_dir=tmp_path,
                            file_path=ventas_simple_path)
        # Second call should hit cache
        plan2 = plan_presentation(wb, blocks, "p", "a", "es", "key", cache_dir=tmp_path,
                                    file_path=ventas_simple_path)
    assert plan2["_meta"]["cache_hit"] is True
    # Only one AI call across both invocations
    assert MockChain.return_value.call.call_count == 1
```

- [ ] **Step 2: Run, expect failure**

```bash
pytest __tests__/python/test_planner.py -v
```

- [ ] **Step 3: Implement `socya_pipeline/planner.py`**

```python
"""Single-call IA planner. Takes inventory + prompt, returns full slide plan."""
import json
import re
from pathlib import Path
from typing import Optional
from socya_pipeline.ai_cache import PlanCache, compute_cache_key
from socya_pipeline.ai_chain import AIChain, AIProfile
from socya_pipeline.errors import PipelineError, ErrorCode
from socya_pipeline.inventory import Block
from socya_pipeline.parser import WorkbookData

PLANNER_VERSION = "p1"  # bump when prompt template changes — invalidates cache

MAX_PAYLOAD_CHARS = 24_000  # rough 6K-token budget (~4 chars/token)
MAX_SAMPLES_PER_COL = 8
MAX_FIRST_ROWS = 8

def build_payload(wb: WorkbookData, blocks, user_prompt: str, audience: str,
                   language: str) -> dict:
    sheets_payload = []
    for s in wb.sheets:
        sheets_payload.append({
            "name": s.name,
            "shape": list(s.shape),
            "fill_ratio": round(s.fill_ratio, 3),
            "columns": [_col_payload(c) for c in s.columns],
            "first_rows": s.first_rows[:MAX_FIRST_ROWS],
        })
    blocks_payload = [_block_payload(b) for b in blocks]

    payload = {
        "user_prompt": user_prompt or "",
        "audience": audience or "ejecutivos",
        "language": language or "es",
        "workbook": {
            "filename": wb.filename,
            "sheets": sheets_payload,
            "blocks": blocks_payload,
        },
    }

    # Adaptive truncation if too large
    serialized = json.dumps(payload, ensure_ascii=False, default=str)
    if len(serialized) > MAX_PAYLOAD_CHARS:
        for s in payload["workbook"]["sheets"]:
            s["first_rows"] = s["first_rows"][:3]
            for c in s["columns"]:
                if "samples" in c:
                    c["samples"] = c["samples"][:3]
                if "top_values" in c:
                    c["top_values"] = c["top_values"][:3]
    return payload


def _col_payload(c) -> dict:
    out = {"name": c.name, "dtype": c.dtype,
           "n_unique": c.n_unique, "fill_ratio": round(c.fill_ratio, 3),
           "samples": c.samples[:MAX_SAMPLES_PER_COL]}
    for k in ("min", "max", "mean", "sum"):
        v = getattr(c, k, None)
        if v is not None:
            out[k] = v
    if c.top_values:
        out["top_values"] = c.top_values[:MAX_SAMPLES_PER_COL]
    return out


def _block_payload(b: Block) -> dict:
    return {"id": b.id, "kind": b.kind, "label": b.label,
            "provenance": {"sheet": b.provenance.sheet,
                            "columns": b.provenance.columns,
                            "rows": list(b.provenance.rows)},
            "quality_flags": sorted(b.quality_flags),
            **{k: v for k, v in b.extra.items() if k != "raw"}}


PROMPT_TEMPLATE = """Eres un planificador de presentaciones ejecutivas que genera slides a partir de datos reales de Excel.

Reglas estrictas:
1. Cada slide DEBE referenciar bloques existentes por su `id` (campo `block_ref` o `block_refs`).
2. NO inventes números, nombres ni hechos. Si necesitas un dato, viene del bloque referenciado.
3. Si el prompt del usuario pide algo que los bloques no soportan, omite ese slide.
4. Bullets deben ser específicos (con cifras o nombres del bloque). No bullets genéricos como "los datos muestran variabilidad".
5. Devuelve EXCLUSIVAMENTE un JSON válido con esta forma exacta:

{{
  "presentation_meta": {{ "title": "...", "subtitle": "..." }},
  "slides": [
    {{ "type": "title", "title": "...", "subtitle": "..." }},
    {{ "type": "kpi_row", "title": "...", "block_refs": ["K1", "K2"] }},
    {{ "type": "chart", "chart_type": "bar|line|pie", "title": "...",
       "block_ref": "C1", "narrative": "una frase basada en el bloque" }},
    {{ "type": "table", "title": "...", "block_ref": "T1",
       "columns_subset": ["..."], "max_rows": 12 }},
    {{ "type": "text_bullets", "title": "...", "supports_block": "T1",
       "bullets": ["bullet con cifra", "bullet con nombre"] }}
  ],
  "prompt_suggestions": ["sugerencia 1", "sugerencia 2", "sugerencia 3"]
}}

Datos disponibles:
{payload_json}
"""

def plan_presentation(wb: WorkbookData, blocks, user_prompt: str, audience: str,
                       language: str, api_key: str,
                       profile: AIProfile = AIProfile.PATIENT,
                       cache_dir: Optional[Path] = None,
                       file_path: Optional[Path] = None) -> dict:
    cache_key = None
    cache: Optional[PlanCache] = None
    if file_path is not None:
        try:
            file_bytes = Path(file_path).read_bytes()
            cache_key = compute_cache_key(file_bytes, user_prompt, audience,
                                            language, PLANNER_VERSION)
            cache = PlanCache(cache_dir=cache_dir)
            cached = cache.get(cache_key)
            if cached:
                cached["_meta"] = {**cached.get("_meta", {}),
                                    "cache_hit": True}
                return cached
        except OSError:
            pass

    payload = build_payload(wb, blocks, user_prompt, audience, language)
    payload_json = json.dumps(payload, ensure_ascii=False, default=str)
    prompt = PROMPT_TEMPLATE.format(payload_json=payload_json)

    chain = AIChain(api_key=api_key, profile=profile)
    result = chain.call(prompt)

    parsed = _parse_json_loose(result.content)
    if not isinstance(parsed, dict) or "slides" not in parsed:
        raise PipelineError(
            ErrorCode.AI_RESPONSE_INVALID,
            "El planificador devolvió un JSON sin la forma esperada.",
            details=str(result.content)[:300],
            user_action="retry",
        )

    parsed["_meta"] = {
        "model": result.model,
        "cache_hit": False,
        "fallback_steps": result.fallback_steps,
        "planner_version": PLANNER_VERSION,
    }

    if cache and cache_key:
        cache.set(cache_key, parsed)

    return parsed


def _parse_json_loose(text: str):
    raw = (text or "").strip()
    if not raw:
        return None
    raw = raw.replace("```json", "").replace("```", "").strip()
    m = re.search(r"\{.*\}", raw, re.DOTALL)
    if m:
        raw = m.group(0)
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return None
```

- [ ] **Step 4: Run tests, expect pass**

```bash
pytest __tests__/python/test_planner.py -v
```

Expected: 4 passed.

- [ ] **Step 5: Commit**

```bash
git add socya_pipeline/planner.py __tests__/python/test_planner.py
git commit -m "feat(pipeline): add single-call planner with cache + payload truncation

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task 2.4: Implement `validator.py`

**Files:**
- Create: `socya_pipeline/validator.py`
- Test: `__tests__/python/test_validator.py`

- [ ] **Step 1: Write the failing test**

```python
# __tests__/python/test_validator.py
from socya_pipeline.parser import parse_workbook
from socya_pipeline.inventory import build_inventory
from socya_pipeline.validator import validate_plan, ValidationOutcome

def _plan_with(slide_list):
    return {"presentation_meta": {"title": "T", "subtitle": "S"},
            "slides": slide_list,
            "prompt_suggestions": []}

def test_drops_slides_with_unknown_block_ref(ventas_simple_path):
    wb = parse_workbook(ventas_simple_path)
    inv = build_inventory(wb)
    plan = _plan_with([
        {"type": "chart", "chart_type": "bar", "title": "x",
         "block_ref": "NONEXISTENT", "narrative": "n"},
    ])
    outcome = validate_plan(plan, inv, wb)
    assert outcome.slides == []
    assert outcome.dropped[0]["reason"] == "block_ref_not_found"

def test_keeps_valid_chart(ventas_simple_path):
    wb = parse_workbook(ventas_simple_path)
    inv = build_inventory(wb)
    cat_block = next(b for b in inv if b.kind == "categorical_distribution")
    plan = _plan_with([
        {"type": "chart", "chart_type": "bar", "title": "x",
         "block_ref": cat_block.id, "narrative": "n"},
    ])
    outcome = validate_plan(plan, inv, wb)
    assert len(outcome.slides) == 1

def test_drops_chart_when_block_flagged_low_quality(ventas_simple_path):
    wb = parse_workbook(ventas_simple_path)
    inv = build_inventory(wb)
    # Force a single_dominant_category flag on first cat block
    cat_block = next(b for b in inv if b.kind == "categorical_distribution")
    cat_block.quality_flags.add("single_dominant_category")
    plan = _plan_with([
        {"type": "chart", "chart_type": "pie", "title": "x",
         "block_ref": cat_block.id, "narrative": "n"},
    ])
    outcome = validate_plan(plan, inv, wb)
    assert outcome.slides == []
    assert outcome.dropped[0]["reason"] == "single_dominant_category"

def test_returns_at_least_one_slide_or_raises(ventas_simple_path):
    wb = parse_workbook(ventas_simple_path)
    inv = build_inventory(wb)
    plan = _plan_with([])  # empty plan
    outcome = validate_plan(plan, inv, wb)
    assert outcome.slides == []
    # Caller decides whether to raise EXCEL_INSUFFICIENT_DATA
```

- [ ] **Step 2: Run test, expect failure**

```bash
pytest __tests__/python/test_validator.py -v
```

- [ ] **Step 3: Implement `socya_pipeline/validator.py`**

```python
"""Strict validator: drops slides without traceable provenance or low quality."""
from dataclasses import dataclass, field
from typing import List, Dict
import re
from socya_pipeline.parser import WorkbookData

UGLY_LITERALS = {"nan", "none", "null", "nat", "???", "—", "s/d", "n/a", "na"}
DROP_REASONS = {
    "block_ref_not_found",
    "single_dominant_category",
    "all_zero_values",
    "too_few_categories",
    "too_few_rows",
    "low_fill_ratio",
    "all_bullets_failed_provenance",
    "missing_required_field",
    "no_columns_after_filter",
    "no_rows_after_filter",
}

@dataclass
class ValidationOutcome:
    slides: List[dict] = field(default_factory=list)
    dropped: List[dict] = field(default_factory=list)   # {type, reason, block_ref?}
    bullets_dropped: int = 0


def validate_plan(plan: dict, inventory: List, wb: WorkbookData) -> ValidationOutcome:
    blocks_by_id = {b.id: b for b in inventory}
    out = ValidationOutcome()

    for slide in plan.get("slides", []):
        stype = slide.get("type")
        if stype == "title":
            # title slide always kept (informational, no block ref required)
            out.slides.append(_attach_provenance(slide, None))
            continue

        # Resolve block ref(s)
        ref_keys = []
        if "block_ref" in slide:
            ref_keys.append(slide["block_ref"])
        if "block_refs" in slide:
            ref_keys.extend(slide["block_refs"])
        if "supports_block" in slide:
            ref_keys.append(slide["supports_block"])

        if not ref_keys:
            out.dropped.append({"type": stype, "reason": "missing_required_field"})
            continue

        unknown = [r for r in ref_keys if r not in blocks_by_id]
        if unknown:
            out.dropped.append({"type": stype, "reason": "block_ref_not_found",
                                  "block_refs": unknown})
            continue

        ref_blocks = [blocks_by_id[r] for r in ref_keys]

        if stype == "chart":
            primary = ref_blocks[0]
            bad_flag = next((f for f in primary.quality_flags
                              if f in {"single_dominant_category", "all_zero",
                                        "low_fill_ratio", "too_few_rows"}), None)
            if bad_flag:
                out.dropped.append({"type": stype, "reason": bad_flag,
                                      "block_ref": primary.id})
                continue

        if stype == "table":
            primary = ref_blocks[0]
            if "low_fill_ratio" in primary.quality_flags:
                out.dropped.append({"type": stype, "reason": "low_fill_ratio",
                                      "block_ref": primary.id})
                continue

        if stype == "text_bullets":
            block = ref_blocks[0]
            kept_bullets = []
            for bullet in slide.get("bullets", []):
                if _bullet_has_provenance(bullet, block, wb):
                    kept_bullets.append(_clean_bullet(bullet))
                else:
                    out.bullets_dropped += 1
            if not kept_bullets:
                out.dropped.append({"type": stype,
                                      "reason": "all_bullets_failed_provenance",
                                      "block_ref": block.id})
                continue
            slide = {**slide, "bullets": kept_bullets}

        out.slides.append(_attach_provenance(slide, ref_blocks[0] if ref_blocks else None))

    return out


def _attach_provenance(slide: dict, block) -> dict:
    if block is None:
        return slide
    return {**slide, "provenance": {
        "sheet": block.provenance.sheet,
        "columns": block.provenance.columns,
        "rows": list(block.provenance.rows),
    }}


_NUM_RE = re.compile(r"-?\d+(?:[\.,]\d+)?")
_NAME_RE = re.compile(r"\b[A-ZÁÉÍÓÚÑ][a-záéíóúñ]{2,}\b")

def _bullet_has_provenance(bullet: str, block, wb: WorkbookData) -> bool:
    """Returns True if every number / proper name in the bullet exists in the block."""
    if not bullet or not bullet.strip():
        return False
    text = str(bullet).strip()
    nums = _NUM_RE.findall(text)
    names = _NAME_RE.findall(text)
    if not nums and not names:
        # Pure-qualitative bullet → not allowed (would be generic)
        return False

    # Get the actual data slice for this block
    sheet = next((s for s in wb.sheets if s.name == block.provenance.sheet), None)
    if not sheet:
        return False
    cols_set = set(block.provenance.columns)

    # Build a flat searchable string of relevant cells
    haystack_str = ""
    haystack_nums = set()
    for col in sheet.columns:
        if col.name not in cols_set:
            continue
        for s in col.samples:
            if s is None:
                continue
            if isinstance(s, (int, float)):
                haystack_nums.add(round(float(s), 2))
            else:
                haystack_str += " " + str(s).lower()
        for tv in (col.top_values or []):
            haystack_str += " " + str(tv[0]).lower()

    for n in nums:
        try:
            n_val = float(n.replace(",", "."))
            if not any(abs(n_val - h) <= max(1.0, abs(h) * 0.005) for h in haystack_nums):
                return False
        except ValueError:
            continue

    text_lower = text.lower()
    for name in names:
        if name.lower() not in haystack_str:
            return False

    return True


def _clean_bullet(text: str) -> str:
    s = str(text).strip()
    for ugly in UGLY_LITERALS:
        s = re.sub(rf"\b{re.escape(ugly)}\b", "", s, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", s).strip()
```

- [ ] **Step 4: Run tests, expect pass**

```bash
pytest __tests__/python/test_validator.py -v
```

Expected: 4 passed.

- [ ] **Step 5: Commit**

```bash
git add socya_pipeline/validator.py __tests__/python/test_validator.py
git commit -m "feat(pipeline): strict validator with bullet provenance enforcement

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task 2.5: Implement `extractor.py`

**Files:**
- Create: `socya_pipeline/extractor.py`
- Test: `__tests__/python/test_extractor.py`

- [ ] **Step 1: Write the failing test**

```python
# __tests__/python/test_extractor.py
from socya_pipeline.parser import parse_workbook
from socya_pipeline.inventory import build_inventory
from socya_pipeline.validator import validate_plan
from socya_pipeline.extractor import extract_for_render

def test_extracts_table_with_clean_cells(ventas_simple_path):
    wb = parse_workbook(ventas_simple_path)
    inv = build_inventory(wb)
    table_block = next(b for b in inv if b.kind == "table")
    plan = {"presentation_meta": {"title":"T","subtitle":"S"},
            "slides": [{"type": "table", "title": "x",
                         "block_ref": table_block.id,
                         "columns_subset": [c for c in table_block.provenance.columns[:3]],
                         "max_rows": 5}]}
    outcome = validate_plan(plan, inv, wb)
    rendered = extract_for_render(outcome.slides, inv, wb,
                                    file_path=ventas_simple_path)
    assert len(rendered) == 1
    table_slide = rendered[0]
    assert "rows" in table_slide["data"]
    assert len(table_slide["data"]["rows"]) <= 5
    # No ugly literals
    flat = [str(c) for row in table_slide["data"]["rows"] for c in row]
    for forbidden in ("nan", "None", "NaT", "null"):
        assert forbidden not in flat

def test_kpi_extraction(ventas_simple_path):
    wb = parse_workbook(ventas_simple_path)
    inv = build_inventory(wb)
    kpi_block = next(b for b in inv if b.kind == "kpi_candidate")
    plan = {"presentation_meta":{"title":"T","subtitle":"S"},
            "slides": [{"type": "kpi_row", "title": "k",
                         "block_refs": [kpi_block.id]}]}
    outcome = validate_plan(plan, inv, wb)
    rendered = extract_for_render(outcome.slides, inv, wb,
                                    file_path=ventas_simple_path)
    assert rendered[0]["data"]["kpis"][0]["label"]
    assert rendered[0]["data"]["kpis"][0]["value"]
```

- [ ] **Step 2: Run, expect failure**

```bash
pytest __tests__/python/test_extractor.py -v
```

- [ ] **Step 3: Implement `socya_pipeline/extractor.py`**

```python
"""Pulls real cell data per validated slide. Strips ugly literals."""
import math
from pathlib import Path
import pandas as pd
from socya_pipeline.parser import WorkbookData

UGLY_LITERALS_LOWER = {"nan", "none", "null", "nat", "???", "—", "s/d", "n/a", "na"}

def extract_for_render(validated_slides, inventory, wb: WorkbookData,
                         file_path) -> list:
    blocks_by_id = {b.id: b for b in inventory}
    xls = pd.ExcelFile(Path(file_path))
    sheets_cache = {}

    rendered = []
    for slide in validated_slides:
        stype = slide.get("type")
        if stype == "title":
            rendered.append({**slide, "data": {
                "title": slide.get("title", ""),
                "subtitle": slide.get("subtitle", ""),
            }})
            continue

        primary_id = slide.get("block_ref") or (slide.get("block_refs") or [None])[0] \
                       or slide.get("supports_block")
        block = blocks_by_id.get(primary_id) if primary_id else None
        if block is None:
            continue

        sheet_name = block.provenance.sheet
        if sheet_name not in sheets_cache:
            sheets_cache[sheet_name] = xls.parse(sheet_name)
        df = sheets_cache[sheet_name]

        if stype == "kpi_row":
            kpis = []
            for ref in slide.get("block_refs", []):
                b = blocks_by_id.get(ref)
                if not b or b.kind != "kpi_candidate":
                    continue
                value = b.extra.get("value")
                if value is None:
                    continue
                kpis.append({"label": b.label, "value": _format_kpi_value(value)})
            if kpis:
                rendered.append({**slide, "data": {"kpis": kpis}})

        elif stype == "chart":
            chart_data = _build_chart_data(block, df, slide.get("chart_type", "bar"))
            if chart_data:
                rendered.append({**slide, "data": chart_data})

        elif stype == "table":
            cols = slide.get("columns_subset") or block.provenance.columns
            cols = [c for c in cols if c in df.columns]
            sub = df[cols].copy()
            sub = _clean_dataframe(sub)
            max_rows = int(slide.get("max_rows") or 12)
            if sub.empty or len(sub.columns) < 2:
                continue
            sub = sub.head(max_rows)
            rendered.append({**slide, "data": {
                "headers": list(sub.columns),
                "rows": sub.values.tolist(),
            }})

        elif stype == "text_bullets":
            bullets = slide.get("bullets") or []
            if bullets:
                rendered.append({**slide, "data": {"bullets": bullets}})

    return rendered


def _format_kpi_value(value):
    try:
        f = float(value)
    except (TypeError, ValueError):
        return str(value)
    if math.isnan(f):
        return "—"
    if abs(f) >= 1_000_000:
        return f"{f/1_000_000:.1f}M"
    if abs(f) >= 1_000:
        return f"{f/1_000:.1f}K"
    if f.is_integer():
        return str(int(f))
    return f"{f:.2f}"


def _build_chart_data(block, df, chart_type):
    if block.kind == "categorical_distribution":
        col = block.provenance.columns[0]
        if col not in df.columns:
            return None
        vc = df[col].dropna().astype(str).value_counts().head(6)
        if len(vc) < 2:
            return None
        return {
            "chart_type": chart_type,
            "name": col,
            "labels": vc.index.tolist(),
            "values": vc.values.tolist(),
        }
    if block.kind == "time_series_candidate":
        x_col = block.extra.get("x")
        y_col = block.extra.get("y")
        if x_col not in df.columns or y_col not in df.columns:
            return None
        sub = df[[x_col, y_col]].dropna()
        if len(sub) < 2:
            return None
        sub = sub.sort_values(x_col).head(20)
        return {
            "chart_type": "line",
            "name": y_col,
            "labels": [str(x) for x in sub[x_col]],
            "values": [float(v) for v in pd.to_numeric(sub[y_col], errors="coerce").fillna(0)],
        }
    return None


def _clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    def clean_cell(v):
        if v is None:
            return ""
        if isinstance(v, float) and math.isnan(v):
            return ""
        s = str(v).strip()
        if s.lower() in UGLY_LITERALS_LOWER:
            return ""
        return s

    cleaned = df.applymap(clean_cell)
    # Drop rows with <50% filled
    row_fill = cleaned.apply(lambda r: sum(1 for v in r if v != "") / max(1, len(r)),
                                axis=1)
    cleaned = cleaned[row_fill >= 0.5]
    # Drop columns with <30% filled
    col_fill = cleaned.apply(lambda c: sum(1 for v in c if v != "") / max(1, len(c)),
                                axis=0)
    cleaned = cleaned.loc[:, col_fill >= 0.3]
    return cleaned
```

- [ ] **Step 4: Run tests, expect pass**

```bash
pytest __tests__/python/test_extractor.py -v
```

Expected: 2 passed.

- [ ] **Step 5: Commit**

```bash
git add socya_pipeline/extractor.py __tests__/python/test_extractor.py
git commit -m "feat(pipeline): add extractor that strips ugly literals and applies fill ratios

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task 2.6: Implement `cli.py` and feature-flag wire into routes

**Files:**
- Create: `socya_pipeline/cli.py`
- Modify: `app/api/advanced-generate/route.ts` (use new pipeline behind flag)

- [ ] **Step 1: Write `socya_pipeline/cli.py`**

```python
"""CLI entry: `python -m socya_pipeline {analyze|generate} <args>`."""
import argparse
import json
import sys
from pathlib import Path
from socya_pipeline.errors import PipelineError, ErrorCode
from socya_pipeline.parser import parse_workbook
from socya_pipeline.inventory import build_inventory
from socya_pipeline.planner import plan_presentation
from socya_pipeline.validator import validate_plan
from socya_pipeline.extractor import extract_for_render
from socya_pipeline.ai_chain import AIProfile

import os
import re

def _emit_error(err: PipelineError):
    sys.stdout.write(json.dumps({"error": err.to_dict()}, ensure_ascii=False))
    sys.exit(2)

def _load_request(raw: str) -> dict:
    try:
        return json.loads(raw or "{}")
    except json.JSONDecodeError:
        return {}

def cmd_plan(args):
    request = _load_request(args.request)
    api_key = _resolve_api_key()
    profile = (AIProfile.PATIENT if (os.environ.get("SOCYA_AI_PROFILE", "fast")
                                       .lower() == "patient")
                 else AIProfile.FAST)
    try:
        wb = parse_workbook(args.input)
        inv = build_inventory(wb)
        plan = plan_presentation(
            wb, inv,
            user_prompt=request.get("prompt", ""),
            audience=request.get("audience", "ejecutivos"),
            language=request.get("language", "es"),
            api_key=api_key,
            profile=profile,
            file_path=Path(args.input),
        )
        outcome = validate_plan(plan, inv, wb)
        if not outcome.slides:
            raise PipelineError(
                ErrorCode.EXCEL_INSUFFICIENT_DATA,
                f"El Excel '{wb.filename}' no tiene datos suficientes para "
                "generar slides validados.",
                details=f"Slides planeados: {len(plan.get('slides', []))}, "
                f"descartados: {len(outcome.dropped)}.",
                user_action="improve_excel_or_change_prompt",
            )
        rendered = extract_for_render(outcome.slides, inv, wb, args.input)
        result = {
            "presentation_meta": plan.get("presentation_meta", {}),
            "slides": rendered,
            "prompt_suggestions": plan.get("prompt_suggestions", []),
            "ai_status": plan.get("_meta", {}),
            "audit": {
                "slides_planned": len(plan.get("slides", [])),
                "slides_validated": len(outcome.slides),
                "slides_dropped": outcome.dropped,
                "bullets_dropped": outcome.bullets_dropped,
            },
        }
        sys.stdout.write(json.dumps(result, ensure_ascii=False, default=str))
    except PipelineError as e:
        _emit_error(e)
    except Exception as e:
        _emit_error(PipelineError(ErrorCode.PYTHON_RUNTIME_ERROR,
                                    "Error inesperado en el planificador.",
                                    details=str(e)[:300]))

def _resolve_api_key() -> str:
    key = os.environ.get("OPENROUTER_API_KEY", "").strip()
    if key:
        return key
    for env_file in (".env", ".env.local"):
        if os.path.exists(env_file):
            with open(env_file, "r", encoding="utf-8") as f:
                for line in f:
                    if line.startswith("OPENROUTER_API_KEY="):
                        return line.split("=", 1)[1].strip().strip('"').strip("'")
    return ""

def main():
    p = argparse.ArgumentParser()
    sub = p.add_subparsers(dest="command", required=True)
    plan_p = sub.add_parser("plan")
    plan_p.add_argument("--input", required=True)
    plan_p.add_argument("--request", default="{}")
    plan_p.set_defaults(func=cmd_plan)
    args = p.parse_args()
    args.func(args)

if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Smoke-test the CLI directly**

```bash
python -m socya_pipeline plan --input __tests__/fixtures/ventas_simple.xlsx \
  --request '{"prompt":"Resume las ventas por ciudad","audience":"ejecutivos","language":"es"}' \
  | python -m json.tool | head -40
```

Expected: a JSON with `presentation_meta`, `slides`, `prompt_suggestions`, `ai_status`, `audit`. (Or, if `OPENROUTER_API_KEY` missing, an error JSON with `code=AI_SATURATED`.)

- [ ] **Step 3: Add feature-flag branch in `/api/advanced-generate/route.ts`**

In the `try` block where `execFileAsync('python', args, ...)` is called, BEFORE the call, check the flag:

```typescript
const useNewPipeline = process.env.SOCYA_USE_NEW_PIPELINE === '1';
if (useNewPipeline) {
  const newArgs = ['-X', 'utf8', '-m', 'socya_pipeline', 'plan',
    '--input', filePath, '--request', presentationRequest];
  const { stdout, stderr } = await execFileAsync('python', newArgs, {
    encoding: 'utf8', timeout: pythonTimeoutMs, maxBuffer: 20 * 1024 * 1024,
    windowsHide: true,
    env: { ...process.env, PYTHONUTF8: '1', SOCYA_AI_PROFILE: 'patient' },
  });
  if (stderr?.trim() && !stdout?.trim()) throw new Error(stderr.trim());
  return NextResponse.json(JSON.parse(stdout), { headers: { 'Cache-Control': 'no-store' } });
}
// else: existing behavior continues unchanged
```

- [ ] **Step 4: Manual smoke test with feature flag**

```bash
SOCYA_USE_NEW_PIPELINE=1 npm run dev &
curl -X POST http://localhost:3001/api/advanced-generate \
  -F "file=@__tests__/fixtures/ventas_simple.xlsx" \
  -F "userPrompt=Resume las ventas por ciudad" \
  -F "audience=ejecutivos" -F "language=Español" \
  -o /tmp/new_plan.json
python -m json.tool /tmp/new_plan.json | head -50
```

Expected: real plan with validated slides + audit info.

- [ ] **Step 5: Commit**

```bash
git add socya_pipeline/cli.py app/api/advanced-generate/route.ts
git commit -m "feat(pipeline): wire new pipeline behind SOCYA_USE_NEW_PIPELINE flag

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task 2.7: F2 acceptance check

- [ ] **Step 1: Run all tests**

```bash
pytest __tests__/python/ -v
npm test
npm run lint
```

All must pass.

- [ ] **Step 2: Tag F2**

```bash
git tag -a f2-planner-provenance -m "F2 complete: planner with strict provenance behind flag"
```

---

## Phase 3 — Strict Renderer + Drop Generic Templates

**Goal of F3:** New pipeline becomes the only path. Old organizer fallback paths deleted. Renderer is execute-only.

**Deployable after F3:** Yes — single pipeline, all slides go through validator/extractor.

### Task 3.1: Implement `renderer.py`

**Files:**
- Create: `socya_pipeline/renderer.py`
- Test: `__tests__/python/test_renderer.py`

- [ ] **Step 1: Write the failing test**

```python
# __tests__/python/test_renderer.py
from pathlib import Path
import pytest
from pptx import Presentation
from socya_pipeline.renderer import render_pptx

TEMPLATE = Path(__file__).resolve().parent.parent.parent / "Plantilla_Presentacion_Socya (1) (1).pptx"

@pytest.mark.skipif(not TEMPLATE.exists(), reason="template not available")
def test_renders_minimal_deck(tmp_path):
    rendered_slides = [
        {"type": "title", "data": {"title": "Test", "subtitle": "Sub"}},
        {"type": "kpi_row", "data": {"kpis": [{"label": "Total", "value": "12.4M"}]}},
        {"type": "table", "data": {"headers": ["Ciudad", "Total"],
                                     "rows": [["Bogotá", "100"], ["Medellín", "80"]]}},
    ]
    out = tmp_path / "test.pptx"
    render_pptx(rendered_slides,
                presentation_meta={"title": "T", "subtitle": "S"},
                template_path=TEMPLATE, output_path=out)
    assert out.exists() and out.stat().st_size > 1000
    prs = Presentation(out)
    assert len(prs.slides) == 3
```

- [ ] **Step 2: Run, expect failure**

```bash
pytest __tests__/python/test_renderer.py -v
```

- [ ] **Step 3: Implement `socya_pipeline/renderer.py`**

```python
"""Pure execute renderer. No fallbacks. No template fillers."""
import io
from pathlib import Path
from typing import List, Optional
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

PRIMARY = RGBColor(0x0F, 0x17, 0x2A)
ACCENT = RGBColor(0x25, 0x63, 0xEB)
TEXT_DARK = RGBColor(0x33, 0x33, 0x33)

def render_pptx(slides: List[dict], presentation_meta: dict,
                  template_path: Path, output_path: Path) -> None:
    prs = Presentation(str(template_path))
    # Strip any pre-existing slides from the template so we control the deck
    while prs.slides and len(prs.slides) > 0:
        rId = prs.slides._sldIdLst[0].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[0]

    blank_layout = _pick_blank_layout(prs)

    for slide_def in slides:
        stype = slide_def.get("type")
        if stype == "title":
            _add_title(prs, blank_layout, slide_def["data"])
        elif stype == "kpi_row":
            _add_kpi_row(prs, blank_layout, slide_def)
        elif stype == "chart":
            _add_chart(prs, blank_layout, slide_def)
        elif stype == "table":
            _add_table(prs, blank_layout, slide_def)
        elif stype == "text_bullets":
            _add_bullets(prs, blank_layout, slide_def)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))


def _pick_blank_layout(prs):
    for layout in prs.slide_layouts:
        if "Blank" in layout.name or "blanco" in layout.name.lower():
            return layout
    return prs.slide_layouts[0]


def _add_title(prs, layout, data):
    s = prs.slides.add_slide(layout)
    _put_text(s, data.get("title", ""), Inches(0.7), Inches(2.2), Inches(8.5),
                Inches(1.5), font_size=44, bold=True, color=PRIMARY)
    _put_text(s, data.get("subtitle", ""), Inches(0.7), Inches(3.8), Inches(8.5),
                Inches(1.0), font_size=20, color=TEXT_DARK)


def _add_kpi_row(prs, layout, slide_def):
    s = prs.slides.add_slide(layout)
    _put_text(s, slide_def.get("title", ""), Inches(0.5), Inches(0.5), Inches(9),
                Inches(0.7), font_size=24, bold=True, color=PRIMARY)
    kpis = slide_def["data"].get("kpis", [])[:4]
    if not kpis:
        return
    width = Inches(8.5 / max(1, len(kpis)))
    for i, k in enumerate(kpis):
        x = Inches(0.5) + width * i
        _put_text(s, k["value"], x, Inches(2.2), width, Inches(1.0),
                    font_size=36, bold=True, color=ACCENT, align="center")
        _put_text(s, k["label"], x, Inches(3.4), width, Inches(0.6),
                    font_size=14, color=TEXT_DARK, align="center")


def _add_chart(prs, layout, slide_def):
    s = prs.slides.add_slide(layout)
    _put_text(s, slide_def.get("title", ""), Inches(0.5), Inches(0.4), Inches(9),
                Inches(0.7), font_size=22, bold=True, color=PRIMARY)
    data = slide_def["data"]
    img = _build_chart_png(data)
    s.shapes.add_picture(img, Inches(0.7), Inches(1.4), Inches(8.5), Inches(5.0))
    if slide_def.get("narrative"):
        _put_text(s, slide_def["narrative"], Inches(0.7), Inches(6.6),
                    Inches(8.5), Inches(0.6), font_size=12, color=TEXT_DARK)


def _build_chart_png(data: dict) -> io.BytesIO:
    chart_type = data.get("chart_type", "bar")
    labels = data.get("labels", [])
    values = data.get("values", [])
    fig, ax = plt.subplots(figsize=(8, 4.5), dpi=120)
    if chart_type == "pie":
        ax.pie(values, labels=labels, autopct="%1.1f%%",
               colors=plt.cm.Blues([0.4 + 0.5 * i / max(1, len(labels))
                                       for i in range(len(labels))]))
    elif chart_type == "line":
        ax.plot(labels, values, marker="o", color="#2563EB", linewidth=2)
        ax.tick_params(axis="x", rotation=45)
    else:  # bar
        ax.bar(labels, values, color="#2563EB")
        ax.tick_params(axis="x", rotation=30)
    ax.set_title(data.get("name", ""), fontsize=12)
    plt.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def _add_table(prs, layout, slide_def):
    s = prs.slides.add_slide(layout)
    _put_text(s, slide_def.get("title", ""), Inches(0.5), Inches(0.4), Inches(9),
                Inches(0.7), font_size=22, bold=True, color=PRIMARY)
    data = slide_def["data"]
    headers = data.get("headers", [])
    rows = data.get("rows", [])
    if not headers or not rows:
        return
    n_cols = len(headers)
    n_rows = len(rows) + 1
    table_shape = s.shapes.add_table(n_rows, n_cols, Inches(0.5), Inches(1.4),
                                       Inches(9), Inches(0.4 * n_rows + 0.4))
    tbl = table_shape.table
    for i, h in enumerate(headers):
        cell = tbl.cell(0, i)
        cell.text = str(h)
        for r in cell.text_frame.paragraphs:
            for run in r.runs:
                run.font.bold = True
                run.font.size = Pt(11)
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        cell.fill.solid(); cell.fill.fore_color.rgb = PRIMARY
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, val in enumerate(row[:n_cols]):
            tbl.cell(r_idx, c_idx).text = "" if val is None else str(val)


def _add_bullets(prs, layout, slide_def):
    s = prs.slides.add_slide(layout)
    _put_text(s, slide_def.get("title", ""), Inches(0.5), Inches(0.4), Inches(9),
                Inches(0.7), font_size=22, bold=True, color=PRIMARY)
    bullets = slide_def["data"].get("bullets", [])
    body = s.shapes.add_textbox(Inches(0.7), Inches(1.5), Inches(8.5), Inches(5))
    tf = body.text_frame
    tf.word_wrap = True
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"• {b}"
        for run in p.runs:
            run.font.size = Pt(16)
            run.font.color.rgb = TEXT_DARK


def _put_text(slide, text, x, y, w, h, *, font_size=14, bold=False,
                color=None, align="left"):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = str(text or "")
    if align == "center":
        from pptx.enum.text import PP_ALIGN
        p.alignment = PP_ALIGN.CENTER
    for run in p.runs:
        run.font.size = Pt(font_size)
        run.font.bold = bold
        if color is not None:
            run.font.color.rgb = color
```

- [ ] **Step 4: Run test, expect pass**

```bash
pytest __tests__/python/test_renderer.py -v
```

Expected: 1 passed.

- [ ] **Step 5: Commit**

```bash
git add socya_pipeline/renderer.py __tests__/python/test_renderer.py
git commit -m "feat(pipeline): add execute-only renderer with template-stripping

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task 3.2: Add `generate` CLI command

**Files:**
- Modify: `socya_pipeline/cli.py` (add `generate` subcommand)

- [ ] **Step 1: Add `cmd_generate` and register subcommand in `cli.py`**

In `socya_pipeline/cli.py`, after `cmd_plan`, add:

```python
def cmd_generate(args):
    request = _load_request(args.request)
    api_key = _resolve_api_key()
    profile = AIProfile.PATIENT
    try:
        wb = parse_workbook(args.input)
        inv = build_inventory(wb)
        plan = plan_presentation(
            wb, inv,
            user_prompt=request.get("prompt", ""),
            audience=request.get("audience", "ejecutivos"),
            language=request.get("language", "es"),
            api_key=api_key,
            profile=profile,
            file_path=Path(args.input),
        )
        outcome = validate_plan(plan, inv, wb)
        if not outcome.slides:
            raise PipelineError(
                ErrorCode.EXCEL_INSUFFICIENT_DATA,
                f"El Excel '{wb.filename}' no produjo slides validados.",
                details=f"Descartados: {len(outcome.dropped)}.",
                user_action="improve_excel_or_change_prompt",
            )
        rendered = extract_for_render(outcome.slides, inv, wb, args.input)

        from socya_pipeline.renderer import render_pptx
        template = Path(args.template)
        render_pptx(rendered, plan.get("presentation_meta", {}),
                     template_path=template, output_path=Path(args.output))

        # Write audit JSON next to output
        audit = {
            "model_used": plan.get("_meta", {}).get("model"),
            "cache_hit": plan.get("_meta", {}).get("cache_hit", False),
            "fallback_chain_steps": plan.get("_meta", {}).get("fallback_steps", []),
            "slides_planned": len(plan.get("slides", [])),
            "slides_validated": len(outcome.slides),
            "slides_dropped": outcome.dropped,
            "bullets_dropped": outcome.bullets_dropped,
            "provenance_per_slide": [s.get("provenance") for s in rendered],
        }
        audit_path = Path(args.output).with_suffix(".audit.json")
        audit_path.write_text(json.dumps(audit, ensure_ascii=False, indent=2),
                                encoding="utf-8")

        sys.stdout.write(json.dumps({"ok": True, "audit": audit}, ensure_ascii=False))
    except PipelineError as e:
        _emit_error(e)
    except Exception as e:
        _emit_error(PipelineError(ErrorCode.PYTHON_RUNTIME_ERROR,
                                    "Error inesperado en la generación.",
                                    details=str(e)[:300]))
```

And in `main()`:

```python
gen_p = sub.add_parser("generate")
gen_p.add_argument("--input", required=True)
gen_p.add_argument("--output", required=True)
gen_p.add_argument("--template", required=True)
gen_p.add_argument("--request", default="{}")
gen_p.set_defaults(func=cmd_generate)
```

- [ ] **Step 2: Smoke test CLI generate**

```bash
python -m socya_pipeline generate \
  --input __tests__/fixtures/ventas_simple.xlsx \
  --output /tmp/test_new.pptx \
  --template "Plantilla_Presentacion_Socya (1) (1).pptx" \
  --request '{"prompt":"Resume las ventas por ciudad","audience":"ejecutivos","language":"es"}'
ls -la /tmp/test_new.pptx /tmp/test_new.audit.json
```

Expected: both files created.

- [ ] **Step 3: Commit**

```bash
git add socya_pipeline/cli.py
git commit -m "feat(pipeline): add generate CLI command with audit output

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task 3.3: Switch `/api/generate-pptx` to call new pipeline (still legacy fallback retained)

**Files:**
- Modify: `app/api/generate-pptx/route.ts`

- [ ] **Step 1: Add new-pipeline branch in route handler**

In `app/api/generate-pptx/route.ts`, before the existing `execFileAsync('python', args, ...)` call, add:

```typescript
const useNewPipeline = process.env.SOCYA_USE_NEW_PIPELINE === '1';
if (useNewPipeline) {
  const templatePath = path.join(process.cwd(), 'Plantilla_Presentacion_Socya (1) (1).pptx');
  const newArgs = ['-X', 'utf8', '-m', 'socya_pipeline', 'generate',
    '--input', inputPath, '--output', outputPath,
    '--template', templatePath, '--request', presentationRequest];
  await execFileAsync('python', newArgs, {
    encoding: 'utf8', timeout: pythonTimeoutMs, maxBuffer: 20 * 1024 * 1024,
    windowsHide: true,
    env: { ...process.env, PYTHONUTF8: '1', SOCYA_AI_PROFILE: 'patient' },
  });
  await fs.access(outputPath);
  const pptxBuffer = await fs.readFile(outputPath);
  return new NextResponse(pptxBuffer, {
    status: 200,
    headers: {
      'Cache-Control': 'no-store',
      'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      'Content-Disposition': `attachment; filename="${path.basename(outputPath)}"`,
    },
  });
}
// else: existing legacy behavior (kept for rollback safety)
```

- [ ] **Step 2: Smoke test**

```bash
SOCYA_USE_NEW_PIPELINE=1 npm run dev &
curl -X POST http://localhost:3001/api/generate-pptx \
  -F "file=@__tests__/fixtures/ventas_simple.xlsx" \
  -F "userPrompt=Genera el PPT por ciudad" \
  --output /tmp/post_F3.pptx
ls -la /tmp/post_F3.pptx
```

Open `/tmp/post_F3.pptx` in PowerPoint. Verify: real data, no `nan`/`null`, real chart, real table.

- [ ] **Step 3: Commit**

```bash
git add app/api/generate-pptx/route.ts
git commit -m "feat(api): wire generate-pptx to new pipeline behind flag

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task 3.4: Make new pipeline the default; remove legacy fallback paths from organizer

**Files:**
- Modify: `app/api/advanced-generate/route.ts` (drop legacy branch)
- Modify: `app/api/generate-pptx/route.ts` (drop legacy branch)
- Modify: `app/api/excel-intelligence/route.ts` (call new `plan` for preview)
- Modify: `organizer.py` (remove `build_local_*` template functions)

- [ ] **Step 1: Verify the flag-based rollouts behave identically to legacy on 3 fixtures**

Run with and without `SOCYA_USE_NEW_PIPELINE=1` for `ventas_simple`, `comisiones_real`, `dominio_raro`. Side-by-side compare PPTs in PowerPoint. Document any regression and stop here if found (loop until matching quality or better).

- [ ] **Step 2: Remove the conditional, make new pipeline the only path**

In `app/api/generate-pptx/route.ts`, replace the old `args = ['-X', 'utf8', GENERATOR_SCRIPT_NAME, ...]` logic with the new pipeline call (drop the `useNewPipeline` flag check; new pipeline is now unconditional).

In `app/api/advanced-generate/route.ts`, do the same — drop the legacy `args = ['-X', 'utf8', ORGANIZER_SCRIPT_NAME, filePath, presentationRequest]` branch and the `buildGenericSlides` function. Replace with calling `socya_pipeline plan`. The route now just forwards the JSON output from the CLI.

In `app/api/excel-intelligence/route.ts`, replace the call to `organizer --panel-report` with `socya_pipeline plan` (profile=fast). Map the output to the `IntelligenceResponse` shape the UI expects (semanticSummary, executiveSummary, suggestions, etc.) — see Task 4.1 for adapter helper.

- [ ] **Step 3: Strip legacy fallback functions from `organizer.py`**

Delete or stub-out (raise NotImplementedError) these top-level functions in `organizer.py`:
- `build_local_executive_summary`
- `build_local_briefing_payload`
- `enrich_textual_blocks_with_ai`
- `build_unified_ai_response` and similar multi-call AI helpers
- `buildGenericSlides`'s Python equivalents — any function that constructs a slide payload outside the new pipeline

If any function is called from `generate_template_presentation.py` or elsewhere and we can't yet remove it without breaking, leave a `# TODO: remove in F5` comment and continue. (Will be cleaned in F5.)

- [ ] **Step 4: Run full test suite + manual verification**

```bash
pytest __tests__/python/ -v
npm test
npm run lint
npm run build
```

Then manual: 3 fixtures end-to-end via UI. Confirm:
- No "Análisis ejecutivo del archivo" in outputs.
- No `nan`/`None`/`null` in any slide.
- Charts have real categories.
- Tables have real columns and rows.

- [ ] **Step 5: If something fails, loop**

The user's directive: **don't deliver broken work**. If any verification fails:
- Diagnose root cause (is it parser? validator? renderer? prompt template?).
- Fix and re-run all verification.
- Don't move forward until F3 is solid.

- [ ] **Step 6: Commit**

```bash
git add organizer.py app/api/
git commit -m "feat: make new pipeline the only path; drop legacy local fallbacks

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task 3.5: F3 acceptance + tag

- [ ] **Step 1: All tests pass**

```bash
pytest __tests__/python/ -v && npm test && npm run lint && npm run build
```

- [ ] **Step 2: Tag**

```bash
git tag -a f3-strict-renderer -m "F3 complete: strict renderer, legacy fallbacks gone"
```

---

## Phase 4 — UI: SSE streaming + audit modal + dynamic AI status

**Goal of F4:** UI reflects real backend state honestly. Progress bar streams. Errors actionable. Audit visible.

**Deployable after F4:** Yes — UX layer matches backend capabilities.

### Task 4.1: TypeScript error contract + AI status helpers

**Files:**
- Create: `utils/error-codes.ts`
- Create: `utils/ai-status.ts`
- Test: `__tests__/error-contract.test.js`

- [ ] **Step 1: Write the failing test**

```javascript
// __tests__/error-contract.test.js
const { test } = require('node:test');
const assert = require('node:assert');

test('error contract has required fields', async () => {
  const { isPipelineError, formatErrorForUser } = require('../utils/error-codes.ts');
  const err = { code: 'AI_SATURATED', message: 'all saturated',
                user_action: 'retry_later', retry_after_seconds: 300 };
  assert.equal(isPipelineError(err), true);
  const formatted = formatErrorForUser(err);
  assert.match(formatted.title, /saturad/i);
  assert.match(formatted.action, /reintenta/i);
});
```

- [ ] **Step 2: Run, expect failure**

```bash
npm test -- --test-name-pattern="error contract"
```

- [ ] **Step 3: Implement `utils/error-codes.ts`**

```typescript
export type ErrorCode =
  | 'EXCEL_INVALID' | 'EXCEL_EMPTY' | 'EXCEL_INSUFFICIENT_DATA'
  | 'AI_SATURATED' | 'AI_RESPONSE_INVALID' | 'PLANNER_REJECTED_PROMPT'
  | 'PYTHON_RUNTIME_ERROR' | 'TIMEOUT';

export type UserAction =
  | 'retry' | 'retry_later' | 'upload_again' | 'change_prompt'
  | 'improve_excel_or_change_prompt' | 'upload_smaller' | 'report_bug';

export interface PipelineErrorPayload {
  code: ErrorCode;
  message: string;
  details?: string;
  user_action: UserAction;
  retry_after_seconds?: number;
}

export function isPipelineError(x: unknown): x is PipelineErrorPayload {
  return Boolean(x && typeof x === 'object' && 'code' in x && 'message' in x
                  && 'user_action' in x);
}

const FRIENDLY: Record<ErrorCode, { title: string; action: string }> = {
  EXCEL_INVALID: {
    title: 'No pude leer este Excel',
    action: 'Sube otro archivo o revisa que no esté dañado.',
  },
  EXCEL_EMPTY: {
    title: 'El Excel no tiene datos legibles',
    action: 'Asegúrate de que tenga al menos una hoja con tablas.',
  },
  EXCEL_INSUFFICIENT_DATA: {
    title: 'Los datos no alcanzan para una presentación',
    action: 'Mejora el Excel (más filas, menos vacíos) o ajusta el prompt.',
  },
  AI_SATURATED: {
    title: 'Servicio IA con mucha demanda ahora',
    action: 'Espera 5-10 min y reintenta. Tu Excel queda listo.',
  },
  AI_RESPONSE_INVALID: {
    title: 'La IA devolvió una respuesta inválida',
    action: 'Reintenta — suele resolverse en el siguiente intento.',
  },
  PLANNER_REJECTED_PROMPT: {
    title: 'El prompt no encaja con este Excel',
    action: 'Cambia el prompt para enfocarte en datos disponibles.',
  },
  PYTHON_RUNTIME_ERROR: {
    title: 'Error técnico inesperado',
    action: 'Reporta el problema con el archivo que usaste.',
  },
  TIMEOUT: {
    title: 'La generación tomó demasiado',
    action: 'Reintenta o sube un archivo más pequeño.',
  },
};

export function formatErrorForUser(err: PipelineErrorPayload) {
  return FRIENDLY[err.code] ?? {
    title: 'Algo salió mal',
    action: err.message || 'Reintenta más tarde.',
  };
}
```

- [ ] **Step 4: Implement `utils/ai-status.ts`**

```typescript
export interface AIStatus {
  model?: string;
  cache_hit?: boolean;
  fallback_steps?: { from: string; reason: string; message?: string }[];
}

const MODEL_LABELS: Record<string, string> = {
  'nousresearch/hermes-3-llama-3.1-405b:free': 'Hermes 3 405B',
  'meta-llama/llama-3.3-70b-instruct:free': 'Llama 3.3 70B',
  'google/gemma-2-9b-it:free': 'Gemma 2 9B',
  'qwen/qwen-2.5-72b-instruct:free': 'Qwen 2.5 72B',
};

export function aiStatusBadge(status?: AIStatus): {
  label: string; tone: 'ok' | 'warn' | 'cache';
} {
  if (!status) return { label: 'Esperando IA…', tone: 'warn' };
  if (status.cache_hit) return {
    label: 'Resultado en caché (sin consumo IA)', tone: 'cache',
  };
  const label = MODEL_LABELS[status.model || ''] || status.model || 'IA';
  if ((status.fallback_steps?.length || 0) > 0) {
    const skipped = status.fallback_steps?.[0]?.from || 'modelo primario';
    const skippedLabel = MODEL_LABELS[skipped] || skipped;
    return { label: `${skippedLabel} saturado → usando ${label}`, tone: 'warn' };
  }
  return { label: `${label} · listo`, tone: 'ok' };
}
```

- [ ] **Step 5: Run test, expect pass**

```bash
npm test
```

- [ ] **Step 6: Commit**

```bash
git add utils/error-codes.ts utils/ai-status.ts __tests__/error-contract.test.js
git commit -m "feat(ui): add error contract types and AI status badge helpers

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task 4.2: SSE streaming for `/api/generate-pptx`

**Files:**
- Create: `utils/sse-stream.ts`
- Modify: `app/api/generate-pptx/route.ts` (split into two endpoints: SSE progress + binary fetch)
- Test: `__tests__/sse-stream.test.js`

**Note:** Read `node_modules/next/dist/docs/` for Next.js 16 streaming patterns before implementing. The codebase uses Next.js 16 which has updated streaming APIs.

- [ ] **Step 1: Read Next.js 16 docs on streaming**

```bash
ls node_modules/next/dist/docs/
# look for files like streaming.md, route-handlers.md
```

Confirm the recommended `Response` + `ReadableStream` shape. Make a note if Next.js 16 expects a different signature than older versions.

- [ ] **Step 2: Write `utils/sse-stream.ts`**

```typescript
export type SSEEvent = {
  phase: 'parsing' | 'inventory' | 'planning' | 'validating' | 'rendering' | 'done' | 'error';
  step?: string;
  message?: string;
  data?: unknown;
};

export function makeSSEStream(): {
  stream: ReadableStream<Uint8Array>;
  send: (event: SSEEvent) => void;
  close: () => void;
} {
  const encoder = new TextEncoder();
  let controller: ReadableStreamDefaultController<Uint8Array>;
  const stream = new ReadableStream<Uint8Array>({
    start(c) { controller = c; },
  });
  const send = (event: SSEEvent) => {
    const payload = `data: ${JSON.stringify(event)}\n\n`;
    controller.enqueue(encoder.encode(payload));
  };
  const close = () => { try { controller.close(); } catch {} };
  return { stream, send, close };
}

export function sseHeaders() {
  return {
    'Content-Type': 'text/event-stream; charset=utf-8',
    'Cache-Control': 'no-cache, no-transform',
    'Connection': 'keep-alive',
    'X-Accel-Buffering': 'no',
  };
}
```

- [ ] **Step 3: Refactor `/api/generate-pptx` to two-phase: SSE progress, then download by token**

In `app/api/generate-pptx/route.ts`:
- New shape: POST starts the generation, streams SSE progress; final SSE event has `data: { downloadToken }`.
- Add a new GET endpoint `/api/generate-pptx?token=...` that returns the binary file.
- In-memory map (`Map<string, { path: string; expires: number }>`) tracks tokens with 5 min TTL.

Key code changes:

```typescript
import { randomUUID } from 'crypto';
import { makeSSEStream, sseHeaders } from '@/utils/sse-stream';

const PENDING: Map<string, { path: string; expires: number }> = (globalThis as any).__SOCYA_PENDING__ ||= new Map();

export async function POST(req: NextRequest) {
  const { stream, send, close } = makeSSEStream();
  const formData = await req.formData();
  const file = formData.get('file');
  // ...validate file...

  // Run generation in background; don't await
  (async () => {
    try {
      send({ phase: 'parsing', step: '1/5', message: 'Leyendo Excel…' });
      const { tempDir, inputPath, outputPath } = await prepareTempPaths(file);

      send({ phase: 'inventory', step: '2/5', message: 'Construyendo inventario…' });
      send({ phase: 'planning', step: '3/5', message: 'Consultando IA (Hermes)…' });

      const presentationRequest = JSON.stringify({ /* ...as before... */ });
      const templatePath = path.join(process.cwd(), 'Plantilla_Presentacion_Socya (1) (1).pptx');
      const args = ['-X', 'utf8', '-m', 'socya_pipeline', 'generate',
        '--input', inputPath, '--output', outputPath,
        '--template', templatePath, '--request', presentationRequest];

      await execFileAsync('python', args, { /* ... */ });

      send({ phase: 'validating', step: '4/5', message: 'Validando datos…' });
      send({ phase: 'rendering', step: '5/5', message: 'Renderizando PPTX…' });

      const token = randomUUID();
      PENDING.set(token, { path: outputPath, expires: Date.now() + 5 * 60_000 });
      // Read audit JSON written by the Python CLI alongside the pptx
      const auditPath = outputPath.replace(/\.pptx$/i, '.audit.json');
      let audit: any = null;
      try {
        const auditText = await fs.readFile(auditPath, 'utf-8');
        audit = JSON.parse(auditText);
      } catch { /* audit is best-effort; UI handles null */ }
      send({ phase: 'done', message: 'Listo.', data: {
        downloadToken: token,
        filename: path.basename(outputPath),
        audit,
      } });
    } catch (err) {
      send({ phase: 'error', message: getErrorMessage(err),
              data: parseErrorPayload(err) });
    } finally {
      close();
    }
  })();

  return new Response(stream, { headers: sseHeaders() });
}

export async function GET(req: NextRequest) {
  const token = new URL(req.url).searchParams.get('token');
  const entry = token ? PENDING.get(token) : null;
  if (!entry || entry.expires < Date.now()) {
    return NextResponse.json({ error: 'token_expired' }, { status: 404 });
  }
  PENDING.delete(token);
  const buffer = await fs.readFile(entry.path);
  // cleanup tempdir asynchronously
  fs.unlink(entry.path).catch(() => {});
  return new NextResponse(buffer, { /* ...headers as before... */ });
}
```

(Helper functions `prepareTempPaths`, `parseErrorPayload`, `getErrorMessage` are extracted from existing code.)

- [ ] **Step 4: Write SSE test**

```javascript
// __tests__/sse-stream.test.js
const { test } = require('node:test');
const assert = require('node:assert');
const { makeSSEStream } = require('../utils/sse-stream.ts');

test('sse stream emits and closes', async () => {
  const { stream, send, close } = makeSSEStream();
  const reader = stream.getReader();
  send({ phase: 'parsing', message: 'go' });
  send({ phase: 'done' });
  close();
  const decoder = new TextDecoder();
  let collected = '';
  while (true) {
    const { value, done } = await reader.read();
    if (done) break;
    collected += decoder.decode(value);
  }
  assert.match(collected, /phase":"parsing/);
  assert.match(collected, /phase":"done/);
});
```

- [ ] **Step 5: Run tests + manual smoke**

```bash
npm test
SOCYA_USE_NEW_PIPELINE=1 npm run dev &
curl -N -X POST http://localhost:3001/api/generate-pptx \
  -F "file=@__tests__/fixtures/ventas_simple.xlsx" \
  -F "userPrompt=Genera el PPT"
```

Expected: SSE events stream in real time, ending with `phase: done` + token. Then GET that token returns the binary.

- [ ] **Step 6: Commit**

```bash
git add utils/sse-stream.ts app/api/generate-pptx/route.ts __tests__/sse-stream.test.js
git commit -m "feat(api): convert generate-pptx to SSE-streamed two-phase flow

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task 4.3: `GenerationProgress` component

**Files:**
- Create: `components/GenerationProgress.tsx`

- [ ] **Step 1: Implement the component**

```typescript
"use client";
import React, { useEffect, useState } from 'react';
import { FileSpreadsheet, Database, Brain, Check, Palette, Loader2 } from 'lucide-react';

const PHASES = [
  { key: 'parsing',    icon: FileSpreadsheet, label: 'Leyendo Excel' },
  { key: 'inventory',  icon: Database,        label: 'Construyendo inventario' },
  { key: 'planning',   icon: Brain,           label: 'Consultando IA' },
  { key: 'validating', icon: Check,           label: 'Validando datos' },
  { key: 'rendering',  icon: Palette,         label: 'Renderizando PPTX' },
] as const;

interface Props {
  currentPhase: typeof PHASES[number]['key'] | 'done' | 'error' | null;
  message?: string;
}

export default function GenerationProgress({ currentPhase, message }: Props) {
  const idx = PHASES.findIndex(p => p.key === currentPhase);
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
      {PHASES.map((p, i) => {
        const done = idx > i || currentPhase === 'done';
        const active = idx === i;
        const Icon = p.icon;
        return (
          <div key={p.key} style={{
            display: 'flex', alignItems: 'center', gap: '0.6rem',
            padding: '0.5rem 0.7rem', borderRadius: '8px',
            background: active ? 'rgba(124,58,237,0.12)' : 'transparent',
            color: done ? '#86EFAC' : active ? '#A78BFA' : 'rgba(255,255,255,0.35)',
            fontSize: '0.78rem',
          }}>
            {active ? <Loader2 size={14} className="spin" /> : <Icon size={14} />}
            <span>{p.label}</span>
            {active && message && (
              <span style={{ marginLeft: 'auto', opacity: 0.7, fontSize: '0.7rem' }}>
                {message}
              </span>
            )}
          </div>
        );
      })}
    </div>
  );
}
```

- [ ] **Step 2: Commit**

```bash
git add components/GenerationProgress.tsx
git commit -m "feat(ui): add GenerationProgress 5-phase indicator

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task 4.4: `AuditModal` component

**Files:**
- Create: `components/AuditModal.tsx`

- [ ] **Step 1: Implement**

```typescript
"use client";
import React from 'react';
import { X } from 'lucide-react';
import { aiStatusBadge, AIStatus } from '@/utils/ai-status';

const REASON_LABELS: Record<string, string> = {
  block_ref_not_found: 'la IA referenció un bloque que no existe',
  single_dominant_category: 'una categoría dominaba >85% (gráfica sin sentido)',
  all_zero_values: 'todos los valores eran cero',
  too_few_categories: 'no había suficientes categorías para una gráfica',
  too_few_rows: 'la tabla tenía muy pocas filas',
  low_fill_ratio: 'demasiadas celdas vacías',
  all_bullets_failed_provenance: 'ningún bullet citaba datos del Excel',
  missing_required_field: 'faltó un campo obligatorio',
};

interface Audit {
  model_used?: string;
  cache_hit?: boolean;
  fallback_chain_steps?: any[];
  slides_planned: number;
  slides_validated: number;
  slides_dropped: { type: string; reason: string; block_ref?: string }[];
  bullets_dropped: number;
}

export default function AuditModal({ audit, onClose }: {
  audit: Audit; onClose: () => void
}) {
  const status = aiStatusBadge({ model: audit.model_used,
                                  cache_hit: audit.cache_hit,
                                  fallback_steps: audit.fallback_chain_steps as any });
  return (
    <div style={{
      position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.6)',
      display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 100,
    }} onClick={onClose}>
      <div onClick={e => e.stopPropagation()} style={{
        background: '#0F172A', borderRadius: '14px', padding: '1.5rem',
        maxWidth: '560px', width: '92%', color: 'white',
      }}>
        <div style={{ display: 'flex', justifyContent: 'space-between',
                       alignItems: 'center', marginBottom: '1rem' }}>
          <h2 style={{ fontSize: '1.1rem', margin: 0 }}>Detalles de la generación</h2>
          <button onClick={onClose} style={{ background: 'none', border: 'none',
                                                color: 'white', cursor: 'pointer' }}>
            <X size={18} />
          </button>
        </div>
        <p style={{ fontSize: '0.85rem', margin: '0 0 1rem',
                     color: '#A78BFA' }}>{status.label}</p>
        <div style={{ fontSize: '0.8rem', display: 'flex', gap: '1rem',
                       marginBottom: '1rem' }}>
          <span><strong>{audit.slides_validated}</strong> slides en el PPT</span>
          <span><strong>{audit.slides_dropped.length}</strong> descartados</span>
          <span><strong>{audit.bullets_dropped}</strong> bullets descartados</span>
        </div>
        {audit.slides_dropped.length > 0 && (
          <>
            <h3 style={{ fontSize: '0.85rem', margin: '0 0 0.5rem' }}>
              Slides que omitimos
            </h3>
            <ul style={{ fontSize: '0.78rem', color: 'rgba(255,255,255,0.7)',
                          paddingLeft: '1.2rem' }}>
              {audit.slides_dropped.map((d, i) => (
                <li key={i} style={{ marginBottom: '0.3rem' }}>
                  Slide tipo <strong>{d.type}</strong>: {REASON_LABELS[d.reason] || d.reason}
                  {d.block_ref ? ` (bloque ${d.block_ref})` : ''}
                </li>
              ))}
            </ul>
          </>
        )}
      </div>
    </div>
  );
}
```

- [ ] **Step 2: Commit**

```bash
git add components/AuditModal.tsx
git commit -m "feat(ui): add AuditModal with plain-spanish reason labels

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task 4.5: Wire UI components into `ExcelUploader.tsx`

**Files:**
- Modify: `components/ExcelUploader.tsx` (consume SSE, show progress, retry, audit modal)

- [ ] **Step 1: Read current `ExcelUploader.tsx` to find the generation flow**

```bash
# Find where the existing Generate button calls /api/generate-pptx
```

Identify the function (likely `handleGenerate` or similar).

- [ ] **Step 2: Replace `handleGenerate` to use SSE**

Replace the `fetch('/api/generate-pptx', ...)` call with an SSE consumer:

```typescript
const [progress, setProgress] = useState<{phase: string; message?: string} | null>(null);
const [audit, setAudit] = useState<any>(null);
const [showAudit, setShowAudit] = useState(false);
const [retryError, setRetryError] = useState<PipelineErrorPayload | null>(null);

async function handleGenerate() {
  if (!file) return;
  setProgress({ phase: 'parsing' });
  setRetryError(null);
  const formData = new FormData();
  formData.append('file', file);
  formData.append('userPrompt', prompt);
  // ...other fields...

  try {
    const res = await fetch('/api/generate-pptx', { method: 'POST', body: formData });
    if (!res.body) throw new Error('No stream');
    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    while (true) {
      const { value, done } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value);
      const events = buffer.split('\n\n');
      buffer = events.pop() || '';
      for (const ev of events) {
        const m = /^data: (.+)$/m.exec(ev.trim());
        if (!m) continue;
        const payload = JSON.parse(m[1]);
        setProgress(payload);
        if (payload.phase === 'done') {
          // Download the file
          const dlRes = await fetch(`/api/generate-pptx?token=${payload.data.downloadToken}`);
          const blob = await dlRes.blob();
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url; a.download = payload.data.filename;
          a.click(); URL.revokeObjectURL(url);
          setProgress(null);
          if (payload.data.audit) {
            setAudit(payload.data.audit);
            // Optionally auto-open the audit modal: setShowAudit(true);
          }
        }
        if (payload.phase === 'error' && payload.data) {
          setRetryError(payload.data);
          setProgress(null);
        }
      }
    }
  } catch (err) {
    setRetryError({
      code: 'PYTHON_RUNTIME_ERROR',
      message: 'Algo salió mal en el navegador.',
      user_action: 'retry',
    } as any);
    setProgress(null);
  }
}
```

- [ ] **Step 3: Insert progress UI and retry banner**

In the JSX, conditionally render:

```tsx
{progress && <GenerationProgress currentPhase={progress.phase as any}
                                    message={progress.message} />}
{retryError && <RetryBanner error={retryError} onRetry={handleGenerate}
                              onCancel={() => setRetryError(null)} />}
{showAudit && audit && <AuditModal audit={audit} onClose={() => setShowAudit(false)} />}
```

(Define a small inline `RetryBanner` using `formatErrorForUser`.)

- [ ] **Step 4: Manual smoke test**

```bash
SOCYA_USE_NEW_PIPELINE=1 npm run dev
# In browser, upload fixture, generate, watch progress, download.
```

Verify: 5 phases tick through, file downloads, retry works on error.

- [ ] **Step 5: Commit**

```bash
git add components/ExcelUploader.tsx
git commit -m "feat(ui): wire ExcelUploader to SSE progress + retry + audit modal

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task 4.6: Update `AIControlPanel.tsx` with dynamic AI status

**Files:**
- Modify: `components/AIControlPanel.tsx`

- [ ] **Step 1: Replace static "Hermes 3 · OpenRouter" header text**

Find the line in `AIControlPanel.tsx` that renders `Hermes 3 · OpenRouter` and replace with dynamic badge based on `IntelligenceResponse.ai_status`:

```tsx
import { aiStatusBadge } from '@/utils/ai-status';

// in component, after fetching intelligence response:
const badge = aiStatusBadge(data.ai_status);
// in JSX:
<p style={{ color: 'rgba(255,255,255,0.35)', fontSize: '0.62rem', margin: 0 }}>
  ✦ {badge.label}
</p>
```

- [ ] **Step 2: Replace `PLACEHOLDER_SUGGESTIONS` with backend suggestions only**

Remove the constant. In the `setSuggestions` flow, drop the `else { setSuggestions(PLACEHOLDER_SUGGESTIONS); }` branch and instead show a friendly "Sin sugerencias IA por ahora — escribe tu propio prompt" empty state.

- [ ] **Step 3: Update `IntelligenceResponse` type to include `ai_status`**

```typescript
interface IntelligenceResponse {
  // ...existing fields...
  ai_status?: { model?: string; cache_hit?: boolean; fallback_steps?: any[] };
  audit?: { slides_planned: number; slides_validated: number;
             slides_dropped: any[]; bullets_dropped: number };
}
```

- [ ] **Step 4: Manual verification**

In browser: upload Excel, confirm header shows real model name (e.g. `✦ Hermes 3 405B · listo`). Force a fallback by renaming your API key (`OPENROUTER_API_KEY=invalid`) and verify the badge reads `Vista previa sin IA`.

- [ ] **Step 5: Commit**

```bash
git add components/AIControlPanel.tsx
git commit -m "feat(ui): dynamic AI status badge and remove placeholder suggestions

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task 4.7: F4 acceptance + tag

- [ ] **Step 1: All tests + build**

```bash
pytest __tests__/python/ -v && npm test && npm run lint && npm run build
```

- [ ] **Step 2: Manual UX checklist**

- [ ] Upload `ventas_simple.xlsx` → panel muestra resumen real, badge "Hermes 3 405B · listo".
- [ ] Click Generar → 5 fases tickean en pantalla, descarga al final.
- [ ] Abrir audit modal → razones legibles en español.
- [ ] Eliminar API key → preview muestra "Vista previa sin IA", retry banner aparece al generar.
- [ ] Re-subir mismo Excel + mismo prompt → badge "Resultado en caché".

- [ ] **Step 3: Tag**

```bash
git tag -a f4-ui-streaming -m "F4 complete: UI with SSE progress, audit modal, dynamic AI status"
```

---

## Phase 5 — Cleanup

**Goal of F5:** Delete dead code. Update docs. Final regression run. Ready for merge.

### Task 5.1: Delete legacy code

**Files to clean:**
- `organizer.py`: remove all `# TODO: remove in F5` markers, delete bodies. Reduce file from ~4400 lines to <800 lines (parser + inventory shim only). The new pipeline does the work.
- `generate_template_presentation.py`: collapse to a 5-line shim that calls `python -m socya_pipeline generate` (or delete entirely; keep if other tooling references it).
- `utils/excel-ai-panel.cjs`: keep only `buildProcessingProfile` (still used by route handlers for timeout calc). Delete `buildExcelIntelligenceReport`, `buildSemanticSummary`, `inferExcelTopic`, `inferInformationType`, `buildKeyFindings`, `buildPatterns`, `buildTrends`, `buildPromptHints`, `inferPalette`.
- `app/api/ai-suggestions/route.ts`: delete entire file. Suggestions now come bundled with planner output.
- `app/api/excel-intelligence/route.ts`: simplify — no longer needs `buildExcelIntelligenceReport` import, just forwards CLI output.

- [ ] **Step 1: For each file, identify exact functions/exports to remove**

```bash
grep -n "TODO: remove in F5" organizer.py
```

For each match: read the function, confirm it has no live caller (`grep -r 'function_name'`), delete.

- [ ] **Step 2: Delete `app/api/ai-suggestions/`**

```bash
rm -rf app/api/ai-suggestions
grep -rn "ai-suggestions" app/ components/ utils/ || echo "no remaining references"
```

If references exist in `AIControlPanel.tsx` or elsewhere, remove them too.

- [ ] **Step 3: Trim `utils/excel-ai-panel.cjs`**

Open the file. Delete every export except `buildProcessingProfile` (and its helpers). Update `module.exports = { buildProcessingProfile, constants: {...} };`.

- [ ] **Step 4: Run all tests + build**

```bash
pytest __tests__/python/ -v && npm test && npm run lint && npm run build
```

If anything fails, an unused export was actually used somewhere. Restore minimally and try again.

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "chore: remove legacy fallback code in organizer, panel, ai-suggestions

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task 5.2: Update documentation

**Files:**
- Modify: `README.md`
- Modify: `GEMINI.md`
- Modify: `AGENTS.md` (if needed)

- [ ] **Step 1: Update README.md**

Replace the "Flujo principal" section to describe the new 5-step pipeline. Add a "Configuración IA" section explaining `OPENROUTER_API_KEY`, `SOCYA_AI_PROFILE`, and the model fallback chain.

- [ ] **Step 2: Update GEMINI.md**

Same updates — keep this in sync as it's the project guide.

- [ ] **Step 3: Commit**

```bash
git add README.md GEMINI.md AGENTS.md
git commit -m "docs: update README and GEMINI to describe new pipeline

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task 5.3: Full regression suite + acceptance

- [ ] **Step 1: Generate fresh PPTs from all 7 fixtures**

```bash
for fixture in ventas_simple comisiones_real casi_vacio enorme dominio_raro; do
  python -m socya_pipeline generate \
    --input __tests__/fixtures/${fixture}.xlsx \
    --output __tests__/output/${fixture}_F5.pptx \
    --template "Plantilla_Presentacion_Socya (1) (1).pptx" \
    --request '{"prompt":"Genera una presentación ejecutiva con datos reales","audience":"ejecutivos","language":"es"}' \
    || echo "FAIL: ${fixture}"
done
```

Expected:
- `ventas_simple`, `comisiones_real`, `dominio_raro`: PPT created.
- `casi_vacio`: error JSON with `EXCEL_INSUFFICIENT_DATA`.
- `enorme`: PPT created (with adaptive truncation).
- `corrupto`: error JSON with `EXCEL_INVALID`.

- [ ] **Step 2: Open each PPT in PowerPoint manually and verify**

Checklist per PPT:
- [ ] No `nan`, `None`, `null`, `NaT`, `???` in any text.
- [ ] All charts have ≥2 categories with non-zero values.
- [ ] All tables have ≥2 columns and ≥2 rows of real data.
- [ ] No "Análisis ejecutivo del archivo" or similar template text.
- [ ] Bullets reference real numbers/names from the source Excel.
- [ ] Audit JSON next to PPT shows model used and validation stats.

- [ ] **Step 3: User review**

Bring the user 3 generated PPTs (`ventas_simple`, `comisiones_real`, `dominio_raro`). Wait for explicit approval that the data fidelity issues are resolved.

If user rejects: identify which problem (validator too laxo? planner prompt weak? renderer styling?) and fix. Loop until approval.

- [ ] **Step 4: Tag final**

```bash
git tag -a v1.0-hermes-driven -m "v1.0: hermes-driven pipeline complete and validated"
```

- [ ] **Step 5: Merge to master**

Only after user approval:

```bash
git checkout master
git merge --no-ff feature/hermes-driven-pipeline -m "merge: hermes-driven pipeline redesign

Implements docs/superpowers/specs/2026-05-07-hermes-driven-pptx-design.md
in 5 phases (F1 anti-quota → F2 planner → F3 strict renderer →
F4 UI streaming → F5 cleanup).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

Do NOT push to origin. Wait for explicit user instruction.

---

## Quick reference for executing agent

- **TDD discipline:** Write the test FIRST, watch it fail, then write code. Don't skip the failure step.
- **Don't fabricate "passing" results:** Only mark a step done after running the actual command and seeing the actual output. The user explicitly said no broken deliveries.
- **Phase boundaries:** Each phase ends with `pytest && npm test && npm run lint && npm run build`. All must pass.
- **Git hygiene:** One commit per task. Use the exact commit messages provided. Never amend committed work.
- **If stuck:** Read the spec at `docs/superpowers/specs/2026-05-07-hermes-driven-pptx-design.md`. Re-read the relevant section before making decisions outside the plan.
