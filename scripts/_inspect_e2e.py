"""One-shot QA script for the E2E generated PPTX (delete after run)."""
import re
import subprocess
import sys

pptx = sys.argv[1] if len(sys.argv) > 1 else "E2E_TEST_OUTPUT.pptx"
result = subprocess.run(
    ["python", "-m", "markitdown", pptx],
    capture_output=True,
)
text = result.stdout.decode("utf-8", errors="replace")
with open("e2e_pptx_md.txt", "w", encoding="utf-8") as f:
    f.write(text)

print("Total length:", len(text))
print("Slide markers:", text.count("<!-- Slide number:"))

print("\n=== PLACEHOLDER LEAK CHECK ===")
patterns = [
    r"xxxx",
    r"lorem",
    r"ipsum",
    r"this.*(page|slide).*layout",
    r"\$1\.50B",
    r"\$1\.5B(?!\d)",
    r"TODO",
    r"PLACEHOLDER",
    r"FIXME",
    r"undefined",
    r"\bNaN\b",
    r"\bnull\b",
    r"\{[a-z_]+\}",
    r"\bNone\b",
    r"placeholder",
]
for p in patterns:
    hits = re.findall(p, text, re.IGNORECASE)
    if hits:
        print(f"  HIT  {p!r:30}  -> {len(hits)} matches: {hits[:5]}")

print("\n=== DOLLAR AMOUNTS FOUND ===")
for m in re.finditer(r"\$[0-9.,]+[KMBkmb]?", text):
    print(" ", m.group())

print("\n=== B-SUFFIX ABUSE CHECK (any '$NNNB' standalone) ===")
for m in re.finditer(r"\$[0-9.,]+B(?![a-zA-Z])", text):
    print("  POSSIBLE FAKE B:", m.group())

print("\n=== SLIDE TITLES (heuristic) ===")
slides = re.split(r"<!-- Slide number: (\d+) -->", text)
for i in range(1, len(slides), 2):
    num = slides[i]
    body = slides[i + 1]
    # First non-empty meaningful line (skip image refs, page footer)
    lines = [
        line.strip()
        for line in body.splitlines()
        if line.strip()
        and not line.strip().startswith("![")
        and not line.strip().startswith("###")
        and "PRESENTACI" not in line
        and not re.match(r"^\d+ / \d+$", line.strip())
    ]
    head = next((l for l in lines if len(l) > 3), "(empty)")
    print(f"  Slide {num}: {head[:80]}")
