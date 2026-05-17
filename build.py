"""
Build script: compiles pph23_extractor.py and ppn_extractor.py into
standalone Windows .exe files using Nuitka.

Usage (on Windows, inside the project folder):
    pip install nuitka
    python build.py
"""

import subprocess
import sys
from pathlib import Path

TARGETS = [
    "pph23_extractor.py",
    "ppn_extractor.py",
]

# Packages that must be bundled (pdfplumber pulls in pdfminer + PIL)
INCLUDE_PACKAGES = [
    "pdfplumber",
    "pdfminer",
    "openpyxl",
    "PIL",
]

OUTPUT_DIR = "dist"


def build(target: str):
    name = Path(target).stem
    print(f"\n{'='*60}")
    print(f"Building: {target}  →  {OUTPUT_DIR}/{name}.exe")
    print("=" * 60)

    cmd = [
        sys.executable, "-m", "nuitka",
        "--onefile",
        "--windows-console-mode=force",
        f"--output-dir={OUTPUT_DIR}",
        f"--output-filename={name}.exe",
    ]

    for pkg in INCLUDE_PACKAGES:
        cmd.append(f"--include-package={pkg}")

    # openpyxl ships data files (styles, templates) that must travel with the exe
    cmd.append("--include-package-data=openpyxl")

    cmd.append(target)

    result = subprocess.run(cmd)
    if result.returncode != 0:
        print(f"\nERROR: build failed for {target} (exit code {result.returncode})")
        sys.exit(result.returncode)

    print(f"\nOK: {OUTPUT_DIR}/{name}.exe")


if __name__ == "__main__":
    Path(OUTPUT_DIR).mkdir(exist_ok=True)
    for t in TARGETS:
        build(t)
    print(f"\nDone. Executables are in ./{OUTPUT_DIR}/")
