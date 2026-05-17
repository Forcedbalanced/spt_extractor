"""
Build script: compiles app.py, pph23_extractor.py, and ppn_extractor.py
into standalone Windows .exe files using Nuitka.

Usage (on Windows, inside the project folder):
    pip install nuitka
    python build.py
"""

import subprocess
import sys
from pathlib import Path

# Packages that must be bundled
INCLUDE_PACKAGES = [
    "pdfplumber",
    "pdfminer",
    "openpyxl",
    "PIL",
]

# Windows EXE metadata (visible in right-click > Properties > Details)
METADATA = {
    "company-name":   "Maulana Automation Consultant",
    "product-name":   "SPT Extractor",
    "file-version":   "1.0.0.0",
    "product-version": "1.0.0.0",
    "copyright":      "© 2026 Achmad Maulana — Maulana Automation Consultant",
}

OUTPUT_DIR = "dist"

# (target, output_name, is_gui)
TARGETS = [
    ("app.py", "SPT_Extractor", True),
]


def build(target: str, output_name: str, is_gui: bool):
    print(f"\n{'='*60}")
    print(f"Building: {target}  ->  {OUTPUT_DIR}/{output_name}.exe")
    print("=" * 60)

    cmd = [
        sys.executable, "-m", "nuitka",
        "--onefile",
        "--assume-yes-for-downloads",
        f"--output-dir={OUTPUT_DIR}",
        f"--output-filename={output_name}.exe",
    ]

    # GUI apps hide the console; CLI tools keep it
    if is_gui:
        cmd.append("--windows-console-mode=disable")
        cmd.append("--enable-plugin=tk-inter")
    else:
        cmd.append("--windows-console-mode=force")

    for pkg in INCLUDE_PACKAGES:
        cmd.append(f"--include-package={pkg}")

    # openpyxl ships data files (styles, templates) that must travel with the exe
    cmd.append("--include-package-data=openpyxl")

    # When building app.py, bundle the extractor modules too
    if target == "app.py":
        cmd.append("--include-module=pph23_extractor")
        cmd.append("--include-module=ppn_extractor")

    # Windows metadata
    for key, value in METADATA.items():
        cmd.append(f"--{key}={value}")

    # Set description per target
    descriptions = {
        "app.py":            "SPT Extractor GUI (PPN & PPh 23)",
        "pph23_extractor.py": "SPT PPh 23 PDF Extractor (CLI)",
        "ppn_extractor.py":   "SPT PPN PDF Extractor (CLI)",
    }
    cmd.append(f"--file-description={descriptions[target]}")

    cmd.append(target)

    result = subprocess.run(cmd)
    if result.returncode != 0:
        print(f"\nERROR: build failed for {target} (exit code {result.returncode})")
        sys.exit(result.returncode)

    print(f"\nOK: {OUTPUT_DIR}/{output_name}.exe")


if __name__ == "__main__":
    Path(OUTPUT_DIR).mkdir(exist_ok=True)
    for t, name, gui in TARGETS:
        build(t, name, gui)
    print(f"\nDone. Executables are in ./{OUTPUT_DIR}/")
