"""
S01_extract_leap.py
Decrypt a .leap file (AES-encrypted ZIP, password "LEAP") and dump its contents
to data/extracted/ for inspection.

Usage:
    python scripts/01_scout/S01_extract_leap.py data/snapshots/cycle_000_colleague_baseline.leap

Output:
    - data/extracted/<filename_without_ext>/  (all files inside the .leap archive)
    - data/audit_reports/extract_<filename>_<timestamp>.md (summary report)

Requirements:
    pip install pyzipper

Idempotent: re-running overwrites previous extraction.
"""

import sys
import os
import argparse
from pathlib import Path
from datetime import datetime

try:
    import pyzipper
except ImportError:
    print("ERROR: pyzipper not installed.")
    print("Run: pip install pyzipper")
    sys.exit(1)


LEAP_PASSWORD = b"LEAP"


def extract_leap(leap_path: Path, out_dir: Path):
    """Extract all files from a .leap archive."""
    out_dir.mkdir(parents=True, exist_ok=True)

    text_files = []  # files we can read as text
    binary_files = []  # NexusDB tables, etc.
    sub_archives = []  # nested zip/xlsx files
    summary = {
        "total_files": 0,
        "total_size_bytes": 0,
        "vbs_files": [],
        "txt_files": [],
        "nx1_files": [],  # NexusDB tables
        "xlsx_files": [],  # availability data
        "other": [],
    }

    with pyzipper.AESZipFile(leap_path) as zf:
        zf.pwd = LEAP_PASSWORD
        for info in zf.infolist():
            if info.is_dir():
                continue
            summary["total_files"] += 1
            summary["total_size_bytes"] += info.file_size

            name = info.filename
            ext = Path(name).suffix.lower()

            if ext == ".vbs_safe":
                summary["vbs_files"].append((name, info.file_size))
            elif ext == ".txt" or "ini" in name.lower():
                summary["txt_files"].append((name, info.file_size))
            elif ext == ".nx1":
                summary["nx1_files"].append((name, info.file_size))
            elif ext in (".xlsx", ".xls"):
                summary["xlsx_files"].append((name, info.file_size))
            else:
                summary["other"].append((name, info.file_size))

            try:
                zf.extract(name, out_dir)
            except Exception as e:
                print(f"  ! FAILED to extract {name}: {e}")

    return summary


def write_report(leap_path: Path, summary: dict, report_path: Path):
    """Write a markdown summary of what's in the archive."""
    report_path.parent.mkdir(parents=True, exist_ok=True)

    lines = []
    lines.append(f"# Extraction Report — {leap_path.name}")
    lines.append("")
    lines.append(f"**Extracted:** {datetime.now().isoformat()}")
    lines.append(f"**Source file:** `{leap_path}`")
    lines.append(f"**Source size:** {leap_path.stat().st_size / 1024 / 1024:.1f} MB")
    lines.append(f"**Files inside:** {summary['total_files']}")
    lines.append(f"**Total uncompressed:** {summary['total_size_bytes'] / 1024 / 1024:.1f} MB")
    lines.append("")

    lines.append("## VBScript hooks (calculation customizations)")
    lines.append("")
    if summary["vbs_files"]:
        lines.append("| File | Size (bytes) |")
        lines.append("|---|---|")
        for name, size in sorted(summary["vbs_files"], key=lambda x: -x[1]):
            lines.append(f"| `{name}` | {size:,} |")
    else:
        lines.append("_None_")
    lines.append("")

    lines.append("## Configuration / text files")
    lines.append("")
    lines.append("| File | Size (bytes) |")
    lines.append("|---|---|")
    for name, size in sorted(summary["txt_files"], key=lambda x: -x[1]):
        lines.append(f"| `{name}` | {size:,} |")
    lines.append("")

    lines.append("## NexusDB tables (model data)")
    lines.append("")
    lines.append(f"_{len(summary['nx1_files'])} table files. The big ones contain the actual model:_")
    lines.append("")
    lines.append("| File | Size (bytes) | Likely purpose |")
    lines.append("|---|---|---|")
    table_hints = {
        "datanames": "branch / variable name strings",
        "deviceparents": "branch hierarchy parent-child links",
        "hiddenbranches": "branches hidden from views",
        "branchtags": "branch tags / categories",
        "fuels": "fuel definitions",
        "tsprofiles": "time-slice profiles",
        "tsprofiledata": "time-slice profile data",
        "uservardata": "user variable data values",
        "Diagnostics": "model error diagnostics",
        "Constraints": "constraint definitions",
        "DeviceBranchTable": "device-to-branch mapping",
        "allbutdeviceparents": "non-device branches",
        "units": "unit definitions",
        "variabledata": "variable value data",
    }
    for name, size in sorted(summary["nx1_files"], key=lambda x: -x[1]):
        stem = Path(name).stem
        hint = table_hints.get(stem, "")
        lines.append(f"| `{name}` | {size:,} | {hint} |")
    lines.append("")

    if summary["xlsx_files"]:
        lines.append("## Embedded Excel files (hydro availability data)")
        lines.append("")
        lines.append(f"_{len(summary['xlsx_files'])} XLSX files. These are seasonal availability profiles for hydro plants._")
        lines.append("")

    if summary["other"]:
        lines.append("## Other files")
        lines.append("")
        lines.append("| File | Size (bytes) |")
        lines.append("|---|---|")
        for name, size in sorted(summary["other"], key=lambda x: -x[1])[:20]:
            lines.append(f"| `{name}` | {size:,} |")
        if len(summary["other"]) > 20:
            lines.append(f"| _...{len(summary['other']) - 20} more not shown_ | |")
        lines.append("")

    lines.append("---")
    lines.append("")
    lines.append("## Next steps")
    lines.append("")
    lines.append("1. Read `beforeCalculation.vbs_Safe` — this is where the nodal distribution bug lives")
    lines.append("2. Read `AreaSettingsINI.txt` — confirm BaseYear / FirstScenarioYear / EndYear")
    lines.append("3. Run `S02_audit_model.ps1` against the LIVE area in LEAP for COM-based audit")

    report_path.write_text("\n".join(lines), encoding="utf-8")


def main():
    ap = argparse.ArgumentParser(description="Decrypt and extract a LEAP .leap file")
    ap.add_argument("leap_file", help="Path to .leap file")
    ap.add_argument("--out-dir", default=None, help="Output directory (default: data/extracted/<name>/)")
    ap.add_argument("--report-dir", default="data/audit_reports", help="Where to write summary report")
    args = ap.parse_args()

    leap_path = Path(args.leap_file).resolve()
    if not leap_path.exists():
        print(f"ERROR: {leap_path} not found")
        sys.exit(1)

    project_root = Path.cwd()
    out_dir = (Path(args.out_dir) if args.out_dir
               else project_root / "data" / "extracted" / leap_path.stem)

    print(f"Extracting: {leap_path}")
    print(f"      to  : {out_dir}")
    print()

    summary = extract_leap(leap_path, out_dir)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = project_root / args.report_dir / f"extract_{leap_path.stem}_{timestamp}.md"
    write_report(leap_path, summary, report_path)

    print(f"Done. {summary['total_files']} files extracted.")
    print(f"Report: {report_path}")
    print()
    print("Quick stats:")
    print(f"  VBScript hooks : {len(summary['vbs_files'])}")
    print(f"  Text/INI files : {len(summary['txt_files'])}")
    print(f"  NexusDB tables : {len(summary['nx1_files'])}")
    print(f"  Excel files    : {len(summary['xlsx_files'])}")
    print(f"  Other          : {len(summary['other'])}")


if __name__ == "__main__":
    main()
