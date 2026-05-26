"""
X02_pack_area_to_leap.py
Inverse of S01_extract_leap.py.
Take a LEAP area folder and package it as an AES-encrypted .leap file
(password "LEAP", same format LEAP uses for installable area archives).

Usage:
    python scripts/04_run/X02_pack_area_to_leap.py "C:\\Users\\User\\Documents\\LEAP_16_04\\LEAP Areas\\kaz_workshop exercise\\KAZ_2024" data/snapshots/cycle_009_KAZ_2024_current.leap

Requirements:
    pip install pyzipper

What it does:
    - Walks the area folder recursively
    - Adds each file to an AES-encrypted ZIP, password "LEAP"
    - Skips known-noise files: *.bak_cycle*, *.sqlite-shm, *.sqlite-wal, *.lock
    - Preserves relative paths inside the archive

What LEAP does on install:
    - Decrypts with password "LEAP"
    - Extracts to its LEAP Areas folder
    - Area is browsable in LEAP UI thereafter
"""

import sys
import os
import argparse
from pathlib import Path

try:
    import pyzipper
except ImportError:
    print("ERROR: pyzipper not installed.")
    print("Run: pip install pyzipper")
    sys.exit(1)


LEAP_PASSWORD = b"LEAP"

# Files that are runtime/project artifacts and should not be packaged
EXCLUDE_PATTERNS = [
    ".bak_cycle",       # our backup files
    ".sqlite-shm",      # SQLite shared memory (runtime)
    ".sqlite-wal",      # SQLite write-ahead log (runtime)
    ".lock",            # any lock files
    "Thumbs.db",        # Windows thumbnail cache
    ".DS_Store",        # macOS metadata
]


def should_skip(name: str) -> bool:
    for pat in EXCLUDE_PATTERNS:
        if pat in name:
            return True
    return False


def pack_area(area_dir: Path, out_path: Path) -> dict:
    """Pack area_dir into an AES-encrypted .leap file at out_path."""
    out_path.parent.mkdir(parents=True, exist_ok=True)

    summary = {
        "added": 0,
        "skipped": 0,
        "total_uncompressed_bytes": 0,
        "skipped_files": [],
        "added_files": [],
    }

    area_dir = area_dir.resolve()

    with pyzipper.AESZipFile(
        out_path,
        mode="w",
        compression=pyzipper.ZIP_DEFLATED,
        encryption=pyzipper.WZ_AES,
    ) as zf:
        zf.setpassword(LEAP_PASSWORD)
        zf.setencryption(pyzipper.WZ_AES, nbits=256)

        for root, dirs, files in os.walk(area_dir):
            root_path = Path(root)
            for fname in files:
                file_path = root_path / fname
                rel_path = file_path.relative_to(area_dir)
                rel_str = str(rel_path).replace("\\", "/")

                if should_skip(rel_str):
                    summary["skipped"] += 1
                    summary["skipped_files"].append(rel_str)
                    continue

                try:
                    size = file_path.stat().st_size
                    zf.write(file_path, arcname=rel_str)
                    summary["added"] += 1
                    summary["total_uncompressed_bytes"] += size
                    summary["added_files"].append((rel_str, size))
                except Exception as e:
                    print(f"  ! FAILED to add {rel_str}: {e}")

    return summary


def main():
    ap = argparse.ArgumentParser(description="Pack a LEAP area folder into an AES-encrypted .leap file")
    ap.add_argument("area_dir", help="Path to area folder (e.g. ...\\LEAP Areas\\KAZ_2024)")
    ap.add_argument("out_leap", help="Output .leap path")
    args = ap.parse_args()

    area_dir = Path(args.area_dir).resolve()
    out_path = Path(args.out_leap).resolve()

    if not area_dir.exists() or not area_dir.is_dir():
        print(f"ERROR: {area_dir} is not a directory")
        sys.exit(1)

    print(f"Packing area : {area_dir}")
    print(f"        to   : {out_path}")
    print()

    summary = pack_area(area_dir, out_path)

    archive_size = out_path.stat().st_size
    print(f"Done.")
    print(f"  Files added            : {summary['added']}")
    print(f"  Files skipped (runtime): {summary['skipped']}")
    print(f"  Uncompressed total     : {summary['total_uncompressed_bytes'] / 1024 / 1024:.1f} MB")
    print(f"  Archive size           : {archive_size / 1024 / 1024:.1f} MB")
    print(f"  Output                 : {out_path}")
    if summary["skipped_files"]:
        print()
        print("Skipped files:")
        for s in summary["skipped_files"][:20]:
            print(f"  - {s}")
        if len(summary["skipped_files"]) > 20:
            print(f"  ... {len(summary['skipped_files']) - 20} more")


if __name__ == "__main__":
    main()
