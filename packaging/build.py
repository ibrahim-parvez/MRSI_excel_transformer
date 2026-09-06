"""Build the app and the installer, and tidy up after them.

Replaces the pile of pyinstaller command lines this project used to keep in
a text file. The per-target options now live in the .spec files next to this
script, so a build is the same command on every machine:

    python packaging/build.py app          # the tool itself
    python packaging/build.py installer    # the standalone installer
    python packaging/build.py all
    python packaging/build.py clean        # drop build/, dist/ and __pycache__

Output lands in dist/ at the project root: a single .exe on Windows, a .app
bundle on macOS. Run it from anywhere; paths are resolved off this file.
"""

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

PACKAGING_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = PACKAGING_DIR.parent
BUILD_DIR = PROJECT_ROOT / "build"
DIST_DIR = PROJECT_ROOT / "dist"

SPECS = {
    "app": PACKAGING_DIR / "app.spec",
    "installer": PACKAGING_DIR / "installer.spec",
}


def run_pyinstaller(spec: Path) -> None:
    print(f"\n=== Building {spec.name} ===")
    subprocess.run(
        [
            sys.executable,
            "-m",
            "PyInstaller",
            "--noconfirm",
            "--clean",
            "--workpath",
            str(BUILD_DIR),
            "--distpath",
            str(DIST_DIR),
            str(spec),
        ],
        cwd=PROJECT_ROOT,
        check=True,
    )


def clean() -> None:
    for directory in (BUILD_DIR, DIST_DIR):
        if directory.exists():
            print(f"Removing {directory.relative_to(PROJECT_ROOT)}/")
            shutil.rmtree(directory)

    for cache in PROJECT_ROOT.rglob("__pycache__"):
        if ".git" not in cache.parts:
            shutil.rmtree(cache, ignore_errors=True)
    print("Removed __pycache__ directories.")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "target",
        choices=["app", "installer", "all", "clean"],
        help="what to build, or 'clean' to remove build output",
    )
    args = parser.parse_args()

    if args.target == "clean":
        clean()
        return

    targets = list(SPECS) if args.target == "all" else [args.target]
    for target in targets:
        run_pyinstaller(SPECS[target])

    print(f"\nDone. Output is in {DIST_DIR}")


if __name__ == "__main__":
    main()
