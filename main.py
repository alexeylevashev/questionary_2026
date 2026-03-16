"""Entry point for PyInstaller executable."""
import sys
import os

# Force UTF-8 for console I/O on Windows (avoids cp1251 UnicodeEncodeError)
if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if sys.stderr and hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# Ensure the project root is on sys.path when running as a frozen exe
if getattr(sys, "frozen", False):
    # Running inside PyInstaller bundle
    bundle_dir = sys._MEIPASS
    project_dir = os.path.dirname(sys.executable)
    sys.path.insert(0, bundle_dir)
    # Change working dir so relative paths in config.yaml work
    os.chdir(project_dir)

from src.cli import main

if __name__ == "__main__":
    main()
