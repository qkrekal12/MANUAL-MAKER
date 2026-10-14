"""Streamlit Cloud and CLI entry point. Configure this file as the app file."""
from pathlib import Path
import runpy
import sys


APP_DIR = Path(__file__).resolve().parent
APP_FILE = APP_DIR / "app.py"
if not APP_FILE.is_file():
    raise FileNotFoundError(f"Streamlit app module was not found: {APP_FILE}")

# Execute the app source on every Streamlit rerun (an ordinary import is cached).
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))
runpy.run_path(str(APP_FILE), run_name="__main__")
