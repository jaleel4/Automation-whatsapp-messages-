import os
import sys
import runpy
import traceback
import requests

# Force PyInstaller to bundle all libraries needed by quiz_forwarder.py
import pandas
import openpyxl
from playwright.sync_api import sync_playwright

# ================= GITHUB AUTO-UPDATE CONFIG =================
# Paste your GitHub "Raw" URL for quiz_forwarder.py inside the quotes below.
# If left empty (""), it will automatically run your local quiz_forwarder.py file.
GITHUB_RAW_URL = ""

APP_DATA_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", r"C:\EnglishCourseAutomation"),
    "EnglishCourseAutomation"
)
os.makedirs(APP_DATA_DIR, exist_ok=True)
os.makedirs(r"C:\EnglishCourseAutomation", exist_ok=True)

LOCAL_SCRIPT_PATH = os.path.join(APP_DATA_DIR, "quiz_forwarder.py")
FALLBACK_SCRIPT_PATH = r"C:\EnglishCourseAutomation\quiz_forwarder.py"
# =============================================================


def sync_latest_script():
    print("Checking for Quiz Forwarder updates...")
    if GITHUB_RAW_URL.strip():
        try:
            response = requests.get(GITHUB_RAW_URL.strip(), timeout=15)
            if response.status_code == 200 and "def run_quiz_automation" in response.text:
                with open(LOCAL_SCRIPT_PATH, "w", encoding="utf-8") as f:
                    f.write(response.text)
                print("✅ Downloaded latest quiz_forwarder.py from GitHub.\n")
                return LOCAL_SCRIPT_PATH
        except Exception as e:
            print(f"⚠️ Could not reach GitHub ({e}). Using local version.\n")

    if os.path.exists(FALLBACK_SCRIPT_PATH):
        print("✅ System up to date.\n")
        return FALLBACK_SCRIPT_PATH
    elif os.path.exists(LOCAL_SCRIPT_PATH):
        print("✅ System up to date.\n")
        return LOCAL_SCRIPT_PATH
    else:
        raise FileNotFoundError(
            f"Could not find quiz_forwarder.py in {FALLBACK_SCRIPT_PATH} or {LOCAL_SCRIPT_PATH}."
        )


if __name__ == "__main__":
    try:
        target_script = sync_latest_script()
        if APP_DATA_DIR not in sys.path:
            sys.path.insert(0, APP_DATA_DIR)
        runpy.run_path(target_script, run_name="__main__")
    except Exception:
        print("\n❌ Critical Error launching the Quiz Forwarder engine:")
        traceback.print_exc()
        input("\nPress Enter to exit...")