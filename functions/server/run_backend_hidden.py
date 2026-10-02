from __future__ import annotations

import os
import runpy
import sys
from datetime import datetime


SERVER_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(SERVER_DIR)
LOG_PATH = os.path.join(SERVER_DIR, "backend_hidden.log")


def main() -> None:
    os.chdir(ROOT_DIR)
    if SERVER_DIR not in sys.path:
        sys.path.insert(0, SERVER_DIR)

    log = open(LOG_PATH, "a", encoding="utf-8", buffering=1)
    sys.stdout = log
    sys.stderr = log
    print(f"\n[{datetime.now().isoformat(timespec='seconds')}] backend start", flush=True)
    runpy.run_path(os.path.join(SERVER_DIR, "main.py"), run_name="__main__")


if __name__ == "__main__":
    main()
