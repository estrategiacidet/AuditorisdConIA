from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def main() -> int:
    app_path = Path(__file__).with_name("app.py")
    result = subprocess.run(
        [sys.executable, "-m", "streamlit", "run", str(app_path)],
        check=False,
    )
    return result.returncode


if __name__ == "__main__":
    raise SystemExit(main())
