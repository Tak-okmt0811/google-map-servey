"""PyInstaller用の起動エントリポイント。

`streamlit run app.py` を直接呼べないexe環境向けに、Streamlitランタイムを
プログラム的に起動する薄いラッパー。app.py自体はAPIキーを一切参照しない。
"""

from __future__ import annotations

import sys
from pathlib import Path


def resolve_app_path() -> str:
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidate = Path(meipass) / "app.py"
        if candidate.exists():
            return str(candidate)

    if getattr(sys, "frozen", False):
        candidate = Path(sys.executable).resolve().parent / "app.py"
        if candidate.exists():
            return str(candidate)

    return str(Path(__file__).resolve().parent / "app.py")


def main() -> None:
    from streamlit.web import cli as stcli

    sys.argv = [
        "streamlit",
        "run",
        resolve_app_path(),
        "--global.developmentMode=false",
        "--browser.gatherUsageStats=false",
    ]
    sys.exit(stcli.main())


if __name__ == "__main__":
    main()
