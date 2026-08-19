"""PyInstaller用の起動エントリポイント。

`streamlit run app.py` を直接呼べないexe環境向けに、Streamlitランタイムを
プログラム的に起動する薄いラッパー。app.py自体はAPIキーを一切参照しない。
"""

from __future__ import annotations

import sys
from pathlib import Path

# Streamlitは初回起動時、~/.streamlit/credentials.toml が無いと
# 「メールアドレスを入力してください」という対話プロンプトを標準入力待ちで
# 表示する。コンソールを持たない配布exeではこれが「固まった」ように見えるため、
# 事前にファイルを用意してプロンプト自体を発生させない。
DEFAULT_CREDENTIALS_EMAIL = "nightbar.lumel@gmail.com"


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


def ensure_streamlit_credentials() -> None:
    creds_path = Path.home() / ".streamlit" / "credentials.toml"
    if creds_path.exists():
        return
    creds_path.parent.mkdir(parents=True, exist_ok=True)
    creds_path.write_text(
        f'[general]\nemail = "{DEFAULT_CREDENTIALS_EMAIL}"\n',
        encoding="utf-8",
    )


def show_fatal_error(message: str) -> None:
    """コンソールを隠しているため、致命的エラーはネイティブダイアログで通知する。"""
    if sys.platform == "win32":
        import ctypes

        MB_ICONERROR = 0x10
        ctypes.windll.user32.MessageBoxW(
            0, message, "競合分析ダッシュボード - 起動エラー", MB_ICONERROR
        )
    else:
        print(message, file=sys.stderr)


def main() -> None:
    try:
        ensure_streamlit_credentials()

        from streamlit.web import cli as stcli

        sys.argv = [
            "streamlit",
            "run",
            resolve_app_path(),
            "--global.developmentMode=false",
            "--browser.gatherUsageStats=false",
        ]
        sys.exit(stcli.main())
    except SystemExit:
        raise
    except Exception as exc:  # noqa: BLE001 - 非エンジニア向けに必ず画面で知らせる
        show_fatal_error(f"アプリの起動に失敗しました。\n\n{exc}")
        raise


if __name__ == "__main__":
    main()
