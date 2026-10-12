#!/usr/bin/env python3
"""docs/manual.md（画像つき）から docs/manual.pdf を作る。

マニュアルを更新したら、このスクリプトでPDFを作り直してcommitする。
markdown ライブラリと Google Chrome を使う（プロジェクトの依存関係には追加しない）。

    uv run --with markdown python scripts/build_manual_pdf.py

Chromeの場所が違う場合は、環境変数 CHROME_PATH で指定する。
"""

from __future__ import annotations

import os
import subprocess
import sys
import tempfile
from pathlib import Path

import markdown

ROOT = Path(__file__).resolve().parent.parent
DOCS = ROOT / "docs"
SOURCE = DOCS / "manual.md"
OUTPUT = DOCS / "manual.pdf"

CHROME_CANDIDATES = [
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/usr/bin/google-chrome",
    "/usr/bin/chromium",
    "/usr/bin/chromium-browser",
]

CSS = """
@page { size: A4; margin: 16mm 14mm; }
body {
  font-family: "Hiragino Sans", "Hiragino Kaku Gothic ProN", "Noto Sans JP", "Yu Gothic", sans-serif;
  font-size: 10pt; line-height: 1.75; color: #222;
}
h1 { font-size: 20pt; border-bottom: 3px solid #e53935; padding-bottom: 6px; margin: 0 0 14px; }
h2 { font-size: 15pt; margin: 26px 0 10px; padding-left: 10px; border-left: 5px solid #e53935; break-after: avoid; }
h2.chapter { break-before: page; margin-top: 0; }
h3 { font-size: 11.5pt; margin: 18px 0 6px; break-after: avoid; }
p { margin: 6px 0; }
img { max-width: 100%; border: 1px solid #ddd; break-inside: avoid; }
table { border-collapse: collapse; width: 100%; margin: 8px 0 12px; font-size: 9pt; break-inside: avoid; }
th, td { border: 1px solid #ccc; padding: 4px 7px; vertical-align: top; text-align: left; }
th { background: #f3f4f6; }
td img { border: none; }
blockquote { margin: 10px 0; padding: 6px 12px; background: #fff5f5; border-left: 4px solid #e57373; break-inside: avoid; }
blockquote p { margin: 3px 0; }
code { font-family: Menlo, monospace; font-size: 8.5pt; background: #f3f4f6; padding: 1px 4px; border-radius: 3px; }
pre { background: #f3f4f6; padding: 8px 10px; border-radius: 4px; overflow-wrap: anywhere; white-space: pre-wrap; break-inside: avoid; }
pre code { background: none; padding: 0; }
hr { display: none; }  /* 章ごとに改ページするので区切り線は不要（ページ末に残ると空白ページになる） */
a { color: #1565c0; text-decoration: none; }
"""


def find_chrome() -> str:
    env = os.environ.get("CHROME_PATH")
    candidates = ([env] if env else []) + CHROME_CANDIDATES
    for path in candidates:
        if path and Path(path).exists():
            return path
    sys.exit("Google Chrome が見つかりません。環境変数 CHROME_PATH で場所を指定してください。")


def build_html() -> str:
    body = markdown.markdown(
        SOURCE.read_text(encoding="utf-8"),
        extensions=["tables", "fenced_code", "toc", "sane_lists"],
    )
    # 「## 1. 〜」以降の章は、新しいページから始める（先頭の目次は除く）
    parts = body.split("<h2 ")
    body = parts[0] + "".join(
        ("<h2 class=\"chapter\" " if i >= 2 else "<h2 ") + part for i, part in enumerate(parts[1:], start=1)
    )
    return (
        "<!doctype html><html lang=\"ja\"><head><meta charset=\"utf-8\">"
        f"<base href=\"{DOCS.as_uri()}/\"><title>競合分析ダッシュボード 操作マニュアル</title>"
        f"<style>{CSS}</style></head><body>{body}</body></html>"
    )


def main() -> None:
    chrome = find_chrome()
    with tempfile.TemporaryDirectory() as tmp:
        html_path = Path(tmp) / "manual.html"
        html_path.write_text(build_html(), encoding="utf-8")
        subprocess.run(
            [
                chrome,
                "--headless=new",
                "--disable-gpu",
                "--no-pdf-header-footer",
                f"--user-data-dir={tmp}/profile",
                "--virtual-time-budget=15000",
                f"--print-to-pdf={OUTPUT}",
                html_path.as_uri(),
            ],
            check=True,
            capture_output=True,
        )
    print(f"wrote {OUTPUT} ({OUTPUT.stat().st_size / 1024 / 1024:.1f} MB)")


if __name__ == "__main__":
    main()
