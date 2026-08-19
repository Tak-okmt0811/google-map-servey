# -*- mode: python ; coding: utf-8 -*-
#
# Windows向けexeビルド用スペックファイル（onefile: 単一exeに全て同梱）。
# ビルドはWindows環境で実行してください（PyInstallerはクロスコンパイル不可）。
#
#   uv sync --group dev
#   uv run pyinstaller viewer/build_exe.spec
#
# 出力: dist/competitor-dashboard.exe （単一ファイル。_internal 等のフォルダは生成されない）
#
# 配布方法:
#   competitor-dashboard.exe を1つだけクライアントに渡す。
#   同じフォルダ内、または同フォルダ内の input/ サブフォルダに
#   collector/export.py が出力したCSV/Excelを置けば起動時に自動検出される。
#   APIキーはこのexeに一切含まれず、実行時にもGoogleへ通信しない。
#
# 注意:
#   - onefileは起動のたびに一時フォルダへ展開するため、初回表示まで数秒かかる。
#   - console=False のためコマンドプロンプトは表示されない。アプリを終了する際は
#     ブラウザタブを閉じるだけでは裏のプロセスが残るため、タスクバーの通知領域や
#     タスクマネージャーからの終了が必要（launcher.pyの制約）。

import os

from PyInstaller.utils.hooks import collect_all

HERE = os.path.dirname(os.path.abspath(SPEC))

datas = [(os.path.join(HERE, "app.py"), ".")]
binaries = []
hiddenimports = []

for pkg in ("streamlit", "pydeck", "plotly", "altair"):
    pkg_datas, pkg_binaries, pkg_hiddenimports = collect_all(pkg)
    datas += pkg_datas
    binaries += pkg_binaries
    hiddenimports += pkg_hiddenimports

a = Analysis(
    [os.path.join(HERE, "launcher.py")],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="competitor-dashboard",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    # コマンドプロンプトを表示しない。致命的エラーはlauncher.py側でダイアログ表示する。
    console=False,
)
