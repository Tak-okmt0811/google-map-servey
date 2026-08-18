# -*- mode: python ; coding: utf-8 -*-
#
# Windows向けexeビルド用スペックファイル。
# ビルドはWindows環境で実行してください（PyInstallerはクロスコンパイル不可）。
#
#   uv sync --group dev
#   uv run pyinstaller viewer/build_exe.spec
#
# 出力: dist/competitor-dashboard/competitor-dashboard.exe （フォルダ一式）
#
# 配布方法:
#   dist/competitor-dashboard/ フォルダをまるごとクライアントに渡す。
#   同フォルダ直下、または同フォルダ内の input/ サブフォルダに
#   collector/export.py が出力したCSV/Excelを置けば起動時に自動検出される。
#   APIキーはこのexeに一切含まれず、実行時にもGoogleへ通信しない。

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
    [],
    exclude_binaries=True,
    name="competitor-dashboard",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    # 初回ビルド時はTrueのままにし、エラーがコンソールに出るようにしておくと
    # 動作確認しやすい。安定してからFalseに切り替えても良い。
    console=True,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="competitor-dashboard",
)
