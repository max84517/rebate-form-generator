# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec — onedir build for Rebate Form Generator."""

import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# Collect customtkinter assets (themes, images, fonts)
ctk_datas = collect_data_files("customtkinter")
docx_datas = collect_data_files("docx")

a = Analysis(
    ["src/rebate_form_generator/main.py"],
    pathex=[],
    binaries=[],
    datas=ctk_datas + docx_datas,
    hiddenimports=[
        "customtkinter",
        "PIL._tkinter_finder",
        "docx",
        "lxml",
        "lxml.etree",
        "lxml._elementpath",
        "openpyxl",
        "openpyxl.cell._writer",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="RebateFormGenerator",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,           # no console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="RebateFormGenerator",
)
