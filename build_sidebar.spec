# -*- mode: python ; coding: utf-8 -*-

import os
import sys
from pathlib import Path

block_cipher = None

# 專案根目錄
BASE_DIR = Path(r"g:\我的雲端硬碟\Development\Codeing\Python\Auto_Mount_Tool").resolve()

# 收集所有資料檔案
datas = [
    # 設定檔
    (str(BASE_DIR / "settings.ini"), "."),
    (str(BASE_DIR / "config.ini"), "."),
]

# 檢查 icon 是否存在
icon_path = BASE_DIR / "icon.ico"
if icon_path.exists():
    datas.append((str(icon_path), "."))

# CustomTkinter 資源
import customtkinter
ctk_path = Path(customtkinter.__file__).parent
datas.append((str(ctk_path), "customtkinter"))

a = Analysis(
    [r"g:\我的雲端硬碟\Development\Codeing\Python\Auto_Mount_Tool\main_sidebar.py"],
    pathex=[str(BASE_DIR)],
    binaries=[],
    datas=datas,
    hiddenimports=['tkinter', 'tkinter.ttk', 'tkinter.messagebox', 'tkinter.filedialog', 'customtkinter', 'subprocess', 'threading', 'pathlib', 'configparser', 'shutil', 'ctypes', 'datetime', 're', 'json'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'numpy', 'pandas', 'scipy', 'PIL'],
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
    name='WIM_Driver_Manager',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(icon_path) if icon_path.exists() else None,
    uac_admin=True,  # 要求管理員權限
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='WIM_Driver_Manager',
)
