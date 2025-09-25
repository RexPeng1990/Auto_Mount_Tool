# PyInstaller 配置文件
# -*- mode: python ; coding: utf-8 -*-

import os
import random
import string

# 生成隨機密鑰用於加密
key = ''.join(random.choices(string.ascii_letters + string.digits, k=16))

block_cipher = None

a = Analysis(
    ['main_obfuscated.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.filedialog',
        'subprocess',
        'threading',
        'os',
        'sys',
        'pathlib',
        'base64',
        'codecs'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'PIL',
        'pygame'
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# 過濾掉不需要的 Python 庫以減小體積
a.binaries = [x for x in a.binaries if not x[0].startswith('api-ms-win')]
a.binaries = [x for x in a.binaries if not x[0].startswith('ucrtbase')]

pyz = PYZ(
    a.pure, 
    a.zipped_data,
    cipher=block_cipher
)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='WIM_Driver_Manager',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,  # 使用 UPX 壓縮
    upx_exclude=[],
    console=False,  # 隱藏控制台窗口
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    uac_admin=True,  # 請求管理員權限
    icon='icon.ico' if os.path.exists('icon.ico') else None,
    version_file='version.txt' if os.path.exists('version.txt') else None,
)
