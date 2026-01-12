# !/usr/bin/env python3
"""
WIM 管理工具打包腳本（側邊欄版本）
使用 Cython 編譯核心模組 + PyInstaller 打包
可在其他電腦上使用
"""

import os
import sys
import subprocess
import shutil
import tempfile
from pathlib import Path

# 專案設定
PROJECT_NAME = "WIM_Driver_Manager"
MAIN_SCRIPT = "main.py"
VERSION = "3.0"

# 目錄設定
BASE_DIR = Path(__file__).parent
BUILD_DIR = BASE_DIR / "build_output"
DIST_DIR = BASE_DIR / "dist_sidebar"
TEMP_DIR = BASE_DIR / "temp_build"

# 需要編譯成 Cython 的核心模組
CYTHON_MODULES = [
    "app/wim_manager.py",
    "app/driver_manager.py",
    "app/config.py",
]

# 需要打包的 Python 模組（不編譯成 Cython）
PYTHON_MODULES = [
    "ui/theme.py",
    "ui/components.py",
    "ui/sidebar.py",
    "ui/collapsible.py",
    "ui/log_panel.py",
    "ui/wim_slot_card.py",
    "ui/pages/wim_page.py",
    "ui/pages/driver_page.py",
    "ui/pages/script_page.py",
]


def check_requirements():
    """檢查必要的套件"""
    print("檢查必要套件...")
    
    required = ['pyinstaller', 'customtkinter']
    missing = []
    
    for pkg in required:
        try:
            __import__(pkg.replace('-', '_'))
        except ImportError:
            missing.append(pkg)
    
    if missing:
        print(f"缺少套件: {', '.join(missing)}")
        print("正在安裝...")
        subprocess.run([sys.executable, '-m', 'pip', 'install'] + missing, check=True)
    
    print("✓ 套件檢查完成")


def clean_build():
    """清理舊的建置檔案"""
    print("清理舊的建置檔案...")
    
    dirs_to_clean = [BUILD_DIR, DIST_DIR, TEMP_DIR]
    for d in dirs_to_clean:
        if d.exists():
            shutil.rmtree(d, ignore_errors=True)
    
    # 清理 .pyc 檔案
    for pyc in BASE_DIR.rglob("*.pyc"):
        try:
            pyc.unlink()
        except:
            pass
    
    # 清理 __pycache__
    for cache in BASE_DIR.rglob("__pycache__"):
        try:
            shutil.rmtree(cache)
        except:
            pass
    
    print("✓ 清理完成")


def try_cython_compile():
    """嘗試使用 Cython 編譯核心模組"""
    print("\n嘗試 Cython 編譯...")
    
    try:
        import Cython
        print(f"  Cython 版本: {Cython.__version__}")
    except ImportError:
        print("  ⚠ Cython 未安裝，跳過編譯")
        return False, []
    
    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    compiled_files = []
    
    for module_path in CYTHON_MODULES:
        src = BASE_DIR / module_path
        if not src.exists():
            print(f"  ⚠ 找不到: {module_path}")
            continue
        
        # 複製到臨時目錄
        module_name = Path(module_path).stem
        pyx_file = TEMP_DIR / f"{module_name}.pyx"
        shutil.copy(src, pyx_file)
        
        # 建立 setup.py
        setup_content = f'''
from setuptools import setup
from Cython.Build import cythonize

setup(
    ext_modules=cythonize(
        "{pyx_file.name}",
        compiler_directives={{"language_level": "3"}},
        annotate=False
    )
)
'''
        setup_file = TEMP_DIR / "setup.py"
        setup_file.write_text(setup_content)
        
        # 編譯
        print(f"  編譯: {module_path}")
        result = subprocess.run(
            [sys.executable, str(setup_file), "build_ext", "--inplace"],
            cwd=str(TEMP_DIR),
            capture_output=True,
            text=True
        )
        
        if result.returncode == 0:
            # 找到 .pyd 檔案
            for pyd in TEMP_DIR.glob(f"{module_name}*.pyd"):
                compiled_files.append((module_path, pyd))
                print(f"    ✓ 成功: {pyd.name}")
                break
        else:
            print(f"    ✗ 失敗: {result.stderr[:200] if result.stderr else 'Unknown error'}")
    
    return len(compiled_files) > 0, compiled_files


def build_with_pyinstaller(use_cython=False, cython_files=None):
    """使用 PyInstaller 打包"""
    print("\n使用 PyInstaller 打包...")
    
    # 準備打包目錄
    DIST_DIR.mkdir(parents=True, exist_ok=True)
    
    # 隱藏導入
    hidden_imports = [
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.filedialog',
        'customtkinter',
        'subprocess',
        'threading',
        'pathlib',
        'configparser',
        'shutil',
        'ctypes',
        'datetime',
        're',
        'json',
    ]
    
    # 建立 spec 檔案內容
    spec_content = f'''# -*- mode: python ; coding: utf-8 -*-

import os
import sys
from pathlib import Path

block_cipher = None

# 專案根目錄
BASE_DIR = Path(r"{BASE_DIR}").resolve()

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
    [r"{BASE_DIR / MAIN_SCRIPT}"],
    pathex=[str(BASE_DIR)],
    binaries=[],
    datas=datas,
    hiddenimports={hidden_imports},
    hookspath=[],
    hooksconfig={{}},
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
    name='{PROJECT_NAME}',
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
    name='{PROJECT_NAME}',
)
'''
    
    # 寫入 spec 檔案
    spec_file = BASE_DIR / "build_sidebar.spec"
    spec_file.write_text(spec_content, encoding='utf-8')
    
    # 執行 PyInstaller
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--clean',
        '--noconfirm',
        f'--distpath={DIST_DIR}',
        f'--workpath={BUILD_DIR}',
        str(spec_file),
    ]
    
    print(f"執行: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.returncode != 0:
        print(f"✗ PyInstaller 失敗:")
        print(result.stderr)
        return False
    
    print("✓ PyInstaller 完成")
    return True


def copy_additional_files():
    """複製額外需要的檔案"""
    print("\n複製額外檔案...")
    
    output_dir = DIST_DIR / PROJECT_NAME
    if not output_dir.exists():
        print("✗ 找不到輸出目錄")
        return False
    
    # 複製設定檔（如果不存在）
    for config_file in ["settings.ini", "config.ini"]:
        src = BASE_DIR / config_file
        dst = output_dir / config_file
        if src.exists() and not dst.exists():
            shutil.copy(src, dst)
            print(f"  複製: {config_file}")
    
    # 建立 output 目錄結構
    for subdir in ["driver", "log"]:
        (output_dir / "output" / subdir).mkdir(parents=True, exist_ok=True)
        print(f"  建立目錄: output/{subdir}")
    
    print("✓ 額外檔案處理完成")
    return True


def create_readme():
    """建立使用說明"""
    readme_content = f"""# {PROJECT_NAME} v{VERSION}

## 使用說明

1. 右鍵點擊 `{PROJECT_NAME}.exe`，選擇「以系統管理員身分執行」
2. 程式需要管理員權限才能操作 WIM 映像

## 功能

- WIM 掛載管理：掛載/卸載 WIM 映像
- 驅動管理：查看、匯出、新增、移除驅動程式
- 腳本管理：編輯 WinPE startnet.cmd 腳本

## 注意事項

- 操作 WIM 映像前，請確保映像檔案未被其他程式使用
- 建議在操作前先備份重要檔案
- 輸出的驅動程式會儲存在 output/driver 目錄

## 系統需求

- Windows 10/11
- 管理員權限
- DISM 工具（Windows 內建）

---
建置時間: {__import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
"""
    
    readme_path = DIST_DIR / PROJECT_NAME / "README.txt"
    readme_path.write_text(readme_content, encoding='utf-8')
    print(f"✓ 建立 README.txt")


def main():
    print(f"=" * 50)
    print(f"  {PROJECT_NAME} 打包工具 v{VERSION}")
    print(f"=" * 50)
    print()
    
    # 檢查主程式存在
    if not (BASE_DIR / MAIN_SCRIPT).exists():
        print(f"✗ 錯誤: 找不到 {MAIN_SCRIPT}")
        sys.exit(1)
    
    # 步驟 1: 檢查套件
    check_requirements()
    
    # 步驟 2: 清理
    clean_build()
    
    # 步驟 3: 嘗試 Cython 編譯（可選）
    use_cython, cython_files = try_cython_compile()
    
    # 步驟 4: PyInstaller 打包
    if not build_with_pyinstaller(use_cython, cython_files):
        print("\n✗ 打包失敗")
        sys.exit(1)
    
    # 步驟 5: 複製額外檔案
    copy_additional_files()
    
    # 步驟 6: 建立說明文件
    create_readme()
    
    # 清理臨時檔案
    if TEMP_DIR.exists():
        shutil.rmtree(TEMP_DIR, ignore_errors=True)
    
    # 完成
    print()
    print("=" * 50)
    print("  打包完成！")
    print("=" * 50)
    print()
    
    output_exe = DIST_DIR / PROJECT_NAME / f"{PROJECT_NAME}.exe"
    if output_exe.exists():
        size_mb = output_exe.stat().st_size / 1024 / 1024
        print(f"輸出目錄: {DIST_DIR / PROJECT_NAME}")
        print(f"執行檔:   {output_exe}")
        print(f"檔案大小: {size_mb:.1f} MB")
    
    print()
    print("使用方式: 右鍵以系統管理員身分執行")


if __name__ == "__main__":
    main()
