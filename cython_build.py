#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WIM Driver Manager - Cython 編譯打包系統
將 Python 代碼編譯成 C 擴展，提供更強的保護

使用前請確保已安裝:
  pip install cython pyinstaller

使用方式:
  python cython_build.py
  python cython_build.py --keep-c    # 保留生成的 C 文件
"""

import os
import sys
import shutil
import subprocess
import tempfile
from pathlib import Path
from datetime import datetime

# ==================== 配置 ====================
APP_NAME = "WIM_Driver_Manager"
SOURCE_FILE = "main.py"
APP_MODULES = ["app"]  # 需要編譯的模組目錄

# PyInstaller 隱藏導入
HIDDEN_IMPORTS = [
    'tkinter', 'tkinter.ttk', 'tkinter.messagebox', 'tkinter.filedialog',
    'subprocess', 'threading', 'pathlib', 'configparser', 'datetime',
    're', 'os', 'sys', 'base64', 'zlib', 'hashlib', 'ctypes',
    'json', 'shutil', 'time', 'traceback'
]


def check_dependencies():
    """檢查依賴"""
    print("[檢查] 檢查必要依賴...")
    
    # 檢查 Cython
    try:
        import Cython
        print(f"  ✓ Cython {Cython.__version__}")
    except ImportError:
        print("  ✗ 未安裝 Cython")
        print("    請執行: pip install cython")
        return False
    
    # 檢查 PyInstaller
    try:
        import PyInstaller
        print(f"  ✓ PyInstaller {PyInstaller.__version__}")
    except ImportError:
        print("  ✗ 未安裝 PyInstaller")
        print("    請執行: pip install pyinstaller")
        return False
    
    # 檢查 C 編譯器
    try:
        result = subprocess.run(['cl'], capture_output=True, text=True)
        print("  ✓ MSVC 編譯器可用")
    except FileNotFoundError:
        try:
            result = subprocess.run(['gcc', '--version'], capture_output=True, text=True)
            print("  ✓ GCC 編譯器可用")
        except FileNotFoundError:
            print("  ⚠ 未檢測到 C 編譯器 (將嘗試使用 distutils 自動尋找)")
    
    return True


def create_setup_py(build_dir: str, py_files: list) -> str:
    """創建 setup.py 用於 Cython 編譯"""
    
    # 構建模組列表
    ext_modules = []
    for py_file in py_files:
        # 將路徑轉換為模組名
        rel_path = os.path.relpath(py_file, build_dir)
        module_name = rel_path.replace(os.sep, '.').replace('.py', '')
        ext_modules.append(f"    Extension('{module_name}', [r'{py_file}']),")
    
    ext_modules_str = '\n'.join(ext_modules)
    
    setup_content = f'''# -*- coding: utf-8 -*-
from setuptools import setup
from Cython.Build import cythonize
from setuptools.extension import Extension
import sys

# 強制使用 C 語言編譯
ext_modules = [
{ext_modules_str}
]

setup(
    name='{APP_NAME}',
    ext_modules=cythonize(
        ext_modules,
        compiler_directives={{
            'language_level': '3',
            'boundscheck': False,
            'wraparound': False,
            'initializedcheck': False,
            'cdivision': True,
        }},
        annotate=False,  # 不生成 HTML 注釋文件
    ),
    zip_safe=False,
)
'''
    
    setup_path = os.path.join(build_dir, 'setup.py')
    with open(setup_path, 'w', encoding='utf-8') as f:
        f.write(setup_content)
    
    return setup_path


def collect_python_files(source_dir: str, build_dir: str) -> list:
    """收集並複製所有 Python 文件"""
    py_files = []
    
    # 複製主文件
    main_src = os.path.join(source_dir, SOURCE_FILE)
    main_dst = os.path.join(build_dir, SOURCE_FILE)
    if os.path.exists(main_src):
        shutil.copy2(main_src, main_dst)
        py_files.append(main_dst)
        print(f"  複製: {SOURCE_FILE}")
    
    # 複製 app 模組
    for module in APP_MODULES:
        module_src = os.path.join(source_dir, module)
        module_dst = os.path.join(build_dir, module)
        
        if os.path.exists(module_src) and os.path.isdir(module_src):
            # 創建目標目錄
            os.makedirs(module_dst, exist_ok=True)
            
            # 複製所有 .py 文件
            for root, dirs, files in os.walk(module_src):
                # 跳過 __pycache__
                dirs[:] = [d for d in dirs if d != '__pycache__']
                
                rel_root = os.path.relpath(root, module_src)
                dst_root = os.path.join(module_dst, rel_root) if rel_root != '.' else module_dst
                os.makedirs(dst_root, exist_ok=True)
                
                for file in files:
                    if file.endswith('.py'):
                        src_file = os.path.join(root, file)
                        dst_file = os.path.join(dst_root, file)
                        shutil.copy2(src_file, dst_file)
                        
                        # __init__.py 不編譯，只複製
                        if file != '__init__.py':
                            py_files.append(dst_file)
                        print(f"  複製: {module}/{os.path.relpath(dst_file, module_dst)}")
    
    return py_files


def compile_with_cython(build_dir: str, py_files: list) -> bool:
    """使用 Cython 編譯 Python 文件"""
    print("\n[編譯] 使用 Cython 編譯...")
    
    # 創建 setup.py
    setup_path = create_setup_py(build_dir, py_files)
    
    # 執行編譯
    original_dir = os.getcwd()
    try:
        os.chdir(build_dir)
        
        cmd = [sys.executable, 'setup.py', 'build_ext', '--inplace']
        print(f"  執行: {' '.join(cmd)}")
        
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace'
        )
        
        if result.returncode != 0:
            print(f"  ✗ Cython 編譯失敗")
            print(result.stderr[-1000:] if len(result.stderr) > 1000 else result.stderr)
            return False
        
        print("  ✓ Cython 編譯成功")
        return True
        
    finally:
        os.chdir(original_dir)


def create_launcher(build_dir: str) -> str:
    """創建啟動器腳本"""
    
    launcher_content = '''# -*- coding: utf-8 -*-
# WIM Driver Manager - Cython Compiled Launcher
import sys
import os

# 確保模組路徑正確
if hasattr(sys, '_MEIPASS'):
    # PyInstaller 打包環境
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

sys.path.insert(0, base_path)

# 導入編譯後的主模組
try:
    import main
except ImportError as e:
    print(f"Failed to import compiled module: {e}")
    sys.exit(1)
'''
    
    launcher_path = os.path.join(build_dir, '_launcher.py')
    with open(launcher_path, 'w', encoding='utf-8') as f:
        f.write(launcher_content)
    
    return launcher_path


def package_with_pyinstaller(build_dir: str, output_dir: str) -> bool:
    """使用 PyInstaller 打包"""
    print("\n[打包] 使用 PyInstaller 打包...")
    
    # 找到編譯後的 .pyd 文件
    pyd_files = []
    for root, dirs, files in os.walk(build_dir):
        for file in files:
            if file.endswith('.pyd') or file.endswith('.so'):
                pyd_files.append(os.path.join(root, file))
    
    if not pyd_files:
        print("  ✗ 找不到編譯後的 .pyd 文件")
        return False
    
    print(f"  找到 {len(pyd_files)} 個編譯模組")
    for pyd in pyd_files:
        print(f"    - {os.path.basename(pyd)}")
    
    # 創建啟動器
    launcher_path = create_launcher(build_dir)
    
    # 構建 PyInstaller 命令
    output_name = f"{APP_NAME}_Cython"
    
    cmd = [
        'pyinstaller',
        '--onedir',
        '--windowed',
        f'--distpath={output_dir}',
        f'--workpath={os.path.join(output_dir, "build")}',
        f'--specpath={os.path.join(output_dir, "spec")}',
        f'--name={output_name}',
        '--noconfirm',
    ]
    
    # 添加隱藏導入
    for imp in HIDDEN_IMPORTS:
        cmd.extend(['--hidden-import', imp])
    
    # 添加編譯後的 .pyd 文件
    for pyd in pyd_files:
        rel_path = os.path.relpath(pyd, build_dir)
        # 確定目標路徑
        if os.sep in rel_path:
            dest = os.path.dirname(rel_path)
        else:
            dest = '.'
        cmd.extend(['--add-binary', f'{pyd};{dest}'])
    
    # 添加 app 目錄的 __init__.py
    app_init = os.path.join(build_dir, 'app', '__init__.py')
    if os.path.exists(app_init):
        cmd.extend(['--add-data', f'{app_init};app'])
    
    cmd.append(launcher_path)
    
    print(f"  執行 PyInstaller...")
    
    result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
    
    if result.returncode != 0:
        print(f"  ✗ PyInstaller 打包失敗")
        print(result.stderr[-1000:] if len(result.stderr) > 1000 else result.stderr)
        return False
    
    print("  ✓ PyInstaller 打包成功")
    return True


def build(keep_c: bool = False):
    """主要構建流程"""
    print(f"\n{'='*60}")
    print(f"  WIM Driver Manager - Cython 編譯打包")
    print(f"  時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*60}\n")
    
    # 檢查依賴
    if not check_dependencies():
        return False
    
    # 獲取源目錄
    source_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 創建臨時構建目錄
    build_dir = tempfile.mkdtemp(prefix="WIM_Cython_Build_")
    output_dir = tempfile.mkdtemp(prefix="WIM_Cython_Output_")
    
    print(f"\n[準備] 構建目錄: {build_dir}")
    print(f"[準備] 輸出目錄: {output_dir}")
    
    try:
        # 收集 Python 文件
        print("\n[收集] 收集 Python 源文件...")
        py_files = collect_python_files(source_dir, build_dir)
        print(f"  共 {len(py_files)} 個文件需要編譯")
        
        # Cython 編譯
        if not compile_with_cython(build_dir, py_files):
            return False
        
        # PyInstaller 打包
        if not package_with_pyinstaller(build_dir, output_dir):
            return False
        
        # 清理
        print("\n[清理] 清理臨時文件...")
        if not keep_c:
            # 刪除 .c 文件
            for root, dirs, files in os.walk(build_dir):
                for file in files:
                    if file.endswith('.c'):
                        os.remove(os.path.join(root, file))
        
        shutil.rmtree(os.path.join(output_dir, "build"), ignore_errors=True)
        shutil.rmtree(os.path.join(output_dir, "spec"), ignore_errors=True)
        
        # 顯示結果
        output_name = f"{APP_NAME}_Cython"
        exe_path = os.path.join(output_dir, output_name, f"{output_name}.exe")
        
        if os.path.exists(exe_path):
            size_mb = os.path.getsize(exe_path) / (1024 * 1024)
            print(f"\n{'='*60}")
            print(f"  ✓ Cython 編譯打包成功!")
            print(f"  輸出位置: {os.path.join(output_dir, output_name)}")
            print(f"  執行檔: {output_name}.exe")
            print(f"  大小: {size_mb:.2f} MB")
            print(f"{'='*60}\n")
            
            # 保留構建目錄供檢查
            if keep_c:
                print(f"  [保留] C 源碼目錄: {build_dir}")
            else:
                shutil.rmtree(build_dir, ignore_errors=True)
            
            return True
        else:
            print("  ✗ 找不到輸出執行檔")
            return False
            
    except Exception as e:
        print(f"\n  ✗ 構建過程發生錯誤: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """主函數"""
    import argparse
    
    parser = argparse.ArgumentParser(description='WIM Driver Manager Cython 編譯打包工具')
    parser.add_argument('--keep-c', action='store_true', 
                       help='保留生成的 C 源碼文件')
    
    args = parser.parse_args()
    
    print("\n" + "="*60)
    print("  Cython 編譯說明")
    print("="*60)
    print("  Cython 將 Python 代碼轉換為 C 語言，然後編譯成")
    print("  原生機器碼 (.pyd 文件)，提供：")
    print("    • 更強的代碼保護 - 無法直接看到 Python 源碼")
    print("    • 更好的性能 - 原生編譯的執行速度更快")
    print("    • 更難逆向工程 - 需要反編譯 C 擴展")
    print("="*60 + "\n")
    
    success = build(keep_c=args.keep_c)
    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
