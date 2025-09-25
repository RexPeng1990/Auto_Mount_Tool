#!/usr/bin/env python3
"""
最簡直接打包工具
直接打包原始代碼，確保最高穩定性
"""

import os
import sys
import subprocess
import shutil

def build_direct_version():
    """直接打包原始代碼"""
    print("=== 直接打包工具 ===\\n")
    
    source_file = "main.py"
    output_name = "WIM_Driver_Manager_Direct"
    
    if not os.path.exists(source_file):
        print(f"錯誤: 找不到 {source_file}")
        return False
    
    print("步驟1: 直接打包原始代碼...")
    
    try:
        cmd = [
            'pyinstaller',
            '--onedir',
            '--windowed',
            '--distpath=direct_release',
            '--workpath=direct_build',
            '--specpath=direct_spec',
            f'--name={output_name}',
            '--hidden-import=tkinter',
            '--hidden-import=tkinter.ttk',
            '--hidden-import=tkinter.messagebox',
            '--hidden-import=tkinter.filedialog',
            '--hidden-import=subprocess',
            '--hidden-import=threading',
            '--hidden-import=pathlib',
            '--hidden-import=configparser',
            '--hidden-import=datetime',
            '--hidden-import=re',
            '--hidden-import=os',
            '--hidden-import=sys',
            source_file
        ]
        
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode == 0:
            print(f"\\n=== 成功 ===")
            print(f"直接打包版本: direct_release/{output_name}/")
            print(f"執行文件: direct_release/{output_name}/{output_name}.exe")
            
            # 檢查生成的文件
            exe_path = f"direct_release/{output_name}/{output_name}.exe"
            if os.path.exists(exe_path):
                size_mb = os.path.getsize(exe_path) // 1024 // 1024
                print(f"執行文件大小: {size_mb} MB")
            
            print("\\n特性:")
            print("✓ 原始代碼直接打包")
            print("✓ 無任何混淆或加密")
            print("✓ 最高穩定性和相容性")
            print("✓ 目錄式打包 (onedir)")
            return True
        else:
            print(f"打包失敗: {result.stderr}")
            return False
            
    except Exception as e:
        print(f"打包出錯: {e}")
        return False

def cleanup():
    """清理臨時文件"""
    try:
        shutil.rmtree('direct_build', ignore_errors=True)
        shutil.rmtree('direct_spec', ignore_errors=True)
    except:
        pass

def main():
    try:
        success = build_direct_version()
        if not success:
            print("直接打包失敗")
            sys.exit(1)
    finally:
        cleanup()

if __name__ == "__main__":
    main()
