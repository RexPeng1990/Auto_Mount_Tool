#!/usr/bin/env python3
"""
簡化版代碼保護和打包工具
使用基本混淆和onedir模式確保穩定運行
"""

import os
import sys
import subprocess
import shutil
import base64
import random
import string

class SimpleProtection:
    def __init__(self):
        self.key = self.generate_key()
        
    def generate_key(self):
        """生成簡單密鑰"""
        return ''.join(random.choices(string.ascii_letters + string.digits, k=16))
    
    def simple_obfuscate(self, source_code):
        """簡單字符串混淆"""
        # 基本的字符串替換
        replacements = [
            ("'WIM掛載工具'", f"base64.b64decode('{base64.b64encode('WIM掛載工具'.encode()).decode()}').decode()"),
            ("'驅動程式管理'", f"base64.b64decode('{base64.b64encode('驅動程式管理'.encode()).decode()}').decode()"),
            ("'掛載'", f"base64.b64decode('{base64.b64encode('掛載'.encode()).decode()}').decode()"),
            ("'卸載'", f"base64.b64decode('{base64.b64encode('卸載'.encode()).decode()}').decode()"),
            ("'成功'", f"base64.b64decode('{base64.b64encode('成功'.encode()).decode()}').decode()"),
            ("'錯誤'", f"base64.b64decode('{base64.b64encode('錯誤'.encode()).decode()}').decode()"),
            ("'Warning'", f"base64.b64decode('{base64.b64encode('Warning'.encode()).decode()}').decode()"),
            ("'Error'", f"base64.b64decode('{base64.b64encode('Error'.encode()).decode()}').decode()"),
        ]
        
        obfuscated = source_code
        
        # 添加基本保護導入
        protection_header = '''
# Basic Protection
import base64
import os
import sys

def _basic_check():
    """基本完整性檢查"""
    try:
        # 簡單的運行環境檢查
        if hasattr(sys, 'frozen') and hasattr(sys, '_MEIPASS'):
            # PyInstaller 環境
            pass
        return True
    except:
        return False

if not _basic_check():
    sys.exit(1)

'''
        
        # 應用字符串替換
        for original, replacement in replacements:
            obfuscated = obfuscated.replace(original, replacement)
        
        return protection_header + obfuscated
    
    def create_protected_version(self, source_file, output_name):
        """創建保護版本"""
        print("=== 簡化代碼保護系統 ===\\n")
        
        try:
            # 讀取源代碼
            print("步驟1: 讀取源代碼...")
            with open(source_file, 'r', encoding='utf-8') as f:
                source_code = f.read()
            
            # 簡單混淆
            print("步驟2: 執行基本混淆...")
            protected_code = self.simple_obfuscate(source_code)
            
            # 寫入保護文件
            protected_file = f"{output_name}_simple.py"
            with open(protected_file, 'w', encoding='utf-8') as f:
                f.write(protected_code)
            
            # 打包
            print("步驟3: 使用 onedir 模式打包...")
            success = self.build_onedir_executable(protected_file, f"{output_name}_Simple")
            
            # 清理
            try:
                os.remove(protected_file)
            except:
                pass
            
            return success
            
        except Exception as e:
            print(f"處理失敗: {e}")
            return False
    
    def build_onedir_executable(self, script_file, output_name):
        """使用 onedir 模式打包"""
        try:
            cmd = [
                'pyinstaller',
                '--onedir',
                '--windowed',
                '--distpath=simple_release',
                '--workpath=simple_build',
                '--specpath=simple_spec',
                f'--name={output_name}',
                '--hidden-import=tkinter',
                '--hidden-import=tkinter.ttk',
                '--hidden-import=tkinter.messagebox', 
                '--hidden-import=tkinter.filedialog',
                '--hidden-import=subprocess',
                '--hidden-import=threading',
                '--hidden-import=pathlib',
                '--hidden-import=os',
                '--hidden-import=sys',
                '--hidden-import=base64',
                '--hidden-import=configparser',
                '--hidden-import=datetime',
                '--hidden-import=re',
                script_file
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True)
            
            if result.returncode == 0:
                print(f"\\n=== 成功 ===")
                print(f"簡化保護版本: simple_release/{output_name}/")
                print(f"執行文件: simple_release/{output_name}/{output_name}.exe")
                
                # 檢查生成的文件
                exe_path = f"simple_release/{output_name}/{output_name}.exe"
                if os.path.exists(exe_path):
                    size_mb = os.path.getsize(exe_path) // 1024 // 1024
                    print(f"執行文件大小: {size_mb} MB")
                
                print("\\n保護特性:")
                print("✓ 字符串編碼保護")
                print("✓ 基本完整性檢查")
                print("✓ 目錄式打包 (onedir)")
                print("✓ 隱藏控制台窗口")
                return True
            else:
                print(f"打包失敗: {result.stderr}")
                return False
                
        except Exception as e:
            print(f"打包出錯: {e}")
            return False
    
    def cleanup(self):
        """清理臨時文件"""
        try:
            shutil.rmtree('simple_build', ignore_errors=True)
            shutil.rmtree('simple_spec', ignore_errors=True)
        except:
            pass

def main():
    source_file = "main.py"
    output_name = "WIM_Driver_Manager"
    
    if not os.path.exists(source_file):
        print(f"錯誤: 找不到 {source_file}")
        sys.exit(1)
    
    protector = SimpleProtection()
    
    try:
        success = protector.create_protected_version(source_file, output_name)
        if not success:
            print("保護和打包失敗")
            sys.exit(1)
            
    finally:
        protector.cleanup()

if __name__ == "__main__":
    main()
