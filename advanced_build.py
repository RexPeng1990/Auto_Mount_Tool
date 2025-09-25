#!/usr/bin/env python3
"""
進階代碼混淆和打包工具
結合多種混淆技術和加密打包
"""

import os
import sys
import subprocess
import shutil
import tempfile
import zipfile
import base64
import random
import string
from pathlib import Path

class AdvancedObfuscator:
    def __init__(self):
        self.temp_dir = tempfile.mkdtemp()
        self.key = self.generate_key()
        
    def generate_key(self):
        """生成隨機加密密鑰"""
        return ''.join(random.choices(string.ascii_letters + string.digits, k=32))
    
    def create_loader(self, encrypted_data):
        """創建加密載入器"""
        loader_code = f'''
import base64
import zlib
import marshal
import sys

def _decrypt_and_load():
    encrypted = """{encrypted_data}"""
    key = "{self.key}"
    
    # 簡單的解密邏輯
    decrypted = base64.b64decode(encrypted.encode()).decode()
    
    # 解壓縮
    decompressed = zlib.decompress(base64.b64decode(decrypted.encode()))
    
    # 執行代碼
    code = marshal.loads(decompressed)
    exec(code)

if __name__ == "__main__":
    _decrypt_and_load()
'''
        return loader_code
    
    def obfuscate_and_encrypt(self, source_file, output_file):
        """混淆並加密Python文件"""
        try:
            # 讀取原始代碼
            with open(source_file, 'r', encoding='utf-8') as f:
                source_code = f.read()
            
            # 編譯代碼
            compiled = compile(source_code, source_file, 'exec')
            marshaled = marshal.dumps(compiled)
            
            # 壓縮
            compressed = zlib.compress(marshaled)
            
            # Base64編碼
            encoded = base64.b64encode(compressed).decode()
            
            # 再次Base64編碼作為簡單加密
            encrypted = base64.b64encode(encoded.encode()).decode()
            
            # 創建載入器
            loader = self.create_loader(encrypted)
            
            # 寫入輸出文件
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(loader)
            
            print(f"成功混淆並加密: {source_file} -> {output_file}")
            return True
            
        except Exception as e:
            print(f"混淆失敗: {e}")
            return False
    
    def build_executable(self, script_file, output_name):
        """使用PyInstaller打包為可執行文件"""
        try:
            # PyInstaller 命令
            cmd = [
                'pyinstaller',
                '--onedir',           # 目錄模式
                '--windowed',          # 不顯示控制台
                '--optimize=1',        # 基本優化等級
                '--distpath=release',  # 輸出目錄
                '--workpath=build_temp', # 臨時文件目錄
                '--specpath=spec',     # spec文件目錄
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
                # '--add-data=version.txt;.',  # 如果存在版本文件
                script_file
            ]
            
            # 執行打包
            result = subprocess.run(cmd, capture_output=True, text=True)
            
            if result.returncode == 0:
                print(f"打包成功: release/{output_name}.exe")
                return True
            else:
                print(f"打包失敗: {result.stderr}")
                return False
                
        except Exception as e:
            print(f"打包過程出錯: {e}")
            return False
    
    def cleanup(self):
        """清理臨時文件"""
        try:
            shutil.rmtree(self.temp_dir, ignore_errors=True)
            shutil.rmtree('build_temp', ignore_errors=True)
            shutil.rmtree('spec', ignore_errors=True)
            shutil.rmtree('__pycache__', ignore_errors=True)
        except:
            pass

def main():
    print("=== 進階代碼混淆和打包工具 ===")
    
    source_file = "main.py"
    obfuscated_file = "main_protected.py"
    output_name = "WIM_Driver_Manager_Protected"
    
    if not os.path.exists(source_file):
        print(f"錯誤: 找不到源文件 '{source_file}'")
        sys.exit(1)
    
    obfuscator = AdvancedObfuscator()
    
    try:
        # 步驟1: 混淆和加密
        print("步驟1: 混淆和加密代碼...")
        if not obfuscator.obfuscate_and_encrypt(source_file, obfuscated_file):
            print("混淆失敗，退出")
            sys.exit(1)
        
        # 步驟2: 打包為可執行文件
        print("步驟2: 打包為可執行文件...")
        if not obfuscator.build_executable(obfuscated_file, output_name):
            print("打包失敗")
            sys.exit(1)
        
        print("\\n=== 完成 ===")
        print(f"受保護的可執行文件: release/{output_name}.exe")
        print("注意: 請妥善保管加密密鑰和源代碼")
        
    finally:
        # 清理臨時文件
        obfuscator.cleanup()
        # 清理混淆後的文件
        try:
            os.remove(obfuscated_file)
        except:
            pass

if __name__ == "__main__":
    import marshal
    import zlib
    main()
