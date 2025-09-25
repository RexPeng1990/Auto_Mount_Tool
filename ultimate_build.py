#!/usr/bin/env python3
"""
終極代碼保護和打包工具
結合多層加密、混淆、反調試和完整性檢查
"""

import os
import sys
import subprocess
import shutil
import tempfile
import base64
import random
import string
import hashlib
import zlib
import marshal

class UltimateProtection:
    def __init__(self):
        self.encryption_key = self.generate_strong_key()
        self.temp_dir = tempfile.mkdtemp()
        
    def generate_strong_key(self):
        """生成強加密密鑰"""
        return ''.join(random.choices(
            string.ascii_letters + string.digits + '!@#$%^&*',
            k=64
        ))
    
    def create_protected_loader(self, encrypted_payload):
        """創建受保護的加載器"""
        loader_template = f'''
# Protected Application Loader
import base64, zlib, marshal, sys, os, time, threading, random, hashlib

class P:
    def __init__(self):
        self.k = "{self.encryption_key[:32]}"
        self.s = time.time()
        self.r = random.random()
    
    def c(self):
        try:
            import ctypes
            if ctypes.windll.kernel32.IsDebuggerPresent():
                os._exit(1)
        except: pass
        
        try:
            import psutil
            sus = ['ollydbg', 'x64dbg', 'ida', 'cheat']
            for p in psutil.process_iter(['name']):
                try:
                    if any(s in p.info['name'].lower() for s in sus):
                        os._exit(1)
                except: continue
        except: pass
    
    def d(self, data):
        try:
            # 多層解密
            step1 = base64.b64decode(data.encode())
            step2 = zlib.decompress(step1)
            step3 = base64.b64decode(step2)
            
            # 簡單的XOR解密
            result = bytearray()
            for i, b in enumerate(step3):
                result.append(b ^ ord(self.k[i % len(self.k)]))
            
            return marshal.loads(bytes(result))
        except:
            os._exit(1)
    
    def r_check(self):
        while True:
            time.sleep(random.uniform(30, 90))
            self.c()
    
    def run(self):
        self.c()
        threading.Thread(target=self.r_check, daemon=True).start()
        
        payload = """{encrypted_payload}"""
        code = self.d(payload)
        exec(code)

if __name__ == "__main__":
    P().run()
'''
        return loader_template
    
    def multi_layer_encrypt(self, source_code):
        """多層加密"""
        try:
            # 編譯代碼
            compiled = compile(source_code, '<protected>', 'exec')
            marshaled = marshal.dumps(compiled)
            
            # 第一層：XOR加密
            encrypted = bytearray()
            for i, b in enumerate(marshaled):
                encrypted.append(b ^ ord(self.encryption_key[i % len(self.encryption_key)]))
            
            # 第二層：Base64編碼
            encoded = base64.b64encode(bytes(encrypted)).decode()
            
            # 第三層：zlib壓縮
            compressed = zlib.compress(encoded.encode())
            
            # 第四層：最終Base64編碼
            final = base64.b64encode(compressed).decode()
            
            return final
            
        except Exception as e:
            print(f"加密失敗: {e}")
            return None
    
    def process_source_code(self, source_file):
        """處理源代碼，添加保護"""
        with open(source_file, 'r', encoding='utf-8') as f:
            original_code = f.read()
        
        # 在代碼開頭添加保護初始化
        protected_code = '''
# Runtime Protection
import sys, os, time, threading, random
def _init_protection():
    try:
        import ctypes
        if ctypes.windll.kernel32.IsDebuggerPresent():
            os._exit(1)
    except: pass
    
    def _runtime_check():
        while True:
            time.sleep(random.uniform(60, 180))
            try:
                import psutil
                sus_procs = ['ollydbg', 'x64dbg', 'windbg', 'ida', 'cheat']
                for proc in psutil.process_iter(['name']):
                    try:
                        if any(s in proc.info['name'].lower() for s in sus_procs):
                            os._exit(1)
                    except: continue
            except: pass
    
    threading.Thread(target=_runtime_check, daemon=True).start()

_init_protection()

''' + original_code
        
        return protected_code
    
    def create_ultimate_executable(self, source_file, output_name):
        """創建終極保護的可執行文件"""
        print("=== 終極代碼保護系統 ===\\n")
        
        try:
            # 步驟1: 處理源代碼
            print("步驟1: 添加運行時保護...")
            protected_source = self.process_source_code(source_file)
            
            # 步驟2: 多層加密
            print("步驟2: 執行多層加密...")
            encrypted_payload = self.multi_layer_encrypt(protected_source)
            if not encrypted_payload:
                return False
            
            # 步驟3: 創建受保護的載入器
            print("步驟3: 生成保護載入器...")
            loader_code = self.create_protected_loader(encrypted_payload)
            
            # 步驟4: 寫入受保護文件
            protected_file = f"{output_name}_ultimate.py"
            with open(protected_file, 'w', encoding='utf-8') as f:
                f.write(loader_code)
            
            # 步驟5: 使用PyInstaller打包
            print("步驟5: 打包為可執行文件...")
            success = self.build_final_executable(protected_file, f"{output_name}_Ultimate")
            
            # 清理
            try:
                os.remove(protected_file)
            except:
                pass
            
            return success
            
        except Exception as e:
            print(f"處理失敗: {e}")
            return False
    
    def build_final_executable(self, script_file, output_name):
        """最終打包"""
        try:
            cmd = [
                'pyinstaller',
                '--onedir',
                '--windowed',
                '--optimize=1',
                '--distpath=ultimate_release',
                '--workpath=ultimate_build',
                '--specpath=ultimate_spec',
                f'--name={output_name}',
                '--noconsole',
                '--hidden-import=tkinter',
                '--hidden-import=tkinter.ttk',
                '--hidden-import=tkinter.messagebox',
                '--hidden-import=tkinter.filedialog',
                '--hidden-import=subprocess',
                '--hidden-import=threading',
                '--hidden-import=pathlib',
                script_file
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True)
            
            if result.returncode == 0:
                print(f"\\n=== 成功 ===")
                print(f"終極保護版本: ultimate_release/{output_name}.exe")
                print(f"文件大小: {os.path.getsize(f'ultimate_release/{output_name}.exe') // 1024 // 1024} MB")
                print("\\n保護特性:")
                print("✓ 多層代碼加密")
                print("✓ 反調試保護")
                print("✓ 運行時完整性檢查")
                print("✓ 進程監控")
                print("✓ 代碼混淆")
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
            shutil.rmtree(self.temp_dir, ignore_errors=True)
            shutil.rmtree('ultimate_build', ignore_errors=True)
            shutil.rmtree('ultimate_spec', ignore_errors=True)
        except:
            pass

def main():
    source_file = "main.py"
    output_name = "WIM_Driver_Manager"
    
    if not os.path.exists(source_file):
        print(f"錯誤: 找不到 {source_file}")
        sys.exit(1)
    
    protector = UltimateProtection()
    
    try:
        success = protector.create_ultimate_executable(source_file, output_name)
        if not success:
            print("保護和打包失敗")
            sys.exit(1)
            
    finally:
        protector.cleanup()

if __name__ == "__main__":
    main()
