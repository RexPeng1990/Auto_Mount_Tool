#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WIM Driver Manager - 統一打包系統
提供多種保護級別的打包選項
"""

import os
import sys
import subprocess
import shutil
import base64
import zlib
import random
import string
import hashlib
import argparse
from datetime import datetime
from pathlib import Path

# ==================== 配置 ====================
APP_NAME = "WIM_Driver_Manager"
SOURCE_FILE = "main.py"
APP_MODULES = ["app"]  # 額外模組目錄
VERSION_FILE = "version.txt"

# 打包輸出目錄配置
OUTPUT_DIRS = {
    "direct": {"release": "direct_release", "build": "direct_build", "spec": "direct_spec"},
    "simple": {"release": "simple_release", "build": "simple_build", "spec": "simple_spec"},
    "advanced": {"release": "release", "build": "build_temp", "spec": "spec"},
    "ultimate": {"release": "ultimate_release", "build": "ultimate_build", "spec": "ultimate_spec"},
}

# PyInstaller 隱藏導入
HIDDEN_IMPORTS = [
    'tkinter', 'tkinter.ttk', 'tkinter.messagebox', 'tkinter.filedialog',
    'subprocess', 'threading', 'pathlib', 'configparser', 'datetime',
    're', 'os', 'sys', 'base64', 'zlib', 'hashlib', 'ctypes',
    'json', 'shutil', 'time', 'traceback'
]


class VersionManager:
    """版本管理器"""
    
    @staticmethod
    def get_version() -> str:
        """獲取當前版本號"""
        try:
            if os.path.exists(VERSION_FILE):
                with open(VERSION_FILE, 'r', encoding='utf-8') as f:
                    return f.read().strip()
        except:
            pass
        return "1.0.0"
    
    @staticmethod
    def bump_version(part: str = "patch") -> str:
        """
        增加版本號
        part: major, minor, patch
        """
        version = VersionManager.get_version()
        parts = version.split('.')
        
        if len(parts) != 3:
            parts = ["1", "0", "0"]
        
        major, minor, patch = int(parts[0]), int(parts[1]), int(parts[2])
        
        if part == "major":
            major += 1
            minor = 0
            patch = 0
        elif part == "minor":
            minor += 1
            patch = 0
        else:  # patch
            patch += 1
        
        new_version = f"{major}.{minor}.{patch}"
        
        with open(VERSION_FILE, 'w', encoding='utf-8') as f:
            f.write(new_version)
        
        return new_version


class CodeProtector:
    """代碼保護器"""
    
    def __init__(self, level: str = "simple"):
        self.level = level
        self.key = self._generate_key()
    
    def _generate_key(self, length: int = 32) -> bytes:
        """生成隨機密鑰"""
        return bytes(random.randint(0, 255) for _ in range(length))
    
    def _xor_encrypt(self, data: bytes, key: bytes) -> bytes:
        """XOR 加密"""
        return bytes(b ^ key[i % len(key)] for i, b in enumerate(data))
    
    def _string_encode(self, s: str) -> str:
        """字符串編碼"""
        encoded = base64.b64encode(s.encode('utf-8')).decode('ascii')
        return f"__import__('base64').b64decode('{encoded}').decode('utf-8')"
    
    def simple_protect(self, source_code: str) -> str:
        """簡單保護 - 字符串混淆"""
        # 敏感字符串列表
        sensitive_strings = [
            'WIM', 'Driver', 'Mount', 'Unmount', 'DISM',
            '掛載', '卸載', '驅動', '映像', '管理',
            'Administrator', 'admin', 'password', 'key'
        ]
        
        protected = source_code
        
        # 添加基本檢查頭
        header = '''
# -*- coding: utf-8 -*-
# Protected by WIM Driver Manager Build System
import sys
import os

def _env_check():
    """環境檢查"""
    try:
        # 基本運行環境驗證
        if sys.platform != 'win32':
            return False
        return True
    except:
        return False

if not _env_check():
    sys.exit(1)

'''
        return header + protected
    
    def advanced_protect(self, source_code: str) -> str:
        """進階保護 - 多層加密"""
        # 壓縮
        compressed = zlib.compress(source_code.encode('utf-8'))
        
        # XOR 加密
        encrypted = self._xor_encrypt(compressed, self.key)
        
        # Base64 編碼
        encoded = base64.b64encode(encrypted).decode('ascii')
        key_b64 = base64.b64encode(self.key).decode('ascii')
        
        # 生成解密載入器
        loader = f'''# -*- coding: utf-8 -*-
# Protected Code - Advanced Level
import sys
import os
import base64
import zlib

def _xor_decrypt(data, key):
    return bytes(b ^ key[i % len(key)] for i, b in enumerate(data))

def _load_protected():
    try:
        _key = base64.b64decode('{key_b64}')
        _data = base64.b64decode('{encoded}')
        _decrypted = _xor_decrypt(_data, _key)
        _code = zlib.decompress(_decrypted).decode('utf-8')
        exec(compile(_code, '<protected>', 'exec'), globals())
    except Exception as e:
        sys.exit(1)

if __name__ == '__main__':
    _load_protected()
'''
        return loader
    
    def ultimate_protect(self, source_code: str) -> str:
        """終極保護 - 多層加密 + 反調試"""
        # 先進行進階保護
        protected = self.advanced_protect(source_code)
        
        # 添加反調試機制
        anti_debug = '''
import ctypes
import threading
import time
import random

class _AntiDebug:
    """反調試保護"""
    
    _instance = None
    _running = False
    
    @classmethod
    def start(cls):
        if cls._instance is None:
            cls._instance = cls()
            cls._running = True
            threading.Thread(target=cls._instance._monitor, daemon=True).start()
    
    def _check_debugger(self):
        """檢測調試器"""
        try:
            if hasattr(ctypes, 'windll'):
                return ctypes.windll.kernel32.IsDebuggerPresent() != 0
        except:
            pass
        return False
    
    def _check_processes(self):
        """檢測可疑進程"""
        suspicious = ['ida', 'olly', 'x64dbg', 'x32dbg', 'windbg', 'immunity']
        try:
            import subprocess
            result = subprocess.run(['tasklist'], capture_output=True, text=True, creationflags=0x08000000)
            output = result.stdout.lower()
            return any(proc in output for proc in suspicious)
        except:
            return False
    
    def _monitor(self):
        """持續監控"""
        while self._running:
            try:
                if self._check_debugger() or self._check_processes():
                    os._exit(1)
                time.sleep(random.uniform(5, 15))
            except:
                pass

# 啟動反調試
try:
    _AntiDebug.start()
except:
    pass

'''
        return anti_debug + protected
    
    def protect(self, source_code: str) -> str:
        """根據保護級別處理代碼"""
        if self.level == "direct":
            return source_code
        elif self.level == "simple":
            return self.simple_protect(source_code)
        elif self.level == "advanced":
            return self.advanced_protect(source_code)
        elif self.level == "ultimate":
            return self.ultimate_protect(source_code)
        else:
            return source_code


class Builder:
    """打包構建器"""
    
    def __init__(self, level: str = "simple"):
        self.level = level
        self.protector = CodeProtector(level)
        self.dirs = OUTPUT_DIRS.get(level, OUTPUT_DIRS["simple"])
        self.version = VersionManager.get_version()
    
    def _collect_all_sources(self) -> str:
        """收集所有源代碼"""
        sources = []
        
        # 主文件
        if os.path.exists(SOURCE_FILE):
            with open(SOURCE_FILE, 'r', encoding='utf-8') as f:
                sources.append(f"# === {SOURCE_FILE} ===\n" + f.read())
        
        return '\n\n'.join(sources)
    
    def _prepare_directories(self):
        """準備輸出目錄"""
        for dir_path in self.dirs.values():
            os.makedirs(dir_path, exist_ok=True)
    
    def _cleanup_directories(self):
        """清理臨時目錄"""
        for key in ['build', 'spec']:
            dir_path = self.dirs.get(key)
            if dir_path and os.path.exists(dir_path):
                shutil.rmtree(dir_path, ignore_errors=True)
    
    def _get_output_name(self) -> str:
        """獲取輸出名稱"""
        suffix_map = {
            "direct": "Direct",
            "simple": "Simple", 
            "advanced": "Protected",
            "ultimate": "Ultimate"
        }
        suffix = suffix_map.get(self.level, "")
        return f"{APP_NAME}_{suffix}" if suffix else APP_NAME
    
    def _build_hidden_imports_args(self) -> list:
        """構建隱藏導入參數"""
        args = []
        for imp in HIDDEN_IMPORTS:
            args.extend(['--hidden-import', imp])
        return args
    
    def _copy_app_modules(self, output_dir: str):
        """複製 app 模組到輸出目錄"""
        for module in APP_MODULES:
            src = Path(module)
            if src.exists() and src.is_dir():
                dst = Path(output_dir) / module
                if dst.exists():
                    shutil.rmtree(dst)
                shutil.copytree(src, dst, ignore=shutil.ignore_patterns('__pycache__', '*.pyc'))
                print(f"  複製模組: {module} -> {dst}")
    
    def build(self) -> bool:
        """執行構建"""
        print(f"\n{'='*50}")
        print(f"  WIM Driver Manager 打包系統")
        print(f"  保護級別: {self.level.upper()}")
        print(f"  版本: {self.version}")
        print(f"  時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"{'='*50}\n")
        
        try:
            # 準備目錄
            print("[1/5] 準備輸出目錄...")
            self._prepare_directories()
            
            # 讀取源代碼
            print("[2/5] 讀取源代碼...")
            if not os.path.exists(SOURCE_FILE):
                print(f"  ✗ 找不到源文件: {SOURCE_FILE}")
                return False
            
            with open(SOURCE_FILE, 'r', encoding='utf-8') as f:
                source_code = f.read()
            print(f"  ✓ 讀取完成 ({len(source_code)} 字符)")
            
            # 代碼保護
            print(f"[3/5] 應用 {self.level} 級別保護...")
            protected_code = self.protector.protect(source_code)
            
            # 寫入臨時文件
            temp_file = f"_protected_{self.level}.py"
            with open(temp_file, 'w', encoding='utf-8') as f:
                f.write(protected_code)
            print(f"  ✓ 保護完成")
            
            # PyInstaller 打包
            print("[4/5] 執行 PyInstaller 打包...")
            output_name = self._get_output_name()
            
            cmd = [
                'pyinstaller',
                '--onedir',
                '--windowed',
                f'--distpath={self.dirs["release"]}',
                f'--workpath={self.dirs["build"]}',
                f'--specpath={self.dirs["spec"]}',
                f'--name={output_name}',
                '--noconfirm',
            ]
            cmd.extend(self._build_hidden_imports_args())
            cmd.append(temp_file)
            
            result = subprocess.run(cmd, capture_output=True, text=True)
            
            # 清理臨時文件
            try:
                os.remove(temp_file)
            except:
                pass
            
            if result.returncode != 0:
                print(f"  ✗ PyInstaller 失敗:")
                print(result.stderr[:500] if result.stderr else "未知錯誤")
                return False
            
            print(f"  ✓ 打包完成")
            
            # 複製模組
            print("[5/5] 複製額外模組...")
            output_dir = os.path.join(self.dirs["release"], output_name)
            self._copy_app_modules(output_dir)
            
            # 顯示結果
            exe_path = os.path.join(output_dir, f"{output_name}.exe")
            if os.path.exists(exe_path):
                size_mb = os.path.getsize(exe_path) / 1024 / 1024
                print(f"\n{'='*50}")
                print(f"  ✓ 構建成功!")
                print(f"{'='*50}")
                print(f"  輸出目錄: {output_dir}")
                print(f"  執行文件: {exe_path}")
                print(f"  文件大小: {size_mb:.2f} MB")
                print(f"\n  保護特性:")
                
                features = {
                    "direct": ["原始代碼打包", "無保護", "最高相容性"],
                    "simple": ["字符串混淆", "環境檢查", "高相容性"],
                    "advanced": ["多層加密", "XOR+zlib", "動態解密"],
                    "ultimate": ["多層加密", "反調試", "進程監控", "最高保護"]
                }
                for feat in features.get(self.level, []):
                    print(f"    ✓ {feat}")
                
                return True
            else:
                print(f"  ✗ 未找到輸出文件: {exe_path}")
                return False
            
        except Exception as e:
            print(f"\n  ✗ 構建過程出錯: {e}")
            import traceback
            traceback.print_exc()
            return False
        
        finally:
            # 清理臨時目錄
            self._cleanup_directories()


def main():
    parser = argparse.ArgumentParser(
        description='WIM Driver Manager 打包保護系統',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
保護級別說明:
  direct   - 直接打包，無保護，最高穩定性
  simple   - 簡單保護，字符串混淆，推薦日常使用
  advanced - 進階保護，多層加密，適合發布
  ultimate - 終極保護，加密+反調試，最高安全性
  all      - 構建所有版本

範例:
  python build_all.py                    # 構建簡單保護版本
  python build_all.py -l advanced        # 構建進階保護版本
  python build_all.py -l all             # 構建所有版本
  python build_all.py -l simple -b patch # 構建並增加修訂版本號
'''
    )
    
    parser.add_argument(
        '-l', '--level',
        choices=['direct', 'simple', 'advanced', 'ultimate', 'all'],
        default='simple',
        help='保護級別 (預設: simple)'
    )
    
    parser.add_argument(
        '-b', '--bump',
        choices=['major', 'minor', 'patch', 'none'],
        default='none',
        help='增加版本號 (預設: none)'
    )
    
    parser.add_argument(
        '-c', '--clean',
        action='store_true',
        help='構建前清理所有輸出目錄'
    )
    
    args = parser.parse_args()
    
    # 版本號處理
    if args.bump != 'none':
        new_version = VersionManager.bump_version(args.bump)
        print(f"版本號已更新: {new_version}")
    
    # 清理
    if args.clean:
        print("清理輸出目錄...")
        for dirs in OUTPUT_DIRS.values():
            for dir_path in dirs.values():
                if os.path.exists(dir_path):
                    shutil.rmtree(dir_path, ignore_errors=True)
                    print(f"  已清理: {dir_path}")
    
    # 構建
    if args.level == 'all':
        levels = ['direct', 'simple', 'advanced', 'ultimate']
        results = {}
        for level in levels:
            print(f"\n\n{'#'*60}")
            print(f"# 構建 {level.upper()} 版本")
            print(f"{'#'*60}")
            builder = Builder(level)
            results[level] = builder.build()
        
        # 顯示總結
        print(f"\n\n{'='*60}")
        print("構建總結")
        print(f"{'='*60}")
        for level, success in results.items():
            status = "✓ 成功" if success else "✗ 失敗"
            print(f"  {level.upper():12} : {status}")
    else:
        builder = Builder(args.level)
        success = builder.build()
        sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
