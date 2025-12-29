#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WIM Driver Manager - 代碼保護核心模組
提供多層次的代碼保護機制
"""

import os
import sys
import base64
import zlib
import hashlib
import random
import string
import struct
from typing import Optional, Tuple
from datetime import datetime


class EncryptionEngine:
    """加密引擎 - 提供多種加密算法"""
    
    @staticmethod
    def generate_key(length: int = 32) -> bytes:
        """生成隨機密鑰"""
        return bytes(random.randint(0, 255) for _ in range(length))
    
    @staticmethod
    def xor_cipher(data: bytes, key: bytes) -> bytes:
        """XOR 加密/解密"""
        return bytes(b ^ key[i % len(key)] for i, b in enumerate(data))
    
    @staticmethod
    def rc4_cipher(data: bytes, key: bytes) -> bytes:
        """RC4 流加密"""
        S = list(range(256))
        j = 0
        
        # KSA (Key Scheduling Algorithm)
        for i in range(256):
            j = (j + S[i] + key[i % len(key)]) % 256
            S[i], S[j] = S[j], S[i]
        
        # PRGA (Pseudo-Random Generation Algorithm)
        i = j = 0
        result = []
        for byte in data:
            i = (i + 1) % 256
            j = (j + S[i]) % 256
            S[i], S[j] = S[j], S[i]
            k = S[(S[i] + S[j]) % 256]
            result.append(byte ^ k)
        
        return bytes(result)
    
    @staticmethod
    def scramble(data: bytes, seed: int = None) -> Tuple[bytes, int]:
        """位元組打散"""
        if seed is None:
            seed = random.randint(1, 0xFFFFFFFF)
        
        random.seed(seed)
        indices = list(range(len(data)))
        random.shuffle(indices)
        
        scrambled = bytes(data[i] for i in indices)
        return scrambled, seed
    
    @staticmethod
    def unscramble(data: bytes, seed: int) -> bytes:
        """位元組還原"""
        random.seed(seed)
        indices = list(range(len(data)))
        random.shuffle(indices)
        
        result = [0] * len(data)
        for i, idx in enumerate(indices):
            result[idx] = data[i]
        
        return bytes(result)


class CodeObfuscator:
    """代碼混淆器"""
    
    def __init__(self):
        self.var_counter = 0
        self.func_counter = 0
    
    def _generate_var_name(self) -> str:
        """生成混淆變量名"""
        self.var_counter += 1
        chars = ''.join(random.choices('_OoIl1', k=8))
        return f"_{chars}{self.var_counter:04x}"
    
    def _generate_func_name(self) -> str:
        """生成混淆函數名"""
        self.func_counter += 1
        chars = ''.join(random.choices('_OoIl1', k=6))
        return f"__{chars}{self.func_counter:03x}"
    
    def obfuscate_strings(self, code: str, strings: list) -> str:
        """混淆指定字符串"""
        result = code
        
        for s in strings:
            if s in result:
                encoded = base64.b64encode(s.encode('utf-8')).decode('ascii')
                replacement = f"__import__('base64').b64decode('{encoded}').decode('utf-8')"
                result = result.replace(f"'{s}'", replacement)
                result = result.replace(f'"{s}"', replacement)
        
        return result
    
    def add_dead_code(self, code: str, density: float = 0.1) -> str:
        """添加死代碼 (增加逆向難度)"""
        dead_code_templates = [
            "_ = lambda: None",
            "__ = type('_', (), {})()",
            "___ = [0] * 0",
            "____ = {}.get('_', '')",
            "_____ = '' if 0 else ''",
        ]
        
        lines = code.split('\n')
        result = []
        
        for line in lines:
            result.append(line)
            if random.random() < density and line.strip() and not line.strip().startswith('#'):
                dead = random.choice(dead_code_templates)
                indent = len(line) - len(line.lstrip())
                result.append(' ' * indent + dead)
        
        return '\n'.join(result)


class ProtectionLayer:
    """保護層 - 組合多種保護機制"""
    
    def __init__(self):
        self.encryption = EncryptionEngine()
        self.obfuscator = CodeObfuscator()
    
    def create_loader(self, encrypted_data: str, key_b64: str, seed: int, 
                      compression: bool = True, rc4: bool = False) -> str:
        """創建解密載入器"""
        
        decrypt_func = "rc4_cipher" if rc4 else "xor_cipher"
        decompress_step = "zlib.decompress(decrypted)" if compression else "decrypted"
        
        loader = f'''# -*- coding: utf-8 -*-
# Protected by WIM Driver Manager Protection System
# Generated: {datetime.now().isoformat()}
import sys
import os
import base64
import zlib
import random

def xor_cipher(data, key):
    return bytes(b ^ key[i % len(key)] for i, b in enumerate(data))

def rc4_cipher(data, key):
    S = list(range(256))
    j = 0
    for i in range(256):
        j = (j + S[i] + key[i % len(key)]) % 256
        S[i], S[j] = S[j], S[i]
    i = j = 0
    result = []
    for byte in data:
        i = (i + 1) % 256
        j = (j + S[i]) % 256
        S[i], S[j] = S[j], S[i]
        k = S[(S[i] + S[j]) % 256]
        result.append(byte ^ k)
    return bytes(result)

def unscramble(data, seed):
    random.seed(seed)
    indices = list(range(len(data)))
    random.shuffle(indices)
    result = [0] * len(data)
    for i, idx in enumerate(indices):
        result[idx] = data[i]
    return bytes(result)

def _load():
    try:
        _k = base64.b64decode('{key_b64}')
        _d = base64.b64decode('{encrypted_data}')
        _u = unscramble(_d, {seed})
        decrypted = {decrypt_func}(_u, _k)
        _c = {decompress_step}
        exec(compile(_c.decode('utf-8'), '<protected>', 'exec'), globals())
    except Exception as e:
        sys.exit(1)

if __name__ == '__main__':
    _load()
'''
        return loader
    
    def protect_code(self, source_code: str, level: int = 2) -> str:
        """
        保護代碼
        level 1: 基本混淆
        level 2: 加密 + 混淆
        level 3: 多層加密 + 混淆 + 打散
        """
        
        if level == 1:
            # 基本混淆
            sensitive = ['WIM', 'Driver', 'Mount', 'DISM', '掛載', '卸載', '驅動']
            return self.obfuscator.obfuscate_strings(source_code, sensitive)
        
        # 壓縮
        compressed = zlib.compress(source_code.encode('utf-8'), 9)
        
        # 生成密鑰
        key = self.encryption.generate_key(32)
        
        if level >= 3:
            # 打散
            scrambled, seed = self.encryption.scramble(compressed)
            # RC4 加密
            encrypted = self.encryption.rc4_cipher(scrambled, key)
        else:
            seed = 0
            # XOR 加密
            encrypted = self.encryption.xor_cipher(compressed, key)
        
        # Base64 編碼
        encrypted_b64 = base64.b64encode(encrypted).decode('ascii')
        key_b64 = base64.b64encode(key).decode('ascii')
        
        # 創建載入器
        use_rc4 = level >= 3
        loader = self.create_loader(encrypted_b64, key_b64, seed, 
                                   compression=True, rc4=use_rc4)
        
        return loader


class IntegrityChecker:
    """完整性檢查器"""
    
    @staticmethod
    def calculate_hash(data: bytes, algorithm: str = 'sha256') -> str:
        """計算哈希值"""
        if algorithm == 'md5':
            return hashlib.md5(data).hexdigest()
        elif algorithm == 'sha1':
            return hashlib.sha1(data).hexdigest()
        else:
            return hashlib.sha256(data).hexdigest()
    
    @staticmethod
    def create_checksum_code(file_path: str) -> str:
        """創建文件校驗代碼"""
        if not os.path.exists(file_path):
            return ""
        
        with open(file_path, 'rb') as f:
            content = f.read()
        
        hash_value = IntegrityChecker.calculate_hash(content)
        
        return f'''
def _verify_integrity():
    """驗證文件完整性"""
    try:
        import hashlib
        import sys
        expected = '{hash_value}'
        with open(sys.executable, 'rb') as f:
            actual = hashlib.sha256(f.read()).hexdigest()
        return actual == expected
    except:
        return True  # 開發環境跳過檢查

if not _verify_integrity():
    import sys
    sys.exit(1)
'''


class AntiDebugModule:
    """反調試模組"""
    
    @staticmethod
    def get_anti_debug_code() -> str:
        """獲取反調試代碼"""
        return '''
import ctypes
import threading
import time
import random
import os
import sys

class _SecurityMonitor:
    """安全監控器"""
    
    _instance = None
    _active = False
    
    SUSPICIOUS_PROCESSES = [
        'ida', 'ida64', 'idaq', 'idaq64',
        'ollydbg', 'x64dbg', 'x32dbg',
        'windbg', 'immunity', 'radare2',
        'ghidra', 'binary ninja', 'hopper',
        'processhacker', 'procmon', 'procexp',
        'wireshark', 'fiddler', 'charles'
    ]
    
    @classmethod
    def start(cls):
        if cls._instance is None:
            cls._instance = cls()
            cls._active = True
            t = threading.Thread(target=cls._instance._monitor_loop, daemon=True)
            t.start()
    
    @classmethod
    def stop(cls):
        cls._active = False
    
    def _check_debugger_api(self) -> bool:
        """使用 Windows API 檢測調試器"""
        try:
            if hasattr(ctypes, 'windll'):
                # IsDebuggerPresent
                if ctypes.windll.kernel32.IsDebuggerPresent():
                    return True
                
                # CheckRemoteDebuggerPresent
                is_debugged = ctypes.c_bool(False)
                ctypes.windll.kernel32.CheckRemoteDebuggerPresent(
                    ctypes.windll.kernel32.GetCurrentProcess(),
                    ctypes.byref(is_debugged)
                )
                if is_debugged.value:
                    return True
        except:
            pass
        return False
    
    def _check_timing(self) -> bool:
        """時間檢測 (調試會導致執行變慢)"""
        try:
            start = time.perf_counter()
            for _ in range(1000):
                pass
            elapsed = time.perf_counter() - start
            # 正常執行應該非常快
            return elapsed > 0.1
        except:
            pass
        return False
    
    def _check_processes(self) -> bool:
        """檢測可疑進程"""
        try:
            import subprocess
            result = subprocess.run(
                ['tasklist', '/FO', 'CSV'],
                capture_output=True, text=True,
                creationflags=0x08000000  # CREATE_NO_WINDOW
            )
            output = result.stdout.lower()
            return any(proc in output for proc in self.SUSPICIOUS_PROCESSES)
        except:
            pass
        return False
    
    def _on_threat_detected(self):
        """檢測到威脅時的處理"""
        try:
            # 靜默退出
            os._exit(1)
        except:
            sys.exit(1)
    
    def _monitor_loop(self):
        """監控循環"""
        while self._active:
            try:
                # 隨機檢查間隔
                time.sleep(random.uniform(3, 10))
                
                # 執行各項檢查
                if self._check_debugger_api():
                    self._on_threat_detected()
                
                if self._check_timing():
                    self._on_threat_detected()
                
                # 進程檢查頻率較低
                if random.random() < 0.3:
                    if self._check_processes():
                        self._on_threat_detected()
                        
            except Exception:
                pass

# 啟動安全監控
try:
    if sys.platform == 'win32':
        _SecurityMonitor.start()
except:
    pass
'''


def create_protected_module(source_code: str, 
                           protection_level: int = 2,
                           add_anti_debug: bool = False,
                           add_integrity_check: bool = False) -> str:
    """
    創建受保護的模組
    
    Args:
        source_code: 原始源代碼
        protection_level: 保護級別 (1-3)
        add_anti_debug: 是否添加反調試
        add_integrity_check: 是否添加完整性檢查
    
    Returns:
        受保護的代碼
    """
    protector = ProtectionLayer()
    
    # 應用代碼保護
    protected = protector.protect_code(source_code, protection_level)
    
    # 添加反調試
    if add_anti_debug:
        anti_debug = AntiDebugModule.get_anti_debug_code()
        protected = anti_debug + '\n' + protected
    
    return protected


# 命令行接口
if __name__ == '__main__':
    import argparse
    
    parser = argparse.ArgumentParser(description='代碼保護工具')
    parser.add_argument('input', help='輸入文件')
    parser.add_argument('-o', '--output', help='輸出文件')
    parser.add_argument('-l', '--level', type=int, default=2, choices=[1, 2, 3],
                       help='保護級別 (1-3)')
    parser.add_argument('--anti-debug', action='store_true', help='添加反調試')
    
    args = parser.parse_args()
    
    with open(args.input, 'r', encoding='utf-8') as f:
        source = f.read()
    
    protected = create_protected_module(
        source,
        protection_level=args.level,
        add_anti_debug=args.anti_debug
    )
    
    output_file = args.output or args.input.replace('.py', '_protected.py')
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(protected)
    
    print(f"Protected code saved to: {output_file}")
