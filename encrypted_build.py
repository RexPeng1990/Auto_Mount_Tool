#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WIM Driver Manager - 加密打包系統
提供多層次加密保護的打包方案
"""

import os
import sys
import subprocess
import shutil
import base64
import zlib
import random
import hashlib
import struct
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Tuple, Optional

# ==================== 配置 ====================
APP_NAME = "WIM_Driver_Manager"
SOURCE_FILE = "main.py"
VERSION = "1.0.0"

# PyInstaller 隱藏導入
HIDDEN_IMPORTS = [
    'tkinter', 'tkinter.ttk', 'tkinter.messagebox', 'tkinter.filedialog',
    'subprocess', 'threading', 'pathlib', 'configparser', 'datetime',
    're', 'os', 'sys', 'base64', 'zlib', 'hashlib', 'ctypes',
    'json', 'shutil', 'time', 'traceback'
]


class AdvancedEncryption:
    """進階加密引擎"""
    
    @staticmethod
    def generate_key(length: int = 32) -> bytes:
        """生成密碼學安全的隨機密鑰"""
        return os.urandom(length)
    
    @staticmethod
    def xor_encrypt(data: bytes, key: bytes) -> bytes:
        """XOR 加密"""
        return bytes(b ^ key[i % len(key)] for i, b in enumerate(data))
    
    @staticmethod
    def rc4_encrypt(data: bytes, key: bytes) -> bytes:
        """RC4 流加密"""
        S = list(range(256))
        j = 0
        
        # KSA
        for i in range(256):
            j = (j + S[i] + key[i % len(key)]) % 256
            S[i], S[j] = S[j], S[i]
        
        # PRGA
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
    def aes_like_encrypt(data: bytes, key: bytes) -> bytes:
        """
        類 AES 加密 (多輪替換-置換網絡)
        注意：這是簡化版本，用於代碼保護而非安全通信
        """
        # S-Box (替換盒)
        sbox = bytes([
            0x63, 0x7c, 0x77, 0x7b, 0xf2, 0x6b, 0x6f, 0xc5, 0x30, 0x01, 0x67, 0x2b, 0xfe, 0xd7, 0xab, 0x76,
            0xca, 0x82, 0xc9, 0x7d, 0xfa, 0x59, 0x47, 0xf0, 0xad, 0xd4, 0xa2, 0xaf, 0x9c, 0xa4, 0x72, 0xc0,
            0xb7, 0xfd, 0x93, 0x26, 0x36, 0x3f, 0xf7, 0xcc, 0x34, 0xa5, 0xe5, 0xf1, 0x71, 0xd8, 0x31, 0x15,
            0x04, 0xc7, 0x23, 0xc3, 0x18, 0x96, 0x05, 0x9a, 0x07, 0x12, 0x80, 0xe2, 0xeb, 0x27, 0xb2, 0x75,
            0x09, 0x83, 0x2c, 0x1a, 0x1b, 0x6e, 0x5a, 0xa0, 0x52, 0x3b, 0xd6, 0xb3, 0x29, 0xe3, 0x2f, 0x84,
            0x53, 0xd1, 0x00, 0xed, 0x20, 0xfc, 0xb1, 0x5b, 0x6a, 0xcb, 0xbe, 0x39, 0x4a, 0x4c, 0x58, 0xcf,
            0xd0, 0xef, 0xaa, 0xfb, 0x43, 0x4d, 0x33, 0x85, 0x45, 0xf9, 0x02, 0x7f, 0x50, 0x3c, 0x9f, 0xa8,
            0x51, 0xa3, 0x40, 0x8f, 0x92, 0x9d, 0x38, 0xf5, 0xbc, 0xb6, 0xda, 0x21, 0x10, 0xff, 0xf3, 0xd2,
            0xcd, 0x0c, 0x13, 0xec, 0x5f, 0x97, 0x44, 0x17, 0xc4, 0xa7, 0x7e, 0x3d, 0x64, 0x5d, 0x19, 0x73,
            0x60, 0x81, 0x4f, 0xdc, 0x22, 0x2a, 0x90, 0x88, 0x46, 0xee, 0xb8, 0x14, 0xde, 0x5e, 0x0b, 0xdb,
            0xe0, 0x32, 0x3a, 0x0a, 0x49, 0x06, 0x24, 0x5c, 0xc2, 0xd3, 0xac, 0x62, 0x91, 0x95, 0xe4, 0x79,
            0xe7, 0xc8, 0x37, 0x6d, 0x8d, 0xd5, 0x4e, 0xa9, 0x6c, 0x56, 0xf4, 0xea, 0x65, 0x7a, 0xae, 0x08,
            0xba, 0x78, 0x25, 0x2e, 0x1c, 0xa6, 0xb4, 0xc6, 0xe8, 0xdd, 0x74, 0x1f, 0x4b, 0xbd, 0x8b, 0x8a,
            0x70, 0x3e, 0xb5, 0x66, 0x48, 0x03, 0xf6, 0x0e, 0x61, 0x35, 0x57, 0xb9, 0x86, 0xc1, 0x1d, 0x9e,
            0xe1, 0xf8, 0x98, 0x11, 0x69, 0xd9, 0x8e, 0x94, 0x9b, 0x1e, 0x87, 0xe9, 0xce, 0x55, 0x28, 0xdf,
            0x8c, 0xa1, 0x89, 0x0d, 0xbf, 0xe6, 0x42, 0x68, 0x41, 0x99, 0x2d, 0x0f, 0xb0, 0x54, 0xbb, 0x16
        ])
        
        # 確保密鑰長度
        key = (key * ((32 // len(key)) + 1))[:32]
        
        # 多輪加密
        result = bytearray(data)
        rounds = 10
        
        for r in range(rounds):
            # 輪密鑰
            round_key = bytes((key[i] + r * 17) % 256 for i in range(len(key)))
            
            # S-Box 替換
            result = bytearray(sbox[b] for b in result)
            
            # XOR 輪密鑰
            result = bytearray(b ^ round_key[i % len(round_key)] for i, b in enumerate(result))
            
            # 位元旋轉
            result = bytearray(((b << 3) | (b >> 5)) & 0xFF for b in result)
        
        return bytes(result)
    
    @staticmethod
    def scramble_bytes(data: bytes, seed: int) -> bytes:
        """位元組打散"""
        random.seed(seed)
        indices = list(range(len(data)))
        random.shuffle(indices)
        return bytes(data[i] for i in indices)
    
    @staticmethod
    def multi_layer_encrypt(data: bytes, key: bytes) -> Tuple[bytes, dict]:
        """
        多層加密
        Layer 1: 壓縮
        Layer 2: AES-like 加密
        Layer 3: RC4 加密
        Layer 4: XOR 加密
        Layer 5: 位元組打散
        Layer 6: Base64 編碼
        """
        metadata = {}
        
        # Layer 1: 壓縮
        compressed = zlib.compress(data, 9)
        metadata['original_size'] = len(data)
        metadata['compressed_size'] = len(compressed)
        
        # 生成子密鑰
        key1 = hashlib.sha256(key + b'layer1').digest()
        key2 = hashlib.sha256(key + b'layer2').digest()
        key3 = hashlib.sha256(key + b'layer3').digest()
        
        # Layer 2: AES-like 加密
        layer2 = AdvancedEncryption.aes_like_encrypt(compressed, key1)
        
        # Layer 3: RC4 加密
        layer3 = AdvancedEncryption.rc4_encrypt(layer2, key2)
        
        # Layer 4: XOR 加密
        layer4 = AdvancedEncryption.xor_encrypt(layer3, key3)
        
        # Layer 5: 位元組打散
        seed = int.from_bytes(key[:4], 'little')
        layer5 = AdvancedEncryption.scramble_bytes(layer4, seed)
        metadata['scramble_seed'] = seed
        
        return layer5, metadata


class CodeProtectionBuilder:
    """代碼保護打包器"""
    
    def __init__(self, encryption_level: int = 3):
        """
        encryption_level:
        1 = 基本 (XOR + 壓縮)
        2 = 標準 (RC4 + 壓縮 + 混淆)
        3 = 進階 (多層加密 + 反調試)
        """
        self.level = encryption_level
        self.key = AdvancedEncryption.generate_key(32)
    
    def _generate_decryption_code(self, encrypted_b64: str, metadata: dict) -> str:
        """生成解密載入代碼"""
        key_b64 = base64.b64encode(self.key).decode('ascii')
        seed = metadata.get('scramble_seed', 0)
        
        # 分割 key 和 data 以增加混淆
        key_parts = [key_b64[i:i+16] for i in range(0, len(key_b64), 16)]
        data_parts = [encrypted_b64[i:i+1000] for i in range(0, len(encrypted_b64), 1000)]
        
        key_parts_str = str(key_parts)
        data_parts_str = str(data_parts)
        
        if self.level == 1:
            return self._basic_loader(encrypted_b64, key_b64)
        elif self.level == 2:
            return self._standard_loader(encrypted_b64, key_b64, seed)
        else:
            return self._advanced_loader(data_parts_str, key_parts_str, seed)
    
    def _basic_loader(self, data_b64: str, key_b64: str) -> str:
        """基本載入器 (Level 1)"""
        return f'''# -*- coding: utf-8 -*-
# WIM Driver Manager - Protected Build
# Level: Basic Encryption
import sys, os, base64, zlib

def _d(d, k):
    return bytes(b ^ k[i % len(k)] for i, b in enumerate(d))

def _r():
    try:
        k = base64.b64decode('{key_b64}')
        d = base64.b64decode('{data_b64}')
        c = zlib.decompress(_d(d, k)).decode('utf-8')
        exec(compile(c, '<main>', 'exec'), globals())
    except: sys.exit(1)

if __name__ == '__main__': _r()
'''
    
    def _standard_loader(self, data_b64: str, key_b64: str, seed: int) -> str:
        """標準載入器 (Level 2)"""
        return f'''# -*- coding: utf-8 -*-
# WIM Driver Manager - Protected Build
# Level: Standard Encryption (RC4 + Scramble)
import sys, os, base64, zlib, random

def _rc4(d, k):
    S = list(range(256))
    j = 0
    for i in range(256):
        j = (j + S[i] + k[i % len(k)]) % 256
        S[i], S[j] = S[j], S[i]
    i = j = 0
    r = []
    for b in d:
        i = (i + 1) % 256
        j = (j + S[i]) % 256
        S[i], S[j] = S[j], S[i]
        r.append(b ^ S[(S[i] + S[j]) % 256])
    return bytes(r)

def _us(d, s):
    random.seed(s)
    idx = list(range(len(d)))
    random.shuffle(idx)
    r = [0] * len(d)
    for i, x in enumerate(idx): r[x] = d[i]
    return bytes(r)

def _r():
    try:
        k = base64.b64decode('{key_b64}')
        d = base64.b64decode('{data_b64}')
        d = _us(d, {seed})
        c = zlib.decompress(_rc4(d, k)).decode('utf-8')
        exec(compile(c, '<main>', 'exec'), globals())
    except: sys.exit(1)

if __name__ == '__main__': _r()
'''
    
    def _advanced_loader(self, data_parts: str, key_parts: str, seed: int) -> str:
        """進階載入器 (Level 3) - 多層解密 + 反調試"""
        return f'''# -*- coding: utf-8 -*-
# WIM Driver Manager - Protected Build
# Level: Advanced Multi-Layer Encryption
import sys, os, base64, zlib, random, hashlib, ctypes, threading, time

# Anti-Debug Protection
class _G:
    _a = False
    @classmethod
    def _s(cls):
        cls._a = True
        threading.Thread(target=cls._m, daemon=True).start()
    @classmethod
    def _m(cls):
        while cls._a:
            try:
                if hasattr(ctypes, 'windll'):
                    if ctypes.windll.kernel32.IsDebuggerPresent(): os._exit(1)
                    _b = ctypes.c_bool(False)
                    ctypes.windll.kernel32.CheckRemoteDebuggerPresent(
                        ctypes.windll.kernel32.GetCurrentProcess(), ctypes.byref(_b))
                    if _b.value: os._exit(1)
            except: pass
            time.sleep(random.uniform(3, 8))

try: _G._s()
except: pass

# Decryption Functions
def _x(d, k):
    return bytes(b ^ k[i % len(k)] for i, b in enumerate(d))

def _rc4(d, k):
    S = list(range(256))
    j = 0
    for i in range(256):
        j = (j + S[i] + k[i % len(k)]) % 256
        S[i], S[j] = S[j], S[i]
    i = j = 0
    r = []
    for b in d:
        i = (i + 1) % 256
        j = (j + S[i]) % 256
        S[i], S[j] = S[j], S[i]
        r.append(b ^ S[(S[i] + S[j]) % 256])
    return bytes(r)

def _al(d, k):
    # AES-like decrypt (inverse S-Box)
    isbox = bytes([
        0x52, 0x09, 0x6a, 0xd5, 0x30, 0x36, 0xa5, 0x38, 0xbf, 0x40, 0xa3, 0x9e, 0x81, 0xf3, 0xd7, 0xfb,
        0x7c, 0xe3, 0x39, 0x82, 0x9b, 0x2f, 0xff, 0x87, 0x34, 0x8e, 0x43, 0x44, 0xc4, 0xde, 0xe9, 0xcb,
        0x54, 0x7b, 0x94, 0x32, 0xa6, 0xc2, 0x23, 0x3d, 0xee, 0x4c, 0x95, 0x0b, 0x42, 0xfa, 0xc3, 0x4e,
        0x08, 0x2e, 0xa1, 0x66, 0x28, 0xd9, 0x24, 0xb2, 0x76, 0x5b, 0xa2, 0x49, 0x6d, 0x8b, 0xd1, 0x25,
        0x72, 0xf8, 0xf6, 0x64, 0x86, 0x68, 0x98, 0x16, 0xd4, 0xa4, 0x5c, 0xcc, 0x5d, 0x65, 0xb6, 0x92,
        0x6c, 0x70, 0x48, 0x50, 0xfd, 0xed, 0xb9, 0xda, 0x5e, 0x15, 0x46, 0x57, 0xa7, 0x8d, 0x9d, 0x84,
        0x90, 0xd8, 0xab, 0x00, 0x8c, 0xbc, 0xd3, 0x0a, 0xf7, 0xe4, 0x58, 0x05, 0xb8, 0xb3, 0x45, 0x06,
        0xd0, 0x2c, 0x1e, 0x8f, 0xca, 0x3f, 0x0f, 0x02, 0xc1, 0xaf, 0xbd, 0x03, 0x01, 0x13, 0x8a, 0x6b,
        0x3a, 0x91, 0x11, 0x41, 0x4f, 0x67, 0xdc, 0xea, 0x97, 0xf2, 0xcf, 0xce, 0xf0, 0xb4, 0xe6, 0x73,
        0x96, 0xac, 0x74, 0x22, 0xe7, 0xad, 0x35, 0x85, 0xe2, 0xf9, 0x37, 0xe8, 0x1c, 0x75, 0xdf, 0x6e,
        0x47, 0xf1, 0x1a, 0x71, 0x1d, 0x29, 0xc5, 0x89, 0x6f, 0xb7, 0x62, 0x0e, 0xaa, 0x18, 0xbe, 0x1b,
        0xfc, 0x56, 0x3e, 0x4b, 0xc6, 0xd2, 0x79, 0x20, 0x9a, 0xdb, 0xc0, 0xfe, 0x78, 0xcd, 0x5a, 0xf4,
        0x1f, 0xdd, 0xa8, 0x33, 0x88, 0x07, 0xc7, 0x31, 0xb1, 0x12, 0x10, 0x59, 0x27, 0x80, 0xec, 0x5f,
        0x60, 0x51, 0x7f, 0xa9, 0x19, 0xb5, 0x4a, 0x0d, 0x2d, 0xe5, 0x7a, 0x9f, 0x93, 0xc9, 0x9c, 0xef,
        0xa0, 0xe0, 0x3b, 0x4d, 0xae, 0x2a, 0xf5, 0xb0, 0xc8, 0xeb, 0xbb, 0x3c, 0x83, 0x53, 0x99, 0x61,
        0x17, 0x2b, 0x04, 0x7e, 0xba, 0x77, 0xd6, 0x26, 0xe1, 0x69, 0x14, 0x63, 0x55, 0x21, 0x0c, 0x7d
    ])
    k = (k * ((32 // len(k)) + 1))[:32]
    r = bytearray(d)
    for rnd in range(9, -1, -1):
        rk = bytes((k[i] + rnd * 17) % 256 for i in range(len(k)))
        r = bytearray(((b >> 3) | (b << 5)) & 0xFF for b in r)
        r = bytearray(b ^ rk[i % len(rk)] for i, b in enumerate(r))
        r = bytearray(isbox[b] for b in r)
    return bytes(r)

def _us(d, s):
    random.seed(s)
    idx = list(range(len(d)))
    random.shuffle(idx)
    r = [0] * len(d)
    for i, x in enumerate(idx): r[x] = d[i]
    return bytes(r)

def _r():
    try:
        # Reconstruct key and data
        _kp = {key_parts}
        _dp = {data_parts}
        _k = base64.b64decode(''.join(_kp))
        _d = base64.b64decode(''.join(_dp))
        
        # Multi-layer decryption
        _k1 = hashlib.sha256(_k + b'layer1').digest()
        _k2 = hashlib.sha256(_k + b'layer2').digest()
        _k3 = hashlib.sha256(_k + b'layer3').digest()
        
        # Reverse: unscramble -> XOR -> RC4 -> AES-like -> decompress
        _l5 = _us(_d, {seed})
        _l4 = _x(_l5, _k3)
        _l3 = _rc4(_l4, _k2)
        _l2 = _al(_l3, _k1)
        _c = zlib.decompress(_l2).decode('utf-8')
        
        exec(compile(_c, '<main>', 'exec'), globals())
    except Exception as _e:
        sys.exit(1)

if __name__ == '__main__': _r()
'''
    
    def protect_code(self, source_code: str) -> str:
        """保護源代碼"""
        data = source_code.encode('utf-8')
        
        if self.level == 1:
            # 基本加密
            compressed = zlib.compress(data, 9)
            encrypted = AdvancedEncryption.xor_encrypt(compressed, self.key)
            encrypted_b64 = base64.b64encode(encrypted).decode('ascii')
            return self._generate_decryption_code(encrypted_b64, {})
        
        elif self.level == 2:
            # 標準加密
            compressed = zlib.compress(data, 9)
            seed = int.from_bytes(self.key[:4], 'little')
            scrambled = AdvancedEncryption.scramble_bytes(compressed, seed)
            encrypted = AdvancedEncryption.rc4_encrypt(scrambled, self.key)
            encrypted_b64 = base64.b64encode(encrypted).decode('ascii')
            return self._generate_decryption_code(encrypted_b64, {'scramble_seed': seed})
        
        else:
            # 進階多層加密
            encrypted, metadata = AdvancedEncryption.multi_layer_encrypt(data, self.key)
            encrypted_b64 = base64.b64encode(encrypted).decode('ascii')
            return self._generate_decryption_code(encrypted_b64, metadata)
    
    def build(self, output_dir: str = None) -> bool:
        """執行加密打包"""
        level_names = {1: "Basic", 2: "Standard", 3: "Advanced"}
        level_name = level_names.get(self.level, "Unknown")
        
        print(f"\n{'='*60}")
        print(f"  WIM Driver Manager 加密打包系統")
        print(f"  加密級別: {self.level} ({level_name})")
        print(f"  時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"{'='*60}\n")
        
        # 使用臨時目錄避免 Google Drive 權限問題
        if output_dir is None:
            output_dir = tempfile.mkdtemp(prefix="WIM_Encrypted_Build_")
        
        try:
            # 讀取源代碼
            print("[1/4] 讀取源代碼...")
            if not os.path.exists(SOURCE_FILE):
                print(f"  ✗ 找不到源文件: {SOURCE_FILE}")
                return False
            
            with open(SOURCE_FILE, 'r', encoding='utf-8') as f:
                source_code = f.read()
            print(f"  ✓ 讀取完成 ({len(source_code):,} 字符)")
            
            # 加密保護
            print(f"[2/4] 應用 Level {self.level} 加密保護...")
            protected_code = self.protect_code(source_code)
            print(f"  ✓ 加密完成 (保護後 {len(protected_code):,} 字符)")
            
            # 寫入臨時文件
            temp_file = os.path.join(output_dir, "_encrypted_main.py")
            with open(temp_file, 'w', encoding='utf-8') as f:
                f.write(protected_code)
            
            # 執行 PyInstaller
            print("[3/4] 執行 PyInstaller 打包...")
            output_name = f"{APP_NAME}_Encrypted_L{self.level}"
            
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
            
            for imp in HIDDEN_IMPORTS:
                cmd.extend(['--hidden-import', imp])
            
            cmd.append(temp_file)
            
            result = subprocess.run(cmd, capture_output=True, text=True)
            
            if result.returncode != 0:
                print(f"  ✗ PyInstaller 失敗")
                print(result.stderr[-500:] if len(result.stderr) > 500 else result.stderr)
                return False
            
            print("  ✓ PyInstaller 打包成功")
            
            # 清理
            print("[4/4] 清理臨時文件...")
            os.remove(temp_file)
            shutil.rmtree(os.path.join(output_dir, "build"), ignore_errors=True)
            shutil.rmtree(os.path.join(output_dir, "spec"), ignore_errors=True)
            print("  ✓ 清理完成")
            
            # 輸出結果
            exe_path = os.path.join(output_dir, output_name, f"{output_name}.exe")
            if os.path.exists(exe_path):
                size_mb = os.path.getsize(exe_path) / (1024 * 1024)
                print(f"\n{'='*60}")
                print(f"  ✓ 打包成功!")
                print(f"  輸出位置: {os.path.join(output_dir, output_name)}")
                print(f"  執行檔: {output_name}.exe")
                print(f"  大小: {size_mb:.2f} MB")
                print(f"{'='*60}\n")
                return True
            else:
                print("  ✗ 找不到輸出執行檔")
                return False
            
        except Exception as e:
            print(f"  ✗ 打包過程發生錯誤: {e}")
            import traceback
            traceback.print_exc()
            return False


def main():
    """主函數"""
    import argparse
    
    parser = argparse.ArgumentParser(description='WIM Driver Manager 加密打包工具')
    parser.add_argument('-l', '--level', type=int, choices=[1, 2, 3], default=3,
                       help='加密級別: 1=基本, 2=標準, 3=進階 (預設: 3)')
    parser.add_argument('-o', '--output', type=str, default=None,
                       help='輸出目錄 (預設: 系統臨時目錄)')
    
    args = parser.parse_args()
    
    print("\n加密級別說明:")
    print("  Level 1 (基本): XOR 加密 + zlib 壓縮")
    print("  Level 2 (標準): RC4 加密 + 位元組打散 + zlib 壓縮")
    print("  Level 3 (進階): 多層加密 (AES-like + RC4 + XOR) + 打散 + 反調試")
    print("")
    
    builder = CodeProtectionBuilder(encryption_level=args.level)
    success = builder.build(output_dir=args.output)
    
    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
