#!/usr/bin/env python3
"""
簡單的代碼混淆工具
使用基本的字符串編碼和變量名混淆來保護代碼
"""

import ast
import base64
import codecs
import keyword
import random
import string
import os
import sys

class SimpleObfuscator:
    def __init__(self):
        self.var_map = {}
        self.string_pool = []
        self.obfuscated_names = set()
        
    def generate_random_name(self, prefix=""):
        """生成隨機變數名"""
        while True:
            name = prefix + ''.join(random.choices(
                string.ascii_letters + '_', 
                k=random.randint(8, 15)
            ))
            if name not in keyword.kwlist and name not in self.obfuscated_names:
                self.obfuscated_names.add(name)
                return name
    
    def encode_string(self, s):
        """編碼字符串"""
        if len(s) < 5:  # 短字符串不編碼
            return repr(s)
        
        # Base64 編碼
        encoded = base64.b64encode(s.encode()).decode()
        var_name = self.generate_random_name("_str_")
        self.string_pool.append(f"{var_name} = base64.b64decode('{encoded}').decode()")
        return var_name
    
    def obfuscate_file(self, input_file, output_file):
        """混淆Python文件"""
        with open(input_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 添加必要的導入
        imports = [
            "import base64",
            "import codecs",
            "import sys",
            "import os"
        ]
        
        # 創建解碼函數
        decode_func = '''
def _decode_data():
    """解碼字符串數據"""
    global ''' + ', '.join(f"_str_{i}" for i in range(10)) + '''
'''
        
        # 基本的字符串替換混淆
        obfuscated_content = content
        
        # 替換一些常見的字符串
        string_replacements = [
            ("'WIM'", "base64.b64decode('V0lN').decode()"),
            ("'Driver'", "base64.b64decode('RHJpdmVy').decode()"),
            ("'Error'", "base64.b64decode('RXJyb3I=').decode()"),
            ("'Success'", "base64.b64decode('U3VjY2Vzcw==').decode()"),
            ("'Mount'", "base64.b64decode('TW91bnQ=').decode()"),
            ("'Unmount'", "base64.b64decode('VW5tb3VudA==').decode()"),
        ]
        
        for original, replacement in string_replacements:
            obfuscated_content = obfuscated_content.replace(original, replacement)
        
        # 寫入混淆後的文件
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write("# Obfuscated by Simple Python Obfuscator\n")
            f.write("# -*- coding: utf-8 -*-\n\n")
            for imp in imports:
                f.write(f"{imp}\n")
            f.write("\n")
            f.write(obfuscated_content)

def main():
    if len(sys.argv) != 3:
        print("Usage: python obfuscate.py <input_file> <output_file>")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    
    if not os.path.exists(input_file):
        print(f"Error: Input file '{input_file}' not found")
        sys.exit(1)
    
    obfuscator = SimpleObfuscator()
    obfuscator.obfuscate_file(input_file, output_file)
    print(f"Obfuscation complete: {input_file} -> {output_file}")

if __name__ == "__main__":
    main()
