#!/usr/bin/env python3
"""
額外安全保護層
添加反調試、完整性檢查和運行時保護
"""

import os
import sys
import hashlib
import time
import random
import threading
import ctypes
from ctypes import wintypes

class SecurityProtection:
    def __init__(self):
        self.start_time = time.time()
        self.check_interval = random.uniform(30, 60)  # 隨機檢查間隔
        
    def anti_debug_check(self):
        """反調試檢查"""
        try:
            # 檢查是否有調試器附加
            kernel32 = ctypes.windll.kernel32
            if kernel32.IsDebuggerPresent():
                return False
                
            # 檢查進程名稱
            import psutil
            current_process = psutil.Process()
            suspicious_names = [
                'ollydbg', 'x64dbg', 'windbg', 'ida', 'cheatengine',
                'processhacker', 'procmon', 'wireshark', 'fiddler'
            ]
            
            for proc in psutil.process_iter(['name']):
                try:
                    if any(sus in proc.info['name'].lower() for sus in suspicious_names):
                        return False
                except:
                    continue
                    
        except:
            pass
        return True
    
    def integrity_check(self):
        """完整性檢查"""
        try:
            # 檢查文件大小和修改時間
            current_file = sys.executable
            stat = os.stat(current_file)
            
            # 簡單的完整性檢查（實際應用中可以更複雜）
            expected_size_min = 10000000  # 最小預期大小
            if stat.st_size < expected_size_min:
                return False
                
            # 檢查是否在虛擬機中運行
            vm_indicators = [
                'VBOX', 'VMWARE', 'QEMU', 'XEN', 'VIRTUAL'
            ]
            
            try:
                import wmi
                c = wmi.WMI()
                for system in c.Win32_ComputerSystem():
                    if any(vm in system.Model.upper() for vm in vm_indicators):
                        return False
            except:
                pass
                
        except:
            pass
        return True
    
    def runtime_protection(self):
        """運行時保護線程"""
        while True:
            time.sleep(self.check_interval)
            
            if not self.anti_debug_check():
                self._exit_protection()
            
            if not self.integrity_check():
                self._exit_protection()
            
            # 更新檢查間隔
            self.check_interval = random.uniform(30, 120)
    
    def _exit_protection(self):
        """保護性退出"""
        # 清理敏感數據
        try:
            import gc
            gc.collect()
        except:
            pass
        
        # 強制退出
        os._exit(1)
    
    def start_protection(self):
        """啟動保護"""
        if not self.anti_debug_check():
            self._exit_protection()
        
        if not self.integrity_check():
            self._exit_protection()
        
        # 啟動保護線程
        protection_thread = threading.Thread(target=self.runtime_protection, daemon=True)
        protection_thread.start()

# 創建全局保護實例
_protection = SecurityProtection()

def init_protection():
    """初始化保護系統"""
    _protection.start_protection()

if __name__ == "__main__":
    init_protection()
