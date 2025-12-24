# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
設定管理模組
- 統一管理應用程式設定常數
- 預設資料夾結構
- 設定檔讀寫功能
"""

import os
import sys
import configparser
from typing import Optional

# ========== 路徑常數 ==========

# 自動判斷 .py 或 .exe 模式，將設定檔放在執行檔同層
if getattr(sys, 'frozen', False):
    # 打包成 .exe 時
    SCRIPT_DIR = os.path.dirname(sys.executable)
else:
    # .py 腳本模式
    SCRIPT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 設定檔路徑
CONFIG_FILE = os.path.join(SCRIPT_DIR, 'settings.ini')

# ========== 預設資料夾結構 ==========

# 輸出根目錄
OUTPUT_DIR = os.path.join(SCRIPT_DIR, 'output')

# 驅動程式提取輸出目錄
DRIVER_EXPORT_DIR = os.path.join(OUTPUT_DIR, 'driver')

# 匯出清單目錄
LIST_EXPORT_DIR = os.path.join(OUTPUT_DIR, 'list')

# Log 目錄
LOG_DIR = os.path.join(SCRIPT_DIR, 'log')


def ensure_output_dirs():
    """確保輸出目錄結構存在"""
    os.makedirs(DRIVER_EXPORT_DIR, exist_ok=True)
    os.makedirs(LIST_EXPORT_DIR, exist_ok=True)
    os.makedirs(LOG_DIR, exist_ok=True)


# ========== 設定管理類別 ==========

class ConfigManager:
    """設定檔管理器"""
    
    def __init__(self, config_file: str = CONFIG_FILE):
        self.config_file = config_file
        self.cfg = configparser.ConfigParser()
        self._load()
    
    def _load(self):
        """載入設定檔"""
        try:
            if os.path.exists(self.config_file):
                self.cfg.read(self.config_file, encoding='utf-8')
        except Exception:
            pass
    
    def get(self, section: str, option: str, fallback: Optional[str] = None) -> Optional[str]:
        """取得設定值"""
        if self.cfg.has_section(section) and self.cfg.has_option(section, option):
            return self.cfg.get(section, option)
        return fallback
    
    def get_bool(self, section: str, option: str, fallback: bool = False) -> bool:
        """取得布林設定值"""
        value = self.get(section, option)
        if value is None:
            return fallback
        return value.lower() in ('1', 'true', 'yes', 'on')
    
    def set(self, section: str, option: str, value: str):
        """設定值"""
        if not self.cfg.has_section(section):
            self.cfg.add_section(section)
        self.cfg.set(section, option, value)
    
    def save(self):
        """儲存設定檔"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                self.cfg.write(f)
        except Exception:
            pass
    
    # ========== 便捷屬性 ==========
    
    @property
    def driver_export_dir(self) -> str:
        """驅動程式提取輸出目錄（使用固定預設值）"""
        return DRIVER_EXPORT_DIR
    
    @property
    def list_export_dir(self) -> str:
        """匯出清單目錄（使用固定預設值）"""
        return LIST_EXPORT_DIR
    
    @property
    def log_dir(self) -> str:
        """Log 目錄"""
        return LOG_DIR


# 初始化時確保目錄存在
ensure_output_dirs()
