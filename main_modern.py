# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WIM/Driver 管理工具 - 現代化 UI 版本
使用 CustomTkinter 框架 + Accordion 折疊面板設計
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
import configparser
import os
import sys
from datetime import datetime

# 確保工作目錄正確（管理員權限啟動時會改變工作目錄）
APP_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(APP_DIR)
sys.path.insert(0, APP_DIR)

# 應用程式模組
from app.config import CONFIG_FILE, LOG_DIR
from app.wim_manager import WIMManager, get_dism_lock

# UI 模組
from ui.theme import theme_manager, ThemeManager, Fonts, Spacing
from ui.components import ModernButton, ModernCard, StatusBadge
from ui.collapsible import CollapsibleSection
from ui.log_panel import CollapsibleLogPanel
from ui.pages.wim_page import WIMSlot
from ui.pages.driver_page import DriverPage

# 設定 CustomTkinter
ctk.set_appearance_mode("dark")  # 預設深色模式
ctk.set_default_color_theme("blue")

# 全域 DISM 操作鎖
_dism_lock = get_dism_lock()


class ModernApp(ctk.CTk):
    """現代化 WIM/Driver 管理工具主程式"""
    
    # === 應用程式常數 ===
    APP_TITLE = "WIM/Driver 管理工具"
    APP_VERSION = "2.0"
    WINDOW_SIZE = "900x750"
    MIN_SIZE = (850, 700)
    WIM_SLOT_COUNT = 2  # WIM 槽位數量（可配置）
    
    def __init__(self):
        super().__init__()
        
        # 基本視窗設定
        self.title(self.APP_TITLE)
        self.geometry(self.WINDOW_SIZE)
        self.minsize(*self.MIN_SIZE)
        
        # 設定視窗圖示（如果有的話）
        try:
            self.iconbitmap("icon.ico")
        except:
            pass
        
        # 初始化 log 檔案
        self._init_log_file()
        
        # 檢查管理員權限
        if not WIMManager.is_admin():
            self._elevate_and_exit()
            return
        
        # 設定檔
        self.cfg = configparser.ConfigParser()
        self._load_config()
        
        # 建立 UI
        self._build_ui()
        
        # 載入設定
        self._load_settings()
        
        # 啟動訊息
        self._log("應用程式已啟動 (管理員權限)")
        self._log(f"版本: {self.APP_VERSION}")
        
        # 視窗關閉事件
        self.protocol("WM_DELETE_WINDOW", self._on_closing)
    
    def _init_log_file(self):
        """初始化 log 檔案"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        log_filename = f"session_{timestamp}.log"
        self._log_file_path = os.path.join(LOG_DIR, log_filename)
        
        # 確保目錄存在
        os.makedirs(LOG_DIR, exist_ok=True)
        
        # 寫入檔案開頭
        with open(self._log_file_path, 'w', encoding='utf-8') as f:
            f.write("=" * 60 + "\n")
            f.write(f"WIM/Driver 管理工具 - 操作日誌\n")
            f.write(f"啟動時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 60 + "\n\n")
    
    def _elevate_and_exit(self):
        """提升權限並重啟"""
        import ctypes
        import subprocess
        
        if sys.platform == 'win32':
            try:
                # 使用 ShellExecute 以管理員身份重新啟動
                script = os.path.abspath(sys.argv[0])
                params = ' '.join(sys.argv[1:])
                
                # 如果是 Python 腳本
                if script.endswith('.py'):
                    cmd = f'"{sys.executable}" "{script}" {params}'
                else:
                    cmd = f'"{script}" {params}'
                
                ret = ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", sys.executable, f'"{script}" {params}', None, 1
                )
                
                if ret > 32:  # 成功
                    self.quit()
                    sys.exit(0)
                else:
                    messagebox.showerror(
                        "權限不足",
                        "此程式需要管理員權限才能執行。\n請以管理員身份重新啟動。"
                    )
                    self.quit()
                    sys.exit(1)
            except Exception as e:
                messagebox.showerror("錯誤", f"無法提升權限: {e}")
                self.quit()
                sys.exit(1)
        else:
            messagebox.showerror("錯誤", "此程式僅支援 Windows 系統")
            self.quit()
            sys.exit(1)
    
    def _build_ui(self):
        """建立 UI"""
        # 設定主題顏色
        self.configure(fg_color=theme_manager.colors.background)
        
        # === 頂部標題列 ===
        self._build_header()
        
        # === 主內容區（可滾動）===
        self.main_scroll = ctk.CTkScrollableFrame(
            self,
            fg_color="transparent",
            scrollbar_button_color=theme_manager.colors.border,
            scrollbar_button_hover_color=theme_manager.colors.text_muted
        )
        self.main_scroll.pack(fill="both", expand=True, padx=12, pady=(8, 0))
        
        # 提升滾動速度（預設太慢）
        self._setup_scroll_speed(self.main_scroll)
        
        # === 1. WIM 掛載區塊 ===
        self.section_wim = CollapsibleSection(
            self.main_scroll,
            title="WIM 映像掛載",
            icon="🗂️",
            default_expanded=True
        )
        self.section_wim.pack(fill="x", pady=(0, 8))
        
        wim_content = self.section_wim.get_content_frame()
        
        # 動態建立 WIM 槽位
        self.wim_slots: list[WIMSlot] = []
        for i in range(1, self.WIM_SLOT_COUNT + 1):
            slot = WIMSlot(
                wim_content,
                slot_number=i,
                on_log=self._log,
                on_status_change=self._on_mount_status_changed,
                get_other_slot_info=lambda idx=i: self._get_other_slot_info(idx)
            )
            # 最後一個槽位不加底部間距
            pady = (0, 8) if i < self.WIM_SLOT_COUNT else (0, 0)
            slot.pack(fill="x", pady=pady)
            self.wim_slots.append(slot)
        
        # 向後相容的屬性別名
        if len(self.wim_slots) >= 1:
            self.wim_slot1 = self.wim_slots[0]
        if len(self.wim_slots) >= 2:
            self.wim_slot2 = self.wim_slots[1]
        
        # === 2. 驅動程式區塊 ===
        self.section_driver = CollapsibleSection(
            self.main_scroll,
            title="驅動程式管理",
            icon="🔧",
            default_expanded=False
        )
        self.section_driver.pack(fill="x", pady=(0, 8))
        
        driver_content = self.section_driver.get_content_frame()
        
        self.driver_page = DriverPage(
            driver_content,
            on_log=self._log,
            get_mounted_dirs=self._get_mounted_dirs,
            show_header=False  # 折疊面板已經有標題了
        )
        self.driver_page.pack(fill="both", expand=True)
        
        # === 3. 日誌區塊 ===
        self.log_panel = CollapsibleLogPanel(
            self.main_scroll,
            title="系統日誌",
            default_expanded=True,
            min_height=150
        )
        self.log_panel.pack(fill="x", pady=(0, 8))
        
        # === 4. 設定區塊 ===
        self.section_settings = CollapsibleSection(
            self.main_scroll,
            title="應用程式設定",
            icon="⚙️",
            default_expanded=False
        )
        self.section_settings.pack(fill="x", pady=(0, 8))
        
        self._build_settings_content(self.section_settings.get_content_frame())
        
        # === 底部狀態列 ===
        self._build_footer()
    
    def _build_header(self):
        """建立頂部標題列"""
        header = ctk.CTkFrame(
            self,
            fg_color=theme_manager.colors.card_bg,
            height=50,
            corner_radius=0
        )
        header.pack(fill="x")
        header.pack_propagate(False)
        
        # 左側：應用標題
        title_frame = ctk.CTkFrame(header, fg_color="transparent")
        title_frame.pack(side="left", padx=16)
        
        ctk.CTkLabel(
            title_frame,
            text="🗃️",
            font=("Segoe UI Emoji", 24)
        ).pack(side="left")
        
        ctk.CTkLabel(
            title_frame,
            text=self.APP_TITLE,
            font=Fonts.to_tuple(Fonts.TITLE),
            text_color=theme_manager.colors.text_primary
        ).pack(side="left", padx=(8, 0))
        
        ctk.CTkLabel(
            title_frame,
            text=f"v{self.APP_VERSION}",
            font=Fonts.to_tuple(Fonts.CAPTION),
            text_color=theme_manager.colors.text_muted
        ).pack(side="left", padx=(8, 0))
        
        # 右側：主題切換
        right_frame = ctk.CTkFrame(header, fg_color="transparent")
        right_frame.pack(side="right", padx=16)
        
        self.theme_switch = ctk.CTkSwitch(
            right_frame,
            text="深色模式",
            font=Fonts.to_tuple(Fonts.BODY),
            text_color=theme_manager.colors.text_secondary,
            progress_color=theme_manager.colors.primary,
            command=self._on_theme_toggle
        )
        self.theme_switch.pack(side="right")
        self.theme_switch.select()  # 預設深色
    
    def _build_settings_content(self, parent):
        """建立設定區塊內容"""
        # === 外觀設定 ===
        appearance_card = ModernCard(parent, padding=16)
        appearance_card.pack(fill="x", pady=(0, 12))
        appearance_content = appearance_card.get_content_frame()
        
        ctk.CTkLabel(
            appearance_content,
            text="🎨 外觀設定",
            font=Fonts.to_tuple(Fonts.TITLE_SMALL),
            text_color=theme_manager.colors.text_primary
        ).pack(anchor="w", pady=(0, 12))
        
        # 主題選擇
        theme_row = ctk.CTkFrame(appearance_content, fg_color="transparent")
        theme_row.pack(fill="x", pady=(0, 8))
        
        ctk.CTkLabel(
            theme_row,
            text="主題模式",
            font=Fonts.to_tuple(Fonts.BODY),
            text_color=theme_manager.colors.text_secondary,
            width=100,
            anchor="w"
        ).pack(side="left")
        
        self.var_theme = ctk.StringVar(value="dark")
        ctk.CTkSegmentedButton(
            theme_row,
            values=["light", "dark", "system"],
            variable=self.var_theme,
            command=self._on_theme_change
        ).pack(side="left", padx=(8, 0))
        
        # === 路徑設定 ===
        path_card = ModernCard(parent, padding=16)
        path_card.pack(fill="x", pady=(0, 12))
        path_content = path_card.get_content_frame()
        
        ctk.CTkLabel(
            path_content,
            text="📁 預設路徑",
            font=Fonts.to_tuple(Fonts.TITLE_SMALL),
            text_color=theme_manager.colors.text_primary
        ).pack(anchor="w", pady=(0, 12))
        
        # 設定檔路徑
        ctk.CTkLabel(
            path_content,
            text=f"設定檔: {CONFIG_FILE}",
            font=Fonts.to_tuple(Fonts.CAPTION),
            text_color=theme_manager.colors.text_muted
        ).pack(anchor="w")
        
        ctk.CTkLabel(
            path_content,
            text=f"日誌目錄: {LOG_DIR}",
            font=Fonts.to_tuple(Fonts.CAPTION),
            text_color=theme_manager.colors.text_muted
        ).pack(anchor="w", pady=(4, 0))
        
        # === 關於 ===
        about_card = ModernCard(parent, padding=16)
        about_card.pack(fill="x", pady=(0, 12))
        about_content = about_card.get_content_frame()
        
        ctk.CTkLabel(
            about_content,
            text="ℹ️ 關於",
            font=Fonts.to_tuple(Fonts.TITLE_SMALL),
            text_color=theme_manager.colors.text_primary
        ).pack(anchor="w", pady=(0, 12))
        
        ctk.CTkLabel(
            about_content,
            text=f"{self.APP_TITLE} v{self.APP_VERSION}",
            font=Fonts.to_tuple(Fonts.BODY),
            text_color=theme_manager.colors.text_primary
        ).pack(anchor="w")
        
        ctk.CTkLabel(
            about_content,
            text="現代化 Windows 映像與驅動程式管理工具",
            font=Fonts.to_tuple(Fonts.CAPTION),
            text_color=theme_manager.colors.text_muted
        ).pack(anchor="w", pady=(4, 0))
        
        ctk.CTkLabel(
            about_content,
            text="Developer: RexPeng",
            font=Fonts.to_tuple(Fonts.CAPTION),
            text_color=theme_manager.colors.text_muted
        ).pack(anchor="w", pady=(8, 0))
    
    def _build_footer(self):
        """建立底部狀態列"""
        footer = ctk.CTkFrame(
            self,
            fg_color=theme_manager.colors.card_bg,
            height=28,
            corner_radius=0
        )
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)
        
        # 左側狀態
        self.var_footer_status = ctk.StringVar(value="就緒")
        ctk.CTkLabel(
            footer,
            textvariable=self.var_footer_status,
            font=Fonts.to_tuple(Fonts.CAPTION),
            text_color=theme_manager.colors.text_muted
        ).pack(side="left", padx=12)
        
        # 右側資訊
        ctk.CTkLabel(
            footer,
            text="Developer: RexPeng",
            font=Fonts.to_tuple(Fonts.CAPTION),
            text_color=theme_manager.colors.text_muted
        ).pack(side="right", padx=12)
    
    # === 事件處理 ===
    
    def _on_theme_toggle(self):
        """切換主題"""
        if self.theme_switch.get():
            self.var_theme.set("dark")
            ctk.set_appearance_mode("dark")
        else:
            self.var_theme.set("light")
            ctk.set_appearance_mode("light")
    
    def _on_theme_change(self, value: str):
        """主題變更"""
        ctk.set_appearance_mode(value)
        self.theme_switch.select() if value == "dark" else self.theme_switch.deselect()
    
    def _on_mount_status_changed(self, slot_number: int = None, is_mounted: bool = None):
        """掛載狀態變更"""
        # 更新驅動程式頁面的目標選項
        self.driver_page.refresh_targets()
    
    def _get_mounted_dirs(self):
        """取得已掛載的目錄列表 [("WIM#1", path), ...]"""
        dirs = []
        for i, slot in enumerate(self.wim_slots, start=1):
            mount_dir = slot.get_mount_dir()
            if mount_dir and slot.is_mounted():
                dirs.append((f"WIM#{i}", mount_dir))
        return dirs
    
    def _get_other_slot_info(self, current_slot_number: int) -> dict:
        """
        取得其他槽位的資訊（供槽位間互相參照）
        
        Args:
            current_slot_number: 當前槽位編號 (1-based)
            
        Returns:
            其他槽位的資訊字典，包含 wim_file 和 mount_dir
        """
        # 簡單實作：返回另一個槽位（僅適用於 2 個槽位的情況）
        other_idx = 1 if current_slot_number == 2 else 0
        if other_idx < len(self.wim_slots):
            slot = self.wim_slots[other_idx]
            return slot.get_config()  # 使用公開方法，符合封裝原則
        return {}
    
    def _setup_scroll_speed(self, scrollable_frame, speed_multiplier: int = 3):
        """
        設定滾動速度
        
        Args:
            scrollable_frame: CTkScrollableFrame 實例
            speed_multiplier: 速度倍率（預設 3 倍）
        """
        # 取得內部的 canvas
        try:
            # CTkScrollableFrame 內部使用 _parent_canvas
            canvas = scrollable_frame._parent_canvas
            
            def on_mousewheel(event):
                # Windows: event.delta 通常是 120 或 -120
                # 將滾動量乘以倍率
                canvas.yview_scroll(int(-1 * (event.delta / 120) * speed_multiplier), "units")
            
            # 綁定滾輪事件到 canvas
            canvas.bind("<MouseWheel>", on_mousewheel)
            
            # 也綁定到內部框架
            scrollable_frame.bind("<MouseWheel>", on_mousewheel)
            
        except AttributeError:
            # 如果取不到 canvas，忽略
            pass
    
    def _on_closing(self):
        """視窗關閉事件"""
        self._save_config()
        self._log("應用程式關閉")
        self.quit()
    
    # === 日誌 ===
    
    def _log(self, message: str, level: str = None):
        """
        寫入日誌
        
        Args:
            message: 日誌訊息
            level: 日誌級別 (可選)
        """
        # 寫入 UI
        if hasattr(self, 'log_panel'):
            self.log_panel.log(message, level)
        
        # 寫入檔案
        try:
            timestamp = datetime.now().strftime('%H:%M:%S')
            with open(self._log_file_path, 'a', encoding='utf-8') as f:
                f.write(f"[{timestamp}] {message}\n")
        except:
            pass
    
    # === 設定檔 ===
    
    def _load_config(self):
        """載入設定檔"""
        try:
            if os.path.exists(CONFIG_FILE):
                self.cfg.read(CONFIG_FILE, encoding='utf-8')
        except Exception:
            pass
    
    def _cfg_get(self, section: str, option: str):
        """取得設定值"""
        if self.cfg.has_section(section) and self.cfg.has_option(section, option):
            return self.cfg.get(section, option)
        return None
    
    def _load_settings(self):
        """載入 UI 設定"""
        try:
            # 動態載入各 WIM 槽位設定
            section_names = ['WIM'] + [f'WIM{i}' for i in range(2, self.WIM_SLOT_COUNT + 1)]
            
            for i, (slot, section) in enumerate(zip(self.wim_slots, section_names)):
                if self.cfg.has_section(section):
                    slot.set_config({
                        'wim_file': self._cfg_get(section, 'wim_file') or '',
                        'mount_dir': self._cfg_get(section, 'mount_dir') or '',
                        'index': self._cfg_get(section, 'index') or '',
                        'readonly': self._cfg_get(section, 'readonly') in ('1', 'true', 'yes'),
                        'commit': self._cfg_get(section, 'unmount_commit') in ('1', 'true', 'yes')
                    })
            
            # 延遲檢查掛載狀態和更新驅動頁面
            self.after(1000, self._delayed_init)
            
        except Exception as e:
            self._log(f"載入設定時發生錯誤: {e}", "WARNING")
    
    def _delayed_init(self):
        """延遲初始化（等待 UI 完全載入）"""
        # 檢查所有 WIM 槽位掛載狀態
        for i, slot in enumerate(self.wim_slots):
            # 使用 after 來錯開檢查時間，避免同時執行
            self.after(i * 300, slot.check_status)
        
        # 更新驅動程式頁面的目標選項
        delay = len(self.wim_slots) * 300 + 200
        self.after(delay, self.driver_page.refresh_targets)
    
    def _save_config(self):
        """儲存設定檔"""
        try:
            # 動態儲存各 WIM 槽位設定
            section_names = ['WIM'] + [f'WIM{i}' for i in range(2, self.WIM_SLOT_COUNT + 1)]
            
            for slot, section in zip(self.wim_slots, section_names):
                if not self.cfg.has_section(section):
                    self.cfg.add_section(section)
                
                config = slot.get_config()
                self.cfg.set(section, 'wim_file', config.get('wim_file', ''))
                self.cfg.set(section, 'mount_dir', config.get('mount_dir', ''))
                self.cfg.set(section, 'index', config.get('index', ''))
                self.cfg.set(section, 'readonly', '1' if config.get('readonly', True) else '0')
                self.cfg.set(section, 'unmount_commit', '1' if config.get('commit', False) else '0')
            
            # 儲存
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                self.cfg.write(f)
                
        except Exception as e:
            self._log(f"儲存設定時發生錯誤: {e}", "ERROR")


def main():
    """主程式入口"""
    app = ModernApp()
    app.mainloop()


if __name__ == '__main__':
    main()
