# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WIM/Driver 管理工具 - 左側導航版本
清爽的淺色系設計
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
import configparser
import os
import sys
from datetime import datetime

# 確保工作目錄正確
APP_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(APP_DIR)
sys.path.insert(0, APP_DIR)

# 應用程式模組
from app.config import CONFIG_FILE, LOG_DIR
from app.wim_manager import WIMManager, get_dism_lock

# UI 模組
from ui.sidebar import Sidebar
from ui.wim_slot_card import WIMSlotCard
from ui.pages.driver_page import DriverPage
from ui.pages.script_page import ScriptPage
from ui.log_panel import LogPanel

# 設定 CustomTkinter
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class SidebarApp(ctk.CTk):
    """左側導航版本的 WIM/Driver 管理工具"""
    
    APP_TITLE = "WIM 管理工具"
    APP_VERSION = "2026.01.12"
    WINDOW_SIZE = "1100x620"
    MIN_SIZE = (1000, 570)
    WIM_SLOT_COUNT = 2
    
    def __init__(self):
        super().__init__()
        
        self.title(f"{self.APP_TITLE} v{self.APP_VERSION}")
        self.minsize(*self.MIN_SIZE)
        self.configure(fg_color="#ffffff")
        
        # 設定視窗位置（置中但稍微偏上，避免被工作列遮住）
        self._center_window()
        
        try:
            self.iconbitmap("icon.ico")
        except:
            pass
        
        # 初始化 log
        try:
            self._init_log_file()
        except Exception as e:
            print(f"初始化日誌失敗: {e}")
            self._log_file_path = None
        
        # 檢查管理員權限
        if not WIMManager.is_admin():
            self._elevate_and_exit()
            return
        
        # 設定檔
        self.cfg = configparser.ConfigParser()
        self._load_config()
        
        # 建立 UI
        try:
            self._build_ui()
        except Exception as e:
            messagebox.showerror("錯誤", f"建立 UI 失敗: {e}")
            import traceback
            traceback.print_exc()
            return
        
        # 載入設定
        self._load_settings()
        
        # 啟動訊息
        self._log("應用程式已啟動 (管理員權限)")
        self._log(f"版本: {self.APP_VERSION}")
        
        self.protocol("WM_DELETE_WINDOW", self._on_closing)
    
    def _center_window(self):
        """將視窗置中並稍微偏上（避免被工作列遮住）"""
        # 解析視窗大小
        w, h = map(int, self.WINDOW_SIZE.split('x'))
        
        # 取得螢幕大小
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        
        # 計算位置（水平置中，垂直稍微偏上）
        x = (screen_w - w) // 2
        y = max(20, (screen_h - h) // 2 - 50)  # 偏上 50px，最小距離頂部 20px
        
        self.geometry(f"{w}x{h}+{x}+{y}")
    
    def _init_log_file(self):
        """初始化 log 檔案"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        log_filename = f"session_{timestamp}.log"
        self._log_file_path = os.path.join(LOG_DIR, log_filename)
        os.makedirs(LOG_DIR, exist_ok=True)
        
        with open(self._log_file_path, 'w', encoding='utf-8') as f:
            f.write("=" * 60 + "\n")
            f.write(f"WIM/Driver 管理工具 - 操作日誌\n")
            f.write(f"啟動時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 60 + "\n\n")
    
    def _elevate_and_exit(self):
        """提升權限"""
        import ctypes
        
        if sys.platform == 'win32':
            try:
                script = os.path.abspath(sys.argv[0])
                
                ret = ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", sys.executable, f'"{script}"', APP_DIR, 1
                )
                
                if ret > 32:
                    self.destroy()
                    sys.exit(0)
                else:
                    messagebox.showerror("權限不足", "此程式需要管理員權限才能執行。")
                    self.destroy()
                    sys.exit(1)
            except Exception as e:
                messagebox.showerror("錯誤", f"無法提升權限: {e}")
                self.destroy()
                sys.exit(1)
    
    def _build_ui(self):
        """建立 UI"""
        # === 左側導航欄 ===
        self.sidebar = Sidebar(
            self,
            title="WIM 管理工具",
            version=self.APP_VERSION,
            on_navigate=self._on_navigate
        )
        self.sidebar.pack(side="left", fill="y")
        
        # 導航區分界線
        self.sidebar_separator = ctk.CTkFrame(
            self,
            width=1,
            fg_color="#e8f4fc"
        )
        self.sidebar_separator.pack(side="left", fill="y")
        
        # 新增導航項目
        self.sidebar.add_item("wim", "WIM 掛載管理", "🗂️")
        self.sidebar.add_item("driver", "驅動管理", "🔧")
        self.sidebar.add_item("script", "腳本管理", "📝")
        
        # === 右側內容區 ===
        self.content_area = ctk.CTkFrame(self, fg_color="#ffffff", corner_radius=0)
        self.content_area.pack(side="right", fill="both", expand=True)
        
        # === 頁面容器 ===
        self.pages_container = ctk.CTkFrame(self.content_area, fg_color="transparent")
        self.pages_container.pack(fill="both", expand=True, padx=20, pady=(15, 10))
        
        # 建立所有頁面
        self.pages = {}
        self._build_wim_page()
        self._build_driver_page()
        self._build_script_page()
        
        # === 底部日誌面板（固定在底部）===
        self._build_bottom_log_panel()
        
        # 顯示預設頁面（第一個）
        self._on_navigate("wim")
        
        # 啟動時讀取驅動資訊（延長到 1 秒確保設定已載入）
        self.after(1000, self._init_driver_list)
    
    def _build_wim_page(self):
        """建立 WIM 頁面"""
        page = ctk.CTkFrame(self.pages_container, fg_color="transparent")
        self.pages["wim"] = page
        
        # 標題
        ctk.CTkLabel(
            page,
            text="WIM 掛載管理",
            font=("Microsoft JhengHei UI", 24, "bold"),
            text_color="#37474f"
        ).pack(anchor="w", pady=(0, 20))
        
        # WIM 槽位容器（並排）
        slots_frame = ctk.CTkFrame(page, fg_color="transparent")
        slots_frame.pack(fill="both", expand=True)
        
        # 配置網格
        slots_frame.grid_columnconfigure(0, weight=1)
        slots_frame.grid_columnconfigure(1, weight=1)
        slots_frame.grid_rowconfigure(0, weight=1)
        
        # 建立 WIM 槽位
        self.wim_slots = []
        for i in range(self.WIM_SLOT_COUNT):
            slot = WIMSlotCard(
                slots_frame,
                slot_number=i + 1,
                on_log=self._log,
                on_status_change=self._on_mount_status_changed,
                get_other_slot_info=lambda idx=i: self._get_other_slot_info(idx)
            )
            slot.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 10, 0))
            self.wim_slots.append(slot)
    
    def _build_driver_page(self):
        """建立驅動頁面"""
        page = ctk.CTkFrame(self.pages_container, fg_color="transparent")
        self.pages["driver"] = page
        
        ctk.CTkLabel(
            page,
            text="驅動程式管理",
            font=("Microsoft JhengHei UI", 24, "bold"),
            text_color="#37474f"
        ).pack(anchor="w", pady=(0, 20))
        
        # 驅動頁面內容
        self.driver_page = DriverPage(
            page,
            on_log=self._log,
            get_mounted_dirs=self._get_mounted_dirs,
            show_header=False
        )
        self.driver_page.pack(fill="both", expand=True)
    
    def _build_script_page(self):
        """建立腳本管理頁面"""
        page = ctk.CTkFrame(self.pages_container, fg_color="transparent")
        self.pages["script"] = page
        
        # 腳本頁面內容
        self.script_page = ScriptPage(
            page,
            on_log=self._log,
            get_mounted_dirs=self._get_mounted_dirs
        )
        self.script_page.pack(fill="both", expand=True)
    
    def _build_bottom_log_panel(self):
        """建立底部日誌面板"""
        # 日誌容器
        self.log_container = ctk.CTkFrame(
            self.content_area,
            fg_color="#ffffff",
            corner_radius=8,
            border_width=1,
            border_color="#e1f0ff",
            height=150
        )
        self.log_container.pack(fill="x", padx=20, pady=(0, 15))
        self.log_container.pack_propagate(False)
        
        # 標題列
        header = ctk.CTkFrame(self.log_container, fg_color="transparent", height=32)
        header.pack(fill="x", padx=10, pady=(6, 0))
        header.pack_propagate(False)
        
        ctk.CTkLabel(
            header,
            text="系統日誌 / DISM 輸出",
            font=("Microsoft JhengHei UI", 16, "bold"),
            text_color="#37474f"
        ).pack(side="left")
        
        # 控制按鈕
        btn_frame = ctk.CTkFrame(header, fg_color="transparent")
        btn_frame.pack(side="right")
        
        # 最小化按鈕
        self.btn_minimize_log = ctk.CTkButton(
            btn_frame,
            text="─",
            font=("Segoe UI", 12),
            width=28,
            height=24,
            corner_radius=4,
            fg_color="transparent",
            hover_color="#e1f0ff",
            text_color="#78909c",
            command=self._toggle_log_panel
        )
        self.btn_minimize_log.pack(side="left", padx=(0, 4))
        
        # 清除按鈕
        ctk.CTkButton(
            btn_frame,
            text="✕",
            font=("Segoe UI", 12),
            width=28,
            height=24,
            corner_radius=4,
            fg_color="transparent",
            hover_color="#fce4ec",
            text_color="#78909c",
            command=self._clear_log
        ).pack(side="left")
        
        # 日誌內容區
        self.log_content = ctk.CTkFrame(self.log_container, fg_color="transparent")
        self.log_content.pack(fill="both", expand=True, padx=10, pady=(4, 8))
        
        # 日誌文字框
        self.log_text = tk.Text(
            self.log_content,
            wrap="word",
            font=("Consolas", 14),
            bg="#fafcff",
            fg="#37474f",
            relief="flat",
            padx=8,
            pady=6,
            state="disabled",
            cursor="arrow"
        )
        self.log_text.pack(side="left", fill="both", expand=True)
        
        # 滾動條
        scrollbar = ctk.CTkScrollbar(self.log_content, command=self.log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        # 配置標籤顏色
        self.log_text.tag_configure("TIMESTAMP", foreground="#90a4ae")
        self.log_text.tag_configure("SUCCESS", foreground="#22c55e")
        self.log_text.tag_configure("ERROR", foreground="#ef4444")
        self.log_text.tag_configure("WARNING", foreground="#f59e0b")
        self.log_text.tag_configure("INFO", foreground="#3b82f6")
        
        # 記錄展開狀態
        self._log_expanded = True
    
    def _toggle_log_panel(self):
        """切換日誌面板展開/收合"""
        if self._log_expanded:
            self.log_container.configure(height=36)
            self.log_content.pack_forget()
            self.btn_minimize_log.configure(text="□")
        else:
            self.log_container.configure(height=150)
            self.log_content.pack(fill="both", expand=True, padx=10, pady=(4, 8))
            self.btn_minimize_log.configure(text="─")
        self._log_expanded = not self._log_expanded
    
    def _clear_log(self):
        """清除日誌"""
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
    
    def _on_navigate(self, page_key: str):
        """導航切換"""
        # 隱藏所有頁面
        for key, page in self.pages.items():
            page.pack_forget()
        
        # 顯示選中頁面
        if page_key in self.pages:
            self.pages[page_key].pack(fill="both", expand=True)
        
        # 切換到驅動頁面時只重新整理選項（不自動載入驅動）
        if page_key == "driver" and hasattr(self, 'driver_page'):
            self.driver_page.refresh_targets(auto_load=False)
        
        # 切換到腳本頁面時重新整理
        if page_key == "script" and hasattr(self, 'script_page'):
            self.script_page.refresh_targets()
    
    def _init_driver_list(self):
        """初始化驅動程式列表（程式啟動時呼叫）"""
        print("[DEBUG] _init_driver_list called")
        if hasattr(self, 'driver_page'):
            # 啟動時自動載入驅動（如果有已掛載的映像）
            self.driver_page.refresh_targets(auto_load=True)
    
    def _on_mount_status_changed(self, slot_number: int, is_mounted: bool):
        """掛載狀態變更 - 處理掛載/卸載後驅動頁面更新"""
        if not hasattr(self, 'driver_page'):
            return
        
        if is_mounted:
            # 掛載時：自動刷新驅動頁面並讀取該槽位的驅動
            self._log(f"WIM#{slot_number} 掛載完成，正在更新驅動管理...")
            
            # 刷新目標選項
            self.driver_page.refresh_targets(auto_load=False)
            
            # 自動選擇剛掛載的 WIM 並載入驅動
            self.after(200, lambda: self._auto_select_and_load_drivers(slot_number))
        else:
            # 卸載時：清除驅動清單，檢查另一個 WIM 是否掛載
            other_slot = 2 if slot_number == 1 else 1
            other_mounted = False
            other_mount_dir = ""
            
            # 檢查另一個槽位
            if len(self.wim_slots) >= other_slot:
                other_idx = other_slot - 1
                other_mount_dir = self.wim_slots[other_idx].get_mount_dir()
                if other_mount_dir:
                    from app.wim_manager import WIMManager
                    other_mounted, _, _ = WIMManager.is_path_mounted(other_mount_dir)
            
            # 更新驅動頁面
            self.driver_page.on_wim_unmounted(slot_number, other_slot if other_mounted else None)
            self._log(f"WIM#{slot_number} 已卸載" + (f"，已切換到 WIM#{other_slot}" if other_mounted else ""))
    
    def _auto_select_and_load_drivers(self, preferred_slot: int):
        """自動選擇並載入驅動"""
        if not hasattr(self, 'driver_page'):
            return
        
        # 取得所有選項
        options = self.driver_page.combo_target.cget("values")
        
        # 優先選擇指定的槽位
        for opt in options:
            if f"WIM#{preferred_slot}" in opt and "✓ 已掛載" in opt:
                self.driver_page.combo_target.set(opt)
                self.driver_page._on_target_changed(opt)
                return
        
        # 如果指定槽位未掛載，選擇 #1（如果已掛載）
        if preferred_slot != 1:
            for opt in options:
                if "WIM#1" in opt and "✓ 已掛載" in opt:
                    self.driver_page.combo_target.set(opt)
                    self.driver_page._on_target_changed(opt)
                    return
    
    def _get_other_slot_info(self, current_idx: int):
        """取得其他槽位資訊"""
        for i, slot in enumerate(self.wim_slots):
            if i != current_idx:
                return slot.get_config()
        return None
    
    def _get_mounted_dirs(self) -> list:
        """取得已掛載目錄 - 返回 [(name, path, readonly), ...]"""
        dirs = []
        for i, slot in enumerate(self.wim_slots, 1):
            mount_dir = slot.get_mount_dir()
            readonly = slot.get_readonly()
            name = f"WIM#{i}"
            dirs.append((name, mount_dir, readonly))
        return dirs
    
    def _log(self, message: str):
        """寫入日誌"""
        # 寫入底部日誌面板
        if hasattr(self, 'log_text') and self.log_text:
            self.log_text.configure(state="normal")
            
            # 時間戳
            timestamp = datetime.now().strftime('[%H:%M:%S] ')
            self.log_text.insert("end", timestamp, "TIMESTAMP")
            
            # 判斷日誌類型並套用顏色
            tag = "INFO"
            if "✓" in message or "成功" in message or "完成" in message:
                tag = "SUCCESS"
            elif "✗" in message or "失敗" in message or "錯誤" in message:
                tag = "ERROR"
            elif "警告" in message or "注意" in message:
                tag = "WARNING"
            
            self.log_text.insert("end", message + "\n", tag)
            self.log_text.configure(state="disabled")
            self.log_text.see("end")
        
        # 寫入檔案
        if hasattr(self, '_log_file_path') and self._log_file_path:
            try:
                with open(self._log_file_path, 'a', encoding='utf-8') as f:
                    timestamp = datetime.now().strftime('%H:%M:%S')
                    f.write(f"[{timestamp}] {message}\n")
            except:
                pass
    
    def _load_config(self):
        """載入設定檔"""
        try:
            if os.path.exists(CONFIG_FILE):
                self.cfg.read(CONFIG_FILE, encoding='utf-8')
                print(f"[DEBUG] 設定檔已載入: {CONFIG_FILE}")
                for section in self.cfg.sections():
                    print(f"[DEBUG] [{section}]")
                    for key, value in self.cfg[section].items():
                        print(f"[DEBUG]   {key} = {value}")
            else:
                print(f"[DEBUG] 設定檔不存在: {CONFIG_FILE}")
        except Exception as e:
            print(f"載入設定檔失敗: {e}")
            import traceback
            traceback.print_exc()
    
    def _save_config(self):
        """儲存設定檔"""
        try:
            # 儲存 WIM 槽位設定
            for i, slot in enumerate(self.wim_slots, 1):
                section = f"WIM{i}"
                if section not in self.cfg:
                    self.cfg[section] = {}
                
                config = slot.get_config()
                self.cfg[section]["wim_path"] = config.get("wim_path", "")
                self.cfg[section]["mount_dir"] = config.get("mount_dir", "")
                self.cfg[section]["index"] = config.get("index", "")
                self.cfg[section]["readonly"] = str(config.get("readonly", True))
            
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                self.cfg.write(f)
            
            print(f"[DEBUG] 設定已儲存到: {CONFIG_FILE}")
            for i in range(1, 3):
                section = f"WIM{i}"
                if section in self.cfg:
                    print(f"[DEBUG] {section}: wim_path={self.cfg[section].get('wim_path', '')}, mount_dir={self.cfg[section].get('mount_dir', '')}")
        except Exception as e:
            print(f"儲存設定檔失敗: {e}")
            import traceback
            traceback.print_exc()
    
    def _load_settings(self):
        """載入設定"""
        try:
            for i, slot in enumerate(self.wim_slots, 1):
                section = f"WIM{i}"
                if section in self.cfg:
                    slot.set_config({
                        "wim_path": self.cfg[section].get("wim_path", ""),
                        "mount_dir": self.cfg[section].get("mount_dir", ""),
                        "index": self.cfg[section].get("index", ""),
                        "readonly": self.cfg[section].getboolean("readonly", True)
                    })
        except Exception as e:
            print(f"載入設定失敗: {e}")
    
    def _on_closing(self):
        """關閉視窗"""
        self._save_config()
        self._log("應用程式已關閉")
        self.quit()


def main():
    app = SidebarApp()
    app.mainloop()


if __name__ == "__main__":
    main()
