# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Windows WIM/Driver 管理工具（tkinter）
- 特色：
  1) WIM 掛載/卸載工具：DISM 離線映像管理
  2) Driver 離線安裝工具：批量安裝驅動程式到離線映像
  3) 自動提升管理員權限
  4) 背景執行緒避免 GUI 卡頓，錯誤訊息人性化
  5) 設定持久化儲存

需求：
- Windows 10/11、Python 3.9+
- 管理員權限（DISM 操作需要）
- 以標準函式庫為主，無第三方相依

作者：Rex 專用版本
"""

import os
import re
import subprocess
import threading
import sys
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import configparser

# 導入模組化的類別
from app.wim_manager import WIMManager, get_dism_lock, is_dism_busy, set_dism_busy

# 全域 DISM 操作鎖 - 從 wim_manager 模組獲取
_dism_lock = get_dism_lock()
_dism_busy = False  # 本地追蹤變數

## 移除網路磁碟相依（專注於 WIM/Driver 功能）

# 設定檔路徑（儲存最近使用的路徑/選項）
# 自動判斷 .py 或 .exe 模式，將設定檔放在執行檔同層
if getattr(sys, 'frozen', False):
    # 打包成 .exe 時
    SCRIPT_DIR = os.path.dirname(sys.executable)
else:
    # .py 腳本模式
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, 'settings.ini')

# -----------------------------
# 工具層：WIM 掛載（使用 DISM）
# -----------------------------
class DriverManager:
    @staticmethod
    def _norm_path(p: str) -> str:
        try:
            return os.path.normpath(p)
        except Exception:
            return p

    @staticmethod
    def _run_dism(args: list[str]) -> tuple[int, str, str]:
        # 直接呼叫系統 dism
        try:
            cp = subprocess.run(["dism", "/English", *args], capture_output=True, text=True)
            return cp.returncode, cp.stdout or "", cp.stderr or ""
        except FileNotFoundError as e:
            return 9001, "", f"找不到 DISM：{e}"
        except Exception as e:
            return 9002, "", str(e)

    @staticmethod
    def add_driver_to_offline_image(mount_dir: str, driver_path: str, recurse: bool = True, force_unsigned: bool = False) -> tuple[bool, str]:
        """
        離線安裝驅動程式到已掛載的映像
        """
        m = DriverManager._norm_path(mount_dir)
        d = DriverManager._norm_path(driver_path)
        
        args = [
            "/Add-Driver",
            f"/Image:{m}",
            f"/Driver:{d}",
        ]
        
        # /Recurse 只能用於資料夾，不能用於單一 .inf 檔案
        is_inf_file = d.lower().endswith('.inf') and os.path.isfile(d)
        if recurse and not is_inf_file:
            args.append("/Recurse")
        
        if force_unsigned:
            args.append("/ForceUnsigned")
            
        rc, out, err = DriverManager._run_dism(args)
        if rc == 0:
            return True, "驅動程式安裝完成"
        return False, err or out

    @staticmethod  
    def export_drivers_from_offline_image(mount_dir: str, export_dir: str) -> tuple[bool, str]:
        """
        從已掛載的映像中萃取所有驅動程式
        """
        m = DriverManager._norm_path(mount_dir)
        e = DriverManager._norm_path(export_dir)
        
        # 確保匯出目錄存在
        os.makedirs(e, exist_ok=True)
        
        args = [
            "/Export-Driver",
            f"/Image:{m}",
            f"/Destination:{e}"
        ]
            
        rc, out, err = DriverManager._run_dism(args)
        if rc == 0:
            return True, "驅動程式萃取完成"
        return False, err or out

    @staticmethod
    def remove_driver_from_offline_image(mount_dir: str, driver_name: str) -> tuple[bool, str]:
        """
        從已掛載的映像中移除指定的驅動程式
        
        Args:
            mount_dir: 映像掛載路徑
            driver_name: 驅動程式名稱 (Published Name，如 oem0.inf)
        
        Returns:
            (success, message)
        """
        m = DriverManager._norm_path(mount_dir)
        
        args = [
            "/Remove-Driver",
            f"/Image:{m}",
            f"/Driver:{driver_name}"
        ]
        
        rc, out, err = DriverManager._run_dism(args)
        if rc == 0:
            return True, f"驅動程式 {driver_name} 移除成功"
        return False, err or out

    @staticmethod
    def remove_drivers_batch(mount_dir: str, driver_names: list[str], callback=None) -> tuple[int, int, list[str]]:
        """
        批量移除多個驅動程式
        
        Args:
            mount_dir: 映像掛載路徑
            driver_names: 驅動程式名稱列表
            callback: 進度回調函數 callback(current, total, driver_name, success, message)
        
        Returns:
            (success_count, fail_count, error_messages)
        """
        success_count = 0
        fail_count = 0
        errors = []
        total = len(driver_names)
        
        for i, driver_name in enumerate(driver_names):
            ok, msg = DriverManager.remove_driver_from_offline_image(mount_dir, driver_name)
            if ok:
                success_count += 1
            else:
                fail_count += 1
                errors.append(f"{driver_name}: {msg}")
            
            if callback:
                callback(i + 1, total, driver_name, ok, msg)
        
        return success_count, fail_count, errors

    @staticmethod
    def get_driver_details(mount_dir: str, driver_name: str) -> tuple[bool, dict, str]:
        """
        取得單一驅動程式的詳細資訊
        
        Args:
            mount_dir: 映像掛載路徑
            driver_name: 驅動程式名稱 (Published Name)
        
        Returns:
            (success, driver_info_dict, error_message)
        """
        m = DriverManager._norm_path(mount_dir)
        
        args = [
            "/Get-DriverInfo",
            f"/Image:{m}",
            f"/Driver:{driver_name}"
        ]
        
        rc, out, err = DriverManager._run_dism(args)
        if rc != 0:
            return False, {}, err or out
        
        # 解析輸出
        info = {"PublishedName": driver_name}
        for line in out.splitlines():
            line = line.strip()
            if ":" in line:
                key, _, value = line.partition(":")
                key = key.strip().replace(" ", "")
                value = value.strip()
                if key and value:
                    info[key] = value
        
        return True, info, ""

    @staticmethod
    def get_driver_info_from_path(driver_path: str) -> tuple[bool, list[dict], str]:
        """
        取得指定路徑中的驅動程式資訊
        """
        if not os.path.exists(driver_path):
            return False, [], "路徑不存在"
            
        drivers = []
        try:
            if os.path.isfile(driver_path) and driver_path.lower().endswith('.inf'):
                # 單一 .inf 檔案
                driver_info = {"path": driver_path, "name": os.path.basename(driver_path)}
                drivers.append(driver_info)
            elif os.path.isdir(driver_path):
                # 資料夾，搜尋所有 .inf 檔案
                for root, dirs, files in os.walk(driver_path):
                    for file in files:
                        if file.lower().endswith('.inf'):
                            full_path = os.path.join(root, file)
                            driver_info = {"path": full_path, "name": file, "folder": root}
                            drivers.append(driver_info)
            
            return True, drivers, f"找到 {len(drivers)} 個驅動程式檔案"
        except Exception as e:
            return False, [], f"掃描驅動程式時發生錯誤: {str(e)}"

    @staticmethod
    def get_drivers_in_offline_image(mount_dir: str) -> tuple[bool, list[dict], str]:
        """
        列出已安裝在離線映像中的驅動程式
        """
        m = DriverManager._norm_path(mount_dir)
        args = ["/Get-Drivers", f"/Image:{m}"]
        
        rc, out, err = DriverManager._run_dism(args)
        if rc != 0:
            return False, [], err or out
            
        drivers = DriverManager._parse_drivers(out)
        return True, drivers, ""

    @staticmethod
    def _parse_drivers(text: str) -> list[dict]:
        """
        解析 DISM 驅動程式輸出
        """
        drivers: list[dict] = []
        cur: dict | None = None
        
        # 欄位對應表：DISM 輸出欄位名稱 -> dict key
        field_map = {
            "Published Name": "PublishedName",
            "Original File Name": "OriginalFileName",
            "Class Name": "ClassName",
            "Provider Name": "Provider",
            "Date": "Date",
            "Version": "Version",
        }
        
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
                
            # 檢測新驅動程式開始
            if line.startswith("Published Name"):
                if cur:
                    drivers.append(cur)
                cur = {"PublishedName": "", "OriginalFileName": "", "ClassName": "", "Provider": "", "Date": "", "Version": ""}
                match = re.search(r"Published Name\s*:\s*(.*)", line, re.IGNORECASE)
                if match:
                    cur["PublishedName"] = match.group(1).strip()
                continue
                
            if cur is not None:
                # 嘗試匹配各欄位
                for dism_field, dict_key in field_map.items():
                    if dict_key == "PublishedName":  # 已處理
                        continue
                    pattern = f"{re.escape(dism_field)}\\s*:\\s*(.*)"
                    match = re.search(pattern, line, re.IGNORECASE)
                    if match:
                        cur[dict_key] = match.group(1).strip()
                        break
                        
        if cur:
            drivers.append(cur)
            
        return drivers

# -----------------------------
# GUI 層
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("WIM/Driver 管理工具")
        self.geometry("750x600")
        self.minsize(750, 580)
        
        # 初始化 log 檔案
        self._init_log_file()
        
        # 檢查並自動提升管理員權限
        if not WIMManager.is_admin():
            self._elevate_and_exit()
            return
            
        # 設定檔
        self.cfg = configparser.ConfigParser()
        self._load_config()
        self._build_ui()
        self._load_wim_config()  # 載入 WIM 分頁配置（在 UI 建構後）
        self._log("應用程式已啟動 (管理員權限)")  # 修改啟動訊息
    
    def _init_log_file(self):
        """初始化 log 檔案"""
        # 建立 log 資料夾
        log_dir = os.path.join(SCRIPT_DIR, "log")
        os.makedirs(log_dir, exist_ok=True)
        
        # 建立 log 檔案（使用日期+時間）
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        log_filename = f"session_{timestamp}.log"
        self._log_file_path = os.path.join(log_dir, log_filename)
        
        # 寫入檔案開頭
        with open(self._log_file_path, 'w', encoding='utf-8') as f:
            f.write(f"=" * 60 + "\n")
            f.write(f"WIM/Driver 管理工具 - 操作日誌\n")
            f.write(f"啟動時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"=" * 60 + "\n\n")

    # UI 組件
    def _build_ui(self):
        # === 標題列 / 工具列 ===
        toolbar_frame = ttk.Frame(self, padding=(8, 4))
        toolbar_frame.pack(fill=tk.X)
        
        # 左側：應用名稱
        title_label = ttk.Label(toolbar_frame, text="WIM/Driver 管理工具", font=('Arial', 11, 'bold'))
        title_label.pack(side=tk.LEFT)
        
        # 右側：設定按鈕
        ttk.Button(toolbar_frame, text="⚙ 設定", command=self._on_open_settings, width=8).pack(side=tk.RIGHT)
        
        # 分隔線
        ttk.Separator(self, orient='horizontal').pack(fill=tk.X)
        
        # 主容器
        main_frame = ttk.Frame(self, padding=4)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 建立 Notebook (分頁)
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 4))

        # 分頁 1：WIM 掛載（使用子分頁）
        wim_frame = ttk.Frame(self.notebook)
        self.notebook.add(wim_frame, text="WIM 掛載")
        
        # 在 WIM 掛載分頁中建立子分頁
        wim_sub_notebook = ttk.Notebook(wim_frame)
        wim_sub_notebook.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        
        # 子分頁 1：WIM 掛載 #1
        wim1_frame = ttk.Frame(wim_sub_notebook)
        wim_sub_notebook.add(wim1_frame, text="掛載 #1")
        self._build_wim1_tab(wim1_frame)
        
        # 子分頁 2：WIM 掛載 #2
        wim2_frame = ttk.Frame(wim_sub_notebook)
        wim_sub_notebook.add(wim2_frame, text="掛載 #2")
        self._build_wim2_tab(wim2_frame)

        # 分頁 2：Driver 管理（使用子分頁）
        driver_frame = ttk.Frame(self.notebook)
        self.notebook.add(driver_frame, text="Driver 管理")
        self._build_driver_tab(driver_frame)

        # Log 視窗（共用）
        log_frame = ttk.LabelFrame(main_frame, text="狀態 / 訊息", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.txt = tk.Text(log_frame, height=12, wrap=tk.WORD, font=('Consolas', 9))
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.txt.yview)
        self.txt.configure(yscrollcommand=scrollbar.set)
        
        self.txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt.configure(state=tk.DISABLED)

        # 版權資訊標籤（右下角）
        copyright_frame = ttk.Frame(main_frame)
        copyright_frame.pack(fill=tk.X)
        
        # 使用 Frame 來控制對齊
        copyright_label = tk.Label(copyright_frame, text="Developed by RexPeng", 
                                 font=('Arial', 8), fg='gray', anchor='e')
        copyright_label.pack(side=tk.RIGHT, padx=(0, 8), pady=(2, 4))

    # WIM 掛載分頁
    def _build_wim1_tab(self, parent: tk.Misc):
        # 使用 padding 的 frame
        content_frame = ttk.Frame(parent, padding=8)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # WIM 掛載 #1
        wim1_frame = ttk.Frame(content_frame, padding=10)
        wim1_frame.pack(fill=tk.X, pady=(0, 12))

        # 行 1：選擇 WIM 檔
        row1 = ttk.Frame(wim1_frame)
        row1.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(row1, text="WIM 檔案", width=12).pack(side=tk.LEFT)
        self.var_wim = tk.StringVar()
        ent_wim = ttk.Entry(row1, textvariable=self.var_wim, width=45)
        ent_wim.pack(side=tk.LEFT, padx=(8, 6), fill=tk.X, expand=True)
        
        # WIM 檔案操作按鈕組
        wim_btn_frame = ttk.Frame(row1)
        wim_btn_frame.pack(side=tk.RIGHT)
        ttk.Button(wim_btn_frame, text="瀏覽…", command=self._on_browse_wim).pack(side=tk.LEFT)
        ttk.Button(wim_btn_frame, text="讀取映像資訊", command=self._on_wim_info).pack(side=tk.LEFT, padx=(8, 0))

        # 行 2：Index / ReadOnly
        row2 = ttk.Frame(wim1_frame)
        row2.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(row2, text="Index", width=12).pack(side=tk.LEFT)
        self.var_wim_index = tk.StringVar()
        self.cbo_wim_index = ttk.Combobox(row2, textvariable=self.var_wim_index, width=8, state="readonly")
        self.cbo_wim_index.pack(side=tk.LEFT, padx=(8, 20))
        self.cbo_wim_index.bind('<<ComboboxSelected>>', self._on_wim1_index_changed)

        self.var_wim_readonly = tk.BooleanVar(value=True)
        ttk.Checkbutton(row2, text="唯讀掛載 (ReadOnly)", variable=self.var_wim_readonly, command=self._save_config).pack(side=tk.LEFT)

        # 行 3：掛載資料夾
        row3 = ttk.Frame(wim1_frame)
        row3.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(row3, text="掛載資料夾", width=12).pack(side=tk.LEFT)
        self.var_mount_dir = tk.StringVar()
        # 監聽掛載路徑變更，自動同步到 Driver 分頁
        self.var_mount_dir.trace_add('write', self._on_mount_dir_changed)
        ent_mdir = ttk.Entry(row3, textvariable=self.var_mount_dir, width=40)
        ent_mdir.pack(side=tk.LEFT, padx=(8, 6), fill=tk.X, expand=True)
        
        # 掛載資料夾操作按鈕組
        mount_btn_frame = ttk.Frame(row3)
        mount_btn_frame.pack(side=tk.RIGHT)
        ttk.Button(mount_btn_frame, text="選擇…", command=self._on_browse_mount_dir).pack(side=tk.LEFT)
        ttk.Button(mount_btn_frame, text="建立", command=self._on_create_mount_dir).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(mount_btn_frame, text="開啟", command=self._on_open_mount_dir).pack(side=tk.LEFT, padx=(6, 0))

        # 行 4：卸載選項
        row4 = ttk.Frame(wim1_frame)
        row4.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(row4, text="卸載模式", width=12).pack(side=tk.LEFT)
        self.var_unmount_commit = tk.BooleanVar(value=False)
        
        # 卸載選項組
        unmount_options_frame = ttk.Frame(row4)
        unmount_options_frame.pack(side=tk.LEFT, padx=(8, 0))
        ttk.Radiobutton(unmount_options_frame, text="丟棄變更 (/Discard)", variable=self.var_unmount_commit, value=False, command=self._save_config).pack(side=tk.LEFT)
        ttk.Radiobutton(unmount_options_frame, text="提交變更 (/Commit)", variable=self.var_unmount_commit, value=True, command=self._save_config).pack(side=tk.LEFT, padx=(20, 0))

        # 掛載狀態顯示
        row4b = ttk.Frame(wim1_frame)
        row4b.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(row4b, text="掛載狀態", width=12).pack(side=tk.LEFT)
        self.var_mount_status1 = tk.StringVar(value="未檢查")
        self.lbl_mount_status1 = ttk.Label(row4b, textvariable=self.var_mount_status1, foreground="gray")
        self.lbl_mount_status1.pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(row4b, text="檢查狀態", command=self._on_check_wim1_status, width=10).pack(side=tk.LEFT, padx=(12, 0))

        # 行 5：動作按鈕
        row5 = ttk.Frame(wim1_frame)
        row5.pack(fill=tk.X, pady=(0, 5))
        
        # WIM 操作按鈕組
        wim_action_frame = ttk.Frame(row5)
        wim_action_frame.pack(side=tk.LEFT)
        self.btn_wim_mount1 = ttk.Button(wim_action_frame, text="掛載 WIM", command=self._on_wim_mount, width=12)
        self.btn_wim_mount1.pack(side=tk.LEFT)
        self.btn_wim_unmount1 = ttk.Button(wim_action_frame, text="卸載 WIM", command=self._on_wim_unmount, width=12)
        self.btn_wim_unmount1.pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(wim_action_frame, text="關閉檔案總管", command=self._on_close_explorer).pack(side=tk.LEFT, padx=(8, 0))
        
        # 一鍵修復按鈕 - 整合所有診斷和修復功能
        smart_fix_btn = ttk.Button(wim_action_frame, text="🔧 一鍵修復", 
                                  command=self._on_smart_cleanup_fix, width=12)
        smart_fix_btn.pack(side=tk.LEFT, padx=(8, 0))
        
        # 添加工具提示
        tooltip_window = None  # 用於追蹤當前的工具提示窗口
        
        def show_tooltip(event):
            nonlocal tooltip_window
            # 如果已經有工具提示窗口存在，先關閉它
            if tooltip_window:
                tooltip_window.destroy()
                tooltip_window = None
            
            tooltip_window = tk.Toplevel()
            tooltip_window.wm_overrideredirect(True)
            tooltip_window.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
            
            # 使用 Frame 來控制寬度和添加邊距
            frame = tk.Frame(tooltip_window, bg="lightyellow", relief="solid", bd=1)
            frame.pack()
            
            # 分行顯示，避免文字過長
            lines = [
                "🔧 智能一鍵修復",
                "自動診斷並修復所有 WIM 掛載問題",
                "",
                "包含功能：",
                "• 狀態檢查與診斷", 
                "• 清理掛載衝突",
                "• 修復損壞掛載",
                "• 系統級清理"
            ]
            
            for line in lines:
                label = tk.Label(frame, text=line, bg="lightyellow", 
                               font=("Arial", 9), anchor="w", justify="left")
                label.pack(anchor="w", padx=8, pady=1)
            
            def hide_tooltip():
                nonlocal tooltip_window
                if tooltip_window:
                    tooltip_window.destroy()
                    tooltip_window = None
                    
            tooltip_window.after(4000, hide_tooltip)  # 延長顯示時間
        
        def hide_tooltip_on_leave(event):
            nonlocal tooltip_window
            if tooltip_window:
                tooltip_window.destroy()
                tooltip_window = None
        
        smart_fix_btn.bind("<Enter>", show_tooltip)
        smart_fix_btn.bind("<Leave>", hide_tooltip_on_leave)

    def _build_wim2_tab(self, parent: tk.Misc):
        # 使用 padding 的 frame
        content_frame = ttk.Frame(parent, padding=8)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # WIM 掛載 #2
        wim2_frame = ttk.Frame(content_frame, padding=10)
        wim2_frame.pack(fill=tk.X, pady=(0, 12))

        # 行 1：選擇 WIM 檔 #2
        row1_2 = ttk.Frame(wim2_frame)
        row1_2.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(row1_2, text="WIM 檔案", width=12).pack(side=tk.LEFT)
        self.var_wim2 = tk.StringVar()
        ent_wim2 = ttk.Entry(row1_2, textvariable=self.var_wim2, width=45)
        ent_wim2.pack(side=tk.LEFT, padx=(8, 6), fill=tk.X, expand=True)
        
        # WIM 檔案操作按鈕組 #2
        wim2_btn_frame = ttk.Frame(row1_2)
        wim2_btn_frame.pack(side=tk.RIGHT)
        ttk.Button(wim2_btn_frame, text="瀏覽…", command=self._on_browse_wim2).pack(side=tk.LEFT)
        ttk.Button(wim2_btn_frame, text="讀取映像資訊", command=self._on_wim_info2).pack(side=tk.LEFT, padx=(8, 0))

        # 行 2：Index / ReadOnly #2
        row2_2 = ttk.Frame(wim2_frame)
        row2_2.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(row2_2, text="Index", width=12).pack(side=tk.LEFT)
        self.var_wim_index2 = tk.StringVar()
        self.cbo_wim_index2 = ttk.Combobox(row2_2, textvariable=self.var_wim_index2, width=8, state="readonly")
        self.cbo_wim_index2.pack(side=tk.LEFT, padx=(8, 20))
        self.cbo_wim_index2.bind('<<ComboboxSelected>>', self._on_wim2_index_changed)

        self.var_wim_readonly2 = tk.BooleanVar(value=True)
        ttk.Checkbutton(row2_2, text="唯讀掛載 (ReadOnly)", variable=self.var_wim_readonly2, command=self._save_config).pack(side=tk.LEFT)

        # 行 3：掛載資料夾 #2
        row3_2 = ttk.Frame(wim2_frame)
        row3_2.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(row3_2, text="掛載資料夾", width=12).pack(side=tk.LEFT)
        self.var_mount_dir2 = tk.StringVar()
        # 監聽掛載路徑變更
        self.var_mount_dir2.trace_add('write', self._on_mount_dir2_changed)
        ent_mdir2 = ttk.Entry(row3_2, textvariable=self.var_mount_dir2, width=40)
        ent_mdir2.pack(side=tk.LEFT, padx=(8, 6), fill=tk.X, expand=True)
        
        # 掛載資料夾操作按鈕組 #2
        mount2_btn_frame = ttk.Frame(row3_2)
        mount2_btn_frame.pack(side=tk.RIGHT)
        ttk.Button(mount2_btn_frame, text="選擇…", command=self._on_browse_mount_dir2).pack(side=tk.LEFT)
        ttk.Button(mount2_btn_frame, text="建立", command=self._on_create_mount_dir2).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(mount2_btn_frame, text="開啟", command=self._on_open_mount_dir2).pack(side=tk.LEFT, padx=(6, 0))

        # 行 4：卸載選項 #2
        row4_2 = ttk.Frame(wim2_frame)
        row4_2.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(row4_2, text="卸載模式", width=12).pack(side=tk.LEFT)
        self.var_unmount_commit2 = tk.BooleanVar(value=False)
        
        # 卸載選項組 #2
        unmount2_options_frame = ttk.Frame(row4_2)
        unmount2_options_frame.pack(side=tk.LEFT, padx=(8, 0))
        ttk.Radiobutton(unmount2_options_frame, text="丟棄變更 (/Discard)", variable=self.var_unmount_commit2, value=False, command=self._save_config).pack(side=tk.LEFT)
        ttk.Radiobutton(unmount2_options_frame, text="提交變更 (/Commit)", variable=self.var_unmount_commit2, value=True, command=self._save_config).pack(side=tk.LEFT, padx=(20, 0))

        # 掛載狀態顯示 #2
        row4b_2 = ttk.Frame(wim2_frame)
        row4b_2.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(row4b_2, text="掛載狀態", width=12).pack(side=tk.LEFT)
        self.var_mount_status2 = tk.StringVar(value="未檢查")
        self.lbl_mount_status2 = ttk.Label(row4b_2, textvariable=self.var_mount_status2, foreground="gray")
        self.lbl_mount_status2.pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(row4b_2, text="檢查狀態", command=self._on_check_wim2_status, width=10).pack(side=tk.LEFT, padx=(12, 0))

        # 行 5：動作按鈕 #2
        row5_2 = ttk.Frame(wim2_frame)
        row5_2.pack(fill=tk.X, pady=(0, 5))
        
        # WIM 操作按鈕組 #2
        wim2_action_frame = ttk.Frame(row5_2)
        wim2_action_frame.pack(side=tk.LEFT)
        self.btn_wim_mount2 = ttk.Button(wim2_action_frame, text="掛載 WIM", command=self._on_wim_mount2, width=12)
        self.btn_wim_mount2.pack(side=tk.LEFT)
        self.btn_wim_unmount2 = ttk.Button(wim2_action_frame, text="卸載 WIM", command=self._on_wim_unmount2, width=12)
        self.btn_wim_unmount2.pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(wim2_action_frame, text="關閉檔案總管", command=self._on_close_explorer2).pack(side=tk.LEFT, padx=(8, 0))
        
        # 一鍵修復按鈕 - 整合所有診斷和修復功能
        smart_fix_btn2 = ttk.Button(wim2_action_frame, text="🔧 一鍵修復", 
                                   command=self._on_smart_cleanup_fix, width=12)
        smart_fix_btn2.pack(side=tk.LEFT, padx=(8, 0))
        
        # 添加工具提示
        tooltip_window2 = None  # 用於追蹤當前的工具提示窗口
        
        def show_tooltip2(event):
            nonlocal tooltip_window2
            # 如果已經有工具提示窗口存在，先關閉它
            if tooltip_window2:
                tooltip_window2.destroy()
                tooltip_window2 = None
            
            tooltip_window2 = tk.Toplevel()
            tooltip_window2.wm_overrideredirect(True)
            tooltip_window2.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
            
            # 使用 Frame 來控制寬度和添加邊距
            frame = tk.Frame(tooltip_window2, bg="lightyellow", relief="solid", bd=1)
            frame.pack()
            
            # 分行顯示，避免文字過長
            lines = [
                "🔧 智能一鍵修復",
                "自動診斷並修復所有 WIM 掛載問題",
                "",
                "包含功能：",
                "• 狀態檢查與診斷", 
                "• 清理掛載衝突",
                "• 修復損壞掛載",
                "• 系統級清理"
            ]
            
            for line in lines:
                label = tk.Label(frame, text=line, bg="lightyellow", 
                               font=("Arial", 9), anchor="w", justify="left")
                label.pack(anchor="w", padx=8, pady=1)
            
            def hide_tooltip():
                nonlocal tooltip_window2
                if tooltip_window2:
                    tooltip_window2.destroy()
                    tooltip_window2 = None
                    
            tooltip_window2.after(4000, hide_tooltip)  # 延長顯示時間
        
        def hide_tooltip2_on_leave(event):
            nonlocal tooltip_window2
            if tooltip_window2:
                tooltip_window2.destroy()
                tooltip_window2 = None
        
        smart_fix_btn2.bind("<Enter>", show_tooltip2)
        smart_fix_btn2.bind("<Leave>", hide_tooltip2_on_leave)

    # WIM 分頁配置載入
    def _load_wim_config(self):
        """載入 WIM 分頁的配置設定"""
        wim = self._cfg_get('WIM', 'wim_file')
        if wim:
            self.var_wim.set(wim)
        mdir = self._cfg_get('WIM', 'mount_dir')
        if mdir:
            self.var_mount_dir.set(mdir)
        idx = self._cfg_get('WIM', 'index')
        if idx:
            self.var_wim_index.set(idx)
        ro = self._cfg_get('WIM', 'readonly')
        if ro is not None:
            self.var_wim_readonly.set(ro.lower() in ('1', 'true', 'yes', 'on'))
        commit = self._cfg_get('WIM', 'unmount_commit')
        if commit is not None:
            self.var_unmount_commit.set(commit.lower() in ('1', 'true', 'yes', 'on'))
            
        # 載入設定值 - WIM #2
        wim2 = self._cfg_get('WIM2', 'wim_file')
        if wim2:
            self.var_wim2.set(wim2)
        mdir2 = self._cfg_get('WIM2', 'mount_dir')
        if mdir2:
            self.var_mount_dir2.set(mdir2)
        idx2 = self._cfg_get('WIM2', 'index')
        if idx2:
            self.var_wim_index2.set(idx2)
        ro2 = self._cfg_get('WIM2', 'readonly')
        if ro2 is not None:
            self.var_wim_readonly2.set(ro2.lower() in ('1', 'true', 'yes', 'on'))
        commit2 = self._cfg_get('WIM2', 'unmount_commit')
        if commit2 is not None:
            self.var_unmount_commit2.set(commit2.lower() in ('1', 'true', 'yes', 'on'))
        
        # 初始化按鈕狀態（預設禁用卸載按鈕）
        self._update_wim1_buttons(False)
        self._update_wim2_buttons(False)
        
        # 延遲自動檢查掛載狀態（讓 UI 先完成載入）
        self.after(1000, self._auto_check_mount_status)

    def _auto_check_mount_status(self):
        """自動檢查 WIM#1 和 WIM#2 的掛載狀態（依序執行）"""
        # 先檢查 WIM#1
        mdir1 = self.var_mount_dir.get().strip()
        if mdir1:
            self._on_check_wim1_status()
        
        # 延遲 500ms 再檢查 WIM#2，避免同時觸發多個 DISM 操作
        mdir2 = self.var_mount_dir2.get().strip()
        if mdir2:
            self.after(500, self._on_check_wim2_status)

    # Driver 管理分頁（整合版：驅動萃取、安裝、移除）
    def _build_driver_tab(self, parent: tk.Misc):
        """建立整合版驅動管理分頁"""
        # 主容器
        main_container = ttk.Frame(parent, padding=4)
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # === 頂部：目標映像選擇 ===
        header_frame = ttk.Frame(main_container)
        header_frame.pack(fill=tk.X, pady=(4, 8))
        
        ttk.Label(header_frame, text="目標映像", width=10).pack(side=tk.LEFT)
        
        # 下拉選單
        self.var_driver_list_mount_dir = tk.StringVar()
        self.cbo_driver_target = ttk.Combobox(header_frame, width=50, state="readonly")
        self.cbo_driver_target.pack(side=tk.LEFT, padx=(8, 6))
        self.cbo_driver_target.bind('<<ComboboxSelected>>', self._on_driver_target_selected)
        
        # 狀態顯示
        self.var_driver_status = tk.StringVar(value="未選擇")
        self.lbl_driver_status = ttk.Label(header_frame, textvariable=self.var_driver_status, width=12)
        self.lbl_driver_status.pack(side=tk.LEFT, padx=(8, 0))
        
        # === 搜尋區域（右側）===
        search_frame = ttk.Frame(header_frame)
        search_frame.pack(side=tk.RIGHT, padx=(8, 0))
        
        ttk.Label(search_frame, text="🔍").pack(side=tk.LEFT)
        self.var_driver_search = tk.StringVar()
        self.ent_driver_search = ttk.Entry(search_frame, textvariable=self.var_driver_search, width=20)
        self.ent_driver_search.pack(side=tk.LEFT, padx=(4, 0))
        self.ent_driver_search.bind('<Return>', lambda e: self._on_driver_search())
        self.ent_driver_search.bind('<Escape>', lambda e: self._on_clear_search())
        
        self.btn_driver_search = ttk.Button(search_frame, text="搜尋", command=self._on_driver_search, width=6)
        self.btn_driver_search.pack(side=tk.LEFT, padx=(4, 0))
        
        # === 驅動清單 Treeview ===
        tree_frame = ttk.Frame(main_container)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 4))
        
        # 新增勾選欄位
        columns = ("select", "name", "inf", "provider", "version", "date", "class")
        self.driver_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="extended")
        
        # 設定欄位標題 - 勾選欄點擊可全選/取消全選
        self._driver_select_all = False  # 追蹤全選狀態
        self.driver_tree.heading("select", text="☐ 選取", command=self._on_toggle_select_all)
        self.driver_tree.heading("name", text="驅動名稱", command=lambda: self._sort_driver_tree("name"))
        self.driver_tree.heading("inf", text="INF 檔案", command=lambda: self._sort_driver_tree("inf"))
        self.driver_tree.heading("provider", text="提供者", command=lambda: self._sort_driver_tree("provider"))
        self.driver_tree.heading("version", text="版本", command=lambda: self._sort_driver_tree("version"))
        self.driver_tree.heading("date", text="日期", command=lambda: self._sort_driver_tree("date"))
        self.driver_tree.heading("class", text="類型", command=lambda: self._sort_driver_tree("class"))
        
        # 設定欄位寬度
        self.driver_tree.column("select", width=60, minwidth=50, anchor="center")
        self.driver_tree.column("name", width=180, minwidth=120)
        self.driver_tree.column("inf", width=90, minwidth=70)
        self.driver_tree.column("provider", width=150, minwidth=100)
        self.driver_tree.column("version", width=100, minwidth=80)
        self.driver_tree.column("date", width=100, minwidth=80)
        self.driver_tree.column("class", width=100, minwidth=80)
        
        # 點擊行切換勾選狀態
        self.driver_tree.bind("<ButtonRelease-1>", self._on_driver_row_click)
        
        # 滾動條
        tree_scroll_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.driver_tree.yview)
        tree_scroll_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.driver_tree.xview)
        self.driver_tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        
        # 佈局
        self.driver_tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll_y.grid(row=0, column=1, sticky="ns")
        tree_scroll_x.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # === 狀態列 ===
        status_frame = ttk.Frame(main_container)
        status_frame.pack(fill=tk.X, pady=(0, 4))
        self.var_driver_list_status = tk.StringVar(value="請選擇已掛載的映像")
        ttk.Label(status_frame, textvariable=self.var_driver_list_status, foreground="gray").pack(side=tk.LEFT)
        
        # === 操作按鈕區 ===
        btn_frame = ttk.Frame(main_container)
        btn_frame.pack(fill=tk.X, pady=(4, 0))
        
        # 左側：主要功能
        left_btn_frame = ttk.Frame(btn_frame)
        left_btn_frame.pack(side=tk.LEFT)
        
        self.btn_extract_all = ttk.Button(left_btn_frame, text="📤 提取", command=self._on_extract_selected_drivers, width=10)
        self.btn_extract_all.pack(side=tk.LEFT)
        self.btn_add_driver = ttk.Button(left_btn_frame, text="➕ 新增", command=self._on_add_driver_dialog, width=10)
        self.btn_add_driver.pack(side=tk.LEFT, padx=(8, 0))
        self.btn_remove_drivers = ttk.Button(left_btn_frame, text="🗑 移除", command=self._on_remove_selected_drivers, width=10)
        self.btn_remove_drivers.pack(side=tk.LEFT, padx=(8, 0))
        
        # 右側：輔助功能
        right_btn_frame = ttk.Frame(btn_frame)
        right_btn_frame.pack(side=tk.RIGHT)
        
        ttk.Button(right_btn_frame, text="查看詳情", command=self._on_view_driver_details, width=10).pack(side=tk.LEFT)
        ttk.Button(right_btn_frame, text="匯出清單", command=self._on_export_driver_list, width=10).pack(side=tk.LEFT, padx=(8, 0))
        
        # 初始化排序狀態
        self._driver_tree_sort_col = None
        self._driver_tree_sort_reverse = False
        self._driver_published_names = {}
        
        # 載入設定
        self._load_driver_config()

    def _get_wim_options(self) -> list[tuple[str, str, str]]:
        """
        取得 WIM 映像選項列表
        Returns: [(顯示文字, 掛載路徑, 狀態), ...]
        """
        options = []
        
        # WIM#1
        mdir1 = self.var_mount_dir.get().strip() if hasattr(self, 'var_mount_dir') else ""
        if mdir1:
            is_mounted, _, _ = WIMManager.is_path_mounted(mdir1)
            status = "✓ 已掛載" if is_mounted else "○ 未掛載"
            options.append((f"WIM#1 - {mdir1} ({status})", mdir1, "mounted" if is_mounted else "not_mounted"))
        else:
            options.append(("WIM#1 - (未設定)", "", "not_set"))
        
        # WIM#2
        mdir2 = self.var_mount_dir2.get().strip() if hasattr(self, 'var_mount_dir2') else ""
        if mdir2:
            is_mounted, _, _ = WIMManager.is_path_mounted(mdir2)
            status = "✓ 已掛載" if is_mounted else "○ 未掛載"
            options.append((f"WIM#2 - {mdir2} ({status})", mdir2, "mounted" if is_mounted else "not_mounted"))
        else:
            options.append(("WIM#2 - (未設定)", "", "not_set"))
        
        return options

    def _update_wim_combobox(self, combobox: ttk.Combobox, status_var: tk.StringVar, status_label: ttk.Label):
        """更新 WIM 下拉選單選項"""
        options = self._get_wim_options()
        values = [opt[0] for opt in options]
        combobox['values'] = values
        
        # 更新狀態顯示
        current_idx = combobox.current()
        if current_idx >= 0 and current_idx < len(options):
            _, path, status = options[current_idx]
            self._update_wim_status_display(status_var, status_label, path, status)

    def _update_wim_status_display(self, status_var: tk.StringVar, status_label: ttk.Label, path: str, status: str):
        """更新掛載狀態顯示"""
        if status == "mounted":
            status_var.set("✓ 已掛載")
            status_label.configure(foreground="green")
        elif status == "not_mounted":
            status_var.set("○ 未掛載")
            status_label.configure(foreground="orange")
        elif status == "not_set":
            status_var.set("未設定")
            status_label.configure(foreground="gray")
        else:
            status_var.set("未知")
            status_label.configure(foreground="gray")

    def _on_wim_combobox_selected(self, event, path_var: tk.StringVar, status_var: tk.StringVar, status_label: ttk.Label):
        """當 WIM 下拉選單選擇變更時"""
        combobox = event.widget
        current_idx = combobox.current()
        options = self._get_wim_options()
        
        if current_idx >= 0 and current_idx < len(options):
            _, path, status = options[current_idx]
            path_var.set(path)
            self._update_wim_status_display(status_var, status_label, path, status)

    def _on_driver_target_selected(self, event):
        """驅動管理目標映像選擇變更"""
        self._on_wim_combobox_selected(event, self.var_driver_list_mount_dir, self.var_driver_status, self.lbl_driver_status)
        self._update_driver_buttons_state()
        
        # 先清空現有清單
        for item in self.driver_tree.get_children():
            self.driver_tree.delete(item)
        self._driver_published_names = {}
        
        # 檢查選擇的映像狀態
        options = self._get_wim_options()
        current_idx = self.cbo_driver_target.current()
        if current_idx >= 0 and current_idx < len(options):
            _, path, status = options[current_idx]
            if status == "mounted" and path:
                # 已掛載 -> 自動載入驅動清單
                self._on_refresh_driver_list()
            else:
                # 未掛載 -> 顯示提示
                self.var_driver_list_status.set("映像未掛載，請先掛載後再操作")

    def _update_driver_buttons_state(self):
        """更新驅動管理按鈕狀態"""
        options = self._get_wim_options()
        current_idx = self.cbo_driver_target.current() if hasattr(self, 'cbo_driver_target') else -1
        
        is_mounted = False
        if current_idx >= 0 and current_idx < len(options):
            _, path, status = options[current_idx]
            if status == "mounted" and path:
                is_mounted = True
        
        state = "normal" if is_mounted else "disabled"
        if hasattr(self, 'btn_extract_all'):
            self.btn_extract_all.configure(state=state)
        if hasattr(self, 'btn_add_driver'):
            self.btn_add_driver.configure(state=state)
        if hasattr(self, 'btn_remove_drivers'):
            self.btn_remove_drivers.configure(state=state)

    def _on_toggle_select_all(self):
        """點擊標題列切換全選/取消全選"""
        self._driver_select_all = not self._driver_select_all
        
        # 更新標題顯示
        if self._driver_select_all:
            self.driver_tree.heading("select", text="☑ 選取")
        else:
            self.driver_tree.heading("select", text="☐ 選取")
        
        # 更新所有行的勾選狀態
        for item in self.driver_tree.get_children():
            values = list(self.driver_tree.item(item, 'values'))
            values[0] = "☑" if self._driver_select_all else "☐"
            self.driver_tree.item(item, values=values)
        
        # 更新選取計數
        self._update_selection_count()

    def _on_driver_row_click(self, event):
        """點擊行切換勾選狀態"""
        # 檢查是否點擊在勾選欄
        region = self.driver_tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        
        column = self.driver_tree.identify_column(event.x)
        if column != "#1":  # 第一欄（select）
            return
        
        item = self.driver_tree.identify_row(event.y)
        if not item:
            return
        
        # 切換勾選狀態
        values = list(self.driver_tree.item(item, 'values'))
        values[0] = "☐" if values[0] == "☑" else "☑"
        self.driver_tree.item(item, values=values)
        
        # 更新選取計數和標題狀態
        self._update_selection_count()

    def _update_selection_count(self):
        """更新選取計數顯示"""
        total = len(self.driver_tree.get_children())
        selected = sum(1 for item in self.driver_tree.get_children() 
                      if self.driver_tree.item(item, 'values')[0] == "☑")
        
        if total > 0:
            self.var_driver_list_status.set(f"共 {total} 個驅動程式，已選取 {selected} 個")
            
            # 更新標題狀態
            if selected == total:
                self._driver_select_all = True
                self.driver_tree.heading("select", text="☑ 選取")
            elif selected == 0:
                self._driver_select_all = False
                self.driver_tree.heading("select", text="☐ 選取")
            else:
                self._driver_select_all = False
                self.driver_tree.heading("select", text="☐ 選取")

    def _get_checked_drivers(self) -> list[str]:
        """取得所有已勾選的驅動程式 PublishedName 列表"""
        checked = []
        for item in self.driver_tree.get_children():
            values = self.driver_tree.item(item, 'values')
            if values[0] == "☑":
                # 從對照表取得 PublishedName
                if item in self._driver_published_names:
                    checked.append(self._driver_published_names[item])
        return checked

    def _load_driver_config(self):
        """載入驅動程式相關設定"""
        # 載入安裝設定
        self.var_driver_source = tk.StringVar()
        driver_source = self._cfg_get('DRIVER', 'source_path')
        if driver_source:
            self.var_driver_source.set(driver_source)
        
        self.var_driver_recurse = tk.BooleanVar(value=True)
        recurse = self._cfg_get('DRIVER', 'recurse')
        if recurse is not None:
            self.var_driver_recurse.set(recurse.lower() in ('1', 'true', 'yes', 'on'))
        
        self.var_driver_force_unsigned = tk.BooleanVar(value=False)
        force_unsigned = self._cfg_get('DRIVER', 'force_unsigned')
        if force_unsigned is not None:
            self.var_driver_force_unsigned.set(force_unsigned.lower() in ('1', 'true', 'yes', 'on'))
        
        # 載入萃取設定
        self.var_extract_output = tk.StringVar()
        extract_output = self._cfg_get('EXTRACT', 'output_path') 
        if extract_output:
            self.var_extract_output.set(extract_output)

        # 初始化下拉選單
        self._init_driver_combobox()

    def _init_driver_combobox(self):
        """初始化 Driver 分頁的下拉選單"""
        # 延遲更長時間，確保狀態檢查完成後再執行
        self.after(2000, self._do_init_driver_combobox)

    def _do_init_driver_combobox(self):
        """實際初始化下拉選單"""
        global _dism_busy
        
        options = self._get_wim_options()
        values = [opt[0] for opt in options]
        
        if hasattr(self, 'cbo_driver_target'):
            self.cbo_driver_target['values'] = values
            # 預設選擇第一個已掛載的映像
            for i, (_, path, status) in enumerate(options):
                if status == "mounted":
                    self.cbo_driver_target.current(i)
                    self.var_driver_list_mount_dir.set(path)
                    self._update_wim_status_display(self.var_driver_status, self.lbl_driver_status, path, status)
                    # 只有在沒有其他 DISM 操作時才自動載入驅動清單
                    if not _dism_busy:
                        self._on_refresh_driver_list()
                    break
            else:
                # 沒有已掛載的，選擇第一個
                if values:
                    self.cbo_driver_target.current(0)
                    _, path, status = options[0]
                    self.var_driver_list_mount_dir.set(path)
                    self._update_wim_status_display(self.var_driver_status, self.lbl_driver_status, path, status)
            
            self._update_driver_buttons_state()

    def _refresh_all_driver_comboboxes(self, auto_load_drivers: bool = False):
        """重新整理所有 Driver 下拉選單並自動選擇已掛載的映像
        
        Args:
            auto_load_drivers: 是否自動載入驅動清單 (預設 False 避免 DISM 衝突)
        """
        options = self._get_wim_options()
        values = [opt[0] for opt in options]
        
        # 找出第一個已掛載的映像索引
        mounted_idx = -1
        for i, (_, path, status) in enumerate(options):
            if status == "mounted" and path:
                mounted_idx = i
                break
        
        # 更新驅動管理下拉選單
        if hasattr(self, 'cbo_driver_target'):
            self.cbo_driver_target['values'] = values
            if mounted_idx >= 0:
                self.cbo_driver_target.current(mounted_idx)
                _, path, status = options[mounted_idx]
                self.var_driver_list_mount_dir.set(path)
                self._update_wim_status_display(self.var_driver_status, self.lbl_driver_status, path, status)
                # 只有明確要求時才自動載入驅動清單
                if auto_load_drivers:
                    self._on_refresh_driver_list()
            self._update_driver_buttons_state()

    # 工具方法
    def _log(self, msg: str):
        ts = datetime.now().strftime('%H:%M:%S')
        log_line = f"[{ts}] {msg}"
        
        # 寫入 UI
        self.txt.configure(state=tk.NORMAL)
        self.txt.insert(tk.END, log_line + "\n")
        self.txt.see(tk.END)
        self.txt.configure(state=tk.DISABLED)
        
        # 寫入檔案
        if hasattr(self, '_log_file_path') and self._log_file_path:
            try:
                with open(self._log_file_path, 'a', encoding='utf-8') as f:
                    f.write(log_line + "\n")
            except Exception:
                pass  # 忽略寫入錯誤

    def show_error_with_advice(self, title: str, error_message: str):
        """
        顯示錯誤訊息並提供針對性建議
        """
        error_type, advice, solutions = WIMManager.get_error_solution_advice(error_message)
        
        # 構建完整的錯誤訊息
        full_message = f"錯誤詳情:\n{error_message}\n\n"
        full_message += f"錯誤類型: {error_type}\n"
        full_message += f"說明: {advice}\n\n"
        full_message += "建議解決方案:\n"
        for solution in solutions:
            full_message += f"{solution}\n"
        
        # 使用自定義對話框顯示
        dialog = tk.Toplevel(self)
        dialog.title(f"{title} - 解決建議")
        dialog.geometry("600x400")
        dialog.resizable(True, True)
        dialog.grab_set()  # 模態對話框
        
        # 主框架
        main_frame = ttk.Frame(dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 錯誤圖標和標題
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(title_frame, text="⚠️", font=("Arial", 24)).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(title_frame, text=f"{title} - {error_type}", 
                 font=("Arial", 14, "bold"), foreground="red").pack(side=tk.LEFT)
        
        # 滾動文本框
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 創建文本框和滾動條
        text_widget = tk.Text(text_frame, wrap=tk.WORD, font=("Consolas", 10))
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 插入文本內容
        text_widget.insert("1.0", full_message)
        text_widget.configure(state="disabled")  # 只讀
        
        # 按鈕框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        # 複製到剪貼板按鈕
        def copy_to_clipboard():
            dialog.clipboard_clear()
            dialog.clipboard_append(full_message)
            messagebox.showinfo("已複製", "錯誤訊息和建議已複製到剪貼板")
        
        ttk.Button(button_frame, text="複製內容", 
                  command=copy_to_clipboard).pack(side=tk.LEFT, padx=(0, 10))
        
        # 關閉按鈕
        ttk.Button(button_frame, text="關閉", 
                  command=dialog.destroy).pack(side=tk.RIGHT)
        
        # 居中顯示對話框
        dialog.transient(self)
        dialog.update_idletasks()
        
        # 計算居中位置
        x = (dialog.winfo_screenwidth() // 2) - (600 // 2)
        y = (dialog.winfo_screenheight() // 2) - (400 // 2)
        dialog.geometry(f"600x400+{x}+{y}")
        
        # 聚焦到對話框
        dialog.focus_set()
        
        # 等待對話框關閉
        dialog.wait_window()

    def _thread(self, target, *args):
        t = threading.Thread(target=target, args=args, daemon=True)
        t.start()

    def _on_mount_dir_changed(self, *args):
        """當 WIM 掛載路徑變更時自動同步到 Driver 分頁並檢查狀態"""
        if hasattr(self, 'var_driver_mount_dir') and hasattr(self, 'var_mount_dir'):
            wim_path = self.var_mount_dir.get().strip()
            current_driver_path = self.var_driver_mount_dir.get().strip()
            
            # 只有當 driver 路徑為空或與 wim 路徑不同時才同步
            if wim_path and (not current_driver_path or current_driver_path != wim_path):
                self.var_driver_mount_dir.set(wim_path)
                self._log(f"自動同步掛載路徑到 Driver 分頁: {wim_path}")
        
        # 當路徑變更時，延遲檢查掛載狀態（避免頻繁檢查）
        if hasattr(self, '_mount_check_timer1'):
            self.after_cancel(self._mount_check_timer1)
        self._mount_check_timer1 = self.after(500, self._on_check_wim1_status)

    def _on_mount_dir2_changed(self, *args):
        """當 WIM#2 掛載路徑變更時檢查狀態"""
        # 當路徑變更時，延遲檢查掛載狀態（避免頻繁檢查）
        if hasattr(self, '_mount_check_timer2'):
            self.after_cancel(self._mount_check_timer2)
        self._mount_check_timer2 = self.after(500, self._on_check_wim2_status)

    def _elevate_and_exit(self):
        """自動提升權限並退出當前程序（靜默執行）"""
        import sys
        import ctypes
        try:
            print("檢測到非管理員權限，正在提升權限...")
            script = os.path.abspath(sys.argv[0])
            params = " ".join([f'"{p}"' if ' ' in p else p for p in sys.argv[1:]])
            
            # 使用 SW_HIDE (0) 參數來隱藏視窗，實現靜默執行
            r = ctypes.windll.shell32.ShellExecuteW(
                None,           # hwnd
                "runas",        # lpOperation (以管理員身分執行)
                sys.executable, # lpFile (python.exe)
                f'"{script}" {params}',  # lpParameters
                None,           # lpDirectory
                0               # nShowCmd (0 = SW_HIDE, 隱藏視窗)
            )
            
            if r <= 32:
                print(f"提升權限失敗，錯誤代碼：{r}")
                messagebox.showerror("權限錯誤", "無法提升管理員權限，程式將退出")
            else:
                print("正在以管理員權限靜默啟動...")
            sys.exit(0)
        except Exception as e:
            print(f"提升權限時發生錯誤：{e}")
            messagebox.showerror("錯誤", f"提升權限失敗：{e}")
            sys.exit(1)

    def _on_create_mount_dir(self):
        """建立掛載資料夾"""
        path = self.var_mount_dir.get().strip()
        if not path:
            messagebox.showwarning("輸入不完整", "請先輸入掛載資料夾路徑")
            return
        
        try:
            if os.path.exists(path):
                if os.path.isdir(path):
                    if os.listdir(path):
                        self._log(f"資料夾已存在但非空：{path}")
                        messagebox.showinfo("資料夾狀態", "資料夾已存在但包含檔案。DISM 需要空的掛載資料夾。")
                    else:
                        self._log(f"資料夾已存在且為空：{path}")
                        messagebox.showinfo("資料夾狀態", "資料夾已存在且為空，可以使用。")
                else:
                    self._log(f"路徑已存在但不是資料夾：{path}")
                    messagebox.showerror("路徑錯誤", "指定路徑已存在但不是資料夾")
            else:
                os.makedirs(path, exist_ok=True)
                self._log(f"成功建立掛載資料夾：{path}")
                messagebox.showinfo("建立成功", f"已建立掛載資料夾：{path}")
                self._save_config()
        except Exception as e:
            self._log(f"建立資料夾失敗：{e}")
            messagebox.showerror("建立失敗", f"無法建立資料夾：{e}")

    # ---------- WIM 事件 ----------
    # WIM Index 防呆檢查
    def _on_wim1_index_changed(self, event=None):
        """WIM1 Index 變更時的防呆檢查"""
        selected_index = self.var_wim_index.get()
        wim2_index = self.var_wim_index2.get() if hasattr(self, 'var_wim_index2') else None
        
        # 檢查是否與 WIM2 的選擇衝突
        if selected_index and selected_index == wim2_index:
            self._log(f"⚠️  Index {selected_index} 已被 WIM#2 使用，請選擇其他 Index")
            # 清空當前選擇
            self.var_wim_index.set('')
            messagebox.showwarning("Index 衝突", f"Index {selected_index} 已被 WIM#2 使用\n請選擇不同的 Index")
            return
        
        if selected_index:
            self._log(f"✓ WIM#1 選擇 Index: {selected_index}")
            # 更新 WIM2 的可用選項
            self._update_wim2_available_indices()
        
        self._save_config()

    def _on_wim2_index_changed(self, event=None):
        """WIM2 Index 變更時的防呆檢查"""
        selected_index = self.var_wim_index2.get()
        wim1_index = self.var_wim_index.get() if hasattr(self, 'var_wim_index') else None
        
        # 檢查是否與 WIM1 的選擇衝突
        if selected_index and selected_index == wim1_index:
            self._log(f"⚠️  Index {selected_index} 已被 WIM#1 使用，請選擇其他 Index")
            # 清空當前選擇
            self.var_wim_index2.set('')
            messagebox.showwarning("Index 衝突", f"Index {selected_index} 已被 WIM#1 使用\n請選擇不同的 Index")
            return
        
        if selected_index:
            self._log(f"✓ WIM#2 選擇 Index: {selected_index}")
            # 更新 WIM1 的可用選項
            self._update_wim1_available_indices()
        
        self._save_config()

    def _update_wim1_available_indices(self):
        """更新 WIM1 的可用 Index 列表"""
        if not hasattr(self, 'wim1_available_indices'):
            return
        
        used_by_wim2 = self.var_wim_index2.get() if hasattr(self, 'var_wim_index2') else None
        available_indices = [idx for idx in self.wim1_available_indices if idx != used_by_wim2]
        
        self.cbo_wim_index['values'] = available_indices
        
        # 檢查當前選擇是否還有效
        current = self.var_wim_index.get()
        if current and current not in available_indices:
            self.var_wim_index.set('')

    def _update_wim2_available_indices(self):
        """更新 WIM2 的可用 Index 列表"""
        if not hasattr(self, 'wim2_available_indices'):
            return
        
        used_by_wim1 = self.var_wim_index.get() if hasattr(self, 'var_wim_index') else None
        available_indices = [idx for idx in self.wim2_available_indices if idx != used_by_wim1]
        
        self.cbo_wim_index2['values'] = available_indices
        
        # 檢查當前選擇是否還有效
        current = self.var_wim_index2.get()
        if current and current not in available_indices:
            self.var_wim_index2.set('')

    def _on_browse_wim(self):
        path = filedialog.askopenfilename(
            title="選擇 WIM 檔案",
            filetypes=[("WIM files", "*.wim"), ("All files", "*.*")],
        )
        if path:
            self.var_wim.set(path)
            self._log(f"已選擇 WIM 檔案：{path}")
            self._save_config()
            # 自動讀取映像資訊
            self._thread(self._do_wim_info, path)

    def _on_browse_mount_dir(self):
        path = filedialog.askdirectory(title="選擇掛載資料夾 (需為空)")
        if path:
            self.var_mount_dir.set(path)
            self._log(f"已選擇掛載資料夾：{path}")
            self._save_config()

    def _on_open_mount_dir(self):
        path = self.var_mount_dir.get().strip()
        if not path or not os.path.exists(path):
            self._log("掛載資料夾不存在或路徑無效")
            return
        try:
            os.startfile(path)
            self._log(f"已開啟掛載資料夾：{path}")
        except Exception as e:
            self._log(f"開啟掛載資料夾失敗：{e}")

    def _on_wim_info(self):
        wim = self.var_wim.get().strip()
        if not wim:
            messagebox.showwarning("輸入不完整", "請先選擇 WIM 檔案")
            return
        self._log("開始讀取 WIM 映像資訊...")
        self._save_config()
        self._thread(self._do_wim_info, wim)

    def _do_wim_info(self, wim: str):
        self._log(f"正在解析 WIM 檔案：{wim}")
        ok, images, err = WIMManager.get_wim_images(wim)
        if not ok:
            self._log(f"WIM 解析失敗：{err}")
            return
        if not images:
            self._log("此 WIM 檔案中未找到任何映像")
            return
        
        self._log(f"成功解析 WIM，找到 {len(images)} 個映像")
        # 更新下拉
        idxes = [str(img["Index"]) + (f" - {img['Name']}" if img.get("Name") else "") for img in images]
        indices_only = [str(img["Index"]) for img in images]
        
        # 儲存 WIM1 的所有可用 Index
        self.wim1_available_indices = indices_only.copy()
        
        def update_combo():
            # 檢查 WIM2 是否已選擇 Index，排除已被使用的
            used_by_wim2 = self.var_wim_index2.get() if hasattr(self, 'var_wim_index2') else None
            available_indices = [idx for idx in indices_only if idx != used_by_wim2]
            
            self.cbo_wim_index['values'] = available_indices
            
            # 若目前選擇的 Index 已被 WIM2 使用，需要重新選擇
            current_selection = self.var_wim_index.get()
            if current_selection and current_selection == used_by_wim2:
                self.var_wim_index.set('')
                self._log(f"⚠️  Index {current_selection} 已被 WIM#2 使用，請重新選擇")
            
            # 若尚未選擇且有可用選項，預設第一個可用的
            if not self.var_wim_index.get() and available_indices:
                self.var_wim_index.set(available_indices[0])
                self._save_config()
                self._log(f"✓ 自動選擇第一個可用映像 Index：{available_indices[0]}")
            elif not available_indices:
                self._log("⚠️  所有 Index 都已被使用，請檢查 WIM#2 的選擇")
        
        self.after(0, update_combo)
        
        for i, img in enumerate(images):
            name = img.get('Name', '(無名稱)')
            desc = img.get('Description', '(無描述)')
            self._log(f"映像 {img['Index']}: {name} - {desc}")
        self._log("映像資訊讀取完成")

    def _on_wim_mount(self):
        wim = self.var_wim.get().strip()
        idx = self.var_wim_index.get().strip()
        mdir = self.var_mount_dir.get().strip()
        ro = self.var_wim_readonly.get()
        
        self._log("開始 WIM#1 掛載前檢查...")
        
        if not wim or not mdir:
            self._log("掛載檢查失敗：缺少 WIM 檔案或掛載資料夾")
            messagebox.showwarning("輸入不完整", "請選擇 WIM 與掛載資料夾")
            return
        
        # Index 衝突檢查
        if idx and hasattr(self, 'var_wim_index2'):
            wim2_index = self.var_wim_index2.get()
            if idx == wim2_index:
                self._log(f"❌ Index 衝突：WIM#1 和 WIM#2 都選擇了 Index {idx}")
                messagebox.showerror("Index 衝突", f"WIM#1 和 WIM#2 不能使用相同的 Index: {idx}\n請選擇不同的 Index")
                return
            
        # 若未選 Index，嘗試自動解析
        if not idx:
            self._log("未選擇 Index，嘗試自動解析...")
            ok, images, err = WIMManager.get_wim_images(wim)
            if not ok or not images:
                self._log(f"自動解析失敗：{err}")
                messagebox.showwarning("缺少 Index", "請按『讀取映像資訊』後選擇 Index")
                return
            if len(images) == 1:
                idx = str(images[0]['Index'])
                self.var_wim_index.set(idx)
                self._save_config()
                self._log(f"自動選擇唯一映像 Index：{idx}")
            else:
                self._log(f"WIM 包含 {len(images)} 個映像，需要手動選擇")
                messagebox.showwarning("需要選擇 Index", "此 WIM 有多個映像，請先選擇 Index")
                return
                
        if not os.path.exists(mdir):
            self._log(f"掛載資料夾不存在：{mdir}")
            messagebox.showwarning("路徑不存在", "掛載資料夾不存在，請先建立")
            return
            
        if os.listdir(mdir):
            self._log(f"掛載資料夾非空：{mdir}")
            messagebox.showwarning("資料夾非空", "DISM 需要空的掛載資料夾，請清空後再試")
            return
            
        try:
            index = int(idx)
        except ValueError:
            self._log(f"Index 格式錯誤：{idx}")
            messagebox.showwarning("Index 錯誤", "Index 必須是數字")
            return
            
        self._log("掛載前檢查通過，開始掛載...")
        self._save_config()
        self._thread(self._do_wim_mount, wim, index, mdir, ro)

    def _do_wim_mount(self, wim: str, index: int, mdir: str, ro: bool):
        readonly_text = "唯讀" if ro else "讀寫"
        self._log(f"正在掛載 WIM...")
        self._log(f"  WIM 檔案: {wim}")
        self._log(f"  映像 Index: {index}")
        self._log(f"  掛載位置: {mdir}")
        self._log(f"  掛載模式: {readonly_text}")
        
        # 先檢查是否已有掛載
        self._log("檢查現有掛載狀態...")
        check_ok, mounted_images, check_err = WIMManager.get_mount_info()
        if check_ok and mounted_images:
            # 檢查是否有衝突的掛載
            conflict_found = False
            for img in mounted_images:
                img_file = img.get('ImageFile', '').lower()
                img_index = img.get('ImageIndex', '')
                img_mount_dir = img.get('MountDir', '')
                
                # 檢查是否相同的 WIM 文件和 Index
                if (os.path.normpath(wim).lower() in img_file or img_file in os.path.normpath(wim).lower()) and str(index) == img_index:
                    conflict_found = True
                    self._log(f"⚠️ 發現衝突: WIM {wim} Index {index} 已掛載到 {img_mount_dir}")
                    
                    def ask_user():
                        response = messagebox.askyesnocancel(
                            "掛載衝突",
                            f"映像 {os.path.basename(wim)} Index {index} 已經掛載到:\n{img_mount_dir}\n\n"
                            f"請選擇處理方式:\n"
                            f"是(Y) = 強制清理後重新掛載\n"
                            f"否(N) = 取消掛載操作\n"
                            f"取消 = 查看所有掛載狀態"
                        )
                        
                        if response is True:  # 是 - 強制清理
                            self._log("使用者選擇強制清理後重新掛載...")
                            cleanup_ok, cleanup_msg = WIMManager.cleanup_mount()
                            if cleanup_ok:
                                self._log(f"✓ {cleanup_msg}")
                                # 清理後重新嘗試掛載
                                self._perform_mount(wim, index, mdir, ro)
                            else:
                                self._log(f"✗ 清理失敗: {cleanup_msg}")
                                messagebox.showerror("清理失敗", f"無法清理掛載狀態:\n{cleanup_msg}")
                                # 為清理錯誤提供詳細建議
                                self.after(100, lambda: self.show_error_with_advice("清理失敗", cleanup_msg))
                        elif response is False:  # 否 - 取消
                            self._log("使用者選擇取消掛載操作")
                            return
                        else:  # 取消 - 查看狀態
                            self._log("顯示所有掛載狀態...")
                            self._do_check_wim_mount_status()
                            return
                    
                    self.after(0, ask_user)
                    return
            
        # 沒有衝突，直接掛載
        self._perform_mount(wim, index, mdir, ro)
    
    def _perform_mount(self, wim: str, index: int, mdir: str, ro: bool):
        """實際執行掛載操作"""
        ok, msg = WIMManager.mount_wim(wim, index, mdir, ro)
        if ok:
            self._log("✓ WIM 掛載成功！")
            self._log(f"掛載位置: {mdir}")
            
            # 自動同步掛載路徑到 Driver 分頁
            if hasattr(self, 'var_driver_mount_dir'):
                self.var_driver_mount_dir.set(mdir)
                self._log(f"✓ 已自動同步掛載路徑到 Driver 分頁: {mdir}")
            
            messagebox.showinfo("掛載成功", f"WIM 已成功掛載到:\n{mdir}\n\n已自動同步路徑到 Driver 分頁")
        else:
            self._log(f"✗ WIM 掛載失敗: {msg}")
            
            # 檢查是否是常見的掛載錯誤
            if "0xc1420127" in msg or "already mounted" in msg.lower():
                def handle_mount_error():
                    response = messagebox.askyesno(
                        "掛載失敗 - 映像已掛載",
                        f"錯誤: 映像已經掛載\n{msg}\n\n是否要清理掛載狀態後重試？"
                    )
                    if response:
                        self._log("嘗試清理掛載狀態後重試...")
                        cleanup_ok, cleanup_msg = WIMManager.cleanup_mount()
                        if cleanup_ok:
                            self._log(f"✓ 清理成功: {cleanup_msg}")
                            self._log("重新嘗試掛載...")
                            self._perform_mount(wim, index, mdir, ro)
                        else:
                            self._log(f"✗ 清理失敗: {cleanup_msg}")
                            messagebox.showerror("清理失敗", f"無法清理掛載狀態:\n{cleanup_msg}")
                            # 為清理錯誤提供詳細建議
                            self.after(100, lambda: self.show_error_with_advice("清理失敗", cleanup_msg))
                
                self.after(0, handle_mount_error)
            else:
                messagebox.showerror("掛載失敗", f"掛載失敗:\n{msg}")
                # 為掛載錯誤提供詳細建議
                self.after(100, lambda: self.show_error_with_advice("掛載失敗", msg))
        
        # 更新掛載狀態
        self.after(500, self._on_check_wim1_status)

    def _on_wim_unmount(self):
        mdir = self.var_mount_dir.get().strip()
        commit = self.var_unmount_commit.get()
        
        if not mdir:
            self._log("卸載失敗：未指定掛載資料夾")
            messagebox.showwarning("輸入不完整", "請先指定掛載資料夾")
            return
        
        # 檢查是否已掛載
        is_mounted, mount_info, err = WIMManager.is_path_mounted(mdir)
        if err:
            self._log(f"⚠ 無法確認掛載狀態: {err}")
        elif not is_mounted:
            result = messagebox.askyesno(
                "未檢測到掛載", 
                f"路徑 '{mdir}' 未在 DISM 掛載清單中。\n\n"
                "可能原因：\n"
                "• 映像尚未掛載\n"
                "• 掛載狀態異常\n\n"
                "是否仍要嘗試卸載？（可能需要使用「一鍵修復」功能）",
                icon='warning'
            )
            if not result:
                return
            
        commit_text = "提交變更" if commit else "丟棄變更"
        self._log(f"準備卸載 WIM (模式: {commit_text})...")
        self._thread(self._do_wim_unmount, mdir, commit)

    def _on_check_wim1_status(self):
        """檢查 WIM#1 的掛載狀態"""
        mdir = self.var_mount_dir.get().strip()
        if not mdir:
            self.var_mount_status1.set("未設定掛載路徑")
            self.lbl_mount_status1.configure(foreground="gray")
            self._update_wim1_buttons(False)
            return
        
        self._thread(self._do_check_wim1_status, mdir)

    def _do_check_wim1_status(self, mdir: str):
        """實際檢查 WIM#1 狀態"""
        status, details = WIMManager.get_mount_status_for_path(mdir)
        
        # 在主執行緒更新 UI
        self.after(0, lambda: self._update_wim1_status_ui(status, details))

    def _update_wim1_status_ui(self, status: str, details: str):
        """更新 WIM#1 狀態 UI"""
        status_map = {
            "mounted": ("✓ 已掛載", "green", True),
            "not_mounted": ("○ 未掛載", "gray", False),
            "needs_remount": ("⚠ 需要修復", "orange", True),
            "orphaned": ("⚠ 狀態異常", "orange", False),
            "error": ("✗ 檢查失敗", "red", False),
        }
        
        text, color, is_mounted = status_map.get(status, ("未知", "gray", False))
        self.var_mount_status1.set(f"{text} - {details}")
        self.lbl_mount_status1.configure(foreground=color)
        self._update_wim1_buttons(is_mounted)
        
        # 當掛載狀態改變時，同步更新驅動分頁的下拉選單
        self._refresh_all_driver_comboboxes()

    def _update_wim1_buttons(self, is_mounted: bool):
        """根據掛載狀態更新 WIM#1 按鈕狀態"""
        if is_mounted:
            self.btn_wim_mount1.configure(state="disabled")
            self.btn_wim_unmount1.configure(state="normal")
        else:
            self.btn_wim_mount1.configure(state="normal")
            self.btn_wim_unmount1.configure(state="disabled")

    def _on_check_wim2_status(self):
        """檢查 WIM#2 的掛載狀態"""
        mdir = self.var_mount_dir2.get().strip()
        if not mdir:
            self.var_mount_status2.set("未設定掛載路徑")
            self.lbl_mount_status2.configure(foreground="gray")
            self._update_wim2_buttons(False)
            return
        
        self._thread(self._do_check_wim2_status, mdir)

    def _do_check_wim2_status(self, mdir: str):
        """實際檢查 WIM#2 狀態"""
        status, details = WIMManager.get_mount_status_for_path(mdir)
        
        # 在主執行緒更新 UI
        self.after(0, lambda: self._update_wim2_status_ui(status, details))

    def _update_wim2_status_ui(self, status: str, details: str):
        """更新 WIM#2 狀態 UI"""
        status_map = {
            "mounted": ("✓ 已掛載", "green", True),
            "not_mounted": ("○ 未掛載", "gray", False),
            "needs_remount": ("⚠ 需要修復", "orange", True),
            "orphaned": ("⚠ 狀態異常", "orange", False),
            "error": ("✗ 檢查失敗", "red", False),
        }
        
        text, color, is_mounted = status_map.get(status, ("未知", "gray", False))
        self.var_mount_status2.set(f"{text} - {details}")
        self.lbl_mount_status2.configure(foreground=color)
        self._update_wim2_buttons(is_mounted)
        
        # 當掛載狀態改變時，同步更新驅動分頁的下拉選單
        self._refresh_all_driver_comboboxes()

    def _update_wim2_buttons(self, is_mounted: bool):
        """根據掛載狀態更新 WIM#2 按鈕狀態"""
        if is_mounted:
            self.btn_wim_mount2.configure(state="disabled")
            self.btn_wim_unmount2.configure(state="normal")
        else:
            self.btn_wim_mount2.configure(state="normal")
            self.btn_wim_unmount2.configure(state="disabled")

    def _on_close_explorer(self):
        """手動關閉指向掛載資料夾的檔案總管視窗"""
        mdir = self.var_mount_dir.get().strip()
        
        if not mdir:
            messagebox.showwarning("輸入不完整", "請先指定掛載資料夾")
            return
            
        self._log("手動關閉檔案總管視窗...")
        self._thread(self._do_close_explorer, mdir)

    def _do_close_explorer(self, mdir: str):
        """執行關閉檔案總管的操作"""
        try:
            self._log(f"正在關閉指向 {mdir} 的檔案總管視窗...")
            ok, msg = WIMManager.close_explorer_windows(mdir)
            if ok:
                self._log(f"✓ {msg}")
                messagebox.showinfo("完成", f"已處理檔案總管視窗\n{msg}")
            else:
                self._log(f"⚠ {msg}")
                messagebox.showwarning("注意", f"處理檔案總管視窗時遇到問題:\n{msg}")
        except Exception as e:
            self._log(f"關閉檔案總管視窗時發生錯誤: {e}")
            messagebox.showerror("錯誤", f"操作失敗: {e}")

    def _on_check_wim_mount_status(self):
        """檢查當前 WIM 掛載狀態"""
        self._log("檢查系統中所有已掛載的映像...")
        self._thread(self._do_check_wim_mount_status)

    def _do_check_wim_mount_status(self):
        """執行檢查 WIM 掛載狀態"""
        try:
            ok, mounted_images, err = WIMManager.get_mount_info()
            if not ok:
                self._log(f"✗ 檢查掛載狀態失敗: {err}")
                messagebox.showerror("檢查失敗", f"無法檢查掛載狀態:\n{err}")
                return
                
            if not mounted_images:
                self._log("✓ 系統中沒有已掛載的映像")
                messagebox.showinfo("掛載狀態", "系統中沒有已掛載的映像")
                return
                
            self._log(f"✓ 找到 {len(mounted_images)} 個已掛載的映像:")
            for i, img in enumerate(mounted_images, 1):
                mount_dir = img.get('MountDir', 'N/A')
                image_file = img.get('ImageFile', 'N/A')
                image_index = img.get('ImageIndex', 'N/A')
                status = img.get('Status', 'N/A')
                read_write = img.get('ReadWrite', 'N/A')
                
                self._log(f"  {i}. 掛載目錄: {mount_dir}")
                self._log(f"     映像檔案: {image_file}")
                self._log(f"     映像索引: {image_index}")
                self._log(f"     狀態: {status}")
                self._log(f"     權限: {read_write}")
                self._log("")
                
            messagebox.showinfo("掛載狀態", f"找到 {len(mounted_images)} 個已掛載的映像\n詳細資訊請查看日誌")
            
        except Exception as e:
            self._log(f"檢查掛載狀態時發生錯誤: {e}")
            messagebox.showerror("錯誤", f"檢查掛載狀態時發生錯誤: {e}")

    def _on_cleanup_mount(self):
        """清理掛載點"""
        response = messagebox.askyesno(
            "確認清理", 
            "此操作將清理所有掛載點\n⚠️ 這會強制卸載所有映像並捨棄未提交的變更\n\n確定要繼續嗎？"
        )
        if response:
            self._log("開始清理掛載點...")
            self._thread(self._do_cleanup_mount)

    def _do_cleanup_mount(self):
        """執行清理掛載點操作"""
        try:
            ok, msg = WIMManager.cleanup_mount()
            if ok:
                self._log(f"✓ 清理完成: {msg}")
                messagebox.showinfo("清理成功", f"掛載點清理完成:\n{msg}")
            else:
                self._log(f"✗ 清理失敗: {msg}")
                messagebox.showerror("清理失敗", f"掛載點清理失敗:\n{msg}")
                # 為掛載點清理錯誤提供詳細建議
                self.after(100, lambda: self.show_error_with_advice("清理失敗", msg))
        except Exception as e:
            self._log(f"清理掛載點時發生錯誤: {e}")
            messagebox.showerror("錯誤", f"清理掛載點時發生錯誤: {e}")

    def _on_fix_broken_mounts(self):
        """修復損壞的掛載點"""
        self._log("🔧 開始修復損壞的掛載點...")
        self._thread(self._do_fix_broken_mounts)

    def _do_fix_broken_mounts(self):
        """執行修復損壞掛載點操作"""
        try:
            ok, msg = WIMManager.fix_broken_mounts()
            
            # 將詳細訊息寫入日誌
            for line in msg.split('\n'):
                if line.strip():
                    self._log(line)
            
            if ok:
                self._log("✅ 損壞掛載點修復完成！")
                messagebox.showinfo("修復完成", 
                    "損壞掛載點修復完成！\n\n"
                    "所有 'Needs Remount' 狀態的掛載點已處理。\n"
                    "詳細資訊請查看日誌。")
            else:
                self._log("⚠️ 修復部分完成或失敗")
                messagebox.showwarning("修復部分完成", 
                    "修復操作已執行但可能未完全成功。\n\n"
                    "建議嘗試以下操作：\n"
                    "1. 使用「強力清理」功能\n"
                    "2. 重新啟動電腦\n\n"
                    "詳細結果請查看日誌視窗。")
                    
        except Exception as e:
            self._log(f"❌ 修復損壞掛載點時發生錯誤: {e}")
            messagebox.showerror("修復錯誤", f"修復損壞掛載點時發生錯誤:\n{e}")
            # 為修復錯誤提供詳細建議
            self.after(100, lambda: self.show_error_with_advice("修復錯誤", str(e)))

    def _on_smart_cleanup_fix(self):
        """一鍵智能修復所有 WIM 掛載問題"""
        self._log("🚀 開始一鍵智能修復...")
        self._thread(self._do_smart_cleanup_fix)

    def _do_smart_cleanup_fix(self):
        """執行一鍵智能修復操作"""
        try:
            ok, msg = WIMManager.smart_cleanup_and_fix()
            
            # 將詳細訊息寫入日誌
            for line in msg.split('\n'):
                if line.strip():
                    self._log(line)
            
            if ok:
                self._log("✅ 一鍵智能修復完成")
                messagebox.showinfo("修復完成", "🎉 一鍵智能修復已完成！\n\n所有 WIM 掛載問題已自動診斷和修復。\n系統現在處於良好狀態，可以正常進行新的掛載操作。")
            else:
                self._log("❌ 一鍵智能修復失敗")
                messagebox.showerror("修復失敗", f"一鍵智能修復過程中遇到問題:\n{msg}")
                
        except Exception as e:
            self._log(f"一鍵智能修復錯誤: {e}")
            messagebox.showerror("修復錯誤", f"一鍵智能修復時發生錯誤:\n{e}")
            # 為修復錯誤提供詳細建議
            self.after(100, lambda: self.show_error_with_advice("修復錯誤", str(e)))

    def _on_force_cleanup(self):
        """強力清理掛載點 - 最後手段"""
        response = messagebox.askyesno(
            "強力清理確認", 
            "⚠️ 強力清理將執行以下操作:\n"
            "• 強制終止相關進程\n"
            "• 清理系統暫存檔\n" 
            "• 重啟系統服務\n"
            "• 清理註冊表項目\n"
            "• 執行多重 DISM 清理\n\n"
            "這個操作比較激進，確定要繼續嗎？"
        )
        if response:
            self._log("🔥 開始強力清理掛載點...")
            # 顯示進度對話框
            progress_msg = "正在執行強力清理...\n這可能需要一些時間，請耐心等待..."
            self._show_progress_and_execute(self._do_force_cleanup, progress_msg)

    def _show_progress_and_execute(self, target_func, message):
        """顯示進度對話框並執行長時間任務"""
        import tkinter.messagebox as mb
        
        # 創建一個簡單的進度提示
        progress_window = tk.Toplevel(self)
        progress_window.title("執行中...")
        progress_window.geometry("400x150")
        progress_window.transient(self)
        progress_window.grab_set()
        
        # 居中顯示
        progress_window.update_idletasks()
        x = (progress_window.winfo_screenwidth() // 2) - (400 // 2)
        y = (progress_window.winfo_screenheight() // 2) - (150 // 2)
        progress_window.geometry(f"400x150+{x}+{y}")
        
        progress_label = tk.Label(progress_window, text=message, wraplength=350, justify=tk.CENTER)
        progress_label.pack(expand=True)
        
        # 執行清理任務
        def execute_task():
            try:
                target_func()
            finally:
                progress_window.destroy()
        
        # 延遲執行以確保進度窗口顯示
        self.after(100, execute_task)

    def _do_force_cleanup(self):
        """執行強力清理操作"""
        try:
            self._log("🚀 啟動終極清理程序...")
            ok, msg = WIMManager.ultimate_cleanup()
            
            # 將詳細訊息寫入日誌
            for line in msg.split('\n'):
                if line.strip():
                    self._log(line)
            
            if ok:
                self._log("✅ 強力清理完成！")
                messagebox.showinfo("強力清理成功", 
                    "強力清理已完成！\n\n"
                    "系統已重啟相關服務並清理掛載狀態。\n"
                    "詳細資訊請查看日誌。\n\n"
                    "建議現在重新嘗試掛載操作。")
            else:
                self._log("⚠️ 強力清理部分完成")
                response = messagebox.askyesnocancel(
                    "強力清理部分完成",
                    "強力清理已執行但可能未完全成功。\n\n"
                    "建議選項:\n"
                    "是 = 重新啟動電腦（最徹底）\n"
                    "否 = 嘗試重新掛載\n" 
                    "取消 = 查看詳細日誌\n\n"
                    "詳細結果請查看日誌視窗。"
                )
                
                if response is True:  # 重啟電腦
                    restart_confirm = messagebox.askyesno(
                        "重啟確認",
                        "確定要重新啟動電腦嗎？\n\n"
                        "重啟將徹底清除所有掛載狀態，\n"
                        "但會中斷當前所有工作。"
                    )
                    if restart_confirm:
                        self._log("🔄 使用者選擇重啟電腦...")
                        try:
                            import subprocess
                            subprocess.run(["shutdown", "/r", "/t", "10", "/c", "WIM工具：重啟清理掛載狀態"], check=True)
                            self._log("⏰ 系統將在 10 秒後重啟...")
                            messagebox.showinfo("重啟排程", "系統將在 10 秒後重啟\n請保存重要工作！")
                        except Exception as e:
                            self._log(f"❌ 重啟失敗: {e}")
                            messagebox.showerror("重啟失敗", f"無法重啟系統: {e}")
                
        except Exception as e:
            self._log(f"❌ 強力清理時發生嚴重錯誤: {e}")
            messagebox.showerror("強力清理錯誤", f"強力清理過程中發生嚴重錯誤:\n{e}")

    # ---------- WIM #2 事件 ----------
    def _on_browse_wim2(self):
        path = filedialog.askopenfilename(
            title="選擇第二個 WIM 檔案",
            filetypes=[("WIM files", "*.wim"), ("All files", "*.*")],
        )
        if path:
            self.var_wim2.set(path)
            self._log(f"已選擇第二個 WIM 檔案：{path}")
            self._save_config()
            # 自動讀取映像資訊
            self._thread(self._do_wim_info2, path)

    def _on_browse_mount_dir2(self):
        path = filedialog.askdirectory(title="選擇第二個掛載資料夾 (需為空)")
        if path:
            self.var_mount_dir2.set(path)
            self._log(f"已選擇第二個掛載資料夾：{path}")
            self._save_config()

    def _on_create_mount_dir2(self):
        """建立第二個掛載資料夾"""
        path = self.var_mount_dir2.get().strip()
        if not path:
            messagebox.showwarning("輸入不完整", "請先輸入第二個掛載資料夾路徑")
            return
        
        try:
            if os.path.exists(path):
                if os.path.isdir(path):
                    if os.listdir(path):
                        self._log(f"資料夾已存在但非空：{path}")
                        messagebox.showinfo("資料夾狀態", "資料夾已存在但包含檔案。DISM 需要空的掛載資料夾。")
                    else:
                        self._log(f"資料夾已存在且為空：{path}")
                        messagebox.showinfo("資料夾狀態", "資料夾已存在且為空，可以使用。")
                else:
                    self._log(f"路徑已存在但不是資料夾：{path}")
                    messagebox.showerror("路徑錯誤", "指定路徑已存在但不是資料夾")
            else:
                os.makedirs(path, exist_ok=True)
                self._log(f"成功建立第二個掛載資料夾：{path}")
                messagebox.showinfo("建立成功", f"已建立第二個掛載資料夾：{path}")
                self._save_config()
        except Exception as e:
            self._log(f"建立資料夾失敗：{e}")
            messagebox.showerror("建立失敗", f"無法建立資料夾：{e}")

    def _on_open_mount_dir2(self):
        path = self.var_mount_dir2.get().strip()
        if not path or not os.path.exists(path):
            self._log("第二個掛載資料夾不存在或路徑無效")
            return
        try:
            os.startfile(path)
            self._log(f"已開啟第二個掛載資料夾：{path}")
        except Exception as e:
            self._log(f"開啟第二個掛載資料夾失敗：{e}")

    def _on_wim_info2(self):
        wim = self.var_wim2.get().strip()
        if not wim:
            messagebox.showwarning("輸入不完整", "請先選擇第二個 WIM 檔案")
            return
        self._log("開始讀取第二個 WIM 映像資訊...")
        self._save_config()
        self._thread(self._do_wim_info2, wim)

    def _do_wim_info2(self, wim: str):
        self._log(f"正在解析第二個 WIM 檔案：{wim}")
        ok, images, err = WIMManager.get_wim_images(wim)
        if not ok:
            self._log(f"第二個 WIM 解析失敗：{err}")
            return
        if not images:
            self._log("此 WIM 檔案中未找到任何映像")
            return
        
        self._log(f"成功解析第二個 WIM，找到 {len(images)} 個映像")
        indices_only = [str(img["Index"]) for img in images]
        
        # 儲存 WIM2 的所有可用 Index
        self.wim2_available_indices = indices_only.copy()
        
        def update_combo():
            # 檢查 WIM1 是否已選擇 Index，排除已被使用的
            used_by_wim1 = self.var_wim_index.get() if hasattr(self, 'var_wim_index') else None
            available_indices = [idx for idx in indices_only if idx != used_by_wim1]
            
            self.cbo_wim_index2['values'] = available_indices
            
            # 若目前選擇的 Index 已被 WIM1 使用，需要重新選擇
            current_selection = self.var_wim_index2.get()
            if current_selection and current_selection == used_by_wim1:
                self.var_wim_index2.set('')
                self._log(f"⚠️  Index {current_selection} 已被 WIM#1 使用，請重新選擇")
            
            # 若尚未選擇且有可用選項，預設第一個可用的
            if not self.var_wim_index2.get() and available_indices:
                self.var_wim_index2.set(available_indices[0])
                self._save_config()
                self._log(f"✓ 自動選擇第一個可用映像 Index：{available_indices[0]}")
            elif not available_indices:
                self._log("⚠️  所有 Index 都已被使用，請檢查 WIM#1 的選擇")
        
        self.after(0, update_combo)
        
        for i, img in enumerate(images):
            name = img.get('Name', '(無名稱)')
            desc = img.get('Description', '(無描述)')
            self._log(f"第二個映像 {img['Index']}: {name} - {desc}")
        self._log("第二個映像資訊讀取完成")

    def _on_wim_mount2(self):
        wim = self.var_wim2.get().strip()
        idx = self.var_wim_index2.get().strip()
        mdir = self.var_mount_dir2.get().strip()
        ro = self.var_wim_readonly2.get()
        
        self._log("開始 WIM#2 掛載前檢查...")
        
        if not wim or not mdir:
            self._log("第二個掛載檢查失敗：缺少 WIM 檔案或掛載資料夾")
            messagebox.showwarning("輸入不完整", "請選擇第二個 WIM 與掛載資料夾")
            return
        
        # Index 衝突檢查
        if idx and hasattr(self, 'var_wim_index'):
            wim1_index = self.var_wim_index.get()
            if idx == wim1_index:
                self._log(f"❌ Index 衝突：WIM#1 和 WIM#2 都選擇了 Index {idx}")
                messagebox.showerror("Index 衝突", f"WIM#1 和 WIM#2 不能使用相同的 Index: {idx}\n請選擇不同的 Index")
                return
            
        # 若未選 Index，嘗試自動解析
        if not idx:
            self._log("未選擇第二個 Index，嘗試自動解析...")
            ok, images, err = WIMManager.get_wim_images(wim)
            if not ok or not images:
                self._log(f"自動解析失敗：{err}")
                messagebox.showwarning("缺少 Index", "請按『讀取映像資訊』後選擇 Index")
                return
            if len(images) == 1:
                idx = str(images[0]['Index'])
                self.var_wim_index2.set(idx)
                self._save_config()
                self._log(f"自動選擇唯一映像 Index：{idx}")
            else:
                self._log(f"第二個 WIM 包含 {len(images)} 個映像，需要手動選擇")
                messagebox.showwarning("需要選擇 Index", "此 WIM 有多個映像，請先選擇 Index")
                return
                
        if not os.path.exists(mdir):
            self._log(f"第二個掛載資料夾不存在：{mdir}")
            messagebox.showwarning("路徑不存在", "第二個掛載資料夾不存在，請先建立")
            return
            
        if os.listdir(mdir):
            self._log(f"第二個掛載資料夾非空：{mdir}")
            messagebox.showwarning("資料夾非空", "DISM 需要空的掛載資料夾，請清空後再試")
            return
            
        try:
            index = int(idx)
        except ValueError:
            self._log(f"第二個 Index 格式錯誤：{idx}")
            messagebox.showwarning("Index 錯誤", "Index 必須是數字")
            return
            
        self._log("第二個掛載前檢查通過，開始掛載...")
        self._save_config()
        self._thread(self._do_wim_mount2, wim, index, mdir, ro)

    def _do_wim_mount2(self, wim: str, index: int, mdir: str, ro: bool):
        readonly_text = "唯讀" if ro else "讀寫"
        self._log(f"正在掛載第二個 WIM...")
        self._log(f"  WIM 檔案: {wim}")
        self._log(f"  映像 Index: {index}")
        self._log(f"  掛載位置: {mdir}")
        self._log(f"  掛載模式: {readonly_text}")
        
        # 先檢查是否已有掛載
        self._log("檢查現有掛載狀態...")
        check_ok, mounted_images, check_err = WIMManager.get_mount_info()
        if check_ok and mounted_images:
            # 檢查是否有衝突的掛載
            conflict_found = False
            for img in mounted_images:
                img_file = img.get('ImageFile', '').lower()
                img_index = img.get('ImageIndex', '')
                img_mount_dir = img.get('MountDir', '')
                
                # 檢查是否相同的 WIM 文件和 Index
                if (os.path.normpath(wim).lower() in img_file or img_file in os.path.normpath(wim).lower()) and str(index) == img_index:
                    conflict_found = True
                    self._log(f"⚠️ 發現衝突: WIM {wim} Index {index} 已掛載到 {img_mount_dir}")
                    
                    def ask_user():
                        response = messagebox.askyesnocancel(
                            "掛載衝突 - WIM #2",
                            f"映像 {os.path.basename(wim)} Index {index} 已經掛載到:\n{img_mount_dir}\n\n"
                            f"請選擇處理方式:\n"
                            f"是(Y) = 強制清理後重新掛載\n"
                            f"否(N) = 取消掛載操作\n"
                            f"取消 = 查看所有掛載狀態"
                        )
                        
                        if response is True:  # 是 - 強制清理
                            self._log("使用者選擇強制清理後重新掛載...")
                            cleanup_ok, cleanup_msg = WIMManager.cleanup_mount()
                            if cleanup_ok:
                                self._log(f"✓ {cleanup_msg}")
                                # 清理後重新嘗試掛載
                                self._perform_mount2(wim, index, mdir, ro)
                            else:
                                self._log(f"✗ 清理失敗: {cleanup_msg}")
                                messagebox.showerror("清理失敗", f"無法清理掛載狀態:\n{cleanup_msg}")
                        elif response is False:  # 否 - 取消
                            self._log("使用者選擇取消第二個掛載操作")
                            return
                        else:  # 取消 - 查看狀態
                            self._log("顯示所有掛載狀態...")
                            self._do_check_wim_mount_status()
                            return
                    
                    self.after(0, ask_user)
                    return
            
        # 沒有衝突，直接掛載
        self._perform_mount2(wim, index, mdir, ro)
    
    def _perform_mount2(self, wim: str, index: int, mdir: str, ro: bool):
        """實際執行第二個 WIM 掛載操作"""
        ok, msg = WIMManager.mount_wim(wim, index, mdir, ro)
        if ok:
            self._log("✓ 第二個 WIM 掛載成功！")
            self._log(f"掛載位置: {mdir}")
            messagebox.showinfo("掛載成功", f"第二個 WIM 已成功掛載到:\n{mdir}")
            # 更新掛載狀態
            self.after(500, self._on_check_wim2_status)
        else:
            self._log(f"✗ 第二個 WIM 掛載失敗: {msg}")
            
            # 檢查是否是常見的掛載錯誤
            if "0xc1420127" in msg or "already mounted" in msg.lower():
                def handle_mount_error():
                    response = messagebox.askyesno(
                        "掛載失敗 - 映像已掛載 (WIM #2)",
                        f"錯誤: 第二個映像已經掛載\n{msg}\n\n是否要清理掛載狀態後重試？"
                    )
                    if response:
                        self._log("嘗試清理掛載狀態後重試...")
                        cleanup_ok, cleanup_msg = WIMManager.cleanup_mount()
                        if cleanup_ok:
                            self._log(f"✓ 清理成功: {cleanup_msg}")
                            self._log("重新嘗試掛載第二個 WIM...")
                            self._perform_mount2(wim, index, mdir, ro)
                        else:
                            self._log(f"✗ 清理失敗: {cleanup_msg}")
                            messagebox.showerror("清理失敗", f"無法清理掛載狀態:\n{cleanup_msg}")
                
                self.after(0, handle_mount_error)
            else:
                messagebox.showerror("掛載失敗", f"第二個掛載失敗:\n{msg}")

    def _on_wim_unmount2(self):
        mdir = self.var_mount_dir2.get().strip()
        commit = self.var_unmount_commit2.get()
        
        if not mdir:
            self._log("第二個卸載失敗：未指定掛載資料夾")
            messagebox.showwarning("輸入不完整", "請先指定第二個掛載資料夾")
            return
        
        # 檢查是否已掛載
        is_mounted, mount_info, err = WIMManager.is_path_mounted(mdir)
        if err:
            self._log(f"⚠ 無法確認掛載狀態: {err}")
        elif not is_mounted:
            result = messagebox.askyesno(
                "未檢測到掛載", 
                f"路徑 '{mdir}' 未在 DISM 掛載清單中。\n\n"
                "可能原因：\n"
                "• 映像尚未掛載\n"
                "• 掛載狀態異常\n\n"
                "是否仍要嘗試卸載？（可能需要使用「一鍵修復」功能）",
                icon='warning'
            )
            if not result:
                return
            
        commit_text = "提交變更" if commit else "丟棄變更"
        self._log(f"準備卸載第二個 WIM (模式: {commit_text})...")
        self._thread(self._do_wim_unmount2, mdir, commit)

    def _do_wim_unmount2(self, mdir: str, commit: bool):
        commit_text = "提交變更 (/Commit)" if commit else "丟棄變更 (/Discard)"
        self._log(f"正在卸載第二個 WIM...")
        self._log(f"  掛載位置: {mdir}")
        self._log(f"  卸載模式: {commit_text}")
        
        # 防呆：嘗試關閉指向掛載資料夾的檔案總管視窗
        self._log("正在檢查並關閉相關檔案總管視窗...")
        try:
            close_ok, close_msg = WIMManager.close_explorer_windows(mdir)
            if close_ok:
                self._log(f"✓ {close_msg}")
            else:
                self._log(f"⚠ 關閉檔案總管視窗時出現問題: {close_msg}")
                self._log("  繼續執行卸載程序...")
        except Exception as e:
            self._log(f"⚠ 關閉檔案總管視窗時發生錯誤: {e}")
            self._log("  繼續執行卸載程序...")
        
        # 短暫等待以確保檔案總管完全關閉
        import time
        time.sleep(1)
        
        ok, msg = WIMManager.unmount_wim(mdir, commit)
        if ok:
            self._log("✓ 第二個 WIM 卸載成功！")
            messagebox.showinfo("卸載成功", f"第二個 WIM 已成功卸載\n模式: {commit_text}")
            # 更新掛載狀態
            self.after(500, self._on_check_wim2_status)
        else:
            self._log(f"✗ 第二個 WIM 卸載失敗: {msg}")
            if "is currently in use" in msg or "正在使用" in msg or "檔案正在使用中" in msg:
                response = messagebox.askyesno(
                    "卸載失敗", 
                    f"第二個卸載失敗，可能有程式正在使用掛載資料夾:\n{msg}\n\n是否要強制重試？"
                )
                if response:
                    self._log("使用者選擇強制重試第二個...")
                    self._force_unmount_retry2(mdir, commit)
            else:
                messagebox.showerror("卸載失敗", f"第二個卸載失敗:\n{msg}")

    def _force_unmount_retry2(self, mdir: str, commit: bool):
        """強制重試卸載第二個 WIM"""
        self._log("正在執行第二個強制卸載重試...")
        
        try:
            self._log("嘗試關閉可能鎖定檔案的程式...")
            
            result = subprocess.run(['taskkill', '/F', '/IM', 'explorer.exe'], 
                                  capture_output=True, text=True)
            if result.returncode == 0:
                self._log("已終止 explorer.exe 程序")
                subprocess.Popen(['explorer.exe'])
                self._log("已重新啟動 explorer.exe")
            
            import time
            time.sleep(2)
            
            self._log("重新嘗試卸載第二個...")
            ok, msg = WIMManager.unmount_wim(mdir, commit)
            
            if ok:
                self._log("✓ 第二個強制卸載成功！")
                messagebox.showinfo("卸載成功", "第二個強制卸載成功！")
            else:
                self._log(f"✗ 第二個強制卸載仍然失敗: {msg}")
                messagebox.showerror("卸載失敗", f"第二個強制卸載仍然失敗:\n{msg}")
                
        except Exception as e:
            self._log(f"第二個強制卸載過程中發生錯誤: {e}")
            messagebox.showerror("錯誤", f"第二個強制卸載過程中發生錯誤: {e}")

    def _on_close_explorer2(self):
        """手動關閉指向第二個掛載資料夾的檔案總管視窗"""
        mdir = self.var_mount_dir2.get().strip()
        
        if not mdir:
            messagebox.showwarning("輸入不完整", "請先指定第二個掛載資料夾")
            return
            
        self._log("手動關閉第二個檔案總管視窗...")
        self._thread(self._do_close_explorer2, mdir)

    def _do_close_explorer2(self, mdir: str):
        """執行關閉第二個檔案總管的操作"""
        try:
            self._log(f"正在關閉指向 {mdir} 的檔案總管視窗...")
            ok, msg = WIMManager.close_explorer_windows(mdir)
            if ok:
                self._log(f"✓ {msg}")
                messagebox.showinfo("完成", f"已處理第二個檔案總管視窗\n{msg}")
            else:
                self._log(f"⚠ {msg}")
                messagebox.showwarning("注意", f"處理第二個檔案總管視窗時遇到問題:\n{msg}")
        except Exception as e:
            self._log(f"關閉第二個檔案總管視窗時發生錯誤: {e}")
            messagebox.showerror("錯誤", f"操作失敗: {e}")

    def _do_wim_unmount(self, mdir: str, commit: bool):
        commit_text = "提交變更 (/Commit)" if commit else "丟棄變更 (/Discard)"
        self._log(f"正在卸載 WIM...")
        self._log(f"  掛載位置: {mdir}")
        self._log(f"  卸載模式: {commit_text}")
        
        # 防呆：嘗試關閉指向掛載資料夾的檔案總管視窗
        self._log("正在檢查並關閉相關檔案總管視窗...")
        try:
            close_ok, close_msg = WIMManager.close_explorer_windows(mdir)
            if close_ok:
                self._log(f"✓ {close_msg}")
            else:
                self._log(f"⚠ 關閉檔案總管視窗時出現問題: {close_msg}")
                self._log("  繼續執行卸載程序...")
        except Exception as e:
            self._log(f"⚠ 關閉檔案總管視窗時發生錯誤: {e}")
            self._log("  繼續執行卸載程序...")
        
        # 短暫等待以確保檔案總管完全關閉
        import time
        time.sleep(1)
        
        ok, msg = WIMManager.unmount_wim(mdir, commit)
        if ok:
            self._log("✓ WIM 卸載成功！")
            messagebox.showinfo("卸載成功", f"WIM 已成功卸載\n模式: {commit_text}")
            # 更新掛載狀態
            self.after(500, self._on_check_wim1_status)
        else:
            self._log(f"✗ WIM 卸載失敗: {msg}")
            if "is currently in use" in msg or "正在使用" in msg or "檔案正在使用中" in msg:
                response = messagebox.askyesno(
                    "卸載失敗", 
                    f"卸載失敗，可能有程式正在使用掛載資料夾:\n{msg}\n\n是否要強制重試？\n（將嘗試更積極地關閉相關程式）"
                )
                if response:
                    self._log("使用者選擇強制重試...")
                    self._force_unmount_retry(mdir, commit)
            else:
                messagebox.showerror("卸載失敗", f"卸載失敗:\n{msg}")

    def _force_unmount_retry(self, mdir: str, commit: bool):
        """強制重試卸載，使用更積極的方法"""
        self._log("正在執行強制卸載重試...")
        
        try:
            # 嘗試使用 taskkill 關閉可能鎖定檔案的程式
            self._log("嘗試關閉可能鎖定檔案的程式...")
            
            # 使用 handle.exe 或 lsof 類似功能（如果可用）
            # 這裡使用簡單的方法：關閉所有 explorer.exe
            result = subprocess.run(['taskkill', '/F', '/IM', 'explorer.exe'], 
                                  capture_output=True, text=True)
            if result.returncode == 0:
                self._log("已終止 explorer.exe 程序")
                # 重新啟動 explorer
                subprocess.Popen(['explorer.exe'])
                self._log("已重新啟動 explorer.exe")
            
            # 等待程序完全終止
            import time
            time.sleep(2)
            
            # 再次嘗試卸載
            self._log("重新嘗試卸載...")
            ok, msg = WIMManager.unmount_wim(mdir, commit)
            
            if ok:
                self._log("✓ 強制卸載成功！")
                messagebox.showinfo("卸載成功", "強制卸載成功！")
            else:
                self._log(f"✗ 強制卸載仍然失敗: {msg}")
                messagebox.showerror("卸載失敗", f"強制卸載仍然失敗:\n{msg}\n\n建議手動重開機後再試")
                
        except Exception as e:
            self._log(f"強制卸載過程中發生錯誤: {e}")
            messagebox.showerror("錯誤", f"強制卸載過程中發生錯誤: {e}")

    # ---------- Driver 事件 ----------
    def _on_browse_driver_mount_dir(self):
        path = filedialog.askdirectory(title="選擇已掛載的映像資料夾")
        if path:
            self.var_driver_mount_dir.set(path)
            self._log(f"已選擇映像掛載路徑：{path}")
            self._save_config()

    def _on_sync_from_wim1(self):
        """從 WIM#1 分頁同步掛載路徑"""
        if not hasattr(self, 'var_mount_dir'):
            messagebox.showwarning("同步失敗", "找不到 WIM#1 分頁的掛載路徑")
            return
            
        wim_mount_dir = self.var_mount_dir.get().strip()
        if not wim_mount_dir:
            messagebox.showwarning("同步失敗", "WIM#1 分頁的掛載路徑為空\n請先在 WIM 分頁設定 WIM#1 掛載路徑")
            return
            
        self.var_driver_mount_dir.set(wim_mount_dir)
        self._log(f"✓ 已從 WIM#1 分頁同步掛載路徑：{wim_mount_dir}")
        self._save_config()
        messagebox.showinfo("同步成功", f"已同步 WIM#1 掛載路徑：\n{wim_mount_dir}")
        
    def _on_sync_from_wim2(self):
        """從 WIM#2 分頁同步掛載路徑"""
        if not hasattr(self, 'var_mount_dir2'):
            messagebox.showwarning("同步失敗", "找不到 WIM#2 分頁的掛載路徑")
            return
            
        wim_mount_dir = self.var_mount_dir2.get().strip()
        if not wim_mount_dir:
            messagebox.showwarning("同步失敗", "WIM#2 分頁的掛載路徑為空\n請先在 WIM 分頁設定 WIM#2 掛載路徑")
            return
            
        self.var_driver_mount_dir.set(wim_mount_dir)
        self._log(f"✓ 已從 WIM#2 分頁同步掛載路徑：{wim_mount_dir}")
        self._save_config()
        messagebox.showinfo("同步成功", f"已同步 WIM#2 掛載路徑：\n{wim_mount_dir}")

    def _on_browse_driver_source(self):
        path = filedialog.askdirectory(title="選擇驅動程式資料夾")
        if path:
            self.var_driver_source.set(path)
            self._log(f"已選擇驅動程式資料夾：{path}")
            self._save_config()

    def _on_browse_driver_file(self):
        # 根據目前路徑智能選擇初始目錄
        current_path = self.var_driver_source.get().strip()
        initial_dir = None
        if current_path:
            if os.path.isfile(current_path):
                initial_dir = os.path.dirname(current_path)
            elif os.path.isdir(current_path):
                initial_dir = current_path
        
        path = filedialog.askopenfilename(
            title="選擇驅動程式檔案 (.inf)",
            filetypes=[("Driver INF files", "*.inf"), ("All files", "*.*")],
            initialdir=initial_dir
        )
        if path:
            self.var_driver_source.set(path)
            self._log(f"已選擇驅動程式檔案：{path}")
            
            # 檢查是否為 .inf 檔案並顯示資訊
            if path.lower().endswith('.inf'):
                try:
                    # 簡單讀取 .inf 檔案的基本資訊
                    with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                        content = f.read(500)  # 只讀取前 500 字元
                        if 'DriverVer' in content:
                            self._log("✓ 偵測到有效的驅動程式 .inf 檔案")
                        else:
                            self._log("⚠ 警告：可能不是標準的驅動程式 .inf 檔案")
                except:
                    self._log("無法讀取 .inf 檔案內容")
            
            self._save_config()

    def _on_check_mount_status(self):
        mount_dir = self.var_driver_mount_dir.get().strip()
        if not mount_dir:
            messagebox.showwarning("輸入不完整", "請先輸入映像掛載路徑")
            return
            
        self._log("檢查映像掛載狀態...")
        self._thread(self._do_check_mount_status, mount_dir)

    def _do_check_mount_status(self, mount_dir: str):
        # 檢查路徑是否存在
        if not os.path.exists(mount_dir):
            self._log(f"路徑不存在：{mount_dir}")
            return
            
        # 檢查是否有 Windows 資料夾（通常表示這是一個掛載的映像）
        windows_path = os.path.join(mount_dir, "Windows")
        system32_path = os.path.join(windows_path, "System32")
        
        if os.path.exists(windows_path) and os.path.exists(system32_path):
            self._log(f"✓ 映像掛載狀態正常：{mount_dir}")
            self._log("  發現 Windows 系統資料夾")
            messagebox.showinfo("掛載狀態", "映像掛載狀態正常，可以進行驅動程式安裝")
        else:
            self._log(f"⚠ 警告：路徑可能不是已掛載的映像：{mount_dir}")
            self._log("  未發現 Windows 系統資料夾")
            messagebox.showwarning("掛載狀態", "此路徑可能不是已掛載的映像\n請確認路徑正確")

    def _on_install_driver(self):
        mount_dir = self.var_driver_mount_dir.get().strip()
        driver_source = self.var_driver_source.get().strip()
        recurse = self.var_driver_recurse.get()
        force_unsigned = self.var_driver_force_unsigned.get()
        
        if not mount_dir or not driver_source:
            messagebox.showwarning("輸入不完整", "請選擇映像掛載路徑和驅動程式來源")
            return
            
        if not os.path.exists(mount_dir):
            messagebox.showerror("路徑錯誤", "映像掛載路徑不存在")
            return
            
        if not os.path.exists(driver_source):
            messagebox.showerror("路徑錯誤", "驅動程式路徑不存在")
            return
        
        # 檢查是否已掛載
        is_mounted, mount_info, err = WIMManager.is_path_mounted(mount_dir)
        if err:
            self._log(f"⚠ 無法確認掛載狀態: {err}")
        elif not is_mounted:
            result = messagebox.askyesno(
                "未檢測到掛載", 
                f"路徑 '{mount_dir}' 未在 DISM 掛載清單中。\n\n"
                "驅動程式安裝需要已掛載的離線映像。\n\n"
                "是否仍要嘗試安裝？",
                icon='warning'
            )
            if not result:
                return
            
        self._log("開始安裝驅動程式...")
        self._save_config()
        self._thread(self._do_install_driver, mount_dir, driver_source, recurse, force_unsigned)

    def _do_install_driver(self, mount_dir: str, driver_source: str, recurse: bool, force_unsigned: bool):
        global _dism_lock, _dism_busy
        
        # 取得 DISM 鎖
        with _dism_lock:
            _dism_busy = True
            try:
                recurse_text = "遞迴" if recurse else "非遞迴"
                unsigned_text = "允許未簽署" if force_unsigned else "僅簽署"
                
                self._log(f"正在安裝驅動程式...")
                self._log(f"  映像路徑: {mount_dir}")
                self._log(f"  驅動來源: {driver_source}")
                self._log(f"  搜尋模式: {recurse_text}")
                self._log(f"  簽署要求: {unsigned_text}")
                
                ok, msg = DriverManager.add_driver_to_offline_image(mount_dir, driver_source, recurse, force_unsigned)
                if ok:
                    self._log("✓ 驅動程式安裝成功！")
                    self.after(0, lambda: messagebox.showinfo("安裝成功", "驅動程式已成功安裝到離線映像"))
                    # 重新載入驅動清單
                    self.after(100, self._on_refresh_driver_list)
                else:
                    self._log(f"✗ 驅動程式安裝失敗: {msg}")
                    self.after(0, lambda: messagebox.showerror("安裝失敗", f"驅動程式安裝失敗:\n{msg}"))
            finally:
                _dism_busy = False

    def _on_use_extracted_drivers(self):
        """使用萃取結果作為驅動程式來源"""
        output_path = self.var_extract_output.get().strip()
        if not output_path:
            messagebox.showwarning("路徑為空", "請先設定萃取輸出目錄")
            return
            
        if not os.path.exists(output_path):
            messagebox.showwarning("路徑無效", "萃取輸出目錄不存在，請先執行萃取")
            return
            
        self.var_driver_source.set(output_path)
        self._log(f"✓ 已設定萃取結果為驅動程式來源：{output_path}")
        self._save_config()
        messagebox.showinfo("設定完成", f"已將萃取結果設為驅動程式來源：\n{output_path}")

    def _on_list_drivers(self):
        mount_dir = self.var_driver_mount_dir.get().strip()
        if not mount_dir:
            messagebox.showwarning("輸入不完整", "請先輸入映像掛載路徑")
            return
            
        if not os.path.exists(mount_dir):
            messagebox.showerror("路徑錯誤", "映像掛載路徑不存在")
            return
            
        self._log("開始列出已安裝的驱動程式...")
        self._thread(self._do_list_drivers, mount_dir)

    def _do_list_drivers(self, mount_dir: str):
        global _dism_lock, _dism_busy
        
        # 取得 DISM 鎖
        with _dism_lock:
            _dism_busy = True
            try:
                self._log(f"正在查詢映像中的驅動程式: {mount_dir}")
                
                ok, drivers, err = DriverManager.get_drivers_in_offline_image(mount_dir)
                if not ok:
                    self._log(f"查詢驅動程式失敗: {err}")
                    self.after(0, lambda: messagebox.showerror("查詢失敗", f"無法查詢驅動程式:\n{err}"))
                    return
                    
                if not drivers:
                    self._log("映像中沒有找到已安裝的驅動程式")
                    self.after(0, lambda: messagebox.showinfo("查詢結果", "映像中沒有找到已安裝的驅動程式"))
                    return
                    
                self._log(f"找到 {len(drivers)} 個已安裝的驅動程式:")
            finally:
                _dism_busy = False
        for i, driver in enumerate(drivers, 1):
            name = driver.get('PublishedName', 'N/A')
            provider = driver.get('Provider', 'N/A')
            version = driver.get('Version', 'N/A')
            date = driver.get('Date', 'N/A')
            self._log(f"  {i:2d}. {name} - {provider} (v{version}, {date})")
            
        messagebox.showinfo("查詢結果", f"找到 {len(drivers)} 個已安裝的驅動程式\n詳細資訊請查看日誌")

    # ---------- Extract 事件 ----------

    def _on_browse_extract_source(self):
        """選擇驅動程式擷取的來源映像掛載目錄"""
        path = filedialog.askdirectory(title="選擇來源映像掛載目錄")
        if path:
            self.var_extract_source.set(path)
            self._log(f"已選擇來源 WIM 檔案：{path}")
            self._save_config()

    def _on_sync_extract_from_wim1(self):
        """從 WIM#1 分頁同步來源路徑"""
        if not hasattr(self, 'var_mount_dir'):
            messagebox.showwarning("同步失敗", "找不到 WIM#1 分頁的掛載路徑")
            return
            
        wim_mount_dir = self.var_mount_dir.get().strip()
        if not wim_mount_dir:
            messagebox.showwarning("同步失敗", "WIM#1 分頁的掛載路徑為空")
            return
            
        self.var_extract_source.set(wim_mount_dir)
        self._log(f"✓ 已同步來源映像路徑（WIM#1）：{wim_mount_dir}")
        self._save_config()

    def _on_sync_extract_from_wim2(self):
        """從 WIM#2 分頁同步來源路徑"""
        if not hasattr(self, 'var_mount_dir2'):
            messagebox.showwarning("同步失敗", "找不到 WIM#2 分頁的掛載路徑")
            return
            
        wim_mount_dir = self.var_mount_dir2.get().strip()
        if not wim_mount_dir:
            messagebox.showwarning("同步失敗", "WIM#2 分頁的掛載路徑為空")
            return
            
        self.var_extract_source.set(wim_mount_dir)
        self._log(f"✓ 已同步來源映像路徑（WIM#2）：{wim_mount_dir}")
        self._save_config()

    def _on_browse_extract_output(self):
        path = filedialog.askdirectory(title="選擇驅動萃取輸出目錄")
        if path:
            self.var_extract_output.set(path)
            self._log(f"已選擇萃取輸出目錄：{path}")
            self._save_config()

    def _on_create_extract_dir(self):
        """建立萃取目錄"""
        path = self.var_extract_output.get().strip()
        if not path:
            messagebox.showwarning("輸入不完整", "請先輸入萃取目錄路徑")
            return
        
        try:
            if os.path.exists(path):
                messagebox.showinfo("目錄狀態", f"目錄已存在：{path}")
            else:
                os.makedirs(path, exist_ok=True)
                self._log(f"✓ 已建立萃取目錄：{path}")
                messagebox.showinfo("建立成功", f"已建立萃取目錄：{path}")
                self._save_config()
        except Exception as e:
            self._log(f"建立目錄失敗：{e}")
            messagebox.showerror("建立失敗", f"無法建立目錄：{e}")

    def _on_open_extract_dir(self):
        """開啟萃取目錄"""
        path = self.var_extract_output.get().strip()
        if not path or not os.path.exists(path):
            self._log("萃取目錄不存在或路徑無效")
            messagebox.showwarning("路徑無效", "萃取目錄不存在，請先建立目錄")
            return
        try:
            os.startfile(path)
            self._log(f"已開啟萃取目錄：{path}")
        except Exception as e:
            self._log(f"開啟萃取目錄失敗：{e}")
            messagebox.showerror("開啟失敗", f"無法開啟目錄：{e}")

    def _on_extract_drivers(self):
        source_path = self.var_extract_source.get().strip()
        output_path = self.var_extract_output.get().strip()
        
        if not source_path or not output_path:
            messagebox.showwarning("輸入不完整", "請選擇來源映像掛載目錄和萃取輸出目錄")
            return
            
        if not os.path.exists(source_path):
            messagebox.showerror("路徑錯誤", "來源映像掛載目錄不存在")
            return
            
        self._log("開始萃取驅動程式...")
        self._save_config()
        self._thread(self._do_extract_drivers, source_path, output_path)

    def _do_extract_drivers(self, source_path: str, output_path: str):
        self._log(f"正在從映像萃取驅動程式...")
        self._log(f"  來源映像目錄: {source_path}")
        self._log(f"  輸出目錄: {output_path}")
        
        ok, msg = DriverManager.export_drivers_from_offline_image(source_path, output_path)
        if ok:
            self._log("✓ 驅動程式萃取成功！")
            self._log(f"驅動程式已萃取到: {output_path}")
            
            # 自動將萃取結果設為驅動程式來源
            if hasattr(self, 'var_driver_source'):
                self.var_driver_source.set(output_path)
                self._log("✓ 已自動設定為驅動程式來源")
                self._save_config()
            
            messagebox.showinfo("萃取成功", f"驅動程式已成功萃取到:\n{output_path}\n\n已自動設為驅動程式來源")
        else:
            self._log(f"✗ 驅動程式萃取失敗: {msg}")
            messagebox.showerror("萃取失敗", f"驅動程式萃取失敗:\n{msg}")

    def _on_view_extracted_drivers(self):
        output_path = self.var_extract_output.get().strip()
        if not output_path or not os.path.exists(output_path):
            messagebox.showwarning("路徑無效", "萃取目錄不存在或無效")
            return
            
        self._log("正在掃描萃取的驅動程式...")
        self._thread(self._do_view_extracted_drivers, output_path)

    def _do_view_extracted_drivers(self, output_path: str):
        ok, drivers, msg = DriverManager.get_driver_info_from_path(output_path)
        if ok:
            self._log(f"✓ {msg}")
            if drivers:
                self._log("萃取的驅動程式清單:")
                for i, driver in enumerate(drivers, 1):
                    self._log(f"  {i:2d}. {driver['name']} ({driver.get('folder', driver['path'])})")
                messagebox.showinfo("掃描結果", f"{msg}\n詳細清單請查看日誌")
            else:
                messagebox.showinfo("掃描結果", "未找到任何 .inf 驅動程式檔案")
        else:
            self._log(f"✗ 掃描失敗: {msg}")
            messagebox.showerror("掃描失敗", f"掃描驅動程式失敗:\n{msg}")

    # ---------- 整合版驅動管理功能 ----------
    
    def _on_extract_selected_drivers(self):
        """提取驅動程式（整合版）- 根據選取狀態決定提取全部或選定"""
        mount_dir = self.var_driver_list_mount_dir.get().strip()
        if not mount_dir:
            messagebox.showwarning("未選擇映像", "請先選擇一個已掛載的映像")
            return
        
        # 檢查是否有勾選的項目
        checked_drivers = self._get_checked_drivers()
        total_count = len(self.driver_tree.get_children())
        
        if not total_count:
            messagebox.showwarning("清單為空", "沒有可提取的驅動程式")
            return
        
        # 取得預設輸出目錄（從設定）
        default_path = self._cfg_get('EXTRACT', 'output_path') or ""
        
        # 彈出對話框選擇輸出目錄
        output_path = filedialog.askdirectory(
            title="選擇驅動程式提取輸出目錄",
            initialdir=default_path if default_path and os.path.exists(default_path) else None
        )
        if not output_path:
            return
        
        # 確認訊息根據是否有選取來調整
        if checked_drivers:
            # 有勾選 -> 只提取選定的 (注意: DISM 只能全部萃取，所以實際還是全部)
            confirm_msg = (f"將從映像提取驅動程式到:\n{output_path}\n\n"
                          f"已選取 {len(checked_drivers)} 個驅動程式\n"
                          f"（注意：DISM 會提取所有驅動程式）\n\n繼續？")
        else:
            # 沒勾選 -> 提取全部
            confirm_msg = (f"將從映像提取所有驅動程式到:\n{output_path}\n\n"
                          f"共 {total_count} 個驅動程式\n\n繼續？")
        
        if not messagebox.askyesno("確認提取", confirm_msg):
            return
        
        self._log(f"開始提取驅動程式到: {output_path}")
        self.var_driver_list_status.set("正在提取驅動程式...")
        self._thread(self._do_extract_all_drivers, mount_dir, output_path)

    def _do_extract_all_drivers(self, mount_dir: str, output_path: str):
        """實際執行提取全部驅動"""
        global _dism_lock, _dism_busy
        
        # 取得 DISM 鎖
        with _dism_lock:
            _dism_busy = True
            try:
                ok, msg = DriverManager.export_drivers_from_offline_image(mount_dir, output_path)
                
                if ok:
                    self._log("✓ 驅動程式提取成功！")
                    self._log(f"驅動程式已提取到: {output_path}")
                    self.after(0, lambda: self._update_selection_count())
                    
                    def show_success():
                        result = messagebox.askyesno("提取成功", 
                            f"驅動程式已成功提取到:\n{output_path}\n\n是否開啟該資料夾？")
                        if result:
                            os.startfile(output_path)
                    
                    self.after(0, show_success)
                else:
                    self._log(f"✗ 驅動程式提取失敗: {msg}")
                    self.after(0, lambda: self._update_selection_count())
                    self.after(0, lambda: messagebox.showerror("提取失敗", f"驅動程式提取失敗:\n{msg}"))
            finally:
                _dism_busy = False

    def _on_add_driver_dialog(self):
        """新增驅動程式對話框 - 重新設計 UX"""
        mount_dir = self.var_driver_list_mount_dir.get().strip()
        if not mount_dir:
            messagebox.showwarning("未選擇映像", "請先選擇一個已掛載的映像")
            return
        
        # 建立對話框
        dialog = tk.Toplevel(self)
        dialog.title("新增驅動程式")
        dialog.geometry("500x400")
        dialog.resizable(False, False)
        dialog.grab_set()
        dialog.transient(self)
        
        # 主框架
        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 目標映像顯示
        target_frame = ttk.Frame(main_frame)
        target_frame.pack(fill=tk.X, pady=(0, 12))
        ttk.Label(target_frame, text="目標映像:", font=("Microsoft JhengHei UI", 9)).pack(side=tk.LEFT)
        ttk.Label(target_frame, text=mount_dir, foreground="blue", font=("Microsoft JhengHei UI", 9)).pack(side=tk.LEFT, padx=(8, 0))
        
        # 分隔線
        ttk.Separator(main_frame, orient='horizontal').pack(fill=tk.X, pady=(0, 12))
        
        # === 安裝模式選擇 ===
        var_mode = tk.StringVar(value="single")
        var_source = tk.StringVar()
        var_force = tk.BooleanVar(value=False)
        
        # 單支安裝區塊
        single_frame = ttk.LabelFrame(main_frame, text="📄 單支安裝", padding=10)
        single_frame.pack(fill=tk.X, pady=(0, 10))
        
        single_inner = ttk.Frame(single_frame)
        single_inner.pack(fill=tk.X)
        
        rb_single = ttk.Radiobutton(single_inner, text="選擇單一 .inf 檔案", variable=var_mode, value="single")
        rb_single.pack(side=tk.LEFT)
        
        var_single_path = tk.StringVar()
        ent_single = ttk.Entry(single_inner, textvariable=var_single_path, width=28, state="readonly")
        ent_single.pack(side=tk.LEFT, padx=(10, 6))
        
        def browse_inf():
            path = filedialog.askopenfilename(
                title="選擇驅動程式檔案 (.inf)",
                filetypes=[("Driver INF files", "*.inf"), ("All files", "*.*")]
            )
            if path:
                var_single_path.set(path.replace('/', '\\'))
                var_mode.set("single")
        
        ttk.Button(single_inner, text="瀏覽...", command=browse_inf, width=8).pack(side=tk.LEFT)
        
        ttk.Label(single_frame, text="適用於：安裝特定的單一驅動程式", foreground="gray", 
                 font=("Microsoft JhengHei UI", 8)).pack(anchor="w", pady=(6, 0))
        
        # 批量安裝區塊
        batch_frame = ttk.LabelFrame(main_frame, text="📁 批量安裝", padding=10)
        batch_frame.pack(fill=tk.X, pady=(0, 10))
        
        batch_inner = ttk.Frame(batch_frame)
        batch_inner.pack(fill=tk.X)
        
        rb_batch = ttk.Radiobutton(batch_inner, text="選擇驅動程式資料夾", variable=var_mode, value="batch")
        rb_batch.pack(side=tk.LEFT)
        
        var_batch_path = tk.StringVar()
        ent_batch = ttk.Entry(batch_inner, textvariable=var_batch_path, width=28, state="readonly")
        ent_batch.pack(side=tk.LEFT, padx=(10, 6))
        
        def browse_folder():
            path = filedialog.askdirectory(title="選擇驅動程式資料夾")
            if path:
                var_batch_path.set(path.replace('/', '\\'))
                var_mode.set("batch")
        
        ttk.Button(batch_inner, text="瀏覽...", command=browse_folder, width=8).pack(side=tk.LEFT)
        
        # 批量安裝選項
        batch_options = ttk.Frame(batch_frame)
        batch_options.pack(fill=tk.X, pady=(8, 0))
        
        var_recurse = tk.BooleanVar(value=True)
        ttk.Checkbutton(batch_options, text="包含子資料夾 (遞迴搜尋所有 .inf)", variable=var_recurse).pack(side=tk.LEFT, padx=(20, 0))
        
        ttk.Label(batch_frame, text="適用於：一次安裝資料夾內的多個驅動程式", foreground="gray",
                 font=("Microsoft JhengHei UI", 8)).pack(anchor="w", pady=(6, 0))
        
        # 進階選項
        adv_frame = ttk.Frame(main_frame)
        adv_frame.pack(fill=tk.X, pady=(5, 12))
        
        ttk.Checkbutton(adv_frame, text="允許未簽署的驅動程式 (ForceUnsigned)", variable=var_force).pack(side=tk.LEFT)
        
        # 按鈕
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        def do_install():
            mode = var_mode.get()
            
            if mode == "single":
                source = var_single_path.get().strip()
                if not source:
                    messagebox.showwarning("請選擇檔案", "請先選擇要安裝的 .inf 檔案", parent=dialog)
                    return
                recurse = False  # 單支安裝不需要遞迴
            else:
                source = var_batch_path.get().strip()
                if not source:
                    messagebox.showwarning("請選擇資料夾", "請先選擇驅動程式資料夾", parent=dialog)
                    return
                recurse = var_recurse.get()
            
            if not os.path.exists(source):
                messagebox.showerror("路徑錯誤", "指定的路徑不存在", parent=dialog)
                return
            
            # 儲存設定
            self.var_driver_recurse.set(recurse)
            self.var_driver_force_unsigned.set(var_force.get())
            self._save_config()
            
            dialog.destroy()
            
            # 執行安裝
            mode_text = "單支" if mode == "single" else "批量"
            self._log(f"開始{mode_text}安裝驅動程式: {source}")
            self.var_driver_list_status.set("正在安裝驅動程式...")
            self._thread(self._do_add_driver, mount_dir, source, recurse, var_force.get())
        
        ttk.Button(btn_frame, text="新增", command=do_install, width=12).pack(side=tk.RIGHT)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy, width=12).pack(side=tk.RIGHT, padx=(0, 8))
        
        # 居中
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (500 // 2)
        y = (dialog.winfo_screenheight() // 2) - (400 // 2)
        dialog.geometry(f"500x400+{x}+{y}")

    def _do_add_driver(self, mount_dir: str, source: str, recurse: bool, force_unsigned: bool):
        """實際執行驅動安裝"""
        self._log(f"正在安裝驅動程式...")
        self._log(f"  目標映像: {mount_dir}")
        self._log(f"  驅動來源: {source}")
        self._log(f"  遞迴搜尋: {'是' if recurse else '否'}")
        self._log(f"  強制未簽署: {'是' if force_unsigned else '否'}")
        
        ok, msg = DriverManager.add_driver_to_offline_image(mount_dir, source, recurse, force_unsigned)
        
        if ok:
            self._log("✓ 驅動程式安裝成功！")
            self.after(0, lambda: self.var_driver_list_status.set("安裝完成，重新載入清單..."))
            self.after(0, lambda: messagebox.showinfo("安裝成功", "驅動程式已成功安裝\n\n清單將自動重新載入"))
            # 重新載入驅動清單
            self.after(500, self._on_refresh_driver_list)
        else:
            self._log(f"✗ 驅動程式安裝失敗: {msg}")
            self.after(0, lambda: self.var_driver_list_status.set("安裝失敗"))
            self.after(0, lambda: messagebox.showerror("安裝失敗", f"驅動程式安裝失敗:\n{msg}"))

    # ---------- 驅動清單事件 ----------
    
    def _on_browse_driver_list_mount_dir(self):
        """選擇驅動清單的目標映像路徑"""
        path = filedialog.askdirectory(title="選擇映像掛載目錄")
        if path:
            self.var_driver_list_mount_dir.set(path)
            self._log(f"已選擇映像掛載目錄：{path}")
            # 自動載入驅動清單
            self._on_refresh_driver_list()

    def _on_sync_list_from_wim1(self):
        """從 WIM#1 同步目標映像路徑"""
        if not hasattr(self, 'var_mount_dir'):
            messagebox.showwarning("同步失敗", "找不到 WIM#1 分頁的掛載路徑")
            return
        wim_mount_dir = self.var_mount_dir.get().strip()
        if not wim_mount_dir:
            messagebox.showwarning("同步失敗", "WIM#1 分頁的掛載路徑為空")
            return
        self.var_driver_list_mount_dir.set(wim_mount_dir)
        self._log(f"✓ 已同步目標映像路徑（WIM#1）：{wim_mount_dir}")
        # 自動載入驅動清單
        self._on_refresh_driver_list()

    def _on_sync_list_from_wim2(self):
        """從 WIM#2 同步目標映像路徑"""
        if not hasattr(self, 'var_mount_dir2'):
            messagebox.showwarning("同步失敗", "找不到 WIM#2 分頁的掛載路徑")
            return
        wim_mount_dir = self.var_mount_dir2.get().strip()
        if not wim_mount_dir:
            messagebox.showwarning("同步失敗", "WIM#2 分頁的掛載路徑為空")
            return
        self.var_driver_list_mount_dir.set(wim_mount_dir)
        self._log(f"✓ 已同步目標映像路徑（WIM#2）：{wim_mount_dir}")
        # 自動載入驅動清單
        self._on_refresh_driver_list()

    def _on_refresh_driver_list(self):
        """重新整理驅動清單"""
        mount_dir = self.var_driver_list_mount_dir.get().strip()
        if not mount_dir:
            messagebox.showwarning("輸入不完整", "請先輸入或選擇映像掛載路徑")
            return
        if not os.path.exists(mount_dir):
            messagebox.showerror("路徑錯誤", "映像掛載路徑不存在")
            return
        
        self._log("開始載入驅動清單...")
        self.var_driver_list_status.set("正在載入...")
        self._thread(self._do_refresh_driver_list, mount_dir)

    def _do_refresh_driver_list(self, mount_dir: str):
        """實際執行驅動清單載入"""
        global _dism_lock, _dism_busy
        
        # 嘗試取得 DISM 鎖
        acquired = _dism_lock.acquire(blocking=False)
        if not acquired:
            # 已有其他 DISM 操作正在執行
            self.after(0, lambda: self._log("⏳ 正在等待其他 DISM 操作完成..."))
            # 等待鎖釋放後重試
            _dism_lock.acquire()  # 阻塞等待
        
        try:
            _dism_busy = True
            ok, drivers, err = DriverManager.get_drivers_in_offline_image(mount_dir)
            
            # 在主執行緒更新 UI
            self.after(0, lambda: self._update_driver_tree(ok, drivers, err))
        finally:
            _dism_busy = False
            _dism_lock.release()

    def _update_driver_tree(self, ok: bool, drivers: list[dict], err: str):
        """更新驅動清單 Treeview"""
        # 清空現有項目
        for item in self.driver_tree.get_children():
            self.driver_tree.delete(item)
        
        # 清空 PublishedName 對照表
        self._driver_published_names = {}
        
        if not ok:
            self._log(f"✗ 載入驅動清單失敗: {err}")
            self.var_driver_list_status.set(f"載入失敗: {err}")
            messagebox.showerror("載入失敗", f"無法載入驅動清單:\n{err}")
            return
        
        if not drivers:
            self._log("映像中沒有安裝任何驅動程式")
            self.var_driver_list_status.set("沒有找到任何驅動程式")
            return
        
        # 填充資料
        for driver in drivers:
            published_name = driver.get('PublishedName', 'N/A')  # oemX.inf
            original_name = driver.get('OriginalFileName', '')   # 原始檔名
            provider = driver.get('Provider', 'N/A')
            version = driver.get('Version', 'N/A')
            date = driver.get('Date', 'N/A')
            class_name = driver.get('ClassName', 'N/A')
            
            # 驅動名稱：優先顯示原始檔名，沒有則顯示 oemX.inf
            display_name = original_name if original_name else published_name
            
            # 使用 PublishedName 作為 item id，方便後續操作
            # 新增勾選欄位 (預設未勾選)
            item_id = self.driver_tree.insert("", tk.END, values=("☐", display_name, published_name, provider, version, date, class_name))
            # 儲存 item_id 與 PublishedName 的對照
            self._driver_published_names[item_id] = published_name
        
        # 重設全選狀態
        self._driver_select_all = False
        self.driver_tree.heading("select", text="☐ 選取")
        
        count = len(drivers)
        self._log(f"✓ 已載入 {count} 個驅動程式")
        self.var_driver_list_status.set(f"共 {count} 個驅動程式，已選取 0 個")

    def _sort_driver_tree(self, col: str):
        """排序驅動清單"""
        import re
        
        # 切換排序方向
        if self._driver_tree_sort_col == col:
            self._driver_tree_sort_reverse = not self._driver_tree_sort_reverse
        else:
            self._driver_tree_sort_col = col
            self._driver_tree_sort_reverse = False
        
        # 取得所有項目
        items = [(self.driver_tree.set(item, col), item) for item in self.driver_tree.get_children("")]
        
        # 定義排序鍵函數
        def sort_key(item_tuple):
            value = item_tuple[0]
            
            # INF 檔案排序：提取 oem 後面的數字
            if col == "inf":
                match = re.search(r'oem(\d+)\.inf', value.lower())
                if match:
                    return (0, int(match.group(1)))  # (優先級, 數字)
                return (1, value.lower())  # 非 oem 格式排在後面
            
            # 日期排序：嘗試解析日期
            if col == "date":
                # 嘗試多種日期格式
                for fmt in ('%Y/%m/%d', '%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y'):
                    try:
                        from datetime import datetime
                        dt = datetime.strptime(value, fmt)
                        return dt
                    except ValueError:
                        continue
                return value  # 無法解析則按字串排序
            
            # 版本排序：嘗試按版本號排序
            if col == "version":
                parts = re.split(r'[.\-]', value)
                try:
                    return tuple(int(p) if p.isdigit() else p for p in parts)
                except:
                    return (value,)
            
            # 其他欄位：字串排序（不分大小寫）
            return value.lower() if isinstance(value, str) else value
        
        # 排序
        items.sort(key=sort_key, reverse=self._driver_tree_sort_reverse)
        
        # 重新排列
        for index, (_, item) in enumerate(items):
            self.driver_tree.move(item, "", index)
        
        # 更新標題顯示排序方向
        arrow = " ▼" if self._driver_tree_sort_reverse else " ▲"
        col_titles = {
            "select": "☐ 選取",
            "name": "驅動名稱",
            "inf": "INF 檔案",
            "provider": "提供者",
            "version": "版本",
            "date": "日期",
            "class": "類型"
        }
        for c, title in col_titles.items():
            if c == col:
                self.driver_tree.heading(c, text=title + arrow)
            elif c != "select":  # 不改變選取欄的標題
                self.driver_tree.heading(c, text=title)

    def _on_select_all_drivers(self):
        """全選驅動"""
        for item in self.driver_tree.get_children():
            self.driver_tree.selection_add(item)
        count = len(self.driver_tree.selection())
        self.var_driver_list_status.set(f"已選取 {count} 個驅動程式")

    def _on_deselect_all_drivers(self):
        """取消全選"""
        self.driver_tree.selection_remove(self.driver_tree.selection())
        self.var_driver_list_status.set(f"共 {len(self.driver_tree.get_children())} 個驅動程式")

    # ===== 搜尋功能 =====
    def _on_driver_search(self):
        """執行驅動名稱搜尋"""
        keyword = self.var_driver_search.get().strip()
        if not keyword:
            messagebox.showinfo("搜尋", "請輸入搜尋關鍵字")
            return
        
        # 搜尋所有驅動名稱（不分大小寫）
        results = []
        keyword_lower = keyword.lower()
        
        for item in self.driver_tree.get_children():
            values = self.driver_tree.item(item, 'values')
            if len(values) >= 2:
                driver_name = values[1]  # 第二欄是驅動名稱
                if keyword_lower in driver_name.lower():
                    results.append((item, driver_name))
        
        if not results:
            messagebox.showinfo("搜尋結果", f"找不到包含 \"{keyword}\" 的驅動程式")
            return
        
        # 顯示搜尋結果視窗
        self._show_search_results_dialog(keyword, results)
    
    def _on_clear_search(self):
        """清除搜尋"""
        self.var_driver_search.set("")
        self.driver_tree.selection_remove(self.driver_tree.selection())
    
    def _show_search_results_dialog(self, keyword: str, results: list):
        """顯示搜尋結果對話框"""
        dialog = tk.Toplevel(self)
        dialog.title(f"搜尋結果 - \"{keyword}\"")
        dialog.geometry("450x200")
        dialog.resizable(True, False)
        dialog.transient(self)
        
        # 搜尋狀態
        current_index = [0]  # 使用 list 以便在內部函數中修改
        
        # 主框架
        main_frame = ttk.Frame(dialog, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 結果統計
        stats_frame = ttk.Frame(main_frame)
        stats_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(stats_frame, text=f"找到 {len(results)} 個符合的結果", 
                 font=("Microsoft JhengHei UI", 10)).pack(side=tk.LEFT)
        
        var_position = tk.StringVar(value=f"第 1/{len(results)} 個")
        lbl_position = ttk.Label(stats_frame, textvariable=var_position, 
                                font=("Microsoft JhengHei UI", 10, "bold"))
        lbl_position.pack(side=tk.RIGHT)
        
        # 當前結果顯示
        result_frame = ttk.LabelFrame(main_frame, text="驅動名稱", padding=10)
        result_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        var_result = tk.StringVar(value=results[0][1] if results else "")
        lbl_result = ttk.Label(result_frame, textvariable=var_result, 
                              font=("Consolas", 11), wraplength=400, anchor="w")
        lbl_result.pack(fill=tk.BOTH, expand=True)
        
        # 高亮當前項目
        def highlight_current():
            item_id, driver_name = results[current_index[0]]
            # 清除之前的選取
            self.driver_tree.selection_remove(self.driver_tree.selection())
            # 選取並顯示當前項目
            self.driver_tree.selection_set(item_id)
            self.driver_tree.see(item_id)
            self.driver_tree.focus(item_id)
            # 更新顯示
            var_result.set(driver_name)
            var_position.set(f"第 {current_index[0] + 1}/{len(results)} 個")
        
        # 上一個
        def go_prev():
            if current_index[0] > 0:
                current_index[0] -= 1
                highlight_current()
                update_buttons_state()
        
        # 下一個
        def go_next():
            if current_index[0] < len(results) - 1:
                current_index[0] += 1
                highlight_current()
                update_buttons_state()
        
        # 按鈕框架
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X)
        
        btn_prev = ttk.Button(btn_frame, text="◀ 上一個", command=go_prev, width=12)
        btn_prev.pack(side=tk.LEFT)
        
        btn_next = ttk.Button(btn_frame, text="下一個 ▶", command=go_next, width=12)
        btn_next.pack(side=tk.LEFT, padx=(8, 0))
        
        ttk.Button(btn_frame, text="關閉", command=dialog.destroy, width=10).pack(side=tk.RIGHT)
        
        # 更新按鈕狀態
        def update_buttons_state():
            btn_prev.config(state=tk.NORMAL if current_index[0] > 0 else tk.DISABLED)
            btn_next.config(state=tk.NORMAL if current_index[0] < len(results) - 1 else tk.DISABLED)
        
        # 綁定鍵盤快捷鍵
        dialog.bind('<Left>', lambda e: go_prev())
        dialog.bind('<Right>', lambda e: go_next())
        dialog.bind('<Escape>', lambda e: dialog.destroy())
        dialog.bind('<Return>', lambda e: go_next())
        
        # 初始化
        highlight_current()
        update_buttons_state()
        
        # 居中顯示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (450 // 2)
        y = (dialog.winfo_screenheight() // 2) - (200 // 2)
        dialog.geometry(f"450x200+{x}+{y}")
        
        dialog.focus_set()

    def _on_view_driver_details(self):
        """查看選定驅動的詳細資訊"""
        selected = self.driver_tree.selection()
        if not selected:
            messagebox.showwarning("未選取", "請先選取一個驅動程式")
            return
        
        # 只取第一個選取的項目
        item = selected[0]
        
        # 從對照表取得真正的 PublishedName
        if hasattr(self, '_driver_published_names') and item in self._driver_published_names:
            driver_name = self._driver_published_names[item]
        else:
            # 嘗試從顯示名稱中提取 (oemX.inf)
            display_name = self.driver_tree.item(item, 'values')[0]
            import re
            match = re.search(r'\(([^)]+\.inf)\)', display_name)
            driver_name = match.group(1) if match else display_name
        
        mount_dir = self.var_driver_list_mount_dir.get().strip()
        if not mount_dir:
            messagebox.showwarning("路徑錯誤", "請先設定映像掛載路徑")
            return
        
        self._log(f"正在查詢驅動程式詳情: {driver_name}")
        self._thread(self._do_view_driver_details, mount_dir, driver_name)

    def _do_view_driver_details(self, mount_dir: str, driver_name: str):
        """實際執行驅動詳情查詢"""
        ok, info, err = DriverManager.get_driver_details(mount_dir, driver_name)
        
        if not ok:
            self._log(f"✗ 查詢驅動詳情失敗: {err}")
            self.after(0, lambda: messagebox.showerror("查詢失敗", f"無法查詢驅動詳情:\n{err}"))
            return
        
        # 格式化詳情
        details = f"驅動程式詳細資訊\n{'='*40}\n\n"
        for key, value in info.items():
            # 格式化 key
            display_key = key
            if key == "PublishedName":
                display_key = "發佈名稱"
            elif key == "OriginalFileName":
                display_key = "原始檔名"
            elif key == "Inbox":
                display_key = "內建驅動"
            elif key == "ClassName":
                display_key = "類型"
            elif key == "Provider":
                display_key = "提供者"
            elif key == "Date":
                display_key = "日期"
            elif key == "Version":
                display_key = "版本"
            details += f"{display_key}: {value}\n"
        
        self._log(f"驅動程式 {driver_name} 詳情:\n{details}")
        self.after(0, lambda: self._show_driver_details_dialog(driver_name, details))

    def _show_driver_details_dialog(self, driver_name: str, details: str):
        """顯示驅動詳情對話框"""
        dialog = tk.Toplevel(self)
        dialog.title(f"驅動程式詳情 - {driver_name}")
        dialog.geometry("500x400")
        dialog.resizable(True, True)
        dialog.grab_set()
        
        # 文字框
        text_frame = ttk.Frame(dialog)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        text_widget = tk.Text(text_frame, wrap=tk.WORD, font=("Consolas", 10))
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        text_widget.insert("1.0", details)
        text_widget.configure(state="disabled")
        
        # 關閉按鈕
        ttk.Button(dialog, text="關閉", command=dialog.destroy).pack(pady=(0, 10))
        
        # 居中
        dialog.transient(self)
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (500 // 2)
        y = (dialog.winfo_screenheight() // 2) - (400 // 2)
        dialog.geometry(f"500x400+{x}+{y}")

    def _on_remove_selected_drivers(self):
        """移除選定的驅動程式（使用勾選的項目）"""
        # 使用勾選的項目而非 selection
        checked_drivers = self._get_checked_drivers()
        if not checked_drivers:
            messagebox.showwarning("未選取", "請先勾選要移除的驅動程式")
            return
        
        mount_dir = self.var_driver_list_mount_dir.get().strip()
        if not mount_dir:
            messagebox.showwarning("路徑錯誤", "請先設定映像掛載路徑")
            return
        
        # 檢查是否已掛載
        is_mounted, mount_info, err = WIMManager.is_path_mounted(mount_dir)
        if err:
            self._log(f"⚠ 無法確認掛載狀態: {err}")
        elif not is_mounted:
            result = messagebox.askyesno(
                "未檢測到掛載", 
                f"路徑 '{mount_dir}' 未在 DISM 掛載清單中。\n\n"
                "驅動程式移除需要已掛載的離線映像。\n\n"
                "是否仍要嘗試移除？",
                icon='warning'
            )
            if not result:
                return
        
        # 取得勾選項目的顯示名稱（用於確認對話框）
        display_names = []
        driver_name_map = {}  # {oem_inf: display_name}
        for item in self.driver_tree.get_children():
            values = self.driver_tree.item(item, 'values')
            if values[0] == "☑":
                oem_inf = values[2]       # INF 檔案欄位 (oemX.inf)
                display_name = values[1]  # 驅動名稱欄位
                display_names.append(display_name)
                driver_name_map[oem_inf] = display_name
        
        # 確認對話框
        count = len(checked_drivers)
        if count == 1:
            confirm_msg = f"確定要移除驅動程式 '{display_names[0]}'？"
        else:
            confirm_msg = f"確定要移除以下 {count} 個驅動程式？\n\n"
            for name in display_names[:10]:  # 最多顯示10個
                confirm_msg += f"  • {name}\n"
            if count > 10:
                confirm_msg += f"  ... 還有 {count - 10} 個\n"
        
        confirm_msg += "\n此操作無法復原！"
        
        if not messagebox.askyesno("確認移除", confirm_msg, icon='warning'):
            return
        
        self._log(f"開始移除 {count} 個驅動程式...")
        self.var_driver_list_status.set(f"正在移除 {count} 個驅動程式...")
        self._thread(self._do_remove_drivers, mount_dir, checked_drivers, driver_name_map)

    def _do_remove_drivers(self, mount_dir: str, driver_names: list[str], driver_name_map: dict = None):
        """實際執行驅動移除"""
        global _dism_lock, _dism_busy
        
        if driver_name_map is None:
            driver_name_map = {}
        
        # 取得 DISM 鎖
        with _dism_lock:
            _dism_busy = True
            try:
                def progress_callback(current, total, driver_name, success, message):
                    # 使用驅動名稱顯示（如果有對應）
                    display_name = driver_name_map.get(driver_name, driver_name)
                    status = "✓" if success else "✗"
                    self._log(f"  [{current}/{total}] {status} {display_name}")
                    self.after(0, lambda: self.var_driver_list_status.set(f"正在移除... ({current}/{total})"))
                
                success_count, fail_count, errors = DriverManager.remove_drivers_batch(
                    mount_dir, driver_names, progress_callback
                )
                
                # 顯示結果
                result_msg = f"移除完成\n\n成功: {success_count} 個\n失敗: {fail_count} 個"
                if errors:
                    result_msg += f"\n\n失敗詳情:\n"
                    for err in errors[:5]:
                        result_msg += f"  • {err}\n"
                    if len(errors) > 5:
                        result_msg += f"  ... 還有 {len(errors) - 5} 個錯誤"
                
                self._log(f"驅動移除完成: 成功 {success_count}, 失敗 {fail_count}")
                
                if fail_count > 0:
                    self.after(0, lambda: messagebox.showwarning("移除完成（有錯誤）", result_msg))
                else:
                    self.after(0, lambda: messagebox.showinfo("移除完成", result_msg))
            finally:
                _dism_busy = False
        
        # 重新整理清單（在鎖釋放後）
        self.after(100, self._on_refresh_driver_list)

    def _on_export_driver_list(self):
        """匯出驅動清單到 CSV"""
        items = self.driver_tree.get_children()
        if not items:
            messagebox.showwarning("清單為空", "沒有可匯出的驅動程式")
            return
        
        # 取得輸出目錄
        output_dir = self._cfg_get('EXPORT', 'output_path') or SCRIPT_DIR
        output_dir = output_dir.replace('/', '\\')
        
        # 生成檔名（包含時間戳）
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"driver_list_{timestamp}.csv"
        
        # 完整路徑
        file_path = os.path.join(output_dir, filename)
        
        # 確保目錄存在
        os.makedirs(output_dir, exist_ok=True)
        
        try:
            with open(file_path, 'w', encoding='utf-8-sig') as f:
                # 寫入標題（跳過勾選欄位）
                f.write("驅動名稱,INF 檔案,提供者,版本,日期,類型\n")
                
                # 寫入資料（跳過第一個勾選欄位）
                for item in items:
                    values = self.driver_tree.item(item, 'values')
                    # 跳過第一欄（勾選狀態），取後面的欄位
                    data_values = values[1:] if len(values) > 1 else values
                    # 處理可能包含逗號的值
                    row = [f'"{v}"' if ',' in str(v) else str(v) for v in data_values]
                    f.write(",".join(row) + "\n")
            
            self._log(f"✓ 驅動清單已匯出到: {file_path}")
            
            # 顯示成功，提供開啟資料夾或確定選項
            result = messagebox.askyesno("匯出成功", 
                f"驅動清單已成功匯出到:\n{file_path}\n\n共匯出 {len(items)} 筆資料\n\n是否開啟資料夾？")
            if result:
                # 開啟資料夾並選取檔案
                subprocess.run(['explorer', '/select,', file_path])
                
        except Exception as e:
            self._log(f"✗ 匯出失敗: {e}")
            messagebox.showerror("匯出失敗", f"無法匯出驅動清單:\n{e}")

    # ========== 設定對話框 ==========
    def _on_open_settings(self):
        """開啟設定對話框"""
        dialog = tk.Toplevel(self)
        dialog.title("設定")
        dialog.geometry("550x200")
        dialog.resizable(False, False)
        dialog.grab_set()
        dialog.transient(self)
        
        # 置中顯示
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 550) // 2
        y = self.winfo_y() + (self.winfo_height() - 200) // 2
        dialog.geometry(f"550x200+{x}+{y}")
        
        # 主框架
        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # === 驅動程式設定 ===
        driver_frame = ttk.LabelFrame(main_frame, text="驅動程式設定", padding=10)
        driver_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 提取輸出目錄
        ttk.Label(driver_frame, text="提取輸出目錄:").grid(row=0, column=0, sticky="w", pady=4)
        
        extract_path_frame = ttk.Frame(driver_frame)
        extract_path_frame.grid(row=0, column=1, sticky="ew", padx=(8, 0), pady=4)
        
        self.var_settings_extract_path = tk.StringVar()
        # 載入現有設定，預設為程式根目錄
        saved_extract_path = self._cfg_get('EXTRACT', 'output_path')
        if saved_extract_path:
            self.var_settings_extract_path.set(saved_extract_path.replace('/', '\\'))
        else:
            self.var_settings_extract_path.set(SCRIPT_DIR)
        
        extract_entry = ttk.Entry(extract_path_frame, textvariable=self.var_settings_extract_path, width=40)
        extract_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        def browse_extract_path():
            path = filedialog.askdirectory(title="選擇驅動程式提取輸出目錄")
            if path:
                self.var_settings_extract_path.set(path.replace('/', '\\'))
        
        ttk.Button(extract_path_frame, text="瀏覽...", command=browse_extract_path, width=8).pack(side=tk.LEFT, padx=(8, 0))
        
        # 匯出清單目錄
        ttk.Label(driver_frame, text="匯出清單目錄:").grid(row=1, column=0, sticky="w", pady=4)
        
        export_path_frame = ttk.Frame(driver_frame)
        export_path_frame.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=4)
        
        self.var_settings_export_path = tk.StringVar()
        # 載入現有設定，預設為程式根目錄
        saved_export_path = self._cfg_get('EXPORT', 'output_path')
        if saved_export_path:
            self.var_settings_export_path.set(saved_export_path.replace('/', '\\'))
        else:
            self.var_settings_export_path.set(SCRIPT_DIR)
        
        export_entry = ttk.Entry(export_path_frame, textvariable=self.var_settings_export_path, width=40)
        export_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        def browse_export_path():
            path = filedialog.askdirectory(title="選擇匯出清單輸出目錄")
            if path:
                self.var_settings_export_path.set(path.replace('/', '\\'))
        
        ttk.Button(export_path_frame, text="瀏覽...", command=browse_export_path, width=8).pack(side=tk.LEFT, padx=(8, 0))
        
        driver_frame.columnconfigure(1, weight=1)
        
        # === 按鈕區 ===
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        def save_settings():
            # 儲存提取輸出目錄
            if not self.cfg.has_section('EXTRACT'):
                self.cfg.add_section('EXTRACT')
            self.cfg.set('EXTRACT', 'output_path', self.var_settings_extract_path.get().strip())
            
            # 儲存匯出清單目錄
            if not self.cfg.has_section('EXPORT'):
                self.cfg.add_section('EXPORT')
            self.cfg.set('EXPORT', 'output_path', self.var_settings_export_path.get().strip())
            
            # 同步到主程式的變數
            if hasattr(self, 'var_extract_output'):
                self.var_extract_output.set(self.var_settings_extract_path.get().strip())
            
            # 儲存到檔案
            self._save_config()
            self._log("✓ 設定已儲存")
            messagebox.showinfo("儲存成功", "設定已儲存")
            dialog.destroy()
        
        def cancel_settings():
            dialog.destroy()
        
        ttk.Button(btn_frame, text="儲存", command=save_settings, width=10).pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(btn_frame, text="取消", command=cancel_settings, width=10).pack(side=tk.RIGHT)

    # ========== 設定檔 ==========
    def _load_config(self):
        try:
            if os.path.exists(CONFIG_FILE):
                self.cfg.read(CONFIG_FILE, encoding='utf-8')
        except Exception:
            pass

    def _cfg_get(self, section: str, option: str):
        if self.cfg.has_section(section) and self.cfg.has_option(section, option):
            return self.cfg.get(section, option)
        return None

    def _save_config(self):
        try:
            # WIM 設定
            if not self.cfg.has_section('WIM'):
                self.cfg.add_section('WIM')
            self.cfg.set('WIM', 'wim_file', self.var_wim.get().strip() if hasattr(self, 'var_wim') else '')
            self.cfg.set('WIM', 'mount_dir', self.var_mount_dir.get().strip() if hasattr(self, 'var_mount_dir') else '')
            self.cfg.set('WIM', 'index', self.var_wim_index.get().strip() if hasattr(self, 'var_wim_index') else '')
            self.cfg.set('WIM', 'readonly', '1' if (hasattr(self, 'var_wim_readonly') and self.var_wim_readonly.get()) else '0')
            self.cfg.set('WIM', 'unmount_commit', '1' if (hasattr(self, 'var_unmount_commit') and self.var_unmount_commit.get()) else '0')
            
            # WIM #2 設定
            if not self.cfg.has_section('WIM2'):
                self.cfg.add_section('WIM2')
            self.cfg.set('WIM2', 'wim_file', self.var_wim2.get().strip() if hasattr(self, 'var_wim2') else '')
            self.cfg.set('WIM2', 'mount_dir', self.var_mount_dir2.get().strip() if hasattr(self, 'var_mount_dir2') else '')
            self.cfg.set('WIM2', 'index', self.var_wim_index2.get().strip() if hasattr(self, 'var_wim_index2') else '')
            self.cfg.set('WIM2', 'readonly', '1' if (hasattr(self, 'var_wim_readonly2') and self.var_wim_readonly2.get()) else '0')
            self.cfg.set('WIM2', 'unmount_commit', '1' if (hasattr(self, 'var_unmount_commit2') and self.var_unmount_commit2.get()) else '0')
            
            # Driver 設定
            if not self.cfg.has_section('DRIVER'):
                self.cfg.add_section('DRIVER')
            self.cfg.set('DRIVER', 'mount_dir', self.var_driver_mount_dir.get().strip() if hasattr(self, 'var_driver_mount_dir') else '')
            self.cfg.set('DRIVER', 'source_path', self.var_driver_source.get().strip() if hasattr(self, 'var_driver_source') else '')
            self.cfg.set('DRIVER', 'recurse', '1' if (hasattr(self, 'var_driver_recurse') and self.var_driver_recurse.get()) else '0')
            self.cfg.set('DRIVER', 'force_unsigned', '1' if (hasattr(self, 'var_driver_force_unsigned') and self.var_driver_force_unsigned.get()) else '0')
            
            # Extract 設定
            if not self.cfg.has_section('EXTRACT'):
                self.cfg.add_section('EXTRACT')
            self.cfg.set('EXTRACT', 'source_path', self.var_extract_source.get().strip() if hasattr(self, 'var_extract_source') else '')
            self.cfg.set('EXTRACT', 'output_path', self.var_extract_output.get().strip() if hasattr(self, 'var_extract_output') else '')
            
            # Driver List 設定
            if not self.cfg.has_section('DRIVER_LIST'):
                self.cfg.add_section('DRIVER_LIST')
            self.cfg.set('DRIVER_LIST', 'mount_dir', self.var_driver_list_mount_dir.get().strip() if hasattr(self, 'var_driver_list_mount_dir') else '')
            
            # 設定檔直接放在程式同層，不需要建立額外資料夾
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                self.cfg.write(f)
        except Exception:
            pass


def main():
    app = App()
    app.mainloop()


if __name__ == '__main__':
    main()
