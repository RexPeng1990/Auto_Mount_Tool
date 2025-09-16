#!/usr/bin/env python3
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
class WIMManager:
    @staticmethod
    def _norm_path(p: str) -> str:
        try:
            return os.path.normpath(p)
        except Exception:
            return p
    @staticmethod
    def is_admin() -> bool:
        try:
            import ctypes
            return bool(ctypes.windll.shell32.IsUserAnAdmin())
        except Exception:
            return False

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
    def get_wim_images(wim_path: str) -> tuple[bool, list[dict], str]:
        # 優先使用 /Get-WimInfo
        w = WIMManager._norm_path(wim_path)
        rc, out, err = WIMManager._run_dism(["/Get-WimInfo", f"/WimFile:{w}"])
        if rc != 0:
            # 兼容舊參數 /Get-ImageInfo
            rc2, out2, err2 = WIMManager._run_dism(["/Get-ImageInfo", f"/ImageFile:{w}"])
            if rc2 != 0:
                return False, [], err or err2 or out2 or out
            out = out2
        images = WIMManager._parse_wiminfo(out)
        return True, images, ""

    @staticmethod
    def _parse_wiminfo(text: str) -> list[dict]:
        # 解析 DISM 輸出，擷取 Index/Name/Description
        imgs: list[dict] = []
        cur: dict | None = None
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
            m = re.match(r"Index\s*:\s*(\d+)", line, re.IGNORECASE)
            if m:
                if cur:
                    imgs.append(cur)
                cur = {"Index": int(m.group(1)), "Name": "", "Description": ""}
                continue
            if cur is not None:
                m = re.match(r"Name\s*:\s*(.*)", line, re.IGNORECASE)
                if m:
                    cur["Name"] = m.group(1).strip()
                    continue
                m = re.match(r"Description\s*:\s*(.*)", line, re.IGNORECASE)
                if m:
                    cur["Description"] = m.group(1).strip()
                    continue
        if cur:
            imgs.append(cur)
        return imgs

    @staticmethod
    def mount_wim(wim_path: str, index: int, mount_dir: str, readonly: bool) -> tuple[bool, str]:
        w = WIMManager._norm_path(wim_path)
        m = WIMManager._norm_path(mount_dir)
        args = [
            "/Mount-Image",
            f"/ImageFile:{w}",
            f"/Index:{index}",
            f"/MountDir:{m}",
        ]
        if readonly:
            args.append("/ReadOnly")
        rc, out, err = WIMManager._run_dism(args)
        if rc == 0:
            return True, "WIM 掛載完成"
        return False, err or out

    @staticmethod
    def unmount_wim(mount_dir: str, commit: bool = False) -> tuple[bool, str]:
        m = WIMManager._norm_path(mount_dir)
        args = [
            "/Unmount-Image",
            f"/MountDir:{m}",
            "/Commit" if commit else "/Discard",
        ]
        rc, out, err = WIMManager._run_dism(args)
        if rc == 0:
            return True, "WIM 卸載完成"
        return False, err or out
    
    @staticmethod
    def close_explorer_windows(target_path: str) -> tuple[bool, str]:
        """
        關閉指向特定路徑的檔案總管視窗
        """
        try:
            import ctypes
            from ctypes import wintypes
            
            # 正規化路徑
            target_path = os.path.normpath(target_path).lower()
            
            # 使用 tasklist 找到 explorer.exe 進程
            result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq explorer.exe', '/FO', 'CSV'], 
                                  capture_output=True, text=True)
            
            if result.returncode != 0:
                return False, "無法查詢 explorer 進程"
            
            # 嘗試使用 PowerShell 關閉特定資料夾的檔案總管視窗
            ps_script = f'''
$shell = New-Object -ComObject Shell.Application
$windows = $shell.Windows()
$closed = 0
foreach ($window in $windows) {{
    try {{
        $path = $window.LocationURL
        if ($path -like "*file:///*") {{
            $localPath = $window.Document.Folder.Self.Path
            if ($localPath -and $localPath.ToLower().StartsWith("{target_path}")) {{
                $window.Quit()
                $closed++
            }}
        }}
    }} catch {{
        # 忽略錯誤，繼續下一個視窗
    }}
}}
Write-Output "已關閉 $closed 個檔案總管視窗"
'''
            
            result = subprocess.run(['powershell', '-Command', ps_script], 
                                  capture_output=True, text=True)
            
            if result.returncode == 0:
                output = result.stdout.strip()
                return True, output
            else:
                return False, f"PowerShell 執行失敗: {result.stderr}"
                
        except Exception as e:
            return False, f"關閉檔案總管視窗時發生錯誤: {str(e)}"

# -----------------------------
# 工具層：Driver 離線安裝（使用 DISM）
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
        
        if recurse:
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
                for field in ["OriginalFileName", "ClassName", "Provider", "Date", "Version"]:
                    pattern = f"{field.replace('Name', ' Name')}\\s*:\\s*(.*)"
                    match = re.search(pattern, line, re.IGNORECASE)
                    if match:
                        cur[field] = match.group(1).strip()
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
        self.title("WIM/Driver 管理工具 - Rex 版")
        self.geometry("750x600")
        self.minsize(750, 580)
        
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

    # UI 組件
    def _build_ui(self):
        # 主容器
        main_frame = ttk.Frame(self, padding=8)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 建立 Notebook (分頁)
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        # 分頁 1：WIM 掛載（使用子分頁）
        wim_frame = ttk.Frame(self.notebook)
        self.notebook.add(wim_frame, text="WIM 掛載")
        
        # 在 WIM 掛載分頁中建立子分頁
        wim_sub_notebook = ttk.Notebook(wim_frame)
        wim_sub_notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        
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

    # WIM 掛載分頁
    def _build_wim1_tab(self, parent: tk.Misc):
        # 使用 padding 的 frame
        content_frame = ttk.Frame(parent, padding=12)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # 說明文字
        desc_frame = ttk.LabelFrame(content_frame, text="功能說明", padding=8)
        desc_frame.pack(fill=tk.X, pady=(0, 12))
        desc_text = "此功能用於掛載/卸載第一組 Windows 映像檔 (WIM)。\n操作流程：選擇 WIM 檔案 → 讀取映像資訊 → 設定掛載資料夾 → 掛載映像 → 修改檔案 → 卸載並提交變更。"
        ttk.Label(desc_frame, text=desc_text, wraplength=600).pack(anchor=tk.W)

        # WIM 掛載 #1
        wim1_frame = ttk.LabelFrame(content_frame, text="WIM 掛載 #1", padding=10)
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

        # 行 5：動作按鈕
        row5 = ttk.Frame(wim1_frame)
        row5.pack(fill=tk.X, pady=(0, 5))
        
        # WIM 操作按鈕組
        wim_action_frame = ttk.Frame(row5)
        wim_action_frame.pack(side=tk.LEFT)
        ttk.Button(wim_action_frame, text="掛載 WIM", command=self._on_wim_mount, width=12).pack(side=tk.LEFT)
        ttk.Button(wim_action_frame, text="卸載 WIM", command=self._on_wim_unmount, width=12).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(wim_action_frame, text="關閉檔案總管", command=self._on_close_explorer).pack(side=tk.LEFT, padx=(8, 0))

    def _build_wim2_tab(self, parent: tk.Misc):
        # 使用 padding 的 frame
        content_frame = ttk.Frame(parent, padding=12)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # 說明文字
        desc_frame = ttk.LabelFrame(content_frame, text="功能說明", padding=8)
        desc_frame.pack(fill=tk.X, pady=(0, 12))
        desc_text = "此功能用於掛載/卸載第二組 Windows 映像檔 (WIM)。\n操作流程：選擇 WIM 檔案 → 讀取映像資訊 → 設定掛載資料夾 → 掛載映像 → 修改檔案 → 卸載並提交變更。"
        ttk.Label(desc_frame, text=desc_text, wraplength=600).pack(anchor=tk.W)

        # WIM 掛載 #2
        wim2_frame = ttk.LabelFrame(content_frame, text="WIM 掛載 #2", padding=10)
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

        # 行 5：動作按鈕 #2
        row5_2 = ttk.Frame(wim2_frame)
        row5_2.pack(fill=tk.X, pady=(0, 5))
        
        # WIM 操作按鈕組 #2
        wim2_action_frame = ttk.Frame(row5_2)
        wim2_action_frame.pack(side=tk.LEFT)
        ttk.Button(wim2_action_frame, text="掛載 WIM", command=self._on_wim_mount2, width=12).pack(side=tk.LEFT)
        ttk.Button(wim2_action_frame, text="卸載 WIM", command=self._on_wim_unmount2, width=12).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(wim2_action_frame, text="關閉檔案總管", command=self._on_close_explorer2).pack(side=tk.LEFT, padx=(8, 0))

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

    # Driver 管理分頁（使用子分頁：萃取和安裝）
    def _build_driver_tab(self, parent: tk.Misc):
        # 使用 padding 的 frame
        content_frame = ttk.Frame(parent, padding=8)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # 說明文字
        desc_frame = ttk.LabelFrame(content_frame, text="功能說明", padding=8)
        desc_frame.pack(fill=tk.X, pady=(0, 8))
        desc_text = "此功能提供驅動程式的萃取與安裝。可以從一個映像萃取驅動，然後安裝到另一個映像，或直接安裝外部驅動程式。"
        ttk.Label(desc_frame, text=desc_text, wraplength=600).pack(anchor=tk.W)

        # 建立子分頁
        driver_sub_notebook = ttk.Notebook(content_frame)
        driver_sub_notebook.pack(fill=tk.BOTH, expand=True, pady=(8, 0))

        # 子分頁 1：驅動程式萃取
        extract_frame = ttk.Frame(driver_sub_notebook)
        driver_sub_notebook.add(extract_frame, text="驅動萃取")
        self._build_extract_subtab(extract_frame)

        # 子分頁 2：驅動程式安裝
        install_frame = ttk.Frame(driver_sub_notebook)
        driver_sub_notebook.add(install_frame, text="驅動安裝")
        self._build_install_subtab(install_frame)

        # 載入設定
        self._load_driver_config()

    def _build_extract_subtab(self, parent: tk.Misc):
        # 使用 padding 的 frame
        content_frame = ttk.Frame(parent, padding=12)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # 萃取說明
        desc_frame = ttk.LabelFrame(content_frame, text="萃取說明", padding=8)
        desc_frame.pack(fill=tk.X, pady=(0, 12))
        desc_text = "從已掛載的 Windows 映像中萃取所有驅動程式到指定目錄。\n萃取完成後可在「驅動安裝」分頁中使用這些驅動程式。"
        ttk.Label(desc_frame, text=desc_text, wraplength=600).pack(anchor=tk.W)

        # 來源映像路徑
        row1 = ttk.Frame(content_frame)
        row1.pack(fill=tk.X, pady=(0, 12))
        ttk.Label(row1, text="來源映像路徑", width=14).pack(side=tk.LEFT)
        self.var_extract_source = tk.StringVar()
        ent_extract_source = ttk.Entry(row1, textvariable=self.var_extract_source, width=40)
        ent_extract_source.pack(side=tk.LEFT, padx=(8, 6), fill=tk.X, expand=True)
        
        # 來源同步按鈕組
        source_sync_frame = ttk.Frame(row1)
        source_sync_frame.pack(side=tk.RIGHT)
        ttk.Button(source_sync_frame, text="選擇…", command=self._on_browse_extract_source).pack(side=tk.LEFT)
        ttk.Button(source_sync_frame, text="從 WIM#1", command=self._on_sync_extract_from_wim1).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(source_sync_frame, text="從 WIM#2", command=self._on_sync_extract_from_wim2).pack(side=tk.LEFT, padx=(6, 0))

        # 萃取輸出目錄
        row2 = ttk.Frame(content_frame)
        row2.pack(fill=tk.X, pady=(0, 12))
        ttk.Label(row2, text="驅動萃取目錄", width=14).pack(side=tk.LEFT)
        self.var_extract_output = tk.StringVar()
        ent_extract_output = ttk.Entry(row2, textvariable=self.var_extract_output, width=40)
        ent_extract_output.pack(side=tk.LEFT, padx=(8, 6), fill=tk.X, expand=True)
        
        # 萃取目錄按鈕組
        output_btn_frame = ttk.Frame(row2)
        output_btn_frame.pack(side=tk.RIGHT)
        ttk.Button(output_btn_frame, text="選擇…", command=self._on_browse_extract_output).pack(side=tk.LEFT)
        ttk.Button(output_btn_frame, text="建立", command=self._on_create_extract_dir).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(output_btn_frame, text="開啟", command=self._on_open_extract_dir).pack(side=tk.LEFT, padx=(6, 0))

        # 萃取操作按鈕
        row3 = ttk.Frame(content_frame)
        row3.pack(fill=tk.X, pady=(0, 8))
        extract_action_frame = ttk.Frame(row3)
        extract_action_frame.pack(side=tk.LEFT)
        ttk.Button(extract_action_frame, text="萃取驅動程式", command=self._on_extract_drivers, width=15).pack(side=tk.LEFT)
        ttk.Button(extract_action_frame, text="查看萃取結果", command=self._on_view_extracted_drivers).pack(side=tk.LEFT, padx=(10, 0))

    def _build_install_subtab(self, parent: tk.Misc):
        # 使用 padding 的 frame
        content_frame = ttk.Frame(parent, padding=12)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # 安裝說明
        desc_frame = ttk.LabelFrame(content_frame, text="安裝說明", padding=8)
        desc_frame.pack(fill=tk.X, pady=(0, 12))
        desc_text = "將驅動程式離線安裝到已掛載的 Windows 映像中。\n驅動來源可以是萃取的結果、外部驅動資料夾或單一 .inf 檔案。"
        ttk.Label(desc_frame, text=desc_text, wraplength=600).pack(anchor=tk.W)

        # 目標映像路徑
        row1 = ttk.Frame(content_frame)
        row1.pack(fill=tk.X, pady=(0, 12))
        ttk.Label(row1, text="目標映像路徑", width=14).pack(side=tk.LEFT)
        self.var_driver_mount_dir = tk.StringVar()
        ent_driver_mount = ttk.Entry(row1, textvariable=self.var_driver_mount_dir, width=40)
        ent_driver_mount.pack(side=tk.LEFT, padx=(8, 6), fill=tk.X, expand=True)
        
        # 目標同步按鈕組
        target_sync_frame = ttk.Frame(row1)
        target_sync_frame.pack(side=tk.RIGHT)
        ttk.Button(target_sync_frame, text="選擇…", command=self._on_browse_driver_mount_dir).pack(side=tk.LEFT)
        ttk.Button(target_sync_frame, text="從 WIM#1", command=self._on_sync_from_wim1).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(target_sync_frame, text="從 WIM#2", command=self._on_sync_from_wim2).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(target_sync_frame, text="檢查掛載", command=self._on_check_mount_status).pack(side=tk.LEFT, padx=(6, 0))

        # 驅動程式來源
        row2 = ttk.Frame(content_frame)
        row2.pack(fill=tk.X, pady=(0, 12))
        ttk.Label(row2, text="驅動程式來源", width=14).pack(side=tk.LEFT)
        self.var_driver_source = tk.StringVar()
        ent_driver_source = ttk.Entry(row2, textvariable=self.var_driver_source, width=40)
        ent_driver_source.pack(side=tk.LEFT, padx=(8, 6), fill=tk.X, expand=True)
        
        # 驅動程式選擇按鈕組
        driver_btn_frame = ttk.Frame(row2)
        driver_btn_frame.pack(side=tk.RIGHT)
        ttk.Button(driver_btn_frame, text="選擇 .inf 檔…", command=self._on_browse_driver_file).pack(side=tk.LEFT)
        ttk.Button(driver_btn_frame, text="選擇資料夾…", command=self._on_browse_driver_source).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(driver_btn_frame, text="使用萃取結果", command=self._on_use_extracted_drivers).pack(side=tk.LEFT, padx=(6, 0))

        # 安裝選項
        row3 = ttk.Frame(content_frame)
        row3.pack(fill=tk.X, pady=(0, 12))
        ttk.Label(row3, text="安裝選項", width=14).pack(side=tk.LEFT)
        
        # 安裝選項框
        options_frame = ttk.Frame(row3)
        options_frame.pack(side=tk.LEFT, padx=(8, 0))
        self.var_driver_recurse = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="遞迴搜尋子資料夾 (/Recurse)", variable=self.var_driver_recurse, command=self._save_config).pack(side=tk.LEFT)
        
        self.var_driver_force_unsigned = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="強制未簽署驱動 (/ForceUnsigned)", variable=self.var_driver_force_unsigned, command=self._save_config).pack(side=tk.LEFT, padx=(20, 0))

        # 安裝操作按鈕
        row4 = ttk.Frame(content_frame)
        row4.pack(fill=tk.X, pady=(0, 8))
        
        # 驅動操作按鈕組
        driver_action_frame = ttk.Frame(row4)
        driver_action_frame.pack(side=tk.LEFT)
        ttk.Button(driver_action_frame, text="安裝驅動程式", command=self._on_install_driver, width=15).pack(side=tk.LEFT)
        ttk.Button(driver_action_frame, text="列出已安裝驅動", command=self._on_list_drivers, width=15).pack(side=tk.LEFT, padx=(10, 0))

    def _load_driver_config(self):
        """載入驅動程式相關設定"""
        # 載入安裝設定
        driver_mount = self._cfg_get('DRIVER', 'mount_dir')
        if driver_mount:
            self.var_driver_mount_dir.set(driver_mount)
        else:
            # 如果沒有設定且 WIM 路徑已設定，則自動同步
            wim_mount = self._cfg_get('WIM', 'mount_dir')
            if wim_mount:
                self.var_driver_mount_dir.set(wim_mount)
                
        driver_source = self._cfg_get('DRIVER', 'source_path')
        if driver_source:
            self.var_driver_source.set(driver_source)
        recurse = self._cfg_get('DRIVER', 'recurse')
        if recurse is not None:
            self.var_driver_recurse.set(recurse.lower() in ('1', 'true', 'yes', 'on'))
        force_unsigned = self._cfg_get('DRIVER', 'force_unsigned')
        if force_unsigned is not None:
            self.var_driver_force_unsigned.set(force_unsigned.lower() in ('1', 'true', 'yes', 'on'))
        
        # 載入萃取設定
        extract_source = self._cfg_get('EXTRACT', 'source_path')
        if extract_source:
            self.var_extract_source.set(extract_source)
        
        extract_output = self._cfg_get('EXTRACT', 'output_path') 
        if extract_output:
            self.var_extract_output.set(extract_output)

    # 工具方法
    def _log(self, msg: str):
        ts = datetime.now().strftime('%H:%M:%S')
        self.txt.configure(state=tk.NORMAL)
        self.txt.insert(tk.END, f"[{ts}] {msg}\n")
        self.txt.see(tk.END)
        self.txt.configure(state=tk.DISABLED)

    def _thread(self, target, *args):
        t = threading.Thread(target=target, args=args, daemon=True)
        t.start()

    def _on_mount_dir_changed(self, *args):
        """當 WIM 掛載路徑變更時自動同步到 Driver 分頁"""
        if hasattr(self, 'var_driver_mount_dir') and hasattr(self, 'var_mount_dir'):
            wim_path = self.var_mount_dir.get().strip()
            current_driver_path = self.var_driver_mount_dir.get().strip()
            
            # 只有當 driver 路徑為空或與 wim 路徑不同時才同步
            if wim_path and (not current_driver_path or current_driver_path != wim_path):
                self.var_driver_mount_dir.set(wim_path)
                self._log(f"自動同步掛載路徑到 Driver 分頁: {wim_path}")

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
            messagebox.showerror("掛載失敗", f"掛載失敗:\n{msg}")

    def _on_wim_unmount(self):
        mdir = self.var_mount_dir.get().strip()
        commit = self.var_unmount_commit.get()
        
        if not mdir:
            self._log("卸載失敗：未指定掛載資料夾")
            messagebox.showwarning("輸入不完整", "請先指定掛載資料夾")
            return
            
        commit_text = "提交變更" if commit else "丟棄變更"
        self._log(f"準備卸載 WIM (模式: {commit_text})...")
        self._thread(self._do_wim_unmount, mdir, commit)

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
        
        ok, msg = WIMManager.mount_wim(wim, index, mdir, ro)
        if ok:
            self._log("✓ 第二個 WIM 掛載成功！")
            self._log(f"掛載位置: {mdir}")
            messagebox.showinfo("掛載成功", f"第二個 WIM 已成功掛載到:\n{mdir}")
        else:
            self._log(f"✗ 第二個 WIM 掛載失敗: {msg}")
            messagebox.showerror("掛載失敗", f"第二個掛載失敗:\n{msg}")

    def _on_wim_unmount2(self):
        mdir = self.var_mount_dir2.get().strip()
        commit = self.var_unmount_commit2.get()
        
        if not mdir:
            self._log("第二個卸載失敗：未指定掛載資料夾")
            messagebox.showwarning("輸入不完整", "請先指定第二個掛載資料夾")
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
            
        self._log("開始安裝驅動程式...")
        self._save_config()
        self._thread(self._do_install_driver, mount_dir, driver_source, recurse, force_unsigned)

    def _do_install_driver(self, mount_dir: str, driver_source: str, recurse: bool, force_unsigned: bool):
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
            messagebox.showinfo("安裝成功", "驅動程式已成功安裝到離線映像")
        else:
            self._log(f"✗ 驅動程式安裝失敗: {msg}")
            messagebox.showerror("安裝失敗", f"驅動程式安裝失敗:\n{msg}")

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
        self._log(f"正在查詢映像中的驅動程式: {mount_dir}")
        
        ok, drivers, err = DriverManager.get_drivers_in_offline_image(mount_dir)
        if not ok:
            self._log(f"查詢驅動程式失敗: {err}")
            messagebox.showerror("查詢失敗", f"無法查詢驅動程式:\n{err}")
            return
            
        if not drivers:
            self._log("映像中沒有找到已安裝的驅動程式")
            messagebox.showinfo("查詢結果", "映像中沒有找到已安裝的驅動程式")
            return
            
        self._log(f"找到 {len(drivers)} 個已安裝的驅動程式:")
        for i, driver in enumerate(drivers, 1):
            name = driver.get('PublishedName', 'N/A')
            provider = driver.get('Provider', 'N/A')
            version = driver.get('Version', 'N/A')
            date = driver.get('Date', 'N/A')
            self._log(f"  {i:2d}. {name} - {provider} (v{version}, {date})")
            
        messagebox.showinfo("查詢結果", f"找到 {len(drivers)} 個已安裝的驅動程式\n詳細資訊請查看日誌")

    # ---------- Extract 事件 ----------

    def _on_browse_extract_source(self):
        path = filedialog.askdirectory(title="選擇來源映像掛載目錄")
        if path:
            self.var_extract_source.set(path)
            self._log(f"已選擇來源映像路徑：{path}")
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
            messagebox.showwarning("輸入不完整", "請選擇來源映像路徑和萃取輸出目錄")
            return
            
        if not os.path.exists(source_path):
            messagebox.showerror("路徑錯誤", "來源映像路徑不存在")
            return
            
        self._log("開始萃取驅動程式...")
        self._save_config()
        self._thread(self._do_extract_drivers, source_path, output_path)

    def _do_extract_drivers(self, source_path: str, output_path: str):
        self._log(f"正在從映像萃取驅動程式...")
        self._log(f"  來源映像: {source_path}")
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

    # 設定檔
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
