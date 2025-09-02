#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Windows 網路磁碟 Mount/Unmount GUI 工具（tkinter）
- 特色：
  1) 以 Win32 API (mpr.dll) 直接連線/中斷，不依賴 net use 指令
  2) 支援選擇磁碟機代號、設定帳密、是否開機自動還原 (Persistent)
  3) 一鍵列出目前所有對應的網路磁碟，並可直接開啟檔案總管
  4) 連線前快速 Ping 主機偵測
  5) 背景執行緒避免 GUI 卡頓，錯誤訊息人性化

需求：
- Windows 10/11、Python 3.9+
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

## 移除網路磁碟相依（WIM 功能不需要 Win32 mpr/k32）

# 設定檔路徑（儲存最近使用的路徑/選項）
# 自動判斷 .py 或 .exe 模式，將設定檔放在執行檔同層
if getattr(sys, 'frozen', False):
    # 打包成 .exe 時
    SCRIPT_DIR = os.path.dirname(sys.executable)
else:
    # .py 腳本模式
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, 'settings.ini')

## 只保留 WIM 掛載

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

# -----------------------------
# GUI 層
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("WIM 掛載工具 - Rex 版")
        self.geometry("720x520")
        self.minsize(700, 500)
        
        # 檢查並自動提升管理員權限
        if not WIMManager.is_admin():
            self._elevate_and_exit()
            return
            
        # 設定檔
        self.cfg = configparser.ConfigParser()
        self._load_config()
        self._build_ui()
        self._log("應用程式已啟動 (管理員權限)")

    # UI 組件
    def _build_ui(self):
        frm = ttk.Frame(self, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        # 只有 WIM 掛載頁
        self._build_wim_tab(frm)

        # Log 視窗（共用）
        lab = ttk.Label(frm, text="狀態 / 訊息")
        lab.pack(anchor=tk.W, pady=(8, 4))
        self.txt = tk.Text(frm, height=16, wrap=tk.WORD)
        self.txt.pack(fill=tk.BOTH, expand=True)
        self.txt.configure(state=tk.DISABLED)

    # WIM 掛載
    def _build_wim_tab(self, parent: tk.Misc):
        # 行 1：選擇 WIM 檔
        row1 = ttk.Frame(parent)
        row1.pack(fill=tk.X, pady=(8, 6))
        ttk.Label(row1, text="WIM 檔案").pack(side=tk.LEFT)
        self.var_wim = tk.StringVar()
        ent_wim = ttk.Entry(row1, textvariable=self.var_wim, width=50)
        ent_wim.pack(side=tk.LEFT, padx=8, fill=tk.X, expand=True)
        ttk.Button(row1, text="瀏覽…", command=self._on_browse_wim).pack(side=tk.LEFT)
        ttk.Button(row1, text="讀取映像資訊", command=self._on_wim_info).pack(side=tk.LEFT, padx=8)

        # 行 2：Index / ReadOnly
        row2 = ttk.Frame(parent)
        row2.pack(fill=tk.X, pady=6)
        ttk.Label(row2, text="Index").pack(side=tk.LEFT)
        self.var_wim_index = tk.StringVar()
        self.cbo_wim_index = ttk.Combobox(row2, textvariable=self.var_wim_index, width=6, state="readonly")
        self.cbo_wim_index.pack(side=tk.LEFT, padx=(6, 12))
        self.cbo_wim_index.bind('<<ComboboxSelected>>', lambda e: self._save_config())

        self.var_wim_readonly = tk.BooleanVar(value=True)
        ttk.Checkbutton(row2, text="唯讀掛載 (ReadOnly)", variable=self.var_wim_readonly, command=self._save_config).pack(side=tk.LEFT)

        # 行 3：掛載資料夾
        row3 = ttk.Frame(parent)
        row3.pack(fill=tk.X, pady=6)
        ttk.Label(row3, text="掛載資料夾").pack(side=tk.LEFT)
        self.var_mount_dir = tk.StringVar()
        ent_mdir = ttk.Entry(row3, textvariable=self.var_mount_dir, width=50)
        ent_mdir.pack(side=tk.LEFT, padx=8, fill=tk.X, expand=True)
        ttk.Button(row3, text="選擇…", command=self._on_browse_mount_dir).pack(side=tk.LEFT)
        ttk.Button(row3, text="建立", command=self._on_create_mount_dir).pack(side=tk.LEFT, padx=3)
        ttk.Button(row3, text="開啟", command=self._on_open_mount_dir).pack(side=tk.LEFT, padx=3)

        # 行 4：卸載選項
        row4 = ttk.Frame(parent)
        row4.pack(fill=tk.X, pady=6)
        ttk.Label(row4, text="卸載模式").pack(side=tk.LEFT)
        self.var_unmount_commit = tk.BooleanVar(value=False)
        ttk.Radiobutton(row4, text="丟棄變更 (/Discard)", variable=self.var_unmount_commit, value=False, command=self._save_config).pack(side=tk.LEFT, padx=(8, 12))
        ttk.Radiobutton(row4, text="提交變更 (/Commit)", variable=self.var_unmount_commit, value=True, command=self._save_config).pack(side=tk.LEFT)

        # 行 5：動作按鈕
        row5 = ttk.Frame(parent)
        row5.pack(fill=tk.X, pady=6)
        ttk.Button(row5, text="掛載 WIM", command=self._on_wim_mount).pack(side=tk.LEFT)
        ttk.Button(row5, text="卸載 WIM", command=self._on_wim_unmount).pack(side=tk.LEFT, padx=8)

        # 載入設定值
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

    def _elevate_and_exit(self):
        """自動提升權限並退出當前程序"""
        import sys
        import ctypes
        try:
            print("檢測到非管理員權限，正在提升權限...")
            script = os.path.abspath(sys.argv[0])
            params = " ".join([f'"{p}"' if ' ' in p else p for p in sys.argv[1:]])
            r = ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{script}" {params}', None, 1)
            if r <= 32:
                print(f"提升權限失敗，錯誤代碼：{r}")
                messagebox.showerror("權限錯誤", "無法提升管理員權限，程式將退出")
            else:
                print("正在以管理員權限重新啟動...")
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
        def update_combo():
            self.cbo_wim_index['values'] = indices_only
            # 若尚未選擇，預設第一個
            if not self.var_wim_index.get() and indices_only:
                self.var_wim_index.set(indices_only[0])
                self._save_config()
                self._log(f"自動選擇第一個映像 Index：{indices_only[0]}")
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
        
        self._log("開始 WIM 掛載前檢查...")
        
        if not wim or not mdir:
            self._log("掛載檢查失敗：缺少 WIM 檔案或掛載資料夾")
            messagebox.showwarning("輸入不完整", "請選擇 WIM 與掛載資料夾")
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
            self._log("WIM 掛載成功！")
            self._log(f"掛載位置: {mdir}")
            messagebox.showinfo("掛載成功", f"WIM 已成功掛載到:\n{mdir}")
        else:
            self._log(f"WIM 掛載失敗: {msg}")
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

    def _do_wim_unmount(self, mdir: str, commit: bool):
        commit_text = "提交變更 (/Commit)" if commit else "丟棄變更 (/Discard)"
        self._log(f"正在卸載 WIM...")
        self._log(f"  掛載位置: {mdir}")
        self._log(f"  卸載模式: {commit_text}")
        
        ok, msg = WIMManager.unmount_wim(mdir, commit)
        if ok:
            self._log("WIM 卸載成功！")
            messagebox.showinfo("卸載成功", f"WIM 已成功卸載\n模式: {commit_text}")
        else:
            self._log(f"WIM 卸載失敗: {msg}")
            messagebox.showerror("卸載失敗", f"卸載失敗:\n{msg}")

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
            if not self.cfg.has_section('WIM'):
                self.cfg.add_section('WIM')
            self.cfg.set('WIM', 'wim_file', self.var_wim.get().strip() if hasattr(self, 'var_wim') else '')
            self.cfg.set('WIM', 'mount_dir', self.var_mount_dir.get().strip() if hasattr(self, 'var_mount_dir') else '')
            self.cfg.set('WIM', 'index', self.var_wim_index.get().strip() if hasattr(self, 'var_wim_index') else '')
            self.cfg.set('WIM', 'readonly', '1' if (hasattr(self, 'var_wim_readonly') and self.var_wim_readonly.get()) else '0')
            self.cfg.set('WIM', 'unmount_commit', '1' if (hasattr(self, 'var_unmount_commit') and self.var_unmount_commit.get()) else '0')
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
