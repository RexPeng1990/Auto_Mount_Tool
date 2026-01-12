# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WIM 槽位卡片 - 側邊導航版本
適合並排顯示的緊湊設計
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import os
from typing import Optional, Callable, Any

from ui.theme import theme_manager, Fonts
from ui.components import StatusBadge
from ui.widgets import ModernComboBox, UnmountWarningDialog
from app.wim_manager import WIMManager
from app.utils import open_directory


class WIMSlotCard(ctk.CTkFrame):
    """WIM 槽位卡片 - 並排版本"""
    
    def __init__(
        self,
        master: Any,
        slot_number: int,
        on_log: Optional[Callable[[str], None]] = None,
        on_status_change: Optional[Callable[[int, bool], None]] = None,
        get_other_slot_info: Optional[Callable] = None,
        **kwargs
    ):
        super().__init__(
            master,
            fg_color="#ffffff",
            corner_radius=12,
            border_width=1,
            border_color="#e1f0ff",
            **kwargs
        )
        
        self.slot_number = slot_number
        self._on_log = on_log or (lambda x: None)
        self._on_status_change = on_status_change
        self._get_other_slot_info = get_other_slot_info
        
        # 變數
        self.var_wim_path = ctk.StringVar()
        self.var_mount_dir = ctk.StringVar()
        self.var_index = ctk.StringVar()
        self.var_readonly = ctk.BooleanVar(value=True)
        self.var_commit = ctk.BooleanVar(value=False)
        self.var_status = ctk.StringVar(value="未掛載")
        
        # Index 選項
        self._index_options = []
        
        self._build_ui()
    
    def _build_ui(self):
        """建立 UI"""
        # === 標題列 ===
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(12, 8))
        
        ctk.CTkLabel(
            header,
            text=f"WIM 掛載點 #{self.slot_number}",
            font=Fonts.to_tuple(Fonts.TITLE_SMALL),
            text_color="#37474f"
        ).pack(side="left")
        
        # 狀態徽章
        self.status_badge = StatusBadge(header, "未掛載", "default")
        self.status_badge.pack(side="right")
        
        # === 內容區 ===
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        
        # WIM 檔案
        self._create_field_row(content, "WIM 檔案", self.var_wim_path, 
                              browse_cmd=self._on_browse_wim)
        
        # 掛載目錄
        self._create_field_row(content, "掛載目錄", self.var_mount_dir,
                              browse_cmd=self._on_browse_mount)
        
        # 選項列：儲存模式 + 唯讀
        options_frame = ctk.CTkFrame(content, fg_color="transparent")
        options_frame.pack(fill="x", pady=(0, 8))
        
        # 卸載模式（左側）
        ctk.CTkLabel(
            options_frame,
            text="卸載模式",
            font=Fonts.to_tuple(Fonts.LABEL),
            text_color="#78909c",
            width=90,
            anchor="w"
        ).pack(side="left")
        
        self.combo_unmount = ModernComboBox(
            options_frame,
            width=110,
            height=32,
            values=["放棄變更", "保留變更"]
        )
        self.combo_unmount.set("放棄變更")
        self.combo_unmount.pack(side="left", padx=(0, 16))
        
        # 唯讀開關（右側）
        ctk.CTkLabel(
            options_frame,
            text="唯讀",
            font=Fonts.to_tuple(Fonts.LABEL),
            text_color="#78909c"
        ).pack(side="left")
        
        self.switch_readonly = ctk.CTkSwitch(
            options_frame,
            text="",
            variable=self.var_readonly,
            width=40,
            height=20,
            progress_color="#64b5f6",
            button_color="#ffffff",
            button_hover_color="#e3f2fd",
            fg_color="#e0e0e0",
            command=self._on_readonly_changed
        )
        self.switch_readonly.pack(side="left", padx=(4, 0))
        
        # 初始化：唯讀模式下鎖定儲存模式
        self._on_readonly_changed()
        
        # Index 固定為 1
        self.var_index.set("1")
        
        # === 按鈕區 ===
        btn_frame = ctk.CTkFrame(content, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(4, 0))
        
        # 根據狀態顯示不同按鈕
        self.btn_mount = ctk.CTkButton(
            btn_frame,
            text="掛載 WIM",
            font=Fonts.to_tuple(Fonts.BODY),
            height=36,
            corner_radius=8,
            fg_color="#64b5f6",
            hover_color="#42a5f5",
            text_color="#ffffff",
            command=self._on_mount
        )
        self.btn_mount.pack(fill="x")
        
        # 卸載+修復按鈕容器（左右並排）
        self.unmount_fix_frame = ctk.CTkFrame(btn_frame, fg_color="transparent")
        
        self.btn_unmount = ctk.CTkButton(
            self.unmount_fix_frame,
            text="卸載 WIM",
            font=Fonts.to_tuple(Fonts.BODY),
            height=36,
            corner_radius=8,
            fg_color="#64b5f6",
            hover_color="#42a5f5",
            text_color="#ffffff",
            command=self._on_unmount
        )
        
        # 一鍵修復按鈕
        self.btn_fix = ctk.CTkButton(
            self.unmount_fix_frame,
            text="修復卸載",
            font=Fonts.to_tuple(Fonts.BODY),
            height=36,
            corner_radius=8,
            fg_color="#81c784",  # 綠色
            hover_color="#66bb6a",
            text_color="#ffffff",
            command=self._on_smart_fix
        )
        
        self._update_button_state()
        
        # 監聽欄位變更，即時更新按鈕狀態
        self.var_wim_path.trace_add("write", lambda *_: self._update_button_state())
        self.var_mount_dir.trace_add("write", lambda *_: self._update_button_state())
    
    def _create_field_row(self, parent, label: str, var: ctk.StringVar, 
                         browse_cmd: Optional[Callable] = None):
        """建立欄位行"""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", pady=(0, 8))
        
        ctk.CTkLabel(
            frame,
            text=label,
            font=Fonts.to_tuple(Fonts.LABEL),
            text_color="#78909c",
            width=90,
            anchor="w"
        ).pack(side="left")
        
        entry_frame = ctk.CTkFrame(
            frame, 
            fg_color="#ffffff", 
            corner_radius=6,
            border_width=1,
            border_color="#cce5ff",
            height=36
        )
        entry_frame.pack(side="left", fill="x", expand=True)
        entry_frame.pack_propagate(False)
        
        entry = ctk.CTkEntry(
            entry_frame,
            textvariable=var,
            font=Fonts.to_tuple(Fonts.BODY_SMALL),
            height=32,
            border_width=0,
            fg_color="transparent",
            text_color="#37474f"
        )
        entry.pack(side="left", fill="x", expand=True, padx=(8, 0))
        
        if browse_cmd:
            btn = ctk.CTkButton(
                entry_frame,
                text="📁",
                font=("Segoe UI Emoji", 12),
                width=32,
                height=28,
                corner_radius=4,
                fg_color="transparent",
                hover_color="#e0e0e0",
                text_color="#78909c",
                command=browse_cmd
            )
            btn.pack(side="right", padx=2, pady=2)
    
    def _on_browse_wim(self):
        """瀏覽 WIM 檔案"""
        path = filedialog.askopenfilename(
            title="選擇 WIM 檔案",
            filetypes=[("WIM 映像", "*.wim"), ("ESD 映像", "*.esd"), ("所有檔案", "*.*")]
        )
        if path:
            # 統一使用正斜線
            path = path.replace("\\", "/")
            self.var_wim_path.set(path)
            self._on_log(f"WIM#{self.slot_number} 選擇檔案: {path}")
            self._read_wim_info()
    
    def _on_browse_mount(self):
        """瀏覽掛載目錄"""
        path = filedialog.askdirectory(title="選擇掛載目錄")
        if path:
            # 統一使用正斜線
            path = path.replace("\\", "/")
            self.var_mount_dir.set(path)
            self._on_log(f"WIM#{self.slot_number} 選擇目錄: {path}")
            # 檢查此目錄是否已有掛載
            self._check_and_update_mount_status()
    
    def _read_wim_info(self):
        """讀取 WIM 資訊"""
        wim_path = self.var_wim_path.get().strip()
        if not wim_path or not os.path.isfile(wim_path):
            return
        
        def do_read():
            success, images, error = WIMManager.get_wim_images(wim_path)
            self.after(0, lambda: self._update_index_options(success, images, error))
        
        threading.Thread(target=do_read, daemon=True).start()
    
    def _update_index_options(self, success: bool, images: list, error: str):
        """更新 Index 選項（已簡化，Index 固定為 1）"""
        if success and images:
            self._index_options = images
            self._on_log(f"WIM#{self.slot_number} 讀取到 {len(images)} 個映像")
        else:
            self._on_log(f"WIM#{self.slot_number} 讀取失敗: {error}")
    
    def _on_mount(self):
        """掛載 WIM"""
        wim_path = self.var_wim_path.get().strip()
        mount_dir = self.var_mount_dir.get().strip()
        index = "1"  # Index 固定為 1
        readonly = self.var_readonly.get()
        
        if not wim_path:
            messagebox.showwarning("提示", "請選擇 WIM 檔案")
            return
        if not mount_dir:
            messagebox.showwarning("提示", "請選擇掛載目錄")
            return
        
        # 建立目錄
        os.makedirs(mount_dir, exist_ok=True)
        
        self._on_log(f"正在掛載 WIM#{self.slot_number}...")
        self.btn_mount.configure(state="disabled", text="掛載中...")
        
        def do_mount():
            success, message = WIMManager.mount_wim(wim_path, index, mount_dir, readonly)
            self.after(0, lambda: self._mount_complete(success, message))
        
        threading.Thread(target=do_mount, daemon=True).start()
    
    def _mount_complete(self, success: bool, message: str):
        """掛載完成"""
        self.btn_mount.configure(state="normal", text="掛載 WIM")
        
        if success:
            self._on_log(f"✓ WIM#{self.slot_number} 掛載成功")
            self._update_status("已掛載", "success")
        else:
            self._on_log(f"✗ WIM#{self.slot_number} 掛載失敗: {message}")
            messagebox.showerror("掛載失敗", message)
    
    def _on_unmount(self):
        """卸載 WIM"""
        mount_dir = self.var_mount_dir.get().strip()
        if not mount_dir:
            return
        
        # 顯示卸載警告對話框
        dialog = UnmountWarningDialog(self.winfo_toplevel(), mount_dir)
        if not dialog.confirmed:
            return
        
        commit = self.combo_unmount.get() == "保留變更"
        action = "保留變更" if commit else "放棄變更"
        
        self._on_log(f"正在卸載 WIM#{self.slot_number} ({action})...")
        self.btn_unmount.configure(state="disabled", text="卸載中...")
        
        def do_unmount():
            success, message = WIMManager.unmount_wim(mount_dir, commit)
            self.after(0, lambda: self._unmount_complete(success, message))
        
        threading.Thread(target=do_unmount, daemon=True).start()
    
    def _unmount_complete(self, success: bool, message: str):
        """卸載完成"""
        self.btn_unmount.configure(state="normal", text="卸載 WIM")
        
        if success:
            self._on_log(f"✓ WIM#{self.slot_number} 卸載成功")
            self._update_status("未掛載", "default")
        else:
            self._on_log(f"✗ WIM#{self.slot_number} 卸載失敗: {message}")
            messagebox.showerror("卸載失敗", message)
    
    def _on_smart_fix(self):
        """一鍵修復"""
        self._on_log(f"WIM#{self.slot_number} 開始智能修復...")
        self.btn_fix.configure(state="disabled", text="修復中...")
        
        def do_fix():
            success, message = WIMManager.smart_cleanup_and_fix()
            self.after(0, lambda: self._fix_complete(success, message))
        
        threading.Thread(target=do_fix, daemon=True).start()
    
    def _fix_complete(self, success: bool, message: str):
        """修復完成"""
        self.btn_fix.configure(state="normal", text="⚠ 一鍵修復")
        
        if success:
            self._on_log(f"✓ WIM#{self.slot_number} 修復完成")
            messagebox.showinfo("修復完成", message)
            self._check_status()
        else:
            self._on_log(f"✗ WIM#{self.slot_number} 修復失敗: {message}")
            messagebox.showerror("修復失敗", message)
    
    def _check_status(self):
        """檢查狀態"""
        mount_dir = self.var_mount_dir.get().strip()
        if not mount_dir:
            self._update_status("未掛載", "default")
            return
        
        is_mounted, wim_path, status = WIMManager.is_path_mounted(mount_dir)
        
        if is_mounted:
            self._update_status("已掛載", "success")
        elif status == "needs_remount":
            self._update_status("需重新掛載", "warning")
        else:
            self._update_status("未掛載", "default")
    
    def _update_status(self, text: str, status: str):
        """更新狀態"""
        self.var_status.set(text)
        self.status_badge.set_status(text, status)
        self._update_button_state()
        
        if self._on_status_change:
            self._on_status_change(self.slot_number, status == "success")
    
    def _on_readonly_changed(self):
        """唯讀開關變更時的處理"""
        is_readonly = self.var_readonly.get()
        
        if is_readonly:
            # 唯讀模式：強制選擇「放棄變更」並鎖定下拉選單
            self.combo_unmount.set("放棄變更")
            self.combo_unmount.configure(state="disabled")
        else:
            # 可寫模式：解鎖下拉選單
            self.combo_unmount.configure(state="normal")
    
    def _update_button_state(self):
        """更新按鈕狀態"""
        is_mounted = self.var_status.get() == "已掛載"
        
        # 檢查欄位是否有值
        wim_path = self.var_wim_path.get().strip()
        mount_dir = self.var_mount_dir.get().strip()
        has_required_fields = bool(wim_path and mount_dir)
        
        if is_mounted:
            self.btn_mount.pack_forget()
            self.unmount_fix_frame.pack(fill="x")
            self.btn_unmount.pack(side="left", fill="x", expand=True, padx=(0, 4))
            self.btn_fix.pack(side="right", padx=(4, 0))
        else:
            self.unmount_fix_frame.pack_forget()
            self.btn_unmount.pack_forget()
            self.btn_fix.pack_forget()
            self.btn_mount.pack(fill="x")
            
            # 根據欄位是否有值設定按鈕狀態
            if has_required_fields:
                self.btn_mount.configure(
                    state="normal",
                    fg_color="#64b5f6",
                    hover_color="#42a5f5",
                    text_color="#ffffff"
                )
            else:
                self.btn_mount.configure(
                    state="disabled",
                    fg_color="#d0d0d0",
                    hover_color="#d0d0d0",
                    text_color="#ffffff"
                )
    
    # === 公開方法 ===
    
    def get_config(self) -> dict:
        """取得設定"""
        return {
            "wim_path": self.var_wim_path.get().replace("\\", "/"),
            "mount_dir": self.var_mount_dir.get().replace("\\", "/"),
            "index": self.var_index.get(),
            "readonly": self.var_readonly.get()
        }
    
    def set_config(self, config: dict):
        """設定值"""
        if "wim_path" in config:
            # 統一使用正斜線
            path = config["wim_path"].replace("\\", "/") if config["wim_path"] else ""
            self.var_wim_path.set(path)
        if "mount_dir" in config:
            # 統一使用正斜線
            path = config["mount_dir"].replace("\\", "/") if config["mount_dir"] else ""
            self.var_mount_dir.set(path)
        if "index" in config:
            self.var_index.set(config["index"])
        if "readonly" in config:
            self.var_readonly.set(config["readonly"])
            # 設定後立即更新儲存模式狀態
            self._on_readonly_changed()
        
        # 設定後檢查掛載狀態
        self.after(100, self._check_and_update_mount_status)
    
    def get_mount_dir(self) -> str:
        """取得掛載目錄"""
        return self.var_mount_dir.get().strip().replace("\\", "/")
    
    def get_readonly(self) -> bool:
        """取得唯讀狀態"""
        return self.var_readonly.get()
    
    def is_mounted(self) -> bool:
        """是否已掛載"""
        return self.var_status.get() == "已掛載"
    
    def _check_and_update_mount_status(self):
        """檢查掛載目錄是否已有 WIM 掛載，並更新狀態"""
        mount_dir = self.var_mount_dir.get().strip()
        if not mount_dir:
            return
        
        def do_check():
            is_mounted, mount_info = WIMManager.check_mount_status(mount_dir)
            self.after(0, lambda: self._update_detected_mount_status(is_mounted, mount_info))
        
        threading.Thread(target=do_check, daemon=True).start()
    
    def _update_detected_mount_status(self, is_mounted: bool, mount_info: dict | None):
        """更新偵測到的掛載狀態"""
        if is_mounted and mount_info:
            # 偵測到已掛載
            wim_file = mount_info.get("image_file", "")
            index = mount_info.get("image_index", "")
            status = mount_info.get("status", "Ok")
            
            # 統一路徑斜線
            if wim_file:
                wim_file = wim_file.replace("\\", "/")
            
            # 判斷掛載狀態類型
            if status in ["Needs Remount", "Invalid", "Corrupted"]:
                # 異常掛載
                self._update_status("需重新掛載", "warning")
                self._on_log(f"⚠️ WIM#{self.slot_number} 偵測到異常掛載 - {os.path.basename(wim_file)} (Index {index}, 狀態: {status})")
            else:
                # 正常掛載
                self._update_status("已掛載", "success")
                self._on_log(f"WIM#{self.slot_number} 偵測到已掛載 - {os.path.basename(wim_file)} (Index {index}, 狀態正常)")
            
            # 如果 WIM 路徑為空，自動填入
            if not self.var_wim_path.get().strip() and wim_file:
                self.var_wim_path.set(wim_file)
            
            # 如果 Index 為空，自動填入
            if not self.var_index.get().strip() and index:
                self.var_index.set(index)
        else:
            # 未偵測到掛載，只有在之前顯示已掛載時才更新
            current_status = self.var_status.get()
            if current_status in ["已掛載", "需重新掛載"]:
                self._update_status("未掛載", "default")
                self._on_log(f"WIM#{self.slot_number} 掛載目錄無掛載")
