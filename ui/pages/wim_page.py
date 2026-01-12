# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WIM 掛載頁面
現代化 UI 設計
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Optional, Callable, Any
import os
import threading

from ui.theme import ThemeManager, Fonts, Spacing, theme_manager
from ui.components import (
    ModernButton, ModernCard, ModernEntry, ModernTooltip, 
    StatusBadge, FormField, SectionTitle, IconButton
)
from app.wim_manager import WIMManager
from app.utils import create_mount_directory, open_directory


class WIMSlot(ctk.CTkFrame):
    """
    單一 WIM 掛載槽位組件
    封裝所有單個 WIM 掛載相關的 UI 和邏輯
    """
    
    def __init__(
        self,
        master: Any,
        slot_number: int,
        on_log: Callable[[str], None],
        on_status_change: Optional[Callable] = None,
        get_other_slot_info: Optional[Callable] = None,
        expanded: bool = True,
        **kwargs
    ):
        """
        初始化 WIM 掛載槽位
        
        Args:
            master: 父組件
            slot_number: 槽位編號 (1 或 2)
            on_log: 日誌回調函數
            on_status_change: 狀態變更回調
            get_other_slot_info: 取得另一個槽位資訊的回調
            expanded: 是否預設展開
        """
        super().__init__(master, fg_color="transparent", **kwargs)
        
        self.slot_number = slot_number
        self._on_log = on_log
        self._on_status_change = on_status_change
        self._get_other_slot_info = get_other_slot_info
        self._expanded = expanded
        
        # 變數
        self.var_wim_path = ctk.StringVar()
        self.var_mount_dir = ctk.StringVar()
        self.var_index = ctk.StringVar()
        self.var_readonly = ctk.BooleanVar(value=True)
        self.var_commit = ctk.BooleanVar(value=False)
        self.var_status = ctk.StringVar(value="未掛載")
        
        # Index 選項
        self._index_options = []
        
        # 建立 UI
        self._build_ui()
    
    def _build_ui(self):
        """建立 UI - 可摺疊緊湊版"""
        # 主卡片
        self.card = ModernCard(self, padding=8)
        self.card.pack(fill="x", pady=(0, 6))
        content = self.card.get_content_frame()
        
        # === 頂部：可點擊的標題列（用於摺疊） ===
        header = ctk.CTkFrame(content, fg_color="transparent", cursor="hand2", height=28)
        header.pack(fill="x")
        header.pack_propagate(False)
        
        # 摺疊指示器
        self.collapse_icon = ctk.CTkLabel(
            header, text="▼" if self._expanded else "▶",
            font=Fonts.to_tuple(Fonts.CAPTION), width=16,
            text_color=theme_manager.colors.text_secondary
        )
        self.collapse_icon.pack(side="left")
        
        # 標題
        self.title_label = ctk.CTkLabel(
            header,
            text=f"🗂️ WIM 掛載 #{self.slot_number}",
            font=Fonts.to_tuple(Fonts.BODY),
            text_color=theme_manager.colors.text_primary
        )
        self.title_label.pack(side="left", padx=(2, 0))
        
        # 狀態徽章
        self.status_badge = StatusBadge(header, "未掛載", "default")
        self.status_badge.pack(side="right")
        
        # 綁定點擊事件到 header
        for widget in [header, self.collapse_icon, self.title_label]:
            widget.bind("<Button-1>", lambda e: self._toggle_collapse())
        
        # === 可摺疊內容區 ===
        self.content_frame = ctk.CTkFrame(content, fg_color="transparent")
        if self._expanded:
            self.content_frame.pack(fill="x", pady=(6, 0))
        
        # === 第一行：WIM 檔案 + 讀取資訊 ===
        row1 = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        row1.pack(fill="x", pady=(0, 4))
        
        ctk.CTkLabel(
            row1, text="WIM 檔案", font=Fonts.to_tuple(Fonts.LABEL),
            text_color=theme_manager.colors.text_secondary, width=70, anchor="w"
        ).pack(side="left")
        
        self.entry_wim = ModernEntry(row1, placeholder="選擇 WIM 映像檔...", width=280)
        self.entry_wim.pack(side="left", padx=(4, 0), fill="x", expand=True)
        self.entry_wim.configure(textvariable=self.var_wim_path)
        
        ModernButton(row1, text="瀏覽", variant="outline", size="sm",
                     command=self._on_browse_wim).pack(side="left", padx=(4, 0))
        ModernButton(row1, text="讀取資訊", variant="secondary", size="sm",
                     command=self._on_read_wim_info).pack(side="left", padx=(4, 0))
        
        # === 第二行：掛載目錄 + 建立/開啟 ===
        row2 = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        row2.pack(fill="x", pady=(0, 4))
        
        ctk.CTkLabel(
            row2, text="掛載目錄", font=Fonts.to_tuple(Fonts.LABEL),
            text_color=theme_manager.colors.text_secondary, width=70, anchor="w"
        ).pack(side="left")
        
        self.entry_mount = ModernEntry(row2, placeholder="選擇掛載目錄...", width=280)
        self.entry_mount.pack(side="left", padx=(4, 0), fill="x", expand=True)
        self.entry_mount.configure(textvariable=self.var_mount_dir)
        
        ModernButton(row2, text="選擇", variant="outline", size="sm",
                     command=self._on_browse_mount_dir).pack(side="left", padx=(4, 0))
        ModernButton(row2, text="建立", variant="outline", size="sm",
                     command=self._on_create_mount_dir).pack(side="left", padx=(4, 0))
        ModernButton(row2, text="開啟", variant="ghost", size="sm",
                     command=self._on_open_mount_dir).pack(side="left", padx=(4, 0))
        
        # === 第三行：Index + 唯讀 + 卸載模式 ===
        row3 = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        row3.pack(fill="x", pady=(0, 6))
        
        # Index
        ctk.CTkLabel(
            row3, text="Index", font=Fonts.to_tuple(Fonts.LABEL),
            text_color=theme_manager.colors.text_secondary, width=70, anchor="w"
        ).pack(side="left")
        
        self.combo_index = ctk.CTkComboBox(
            row3, width=80, variable=self.var_index, state="readonly", values=[""],
            font=Fonts.to_tuple(Fonts.BODY), dropdown_font=Fonts.to_tuple(Fonts.BODY),
            command=self._on_index_changed
        )
        self.combo_index.pack(side="left", padx=(4, 0))
        
        # 唯讀開關
        self.switch_readonly = ctk.CTkSwitch(
            row3, text="唯讀", variable=self.var_readonly, width=60,
            font=Fonts.to_tuple(Fonts.BODY), text_color=theme_manager.colors.text_primary,
            progress_color=theme_manager.colors.primary, command=self._on_readonly_changed
        )
        self.switch_readonly.pack(side="left", padx=(16, 0))
        ModernTooltip(self.switch_readonly, "唯讀模式掛載，無法修改映像內容")
        
        # 分隔
        ctk.CTkLabel(row3, text="│", text_color=theme_manager.colors.border).pack(side="left", padx=(12, 12))
        
        # 卸載模式
        ctk.CTkLabel(
            row3, text="卸載:", font=Fonts.to_tuple(Fonts.LABEL),
            text_color=theme_manager.colors.text_secondary
        ).pack(side="left")
        
        self.radio_discard = ctk.CTkRadioButton(
            row3, text="丟棄", variable=self.var_commit, value=False,
            font=Fonts.to_tuple(Fonts.BODY), text_color=theme_manager.colors.text_primary, width=60
        )
        self.radio_discard.pack(side="left", padx=(4, 0))
        
        self.radio_commit = ctk.CTkRadioButton(
            row3, text="提交", variable=self.var_commit, value=True,
            font=Fonts.to_tuple(Fonts.BODY), text_color=theme_manager.colors.text_primary, width=60
        )
        self.radio_commit.pack(side="left", padx=(4, 0))
        
        # 檢查狀態按鈕（右側）
        ModernButton(row3, text="檢查狀態", variant="ghost", size="sm",
                     command=self._on_check_status).pack(side="right")
        
        # === 第四行：操作按鈕 ===
        btn_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        btn_frame.pack(fill="x")
        
        self.btn_mount = ModernButton(
            btn_frame, text="掛載 WIM", icon="📁", variant="primary", command=self._on_mount
        )
        self.btn_mount.pack(side="left")
        
        self.btn_unmount = ModernButton(
            btn_frame, text="卸載 WIM", icon="📤", variant="secondary", command=self._on_unmount
        )
        self.btn_unmount.pack(side="left", padx=(8, 0))
        
        ModernButton(btn_frame, text="關閉檔案總管", variant="outline", size="sm",
                     command=self._on_close_explorer).pack(side="left", padx=(8, 0))
        
        ModernButton(btn_frame, text="🔧 一鍵修復", variant="warning", size="sm",
                     command=self._on_smart_fix).pack(side="right")
    
    def _toggle_collapse(self):
        """切換摺疊狀態"""
        self._expanded = not self._expanded
        
        if self._expanded:
            self.collapse_icon.configure(text="▼")
            self.content_frame.pack(fill="x", pady=(6, 0))
        else:
            self.collapse_icon.configure(text="▶")
            self.content_frame.pack_forget()
    
    def expand(self):
        """展開面板"""
        if not self._expanded:
            self._toggle_collapse()
    
    def collapse(self):
        """摺疊面板"""
        if self._expanded:
            self._toggle_collapse()
    
    # === 事件處理 ===
    
    def _on_browse_wim(self):
        """瀏覽 WIM 檔案"""
        path = filedialog.askopenfilename(
            title="選擇 WIM 檔案",
            filetypes=[("WIM 映像", "*.wim"), ("ESD 映像", "*.esd"), ("所有檔案", "*.*")]
        )
        if path:
            self.var_wim_path.set(path)
            self._on_log(f"WIM#{self.slot_number} 選擇檔案: {path}")
            # 自動讀取資訊
            self._on_read_wim_info()
    
    def _on_read_wim_info(self):
        """讀取 WIM 資訊"""
        wim_path = self.var_wim_path.get().strip()
        if not wim_path:
            messagebox.showwarning("提示", "請先選擇 WIM 檔案")
            return
        
        if not os.path.isfile(wim_path):
            messagebox.showerror("錯誤", f"檔案不存在: {wim_path}")
            return
        
        self._on_log(f"正在讀取 WIM 資訊: {wim_path}")
        
        def do_read():
            success, images, error = WIMManager.get_wim_images(wim_path)
            result = {"success": success, "images": images, "error": error}
            
            # 取得另一個槽位使用的 Index（如果是相同檔案）
            exclude_indices = []
            if self._get_other_slot_info:
                other_info = self._get_other_slot_info()
                if other_info:
                    # 支援兩種格式：tuple (wim_path, index) 或 dict {wim_file, index, ...}
                    if isinstance(other_info, dict):
                        other_path = other_info.get("wim_file", "")
                        other_index = other_info.get("index", "")
                    else:
                        other_path, other_index = other_info
                    # 檢查是否為同一檔案
                    if self._is_same_file(wim_path, other_path) and other_index:
                        exclude_indices.append(other_index)
            
            self.after(0, lambda: self._update_index_options(result, exclude_indices))
        
        threading.Thread(target=do_read, daemon=True).start()
    
    def _is_same_file(self, path1: str, path2: str) -> bool:
        """檢查是否為同一檔案"""
        if not path1 or not path2:
            return False
        try:
            return os.path.normpath(os.path.abspath(path1)).lower() == \
                   os.path.normpath(os.path.abspath(path2)).lower()
        except Exception:
            return False
    
    def _update_index_options(self, result: dict, exclude_indices: list = None):
        """更新 Index 下拉選項"""
        if result.get("success"):
            images = result.get("images", [])
            self._index_options = []
            
            for img in images:
                idx = img.get("Index", "")
                name = img.get("Name", "")
                option = f"{idx}: {name}"
                
                # 排除另一槽位使用的 Index
                if exclude_indices and idx in exclude_indices:
                    continue
                
                self._index_options.append(option)
            
            self.combo_index.configure(values=self._index_options)
            
            if self._index_options:
                self.combo_index.set(self._index_options[0])
                self.var_index.set(self._index_options[0].split(":")[0])
            
            self._on_log(f"✓ WIM#{self.slot_number} 讀取到 {len(images)} 個映像")
        else:
            self._on_log(f"✗ 讀取失敗: {result.get('error', '未知錯誤')}")
    
    def _on_index_changed(self, selection):
        """Index 選擇變更"""
        if selection and ":" in selection:
            self.var_index.set(selection.split(":")[0])
    
    def _on_readonly_changed(self):
        """唯讀選項變更"""
        if self.var_readonly.get():
            self.var_commit.set(False)
            self.radio_commit.configure(state="disabled")
        else:
            self.radio_commit.configure(state="normal")
    
    def _on_browse_mount_dir(self):
        """選擇掛載目錄"""
        path = filedialog.askdirectory(title="選擇掛載目錄")
        if path:
            self.var_mount_dir.set(path)
    
    def _on_create_mount_dir(self):
        """建立掛載目錄"""
        path = self.var_mount_dir.get().strip()
        success, msg = create_mount_directory(path, f"WIM#{self.slot_number}")
        if success:
            self._on_log(f"✓ {msg}")
        else:
            messagebox.showwarning("提示", msg)
    
    def _on_open_mount_dir(self):
        """開啟掛載目錄"""
        path = self.var_mount_dir.get().strip()
        if path and os.path.isdir(path):
            open_directory(path)
        else:
            messagebox.showwarning("提示", "目錄不存在")
    
    def _on_check_status(self):
        """檢查掛載狀態"""
        mount_dir = self.var_mount_dir.get().strip()
        if not mount_dir:
            self._update_status("未設定", "default")
            return
        
        is_mounted, wim_path, status = WIMManager.is_path_mounted(mount_dir)
        
        if is_mounted:
            self._update_status("已掛載", "success")
            self._on_log(f"WIM#{self.slot_number} 狀態: 已掛載 - {wim_path}")
        elif status == "needs_remount":
            self._update_status("需重新掛載", "warning")
        else:
            self._update_status("未掛載", "default")
    
    def _update_status(self, text: str, status: str):
        """更新狀態顯示"""
        self.var_status.set(text)
        self.status_badge.set_status(text, status)
        
        # 更新按鈕狀態
        is_mounted = status == "success"
        self.btn_mount.configure(state="disabled" if is_mounted else "normal")
        self.btn_unmount.configure(state="normal" if is_mounted else "disabled")
        
        if self._on_status_change:
            self._on_status_change(self.slot_number, is_mounted)
    
    def _on_mount(self):
        """掛載 WIM"""
        wim_path = self.var_wim_path.get().strip()
        mount_dir = self.var_mount_dir.get().strip()
        index = self.var_index.get().strip()
        readonly = self.var_readonly.get()
        
        if not wim_path:
            messagebox.showwarning("提示", "請選擇 WIM 檔案")
            return
        if not mount_dir:
            messagebox.showwarning("提示", "請設定掛載目錄")
            return
        if not index:
            messagebox.showwarning("提示", "請選擇 Index")
            return
        
        self._on_log(f"正在掛載 WIM#{self.slot_number}...")
        self.btn_mount.set_loading(True)
        
        def do_mount():
            # mount_wim(wim_path, index, mount_dir, readonly) -> (success, message)
            success, message = WIMManager.mount_wim(wim_path, index, mount_dir, readonly)
            self.after(0, lambda: self._mount_complete(success, message))
        
        threading.Thread(target=do_mount, daemon=True).start()
    
    def _mount_complete(self, success: bool, message: str):
        """掛載完成回調"""
        self.btn_mount.set_loading(False)
        
        if success:
            self._on_log(f"✓ WIM#{self.slot_number} 掛載成功")
            self._update_status("已掛載", "success")
        else:
            self._on_log(f"✗ WIM#{self.slot_number} 掛載失敗: {message}")
            messagebox.showerror("掛載失敗", message or "未知錯誤")
    
    def _on_unmount(self):
        """卸載 WIM"""
        mount_dir = self.var_mount_dir.get().strip()
        commit = self.var_commit.get()
        
        if not mount_dir:
            messagebox.showwarning("提示", "請設定掛載目錄")
            return
        
        action = "提交變更" if commit else "丟棄變更"
        if not messagebox.askyesno("確認卸載", f"確定要卸載並{action}嗎？"):
            return
        
        self._on_log(f"正在卸載 WIM#{self.slot_number} ({action})...")
        self.btn_unmount.set_loading(True)
        
        def do_unmount():
            # unmount_wim(mount_dir, commit) -> (success, message)
            success, message = WIMManager.unmount_wim(mount_dir, commit)
            self.after(0, lambda: self._unmount_complete(success, message))
        
        threading.Thread(target=do_unmount, daemon=True).start()
    
    def _unmount_complete(self, success: bool, message: str):
        """卸載完成回調"""
        self.btn_unmount.set_loading(False)
        
        if success:
            self._on_log(f"✓ WIM#{self.slot_number} 卸載成功")
            self._update_status("未掛載", "default")
        else:
            self._on_log(f"✗ WIM#{self.slot_number} 卸載失敗: {message}")
            messagebox.showerror("卸載失敗", message)
    
    def _on_close_explorer(self):
        """關閉檔案總管"""
        mount_dir = self.var_mount_dir.get().strip()
        if mount_dir:
            WIMManager.close_explorer_for_path(mount_dir)
            self._on_log(f"已嘗試關閉 WIM#{self.slot_number} 的檔案總管視窗")
    
    def _on_smart_fix(self):
        """一鍵修復"""
        mount_dir = self.var_mount_dir.get().strip()
        self._on_log(f"開始 WIM#{self.slot_number} 智能修復...")
        
        def do_fix():
            # smart_cleanup_and_fix() -> (success, message)
            success, message = WIMManager.smart_cleanup_and_fix()
            self.after(0, lambda: self._fix_complete(success, message))
        
        threading.Thread(target=do_fix, daemon=True).start()
    
    def _fix_complete(self, success: bool, message: str):
        """修復完成回調"""
        if success:
            self._on_log(f"✓ WIM#{self.slot_number} 修復完成")
            self._on_check_status()
        else:
            self._on_log(f"✗ 修復過程有錯誤: {message}")
    
    # === 公開方法 ===
    
    def get_config(self) -> dict:
        """取得設定"""
        return {
            "wim_file": self.var_wim_path.get(),
            "mount_dir": self.var_mount_dir.get(),
            "index": self.var_index.get(),
            "readonly": self.var_readonly.get(),
            "commit": self.var_commit.get()
        }
    
    def set_config(self, config: dict):
        """設定配置"""
        if config.get("wim_file"):
            self.var_wim_path.set(config["wim_file"])
        if config.get("mount_dir"):
            self.var_mount_dir.set(config["mount_dir"])
        if config.get("index"):
            self.var_index.set(config["index"])
        if config.get("readonly") is not None:
            self.var_readonly.set(config["readonly"])
        if config.get("commit") is not None:
            self.var_commit.set(config["commit"])
        
        self._on_readonly_changed()
    
    def get_slot_info(self) -> tuple:
        """取得槽位資訊 (wim_path, index)"""
        return (self.var_wim_path.get().strip(), self.var_index.get().strip())
    
    def get_mount_dir(self) -> str:
        """取得掛載目錄"""
        return self.var_mount_dir.get().strip()
    
    def is_mounted(self) -> bool:
        """是否已掛載"""
        return self.var_status.get() == "已掛載"
    
    def check_status(self):
        """檢查掛載狀態（公開方法）"""
        self._on_check_status()


class WIMPage(ctk.CTkFrame):
    """
    WIM 掛載頁面
    包含兩個 WIM 掛載槽位
    """
    
    def __init__(
        self,
        master: Any,
        on_log: Callable[[str], None],
        on_mount_change: Optional[Callable] = None,
        **kwargs
    ):
        """
        初始化 WIM 頁面
        
        Args:
            master: 父組件
            on_log: 日誌回調
            on_mount_change: 掛載狀態變更回調
        """
        super().__init__(master, fg_color="transparent", **kwargs)
        
        self._on_log = on_log
        self._on_mount_change = on_mount_change
        
        self._build_ui()
    
    def _build_ui(self):
        """建立 UI"""
        # 頁面標題
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=20, pady=(20, 16))
        
        ctk.CTkLabel(
            header,
            text="WIM 映像掛載",
            font=Fonts.to_tuple(Fonts.TITLE),
            text_color=theme_manager.colors.text_primary
        ).pack(side="left")
        
        ctk.CTkLabel(
            header,
            text="管理 Windows 映像檔的掛載與卸載",
            font=Fonts.to_tuple(Fonts.BODY),
            text_color=theme_manager.colors.text_muted
        ).pack(side="left", padx=(12, 0))
        
        # 滾動區域
        scroll_frame = ctk.CTkScrollableFrame(
            self,
            fg_color="transparent"
        )
        scroll_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
        # WIM 槽位 #1 (預設展開)
        self.slot1 = WIMSlot(
            scroll_frame,
            slot_number=1,
            on_log=self._on_log,
            on_status_change=self._on_slot_status_change,
            get_other_slot_info=lambda: self.slot2.get_slot_info() if hasattr(self, 'slot2') else None,
            expanded=True
        )
        self.slot1.pack(fill="x")
        
        # WIM 槽位 #2 (預設摺疊)
        self.slot2 = WIMSlot(
            scroll_frame,
            slot_number=2,
            on_log=self._on_log,
            on_status_change=self._on_slot_status_change,
            get_other_slot_info=lambda: self.slot1.get_slot_info(),
            expanded=False
        )
        self.slot2.pack(fill="x")
    
    def _on_slot_status_change(self, slot_number: int, is_mounted: bool):
        """槽位狀態變更"""
        if self._on_mount_change:
            self._on_mount_change()
    
    # === 公開方法 ===
    
    def get_config(self) -> dict:
        """取得所有設定"""
        return {
            "WIM": self.slot1.get_config(),
            "WIM2": self.slot2.get_config()
        }
    
    def set_config(self, config: dict):
        """設定配置"""
        if "WIM" in config:
            self.slot1.set_config(config["WIM"])
        if "WIM2" in config:
            self.slot2.set_config(config["WIM2"])
    
    def get_mounted_dirs(self) -> list:
        """取得所有已掛載的目錄"""
        dirs = []
        if self.slot1.is_mounted():
            dirs.append(("WIM#1", self.slot1.get_mount_dir()))
        if self.slot2.is_mounted():
            dirs.append(("WIM#2", self.slot2.get_mount_dir()))
        return dirs
    
    def check_all_status(self):
        """檢查所有槽位狀態"""
        self.slot1._on_check_status()
        self.after(500, self.slot2._on_check_status)
