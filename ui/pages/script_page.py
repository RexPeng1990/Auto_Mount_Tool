# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WinPE 腳本管理頁面
管理 startnet.cmd 啟動腳本
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
from typing import Optional, Callable, List
import os
import shutil

from ui.theme import ThemeManager, Fonts, theme_manager
from ui.components import ModernButton, ModernCard, StatusBadge


class ScriptPage(ctk.CTkFrame):
    """WinPE 腳本管理頁面"""
    
    STARTNET_PATH = "Windows\\System32\\startnet.cmd"
    BACKUP_SUFFIX = ".backup"
    
    def __init__(
        self,
        master,
        get_mounted_dirs: Callable[[], List[tuple]],
        on_log: Optional[Callable[[str], None]] = None,
        **kwargs
    ):
        super().__init__(master, fg_color="transparent", **kwargs)
        
        self._get_mounted_dirs = get_mounted_dirs
        self._on_log = on_log or print
        self._current_file_path = ""
        self._original_content = ""
        
        self._build_ui()
    
    def _build_ui(self):
        """建立 UI"""
        # 標題
        ctk.CTkLabel(
            self,
            text="WinPE 腳本管理",
            font=("Microsoft JhengHei UI", 24, "bold"),
            text_color=theme_manager.colors.text_primary
        ).pack(anchor="w", pady=(0, 16))
        
        # 目標選擇區
        self._build_target_section()
        
        # 編輯器區
        self._build_editor_section()
    
    def _build_target_section(self):
        """建立目標選擇區"""
        target_card = ModernCard(self, title="目標檔案")
        target_card.pack(fill="x", pady=(0, 12))
        
        target_row = ctk.CTkFrame(target_card.content_frame, fg_color="transparent")
        target_row.pack(fill="x", pady=4)
        
        ctk.CTkLabel(
            target_row,
            text="掛載映像",
            font=Fonts.to_tuple(Fonts.LABEL),
            width=80
        ).pack(side="left")
        
        self.var_target = ctk.StringVar()
        self.combo_target = ctk.CTkComboBox(
            target_row,
            variable=self.var_target,
            values=[""],
            width=350,
            state="readonly",
            command=self._on_target_changed
        )
        self.combo_target.pack(side="left", padx=(8, 0))
        
        self.status_badge = StatusBadge(target_row, "未選擇", "default")
        self.status_badge.pack(side="left", padx=(12, 0))
        
        ModernButton(
            target_row,
            text="重新整理",
            variant="outline",
            size="sm",
            command=self.refresh_targets
        ).pack(side="left", padx=(12, 0))
        
        ModernButton(
            target_row,
            text="開啟檔案位置",
            variant="ghost",
            size="sm",
            command=self._open_file_location
        ).pack(side="right")
        
        # 顯示檔案路徑
        self.lbl_path = ctk.CTkLabel(
            target_card.content_frame,
            text="",
            font=Fonts.to_tuple(Fonts.BODY_SMALL),
            text_color="#888888",
            anchor="w"
        )
        self.lbl_path.pack(fill="x", pady=(4, 0))
    
    def _build_editor_section(self):
        """建立編輯器區"""
        editor_card = ModernCard(self, title="startnet.cmd 編輯器")
        editor_card.pack(fill="both", expand=True)
        
        # 工具列
        toolbar = ctk.CTkFrame(editor_card.content_frame, fg_color="transparent")
        toolbar.pack(fill="x", pady=(0, 8))
        
        self.btn_save = ModernButton(
            toolbar,
            text="💾 儲存",
            variant="primary",
            size="sm",
            command=self._save_file
        )
        self.btn_save.pack(side="left")
        
        self.btn_reload = ModernButton(
            toolbar,
            text="🔄 重新載入",
            variant="outline",
            size="sm",
            command=self._reload_file
        )
        self.btn_reload.pack(side="left", padx=(8, 0))
        
        # 分隔線
        ctk.CTkFrame(toolbar, width=1, height=24, fg_color="#e0e0e0").pack(side="left", padx=16)
        
        self.btn_backup = ModernButton(
            toolbar,
            text="📋 建立備份",
            variant="outline",
            size="sm",
            command=self._backup_file
        )
        self.btn_backup.pack(side="left")
        
        self.btn_restore = ModernButton(
            toolbar,
            text="↩️ 還原備份",
            variant="ghost",
            size="sm",
            command=self._restore_backup,
            state="disabled"
        )
        self.btn_restore.pack(side="left", padx=(8, 0))
        
        # 修改狀態
        self.lbl_status = ctk.CTkLabel(
            toolbar,
            text="",
            font=Fonts.to_tuple(Fonts.BODY_SMALL),
            text_color="#888888"
        )
        self.lbl_status.pack(side="right")
        
        # 快捷鍵提示
        ctk.CTkLabel(
            toolbar,
            text="Ctrl+S 儲存",
            font=Fonts.to_tuple(Fonts.BODY_SMALL),
            text_color="#aaaaaa"
        ).pack(side="right", padx=(0, 16))
        
        # 文字編輯器
        editor_container = ctk.CTkFrame(editor_card.content_frame, fg_color="#1e1e1e", corner_radius=8)
        editor_container.pack(fill="both", expand=True)
        
        # 使用 Text widget
        self.text_editor = tk.Text(
            editor_container,
            wrap="none",
            font=("Consolas", 11),
            bg="#1e1e1e",
            fg="#d4d4d4",
            insertbackground="#ffffff",
            selectbackground="#264f78",
            selectforeground="#ffffff",
            padx=12,
            pady=12,
            undo=True
        )
        
        # 滾動條
        scroll_y = ctk.CTkScrollbar(editor_container, command=self.text_editor.yview)
        scroll_x = ctk.CTkScrollbar(editor_container, orientation="horizontal", command=self.text_editor.xview)
        self.text_editor.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        self.text_editor.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        
        editor_container.grid_rowconfigure(0, weight=1)
        editor_container.grid_columnconfigure(0, weight=1)
        
        # 綁定事件
        self.text_editor.bind("<<Modified>>", self._on_text_modified)
        self.text_editor.bind("<Control-s>", lambda e: self._save_file())
    
    # === 事件處理 ===
    
    def _on_target_changed(self, selection: str):
        """目標變更"""
        mount_dir = ""
        if selection and " - " in selection:
            parts = selection.split(" - ")
            if len(parts) >= 2:
                mount_dir = parts[1].split(" (")[0].strip()
        
        if mount_dir and os.path.isdir(mount_dir):
            startnet_path = os.path.join(mount_dir, self.STARTNET_PATH)
            
            if os.path.exists(startnet_path):
                self._current_file_path = startnet_path
                self._load_file()
                self.status_badge.set_status("已載入", "success")
                self.lbl_path.configure(text=f"📁 {startnet_path}")
                
                # 檢查是否有備份
                backup_path = startnet_path + self.BACKUP_SUFFIX
                if os.path.exists(backup_path):
                    self.btn_restore.configure(state="normal")
                else:
                    self.btn_restore.configure(state="disabled")
            else:
                self._current_file_path = ""
                self.text_editor.delete("1.0", "end")
                self.status_badge.set_status("檔案不存在", "warning")
                self.lbl_path.configure(text=f"⚠️ 找不到: {startnet_path}")
                self._on_log(f"找不到 startnet.cmd: {startnet_path}")
        else:
            self._current_file_path = ""
            self.text_editor.delete("1.0", "end")
            self.status_badge.set_status("未選擇", "default")
            self.lbl_path.configure(text="")
    
    def _on_text_modified(self, event=None):
        """文字修改事件"""
        if self.text_editor.edit_modified():
            current = self.text_editor.get("1.0", "end-1c")
            if current != self._original_content:
                self.lbl_status.configure(text="● 已修改", text_color="#ffa500")
            else:
                self.lbl_status.configure(text="")
            self.text_editor.edit_modified(False)
    
    def _load_file(self):
        """載入檔案"""
        if not self._current_file_path:
            return
        
        try:
            with open(self._current_file_path, 'r', encoding='utf-8', errors='replace') as f:
                content = f.read()
            
            self.text_editor.delete("1.0", "end")
            self.text_editor.insert("1.0", content)
            self._original_content = content
            self.lbl_status.configure(text="")
            self.text_editor.edit_modified(False)
            
            self._on_log(f"已載入: {self._current_file_path}")
        except Exception as e:
            messagebox.showerror("錯誤", f"載入檔案失敗: {e}")
            self._on_log(f"✗ 載入失敗: {e}")
    
    def _save_file(self):
        """儲存檔案"""
        if not self._current_file_path:
            messagebox.showwarning("提示", "請先選擇目標映像")
            return
        
        try:
            content = self.text_editor.get("1.0", "end-1c")
            
            with open(self._current_file_path, 'w', encoding='utf-8') as f:
                f.write(content)
            
            self._original_content = content
            self.lbl_status.configure(text="✓ 已儲存", text_color="#4caf50")
            self.text_editor.edit_modified(False)
            
            self._on_log(f"✓ 已儲存: {self._current_file_path}")
            
            # 2秒後清除狀態
            self.after(2000, lambda: self.lbl_status.configure(text=""))
        except Exception as e:
            messagebox.showerror("錯誤", f"儲存失敗: {e}")
            self._on_log(f"✗ 儲存失敗: {e}")
    
    def _reload_file(self):
        """重新載入檔案"""
        if not self._current_file_path:
            return
        
        current = self.text_editor.get("1.0", "end-1c")
        if current != self._original_content:
            if not messagebox.askyesno("確認", "有未儲存的修改，確定要重新載入嗎？"):
                return
        
        self._load_file()
    
    def _backup_file(self):
        """備份檔案"""
        if not self._current_file_path:
            messagebox.showwarning("提示", "請先選擇目標映像")
            return
        
        try:
            backup_path = self._current_file_path + self.BACKUP_SUFFIX
            shutil.copy2(self._current_file_path, backup_path)
            
            self.btn_restore.configure(state="normal")
            messagebox.showinfo("完成", f"已建立備份:\n{backup_path}")
            self._on_log(f"✓ 已備份: {backup_path}")
        except Exception as e:
            messagebox.showerror("錯誤", f"備份失敗: {e}")
            self._on_log(f"✗ 備份失敗: {e}")
    
    def _restore_backup(self):
        """還原備份"""
        if not self._current_file_path:
            return
        
        backup_path = self._current_file_path + self.BACKUP_SUFFIX
        if not os.path.exists(backup_path):
            messagebox.showwarning("提示", "找不到備份檔案")
            return
        
        if not messagebox.askyesno("確認", "確定要還原備份嗎？\n目前的內容將被覆蓋！"):
            return
        
        try:
            shutil.copy2(backup_path, self._current_file_path)
            self._load_file()
            messagebox.showinfo("完成", "已還原備份")
            self._on_log("✓ 已還原備份")
        except Exception as e:
            messagebox.showerror("錯誤", f"還原失敗: {e}")
            self._on_log(f"✗ 還原失敗: {e}")
    
    def _open_file_location(self):
        """開啟檔案位置"""
        if not self._current_file_path:
            messagebox.showwarning("提示", "請先選擇目標映像")
            return
        
        folder = os.path.dirname(self._current_file_path)
        if os.path.isdir(folder):
            os.startfile(folder)
    
    # === 公開方法 ===
    
    def refresh_targets(self):
        """重新整理目標選項"""
        mounted_dirs = self._get_mounted_dirs()
        
        from app.wim_manager import WIMManager
        
        options = []
        for name, path in mounted_dirs:
            if path:
                is_mounted, _, _ = WIMManager.is_path_mounted(path)
                status = "✓ 已掛載" if is_mounted else "○ 未掛載"
                options.append(f"{name} - {path} ({status})")
            else:
                options.append(f"{name} - (未設定)")
        
        if not options:
            options = ["(無可用映像)"]
        
        self.combo_target.configure(values=options)
        
        # 自動選擇第一個已掛載的
        for opt in options:
            if "✓ 已掛載" in opt:
                self.combo_target.set(opt)
                self._on_target_changed(opt)
                break
        else:
            if options:
                self.combo_target.set(options[0])
                self._on_target_changed(options[0])
