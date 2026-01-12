# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
驅動程式管理頁面
現代化 UI 設計
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Optional, Callable, Any, List, Dict
import os
import threading

from ui.theme import Fonts, theme_manager
from ui.components import (
    ModernButton, ModernCard, ModernEntry, ModernTooltip, 
    StatusBadge, ModernComboBox
)
from ui.widgets.driver_table import DriverTable
from app.driver_manager import DriverManager
from app.config import DRIVER_EXPORT_DIR, ensure_output_dirs

# 用於追蹤當前選擇的目標
_DRIVER_PAGE_CURRENT_TARGET = ""


class DriverPage(ctk.CTkFrame):
    """
    驅動程式管理頁面
    """
    
    def __init__(
        self,
        master: Any,
        on_log: Callable[[str], None],
        get_mounted_dirs: Callable[[], List[tuple]],
        show_header: bool = True,
        **kwargs
    ):
        """
        初始化驅動程式頁面
        
        Args:
            master: 父組件
            on_log: 日誌回調
            get_mounted_dirs: 取得已掛載目錄的回調 [("WIM#1", path), ...]
            show_header: 是否顯示頁面標題
        """
        super().__init__(master, fg_color="transparent", **kwargs)
        
        self._on_log = on_log
        self._get_mounted_dirs = get_mounted_dirs
        self._show_header = show_header
        
        # 變數
        self.var_target_mount = ctk.StringVar()
        self.var_search = ctk.StringVar()
        self.var_status = ctk.StringVar(value="請選擇已掛載的映像")
        
        self._build_ui()
    
    def _build_ui(self):
        """建立 UI"""
        # 頁面標題（可選）
        if self._show_header:
            header = ctk.CTkFrame(self, fg_color="transparent")
            header.pack(fill="x", padx=20, pady=(20, 16))
            
            ctk.CTkLabel(
                header,
                text="驅動程式管理",
                font=Fonts.to_tuple(Fonts.TITLE),
                text_color=theme_manager.colors.text_primary
            ).pack(side="left")
            
            ctk.CTkLabel(
                header,
                text="安裝、移除、萃取映像中的驅動程式",
                font=Fonts.to_tuple(Fonts.BODY),
                text_color=theme_manager.colors.text_muted
            ).pack(side="left", padx=(12, 0))
        
        # 根據是否顯示標題來調整 padding
        pad_x = 20 if self._show_header else 0
        pad_y = (0, 12) if self._show_header else (0, 8)
        
        # === 目標映像選擇區 ===
        target_card = ModernCard(self, padding=12)
        target_card.pack(fill="x", padx=pad_x, pady=pad_y)
        target_content = target_card.get_content_frame()
        
        target_row = ctk.CTkFrame(target_content, fg_color="transparent")
        target_row.pack(fill="x")
        
        ctk.CTkLabel(
            target_row,
            text="目標映像",
            font=Fonts.to_tuple(Fonts.LABEL),
            text_color=theme_manager.colors.text_secondary,
            width=80,
            anchor="w"
        ).pack(side="left")
        
        self.combo_target = ModernComboBox(
            target_row,
            width=300,
            variable=self.var_target_mount,
            values=[""],
            command=self._on_target_changed
        )
        self.combo_target.pack(side="left", padx=(8, 0))
        
        self.status_badge = StatusBadge(target_row, "未選擇", "default")
        self.status_badge.pack(side="left", padx=(12, 0))
        
        # 重新整理按鈕
        ModernButton(
            target_row,
            text="重新整理",
            variant="outline",
            size="sm",
            command=self._refresh_driver_list
        ).pack(side="left", padx=(12, 0))
        
        # 搜尋
        search_frame = ctk.CTkFrame(target_row, fg_color="transparent")
        search_frame.pack(side="right")
        
        ctk.CTkLabel(
            search_frame,
            text="🔍",
            font=Fonts.to_tuple(Fonts.BODY)
        ).pack(side="left")
        
        self.entry_search = ModernEntry(
            search_frame,
            placeholder="搜尋驅動...",
            width=180
        )
        self.entry_search.pack(side="left", padx=(4, 0))
        self.entry_search.configure(textvariable=self.var_search)
        self.entry_search.bind("<Return>", lambda e: self._on_search())
        self.entry_search.bind("<Escape>", lambda e: self._on_clear_search())
        
        ModernButton(
            search_frame,
            text="搜尋",
            variant="ghost",
            size="sm",
            command=self._on_search
        ).pack(side="left", padx=(4, 0))
        
        # === 驅動程式表格 ===
        table_frame = ctk.CTkFrame(self, fg_color="transparent")
        table_frame.pack(fill="both", expand=True, padx=20, pady=(0, 12))
        
        self.driver_table = DriverTable(
            table_frame,
            on_selection_change=self._on_selection_change
        )
        self.driver_table.pack(fill="both", expand=True)
        
        # === 操作按鈕區 ===
        btn_card = ModernCard(self, padding=12)
        btn_card.pack(fill="x", padx=20, pady=(0, 20))
        btn_content = btn_card.get_content_frame()
        
        btn_row = ctk.CTkFrame(btn_content, fg_color="transparent")
        btn_row.pack(fill="x")
        
        # 左側主要操作
        left_btns = ctk.CTkFrame(btn_row, fg_color="transparent")
        left_btns.pack(side="left")
        
        self.btn_extract = ModernButton(
            left_btns,
            text="📤 萃取驅動",
            variant="primary",
            command=self._on_extract_drivers
        )
        self.btn_extract.pack(side="left")
        ModernTooltip(self.btn_extract, "將選取的驅動程式萃取到指定目錄")
        
        self.btn_add = ModernButton(
            left_btns,
            text="➕ 新增驅動",
            variant="secondary",
            command=self._on_add_driver
        )
        self.btn_add.pack(side="left", padx=(12, 0))
        ModernTooltip(self.btn_add, "安裝新的驅動程式到映像中")
        
        self.btn_remove = ModernButton(
            left_btns,
            text="🗑 移除驅動",
            variant="danger",
            command=self._on_remove_drivers
        )
        self.btn_remove.pack(side="left", padx=(12, 0))
        ModernTooltip(self.btn_remove, "從映像中移除選取的驅動程式")
        
        # 右側輔助操作
        right_btns = ctk.CTkFrame(btn_row, fg_color="transparent")
        right_btns.pack(side="right")
        
        # 狀態文字
        self.lbl_status = ctk.CTkLabel(
            right_btns,
            textvariable=self.var_status,
            font=Fonts.to_tuple(Fonts.CAPTION),
            text_color=theme_manager.colors.text_muted
        )
        self.lbl_status.pack(side="left", padx=(0, 16))
        
        ModernButton(
            right_btns,
            text="查看詳情",
            variant="ghost",
            size="sm",
            command=self._on_view_details
        ).pack(side="left")
        
        ModernButton(
            right_btns,
            text="匯出清單",
            variant="ghost",
            size="sm",
            command=self._on_export_list
        ).pack(side="left", padx=(8, 0))
        
        # 初始化按鈕狀態
        self._update_button_states()
    
    # === 事件處理 ===
    
    def _on_target_changed(self, selection: str):
        """目標映像選擇變更"""
        global _DRIVER_PAGE_CURRENT_TARGET
        
        # 先清空前一次的結果
        self.driver_table.clear()
        self.var_status.set("讀取中...")
        
        # 解析選擇
        mount_dir = self._extract_mount_dir(selection)
        
        # 更新追蹤變數
        _DRIVER_PAGE_CURRENT_TARGET = mount_dir
        
        if mount_dir and os.path.isdir(mount_dir):
            # 檢查掛載狀態
            from app.wim_manager import WIMManager
            is_mounted, _, _ = WIMManager.is_path_mounted(mount_dir)
            
            if is_mounted:
                self.status_badge.set_status("已掛載", "success")
                self._refresh_driver_list()
            else:
                self.status_badge.set_status("未掛載", "default")
                self.driver_table.clear()
                self.var_status.set("映像未掛載，請先掛載後再操作")
        else:
            self.status_badge.set_status("未選擇", "default")
            self.driver_table.clear()
            self.var_status.set("請選擇已掛載的映像")
        
        self._update_button_states()
    
    def _on_selection_change(self, selected_count: int):
        """選取變更"""
        total, selected = self.driver_table.get_count()
        self.var_status.set(f"共 {total} 個驅動程式，已選取 {selected} 個")
    
    def _on_search(self):
        """執行搜尋"""
        text = self.var_search.get().strip()
        self.driver_table.search(text)
        if text:
            self._on_log(f"搜尋: {text}")
    
    def _on_clear_search(self):
        """清除搜尋"""
        self.var_search.set("")
        self.driver_table.clear_search()
    
    def _refresh_driver_list(self):
        """重新整理驅動程式列表"""
        mount_dir = self._get_current_mount_dir()
        if not mount_dir:
            messagebox.showwarning("提示", "請先選擇目標映像")
            return
        
        self._on_log("正在讀取驅動程式列表...")
        self.var_status.set("正在讀取...")
        
        def do_refresh():
            success, drivers, error = DriverManager.get_drivers_in_offline_image(mount_dir)
            self.after(0, lambda: self._refresh_complete(success, drivers, error))
        
        threading.Thread(target=do_refresh, daemon=True).start()
    
    def _refresh_complete(self, success: bool, drivers: List[Dict], error: str):
        """重新整理完成"""
        if success:
            self.driver_table.set_drivers(drivers)
            total, _ = self.driver_table.get_count()
            self.var_status.set(f"共 {total} 個驅動程式")
            self._on_log(f"✓ 已載入 {total} 個驅動程式")
        else:
            self.driver_table.clear()
            self.var_status.set(f"讀取失敗: {error}")
            self._on_log(f"✗ 讀取驅動程式失敗: {error}")
        
        self._update_button_states()
    
    def _on_extract_drivers(self):
        """萃取驅動程式"""
        mount_dir = self._get_current_mount_dir()
        if not mount_dir:
            messagebox.showwarning("提示", "請先選擇目標映像")
            return
        
        selected = self.driver_table.get_selected_drivers()
        
        if not selected:
            messagebox.showinfo("提示", "請先勾選要萃取的驅動程式")
            return
        
        # 自動使用預設輸出目錄
        ensure_output_dirs()
        output_dir = DRIVER_EXPORT_DIR
        
        self.btn_extract.set_loading(True)
        
        def do_extract():
            # 萃取選定的驅動
            self._on_log(f"正在萃取 {len(selected)} 個驅動程式到: {output_dir}")
            
            def callback(current, total, name, success, msg):
                status = "✓" if success else "✗"
                self.after(0, lambda: self._on_log(f"  {status} [{current}/{total}] {name}"))
            
            success_count, fail_count, errors = DriverManager.export_selected_drivers(
                mount_dir, output_dir, selected, callback
            )
            self.after(0, lambda: self._extract_complete(success_count, fail_count, errors, output_dir))
        
        threading.Thread(target=do_extract, daemon=True).start()
    
    def _extract_complete(self, success_count: int, fail_count: int, errors: List[str], output_dir: str):
        """萃取完成"""
        self.btn_extract.set_loading(False)
        
        self._on_log(f"萃取完成: 成功 {success_count}，失敗 {fail_count}")
        
        if fail_count > 0:
            error_msg = "\n".join(errors[:5])
            if len(errors) > 5:
                error_msg += f"\n... 還有 {len(errors) - 5} 個錯誤"
            messagebox.showwarning("部分失敗", f"成功: {success_count}\n失敗: {fail_count}\n\n{error_msg}")
        else:
            if messagebox.askyesno("完成", f"已成功萃取 {success_count} 個驅動程式到：\n{output_dir}\n\n是否開啟輸出目錄？"):
                os.startfile(output_dir)
    
    def _on_add_driver(self):
        """新增驅動程式 - 簡化操作：直接選擇檔案或資料夾"""
        mount_dir = self._get_current_mount_dir()
        if not mount_dir:
            messagebox.showwarning("提示", "請先選擇目標映像")
            return
        
        # 詢問選擇方式
        choice = messagebox.askquestion(
            "新增驅動程式",
            "選擇驅動來源類型：\n\n" +
            "「是」= 選擇資料夾（自動搜尋所有 .inf）\n" +
            "「否」= 選擇單一 .inf 檔案",
            icon="question"
        )
        
        if choice == "yes":
            # 選擇資料夾
            path = filedialog.askdirectory(title="選擇驅動程式資料夾")
            if path:
                self._do_add_driver(mount_dir, path, recurse=True, force=False)
        else:
            # 選擇檔案
            path = filedialog.askopenfilename(
                title="選擇驅動程式",
                filetypes=[("INF 檔案", "*.inf"), ("所有檔案", "*.*")]
            )
            if path:
                self._do_add_driver(mount_dir, path, recurse=False, force=False)
    
    def _do_add_driver(self, mount_dir: str, source_path: str, recurse: bool, force: bool):
        """執行新增驅動"""
        self._on_log(f"正在安裝驅動: {source_path}")
        self.btn_add.set_loading(True)
        
        def add_thread():
            success, msg = DriverManager.add_driver_to_offline_image(
                mount_dir, source_path,
                recurse=recurse,
                force_unsigned=force
            )
            self.after(0, lambda: add_complete(success, msg))
        
        def add_complete(success, msg):
            self.btn_add.set_loading(False)
            if success:
                self._on_log("✓ 驅動程式安裝完成")
                messagebox.showinfo("完成", "驅動程式安裝完成！")
                self._refresh_driver_list()
            else:
                self._on_log(f"✗ 安裝失敗: {msg}")
                messagebox.showerror("錯誤", f"安裝失敗: {msg}")
        
        threading.Thread(target=add_thread, daemon=True).start()
    
    def _on_remove_drivers(self):
        """移除驅動程式"""
        mount_dir = self._get_current_mount_dir()
        if not mount_dir:
            messagebox.showwarning("提示", "請先選擇目標映像")
            return
        
        selected = self.driver_table.get_selected_published_names()
        if not selected:
            messagebox.showwarning("提示", "請先選取要移除的驅動程式")
            return
        
        if not messagebox.askyesno("確認", f"確定要移除 {len(selected)} 個驅動程式嗎？\n此操作無法復原！"):
            return
        
        self._on_log(f"正在移除 {len(selected)} 個驅動程式...")
        self.btn_remove.set_loading(True)
        
        def do_remove():
            def callback(current, total, name, success, msg):
                status = "✓" if success else "✗"
                self.after(0, lambda: self._on_log(f"  {status} [{current}/{total}] {name}"))
            
            success_count, fail_count, errors = DriverManager.remove_drivers_batch(
                mount_dir, selected, callback
            )
            self.after(0, lambda: self._remove_complete(success_count, fail_count, errors))
        
        threading.Thread(target=do_remove, daemon=True).start()
    
    def _remove_complete(self, success_count: int, fail_count: int, errors: List[str]):
        """移除完成"""
        self.btn_remove.set_loading(False)
        
        self._on_log(f"移除完成: 成功 {success_count}，失敗 {fail_count}")
        
        if fail_count > 0:
            error_msg = "\n".join(errors[:5])
            if len(errors) > 5:
                error_msg += f"\n... 還有 {len(errors) - 5} 個錯誤"
            messagebox.showwarning("部分失敗", f"成功: {success_count}\n失敗: {fail_count}\n\n{error_msg}")
        else:
            messagebox.showinfo("完成", f"已成功移除 {success_count} 個驅動程式")
        
        self._refresh_driver_list()
    
    def _on_view_details(self):
        """查看詳情"""
        selected = self.driver_table.get_selected_drivers()
        if not selected:
            messagebox.showinfo("提示", "請先選取驅動程式")
            return
        
        # 顯示第一個選取的詳情
        driver = selected[0]
        details = "\n".join([f"{k}: {v}" for k, v in driver.items()])
        messagebox.showinfo("驅動程式詳情", details)
    
    def _on_export_list(self):
        """匯出清單"""
        total, _ = self.driver_table.get_count()
        if total == 0:
            messagebox.showinfo("提示", "沒有驅動程式可匯出")
            return
        
        path = filedialog.asksaveasfilename(
            title="匯出驅動清單",
            defaultextension=".txt",
            filetypes=[("文字檔", "*.txt"), ("CSV", "*.csv")]
        )
        if not path:
            return
        
        try:
            with open(path, "w", encoding="utf-8") as f:
                drivers = self.driver_table._filtered_drivers
                f.write("PublishedName\tOriginalFileName\tProvider\tVersion\tDate\tClassName\n")
                for d in drivers:
                    f.write(f"{d.get('PublishedName', '')}\t")
                    f.write(f"{d.get('OriginalFileName', '')}\t")
                    f.write(f"{d.get('Provider', '')}\t")
                    f.write(f"{d.get('Version', '')}\t")
                    f.write(f"{d.get('Date', '')}\t")
                    f.write(f"{d.get('ClassName', '')}\n")
            
            self._on_log(f"✓ 驅動清單已匯出至: {path}")
            messagebox.showinfo("完成", "清單匯出成功！")
        except Exception as e:
            self._on_log(f"✗ 匯出失敗: {e}")
            messagebox.showerror("錯誤", f"匯出失敗: {e}")
    
    def _get_current_mount_dir(self) -> str:
        """取得當前選擇的掛載目錄"""
        selection = self.var_target_mount.get()
        if selection:
            parts = selection.split(" - ")
            if len(parts) >= 2:
                path_part = parts[1].split(" (")[0]
                return path_part.strip()
        return ""
    
    def _update_button_states(self):
        """更新按鈕狀態"""
        mount_dir = self._get_current_mount_dir()
        has_mount = bool(mount_dir and os.path.isdir(mount_dir))
        
        # 檢查是否已掛載
        is_mounted = False
        if has_mount:
            from app.wim_manager import WIMManager
            is_mounted, _, _ = WIMManager.is_path_mounted(mount_dir)
        
        # 檢查是否為唯讀模式
        is_readonly = self._get_current_readonly()
        
        # 萃取按鈕：只要已掛載就可用
        extract_state = "normal" if is_mounted else "disabled"
        self.btn_extract.configure(state=extract_state)
        
        # 新增/移除按鈕：已掛載且非唯讀才可用
        modify_state = "normal" if (is_mounted and not is_readonly) else "disabled"
        self.btn_add.configure(state=modify_state)
        self.btn_remove.configure(state=modify_state)
    
    def _get_current_readonly(self) -> bool:
        """取得當前選擇映像的唯讀狀態"""
        selection = self.combo_target.get()
        if not selection or "WIM#" not in selection:
            return True  # 預設為唯讀（安全起見）
        
        # 從選項中提取 WIM 編號
        try:
            wim_num = int(selection.split("WIM#")[1].split(" ")[0])
        except (ValueError, IndexError):
            return True
        
        # 從 _get_mounted_dirs 獲取唯讀狀態
        mounted_dirs = self._get_mounted_dirs()
        for name, path, readonly in mounted_dirs:
            if f"WIM#{wim_num}" == name:
                return readonly
        
        return True  # 找不到則預設為唯讀
    
    # === 公開方法 ===
    
    def refresh_targets(self, auto_load: bool = False):
        """
        重新整理目標映像選項
        
        Args:
            auto_load: 是否自動載入驅動（僅在目標變更時）
        """
        global _DRIVER_PAGE_CURRENT_TARGET
        
        mounted_dirs = self._get_mounted_dirs()
        
        from app.wim_manager import WIMManager
        
        options = []
        first_mounted_idx = -1
        for i, item in enumerate(mounted_dirs):
            # 支援舊格式 (name, path) 和新格式 (name, path, readonly)
            if len(item) >= 3:
                name, path, readonly = item[0], item[1], item[2]
            else:
                name, path = item[0], item[1]
                readonly = True
            
            if path:
                is_mounted, _, _ = WIMManager.is_path_mounted(path)
                status = "✓ 已掛載" if is_mounted else "○ 未掛載"
                opt = f"{name} - {path} ({status})"
                options.append(opt)
                if is_mounted and first_mounted_idx < 0:
                    first_mounted_idx = i
            else:
                options.append(f"{name} - (未設定)")
        
        if not options:
            options = ["(無可用映像)"]
        
        self.combo_target.configure(values=options)
        
        # 取得當前選擇或自動選擇第一個已掛載的
        current_selection = self.combo_target.get()
        new_selection = ""
        
        # 如果當前選擇仍在選項中，保持不變
        if current_selection in options:
            new_selection = current_selection
        # 否則選擇第一個已掛載的
        elif first_mounted_idx >= 0:
            new_selection = options[first_mounted_idx]
        elif options:
            new_selection = options[0]
        
        if new_selection:
            self.combo_target.set(new_selection)
            
            # 提取 mount_dir
            mount_dir = self._extract_mount_dir(new_selection)
            
            # 只有在目標真正變更時才載入驅動
            if auto_load and mount_dir and mount_dir != _DRIVER_PAGE_CURRENT_TARGET:
                _DRIVER_PAGE_CURRENT_TARGET = mount_dir
                self._on_target_changed(new_selection)
            elif not auto_load:
                # 只更新狀態，不載入驅動
                self._update_status_only(new_selection)
    
    def _extract_mount_dir(self, selection: str) -> str:
        """從選項字串中提取掛載目錄"""
        if selection and " - " in selection:
            parts = selection.split(" - ")
            if len(parts) >= 2:
                return parts[1].split(" (")[0].strip()
        return ""
    
    def _update_status_only(self, selection: str):
        """只更新狀態顯示，不載入驅動"""
        mount_dir = self._extract_mount_dir(selection)
        
        if mount_dir and os.path.isdir(mount_dir):
            from app.wim_manager import WIMManager
            is_mounted, _, _ = WIMManager.is_path_mounted(mount_dir)
            
            if is_mounted:
                self.status_badge.set_status("已掛載", "success")
            else:
                self.status_badge.set_status("未掛載", "default")
        else:
            self.status_badge.set_status("未選擇", "default")
    
    def on_wim_unmounted(self, unmounted_slot: int, switch_to_slot: int = None):
        """
        WIM 卸載時的處理
        
        Args:
            unmounted_slot: 被卸載的 WIM 槽位編號
            switch_to_slot: 要切換到的槽位編號（如果有另一個已掛載）
        """
        global _DRIVER_PAGE_CURRENT_TARGET
        
        # 清除驅動清單
        self.driver_table.clear()
        
        # 重新整理選項
        self.refresh_targets(auto_load=False)
        
        if switch_to_slot:
            # 自動選擇另一個已掛載的 WIM
            options = self.combo_target.cget("values")
            for opt in options:
                if f"WIM#{switch_to_slot}" in opt and "✓ 已掛載" in opt:
                    self.combo_target.set(opt)
                    # 觸發載入
                    self._on_target_changed(opt)
                    break
        else:
            # 沒有其他已掛載的，顯示空狀態
            _DRIVER_PAGE_CURRENT_TARGET = ""
            self.status_badge.set_status("未掛載", "default")
            self.var_status.set("請掛載映像後操作")
            self._update_button_states()
        
        self._update_button_states()
